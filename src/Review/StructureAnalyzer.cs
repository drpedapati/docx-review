using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace DocxReview.Review;

public sealed class StructureAnalyzer
{
    private const int StructureMaxInputChars = 100_000;
    private static readonly TimeSpan StructureTimeout = TimeSpan.FromSeconds(30);
    private static readonly Regex MultiLineSpacing = new(@"\n{3,}", RegexOptions.Compiled);

    private readonly IResponsesClient? _client;
    private readonly string _model;
    private readonly JsonElement _schema;

    public StructureAnalyzer(IResponsesClient? client, string model)
    {
        _client = client;
        _model = string.IsNullOrWhiteSpace(model) ? ReviewOptions.DefaultStructureModel : model.Trim();
        _schema = ReviewEmbeddedResources.LoadSchemaElement("structure-analysis.schema.json");
    }

    public async Task<StructureAnalysisResult> AnalyzeAsync(
        ReviewTextDocument document,
        DocumentProfile profile,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(document);

        if (string.IsNullOrWhiteSpace(document.Text))
        {
            return new StructureAnalysisResult
            {
                UsedFallback = true,
                Warnings = { "Structure analysis skipped because the extracted document text was empty." }
            };
        }

        if (_client is null)
        {
            return BuildFallbackResult(document, profile, "No API client configured for structure analysis.");
        }

        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        linkedCts.CancelAfter(StructureTimeout);

        try
        {
            var prompt = BuildPrompt(document.Text, profile);
            var response = await _client.CreateResponseAsync(
                new OpenAiResponseRequest
                {
                    Model = _model,
                    Input = prompt,
                    ReasoningEffort = "low",
                    JsonSchema = _schema,
                    JsonSchemaName = "structure_analysis",
                    JsonSchemaDescription = "Top-level sections and overall document context for review chunking."
                },
                linkedCts.Token).ConfigureAwait(false);

            var parsed = ParseResponse(response.OutputText, document.Text);
            parsed.Model = _model;
            parsed.Usage = response.Usage;
            return parsed;
        }
        catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
        {
            var result = BuildFallbackResult(document, profile, $"Structure analysis fell back to local heuristics: {ex.Message}");
            result.Model = _model;
            return result;
        }
    }

    private static string BuildPrompt(string fullText, DocumentProfile profile)
    {
        var profileHint = profile == DocumentProfile.Auto
            ? string.Empty
            : Environment.NewLine + "Document profile hint: " + profile.ToCliString();

        var limitedText = Truncate(fullText, StructureMaxInputChars);
        return """
Analyze this document's structure and classify it. Return a JSON object with two keys: "sections" and "document_context".

"sections" is an array of TOP-LEVEL sections only (not subsections or sub-subsections). Aim for 4-8 major sections.
For each section, identify the heading text, classify it (introduction/methods/results/discussion/conclusion/abstract/references/body/frontmatter), provide the exact first 5-8 words of the section body, and mark whether it's a major section.
Front matter (title, authors, affiliations) should be classified as "frontmatter" and marked not major.
References/bibliography should be classified as "references" and marked not major.
Do NOT split a section like Methods or Results into sub-parts. Keep the granularity at the top heading level.

"document_context" classifies the document overall:
- document_type: one of biomedical_manuscript, clinical_protocol, regulatory_submission, grant_proposal, legal_contract, technical_report, general
- study_design: one of observational, rct, case_control, cohort, meta_analysis, case_report, or n/a
- reporting_standard: the most applicable reporting guideline (STROBE, CONSORT, PRISMA, SPIRIT, ARRIVE, CARE) or n/a
- discipline: the primary scientific/professional discipline (e.g., clinical_neuroscience, pharmacology, contract_law)
- context_note: one sentence describing what this document is about, for use as context in downstream review

Only return the JSON, nothing else.
""" + profileHint + """

Document text:
""" + limitedText;
    }

    private static StructureAnalysisResult ParseResponse(string outputText, string fullText)
    {
        var normalized = StripCodeFences(outputText);
        var parsed = JsonSerializer.Deserialize(normalized, ReviewJsonContext.Default.LlmStructureResponse)
            ?? throw new JsonException("Structure response deserialized to null.");

        var sections = MapSections(parsed.Sections, fullText);
        if (sections.Count == 0)
            throw new JsonException("Structure analysis returned no mappable sections.");

        return new StructureAnalysisResult
        {
            Sections = sections,
            DocumentContext = parsed.DocumentContext
        };
    }

    private static List<ReviewSection> MapSections(IReadOnlyList<LlmStructureSection> llmSections, string fullText)
    {
        var mapped = new List<ReviewSection>();
        foreach (var (section, index) in llmSections.Select((value, index) => (value, index)))
        {
            var start = FindPhraseOffset(fullText, section.StartPhrase, section.Heading);
            if (start < 0)
                continue;

            mapped.Add(new ReviewSection
            {
                Id = $"{ReviewTextBuilder.NormalizeSectionId(string.IsNullOrWhiteSpace(section.Heading) ? section.Key ?? "section" : section.Heading)}-{index + 1}",
                Heading = string.IsNullOrWhiteSpace(section.Heading) ? section.Key ?? "Section" : section.Heading.Trim(),
                Key = NormalizeSectionKey(section.Key),
                Start = start,
                Major = section.Major
            });
        }

        mapped.Sort(static (left, right) => left.Start.CompareTo(right.Start));
        if (mapped.Count == 0)
            return mapped;

        for (var index = 0; index < mapped.Count; index++)
        {
            var end = index == mapped.Count - 1 ? fullText.Length : mapped[index + 1].Start;
            if (end <= mapped[index].Start)
                continue;

            mapped[index].End = end;
            mapped[index].Text = fullText[mapped[index].Start..end].Trim();
            mapped[index].Chunked = false;
        }

        return mapped.Where(static section => section.End > section.Start && !string.IsNullOrWhiteSpace(section.Text)).ToList();
    }

    private static int FindPhraseOffset(string fullText, string? startPhrase, string? heading)
    {
        if (!string.IsNullOrWhiteSpace(startPhrase))
        {
            var index = fullText.IndexOf(startPhrase.Trim(), StringComparison.OrdinalIgnoreCase);
            if (index >= 0)
                return index;
        }

        if (!string.IsNullOrWhiteSpace(heading))
        {
            var index = fullText.IndexOf(heading.Trim(), StringComparison.OrdinalIgnoreCase);
            if (index >= 0)
                return index;
        }

        return -1;
    }

    private static StructureAnalysisResult BuildFallbackResult(ReviewTextDocument document, DocumentProfile profile, string warning)
    {
        var sections = BuildFallbackSections(document);
        var context = InferFallbackDocumentContext(document.Text, profile);

        return new StructureAnalysisResult
        {
            Sections = sections,
            DocumentContext = context,
            UsedFallback = true,
            Warnings = { warning }
        };
    }

    private static List<ReviewSection> BuildFallbackSections(ReviewTextDocument document)
    {
        var headings = document.Lines
            .Where(static line => line.IsHeading)
            .ToList();

        if (headings.Count == 0)
        {
            return new List<ReviewSection>
            {
                new()
                {
                    Id = "body-1",
                    Heading = "Body",
                    Key = "body",
                    Start = 0,
                    End = document.Text.Length,
                    Text = document.Text.Trim(),
                    Major = true
                }
            };
        }

        var sections = new List<ReviewSection>();

        if (headings[0].Start > 0)
        {
            var frontMatterText = document.Text[..headings[0].Start].Trim();
            if (!string.IsNullOrWhiteSpace(frontMatterText))
            {
                sections.Add(new ReviewSection
                {
                    Id = "frontmatter-1",
                    Heading = "Front Matter",
                    Key = "frontmatter",
                    Start = 0,
                    End = headings[0].Start,
                    Text = frontMatterText,
                    Major = false
                });
            }
        }

        for (var index = 0; index < headings.Count; index++)
        {
            var line = headings[index];
            var start = line.Start;
            var end = index == headings.Count - 1 ? document.Text.Length : headings[index + 1].Start;
            if (end <= start)
                continue;

            var key = NormalizeSectionKey(line.Text);
            sections.Add(new ReviewSection
            {
                Id = line.SectionId ?? $"section-{index + 1}",
                Heading = line.Text,
                Key = key,
                Style = line.Style,
                Start = start,
                End = end,
                Text = document.Text[start..end].Trim(),
                Major = key is not ("references" or "frontmatter")
            });
        }

        if (sections.Count == 0)
        {
            sections.Add(new ReviewSection
            {
                Id = "body-1",
                Heading = "Body",
                Key = "body",
                Start = 0,
                End = document.Text.Length,
                Text = document.Text.Trim(),
                Major = true
            });
        }

        return sections;
    }

    private static DocumentContext InferFallbackDocumentContext(string text, DocumentProfile profile)
    {
        var lowered = text.ToLowerInvariant();
        var context = new DocumentContext
        {
            DocumentType = "general",
            StudyDesign = "n/a",
            ReportingStandard = "n/a",
            Discipline = "general",
            ContextNote = "General document for editorial review."
        };

        if (profile == DocumentProfile.Contract ||
            ContainsAny(lowered, "agreement", "termination", "governing law", "force majeure", "party", "parties"))
        {
            context.DocumentType = "legal_contract";
            context.Discipline = "contract_law";
            context.ContextNote = "Contract or agreement text with obligation and term language.";
        }
        else if (profile == DocumentProfile.Regulatory ||
                 ContainsAny(lowered, "fda", "ema", "ich", "module 2", "regulatory submission", "compliance"))
        {
            context.DocumentType = "regulatory_submission";
            context.Discipline = "regulatory_affairs";
            context.ContextNote = "Regulatory or quality document that references guidance, modules, or compliance.";
        }
        else if (ContainsAny(lowered, "protocol", "eligibility criteria", "study visit", "randomized", "screening"))
        {
            context.DocumentType = "clinical_protocol";
            context.Discipline = "clinical";
            context.ContextNote = "Clinical protocol or trial operations document.";
        }
        else if (ContainsAny(lowered, "systematic review", "meta-analysis", "prisma"))
        {
            context.DocumentType = "general";
            context.StudyDesign = "meta_analysis";
            context.ReportingStandard = "PRISMA";
            context.Discipline = "biomedical";
            context.ContextNote = "Systematic review or meta-analysis manuscript.";
            return context;
        }
        else if (ContainsAny(lowered, "manuscript", "abstract", "methods", "results", "discussion", "adverse event", "clinical trial"))
        {
            context.DocumentType = "biomedical_manuscript";
            context.Discipline = "biomedical";
            context.ContextNote = "Biomedical or clinical manuscript.";
        }

        if (ContainsAny(lowered, "randomized", "randomised", "allocation concealment"))
        {
            context.StudyDesign = "rct";
            context.ReportingStandard = "CONSORT";
        }
        else if (ContainsAny(lowered, "cohort"))
        {
            context.StudyDesign = "cohort";
            context.ReportingStandard = "STROBE";
        }
        else if (ContainsAny(lowered, "case-control", "case control"))
        {
            context.StudyDesign = "case_control";
            context.ReportingStandard = "STROBE";
        }
        else if (ContainsAny(lowered, "case report"))
        {
            context.StudyDesign = "case_report";
            context.ReportingStandard = "CARE";
        }
        else if (context.DocumentType is "clinical_protocol")
        {
            context.StudyDesign = "rct";
            context.ReportingStandard = "SPIRIT";
        }
        else if (context.DocumentType is "biomedical_manuscript")
        {
            context.StudyDesign = "observational";
            context.ReportingStandard = "STROBE";
        }

        return context;
    }

    private static bool ContainsAny(string text, params string[] needles) =>
        needles.Any(needle => text.Contains(needle, StringComparison.Ordinal));

    private static string NormalizeSectionKey(string? keyOrHeading)
    {
        var value = (keyOrHeading ?? string.Empty).Trim().ToLowerInvariant();
        return value switch
        {
            "abstract" => "abstract",
            "introduction" => "introduction",
            "background" => "introduction",
            "methods" => "methods",
            "methodology" => "methods",
            "materials and methods" => "methods",
            "results" => "results",
            "findings" => "results",
            "discussion" => "discussion",
            "conclusion" => "conclusion",
            "conclusions" => "conclusion",
            "references" => "references",
            "bibliography" => "references",
            "frontmatter" => "frontmatter",
            "front matter" => "frontmatter",
            _ when value.Contains("abstract", StringComparison.Ordinal) => "abstract",
            _ when value.Contains("introduction", StringComparison.Ordinal) => "introduction",
            _ when value.Contains("method", StringComparison.Ordinal) => "methods",
            _ when value.Contains("result", StringComparison.Ordinal) => "results",
            _ when value.Contains("discussion", StringComparison.Ordinal) => "discussion",
            _ when value.Contains("conclusion", StringComparison.Ordinal) => "conclusion",
            _ when value.Contains("reference", StringComparison.Ordinal) || value.Contains("bibliography", StringComparison.Ordinal) => "references",
            _ => "body"
        };
    }

    private static string StripCodeFences(string text)
    {
        var normalized = (text ?? string.Empty).Trim();
        if (!normalized.StartsWith("```", StringComparison.Ordinal))
            return normalized;

        var newlineIndex = normalized.IndexOf('\n');
        if (newlineIndex < 0)
            return normalized.Trim('`').Trim();

        normalized = normalized[(newlineIndex + 1)..];
        var closingFence = normalized.LastIndexOf("```", StringComparison.Ordinal);
        if (closingFence >= 0)
            normalized = normalized[..closingFence];

        normalized = MultiLineSpacing.Replace(normalized, Environment.NewLine + Environment.NewLine);
        return normalized.Trim();
    }

    private static string Truncate(string value, int maxLength)
    {
        if (value.Length <= maxLength)
            return value;

        return value[..maxLength];
    }
}

public sealed class LlmStructureSection
{
    [JsonPropertyName("heading")]
    public string Heading { get; set; } = string.Empty;

    [JsonPropertyName("key")]
    public string? Key { get; set; }

    [JsonPropertyName("start_phrase")]
    public string? StartPhrase { get; set; }

    [JsonPropertyName("major")]
    public bool Major { get; set; }
}

public sealed class LlmStructureResponse
{
    [JsonPropertyName("sections")]
    public List<LlmStructureSection> Sections { get; set; } = new();

    [JsonPropertyName("document_context")]
    public DocumentContext? DocumentContext { get; set; }
}
