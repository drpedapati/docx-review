using System;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DocxReview.Review;

[JsonConverter(typeof(ReviewModeJsonConverter))]
public enum ReviewMode
{
    Substantive,

    Proofread,

    PeerReview
}

[JsonConverter(typeof(DocumentProfileJsonConverter))]
public enum DocumentProfile
{
    Auto,

    General,

    Medical,

    Regulatory,

    Reference,

    Legal,

    Contract
}

public static class ReviewPassNames
{
    public const string StructureAnalysis = "structure_analysis";
    public const string SectionReview = "section_review";
    public const string Integration = "integration";
}

public static class ReviewStatuses
{
    public const string Completed = "completed";
    public const string Degraded = "degraded";
    public const string Failed = "failed";
    public const string Skipped = "skipped";
}

public static class ReviewModeExtensions
{
    public static string ToCliString(this ReviewMode mode) => mode switch
    {
        ReviewMode.Substantive => "substantive",
        ReviewMode.Proofread => "proofread",
        ReviewMode.PeerReview => "peer_review",
        _ => "substantive"
    };

    public static bool TryParse(string? value, out ReviewMode mode)
    {
        switch ((value ?? string.Empty).Trim().ToLowerInvariant())
        {
            case "substantive":
                mode = ReviewMode.Substantive;
                return true;
            case "proofread":
                mode = ReviewMode.Proofread;
                return true;
            case "peer_review":
                mode = ReviewMode.PeerReview;
                return true;
            default:
                mode = ReviewMode.Substantive;
                return false;
        }
    }
}

public sealed class ReviewModeJsonConverter : JsonConverter<ReviewMode>
{
    public override ReviewMode Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType != JsonTokenType.String)
            throw new JsonException("Review mode must be a string.");

        var value = reader.GetString();
        if (!ReviewModeExtensions.TryParse(value, out var mode))
            throw new JsonException($"Invalid review mode '{value}'.");

        return mode;
    }

    public override void Write(Utf8JsonWriter writer, ReviewMode value, JsonSerializerOptions options) =>
        writer.WriteStringValue(value.ToCliString());
}

public static class DocumentProfileExtensions
{
    public static string ToCliString(this DocumentProfile profile) => profile switch
    {
        DocumentProfile.Auto => "auto",
        DocumentProfile.General => "general",
        DocumentProfile.Medical => "medical",
        DocumentProfile.Regulatory => "regulatory",
        DocumentProfile.Reference => "reference",
        DocumentProfile.Legal => "legal",
        DocumentProfile.Contract => "contract",
        _ => "general"
    };

    public static bool TryParse(string? value, out DocumentProfile profile)
    {
        switch ((value ?? string.Empty).Trim().ToLowerInvariant())
        {
            case "":
            case "auto":
                profile = DocumentProfile.Auto;
                return true;
            case "general":
                profile = DocumentProfile.General;
                return true;
            case "medical":
                profile = DocumentProfile.Medical;
                return true;
            case "regulatory":
                profile = DocumentProfile.Regulatory;
                return true;
            case "reference":
                profile = DocumentProfile.Reference;
                return true;
            case "legal":
                profile = DocumentProfile.Legal;
                return true;
            case "contract":
                profile = DocumentProfile.Contract;
                return true;
            default:
                profile = DocumentProfile.Auto;
                return false;
        }
    }
}

public sealed class DocumentProfileJsonConverter : JsonConverter<DocumentProfile>
{
    public override DocumentProfile Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType != JsonTokenType.String)
            throw new JsonException("Document profile must be a string.");

        var value = reader.GetString();
        if (!DocumentProfileExtensions.TryParse(value, out var profile))
            throw new JsonException($"Invalid document profile '{value}'.");

        return profile;
    }

    public override void Write(Utf8JsonWriter writer, DocumentProfile value, JsonSerializerOptions options) =>
        writer.WriteStringValue(value.ToCliString());
}

public sealed class ReviewSuggestion
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("type")]
    public string Type { get; set; } = "edit";

    [JsonPropertyName("severity")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Severity { get; set; }

    [JsonPropertyName("original")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Original { get; set; }

    [JsonPropertyName("revised")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Revised { get; set; }

    [JsonPropertyName("rationale")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Rationale { get; set; }

    [JsonPropertyName("anchor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Anchor { get; set; }

    [JsonPropertyName("pass")]
    public string Pass { get; set; } = string.Empty;

    [JsonPropertyName("chunk")]
    public int Chunk { get; set; }

    [JsonPropertyName("metadata")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public Dictionary<string, string>? Metadata { get; set; }
}

public sealed class ReviewLine
{
    [JsonPropertyName("paragraph_index")]
    public int ParagraphIndex { get; set; }

    [JsonPropertyName("start")]
    public int Start { get; set; }

    [JsonPropertyName("end")]
    public int End { get; set; }

    [JsonPropertyName("text")]
    public string Text { get; set; } = string.Empty;

    [JsonPropertyName("style")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Style { get; set; }

    [JsonPropertyName("section_id")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SectionId { get; set; }

    [JsonPropertyName("is_heading")]
    public bool IsHeading { get; set; }
}

public sealed class ReviewTextDocument
{
    [JsonPropertyName("text")]
    public string Text { get; set; } = string.Empty;

    [JsonPropertyName("lines")]
    public List<ReviewLine> Lines { get; set; } = new();
}

public sealed class ReviewChunk
{
    [JsonPropertyName("index")]
    public int Index { get; set; }

    [JsonPropertyName("start")]
    public int Start { get; set; }

    [JsonPropertyName("end")]
    public int End { get; set; }

    [JsonPropertyName("text")]
    public string Text { get; set; } = string.Empty;

    [JsonPropertyName("section_id")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SectionId { get; set; }

    [JsonPropertyName("section_key")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SectionKey { get; set; }

    [JsonPropertyName("section_name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SectionName { get; set; }

    [JsonPropertyName("section_style")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SectionStyle { get; set; }
}

public sealed class ReviewSection
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("key")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Key { get; set; }

    [JsonPropertyName("heading")]
    public string Heading { get; set; } = string.Empty;

    [JsonPropertyName("style")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Style { get; set; }

    [JsonPropertyName("start")]
    public int Start { get; set; }

    [JsonPropertyName("end")]
    public int End { get; set; }

    [JsonPropertyName("text")]
    public string Text { get; set; } = string.Empty;

    [JsonPropertyName("major")]
    public bool Major { get; set; }

    [JsonPropertyName("chunked")]
    public bool Chunked { get; set; }
}

public sealed class DocumentContext
{
    [JsonPropertyName("document_type")]
    public string DocumentType { get; set; } = "general";

    [JsonPropertyName("study_design")]
    public string StudyDesign { get; set; } = "n/a";

    [JsonPropertyName("reporting_standard")]
    public string ReportingStandard { get; set; } = "n/a";

    [JsonPropertyName("discipline")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Discipline { get; set; }

    [JsonPropertyName("context_note")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ContextNote { get; set; }

    public DocumentProfile InferProfile() => DocumentType.Trim().ToLowerInvariant() switch
    {
        "biomedical_manuscript" => DocumentProfile.Medical,
        "clinical_protocol" => DocumentProfile.Medical,
        "case_report" => DocumentProfile.Medical,
        "regulatory_submission" => DocumentProfile.Regulatory,
        "sop" => DocumentProfile.Regulatory,
        "systematic_review" => DocumentProfile.Reference,
        "legal_contract" => DocumentProfile.Contract,
        "grant_proposal" => DocumentProfile.General,
        "technical_report" => DocumentProfile.General,
        "thesis" => DocumentProfile.General,
        _ => DocumentProfile.General
    };
}

public sealed class PassSummary
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("status")]
    public string Status { get; set; } = ReviewStatuses.Completed;

    [JsonPropertyName("model")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Model { get; set; }

    [JsonPropertyName("calls")]
    public int Calls { get; set; }

    [JsonPropertyName("retries")]
    public int Retries { get; set; }

    [JsonPropertyName("chunks_total")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int ChunksTotal { get; set; }

    [JsonPropertyName("chunks_succeeded")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int ChunksSucceeded { get; set; }

    [JsonPropertyName("duration_ms")]
    public long DurationMs { get; set; }

    [JsonPropertyName("usage")]
    public TokenUsage Usage { get; set; } = new();

    [JsonPropertyName("errors")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? Errors { get; set; }
}

public sealed class TokenUsage
{
    [JsonPropertyName("input_tokens")]
    public int InputTokens { get; set; }

    [JsonPropertyName("output_tokens")]
    public int OutputTokens { get; set; }

    [JsonPropertyName("total_tokens")]
    public int TotalTokens { get; set; }

    public static TokenUsage operator +(TokenUsage left, TokenUsage right) => new()
    {
        InputTokens = left.InputTokens + right.InputTokens,
        OutputTokens = left.OutputTokens + right.OutputTokens,
        TotalTokens = left.TotalTokens + right.TotalTokens
    };
}

public sealed class ReviewLetter
{
    [JsonPropertyName("overall_assessment")]
    public string OverallAssessment { get; set; } = string.Empty;

    [JsonPropertyName("key_findings")]
    public List<LetterFinding> KeyFindings { get; set; } = new();

    [JsonPropertyName("highlights")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? Highlights { get; set; }

    [JsonPropertyName("concerns")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? Concerns { get; set; }

    [JsonPropertyName("recommendations")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? Recommendations { get; set; }

    [JsonPropertyName("patterns")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<LetterPattern>? Patterns { get; set; }

    [JsonPropertyName("major_comments")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PeerComment>? MajorComments { get; set; }

    [JsonPropertyName("minor_comments")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PeerComment>? MinorComments { get; set; }

    [JsonPropertyName("questions_for_authors")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? QuestionsForAuthors { get; set; }

    [JsonPropertyName("recommendation")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Recommendation { get; set; }
}

public sealed class LetterFinding
{
    [JsonPropertyName("category")]
    public string Category { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    public string Description { get; set; } = string.Empty;

    [JsonPropertyName("severity")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Severity { get; set; }
}

public sealed class LetterPattern
{
    [JsonPropertyName("pattern")]
    public string Pattern { get; set; } = string.Empty;

    [JsonPropertyName("examples")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? Examples { get; set; }

    [JsonPropertyName("category")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Category { get; set; }
}

public sealed class PeerComment
{
    [JsonPropertyName("number")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int Number { get; set; }

    [JsonPropertyName("section")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Section { get; set; }

    [JsonPropertyName("description")]
    public string Description { get; set; } = string.Empty;

    [JsonPropertyName("severity")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Severity { get; set; }

    [JsonPropertyName("comment")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LegacyComment { get; set; }

    [JsonPropertyName("section_ref")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LegacySectionRef { get; set; }

    public void Normalize()
    {
        if (string.IsNullOrWhiteSpace(Description) && !string.IsNullOrWhiteSpace(LegacyComment))
            Description = LegacyComment;

        if (string.IsNullOrWhiteSpace(Section) && !string.IsNullOrWhiteSpace(LegacySectionRef))
            Section = LegacySectionRef;

        LegacyComment = null;
        LegacySectionRef = null;
    }
}

public sealed class ReviewRunResult
{
    [JsonPropertyName("status")]
    public string Status { get; set; } = ReviewStatuses.Completed;

    [JsonPropertyName("degraded")]
    public bool Degraded { get; set; }

    [JsonPropertyName("input")]
    public string Input { get; set; } = string.Empty;

    [JsonPropertyName("output")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Output { get; set; }

    [JsonPropertyName("review_mode")]
    public ReviewMode ReviewMode { get; set; }

    [JsonPropertyName("profile")]
    public DocumentProfile Profile { get; set; } = DocumentProfile.Auto;

    [JsonPropertyName("model")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Model { get; set; }

    [JsonPropertyName("structure_model")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StructureModel { get; set; }

    [JsonPropertyName("document_context")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DocumentContext? DocumentContext { get; set; }

    [JsonPropertyName("changes_attempted")]
    public int ChangesAttempted { get; set; }

    [JsonPropertyName("changes_succeeded")]
    public int ChangesSucceeded { get; set; }

    [JsonPropertyName("comments_attempted")]
    public int CommentsAttempted { get; set; }

    [JsonPropertyName("comments_succeeded")]
    public int CommentsSucceeded { get; set; }

    [JsonPropertyName("changes_failed")]
    public int ChangesFailed { get; set; }

    [JsonPropertyName("fallback_comments")]
    public int FallbackComments { get; set; }

    [JsonPropertyName("review_letter_path")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ReviewLetterPath { get; set; }

    [JsonPropertyName("manifest_path")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ManifestPath { get; set; }

    [JsonPropertyName("summary_path")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SummaryPath { get; set; }

    [JsonPropertyName("token_usage")]
    public TokenUsage TokenUsage { get; set; } = new();

    [JsonPropertyName("passes")]
    public List<PassSummary> Passes { get; set; } = new();

    [JsonPropertyName("durations_ms")]
    public Dictionary<string, long> DurationsMs { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    [JsonPropertyName("warnings")]
    public List<string> Warnings { get; set; } = new();

    [JsonPropertyName("errors")]
    public List<string> Errors { get; set; } = new();

    [JsonPropertyName("success")]
    public bool Success { get; set; }
}

public sealed class StructureAnalysisResult
{
    [JsonPropertyName("sections")]
    public List<ReviewSection> Sections { get; set; } = new();

    [JsonPropertyName("document_context")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DocumentContext? DocumentContext { get; set; }

    [JsonPropertyName("usage")]
    public TokenUsage Usage { get; set; } = new();

    [JsonPropertyName("model")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Model { get; set; }

    [JsonPropertyName("used_fallback")]
    public bool UsedFallback { get; set; }

    [JsonPropertyName("warnings")]
    public List<string> Warnings { get; set; } = new();
}
