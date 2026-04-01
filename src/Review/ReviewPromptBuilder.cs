using System;

namespace DocxReview.Review;

public sealed class ReviewPromptBuilder
{
    private readonly ReviewPromptSet _promptSet;

    public ReviewPromptBuilder(ReviewPromptSet promptSet)
    {
        _promptSet = promptSet ?? throw new ArgumentNullException(nameof(promptSet));
    }

    public string BuildSectionPrompt(
        ReviewMode mode,
        DocumentProfile profile,
        string? sectionContext,
        int chunkIndex,
        string chunkText,
        string? customInstructions)
    {
        var entry = _promptSet.GetProfileEntry(profile);
        var preamble = _promptSet.GetSectionPreamble(mode);
        var task = _promptSet.GetSectionTask(mode);
        var overlay = mode == ReviewMode.Proofread ? ProofreadSafeOverlay(entry) : entry.Overlay;
        var custom = FormatCustomInstructions(customInstructions);

        var body = $$"""
        {{preamble}}
        {{entry.Prefix}}
        {{task}}
        {{(string.IsNullOrWhiteSpace(sectionContext) ? string.Empty : sectionContext.Trim())}}
        {{overlay}}
        {{custom}}
        Output schema (strict JSON, no markdown):
        {{_promptSet.SectionSchema}}
        Chunk index: {{chunkIndex}}
        Chunk text:
        {{chunkText}}
        """;

        return mode switch
        {
            ReviewMode.Proofread => WrapWithProofreadScaffold(body),
            ReviewMode.PeerReview => WrapWithPeerReviewScaffold(body),
            _ => WrapWithDeepScaffold(body)
        };
    }

    public string BuildIntegrationPrompt(
        ReviewMode mode,
        string fullDocument,
        string editSummaryJson,
        DocumentProfile profile,
        string? customInstructions)
    {
        var entry = _promptSet.GetProfileEntry(profile);
        var preamble = _promptSet.GetIntegrationPreamble(mode);
        var task = _promptSet.GetIntegrationTask(mode);
        var schema = _promptSet.GetIntegrationSchema(mode);
        var overlay = mode == ReviewMode.Proofread ? string.Empty : entry.Overlay;
        var custom = FormatCustomInstructions(customInstructions);

        return $$"""
        {{preamble}}
        {{entry.Prefix}}
        {{task}}
        {{overlay}}
        {{custom}}
        Output schema (strict JSON, no markdown):
        {{schema}}
        Document text:
        {{fullDocument}}
        Per-chunk edits already made (compact summary):
        {{editSummaryJson}}
        """;
    }

    internal static string WrapWithDeepScaffold(string prompt) =>
        """
        Follow these steps internally. Do NOT include the step labels or analysis text in your JSON output — output ONLY the final JSON manifest.

        STEP 1 — DOCUMENT ANALYSIS:
        Before making any changes, analyze the chunk:
        - What type of section is this? (methods, results, discussion, etc.)
        - What discipline does it belong to?
        - What is the overall quality level?

        STEP 2 — ISSUE IDENTIFICATION:
        Scan the chunk systematically for:
        - Language and tone issues (academic register, hedging, clarity)
        - Grammar, punctuation, and style errors
        - Content verification issues flagged by the profile instructions below
        - Claims that may be overclaiming, unsupported, or contradicted within the chunk

        """ + prompt + """

        STEP 3 — EDIT PLANNING:
        For each issue found, decide:
        - Is this a clear error that should be a tracked change? → add to "changes"
        - Is this a subjective suggestion the author should decide? → add to "comments"
        - How targeted can the edit be? (prefer short phrases over full sentences)

        STEP 4 — MANIFEST GENERATION:
        Generate the JSON manifest with all changes and comments. Output ONLY the JSON object, nothing else.
        """;

    internal static string WrapWithProofreadScaffold(string prompt) =>
        """
        Follow these steps internally. Do NOT include the step labels or analysis text in your JSON output — output ONLY the final JSON manifest.

        STEP 1 — SCAN:
        Read the chunk once through, marking every mechanical issue you find: misspellings, grammar errors, punctuation problems, inconsistent style choices, wrong words, hyphenation issues, number-style drift, abbreviation problems.

        """ + prompt + """

        STEP 2 — GENERATE:
        For each issue found, create the most targeted fix possible (1-3 word "find" spans). Output ONLY the JSON object, nothing else.
        """;

    internal static string WrapWithPeerReviewScaffold(string prompt) =>
        """
        Follow these steps internally. Do NOT include the step labels or analysis text in your JSON output — output ONLY the final JSON manifest.

        STEP 1 — SECTION CLASSIFICATION:
        What type of section is this? (introduction, methods, results, discussion, abstract, etc.)
        What claims does it make? What evidence does it present?

        STEP 2 — METHODOLOGY ASSESSMENT:
        Evaluate the study design, controls, sampling, blinding, and reproducibility described in this section. Are the methods appropriate for the research question? Are there gaps?

        STEP 3 — EVIDENCE EVALUATION:
        For each claim in this section:
        - What evidence supports it?
        - Is the claim proportional to the evidence? (correlation vs causation, sample size limitations, effect size)
        - Are the statistical methods appropriate?
        - Are there alternative explanations the authors haven't addressed?

        STEP 4 — ISSUE CATEGORIZATION:
        For each issue found, decide:
        - Is this a clear factual error (wrong number, broken reference, typo)? → add to "changes" (rare)
        - Is this a methodological concern, unsupported claim, missing information, or suggestion? → add to "comments"
        - How severe is this issue? Would it affect the paper's conclusions?

        """ + prompt + """

        STEP 5 — MANIFEST GENERATION:
        Generate the JSON manifest with all changes and comments. Output ONLY the JSON object, nothing else.
        """;

    internal static string ProofreadSafeOverlay(ReviewProfilePromptEntry entry)
    {
        if (string.IsNullOrWhiteSpace(entry.Overlay))
            return string.Empty;

        var split = entry.Overlay.Split(
            new[] { Environment.NewLine + Environment.NewLine, "\n\n" },
            2,
            StringSplitOptions.None);
        var firstParagraph = split[0].Trim();
        return string.IsNullOrWhiteSpace(firstParagraph)
            ? string.Empty
            : "Copy editing context: " + firstParagraph;
    }

    internal static string FormatCustomInstructions(string? instructions) =>
        string.IsNullOrWhiteSpace(instructions) ? string.Empty : "Additional instructions: " + instructions.Trim();
}
