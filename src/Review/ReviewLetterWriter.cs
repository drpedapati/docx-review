using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DocxReview.Review;

public sealed class ReviewLetterContext
{
    public string InputPath { get; init; } = string.Empty;
    public string? OutputPath { get; init; }
    public ReviewMode ReviewMode { get; init; }
    public DocumentProfile Profile { get; init; } = DocumentProfile.Auto;
    public DocumentContext? DocumentContext { get; init; }
    public string Status { get; init; } = ReviewStatuses.Completed;
    public bool Degraded { get; init; }
    public int ChangesAttempted { get; init; }
    public int ChangesSucceeded { get; init; }
    public int CommentsAttempted { get; init; }
    public int CommentsSucceeded { get; init; }
    public int ChangesFailed { get; init; }
    public int FallbackComments { get; init; }
    public TokenUsage TokenUsage { get; init; } = new();
    public IReadOnlyList<PassSummary>? Passes { get; init; }
}

public static class ReviewLetterWriter
{
    public static void WriteMarkdown(string path, ReviewLetter letter, ReviewLetterContext context)
    {
        ArgumentNullException.ThrowIfNull(letter);
        ArgumentNullException.ThrowIfNull(context);

        EnsureParentDirectory(path);
        File.WriteAllText(path, BuildMarkdown(letter, context));
    }

    public static string BuildMarkdown(ReviewLetter letter, ReviewLetterContext context)
    {
        ArgumentNullException.ThrowIfNull(letter);
        ArgumentNullException.ThrowIfNull(context);

        var builder = new StringBuilder();
        builder.AppendLine(GetTitle(context.ReviewMode));
        builder.AppendLine();

        WriteMetadata(builder, context);

        switch (context.ReviewMode)
        {
            case ReviewMode.Proofread:
                WriteProofreadLetter(builder, letter);
                break;

            case ReviewMode.PeerReview:
                WritePeerReviewLetter(builder, letter);
                break;

            default:
                WriteSubstantiveLetter(builder, letter);
                break;
        }

        return builder.ToString().TrimEnd() + Environment.NewLine;
    }

    private static string GetTitle(ReviewMode mode) => mode switch
    {
        ReviewMode.Proofread => "# Copy Editor's Summary",
        ReviewMode.PeerReview => "# Peer Review Letter",
        _ => "# Editorial Review Letter"
    };

    private static void WriteMetadata(StringBuilder builder, ReviewLetterContext context)
    {
        builder.AppendLine("## Review Metadata");
        builder.AppendLine();
        builder.AppendLine($"- Input: {context.InputPath}");
        if (!string.IsNullOrWhiteSpace(context.OutputPath))
            builder.AppendLine($"- Output: {context.OutputPath}");
        builder.AppendLine($"- Mode: {context.ReviewMode.ToCliString()}");
        builder.AppendLine($"- Profile: {context.Profile.ToCliString()}");
        builder.AppendLine($"- Status: {context.Status}{(context.Degraded ? " (degraded)" : string.Empty)}");

        if (context.DocumentContext is not null)
        {
            builder.AppendLine($"- Document type: {context.DocumentContext.DocumentType}");
            builder.AppendLine($"- Study design: {context.DocumentContext.StudyDesign}");
            builder.AppendLine($"- Reporting standard: {context.DocumentContext.ReportingStandard}");
            if (!string.IsNullOrWhiteSpace(context.DocumentContext.Discipline))
                builder.AppendLine($"- Discipline: {context.DocumentContext.Discipline}");
        }

        builder.AppendLine($"- Changes applied: {context.ChangesSucceeded}/{context.ChangesAttempted}");
        builder.AppendLine($"- Comments applied: {context.CommentsSucceeded}/{context.CommentsAttempted}");
        if (context.ChangesFailed > 0)
            builder.AppendLine($"- Changes converted to fallback comments: {context.ChangesFailed}");
        if (context.FallbackComments > 0)
            builder.AppendLine($"- Fallback comments added: {context.FallbackComments}");
        builder.AppendLine($"- Token usage: {context.TokenUsage.InputTokens} input, {context.TokenUsage.OutputTokens} output, {context.TokenUsage.TotalTokens} total");

        if (context.Passes is { Count: > 0 })
            builder.AppendLine($"- Passes: {string.Join(", ", context.Passes.Select(static pass => $"{pass.Name}={pass.Status}"))}");

        builder.AppendLine();
    }

    private static void WriteSubstantiveLetter(StringBuilder builder, ReviewLetter letter)
    {
        WriteParagraphSection(builder, "Overall Assessment", letter.OverallAssessment);
        WriteFindings(builder, letter.KeyFindings);
        WriteListSection(builder, "Highlights", letter.Highlights);
        WriteListSection(builder, "Concerns", letter.Concerns);
        WriteListSection(builder, "Recommendations", letter.Recommendations);
    }

    private static void WriteProofreadLetter(StringBuilder builder, ReviewLetter letter)
    {
        WriteParagraphSection(builder, "Overall Assessment", letter.OverallAssessment);
        WritePatterns(builder, letter.Patterns);
        WriteFindings(builder, letter.KeyFindings);
        WriteListSection(builder, "Highlights", letter.Highlights);
        WriteListSection(builder, "Concerns", letter.Concerns);
        WriteListSection(builder, "Recommendations", letter.Recommendations);
    }

    private static void WritePeerReviewLetter(StringBuilder builder, ReviewLetter letter)
    {
        WriteParagraphSection(builder, "Overall Assessment", letter.OverallAssessment);

        if (!string.IsNullOrWhiteSpace(letter.Recommendation))
        {
            builder.AppendLine("## Recommendation");
            builder.AppendLine();
            builder.AppendLine($"**{letter.Recommendation.Trim()}**");
            builder.AppendLine();
        }

        WriteFindings(builder, letter.KeyFindings);
        WritePeerComments(builder, "Major Comments", letter.MajorComments);
        WritePeerComments(builder, "Minor Comments", letter.MinorComments);
        WriteListSection(builder, "Questions for Authors", letter.QuestionsForAuthors);
        WriteListSection(builder, "Highlights", letter.Highlights);
        WriteListSection(builder, "Concerns", letter.Concerns);
        WriteListSection(builder, "Recommendations", letter.Recommendations);
    }

    private static void WriteParagraphSection(StringBuilder builder, string heading, string? body)
    {
        if (string.IsNullOrWhiteSpace(body))
            return;

        builder.AppendLine($"## {heading}");
        builder.AppendLine();
        builder.AppendLine(body.Trim());
        builder.AppendLine();
    }

    private static void WriteFindings(StringBuilder builder, IReadOnlyList<LetterFinding>? findings)
    {
        if (findings is null || findings.Count == 0)
            return;

        builder.AppendLine("## Key Findings");
        builder.AppendLine();
        foreach (var finding in findings)
        {
            var severity = string.IsNullOrWhiteSpace(finding.Severity) ? string.Empty : $" ({finding.Severity.Trim()})";
            builder.AppendLine($"### {finding.Category.Trim()}{severity}");
            builder.AppendLine();
            builder.AppendLine(finding.Description.Trim());
            builder.AppendLine();
        }
    }

    private static void WritePatterns(StringBuilder builder, IReadOnlyList<LetterPattern>? patterns)
    {
        if (patterns is null || patterns.Count == 0)
            return;

        builder.AppendLine("## Recurring Patterns");
        builder.AppendLine();
        foreach (var pattern in patterns.Where(static item => !string.IsNullOrWhiteSpace(item.Pattern)))
        {
            var category = string.IsNullOrWhiteSpace(pattern.Category) ? string.Empty : $" ({pattern.Category.Trim()})";
            builder.AppendLine($"- **{pattern.Pattern.Trim()}**{category}");
            if (pattern.Examples is { Count: > 0 })
            {
                foreach (var example in pattern.Examples.Where(static item => !string.IsNullOrWhiteSpace(item)))
                    builder.AppendLine($"  - {example.Trim()}");
            }
        }
        builder.AppendLine();
    }

    private static void WritePeerComments(StringBuilder builder, string heading, IReadOnlyList<PeerComment>? comments)
    {
        if (comments is null || comments.Count == 0)
            return;

        builder.AppendLine($"## {heading}");
        builder.AppendLine();
        foreach (var comment in comments.Where(static item => !string.IsNullOrWhiteSpace(item.Description)))
        {
            var prefix = comment.Number > 0 ? $"{comment.Number}. " : "- ";
            var severity = string.IsNullOrWhiteSpace(comment.Severity) ? string.Empty : $" ({comment.Severity.Trim()})";
            var section = string.IsNullOrWhiteSpace(comment.Section) ? string.Empty : $" [{comment.Section.Trim()}]";
            builder.AppendLine($"{prefix}{comment.Description.Trim()}{section}{severity}");
        }
        builder.AppendLine();
    }

    private static void WriteListSection(StringBuilder builder, string heading, IReadOnlyList<string>? items)
    {
        if (items is null || items.Count == 0)
            return;

        builder.AppendLine($"## {heading}");
        builder.AppendLine();
        foreach (var item in items.Where(static value => !string.IsNullOrWhiteSpace(value)))
            builder.AppendLine($"- {item.Trim()}");
        builder.AppendLine();
    }

    private static void EnsureParentDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);
    }
}
