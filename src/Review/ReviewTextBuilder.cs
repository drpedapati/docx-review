using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace DocxReview.Review;

public static class ReviewTextBuilder
{
    private static readonly Regex MultiWhitespace = new(@"[ \t\r\f\v\u00A0]+", RegexOptions.Compiled);
    private static readonly Regex SlugSanitizer = new(@"[^a-z0-9]+", RegexOptions.Compiled);

    public static ReviewTextDocument Build(DocumentExtraction extraction)
    {
        ArgumentNullException.ThrowIfNull(extraction);

        var items = new List<(int flowIndex, string text, string? style, bool isHeading)>();
        var tableIndex = 0;

        foreach (var paragraph in extraction.Paragraphs.OrderBy(static paragraph => paragraph.Index))
        {
            while (tableIndex < extraction.Tables.Count &&
                   extraction.Tables[tableIndex].ParagraphIndex <= paragraph.Index)
            {
                var tableText = BuildTableText(extraction.Tables[tableIndex]);
                if (!string.IsNullOrWhiteSpace(tableText))
                {
                    items.Add((extraction.Tables[tableIndex].ParagraphIndex, tableText, "Table", false));
                }

                tableIndex++;
            }

            var text = NormalizeInlineText(paragraph.Text);
            if (string.IsNullOrWhiteSpace(text))
                continue;

            items.Add((paragraph.Index, text, paragraph.Style, IsLikelySectionHeading(text, paragraph.Style)));
        }

        while (tableIndex < extraction.Tables.Count)
        {
            var tableText = BuildTableText(extraction.Tables[tableIndex]);
            if (!string.IsNullOrWhiteSpace(tableText))
            {
                items.Add((extraction.Tables[tableIndex].ParagraphIndex, tableText, "Table", false));
            }

            tableIndex++;
        }

        if (extraction.Footnotes.Count > 0)
        {
            items.Add((int.MaxValue - 1, "Footnotes", "Heading1", true));
            foreach (var footnote in extraction.Footnotes)
            {
                var text = NormalizeInlineText($"[{footnote.Id}] {footnote.Text}");
                if (!string.IsNullOrWhiteSpace(text))
                    items.Add((int.MaxValue - 1, text, "Footnote", false));
            }
        }

        if (extraction.Endnotes.Count > 0)
        {
            items.Add((int.MaxValue, "Endnotes", "Heading1", true));
            foreach (var endnote in extraction.Endnotes)
            {
                var text = NormalizeInlineText($"[{endnote.Id}] {endnote.Text}");
                if (!string.IsNullOrWhiteSpace(text))
                    items.Add((int.MaxValue, text, "Endnote", false));
            }
        }

        items.Sort(static (left, right) => left.flowIndex.CompareTo(right.flowIndex));

        var lines = new List<ReviewLine>(items.Count);
        var builder = new StringBuilder();
        var paragraphIndex = 0;

        foreach (var item in items)
        {
            if (builder.Length > 0)
                builder.AppendLine().AppendLine();

            var start = builder.Length;
            builder.Append(item.text);
            var end = builder.Length;

            lines.Add(new ReviewLine
            {
                ParagraphIndex = paragraphIndex,
                Start = start,
                End = end,
                Text = item.text,
                Style = item.style,
                SectionId = item.isHeading ? NormalizeSectionId(item.text) : null,
                IsHeading = item.isHeading
            });

            paragraphIndex++;
        }

        return new ReviewTextDocument
        {
            Text = builder.ToString(),
            Lines = lines
        };
    }

    public static string NormalizeInlineText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        var normalized = text
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(static line => MultiWhitespace.Replace(line, " ").Trim())
            .Where(static line => !string.IsNullOrWhiteSpace(line));

        return string.Join(" ", normalized);
    }

    public static bool IsLikelySectionHeading(string text, string? style)
    {
        if (!string.IsNullOrWhiteSpace(style))
        {
            var styleLower = style.Trim().ToLowerInvariant();
            if (styleLower.Contains("heading", StringComparison.Ordinal) ||
                styleLower.Contains("title", StringComparison.Ordinal) ||
                styleLower.Contains("subtitle", StringComparison.Ordinal))
            {
                return true;
            }
        }

        if (string.IsNullOrWhiteSpace(text))
            return false;

        if (text.Length > 120 || text.EndsWith(".", StringComparison.Ordinal) || text.EndsWith(":", StringComparison.Ordinal))
            return false;

        if (char.IsDigit(text[0]))
            return true;

        var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length is < 1 or > 12)
            return false;

        var titleCaseWords = words.Count(static word => word.Length > 0 && char.IsUpper(word[0]));
        return titleCaseWords >= Math.Max(1, words.Length - 2);
    }

    public static string NormalizeSectionId(string heading)
    {
        var lowered = heading.Trim().ToLowerInvariant();
        lowered = SlugSanitizer.Replace(lowered, "-").Trim('-');
        return string.IsNullOrWhiteSpace(lowered) ? "section" : lowered;
    }

    private static string BuildTableText(TableInfo table)
    {
        if (table.Cells.Count == 0)
            return string.Empty;

        var lines = table.Cells
            .Select(static row => "| " + string.Join(" | ", row.Select(static cell => NormalizeInlineText(cell))) + " |")
            .Where(static row => !string.IsNullOrWhiteSpace(row));

        return string.Join(" ", lines);
    }
}
