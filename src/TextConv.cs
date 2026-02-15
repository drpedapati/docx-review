using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocxReview;

/// <summary>
/// Produces a normalized text representation of a .docx file suitable for
/// use as a git textconv driver. This allows `git diff` to show meaningful
/// changes for Word documents.
///
/// Output format captures all document layers:
/// - Body text with inline formatting markers [B]/[I]/[U]
/// - Tracked changes as [-deleted-] / [+inserted+]
/// - Comments as /* [Author] text */ inline
/// - Tables as pipe-delimited rows
/// - Images as [IMG: filename (hash)]
/// - Headers/footers
/// - Metadata
/// </summary>
public static class TextConv
{
    public static string Convert(DocumentExtraction doc)
    {
        var sb = new StringBuilder();

        // ── Metadata ───────────────────────────────────────
        sb.AppendLine("=== METADATA ===");
        if (doc.Metadata.Title != null)
            sb.AppendLine($"Title: {doc.Metadata.Title}");
        if (doc.Metadata.Author != null)
            sb.AppendLine($"Author: {doc.Metadata.Author}");
        if (doc.Metadata.LastModifiedBy != null)
            sb.AppendLine($"LastModifiedBy: {doc.Metadata.LastModifiedBy}");
        if (doc.Metadata.Modified != null)
            sb.AppendLine($"Modified: {doc.Metadata.Modified}");
        if (doc.Metadata.Revision != null)
            sb.AppendLine($"Revision: {doc.Metadata.Revision}");
        sb.AppendLine($"Words: {doc.Metadata.WordCount}");
        sb.AppendLine($"Paragraphs: {doc.Metadata.ParagraphCount}");
        sb.AppendLine();

        // ── Body ───────────────────────────────────────────
        sb.AppendLine("=== BODY ===");

        // Build comment index: paragraphIndex → list of comments
        var commentsByPara = doc.Comments
            .GroupBy(c => c.ParagraphIndex)
            .ToDictionary(g => g.Key, g => g.ToList());

        int tableIdx = 0;
        foreach (var para in doc.Paragraphs)
        {
            // Insert any tables that appear before this paragraph
            while (tableIdx < doc.Tables.Count && doc.Tables[tableIdx].ParagraphIndex <= para.Index)
            {
                AppendTable(sb, doc.Tables[tableIdx]);
                tableIdx++;
            }

            string styleTag = para.Style != null ? $" [{para.Style}]" : "";

            // Build paragraph text with inline formatting and tracked changes
            var text = BuildRichText(para);

            // Append inline comments
            if (commentsByPara.TryGetValue(para.Index, out var comments))
            {
                foreach (var c in comments)
                    text += $" /* [{c.Author}] {c.Text} */";
            }

            sb.AppendLine($"¶{para.Index}{styleTag} {text}");
        }

        // Remaining tables
        while (tableIdx < doc.Tables.Count)
        {
            AppendTable(sb, doc.Tables[tableIdx]);
            tableIdx++;
        }

        sb.AppendLine();

        // ── Tables summary ─────────────────────────────────
        if (doc.Tables.Count > 0)
        {
            sb.AppendLine("=== TABLES ===");
            foreach (var table in doc.Tables)
            {
                sb.AppendLine($"Table {table.Index} ({table.Rows}×{table.Columns}) at ¶{table.ParagraphIndex}:");
                AppendTable(sb, table);
            }
            sb.AppendLine();
        }

        // ── Comments ───────────────────────────────────────
        if (doc.Comments.Count > 0)
        {
            sb.AppendLine("=== COMMENTS ===");
            foreach (var c in doc.Comments)
            {
                sb.AppendLine($"#{c.Id} [{c.Author}] on \"{Trunc(c.AnchorText, 60)}\" (¶{c.ParagraphIndex}): {c.Text}");
            }
            sb.AppendLine();
        }

        // ── Images ─────────────────────────────────────────
        if (doc.Images.Count > 0)
        {
            sb.AppendLine("=== IMAGES ===");
            foreach (var img in doc.Images)
            {
                sb.AppendLine($"[IMG] {img.FileName} ({img.ContentType}, {img.SizeBytes} bytes, sha256:{img.Sha256[..12]}...)");
            }
            sb.AppendLine();
        }

        // ── Headers/Footers ────────────────────────────────
        if (doc.HeadersFooters.Count > 0)
        {
            sb.AppendLine("=== HEADERS/FOOTERS ===");
            foreach (var hf in doc.HeadersFooters)
            {
                sb.AppendLine($"[{hf.Type}/{hf.Scope}] {hf.Text}");
            }
            sb.AppendLine();
        }

        return sb.ToString();
    }

    private static string BuildRichText(RichParagraphInfo para)
    {
        var sb = new StringBuilder();

        // First, emit tracked deletions
        foreach (var tc in para.TrackedChanges.Where(t => t.Type == "delete"))
        {
            sb.Append($"[-{tc.Text}-]");
        }

        // Then emit runs with formatting markers
        foreach (var run in para.Runs)
        {
            string text = run.Text;

            // Check if this text is part of a tracked insertion
            bool isInserted = para.TrackedChanges.Any(
                tc => tc.Type == "insert" && tc.Text.Contains(text));

            if (isInserted)
                text = $"[+{text}+]";

            if (run.Bold) text = $"[B]{text}[/B]";
            if (run.Italic) text = $"[I]{text}[/I]";
            if (run.Underline) text = $"[U]{text}[/U]";
            if (run.Strikethrough) text = $"[S]{text}[/S]";

            sb.Append(text);
        }

        // If no runs but there's plain text (e.g., from tracked insertions)
        if (para.Runs.Count == 0 && !string.IsNullOrEmpty(para.Text))
            sb.Append(para.Text);

        return sb.ToString();
    }

    private static void AppendTable(StringBuilder sb, TableInfo table)
    {
        foreach (var row in table.Cells)
        {
            sb.Append("| ");
            sb.Append(string.Join(" | ", row.Select(c => c.PadRight(1))));
            sb.AppendLine(" |");
        }
    }

    private static string Trunc(string s, int max) =>
        s.Length <= max ? s : s[..max] + "…";
}
