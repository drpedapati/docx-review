using System;
using System.Collections.Generic;
using System.Linq;

namespace DocxReview;

/// <summary>
/// Compares two DocumentExtractions and produces a semantic DiffResult
/// covering text, comments, tracked changes, formatting, styles, and metadata.
/// </summary>
public static class DocumentDiffer
{
    public static DiffResult Diff(DocumentExtraction oldDoc, DocumentExtraction newDoc)
    {
        var result = new DiffResult
        {
            OldFile = oldDoc.FileName,
            NewFile = newDoc.FileName
        };

        result.Metadata = DiffMetadata(oldDoc.Metadata, newDoc.Metadata);
        result.Paragraphs = DiffParagraphs(oldDoc.Paragraphs, newDoc.Paragraphs);
        result.Comments = DiffComments(oldDoc.Comments, newDoc.Comments);
        result.TrackedChanges = DiffTrackedChanges(oldDoc.Paragraphs, newDoc.Paragraphs);

        // Build summary
        int styleChanges = result.Paragraphs.Modified.Count(m => m.StyleChange != null);
        int formattingChanges = result.Paragraphs.Modified.Sum(m => m.FormattingChanges.Count);

        result.Summary = new DiffSummary
        {
            TextChanges = result.Paragraphs.Added.Count + result.Paragraphs.Deleted.Count
                        + result.Paragraphs.Modified.Count,
            ParagraphsAdded = result.Paragraphs.Added.Count,
            ParagraphsDeleted = result.Paragraphs.Deleted.Count,
            ParagraphsModified = result.Paragraphs.Modified.Count,
            CommentChanges = result.Comments.Added.Count + result.Comments.Deleted.Count
                           + result.Comments.Modified.Count,
            TrackedChangeChanges = result.TrackedChanges.Added.Count
                                 + result.TrackedChanges.Deleted.Count,
            FormattingChanges = formattingChanges,
            StyleChanges = styleChanges,
            MetadataChanges = result.Metadata.Changes.Count,
            Identical = result.Metadata.Changes.Count == 0
                     && result.Paragraphs.Added.Count == 0
                     && result.Paragraphs.Deleted.Count == 0
                     && result.Paragraphs.Modified.Count == 0
                     && result.Comments.Added.Count == 0
                     && result.Comments.Deleted.Count == 0
                     && result.Comments.Modified.Count == 0
                     && result.TrackedChanges.Added.Count == 0
                     && result.TrackedChanges.Deleted.Count == 0
        };

        return result;
    }

    // ── Metadata ───────────────────────────────────────────────

    private static MetadataDiff DiffMetadata(DocumentMetadata oldMeta, DocumentMetadata newMeta)
    {
        var diff = new MetadataDiff();

        CompareField(diff.Changes, "title", oldMeta.Title, newMeta.Title);
        CompareField(diff.Changes, "author", oldMeta.Author, newMeta.Author);
        CompareField(diff.Changes, "last_modified_by", oldMeta.LastModifiedBy, newMeta.LastModifiedBy);
        CompareField(diff.Changes, "created", oldMeta.Created, newMeta.Created);
        CompareField(diff.Changes, "modified", oldMeta.Modified, newMeta.Modified);

        if (oldMeta.Revision != newMeta.Revision)
            diff.Changes.Add(new FieldChange { Field = "revision", Old = oldMeta.Revision, New = newMeta.Revision });

        if (oldMeta.WordCount != newMeta.WordCount)
            diff.Changes.Add(new FieldChange { Field = "word_count", Old = oldMeta.WordCount, New = newMeta.WordCount });

        return diff;
    }

    private static void CompareField(List<FieldChange> changes, string name, string? oldVal, string? newVal)
    {
        if (oldVal != newVal)
            changes.Add(new FieldChange { Field = name, Old = oldVal, New = newVal });
    }

    // ── Paragraphs ─────────────────────────────────────────────

    private static ParagraphDiff DiffParagraphs(
        List<RichParagraphInfo> oldParas, List<RichParagraphInfo> newParas)
    {
        var diff = new ParagraphDiff();

        // Use LCS on paragraph text to align paragraphs
        var oldTexts = oldParas.Select(p => p.Text).ToList();
        var newTexts = newParas.Select(p => p.Text).ToList();
        var alignment = AlignParagraphs(oldTexts, newTexts);

        foreach (var (oi, ni) in alignment)
        {
            if (oi < 0)
            {
                // Added paragraph
                diff.Added.Add(new ParagraphEntry
                {
                    Index = ni,
                    Style = newParas[ni].Style,
                    Text = newParas[ni].Text
                });
            }
            else if (ni < 0)
            {
                // Deleted paragraph
                diff.Deleted.Add(new ParagraphEntry
                {
                    Index = oi,
                    Style = oldParas[oi].Style,
                    Text = oldParas[oi].Text
                });
            }
            else
            {
                // Matched — check for modifications
                var oldPara = oldParas[oi];
                var newPara = newParas[ni];

                bool textChanged = oldPara.Text != newPara.Text;
                bool styleChanged = oldPara.Style != newPara.Style;
                var fmtChanges = CompareFormatting(oldPara.Runs, newPara.Runs);

                if (textChanged || styleChanged || fmtChanges.Count > 0)
                {
                    var mod = new ParagraphModification
                    {
                        OldIndex = oi,
                        NewIndex = ni,
                        OldText = oldPara.Text,
                        NewText = newPara.Text,
                        FormattingChanges = fmtChanges
                    };

                    if (styleChanged)
                        mod.StyleChange = new StyleChange { Old = oldPara.Style, New = newPara.Style };

                    if (textChanged)
                        mod.WordChanges = ComputeWordDiff(oldPara.Text, newPara.Text);

                    diff.Modified.Add(mod);
                }
            }
        }

        return diff;
    }

    /// <summary>
    /// Align paragraphs using a similarity-based approach.
    /// Returns list of (oldIndex, newIndex) pairs. -1 means added/deleted.
    /// Uses LCS on text hashes for alignment.
    /// </summary>
    private static List<(int oldIdx, int newIdx)> AlignParagraphs(
        List<string> oldTexts, List<string> newTexts)
    {
        int m = oldTexts.Count;
        int n = newTexts.Count;

        // LCS table
        var dp = new int[m + 1, n + 1];
        for (int i = 1; i <= m; i++)
        {
            for (int j = 1; j <= n; j++)
            {
                if (AreSimilar(oldTexts[i - 1], newTexts[j - 1]))
                    dp[i, j] = dp[i - 1, j - 1] + 1;
                else
                    dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
            }
        }

        // Backtrack to get alignment
        var result = new List<(int, int)>();
        int oi = m, ni = n;
        var matched = new List<(int, int)>();

        while (oi > 0 && ni > 0)
        {
            if (AreSimilar(oldTexts[oi - 1], newTexts[ni - 1]))
            {
                matched.Add((oi - 1, ni - 1));
                oi--; ni--;
            }
            else if (dp[oi - 1, ni] > dp[oi, ni - 1])
            {
                oi--;
            }
            else
            {
                ni--;
            }
        }

        matched.Reverse();

        // Build full alignment including unmatched
        var oldMatched = new HashSet<int>(matched.Select(p => p.Item1));
        var newMatched = new HashSet<int>(matched.Select(p => p.Item2));

        int mi = 0, oPtr = 0, nPtr = 0;

        while (mi < matched.Count || oPtr < m || nPtr < n)
        {
            if (mi < matched.Count)
            {
                var (mo, mn) = matched[mi];
                // Emit unmatched before this match
                while (oPtr < mo)
                    result.Add((oPtr++, -1));
                while (nPtr < mn)
                    result.Add((-1, nPtr++));
                // Emit match
                result.Add((mo, mn));
                oPtr = mo + 1;
                nPtr = mn + 1;
                mi++;
            }
            else
            {
                while (oPtr < m)
                    result.Add((oPtr++, -1));
                while (nPtr < n)
                    result.Add((-1, nPtr++));
                break;
            }
        }

        return result;
    }

    /// <summary>
    /// Two paragraphs are "similar" if they share the same text or are close enough.
    /// Exact match or >70% word overlap.
    /// </summary>
    private static bool AreSimilar(string a, string b)
    {
        if (a == b) return true;
        if (string.IsNullOrWhiteSpace(a) && string.IsNullOrWhiteSpace(b)) return true;
        if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b)) return false;

        // Jaccard similarity on words
        var wordsA = new HashSet<string>(a.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        var wordsB = new HashSet<string>(b.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

        if (wordsA.Count == 0 && wordsB.Count == 0) return true;

        int intersection = wordsA.Intersect(wordsB).Count();
        int union = wordsA.Union(wordsB).Count();

        return union > 0 && (double)intersection / union >= 0.5;
    }

    /// <summary>
    /// Compute word-level diff between two strings.
    /// </summary>
    private static List<WordChange> ComputeWordDiff(string oldText, string newText)
    {
        var oldWords = oldText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var newWords = newText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

        var changes = new List<WordChange>();
        var lcs = WordLCS(oldWords, newWords);

        int oi = 0, ni = 0, li = 0;

        while (oi < oldWords.Length || ni < newWords.Length)
        {
            if (li < lcs.Count)
            {
                var (lo, ln) = lcs[li];

                // Deleted words before match
                while (oi < lo)
                {
                    changes.Add(new WordChange { Type = "delete", Old = oldWords[oi], Position = oi });
                    oi++;
                }

                // Added words before match
                while (ni < ln)
                {
                    changes.Add(new WordChange { Type = "add", New = newWords[ni], Position = ni });
                    ni++;
                }

                // Skip matched word
                oi = lo + 1;
                ni = ln + 1;
                li++;
            }
            else
            {
                while (oi < oldWords.Length)
                {
                    changes.Add(new WordChange { Type = "delete", Old = oldWords[oi], Position = oi });
                    oi++;
                }
                while (ni < newWords.Length)
                {
                    changes.Add(new WordChange { Type = "add", New = newWords[ni], Position = ni });
                    ni++;
                }
            }
        }

        // Collapse adjacent delete+add into replace
        return CollapseToReplace(changes);
    }

    private static List<(int, int)> WordLCS(string[] a, string[] b)
    {
        int m = a.Length, n = b.Length;
        var dp = new int[m + 1, n + 1];

        for (int i = 1; i <= m; i++)
            for (int j = 1; j <= n; j++)
                dp[i, j] = a[i - 1] == b[j - 1]
                    ? dp[i - 1, j - 1] + 1
                    : Math.Max(dp[i - 1, j], dp[i, j - 1]);

        var result = new List<(int, int)>();
        int oi2 = m, ni2 = n;
        while (oi2 > 0 && ni2 > 0)
        {
            if (a[oi2 - 1] == b[ni2 - 1])
            {
                result.Add((oi2 - 1, ni2 - 1));
                oi2--; ni2--;
            }
            else if (dp[oi2 - 1, ni2] > dp[oi2, ni2 - 1])
                oi2--;
            else
                ni2--;
        }

        result.Reverse();
        return result;
    }

    private static List<WordChange> CollapseToReplace(List<WordChange> changes)
    {
        var result = new List<WordChange>();

        for (int i = 0; i < changes.Count; i++)
        {
            if (i + 1 < changes.Count
                && changes[i].Type == "delete"
                && changes[i + 1].Type == "add")
            {
                result.Add(new WordChange
                {
                    Type = "replace",
                    Old = changes[i].Old,
                    New = changes[i + 1].New,
                    Position = changes[i].Position
                });
                i++; // skip the add
            }
            else
            {
                result.Add(changes[i]);
            }
        }

        return result;
    }

    // ── Formatting comparison ──────────────────────────────────

    private static List<FormattingChange> CompareFormatting(
        List<RichRunInfo> oldRuns, List<RichRunInfo> newRuns)
    {
        var changes = new List<FormattingChange>();

        // Build word→formatting maps from runs
        var oldFmt = BuildFormattingMap(oldRuns);
        var newFmt = BuildFormattingMap(newRuns);

        // Compare formatting for words that exist in both
        foreach (var word in oldFmt.Keys.Intersect(newFmt.Keys))
        {
            var of = oldFmt[word];
            var nf = newFmt[word];

            if (of.Bold != nf.Bold)
                changes.Add(new FormattingChange
                { Word = word, Property = "bold", OldValue = of.Bold.ToString(), NewValue = nf.Bold.ToString() });

            if (of.Italic != nf.Italic)
                changes.Add(new FormattingChange
                { Word = word, Property = "italic", OldValue = of.Italic.ToString(), NewValue = nf.Italic.ToString() });

            if (of.Underline != nf.Underline)
                changes.Add(new FormattingChange
                { Word = word, Property = "underline", OldValue = of.Underline.ToString(), NewValue = nf.Underline.ToString() });

            if (of.FontName != nf.FontName)
                changes.Add(new FormattingChange
                { Word = word, Property = "font", OldValue = of.FontName, NewValue = nf.FontName });

            if (of.FontSize != nf.FontSize)
                changes.Add(new FormattingChange
                { Word = word, Property = "size", OldValue = of.FontSize, NewValue = nf.FontSize });

            if (of.Color != nf.Color)
                changes.Add(new FormattingChange
                { Word = word, Property = "color", OldValue = of.Color, NewValue = nf.Color });
        }

        return changes;
    }

    private static Dictionary<string, RichRunInfo> BuildFormattingMap(List<RichRunInfo> runs)
    {
        var map = new Dictionary<string, RichRunInfo>();
        foreach (var run in runs)
        {
            foreach (var word in run.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
            {
                if (!map.ContainsKey(word))
                    map[word] = run;
            }
        }
        return map;
    }

    // ── Comments ───────────────────────────────────────────────

    private static CommentDiff DiffComments(List<CommentInfo> oldComments, List<CommentInfo> newComments)
    {
        var diff = new CommentDiff();

        // Match comments by author + anchor text
        var oldByKey = oldComments.ToDictionary(c => $"{c.Author}|{c.AnchorText}", c => c);
        var newByKey = newComments.ToDictionary(c => $"{c.Author}|{c.AnchorText}", c => c);

        foreach (var key in oldByKey.Keys.Except(newByKey.Keys))
            diff.Deleted.Add(oldByKey[key]);

        foreach (var key in newByKey.Keys.Except(oldByKey.Keys))
            diff.Added.Add(newByKey[key]);

        foreach (var key in oldByKey.Keys.Intersect(newByKey.Keys))
        {
            var oldC = oldByKey[key];
            var newC = newByKey[key];
            if (oldC.Text != newC.Text)
            {
                diff.Modified.Add(new CommentModification
                {
                    Author = oldC.Author,
                    AnchorText = oldC.AnchorText,
                    OldText = oldC.Text,
                    NewText = newC.Text
                });
            }
        }

        return diff;
    }

    // ── Tracked changes ────────────────────────────────────────

    private static TrackedChangeDiff DiffTrackedChanges(
        List<RichParagraphInfo> oldParas, List<RichParagraphInfo> newParas)
    {
        var diff = new TrackedChangeDiff();

        var oldTCs = oldParas.SelectMany(p => p.TrackedChanges).ToList();
        var newTCs = newParas.SelectMany(p => p.TrackedChanges).ToList();

        var oldSet = new HashSet<string>(oldTCs.Select(tc => $"{tc.Type}|{tc.Text}|{tc.Author}"));
        var newSet = new HashSet<string>(newTCs.Select(tc => $"{tc.Type}|{tc.Text}|{tc.Author}"));

        diff.Deleted = oldTCs.Where(tc => !newSet.Contains($"{tc.Type}|{tc.Text}|{tc.Author}")).ToList();
        diff.Added = newTCs.Where(tc => !oldSet.Contains($"{tc.Type}|{tc.Text}|{tc.Author}")).ToList();

        return diff;
    }

    // ── Human-readable output ──────────────────────────────────

    public static void PrintHumanReadable(DiffResult result)
    {
        Console.WriteLine();
        Console.WriteLine($"docx-review diff: {result.OldFile} → {result.NewFile}");
        Console.WriteLine(new string('═', 60));

        if (result.Summary.Identical)
        {
            Console.WriteLine("\n  Documents are identical.");
            return;
        }

        // Metadata
        if (result.Metadata.Changes.Count > 0)
        {
            Console.WriteLine("\nMetadata");
            Console.WriteLine(new string('─', 40));
            foreach (var c in result.Metadata.Changes)
            {
                string oldVal = c.Old?.ToString() ?? "(none)";
                string newVal = c.New?.ToString() ?? "(none)";
                Console.WriteLine($"  {c.Field}: {Trunc(oldVal, 40)} → {Trunc(newVal, 40)}");
            }
        }

        // Paragraphs
        int totalTextChanges = result.Paragraphs.Added.Count + result.Paragraphs.Deleted.Count
                             + result.Paragraphs.Modified.Count;
        if (totalTextChanges > 0)
        {
            Console.WriteLine($"\nBody Text ({totalTextChanges} changes)");
            Console.WriteLine(new string('─', 40));

            foreach (var p in result.Paragraphs.Deleted)
            {
                string styleTag = p.Style != null ? $" [{p.Style}]" : "";
                Console.WriteLine($"\n  ¶{p.Index} DELETED{styleTag}:");
                Console.WriteLine($"    - \"{Trunc(p.Text, 72)}\"");
            }

            foreach (var p in result.Paragraphs.Added)
            {
                string styleTag = p.Style != null ? $" [{p.Style}]" : "";
                Console.WriteLine($"\n  ¶{p.Index} ADDED{styleTag}:");
                Console.WriteLine($"    + \"{Trunc(p.Text, 72)}\"");
            }

            foreach (var m in result.Paragraphs.Modified)
            {
                bool textOnly = m.WordChanges.Count > 0;
                bool fmtOnly = m.FormattingChanges.Count > 0 && m.WordChanges.Count == 0;
                bool styleOnly = m.StyleChange != null && m.WordChanges.Count == 0 && m.FormattingChanges.Count == 0;

                string label = styleOnly ? "MODIFIED (style only)" :
                               fmtOnly ? "MODIFIED (formatting only)" :
                               "MODIFIED";

                Console.WriteLine($"\n  ¶{m.OldIndex}→¶{m.NewIndex} {label}:");

                if (m.WordChanges.Count > 0)
                {
                    Console.WriteLine($"    - \"{Trunc(m.OldText, 72)}\"");
                    Console.WriteLine($"    + \"{Trunc(m.NewText, 72)}\"");

                    if (m.WordChanges.Count <= 5)
                    {
                        foreach (var wc in m.WordChanges)
                        {
                            string desc = wc.Type switch
                            {
                                "replace" => $"\"{wc.Old}\" → \"{wc.New}\"",
                                "add" => $"+ \"{wc.New}\"",
                                "delete" => $"- \"{wc.Old}\"",
                                _ => wc.Type
                            };
                            Console.WriteLine($"      {desc}");
                        }
                    }
                }

                if (m.StyleChange != null)
                    Console.WriteLine($"    Style: {m.StyleChange.Old ?? "Normal"} → {m.StyleChange.New ?? "Normal"}");

                foreach (var fc in m.FormattingChanges)
                    Console.WriteLine($"    [{fc.Property}] \"{fc.Word}\": {fc.OldValue} → {fc.NewValue}");
            }
        }

        // Comments
        int commentChanges = result.Comments.Added.Count + result.Comments.Deleted.Count
                           + result.Comments.Modified.Count;
        if (commentChanges > 0)
        {
            Console.WriteLine($"\nComments ({result.Comments.Added.Count} added, "
                + $"{result.Comments.Deleted.Count} removed, {result.Comments.Modified.Count} modified)");
            Console.WriteLine(new string('─', 40));

            foreach (var c in result.Comments.Added)
                Console.WriteLine($"  + [{c.Author}] on \"{Trunc(c.AnchorText, 40)}\" (¶{c.ParagraphIndex}): {c.Text}");

            foreach (var c in result.Comments.Deleted)
                Console.WriteLine($"  - [{c.Author}] on \"{Trunc(c.AnchorText, 40)}\" (¶{c.ParagraphIndex}): {c.Text}");

            foreach (var c in result.Comments.Modified)
            {
                Console.WriteLine($"  ~ [{c.Author}] on \"{Trunc(c.AnchorText, 40)}\":");
                Console.WriteLine($"    old: {Trunc(c.OldText, 60)}");
                Console.WriteLine($"    new: {Trunc(c.NewText, 60)}");
            }
        }

        // Tracked changes
        int tcChanges = result.TrackedChanges.Added.Count + result.TrackedChanges.Deleted.Count;
        if (tcChanges > 0)
        {
            Console.WriteLine($"\nTracked Changes ({result.TrackedChanges.Added.Count} new, "
                + $"{result.TrackedChanges.Deleted.Count} resolved)");
            Console.WriteLine(new string('─', 40));

            foreach (var tc in result.TrackedChanges.Added)
                Console.WriteLine($"  + [{tc.Type}] \"{Trunc(tc.Text, 50)}\" by {tc.Author} {tc.Date ?? ""}");

            foreach (var tc in result.TrackedChanges.Deleted)
                Console.WriteLine($"  - [{tc.Type}] \"{Trunc(tc.Text, 50)}\" by {tc.Author} {tc.Date ?? ""}");
        }

        // Summary
        Console.WriteLine($"\nSummary: {result.Summary.TextChanges} text, "
            + $"{result.Summary.CommentChanges} comment, "
            + $"{result.Summary.TrackedChangeChanges} tracked change, "
            + $"{result.Summary.StyleChanges} style, "
            + $"{result.Summary.FormattingChanges} formatting, "
            + $"{result.Summary.MetadataChanges} metadata");
        Console.WriteLine();
    }

    private static string Trunc(string s, int max) =>
        s.Length <= max ? s : s[..max] + "…";
}
