using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocxReview;

namespace DocxReview.Review;

public sealed class ManifestPostProcessResult
{
    public EditManifest Manifest { get; init; } = new();
    public List<int> SuggestionMap { get; init; } = new();
    public List<ReviewSuggestion> Suggestions { get; init; } = new();
    public ManifestPostProcessMetrics Metrics { get; init; } = new();
    public List<string> Warnings { get; init; } = new();
}

public sealed class ManifestPostProcessMetrics
{
    public int InputSuggestions { get; init; }
    public int SuggestionsAfterDedup { get; init; }
    public int SuggestionsDroppedByDedup { get; init; }
    public int SuggestionsDroppedByCap { get; init; }
    public int NoOpChangesDropped { get; set; }
    public int CommentsSkippedForAnchor { get; set; }
    public int DuplicateChangesSkipped { get; set; }
    public int DuplicateCommentsSkipped { get; set; }
}

public sealed class DeduplicateStats
{
    public int InputCount { get; init; }
    public int OutputCount { get; init; }
    public int Removed => InputCount - OutputCount;
}

public sealed class ManifestPostProcessor
{
    public const int MaxTotalEdits = 80;

    private readonly string _sourceText;
    private readonly IReadOnlyDictionary<int, ReviewChunk> _chunksByIndex;
    private readonly IReadOnlyList<ReviewSection> _sections;

    public ManifestPostProcessor(
        string sourceText,
        IReadOnlyList<ReviewChunk> chunks,
        IReadOnlyList<ReviewSection> sections)
    {
        _sourceText = sourceText ?? string.Empty;
        _chunksByIndex = (chunks ?? Array.Empty<ReviewChunk>()).ToDictionary(static chunk => chunk.Index);
        _sections = sections ?? Array.Empty<ReviewSection>();
    }

    public ManifestPostProcessResult Process(
        IReadOnlyList<ReviewSuggestion> suggestions,
        string? author,
        ReviewMode mode = ReviewMode.Substantive)
    {
        ArgumentNullException.ThrowIfNull(suggestions);

        var deduped = DeduplicateEdits(suggestions);
        var capped = deduped.Suggestions;
        var droppedByCap = 0;
        // No edit cap for proofread mode — spelling/grammar fixes are numerous and all valid
        if (mode != ReviewMode.Proofread && capped.Count > MaxTotalEdits)
        {
            droppedByCap = capped.Count - MaxTotalEdits;
            capped = capped.Take(MaxTotalEdits).ToList();
        }

        var metrics = new ManifestPostProcessMetrics
        {
            InputSuggestions = suggestions.Count,
            SuggestionsAfterDedup = capped.Count,
            SuggestionsDroppedByDedup = deduped.Stats.Removed,
            SuggestionsDroppedByCap = droppedByCap
        };

        var warnings = new List<string>();
        if (deduped.Stats.Removed > 0)
            warnings.Add($"Deduplicated {deduped.Stats.Removed} overlapping suggestions from adjacent chunks.");
        if (droppedByCap > 0)
            warnings.Add($"Capped the review output at {MaxTotalEdits} suggestions.");

        var manifest = new EditManifest
        {
            Author = string.IsNullOrWhiteSpace(author) ? "Reviewer" : author.Trim(),
            Changes = new List<Change>(),
            Comments = new List<CommentDef>()
        };

        var suggestionMap = new List<int>();
        var seenChanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenComments = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var index = 0; index < capped.Count; index++)
        {
            var suggestion = capped[index];
            if (IsCommentOnlySuggestion(suggestion))
            {
                var anchor = ResolveCommentAnchor(suggestion);
                if (string.IsNullOrWhiteSpace(anchor))
                {
                    metrics.CommentsSkippedForAnchor++;
                    continue;
                }

                var commentText = NormalizeLineText(suggestion.Rationale);
                if (string.IsNullOrWhiteSpace(commentText))
                    continue;

                var comment = new CommentDef
                {
                    Anchor = anchor,
                    Text = commentText
                };

                var commentKey = BuildCommentKey(comment);
                if (!seenComments.Add(commentKey))
                {
                    metrics.DuplicateCommentsSkipped++;
                    continue;
                }

                manifest.Comments!.Add(comment);
                continue;
            }

            if (!TryConvertChangeSuggestion(suggestion, out var change, out var noOp))
            {
                if (noOp)
                    metrics.NoOpChangesDropped++;
                continue;
            }

            var changeKey = BuildChangeKey(change!);
            if (!seenChanges.Add(changeKey))
            {
                metrics.DuplicateChangesSkipped++;
                continue;
            }

            manifest.Changes!.Add(change!);
            suggestionMap.Add(index);
        }

        SortChangesBackToFront(manifest.Changes!, suggestionMap, capped);

        return new ManifestPostProcessResult
        {
            Manifest = manifest,
            SuggestionMap = suggestionMap,
            Suggestions = capped,
            Metrics = metrics,
            Warnings = warnings
        };
    }

    public (List<ReviewSuggestion> Suggestions, DeduplicateStats Stats) DeduplicateEdits(IReadOnlyList<ReviewSuggestion> suggestions)
    {
        var locations = new Dictionary<string, List<(int Chunk, int Index)>>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < suggestions.Count; index++)
        {
            var key = NormalizeForDedupKey(suggestions[index]);
            if (string.IsNullOrWhiteSpace(key))
                continue;

            if (!locations.TryGetValue(key, out var list))
            {
                list = new List<(int Chunk, int Index)>();
                locations[key] = list;
            }

            list.Add((suggestions[index].Chunk, index));
        }

        var remove = new HashSet<int>();
        foreach (var group in locations.Values)
        {
            if (group.Count < 2)
                continue;

            for (var index = 1; index < group.Count; index++)
            {
                if (Math.Abs(group[index].Chunk - group[index - 1].Chunk) <= 1)
                    remove.Add(group[index].Index);
            }
        }

        var filtered = new List<ReviewSuggestion>(suggestions.Count - remove.Count);
        for (var index = 0; index < suggestions.Count; index++)
        {
            if (!remove.Contains(index))
                filtered.Add(suggestions[index]);
        }

        return (filtered, new DeduplicateStats
        {
            InputCount = suggestions.Count,
            OutputCount = filtered.Count
        });
    }

    public List<CommentDef> BuildFallbackComments(
        ProcessingResult processingResult,
        ManifestPostProcessResult postProcessResult)
    {
        ArgumentNullException.ThrowIfNull(processingResult);
        ArgumentNullException.ThrowIfNull(postProcessResult);

        var fallbackComments = new List<CommentDef>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var existing in postProcessResult.Manifest.Comments ?? Enumerable.Empty<CommentDef>())
        {
            seen.Add(BuildCommentKey(existing));
        }

        foreach (var result in processingResult.Results)
        {
            if (result.Success || !IsTrackedChangeResult(result.Type))
                continue;

            if (result.Index < 0 || result.Index >= (postProcessResult.Manifest.Changes?.Count ?? 0))
                continue;

            var change = postProcessResult.Manifest.Changes![result.Index];
            ReviewSuggestion? suggestion = null;
            if (result.Index < postProcessResult.SuggestionMap.Count)
            {
                var suggestionIndex = postProcessResult.SuggestionMap[result.Index];
                if (suggestionIndex >= 0 && suggestionIndex < postProcessResult.Suggestions.Count)
                    suggestion = postProcessResult.Suggestions[suggestionIndex];
            }

            var anchor = ResolveFallbackAnchor(change, suggestion);
            if (string.IsNullOrWhiteSpace(anchor))
                continue;

            var comment = new CommentDef
            {
                Anchor = anchor,
                Text = BuildFallbackCommentText(change, suggestion)
            };

            var key = BuildCommentKey(comment);
            if (!seen.Add(key))
                continue;

            fallbackComments.Add(comment);
        }

        return fallbackComments;
    }

    private bool TryConvertChangeSuggestion(ReviewSuggestion suggestion, out Change? change, out bool noOp)
    {
        change = null;
        noOp = false;

        var original = NormalizeLineText(suggestion.Original);
        var revised = NormalizeLineText(suggestion.Revised);
        var anchor = NormalizeLineText(suggestion.Anchor);
        var type = (suggestion.Type ?? string.Empty).Trim().ToLowerInvariant();

        if (!string.IsNullOrWhiteSpace(original) && revised is not null && revised.Length > 0)
        {
            if (NormalizeDocxText(original) == NormalizeDocxText(revised))
            {
                noOp = true;
                return false;
            }

            change = new Change
            {
                Type = "replace",
                Find = original,
                Replace = revised
            };
            return true;
        }

        if (!string.IsNullOrWhiteSpace(original) && string.IsNullOrWhiteSpace(revised))
        {
            change = new Change
            {
                Type = "delete",
                Find = original
            };
            return true;
        }

        if (!string.IsNullOrWhiteSpace(revised) && !string.IsNullOrWhiteSpace(anchor))
        {
            change = new Change
            {
                Type = type == "insert_before" ? "insert_before" : "insert_after",
                Anchor = anchor,
                Text = revised
            };
            return true;
        }

        return false;
    }

    private bool IsCommentOnlySuggestion(ReviewSuggestion suggestion)
    {
        if (string.Equals(suggestion.Type, "comment", StringComparison.OrdinalIgnoreCase))
            return true;

        return string.IsNullOrWhiteSpace(suggestion.Original) &&
               string.IsNullOrWhiteSpace(suggestion.Revised) &&
               !string.IsNullOrWhiteSpace(suggestion.Rationale);
    }

    private string? ResolveCommentAnchor(ReviewSuggestion suggestion)
    {
        var candidate = NormalizeLineText(FirstNonEmpty(suggestion.Anchor, suggestion.Original, suggestion.Revised));
        if (string.IsNullOrWhiteSpace(candidate))
            return ResolveChunkOrSectionFallback(suggestion);

        var exact = ResolveExactAnchor(candidate);
        if (!string.IsNullOrWhiteSpace(exact))
            return exact;

        var fuzzy = ResolveFuzzyAnchor(candidate, _sourceText);
        if (!string.IsNullOrWhiteSpace(fuzzy))
            return fuzzy;

        return ResolveChunkOrSectionFallback(suggestion);
    }

    private string? ResolveFallbackAnchor(Change change, ReviewSuggestion? suggestion)
    {
        var candidate = NormalizeLineText(FirstNonEmpty(
            suggestion?.Anchor,
            change.Anchor,
            change.Find,
            suggestion?.Original,
            suggestion?.Revised,
            change.Text));

        if (!string.IsNullOrWhiteSpace(candidate))
        {
            var exact = ResolveExactAnchor(candidate);
            if (!string.IsNullOrWhiteSpace(exact))
                return exact;

            var fuzzy = ResolveFuzzyAnchor(candidate, _sourceText);
            if (!string.IsNullOrWhiteSpace(fuzzy))
                return fuzzy;
        }

        return suggestion is null ? null : ResolveChunkOrSectionFallback(suggestion);
    }

    private string? ResolveChunkOrSectionFallback(ReviewSuggestion suggestion)
    {
        if (_chunksByIndex.TryGetValue(suggestion.Chunk, out var chunk))
        {
            var chunkSegment = SliceSource(chunk.Start, chunk.End);
            var chunkAnchor = ResolveFuzzyAnchor(NormalizeLineText(suggestion.Anchor), chunkSegment);
            if (!string.IsNullOrWhiteSpace(chunkAnchor))
                return chunkAnchor;

            var snippet = SnippetFromText(chunkSegment, 80);
            if (!string.IsNullOrWhiteSpace(snippet))
                return snippet;
        }

        var section = ResolveSection(suggestion);
        if (section is not null)
        {
            var sectionAnchor = ResolveFuzzyAnchor(NormalizeLineText(suggestion.Anchor), section.Text);
            if (!string.IsNullOrWhiteSpace(sectionAnchor))
                return sectionAnchor;

            var snippet = SnippetFromText(section.Text, 80);
            if (!string.IsNullOrWhiteSpace(snippet))
                return snippet;
        }

        return null;
    }

    private ReviewSection? ResolveSection(ReviewSuggestion suggestion)
    {
        if (suggestion.Metadata is not null)
        {
            if (suggestion.Metadata.TryGetValue("section_id", out var sectionId) &&
                !string.IsNullOrWhiteSpace(sectionId))
            {
                var byId = _sections.FirstOrDefault(section => string.Equals(section.Id, sectionId, StringComparison.OrdinalIgnoreCase));
                if (byId is not null)
                    return byId;
            }

            if (suggestion.Metadata.TryGetValue("section_name", out var sectionName) &&
                !string.IsNullOrWhiteSpace(sectionName))
            {
                var byName = _sections.FirstOrDefault(section => string.Equals(section.Heading, sectionName, StringComparison.OrdinalIgnoreCase));
                if (byName is not null)
                    return byName;
            }
        }

        if (_chunksByIndex.TryGetValue(suggestion.Chunk, out var chunk) && !string.IsNullOrWhiteSpace(chunk.SectionId))
        {
            return _sections.FirstOrDefault(section => string.Equals(section.Id, chunk.SectionId, StringComparison.OrdinalIgnoreCase));
        }

        return null;
    }

    private string? ResolveExactAnchor(string candidate)
    {
        var index = _sourceText.IndexOf(candidate, StringComparison.Ordinal);
        if (index >= 0)
            return _sourceText.Substring(index, candidate.Length);

        index = _sourceText.IndexOf(candidate, StringComparison.OrdinalIgnoreCase);
        return index >= 0 ? _sourceText.Substring(index, candidate.Length) : null;
    }

    private static string? ResolveFuzzyAnchor(string? candidate, string? sourceText)
    {
        if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(sourceText))
            return null;

        var trimmed = TrimFuzzyAnchor(candidate);
        if (string.IsNullOrWhiteSpace(trimmed))
            return null;

        var exactIndex = sourceText.IndexOf(trimmed, StringComparison.OrdinalIgnoreCase);
        if (exactIndex >= 0)
            return sourceText.Substring(exactIndex, trimmed.Length);

        var words = trimmed
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(static word => word.Length > 0)
            .ToArray();

        for (var width = Math.Min(words.Length, 8); width >= 3; width--)
        {
            for (var start = 0; start <= words.Length - width; start++)
            {
                var phrase = string.Join(" ", words, start, width);
                var index = sourceText.IndexOf(phrase, StringComparison.OrdinalIgnoreCase);
                if (index >= 0)
                    return sourceText.Substring(index, phrase.Length);
            }
        }

        return null;
    }

    private void SortChangesBackToFront(List<Change> changes, List<int> suggestionMap, IReadOnlyList<ReviewSuggestion> suggestions)
    {
        if (changes.Count <= 1)
            return;

        var sortable = changes
            .Select((change, index) => new SortableChange
            {
                Change = change,
                OriginalIndex = index,
                SuggestionIndex = index < suggestionMap.Count ? suggestionMap[index] : -1,
                Position = ResolveAbsolutePosition(change, index < suggestionMap.Count ? suggestions[suggestionMap[index]] : null)
            })
            .OrderBy(static item => item.Position < 0 ? 1 : 0)
            .ThenByDescending(static item => item.Position)
            .ThenBy(static item => item.OriginalIndex)
            .ToList();

        for (var index = 0; index < sortable.Count; index++)
        {
            changes[index] = sortable[index].Change;
            suggestionMap[index] = sortable[index].SuggestionIndex;
        }
    }

    private int ResolveAbsolutePosition(Change change, ReviewSuggestion? suggestion)
    {
        var fragment = FirstNonEmpty(change.Find, change.Anchor);
        if (!string.IsNullOrWhiteSpace(fragment))
        {
            var exact = _sourceText.IndexOf(fragment, StringComparison.Ordinal);
            if (exact >= 0)
                return exact;

            var ignoreCase = _sourceText.IndexOf(fragment, StringComparison.OrdinalIgnoreCase);
            if (ignoreCase >= 0)
                return ignoreCase;

            var fuzzy = ResolveFuzzyAnchor(fragment, _sourceText);
            if (!string.IsNullOrWhiteSpace(fuzzy))
            {
                var fuzzyIndex = _sourceText.IndexOf(fuzzy, StringComparison.Ordinal);
                if (fuzzyIndex >= 0)
                    return fuzzyIndex;
            }
        }

        if (suggestion is not null && _chunksByIndex.TryGetValue(suggestion.Chunk, out var chunk))
            return chunk.Start;

        return -1;
    }

    private string SliceSource(int start, int end)
    {
        if (_sourceText.Length == 0)
            return string.Empty;

        start = Math.Clamp(start, 0, _sourceText.Length);
        end = Math.Clamp(end, start, _sourceText.Length);
        return start >= end ? string.Empty : _sourceText[start..end];
    }

    private static string BuildChangeKey(Change change) =>
        string.Join(
            "|",
            (change.Type ?? string.Empty).Trim().ToLowerInvariant(),
            NormalizeDocxText(change.Find),
            NormalizeDocxText(change.Replace),
            NormalizeDocxText(change.Anchor),
            NormalizeDocxText(change.Text));

    private static string BuildCommentKey(CommentDef comment) =>
        string.Join(
            "|",
            NormalizeDocxText(comment.Anchor),
            NormalizeDocxText(comment.Text));

    private static string NormalizeForDedupKey(ReviewSuggestion suggestion)
    {
        var anchor = FirstNonEmpty(suggestion.Anchor, suggestion.Original, suggestion.Revised);
        if (string.IsNullOrWhiteSpace(anchor))
            return string.Empty;

        return string.Join(
            "|",
            (suggestion.Type ?? string.Empty).Trim().ToLowerInvariant(),
            NormalizeDocxText(anchor));
    }

    private static string NormalizeDocxText(string? value) =>
        string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));

    private static string NormalizeLineText(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var builder = new StringBuilder();
        foreach (var line in value
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (builder.Length > 0)
                builder.Append(' ');
            builder.Append(string.Join(" ", line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)));
        }

        return builder.ToString();
    }

    private static string? SnippetFromText(string? text, int maxChars)
    {
        var normalized = NormalizeLineText(text);
        if (string.IsNullOrWhiteSpace(normalized))
            return null;

        if (normalized.Length <= maxChars)
            return normalized;

        var cut = normalized.LastIndexOf(' ', maxChars);
        if (cut < 20)
            cut = maxChars;

        return normalized[..cut].Trim();
    }

    private static string TrimFuzzyAnchor(string value)
    {
        var trimmed = value.Trim();
        while (trimmed.StartsWith("...", StringComparison.Ordinal) || trimmed.StartsWith("…", StringComparison.Ordinal))
        {
            trimmed = trimmed.StartsWith("...", StringComparison.Ordinal)
                ? trimmed[3..].Trim()
                : trimmed[1..].Trim();
        }

        while (trimmed.EndsWith("...", StringComparison.Ordinal) || trimmed.EndsWith("…", StringComparison.Ordinal))
        {
            trimmed = trimmed.EndsWith("...", StringComparison.Ordinal)
                ? trimmed[..^3].Trim()
                : trimmed[..^1].Trim();
        }

        return trimmed.Trim('"', '\'').Trim();
    }

    private static string BuildFallbackCommentText(Change change, ReviewSuggestion? suggestion)
    {
        string text = (change.Type ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "replace" => $"Consider changing '{SnippetFromText(change.Find, 60)}' to '{SnippetFromText(change.Replace, 60)}'.",
            "delete" => $"Consider removing '{SnippetFromText(change.Find, 60)}'.",
            "insert_before" => $"Consider adding '{SnippetFromText(change.Text, 60)}' before '{SnippetFromText(change.Anchor, 60)}'.",
            _ => $"Consider adding '{SnippetFromText(change.Text, 60)}' after '{SnippetFromText(change.Anchor, 60)}'."
        };

        if (!string.IsNullOrWhiteSpace(suggestion?.Rationale))
            text = text.TrimEnd('.') + " " + suggestion.Rationale!.Trim();

        return text;
    }

    private static bool IsTrackedChangeResult(string? type) =>
        string.Equals(type, "replace", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(type, "delete", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(type, "insert_after", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(type, "insert_before", StringComparison.OrdinalIgnoreCase);

    private static string FirstNonEmpty(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
                return value.Trim();
        }

        return string.Empty;
    }

    private sealed class SortableChange
    {
        public Change Change { get; set; } = new();
        public int OriginalIndex { get; set; }
        public int SuggestionIndex { get; set; }
        public int Position { get; set; }
    }
}
