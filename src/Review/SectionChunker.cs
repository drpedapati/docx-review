using System;
using System.Collections.Generic;
using System.Linq;

namespace DocxReview.Review;

public static class SectionChunker
{
    public const int DefaultChunkSize = 24_000;
    public const int ProofreadChunkSize = 8_000;
    public const int DefaultChunkOverlap = 600;
    public const int DefaultMaxChunks = 32;

    public static List<ReviewChunk> ChunkDocument(
        ReviewTextDocument document,
        IReadOnlyList<ReviewSection>? sections,
        ReviewMode mode,
        int chunkSize = DefaultChunkSize,
        int chunkOverlap = DefaultChunkOverlap,
        int maxChunks = DefaultMaxChunks)
    {
        ArgumentNullException.ThrowIfNull(document);
        return ChunkDocument(document.Text, sections, mode, chunkSize, chunkOverlap, maxChunks);
    }

    public static List<ReviewChunk> ChunkDocument(
        string text,
        IReadOnlyList<ReviewSection>? sections,
        ReviewMode mode,
        int chunkSize = DefaultChunkSize,
        int chunkOverlap = DefaultChunkOverlap,
        int maxChunks = DefaultMaxChunks)
    {
        if (string.IsNullOrWhiteSpace(text))
            return new List<ReviewChunk>();

        chunkSize = chunkSize <= 0 ? DefaultChunkSize : chunkSize;
        chunkOverlap = chunkOverlap < 0 ? DefaultChunkOverlap : chunkOverlap;
        if (chunkOverlap >= chunkSize)
            chunkOverlap = Math.Max(0, chunkSize / 8);

        maxChunks = maxChunks <= 0 ? DefaultMaxChunks : maxChunks;

        if (sections is null || sections.Count == 0)
            return ChunkPlainText(text, chunkSize, chunkOverlap, maxChunks);

        var orderedSections = sections
            .Where(static section => !string.IsNullOrWhiteSpace(section.Text))
            .OrderBy(static section => section.Start)
            .ToList();

        if (orderedSections.Count == 0)
            return ChunkPlainText(text, chunkSize, chunkOverlap, maxChunks);

        var forcePerSection = mode is ReviewMode.Substantive or ReviewMode.PeerReview;
        var chunks = new List<ReviewChunk>();
        var nextChunkIndex = 1;

        var buffer = string.Empty;
        var bufferStart = 0;
        var bufferSections = new List<string>();

        void FlushBuffer()
        {
            if (string.IsNullOrEmpty(buffer))
                return;

            var label = bufferSections.Count switch
            {
                0 => null,
                1 => bufferSections[0],
                _ => $"{bufferSections[0]} ... {bufferSections[^1]}"
            };

            chunks.Add(new ReviewChunk
            {
                Index = nextChunkIndex++,
                Start = bufferStart,
                End = Math.Min(text.Length, bufferStart + buffer.Length),
                Text = buffer,
                SectionName = label
            });

            buffer = string.Empty;
            bufferSections.Clear();
        }

        foreach (var section in orderedSections)
        {
            var sectionText = section.Text;
            if (sectionText.Length > chunkSize)
            {
                FlushBuffer();
                foreach (var piece in ChunkSection(section, nextChunkIndex, chunkSize, chunkOverlap, maxChunks - chunks.Count))
                {
                    chunks.Add(piece);
                    nextChunkIndex = piece.Index + 1;
                    if (chunks.Count >= maxChunks)
                        return chunks;
                }

                continue;
            }

            var isStandaloneMajor = forcePerSection &&
                section.Text.Length >= 500 &&
                IsStandaloneMajorSection(section);

            if (isStandaloneMajor)
            {
                FlushBuffer();
                chunks.Add(new ReviewChunk
                {
                    Index = nextChunkIndex++,
                    Start = Math.Max(0, section.Start),
                    End = Math.Min(text.Length, section.End),
                    Text = sectionText,
                    SectionId = section.Id,
                    SectionKey = NormalizeSectionKey(section.Key),
                    SectionName = string.IsNullOrWhiteSpace(section.Heading) ? section.Key : section.Heading,
                    SectionStyle = section.Style
                });

                if (chunks.Count >= maxChunks)
                    return chunks;

                continue;
            }

            var separator = string.IsNullOrEmpty(buffer) ? string.Empty : Environment.NewLine + Environment.NewLine;
            if (!string.IsNullOrEmpty(buffer) && buffer.Length + separator.Length + sectionText.Length > chunkSize)
                FlushBuffer();

            if (string.IsNullOrEmpty(buffer))
                bufferStart = Math.Max(0, section.Start);
            else
                buffer += Environment.NewLine + Environment.NewLine;

            buffer += sectionText;

            var sectionName = string.IsNullOrWhiteSpace(section.Heading) ? section.Key : section.Heading;
            if (!string.IsNullOrWhiteSpace(sectionName))
                bufferSections.Add(sectionName!);
        }

        FlushBuffer();

        if (chunks.Count == 0)
            return ChunkPlainText(text, chunkSize, chunkOverlap, maxChunks);

        if (chunks.Count > maxChunks)
            chunks = chunks.Take(maxChunks).ToList();

        return chunks;
    }

    private static List<ReviewChunk> ChunkPlainText(string text, int chunkSize, int chunkOverlap, int maxChunks) =>
        ChunkIntoPieces(
            text,
            baseOffset: 0,
            chunkSize,
            chunkOverlap,
            maxChunks,
            static (start, end, pieceText, index) => new ReviewChunk
            {
                Index = index,
                Start = start,
                End = end,
                Text = pieceText
            });

    private static IEnumerable<ReviewChunk> ChunkSection(
        ReviewSection section,
        int startingChunkIndex,
        int chunkSize,
        int chunkOverlap,
        int maxChunks)
    {
        if (maxChunks <= 0)
            yield break;

        var pieces = ChunkIntoPieces(
            section.Text,
            Math.Max(0, section.Start),
            chunkSize,
            chunkOverlap,
            maxChunks,
            (start, end, pieceText, index) => new ReviewChunk
            {
                Index = startingChunkIndex + index - 1,
                Start = start,
                End = end,
                Text = pieceText,
                SectionId = section.Id,
                SectionKey = NormalizeSectionKey(section.Key),
                SectionName = string.IsNullOrWhiteSpace(section.Heading) ? section.Key : section.Heading,
                SectionStyle = section.Style
            });

        foreach (var piece in pieces)
            yield return piece;
    }

    private static List<ReviewChunk> ChunkIntoPieces(
        string text,
        int baseOffset,
        int chunkSize,
        int chunkOverlap,
        int maxChunks,
        Func<int, int, string, int, ReviewChunk> createChunk)
    {
        var chunks = new List<ReviewChunk>();
        var start = 0;

        while (start < text.Length && chunks.Count < maxChunks)
        {
            var rawEnd = Math.Min(text.Length, start + chunkSize);
            var end = BestBreak(text, start, rawEnd);
            if (end <= start)
                end = rawEnd;

            var pieceText = text[start..end];
            chunks.Add(createChunk(baseOffset + start, baseOffset + end, pieceText, chunks.Count + 1));

            if (end >= text.Length)
                break;

            var next = end - chunkOverlap;
            if (next <= start)
                next = end;

            start = next;
        }

        return chunks;
    }

    private static int BestBreak(string text, int start, int end)
    {
        if (end >= text.Length)
            return text.Length;

        var minBreak = start + ((end - start) / 2);
        if (minBreak < start)
            minBreak = start;

        for (var index = end - 1; index >= minBreak; index--)
        {
            var ch = text[index];
            if (ch == '\n' || char.IsWhiteSpace(ch))
                return index + 1;
        }

        return end;
    }

    private static bool IsStandaloneMajorSection(ReviewSection section)
    {
        if (section.Major)
            return true;

        return NormalizeSectionKey(section.Key) is
            "abstract" or
            "introduction" or
            "methods" or
            "results" or
            "discussion" or
            "conclusion";
    }

    private static string? NormalizeSectionKey(string? rawKey)
    {
        var value = (rawKey ?? string.Empty).Trim().ToLowerInvariant();
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
            _ => null
        };
    }
}
