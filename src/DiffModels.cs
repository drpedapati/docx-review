using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace DocxReview;

// ── Top-level diff result ──────────────────────────────────────

public class DiffResult
{
    [JsonPropertyName("old_file")]
    public string OldFile { get; set; } = "";

    [JsonPropertyName("new_file")]
    public string NewFile { get; set; } = "";

    [JsonPropertyName("metadata")]
    public MetadataDiff Metadata { get; set; } = new();

    [JsonPropertyName("paragraphs")]
    public ParagraphDiff Paragraphs { get; set; } = new();

    [JsonPropertyName("comments")]
    public CommentDiff Comments { get; set; } = new();

    [JsonPropertyName("tracked_changes")]
    public TrackedChangeDiff TrackedChanges { get; set; } = new();

    [JsonPropertyName("summary")]
    public DiffSummary Summary { get; set; } = new();
}

// ── Metadata diff ──────────────────────────────────────────────

public class MetadataDiff
{
    [JsonPropertyName("changes")]
    public List<FieldChange> Changes { get; set; } = new();
}

public class FieldChange
{
    [JsonPropertyName("field")]
    public string Field { get; set; } = "";

    [JsonPropertyName("old")]
    public object? Old { get; set; }

    [JsonPropertyName("new")]
    public object? New { get; set; }
}

// ── Paragraph diff ─────────────────────────────────────────────

public class ParagraphDiff
{
    [JsonPropertyName("added")]
    public List<ParagraphEntry> Added { get; set; } = new();

    [JsonPropertyName("deleted")]
    public List<ParagraphEntry> Deleted { get; set; } = new();

    [JsonPropertyName("modified")]
    public List<ParagraphModification> Modified { get; set; } = new();
}

public class ParagraphEntry
{
    [JsonPropertyName("index")]
    public int Index { get; set; }

    [JsonPropertyName("style")]
    public string? Style { get; set; }

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";
}

public class ParagraphModification
{
    [JsonPropertyName("old_index")]
    public int OldIndex { get; set; }

    [JsonPropertyName("new_index")]
    public int NewIndex { get; set; }

    [JsonPropertyName("old_text")]
    public string OldText { get; set; } = "";

    [JsonPropertyName("new_text")]
    public string NewText { get; set; } = "";

    [JsonPropertyName("style_change")]
    public StyleChange? StyleChange { get; set; }

    [JsonPropertyName("formatting_changes")]
    public List<FormattingChange> FormattingChanges { get; set; } = new();

    [JsonPropertyName("word_changes")]
    public List<WordChange> WordChanges { get; set; } = new();
}

public class StyleChange
{
    [JsonPropertyName("old")]
    public string? Old { get; set; }

    [JsonPropertyName("new")]
    public string? New { get; set; }
}

public class FormattingChange
{
    [JsonPropertyName("word")]
    public string Word { get; set; } = "";

    [JsonPropertyName("property")]
    public string Property { get; set; } = "";

    [JsonPropertyName("old_value")]
    public string? OldValue { get; set; }

    [JsonPropertyName("new_value")]
    public string? NewValue { get; set; }
}

public class WordChange
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "";  // "add", "delete", "replace"

    [JsonPropertyName("old")]
    public string? Old { get; set; }

    [JsonPropertyName("new")]
    public string? New { get; set; }

    [JsonPropertyName("position")]
    public int Position { get; set; }
}

// ── Comment diff ───────────────────────────────────────────────

public class CommentDiff
{
    [JsonPropertyName("added")]
    public List<CommentInfo> Added { get; set; } = new();

    [JsonPropertyName("deleted")]
    public List<CommentInfo> Deleted { get; set; } = new();

    [JsonPropertyName("modified")]
    public List<CommentModification> Modified { get; set; } = new();
}

public class CommentModification
{
    [JsonPropertyName("author")]
    public string Author { get; set; } = "";

    [JsonPropertyName("anchor_text")]
    public string AnchorText { get; set; } = "";

    [JsonPropertyName("old_text")]
    public string OldText { get; set; } = "";

    [JsonPropertyName("new_text")]
    public string NewText { get; set; } = "";
}

// ── Tracked changes diff ──────────────────────────────────────

public class TrackedChangeDiff
{
    [JsonPropertyName("added")]
    public List<TrackedChangeInfo> Added { get; set; } = new();

    [JsonPropertyName("deleted")]
    public List<TrackedChangeInfo> Deleted { get; set; } = new();
}

// ── Summary ────────────────────────────────────────────────────

public class DiffSummary
{
    [JsonPropertyName("text_changes")]
    public int TextChanges { get; set; }

    [JsonPropertyName("paragraphs_added")]
    public int ParagraphsAdded { get; set; }

    [JsonPropertyName("paragraphs_deleted")]
    public int ParagraphsDeleted { get; set; }

    [JsonPropertyName("paragraphs_modified")]
    public int ParagraphsModified { get; set; }

    [JsonPropertyName("comment_changes")]
    public int CommentChanges { get; set; }

    [JsonPropertyName("tracked_change_changes")]
    public int TrackedChangeChanges { get; set; }

    [JsonPropertyName("formatting_changes")]
    public int FormattingChanges { get; set; }

    [JsonPropertyName("style_changes")]
    public int StyleChanges { get; set; }

    [JsonPropertyName("metadata_changes")]
    public int MetadataChanges { get; set; }

    [JsonPropertyName("identical")]
    public bool Identical { get; set; }
}

// ── Enhanced paragraph info with formatting ────────────────────

public class RichParagraphInfo
{
    public int Index { get; set; }
    public string? Style { get; set; }
    public string Text { get; set; } = "";
    public List<RichRunInfo> Runs { get; set; } = new();
    public List<TrackedChangeInfo> TrackedChanges { get; set; } = new();
}

public class RichRunInfo
{
    public string Text { get; set; } = "";
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public bool Strikethrough { get; set; }
    public string? FontName { get; set; }
    public string? FontSize { get; set; }
    public string? Color { get; set; }
    public string? Highlight { get; set; }
}
