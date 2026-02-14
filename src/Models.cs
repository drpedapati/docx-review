using System.Text.Json;
using System.Text.Json.Serialization;
using System.Collections.Generic;

namespace DocxReview;

/// <summary>
/// JSON source generator context for trim-safe / AOT-compatible serialization.
/// </summary>
[JsonSerializable(typeof(EditManifest))]
[JsonSerializable(typeof(ProcessingResult))]
[JsonSerializable(typeof(ReadResult))]
[JsonSourceGenerationOptions(
    PropertyNameCaseInsensitive = true,
    WriteIndented = true
)]
public partial class DocxReviewJsonContext : JsonSerializerContext { }

/// <summary>
/// Root manifest model deserialized from the JSON input.
/// </summary>
public class EditManifest
{
    [JsonPropertyName("author")]
    public string? Author { get; set; }

    [JsonPropertyName("changes")]
    public List<Change>? Changes { get; set; }

    [JsonPropertyName("comments")]
    public List<CommentDef>? Comments { get; set; }
}

/// <summary>
/// A single tracked change (replace, delete, insert_after, insert_before).
/// </summary>
public class Change
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "replace";

    [JsonPropertyName("find")]
    public string? Find { get; set; }

    [JsonPropertyName("replace")]
    public string? Replace { get; set; }

    [JsonPropertyName("anchor")]
    public string? Anchor { get; set; }

    [JsonPropertyName("text")]
    public string? Text { get; set; }
}

/// <summary>
/// A comment anchored to specific text.
/// </summary>
public class CommentDef
{
    [JsonPropertyName("anchor")]
    public string Anchor { get; set; } = "";

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";
}

/// <summary>
/// Result of processing a single edit or comment.
/// </summary>
public class EditResult
{
    [JsonPropertyName("index")]
    public int Index { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
}

/// <summary>
/// Overall result summary for JSON output mode.
/// </summary>
public class ProcessingResult
{
    [JsonPropertyName("input")]
    public string Input { get; set; } = "";

    [JsonPropertyName("output")]
    public string? Output { get; set; }

    [JsonPropertyName("author")]
    public string Author { get; set; } = "";

    [JsonPropertyName("changes_attempted")]
    public int ChangesAttempted { get; set; }

    [JsonPropertyName("changes_succeeded")]
    public int ChangesSucceeded { get; set; }

    [JsonPropertyName("comments_attempted")]
    public int CommentsAttempted { get; set; }

    [JsonPropertyName("comments_succeeded")]
    public int CommentsSucceeded { get; set; }

    [JsonPropertyName("results")]
    public List<EditResult> Results { get; set; } = new();

    [JsonPropertyName("success")]
    public bool Success { get; set; }
}

// ── Read-mode models ──────────────────────────────────────────────

/// <summary>
/// Top-level result for --read mode.
/// </summary>
public class ReadResult
{
    [JsonPropertyName("file")]
    public string File { get; set; } = "";

    [JsonPropertyName("paragraphs")]
    public List<ParagraphInfo> Paragraphs { get; set; } = new();

    [JsonPropertyName("comments")]
    public List<CommentInfo> Comments { get; set; } = new();

    [JsonPropertyName("metadata")]
    public DocumentMetadata Metadata { get; set; } = new();

    [JsonPropertyName("summary")]
    public ReadSummary Summary { get; set; } = new();
}

/// <summary>
/// A single paragraph with its text, style, and tracked changes.
/// </summary>
public class ParagraphInfo
{
    [JsonPropertyName("index")]
    public int Index { get; set; }

    [JsonPropertyName("style")]
    public string? Style { get; set; }

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";

    [JsonPropertyName("tracked_changes")]
    public List<TrackedChangeInfo> TrackedChanges { get; set; } = new();
}

/// <summary>
/// An individual tracked change (insertion or deletion).
/// </summary>
public class TrackedChangeInfo
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "";

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";

    [JsonPropertyName("author")]
    public string Author { get; set; } = "";

    [JsonPropertyName("date")]
    public string? Date { get; set; }

    [JsonPropertyName("id")]
    public string Id { get; set; } = "";
}

/// <summary>
/// A comment extracted from the document.
/// </summary>
public class CommentInfo
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = "";

    [JsonPropertyName("author")]
    public string Author { get; set; } = "";

    [JsonPropertyName("date")]
    public string? Date { get; set; }

    [JsonPropertyName("anchor_text")]
    public string AnchorText { get; set; } = "";

    [JsonPropertyName("text")]
    public string Text { get; set; } = "";

    [JsonPropertyName("paragraph_index")]
    public int ParagraphIndex { get; set; }
}

/// <summary>
/// Document metadata from core properties.
/// </summary>
public class DocumentMetadata
{
    [JsonPropertyName("title")]
    public string? Title { get; set; }

    [JsonPropertyName("author")]
    public string? Author { get; set; }

    [JsonPropertyName("last_modified_by")]
    public string? LastModifiedBy { get; set; }

    [JsonPropertyName("created")]
    public string? Created { get; set; }

    [JsonPropertyName("modified")]
    public string? Modified { get; set; }

    [JsonPropertyName("revision")]
    public int? Revision { get; set; }

    [JsonPropertyName("word_count")]
    public int WordCount { get; set; }

    [JsonPropertyName("paragraph_count")]
    public int ParagraphCount { get; set; }
}

/// <summary>
/// Aggregated summary statistics for --read mode.
/// </summary>
public class ReadSummary
{
    [JsonPropertyName("total_tracked_changes")]
    public int TotalTrackedChanges { get; set; }

    [JsonPropertyName("insertions")]
    public int Insertions { get; set; }

    [JsonPropertyName("deletions")]
    public int Deletions { get; set; }

    [JsonPropertyName("total_comments")]
    public int TotalComments { get; set; }

    [JsonPropertyName("change_authors")]
    public List<string> ChangeAuthors { get; set; } = new();

    [JsonPropertyName("comment_authors")]
    public List<string> CommentAuthors { get; set; } = new();
}
