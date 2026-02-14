using System.Text.Json;
using System.Text.Json.Serialization;
using System.Collections.Generic;

namespace DocxReview;

/// <summary>
/// JSON source generator context for trim-safe / AOT-compatible serialization.
/// </summary>
[JsonSerializable(typeof(EditManifest))]
[JsonSerializable(typeof(ProcessingResult))]
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
