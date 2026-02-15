# docx-review --read --json Output Schema

## Top-level structure

```json
{
  "file": "document.docx",
  "paragraphs": [...],
  "comments": [...],
  "metadata": {...},
  "summary": {...}
}
```

## paragraphs[]

```json
{
  "index": 0,
  "style": "Heading1",        // null for Normal/default
  "text": "visible text",     // includes insertions, excludes deletions
  "tracked_changes": [
    {
      "type": "insert",       // "insert" or "delete"
      "text": "changed text",
      "author": "Author Name",
      "date": "2026-02-15T12:00:00Z",
      "id": "100"
    }
  ]
}
```

## comments[]

```json
{
  "id": "0",
  "author": "Reviewer",
  "date": "2026-02-15T12:00:00Z",
  "anchor_text": "the text this comment is attached to",
  "text": "The comment content",
  "paragraph_index": 5
}
```

## metadata

```json
{
  "title": "Document Title",
  "author": "Original Author",
  "last_modified_by": "Last Editor",
  "created": "2026-01-01T00:00:00Z",
  "modified": "2026-02-15T00:00:00Z",
  "revision": 12,
  "word_count": 5432,
  "paragraph_count": 87
}
```

## summary

```json
{
  "total_tracked_changes": 15,
  "insertions": 10,
  "deletions": 5,
  "total_comments": 3,
  "change_authors": ["Author A", "Author B"],
  "comment_authors": ["Reviewer"]
}
```

## Diff mode output (--diff --json)

```json
{
  "old_file": "v1.docx",
  "new_file": "v2.docx",
  "summary": {
    "text_changes": 5,
    "paragraphs_added": 1,
    "paragraphs_deleted": 0,
    "paragraphs_modified": 4,
    "comment_changes": 2,
    "tracked_change_changes": 3,
    "formatting_changes": 1,
    "style_changes": 0,
    "metadata_changes": 2,
    "identical": false
  },
  "metadata": { "changes": [...] },
  "paragraphs": { "added": [...], "deleted": [...], "modified": [...] },
  "comments": { "added": [...], "deleted": [...], "modified": [...] },
  "tracked_changes": { "added": [...], "deleted": [...] }
}
```
