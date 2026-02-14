# docx-review

A CLI tool that adds **tracked changes** and **comments** to Word (.docx) documents using Microsoft's [Open XML SDK](https://github.com/dotnet/Open-XML-SDK). Takes a `.docx` file and a JSON edit manifest, produces a reviewed document with proper `w:del`/`w:ins` markup and comment anchors that render natively in Microsoft Word — no macros, no compatibility issues.

## Why Open XML SDK?

We evaluated three approaches for programmatic document review:

| Approach | Tracked Changes | Comments | Formatting |
|----------|:-:|:-:|:-:|
| **Open XML SDK (.NET)** | ✅ 100% | ✅ 100% | ✅ Preserved |
| python-docx (Python) | ✅ 100% | ⚠️ ~80% | ✅ Preserved |
| pandoc + Lua filters | ❌ Lossy | ❌ Limited | ⚠️ Degraded |

Open XML SDK is the gold standard — it's Microsoft's own library for manipulating Office documents. Comments anchor correctly 100% of the time, tracked changes use proper revision markup, and formatting is always preserved.

## Quick Start

### Build

```bash
git clone https://github.com/henrybloomingdale/docx-review.git
cd docx-review
docker build -t docx-review .
```

### Run

```bash
# Basic usage
docker run --rm -v "$(pwd):/work" -w /work docx-review input.docx edits.json -o reviewed.docx

# Or use the shell wrapper (after adding to PATH)
./docx-review input.docx edits.json -o reviewed.docx

# Pipe JSON from stdin
cat edits.json | ./docx-review input.docx -o reviewed.docx

# Custom author name
./docx-review input.docx edits.json -o reviewed.docx --author "Dr. Smith"

# Dry run (validate without modifying)
./docx-review input.docx edits.json --dry-run

# JSON output for programmatic use
./docx-review input.docx edits.json -o reviewed.docx --json
```

## JSON Manifest Format

```json
{
  "author": "Reviewer Name",
  "changes": [
    {
      "type": "replace",
      "find": "original text in the document",
      "replace": "replacement text with tracked change"
    },
    {
      "type": "delete",
      "find": "text to mark as deleted"
    },
    {
      "type": "insert_after",
      "anchor": "text to find",
      "text": "new text inserted after the anchor"
    },
    {
      "type": "insert_before",
      "anchor": "text to find",
      "text": "new text inserted before the anchor"
    }
  ],
  "comments": [
    {
      "anchor": "text to anchor the comment to",
      "text": "Comment content displayed in the margin"
    }
  ]
}
```

### Change Types

| Type | Required Fields | Description |
|------|----------------|-------------|
| `replace` | `find`, `replace` | Finds text and creates a tracked replacement (w:del + w:ins) |
| `delete` | `find` | Finds text and marks as deleted (w:del only) |
| `insert_after` | `anchor`, `text` | Finds anchor text, inserts new text after it (w:ins) |
| `insert_before` | `anchor`, `text` | Finds anchor text, inserts new text before it (w:ins) |

### Comment Format

Each comment needs:
- `anchor` — text in the document to attach the comment to (CommentRangeStart/End markers)
- `text` — the comment content shown in Word's review pane

## CLI Flags

| Flag | Description |
|------|-------------|
| `-o`, `--output <path>` | Output file path (default: `<input>_reviewed.docx`) |
| `--author <name>` | Reviewer name for tracked changes (overrides manifest `author`) |
| `--json` | Output results as JSON (for scripting/pipelines) |
| `--dry-run` | Validate the manifest without modifying the document |
| `-h`, `--help` | Show help |

## Exit Codes

- `0` — All changes and comments applied successfully
- `1` — One or more edits failed (partial success possible)

## JSON Output Mode

With `--json`, the tool outputs structured results:

```json
{
  "input": "paper.docx",
  "output": "paper_reviewed.docx",
  "author": "Dr. Smith",
  "changes_attempted": 5,
  "changes_succeeded": 5,
  "comments_attempted": 3,
  "comments_succeeded": 3,
  "success": true,
  "results": [
    { "index": 0, "type": "comment", "success": true, "message": "Comment added" },
    { "index": 0, "type": "replace", "success": true, "message": "Replaced" }
  ]
}
```

## How It Works

1. Copies the input `.docx` to the output path
2. Opens the document using Open XML SDK
3. Adds **comments first** (before tracked changes modify the XML tree)
4. Applies tracked changes (replace → w:del + w:ins, delete → w:del, insert → w:ins)
5. Handles multi-run text matching (text spanning multiple XML runs)
6. Preserves original run formatting (RunProperties cloned from source)
7. Saves and reports results

## Development

Built with .NET 8 and packaged as a Docker container for zero-dependency usage.

```bash
# Build locally (requires .NET 8 SDK)
dotnet build
dotnet run -- input.docx edits.json -o reviewed.docx

# Build Docker image
docker build -t docx-review .
```

## License

MIT — see [LICENSE](LICENSE).

---

*Built by [CinciNeuro](https://github.com/henrybloomingdale) for AI-assisted manuscript review workflows.*
