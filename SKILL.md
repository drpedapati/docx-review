---
name: docx-review
description: "Read, edit, and diff Word documents (.docx) with tracked changes and comments using the docx-review CLI — a .NET 8 tool built on Microsoft's Open XML SDK. Ships as a single 12MB native binary (no runtime). Use when: (1) Adding tracked changes (replace, delete, insert) to a .docx, (2) Adding anchored comments to a .docx, (3) Reading/extracting text, tracked changes, comments, and metadata from a .docx, (4) Diffing two .docx files semantically, (5) Responding to peer reviewer comments with tracked revisions, (6) Proofreading or revising manuscripts with reviewable output, (7) Creating new documents from the NIH template, (8) Any task requiring valid .docx output with proper w:del/w:ins markup that renders natively in Word."
metadata: {"nanobot":{"emoji":"📝","requires":{"bins":["docx-review"]},"install":[{"id":"brew","kind":"brew","formula":"henrybloomingdale/tools/docx-review","bins":["docx-review"],"label":"Install docx-review (brew)"}]}}
---

# docx-review

CLI tool for Word document review: tracked changes, comments, read, diff, and git integration. Built on Microsoft's Open XML SDK — 100% compatible tracked changes and comments.

## Install

```bash
brew install henrybloomingdale/tools/docx-review
```

Binary: `/opt/homebrew/bin/docx-review` (12MB, self-contained, no runtime)

Verify: `docx-review --version`

## Template

A NIH-standard Word template is included at `templates/nih-standard.docx`:

- **Font:** Arial 11pt body, Arial Bold 14pt/12pt/11pt for Heading 1/2/3
- **Margins:** 0.75" all sides
- **Colors:** Black only — no color headings, no shading
- **Spacing:** 1.15 line spacing, 6pt after paragraphs
- **Page numbers:** Bottom center, Arial 10pt
- **Sections:** Specific Aims, Significance, Innovation, Approach, Preliminary Data, References

### Creating a new document from template

```bash
# Create a blank NIH template
docx-review --create -o manuscript.docx

# Create and populate in one step
docx-review --create -o manuscript.docx populate.json --json

# Validate a populate manifest first
docx-review --create populate.json --dry-run --json

# Use a custom template instead of the built-in NIH template
docx-review --create --template custom.docx -o output.docx
```

The NIH template is embedded in the binary — no need to locate files on disk.

## Modes

### Create: New document from NIH template

```bash
docx-review --create -o manuscript.docx                           # blank template
docx-review --create -o manuscript.docx populate.json             # create + populate
docx-review --create -o manuscript.docx populate.json --json      # structured output
docx-review --create populate.json --dry-run --json               # validate first
docx-review --create --template custom.docx -o output.docx        # custom template
```

Creates a new `.docx` from the bundled NIH template (Arial 11pt, 0.75" margins, proper heading styles). Optionally populates it with content via a standard edit manifest. Populate edits appear as tracked changes.

### Edit: Apply tracked changes and comments

Takes a `.docx` + JSON manifest, produces a reviewed `.docx` with proper OOXML markup.

```bash
docx-review input.docx edits.json -o reviewed.docx
docx-review input.docx edits.json -o reviewed.docx --json    # structured output
docx-review input.docx edits.json --dry-run --json           # validate without modifying
cat edits.json | docx-review input.docx -o reviewed.docx     # stdin pipe
docx-review input.docx edits.json -o reviewed.docx --author "Dr. Smith"
```

### Read: Extract document content as JSON

```bash
docx-review input.docx --read --json
```

Returns: paragraphs (with styles), tracked changes (type/text/author/date), comments (anchor text/content/author), metadata (title/author/word count/revision), and summary statistics.

### Diff: Semantic comparison of two documents

```bash
docx-review --diff old.docx new.docx
docx-review --diff old.docx new.docx --json
```

Detects: text changes (word-level), formatting (bold/italic/font/color), comment modifications, tracked change differences, metadata changes, structural additions/removals.

### Git: Textconv driver for meaningful Word diffs

```bash
docx-review --textconv document.docx    # normalized text output
docx-review --git-setup                 # print .gitattributes/.gitconfig instructions
```

## JSON Manifest Format

This is the edit contract. Build this JSON, pass it to `docx-review`.

```json
{
  "author": "Reviewer Name",
  "changes": [
    { "type": "replace", "find": "exact text in document", "replace": "new text" },
    { "type": "delete", "find": "exact text to delete" },
    { "type": "insert_after", "anchor": "exact anchor text", "text": "text to insert after" },
    { "type": "insert_before", "anchor": "exact anchor text", "text": "text to insert before" }
  ],
  "comments": [
    { "anchor": "exact text to attach comment to", "text": "Comment content" }
  ]
}
```

### Change types

| Type | Fields | Result in Word |
|------|--------|---------------|
| `replace` | `find`, `replace` | Red strikethrough old + blue new text |
| `delete` | `find` | Red strikethrough |
| `insert_after` | `anchor`, `text` | Blue inserted text after anchor |
| `insert_before` | `anchor`, `text` | Blue inserted text before anchor |

### Critical rules for `find` and `anchor` text

1. **Must be exact copy-paste from the document.** The tool does ordinal string matching.
2. **Include enough context for uniqueness** — 15+ words when the phrase is common.
3. **First occurrence wins.** The tool replaces/anchors at the first match only.
4. Use `--dry-run --json` to validate all matches before applying.

## JSON Output (--json)

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

Exit code 0 = all succeeded. Exit code 1 = at least one failed (partial success possible).

## Workflow: AI-Assisted Document Revision

Standard pattern for using docx-review with AI-generated edits:

### Step 1: Extract text

```bash
docx-review manuscript.docx --read --json > doc_content.json
```

Or use pandoc for markdown extraction:

```bash
pandoc manuscript.docx -t markdown -o manuscript.md
```

### Step 2: Generate the manifest

Feed the extracted text + instructions to the AI. Request output as a docx-review JSON manifest.

Use this system context when prompting for manifest generation:

```
Generate a JSON edit manifest for docx-review. Output format:
{
  "author": "...",
  "changes": [{"type": "replace|delete|insert_after|insert_before", ...}],
  "comments": [{"anchor": "...", "text": "..."}]
}
CRITICAL: "find" and "anchor" values must be EXACT text from the document.
Include 15+ words of surrounding context for uniqueness. First match wins.
```

### Step 3: Validate with dry run

```bash
docx-review manuscript.docx manifest.json --dry-run --json
```

Check for failures. If any edits fail (`"success": false`), fix the manifest (usually the `find`/`anchor` text doesn't match exactly) and retry.

### Step 4: Apply

```bash
docx-review manuscript.docx manifest.json -o manuscript_reviewed.docx --json
```

### Step 5: Verify (optional)

```bash
docx-review manuscript_reviewed.docx --read --json | jq '.summary'
docx-review --diff manuscript.docx manuscript_reviewed.docx
```

## Workflow: Peer Review Response

For addressing reviewer comments on a manuscript:

1. Extract manuscript text (`--read --json` or pandoc)
2. Build manifest addressing each reviewer point — use `replace` for text changes, `comments` to explain changes to the author
3. Dry-run validate
4. Apply edits
5. The output `.docx` has tracked changes the author can review in Word

## Workflow: Proofreading

1. Extract text
2. Generate manifest with grammar/style fixes as `replace` changes and suggestions as `comments`
3. Validate + apply
4. Author opens in Word, accepts/rejects each change individually

## Workflow: New Document from Template

For creating a fresh manuscript from the NIH template:

1. Create the document: `docx-review --create -o manuscript.docx`
2. Build a populate manifest replacing placeholder text with real content
3. Validate: `docx-review --create populate.json --dry-run --json`
4. Apply: `docx-review --create -o manuscript.docx populate.json --json`
5. Open in Word — accept all tracked changes to finalize

Or combine steps 1+4 into a single command. The template's formatting (Arial, margins, page numbers, heading styles) is preserved.

## Key behaviors

- **Comments applied first**, then tracked changes. Ensures anchors resolve before XML is modified.
- **Formatting preserved.** RunProperties cloned from source runs onto both deleted and inserted text.
- **Multi-run text matching.** Text spanning multiple XML `<w:r>` elements (common in previously edited documents) is found and handled correctly.
- **Everything untouched is preserved.** Images, charts, bibliographies, footnotes, cross-references, styles, headers/footers survive intact.

## Read mode output structure

For programmatic processing of `--read --json` output, see `skill/references/read-schema.md`.

## Companion tools

The Open XML SDK ecosystem:

| Tool | Install | Purpose |
|------|---------|---------|
| `pptx-review` | `brew install henrybloomingdale/tools/pptx-review` | PowerPoint read/edit |
| `xlsx-review` | `brew install henrybloomingdale/tools/xlsx-review` | Excel read/edit |

Same architecture: .NET 8, Open XML SDK, single binary, JSON in/out.
