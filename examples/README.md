# Examples

## sample-edits.json

A template manifest showing all supported change types:

- **replace** — Find text and replace with tracked change (w:del + w:ins)
- **delete** — Find text and mark as deleted (w:del only)
- **insert_after** — Insert new text after an anchor phrase (w:ins)
- **insert_before** — Insert new text before an anchor phrase (w:ins)
- **comments** — Anchor a comment to specific text

## Usage

```bash
# Copy sample and customize for your document
cp sample-edits.json my-edits.json
# Edit my-edits.json with your actual changes

# Run against your document
docx-review mydoc.docx my-edits.json -o mydoc_reviewed.docx

# Validate first with dry-run
docx-review mydoc.docx my-edits.json --dry-run
```
