# docx-review --diff: Semantic Document Differencing

## Problem

Git treats `.docx` files as opaque binary blobs. Existing tools (pandoc → text → diff) only capture body text changes and miss the **rich document layers** that matter to reviewers:

- Comments added, deleted, modified, or resolved
- Tracked changes (pending revisions)
- Formatting changes (bold, italic, font, size, color)
- Style changes (Normal → Heading 2)
- Table structure changes (rows/columns added, cell content)
- Image additions/removals/replacements
- Headers and footers
- Metadata (title, author, keywords)
- Section properties (margins, orientation)

## Solution

Add `--diff` and `--textconv` modes to `docx-review`, leveraging the existing Open XML SDK infrastructure.

## CLI Interface

```bash
# Two-file semantic diff (human-readable)
docx-review --diff old.docx new.docx

# Two-file semantic diff (JSON)
docx-review --diff old.docx new.docx --json

# Git textconv mode (single file → normalized text for git diff)
docx-review --textconv file.docx

# Print git configuration instructions
docx-review --git-setup
```

## Architecture

### Document Layers to Compare

| Layer | Priority | What's Compared |
|-------|----------|-----------------|
| **Body Text** | P0 | Paragraph content, paragraph-level LCS diff |
| **Comments** | P0 | Added/removed/modified comments, author, anchor text |
| **Tracked Changes** | P0 | Pending insertions/deletions across versions |
| **Formatting** | P1 | Bold, italic, underline, strikethrough, font name, font size, color, highlight |
| **Styles** | P1 | Paragraph style changes (Normal, Heading1, etc.) |
| **Tables** | P1 | Row/column count, cell content changes |
| **Metadata** | P1 | Title, author, last modified by, revision count |
| **Images** | P2 | Content hash comparison, added/removed |
| **Headers/Footers** | P2 | Content changes per section |
| **Footnotes/Endnotes** | P2 | Added/removed/modified |
| **Section Props** | P2 | Page size, margins, orientation |
| **Hyperlinks** | P2 | Added/removed/changed URLs |

P0 = v1.2.0 (this branch), P1 = v1.3.0, P2 = v1.4.0

### New Source Files

```
src/
├── DocumentDiffer.cs      # Core diff engine
├── DocumentExtractor.cs   # Enhanced extraction (formatting, tables, images)
├── DiffModels.cs          # DiffResult, DiffSection, DiffEntry models
└── TextConv.cs            # Git textconv normalized output
```

### Key Design Decisions

1. **Paragraph-level alignment first, then word-level diff within matched paragraphs.**
   - Use LCS (Longest Common Subsequence) on paragraph hashes to align paragraphs
   - Within matched paragraphs, do word-level diff for precise change detection
   - Unmatched paragraphs = added or deleted

2. **Comments are first-class diff objects.**
   - Match by anchor text + author
   - Show comment text changes, not just presence/absence

3. **Formatting changes are reported per-run, not per-paragraph.**
   - "Word 'methodology' changed from normal to bold in ¶5"
   - Aggregate formatting changes when they span entire paragraphs

4. **Git textconv produces a stable, diffable text representation.**
   - Body text with `[B]`/`[I]`/`[U]` inline markers
   - Comments as `/* [Author] comment text */` after anchor
   - Tracked changes as `[-deleted-]` / `[+inserted+]`
   - Tables as pipe-delimited rows
   - Images as `[IMG: filename (hash)]`
   - This text format is what `git diff` actually processes

## Output Format (Human-Readable)

```
docx-review diff: old.docx → new.docx
══════════════════════════════════════════════════════

Metadata
────────────────────────────────────
  Title:      "Draft v1" → "Final Draft"
  Author:     (unchanged) "Dr. Smith"
  Revision:   3 → 7

Body Text (4 changes)
────────────────────────────────────
  ¶3 MODIFIED:
    - "The methodology was applied to all subjects"
    + "The methods were applied to all participants"

  ¶7 DELETED:
    - "This paragraph was removed entirely."

  ¶12 ADDED:
    + "New paragraph added discussing limitations."

  ¶15 MODIFIED (formatting only):
    "Results" changed: Normal → Heading2

Comments (2 added, 1 removed)
────────────────────────────────────
  + [Dr. Chen] on "all participants" (¶3): "Consider specifying the cohort"
  + [Dr. Chen] on "limitations" (¶12): "Expand this section"
  - [Dr. Smith] on "methodology" (¶3): "Use APA terminology"

Tracked Changes (3 new)
────────────────────────────────────
  [ins] "participants" by Dr. Chen (¶3) — 2024-03-15
  [del] "subjects" by Dr. Chen (¶3) — 2024-03-15
  [ins] "New paragraph..." by Dr. Chen (¶12) — 2024-03-15

Summary: 4 text changes, 3 comment changes, 3 tracked changes, 1 style change
```

## Output Format (JSON)

```json
{
  "old_file": "old.docx",
  "new_file": "new.docx",
  "metadata": {
    "changes": [
      {"field": "title", "old": "Draft v1", "new": "Final Draft"},
      {"field": "revision", "old": 3, "new": 7}
    ]
  },
  "paragraphs": {
    "added": [...],
    "deleted": [...],
    "modified": [
      {
        "index": 3,
        "old_text": "The methodology was applied to all subjects",
        "new_text": "The methods were applied to all participants",
        "word_changes": [
          {"type": "replace", "old": "methodology", "new": "methods", "position": 4},
          {"type": "replace", "old": "subjects", "new": "participants", "position": 38}
        ],
        "formatting_changes": [],
        "style_change": null
      }
    ]
  },
  "comments": {
    "added": [...],
    "deleted": [...],
    "modified": [...]
  },
  "tracked_changes": {
    "added": [...],
    "deleted": [...]
  },
  "summary": {
    "text_changes": 4,
    "comment_changes": 3,
    "tracked_change_changes": 3,
    "formatting_changes": 0,
    "style_changes": 1
  }
}
```

## Git Integration

### textconv mode

`docx-review --textconv` outputs a normalized text that captures all document layers:

```
=== METADATA ===
Title: Final Draft
Author: Dr. Smith
Modified: 2024-03-15T14:30:00Z
Revision: 7

=== BODY ===
¶0 [Heading1] Introduction
¶1 This is the opening paragraph with [B]bold text[/B] and [I]italic text[/I].
¶2 [-deleted text-] [+inserted text+]
¶3 The methods were applied to all participants /* [Dr. Chen] Consider specifying the cohort */

=== TABLES ===
Table 1 (3×4):
| Header 1 | Header 2 | Header 3 | Header 4 |
| Cell 1   | Cell 2   | Cell 3   | Cell 4   |
| Cell 5   | Cell 6   | Cell 7   | Cell 8   |

=== COMMENTS ===
#0 [Dr. Chen] on "all participants" (¶3): Consider specifying the cohort
#1 [Dr. Chen] on "limitations" (¶12): Expand this section

=== IMAGES ===
[IMG] image1.png (sha256: abc123...)
[IMG] figure2.jpg (sha256: def456...)
```

### Setup

```bash
# .gitattributes (add to repo)
*.docx diff=docx

# .gitconfig (add globally or per-repo)
[diff "docx"]
    textconv = docx-review --textconv
```

`docx-review --git-setup` prints these instructions.

## Implementation Plan

### Phase 1: Core Diff (this PR)
1. [x] Create feature/diff branch
2. [ ] Add DiffModels.cs (DiffResult + sub-models)
3. [ ] Add DocumentExtractor.cs (enhanced extraction with formatting)
4. [ ] Add DocumentDiffer.cs (paragraph alignment + comparison)
5. [ ] Add TextConv.cs (normalized text output)
6. [ ] Wire up --diff and --textconv in Program.cs
7. [ ] Test with sample documents
8. [ ] Build and release

### Phase 2: Deep Formatting + Tables
- Run-level formatting comparison
- Table structure and cell-level diffs
- Image content hashing

### Phase 3: Full Coverage
- Headers/footers
- Footnotes/endnotes
- Section properties
- Hyperlinks
- Embedded objects

## Testing Strategy

Create test document pairs:
1. `test_text_changes.docx` — paragraphs added/removed/modified
2. `test_comments.docx` — comments added/removed/modified
3. `test_formatting.docx` — formatting changes only
4. `test_styles.docx` — style changes only
5. `test_combined.docx` — all of the above
6. `test_identical.docx` — no changes (should produce empty diff)
