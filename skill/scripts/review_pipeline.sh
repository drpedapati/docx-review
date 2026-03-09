#!/usr/bin/env bash
# Full review pipeline: validate manifest, apply edits, verify output.
# Usage: review_pipeline.sh <input.docx> <manifest.json> [output.docx]
# Exit 0 = success. Exit 1 = validation or apply failure.
set -euo pipefail

if [ $# -lt 2 ]; then
  echo "Usage: review_pipeline.sh <input.docx> <manifest.json> [output.docx]" >&2
  exit 2
fi

INPUT="$1"
MANIFEST="$2"
OUTPUT="${3:-${INPUT%.docx}_reviewed.docx}"

if ! command -v docx-review &>/dev/null; then
  echo "Error: docx-review not found. Install with: brew install drpedapati/tools/docx-review" >&2
  exit 2
fi

echo "==> Validating manifest (dry run)..."
DRY=$(docx-review "$INPUT" "$MANIFEST" --dry-run --json 2>&1 || true)
CHANGES=$(echo "$DRY" | python3 -c "import json,sys; d=json.load(sys.stdin); print(f'{d[\"changes_succeeded\"]}/{d[\"changes_attempted\"]} changes, {d[\"comments_succeeded\"]}/{d[\"comments_attempted\"]} comments')")
echo "    $CHANGES"

FAILED=$(echo "$DRY" | python3 -c "import json,sys; d=json.load(sys.stdin); print(len([r for r in d.get('results',[]) if not r.get('success')]))")
if [ "$FAILED" -gt 0 ]; then
  echo "    $FAILED edit(s) failed validation. Fix manifest and retry." >&2
  echo "$DRY" | python3 -c "
import json,sys
for r in json.load(sys.stdin).get('results',[]):
    if not r.get('success'):
        print(f'    [{r[\"index\"]}] {r[\"type\"]}: {r[\"message\"]}', file=sys.stderr)
"
  exit 1
fi

echo "==> Applying edits..."
RESULT=$(docx-review "$INPUT" "$MANIFEST" -o "$OUTPUT" --json 2>&1)
SUCCESS=$(echo "$RESULT" | python3 -c "import json,sys; print(json.load(sys.stdin).get('success', False))")

if [ "$SUCCESS" = "True" ]; then
  echo "==> Output: $OUTPUT"
  echo "==> Verify: docx-review --diff \"$INPUT\" \"$OUTPUT\""
else
  echo "    Apply failed. Check output:" >&2
  echo "$RESULT" >&2
  exit 1
fi
