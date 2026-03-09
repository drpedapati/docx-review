#!/usr/bin/env bash
# Validate a docx-review manifest against a document without modifying it.
# Usage: validate_manifest.sh <input.docx> <manifest.json>
# Exit 0 = all edits match. Exit 1 = at least one failed.
set -euo pipefail

if [ $# -lt 2 ]; then
  echo "Usage: validate_manifest.sh <input.docx> <manifest.json>" >&2
  exit 2
fi

INPUT="$1"
MANIFEST="$2"

if ! command -v docx-review &>/dev/null; then
  echo "Error: docx-review not found. Install with: brew install drpedapati/tools/docx-review" >&2
  exit 2
fi

RESULT=$(docx-review "$INPUT" "$MANIFEST" --dry-run --json 2>&1 || true)
echo "$RESULT" | python3 -c "
import json, sys
data = json.load(sys.stdin)
total = data.get('changes_attempted', 0) + data.get('comments_attempted', 0)
ok = data.get('changes_succeeded', 0) + data.get('comments_succeeded', 0)
failed = [r for r in data.get('results', []) if not r.get('success')]
print(f'{ok}/{total} edits matched')
if failed:
    print('Failed:')
    for f in failed:
        print(f'  [{f.get(\"index\")}] {f.get(\"type\")}: {f.get(\"message\")}')
    sys.exit(1)
"
