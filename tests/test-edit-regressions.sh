#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
BUILD_DIR="$PROJECT_DIR/build"
BIN="$BUILD_DIR/docx-review"
PASSED=0
FAILED=0

run_docx() {
    "$BIN" "$@" 2>&1
}

run_docx_stdin() {
    "$BIN" "$@" 2>&1
}

assert_pass() {
    local name="$1"
    echo "  [PASS] $name"
    PASSED=$((PASSED + 1))
}

assert_fail() {
    local name="$1"
    local detail="${2:-}"
    echo "  [FAIL] $name"
    [[ -n "$detail" ]] && echo "    $detail"
    FAILED=$((FAILED + 1))
}

if ! command -v jq >/dev/null 2>&1; then
    echo "Error: jq is required for this test"
    exit 1
fi

echo "=== Edit Regression Tests ==="

mkdir -p "$BUILD_DIR"
(cd "$PROJECT_DIR" && make build >/dev/null)

INPLACE_DOC="$BUILD_DIR/in_place_regression.docx"
CHAIN_BASE="$BUILD_DIR/chained_insert_base.docx"
CHAIN_OUT="$BUILD_DIR/chained_insert_out.docx"

run_docx --create -o "$INPLACE_DOC" --json >/dev/null
BEFORE_TEXT=$(run_docx "$INPLACE_DOC" --textconv)

RESULT=$(cat <<'JSON' | run_docx_stdin "$INPLACE_DOC" --in-place --json || true
{
  "changes": [
    {"type": "replace", "find": "Specific Aims", "replace": "Specific Goals"},
    {"type": "replace", "find": "__missing_anchor__", "replace": "Should fail"}
  ]
}
JSON
)

SUCCESS=$(echo "$RESULT" | jq -r '.success')
FIRST_OK=$(echo "$RESULT" | jq -r '.results[0].success')
SECOND_OK=$(echo "$RESULT" | jq -r '.results[1].success')
AFTER_TEXT=$(run_docx "$INPLACE_DOC" --textconv)

if [[ "$SUCCESS" == "false" && "$FIRST_OK" == "true" && "$SECOND_OK" == "false" && "$BEFORE_TEXT" == "$AFTER_TEXT" ]]; then
    assert_pass "--in-place keeps original file when any edit fails"
else
    assert_fail "--in-place overwrote the source on partial failure" "success=$SUCCESS first=$FIRST_OK second=$SECOND_OK"
fi

run_docx --create -o "$CHAIN_BASE" --json >/dev/null

RESULT=$(cat <<'JSON' | run_docx_stdin "$CHAIN_BASE" -o "$CHAIN_OUT" --json
{
  "changes": [
    {"type": "insert_after", "anchor": "Specific Aims", "text": "Lead-in\n__CHAIN_TARGET__"},
    {"type": "insert_after", "anchor": "__CHAIN_TARGET__", "text": "Chained insert reached the new paragraph"}
  ]
}
JSON
)

SUCCESS=$(echo "$RESULT" | jq -r '.success')
CHANGES_SUCCEEDED=$(echo "$RESULT" | jq -r '.changes_succeeded')
CHAIN_TEXT=$(run_docx "$CHAIN_OUT" --textconv)

if [[ "$SUCCESS" == "true" && "$CHANGES_SUCCEEDED" == "2" && "$CHAIN_TEXT" == *"__CHAIN_TARGET__"* && "$CHAIN_TEXT" == *"Chained insert reached the new paragraph"* ]]; then
    assert_pass "later edits can target paragraphs created earlier in the manifest"
else
    assert_fail "chained multi-paragraph edit failed" "success=$SUCCESS changes_succeeded=$CHANGES_SUCCEEDED"
fi

rm -f "$INPLACE_DOC" "$CHAIN_BASE" "$CHAIN_OUT"

echo "=== Edit Regression Tests: $PASSED passed, $FAILED failed ==="
if [[ $FAILED -gt 0 ]]; then
    exit 1
fi
