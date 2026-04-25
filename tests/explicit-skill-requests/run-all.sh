#!/usr/bin/env bash
set -euo pipefail

script_dir="$(cd "$(dirname "$0")" && pwd)"
prompts="$script_dir/prompts"
runner="$script_dir/run-behavior-test.sh"

"$runner" "$prompts/docx-tracked-changes.txt" "docx-tracked-changes"
"$runner" "$prompts/word-default-formatting.txt" "word-default-formatting"
"$runner" "$prompts/word-formula-writing.txt" "word-formula-writing"
"$runner" "$prompts/multi-formatting-formulas.txt" "word-default-formatting,word-formula-writing"
"$runner" "$prompts/tracked-formula-combo.txt" "docx-tracked-changes,word-formula-writing"
"$runner" "$prompts/multi-all.txt" "word-default-formatting,word-formula-writing,docx-tracked-changes"
"$runner" "$prompts/formula-only-no-formatting.txt" "-" "word-default-formatting,word-formula-writing,docx-tracked-changes"
"$runner" "$prompts/template-no-default-formatting.txt" "-" "word-default-formatting" 4 "not-applied"
