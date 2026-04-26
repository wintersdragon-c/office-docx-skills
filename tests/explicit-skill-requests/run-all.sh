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
"$runner" "$prompts/docx-bilingual-translation.txt" "docx-bilingual-translation"
"$runner" "$prompts/docx-format-audit.txt" "docx-format-audit"
"$runner" "$prompts/formula-embedded-object-guidance.txt" "word-formula-writing"
"$runner" "$prompts/translation-format-formula-combo.txt" "docx-bilingual-translation,word-default-formatting,word-formula-writing,docx-format-audit"
"$runner" "$prompts/translation-audit-tracked-changes-combo.txt" "docx-bilingual-translation,docx-format-audit,docx-tracked-changes"
"$runner" "$prompts/negative-chat-translation-no-docx-skill.txt" "-" "docx-bilingual-translation,docx-format-audit"
"$runner" "$prompts/negative-non-docx-audit-no-format-audit.txt" "-" "docx-format-audit,docx-bilingual-translation"
"$runner" "$prompts/negative-formula-explanation-no-translation.txt" "-" "docx-bilingual-translation,docx-format-audit,word-formula-writing"
"$runner" "$prompts/formula-only-no-formatting.txt" "-" "word-default-formatting,word-formula-writing,docx-tracked-changes"
"$runner" "$prompts/template-no-default-formatting.txt" "-" "word-default-formatting" 4 "not-applied"
