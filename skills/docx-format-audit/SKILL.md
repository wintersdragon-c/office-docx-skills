---
name: docx-format-audit
description: Use when a `.docx` or Word task requires paragraph-by-paragraph translation QA, format review, bilingual document audit, formula preservation checks, 逐段核对, 翻译审校, or validation before deleting Chinese text.
---

# DOCX Format Audit

Audit bilingual or English-only Word translation outputs before finalization. This skill checks structure and formatting; it does not perform translation by itself.

## When To Use

Use this skill when:

1. a translated Word document needs paragraph-by-paragraph review;
2. Chinese and English paragraphs must be compared;
3. table translation coverage must be checked;
4. superscripts, footnotes, formulas, or embedded formula images may have been lost;
5. Chinese text will be deleted after a bilingual review pass.

Do not use this skill for non-DOCX code, finance, or general audit tasks.

## Quick Reference

| Need | Required action |
|------|-----------------|
| Bilingual translation QA | Match `btx_<id>_src` to `btx_<id>_en`; missing pair markers are blocking |
| Table-cell review | Require source and English pair in the same `w:tc` |
| Formula preservation | Compare per-pair OMML, drawing, VML, and OLE signatures |
| English typography check | Confirm inserted English resolves to `Times New Roman` |
| English-only cleanup | Run audit first, remove source paragraphs, then audit with `--mode english-only` |

## Baseline Failures Addressed

The RED baseline capture must show at least one omitted guardrail before this skill is created. This skill closes the observed baseline gaps for DOCX audit work: deleting Chinese text without pair-ID checks, accepting missing markers, auditing only body paragraphs while missing tables, counting formulas at document level instead of per pair, or accepting English-only output that still contains Chinese text. If a future baseline covers one of these without the skill, remove that claim from this section instead of keeping untested guidance.

## Blocking Checks

Fail the audit when:

1. a Chinese source paragraph has no pair-ID-linked English translation;
2. a table-cell Chinese paragraph has no pair-ID-linked English translation in the same cell;
3. an unmarked Chinese body or table-cell paragraph remains in a bilingual output;
4. a source formula signature has no corresponding English pair signature;
5. a superscript, subscript, footnote reference, or endnote reference is missing from the English pair;
6. inserted English text does not resolve to `Times New Roman` through direct formatting or style inheritance;
7. English-only output still contains Chinese characters outside allowed names or citations;
8. the output DOCX cannot be opened by `python-docx`.

## Warning Checks

Warn on:

1. possible terminology inconsistency;
2. duplicate phrases such as `the the`;
3. large drawings that may be figures rather than formulas;
4. missing LibreOffice when headless conversion was requested.

## Pair And Formula Model

Use stable pair IDs. Missing pair markers are blocking because adjacency is not safe enough for automated audit or deletion. Formula preservation should compare per-pair signatures, including object type, relationship target, dimensions, media hash, file name, and paragraph/table-cell context.

## Runnable Helper Example

From the package root, this verifies that the audit helper is importable:

```bash
python3 - <<'PY'
import importlib.util
import sys
from pathlib import Path

audit_path = Path("skills/docx-format-audit/audit_docx_translation.py")
spec = importlib.util.spec_from_file_location("audit_docx_translation", audit_path)
assert spec is not None and spec.loader is not None
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
spec.loader.exec_module(module)
assert callable(module.main)
print("PASS: audit helper import")
PY
```

## Works With

- Use after `docx-bilingual-translation`.
- Use with `word-formula-writing` for formula-heavy documents.
- Use with `word-default-formatting` when typography and layout must be validated.
- Use before using any English-only deletion helper.

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Auditing only marked pairs | Also block unmarked Chinese body and table-cell paragraphs |
| Treating adjacency as proof | Treat missing pair IDs as blocking; add bookmarks before audit cleanup |
| Counting document-level formulas | Compare formula signatures for each source/English pair |
| Deleting Chinese before audit | Run bilingual audit first and refuse deletion on blocking findings |
| Ignoring English-only output | Re-run audit in English-only mode after removal |
