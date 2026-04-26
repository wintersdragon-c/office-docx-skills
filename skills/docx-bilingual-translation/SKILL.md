---
name: docx-bilingual-translation
description: Use when a `.docx` or Word task requires Chinese-to-English translation, bilingual paragraph insertion, table translation, 中英对照, 逐段翻译, or preserving Word formatting while adding English below Chinese source text.
---

# DOCX Bilingual Translation

Translate Chinese Word documents into English while preserving DOCX structure. This skill owns bilingual insertion and structural preservation; it does not guarantee final prose quality by itself.

## When To Use

Use this skill when:

1. the input or output is `.docx` or Microsoft Word;
2. Chinese text must be translated into English;
3. English translations should be inserted below Chinese source paragraphs;
4. Chinese table-cell content must be translated without changing table structure;
5. formatting, formulas, footnotes, superscripts, or embedded objects must survive translation.

Do not use this skill for ordinary chat translation outside Word.

## Required Output Pattern

1. Body paragraphs: insert the English paragraph immediately after the Chinese source paragraph.
2. Table cells: keep the same cell, row, column, and merge structure; insert the English paragraph immediately after the Chinese paragraph inside the cell.
3. Do not put Chinese and English into the same paragraph separated by a line break.
4. Do not add a new translation column unless the user explicitly asks for a side-by-side table.
5. Mark each Chinese/English pair with stable pair IDs.

## Quick Reference

| Need | Required action |
|------|-----------------|
| Translate Chinese body paragraph | Insert English in a new `w:p` immediately after source |
| Translate Chinese in a table cell | Insert English in the same `w:tc`; never restructure rows or columns |
| Preserve formulas | Copy OMML, drawing, VML, and OLE-bearing runs with relationship rewrites |
| Prepare deletion of Chinese text | Run `docx-format-audit` first and delete only pair-ID source paragraphs |
| Combine with tracked changes | Use `docx-tracked-changes` for reviewer-visible insertions/deletions |

## Baseline Failures Addressed

The RED baseline capture must show at least one omitted guardrail before this skill is created. This skill closes the observed baseline gaps for DOCX translation work: missing stable pair markers, unsafe table handling, weak formula/object preservation, or deleting Chinese source text before audit. If a future baseline covers one of these without the skill, remove that claim from this section instead of keeping untested guidance.

## Pair Identity

Do not rely on adjacency for automated audit or deletion decisions. Add invisible bookmark markers:

- `btx_<id>_src` on the Chinese source paragraph
- `btx_<id>_en` on the English translation paragraph

Audit and English-only deletion require pair IDs. Physical adjacency is useful for human diagnosis, but helpers treat missing pair markers as blocking because deletion by adjacency is unsafe.

## Formatting Rules

1. Clone paragraph properties (`w:pPr`) from the source paragraph.
2. Clone relevant run properties before changing inserted English text.
3. Set inserted English text to `Times New Roman` for `w:ascii`, `w:hAnsi`, and `w:cs`.
4. Set `w:eastAsia="Times New Roman"` only for English-only inserted runs when portability requires it.
5. Preserve superscript, subscript, footnote references, endnote references, and author markers.

## Formula And Object Preservation

Preserve formulas and formula-like inline objects:

- `m:oMath`
- `m:oMathPara`
- `w:object`
- `w:drawing`
- `v:shape`
- `v:imagedata`

When cloning image/OLE formulas across parts or documents, copy related media or embedding targets, update the destination `.rels`, generate a new `rId` if needed, rewrite `r:embed` or `r:id`, and update `[Content_Types].xml` when a copied target introduces a new content type.

## Runnable Helper Example

From the package root, this verifies that the helper is importable and that a document paragraph should enter the DOCX bilingual workflow:

```bash
python3 - <<'PY'
import importlib.util
import sys
from pathlib import Path

helper_path = Path("skills/docx-bilingual-translation/translation_docx_helpers.py")
spec = importlib.util.spec_from_file_location("translation_docx_helpers", helper_path)
assert spec is not None and spec.loader is not None
module = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = module
spec.loader.exec_module(module)
assert module.has_chinese("这是中文段落。")
assert not module.has_chinese("English only.")
print("PASS: bilingual helper import and Chinese detection")
PY
```

## Workflow

1. Work on a copy unless the user explicitly allows direct editing.
2. Inspect document paragraphs, tables, relationships, and formula objects before editing.
3. Insert English paragraphs with cloned paragraph formatting and copied formula objects when the source paragraph contains formula candidates.
4. Mark every inserted source/translation pair.
5. Save a bilingual review document.
6. Use `docx-format-audit` before producing an English-only document.

## Works With

- Use with `word-default-formatting` when default Word typography and layout matter.
- Use with `word-formula-writing` when formulas or image/OLE formulas appear.
- Use with `docx-format-audit` before deleting Chinese text.
- Use with `docx-tracked-changes` when insertions or deletions must appear as Word-visible revisions.

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Translating in chat only | Produce or edit a `.docx` when the user asks for Word output |
| Inserting English with a line break in the Chinese paragraph | Create a separate paragraph below the source paragraph |
| Moving table translations outside the table | Insert translations inside the same cell |
| Counting formulas only at document level | Compare per-pair formula signatures |
| Removing Chinese immediately | Audit first, then remove only source paragraphs that have pair IDs |
| Relying on paragraph adjacency for deletion | Treat missing pair markers as blocking and add bookmarks before cleanup |
