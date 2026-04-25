---
name: docx-tracked-changes
description: Use when editing `.docx` or Word documents that need visible Track Changes, revision marks, 修订模式, insertions, deletions, reviewer-facing redlines, or preserved change history.
---

# Editing .docx with Word Track Changes (Revision Mode)

## Overview

Edit `.docx` files programmatically so that changes appear as tracked revisions in Microsoft Word — deletions show as red strikethrough, insertions as red underline, with author/date metadata in the margin. This uses direct XML manipulation of the OOXML `w:ins` / `w:del` elements via `lxml`, NOT `python-docx`'s high-level API (which has no tracked-change support).

## Dependencies

```bash
pip install python-docx lxml
```

- `python-docx` >= 1.2.0 — used only for initial inspection (paragraph styles, run formatting). NOT used for the actual tracked-change writes.
- `lxml` >= 5.0 — the real workhorse; parses and modifies `word/document.xml` inside the `.docx` zip.

Both are pure-Python installs, no system-level dependencies.

## When to Use

- User asks to edit a `.docx` and wants to SEE the changes in Word (修订模式 / Track Changes)
- User needs to send a revised manuscript where reviewers can accept/reject individual edits
- Any scenario where direct text replacement in `.docx` is insufficient because the change history matters

## When NOT to Use

- User just wants the final text replaced (no revision marks needed) — use `python-docx` directly
- File is `.doc` (old binary format) — not supported
- User wants to edit headers/footers/footnotes with tracked changes — this technique covers body paragraphs only

## Works With

This skill can be used alone or together with other skills in `office-docx-skills`.

- Use with `word-default-formatting` when revised text should also follow the package's default Word formatting profile.
- Use with `word-formula-writing` when formula-related edits should be represented with Word-native editable equations and visible revision marks.
- If both formatting and formulas matter, use all three skills together.

## OOXML Tracked-Change Structure

A `.docx` is a zip containing `word/document.xml`. Tracked changes use two wrapper elements around `w:r` (run) elements inside `w:p` (paragraph):

```xml
<!-- Deletion: wraps original runs, text tag becomes w:delText -->
<w:del w:id="102" w:author="Claude" w:date="2026-04-17T10:00:00Z">
  <w:r>
    <w:rPr><!-- original formatting preserved --></w:rPr>
    <w:delText xml:space="preserve">old text here</w:delText>
  </w:r>
</w:del>

<!-- Insertion: wraps new runs with cloned formatting -->
<w:ins w:id="103" w:author="Claude" w:date="2026-04-17T10:00:00Z">
  <w:r>
    <w:rPr><!-- cloned from original run --></w:rPr>
    <w:t xml:space="preserve">new text here</w:t>
  </w:r>
</w:ins>
```

Key rules:
- `w:del` and `w:ins` are direct children of `w:p`, siblings of `w:pPr`
- Each needs a unique `w:id` (integer, must not collide with existing IDs in the document)
- `w:author` and `w:date` populate the margin annotation in Word
- Inside `w:del`, text elements MUST be `w:delText` (not `w:t`)
- Inside `w:ins`, text elements are normal `w:t`
- Run properties (`w:rPr`) must be cloned from original runs to preserve font/size/bold

## Core Implementation

The reusable implementation lives in `tracked_change_editor.py` beside this skill. Copy that helper into your working directory when you want code you can import directly.

```python
from tracked_change_editor import TrackedChangeEditor

editor = TrackedChangeEditor('input.docx', author='Your Name')

for i, p in enumerate(editor.body_paras):
    if 'target text' in editor._get_para_text(p):
        editor.replace_paragraph_text(i, 'replacement text')
        break

editor.save('output.docx')
```

`tracked_change_editor.py` handles the important mechanics:
- scans existing revision IDs before creating new `w:ins` / `w:del` nodes
- preserves paragraph formatting by cloning the paragraph's default `w:rPr`
- writes via a safe temporary file before moving to the final `.docx`

## Usage Flow

```python
editor = TrackedChangeEditor('input.docx', author='Your Name')

# Find target paragraph by text content
for i, p in enumerate(editor.body_paras):
    if 'target text' in editor._get_para_text(p):
        editor.replace_paragraph_text(i, 'replacement text')
        break

editor.save('output.docx')  # Always save to a NEW file first
```

## Critical Gotchas — MUST READ BEFORE EDITING

### 0. ALL edits in ONE script, or chain from the last saved file

If you need to make 7 edits, do them ALL in a single script run against a single `TrackedChangeEditor` instance, then save once. Do NOT:
- Run script A → save to `_revised.docx`
- Run script B against the ORIGINAL file → save to `_revised.docx` (OVERWRITES script A's work)

If you must split across multiple script runs, each subsequent run MUST load from the previous output:
```python
# Round 1
editor = TrackedChangeEditor('original.docx')
# ... edits ...
editor.save('revised.docx')

# Round 2 — load from revised, NOT original
editor2 = TrackedChangeEditor('revised.docx')
# ... more edits ...
editor2.save('revised.docx')
```

**Why this matters:** Loading from the original file discards all tracked changes from prior rounds. This is the single most common mistake — it silently drops work with no error.

### 0b. Inserting new paragraphs: find the SECTION, not just the text

When inserting a new paragraph (e.g., a reference entry), NEVER search the entire document for a text match. Body text citations contain the same author names as reference entries. A naive search like `if 'Acharya' in text` will match a body paragraph first and insert the reference into the middle of Section 2.

**Correct approach — find the section boundary first:**
```python
# Find where References section starts
refs_start_idx = None
for i, p in enumerate(editor.body_paras):
    t = editor._get_para_text(p)
    if t.strip() == 'References':
        refs_start_idx = i
        break

# Search ONLY within References
for i in range(refs_start_idx + 1, len(editor.body_paras)):
    t = editor._get_para_text(editor.body_paras[i])
    if t.startswith('Acharya') and 'Johnson' in t:
        # Insert after this paragraph
        break
```

**Why this matters:** Author names appear in both in-text citations and reference entries. Without section-aware search, new reference paragraphs end up in the body text — a corruption that's hard to spot until the user opens Word.

### 0c. Rewriting paragraphs in academic papers: preserve ALL citations

When replacing a paragraph's text, you MUST preserve every `(Author Year)` citation from the original. Academic papers have a reference list that must match in-text citations — dropping a citation orphans the reference entry or weakens the argument.

**Correct approach — extract before rewriting:**
```python
import re
old_text = editor._get_para_text(editor.body_paras[idx])
old_cites = re.findall(r'\([^)]*\d{4}[a-z]?\)', old_text)
# Write new_text ensuring every item in old_cites appears
# After writing, verify:
for c in old_cites:
    assert c in new_text, f'LOST CITATION: {c}'
```

**Why this matters:** A rewritten introduction that drops `(Malmendier and Nagel 2011; Choi, Gao, and Jiang 2020)` silently breaks the paper's citation integrity. The reference list still contains these entries but nothing in the text points to them. Reviewers and copy-editors will flag this.

### 1. Paragraphs with existing tracked changes

If a paragraph already contains `w:ins`/`w:del` from prior edits, `_collect_direct_runs` only grabs direct `w:r` children — runs inside existing `w:ins`/`w:del` wrappers are skipped. This avoids nesting revisions but means the replacement only covers the "accepted" portion of text. For paragraphs with heavy prior revisions, accept all changes in Word first, then re-edit.

### 2. Always save to a new file

Never overwrite the original. Save as `_revised.docx`, let the user verify in Word, then replace manually.

### 3. TOC won't auto-update

`python-docx` / `lxml` cannot refresh Word's Table of Contents. After editing, the user must open in Word → right-click TOC → Update Field.

### 4. Revision ID uniqueness

The `_find_max_id()` method scans all existing `w:ins`, `w:del`, `w:rPrChange`, `w:pPrChange`, `w:sectPrChange` elements. New IDs increment from the max. If IDs collide, Word may corrupt the revision history.

### 5. Formatting preservation

`_get_default_rpr` clones the `w:rPr` from the first run in the paragraph. This works for paragraphs with uniform formatting. For paragraphs with mixed formatting (e.g., bold "Keywords:" followed by normal text), do surgical per-run replacement instead of whole-paragraph replacement.

### 6. Roundtrip fidelity

Saving via `lxml` produces minor XML cosmetic differences (quote style `"` vs `'`, line endings `\r\n` vs `\n`). Content and formatting are preserved. File size may differ due to zip compression level — this is harmless.

## Verification Checklist

After saving, open in Word and confirm:
- [ ] Deleted text shows red strikethrough
- [ ] Inserted text shows red underline
- [ ] Author name appears in margin annotations
- [ ] Surrounding paragraphs retain original formatting
- [ ] No duplicate or nested revision marks
- [ ] **Count insertions for your chosen author matches expected** — run a verification script to count `w:ins` with your author name and print each one's first 100 chars. If the count is less than expected, prior edits were lost (see Gotcha 0).
- [ ] **No reference entries in body text** — scan paragraphs before the References heading for any paragraph that looks like a full bibliographic entry (author, year, journal, volume, pages). If found, a reference was inserted in the wrong location (see Gotcha 0b).

```python
# Post-save verification snippet
target_author = 'Codex'
with zipfile.ZipFile('output.docx') as z:
    xml = z.read('word/document.xml')
root = etree.fromstring(xml)
author_ins = [e for e in root.findall('.//w:ins', nsmap)
              if e.get(qn('w:author')) == target_author]
print(f'{target_author} insertions: {len(author_ins)}')
for i, ins in enumerate(author_ins):
    texts = ''.join(t.text for t in ins.findall('.//w:t', nsmap) if t.text)
    print(f'  [{i+1}] {texts[:100]}')
```
