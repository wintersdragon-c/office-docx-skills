---
name: docx-tracked-changes
description: Use when editing .docx files and the user wants changes visible in Word's Track Changes / revision mode (red strikethrough for deletions, red underline for insertions, author and timestamp in margin). Also use when user says "修订模式".
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

```python
import zipfile, shutil, copy, tempfile
from lxml import etree
from datetime import datetime, timezone

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
nsmap = {'w': W}

def qn(tag):
    return f'{{{W}}}{tag.split(":")[-1]}'

class TrackedChangeEditor:
    def __init__(self, docx_path):
        self.docx_path = docx_path
        with zipfile.ZipFile(docx_path, 'r') as z:
            self.doc_xml = z.read('word/document.xml')
        self.root = etree.fromstring(self.doc_xml)
        self.body = self.root.find('w:body', nsmap)
        self._find_max_id()
        self._index_body_paragraphs()
        self.author = 'Claude'
        self.date = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

    def _find_max_id(self):
        self.max_id = 0
        for tag in ['ins', 'del', 'rPrChange', 'pPrChange', 'sectPrChange']:
            for el in self.root.findall(f'.//w:{tag}', nsmap):
                rid = el.get(qn('w:id'))
                if rid:
                    try:
                        v = int(rid)
                        if v > self.max_id:
                            self.max_id = v
                    except ValueError:
                        pass

    def _next_id(self):
        self.max_id += 1
        return str(self.max_id)

    def _index_body_paragraphs(self):
        """Index paragraphs after the TOC (w:sdt block)."""
        self.body_paras = []
        past_toc = False
        for child in self.body:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'sdt':
                past_toc = True
                continue
            if past_toc and tag == 'p':
                self.body_paras.append(child)
        # Fallback: if no TOC found, index all paragraphs
        if not self.body_paras:
            self.body_paras = self.body.findall('w:p', nsmap)

    def _get_para_text(self, p):
        return ''.join(t.text for t in p.findall('.//w:t', nsmap) if t.text)

    def _get_default_rpr(self, p):
        for r in p.findall('.//w:r', nsmap):
            rPr = r.find('w:rPr', nsmap)
            if rPr is not None:
                return copy.deepcopy(rPr)
        return None

    def _collect_direct_runs(self, p):
        return [ch for ch in p if ch.tag == qn('w:r')]

    def replace_paragraph_text(self, para_index, new_text):
        """Replace all text in a body paragraph using tracked changes."""
        p = self.body_paras[para_index]
        direct_runs = self._collect_direct_runs(p)
        if not direct_runs:
            return
        default_rpr = self._get_default_rpr(p)
        first_run_pos = list(p).index(direct_runs[0])

        # Wrap original runs in w:del
        del_elem = etree.Element(qn('w:del'))
        del_elem.set(qn('w:id'), self._next_id())
        del_elem.set(qn('w:author'), self.author)
        del_elem.set(qn('w:date'), self.date)
        for r in direct_runs:
            for t in r.findall('w:t', nsmap):
                t.tag = qn('w:delText')
            p.remove(r)
            del_elem.append(r)

        # Create w:ins with new text
        ins_elem = etree.Element(qn('w:ins'))
        ins_elem.set(qn('w:id'), self._next_id())
        ins_elem.set(qn('w:author'), self.author)
        ins_elem.set(qn('w:date'), self.date)
        new_run = etree.SubElement(ins_elem, qn('w:r'))
        if default_rpr is not None:
            new_run.append(copy.deepcopy(default_rpr))
        new_t = etree.SubElement(new_run, qn('w:t'))
        new_t.set(XML_SPACE, 'preserve')
        new_t.text = new_text

        p.insert(first_run_pos, del_elem)
        p.insert(first_run_pos + 1, ins_elem)

    def save(self, output_path):
        doc_xml_new = etree.tostring(
            self.root, xml_declaration=True, encoding='UTF-8', standalone=True
        )
        tmp = tempfile.mktemp(suffix='.docx')
        with zipfile.ZipFile(self.docx_path, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename == 'word/document.xml':
                        zout.writestr(item, doc_xml_new)
                    else:
                        zout.writestr(item, zin.read(item.filename))
        shutil.move(tmp, output_path)
```

## Usage Flow

```python
editor = TrackedChangeEditor('input.docx')

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
- [ ] **Count Claude insertions matches expected** — run a verification script to count `w:ins` with `w:author="Claude"` and print each one's first 100 chars. If the count is less than expected, prior edits were lost (see Gotcha 0).
- [ ] **No reference entries in body text** — scan paragraphs before the References heading for any paragraph that looks like a full bibliographic entry (author, year, journal, volume, pages). If found, a reference was inserted in the wrong location (see Gotcha 0b).

```python
# Post-save verification snippet
with zipfile.ZipFile('output.docx') as z:
    xml = z.read('word/document.xml')
root = etree.fromstring(xml)
claude_ins = [e for e in root.findall('.//w:ins', nsmap)
              if e.get(qn('w:author')) == 'Claude']
print(f'Claude insertions: {len(claude_ins)}')
for i, ins in enumerate(claude_ins):
    texts = ''.join(t.text for t in ins.findall('.//w:t', nsmap) if t.text)
    print(f'  [{i+1}] {texts[:100]}')
```
