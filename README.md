# docx-tracked-changes

Programmatically edit `.docx` files with **Word Track Changes** (修订模式) — so reviewers see red strikethrough for deletions and red underline for insertions, with author/date metadata in the margin.

## Why

`python-docx` has no API for tracked changes. This skill uses direct OOXML XML manipulation via `lxml` to write proper `w:ins` / `w:del` revision markup that Microsoft Word renders natively.

Built and battle-tested during a real academic paper revision workflow (climate finance literature review under peer review).

## What's Inside

```
skills/docx-tracked-changes/
  SKILL.md    # Complete reference: XML structure, Python implementation,
              # usage patterns, and critical gotchas from real-world bugs
```

The `SKILL.md` contains:

- **Dependencies** — `python-docx` + `lxml`, pure Python
- **OOXML structure** — exact XML format for `w:ins` / `w:del` elements
- **`TrackedChangeEditor` class** — ready-to-use Python implementation
- **Critical Gotchas** — 9 documented pitfalls, including:
  - Multi-step edits silently dropping prior changes
  - Reference entries inserted into body text instead of References section
  - Academic citations lost when rewriting paragraphs
  - Paragraphs with existing tracked changes
  - Mixed formatting preservation

## Quick Start

```bash
pip install python-docx lxml
```

```python
from tracked_change_editor import TrackedChangeEditor  # or copy class from SKILL.md

editor = TrackedChangeEditor('manuscript.docx')

# Replace a paragraph with tracked changes
for i, p in enumerate(editor.body_paras):
    if 'old text' in editor._get_para_text(p):
        editor.replace_paragraph_text(i, 'new text')
        break

editor.save('manuscript_revised.docx')  # Always save to a NEW file
```

Open `manuscript_revised.docx` in Word — you'll see the revision marks.

## Designed for AI Agents

This is a [Claude Code skill](https://docs.anthropic.com/en/docs/claude-code) — a reference document that teaches AI coding agents how to perform tracked-change editing correctly. The gotchas section is specifically written to prevent mistakes that AI agents commonly make when manipulating `.docx` files programmatically.

To use as a Claude Code skill, copy `skills/docx-tracked-changes/` into `~/.claude/skills/`.

## Requirements

- Python 3.8+
- `lxml` >= 5.0
- `python-docx` >= 1.2.0 (for inspection only)
- Microsoft Word or compatible editor (to view tracked changes)

## License

MIT
