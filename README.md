# docx-tracked-changes

`docx-tracked-changes` is an agent skill for editing `.docx` files while preserving **Microsoft Word Track Changes** markup. It is designed for cases where the user must see insertions, deletions, author names, and timestamps inside Word instead of receiving a silently rewritten document.

## Why This Exists

`python-docx` does not expose an API for tracked revisions. This repository documents a working OOXML-based approach that writes `w:ins` and `w:del` elements directly, plus the failure modes that matter in real review workflows.

This is useful for:
- academic paper revisions
- legal or policy document redlines
- reviewer-facing manuscript updates
- any agent workflow where visible revision history is required

## Repository Layout

```text
docx-tracked-changes/
  SKILL.md                   Main skill reference
  tracked_change_editor.py   Reusable Python helper extracted from the skill
LICENSE
README.md
```

## What The Skill Covers

- when to use tracked-change editing instead of plain text replacement
- the OOXML structure for `w:ins`, `w:del`, and `w:delText`
- a reusable `TrackedChangeEditor` helper
- operational gotchas from real manuscript revision work
- verification steps before handing the `.docx` back to a user

## Install As A Skill

### Codex

Copy the skill folder into `~/.agents/skills/`:

```bash
cp -R docx-tracked-changes ~/.agents/skills/
```

### Claude Code

Copy the skill folder into `~/.claude/skills/`:

```bash
cp -R docx-tracked-changes ~/.claude/skills/
```

## Use The Python Helper

Install dependencies:

```bash
python3 -m pip install python-docx lxml
```

Copy the helper into the directory where your editing script lives:

```bash
cp docx-tracked-changes/tracked_change_editor.py /path/to/your/project/
```

Example:

```python
from tracked_change_editor import TrackedChangeEditor

editor = TrackedChangeEditor("manuscript.docx", author="Your Name")

for i, paragraph in enumerate(editor.body_paras):
    if "old text" in editor._get_para_text(paragraph):
        editor.replace_paragraph_text(i, "new text")
        break

editor.save("manuscript_revised.docx")
```

Open the output in Word and confirm the insertion/deletion marks are visible.

## Limitations

- supports body-paragraph edits only
- does not handle legacy `.doc` files
- does not refresh Word table-of-contents fields
- whole-paragraph replacement is safest when formatting is uniform
- paragraphs with existing tracked revisions may need cleanup first

## Review Notes

This repository was cleaned up for open-source publication with three important changes:
- the reusable Python implementation now exists as a real file instead of only a code block inside `SKILL.md`
- the helper now writes through a safe temporary file rather than `tempfile.mktemp`
- the README quick start now matches the actual repository contents

## License

MIT
