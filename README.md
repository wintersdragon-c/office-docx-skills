# Office DOCX Skills

Office DOCX Skills is a multi-skill package for LLM-assisted Microsoft Word `.docx` editing. It focuses on Word-visible tracked changes, default Word formatting, and editable Word equations.

The package keeps each capability as its own triggerable skill so agents can use one skill or combine several in the same document workflow.

## What It Provides

- visible Microsoft Word Track Changes markup for `.docx` edits
- default Word formatting for formal documents and Chinese academic papers
- Word-native editable equation guidance
- explicit human-trigger examples for Codex and Claude Code
- package metadata modeled after Superpowers for Codex and Claude Code

## Included Skills

```text
docx-tracked-changes     Edit DOCX files with visible Word revisions.
word-default-formatting  Apply default Word formatting and Chinese paper layout rules.
word-formula-writing     Write formulas as editable Word equation objects.
```

## Installation

### Codex

Codex can use this package through native skill discovery:

```bash
git clone https://github.com/wintersdragon-c/office-docx-skills.git ~/.codex/office-docx-skills
mkdir -p ~/.agents/skills
ln -s ~/.codex/office-docx-skills/skills ~/.agents/skills/office-docx-skills
```

Restart Codex after installation.

Windows PowerShell:

```powershell
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.agents\skills"
cmd /c mklink /J "$env:USERPROFILE\.agents\skills\office-docx-skills" "$env:USERPROFILE\.codex\office-docx-skills\skills"
```

### Claude Code

Marketplace status: this package is not currently published to the official Claude plugin marketplace or to a public third-party marketplace. Do not use a package-specific Claude marketplace install command unless a real marketplace channel is published later.

Current manual installation:

```bash
git clone https://github.com/wintersdragon-c/office-docx-skills.git ~/.claude/office-docx-skills
mkdir -p ~/.claude/skills
ln -s ~/.claude/office-docx-skills/skills ~/.claude/skills/office-docx-skills
```

Restart Claude Code after installation.

The `.claude-plugin/marketplace.json` file is included for local development and future publication metadata. It is not a claim that the package is already available from a public marketplace.

## Explicit Skill Triggering

Ask the agent to use a named skill directly:

```text
Use $docx-tracked-changes to edit this DOCX with visible Word revisions.
Use $word-default-formatting before finalizing this Word document.
Use $word-formula-writing to convert formulas into editable Word equations.
```

Chinese examples:

```text
请使用 docx-tracked-changes 给这个 Word 文档保留修订痕迹。
请使用 word-default-formatting 按默认 Word 格式整理文档。
请使用 word-formula-writing 把公式写成可编辑 Word 公式。
```

## Combining Skills

Multiple skills can be used in the same task:

```text
Visible revisions only:
  docx-tracked-changes

Default Word formatting:
  word-default-formatting

Editable Word formulas:
  word-formula-writing

Default formatting + formulas:
  word-default-formatting + word-formula-writing

Default formatting + visible revisions:
  word-default-formatting + docx-tracked-changes

Formulas + visible revisions:
  word-formula-writing + docx-tracked-changes

Chinese academic paper or formal document + formulas + revisions:
  word-default-formatting + word-formula-writing + docx-tracked-changes
```

Example prompt:

```text
请同时使用 word-default-formatting、word-formula-writing 和 docx-tracked-changes 处理这个文档。
```

## Dependencies

Tracked changes:

```bash
python3 -m pip install python-docx lxml
```

Formula-heavy workflows may also need:

```bash
python3 -m pip install math2docx latex2mathml mathml2omml
```

## Verification

Verify package structure:

```bash
python3 tests/skill-structure/validate_package.py
```

Verify tracked-change DOCX smoke behavior:

```bash
python3 tests/docx-smoke/test_tracked_changes.py
```

## Updating

Codex:

```bash
cd ~/.codex/office-docx-skills
git pull
```

Claude Code:

```bash
cd ~/.claude/office-docx-skills
git pull
```

Restart the agent after updating.

## Uninstalling

Codex:

```bash
rm ~/.agents/skills/office-docx-skills
rm -rf ~/.codex/office-docx-skills
```

Claude Code:

```bash
rm ~/.claude/skills/office-docx-skills
rm -rf ~/.claude/office-docx-skills
```

## Limitations

- `docx-tracked-changes` supports body paragraph tracked changes and does not handle legacy `.doc` files.
- Word table-of-contents fields still need to be refreshed in Word.
- `word-formula-writing` provides equation workflow guidance; exact equation insertion depends on the available local Python packages.
- Claude Code marketplace installation is not documented until a real marketplace target exists.

## License

MIT. See [LICENSE](LICENSE).
