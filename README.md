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

Claude Code can install this repository as a plugin marketplace, matching the Superpowers installation model:

```bash
claude plugin marketplace add https://github.com/wintersdragon-c/office-docx-skills.git
claude plugin install office-docx-skills@office-docx-skills-dev
```

Restart Claude Code after installation.

For local development, point Claude Code at this checkout instead:

```bash
claude plugin marketplace add /path/to/office-docx-skills
claude plugin install office-docx-skills@office-docx-skills-dev
```

Marketplace status: this package is not currently published to the official Claude plugin marketplace. The repository's `.claude-plugin/marketplace.json` provides the installable marketplace metadata.

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

Verify DOCX smoke behavior:

```bash
python3 -m pytest tests/docx-smoke
```

Verify explicit skill triggering and multi-skill behavior in Claude Code:

```bash
tests/explicit-skill-requests/run-all.sh
```

## Updating

Codex:

```bash
cd ~/.codex/office-docx-skills
git pull
```

Claude Code:

```bash
claude plugin marketplace update office-docx-skills-dev
claude plugin update office-docx-skills
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
claude plugin uninstall office-docx-skills
claude plugin marketplace remove office-docx-skills-dev
```

## Limitations

- `docx-tracked-changes` supports body paragraph tracked changes and does not handle legacy `.doc` files.
- Word table-of-contents fields still need to be refreshed in Word.
- `word-formula-writing` provides equation workflow guidance; exact equation insertion depends on the available local Python packages.
- Claude Code installation uses this repository as a plugin marketplace; it is not published in Anthropic's official marketplace.

## License

MIT. See [LICENSE](LICENSE).

This skill was developed by Dongyao Chen, Guozhi Niu, and Jiayun Lei.
