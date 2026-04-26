# Office DOCX Skills

Office DOCX Skills is a multi-skill package for LLM-assisted Microsoft Word `.docx` editing. It focuses on Word-visible tracked changes, default Word formatting, editable Word equations, bilingual Chinese-English DOCX translation, and translation-format audits.

The package keeps each capability as its own triggerable skill so agents can use one skill or combine several in the same document workflow.

## What It Provides

- visible Microsoft Word Track Changes markup for `.docx` edits
- default Word formatting for formal documents and Chinese academic papers
- Word-native editable equation guidance
- bilingual Chinese-English DOCX translation with preserved Word structure
- translation-format audits before deleting Chinese source text
- embedded image/OLE formula preservation guidance
- explicit human-trigger examples for Codex and Claude Code
- package metadata modeled after Superpowers for Codex and Claude Code

## Included Skills

```text
docx-bilingual-translation  Translate Chinese DOCX content into bilingual Chinese-English Word output.
docx-format-audit           Audit translated DOCX formatting, pairing, formulas, tables, and English-only cleanup readiness.
docx-tracked-changes        Edit DOCX files with visible Word revisions.
word-default-formatting     Apply default Word formatting and Chinese paper layout rules.
word-formula-writing        Write formulas as editable Word equation objects and preserve embedded formula objects.
```

## New DOCX Translation Skills

This package now includes two DOCX translation-focused skills:

`docx-bilingual-translation` is for Chinese `.docx` files that need paragraph-by-paragraph Chinese-English output while preserving Word structure. It guides agents to keep the original Chinese source paragraph or table cell content, insert English translations next to the corresponding source content, and preserve Word-native structures such as tables, footnotes, endnotes, superscript/subscript markers, formulas, images, embedded objects, and other non-text runs.

`docx-format-audit` is for reviewing translated DOCX files before delivery or before deleting Chinese source text. It checks Chinese-English paragraph pairing, missing or extra translations, residual Chinese in English-only output, table and cell alignment, formula and embedded object preservation, footnotes/endnotes, headings, numbering, and other structure-sensitive Word formatting issues. Its Chinese removal helper is audit-gated: deletion of Chinese source paragraphs should happen only after bilingual audit passes, and the resulting English-only document is audited again.

Use them together for the full translation workflow:

```text
1. docx-bilingual-translation creates the bilingual Chinese-English DOCX.
2. docx-format-audit checks pairing, formatting, and preserved Word structures.
3. docx-format-audit performs audit-gated Chinese source removal when an English-only final copy is required.
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

After installation, Codex may display these skills with the package namespace, such as `office-docx-skills:word-default-formatting`.

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
Use $docx-bilingual-translation to translate this Chinese DOCX into English below each paragraph.
Use $docx-format-audit to review the translated DOCX before deleting Chinese text.
```

Chinese examples:

```text
请使用 docx-tracked-changes 给这个 Word 文档保留修订痕迹。
请使用 word-default-formatting 按默认 Word 格式整理文档。
请使用 word-formula-writing 把公式写成可编辑 Word 公式。
请使用 docx-bilingual-translation 将这个 Word 文档逐段翻译成中英对照稿。
请使用 docx-format-audit 在删除中文前逐段审查翻译稿格式、表格、脚注和公式。
```

## Combining Skills

Multiple skills can be used in the same task:

```text
Bilingual DOCX translation:
  docx-bilingual-translation

Bilingual translation + default formatting:
  docx-bilingual-translation + word-default-formatting

Bilingual translation + formula preservation:
  docx-bilingual-translation + word-formula-writing

Bilingual translation + audit:
  docx-bilingual-translation + docx-format-audit

English-only cleanup after audit:
  docx-format-audit

Translation with visible revisions:
  docx-bilingual-translation + docx-tracked-changes

Full Chinese academic Word translation workflow:
  docx-bilingual-translation + word-default-formatting + word-formula-writing + docx-format-audit + docx-tracked-changes
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
claude plugin update office-docx-skills@office-docx-skills-dev
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
