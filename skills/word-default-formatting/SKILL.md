---
name: word-default-formatting
description: Use when creating or editing `.docx` or Word documents that explicitly need the default/house Word format, 默认格式, Chinese academic paper layout, mixed Chinese-English fonts, headings, spacing, margins, or formal style cleanup, and no supplied template or custom style overrides it.
---

# Word Default Formatting

Apply the default Word formatting standard for this workspace when the user wants a Word document but has not provided a different style guide. This skill has two layers:

1. a general default layer for formal Word documents;
2. a Chinese paper profile for academic papers.

It does not replace the general DOCX workflow and it does not specialize in equation writing.

## When To Use

Use this skill when:

1. the task is to create or edit a `.docx` file;
2. the user explicitly expects the usual house style, says to use the default Word format, or gives no competing template/style requirement;
3. the user wants consistent Chinese/English typography, heading hierarchy, paragraph spacing, and page layout;
4. the task may later be extended to other format profiles such as paper style or referee-response style, but currently should use the default profile.

Do not use this skill alone for:

1. non-Word outputs such as PDF-only, LaTeX-only, or plain text;
2. tasks centered on editable Word equations; for those, also use `word-formula-writing`;
3. cases where the user has explicitly provided another template, journal style, school format, or custom layout requirement that overrides the default;
4. requests that mention the default profile only to reject it, avoid it, or compare it against a provided template.

## General Default Layer

Unless the user states otherwise, apply the following rules to formal Word documents.

### Typography

1. Chinese text: `宋体`
2. English text and Arabic numerals: `Times New Roman`
3. Word equations: `Cambria Math`
4. default font color: black unless the user or template explicitly requires another color

### Body Text

1. body size: `12 pt` (`小四`)
2. line spacing: `1.5`
3. first-line indent: `2` Chinese characters
4. default alignment: justified

### Structure

1. split title, headings, body, and formula paragraphs into separate styles
2. do not force the whole document into one `Normal` style
3. keep formula paragraphs separate from body paragraphs

### Headings

1. document title: centered, bold, larger than body text
2. first-level headings: left aligned, bold, no first-line indent
3. second-level headings: left aligned, bold, no first-line indent
4. third-level headings: left aligned, bold, no first-line indent
5. when a heading immediately follows a display formula or a table, increase the heading's `space_before` locally instead of loosening heading spacing everywhere
6. as a practical default, start from about `10 pt` before a heading after a formula and about `12 pt` before a heading after a table, then adjust slightly if the page looks too tight or too loose

### Formula Placement

1. short or medium formulas may appear as standalone display equations
2. long formulas should be manually split instead of relying on Word auto-wrap
3. formula paragraphs should use a dedicated paragraph style rather than body style
4. if a display formula is followed immediately by a heading, leave a visibly comfortable gap so the heading does not look glued to the formula block

### Page Setup

1. preserve the user's template if one is provided
2. otherwise use A4 and a formal academic/professional page layout
3. keep the visual hierarchy stable across title, headings, body, and formula blocks

## Chinese Paper Profile

When the user is writing a Chinese academic paper and has not given another journal or school template, use this profile by default.

### Page Setup

1. paper size: `A4`
2. margins: top `2.54 cm`, bottom `2.54 cm`, left `3.18 cm`, right `3.18 cm`
3. default layout: single column

### Fonts

1. Chinese body text: `宋体`
2. English text, Arabic numerals, variable names, and reference metadata: `Times New Roman`
3. Word equations: `Cambria Math`
4. abstract and keywords body may use `仿宋`
5. abstract and keywords labels may use `黑体`
6. default text color remains black unless an external template or explicit instruction overrides it

### Title And Front Matter

1. paper title: centered, bold, `16 pt`
2. author line: centered, usually `12 pt`
3. abstract label: `【摘要】`, bold
4. abstract body: `12 pt`, formal academic prose
5. keywords label: `【关键词】`, bold
6. keywords body: `12 pt`

### Body Text

1. body size: `12 pt`
2. alignment: justified
3. first-line indent: `2` Chinese characters
4. line spacing: `1.5`
5. paragraph spacing: `0` before and after unless a template overrides it

### Heading Hierarchy

1. first-level heading: centered, bold, `12 pt`, typically `一、二、三……`
2. second-level heading: left aligned, `12 pt`, typically `（一）（二）（三）……`
3. third-level heading: left aligned, `12 pt`, typically `1. 2. 3.` or short inline labels
4. headings do not use first-line indent

### Figures And Tables

1. figures and tables should be centered
2. figure captions use the form `图1：……`
3. table captions use the form `表1：……`
4. captions should be visually separated from surrounding body text without excessive blank space
5. if a heading comes immediately after a table, widen the heading's top spacing slightly more than in ordinary heading transitions so the table block and the next section remain visually distinct

### Equations

1. equations must remain editable Word equations
2. use `Cambria Math`
3. display equations should be placed in dedicated equation paragraphs
4. long equations should be manually split
5. if equations are present, also use `word-formula-writing`

### References

1. section title: `参考文献`
2. reference title is a standalone heading, preferably centered, bold, `12 pt`
3. reference entries use `12 pt`
4. line spacing: `1.5`
5. do not force first-line indent for reference entries by default
6. keep mixed Chinese-English font handling inside each reference entry
7. prefer numbered references such as `[1] [2] [3]` unless the user or target journal requires another style

### Notes

1. this profile extends the general default layer; it does not replace it
2. this is a default Chinese paper profile, not a journal-specific template
3. if the user provides a journal, school, or conference format, that external format overrides this profile
4. if the document contains formulas, combine this profile with `word-formula-writing`

## Workflow

### 1. Determine whether the default profile applies

Check the user request and source files first.

1. If the user provided a template, inherit the template unless it conflicts with an explicit instruction.
2. If the user named another style standard, do not apply this default profile blindly.
3. If the user explicitly wants a Chinese paper and there is no override, apply the general default layer and then activate the `Chinese Paper Profile`.
4. If there is no override and the task is not clearly a paper, use the general default layer only.

### 2. Split the document into style layers

At minimum, separate the document into:

1. title
2. heading levels
3. body paragraphs
4. formula paragraphs, if present

Do not force the whole document into a single `Normal` style and patch individual paragraphs ad hoc.

### 3. Apply mixed Chinese-English font rules carefully

When writing runs or paragraph styles:

1. set western fonts to `Times New Roman`
2. set East Asian fonts to `宋体`
3. keep math font separate from body font
4. set the default text color to black unless the source template explicitly requires another color

The default requirement is mixed typography, not a document-wide single font.

Use `formatting_helpers.py` beside this skill when applying fonts through `python-docx`; setting only `run.font.name` is not enough for Chinese text because Word also needs `w:eastAsia`.

```python
from formatting_helpers import set_run_mixed_fonts, set_style_mixed_fonts

set_style_mixed_fonts(document.styles["Normal"])
run = paragraph.add_run("中文 ABC 123")
set_run_mixed_fonts(run)
```

After generating a document, inspect `word/document.xml` or `word/styles.xml` and confirm `w:rFonts` has `w:ascii`, `w:hAnsi`, and `w:eastAsia`.

```bash
python3 tests/docx-smoke/test_default_formatting.py
```

### 4. Preserve structure before decoration

Prioritize:

1. correct section order
2. clean paragraph hierarchy
3. stable indentation and line spacing
4. font consistency
5. local visual refinements
6. use local spacing overrides for special transitions such as `formula -> heading` and `table -> heading` instead of changing the global heading style unless the whole document needs a looser rhythm

## Relationship To Other Skills

This skill can be used alone or together with other skills in `office-docx-skills`.

- Use with `word-formula-writing` when a Word document needs both default formatting and editable formulas.
- Use with `docx-tracked-changes` when edits should appear as Word-visible revisions.
- If the user's environment has a separate general DOCX production skill, it may be used as an optional companion, but it is not required by this package.

## Calling Rules

Use the following routing logic.

1. General `.docx` creation or editing with no special format request and no provided template:
   use `word-default-formatting`, and apply the general default layer
2. `.docx` creation or editing with formulas:
   use `word-default-formatting` + `word-formula-writing`, and apply the general default layer
3. Chinese academic paper with no journal-specific template:
   use `word-default-formatting`, apply the general default layer, and then activate the `Chinese Paper Profile`
4. Chinese academic paper with formulas and no journal-specific template:
   use `word-default-formatting` + `word-formula-writing`, apply the general default layer, and then activate the `Chinese Paper Profile`
5. `.docx` creation or editing with a user-provided non-default template or journal style:
   do not use this skill unless the user explicitly asks to apply, compare, or merge the package default profile
6. Formula-only question without Word output:
   do not use this skill

## Future Extension

This skill should stay layered. Keep the general default layer short and stable. Add document-specific profiles separately. If the user later needs other profiles such as:

1. paper submission format
2. referee response format
3. school application format
4. course handout format

add them as explicit alternative profiles rather than diluting the default profile.
