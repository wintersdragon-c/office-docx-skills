---
name: word-formula-writing
description: Use when a `.docx` or Word task needs editable equations, Word-native formulas, OMML, Cambria Math, formula line breaking, or converting plain-text formulas into editable Word equation objects.
---

# Word Formula Writing

Write formulas in Word as native editable equation objects. This skill is for equation production inside `.docx`, not for general document layout by itself.

## When To Use

Use this skill when:

1. a `.docx` document must contain editable formulas;
2. the user asks for Word-native equations, OMML output, or conversion from plain-text formulas;
3. a Word document needs equation formatting, equation line breaking, or formula-style cleanup;
4. formulas must follow the default math presentation standard in this workspace.

Do not use this skill alone for:

1. ordinary Word editing without formulas;
2. LaTeX-only outputs;
3. PDF-only math typesetting when Word editing is irrelevant.

## Default Equation Standard

### Equation Object

1. use Word native editable equations;
2. do not leave formulas as plain text such as `a^2`, `x_t`, or slash-style fractions when native equations are required;
3. target OMML-compatible equation objects.

### Equation Font and Layout

1. math font: `Cambria Math`
2. formula paragraphs use a dedicated equation style
3. long formulas should be split manually
4. do not rely on Word auto-wrap to rescue overlong equations
5. when a heading immediately follows a display formula, add extra top spacing to that heading so the formula block and the next section do not visually stick together

### Formula Context

1. introduce formulas with a short lead-in sentence when needed
2. explain symbols at first appearance when the user expects a formal document
3. keep variable explanations concise and embedded naturally in the text

## Toolchain

Use this toolchain when formula generation is required:

1. `python-docx` for document structure, paragraphs, and styles
2. `math2docx` for inserting Word-native equations when available
3. `latex2mathml` and `mathml2omml` as backup conversion tools when needed
4. `lxml` for XML-level inspection and validation

Install missing packages in the active project environment before generating formula-heavy DOCX files:

```bash
python3 -m pip install python-docx lxml math2docx latex2mathml mathml2omml
```

## Workflow

### 1. Confirm that Word-native formulas are required

Check whether the user wants:

1. editable Word equations
2. equation formatting inside `.docx`
3. replacement of fake text formulas

If the user only wants explanation in chat, this skill is unnecessary.

### 2. Build the document around equation paragraphs

Separate:

1. body paragraphs
2. formula paragraphs
3. heading paragraphs

Do not merge equation blocks into ordinary body style.

### 3. Convert formulas through the Word equation path

Preferred path:

1. prepare the formula in LaTeX-like syntax
2. insert it into the document with `math2docx`
3. verify that the resulting `.docx` contains equation objects rather than text substitutes

If needed, use the fallback chain:

1. LaTeX
2. MathML
3. OMML

### 4. Split long formulas manually

When a formula risks exceeding the page width:

1. break it into multiple display equations
2. separate definitions from the main equation
3. break first-order conditions or derivations into steps
4. avoid a single overlong equation line

### 5. Keep equation styling stable

By default:

1. use `Cambria Math` for formulas
2. use compact line spacing for formula paragraphs
3. keep formulas visually distinct from body text
4. align with the surrounding Word style rather than making formulas visually oversized
5. if a heading follows directly after a formula, prefer a local heading-spacing override instead of inflating the formula paragraph itself
6. a good default starting point is roughly `10 pt` of heading `space_before` after a display formula, with small page-level adjustment when needed

## Validation

After generating the document, check at least these points:

1. equations are editable in Word
2. equations do not overflow the text area
3. long equations are split cleanly
4. body text and formulas use different style layers
5. the document XML contains math objects rather than only literal formula text
6. if a heading follows a formula, the gap is visibly comfortable and does not read like one merged block

## Relationship To Other Skills

This skill can be used alone or together with other skills in `office-docx-skills`.

- Use with `word-default-formatting` when equations must match the document's default Word formatting profile.
- Use with `docx-tracked-changes` when formula-related edits should appear as Word-visible revisions.
- If the user's environment has a separate general DOCX production skill, it may be used as an optional companion, but it is not required by this package.

## Calling Rules

Use the following routing logic.

1. Word document without formulas:
   do not use this skill unless editable equations are required
2. Word document with formulas:
   use `word-formula-writing`
3. Word document with formulas and no custom style:
   use `word-default-formatting` + `word-formula-writing`
4. Formula discussion outside Word:
   do not use this skill unless the user explicitly wants `.docx` equation output
