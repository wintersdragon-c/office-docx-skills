# Word Formula Writing Embedded Objects RED Baseline

Prompt: `formula-embedded-object-guidance.txt`

Observed missing guardrails:
- w:object
- v:imagedata
- w:drawing
- relationship IDs
- [Content_Types].xml
- per-pair formula signatures

Assistant excerpt:

```text
I'll invoke the word-formula-writing skill to get the guidance on handling formulas in Word documents.
Let me find the Word document in the workspace first.
No `.docx` files in the workspace currently. Let me check if there's a test document or if the audit script reveals what structures it inspects — that will tell us exactly what OOXML nodes and relationships the tooling expects to preserve.
Now let me also check the test files and the docx-format-audit skill definition for the full picture.
```

Raw log: `/tmp/formula-embedded-guidance-red.cPXaDZ`
