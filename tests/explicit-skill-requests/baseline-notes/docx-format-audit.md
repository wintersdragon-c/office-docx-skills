# DOCX Format Audit RED Baseline

Prompt: `baseline-docx-format-audit-no-skill.txt`

Observed missing guardrails:
- pair-ID first audit
- missing marker is blocking
- same-cell table pairing
- per-pair formula signatures
- English-only residual Chinese check

Assistant excerpt:

```text
I'll check the bilingual translation skill first — it's directly relevant to working with Chinese-English .docx files.
The bilingual translation skill requires running `docx-format-audit` before deleting Chinese text. Let me explore what's available in the codebase and find the target .docx file.
```

Raw log: `/tmp/docx-format-audit-baseline.7DmS7R`
