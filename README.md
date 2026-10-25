# Tool Audit & Roadmap Toolkit

A reusable framework for auditing a company's software subscriptions and manual processes and
deciding what to replace or build first. Not code to deploy: a method, a scored workbook, and a
worked example.

| File | Purpose |
|------|---------|
| [guide.md](guide.md) | The five-step method and interview questions (EN, with FR summary) |
| `tool-audit-template.xlsx` | Blank inventory with scoring, ranking and checklist sheets |
| `tool-audit-example.xlsx` | The same workbook filled with a fictional company |
| [decision-checklist.md](decision-checklist.md) | Keep vs replace checks, including critical-system sign-off |
| [worked-example.md](worked-example.md) | The fictional company, read through |
| [proposed-approach.md](proposed-approach.md) | One-page week 1 / 2 / 4 plan (draft application text) |

**Score = value x confidence / (effort + risk).** A critical system is never marked for
replacement, however it scores. Effort and risk of 0 show `CHECK INPUT` instead of a number.

## Regenerate and test

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python build_workbook.py     # rewrites the workbooks and generated pages
.venv/bin/pytest                       # recalculation tests need LibreOffice (skipped without)
```

`decision-checklist.md` and `worked-example.md` are generated from `build_workbook.py`; edit
the data there, not the markdown. The example company is fictional and illustrative.
