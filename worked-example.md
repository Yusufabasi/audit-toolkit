# Worked example: Soleil Grid (fictional)

A made-up solar mini-grid operator with ten tools. Nothing here describes a real company. Generated from `build_workbook.py`, the same data as `tool-audit-example.xlsx`.

| Rank | Tool | Cost/mo (USD) | V | C | E | R | Score | Decision |
|------|------|---------------|---|---|---|---|-------|----------|
| 1 | Weekly KPI report (hand-built in spreadsheets) | 0 | 5 | 0.9 | 2 | 1 | 1.50 | Replace / automate candidate |
| 2 | Forms / survey SaaS | 120 | 3 | 0.9 | 1 | 1 | 1.35 | Replace / automate candidate |
| 3 | Site monitoring portal (vendor) | 350 | 5 | 1.0 | 2 | 2 | 1.25 | Do not replace without sign-off |
| 4 | Mobile-money reconciliation (spreadsheets) | 0 | 5 | 0.8 | 3 | 3 | 0.67 | Review |
| 5 | Helpdesk SaaS | 450 | 4 | 0.7 | 3 | 2 | 0.56 | Review |
| 6 | Field work-order app (SaaS) | 380 | 4 | 0.7 | 4 | 3 | 0.40 | Keep |
| 7 | E-signature SaaS | 90 | 2 | 0.9 | 2 | 3 | 0.36 | Keep |
| 8 | Project tracker SaaS | 200 | 2 | 0.8 | 3 | 2 | 0.32 | Keep |
| 9 | Email and office suite | 1500 | 3 | 0.9 | 5 | 5 | 0.27 | Keep |
| 10 | Accounting software | 800 | 4 | 0.6 | 5 | 5 | 0.24 | Do not replace without sign-off |

Total monthly spend 3890 USD; spend on replace candidates 120 USD.

## Reading the result

- **The best first project is not the most expensive tool.** The hand-built weekly report costs nothing in licences but eats about 12 hours a week, so it ranks first: high value, low effort, low risk. Licence savings understate the gain; staff time is usually the bigger number.
- **The forms SaaS is the only licence worth replacing now.** It is small, safe and quick, so it makes a good early win.
- **A high score does not override a critical system.** The vendor monitoring portal scores 1.25 because its CSV export looks easy to copy, but it is the read path for site operations. The sheet forces `Do not replace without sign-off`; the decision goes to the IT Senior Manager and, for anything touching money, the FP&A Lead.
- **Expensive is not the same as replaceable.** The email suite costs the most and scores lowest because it is the identity backbone: effort 5, risk 5.
- **Next:** run the decision checklist on the two candidates, pilot in parallel with a rollback, then revisit the middle of the table (mobile-money reconciliation, helpdesk) once the first tool is stable.

## Resume (FR)

Le meilleur premier projet n'est pas l'outil le plus cher: le rapport hebdomadaire fait a la main ne coute aucune licence mais prend environ 12 heures par semaine. Un score eleve ne leve jamais la regle sur les systemes critiques: le portail de supervision reste soumis a validation.
