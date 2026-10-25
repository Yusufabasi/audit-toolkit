# Keep vs replace: decision checklist

Generated from `build_workbook.py` so the workbook and this page never disagree.
Answer every item for each replace candidate before any build starts.

| ID | Question | Passes when |
|----|----------|-------------|
| C1 | Is it a critical system (site control, metering, billing, finance books)? *(Critique ? Ne pas remplacer sans validation.)* | If yes: no replacement without IT Senior Manager and FP&A Lead sign-off; access read-only by default. |
| C2 | Can all data be exported in an open format (CSV, JSON, SQL)? *(Export complet possible ?)* | A full export has been tested, not just promised. |
| C3 | Are the accounts and licences owned by the company, not a person? *(Comptes et licences au nom de l'entreprise ?)* | Admin accounts sit on company-owned identities with MFA. |
| C4 | What personal or sensitive data does it hold, and where is it stored? *(Quelles donnees personnelles, et ou ?)* | Covered by GDPR principles and DRC law; a data processing agreement exists or is planned. |
| C5 | Can the replacement run in company-owned infrastructure with least-privilege access? *(Hebergement et acces minimaux ?)* | Secrets manager, MFA, separate environments, backups, access granted by IT/Cloud managers. |
| C6 | Does the replacement use AI? *(Usage d'IA ?)* | Approved provider only, no training on company data, no sensitive data typed into AI tools, a person reviews any AI output that affects money. |
| C7 | Who maintains it after handover? *(Qui le maintient apres la livraison ?)* | Named owner, documentation, runbook, training session delivered. |
| C8 | What are the contract notice period and exit cost? *(Preavis et cout de sortie ?)* | Cancellation date known; parallel running fits before it. |
| C9 | What does it integrate with (APIs, webhooks, Google Workspace, accounting)? *(Quelles integrations ?)* | Every integration is listed and has a replacement path. |
| C10 | Is there a tested backup/restore and an availability target? *(Sauvegarde, restauration, disponibilite ?)* | Restore was tested once; target agreed with IT Senior Manager. |
| C11 | Is there a pilot with a rollback? *(Pilote avec retour arriere ?)* | Run in parallel for an agreed period; rollback procedure written and rehearsed. |

## Resume (FR)

Avant de remplacer un outil: verifier s'il est critique, si les donnees sont exportables, si les comptes appartiennent a l'entreprise, ou sont stockees les donnees personnelles, qui maintiendra l'outil, le cout de sortie, et prevoir un pilote avec retour arriere. Un systeme critique n'est jamais remplace sans validation du responsable IT et du FP&A.
