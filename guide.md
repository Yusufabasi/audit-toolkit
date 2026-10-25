# Tool Audit Guide

A repeatable way to answer: *which software should we replace or build, and in what order?*
Built for the first week of an internal-tools engagement. Takes about five working days for a
company of Nuru's size; the output is a ranked, justified opportunity list a CEO can approve.

## The five steps

1. **Inventory** (day 1). List every tool, including spreadsheets and manual processes.
   Pull from invoices, card statements, SSO/admin consoles and "how do you actually do X?".
   Fill `tool-audit-template.xlsx` > `Inventory` (columns A-G first).
2. **Interview** (days 1-3). 30 minutes per team using the questions below. Record pain in the
   user's own words in the "What it is used for / pain" column.
3. **Score** (day 3-4). For each row: value (1-5), confidence (0.1-1, how sure you are about the
   value), effort (1-5), risk (1-5). Mark critical systems `Y`. The sheet computes
   `score = value x confidence / (effort + risk)`, a decision and a rank.
4. **Check** (day 4). Run [the decision checklist](decision-checklist.md) on every
   "Replace / automate candidate". Anything critical goes to sign-off, not to a build.
5. **Roadmap** (day 5). Take the top candidates that pass, estimate effort in weeks, agree order
   with the CEO, and name an owner for each. See [proposed-approach.md](proposed-approach.md).

## Interview questions

Ask "show me" rather than "tell me" wherever you can: watching someone do the task finds
more than asking about it.

### Everyone
- Which tool or spreadsheet do you open first each morning, and what for?
- What did you do by hand last week that felt like it should be automatic?
- Which task do you avoid or postpone because the tool makes it painful?
- What breaks when the person who built that spreadsheet is away?
- What do you pay for that you rarely use?

### Finance
- How does a mobile-money payment become a reconciled invoice? Who touches it?
- Which reports are rebuilt by hand every week or month? How long do they take?
- Where could a wrong number cost us money, and who checks it?

### Customer care
- Where do customer requests arrive (phone, WhatsApp, in person), and where are they recorded?
- How do you find a customer's history in under a minute?
- Which questions do you answer a hundred times?

### Field operations
- How are jobs assigned and closed, and how do you know a site fault has been fixed?
- What happens when there is no signal on site?
- How is stock (meters, cables, batteries) tracked?

### Reporting and management
- Which decisions do you make from which numbers? How old are those numbers when you see them?
- Which investor or partner reports are built by hand?

### IT, cloud and security
- Who owns each account and licence? Which are in a person's name?
- Where is personal data stored? Who can export it?
- What is backed up, how often, and when did you last restore it?
- Which systems control or read from site equipment? Who has write access?
- Which AI tools are already in use, and with what data?

## Résumé (FR)

Cinq etapes: inventaire, entretiens, notation (valeur x confiance / (effort + risque)),
verification avec la liste de controle, puis feuille de route. Poser des questions concretes
(« montrez-moi ») plutot que generales. Les systemes critiques ne sont jamais remplaces sans
validation du responsable IT, et du responsable FP&A pour tout ce qui touche a l'argent.

## Common mistakes

- Ranking by licence price. Staff time and error cost usually dominate.
- Scoring confidence at 1.0 for everything. Use it to mark what you have not verified.
- Replacing a critical system because it scored well.
- Skipping the pilot. Run the new tool in parallel and keep a rollback until it has earned trust.
