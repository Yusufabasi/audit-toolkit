"""Builds the tool-audit workbook (real formulas) and the decision checklist.

    python build_workbook.py            # writes tool-audit-example.xlsx, tool-audit-template.xlsx,
                                        # decision-checklist.md next to this file

Score = value * confidence / (effort + risk). A tool flagged critical is never "replace":
the decision column forces sign-off, however well it scores.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

MAX_ROWS = 200  # formulas cover this many inventory rows so added rows are picked up
RANK_SLOTS = 30

HEADERS = [
    "Tool",
    "Category",
    "Monthly cost (USD)",
    "Owner team",
    "Users",
    "What it is used for / pain",
    "Critical system? (Y/N)",
    "Value (1-5)",
    "Confidence (0.1-1)",
    "Effort (1-5)",
    "Risk (1-5)",
    "Score",
    "Decision",
    "Rank",
    "Monthly saving if replaced (USD)",
]
_KEYS = [
    "tool", "category", "cost", "owner", "users", "use", "critical", "value", "conf",
    "effort", "risk", "score", "decision", "rank", "saving",
]
COL = {key: chr(ord("A") + i) for i, key in enumerate(_KEYS)}

REPLACE_LABEL = "Replace / automate candidate"
OVERRIDE_LABEL = "Do not replace without sign-off"

CHECKLIST = [
    ("C1", "Is it a critical system (site control, metering, billing, finance books)?",
     "Critique ? Ne pas remplacer sans validation.",
     "If yes: no replacement without IT Senior Manager and FP&A Lead sign-off; "
     "access read-only by default."),
    ("C2", "Can all data be exported in an open format (CSV, JSON, SQL)?",
     "Export complet possible ?",
     "A full export has been tested, not just promised."),
    ("C3", "Are the accounts and licences owned by the company, not a person?",
     "Comptes et licences au nom de l'entreprise ?",
     "Admin accounts sit on company-owned identities with MFA."),
    ("C4", "What personal or sensitive data does it hold, and where is it stored?",
     "Quelles donnees personnelles, et ou ?",
     "Covered by GDPR principles and DRC law; a data processing agreement exists or is planned."),
    ("C5", "Can the replacement run in company-owned infrastructure with least-privilege access?",
     "Hebergement et acces minimaux ?",
     "Secrets manager, MFA, separate environments, backups, access granted by IT/Cloud managers."),
    ("C6", "Does the replacement use AI?",
     "Usage d'IA ?",
     "Approved provider only, no training on company data, no sensitive data typed into AI tools, "
     "a person reviews any AI output that affects money."),
    ("C7", "Who maintains it after handover?",
     "Qui le maintient apres la livraison ?",
     "Named owner, documentation, runbook, training session delivered."),
    ("C8", "What are the contract notice period and exit cost?",
     "Preavis et cout de sortie ?",
     "Cancellation date known; parallel running fits before it."),
    ("C9", "What does it integrate with (APIs, webhooks, Google Workspace, accounting)?",
     "Quelles integrations ?",
     "Every integration is listed and has a replacement path."),
    ("C10", "Is there a tested backup/restore and an availability target?",
     "Sauvegarde, restauration, disponibilite ?",
     "Restore was tested once; target agreed with IT Senior Manager."),
    ("C11", "Is there a pilot with a rollback?",
     "Pilote avec retour arriere ?",
     "Run in parallel for an agreed period; rollback procedure written and rehearsed."),
]

EXAMPLE_ROWS = [
    # Fictional company "Soleil Grid". Figures are illustrative.
    dict(tool="Weekly KPI report (hand-built in spreadsheets)", category="Reporting",
         monthly_cost_usd=0, owner="Finance", users=4,
         use="About 12 hours a week copying numbers between sheets; frequent mistakes.",
         critical="N", value=5, confidence=0.9, effort=2, risk=1),
    dict(tool="Forms / survey SaaS", category="Data collection", monthly_cost_usd=120,
         owner="Operations", users=15, use="Field visit checklists; exports to sheets by hand.",
         critical="N", value=3, confidence=0.9, effort=1, risk=1),
    dict(tool="Site monitoring portal (vendor)", category="Operations", monthly_cost_usd=350,
         owner="Operations", users=6,
         use="Live generation and outage view; CSV export looks easy to copy.",
         critical="Y", value=5, confidence=1.0, effort=2, risk=2),
    dict(tool="Mobile-money reconciliation (spreadsheets)", category="Finance",
         monthly_cost_usd=0, owner="Finance", users=3,
         use="Match payments to invoices daily by hand.",
         critical="N", value=5, confidence=0.8, effort=3, risk=3),
    dict(tool="Helpdesk SaaS", category="Customer support", monthly_cost_usd=450,
         owner="Customer care", users=8, use="Tickets and macros; reports are limited.",
         critical="N", value=4, confidence=0.7, effort=3, risk=2),
    dict(tool="Field work-order app (SaaS)", category="Operations", monthly_cost_usd=380,
         owner="Operations", users=20, use="Technician jobs; needs to work offline.",
         critical="N", value=4, confidence=0.7, effort=4, risk=3),
    dict(tool="E-signature SaaS", category="Legal", monthly_cost_usd=90, owner="Admin", users=5,
         use="Contracts; low volume.", critical="N", value=2, confidence=0.9, effort=2, risk=3),
    dict(tool="Project tracker SaaS", category="Productivity", monthly_cost_usd=200,
         owner="Management", users=25, use="Task boards; works well.",
         critical="N", value=2, confidence=0.8, effort=3, risk=2),
    dict(tool="Email and office suite", category="Productivity", monthly_cost_usd=1500,
         owner="IT", users=60, use="Mail, docs, storage; the company's identity backbone.",
         critical="N", value=3, confidence=0.9, effort=5, risk=5),
    dict(tool="Accounting software", category="Finance", monthly_cost_usd=800, owner="Finance",
         users=6, use="General ledger and statutory reporting.",
         critical="Y", value=4, confidence=0.6, effort=5, risk=5),
]


def score(value: int, confidence: float, effort: int, risk: int) -> float | None:
    """value * confidence / (effort + risk), or None when effort + risk is 0.

    value must be 1-5 and confidence 0.1-1 (a percentage like 80 is rejected). Effort and risk
    may be 0 only so this mirrors the sheet, which shows CHECK INPUT instead of dividing by zero.
    """
    if not 1 <= value <= 5:
        raise ValueError(f"value must be 1-5, got {value}")
    if not 0.1 <= confidence <= 1.0:
        raise ValueError(
            f"confidence must be between 0.1 and 1 (not a percentage), got {confidence}"
        )
    if not (0 <= effort <= 5 and 0 <= risk <= 5):
        raise ValueError("effort and risk must be 0-5")
    if effort + risk == 0:
        return None
    return value * confidence / (effort + risk)


def _score_formula(r: int) -> str:
    v, c, e, k = (f"{COL[x]}{r}" for x in ("value", "conf", "effort", "risk"))
    blank = f'OR({v}="",{c}="",{e}="",{k}="")'
    return f'=IF({blank},"",IF({e}+{k}=0,"CHECK INPUT",{v}*{c}/({e}+{k})))'


def _decision_formula(r: int) -> str:
    s, crit = f"{COL['score']}{r}", f"{COL['critical']}{r}"
    return (
        f'=IF({s}="","",IF({s}="CHECK INPUT","CHECK INPUT",IF({crit}="Y","{OVERRIDE_LABEL}",'
        f'IF({s}>=Settings!$B$1,"{REPLACE_LABEL}",IF({s}>=Settings!$B$2,"Review","Keep")))))'
    )


def _rank_formula(r: int) -> str:
    s = COL["score"]
    return (
        f'=IF(ISNUMBER({s}{r}),RANK({s}{r},${s}$2:${s}${MAX_ROWS})'
        f'+COUNTIF(${s}$2:{s}{r},{s}{r})-1,"")'
    )


def _saving_formula(r: int) -> str:
    return f'=IF({COL["decision"]}{r}="{REPLACE_LABEL}",N({COL["cost"]}{r}),0)'


def build_workbook(path: Path, rows: list[dict]) -> None:
    for row in rows:  # validate everything before writing anything
        score(row["value"], row["confidence"], row["effort"], row["risk"])

    wb = Workbook()
    inv = wb.active
    inv.title = "Inventory"
    inv.append(HEADERS)
    for cell in inv[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F6F43")
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for row in rows:
        inv.append([
            row["tool"], row["category"], row["monthly_cost_usd"], row["owner"], row["users"],
            row["use"], row["critical"], row["value"], row["confidence"], row["effort"],
            row["risk"],
        ])
    for r in range(2, MAX_ROWS + 1):  # formulas on every row so appended rows score themselves
        inv[f"{COL['score']}{r}"] = _score_formula(r)
        inv[f"{COL['decision']}{r}"] = _decision_formula(r)
        inv[f"{COL['rank']}{r}"] = _rank_formula(r)
        inv[f"{COL['saving']}{r}"] = _saving_formula(r)

    for col, kind, lo, hi in (
        ("value", "whole", 1, 5),
        ("effort", "whole", 1, 5),
        ("risk", "whole", 1, 5),
        ("conf", "decimal", 0.1, 1),
    ):
        dv = DataValidation(
            type=kind, operator="between", formula1=lo, formula2=hi,
            showErrorMessage=True, error=f"Enter a number from {lo} to {hi}",
        )
        dv.add(f"{COL[col]}2:{COL[col]}{MAX_ROWS}")
        inv.add_data_validation(dv)
    crit = DataValidation(type="list", formula1='"Y,N"', showErrorMessage=True)
    crit.add(f"{COL['critical']}2:{COL['critical']}{MAX_ROWS}")
    inv.add_data_validation(crit)
    widths = (34, 16, 14, 14, 8, 48, 12, 9, 12, 9, 9, 9, 30, 7, 16)
    for letter, width in zip("ABCDEFGHIJKLMNO", widths, strict=True):
        inv.column_dimensions[letter].width = width
    inv.freeze_panes = "B2"

    settings = wb.create_sheet("Settings")
    settings["A1"], settings["B1"] = "Replace threshold (score >=)", 1.0
    settings["A2"], settings["B2"] = "Review threshold (score >=)", 0.5
    settings["A4"] = "Score = value x confidence / (effort + risk). Critical systems never qualify."
    settings.column_dimensions["A"].width = 34

    rank = wb.create_sheet("Ranking", 1)
    rank.append(["Rank", "Tool", "Score", "Decision", "Monthly cost (USD)"])
    for cell in rank[1]:
        cell.font = Font(bold=True)
    rank_col = f"Inventory!${COL['rank']}$2:${COL['rank']}${MAX_ROWS}"
    for n in range(1, RANK_SLOTS + 1):
        r = n + 1
        rank[f"A{r}"] = n
        for out, key in (("B", "tool"), ("C", "score"), ("D", "decision"), ("E", "cost")):
            src = f"Inventory!${COL[key]}$2:${COL[key]}${MAX_ROWS}"
            rank[f"{out}{r}"] = f'=IFERROR(INDEX({src},MATCH({n},{rank_col},0)),"")'
    cost_rng = f"Inventory!{COL['cost']}2:{COL['cost']}{MAX_ROWS}"
    saving_rng = f"Inventory!{COL['saving']}2:{COL['saving']}{MAX_ROWS}"
    rank["G1"], rank["H1"] = "Total monthly spend (USD)", f"=SUM({cost_rng})"
    rank["G2"], rank["H2"] = "Spend on replace candidates (USD)", f"=SUM({saving_rng})"
    rank["G3"] = "Savings exclude staff time, which is often the bigger gain."
    for letter, width in zip("ABCDEGH", (7, 44, 9, 32, 16, 34, 12), strict=True):
        rank.column_dimensions[letter].width = width

    chk = wb.create_sheet("Checklist")
    chk.append(["ID", "Question", "Question (FR)", "Passes when"])
    for cell in chk[1]:
        cell.font = Font(bold=True)
    for item in CHECKLIST:
        chk.append(list(item))
    for letter, width in zip("ABCD", (6, 60, 40, 80), strict=True):
        chk.column_dimensions[letter].width = width

    wb.save(path)


def checklist_markdown() -> str:
    lines = [
        "# Keep vs replace: decision checklist",
        "",
        "Generated from `build_workbook.py` so the workbook and this page never disagree.",
        "Answer every item for each replace candidate before any build starts.",
        "",
        "| ID | Question | Passes when |",
        "|----|----------|-------------|",
    ]
    lines += [f"| {i} | {q} *({fr})* | {ok} |" for i, q, fr, ok in CHECKLIST]
    lines += [
        "",
        "## Resume (FR)",
        "",
        "Avant de remplacer un outil: verifier s'il est critique, si les donnees sont "
        "exportables, si les comptes appartiennent a l'entreprise, ou sont stockees les donnees "
        "personnelles, qui maintiendra l'outil, le cout de sortie, et prevoir un pilote avec "
        "retour arriere. Un systeme critique n'est jamais remplace sans validation du "
        "responsable IT et du FP&A.",
        "",
    ]
    return "\n".join(lines)


def decide(row: dict, replace_at: float = 1.0, review_at: float = 0.5) -> str:
    """Python mirror of the sheet's Decision column."""
    s = score(row["value"], row["confidence"], row["effort"], row["risk"])
    if s is None:
        return "CHECK INPUT"
    if row["critical"] == "Y":
        return OVERRIDE_LABEL
    if s >= replace_at:
        return REPLACE_LABEL
    return "Review" if s >= review_at else "Keep"


def worked_example_markdown() -> str:
    ranked = sorted(
        EXAMPLE_ROWS,
        key=lambda r: -score(r["value"], r["confidence"], r["effort"], r["risk"]),
    )
    table = [
        "| Rank | Tool | Cost/mo (USD) | V | C | E | R | Score | Decision |",
        "|------|------|---------------|---|---|---|---|-------|----------|",
    ]
    for n, r in enumerate(ranked, start=1):
        s = score(r["value"], r["confidence"], r["effort"], r["risk"])
        table.append(
            f"| {n} | {r['tool']} | {r['monthly_cost_usd']} | {r['value']} | {r['confidence']} "
            f"| {r['effort']} | {r['risk']} | {s:.2f} | {decide(r)} |"
        )
    total = sum(r["monthly_cost_usd"] for r in EXAMPLE_ROWS)
    saving = sum(r["monthly_cost_usd"] for r in EXAMPLE_ROWS if decide(r) == REPLACE_LABEL)
    return "\n".join(
        [
            "# Worked example: Soleil Grid (fictional)",
            "",
            "A made-up solar mini-grid operator with ten tools. Nothing here describes a real "
            "company. Generated from `build_workbook.py`, the same data as "
            "`tool-audit-example.xlsx`.",
            "",
            *table,
            "",
            f"Total monthly spend {total} USD; spend on replace candidates {saving} USD.",
            "",
            "## Reading the result",
            "",
            "- **The best first project is not the most expensive tool.** The hand-built weekly "
            "report costs nothing in licences but eats about 12 hours a week, so it ranks first: "
            "high value, low effort, low risk. Licence savings understate the gain; staff time is "
            "usually the bigger number.",
            "- **The forms SaaS is the only licence worth replacing now.** It is small, safe and "
            "quick, so it makes a good early win.",
            "- **A high score does not override a critical system.** The vendor monitoring portal "
            "scores 1.25 because its CSV export looks easy to copy, but it is the read path for "
            "site operations. The sheet forces `Do not replace without sign-off`; the decision "
            "goes to the IT Senior Manager and, for anything touching money, the FP&A Lead.",
            "- **Expensive is not the same as replaceable.** The email suite costs the most and "
            "scores lowest because it is the identity backbone: effort 5, risk 5.",
            "- **Next:** run the decision checklist on the two candidates, pilot in parallel with "
            "a rollback, then revisit the middle of the table (mobile-money reconciliation, "
            "helpdesk) once the first tool is stable.",
            "",
            "## Resume (FR)",
            "",
            "Le meilleur premier projet n'est pas l'outil le plus cher: le rapport hebdomadaire "
            "fait a la main ne coute aucune licence mais prend environ 12 heures par semaine. "
            "Un score eleve ne leve jamais la regle sur les systemes critiques: le portail de "
            "supervision reste soumis a validation.",
            "",
        ]
    )


def main() -> None:
    here = Path(__file__).resolve().parent
    build_workbook(here / "tool-audit-example.xlsx", EXAMPLE_ROWS)
    build_workbook(here / "tool-audit-template.xlsx", [])
    (here / "decision-checklist.md").write_text(checklist_markdown(), encoding="utf-8")
    (here / "worked-example.md").write_text(worked_example_markdown(), encoding="utf-8")
    print("wrote example + template workbooks, decision-checklist.md, worked-example.md")


if __name__ == "__main__":
    main()
