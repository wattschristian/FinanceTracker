"""
Renders the workbook.

The Budget tab (the original planner sheet) is never touched. The generated tabs
(Transactions, Monthly Stats, Accounts, Sync Log) are rebuilt from the local ledger
on every run, which is safer than inserting rows into a live sheet. Before rebuilding
we read back the things a human may have typed into the generated tabs and keep them:

  * Transactions!Category  -> becomes a per-transaction category override
  * Accounts!Owner/Label   -> updates config.json
  * Monthly Stats!Advice   -> preserved per month
  * Monthly Stats goal amount / target date (blue cells) -> preserved
"""
from __future__ import annotations

import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "Arial"
SHEET_TXN, SHEET_STATS, SHEET_ACCTS, SHEET_LOG, SHEET_LISTS = "Transactions", "Monthly Stats", "Accounts", "Sync Log", "Lists"
SHEET_BUDGETS = "Category Budgets"
GENERATED = [SHEET_TXN, SHEET_STATS, SHEET_BUDGETS, SHEET_ACCTS, SHEET_LISTS]

TXN_HEADERS = ["Date", "Account", "Owner", "Description", "Category", "Income", "Expense",
               "Balance Before", "Balance After", "Month", "Txn ID"]
TXN_WIDTHS = [12, 34, 11, 42, 19, 13, 13, 15, 15, 10, 14]
COL = {h: get_column_letter(i + 1) for i, h in enumerate(TXN_HEADERS)}   # e.g. COL["Income"] == "F"
MAX_ROW = 50000  # bounded ranges keep SUMIFS fast and LibreOffice-friendly

MONEY = '$#,##0.00;[Red]-$#,##0.00;"-"'
PCT = "0.0%"
DATE_FMT = "mm/dd/yyyy"

ADVICE_HEADER = "Advice for next month"
STATS_HEADERS = ["Month", "Income", "Expenses", "Net", "Savings Rate", "Top Expense Category",
                 "Top Category $", "Budgeted Expenses", "Over / (Under) Budget", "# Transactions", ADVICE_HEADER]


# ----------------------------------------------------------------- style helpers
def fill(hex_rgb: str) -> PatternFill:
    return PatternFill("solid", start_color=hex_rgb, end_color=hex_rgb)


def font(bold=False, color="000000", size=10, italic=False) -> Font:
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)


THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BLUE_INPUT = font(color="0000FF")
HEADER_FILL = fill("D9D9D9")


def month_label(iso_date: str) -> str:
    d = date.fromisoformat(iso_date)
    return d.strftime("%b %Y")                      # "Sep 2026"  (used as the join key everywhere)


def month_sort_key(label: str) -> tuple:
    return (datetime.strptime(label, "%b %Y").year, datetime.strptime(label, "%b %Y").month)


def _set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _header_row(ws, row, headers, widths=None):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = font(bold=True)
        cell.fill = HEADER_FILL
        cell.border = BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        _set_widths(ws, widths)


# ------------------------------------------------------------ read user edits
def capture_user_edits(path: Path) -> dict:
    """Read back human edits from the generated tabs. Safe to call when the file doesn't exist."""
    out = {"category_overrides": {}, "account_edits": {}, "advice": {}, "goal_amount": None, "goal_date": None,
           "budget_inputs": {}}
    if not path.exists():
        return out
    wb = load_workbook(path, data_only=True)
    if SHEET_TXN in wb.sheetnames:
        ws = wb[SHEET_TXN]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) >= 11 and row[10] and row[4]:
                out["category_overrides"][str(row[10])] = str(row[4]).strip()
    if SHEET_ACCTS in wb.sheetnames:
        ws = wb[SHEET_ACCTS]
        for row in ws.iter_rows(min_row=1, values_only=True):
            if row and row[0] == "Owner":
                continue
            if row and len(row) >= 11 and row[10]:
                out["account_edits"][str(row[10])] = {
                    "owner": (row[0] or "").strip() if isinstance(row[0], str) else None,
                    "label": (row[1] or "").strip() if isinstance(row[1], str) else None,
                    "hidden": str(row[9]).strip().lower() in ("yes", "y", "true", "x") if row[9] else False,
                }
    if SHEET_STATS in wb.sheetnames:
        ws = wb[SHEET_STATS]
        advice_col = None
        section = None
        for row in ws.iter_rows(min_row=1):
            vals = [c.value for c in row]
            if vals and isinstance(vals[0], str) and vals[0].endswith(SECTION_SUFFIX):
                section = vals[0][: -len(SECTION_SUFFIX)]
                continue
            if ADVICE_HEADER in vals:
                advice_col = vals.index(ADVICE_HEADER)
                continue
            if advice_col is not None and vals and isinstance(vals[0], str):
                if vals[0] in ("Average / month", "Total", ""):
                    advice_col = None
                    continue
                try:
                    month_sort_key(vals[0])
                except ValueError:
                    continue
                if len(vals) > advice_col and vals[advice_col]:
                    out["advice"][f"{section}|{vals[0]}" if section else vals[0]] = str(vals[advice_col])
            if vals and vals[0] == "Goal amount" and isinstance(vals[1], (int, float)):
                out["goal_amount"] = float(vals[1])
            if vals and vals[0] == "Target date" and isinstance(vals[1], (datetime, date)):
                out["goal_date"] = vals[1].date() if isinstance(vals[1], datetime) else vals[1]
    if SHEET_BUDGETS in wb.sheetnames:
        out["budget_inputs"] = _capture_budget_inputs(wb[SHEET_BUDGETS])
    return out


def _capture_budget_inputs(ws) -> dict:
    """Blue cells on the Category Budgets tab: month, savings rate, per-category guideline %/floor, per-section split."""
    got: dict = {}
    section = None
    in_guidelines = False
    for row in ws.iter_rows(min_row=1):
        vals = [c.value for c in row]
        head = vals[0] if vals and isinstance(vals[0], str) else None
        if head == "Month in focus" and isinstance(vals[1], str):
            got["month"] = vals[1]
        elif head == "Target savings rate" and isinstance(vals[1], (int, float)):
            got["savings_rate"] = float(vals[1])
        elif head == "Category" and len(vals) > 4 and vals[2] == "% of take-home":
            in_guidelines = True
            continue
        elif head and head.endswith(BUDGET_SECTION_SUFFIX):
            section = head[: -len(BUDGET_SECTION_SUFFIX)]
            in_guidelines = False
        elif head == "Share of shared costs" and section and isinstance(vals[1], (int, float)):
            got[f"split|{section}"] = float(vals[1])
        elif in_guidelines and head and vals[1] == "Flexible":
            if isinstance(vals[2], (int, float)):
                got[f"guideline|{head}|pct"] = float(vals[2])
            if isinstance(vals[3], (int, float)):
                got[f"guideline|{head}|min"] = float(vals[3])
        elif in_guidelines and head == "Total of guideline percentages":
            in_guidelines = False
    return got


# -------------------------------------------------------------------- render
def render(store, txns: list[dict], path: Path, log_entry: dict, user_edits: dict) -> None:
    cfg = store.config
    path = Path(path)
    if path.exists():
        wb = load_workbook(path)
    else:
        wb = Workbook()
    _ensure_budget_sheet(wb, cfg["budget_sheet"]["name"])

    # drop generated sheets; Sync Log is kept and appended to
    for name in GENERATED:
        if name in wb.sheetnames:
            del wb[name]

    ws_lists = wb.create_sheet(SHEET_LISTS)
    ws_txn = wb.create_sheet(SHEET_TXN)
    months, owners_present = _write_transactions(ws_txn, store, txns, cfg)
    _write_lists(ws_lists, cfg, months)
    ws_lists.sheet_state = "hidden"

    ws_stats = wb.create_sheet(SHEET_STATS)
    _write_stats(ws_stats, store, months, owners_present, cfg, user_edits)

    ws_bud = wb.create_sheet(SHEET_BUDGETS)
    _write_category_budgets(ws_bud, store, months, owners_present, cfg, user_edits)

    ws_accts = wb.create_sheet(SHEET_ACCTS)
    _write_accounts(ws_accts, store, cfg)

    ws_log = wb[SHEET_LOG] if SHEET_LOG in wb.sheetnames else wb.create_sheet(SHEET_LOG)
    _append_log(ws_log, log_entry)

    # order: Budget, Transactions, Monthly Stats, Accounts, Sync Log, (hidden Lists)
    order = [cfg["budget_sheet"]["name"], SHEET_TXN, SHEET_STATS, SHEET_BUDGETS, SHEET_ACCTS, SHEET_LOG, SHEET_LISTS]
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [s for s in wb._sheets if s.title not in order]
    wb.active = wb.sheetnames.index(SHEET_TXN)

    _safe_save(wb, path, cfg.get("backups_to_keep", 10))


def _ensure_budget_sheet(wb, budget_name: str):
    if budget_name in wb.sheetnames:
        return
    if "Sheet1" in wb.sheetnames:
        wb["Sheet1"].title = budget_name
    elif "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1:   # brand-new openpyxl workbook
        wb["Sheet"].title = budget_name
    else:
        wb.create_sheet(budget_name, 0)


def _write_lists(ws, cfg, months=()):
    ws["A1"], ws["B1"], ws["C1"], ws["D1"] = "Categories", "Owners", "Months", "Month start"
    for i, c in enumerate(cfg["categories"], start=2):
        ws.cell(row=i, column=1, value=c)
    for i, o in enumerate(cfg["owners"], start=2):
        ws.cell(row=i, column=2, value=o)
    for i, m in enumerate(months, start=2):
        ws.cell(row=i, column=3, value=m)
        d = ws.cell(row=i, column=4, value=datetime.strptime(m, "%b %Y").date())
        d.number_format = DATE_FMT


# ------------------------------------------------------------- Transactions
def _write_transactions(ws, store, txns, cfg):
    owners = cfg["owners"]
    cats = cfg["categories"]
    owner_order = {o: i for i, o in enumerate(owners)}
    _header_row(ws, 1, TXN_HEADERS, TXN_WIDTHS)
    ws.freeze_panes = "A2"

    cat_dv = DataValidation(type="list", formula1=f"={SHEET_LISTS}!$A$2:$A${1 + len(cats)}", allow_blank=True)
    ws.add_data_validation(cat_dv)

    by_month: dict[str, list[dict]] = {}
    for t in txns:
        by_month.setdefault(month_label(t["date"]), []).append(t)
    months = sorted(by_month, key=month_sort_key)

    r = 2
    last_col = len(TXN_HEADERS)
    for m in months:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        c = ws.cell(row=r, column=1, value=datetime.strptime(m, "%b %Y").strftime("%B %Y"))
        c.font = font(bold=True, color="FFFFFF", size=12)
        c.fill = fill("404040")
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 20
        r += 1

        groups: dict[str, list[dict]] = {}
        for t in by_month[m]:
            groups.setdefault(t["account_id"], []).append(t)
        ordered = sorted(groups, key=lambda a: (owner_order.get(store.account_owner(a), 99), store.account_label(a)))

        for aid in ordered:
            owner = store.account_owner(aid)
            style = owners.get(owner, {"header_fill": "7F7F7F", "row_fill": "F2F2F2"})
            label = store.account_label(aid)

            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            c = ws.cell(row=r, column=1, value=f"{owner}  ·  {label}")
            c.font = font(bold=True, color="FFFFFF")
            c.fill = fill(style["header_fill"])
            c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            r += 1

            first = r
            for t in groups[aid]:
                cat = t.get("category_override") or t.get("category") or "Other"
                t["category_written"] = cat          # remembered so later sheet edits can be told apart from our own output
                vals = [
                    date.fromisoformat(t["date"]), label, owner,
                    t.get("merchant") or t.get("name"), cat,
                    round(-t["amount"], 2) if t["amount"] < 0 else None,
                    round(t["amount"], 2) if t["amount"] > 0 else None,
                    t.get("balance_before"), t.get("balance_after"), m, t["id"],
                ]
                for ci, v in enumerate(vals, start=1):
                    cell = ws.cell(row=r, column=ci, value=v)
                    cell.font = font()
                    cell.border = BOX
                ws.cell(row=r, column=1).number_format = DATE_FMT
                for ci in (6, 7, 8, 9):
                    ws.cell(row=r, column=ci).number_format = MONEY
                ws.cell(row=r, column=2).fill = fill(style["row_fill"])
                ws.cell(row=r, column=3).fill = fill(style["row_fill"])
                ws.cell(row=r, column=5).fill = fill(cats.get(cat, "FFFFFF"))
                ws.cell(row=r, column=11).font = font(color="808080", size=8)
                cat_dv.add(ws.cell(row=r, column=5))
                r += 1
            last = r - 1

            # subtotal row (no Month key -> excluded from stats)
            ws.cell(row=r, column=4, value="Subtotal").font = font(bold=True)
            ws.cell(row=r, column=6, value=f"=SUM(F{first}:F{last})")
            ws.cell(row=r, column=7, value=f"=SUM(G{first}:G{last})")
            ws.cell(row=r, column=9, value=f"=I{last}")
            for ci in range(1, last_col + 1):
                cell = ws.cell(row=r, column=ci)
                cell.fill = fill(style["row_fill"])
                cell.font = font(bold=True)
                if ci in (6, 7, 9):
                    cell.number_format = MONEY
            ws.cell(row=r, column=8, value="ending balance →").font = font(italic=True, color="595959", size=9)
            ws.cell(row=r, column=8).alignment = Alignment(horizontal="right")
            r += 2

    if not months:
        ws.cell(row=2, column=1, value="No posted transactions yet. Run the sync again after your first charges settle.").font = font(italic=True)

    owners_present = [o for o in owners if any(store.account_owner(t["account_id"]) == o for t in txns)]
    return months, owners_present


# ------------------------------------------------------------ Monthly Stats
BASE_COLS = ["Month", "Income", "Other Inflows", "Spending", "Net (income − spending)", "Net incl. other inflows",
             "Savings Rate", "Top Spending Category", "Top Category $", "# Transactions"]
HOUSEHOLD_EXTRA = ["Budgeted Expenses", "Over / (Under) Budget"]
SECTION_SUFFIX = " — Monthly Stats"


def _write_stats(ws, store, months, owners_present, cfg, user_edits):
    ws["A1"] = "Monthly Statistics"
    ws["A1"].font = font(bold=True, size=14)
    ws["A2"] = ("Rebuilt on every sync; every number is a live formula over the Transactions tab. "
                "Income = recurring pay only. Other Inflows = one-off money in (loan disbursements, refunds, aid). "
                "Transfers between your own accounts are excluded everywhere. "
                "Only the blue cells and the 'Advice for next month' columns are kept between runs.")
    ws["A2"].font = font(italic=True, color="595959", size=9)
    ws.column_dimensions["A"].width = 26

    r = 4
    for owner in owners_present:
        style = cfg["owners"].get(owner, {"header_fill": "7F7F7F"})
        info = _stats_section(ws, r, owner, owner, style["header_fill"], months, cfg, store, user_edits, household=False)
        r = info["next_row"]
    info = _stats_section(ws, r, "Household (all accounts)", None, "404040", months, cfg, store, user_edits, household=True)
    _savings_projection(ws, info, months, cfg, store, user_edits)
    ws.freeze_panes = "B4"


def _stats_section(ws, r0, title, owner, title_fill, months, cfg, store, user_edits, household):
    cats = cfg["categories"]
    non_spend = list(cfg.get("non_spending_categories", ["Transfer"]))
    T = f"'{SHEET_TXN}'!"
    rng = lambda col: f"{T}${col}$2:${col}${MAX_ROW}"
    R_MONTH, R_CAT, R_INC, R_EXP, R_ID, R_OWN = (rng(COL["Month"]), rng(COL["Category"]), rng(COL["Income"]),
                                                 rng(COL["Expense"]), rng(COL["Txn ID"]), rng(COL["Owner"]))
    excl = "".join(f',{R_CAT},"<>{c}"' for c in non_spend)
    oc = f',{R_OWN},"{owner}"' if owner else ""
    B = f"'{cfg['budget_sheet']['name']}'!"
    n = len(months)

    headers = BASE_COLS + (HOUSEHOLD_EXTRA if household else []) + [ADVICE_HEADER]
    ci = {h: i + 1 for i, h in enumerate(headers)}          # header -> column number
    cl = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}
    adv_col = ci[ADVICE_HEADER]

    # ---- title ------------------------------------------------------------
    ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=len(headers))
    t = ws.cell(row=r0, column=1, value=f"{title}{SECTION_SUFFIX}")
    t.font = font(bold=True, color="FFFFFF", size=12); t.fill = fill(title_fill)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r0].height = 20

    header_row = r0 + 1
    first = header_row + 1
    avg_row = first + n
    total_row = avg_row + 1
    matrix_header = total_row + 3
    matrix_cats = [c for c in cats if c not in non_spend and c not in ("Income", "Refund", "Other Inflow")]
    matrix_first = matrix_header + 1
    matrix_last = matrix_first + len(matrix_cats) - 1
    matrix_total = matrix_last + 1
    mcol = lambda i: get_column_letter(2 + i)
    m_total_col, m_avg_col, m_pct_col = get_column_letter(2 + n), get_column_letter(3 + n), get_column_letter(4 + n)
    m_bud_col, m_avgvb_col = get_column_letter(5 + n), get_column_letter(6 + n)

    widths = [26, 13, 13, 13, 15, 15, 11, 22, 14, 13] + ([15, 16] if household else []) + [70]
    _header_row(ws, header_row, headers, widths)
    ws.row_dimensions[header_row].height = 30

    # ---- month rows ---------------------------------------------------------
    for i, m in enumerate(months):
        r = first + i
        col = mcol(i)
        top_rng = f"${col}${matrix_first}:${col}${matrix_last}"
        vals = {
            "Month": m,
            "Income": f'=SUMIFS({R_INC},{R_MONTH},$A{r},{R_CAT},"Income"{oc})',
            "Other Inflows": f'=SUMIFS({R_INC},{R_MONTH},$A{r},{R_CAT},"<>Income"{excl}{oc})',
            "Spending": f"=SUMIFS({R_EXP},{R_MONTH},$A{r}{excl}{oc})",
            "Net (income − spending)": f"={cl['Income']}{r}-{cl['Spending']}{r}",
            "Net incl. other inflows": f"={cl['Income']}{r}+{cl['Other Inflows']}{r}-{cl['Spending']}{r}",
            "Savings Rate": f"=IF({cl['Income']}{r}=0,0,{cl['Net (income − spending)']}{r}/{cl['Income']}{r})",
            "Top Spending Category": f'=IF(MAX({top_rng})=0,"",INDEX($A${matrix_first}:$A${matrix_last},MATCH(MAX({top_rng}),{top_rng},0)))',
            "Top Category $": f"=MAX({top_rng})",
            "# Transactions": f'=COUNTIFS({R_MONTH},$A{r},{R_ID},"<>"{oc})',
            ADVICE_HEADER: user_edits.get("advice", {}).get(f"{title}|{m}") or (user_edits.get("advice", {}).get(m, "") if household else ""),
        }
        if household:
            vals["Budgeted Expenses"] = f"={B}${cfg['budget_sheet']['expense_total_cell']}"
            vals["Over / (Under) Budget"] = f"={cl['Spending']}{r}-{cl['Budgeted Expenses']}{r}"
        for h, v in vals.items():
            cell = ws.cell(row=r, column=ci[h], value=v)
            cell.font = font(); cell.border = BOX
            cell.alignment = Alignment(vertical="top", wrap_text=(h == ADVICE_HEADER))
            if h in ("Income", "Other Inflows", "Spending", "Net (income − spending)", "Net incl. other inflows",
                     "Top Category $", "Budgeted Expenses", "Over / (Under) Budget"):
                cell.number_format = MONEY
            elif h == "Savings Rate":
                cell.number_format = PCT
        ws.cell(row=r, column=adv_col).fill = fill("FFFBE6")

    money_cols = ["Income", "Other Inflows", "Spending", "Net (income − spending)", "Net incl. other inflows"]
    if household:
        money_cols += ["Budgeted Expenses", "Over / (Under) Budget"]
    if n:
        seg = lambda c: f"{c}{first}:{c}{avg_row - 1}"
        ws.cell(row=avg_row, column=1, value="Average / month")
        ws.cell(row=total_row, column=1, value="Total")
        for h in money_cols:
            ws.cell(row=avg_row, column=ci[h], value=f"=AVERAGE({seg(cl[h])})")
            ws.cell(row=total_row, column=ci[h], value=f"=SUM({seg(cl[h])})")
        ws.cell(row=avg_row, column=ci["Savings Rate"],
                value=f"=IF({cl['Income']}{avg_row}=0,0,{cl['Net (income − spending)']}{avg_row}/{cl['Income']}{avg_row})")
        ws.cell(row=avg_row, column=ci["# Transactions"], value=f"=AVERAGE({seg(cl['# Transactions'])})")
        ws.cell(row=total_row, column=ci["# Transactions"], value=f"=SUM({seg(cl['# Transactions'])})")
        for r in (avg_row, total_row):
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=r, column=c)
                cell.font = font(bold=True); cell.fill = fill("F2F2F2"); cell.border = BOX
            for h in money_cols:
                ws.cell(row=r, column=ci[h]).number_format = MONEY
            ws.cell(row=r, column=ci["Savings Rate"]).number_format = PCT
            ws.cell(row=r, column=ci["# Transactions"]).number_format = "0.0"
    else:
        ws.cell(row=first, column=1, value="No months yet.").font = font(italic=True)

    # ---- category x month matrix --------------------------------------------
    ws.cell(row=matrix_header - 1, column=1, value=f"Where the money goes — {title}").font = font(bold=True, size=11)
    mheaders = ["Category"] + months + ["Total", "Avg / month", "% of spending"]
    if household:
        mheaders += ["Monthly budget", "Avg vs budget", "Latest month vs budget"]
    _header_row(ws, matrix_header, mheaders)
    budget_cells = cfg["budget_sheet"].get("category_cells", {})
    for i, cat in enumerate(matrix_cats):
        r = matrix_first + i
        c0 = ws.cell(row=r, column=1, value=cat)
        c0.fill = fill(cats[cat]); c0.font = font(); c0.border = BOX
        for j, m in enumerate(months):
            col = mcol(j)
            cell = ws.cell(row=r, column=2 + j, value=f"=SUMIFS({R_EXP},{R_MONTH},{col}${matrix_header},{R_CAT},$A{r}{oc})")
            cell.number_format = MONEY; cell.font = font(); cell.border = BOX
        if n:
            ws.cell(row=r, column=2 + n, value=f"=SUM({mcol(0)}{r}:{mcol(n - 1)}{r})")
            ws.cell(row=r, column=3 + n, value=f"=AVERAGE({mcol(0)}{r}:{mcol(n - 1)}{r})")
            ws.cell(row=r, column=4 + n, value=f"=IF({m_total_col}{matrix_total}=0,0,{m_total_col}{r}/{m_total_col}{matrix_total})")
        else:
            for k in (2, 3, 4):
                ws.cell(row=r, column=k, value=0)
        ncols = 5 + n
        if household:
            refs = budget_cells.get(cat)
            if refs:
                ws.cell(row=r, column=5 + n, value="=" + "+".join(f"{B}${ref[0]}${ref[1:]}" for ref in refs))
                ws.cell(row=r, column=6 + n, value=f"={m_avg_col}{r}-{m_bud_col}{r}")
                if n:
                    ws.cell(row=r, column=7 + n, value=f"={mcol(n - 1)}{r}-{m_bud_col}{r}")
            ncols = 8 + n
        for k in range(2 + n, ncols):
            cell = ws.cell(row=r, column=k)
            cell.number_format = PCT if k == 4 + n else MONEY
            cell.font = font(); cell.border = BOX
    ws.cell(row=matrix_total, column=1, value="Total spending")
    for j in range(n + 1):
        col = get_column_letter(2 + j)
        ws.cell(row=matrix_total, column=2 + j, value=f"=SUM({col}{matrix_first}:{col}{matrix_last})")
    if n:
        ws.cell(row=matrix_total, column=3 + n, value=f"=SUM({m_avg_col}{matrix_first}:{m_avg_col}{matrix_last})")
        if household:
            ws.cell(row=matrix_total, column=5 + n, value=f"=SUM({m_bud_col}{matrix_first}:{m_bud_col}{matrix_last})")
    for k in range(1, len(mheaders) + 1):
        cell = ws.cell(row=matrix_total, column=k)
        cell.font = font(bold=True); cell.fill = fill("F2F2F2"); cell.border = BOX
        if k >= 2 and k != 4 + n:
            cell.number_format = MONEY
    for j in range(len(mheaders) - 1):
        letter = get_column_letter(2 + j)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, 14)
    if household:
        ws.cell(row=matrix_total + 1, column=1,
                value="Budget figures come from the Budget tab (config.json → budget_sheet.category_cells says which cells).").font = font(italic=True, color="595959", size=9)

    return {"next_row": matrix_total + 4, "avg_row": avg_row, "net_col": cl["Net (income − spending)"],
            "matrix_header": matrix_header, "matrix_first": matrix_first, "matrix_last": matrix_last,
            "header_row": header_row, "first": first, "income_col": ci["Income"], "spend_col": ci["Spending"], "n": n}


def _savings_projection(ws, info, months, cfg, store, user_edits):
    B = f"'{cfg['budget_sheet']['name']}'!"
    n = info["n"]
    g = cfg.get("savings_goal", {})
    goal_amount = user_edits.get("goal_amount") or g.get("amount", 0)
    goal_date = user_edits.get("goal_date") or date.fromisoformat(g.get("target_date", date.today().isoformat()))
    goal_acct = _find_goal_account(store, g.get("account_match", ""))
    cur_bal = goal_acct["balance_current"] if goal_acct else 0
    as_of = goal_acct["balance_as_of"] if goal_acct else "—"

    s = info["next_row"]
    ws.cell(row=s, column=1, value="Savings projection (household)").font = font(bold=True, size=12)
    ws.cell(row=s, column=3, value="Blue cells are yours to edit; they survive re-syncs.").font = font(italic=True, color="595959", size=9)
    avg_net_ref = f"${info['net_col']}${info['avg_row']}" if n else "0"
    rows = [
        ("Goal", g.get("label", "Savings goal"), None, False),
        ("Goal amount", goal_amount, MONEY, True),
        ("Target date", goal_date, DATE_FMT, True),
        ("Goal account", store.account_label(goal_acct["id"]) if goal_acct else "(no account matched savings_goal.account_match)", None, False),
        ("Current balance", cur_bal, MONEY, False),
        ("Balance as of", as_of, None, False),
        ("Months remaining", f"=MAX(0,(YEAR(B{s + 3})-YEAR(TODAY()))*12+MONTH(B{s + 3})-MONTH(TODAY()))", "0", False),
        ("Average monthly net (income − spending)", f"={avg_net_ref}", MONEY, False),
        ("Projected balance at target date", f"=B{s + 5}+B{s + 8}*B{s + 7}", MONEY, False),
        ("Surplus / (shortfall) vs goal", f"=B{s + 9}-B{s + 2}", MONEY, False),
        ("Required monthly savings to hit goal", f"=IF(B{s + 7}=0,B{s + 2}-B{s + 5},(B{s + 2}-B{s + 5})/B{s + 7})", MONEY, False),
        ("Planned leftover per month (Budget tab)", f"={B}${cfg['budget_sheet']['leftover_cell']}", MONEY, False),
        ("Projected balance if planned leftover is saved", f"=B{s + 5}+B{s + 12}*B{s + 7}", MONEY, False),
    ]
    for i, (label, val, fmt, is_input) in enumerate(rows, start=1):
        a = ws.cell(row=s + i, column=1, value=label); a.font = font(bold=True); a.border = BOX
        b = ws.cell(row=s + i, column=2, value=val); b.border = BOX
        b.font = BLUE_INPUT if is_input else font()
        if is_input:
            b.fill = fill("FFFF99")
        if fmt:
            b.number_format = fmt
        b.alignment = Alignment(horizontal="right" if fmt else "left")
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width or 0, 16)

    if n:
        bar = BarChart(); bar.type = "bar"; bar.title = "Household spending by category (all months)"; bar.style = 10
        bar.add_data(Reference(ws, min_col=2 + n, min_row=info["matrix_header"], max_row=info["matrix_last"]), titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=1, min_row=info["matrix_first"], max_row=info["matrix_last"]))
        bar.height, bar.width, bar.legend = 9, 16, None
        ws.add_chart(bar, f"D{s + 1}")
        line = LineChart(); line.title = "Household income vs spending by month"; line.style = 12
        line.add_data(Reference(ws, min_col=info["income_col"], min_row=info["header_row"], max_row=info["avg_row"] - 1), titles_from_data=True)
        line.add_data(Reference(ws, min_col=info["spend_col"], min_row=info["header_row"], max_row=info["avg_row"] - 1), titles_from_data=True)
        line.set_categories(Reference(ws, min_col=1, min_row=info["first"], max_row=info["avg_row"] - 1))
        line.height, line.width = 9, 16
        ws.add_chart(line, f"L{s + 1}")


BUDGET_SECTION_SUFFIX = " — Budget vs Actual"
BUD_HEADERS = ["Category", "Type", "Guideline", "Guideline target", "Floor", "Suggested cap",
               "Pace target", f"Actual", "Over / (Under)", "% of cap used", "Status", "Avg / mo"]
BUD_WIDTHS = [33, 10, 46, 15, 12, 15, 14, 14, 15, 13, 12, 13]


def _complete_months(months: list[str], n_max: int) -> list[str]:
    """The most recent months that have fully elapsed (so averages aren't dragged down by a part-month)."""
    today = date.today()
    done = []
    for m in months:
        start = datetime.strptime(m, "%b %Y").date()
        nxt = (start.replace(day=28) + timedelta(days=4)).replace(day=1)
        if nxt <= today:
            done.append(m)
    if not done:                       # only a part-month of history so far
        done = list(months)
    return done[-n_max:] if n_max else done


def _avg_formula(terms: list[str], n: int) -> str:
    if not terms or not n:
        return "=0"
    return "=(" + "+".join(terms) + f")/{n}"


N_DERIV = 7          # derivation rows in every section (keeps the layout identical section to section)


def _section_rows(r0: int, n_flex: int, n_fixed: int) -> dict:
    d0 = r0 + 1
    flex_header = d0 + N_DERIV + 3
    flex_first = flex_header + 1
    flex_last = flex_first + n_flex - 1
    flex_total = flex_last + 1
    fix_header = flex_total + 2
    fix_first = fix_header + 1
    fix_last = fix_first + n_fixed - 1
    fix_total = fix_last + 1
    summary = fix_total + 2
    return {"d0": d0, "flex_header": flex_header, "flex_first": flex_first, "flex_last": flex_last,
            "flex_total": flex_total, "fix_header": fix_header, "fix_first": fix_first, "fix_last": fix_last,
            "fix_total": fix_total, "summary": summary, "end": summary + 8}


def _write_category_budgets(ws, store, months, owners_present, cfg, user_edits):
    """
    One section per owner plus a household section. Caps are derived, not typed:
        take-home income − fixed commitments − target savings = discretionary pool
    Flexible categories get a % -of-income target with a dollar floor, then the whole
    set is scaled to fit the pool. Everything is a formula, so editing a blue cell
    re-derives the caps immediately.
    """
    bm = cfg.get("budget_model", {})
    guidelines = bm.get("guidelines", {})
    fixed_cats = [c for c in bm.get("fixed_categories", []) if c in cfg["categories"]]
    flex_cats = [c for c in guidelines if c in cfg["categories"]]
    saved = user_edits.get("budget_inputs", {})
    n_hist = int(bm.get("months_of_history", 6))
    hist = _complete_months(months, n_hist)

    T = f"'{SHEET_TXN}'!"
    rng = lambda col: f"{T}${col}$2:${col}${MAX_ROW}"
    R_MONTH, R_CAT, R_INC, R_EXP, R_OWN = (rng(COL["Month"]), rng(COL["Category"]), rng(COL["Income"]),
                                           rng(COL["Expense"]), rng(COL["Owner"]))

    ws["A1"] = "Category Budgets — did we over- or underspend?"
    ws["A1"].font = font(bold=True, size=14)
    ws["A2"] = ("Caps are derived, not guessed: take-home income (recurring pay only, averaged over the complete months "
                "listed below) minus the bills that go out regardless, minus your savings target, leaves the discretionary "
                "pool. That pool is split across the flexible categories using the guideline percentages, floored at the "
                "dollar minimums, and scaled down proportionally if the guidelines add up to more than the pool.")
    ws["A2"].font = font(italic=True, color="595959", size=9)
    ws["A3"] = ("Blue cells are yours — change a percentage, a floor, the savings rate or the month and every cap "
                "recalculates. Your edits are kept on the next sync.")
    ws["A3"].font = font(italic=True, color="595959", size=9)
    _set_widths(ws, BUD_WIDTHS)

    # ---------------------------------------------------------------- controls
    sel_default = months[-1] if months else ""
    sel_month = saved.get("month") if saved.get("month") in months else sel_default
    ws["A5"] = "Month in focus"; ws["A5"].font = font(bold=True)
    c = ws["B5"]; c.value = sel_month; c.font = BLUE_INPUT; c.fill = fill("FFFF99"); c.border = BOX
    if months:
        dv = DataValidation(type="list", formula1=f"={SHEET_LISTS}!$C$2:$C${1 + len(months)}", allow_blank=True)
        ws.add_data_validation(dv); dv.add(c)
    ws["C5"] = "← pick any month; every section below follows it"
    ws["C5"].font = font(italic=True, color="595959", size=9)

    ws["A6"] = "Target savings rate"; ws["A6"].font = font(bold=True)
    c = ws["B6"]; c.value = saved.get("savings_rate", bm.get("savings_rate", 0.20))
    c.number_format = PCT; c.font = BLUE_INPUT; c.fill = fill("FFFF99"); c.border = BOX
    ws["C6"] = "of take-home pay, taken off the top before any category is funded (50/30/20 uses 20%)"
    ws["C6"].font = font(italic=True, color="595959", size=9)
    SAVE_RATE = "$B$6"

    # month-progress factor: a part-month is compared against a pro-rated target
    if months:
        L = f"'{SHEET_LISTS}'!"
        ws["A7"] = "Month progress"; ws["A7"].font = font(bold=True)
        ws["D7"] = f"=INDEX({L}$D$2:$D${1 + len(months)},MATCH($B$5,{L}$C$2:$C${1 + len(months)},0))"
        ws["D7"].number_format = DATE_FMT
        ws["E7"] = "=DAY(EOMONTH($D$7,0))"
        ws["F7"] = "=MIN($E$7,MAX(0,TODAY()-$D$7+1))"
        ws["B7"] = "=IF($E$7=0,1,$F$7/$E$7)"
        ws["B7"].number_format = PCT
        for ref in ("D7", "E7", "F7"):
            ws[ref].font = font(color="BFBFBF", size=8)
        ws["C7"] = "share of the month elapsed — a current month is judged against this share of its cap"
        ws["C7"].font = font(italic=True, color="595959", size=9)
    PACE = "$B$7"

    ws["A8"] = "Months averaged"; ws["A8"].font = font(bold=True)
    ws["B8"] = ", ".join(hist) if hist else "—"
    ws["B8"].font = font(italic=True, color="595959", size=9)

    # -------------------------------------------------------------- guidelines
    g_header = 10
    ws.cell(row=g_header - 1, column=1, value="Guidelines (shared by every section)").font = font(bold=True, size=12)
    _header_row(ws, g_header, ["Category", "Type", "% of take-home", "Floor $/mo (household)", "Where the number comes from"])
    g_first = g_header + 1
    grow = {}
    r = g_first
    for cat in flex_cats:
        g = guidelines[cat]
        key = f"guideline|{cat}"
        pct = saved.get(key + "|pct", g.get("pct", 0))
        floor = saved.get(key + "|min", g.get("min_household", 0))
        a = ws.cell(row=r, column=1, value=cat); a.fill = fill(cfg["categories"][cat]); a.border = BOX; a.font = font()
        ws.cell(row=r, column=2, value="Flexible").font = font()
        p = ws.cell(row=r, column=3, value=pct); p.number_format = PCT
        f = ws.cell(row=r, column=4, value=floor); f.number_format = MONEY
        for cell in (p, f):
            cell.font = BLUE_INPUT; cell.fill = fill("FFFF99"); cell.border = BOX
        n = ws.cell(row=r, column=5, value=g.get("note", "")); n.font = font(size=9, color="595959")
        ws.cell(row=r, column=2).border = BOX
        grow[cat] = r
        r += 1
    g_last = r - 1
    ws.cell(row=r, column=1, value="Total of guideline percentages").font = font(bold=True)
    ws.cell(row=r, column=3, value=f"=SUM(C{g_first}:C{g_last})").number_format = PCT
    ws.cell(row=r, column=3).font = font(bold=True)
    ws.cell(row=r, column=5, value="Fixed categories are not budgeted here — they are what they are. Their caps below are your own recent average.").font = font(italic=True, color="595959", size=9)
    r += 2

    # an owner with no recurring income of their own has no budget of their own — their spending is household spending
    budget_owners = [o for o in owners_present if _income_share(store, o, hist) > 0]
    skipped = [o for o in owners_present if o not in budget_owners]
    if skipped:
        note = ws.cell(row=r, column=1, value="No section for " + ", ".join(skipped) +
                       " — no recurring income lands there, so that spending is judged in the household section below.")
        note.font = font(italic=True, color="595959", size=9)
        r += 2

    hh_start = r
    for _ in budget_owners:
        hh_start = _section_rows(hh_start, len(flex_cats), len(fixed_cats))["end"]
    HH_FIXED = f"$F${_section_rows(hh_start, len(flex_cats), len(fixed_cats))['fix_total']}"

    common = dict(months=months, hist=hist, flex_cats=flex_cats, fixed_cats=fixed_cats, grow=grow, cfg=cfg,
                  store=store, saved=saved, SAVE_RATE=SAVE_RATE, PACE=PACE, HH_FIXED=HH_FIXED,
                  R_MONTH=R_MONTH, R_CAT=R_CAT, R_INC=R_INC, R_EXP=R_EXP, R_OWN=R_OWN)
    for owner in budget_owners:
        style = cfg["owners"].get(owner, {"header_fill": "7F7F7F"})
        r = _budget_section(ws, r, owner, owner, style["header_fill"], household=False, **common)
    _budget_section(ws, hh_start, "Household (both of you)", None, "404040", household=True, **common)
    ws.freeze_panes = "A5"


def _budget_section(ws, r0, title, owner, title_fill, months, hist, flex_cats, fixed_cats, grow,
                    cfg, store, saved, SAVE_RATE, PACE, HH_FIXED, R_MONTH, R_CAT, R_INC, R_EXP, R_OWN,
                    household):
    oc = f',{R_OWN},"{owner}"' if owner else ""
    nh = len(hist)
    sel = "$B$5"

    ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=len(BUD_HEADERS))
    t = ws.cell(row=r0, column=1, value=f"{title}{BUDGET_SECTION_SUFFIX}")
    t.font = font(bold=True, color="FFFFFF", size=12); t.fill = fill(title_fill)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r0].height = 20

    # rows are laid out first so the derivation block can forward-reference the tables
    g = _section_rows(r0, len(flex_cats), len(fixed_cats))
    d0, flex_header, flex_first, flex_last, flex_total = g["d0"], g["flex_header"], g["flex_first"], g["flex_last"], g["flex_total"]
    fix_header, fix_first, fix_last, fix_total, summary = g["fix_header"], g["fix_first"], g["fix_last"], g["fix_total"], g["summary"]

    INCOME, MYFIXED, SAVE, POOL, SPLIT, SCALE = (f"$B${d0 + 1}", f"$B${d0 + 3}", f"$B${d0 + 4}",
                                                 f"$B${d0 + 5}", f"$B${d0 + 6}", f"$B${d0 + 7}")
    SUM_RAW, SUM_MIN = f"$D${flex_total}", f"$E${flex_total}"

    inc_terms = [f'SUMIFS({R_INC},{R_MONTH},"{m}",{R_CAT},"Income"{oc})' for m in hist]
    split_default = saved.get(f"split|{title}")
    if split_default is None:
        split_default = 1.0 if household else round(_income_share(store, owner, hist), 4)

    ws.cell(row=d0, column=1, value="How the caps are derived").font = font(bold=True, size=11)
    deriv = [
        ("Take-home income (avg / mo)", _avg_formula(inc_terms, nh), MONEY, False,
         "recurring pay and interest only — one-off inflows are excluded on purpose"),
        ("Household fixed commitments (avg / mo)", f"={HH_FIXED}", MONEY, False,
         "rent, utilities, phone, insurance, loan payments, subscriptions — out the door regardless"),
        ("Your share of fixed commitments", f"={HH_FIXED}*{SPLIT}", MONEY, False,
         "charged by share, not by whose account paid — one of you covering rent while the other transfers money should not skew either budget"),
        ("Target savings", f"={INCOME}*{SAVE_RATE}", MONEY, False, "taken off the top, before any category is funded"),
        ("Discretionary pool", f"=MAX(0,{INCOME}-{MYFIXED}-{SAVE})", MONEY, False, "what is genuinely available for the flexible categories"),
        ("Share of shared costs", split_default, PCT, True,
         "your slice of household fixed bills and of shared floors like groceries. Defaults to your share of household income — change it to match who actually pays."),
        ("Allocation scale",
         f"=IF({SUM_RAW}<={POOL},1,"
         f"IF({POOL}<{SUM_MIN},IF({SUM_RAW}=0,0,{POOL}/{SUM_RAW}),"
         f"IF({SUM_RAW}-{SUM_MIN}<=0,0,({POOL}-{SUM_MIN})/({SUM_RAW}-{SUM_MIN}))))",
         PCT, False, "100% = the pool covers every guideline target. Below that, targets are trimmed to fit — floors are protected first, and only if the pool cannot even cover the floors is everything scaled together."),
    ]
    for i, (label, val, fmt, is_input, note) in enumerate(deriv, start=1):
        a = ws.cell(row=d0 + i, column=1, value=label); a.font = font(bold=True); a.border = BOX
        b = ws.cell(row=d0 + i, column=2, value=val); b.border = BOX; b.number_format = fmt
        b.font = BLUE_INPUT if is_input else font()
        if is_input:
            b.fill = fill("FFFF99")
        n = ws.cell(row=d0 + i, column=3, value=note); n.font = font(italic=True, color="595959", size=9)
    warn = ws.cell(row=d0 + N_DERIV + 1, column=1,
                   value=f'=IF({POOL}<{SUM_MIN},"⚠ The floors alone (" & TEXT({SUM_MIN},"$#,##0") & ") cost more than the discretionary pool (" & TEXT({POOL},"$#,##0") & "). Every cap below is scaled to what is actually available — to fund the benchmarks, fixed bills or the savings rate have to come down.",'
                         f'IF({SCALE}<1,"Guideline targets exceed the pool, so caps above their floors are trimmed to " & TEXT({SCALE},"0%") & ". Floors are funded in full.",'
                         f'"The pool covers every guideline target; " & TEXT({POOL}-{SUM_RAW},"$#,##0") & " a month is left over and rolls into savings."))')
    warn.font = font(italic=True, color="833C00")

    # ------------------------------------------------------------ flexible
    headers = list(BUD_HEADERS)
    headers[7] = "Actual (month in focus)"
    _header_row(ws, flex_header, headers, BUD_WIDTHS)
    ws.row_dimensions[flex_header].height = 28
    for i, cat in enumerate(flex_cats):
        r = flex_first + i
        gr = grow[cat]
        avg_terms = [f'SUMIFS({R_EXP},{R_MONTH},"{m}",{R_CAT},"{cat}"{oc})' for m in hist]
        vals = [
            cat, "Flexible",
            f'=TEXT($C${gr},"0%") & " of take-home" & IF($D${gr}>0," · floor " & TEXT($D${gr}*{SPLIT},"$#,##0") & " (your share)","") & " · " & {"'" + SHEET_BUDGETS + "'"}!$E${gr}',
            f"=MAX({INCOME}*$C${gr},$E{r})",
            f"=$D${gr}*{SPLIT}",
            f"=IF({POOL}>={SUM_RAW},$D{r},IF({POOL}>={SUM_MIN},$E{r}+($D{r}-$E{r})*{SCALE},$D{r}*{SCALE}))",
            f"=$F{r}*{PACE}",
            f'=SUMIFS({R_EXP},{R_MONTH},{sel},{R_CAT},"{cat}"{oc})',
            f"=$H{r}-$G{r}",
            f"=IF($F{r}=0,0,$H{r}/$F{r})",
            f'=IF($F{r}=0,"no cap",IF($H{r}>$G{r},"OVER",IF($H{r}>=0.9*$G{r},"At the line","Under")))',
            _avg_formula(avg_terms, nh),
        ]
        _write_bud_row(ws, r, vals, cfg["categories"][cat])
    _bud_total(ws, flex_total, flex_first, flex_last, "Flexible total", ("D", "E", "F", "G", "H", "I", "L"))

    # --------------------------------------------------------------- fixed
    fx_note = ("Fixed commitments — what actually left these accounts (not discretionary; shown so a spike is visible)"
               if not household else "Fixed commitments (not discretionary — shown so a spike is visible)")
    ws.cell(row=fix_header - 1, column=1, value=fx_note).font = font(bold=True, size=11)
    _header_row(ws, fix_header, headers, BUD_WIDTHS)
    ws.row_dimensions[fix_header].height = 28
    for i, cat in enumerate(fixed_cats):
        r = fix_first + i
        avg_terms = [f'SUMIFS({R_EXP},{R_MONTH},"{m}",{R_CAT},"{cat}"{oc})' for m in hist]
        vals = [
            cat, "Fixed", "your own average over the months listed above", None, None,
            f"=$L{r}", f"=$F{r}*{PACE}",
            f'=SUMIFS({R_EXP},{R_MONTH},{sel},{R_CAT},"{cat}"{oc})',
            f"=$H{r}-$G{r}", f"=IF($F{r}=0,0,$H{r}/$F{r})",
            f'=IF($F{r}=0,"none",IF($H{r}>$G{r}*1.05,"Spike",IF($H{r}<$G{r}*0.5,"Not yet paid","Normal")))',
            _avg_formula(avg_terms, nh),
        ]
        _write_bud_row(ws, r, vals, cfg["categories"][cat])
    _bud_total(ws, fix_total, fix_first, fix_last, "Fixed total", ("F", "G", "H", "I", "L"))

    # ------------------------------------------------------------- summary
    ws.cell(row=summary, column=1, value="Month verdict").font = font(bold=True, size=11)
    items = [
        ("Flexible spending vs pace target", f"=$H${flex_total}-$G${flex_total}", MONEY),
        ("Flexible spending vs full-month cap", f"=$H${flex_total}-$F${flex_total}", MONEY),
        ("Everything vs plan (flexible + fixed)", f"=($H${flex_total}+$H${fix_total})-($F${flex_total}+$F${fix_total})", MONEY),
        ("Left to spend this month", f"=MAX(0,$F${flex_total}-$H${flex_total})", MONEY),
        ("Categories over pace", f'=COUNTIF($K${flex_first}:$K${flex_last},"OVER")', "0"),
    ]
    for i, (label, formula, fmt) in enumerate(items, start=1):
        a = ws.cell(row=summary + i, column=1, value=label); a.font = font(bold=True); a.border = BOX
        b = ws.cell(row=summary + i, column=2, value=formula); b.number_format = fmt; b.border = BOX; b.font = font()
    v = ws.cell(row=summary + 1, column=3,
                value=f'=IF($H${flex_total}>$G${flex_total},"Ahead of pace by " & TEXT($H${flex_total}-$G${flex_total},"$#,##0") & " — see the OVER rows above.",'
                      f'"On or under pace, with " & TEXT($F${flex_total}-$H${flex_total},"$#,##0") & " of the month'"'"'s flexible budget still unspent.")')
    v.font = font(italic=True, color="595959")
    ws.cell(row=summary + 2, column=3, value="pace target = cap × share of the month elapsed, so a part-month is judged fairly").font = font(italic=True, color="595959", size=9)

    for first, last in ((flex_first, flex_last), (fix_first, fix_last)):
        ws.conditional_formatting.add(f"I{first}:I{last}", CellIsRule(operator="greaterThan", formula=["0"], fill=fill("FFC7CE"), font=font(color="9C0006")))
        ws.conditional_formatting.add(f"I{first}:I{last}", CellIsRule(operator="lessThanOrEqual", formula=["0"], fill=fill("C6EFCE"), font=font(color="006100")))
    return g["end"]


def _write_bud_row(ws, r, vals, cat_color):
    for ci, v in enumerate(vals, start=1):
        cell = ws.cell(row=r, column=ci, value=v)
        cell.font = font(); cell.border = BOX
        if ci in (4, 5, 6, 7, 8, 9, 12):
            cell.number_format = MONEY
        elif ci == 10:
            cell.number_format = PCT
    ws.cell(row=r, column=1).fill = fill(cat_color)
    ws.cell(row=r, column=3).font = font(size=9, color="595959")
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="center")
    ws.cell(row=r, column=11).alignment = Alignment(horizontal="center")


def _bud_total(ws, row, first, last, label, cols):
    ws.cell(row=row, column=1, value=label)
    for c in cols:
        cell = ws.cell(row=row, column=ord(c) - 64, value=f"=SUM({c}{first}:{c}{last})")
        cell.number_format = MONEY
    ws.cell(row=row, column=10, value=f"=IF(F{row}=0,0,H{row}/F{row})").number_format = PCT
    for c in range(1, len(BUD_HEADERS) + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = font(bold=True); cell.fill = fill("F2F2F2"); cell.border = BOX


def _income_share(store, owner, hist) -> float:
    """Owner's share of household recurring income over the averaged months (used as the default split)."""
    tot = own = 0.0
    for t in store.ledger["transactions"].values():
        if t.get("removed") or t.get("pending") or t["amount"] >= 0:
            continue
        if (t.get("category_override") or t.get("category")) != "Income":
            continue
        if t["account_id"] not in store.state["accounts"] or store.account_hidden(t["account_id"]):
            continue
        if month_label(t["date"]) not in hist:
            continue
        amt = -t["amount"]
        tot += amt
        if store.account_owner(t["account_id"]) == owner:
            own += amt
    return (own / tot) if tot else 0.0


def _find_goal_account(store, match: str):
    match = (match or "").lower().strip()
    if not match:
        return None
    for aid, a in store.state["accounts"].items():
        hay = " ".join(filter(None, [a.get("name"), a.get("official_name"), store.account_label(aid)])).lower()
        if match in hay:
            return {"id": aid, **a}
    return None


# ----------------------------------------------------------------- Accounts
def _write_accounts(ws, store, cfg):
    owners = cfg["owners"]
    ws["A1"] = "Linked accounts"
    ws["A1"].font = font(bold=True, size=14)
    ws["A2"] = ("You can edit Owner, Label and Hidden here (or in config.json). Changes are picked up on the next sync. "
                "Balances are what the bank reported at the last sync.")
    ws["A2"].font = font(italic=True, color="595959", size=9)
    headers = ["Owner", "Label", "Institution", "Type", "Last 4", "Current Balance", "Available", "As of", "Item", "Hidden?", "Account ID"]
    _header_row(ws, 4, headers, [12, 34, 18, 12, 8, 16, 14, 18, 26, 9, 40])
    dv = DataValidation(type="list", formula1=f"={SHEET_LISTS}!$B$2:$B${1 + len(owners)}", allow_blank=False)
    ws.add_data_validation(dv)
    owner_order = {o: i for i, o in enumerate(owners)}
    ids = sorted(store.state["accounts"], key=lambda a: (owner_order.get(store.account_owner(a), 99), store.account_label(a)))
    r = 5
    for aid in ids:
        a = store.state["accounts"][aid]
        owner = store.account_owner(aid)
        inst = store.state["items"].get(a.get("item_id", ""), {}).get("institution", "")
        vals = [owner, store.account_label(aid), inst, f"{a.get('type') or ''}/{a.get('subtype') or ''}", a.get("mask"),
                a.get("balance_current"), a.get("balance_available"), a.get("balance_as_of"), a.get("item_id"),
                "yes" if store.account_hidden(aid) else "", aid]
        for ci, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=ci, value=v)
            cell.font = font(); cell.border = BOX
        ws.cell(row=r, column=1).fill = fill(owners.get(owner, {}).get("row_fill", "F2F2F2"))
        ws.cell(row=r, column=6).number_format = MONEY
        ws.cell(row=r, column=7).number_format = MONEY
        ws.cell(row=r, column=11).font = font(color="808080", size=8)
        dv.add(ws.cell(row=r, column=1))
        r += 1
    if not ids:
        ws.cell(row=5, column=1, value="No accounts linked yet.").font = font(italic=True)
    ws.freeze_panes = "A5"


# ----------------------------------------------------------------- Sync Log
LOG_HEADERS = ["Timestamp", "Mode", "Items synced", "Added", "Modified", "Removed", "Pending (hidden)", "Notes"]


def _append_log(ws, entry: dict):
    if ws.max_row < 1 or ws.cell(row=1, column=1).value != "Timestamp":
        _header_row(ws, 1, LOG_HEADERS, [20, 12, 12, 9, 9, 9, 16, 80])
    r = ws.max_row + 1
    vals = [entry.get("timestamp"), entry.get("mode"), entry.get("items"), entry.get("added"), entry.get("modified"),
            entry.get("removed"), entry.get("pending"), entry.get("notes")]
    for ci, v in enumerate(vals, start=1):
        cell = ws.cell(row=r, column=ci, value=v)
        cell.font = font()
        cell.alignment = Alignment(wrap_text=(ci == 8), vertical="top")


# --------------------------------------------------------------------- save
def _safe_save(wb, path: Path, keep: int):
    if path.exists():
        bdir = path.parent / "backups"
        bdir.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(path, bdir / f"{path.stem}_{stamp}{path.suffix}")
        old = sorted(bdir.glob(f"{path.stem}_*{path.suffix}"))
        for p in old[:-keep] if keep > 0 else []:
            p.unlink(missing_ok=True)
    tmp = path.with_name(path.stem + ".saving.xlsx")
    wb.save(tmp)
    try:
        tmp.replace(path)
    except PermissionError:
        tmp.unlink(missing_ok=True)
        raise PermissionError(f"'{path.name}' is open in Excel. Close it and run the sync again.")
