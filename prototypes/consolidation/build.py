"""Phase 0 proof: group consolidation as the add-in will build it, on a fictional group shaped like HFG's.

    python prototypes/consolidation/build.py [out.xlsx]

Kelvin (6 October 2026): the add-in carries the whole consolidation feature set, applied to
how HFG works: a lot of intergroup transactions that carry margin, at-cost on-charges that pass
through the "Intergroup AP/AR - [Entity]" netting accounts, and costs moved into a development
LP's work in progress.

The workbook holds what Home Hub (or each entity's saved version) hands over: the entity
register, the group chart, each entity's trial balance by year and the intercompany register
from matching. Everything after that is formulas, so it consolidates without the add-in:

- every intercompany row is tagged with the lowest group holding both sides, and eliminated
  in that group and every group above it; below it, it is a related party;
- trading the buyer expenses comes out in full; trading the buyer capitalises comes out at the
  seller's cost, and its margin comes out of the asset (WIP, then investment property after a
  portfolio sale) for as long as the asset stays in the group, released when homes are sold
  outside it; at-cost pass-throughs and moves at cost eliminate balances only;
- investments are eliminated against capital; non-controlling interests are worked out once at
  each node from its net assets and carried up, taking their share of margin on sales from a
  partly owned seller;
- checks catch a break between the two sides, a netting account left with a balance, an
  investment that is not the parent's share of capital, eliminations that do not balance, and
  NCI that does not tie;
- Group structure shows the entities as a tree (share held, owned by the top, outside
  investors, status, member from, the groups each rolls into) and each group's surplus and net
  assets added up the tree: members' own figures, eliminations made in the groups below it and
  eliminations made in it. The register stays in tree order; group.add_entity places a planned
  entity under its parent as Add entity will, and the checks catch a register out of tree order
  or figures before an entity joins.

group.py builds the same group by re-recording every event as each group sees it; the tests
compare the two.
"""

from __future__ import annotations

import importlib.util
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "models"))
import group as G  # noqa: E402
import navigation  # noqa: E402

TEXT = "404040"
F_BODY = Font(name="Segoe UI", size=9, color=TEXT)
F_BOLD = Font(name="Segoe UI", size=9, color=TEXT, bold=True)
F_HEAD = Font(name="Segoe UI", size=10, color=TEXT, bold=True)
F_NOTE = Font(name="Segoe UI", size=9, color="808080", italic=True)
F_CHECK = Font(name="Segoe UI", size=9, color="9C0006")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
UNDER = Border(bottom=Side(style="thin", color=TEXT))
TOP = Border(top=Side(style="thin", color=TEXT))
NUM = '#,##0.0;(#,##0.0);"-"'
DATE = "d mmmm yyyy"
DOT = '"●";"●";""'
PCT = "0%"
RIGHT = Alignment(horizontal="right")
CENTRE = Alignment(horizontal="center")
IND = Alignment(horizontal="left", indent=1)   # text beside a right-aligned code
Y = G.PERIODS
FIRST = 10                                        # column J: the first year on every sheet
NOTE = "Fictional group shaped like HFG's ($000)"
HEADINGS: dict[str, list[tuple[int, str]]] = {}
TIERS = sorted(G.TIERS, key=lambda t: -len(G.TIER_MEMBERS[t[0]]))      # largest group first; the lowest common group is the last flagged
K = len(TIERS)


def put(ws, r, c, v, font=None, fmt=None, fill=None, border=None, align=None):
    cell = ws.cell(r, c)
    cell.value = v
    cell.font = font or F_BODY
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if border:
        cell.border = border
    if align:
        cell.alignment = align
    return cell


def q(ws) -> str:
    return "'" + ws.title.replace("'", "''") + "'"


def ref(ws, c1, r1, c2=None, r2=None) -> str:
    a = f"{q(ws)}!${L(c1)}${r1}"
    return a if c2 is None else a + f":${L(c2)}${r2 if r2 is not None else r1}"


def define(wb, nm, text):
    wb.defined_names[nm] = DefinedName(nm, attr_text=text)


def sheet(wb, title, purpose):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    put(ws, 1, 2, title, F_HEAD)
    put(ws, 2, 2, purpose)
    put(ws, 3, 2, NOTE, F_NOTE)
    ws.column_dimensions["A"].width = 2.5
    return ws


def section(ws, r, text, toc=None, last_col=12):
    HEADINGS.setdefault(ws.title, []).append((r, toc or text))
    put(ws, r, 2, text, F_HEAD, border=UNDER)
    for c in range(3, last_col + 1):
        ws.cell(r, c).border = UNDER


def year_heads(ws, r, c0, label=None):
    for p in range(Y):
        put(ws, r, c0 + p, G.YEAR_LABELS[p], F_BOLD, align=RIGHT, border=UNDER)
    if label:
        put(ws, r - 1, c0, label, F_BOLD)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


# ----------------------------------------------------------------------------------- inputs
def entities_sheet(wb, tb):
    ws = sheet(wb, "Entities", "Entity register: parent, share held, GST, NCI nodes, status, and the groups each entity consolidates into")
    widths(ws, {"B": 7, "C": 19, "D": 8, "E": 13, "F": 13, "G": 46, "H": 9, "I": 9, "J": 15, "K": 8, "L": 11, "M": 8, "N": 8})
    section(ws, 5, "Entity register", last_col=14 + K + 5 * Y + 2)
    heads = ["Code", "Name", "Parent", "Held by parent", "GST registered", "Role", "NCI node", "Status", "Member from",
             "Parent", "Grandparent", "Level 3", "Level 4"]
    A0 = 11                                                       # first ancestor column (K)
    for j, h in enumerate(heads):
        put(ws, 7, 2 + j, h, F_BOLD, border=UNDER,
            align=RIGHT if j in (0, 2, 3, 8, 9, 10, 11, 12) else (IND if j == 1 else (CENTRE if j in (4, 6, 7) else None)))
    put(ws, 6, A0, "Ancestors (worked out)", F_NOTE)
    c_in = 2 + len(heads)                                        # membership columns start here
    put(ws, 6, c_in, "In group (1) or not (0), by head", F_NOTE)
    for k, (head, name) in enumerate(TIERS):
        put(ws, 7, c_in + k, head, F_BOLD, "0", border=UNDER, align=RIGHT)
        ws.column_dimensions[L(c_in + k)].width = 8
    c_tb = c_in + K
    put(ws, 6, c_tb, "Trial balance total (must be nil)", F_NOTE)
    for p in range(Y):
        put(ws, 7, c_tb + p, G.YEAR_LABELS[p], F_BOLD, border=UNDER, align=RIGHT)
        ws.column_dimensions[L(c_tb + p)].width = 9
    first, last = 8, 8 + len(G.ENTITIES) - 1
    anc = [L(A0 + j) for j in range(4)]
    for i, (code, name, parent, held, gst, role) in enumerate(G.ENTITIES):
        r = first + i
        put(ws, r, 2, code, fmt="0", fill=FILL_IN)
        put(ws, r, 3, name, fill=FILL_IN, align=IND)
        put(ws, r, 4, parent, fmt="0", fill=FILL_IN)
        put(ws, r, 5, held, fmt=PCT, fill=FILL_IN)
        put(ws, r, 6, "Yes" if gst else "No", fill=FILL_IN, align=CENTRE)
        put(ws, r, 7, role, fill=FILL_IN)
        put(ws, r, 8, "Yes" if code in G.NODES else "No", fill=FILL_IN, align=CENTRE)
        put(ws, r, 9, G.STATUS[code], fill=FILL_IN, align=CENTRE)
        put(ws, r, 10, datetime.strptime(G.YEAR_START[G.START[code] - 1], "%d %B %Y"), fmt=DATE, fill=FILL_IN)
        put(ws, r, A0, f'=IF($D{r}="","",$D{r})', fmt="0")
        for j in range(3):                                      # each ancestor is the parent of the one before
            prev = anc[j]
            up = f"INDEX($D${first}:$D${last},MATCH({prev}{r},$B${first}:$B${last},0))"
            put(ws, r, A0 + 1 + j, f'=IF({prev}{r}="","",IF(IFERROR({up},0)=0,"",{up}))', fmt="0")
        for k in range(K):
            h = f"{L(c_in + k)}$7"
            test = ",".join(f"${c}{r}={h}" for c in ["B"] + anc)
            put(ws, r, c_in + k, f"=IF(OR({test}),1,0)", fmt="0", align=RIGHT)
        for p in range(Y):
            put(ws, r, c_tb + p, f"=SUMPRODUCT((Data_Ent=$B{r})*INDEX(Data_Val,0,{p + 1}))", fmt=NUM)
    for col, lst, title in (("D", "Ent_Codes", "Parent"), ("F", "LU_YesNo", "GST registered"), ("H", "LU_YesNo", "NCI node"),
                            ("I", "LU_Status", "Status")):
        dv = DataValidation(type="list", formula1=lst, allow_blank=col == "D")
        dv.error, dv.errorTitle = "Pick from the list", title
        ws.add_data_validation(dv)
        dv.add(f"{col}{first}:{col}{last}")
    define(wb, "Ent_Codes", ref(ws, 2, first, 2, last))
    define(wb, "Ent_Names", ref(ws, 3, first, 3, last))
    define(wb, "Ent_Parent", ref(ws, 4, first, 4, last))
    define(wb, "Ent_Held", ref(ws, 5, first, 5, last))
    define(wb, "Ent_GST", ref(ws, 6, first, 6, last))
    define(wb, "Ent_Status", ref(ws, 9, first, 9, last))
    define(wb, "Ent_From", ref(ws, 10, first, 10, last))
    define(wb, "Ent_Anc", ref(ws, A0, first, A0 + 3, last))
    define(wb, "Ent_In", ref(ws, c_in, first, c_in + K - 1, last))
    define(wb, "Ent_TB_Sum", ref(ws, c_tb, first, c_tb + Y - 1, last))
    define(wb, "Tier_Heads", ref(ws, c_in, 7, c_in + K - 1, 7))
    # Groups that consolidate
    r0 = last + 3
    section(ws, r0, "Groups that consolidate", last_col=14 + K + 5 * Y + 2)
    for j, h in enumerate(["Head", "Group", "Members", "Note"]):
        put(ws, r0 + 1, 2 + j, h, F_BOLD, border=UNDER, align=IND if j == 1 else (RIGHT if j in (0, 2) else IND))
    for k, (head, name) in enumerate(TIERS):
        r = r0 + 2 + k
        subs = len(G.TIER_MEMBERS[head]) > 1
        put(ws, r, 2, head, fmt="0", fill=FILL_IN)
        put(ws, r, 3, name, fill=FILL_IN, align=IND)
        put(ws, r, 4, f"=SUM(INDEX(Ent_In,0,{k + 1}))", fmt="0")
        note = ("Sub-group with outside investors (NCI node)" if subs else "Node only: its NCI is worked out here and carried up") \
            if head in G.NODES else ""
        put(ws, r, 5, note, F_NOTE, align=IND)
    define(wb, "Tier_Names", ref(ws, 3, r0 + 2, 3, r0 + 1 + K))
    define(wb, "Tier_Sizes", ref(ws, 4, r0 + 2, 4, r0 + 1 + K))
    put(ws, r0 + 3 + K, 2, "A pair is eliminated in the lowest group holding both sides and every group above it; below that it is a "
                           "related party. Groups are listed largest first, so the lowest common group is the last one flagged.", F_NOTE)
    put(ws, r0 + 4 + K, 2, "Home Hub is the master for actual entities. Add entity puts a planned one (a future LP, a new fund) under its "
                           "parent, in tree order, until Home Hub has it.", F_NOTE)
    ws.freeze_panes = "D8"
    return ws, c_in


def accounts_sheet(wb):
    ws = sheet(wb, "Accounts", "Group chart of accounts: every entity's accounts are mapped to these (debit positive)")
    widths(ws, {"B": 7, "C": 44, "D": 12, "E": 10})
    section(ws, 5, "Group chart", last_col=5)
    for j, h in enumerate(["Code", "Account", "Class", "Shown as"]):
        put(ws, 6, 2 + j, h, F_BOLD, border=UNDER)
    for i, (code, name, cls) in enumerate(G.ACCOUNTS):
        r = 7 + i
        put(ws, r, 2, code, fmt="0", fill=FILL_IN)
        put(ws, r, 3, name, fill=FILL_IN, align=IND)
        put(ws, r, 4, cls, fill=FILL_IN)
        put(ws, r, 5, 1 if cls in ("asset",) else -1, fmt="0")
    last = 6 + len(G.ACCOUNTS)
    define(wb, "Acct_Codes", ref(ws, 2, 7, 2, last))
    define(wb, "Acct_Names", ref(ws, 3, 7, 3, last))
    define(wb, "Acct_Class", ref(ws, 4, 7, 4, last))
    put(ws, last + 2, 2, "2900 Intercompany differences, 3300 Non-controlling interests and 5900 Surplus attributable to NCI are used "
                         "by the consolidation only. Retained surplus is held at the start of the year; the year's surplus is the P&L.", F_NOTE)
    return ws


def sites_sheet(wb):
    ws = sheet(wb, "Sites", "Where each site's asset sits at each year end, and how many homes are still held")
    widths(ws, {"B": 6, "C": 44, "H": 7})
    section(ws, 5, "Sites", last_col=FIRST + 3 * Y)
    put(ws, 6, 2, "Site", F_BOLD, border=UNDER)
    put(ws, 6, 3, "Name", F_BOLD, border=UNDER)
    put(ws, 6, 8, "Homes", F_BOLD, border=UNDER, align=RIGHT)
    blocks = [("Held by (entity)", 0), ("In account", Y), ("Homes held at year end", 2 * Y)]
    for label, off in blocks:
        put(ws, 5, FIRST + off, label, F_BOLD)
        for p in range(Y):
            put(ws, 6, FIRST + off + p, G.YEAR_LABELS[p], F_BOLD, border=UNDER, align=RIGHT)
            ws.column_dimensions[L(FIRST + off + p)].width = 8
    for i, (code, s) in enumerate(G.SITES.items()):
        r = 7 + i
        put(ws, r, 2, code, fill=FILL_IN)
        put(ws, r, 3, s["name"], fill=FILL_IN, align=IND)
        put(ws, r, 8, s["units"], fmt="0", fill=FILL_IN)
        for p in range(Y):
            put(ws, r, FIRST + p, s["holder"][p], fmt="0", fill=FILL_IN)
            put(ws, r, FIRST + Y + p, s["account"][p], fmt="0", fill=FILL_IN)
            put(ws, r, FIRST + 2 * Y + p, s["held"][p], fmt="0", fill=FILL_IN)
    last = 6 + len(G.SITES)
    define(wb, "Site_Codes", ref(ws, 2, 7, 2, last))
    define(wb, "Site_Units", ref(ws, 8, 7, 8, last))
    define(wb, "Site_Holder", ref(ws, FIRST, 7, FIRST + Y - 1, last))
    define(wb, "Site_Acct", ref(ws, FIRST + Y, 7, FIRST + 2 * Y - 1, last))
    define(wb, "Site_Held", ref(ws, FIRST + 2 * Y, 7, FIRST + 3 * Y - 1, last))
    put(ws, last + 2, 2, "Margin on a site stays out of the asset while the asset is held inside the group; it is released as homes are "
                         "sold outside the group, or when the asset leaves the group.", F_NOTE)
    return ws


def data_sheet(wb, tb, c_in):
    ws = sheet(wb, "Entity data", "Each entity's trial balance by group account and year, as Home Hub (or the entity's saved version) hands it over")
    widths(ws, {"B": 7, "C": 7, "D": 40, "H": 6})
    section(ws, 5, "Trial balances by entity", toc="Trial balances", last_col=FIRST + Y + K)
    put(ws, 6, 2, "Entity", F_BOLD, border=UNDER)
    put(ws, 6, 3, "Account", F_BOLD, border=UNDER)
    put(ws, 6, 4, "Name", F_BOLD, border=UNDER)
    year_heads(ws, 6, FIRST)
    put(ws, 5, FIRST + Y, "In group", F_NOTE)
    for k, (head, name) in enumerate(TIERS):
        put(ws, 6, FIRST + Y + k, head, F_BOLD, "0", border=UNDER, align=RIGHT)
        ws.column_dimensions[L(FIRST + Y + k)].width = 7
    for p in range(Y):
        ws.column_dimensions[L(FIRST + p)].width = 10
    r = 7
    first = r
    for code in G.CODES:
        put(ws, r, 2, f"{code}  {G.NAME[code]}", F_BOLD)
        r += 1
        for acct, name, cls in G.ACCOUNTS:
            if cls == "attribution" or acct in (G.NCI, G.DIFF):
                continue
            put(ws, r, 2, code, fmt="0")
            put(ws, r, 3, acct, fmt="0")
            put(ws, r, 4, f"=INDEX(Acct_Names,MATCH(C{r},Acct_Codes,0))", align=IND)
            for p in range(Y):
                v = round(tb[(code, acct)][p], 9)
                put(ws, r, FIRST + p, v if abs(v) > 1e-12 else 0, fmt=NUM, fill=FILL_IN)
            for k in range(K):
                put(ws, r, FIRST + Y + k, f"=INDEX(Ent_In,MATCH($B{r},Ent_Codes,0),{k + 1})", fmt="0")
            r += 1
        r += 1
    last = r - 1
    define(wb, "Data_Ent", ref(ws, 2, first, 2, last))
    define(wb, "Data_Acct", ref(ws, 3, first, 3, last))
    define(wb, "Data_Val", ref(ws, FIRST, first, FIRST + Y - 1, last))
    define(wb, "Data_In", ref(ws, FIRST + Y, first, FIRST + Y + K - 1, last))
    ws.freeze_panes = f"E{first}"
    return ws, first, last


IC_COLS = ["Id", "Year", "Type", "Seller", "Buyer", "Description", "Site", "Seller account", "Buyer account",
           "Seller cost account", "Seller amount", "Buyer amount", "GST", "Margin"]


def intercompany_sheet(wb, register):
    ws = sheet(wb, "Intercompany", "Intercompany register from Home Hub's matching: every pair, its type and both sides, "
                                   "and the group it is eliminated in")
    w = [5, 6, 18, 8, 8, 40, 6, 13, 13, 17, 13, 13, 8, 9]
    for j, x in enumerate(w):
        ws.column_dimensions[L(2 + j)].width = x
    c = 2 + len(IC_COLS)
    computed = ["Type no", "Difference", "Status", "Eliminated in"] + [f"In {h}" for h, _ in TIERS] + \
               [f"Upstream {n}" for n in G.NODES] + [f"Related {h}" for h, _ in TIERS]
    section(ws, 5, "Intercompany register", toc="Register", last_col=c + len(computed))
    for j, h in enumerate(IC_COLS):
        put(ws, 7, 2 + j, h, F_BOLD, border=UNDER,
            align=CENTRE if j in (1, 3, 4, 6) else (RIGHT if j in (0, 7, 8, 9, 10, 11, 12, 13) else IND))
    put(ws, 6, c, "Worked out", F_NOTE)
    for j, h in enumerate(computed):
        put(ws, 7, c + j, h, F_BOLD, border=UNDER, align=CENTRE if j == 2 else (IND if j == 3 else RIGHT))
        ws.column_dimensions[L(c + j)].width = 18 if j == 3 else (14 if j >= 4 + K else 10)
    first = 8
    for i, row in enumerate(register):
        r = first + i
        vals = [i + 1, row["period"], row["type"], row["seller"], row["buyer"], row["desc"], row["site"] or "",
                row["seller_acct"], row["buyer_acct"], row["cost_acct"], row["seller_amt"], row["buyer_amt"], row["gst"],
                row["margin"]]
        fmts = ["0", "0", None, "0", "0", None, None, "0", "0", "0", NUM, NUM, NUM, NUM]
        for j, (v, f) in enumerate(zip(vals, fmts)):
            put(ws, r, 2 + j, v, fmt=f, fill=FILL_IN if j else None,
                align=CENTRE if j in (1, 3, 4, 6) else (IND if j in (2, 5) else None))
        d = L(c + 1)
        put(ws, r, c, f"=MATCH($D{r},LU_IC_Types,0)", fmt="0")
        put(ws, r, c + 1, f"=ROUND($L{r}-$M{r},6)", fmt=NUM)
        put(ws, r, c + 2, f'=IF(ABS({d}{r})<0.001,"Matched","Break")', align=CENTRE)
        flags = [L(c + 4 + k) for k in range(K)]
        put(ws, r, c + 3, "=" + _lowest(flags, r), align=IND)
        for k in range(K):
            put(ws, r, c + 4 + k, f"=INDEX(Ent_In,MATCH($E{r},Ent_Codes,0),{k + 1})*INDEX(Ent_In,MATCH($F{r},Ent_Codes,0),{k + 1})",
                fmt="0")
        for j, node in enumerate(G.NODES):
            k = [h for h, _ in TIERS].index(node) + 1
            put(ws, r, c + 4 + K + j, f"=INDEX(Ent_In,MATCH($E{r},Ent_Codes,0),{k})*(1-INDEX(Ent_In,MATCH($F{r},Ent_Codes,0),{k}))",
                fmt="0")
        for k in range(K):
            col = c + 4 + K + len(G.NODES) + k
            put(ws, r, col, f"=ABS(INDEX(Ent_In,MATCH($E{r},Ent_Codes,0),{k + 1})-INDEX(Ent_In,MATCH($F{r},Ent_Codes,0),{k + 1}))",
                fmt="0")
    last = first + len(register) - 1
    cols = {h: 2 + j for j, h in enumerate(IC_COLS)}
    for nm, h in (("IC_Id", "Id"), ("IC_Year", "Year"), ("IC_Type", "Type"), ("IC_Seller", "Seller"), ("IC_Buyer", "Buyer"),
                  ("IC_Site", "Site"), ("IC_Seller_Acct", "Seller account"), ("IC_Buyer_Acct", "Buyer account"),
                  ("IC_Cost_Acct", "Seller cost account"), ("IC_Seller_Amt", "Seller amount"),
                  ("IC_Buyer_Amt", "Buyer amount"), ("IC_GST", "GST"), ("IC_Margin", "Margin")):
        define(wb, nm, ref(ws, cols[h], first, cols[h], last))
    define(wb, "IC_Type_No", ref(ws, c, first, c, last))
    define(wb, "IC_Status", ref(ws, c + 2, first, c + 2, last))
    define(wb, "IC_Group", ref(ws, c + 3, first, c + 3, last))
    define(wb, "IC_In", ref(ws, c + 4, first, c + 3 + K, last))
    define(wb, "IC_Up", ref(ws, c + 4 + K, first, c + 3 + K + len(G.NODES), last))
    define(wb, "IC_Related", ref(ws, c + 4 + K + len(G.NODES), first, c + 3 + 2 * K + len(G.NODES), last))
    r = last + 2
    section(ws, r, "Types", last_col=c + len(computed))
    for i, (name, rule) in enumerate(G.TYPES.items()):
        put(ws, r + 1 + i, 3, name, F_BOLD, align=IND)
        put(ws, r + 1 + i, 7, rule, align=IND)
    ws.freeze_panes = f"E{first}"
    return ws, first, last, c


def _lowest(flags, r):
    """The smallest group holding both sides. Groups run largest first and an IF chain tests its outermost
    condition first, so building it from the largest group outward puts the smallest group's test outermost."""
    out = '"None"'
    for k in range(K):
        out = f"IF({flags[k]}{r}=1,INDEX(Tier_Names,{k + 1}),{out})"
    return out


# ----------------------------------------------------------------------------------- consolidation workings
def investments_sheet(wb, investments):
    ws = sheet(wb, "Investments", "Each parent's investment eliminated against the sub's capital; the outside investors' share becomes NCI")
    widths(ws, {"B": 7, "C": 7, "D": 30, "G": 9, "H": 8, "I": 10})
    section(ws, 5, "Investments in group entities", toc="Investments", last_col=FIRST + 4 * Y + K)
    for c, h in ((2, "Parent"), (3, "Sub"), (4, "Sub's name"), (7, "Made in"), (8, "Held"), (9, "Cost")):
        put(ws, 7, c, h, F_BOLD, border=UNDER, align=RIGHT if c != 4 else IND)
    blocks = [("Sub's capital", 0), ("Parent's share of capital", Y), ("Cost held", 2 * Y), ("Cost less share (must be nil)", 3 * Y)]
    for label, off in blocks:
        put(ws, 6, FIRST + off, label, F_BOLD)
        for p in range(Y):
            put(ws, 7, FIRST + off + p, G.YEAR_LABELS[p], F_BOLD, border=UNDER, align=RIGHT)
            ws.column_dimensions[L(FIRST + off + p)].width = 9
    c_in = FIRST + 4 * Y
    put(ws, 6, c_in, "In group", F_NOTE)
    for k, (head, _) in enumerate(TIERS):
        put(ws, 7, c_in + k, head, F_BOLD, "0", border=UNDER, align=RIGHT)
        ws.column_dimensions[L(c_in + k)].width = 7
    first = 8
    for i, (parent, sub, cost, made) in enumerate(investments):
        r = first + i
        put(ws, r, 2, parent, fmt="0", fill=FILL_IN)
        put(ws, r, 3, sub, fmt="0", fill=FILL_IN)
        put(ws, r, 4, f"=INDEX(Ent_Names,MATCH(C{r},Ent_Codes,0))", align=IND)
        put(ws, r, 7, G.YEAR_LABELS[made - 1], fill=FILL_IN, align=RIGHT)
        put(ws, r, 8, f"=INDEX(Ent_Held,MATCH(C{r},Ent_Codes,0))", fmt=PCT)
        put(ws, r, 9, cost, fmt=NUM, fill=FILL_IN)
        for p in range(Y):
            put(ws, r, FIRST + p, f"=-SUMPRODUCT((Data_Ent=$C{r})*(Data_Acct=3000)*INDEX(Data_Val,0,{p + 1}))", fmt=NUM)
            put(ws, r, FIRST + Y + p, f"=$H{r}*{L(FIRST + p)}{r}", fmt=NUM)
            put(ws, r, FIRST + 2 * Y + p, f"=IF(MATCH($G{r},LU_Years,0)<={p + 1},$I{r},0)", fmt=NUM)
            put(ws, r, FIRST + 3 * Y + p, f"=ROUND({L(FIRST + 2 * Y + p)}{r}-{L(FIRST + Y + p)}{r},6)", fmt=NUM)
        for k in range(K):
            put(ws, r, c_in + k, f"=INDEX(Ent_In,MATCH($B{r},Ent_Codes,0),{k + 1})*INDEX(Ent_In,MATCH($C{r},Ent_Codes,0),{k + 1})",
                fmt="0")
    last = first + len(investments) - 1
    define(wb, "Inv_Parent", ref(ws, 2, first, 2, last))
    define(wb, "Inv_Cost_Y", ref(ws, FIRST + 2 * Y, first, FIRST + 3 * Y - 1, last))
    define(wb, "Inv_Diff", ref(ws, FIRST + 3 * Y, first, FIRST + 4 * Y - 1, last))
    put(ws, last + 2, 2, "Every entity here was set up by the group, so there is no goodwill: a cost that is not the parent's share of "
                         "capital is an error, not goodwill.", F_NOTE)
    ws.freeze_panes = f"E{first}"
    return ws, first, last, c_in


def margin_sheet(wb, register):
    ws = sheet(wb, "Unrealised margin", "Margin on intergroup charges and sales still held in each group's assets, and what has been released")
    widths(ws, {"B": 7, "C": 6, "D": 6, "E": 9, "F": 10})
    n = len(register)
    # Part A: one row per register row
    section(ws, 5, "Margin by register row", toc="By register row", last_col=FIRST + Y + K * Y)
    for c, h in ((2, "Id"), (3, "Site"), (4, "Year"), (5, "Margin"), (6, "Homes then")):
        put(ws, 7, c, h, F_BOLD, border=UNDER, align=CENTRE if c == 3 else RIGHT)
    put(ws, 6, FIRST, "Still held (any group)", F_BOLD)
    year_heads(ws, 7, FIRST)
    for k, (head, name) in enumerate(TIERS):
        c0 = FIRST + Y + k * Y
        put(ws, 6, c0, f"In {name}", F_BOLD)
        for p in range(Y):
            put(ws, 7, c0 + p, G.YEAR_LABELS[p], F_BOLD, border=UNDER, align=RIGHT)
    for c in range(FIRST, FIRST + Y + K * Y):
        ws.column_dimensions[L(c)].width = 9
    first = 8
    for i in range(n):
        r = first + i
        put(ws, r, 2, f"=INDEX(IC_Id,{i + 1})", fmt="0")
        put(ws, r, 3, f"=INDEX(IC_Site,{i + 1})", align=CENTRE)
        put(ws, r, 4, f"=INDEX(IC_Year,{i + 1})", fmt="0")
        put(ws, r, 5, f"=IF(OR(INDEX(IC_Type_No,{i + 1})=2,INDEX(IC_Type_No,{i + 1})=3),INDEX(IC_Margin,{i + 1}),0)", fmt=NUM)
        srow = f"MATCH($C{r},Site_Codes,0)"
        put(ws, r, 6, f"=IFERROR(IF($D{r}=1,INDEX(Site_Units,{srow}),INDEX(Site_Held,{srow},$D{r}-1)),1)", fmt="0")
        for p in range(Y):
            put(ws, r, FIRST + p, f"=IF(OR($E{r}=0,{p + 1}<$D{r}),0,$E{r}*INDEX(Site_Held,{srow},{p + 1})/$F{r})", fmt=NUM)
        for k in range(K):
            for p in range(Y):
                holder_in = f"INDEX(Ent_In,MATCH(INDEX(Site_Holder,{srow},{p + 1}),Ent_Codes,0),{k + 1})"
                put(ws, r, FIRST + Y + k * Y + p,
                    f"=IFERROR({L(FIRST + p)}{r}*INDEX(IC_In,{i + 1},{k + 1})*{holder_in},0)", fmt=NUM)
    last = first + n - 1
    define(wb, "Mg_Site", ref(ws, 3, first, 3, last))
    define(wb, "Mg_Year", ref(ws, 4, first, 4, last))
    define(wb, "Mg_Margin", ref(ws, 5, first, 5, last))
    define(wb, "Mg_Uk", ref(ws, FIRST + Y, first, FIRST + Y + K * Y - 1, last))
    # Part B: by group and site
    r = last + 3
    section(ws, r, "Margin by group and site: still held, added and released", toc="By group and site", last_col=FIRST + 5 * Y)
    put(ws, r + 2, 2, "Group", F_BOLD, border=UNDER, align=RIGHT)
    put(ws, r + 2, 3, "Site", F_BOLD, border=UNDER, align=CENTRE)
    for label, off in (("Still held", 0), ("Added", Y), ("Released", 2 * Y), ("Held in account", 3 * Y), ("Held by", 4 * Y)):
        put(ws, r + 1, FIRST + off, label, F_BOLD)
        for p in range(Y):
            put(ws, r + 2, FIRST + off + p, G.YEAR_LABELS[p], F_BOLD, border=UNDER, align=RIGHT)
            ws.column_dimensions[L(FIRST + off + p)].width = 9
    rows_b = {}
    rr = r + 3
    for k, (head, name) in enumerate(TIERS):
        for site in G.SITES:
            rows_b[(head, site)] = rr
            put(ws, rr, 2, head, fmt="0")
            put(ws, rr, 3, site, align=CENTRE)
            srow = f"MATCH($C{rr},Site_Codes,0)"
            for p in range(Y):
                U = f"{L(FIRST + p)}{rr}"
                put(ws, rr, FIRST + p, f"=SUMPRODUCT((Mg_Site=$C{rr})*INDEX(Mg_Uk,0,{k * Y + p + 1}))", fmt=NUM)
                put(ws, rr, FIRST + Y + p,
                    f"=SUMPRODUCT((Mg_Site=$C{rr})*(Mg_Year={p + 1})*Mg_Margin*INDEX(IC_In,0,{k + 1}))", fmt=NUM)
                prev = f"{L(FIRST + p - 1)}{rr}" if p else "0"
                put(ws, rr, FIRST + 2 * Y + p, f"={prev}+{L(FIRST + Y + p)}{rr}-{U}", fmt=NUM)
                put(ws, rr, FIRST + 3 * Y + p, f"=INDEX(Site_Acct,{srow},{p + 1})", fmt="0")
                put(ws, rr, FIRST + 4 * Y + p, f"=INDEX(Site_Holder,{srow},{p + 1})", fmt="0")
            rr += 1
    # Part C: upstream margin by group and node (sales from a partly owned node's group to the rest of the group)
    r = rr + 2
    section(ws, r, "Upstream margin by group and NCI node", toc="Upstream margin", last_col=FIRST + Y)
    put(ws, r + 1, 2, "Group", F_BOLD, border=UNDER, align=RIGHT)
    put(ws, r + 1, 3, "Node", F_BOLD, border=UNDER)
    year_heads(ws, r + 1, FIRST)
    rows_c = {}
    rr = r + 2
    for k, (head, name) in enumerate(TIERS):
        for j, node in enumerate(G.NODES):
            rows_c[(head, node)] = rr
            put(ws, rr, 2, head, fmt="0")
            put(ws, rr, 3, node, fmt="0")
            for p in range(Y):
                put(ws, rr, FIRST + p, f"=SUMPRODUCT(INDEX(IC_Up,0,{j + 1})*INDEX(Mg_Uk,0,{k * Y + p + 1}))", fmt=NUM)
            rr += 1
    put(ws, rr + 1, 2, "Margin counts in a group when both sides and the asset's holder are in it. It is released as homes are sold "
                       "outside the group, or when the asset leaves the group. Upstream margin is on sales from a partly owned "
                       "node's group to the rest of the group: the node's NCI bears its share.", F_NOTE)
    return ws, rows_b, rows_c


ELIM_HEAD = ["Line", "Source", "Account", "Entity", "Amount", "Year", "Description"]


def eliminations_sheet(wb, register, inv_rows, margin_ws, rows_b, rows_c):
    ws = sheet(wb, "Eliminations", "Every elimination line by year and the groups it applies in: from the register, investments, "
                                   "unrealised margin and upstream NCI")
    widths(ws, {"B": 5, "C": 16, "D": 8, "E": 7, "F": 9, "G": 34, "H": 6})
    section(ws, 5, "Elimination lines", toc="Lines", last_col=FIRST + Y + K)
    for c, h in zip((2, 3, 4, 5, 6, 8, 7), ELIM_HEAD):
        put(ws, 7, c, h, F_BOLD, border=UNDER)
    year_heads(ws, 7, FIRST)
    put(ws, 6, FIRST + Y, "Applies in", F_NOTE)
    for k, (head, _) in enumerate(TIERS):
        put(ws, 7, FIRST + Y + k, head, F_BOLD, "0", border=UNDER, align=RIGHT)
        ws.column_dimensions[L(FIRST + Y + k)].width = 7
    for p in range(Y):
        ws.column_dimensions[L(FIRST + p)].width = 10
    r = 8
    first = r
    line = 0

    def emit(source, acct, entity, amount, year, desc, flags, by_year=None):
        nonlocal r, line
        line += 1
        put(ws, r, 2, line, fmt="0")
        put(ws, r, 3, source)
        put(ws, r, 4, acct, fmt="0")
        put(ws, r, 5, entity, fmt="0")
        if amount is not None:
            put(ws, r, 6, amount, fmt=NUM)
        if year is not None:
            put(ws, r, 8, year, fmt="0")
        put(ws, r, 7, desc)
        for p in range(Y):
            if by_year is not None:
                v = by_year(p)
            else:
                v = f"=IF($H{r}={p + 1},$F{r},0)"
            put(ws, r, FIRST + p, v, fmt=NUM)
        for k in range(K):
            put(ws, r, FIRST + Y + k, flags(k), fmt="0")
        r += 1

    # (a) the register: line A, line B and any difference between the two sides
    for i in range(len(register)):
        n = i + 1
        t = f"INDEX(IC_Type_No,{n})"
        sa, ba, ca = f"INDEX(IC_Seller_Acct,{n})", f"INDEX(IC_Buyer_Acct,{n})", f"INDEX(IC_Cost_Acct,{n})"
        se, be = f"INDEX(IC_Seller,{n})", f"INDEX(IC_Buyer,{n})"
        sv, bv, m = f"INDEX(IC_Seller_Amt,{n})", f"INDEX(IC_Buyer_Amt,{n})", f"INDEX(IC_Margin,{n})"
        yr = f"=INDEX(IC_Year,{n})"
        flags = (lambda k, n=n: f"=INDEX(IC_In,{n},{k + 1})")
        emit(f"Register {n}", f'=CHOOSE({t},{sa},{sa},{sa},"","","",{ba},{ba})', f'=CHOOSE({t},{se},{se},{se},"","","",{be},{be})',
             f"=CHOOSE({t},{sv},{sv},{sv},0,0,0,{bv},{bv})", yr, f'=INDEX(IC_Type,{n})&": first side"', flags)
        emit(f"Register {n}", f'=CHOOSE({t},{ba},{ca},{ca},"","","",{sa},{sa})', f'=CHOOSE({t},{be},{se},{se},"","","",{se},{se})',
             f"=CHOOSE({t},-{bv},-({sv}-{m}),-({sv}-{m}),0,0,0,-{bv},-{sv})", yr, f'=INDEX(IC_Type,{n})&": other side"', flags)
        emit(f"Register {n}", f'=IF(OR({t}=1,{t}=8),2900,"")', None,
             f"=CHOOSE({t},{bv}-{sv},0,0,0,0,0,0,{sv}-{bv})", yr, "Difference between the two sides", flags)
    # (b) investments against capital; the outside investors' share to NCI
    inv_ws, inv_first, inv_last, inv_in = inv_rows
    for i in range(inv_last - inv_first + 1):
        ir = inv_first + i
        flags = (lambda k, ir=ir: f"={q(inv_ws)}!{L(inv_in + k)}{ir}")
        sub = f"={q(inv_ws)}!$C{ir}"
        parent = f"={q(inv_ws)}!$B{ir}"
        emit(f"Investment {i + 1}", 3000, sub, None, None, "Sub's capital", flags,
             by_year=lambda p, ir=ir: f"={q(inv_ws)}!{L(FIRST + p)}{ir}")
        emit(f"Investment {i + 1}", 1500, parent, None, None, "Parent's investment", flags,
             by_year=lambda p, ir=ir: f"=-{q(inv_ws)}!{L(FIRST + 2 * Y + p)}{ir}")
        emit(f"Investment {i + 1}", 3300, sub, None, None, "Outside investors' share of capital", flags,
             by_year=lambda p, ir=ir: f"=-({q(inv_ws)}!{L(FIRST + p)}{ir}-{q(inv_ws)}!{L(FIRST + 2 * Y + p)}{ir})")
    # (c) unrealised margin: out of the asset where it sits at the year end, carried forward, released to cost of sale
    mq = q(margin_ws)
    for k, (head, name) in enumerate(TIERS):
        only = (lambda kk, k=k: 1 if kk == k else 0)
        for site in G.SITES:
            br = rows_b[(head, site)]
            emit(f"Margin {head} {site}", 3100, None, None, None, f"Margin held at the start of the year, {name}", only,
                 by_year=lambda p, br=br: f"={mq}!{L(FIRST + p - 1)}{br}" if p else "=0")
            emit(f"Margin {head} {site}", 5000, None, None, None, f"Margin released to cost of sale, {name}", only,
                 by_year=lambda p, br=br: f"=-{mq}!{L(FIRST + 2 * Y + p)}{br}")
            for y in range(Y):
                emit(f"Margin {head} {site}", f"={mq}!{L(FIRST + 3 * Y + y)}{br}", f"={mq}!{L(FIRST + 4 * Y + y)}{br}", None, None,
                     f"Margin out of the asset at {G.YEAR_LABELS[y]} year end, {name}", only,
                     by_year=lambda p, br=br, y=y: f"=-{mq}!{L(FIRST + y)}{br}" if p == y else "=0")
    # (d) upstream margin: the node's NCI bears its share
    for k, (head, name) in enumerate(TIERS):
        members = G.TIER_MEMBERS[head]
        only = (lambda kk, k=k: 1 if kk == k else 0)
        for node in G.NODES:
            if node == head or node not in members:
                continue
            cr = rows_c[(head, node)]
            share = f"(1-INDEX(Ent_Held,MATCH({node},Ent_Codes,0)))"
            U = lambda p, cr=cr: f"{mq}!{L(FIRST + p)}{cr}"             # noqa: E731
            emit(f"Upstream {head} {node}", 3300, node, None, None, f"NCI share of upstream margin still held, {name}", only,
                 by_year=lambda p, U=U, share=share: f"={share}*{U(p)}")
            emit(f"Upstream {head} {node}", 3100, node, None, None, f"NCI share of upstream margin at the start of the year, {name}",
                 only, by_year=lambda p, U=U, share=share: f"=-{share}*{U(p - 1)}" if p else "=0")
            emit(f"Upstream {head} {node}", 5900, node, None, None, f"NCI share of the year's upstream margin, {name}", only,
                 by_year=lambda p, U=U, share=share: f"=-{share}*({U(p)}-{U(p - 1)})" if p else f"=-{share}*{U(p)}")
    last = r - 1
    define(wb, "Elim_Acct", ref(ws, 4, first, 4, last))
    define(wb, "Elim_Val", ref(ws, FIRST, first, FIRST + Y - 1, last))
    define(wb, "Elim_In", ref(ws, FIRST + Y, first, FIRST + Y + K - 1, last))
    ws.freeze_panes = f"D{first}"
    return ws


def nci_sheet(wb, block_rows):
    """NCI at each node, from the node group's own consolidation, carried into the groups above it."""
    ws = sheet(wb, "NCI", "Non-controlling interests: worked out once at each node from its own consolidation, then carried up")
    widths(ws, {"B": 8, "C": 8, "D": 30, "H": 8})
    section(ws, 5, "At each node", toc="Nodes", last_col=FIRST + 5 * Y)
    put(ws, 7, 2, "Node", F_BOLD, border=UNDER)
    put(ws, 7, 4, "Name", F_BOLD, border=UNDER)
    put(ws, 7, 8, "NCI share", F_BOLD, border=UNDER, align=RIGHT)
    parts = [("Capital", 3000), ("Retained at start", 3100), ("Surplus", None), ("Distributions", 3200), ("Net assets", "na")]
    for j, (label, _) in enumerate(parts):
        put(ws, 6, FIRST + j * Y, label, F_BOLD)
        for p in range(Y):
            put(ws, 7, FIRST + j * Y + p, G.YEAR_LABELS[p], F_BOLD, border=UNDER, align=RIGHT)
            ws.column_dimensions[L(FIRST + j * Y + p)].width = 9
    node_rows = {}
    for i, node in enumerate(G.NODES):
        r = 8 + i
        node_rows[node] = r
        put(ws, r, 2, node, fmt="0")
        put(ws, r, 4, f"=INDEX(Ent_Names,MATCH(B{r},Ent_Codes,0))")
        put(ws, r, 8, f"=1-INDEX(Ent_Held,MATCH(B{r},Ent_Codes,0))", fmt=PCT)
        b = block_rows[node]
        for j, (label, acct) in enumerate(parts):
            for p in range(Y):
                col = L(FIRST + 2 * Y + p)                  # consolidated columns on By group
                if acct == "na":
                    v = f"='By group'!{L(FIRST + 2 * Y + p)}{b['na']}"
                elif acct is None:
                    v = f"='By group'!{L(FIRST + 2 * Y + p)}{b['surplus']}"
                else:
                    v = f"='By group'!{col}{b['acct'][acct]}"
                put(ws, r, FIRST + j * Y + p, v, fmt=NUM)
    # Lines for each group above a node
    r = 8 + len(G.NODES) + 2
    blocks = {}
    for k, (head, name) in enumerate(TIERS):
        nodes = [n for n in G.NODES if n != head and n in G.TIER_MEMBERS[head]]
        if not nodes:
            continue
        section(ws, r, f"Lines for the {name}", toc=f"{name}", last_col=FIRST + Y)
        put(ws, r + 1, 2, "Node", F_BOLD, border=UNDER)
        put(ws, r + 1, 3, "Account", F_BOLD, border=UNDER)
        put(ws, r + 1, 4, "Line", F_BOLD, border=UNDER)
        year_heads(ws, r + 1, FIRST)
        rr = r + 2
        first = rr
        for node in nodes:
            nr = node_rows[node]
            s = f"$H${nr}"
            cell = lambda j, p, nr=nr: f"{L(FIRST + j * Y + p)}${nr}"          # noqa: E731
            for acct, label, f in (
                    (3100, "NCI share of retained surplus at the start of the year", lambda p: f"=-{s}*{cell(1, p)}"),
                    (5900, "NCI share of the year's surplus", lambda p: f"={s}*{cell(2, p)}"),
                    (3200, "NCI share of distributions", lambda p: f"=-{s}*{cell(3, p)}"),
                    (3300, "Non-controlling interests", lambda p: f"={s}*({cell(1, p)}-{cell(2, p)}+{cell(3, p)})")):
                put(ws, rr, 2, node, fmt="0")
                put(ws, rr, 3, acct, fmt="0")
                put(ws, rr, 4, label)
                for p in range(Y):
                    put(ws, rr, FIRST + p, f(p), fmt=NUM)
                rr += 1
        last = rr - 1
        define(wb, f"NCI_{head}_Acct", ref(ws, 3, first, 3, last))
        define(wb, f"NCI_{head}_Val", ref(ws, FIRST, first, FIRST + Y - 1, last))
        blocks[head] = (first, last)
        r = rr + 2
    put(ws, r, 2, "The node's capital is split by the investment eliminations; these lines move the outside investors' share of "
                  "retained surplus, the year's surplus and distributions into NCI. Upstream margin adjustments are on Eliminations.",
        F_NOTE)
    return ws, blocks, node_rows


def by_group_sheet(wb, nci_heads):
    """One block per group: combined, eliminations, consolidated, by account and year."""
    ws = sheet(wb, "By group", "Each group's consolidation: its members combined, the eliminations that apply in it, and the result")
    widths(ws, {"B": 7, "C": 40, "H": 6})
    for c in range(FIRST, FIRST + 3 * Y):
        ws.column_dimensions[L(c)].width = 10
    n_acct = len(G.ACCOUNTS)
    # Summary first
    section(ws, 5, "Summary by group (each must be nil where marked)", toc="Summary", last_col=FIRST + 3 * Y + 2)
    sum_first = 8
    labels = [("Eliminations total (nil)", "elim"), ("Consolidated trial balance (nil)", "tb"),
              ("Intercompany differences (nil)", "diff"), ("NCI less the node shares (nil)", "nci")]
    for j, (label, _) in enumerate(labels):
        put(ws, 6, FIRST + j * Y if j < 3 else FIRST + 3 * Y, label, F_BOLD)
    put(ws, 7, 2, "Head", F_BOLD, border=UNDER, align=RIGHT)
    put(ws, 7, 3, "Group", F_BOLD, border=UNDER, align=IND)
    for j in range(4):
        for p in range(Y):
            col = FIRST + j * Y + p
            put(ws, 7, col, G.YEAR_LABELS[p], F_BOLD, border=UNDER, align=RIGHT)
            ws.column_dimensions[L(col)].width = 10
    block_top = sum_first + K + 3
    stride = n_acct + 12
    block_rows = {}
    for k, (head, name) in enumerate(TIERS):
        b0 = block_top + k * stride
        first = b0 + 3
        rows = {acct: first + i for i, (acct, _, _) in enumerate(G.ACCOUNTS)}
        last = first + n_acct - 1
        section(ws, b0, f"{name} (head {head})", toc=name, last_col=FIRST + 3 * Y - 1)
        for j, label in enumerate(["Combined", "Eliminations", "Consolidated"]):
            put(ws, b0 + 1, FIRST + j * Y, label, F_BOLD)
            for p in range(Y):
                put(ws, b0 + 2, FIRST + j * Y + p, G.YEAR_LABELS[p], F_BOLD, border=UNDER, align=RIGHT)
        put(ws, b0 + 2, 2, "Code", F_BOLD, border=UNDER, align=RIGHT)
        put(ws, b0 + 2, 3, "Name", F_BOLD, border=UNDER, align=IND)
        nci = head in nci_heads
        for acct, nm, cls in G.ACCOUNTS:
            r = rows[acct]
            put(ws, r, 2, acct, fmt="0")
            put(ws, r, 3, f"=INDEX(Acct_Names,MATCH(B{r},Acct_Codes,0))", align=IND)
            for p in range(Y):
                put(ws, r, FIRST + p, f"=SUMPRODUCT((Data_Acct=$B{r})*INDEX(Data_In,0,{k + 1})*INDEX(Data_Val,0,{p + 1}))", fmt=NUM)
                e = f"=SUMPRODUCT((Elim_Acct=$B{r})*INDEX(Elim_In,0,{k + 1})*INDEX(Elim_Val,0,{p + 1}))"
                if nci:
                    e += f"+SUMPRODUCT((NCI_{head}_Acct=$B{r})*INDEX(NCI_{head}_Val,0,{p + 1}))"
                put(ws, r, FIRST + Y + p, e, fmt=NUM)
                put(ws, r, FIRST + 2 * Y + p, f"={L(FIRST + p)}{r}+{L(FIRST + Y + p)}{r}", fmt=NUM)
        t = last + 1
        put(ws, t, 3, "Total (must be nil)", F_BOLD)
        for c in range(FIRST, FIRST + 3 * Y):
            put(ws, t, c, f"=ROUND(SUM({L(c)}{first}:{L(c)}{last}),6)", F_BOLD, NUM, border=TOP)
        pl = [rows[a] for a in G.PL]
        put(ws, t + 1, 3, "Surplus for the year")
        put(ws, t + 2, 3, "Attributable to non-controlling interests")
        put(ws, t + 3, 3, "Attributable to the owners")
        put(ws, t + 4, 3, "Net assets")
        assets = [rows[a] for a, _, c in G.ACCOUNTS if c in ("asset", "liability")]
        for p in range(3 * Y):
            col = L(FIRST + p)
            put(ws, t + 1, FIRST + p, "=-(" + "+".join(f"{col}{x}" for x in pl) + ")", fmt=NUM)
            put(ws, t + 2, FIRST + p, f"={col}{rows[G.NCIPL]}", fmt=NUM)
            put(ws, t + 3, FIRST + p, f"={col}{t + 1}-{col}{t + 2}", fmt=NUM)
            put(ws, t + 4, FIRST + p, "=" + "+".join(f"{col}{x}" for x in assets), fmt=NUM)
        block_rows[head] = {"acct": rows, "first": first, "last": last, "total": t, "surplus": t + 1, "nci_pl": t + 2,
                            "owners": t + 3, "na": t + 4, "b0": b0}
    # Summary rows
    for k, (head, name) in enumerate(TIERS):
        r = sum_first + k
        b = block_rows[head]
        put(ws, r, 2, head, fmt="0")
        put(ws, r, 3, name, align=IND)
        for p in range(Y):
            put(ws, r, FIRST + p, f"={L(FIRST + Y + p)}{b['total']}", fmt=NUM)
            put(ws, r, FIRST + Y + p, f"={L(FIRST + 2 * Y + p)}{b['total']}", fmt=NUM)
            put(ws, r, FIRST + 2 * Y + p, f"={L(FIRST + 2 * Y + p)}{b['acct'][G.DIFF]}", fmt=NUM)
    define(wb, "Sum_Elim", ref(ws, FIRST, sum_first, FIRST + Y - 1, sum_first + K - 1))
    define(wb, "Sum_TB", ref(ws, FIRST + Y, sum_first, FIRST + 2 * Y - 1, sum_first + K - 1))
    define(wb, "Sum_Diff", ref(ws, FIRST + 2 * Y, sum_first, FIRST + 3 * Y - 1, sum_first + K - 1))
    define(wb, "Grp_Cons", ref(ws, FIRST + 2 * Y, 1, FIRST + 3 * Y - 1, block_top + K * stride))
    define(wb, "Grp_Comb", ref(ws, FIRST, 1, FIRST + Y - 1, block_top + K * stride))
    ws.freeze_panes = "D6"
    return ws, block_rows, sum_first


def nci_tie(wb, ws, block_rows, sum_first, margin_rows_c, margin_ws, node_rows):
    """Summary: consolidated NCI against each node's share of its net assets less its share of upstream margin."""
    for k, (head, name) in enumerate(TIERS):
        r = sum_first + k
        b = block_rows[head]
        for p in range(Y):
            parts = []
            for node in G.NODES:
                if node == head or node not in G.TIER_MEMBERS[head]:
                    continue
                nr = node_rows[node]
                parts.append(f"NCI!$H${nr}*(NCI!{L(FIRST + 4 * Y + p)}{nr}-{q(margin_ws)}!{L(FIRST + p)}{margin_rows_c[(head, node)]})")
            expect = "+".join(parts) if parts else "0"
            put(ws, r, FIRST + 3 * Y + p, f"=ROUND(-{L(FIRST + 2 * Y + p)}{b['acct'][G.NCI]}-({expect}),6)", fmt=NUM)
    define(wb, "Sum_NCI", ref(ws, FIRST + 3 * Y, sum_first, FIRST + 4 * Y - 1, sum_first + K - 1))


# ----------------------------------------------------------------------------------- reports, lookups and checks
IS_LINES = [("Property sales", [4000]), ("Construction revenue", [4100]), ("Fee revenue", [4200]), ("Rental revenue", [4300]),
            ("Interest income", [4400]), ("Distribution income", [4500])]
EX_LINES = [("Cost of property sold", [5000]), ("Construction costs", [5100]), ("Operating expenses", [5200]),
            ("Interest expense", [5300])]
AS_LINES = [("Cash", [1000]), ("Receivables", [1100]), ("Development work in progress", [1200]), ("Investment property", [1300]),
            ("Loans receivable", [1400]), ("Investments in group entities", [1500])]
LI_LINES = [("Payables", [2000]), ("Intergroup AP/AR clearing", [2050]), ("Accruals", [2100]), ("GST payable", [2200]),
            ("Loans payable", [2300]), ("Intercompany differences", [2900])]


def structure_sheet(wb, block_rows):
    """The group as a tree, and each group's figures added up the tree for the year shown.

    The tree reads the Entities register in its order (kept in tree order by Add entity). Groups run children before
    parents, so each column's figures roll into a column further right."""
    ws = sheet(wb, "Group structure", "Who owns what, the groups each entity rolls into, and how each group's figures add up")
    n, cols = len(G.ENTITIES), [k for k, _ in sorted(enumerate(TIERS), key=lambda t: (-len(G.ancestors(t[1][0])), t[0]))]
    gc = {k: 10 + j for j, k in enumerate(cols)}                   # group column for each TIERS index (from J)
    c_first, c_own = 10 + K, 11 + K
    widths(ws, {"B": 7, "C": 40, "D": 8, "E": 13, "F": 15, "G": 15, "H": 9, "I": 14})
    for c in range(10, 10 + K):
        ws.column_dimensions[L(c)].width = 15
    widths(ws, {L(c_first): 19, L(c_own): 11, L(c_own + 1): 11})
    put(ws, 5, 3, "Year shown", F_BOLD)
    put(ws, 5, 5, G.YEAR_LABELS[-1], fill=FILL_IN, align=RIGHT)
    put(ws, 5, 6, "=MATCH(E5,LU_Years,0)", fmt="0")
    define(wb, "DD_Struct_Year", ref(ws, 5, 5))
    define(wb, "Str_Yr", ref(ws, 6, 5))
    dv = DataValidation(type="list", formula1="LU_Years", allow_blank=False)
    dv.error, dv.errorTitle = "Pick from the list", "Year shown"
    ws.add_data_validation(dv)
    dv.add("E5")
    # ---- the tree
    r = 7
    section(ws, r, "Group tree", last_col=c_own + 1)
    put(ws, r + 1, 10, "Rolls into (\u25cf), smallest group first", F_NOTE)
    put(ws, r + 1, c_own, '="Own figures, "&DD_Struct_Year', F_NOTE)
    hdr = r + 2
    for c, h in ((2, "Code"), (3, "Entity"), (4, "Parent"), (5, "Held by parent"), (6, "Owned by the top"),
                 (7, "Outside investors"), (8, "Status"), (9, "Member from"), (c_first, "First consolidated in"),
                 (c_own, "Surplus"), (c_own + 1, "Net assets")):
        put(ws, hdr, c, h, F_BOLD, border=UNDER, align=IND if c in (3, c_first) else (CENTRE if c == 8 else RIGHT))
    for k in range(K):
        put(ws, hdr, gc[k], f"=INDEX(Tier_Names,{k + 1})", F_BOLD, border=UNDER, align=CENTRE)
    t0, t1 = hdr + 1, hdr + n
    for i in range(n):
        r = t0 + i
        lvl = f"COUNT(INDEX(Ent_Anc,{i + 1},0))"
        put(ws, r, 2, f"=INDEX(Ent_Codes,{i + 1})", fmt="0")
        put(ws, r, 3, f'=IF({lvl}=0,"",REPT("      ",{lvl}-1)&"\u2514  ")&INDEX(Ent_Names,{i + 1})', align=IND)
        put(ws, r, 4, f'=IF(INDEX(Ent_Parent,{i + 1})=0,"",INDEX(Ent_Parent,{i + 1}))', fmt="0")
        put(ws, r, 5, f"=INDEX(Ent_Held,{i + 1})", fmt=PCT)
        put(ws, r, 6, f'=E{r}*IF(D{r}="",1,INDEX($F${t0}:$F${t1},MATCH(D{r},$B${t0}:$B${t1},0)))', fmt=PCT)
        put(ws, r, 7, f"=1-E{r}", fmt='0%;-0%;""')
        put(ws, r, 8, f"=INDEX(Ent_Status,{i + 1})", align=CENTRE)
        put(ws, r, 9, f"=INDEX(Ent_From,{i + 1})", fmt=DATE)
        for k in range(K):
            put(ws, r, gc[k], f"=INDEX(Ent_In,{i + 1},{k + 1})", fmt=DOT, align=CENTRE)
        out = '""'
        for k in range(K):                                       # largest first, so the smallest group's test ends up outermost
            out = f"IF({L(gc[k])}{r}=1,INDEX(Tier_Names,{k + 1}),{out})"
        put(ws, r, c_first, "=" + out, align=IND)
        put(ws, r, c_own, f"=-SUMPRODUCT((Data_Ent=$B{r})*Data_PL*INDEX(Data_Val,0,Str_Yr))", fmt=NUM)
        put(ws, r, c_own + 1, f"=SUMPRODUCT((Data_Ent=$B{r})*Data_NA*INDEX(Data_Val,0,Str_Yr))", fmt=NUM)
    # ---- groups: which group each one rolls up into
    r = t1 + 3
    section(ws, r, "Groups and where they roll up", toc="Groups", last_col=c_own + 1)
    g_hdr = r + 1
    for c, h in ((2, "Head"), (3, "Group"), (4, "Members"), (5, "Held by parent"), (7, "Outside investors"), (10, "Rolls up into")):
        put(ws, g_hdr, c, h, F_BOLD, border=UNDER, align=IND if c in (3, 10) else RIGHT)
    g0 = g_hdr + 1
    for j, k in enumerate(cols):
        r = g0 + j
        head = TIERS[k][0]
        m = f"MATCH($B{r},Ent_Codes,0)"
        put(ws, r, 2, f"=INDEX(Tier_Heads,{k + 1})", fmt="0")
        put(ws, r, 3, f"=INDEX(Tier_Names,{k + 1})", align=IND)
        put(ws, r, 4, f"=INDEX(Tier_Sizes,{k + 1})", fmt="0")
        put(ws, r, 5, f"=INDEX(Ent_Held,{m})", fmt=PCT)
        put(ws, r, 7, f"=1-E{r}", fmt='0%;-0%;""')
        out = '""'
        for kk in range(K):
            out = f"IF(AND(INDEX(Ent_In,{m},{kk + 1})=1,INDEX(Tier_Heads,{kk + 1})<>$B{r}),INDEX(Tier_Names,{kk + 1}),{out})"
        put(ws, r, 10, "=" + out, align=IND)
    g1 = g0 + K - 1
    # ---- roll-ups
    ties = []

    def rollup(r, title, toc, own_col, which, nci_formula):
        section(ws, r, title, toc=toc, last_col=c_own + 1)
        put(ws, r + 1, 3, "Rolls up into", F_NOTE)
        for k in range(K):
            c = gc[k]
            put(ws, r + 1, c, f"=INDEX($J${g0}:$J${g1},{cols.index(k) + 1})", F_NOTE, align=CENTRE)
            put(ws, r + 2, c, f"=INDEX(Tier_Names,{k + 1})", F_BOLD, border=UNDER, align=RIGHT)
        put(ws, r + 2, 2, "Code", F_BOLD, border=UNDER, align=RIGHT)
        put(ws, r + 2, 3, "Entity (own figures)", F_BOLD, border=UNDER, align=IND)
        e0 = r + 3
        for i in range(n):
            rr = e0 + i
            put(ws, rr, 2, f"=B{t0 + i}", fmt="0")
            put(ws, rr, 3, f"=C{t0 + i}", align=IND)
            for k in range(K):
                c = L(gc[k])
                put(ws, rr, gc[k], f'=IF({c}{t0 + i}=1,${L(own_col)}{t0 + i},"")', fmt=NUM)
        e1 = e0 + n - 1
        rows = {}
        labels = [("comb", "Members combined", True), ("below", "Eliminations and adjustments made in the groups below it", False),
                  ("here", "Eliminations and adjustments made in this group", False), ("cons", "Consolidated", True),
                  ("nci", "Attributable to non-controlling interests", False), ("owners", "Attributable to the owners", False),
                  ("total", "Eliminations and adjustments in total (from By group)", False),
                  ("tie", "Consolidated less By group (must be nil)", False)]
        for j, (key, label, bold) in enumerate(labels):
            rows[key] = e1 + 1 + j
            put(ws, rows[key], 3, label, F_NOTE if key in ("total", "tie") else (F_BOLD if bold else F_BODY), align=IND)
        span_row = lambda row: f"${L(gc[cols[0]])}${row}:${L(gc[cols[-1]])}${row}"            # noqa: E731
        span = lambda key: span_row(rows[key])                                                  # noqa: E731
        for k in range(K):
            head, c = TIERS[k][0], L(gc[k])
            b = block_rows[head]
            yr = lambda off, row: f"INDEX('By group'!${L(FIRST + off)}${row}:${L(FIRST + off + Y - 1)}${row},Str_Yr)"  # noqa: E731
            f = F_BOLD
            put(ws, rows["comb"], gc[k], f"=SUM({c}{e0}:{c}{e1})", f, NUM, border=TOP)
            put(ws, rows["total"], gc[k], "=" + yr(Y, b[which]), F_NOTE, NUM)
            put(ws, rows["below"], gc[k], f"=SUMPRODUCT(({span_row(r + 1)}={c}${r + 2})*{span('total')})",
                fmt=NUM)
            put(ws, rows["here"], gc[k], f"={c}{rows['total']}-{c}{rows['below']}", fmt=NUM)
            put(ws, rows["cons"], gc[k], f"={c}{rows['comb']}+{c}{rows['below']}+{c}{rows['here']}", f, NUM, border=TOP)
            put(ws, rows["nci"], gc[k], "=" + nci_formula(b, yr), fmt=NUM)
            put(ws, rows["owners"], gc[k], f"={c}{rows['cons']}-{c}{rows['nci']}", fmt=NUM)
            put(ws, rows["tie"], gc[k], f"=ROUND({c}{rows['cons']}-{yr(2 * Y, b[which])},6)", F_NOTE, NUM)
        ties.append(span("tie"))
        return rows["tie"] + 1

    r = g1 + 3
    r = rollup(r, '="Surplus for the year, rolled up, "&DD_Struct_Year', "Surplus rolled up", c_own, "surplus",
               lambda b, yr: yr(2 * Y, b["nci_pl"]))
    r = rollup(r + 2, '="Net assets, rolled up, "&DD_Struct_Year', "Net assets rolled up", c_own + 1, "na",
               lambda b, yr: "-" + yr(2 * Y, b["acct"][G.NCI]))
    put(ws, r + 1, 2, "Each column adds the members' own figures, the eliminations made in the groups below it (already in those "
                      "groups' columns) and the eliminations made in this group. Home Hub is the master for actual entities; "
                      "planned ones are added with Add entity.", F_NOTE)
    define(wb, "Str_Tie", f"{q(ws)}!{ties[0]}")
    define(wb, "Str_Tie_NA", f"{q(ws)}!{ties[1]}")
    ws.freeze_panes = f"D{t0}"
    return ws


def statements_sheet(wb, block_rows):
    ws = sheet(wb, "Group statements", "Consolidated statements for the group and year shown, the entities in it and its related parties")
    widths(ws, {"B": 3, "C": 3, "D": 3, "E": 3, "F": 3, "G": 40, "H": 12, "I": 12, "J": 12, "K": 12, "L": 12})
    put(ws, 5, 3, "Group shown", F_BOLD)
    put(ws, 5, 8, TIERS[0][1], fill=FILL_IN, align=RIGHT)
    put(ws, 5, 9, "=MATCH(H5,Tier_Names,0)", fmt="0")
    put(ws, 6, 3, "Year shown", F_BOLD)
    put(ws, 6, 8, G.YEAR_LABELS[-1], fill=FILL_IN, align=RIGHT)
    put(ws, 6, 9, "=MATCH(H6,LU_Years,0)", fmt="0")
    define(wb, "DD_Group", ref(ws, 8, 5))
    define(wb, "Grp_Shown", ref(ws, 9, 5))
    define(wb, "DD_Year", ref(ws, 8, 6))
    define(wb, "Yr_Shown", ref(ws, 9, 6))
    for cell, lst, title in (("H5", "Tier_Names", "Group shown"), ("H6", "LU_Years", "Year shown")):
        dv = DataValidation(type="list", formula1=lst, allow_blank=False)
        dv.error, dv.errorTitle = "Pick from the list", title
        ws.add_data_validation(dv)
        dv.add(cell)
    offsets = {a: block_rows[TIERS[0][0]]["acct"][a] - block_rows[TIERS[0][0]]["b0"] for a, _, _ in G.ACCOUNTS}
    stride = block_rows[TIERS[1][0]]["b0"] - block_rows[TIERS[0][0]]["b0"]
    b0 = block_rows[TIERS[0][0]]["b0"]

    def val(acct, sign=1, which="Grp_Cons"):
        off = offsets[acct]
        v = f"INDEX({which},{b0}+(Grp_Shown-1)*{stride}+{off},Yr_Shown)"
        return v if sign == 1 else f"-{v}"

    r = 8
    section(ws, r, '="Income statement, "&DD_Group&", "&DD_Year', toc="Income statement", last_col=9)
    rows = {}

    def line(label, accts, sign, bold=False):
        nonlocal r
        r += 1
        put(ws, r, 3, label, F_BOLD if bold else F_BODY)
        put(ws, r, 8, "=" + "+".join(val(a, sign) for a in accts), F_BOLD if bold else F_BODY, NUM)
        return r

    def total(label, first, last, bold=True):
        nonlocal r
        r += 1
        put(ws, r, 3, label, F_BOLD)
        put(ws, r, 8, f"=SUM(H{first}:H{last})", F_BOLD, NUM, border=TOP)
        return r

    f0 = r + 1
    for label, accts in IS_LINES:
        line(label, accts, -1)
    rev = total("Revenue", f0, r)
    f0 = r + 1
    for label, accts in EX_LINES:
        line(label, accts, -1)
    exp = total("Expenses", f0, r)
    r += 1
    put(ws, r, 3, "Surplus for the year", F_BOLD)
    put(ws, r, 8, f"=H{rev}+H{exp}", F_BOLD, NUM, border=TOP)
    rows["surplus"] = r
    r += 1
    put(ws, r, 4, "Attributable to non-controlling interests")
    put(ws, r, 8, "=" + val(5900), fmt=NUM)
    rows["nci_pl"] = r
    r += 1
    put(ws, r, 4, "Attributable to the owners")
    put(ws, r, 8, f"=H{r - 2}-H{r - 1}", fmt=NUM)
    r += 2
    section(ws, r, '="Balance sheet, "&DD_Group&", "&DD_Year', toc="Balance sheet", last_col=9)
    f0 = r + 1
    for label, accts in AS_LINES:
        line(label, accts, 1)
    ta = total("Assets", f0, r)
    f0 = r + 1
    for label, accts in LI_LINES:
        line(label, accts, -1)
    tl = total("Liabilities", f0, r)
    r += 1
    put(ws, r, 3, "Net assets", F_BOLD)
    put(ws, r, 8, f"=H{ta}-H{tl}", F_BOLD, NUM, border=TOP)
    na = r
    r += 1
    line("Capital", [3000], -1)
    eq0 = r
    r += 1
    put(ws, r, 3, "Retained surplus")
    put(ws, r, 8, "=-(" + "+".join(val(a) for a in [3100, 3200, 5900] + G.PL) + ")", fmt=NUM)
    r += 1
    put(ws, r, 3, "Equity of the owners", F_BOLD)
    put(ws, r, 8, f"=H{eq0}+H{r - 1}", F_BOLD, NUM, border=TOP)
    line("Non-controlling interests", [3300], -1)
    r += 1
    put(ws, r, 3, "Equity", F_BOLD)
    put(ws, r, 8, f"=H{r - 2}+H{r - 1}", F_BOLD, NUM, border=TOP)
    eq = r
    define(wb, "St_Check", f"{q(ws)}!$H${na}")
    r += 1
    put(ws, r, 3, "Net assets less equity (must be nil)", F_NOTE)
    put(ws, r, 8, f"=ROUND(H{na}-H{eq},6)", F_CHECK, NUM)
    define(wb, "St_NA_Less_Eq", ref(ws, 8, r))
    # Entities summary
    r += 2
    section(ws, r, '="Entities in the "&DD_Group&", "&DD_Year', toc="Entities summary", last_col=11)
    r += 1
    for c, h in ((3, "Entity"), (8, "In group"), (9, "Surplus"), (10, "Net assets")):
        put(ws, r, c, h, F_BOLD, border=UNDER, align=RIGHT if c > 3 else None)
    e0 = r + 1
    for code in G.CODES:
        r += 1
        put(ws, r, 3, f'=INDEX(Ent_Codes,MATCH({code},Ent_Codes,0))&"  "&INDEX(Ent_Names,MATCH({code},Ent_Codes,0))')
        put(ws, r, 8, f"=INDEX(Ent_In,MATCH({code},Ent_Codes,0),Grp_Shown)", fmt="0")
        put(ws, r, 9, f"=-H{r}*SUMPRODUCT((Data_Ent={code})*Data_PL*INDEX(Data_Val,0,Yr_Shown))", fmt=NUM)
        put(ws, r, 10, f"=H{r}*SUMPRODUCT((Data_Ent={code})*Data_NA*INDEX(Data_Val,0,Yr_Shown))", fmt=NUM)
    e1 = r
    r += 1
    put(ws, r, 3, "Members combined", F_BOLD)
    put(ws, r, 9, f"=SUM(I{e0}:I{e1})", F_BOLD, NUM, border=TOP)
    put(ws, r, 10, f"=SUM(J{e0}:J{e1})", F_BOLD, NUM, border=TOP)
    comb = r
    r += 1
    put(ws, r, 3, "Eliminations and NCI adjustments")
    put(ws, r, 9, f"=I{r + 1}-I{comb}", fmt=NUM)
    put(ws, r, 10, f"=J{r + 1}-J{comb}", fmt=NUM)
    r += 1
    put(ws, r, 3, "Consolidated", F_BOLD)
    put(ws, r, 9, f"=H{rows['surplus']}", F_BOLD, NUM, border=TOP)
    put(ws, r, 10, f"=H{na}", F_BOLD, NUM, border=TOP)
    # Related parties
    r += 2
    section(ws, r, '="Related parties of the "&DD_Group&", "&DD_Year&": intergroup pairs with one side outside it"',
            toc="Related parties", last_col=11)
    r += 1
    put(ws, r, 3, "Type", F_BOLD, border=UNDER)
    put(ws, r, 8, "Pairs", F_BOLD, border=UNDER, align=RIGHT)
    put(ws, r, 9, "Amount", F_BOLD, border=UNDER, align=RIGHT)
    for t in G.TYPES:
        r += 1
        put(ws, r, 3, t)
        put(ws, r, 8, f'=SUMPRODUCT(INDEX(IC_Related,0,Grp_Shown)*(IC_Year=Yr_Shown)*(IC_Type=C{r}))', fmt="0")
        put(ws, r, 9, f'=SUMPRODUCT(INDEX(IC_Related,0,Grp_Shown)*(IC_Year=Yr_Shown)*(IC_Type=C{r})*IC_Seller_Amt)', fmt=NUM)
    r += 1
    put(ws, r, 3, "Pairs eliminated in a group are related party transactions in the groups below it, where one side sits outside.",
        F_NOTE)
    for rr in range(1, r + 1):
        ws.row_dimensions[rr].height = 15
    return ws


def lookups_sheet(wb, block_rows):
    ws = sheet(wb, "Lookups", "Lists behind the drop-downs")
    widths(ws, {"C": 10, "E": 22})
    put(ws, 5, 3, "Years", F_BOLD)
    for p in range(Y):
        put(ws, 6 + p, 3, G.YEAR_LABELS[p])
    define(wb, "LU_Years", ref(ws, 3, 6, 3, 5 + Y))
    put(ws, 5, 5, "Intercompany types", F_BOLD)
    for i, t in enumerate(G.TYPES):
        put(ws, 6 + i, 5, t)
    define(wb, "LU_IC_Types", ref(ws, 5, 6, 5, 5 + len(G.TYPES)))
    put(ws, 5, 7, "Year ends", F_BOLD)
    for p in range(Y):
        put(ws, 6 + p, 7, datetime.strptime(G.YEAR_START[p], "%d %B %Y").replace(year=int(G.YEAR_LABELS[p][2:]), day=31, month=3),
            fmt=DATE)
    define(wb, "LU_Year_End", ref(ws, 7, 6, 7, 5 + Y))
    put(ws, 5, 9, "Yes or no", F_BOLD)
    for i, v in enumerate(("Yes", "No")):
        put(ws, 6 + i, 9, v)
    define(wb, "LU_YesNo", ref(ws, 9, 6, 9, 7))
    put(ws, 5, 11, "Entity status", F_BOLD)
    for i, v in enumerate(("Actual", "Planned")):
        put(ws, 6 + i, 11, v)
    define(wb, "LU_Status", ref(ws, 11, 6, 11, 7))
    widths(ws, {"G": 14, "I": 10, "K": 12})
    return ws


def checks_sheet(wb, block_rows, ent_ws, c_in):
    ws = sheet(wb, "Checks", "Error checks must be nil; alerts are for review")
    widths(ws, {"C": 96, "H": 8})
    top = block_rows[TIERS[0][0]]
    cons = lambda acct, p: f"'By group'!{L(FIRST + 2 * Y + p)}{top['acct'][acct]}"        # noqa: E731
    errors = [
        ("Each entity's trial balance adds to nil in every year", "=IF(SUMPRODUCT(ABS(Ent_TB_Sum))>0.001,1,0)"),
        ("Every intercompany row names two entities in the register and a group that holds both",
         '=IF(SUMPRODUCT(--ISERROR(IC_Group))+COUNTIF(IC_Group,"None")>0,1,0)'),
        ("The two sides of every intercompany pair agree (a break is a timing difference or an error to clear)",
         '=IF(COUNTIF(IC_Status,"Break")>0,1,0)'),
        ("No intercompany difference is left in any group", "=IF(SUMPRODUCT(ABS(Sum_Diff))>0.001,1,0)"),
        ("Each netting account's balance is in the register as an on-charge still to raise",
         "=IF(SUMPRODUCT(ABS(Ent_Net_Gap))>0.001,1,0)"),
        ("Each investment is the parent's share of the sub's capital", "=IF(SUMPRODUCT(ABS(Inv_Diff))>0.001,1,0)"),
        ("Each parent's investments balance is the sum of its investments", "=IF(SUMPRODUCT(ABS(Ent_Inv_Gap))>0.001,1,0)"),
        ("Eliminations balance in every group and year", "=IF(SUMPRODUCT(ABS(Sum_Elim))>0.001,1,0)"),
        ("Every group's consolidated trial balance balances", "=IF(SUMPRODUCT(ABS(Sum_TB))>0.001,1,0)"),
        ("NCI is each node's share of its net assets less its share of upstream margin still held",
         "=IF(SUMPRODUCT(ABS(Sum_NCI))>0.001,1,0)"),
        ("The top group carries no investments in group entities and no netting account balances",
         "=IF(" + "+".join(f"ABS({cons(a, p)})" for a in (1500,) for p in range(Y)) + ">0.001,1,0)"),
        ("Group statements: net assets equal equity", "=IF(ABS(St_NA_Less_Eq)>0.001,1,0)"),
        ("Group statements: the group and year shown are in their lists", "=IF(OR(ISERROR(Grp_Shown),ISERROR(Yr_Shown)),1,0)"),
        ("Entities register: each entity sits under its parent, in tree order, with one entity at the top",
         "=IF(COUNTIF(Ent_Tree_Ok,0)>0,1,0)"),
        ("No entity has figures before the date it joins the group", "=IF(SUMPRODUCT(ABS(Ent_Before))>0.001,1,0)"),
        ("Group structure: each group's roll-up ties to its consolidation", "=IF(IFERROR(SUMPRODUCT(ABS(Str_Tie))+SUMPRODUCT(ABS(Str_Tie_NA)),1)>0.001,1,0)"),
    ]
    alerts = [
        ("A netting account has a balance at a year end: an on-charge was not raised by the cut-off and the receiver accrued it",
         "=IF(SUMPRODUCT(ABS((Data_Acct=2050)*Data_Val))>0.001,1,0)"),
        ("GST charged to a group entity that cannot claim it is a real cost to the group and is not eliminated",
         "=IF(SUM(IC_GST_Cost)>0,1,0)"),
        ("Planned entities are in the model that Home Hub does not have yet", '=IF(COUNTIF(Ent_Status,"Planned")>0,1,0)'),
    ]
    r = 5
    for title, items, nm in (("Error checks", errors, "Chk_Errors"), ("Alerts", alerts, "Chk_Alerts")):
        section(ws, r, title, last_col=8)
        r += 1
        first = r
        for label, f in items:
            put(ws, r, 3, label)
            put(ws, r, 8, f, F_BODY, "0")
            r += 1
        if nm == "Chk_Errors":                                   # red only when a check is raised
            ws.conditional_formatting.add(f"H{first}:H{r - 1}", CellIsRule(operator="greaterThan", formula=["0"],
                                                                           font=Font(color="9C0006", bold=True)))
        put(ws, r, 3, f"{title} raised", F_BOLD)
        put(ws, r, 8, f"=SUM(H{first}:H{r - 1})", F_BOLD, "0", border=TOP)
        define(wb, nm, ref(ws, 8, r))
        r += 2
    return ws


def entity_gaps(wb, ws, c_in):
    """Per entity and year: netting account balance not in the register, and investments balance not in the investments list."""
    first, last = 8, 8 + len(G.ENTITIES) - 1
    c0 = c_in + K + Y
    put(ws, 6, c0, "Netting not in register (nil)", F_NOTE)
    put(ws, 6, c0 + Y, "Investments not in list (nil)", F_NOTE)
    for j in range(2):
        for p in range(Y):
            put(ws, 7, c0 + j * Y + p, G.YEAR_LABELS[p], F_BOLD, border=UNDER, align=RIGHT)
            ws.column_dimensions[L(c0 + j * Y + p)].width = 9
    for r in range(first, last + 1):
        for p in range(Y):
            put(ws, r, c0 + p, f"=ROUND(SUMPRODUCT((Data_Ent=$B{r})*(Data_Acct=2050)*INDEX(Data_Val,0,{p + 1}))"
                               f"-SUMPRODUCT((IC_Seller=$B{r})*(IC_Type=\"Balance\")*(IC_Seller_Acct=2050)*(IC_Year={p + 1})*IC_Seller_Amt),6)", fmt=NUM)
            put(ws, r, c0 + Y + p, f"=ROUND(SUMPRODUCT((Data_Ent=$B{r})*(Data_Acct=1500)*INDEX(Data_Val,0,{p + 1}))"
                                   f"-SUMPRODUCT((Inv_Parent=$B{r})*INDEX(Inv_Cost_Y,0,{p + 1})),6)", fmt=NUM)
    define(wb, "Ent_Net_Gap", ref(ws, c0, first, c0 + Y - 1, last))
    define(wb, "Ent_Inv_Gap", ref(ws, c0 + Y, first, c0 + 2 * Y - 1, last))
    # Tree order (each entity sits below its parent or under one of the entity above's ancestors) and figures before joining
    c1 = c0 + 2 * Y
    put(ws, 6, c1, "Register order and dates", F_NOTE)
    put(ws, 7, c1, "Tree order", F_BOLD, border=UNDER, align=RIGHT)
    put(ws, 7, c1 + 1, "Before joining", F_BOLD, border=UNDER, align=RIGHT)
    ws.column_dimensions[L(c1)].width = 10
    ws.column_dimensions[L(c1 + 1)].width = 13
    for r in range(first, last + 1):
        if r == first:
            put(ws, r, c1, f'=IF($D{r}="",1,0)', fmt="0")
        else:
            above = ",".join(f"$D{r}=${c}{r - 1}" for c in ("B", "K", "L", "M", "N"))
            put(ws, r, c1, f'=IF(AND($D{r}<>"",OR({above})),1,0)', fmt="0")
        before = "+".join(f"SUMPRODUCT((Data_Ent=$B{r})*ABS(INDEX(Data_Val,0,{p + 1})))*(INDEX(LU_Year_End,{p + 1})<$J{r})"
                          for p in range(Y))
        put(ws, r, c1 + 1, f"=ROUND({before},6)", fmt=NUM)
    define(wb, "Ent_Tree_Ok", ref(ws, c1, first, c1, last))
    define(wb, "Ent_Before", ref(ws, c1 + 1, first, c1 + 1, last))


def data_flags(wb, ws, first, last):
    """P&L and net assets flags beside each trial balance line (for the entities summary)."""
    c = FIRST + Y + K
    put(ws, 6, c, "P&L", F_BOLD, border=UNDER, align=RIGHT)
    put(ws, 6, c + 1, "Net assets", F_BOLD, border=UNDER, align=RIGHT)
    rows = [r for r in range(first, last + 1) if isinstance(ws.cell(r, 3).value, int)]
    for r in range(first, last + 1):
        if r in rows:
            put(ws, r, c, f'=IF(OR(INDEX(Acct_Class,MATCH($C{r},Acct_Codes,0))="revenue",INDEX(Acct_Class,MATCH($C{r},Acct_Codes,0))="expense"),1,0)',
                fmt="0")
            put(ws, r, c + 1, f'=IF(OR(INDEX(Acct_Class,MATCH($C{r},Acct_Codes,0))="asset",INDEX(Acct_Class,MATCH($C{r},Acct_Codes,0))="liability"),1,0)',
                fmt="0")
        else:
            put(ws, r, c, 0, fmt="0")
            put(ws, r, c + 1, 0, fmt="0")
    define(wb, "Data_PL", ref(ws, c, first, c, last))
    define(wb, "Data_NA", ref(ws, c + 1, first, c + 1, last))


def gst_cost(wb, ws, first, last, c):
    col = c + 4 + 2 * K + len(G.NODES)
    put(ws, 7, col, "GST kept as group cost", F_BOLD, border=UNDER, align=RIGHT)
    ws.column_dimensions[L(col)].width = 10
    for r in range(first, last + 1):
        put(ws, r, col, f'=IF(INDEX(Ent_GST,MATCH($F{r},Ent_Codes,0))="No",$N{r},0)', fmt=NUM)
    define(wb, "IC_GST_Cost", ref(ws, col, first, col, last))


def build(path: Path) -> Path:
    HEADINGS.clear()
    bk = G.book()
    tb = G.trial_balances(bk)
    global TIERS, K
    TIERS = sorted(G.TIERS, key=lambda t: -len(G.TIER_MEMBERS[t[0]]))
    K = len(TIERS)
    wb = Workbook()
    wb.active.title = "Contents"
    nci_heads = {h for h, _ in TIERS if any(n != h and n in G.TIER_MEMBERS[h] for n in G.NODES)}
    ent_ws, c_in = entities_sheet(wb, tb)
    accounts_sheet(wb)
    sites_sheet(wb)
    data_ws, d_first, d_last = data_sheet(wb, tb, c_in)
    data_flags(wb, data_ws, d_first, d_last)
    ic_ws, ic_first, ic_last, ic_c = intercompany_sheet(wb, bk.register)
    gst_cost(wb, ic_ws, ic_first, ic_last, ic_c)
    inv = investments_sheet(wb, bk.investments)
    mg_ws, rows_b, rows_c = margin_sheet(wb, bk.register)
    eliminations_sheet(wb, bk.register, inv, mg_ws, rows_b, rows_c)
    grp_ws, block_rows, sum_first = by_group_sheet(wb, nci_heads)
    nci_ws, nci_blocks, node_rows = nci_sheet(wb, block_rows)
    assert set(nci_blocks) == nci_heads
    nci_tie(wb, grp_ws, block_rows, sum_first, rows_c, mg_ws, node_rows)
    entity_gaps(wb, ent_ws, c_in)
    structure_sheet(wb, block_rows)
    statements_sheet(wb, block_rows)
    lookups_sheet(wb, block_rows)
    checks_sheet(wb, block_rows, ent_ws, c_in)
    nav = navigation.Navigation(
        model_name="Demo Group", model_kind="Group consolidation (Phase 0 proof)",
        covers={"Reports": navigation.Cover("Group reports", "The group's structure and roll-up, and consolidated statements for any group and year, the entities in it and its related parties."),
                "Inputs": navigation.Cover("Inputs", "What Home Hub or each entity's saved version hands over: entities, the group chart, sites, trial balances and the intercompany register."),
                "Consolidate": navigation.Cover("Consolidation", "Investments, unrealised margin, eliminations, NCI and each group's consolidation."),
                "Appendices": navigation.Cover("Appendices", "Lookups and the checks.")},
        notes=["Eleven fictional entities in HFG's shape: a parent, a builder, a development manager and a property manager; Holdings, "
               "Devco and two development LPs; a Fund with outside investors and its own development LP; a partly owned LP.",
               "At-cost on-charges pass through the Intergroup AP/AR netting accounts; one is capitalised into a development LP's WIP.",
               "Construction claims, development fees and capitalised interest carry margin into WIP; two portfolios are sold to the Fund; "
               "five homes are sold outside the group in FY2028.",
               "Every pair is eliminated in the lowest group holding both sides and every group above it; NCI is worked out once at each node."],
        headings=HEADINGS, entity=NOTE)
    navigation.apply(wb, nav, [("Reports", ["Group structure", "Group statements"]),
                               ("Inputs", ["Entities", "Accounts", "Sites", "Entity data", "Intercompany"]),
                               ("Consolidate", ["Investments", "Unrealised margin", "Eliminations", "NCI", "By group"]),
                               ("Appendices", ["Lookups", "Checks"])])
    for ws in wb.worksheets:
        if ws.title in ("Contents",) or ws.title in nav.covers:
            continue
        for rr in range(1, ws.max_row + 1):
            ws.row_dimensions[rr].height = 15
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "build" / "consolidation" / "consolidation_demo.xlsx"
    print(build(out))
