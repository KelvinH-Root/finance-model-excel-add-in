"""A fictional group shaped like HFG's, its transactions, and an independent consolidation.

Eleven entities in the same shape as HFG's tiers (numbers and names are fictional):

    9000 Foundation (parent)                          Foundation group
      9001 BuildCo Ltd          construction, charges with a margin
      9002 DevManager Ltd       development management fees, with a margin
      9003 PropManager Ltd      property management, GST registered
      9004 Holdings LP                                  Holdings group
        9005 Devco LP                                   Devco group
          9006 Dev LP A         builds site SA, sells the portfolio to the Fund
          9007 Dev LP B         builds site SB on an intergroup loan
        9008 Housing Fund LP    60% held, 40% external investors (NCI node)   Fund group
          9009 Fund Dev LP      builds site SF
        9010 Partner LP         70% held, 30% external (NCI node), sells site SP to the Fund

Every transaction is an event that writes the entity journals (what each entity's ledger
holds; debit positive, with the counterparty on intergroup lines) and the register rows
Home Hub's matching would hand the model. The reference consolidation does not use the
register: for each group it re-records every event as the group sees it (an intra-group
charge capitalised at the seller's cost, an intra-group sale moving the asset at group cost,
balances and the at-cost netting accounts gone), tracking each site's group cost, and works
out non-controlling interests from each node's net assets. The workbook's elimination
entries have to land on the same numbers.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

PERIODS = 3
YEAR_LABELS = ["FY2026", "FY2027", "FY2028"]

# code, name, parent, share held by the parent, GST registered, role
ENTITIES = [
    (9000, "Foundation", None, 1.0, True, "Parent; pays shared costs and on-charges them at cost"),
    (9001, "BuildCo Ltd", 9000, 1.0, True, "Builds for the development LPs with a margin"),
    (9002, "DevManager Ltd", 9000, 1.0, True, "Development management fees with a margin"),
    (9003, "PropManager Ltd", 9000, 1.0, True, "Manages the Fund's homes for a fee plus GST"),
    (9004, "Holdings LP", 9000, 1.0, False, "Sub-group head; due diligence, loans to LPs"),
    (9005, "Devco LP", 9004, 1.0, False, "Sub-group head for the development LPs"),
    (9006, "Dev LP A", 9005, 1.0, False, "Builds site SA and sells the portfolio to the Fund"),
    (9007, "Dev LP B", 9005, 1.0, False, "Builds site SB on a Holdings loan"),
    (9008, "Housing Fund LP", 9004, 0.6, False, "Holds the portfolios; external investors hold 40%"),
    (9009, "Fund Dev LP", 9008, 1.0, False, "Builds site SF for the Fund"),
    (9010, "Partner LP", 9004, 0.7, False, "Partly owned; sells site SP to the Fund"),
]
NAME = {e[0]: e[1] for e in ENTITIES}
PARENT = {e[0]: e[2] for e in ENTITIES}
HELD = {e[0]: e[3] for e in ENTITIES}
GST_REG = {e[0]: e[4] for e in ENTITIES}
CODES = [e[0] for e in ENTITIES]

# Groups that consolidate (tier heads). Partner LP is a node of its own for its NCI.
TIERS = [(9000, "Foundation group"), (9004, "Holdings group"), (9005, "Devco group"), (9008, "Fund group"),
         (9010, "Partner LP")]
NODES = [9008, 9010]              # entities with non-controlling interests, worked out once at their own node


def ancestors(code: int) -> list[int]:
    out, p = [], PARENT[code]
    while p is not None:
        out.append(p)
        p = PARENT[p]
    return out


def members(head: int) -> set[int]:
    return {c for c in CODES if c == head or head in ancestors(c)}


TIER_MEMBERS = {head: members(head) for head, _ in TIERS}

# Where each entity comes from and when it joined. Home Hub is the master for actual entities; an entity added
# in a model for a plan (a future LP, a new fund) is Planned until Home Hub has it. Financial years run April to March.
YEAR_START = ["1 April 2025", "1 April 2026", "1 April 2027"]
STATUS = {c: "Actual" for c in CODES}
START = {c: 1 for c in CODES}                 # the year the entity joins the group (1 = FY2026)
EXTRA_EVENTS: list = []                       # events that come with entities added after the story was written


def _refresh() -> None:
    """Rebuild the lookups from ENTITIES and TIERS in place, so modules holding them see the change."""
    NAME.clear(); NAME.update({e[0]: e[1] for e in ENTITIES})                       # noqa: E702
    PARENT.clear(); PARENT.update({e[0]: e[2] for e in ENTITIES})                   # noqa: E702
    HELD.clear(); HELD.update({e[0]: e[3] for e in ENTITIES})                       # noqa: E702
    GST_REG.clear(); GST_REG.update({e[0]: e[4] for e in ENTITIES})                 # noqa: E702
    CODES[:] = [e[0] for e in ENTITIES]
    TIER_MEMBERS.clear()
    TIER_MEMBERS.update({head: members(head) for head, _ in TIERS})


_BASE = (list(ENTITIES), list(TIERS), list(NODES), dict(STATUS), dict(START))


def reset() -> None:
    """Back to the eleven entities of the story."""
    ENTITIES[:], TIERS[:], NODES[:] = _BASE[0], _BASE[1], _BASE[2]
    STATUS.clear(); STATUS.update(_BASE[3])                                         # noqa: E702
    START.clear(); START.update(_BASE[4])                                           # noqa: E702
    EXTRA_EVENTS.clear()
    _refresh()


def add_entity(code: int, name: str, parent: int, held: float, gst: bool, role: str, status: str = "Planned",
               start: int = 1, capital: float = 0.0, external: float = 0.0) -> None:
    """What Add entity does: the entity goes under its parent, after the parent's last descendant, so the register
    stays in tree order. An entity its parent holds less than 100% of has outside investors, so it becomes an NCI node
    and a group of its own (named after it, as Partner LP is). Capital from the parent (and outside investors) in the
    year it joins is its opening entry."""
    if code in CODES:
        raise ValueError(f"{code} is already in the group")
    if parent not in CODES:
        raise ValueError(f"parent {parent} is not in the group")
    if not 0 < held <= 1:
        raise ValueError("the share held is more than 0% and at most 100%")
    if not 1 <= start <= PERIODS:
        raise ValueError("the entity joins in one of the model's years")
    after = max(i for i, e in enumerate(ENTITIES) if e[0] == parent or parent in ancestors(e[0]))
    ENTITIES.insert(after + 1, (code, name, parent, held, gst, role))
    STATUS[code], START[code] = status, start
    if held < 1:
        NODES.append(code)
        TIERS.append((code, name))
    _refresh()
    if capital or external:
        EXTRA_EVENTS.append(Capital(start, parent, code, capital, external))

# Group chart of accounts: code, name, class. Debit positive throughout.
ACCOUNTS = [
    (1000, "Cash", "asset"), (1100, "Receivables", "asset"), (1200, "Development work in progress", "asset"),
    (1300, "Investment property", "asset"), (1400, "Loans receivable", "asset"),
    (1500, "Investments in group entities", "asset"),
    (2000, "Payables", "liability"), (2050, "Intergroup AP/AR clearing (netting accounts)", "liability"),
    (2100, "Accruals", "liability"), (2200, "GST payable", "liability"), (2300, "Loans payable", "liability"),
    (2900, "Intercompany differences", "liability"),
    (3000, "Capital", "equity"), (3100, "Retained surplus at the start of the year", "equity"),
    (3200, "Distributions paid", "equity"), (3300, "Non-controlling interests", "equity"),
    (4000, "Property sales", "revenue"), (4100, "Construction revenue", "revenue"), (4200, "Fee revenue", "revenue"),
    (4300, "Rental revenue", "revenue"), (4400, "Interest income", "revenue"), (4500, "Distribution income", "revenue"),
    (5000, "Cost of property sold", "expense"), (5100, "Construction costs", "expense"),
    (5200, "Operating expenses", "expense"), (5300, "Interest expense", "expense"),
    (5900, "Surplus attributable to non-controlling interests", "attribution"),
]
ACCT = {a[0]: a[1] for a in ACCOUNTS}
CLASS = {a[0]: a[2] for a in ACCOUNTS}
PL = [a[0] for a in ACCOUNTS if a[2] in ("revenue", "expense")]
CASH, AR, WIP, PROP, LOANREC, INVEST = 1000, 1100, 1200, 1300, 1400, 1500
AP, CLR, ACCR, GSTPAY, LOANPAY, DIFF = 2000, 2050, 2100, 2200, 2300, 2900
CAPITAL, RE, DIST, NCI = 3000, 3100, 3200, 3300
SALES, CONSTR, FEES, RENT, INTINC, DISTINC = 4000, 4100, 4200, 4300, 4400, 4500
COS, CONSTCOST, OPEX, INTEXP, NCIPL = 5000, 5100, 5200, 5300, 5900

# Sites: code, name, units, then who holds the asset at each year end and in which account.
SITES = {
    "SA": {"name": "Site A (Dev LP A, sold to the Fund in FY2027)", "units": 20,
           "holder": [9006, 9008, 9008], "account": [WIP, PROP, PROP], "held": [20, 20, 15]},
    "SB": {"name": "Site B (Dev LP B)", "units": 12, "holder": [9007, 9007, 9007], "account": [WIP, WIP, WIP],
           "held": [12, 12, 12]},
    "SF": {"name": "Site F (Fund Dev LP)", "units": 10, "holder": [9009, 9009, 9009], "account": [WIP, WIP, WIP],
           "held": [10, 10, 10]},
    "SP": {"name": "Site P (Partner LP, sold to the Fund in FY2027)", "units": 8,
           "holder": [9010, 9008, 9008], "account": [WIP, PROP, PROP], "held": [8, 8, 8]},
}

# Register types (what Home Hub's matching hands the model)
TYPES = {
    "Trading, expensed": "Revenue against expense in full; balances eliminated; any GST the buyer cannot claim stays as a group cost",
    "Trading, capitalised": "Revenue against the seller's cost; the margin comes out of the asset at the tier holding both sides",
    "Asset sale": "Sale against cost of sale; the development margin comes out of the asset while it stays in the tier",
    "At cost, expensed": "Pass-through on the netting accounts: balances only, the P&L is untouched",
    "At cost, capitalised": "Pass-through into the buyer's WIP: balances only, no margin",
    "Moved at cost": "Asset moved between entities at cost (due diligence recharge): balances only",
    "Distribution": "Parent's share of a distribution against the sub's distributions paid",
    "Balance": "Period-end balance pair: receivable, loan or netting account against payable, loan or accrual",
}


@dataclass
class Book:
    """Entity journals and register rows built from the events."""
    lines: list = field(default_factory=list)          # (period, entity, account, amount, counterparty, site)
    register: list = field(default_factory=list)       # dicts
    investments: list = field(default_factory=list)    # (parent, sub, cost, year made)
    events: list = field(default_factory=list)

    def post(self, p, e, a, amt, cp=None, site=None):
        self.lines.append((p, e, a, float(amt), cp, site))

    def reg(self, p, kind, seller, buyer, site, s_acct, b_acct, s_amt, b_amt, gst=0.0, margin=0.0, cost_acct=None, desc=""):
        self.register.append({"period": p, "type": kind, "seller": seller, "buyer": buyer, "site": site,
                              "seller_acct": s_acct, "buyer_acct": b_acct, "cost_acct": cost_acct,
                              "seller_amt": float(s_amt), "buyer_amt": float(b_amt), "gst": float(gst),
                              "margin": float(margin), "desc": desc})


# ----------------------------------------------------------------------------------- events
@dataclass
class Event:
    p: int

    def entity(self, b: Book):                  # entity journals and register rows
        raise NotImplementedError

    def group(self, M: set, g: "GroupView"):    # the same event as the group M records it
        raise NotImplementedError


@dataclass
class External(Event):
    entity_code: int
    entries: list            # (account, amount), balanced, all with outsiders
    site: str | None = None

    def entity(self, b):
        for a, amt in self.entries:
            b.post(self.p, self.entity_code, a, amt, None, self.site)

    def group(self, M, g):
        if self.entity_code in M:
            for a, amt in self.entries:
                g.post(self.p, self.entity_code, a, amt)
                if a == WIP and self.site:
                    g.basis[self.site] += amt


@dataclass
class Capital(Event):
    parent: int | None
    sub: int
    amount: float            # from the parent
    external: float = 0.0    # from outside investors

    def entity(self, b):
        if self.parent is not None:
            b.post(self.p, self.parent, INVEST, self.amount, self.sub)
            b.post(self.p, self.parent, CASH, -self.amount)
            b.investments.append((self.parent, self.sub, self.amount, self.p))
        b.post(self.p, self.sub, CASH, self.amount + self.external)
        b.post(self.p, self.sub, CAPITAL, -(self.amount + self.external), self.parent)

    def group(self, M, g):
        inside = self.parent in M and self.sub in M
        if self.parent in M:
            if not inside:
                g.post(self.p, self.parent, INVEST, self.amount)
            g.post(self.p, self.parent, CASH, -self.amount)
        if self.sub in M:
            g.post(self.p, self.sub, CASH, self.amount + self.external)
            g.post(self.p, self.sub, CAPITAL, -(self.external if inside else self.amount + self.external))


@dataclass
class TradeCap(Event):
    """An intergroup charge the buyer capitalises into WIP: construction claims, development fees, interest."""
    seller: int
    buyer: int
    site: str
    amount: float
    cost: float              # the seller's directly attributable cost (0 for interest)
    rev_acct: int
    cost_acct: int | None
    settle: float            # paid in the period
    desc: str = ""
    via_loan: bool = False   # interest added to an intergroup loan

    def entity(self, b):
        p, s, u = self.p, self.seller, self.buyer
        b.post(p, s, self.rev_acct, -self.amount, u)
        if self.cost:
            b.post(p, s, self.cost_acct, self.cost, None, self.site)
            b.post(p, s, CASH, -self.cost)
        b.post(p, u, WIP, self.amount, s, self.site)
        if self.via_loan:
            b.post(p, s, LOANREC, self.amount, u)
            b.post(p, u, LOANPAY, -self.amount, s)
        else:
            b.post(p, s, AR, self.amount - self.settle, u)
            b.post(p, s, CASH, self.settle)
            b.post(p, u, AP, -(self.amount - self.settle), s)
            b.post(p, u, CASH, -self.settle)
        b.reg(p, "Trading, capitalised", s, u, self.site, self.rev_acct, WIP, self.amount, self.amount,
              margin=self.amount - self.cost, cost_acct=self.cost_acct if self.cost_acct else self.rev_acct, desc=self.desc)

    def group(self, M, g):
        p, s, u = self.p, self.seller, self.buyer
        if s in M and u in M:
            if self.cost:
                g.post(p, s, CASH, -self.cost)
                g.post(p, u, WIP, self.cost)
            g.basis[self.site] += self.cost
            if not self.via_loan:
                g.post(p, s, CASH, self.settle)
                g.post(p, u, CASH, -self.settle)
            return
        if s in M:
            g.post(p, s, self.rev_acct, -self.amount)
            if self.cost:
                g.post(p, s, self.cost_acct, self.cost)
                g.post(p, s, CASH, -self.cost)
            if self.via_loan:
                g.post(p, s, LOANREC, self.amount)
            else:
                g.post(p, s, AR, self.amount - self.settle)
                g.post(p, s, CASH, self.settle)
        if u in M:
            g.post(p, u, WIP, self.amount)
            g.basis[self.site] += self.amount
            if self.via_loan:
                g.post(p, u, LOANPAY, -self.amount)
            else:
                g.post(p, u, AP, -(self.amount - self.settle))
                g.post(p, u, CASH, -self.settle)


@dataclass
class TradeExp(Event):
    """An intergroup charge the buyer expenses, with GST the buyer may not be able to claim."""
    seller: int
    buyer: int
    amount: float
    rev_acct: int
    exp_acct: int
    settle: float
    desc: str = ""

    @property
    def gst(self):
        return round(self.amount * 0.15, 6) if GST_REG[self.seller] else 0.0

    def entity(self, b):
        p, s, u, gst = self.p, self.seller, self.buyer, self.gst
        claim = GST_REG[u]
        b.post(p, s, self.rev_acct, -self.amount, u)
        b.post(p, s, GSTPAY, -gst)
        b.post(p, s, AR, self.amount + gst - self.settle, u)
        b.post(p, s, CASH, self.settle)
        b.post(p, u, self.exp_acct, self.amount + (0 if claim else gst), s)
        if claim and gst:
            b.post(p, u, GSTPAY, gst)
        b.post(p, u, AP, -(self.amount + gst - self.settle), s)
        b.post(p, u, CASH, -self.settle)
        b.reg(p, "Trading, expensed", s, u, None, self.rev_acct, self.exp_acct, self.amount, self.amount, gst=gst, desc=self.desc)

    def group(self, M, g):
        p, s, u, gst = self.p, self.seller, self.buyer, self.gst
        claim = GST_REG[u]
        if s in M and u in M:
            g.post(p, s, GSTPAY, -gst)
            if claim and gst:
                g.post(p, u, GSTPAY, gst)
            else:
                g.post(p, u, self.exp_acct, gst)            # GST the buyer cannot claim is a cost to the group
            g.post(p, s, CASH, self.settle)
            g.post(p, u, CASH, -self.settle)
            return
        if s in M:
            g.post(p, s, self.rev_acct, -self.amount)
            g.post(p, s, GSTPAY, -gst)
            g.post(p, s, AR, self.amount + gst - self.settle)
            g.post(p, s, CASH, self.settle)
        if u in M:
            g.post(p, u, self.exp_acct, self.amount + (0 if claim else gst))
            if claim and gst:
                g.post(p, u, GSTPAY, gst)
            g.post(p, u, AP, -(self.amount + gst - self.settle))
            g.post(p, u, CASH, -self.settle)


@dataclass
class AtCost(Event):
    """Shared costs paid by one entity: its own share expensed, the others' shares coded to the netting accounts
    and on-charged at cost to the same accounts. A receiver expenses its share or capitalises it into WIP."""
    payer: int
    own: float
    shares: list             # (receiver, amount, "expensed" | "capitalised", site, billed, paid)
    desc: str = ""

    def entity(self, b):
        p, f = self.p, self.payer
        b.post(p, f, OPEX, self.own)
        b.post(p, f, CASH, -(self.own + sum(s[1] for s in self.shares)))
        for r, amt, treat, site, billed, paid in self.shares:
            acct = OPEX if treat == "expensed" else WIP
            b.post(p, f, CLR, amt, r)                 # the receiver's share, coded to its netting account
            b.post(p, f, CLR, -billed, r)             # the on-charge invoice, coded to the same account
            b.post(p, f, AR, billed - paid, r)
            b.post(p, f, CASH, paid)
            b.post(p, r, acct, amt, f, site)
            b.post(p, r, AP, -(billed - paid), f)
            b.post(p, r, CASH, -paid)
            if amt - billed:
                b.post(p, r, ACCR, -(amt - billed), f)        # not on-charged by the cut-off: the receiver accrues
            b.reg(p, "At cost, " + treat, f, r, site, CLR, acct, amt, amt, desc=self.desc)

    def group(self, M, g):
        p, f = self.p, self.payer
        if f in M:
            g.post(p, f, OPEX, self.own)
            g.post(p, f, CASH, -(self.own + sum(s[1] for s in self.shares)))
        for r, amt, treat, site, billed, paid in self.shares:
            acct = OPEX if treat == "expensed" else WIP
            if f in M and r in M:
                g.post(p, r, acct, amt)
                if acct == WIP:
                    g.basis[site] += amt
                g.post(p, f, CASH, paid)
                g.post(p, r, CASH, -paid)
                continue
            if f in M:
                g.post(p, f, CLR, amt - billed)
                g.post(p, f, AR, billed - paid)
                g.post(p, f, CASH, paid)
            if r in M:
                g.post(p, r, acct, amt)
                if acct == WIP:
                    g.basis[site] += amt
                g.post(p, r, AP, -(billed - paid))
                g.post(p, r, CASH, -paid)
                if amt - billed:
                    g.post(p, r, ACCR, -(amt - billed))


@dataclass
class BillAccrued(Event):
    """An on-charge accrued at the last cut-off, invoiced and paid this year."""
    payer: int
    receiver: int
    amount: float

    def entity(self, b):
        p, f, r, a = self.p, self.payer, self.receiver, self.amount
        b.post(p, f, CLR, -a, r)
        b.post(p, f, CASH, a)
        b.post(p, r, ACCR, a, f)
        b.post(p, r, CASH, -a)

    def group(self, M, g):
        p, f, r, a = self.p, self.payer, self.receiver, self.amount
        if f in M:
            g.post(p, f, CASH, a)
            if r not in M:
                g.post(p, f, CLR, -a)
        if r in M:
            g.post(p, r, CASH, -a)
            if f not in M:
                g.post(p, r, ACCR, a)


@dataclass
class MoveAtCost(Event):
    """A WIP balance moved to another entity at cost: a due diligence recharge once the business case is approved."""
    sender: int
    receiver: int
    site: str
    amount: float
    paid: float = 0.0

    def entity(self, b):
        p, s, r, a = self.p, self.sender, self.receiver, self.amount
        b.post(p, s, WIP, -a, None, self.site)
        b.post(p, s, AR, a - self.paid, r)
        b.post(p, s, CASH, self.paid)
        b.post(p, r, WIP, a, s, self.site)
        b.post(p, r, AP, -(a - self.paid), s)
        b.post(p, r, CASH, -self.paid)
        b.reg(p, "Moved at cost", s, r, self.site, WIP, WIP, a, a, desc="Due diligence recharged at cost on approval")

    def group(self, M, g):
        p, s, r, a = self.p, self.sender, self.receiver, self.amount
        if s in M and r in M:
            g.post(p, s, WIP, -a)
            g.post(p, r, WIP, a)
            g.post(p, s, CASH, self.paid)
            g.post(p, r, CASH, -self.paid)
            return
        if s in M:
            g.post(p, s, WIP, -a)
            g.post(p, s, AR, a - self.paid)
            g.post(p, s, CASH, self.paid)
            g.basis[self.site] -= a
        if r in M:
            g.post(p, r, WIP, a)
            g.post(p, r, AP, -(a - self.paid))
            g.post(p, r, CASH, -self.paid)
            g.basis[self.site] += a


@dataclass
class Settle(Event):
    """Cash settling an intergroup receivable and payable."""
    payer: int
    payee: int
    amount: float
    ledger: tuple = (AP, AR)        # payer's account, payee's account

    def entity(self, b):
        p = self.p
        b.post(p, self.payer, self.ledger[0], self.amount, self.payee)
        b.post(p, self.payer, CASH, -self.amount)
        b.post(p, self.payee, self.ledger[1], -self.amount, self.payer)
        b.post(p, self.payee, CASH, self.amount)

    def group(self, M, g):
        p = self.p
        both = self.payer in M and self.payee in M
        if self.payer in M:
            if not both:
                g.post(p, self.payer, self.ledger[0], self.amount)
            g.post(p, self.payer, CASH, -self.amount)
        if self.payee in M:
            if not both:
                g.post(p, self.payee, self.ledger[1], -self.amount)
            g.post(p, self.payee, CASH, self.amount)


@dataclass
class Loan(Event):
    lender: int
    borrower: int
    amount: float

    def entity(self, b):
        p = self.p
        b.post(p, self.lender, LOANREC, self.amount, self.borrower)
        b.post(p, self.lender, CASH, -self.amount)
        b.post(p, self.borrower, CASH, self.amount)
        b.post(p, self.borrower, LOANPAY, -self.amount, self.lender)

    def group(self, M, g):
        p = self.p
        both = self.lender in M and self.borrower in M
        if self.lender in M:
            if not both:
                g.post(p, self.lender, LOANREC, self.amount)
            g.post(p, self.lender, CASH, -self.amount)
        if self.borrower in M:
            g.post(p, self.borrower, CASH, self.amount)
            if not both:
                g.post(p, self.borrower, LOANPAY, -self.amount)


@dataclass
class AssetSale(Event):
    """A portfolio sold inside the group: the seller's WIP to the buyer's investment property."""
    seller: int
    buyer: int
    site: str
    price: float
    paid: float

    def entity(self, b):
        p, s, u = self.p, self.seller, self.buyer
        carrying = b.carrying(s, self.site, p)
        self.carrying = carrying
        b.post(p, s, SALES, -self.price, u)
        b.post(p, s, COS, carrying, None, self.site)
        b.post(p, s, WIP, -carrying, None, self.site)
        b.post(p, s, CASH, self.paid)
        b.post(p, s, AR, self.price - self.paid, u)
        b.post(p, u, PROP, self.price, s, self.site)
        b.post(p, u, CASH, -self.paid)
        b.post(p, u, AP, -(self.price - self.paid), s)
        b.reg(p, "Asset sale", s, u, self.site, SALES, PROP, self.price, self.price, margin=self.price - carrying,
              cost_acct=COS, desc="Portfolio sold inside the group")

    def group(self, M, g):
        p, s, u = self.p, self.seller, self.buyer
        if s in M and u in M:
            basis = g.basis[self.site]
            g.post(p, s, WIP, -basis)
            g.post(p, u, PROP, basis)
            g.post(p, s, CASH, self.paid)
            g.post(p, u, CASH, -self.paid)
            return
        if s in M:
            g.post(p, s, SALES, -self.price)
            g.post(p, s, COS, g.basis[self.site])
            g.post(p, s, WIP, -g.basis[self.site])
            g.post(p, s, CASH, self.paid)
            g.post(p, s, AR, self.price - self.paid)
            g.basis[self.site] = 0.0
        if u in M:
            g.post(p, u, PROP, self.price)
            g.post(p, u, CASH, -self.paid)
            g.post(p, u, AP, -(self.price - self.paid))
            g.basis[self.site] = self.price


@dataclass
class UnitSale(Event):
    """Units sold outside the group: the share of the site's carrying amount goes to cost of sale."""
    holder: int
    site: str
    fraction: float
    proceeds: float

    def entity(self, b):
        p, h = self.p, self.holder
        carrying = b.carrying(h, self.site, p, PROP) * self.fraction
        b.post(p, h, SALES, -self.proceeds)
        b.post(p, h, CASH, self.proceeds)
        b.post(p, h, COS, carrying, None, self.site)
        b.post(p, h, PROP, -carrying, None, self.site)

    def group(self, M, g):
        p, h = self.p, self.holder
        if h in M:
            cost = g.basis[self.site] * self.fraction
            g.post(p, h, SALES, -self.proceeds)
            g.post(p, h, CASH, self.proceeds)
            g.post(p, h, COS, cost)
            g.post(p, h, PROP, -cost)
            g.basis[self.site] -= cost


@dataclass
class Distribution(Event):
    sub: int
    parent: int
    total: float
    share: float             # to the parent; the rest to outside investors

    def entity(self, b):
        p = self.p
        b.post(p, self.sub, DIST, self.total, self.parent)
        b.post(p, self.sub, CASH, -self.total)
        b.post(p, self.parent, DISTINC, -self.share, self.sub)
        b.post(p, self.parent, CASH, self.share)
        b.reg(p, "Distribution", self.sub, self.parent, None, DIST, DISTINC, self.share, self.share,
              desc="Parent's share of the distribution")

    def group(self, M, g):
        p = self.p
        both = self.sub in M and self.parent in M
        if self.sub in M:
            g.post(p, self.sub, DIST, self.total - (self.share if both else 0))
            g.post(p, self.sub, CASH, -self.total)
        if self.parent in M:
            if not both:
                g.post(p, self.parent, DISTINC, -self.share)
            g.post(p, self.parent, CASH, self.share)


# ----------------------------------------------------------------------------------- the story
def events() -> list[Event]:
    """Three years of fictional transactions ($000), in the order they happen."""
    ev: list[Event] = []
    add = ev.append
    # FY2026: set-up, development starts
    add(Capital(1, None, 9000, 0, 6000))
    for parent, sub, amt, ext in ((9000, 9004, 3500, 0), (9000, 9001, 500, 0), (9000, 9002, 200, 0), (9000, 9003, 100, 0),
                                  (9004, 9005, 1800, 0), (9005, 9006, 900, 0), (9005, 9007, 700, 0),
                                  (9004, 9008, 1200, 800), (9008, 9009, 600, 0), (9004, 9010, 700, 300)):
        add(Capital(1, parent, sub, amt, ext))
    add(TradeCap(1, 9001, 9006, "SA", 2000, 1700, CONSTR, CONSTCOST, 1800, "Construction claims, site A"))
    add(TradeCap(1, 9001, 9007, "SB", 1000, 850, CONSTR, CONSTCOST, 900, "Construction claims, site B"))
    add(TradeCap(1, 9001, 9009, "SF", 600, 510, CONSTR, CONSTCOST, 600, "Construction claims, site F"))
    add(TradeCap(1, 9002, 9006, "SA", 300, 210, FEES, OPEX, 300, "Development management fee, site A"))
    add(TradeCap(1, 9002, 9007, "SB", 150, 105, FEES, OPEX, 150, "Development management fee, site B"))
    add(External(1, 9004, [(WIP, 120), (CASH, -120)], "SB"))                     # due diligence on site B
    add(MoveAtCost(1, 9004, 9007, "SB", 120))                                     # approved: recharged at cost
    add(External(1, 9004, [(OPEX, 40), (CASH, -40)]))                             # declined prospect, expensed
    add(Loan(1, 9004, 9007, 500))
    add(TradeCap(1, 9004, 9007, "SB", 40, 0, INTINC, None, 0, "Interest on the Holdings loan, capitalised", via_loan=True))
    add(AtCost(1, 9000, 100, [(9003, 80, "expensed", None, 80, 80), (9006, 120, "capitalised", "SA", 100, 100)],
               "Shared costs on-charged at cost"))
    add(External(1, 9003, [(FEES, -400), (OPEX, 300), (CASH, 100)]))
    add(External(1, 9001, [(CONSTR, -1000), (CONSTCOST, 850), (CASH, 150)]))
    add(External(1, 9010, [(WIP, 900), (CASH, -900)], "SP"))
    # FY2027: site A completes and is sold to the Fund; Partner LP sells site P to the Fund
    add(Settle(2, 9006, 9001, 200))
    add(Settle(2, 9007, 9001, 100))
    add(Settle(2, 9007, 9004, 120))
    add(BillAccrued(2, 9000, 9006, 20))
    add(TradeCap(2, 9001, 9006, "SA", 1000, 850, CONSTR, CONSTCOST, 1000, "Construction claims, site A"))
    add(TradeCap(2, 9002, 9006, "SA", 100, 70, FEES, OPEX, 100, "Development management fee, site A"))
    add(AtCost(2, 9000, 90, [(9003, 90, "expensed", None, 90, 90), (9007, 60, "capitalised", "SB", 60, 60)],
               "Shared costs on-charged at cost"))
    add(External(2, 9008, [(CASH, 3200), (LOANPAY, -3200)]))                    # bank facility for the purchases
    add(AssetSale(2, 9006, 9008, "SA", 4200, 3000))
    add(External(2, 9010, [(WIP, 100), (CASH, -100)], "SP"))
    add(AssetSale(2, 9010, 9008, "SP", 1300, 1300))
    add(External(2, 9008, [(RENT, -300), (OPEX, 60), (INTEXP, 96), (CASH, 144)]))
    add(TradeExp(2, 9003, 9008, 100, FEES, OPEX, 0, "Property management fee"))
    add(TradeCap(2, 9001, 9007, "SB", 1200, 1020, CONSTR, CONSTCOST, 1200, "Construction claims, site B"))
    add(TradeCap(2, 9004, 9007, "SB", 43.2, 0, INTINC, None, 0, "Interest on the Holdings loan, capitalised", via_loan=True))
    add(TradeCap(2, 9001, 9009, "SF", 800, 680, CONSTR, CONSTCOST, 800, "Construction claims, site F"))
    add(External(2, 9003, [(FEES, -420), (OPEX, 310), (CASH, 110)]))
    add(External(2, 9001, [(CONSTR, -900), (CONSTCOST, 765), (CASH, 135)]))
    # FY2028: five of site A's twenty homes sold outside the group; the Fund distributes
    add(UnitSale(3, 9008, "SA", 0.25, 1300))
    add(Settle(3, 9008, 9006, 1200))
    add(Settle(3, 9008, 9003, 115))
    add(Distribution(3, 9008, 9004, 200, 120))
    add(TradeExp(3, 9003, 9008, 120, FEES, OPEX, 138, "Property management fee"))
    add(External(3, 9008, [(RENT, -650), (OPEX, 130), (INTEXP, 192), (CASH, 328)]))
    add(TradeCap(3, 9001, 9007, "SB", 500, 425, CONSTR, CONSTCOST, 500, "Construction claims, site B"))
    add(TradeCap(3, 9004, 9007, "SB", 46.656, 0, INTINC, None, 0, "Interest on the Holdings loan, capitalised", via_loan=True))
    add(External(3, 9004, [(OPEX, 30), (CASH, -30)]))
    add(External(3, 9003, [(FEES, -440), (OPEX, 320), (CASH, 120)]))
    add(External(3, 9001, [(CONSTR, -800), (CONSTCOST, 680), (CASH, 120)]))
    add(TradeCap(3, 9001, 9009, "SF", 400, 340, CONSTR, CONSTCOST, 400, "Construction claims, site F"))
    return ev + EXTRA_EVENTS


def _carrying(book, entity, site, p, acct=WIP):
    return sum(amt for (pp, e, a, amt, cp, s) in book.lines if e == entity and s == site and a == acct and pp <= p)


def book() -> Book:
    b = Book()
    b.carrying = lambda e, s, p, acct=WIP: _carrying(b, e, s, p, acct)          # noqa: E731
    for ev in events():
        b.events.append(ev)
        ev.entity(b)
    _balances(b)
    return b


def _balances(b: Book) -> None:
    """Period-end balance pairs, as the matching hands them over (one row per pair and account pair)."""
    pairs = defaultdict(float)
    for p in range(1, PERIODS + 1):
        for (pp, e, a, amt, cp, s) in b.lines:
            if pp <= p and cp is not None and a in (AR, AP, LOANREC, LOANPAY, CLR, ACCR):
                pairs[(p, e, cp, a)] += amt
    rows = []
    for (p, e, cp, a), amt in sorted(pairs.items()):
        if CLASS[a] == "asset" or a == CLR:                 # the seller's side: receivable, loan or netting account
            if abs(amt) < 1e-9:
                continue
            other = {AR: AP, LOANREC: LOANPAY, CLR: ACCR}[a]
            mirror = -pairs.get((p, cp, e, other), 0.0)
            rows.append((p, e, cp, a, other, amt, mirror))
    for p, e, cp, a, other, amt, mirror in rows:
        b.reg(p, "Balance", e, cp, None, a, other, amt, mirror,
              desc={AR: "Receivable against payable", LOANREC: "Loan against loan", CLR: "Netting account against accrual"}[a])


# ----------------------------------------------------------------------------------- entity trial balances
def trial_balances(b: Book) -> dict:
    """(entity, account) -> [value per year]: closing balances for the balance sheet, the year's flows for the P&L,
    retained surplus at the start of the year and distributions paid in the year."""
    tb = {(e, a): [0.0] * PERIODS for e in CODES for a, _, _ in ACCOUNTS}
    for e in CODES:
        re_open = 0.0
        for p in range(1, PERIODS + 1):
            for (pp, ee, a, amt, cp, s) in b.lines:
                if ee != e:
                    continue
                if a in PL or a == DIST:
                    if pp == p:
                        tb[(e, a)][p - 1] += amt
                elif pp <= p:
                    tb[(e, a)][p - 1] += amt
            tb[(e, RE)][p - 1] = re_open
            re_open += sum(tb[(e, a)][p - 1] for a in PL) + tb[(e, DIST)][p - 1]
    return tb


# ----------------------------------------------------------------------------------- reference consolidation
class GroupView:
    def __init__(self):
        self.lines = []
        self.basis = defaultdict(float)          # group cost of each site's asset, as this group sees it

    def post(self, p, e, a, amt):
        self.lines.append((p, e, a, float(amt)))


def reference(b: Book) -> dict:
    """Each group's consolidated trial balance, built by re-recording every event as the group sees it.

    Returns {head: {"tb": {account: [per year]}, "basis": {site: [per year]}, "surplus": [...], "nci": [...],
    "nci_pl": [...], "equity": [...]}}. NCI is each node's share of the node group's net assets, less its share of the
    margin still unrealised on sales from the node's group to the rest of the tier (upstream)."""
    out = {}
    for head, _ in TIERS:
        M = TIER_MEMBERS[head]
        g = GroupView()
        basis_by_year = {s: [0.0] * PERIODS for s in SITES}
        for p in range(1, PERIODS + 1):
            for ev in b.events:
                if ev.p == p:
                    ev.group(M, g)
            for s in SITES:
                basis_by_year[s][p - 1] = g.basis[s] if SITES[s]["holder"][p - 1] in M else 0.0
        tb = {a: [0.0] * PERIODS for a, _, _ in ACCOUNTS}
        for p in range(1, PERIODS + 1):
            for (pp, e, a, amt) in g.lines:
                if a in PL or a == DIST:
                    if pp == p:
                        tb[a][p - 1] += amt
                elif pp <= p:
                    tb[a][p - 1] += amt
        surplus = [-sum(tb[a][i] for a in PL) for i in range(PERIODS)]
        out[head] = {"tb": tb, "basis": basis_by_year, "surplus": surplus, "members": M}
    # Net assets of each group (assets less liabilities), then NCI at the nodes
    for head in out:
        tb = out[head]["tb"]
        out[head]["net_assets"] = [sum(tb[a][i] for a, _, c in ACCOUNTS if c in ("asset", "liability")) for i in range(PERIODS)]
    for head, _ in TIERS:
        M = TIER_MEMBERS[head]
        nci = [0.0] * PERIODS
        nci_pl = [0.0] * PERIODS
        for node in NODES:
            if node == head or node not in M:
                continue
            share = 1 - HELD[node]
            up = upstream_unrealised(b, node, M)
            for i in range(PERIODS):
                nci[i] += share * out[node]["net_assets"][i] - share * up[i]
                nci_pl[i] += share * out[node]["surplus"][i] - share * (up[i] - (up[i - 1] if i else 0.0))
        out[head]["nci"] = nci
        out[head]["nci_pl"] = nci_pl
    return out


def unrealised(b: Book, row: dict, p: int) -> float:
    """Margin from one register row still held at the end of year p (before any tier test)."""
    if row["type"] not in ("Trading, capitalised", "Asset sale") or p < row["period"]:
        return 0.0
    site = SITES[row["site"]]
    start = site["held"][row["period"] - 2] if row["period"] > 1 else site["units"]
    return row["margin"] * site["held"][p - 1] / start


def in_tier(code, M):
    return code in M


def unrealised_tier(b: Book, M: set, p: int, rows=None) -> dict:
    """Unrealised margin by site at the end of year p for group M: both sides and the asset's holder inside M."""
    out = defaultdict(float)
    for row in rows if rows is not None else b.register:
        if row["seller"] in M and row["buyer"] in M and SITES.get(row["site"], {}).get("holder", [None] * 3)[p - 1] in M:
            out[row["site"]] += unrealised(b, row, p)
    return out


def upstream_unrealised(b: Book, node: int, M: set) -> list[float]:
    """Margin unrealised in tier M on sales from the node's group to entities outside it."""
    G = TIER_MEMBERS.get(node, {node})
    rows = [r for r in b.register if r["seller"] in G and r["buyer"] not in G]
    return [sum(unrealised_tier(b, M, p, rows).values()) for p in range(1, PERIODS + 1)]


if __name__ == "__main__":
    bk = book()
    tb = trial_balances(bk)
    for e in CODES:
        for p in range(PERIODS):
            assert abs(sum(tb[(e, a)][p] for a, _, _ in ACCOUNTS)) < 1e-6, (e, p)
    ref = reference(bk)
    for head, name in TIERS:
        r = ref[head]
        print(name, "surplus", [round(x, 3) for x in r["surplus"]], "NCI", [round(x, 3) for x in r["nci"]])
        print("   WIP", [round(x, 3) for x in r["tb"][WIP]], "PROP", [round(x, 3) for x in r["tb"][PROP]])
    print(len(bk.register), "register rows")
