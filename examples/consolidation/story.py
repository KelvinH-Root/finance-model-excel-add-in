"""The consolidation demo: a fictional group shaped like HFG's, month by month, and an independent
consolidation of it.

    python -m examples.consolidation.story        # prints each group's surplus and NCI by year end

The Phase 0 proof (prototypes/consolidation/group.py) told this story by year. Here the same
eleven entities and the same transactions happen in months of FY2026 to FY2028 (April 2025 to
March 2028), and the accounts are those of the HFG library's group chart (the full financial
model's statement lines, with the lines a consolidation needs). Everything is fictional ($000).

Entity figures are handed over as Home Hub would: for every entity and group account an opening
balance (the month before the model starts) and the month's movement, debit positive. The
register is what Home Hub's matching hands over: transactions (each with the month it falls in)
and balance pairs (each side's balance at every month end).

The reference consolidation never reads the register. For each group it re-records every event as
the group sees it (an intra-group charge capitalised at the seller's cost, an intra-group sale moving
the asset at group cost, balances and the at-cost netting accounts gone), tracking each site's group
cost, and works out non-controlling interests from each node's net assets. The workbook's
elimination entries have to land on the same numbers.

Figures are cumulative from the model's start: a balance sheet account at a month end is its
balance, an income statement account is everything since the start (so each month's trial balance
nets to nil). The statements take the month's movement from it.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

T = 36                      # April 2025 to March 2028
START = (2025, 4)
FY_END = 3

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

# Groups that produce accounts. An entity held less than 100% by its parent is an NCI node and a
# group of its own (Housing Fund LP is both).
GROUPS = [(9000, "Foundation group"), (9004, "Holdings group"), (9005, "Devco group"), (9008, "Fund group")]
NODES = [c for c in CODES if HELD[c] < 1]
TIERS = GROUPS + [(n, NAME[n]) for n in NODES if n not in dict(GROUPS)]


def ancestors(code: int) -> list[int]:
    out, p = [], PARENT[code]
    while p is not None:
        out.append(p)
        p = PARENT[p]
    return out


def members(head: int) -> set[int]:
    return {c for c in CODES if c == head or head in ancestors(c)}


TIER_MEMBERS = {head: members(head) for head, _ in TIERS}

# The library's group chart (library/hfg/consolidation.yaml): the accounts this story posts to.
CASH, AR, WIP, PROP, LOANREC, INVEST = 1000, 1100, 1250, 1450, 1600, 1700
AP, CLR, ACCR, GSTPAY, LOANPAY, DIFF = 2000, 2050, 2100, 2200, 2500, 2900
CAPITAL, RE, DIST, NCI = 3000, 3200, 3300, 3400
SALES, CONSTR, FEES, RENT, INTINC, DISTINC = 4000, 4100, 4200, 4300, 4600, 4700
COS, CONSTCOST, OPEX, INTEXP, NCIPL = 5000, 5100, 5300, 5700, 5900
INCOME_STATEMENT = {SALES, CONSTR, FEES, RENT, 4500, INTINC, DISTINC, COS, CONSTCOST, 5200, OPEX, 5400, 5500, 5600, INTEXP,
                    5800}

# Sites: code, name, homes, then month by month who holds the asset, in which account, and how many homes are still held.


def _months(*runs):
    out = []
    for value, n in runs:
        out += [value] * n
    assert len(out) == T
    return out


SITES = {
    "SA": {"name": "Site A (Dev LP A, sold to the Fund in October 2026)", "units": 20,
           "holder": _months((9006, 18), (9008, 18)), "account": _months((WIP, 18), (PROP, 18)), "held": _months((20, 26), (15, 10))},
    "SB": {"name": "Site B (Dev LP B)", "units": 12, "holder": _months((9007, T)), "account": _months((WIP, T)), "held": _months((12, T))},
    "SF": {"name": "Site F (Fund Dev LP)", "units": 10, "holder": _months((9009, T)), "account": _months((WIP, T)),
           "held": _months((10, T))},
    "SP": {"name": "Site P (Partner LP, sold to the Fund in November 2026)", "units": 8,
           "holder": _months((9010, 19), (9008, 17)), "account": _months((WIP, 19), (PROP, 17)), "held": _months((8, T))},
}

TYPES = ["Trading, expensed", "Trading, capitalised", "Asset sale", "At cost, expensed", "At cost, capitalised",
         "Moved at cost", "Distribution", "Balance"]


@dataclass
class Book:
    """Entity journals and register rows built from the events."""
    lines: list = field(default_factory=list)          # (month, entity, account, amount, counterparty, site)
    register: list = field(default_factory=list)       # transactions, dicts
    balances: list = field(default_factory=list)       # balance pairs, dicts with a value for every month
    investments: list = field(default_factory=list)    # (parent, sub, cost, month made)
    events: list = field(default_factory=list)

    def post(self, p, e, a, amt, cp=None, site=None):
        self.lines.append((p, e, a, float(amt), cp, site))

    def reg(self, p, kind, seller, buyer, site, s_acct, b_acct, s_amt, b_amt, gst=0.0, margin=0.0, cost_acct=None, desc=""):
        self.register.append({"period": p, "type": kind, "seller": seller, "buyer": buyer, "site": site,
                              "seller_acct": s_acct, "buyer_acct": b_acct, "cost_acct": cost_acct,
                              "seller_amt": float(s_amt), "buyer_amt": float(b_amt), "gst": float(gst),
                              "margin": float(margin), "desc": desc})

    def carrying(self, entity, site, p, acct=WIP):
        return sum(amt for (pp, e, a, amt, cp, s) in self.lines if e == entity and s == site and a == acct and pp <= p)


class GroupView:
    def __init__(self):
        self.lines = []
        self.basis = defaultdict(float)          # group cost of each site's asset, as this group sees it

    def post(self, p, e, a, amt):
        self.lines.append((p, e, a, float(amt)))


# ----------------------------------------------------------------------------------- events
@dataclass
class Event:
    p: int

    def entity(self, b: Book):
        raise NotImplementedError

    def group(self, M: set, g: GroupView):
        raise NotImplementedError


@dataclass
class External(Event):
    entity_code: int
    entries: list
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
    amount: float
    external: float = 0.0

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
    seller: int
    buyer: int
    site: str
    amount: float
    cost: float
    rev_acct: int
    cost_acct: int | None
    settle: float
    desc: str = ""
    via_loan: bool = False

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
                g.post(p, u, self.exp_acct, gst)
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
            b.post(p, f, CLR, amt, r)
            b.post(p, f, CLR, -billed, r)
            b.post(p, f, AR, billed - paid, r)
            b.post(p, f, CASH, paid)
            b.post(p, r, acct, amt, f, site)
            b.post(p, r, AP, -(billed - paid), f)
            b.post(p, r, CASH, -paid)
            if amt - billed:
                b.post(p, r, ACCR, -(amt - billed), f)
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
    payer: int
    payee: int
    amount: float
    ledger: tuple = (AP, AR)

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
    seller: int
    buyer: int
    site: str
    price: float
    paid: float

    def entity(self, b):
        p, s, u = self.p, self.seller, self.buyer
        carrying = b.carrying(s, self.site, p)
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
    share: float

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
    """Three years of fictional transactions ($000), each in the month it happens (1 = April 2025)."""
    ev: list[Event] = []
    add = ev.append
    # FY2026: set-up, development starts
    add(Capital(1, None, 9000, 0, 6000))
    for parent, sub, amt, ext in ((9000, 9004, 3500, 0), (9000, 9001, 500, 0), (9000, 9002, 200, 0), (9000, 9003, 100, 0),
                                  (9004, 9005, 1800, 0), (9005, 9006, 900, 0), (9005, 9007, 700, 0),
                                  (9004, 9008, 1200, 800), (9008, 9009, 600, 0), (9004, 9010, 700, 300)):
        add(Capital(1, parent, sub, amt, ext))
    add(External(2, 9004, [(WIP, 120), (CASH, -120)], "SB"))                     # due diligence on site B
    add(External(2, 9004, [(OPEX, 40), (CASH, -40)]))                             # declined prospect, expensed
    add(MoveAtCost(3, 9004, 9007, "SB", 120))                                     # approved: recharged at cost
    add(Loan(3, 9004, 9007, 500))
    add(TradeCap(4, 9001, 9006, "SA", 2000, 1700, CONSTR, CONSTCOST, 1800, "Construction claims, site A"))
    add(TradeCap(5, 9001, 9007, "SB", 1000, 850, CONSTR, CONSTCOST, 900, "Construction claims, site B"))
    add(TradeCap(6, 9001, 9009, "SF", 600, 510, CONSTR, CONSTCOST, 600, "Construction claims, site F"))
    add(TradeCap(7, 9002, 9006, "SA", 300, 210, FEES, OPEX, 300, "Development management fee, site A"))
    add(TradeCap(7, 9002, 9007, "SB", 150, 105, FEES, OPEX, 150, "Development management fee, site B"))
    add(External(8, 9010, [(WIP, 900), (CASH, -900)], "SP"))
    add(AtCost(9, 9000, 100, [(9003, 80, "expensed", None, 80, 80), (9006, 120, "capitalised", "SA", 100, 100)],
               "Shared costs on-charged at cost"))
    add(External(10, 9003, [(FEES, -400), (OPEX, 300), (CASH, 100)]))
    add(External(11, 9001, [(CONSTR, -1000), (CONSTCOST, 850), (CASH, 150)]))
    add(TradeCap(12, 9004, 9007, "SB", 40, 0, INTINC, None, 0, "Interest on the Holdings loan, capitalised", via_loan=True))
    # FY2027: site A completes and is sold to the Fund; Partner LP sells site P to the Fund
    add(Settle(13, 9006, 9001, 200))
    add(Settle(13, 9007, 9001, 100))
    add(BillAccrued(13, 9000, 9006, 20))
    add(Settle(14, 9007, 9004, 120))
    add(TradeCap(15, 9001, 9006, "SA", 1000, 850, CONSTR, CONSTCOST, 1000, "Construction claims, site A"))
    add(TradeCap(16, 9002, 9006, "SA", 100, 70, FEES, OPEX, 100, "Development management fee, site A"))
    add(TradeCap(16, 9001, 9009, "SF", 800, 680, CONSTR, CONSTCOST, 800, "Construction claims, site F"))
    add(AtCost(17, 9000, 90, [(9003, 90, "expensed", None, 90, 90), (9007, 60, "capitalised", "SB", 60, 60)],
               "Shared costs on-charged at cost"))
    add(External(18, 9008, [(CASH, 3200), (LOANPAY, -3200)]))                   # bank facility for the purchases
    add(External(18, 9010, [(WIP, 100), (CASH, -100)], "SP"))
    add(AssetSale(19, 9006, 9008, "SA", 4200, 3000))
    add(AssetSale(20, 9010, 9008, "SP", 1300, 1300))
    add(TradeExp(21, 9003, 9008, 100, FEES, OPEX, 0, "Property management fee"))
    add(External(22, 9008, [(RENT, -300), (OPEX, 60), (INTEXP, 96), (CASH, 144)]))
    add(TradeCap(23, 9001, 9007, "SB", 1200, 1020, CONSTR, CONSTCOST, 1200, "Construction claims, site B"))
    add(External(23, 9003, [(FEES, -420), (OPEX, 310), (CASH, 110)]))
    add(TradeCap(24, 9004, 9007, "SB", 43.2, 0, INTINC, None, 0, "Interest on the Holdings loan, capitalised", via_loan=True))
    add(External(24, 9001, [(CONSTR, -900), (CONSTCOST, 765), (CASH, 135)]))
    # FY2028: five of site A's twenty homes sold outside the group; the Fund distributes
    add(Settle(25, 9008, 9003, 115))
    add(Settle(26, 9008, 9006, 1200))
    add(UnitSale(27, 9008, "SA", 0.25, 1300))
    add(TradeExp(28, 9003, 9008, 120, FEES, OPEX, 138, "Property management fee"))
    add(TradeCap(29, 9001, 9007, "SB", 500, 425, CONSTR, CONSTCOST, 500, "Construction claims, site B"))
    add(Distribution(30, 9008, 9004, 200, 120))
    add(External(31, 9004, [(OPEX, 30), (CASH, -30)]))
    add(TradeCap(32, 9001, 9009, "SF", 400, 340, CONSTR, CONSTCOST, 400, "Construction claims, site F"))
    add(External(33, 9008, [(RENT, -650), (OPEX, 130), (INTEXP, 192), (CASH, 328)]))
    add(External(34, 9003, [(FEES, -440), (OPEX, 320), (CASH, 120)]))
    add(External(35, 9001, [(CONSTR, -800), (CONSTCOST, 680), (CASH, 120)]))
    add(TradeCap(36, 9004, 9007, "SB", 46.656, 0, INTINC, None, 0, "Interest on the Holdings loan, capitalised", via_loan=True))
    return ev


def book() -> Book:
    b = Book()
    for ev in events():
        b.events.append(ev)
        ev.entity(b)
    _balances(b)
    return b


BALANCE_PAIRS = {AR: AP, LOANREC: LOANPAY, CLR: ACCR}
PAIR_DESC = {AR: "Receivable against payable", LOANREC: "Loan against loan", CLR: "Netting account against accrual"}


def _balances(b: Book) -> None:
    """Balance pairs as the matching hands them over: one row per pair of entities and accounts, with each side's balance
    at every month end (the receivable side as it stands, the payable side as its negative)."""
    keys = sorted({(e, cp, a) for (pp, e, a, amt, cp, s) in b.lines if cp is not None and a in BALANCE_PAIRS})
    for e, cp, a in keys:
        other = BALANCE_PAIRS[a]
        side = [0.0] * T
        mirror = [0.0] * T
        for t in range(T):
            side[t] = sum(amt for (pp, ee, aa, amt, c, s) in b.lines if ee == e and c == cp and aa == a and pp <= t + 1)
            mirror[t] = -sum(amt for (pp, ee, aa, amt, c, s) in b.lines if ee == cp and c == e and aa == other and pp <= t + 1)
        if all(abs(x) < 1e-9 for x in side + mirror):
            continue
        b.balances.append({"seller": e, "buyer": cp, "seller_acct": a, "buyer_acct": other, "seller": e,
                           "seller_bal": [round(x, 9) for x in side], "buyer_bal": [round(x, 9) for x in mirror],
                           "desc": PAIR_DESC[a]})


# ----------------------------------------------------------------------------------- what the entities hand over
def figures(b: Book, accounts: list[int]) -> dict:
    """(entity, account) -> (opening, [movement per month]), debit positive."""
    out = {}
    for e in CODES:
        for a in accounts:
            mv = [0.0] * T
            for (pp, ee, aa, amt, cp, s) in b.lines:
                if ee == e and aa == a:
                    mv[pp - 1] += amt
            out[(e, a)] = (0.0, [round(x, 9) for x in mv])
    return out


# ----------------------------------------------------------------------------------- reference consolidation
def unrealised(row: dict, t: int) -> float:
    """Margin from one register row still held at the end of month t (any group), before the group tests."""
    if row["type"] not in ("Trading, capitalised", "Asset sale") or t < row["period"]:
        return 0.0
    site = SITES[row["site"]]
    start = site["held"][row["period"] - 2] if row["period"] > 1 else site["units"]
    return row["margin"] * site["held"][t - 1] / start


def unrealised_tier(b: Book, M: set, t: int, rows=None) -> dict:
    out = defaultdict(float)
    for row in rows if rows is not None else b.register:
        if row["seller"] in M and row["buyer"] in M and SITES.get(row["site"], {}).get("holder", [None] * T)[t - 1] in M:
            out[row["site"]] += unrealised(row, t)
    return out


def upstream_unrealised(b: Book, node: int, M: set) -> list[float]:
    G = TIER_MEMBERS.get(node, {node})
    rows = [r for r in b.register if r["seller"] in G and r["buyer"] not in G]
    return [sum(unrealised_tier(b, M, t, rows).values()) for t in range(1, T + 1)]


def reference(b: Book) -> dict:
    """Each group's consolidated trial balance at every month end (cumulative from the start), by re-recording every
    event as the group sees it. {head: {"tb": {account: [36]}, "surplus": [36] (since the start), "net_assets": [36],
    "nci": [36], "nci_pl": [36] (since the start)}}."""
    out = {}
    for head, _ in TIERS:
        M = TIER_MEMBERS[head]
        g = GroupView()
        for ev in b.events:
            ev.group(M, g)
        tb = defaultdict(lambda: [0.0] * T)
        for (pp, e, a, amt) in g.lines:
            for t in range(pp - 1, T):
                tb[a][t] += amt
        surplus = [-sum(tb[a][t] for a in INCOME_STATEMENT) for t in range(T)]
        out[head] = {"tb": tb, "surplus": surplus, "members": M}
    for head in out:
        tb = out[head]["tb"]
        out[head]["net_assets"] = [sum(v[t] for a, v in tb.items() if a < 3000) for t in range(T)]
    for head, _ in TIERS:
        M = TIER_MEMBERS[head]
        nci = [0.0] * T
        nci_pl = [0.0] * T
        for node in NODES:
            if node == head or node not in M:
                continue
            share = 1 - HELD[node]
            up = upstream_unrealised(b, node, M)
            for t in range(T):
                nci[t] += share * out[node]["net_assets"][t] - share * up[t]
                nci_pl[t] += share * out[node]["surplus"][t] - share * up[t]
        out[head]["nci"] = nci
        out[head]["nci_pl"] = nci_pl
    return out


if __name__ == "__main__":
    bk = book()
    ref = reference(bk)
    for head, name in TIERS:
        r = ref[head]
        print(f"{name:18}", "surplus", [round(r["surplus"][t], 3) for t in (11, 23, 35)],
              "NCI", [round(r["nci"][t], 3) for t in (11, 23, 35)])
    print(len(bk.register), "transactions,", len(bk.balances), "balance pairs,", len(bk.investments), "investments")
