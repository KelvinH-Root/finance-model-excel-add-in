"""Independent calculation of the assembly demo's economics (no formulas involved).

Used by tests/test_assembly.py to check built workbooks and by the Impacts proof to check the
live round trip: change one setting, and the statements must move as this calculation says.
"""


def reference(model):
    """Independent calculation of the demo economics (no formulas involved)."""
    n = model.periods
    rev = [0.0] * n
    opex = [0.0] * n
    interest = [0.0] * n
    flow = [0.0] * n
    debt = [0.0] * n
    for inst in model.instances:
        s = inst.settings
        if inst.module == "demo.revenue_line":
            for t in range(n):
                rev[t] += s["base"] * (1 + s["growth"]) ** t
        elif inst.module == "demo.cost_line":
            for t in range(n):
                opex[t] += s["amount"] * (1 + s["inflation"]) ** t
        elif inst.module == "demo.facility":
            bal = 0.0
            for t in range(n):
                draw = s["amount"] if t == 0 else 0.0
                repay = min(bal, s["instalment"]) if t > 0 else 0.0
                interest[t] += bal * s["rate"] / 12
                bal = bal + draw - repay
                flow[t] += draw - repay
                debt[t] += bal
    receipts = [0.0] + rev[:-1]
    cash, equity, debtors = [], [], []
    c = e = d = 0.0
    for t in range(n):
        c += receipts[t] - opex[t] - interest[t] + flow[t]
        e += rev[t] - opex[t] - interest[t]
        d += rev[t] - receipts[t]
        cash.append(c)
        equity.append(e)
        debtors.append(d)
    return {"revenue": rev, "opex": opex, "interest": interest, "receipts": receipts, "cash": cash,
            "assets": [a + b for a, b in zip(cash, debtors)], "debt": debt, "equity": equity}
