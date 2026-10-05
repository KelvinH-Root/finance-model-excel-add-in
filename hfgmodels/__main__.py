"""Command line: python -m hfgmodels build-example [--recipe PATH] [--out PATH] [--template PATH] [--no-values] [--verify]

Builds the development test example, which proves the template editing layer,
the layout engine and the verification harness. It is not a product."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main(argv=None):
    p = argparse.ArgumentParser(prog="hfgmodels")
    sub = p.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build-example", help="build the development test example")
    b.add_argument("--recipe", default=str(ROOT / "examples/development/recipe.yaml"))
    b.add_argument("--out", default=str(ROOT / "build/development_example.xlsx"))
    b.add_argument("--template", default=None, help="path to Budget_Template.xlsx (default: $HFG_TEMPLATE)")
    b.add_argument("--no-values", action="store_true", help="skip the LibreOffice recalculation")
    b.add_argument("--verify", action="store_true", help="compare the workbook with the reference calculation")
    a = p.parse_args(argv)
    if a.cmd == "build-example":
        sys.path.insert(0, str(ROOT))
        from examples.development.assemble import build
        r, model, sheets, cw = build(a.recipe, a.out, a.template, cache_values=not a.no_values)
        print(f"built {a.out}")
        if a.verify:
            from examples.development.compare import compare
            from examples.development.reference import calculate
            from .verify import recalculate
            cw = cw or recalculate(a.out, [s.name for s in model.sheets])
            probs = compare(model, cw, calculate(r), r.sites)
            errs = {k: v for k, v in cw.errors.items() if v}
            for k, v in errs.items():
                print(f"formula errors on {k}: {v[:10]}")
            sh, row = model.index["chk.errors"]
            n_err = cw.get(sh, row, 8)
            sh, row = model.index["chk.alerts"]
            n_alt = cw.get(sh, row, 8)
            print(f"reference mismatches: {len(probs)}; model errors: {n_err}; alerts: {n_alt}")
            for x in probs[:30]:
                print("  " + x)
            if probs or errs or n_err:
                return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
