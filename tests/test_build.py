"""Test example: assemble the development example into the template, recalculate in LibreOffice and
compare every row with the reference calculation. Skipped unless HFG_TEMPLATE
points at Budget_Template.xlsx and LibreOffice is installed."""

from conftest import RECIPE, template_or_skip


def test_build_matches_reference(tmp_path, recipe, reference):
    template = template_or_skip()
    from examples.development.assemble import assemble
    from examples.development.compare import compare
    from hfgmodels.verify import recalculate

    out = tmp_path / "model.xlsx"
    r, model, sheets = assemble(RECIPE, out, template)
    cw = recalculate(out, [s.name for s in model.sheets])
    assert not {k: v for k, v in cw.errors.items() if v}, "formula errors in the new sheets"
    problems = compare(model, cw, reference, r.sites)
    assert not problems, "\n".join(problems[:20])
    sh, row = model.index["chk.errors"]
    assert cw.get(sh, row, 8) == 0
