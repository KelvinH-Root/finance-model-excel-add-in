import os
import shutil
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
RECIPE = ROOT / "examples/development/recipe.yaml"


@pytest.fixture(scope="session")
def recipe():
    from examples.development.recipe import load
    return load(RECIPE)


@pytest.fixture(scope="session")
def reference(recipe):
    from examples.development.reference import calculate
    return calculate(recipe)


def template_or_skip():
    p = os.environ.get("HFG_TEMPLATE")
    if not p or not Path(p).exists():
        pytest.skip("HFG_TEMPLATE not set (Budget_Template.xlsx is not kept in git)")
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice not installed")
    return p
