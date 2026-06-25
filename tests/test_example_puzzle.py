"""Regression anchor for the shipped generic example project.

Guards that projects/example_puzzle.py registers, generates valid framework-free
Airtest, and stays free of game-specific tokens.

The module is loaded by file path under a private module name (not as
``projects.example_puzzle``) so this test never binds the ``projects`` package
in ``sys.modules`` — that would break test_registry's dynamic-discovery test,
which expects ``projects`` to resolve to a temp directory.
"""

import ast
import importlib.util
import os

from excel_to_airtest import Asset, GenCtx, LOCATOR_IMAGE, Step

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load_example():
    path = os.path.join(ROOT, "projects", "example_puzzle.py")
    spec = importlib.util.spec_from_file_location("_example_puzzle_under_test", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # fires @register_project("example_puzzle")
    return mod


def _ctx(app_package="com.example.puzzle"):
    assets = {
        "btn_play": Asset(
            object_id="btn_play",
            locator_type=LOCATOR_IMAGE,
            resource_path="./assets/btn_play.png",
            threshold=0.85,
            timeout=5,
        )
    }
    return GenCtx(assets=assets, app_package=app_package)


def _steps():
    return [
        Step("TC_DEMO", 1, "START_APP", 2, "", "", "App opens"),
        Step("TC_DEMO", 2, "TAP", 3, "btn_play", "", "Game starts"),
    ]


def test_example_puzzle_registers():
    from registry import _PROJECT_REGISTRY

    _load_example()
    assert "example_puzzle" in _PROJECT_REGISTRY


def test_example_puzzle_generates_valid_framework_free_script():
    mod = _load_example()
    gen = mod.ExamplePuzzleGenerator()
    script, issues = gen.generate_suite_script(_steps(), _ctx(), "demo.xlsx", "TC_DEMO")

    assert not issues
    ast.parse(script)  # valid Python
    assert "start_app('com.example.puzzle')" in script
    assert "from airtest.core.api import *" in script
    assert "try:" in script and "snapshot(" in script and "raise" in script
    # No game framework leaked into generated output.
    assert "pixon" not in script
    assert "wrapper." not in script


def test_example_puzzle_comment_only_suite_is_valid_python():
    """A suite whose steps all resolve to comments (e.g. READ_TEXT TODO) must
    still produce a syntactically valid try-block, not an empty one."""
    mod = _load_example()
    gen = mod.ExamplePuzzleGenerator()
    steps = [Step("TC_TODO", 1, "READ_TEXT", 2, "score_text", "", "read score")]
    script, _ = gen.generate_suite_script(steps, _ctx(), "demo.xlsx", "TC_TODO")

    ast.parse(script)  # would raise SyntaxError on an empty try-block
    assert "try:" in script
    assert "pass" in script  # guard line keeps the try-block non-empty
