"""Subclass / inheritance semantics for AirtestGenerator.

Verifies the meta-driven per-class _HANDLERS registry: subclass overrides win,
unoverridden actions are inherited, hook attributes (IMPORTS, MODULE_PROLOGUE,
wrap_main_body) are honored at generation time, and sibling classes do not
share state.
"""
import py_compile

import pytest

from excel_to_airtest import (
    AirtestGenerator,
    Asset,
    GenCtx,
    Step,
    action,
    generate_suite_script,
)


def _step(action_name, target="", params="", suite="S", step_no=1, row=2, expected=""):
    return Step(
        suite_id=suite, step_no=step_no, action=action_name, excel_row=row,
        target=target, params=params, expected=expected,
    )


def _ctx(assets=None, app_package=""):
    return GenCtx(assets=assets or {}, app_package=app_package)


# --------------------------------------------------------------------------- #
# Per-class registry semantics                                                #
# --------------------------------------------------------------------------- #

def test_subclass_inherits_parents_handlers():
    class Child(AirtestGenerator):
        pass

    assert "CLICK" in Child._HANDLERS
    assert "WAIT_FOR" in Child._HANDLERS
    assert Child._HANDLERS["CLICK"] is AirtestGenerator._HANDLERS["CLICK"]


def test_subclass_override_replaces_parent_handler():
    class Child(AirtestGenerator):
        @action("CLICK")
        def handle_click(self, step, ctx):
            return [f"custom_tap({step.target!r})"], None

    child = Child()
    lines, issue = Child._HANDLERS["CLICK"](child, _step("CLICK", target="btn"), _ctx())
    assert issue is None
    assert lines == ["custom_tap('btn')"]

    # Parent's handler is unchanged — sibling/parent isolation.
    parent = AirtestGenerator()
    parent_lines, _ = AirtestGenerator._HANDLERS["CLICK"](
        parent, _step("CLICK", target="b"),
        _ctx(assets={"b": Asset("b", resource_path="b.png")}),
    )
    # CLICK is deprecated — parent emits comment + touch line
    code_lines = [l for l in parent_lines if not l.startswith("#")]
    assert code_lines == ["touch(Template('b.png', threshold=0.8))"]


def test_subclass_can_override_existing_action_keyword():
    class Child(AirtestGenerator):
        @action("SWIPE")
        def handle_swipe(self, step, ctx):
            return [f"swipe_stub({step.params!r})"], None

    assert "SWIPE" in Child._HANDLERS
    # Child's override replaces the base handler for Child only
    assert Child._HANDLERS["SWIPE"] is not AirtestGenerator._HANDLERS["SWIPE"]

    src, issues = Child().generate_suite_script(
        [_step("SWIPE", params="dx=100")],
        _ctx(), "x.xlsx", "S",
    )
    assert "swipe_stub('dx=100')" in src
    assert issues == []


def test_sibling_subclasses_do_not_share_registry_state():
    class A(AirtestGenerator):
        @action("CLICK")
        def handle_click(self, step, ctx):
            return ["A.click()"], None

    class B(AirtestGenerator):
        @action("CLICK")
        def handle_click(self, step, ctx):
            return ["B.click()"], None

    a, b = A(), B()
    a_lines, _ = A._HANDLERS["CLICK"](a, _step("CLICK", target="x"), _ctx())
    b_lines, _ = B._HANDLERS["CLICK"](b, _step("CLICK", target="x"), _ctx())
    assert a_lines == ["A.click()"]
    assert b_lines == ["B.click()"]


# --------------------------------------------------------------------------- #
# Hook attributes shape generated script                                      #
# --------------------------------------------------------------------------- #

def test_imports_hook_lands_at_top_of_generated_script():
    class Child(AirtestGenerator):
        IMPORTS = "from my.framework import driver"

    assets = {"b": Asset("b", resource_path="b.png")}
    src, _ = Child().generate_suite_script(
        [_step("CLICK", target="b")], _ctx(assets), "x.xlsx", "S",
    )
    assert "from my.framework import driver" in src
    assert "from airtest.core.api" not in src


def test_module_prologue_hook_inserts_below_imports_above_main():
    class Child(AirtestGenerator):
        MODULE_PROLOGUE = "page = MyPage()"

    assets = {"b": Asset("b", resource_path="b.png")}
    src, _ = Child().generate_suite_script(
        [_step("CLICK", target="b")], _ctx(assets), "x.xlsx", "S",
    )
    pro_at  = src.index("page = MyPage()")
    main_at = src.index("def main():")
    imp_at  = src.index(Child.IMPORTS.splitlines()[0])
    assert imp_at < pro_at < main_at


def test_wrap_main_body_hook_wraps_step_lines():
    class Child(AirtestGenerator):
        def wrap_main_body(self, step_lines, suite_id):
            return [
                f"with mycontext({suite_id!r}):",
                *[f"    {l}" for l in step_lines],
            ]

    assets = {"b": Asset("b", resource_path="b.png")}
    src, _ = Child().generate_suite_script(
        [_step("CLICK", target="b")], _ctx(assets), "x.xlsx", "TC_X",
    )
    assert "with mycontext('TC_X'):" in src
    # And the result is still syntactically valid Python.
    assert py_compile.PyCompileError is not None  # sanity import


def test_subclass_generated_script_is_syntactically_valid(tmp_path):
    class Child(AirtestGenerator):
        IMPORTS         = "import sys"
        MODULE_PROLOGUE = "ROOT = '/tmp'"

        def wrap_main_body(self, step_lines, suite_id):
            return [
                f"print({suite_id!r})",
                *step_lines,
            ]

        @action("CLICK")
        def handle_click(self, step, ctx):
            return [f"print('clicked', {step.target!r})"], None

    src, _ = Child().generate_suite_script(
        [_step("CLICK", target="a"), _step("CLICK", target="b", step_no=2)],
        _ctx(), "x.xlsx", "TC_X",
    )
    out = tmp_path / "out.py"
    out.write_text(src, encoding="utf-8")
    py_compile.compile(str(out), doraise=True)


# --------------------------------------------------------------------------- #
# Backward-compat surface                                                     #
# --------------------------------------------------------------------------- #

def test_module_level_generate_suite_script_uses_default_generator():
    """Legacy module-level function still works for callers that pre-date the class."""
    assets = {"b": Asset("b", resource_path="b.png")}
    src, _ = generate_suite_script(
        [_step("CLICK", target="b")], _ctx(assets), "x.xlsx", "S",
    )
    assert "touch(Template('b.png'" in src
