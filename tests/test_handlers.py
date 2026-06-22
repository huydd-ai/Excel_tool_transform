"""Action handlers, _resolve_image dispatch, generator dispatch, injection guard.

Handlers are methods on AirtestGenerator. The module-level `_HANDLERS` dict
holds the unbound functions; tests invoke them by passing a generator instance
as the first argument: `_HANDLERS["CLICK"](gen, step, ctx)`.
"""

import ast

import pytest

from excel_to_airtest import (
    LOCATOR_IMAGE,
    LOCATOR_OCR,
    AirtestGenerator,
    Asset,
    GenCtx,
    Step,
    _HANDLERS,
    _resolve_image,
    generate_suite_script,
)


@pytest.fixture
def gen():
    return AirtestGenerator()


def _step(action, target="", params="", suite="S", step_no=1, row=2, expected=""):
    return Step(
        suite_id=suite,
        step_no=step_no,
        action=action,
        excel_row=row,
        target=target,
        params=params,
        expected=expected,
    )


def _ctx(assets=None, app_package=""):
    return GenCtx(assets=assets or {}, app_package=app_package)


# --------------------------------------------------------------------------- #
# _resolve_image branches                                                     #
# --------------------------------------------------------------------------- #


def test_resolve_image_flags_missing_target():
    from models import AirtestError

    with pytest.raises(AirtestError, match="MISSING_TARGET"):
        _resolve_image(_step("CLICK"), _ctx())


def test_resolve_image_flags_unknown_target():
    from models import AirtestError

    with pytest.raises(AirtestError, match="UNKNOWN_TARGET"):
        _resolve_image(_step("CLICK", target="zzz"), _ctx())


def test_resolve_image_rejects_non_image_locator():
    assets = {"x": Asset("x", locator_type=LOCATOR_OCR, resource_path="NONE")}
    asset, lines, issue = _resolve_image(_step("CLICK", target="x"), _ctx(assets))
    assert asset is None
    assert "UNSUPPORTED_LOCATOR" in issue.reason
    assert "OCR" in issue.reason


def test_resolve_image_rejects_empty_resource_path():
    assets = {"x": Asset("x", locator_type=LOCATOR_IMAGE, resource_path="")}
    asset, _, issue = _resolve_image(_step("CLICK", target="x"), _ctx(assets))
    assert asset is None
    assert "MISSING_RESOURCE_PATH" in issue.reason


def test_resolve_image_rejects_none_string_resource_path():
    assets = {"x": Asset("x", locator_type=LOCATOR_IMAGE, resource_path="NONE")}
    asset, _, issue = _resolve_image(_step("CLICK", target="x"), _ctx(assets))
    assert asset is None
    assert "MISSING_RESOURCE_PATH" in issue.reason


def test_resolve_image_succeeds_on_valid_image_asset():
    assets = {"x": Asset("x", resource_path="a.png")}
    asset, lines, issue = _resolve_image(_step("CLICK", target="x"), _ctx(assets))
    assert asset is not None
    assert lines is None
    assert issue is None


# --------------------------------------------------------------------------- #
# CLICK / WAIT_FOR / ASSERT_VISIBLE                                           #
# --------------------------------------------------------------------------- #


def test_click_emits_touch_with_threshold(gen):
    assets = {"b": Asset("b", resource_path="b.png", threshold=0.9)}
    lines, issue = _HANDLERS["CLICK"](gen, _step("CLICK", target="b"), _ctx(assets))
    assert issue is None
    # CLICK is a deprecated alias — it emits a comment + the touch() call
    code_lines = [l for l in lines if not l.startswith("#")]
    assert code_lines == ["touch(Template('b.png', threshold=0.9))"]


def test_wait_for_uses_per_object_timeout(gen):
    assets = {"b": Asset("b", resource_path="b.png", threshold=0.8, timeout=12)}
    lines, issue = _HANDLERS["WAIT_FOR"](
        gen, _step("WAIT_FOR", target="b"), _ctx(assets)
    )
    assert issue is None
    assert lines[0].startswith("wait(Template(")
    assert "timeout=12" in lines[0]


def test_assert_visible_uses_per_object_timeout(gen):
    assets = {"b": Asset("b", resource_path="b.png", timeout=8)}
    lines, issue = _HANDLERS["ASSERT_VISIBLE"](
        gen, _step("ASSERT_VISIBLE", target="b"), _ctx(assets)
    )
    assert issue is None
    assert lines[0].startswith("assert_exists(Template(")
    assert "timeout=8" in lines[0]


# --------------------------------------------------------------------------- #
# START_APP                                                                   #
# --------------------------------------------------------------------------- #


def test_start_app_flags_missing_package(gen):
    lines, issue = _HANDLERS["START_APP"](gen, _step("START_APP"), _ctx())
    assert issue.reason == "START_APP_NEEDS_PACKAGE"
    assert any("START_APP_NEEDS_PACKAGE" in l for l in lines)


def test_start_app_emits_start_app_call_with_package(gen):
    lines, issue = _HANDLERS["START_APP"](
        gen, _step("START_APP"), _ctx(app_package="com.x.y")
    )
    assert issue is None
    assert "start_app('com.x.y')" in lines


def test_start_app_records_valid_json_config_in_comment(gen):
    lines, issue = _HANDLERS["START_APP"](
        gen,
        _step("START_APP", params='{"heart": "5", "coin": "10000"}'),
        _ctx(app_package="com.x"),
    )
    assert issue is None
    assert any(l.startswith("# config:") for l in lines)
    assert any("start_app(" in l for l in lines)


def test_start_app_flags_invalid_json_params(gen):
    lines, issue = _HANDLERS["START_APP"](
        gen,
        _step("START_APP", params="not-json"),
        _ctx(app_package="com.x"),
    )
    assert issue.reason == "INVALID_PARAMS_JSON"
    assert "INVALID_PARAMS_JSON" in lines[0]


# --------------------------------------------------------------------------- #
# INPUT_TEXT / READ_TEXT                                                      #
# --------------------------------------------------------------------------- #


def test_input_text_passes_params_through_repr(gen):
    lines, issue = _HANDLERS["INPUT_TEXT"](
        gen, _step("INPUT_TEXT", params="hello"), _ctx()
    )
    assert issue is None
    assert lines == ["text('hello')"]


def test_read_text_remains_a_todo_stub(gen):
    lines, issue = _HANDLERS["READ_TEXT"](gen, _step("READ_TEXT", target="t"), _ctx())
    assert issue is not None
    assert lines[0].startswith("# TODO:")


# --------------------------------------------------------------------------- #
# Generator dispatch                                                          #
# --------------------------------------------------------------------------- #


def test_generator_flags_unknown_action():
    from models import AirtestError

    with pytest.raises(AirtestError, match="UNSUPPORTED_ACTION"):
        generate_suite_script(
            [_step("BOGUS")],
            _ctx(),
            source_name="x.xlsx",
            suite_id="S",
        )


def test_generator_emits_step_label_with_expected_result():
    assets = {"b": Asset("b", resource_path="b.png")}
    src, _ = generate_suite_script(
        [_step("CLICK", target="b", expected="Tap play")],
        _ctx(assets),
        source_name="x.xlsx",
        suite_id="S",
    )
    assert "# step 1" in src
    assert "Tap play" in src


# --------------------------------------------------------------------------- #
# Injection guard via repr()                                                  #
# --------------------------------------------------------------------------- #


def test_repr_blocks_injection_via_resource_path(gen):
    nasty = 'x"); import os; os.system("calc"); Template(r"'
    assets = {"e": Asset("e", resource_path=nasty)}
    lines, issue = _HANDLERS["CLICK"](gen, _step("CLICK", target="e"), _ctx(assets))

    assert issue is None
    # CLICK now emits a deprecated comment first; find the touch() line.
    touch_line = next(l for l in lines if l.startswith("touch("))
    # Parse the generated line as an expression and assert the AST shape:
    # one outer touch() call, one inner Template() call, and the payload
    # appears only as a single string-literal argument.
    node = ast.parse(touch_line, mode="eval").body
    assert isinstance(node, ast.Call) and node.func.id == "touch"
    inner = node.args[0]
    assert isinstance(inner, ast.Call) and inner.func.id == "Template"
    assert isinstance(inner.args[0], ast.Constant)
    assert inner.args[0].value == nasty


def test_repr_blocks_injection_via_input_text_params(gen):
    nasty = '"); import os; os.system("calc"); ("'
    lines, _ = _HANDLERS["INPUT_TEXT"](gen, _step("INPUT_TEXT", params=nasty), _ctx())
    node = ast.parse(lines[0], mode="eval").body
    assert isinstance(node, ast.Call) and node.func.id == "text"
    assert isinstance(node.args[0], ast.Constant)
    assert node.args[0].value == nasty


@pytest.mark.parametrize("eol", ["\n", "\r", "\r\n"])
def test_step_label_strips_newlines_from_expected(eol):
    """Newline in expected must not escape the Python comment."""
    gen = AirtestGenerator()
    step = Step(suite_id="S1", step_no=1, action="SLEEP",
                expected=f'ok{eol}import os; os.system("evil")')
    label = gen.step_label(step)
    assert "\n" not in label
    assert "\r" not in label
    assert "import os" not in label
