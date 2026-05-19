"""TAP and TOUCH produce identical touch() code. CLICK is a deprecated alias."""
import pytest
from excel_to_airtest import AirtestGenerator, Asset, GenCtx, Step, LOCATOR_IMAGE


@pytest.fixture
def gen():
    return AirtestGenerator()


def _step(action, target="btn", suite="S", step_no=1, row=2):
    return Step(suite_id=suite, step_no=step_no, action=action,
                excel_row=row, target=target, params="", expected="")


def _ctx():
    return GenCtx(
        assets={"btn": Asset("btn", LOCATOR_IMAGE, "./img/btn.png", 0.85, 5.0)},
        app_package="com.test",
    )


def test_tap_generates_touch():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["TAP"](gen, _step("TAP"), _ctx())
    assert issue is None
    assert any("touch(" in l for l in lines)
    assert any("btn.png" in l for l in lines)


def test_touch_generates_touch():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["TOUCH"](gen, _step("TOUCH"), _ctx())
    assert issue is None
    assert any("touch(" in l for l in lines)
    assert any("btn.png" in l for l in lines)


def test_tap_and_touch_generate_same_code():
    gen = AirtestGenerator()
    tap_lines, _ = gen._HANDLERS["TAP"](gen, _step("TAP"), _ctx())
    touch_lines, _ = gen._HANDLERS["TOUCH"](gen, _step("TOUCH"), _ctx())
    code_tap = [l for l in tap_lines if not l.startswith("#")]
    code_touch = [l for l in touch_lines if not l.startswith("#")]
    assert code_tap == code_touch


def test_click_still_works_as_deprecated_alias():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["CLICK"](gen, _step("CLICK"), _ctx())
    assert issue is None
    assert any("touch(" in l for l in lines)


def test_click_output_contains_deprecated_comment():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["CLICK"](gen, _step("CLICK"), _ctx())
    assert any("deprecated" in l.lower() for l in lines)


def test_tap_missing_target_returns_todo():
    gen = AirtestGenerator()
    step = Step(suite_id="S", step_no=1, action="TAP", excel_row=2,
                target="", params="", expected="")
    lines, issue = gen._HANDLERS["TAP"](gen, step, _ctx())
    assert issue is not None
    assert "MISSING_TARGET" in issue.reason


def test_touch_missing_target_returns_todo():
    gen = AirtestGenerator()
    step = Step(suite_id="S", step_no=1, action="TOUCH", excel_row=2,
                target="", params="", expected="")
    lines, issue = gen._HANDLERS["TOUCH"](gen, step, _ctx())
    assert issue is not None
    assert "MISSING_TARGET" in issue.reason
