"""New action handlers: SWIPE, SCROLL, LONG_PRESS, SLEEP, BACK, HOME, SNAPSHOT, STOP_APP."""

import pytest
from excel_to_airtest import AirtestGenerator
from models import Asset, GenCtx, Step, LOCATOR_IMAGE


def _step(action, target="", params="", suite="S", step_no=1, row=2):
    return Step(
        suite_id=suite,
        step_no=step_no,
        action=action,
        excel_row=row,
        target=target,
        params=params,
        expected="",
    )


def _ctx(pkg="com.test"):
    return GenCtx(
        assets={"btn": Asset("btn", LOCATOR_IMAGE, "./img/btn.png", 0.85, 3.0)},
        app_package=pkg,
    )


# --- SWIPE ---


def test_swipe_generates_swipe_call():
    gen = AirtestGenerator()
    params = '{"from": [100, 200], "to": [300, 400], "duration": 0.5}'
    lines, issue = gen._HANDLERS["SWIPE"](gen, _step("SWIPE", params=params), _ctx())
    assert issue is None
    assert len(lines) == 1
    assert lines[0] == "swipe((100, 200), (300, 400), duration=0.5)"


def test_swipe_default_duration():
    gen = AirtestGenerator()
    params = '{"from": [10, 20], "to": [30, 40]}'
    lines, issue = gen._HANDLERS["SWIPE"](gen, _step("SWIPE", params=params), _ctx())
    assert issue is None
    assert "duration=0.5" in lines[0]


def test_swipe_missing_params_returns_todo():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SWIPE"](gen, _step("SWIPE", params=""), _ctx())
    assert issue is not None
    assert "SWIPE_NEEDS_PARAMS" in issue.reason


def test_swipe_invalid_json_returns_todo():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SWIPE"](
        gen, _step("SWIPE", params="not json"), _ctx()
    )
    assert issue is not None
    assert "INVALID_SWIPE_PARAMS" in issue.reason


# --- SCROLL ---


def test_scroll_up_generates_two_lines():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SCROLL"](
        gen, _step("SCROLL", params='{"direction":"up"}'), _ctx()
    )
    assert issue is None
    assert len(lines) == 2
    assert "get_current_resolution" in lines[0]
    assert "swipe" in lines[1]
    assert "h*0.7" in lines[1]
    assert "h*0.3" in lines[1]


def test_scroll_accepts_plain_string_direction():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SCROLL"](gen, _step("SCROLL", params="down"), _ctx())
    assert issue is None
    assert "h*0.3" in lines[1]
    assert "h*0.7" in lines[1]


def test_scroll_left():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SCROLL"](gen, _step("SCROLL", params="left"), _ctx())
    assert issue is None
    assert "w*0.7" in lines[1]
    assert "w*0.3" in lines[1]


def test_scroll_right():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SCROLL"](gen, _step("SCROLL", params="right"), _ctx())
    assert issue is None
    assert "w*0.3" in lines[1]
    assert "w*0.7" in lines[1]


def test_scroll_invalid_direction_returns_todo():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SCROLL"](
        gen, _step("SCROLL", params="diagonal"), _ctx()
    )
    assert issue is not None
    assert "INVALID_SCROLL_DIRECTION" in issue.reason


# --- LONG_PRESS ---


def test_long_press_generates_long_click():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["LONG_PRESS"](
        gen, _step("LONG_PRESS", target="btn"), _ctx()
    )
    assert issue is None
    assert any("long_click(" in l for l in lines)
    assert any("btn.png" in l for l in lines)


def test_long_press_missing_target_returns_todo():
    from models import AirtestError

    gen = AirtestGenerator()
    with pytest.raises(AirtestError, match="MISSING_TARGET"):
        gen._HANDLERS["LONG_PRESS"](gen, _step("LONG_PRESS", target=""), _ctx())


# --- SLEEP ---


def test_sleep_plain_number():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SLEEP"](gen, _step("SLEEP", params="2.5"), _ctx())
    assert issue is None
    assert lines == ["sleep(2.5)"]


def test_sleep_json_seconds():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SLEEP"](
        gen, _step("SLEEP", params='{"seconds": 3}'), _ctx()
    )
    assert issue is None
    assert lines == ["sleep(3.0)"]


def test_sleep_integer_string():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SLEEP"](gen, _step("SLEEP", params="2"), _ctx())
    assert issue is None
    assert lines == ["sleep(2.0)"]


def test_sleep_bad_params_returns_todo():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SLEEP"](gen, _step("SLEEP", params="abc"), _ctx())
    assert issue is not None
    assert "INVALID_SLEEP_PARAMS" in issue.reason


# --- BACK ---


def test_back_generates_keyevent():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["BACK"](gen, _step("BACK"), _ctx())
    assert issue is None
    assert lines == ['keyevent("BACK")']


# --- HOME ---


def test_home_generates_keyevent():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["HOME"](gen, _step("HOME"), _ctx())
    assert issue is None
    assert lines == ['keyevent("HOME")']


# --- SNAPSHOT ---


def test_snapshot_with_params():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SNAPSHOT"](
        gen, _step("SNAPSHOT", params="login_ok"), _ctx()
    )
    assert issue is None
    assert lines == ["snapshot(filename='login_ok.png')"]


def test_snapshot_params_already_has_extension():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SNAPSHOT"](
        gen, _step("SNAPSHOT", params="login.png"), _ctx()
    )
    assert issue is None
    assert lines == ["snapshot(filename='login.png')"]


def test_snapshot_no_params_uses_suite_id():
    gen = AirtestGenerator()
    step = Step(
        suite_id="TC_MISSION_1",
        step_no=1,
        action="SNAPSHOT",
        excel_row=2,
        target="",
        params="",
        expected="",
    )
    lines, issue = gen._HANDLERS["SNAPSHOT"](gen, step, _ctx())
    assert issue is None
    assert "tc_mission_1" in lines[0]


# --- STOP_APP ---


def test_stop_app_generates_stop_app_call():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["STOP_APP"](
        gen, _step("STOP_APP"), _ctx(pkg="com.test")
    )
    assert issue is None
    assert lines == ['stop_app("com.test")']


def test_stop_app_no_package_returns_todo():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["STOP_APP"](gen, _step("STOP_APP"), _ctx(pkg=""))
    assert issue is not None
    assert "STOP_APP_NEEDS_PACKAGE" in issue.reason


def test_scroll_empty_params_defaults_to_up():
    gen = AirtestGenerator()
    lines, issue = gen._HANDLERS["SCROLL"](gen, _step("SCROLL", params=""), _ctx())
    assert issue is None
    assert "h*0.7" in lines[1]
    assert "h*0.3" in lines[1]


def test_snapshot_injection_safe():
    gen = AirtestGenerator()
    payload = 'evil"); import os; os.system("rm -rf /'
    lines, issue = gen._HANDLERS["SNAPSHOT"](
        gen, _step("SNAPSHOT", params=payload), _ctx()
    )
    assert issue is None
    assert lines[0].startswith("snapshot(filename=")
    import ast

    tree = ast.parse(lines[0], mode="eval")
    call = tree.body
    assert isinstance(call, ast.Call)
    kw = call.keywords[0]
    assert isinstance(kw.value, ast.Constant)
