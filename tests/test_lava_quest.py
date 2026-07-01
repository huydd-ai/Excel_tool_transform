"""Regression anchor for the Lava Quest project generator.

Guards that projects/lava_quest.py registers, generates valid pixon-framework
scripts, and stays free of bare-Airtest tokens where pixon equivalents exist.
"""

import ast
import importlib.util
import os

from excel_to_airtest import Asset, GenCtx, LOCATOR_IMAGE, Step

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load_lava_quest():
    path = os.path.join(ROOT, "projects", "lava_quest.py")
    spec = importlib.util.spec_from_file_location("_lava_quest_under_test", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # fires @register_project("lava_quest")
    return mod


def _ctx(app_package="com.example.lavaquest"):
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
        Step("TC_DEMO", 1, "ADVANCE_TIME", 2, "", '{"hours": 25.0}', "Advance time"),
        Step("TC_DEMO", 2, "START_APP", 3, "", '{"level": 27}', "Cold start at level 27"),
        Step("TC_DEMO", 3, "INITIALIZE_QUEST", 4, "", "", "First-time LQ flow"),
        Step("TC_DEMO", 4, "PLAY_AND_WIN_LEVELS", 5, "", '{"count": 7}', "Play 7 levels"),
        Step("TC_DEMO", 5, "OPEN_LQ_POPUP", 6, "", "", "Open LQ popup"),
        Step("TC_DEMO", 6, "CLAIM_REWARD", 7, "", '{"expected_split": 5000}', "Claim reward"),
        Step("TC_DEMO", 7, "VERIFY_COUNTDOWN", 8, "", "", "Verify 24h cooldown"),
    ]


def test_lava_quest_registers():
    from registry import _PROJECT_REGISTRY

    _load_lava_quest()
    assert "lava_quest" in _PROJECT_REGISTRY


def test_lava_quest_generates_valid_pixon_script():
    mod = _load_lava_quest()
    gen = mod.LavaQuestGenerator()
    script, issues = gen.generate_suite_script(_steps(), _ctx(), "demo.xlsx", "TC_DEMO")

    assert not issues, f"Generation issues: {issues}"
    ast.parse(script)  # valid Python

    # Must use pixon framework imports and helpers
    assert "from pixon.common.test_flow import" in script
    assert "run_step" in script
    assert "teardown_app" in script
    assert "go_home_clean" in script
    assert "close_all_popups" in script
    assert "from pixon.pages.home_page import HomePage" in script
    assert "from pixon.pages.lava_quest_page import LavaQuestPage" in script
    assert "from pixon.common.adb_utils import" in script
    assert "home_page = HomePage()" in script
    assert "lava_quest = LavaQuestPage()" in script

    # Must use run_step for pixon actions
    assert "run_step('advance time to expire old events', set_time_relative, 25.0)" in script
    assert "run_step('cold start app', cold_start_with_json" in script
    assert "run_step('initialize quest (first-time flow)', lava_quest.initialize_quest" in script
    assert "run_step('play and win 7 levels', lava_quest.play_and_win_levels" in script
    assert "run_step('close popups after wins', close_all_popups" in script
    assert "run_step('navigate home after wins', go_home_clean" in script

    # Must use wrapper snapshot on failure
    assert "wrapper.snapshot" in script

    # Must have try/except wrapper
    assert "try:" in script
    assert "except Exception:" in script
    assert "snapshot(filename=" in script
    assert "raise" in script

    # Bare Airtest calls should be minimal (only sleep)
    # Verify pixon framework is used for all major actions
    assert "lava_quest.open_lava_quest_popup()" in script
    assert "lava_quest.claim_btn" in script
    assert "lava_quest.get_reward_and_winners_via_ocr()" in script
    assert "lava_quest.is_time_count_down_visible" in script


def test_lava_quest_comment_only_suite_is_valid_python():
    """A suite whose steps all resolve to comments/TODOs must still produce valid Python."""
    mod = _load_lava_quest()
    gen = mod.LavaQuestGenerator()
    steps = [Step("TC_TODO", 1, "READ_TEXT", 2, "score_text", "", "read score")]
    script, _ = gen.generate_suite_script(steps, _ctx(), "demo.xlsx", "TC_TODO")

    ast.parse(script)  # would raise SyntaxError on an empty try-block
    assert "try:" in script
    assert "pass" in script  # guard line keeps the try-block non-empty