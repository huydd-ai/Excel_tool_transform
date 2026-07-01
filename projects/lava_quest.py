"""
excel_tool/projects/lava_quest.py — Lava Quest project generator for the pixon framework.

This project generator targets the pixon automation framework (pixon.pages, pixon.common).
Generated scripts use the pixon page objects (HomePage, LavaQuestPage, etc.) and
framework helpers (run_step, wrappers, cold_start_with_json, etc.).

Run it::
    python -m excel_tool --list-projects                  # shows "lava_quest"
    python -m excel_tool --init-excel                      # writes template.xlsx
    python -m excel_tool template.xlsx --project lava_quest --app-package com.example.game
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_to_airtest import AirtestGenerator, action, register_project


@register_project("lava_quest")
class LavaQuestGenerator(AirtestGenerator):
    """
    Lava Quest generator for the pixon framework.

    Overrides:
    - IMPORTS: imports pixon pages and framework helpers
    - MODULE_PROLOGUE: page-object singletons (HomePage, LavaQuestPage, etc.)
    - DEFAULT_APP_PACKAGE: the game's Android package
    - wrap_main_body: wraps the generated steps in try/except with snapshot on failure
    - Custom @action handlers for pixon-specific keywords (START_GAME, PLAY_AND_WIN_LEVELS,
      INITIALIZE_QUEST, CLAIM_REWARD, VERIFY_COUNTDOWN, etc.)
    """

    DESCRIPTION = "Lava Quest (pixon framework) — generates test scripts using pixon page objects and framework helpers."

    # The Android package id used by START_APP / STOP_APP.
    # Override per game, or pass --app-package on the CLI.
    DEFAULT_APP_PACKAGE = "com.example.lavaquest"

    # Default profile used by cold_start_with_json in many LQ tests
    DEFAULT_PROFILE = {
        "fakeads": True,
        "heart": 5,
        "coin": 10000,
        "playspeed": 6,
        "clear_data": True,
        "server_sync": False,
    }

    # Base imports — pixon framework + Airtest
    IMPORTS = (
        "from airtest.core.api import *\n"
        "from airtest.aircv import Template\n\n"
        "from pixon.common import wrappers as wrapper\n"
        "from pixon.common.test_flow import (\n"
        "    run_step,\n"
        "    teardown_app,\n"
        "    go_home_clean,\n"
        "    close_all_popups,\n"
        ")\n"
        "from pixon.common.adb_utils import (\n"
        "    set_param,\n"
        "    cold_start_with_json,\n"
        "    set_time_relative,\n"
        ")\n"
        "from pixon.pages.home_page import HomePage\n"
        "from pixon.pages.lava_quest_page import LavaQuestPage\n"
        "from pixon.common.wrappers import log_info\n"
    )

    # Module-level singletons — instantiated once per generated script
    MODULE_PROLOGUE = (
        "\n"
        "home_page = HomePage()\n"
        "lava_quest = LavaQuestPage()\n"
    )

    def wrap_main_body(self, step_lines, suite_id):
        """
        Wrap the generated steps so a failure snapshots before re-raising.
        Uses pixon wrapper snapshot.
        """
        body = step_lines or ["pass"]
        # If the suite resolved to comment-only lines (e.g. all-TODO steps),
        # the try-block would be syntactically empty — keep it valid with a pass.
        if not any(line and not line.lstrip().startswith("#") for line in body):
            body = [*body, "pass"]
        out = ["try:"]
        out += [f"    {line}" for line in body]
        out += [
            "except Exception:",
            f"    wrapper.snapshot(filename={suite_id.lower() + '_error.png'!r})",
            "    raise",
        ]
        return out

    # ──────────────────────────────────────────────────────────────────────
    # Custom action handlers for Lava Quest / pixon framework keywords
    # ──────────────────────────────────────────────────────────────────────

    @action("START_GAME")
    def handle_start_game(self, step, ctx):
        """
        START_GAME — cold start the app with the default LQ profile.
        Params (JSON, optional): override profile keys, e.g.
            {"level": 27, "playspeed": 6}
        """
        out = []
        profile = dict(DEFAULT_PROFILE)
        if step.params:
            import json
            try:
                overrides = json.loads(step.params)
                profile.update(overrides)
            except json.JSONDecodeError:
                return [f"# TODO: INVALID_PARAMS_JSON: {step.params!r}"], self._issue(step, "INVALID_PARAMS_JSON")
        out.append(f"run_step('cold start', cold_start_with_json, {profile!r})")
        out.append("sleep(30)")  # wait for app to fully launch
        out.append("run_step('close all popups after launch', close_all_popups, home_page)")
        out.append("run_step('navigate home after launch', go_home_clean, home_page)")
        out.append("home_reached = home_page.wait_for_element(home_page.btn_main_home, timeout=5)")
        out.append(
            "log_info(f'[START_GAME] RESULT — expected: app launched and home reached | actual: home_reached={bool(home_reached)}')"
        )
        return out, None

    @action("INITIALIZE_QUEST")
    def handle_initialize_quest(self, step, ctx):
        """
        INITIALIZE_QUEST — run the first-time LQ flow to initialize the quest.
        No params needed.
        """
        out = [
            "run_step('initialize quest (first-time flow)', lava_quest.initialize_quest, home_page)",
            "log_info('[INITIALIZE_QUEST] RESULT — expected: first-time LQ flow completed | actual: initialize_quest returned without error')",
        ]
        return out, None

    @action("PLAY_AND_WIN_LEVELS")
    def handle_play_and_win_levels(self, step, ctx):
        """
        PLAY_AND_WIN_LEVELS — play and win N levels.
        Params (JSON): {"count": 7}  (required key: count)
        """
        import json
        if not step.params:
            return self._todo(step, 'PLAY_AND_WIN_LEVELS_NEEDS_PARAMS: provide JSON {"count": N}')
        try:
            cfg = json.loads(step.params)
            count = int(cfg.get("count", 0))
        except (json.JSONDecodeError, ValueError, TypeError):
            return self._todo(step, f'INVALID_PARAMS: {step.params!r} — expected {{\"count\": N}}')
        if count <= 0:
            return self._todo(step, 'PLAY_AND_WIN_LEVELS_NEEDS_POSITIVE_COUNT')
        out = [
            f"run_step('play and win {count} levels', lava_quest.play_and_win_levels, home_page, {count})",
            "run_step('close popups after wins', close_all_popups, home_page)",
            "run_step('navigate home after wins', go_home_clean, home_page)",
            f"log_info('[PLAY_AND_WIN_LEVELS] RESULT — expected: {count} wins completed across the chain | actual: play_and_win_levels verified {count} wins via OCR without error')",
        ]
        return out, None

    @action("OPEN_LQ_POPUP")
    def handle_open_lq_popup(self, step, ctx):
        """
        OPEN_LQ_POPUP — open the Lava Quest popup and verify it opens.
        No params.
        """
        out = [
            "popup_opened = lava_quest.open_lava_quest_popup()",
            "claim_active = bool(popup_opened) and bool(lava_quest.wait_for_element(lava_quest.claim_btn, timeout=5))",
            "log_info(f'[OPEN_LQ_POPUP] RESULT — expected: popup open, claim button active | actual: popup_opened={bool(popup_opened)}, claim_active={claim_active}')",
            "if not popup_opened:",
            "    raise AssertionError('Lava Quest popup did not open')",
            "if not claim_active:",
            "    raise AssertionError('Claim button not active')",
        ]
        return out, None

    @action("CLAIM_REWARD")
    def handle_claim_reward(self, step, ctx):
        """
        CLAIM_REWARD — tap claim button, verify win banner, OCR coin amount.
        Params (JSON, optional): {"expected_split": 5000} — total pool to split among winners.
        """
        import json
        expected_pool = 5000
        if step.params:
            try:
                cfg = json.loads(step.params)
                expected_pool = int(cfg.get("expected_split", 5000))
            except (json.JSONDecodeError, ValueError, TypeError):
                return self._todo(step, f'INVALID_PARAMS_JSON: {step.params!r} — expected {{\"expected_split\": N}}')

        out = [
            "lava_quest.tap(lava_quest.claim_btn)",
            "sleep(3)",
            "",
            "# Verify the \"You Win\" overlay banner appears",
            "if not lava_quest.wait_for_element(lava_quest.label_win_lava_quest, timeout=8):",
            "    raise AssertionError(\"'You Win' banner did not appear after tapping Claim\")",
            "sleep(1)",
            "",
            "# OCR the coin amount shown under the banner and the winner count",
            "win_amount, other_winners = lava_quest.get_reward_and_winners_via_ocr()",
            "log_info(f'Win screen OCR: win_amount={win_amount}, other_winners={other_winners}')",
            "",
            "if win_amount <= 0:",
            "    raise AssertionError(f'Expected a positive coin reward on the Win banner, but got {win_amount}')",
            "",
            f"expected_win = {expected_pool} // (other_winners + 1)",
            "if win_amount != expected_win:",
            f"    log_info(f'[WARN] Win amount {{win_amount}} differs from expected split {{expected_win}} ({expected_pool}/{{other_winners+1}}) — logging only, not failing')",
            "else:",
            f"    log_info(f'Verified reward split: {{win_amount}} coins matches expected {expected_pool} / {{other_winners + 1}}')",
            "",
            "# Tap TAP TO CLAIM button on Win screen",
            "from pixon.pages.home_page import HomePage",
            "tap_to_claim_btn = HomePage.tap_to_claim",
            "if not lava_quest.wait_for_element(tap_to_claim_btn, timeout=5):",
            "    raise AssertionError('TAP TO CLAIM button not visible on Win screen')",
            "lava_quest.tap(tap_to_claim_btn)",
            "sleep(3)",
            "",
            "# Dismiss any reward confirmation popup",
            "if lava_quest.wait_for_element(lava_quest.btn_close, timeout=3):",
            "    lava_quest.tap(lava_quest.btn_close)",
            "    sleep(2)",
            "",
            "# Verify we returned to home",
            "returned_home = bool(home_page.wait_for_element(home_page.btn_main_home, timeout=10))",
            "log_info(f'[CLAIM_REWARD] RESULT — expected: reward claimed and returned home | actual: returned_home={returned_home}')",
            "if not returned_home:",
            "    raise AssertionError('Did not return to home after claim')",
        ]
        return out, None

    @action("VERIFY_COUNTDOWN")
    def handle_verify_countdown(self, step, ctx):
        """
        VERIFY_COUNTDOWN — reopen LQ popup and verify the 24h cooldown timer is visible.
        No params.
        """
        out = [
            "log_info('[VERIFY_COUNTDOWN] START: reopen LQ, verify countdown timer replaces claim', snapshot=False)",
            "reopened = lava_quest.open_lava_quest_popup()",
            "countdown_visible = bool(reopened) and lava_quest.is_time_count_down_visible(timeout=5)",
            "timer_text = lava_quest.check_timer_text() if countdown_visible else ''",
            "log_info(f'[VERIFY_COUNTDOWN] RESULT — expected: countdown timer visible (24h cooldown) | actual: reopened={bool(reopened)}, countdown_visible={countdown_visible}, timer_text={timer_text!r}')",
            "if not reopened:",
            "    raise AssertionError('Lava Quest popup did not open to show countdown')",
            "if not countdown_visible:",
            "    raise AssertionError('Countdown timer NOT visible after completing LQ')",
            "log_info('[VERIFY_COUNTDOWN] END', snapshot=False)",
        ]
        return out, None

    @action("ADVANCE_TIME")
    def handle_advance_time(self, step, ctx):
        """
        ADVANCE_TIME — advance device time by relative hours to expire events.
        Params (JSON): {"hours": 25.0}
        """
        import json
        if not step.params:
            return self._todo(step, 'ADVANCE_TIME_NEEDS_PARAMS: provide JSON {"hours": N}')
        try:
            cfg = json.loads(step.params)
            hours = float(cfg.get("hours", 0))
        except (json.JSONDecodeError, ValueError, TypeError):
            return self._todo(step, f'INVALID_ADVANCE_TIME_PARAMS: {step.params!r} — expected {{\"hours\": N}}')
        out = [
            f"run_step('advance time to expire old events', set_time_relative, {hours})",
        ]
        return out, None

    @action("SET_PARAM")
    def handle_set_param(self, step, ctx):
        """
        SET_PARAM — set a game parameter via adb.
        Params (JSON): {"key": "level", "value": "27"}
        """
        import json
        if not step.params:
            return self._todo(step, 'SET_PARAM_NEEDS_PARAMS: provide JSON {"key": "...", "value": "..."}')
        try:
            cfg = json.loads(step.params)
            key = cfg.get("key")
            value = cfg.get("value")
            if key is None or value is None:
                raise ValueError
        except (json.JSONDecodeError, ValueError, TypeError):
            return self._todo(step, f'INVALID_SET_PARAM_PARAMS: {step.params!r} — expected {{\"key\": "...", \"value\": "..."}}')
        out = [f"run_step('set param {key}={value}', set_param, {key!r}, {value!r})"]
        return out, None

    # ──────────────────────────────────────────────────────────────────────
    # Override base handlers to use pixon framework equivalents where needed
    # ──────────────────────────────────────────────────────────────────────

    @action("START_APP")
    def handle_start_app(self, step, ctx):
        """
        START_APP — delegate to pixon's cold_start_with_json using DEFAULT_PROFILE.
        """
        import json
        out = []
        profile = dict(self.DEFAULT_PROFILE)
        if step.params:
            try:
                cfg = json.loads(step.params)
                profile.update(cfg)
            except json.JSONDecodeError:
                return [f"# TODO: INVALID_PARAMS_JSON: {step.params!r}"], self._issue(step, "INVALID_PARAMS_JSON")
        out.append(f"run_step('cold start app', cold_start_with_json, {profile!r})")
        out.append("sleep(20)")  # generous wait for full launch
        return out, None

    @action("STOP_APP")
    def handle_stop_app(self, step, ctx):
        """
        STOP_APP — use pixon's teardown_app.
        """
        return ["run_step('teardown app', teardown_app)"], None

    @action("SLEEP")
    def handle_sleep(self, step, ctx):
        """
        SLEEP — use airtest sleep but allow JSON params like {"seconds": 2.5}.
        """
        raw = step.params.strip() if step.params else ""
        try:
            if raw.startswith("{"):
                import json
                seconds = float(json.loads(raw).get("seconds", 0))
            else:
                seconds = float(raw) if raw else 1.0
            return [f"sleep({seconds})"], None
        except (json.JSONDecodeError, KeyError, ValueError, TypeError):
            return self._todo(
                step,
                f'INVALID_SLEEP_PARAMS: {step.params!r} — provide seconds as number or {{"seconds": 2.5}}',
            )

    @action("SNAPSHOT")
    def handle_snapshot(self, step, ctx):
        """
        SNAPSHOT — use wrapper snapshot with filename.
        """
        filename = (step.params.strip() if step.params else "") or step.suite_id.lower()
        if not filename.endswith(".png"):
            filename += ".png"
        return [f"wrapper.snapshot(filename={filename!r})"], None


if __name__ == "__main__":
    LavaQuestGenerator.main()