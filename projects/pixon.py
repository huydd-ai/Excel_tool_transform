"""
projects/pixon.py — PixonGenerator targeting the pixon test framework.

Registered as project "pixon" via @register_project("pixon").
Run via:
    python excel_to_airtest.py AutomationRebase.xlsx --project pixon --report
"""

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_to_airtest import AirtestGenerator, action, register_project


@register_project("pixon")
class PixonGenerator(AirtestGenerator):
    """Generator targeting the pixon test framework used by AutoRebase / Screw Land."""

    DESCRIPTION = "Convert Excel to .air scripts."
    DEFAULT_APP_PACKAGE = "com.woodpuzzle.pin3d"

    IMPORTS = (
        "import traceback\n"
        "from airtest.core.api import *\n"
        "from airtest.aircv import Template\n"
        "from pixon.common import wrappers as wrapper\n"
        "from pixon.common.test_fixtures_daily import (\n"
        "    open_app_with_fake_ads, teardown_app, reset_progress, go_home_clean,\n"
        ")\n"
        "from pixon.pages.home_page     import HomePage\n"
        "from pixon.pages.game_page     import GamePage\n"
        "from pixon.pages.daily_mission import DailyMissionPage\n"
        "from pixon.pages.setting_page  import SettingPage\n"
        "from pixon.pages.cheat_page    import CheatPage"
    )

    MODULE_PROLOGUE = (
        "home_page = HomePage()\n"
        "game      = GamePage()\n"
        "daily     = DailyMissionPage()\n"
        "setting   = SettingPage()\n"
        "cheat     = CheatPage()"
    )

    def wrap_main_body(self, step_lines, suite_id):
        out = [
            "try:",
        ]
        if step_lines:
            for l in step_lines:
                out.append(f"    {l}")
        else:
            out.append("    pass")
        out += [
            "except Exception as e:",
            f"    wrapper.log_error({suite_id!r} + ' failed: ' + str(e) + chr(10) + traceback.format_exc())",
            f"    snapshot(filename={suite_id.lower() + '_error.png'!r})",
            "finally:",
            "    teardown_app()",
        ]
        return out

    @action("TAP")
    def handle_tap(self, step, ctx):
        asset, lines, issue = self._resolve_image(step, ctx)
        if asset is None:
            return lines, issue
        return [
            f"wrapper.touch(Template({asset.resource_path!r}, threshold={asset.threshold}))"
        ], None

    @action("TOUCH")
    def handle_touch(self, step, ctx):
        return self.handle_tap(step, ctx)

    @action("CLICK")
    def handle_click(self, step, ctx):
        asset, lines, issue = self._resolve_image(step, ctx)
        if asset is None:
            return lines, issue
        return [
            "# deprecated: use TAP or TOUCH",
            f"wrapper.touch(Template({asset.resource_path!r}, threshold={asset.threshold}))",
        ], None

    @action("WAIT_FOR")
    def handle_wait_for(self, step, ctx):
        asset, lines, issue = self._resolve_image(step, ctx)
        if asset is None:
            return lines, issue
        return [
            f"if not wrapper.wait_exists(Template({asset.resource_path!r}, "
            f"threshold={asset.threshold}), timeout={asset.timeout}):",
            f"    raise AssertionError('WAIT_FOR timed out for {step.target}')",
        ], None

    @action("ASSERT_VISIBLE")
    def handle_assert_visible(self, step, ctx):
        asset, lines, issue = self._resolve_image(step, ctx)
        if asset is None:
            return lines, issue
        return [
            f"if not wrapper.wait_exists(Template({asset.resource_path!r}, "
            f"threshold={asset.threshold}), timeout={asset.timeout}):",
            f"    raise AssertionError('ASSERT_VISIBLE failed for {step.target}')",
        ], None

    @action("START_APP")
    def handle_start_app(self, step, ctx):
        out = []
        if step.params:
            try:
                cfg = json.loads(step.params)
                out.append(f"# config (informational): {cfg!r}")
            except json.JSONDecodeError:
                return [f"# TODO: INVALID_PARAMS_JSON: {step.params!r}"], self._issue(
                    step, "INVALID_PARAMS_JSON"
                )
        out.append("open_app_with_fake_ads(home_page)")
        return out, None


if __name__ == "__main__":
    PixonGenerator.main()
