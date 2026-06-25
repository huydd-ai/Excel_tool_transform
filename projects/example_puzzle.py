"""
projects/example_puzzle.py — a generic, game-agnostic example project.

This is the reference plugin shipped with the base tool. Copy it (or run
``python -m excel_tool --init-project <name>``) to target your own puzzle game.

It targets **bare Airtest** only — generated scripts import nothing beyond the
Airtest API, so they run anywhere Airtest is installed, with no game framework.

Run it::

    python -m excel_tool --list-projects                  # shows "example_puzzle"
    python -m excel_tool --init-excel                      # writes a sample template.xlsx
    python -m excel_tool template.xlsx --project example_puzzle

To retarget for a real game, override the class attributes (app package,
IMPORTS, MODULE_PROLOGUE) and any @action handlers whose emitted code should
route through your own test framework instead of bare Airtest.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_to_airtest import AirtestGenerator, action, register_project


@register_project("example_puzzle")
class ExamplePuzzleGenerator(AirtestGenerator):
    """Minimal generic puzzle-game project. Inherits every base (bare-Airtest)
    action handler; overrides only the app package and the body wrapper."""

    DESCRIPTION = "Example generic puzzle-game project (bare Airtest)."

    # The Android package id used by START_APP / STOP_APP. Override per game,
    # or pass --app-package on the CLI.
    DEFAULT_APP_PACKAGE = "com.example.puzzle"

    # Base IMPORTS already pull in the Airtest API; nothing game-specific here.
    # IMPORTS = "from airtest.core.api import *\nfrom airtest.aircv import Template"
    # MODULE_PROLOGUE = ""   # page-object singletons go here, if your game has any.

    def wrap_main_body(self, step_lines, suite_id):
        """Wrap the generated steps so a failure snapshots before re-raising.

        Demonstrates the wrap_main_body hook. Uses only Airtest built-ins
        (snapshot), so the output stays framework-free.
        """
        body = step_lines or ["pass"]
        out = ["try:"]
        out += [f"    {line}" for line in body]
        out += [
            "except Exception:",
            f"    snapshot(filename={suite_id.lower() + '_error.png'!r})",
            "    raise",
        ]
        return out

    # Example of overriding a single action. Uncomment and edit to route a
    # keyword through your own framework instead of bare Airtest:
    #
    # @action("TAP")
    # def handle_tap(self, step, ctx):
    #     asset, lines, issue = self._resolve_image(step, ctx)
    #     if asset is None:
    #         return lines, issue
    #     return [f"my_framework.tap({asset.resource_path!r})"], None


if __name__ == "__main__":
    ExamplePuzzleGenerator.main()
