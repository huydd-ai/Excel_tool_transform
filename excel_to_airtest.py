"""
excel_tool/excel_to_airtest.py — Main AirtestGenerator Engine Facade.

This file provides the `AirtestGenerator` class which translates
parsed Step objects into Airtest Python scripts. It also re-exports
essential types so user subclasses don't break.
"""

from __future__ import annotations

import json
import sys

# ── Backward compatibility exports for existing projects ──
from models import (
    Asset,
    Step,
    GenCtx,
    GenerationIssue,
    HandlerResult,
    LOCATOR_IMAGE,
    LOCATOR_OCR,
    AirtestError,
    _SCROLL_PRESETS,
)
from registry import action, register_project, _GenMeta, _PROJECT_REGISTRY
from hints import diagnostic_hint

_diagnostic_hint = diagnostic_hint


class AirtestGenerator(metaclass=_GenMeta):
    """Father tool. Subclass and override hooks to retarget for any project."""

    # === Sheet names (subclass may rebind) ============================== #
    OBJECTS_SHEET = "Object_Repository"
    ACTIONS_SHEET = "Action_Logic"
    STEPS_SHEET = "Test_Execution"

    # === Generated-script template hooks ================================ #
    IMPORTS = "from airtest.core.api import *\nfrom airtest.aircv import Template"
    MODULE_PROLOGUE = ""

    # === CLI defaults (subclass may rebind) ============================= #
    DESCRIPTION = (
        "Convert AutomationRebase Excel to Airtest .air scripts (one per Suite_ID)."
    )
    DEFAULT_APP_PACKAGE = ""
    DEFAULT_PLAN = None

    # ------------------------------------------------------------------- #
    # Asset resolution helpers
    # ------------------------------------------------------------------- #
    def _issue(self, step: Step, reason: str) -> GenerationIssue:
        return GenerationIssue(step.suite_id, step.step_no, step.excel_row, reason)

    def _todo(self, step: Step, reason: str):
        return [f"# TODO: {reason}"], self._issue(step, reason)

    def inject_config(self, cfg: dict) -> list[str]:
        """Convert a configuration dict into an injected code string."""
        return [f"# config: {cfg}"]

    def _resolve_image(self, step: Step, ctx: GenCtx):
        """Return (Asset, None, None) on success or raise AirtestError on missing/unknown targets to prevent hallucination."""
        if not step.target:
            raise AirtestError(
                f"MISSING_TARGET: target is missing for step {step.step_no} in suite {step.suite_id}"
            )
        asset = ctx.assets.get(step.target)
        if asset is None:
            raise AirtestError(
                f"UNKNOWN_TARGET '{step.target}' in suite {step.suite_id}"
            )
        if asset.locator_type != LOCATOR_IMAGE:
            l, i = self._todo(
                step, f"UNSUPPORTED_LOCATOR '{asset.locator_type}' for '{step.target}'"
            )
            return None, l, i
        if not asset.resource_path or asset.resource_path.upper() == "NONE":
            l, i = self._todo(step, f"MISSING_RESOURCE_PATH for '{step.target}'")
            return None, l, i
        return asset, None, None

    # ------------------------------------------------------------------- #
    # Default action handlers (bare Airtest)
    # ------------------------------------------------------------------- #
    @action("TAP")
    def handle_tap(self, step, ctx):
        asset, lines, issue = self._resolve_image(step, ctx)
        if asset is None:
            return lines, issue
        return [
            f"touch(Template({asset.resource_path!r}, threshold={asset.threshold}))"
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
            f"touch(Template({asset.resource_path!r}, threshold={asset.threshold}))",
        ], None

    @action("WAIT_FOR")
    def handle_wait_for(self, step, ctx):
        asset, lines, issue = self._resolve_image(step, ctx)
        if asset is None:
            return lines, issue
        return [
            f"wait(Template({asset.resource_path!r}, threshold={asset.threshold}), timeout={asset.timeout})"
        ], None

    @action("ASSERT_VISIBLE")
    def handle_assert_visible(self, step, ctx):
        asset, lines, issue = self._resolve_image(step, ctx)
        if asset is None:
            return lines, issue
        return [
            f"assert_exists(Template({asset.resource_path!r}, threshold={asset.threshold}), timeout={asset.timeout})"
        ], None

    @action("START_APP")
    def handle_start_app(self, step, ctx):
        out = []
        if step.params:
            try:
                cfg = json.loads(step.params)
                out.extend(self.inject_config(cfg))
            except json.JSONDecodeError:
                return [f"# TODO: INVALID_PARAMS_JSON: {step.params!r}"], self._issue(
                    step, "INVALID_PARAMS_JSON"
                )
        if not ctx.app_package:
            out.append("# TODO: START_APP_NEEDS_PACKAGE - pass via --app-package")
            return out, self._issue(step, "START_APP_NEEDS_PACKAGE")
        out.append(f"start_app({ctx.app_package!r})")
        return out, None

    @action("INPUT_TEXT")
    def handle_input_text(self, step, ctx):
        return [f"text({step.params!r})"], None

    @action("READ_TEXT")
    def handle_read_text(self, step, ctx):
        return self._todo(step, f"READ_TEXT not implemented (target='{step.target}')")

    @action("SWIPE")
    def handle_swipe(self, step, ctx):
        if not step.params:
            return self._todo(
                step,
                'SWIPE_NEEDS_PARAMS: provide JSON {"from":[x1,y1],"to":[x2,y2],"duration":0.5}',
            )
        try:
            cfg = json.loads(step.params)
            x1, y1 = cfg["from"]
            x2, y2 = cfg["to"]
            duration = cfg.get("duration", 0.5)
            return [f"swipe(({x1}, {y1}), ({x2}, {y2}), duration={duration})"], None
        except (json.JSONDecodeError, KeyError, TypeError, ValueError):
            return self._todo(step, f"INVALID_SWIPE_PARAMS: {step.params!r}")

    @action("SCROLL")
    def handle_scroll(self, step, ctx):
        raw = (step.params.strip() if step.params else "") or "up"
        try:
            direction = json.loads(raw)["direction"].lower()
        except (json.JSONDecodeError, KeyError, TypeError, AttributeError):
            direction = raw.lower()
        if direction not in _SCROLL_PRESETS:
            return self._todo(
                step,
                f"INVALID_SCROLL_DIRECTION: '{direction}' — use up/down/left/right",
            )
        frm, to = _SCROLL_PRESETS[direction]
        return ["w, h = G.DEVICE.get_current_resolution()", f"swipe({frm}, {to})"], None

    @action("LONG_PRESS")
    def handle_long_press(self, step, ctx):
        asset, lines, issue = self._resolve_image(step, ctx)
        if asset is None:
            return lines, issue
        return [
            f"long_click(Template({asset.resource_path!r}, threshold={asset.threshold}))"
        ], None

    @action("SLEEP")
    def handle_sleep(self, step, ctx):
        raw = step.params.strip() if step.params else ""
        try:
            if raw.startswith("{"):
                seconds = float(json.loads(raw)["seconds"])
            else:
                seconds = float(raw)
            return [f"sleep({seconds})"], None
        except (json.JSONDecodeError, KeyError, ValueError, TypeError):
            return self._todo(
                step,
                f'INVALID_SLEEP_PARAMS: {step.params!r} — provide seconds as number or {{"seconds": 2.5}}',
            )

    @action("BACK")
    def handle_back(self, step, ctx):
        return ['keyevent("BACK")'], None

    @action("HOME")
    def handle_home(self, step, ctx):
        return ['keyevent("HOME")'], None

    @action("SNAPSHOT")
    def handle_snapshot(self, step, ctx):
        filename = (step.params.strip() if step.params else "") or step.suite_id.lower()
        if not filename.endswith(".png"):
            filename += ".png"
        return [f"snapshot(filename={filename!r})"], None

    @action("STOP_APP")
    def handle_stop_app(self, step, ctx):
        if not ctx.app_package:
            return self._todo(step, "STOP_APP_NEEDS_PACKAGE — pass via --app-package")
        return [f'stop_app("{ctx.app_package}")'], None

    # ------------------------------------------------------------------- #
    # Generation methods
    # ------------------------------------------------------------------- #
    def step_label(self, step: Step) -> str:
        return f"# step {step.step_no}" + (
            f" - {step.expected}" if step.expected else ""
        )

    def _generate_step(self, step: Step, ctx: GenCtx) -> HandlerResult:
        label = self.step_label(step)
        handler = self._HANDLERS.get(step.action)
        if handler is None:
            raise AirtestError(
                f"UNSUPPORTED_ACTION '{step.action}' in suite {step.suite_id}"
            )
        lines, issue = handler(self, step, ctx)
        body = [l for l in lines if l]
        return [label] + body, issue

    def wrap_main_body(self, step_lines, suite_id):
        return step_lines

    def generate_suite_script(
        self, steps: list[Step], ctx: GenCtx, source_name: str, suite_id: str
    ) -> tuple[str, list[GenerationIssue]]:
        body_lines = []
        issues = []
        for step in steps:
            lines, issue = self._generate_step(step, ctx)
            body_lines.extend(lines)
            if issue:
                issues.append(issue)

        wrapped = self.wrap_main_body(body_lines, suite_id)

        if not any(l and not l.lstrip().startswith("#") for l in wrapped):
            wrapped.append("pass")

        indent = "    "
        body = ("\n" + indent).join(wrapped) if wrapped else "pass"
        prologue = self.MODULE_PROLOGUE.strip()
        prologue_block = ("\n\n" + prologue) if prologue else ""

        return (
            f"# Auto-generated by excel_to_airtest.py ({type(self).__name__})\n"
            f"# Source: {source_name}\n"
            f"# Suite : {suite_id}\n"
            f"{self.IMPORTS}"
            f"{prologue_block}\n"
            f"\n\n"
            f"def main():\n"
            f"    {body}\n"
            f"\n\n"
            f'if __name__ == "__main__":\n'
            f"    main()\n"
        ), issues

    @classmethod
    def main(cls):
        from cli import build_parser, exec_cli, discover_projects, scaffold_project
        import os

        discover_projects()
        parser = build_parser(
            cls.DESCRIPTION, cls.DEFAULT_APP_PACKAGE, cls.DEFAULT_PLAN
        )

        args = parser.parse_args()

        if args.list_projects:
            from registry import list_projects

            list_projects()
            sys.exit(0)

        if args.init_project:
            try:
                out = scaffold_project(args.init_project)
                print(f"Scaffolded new project at {out}")
            except Exception as e:
                sys.exit(f"ERROR: {e}")
            sys.exit(0)

        if args.init_excel:
            import importlib.util
            template_src = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "templates",
                "excel_template.py",
            )
            spec = importlib.util.spec_from_file_location("excel_template", template_src)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            out = mod.generate_template(os.path.join(os.getcwd(), "template.xlsx"))
            print(f"Template generated: {out}")
            sys.exit(0)

        if not args.excel_file:
            parser.print_help()
            sys.exit(1)

        # Select project class if specified, else use self
        gen_cls = cls
        if args.project:
            from registry import _PROJECT_REGISTRY

            key = args.project.lower()
            if key in _PROJECT_REGISTRY:
                gen_cls = _PROJECT_REGISTRY[key]
            else:
                known = ", ".join(sorted(_PROJECT_REGISTRY)) or "(none)"
                sys.exit(
                    f"ERROR: Unknown project '{args.project}'. Known: {known}. Run --list-projects."
                )

        generator = gen_cls()
        exec_cli(generator, args)


# --------------------------------------------------------------------------- #
# Backward-compat module-level surface                                        #
# --------------------------------------------------------------------------- #

def _discover_projects(projects_dir=None):
    """Alias for cli.discover_projects (backward compat)."""
    from cli import discover_projects
    discover_projects(projects_dir)


_default_generator = AirtestGenerator()
_HANDLERS = AirtestGenerator._HANDLERS


def _issue(step, reason):
    return _default_generator._issue(step, reason)


def _todo(step, reason):
    return _default_generator._todo(step, reason)


def _resolve_image(step, ctx):
    return _default_generator._resolve_image(step, ctx)


def generate_suite_script(steps, ctx, source_name, suite_id):
    return _default_generator.generate_suite_script(steps, ctx, source_name, suite_id)


def main():
    AirtestGenerator.main()


if __name__ == "__main__":
    main()
