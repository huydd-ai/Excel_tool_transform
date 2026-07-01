"""
excel_tool/cli.py — Argument parsing, project scaffolding, CLI entry point.

Responsibilities:
  1. Build the ``argparse`` parser (shared by base generator and subclass
     extensions).
  2. Scaffold new project files with ``--init-project``.
  3. Orchestrate ``excel_to_airtest`` execution (parse → validate → generate →
     write → report).
"""

from __future__ import annotations

import argparse
import importlib
import os
import re
import sys

from models import AirtestError, GenCtx
from parsers import (
    parse_action_logic,
    parse_object_repository,
    parse_test_execution,
    validate_assets,
)
from hints import diagnostic_hint
from writer import write_suite, write_report

# ── Project scaffolding ───────────────────────────────────────────────────────


def _classname_from_name(name: str) -> str:
    """Convert ``my_game`` → ``MyGame`` (used by ``--init-project``)."""
    return (
        "".join(part.capitalize() for part in re.split(r"[_\-\s]+", name)) or "Project"
    )


_SCAFFOLD_TEMPLATE = """\
# projects/{name}.py  — generated scaffold
# Edit DEFAULT_APP_PACKAGE and IMPORTS, then run:
#   python excel_to_airtest.py --list-projects
# to confirm your project is discovered.

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_to_airtest import AirtestGenerator, register_project, action


@register_project("{name}")
class {classname}Generator(AirtestGenerator):
    DEFAULT_APP_PACKAGE = "com.example.{name}"   # <- change this
    IMPORTS = "from airtest.core.api import *"    # <- change this
    MODULE_PROLOGUE = ""                           # <- page-object singletons (optional)

    # Override actions only if your framework differs from base Airtest.
    # All base actions are inherited automatically.

    # @action("TAP")
    # def handle_tap(self, step, ctx):
    #     asset, lines, issue = self._resolve_image(step, ctx)
    #     if asset is None: return lines, issue
    #     return [f"my_framework.tap({{asset.resource_path!r}})"], None

    # def wrap_main_body(self, step_lines, suite_id):
    #     return step_lines


if __name__ == "__main__":
    {classname}Generator.main()
"""


def scaffold_project(name: str, projects_dir: str | None = None) -> str:
    """Write ``projects/<name>.py`` scaffold. Raises ``FileExistsError`` if present."""
    if projects_dir is None:
        projects_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "projects"
        )
    os.makedirs(projects_dir, exist_ok=True)
    out_path = os.path.join(projects_dir, f"{name}.py")
    if os.path.exists(out_path):
        raise FileExistsError(f"Project file already exists: {out_path}")
    classname = _classname_from_name(name)
    content = _SCAFFOLD_TEMPLATE.format(name=name, classname=classname)
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(content)
    except OSError as exc:
        raise AirtestError(
            f"Failed to write project scaffold to {out_path}: {exc}"
        ) from exc
    return out_path


# ── Argument parser ───────────────────────────────────────────────────────────


def build_parser(
    description: str = "Convert a structured test-plan Excel to Airtest .air scripts (one per Suite_ID).",
    default_app_package: str = "",
    default_plan_name: str = "",
) -> argparse.ArgumentParser:
    """Create the argument parser with standard flags.

    Subclasses may call this and then add extra flags via ``parser.add_argument()``.
    """
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "excel_file", nargs="?", default=None, help="Path to Excel (.xlsx) file"
    )
    parser.add_argument("--output", default="./output", help="Output root directory")
    parser.add_argument(
        "--plan",
        default=default_plan_name,
        help="Subfolder under output (default: excel filename stem)",
    )
    parser.add_argument(
        "--objects-sheet",
        default="Object_Repository",
        help="Object repository sheet name",
    )
    parser.add_argument(
        "--actions-sheet", default="Action_Logic", help="Action logic sheet name"
    )
    parser.add_argument(
        "--steps-sheet", default="Test_Execution", help="Step sheet name"
    )
    parser.add_argument(
        "--app-package",
        default=default_app_package,
        help="App package id used by START_APP / STOP_APP",
    )
    parser.add_argument(
        "--report", action="store_true", help="Write generation_report.txt"
    )
    parser.add_argument(
        "--project", default=None, help="Project name to use (from projects/ directory)"
    )
    parser.add_argument(
        "--list-projects",
        action="store_true",
        help="List all discovered projects and exit",
    )
    parser.add_argument(
        "--init-project", metavar="NAME", help="Scaffold projects/<NAME>.py and exit"
    )
    parser.add_argument(
        "--init-excel",
        action="store_true",
        help="Generate template.xlsx in current directory and exit",
    )
    return parser


# ── DB ──


def discover_projects(projects_dir: str | None = None) -> None:
    """Import all ``*.py`` files from ``projects/`` so their ``@register_project`` decorators fire."""
    if projects_dir is None:
        projects_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "projects"
        )
    if not os.path.isdir(projects_dir):
        return
    parent = os.path.dirname(projects_dir)
    if parent not in sys.path:
        sys.path.insert(0, parent)
    for fname in sorted(os.listdir(projects_dir)):
        if fname.endswith(".py") and not fname.startswith("_"):
            module_name = f"projects.{fname[:-3]}"
            if module_name not in sys.modules:
                importlib.import_module(module_name)


# ── Entry-point ──────────────────────────────────────────────────────────────


def exec_cli(gen: object, args: argparse.Namespace) -> None:
    """Orchestrate the full generation pipeline (parse → validate → generate → write)."""
    try:
        import openpyxl
    except ImportError:
        sys.exit("ERROR: openpyxl not installed. Run: pip install openpyxl")

    if not os.path.isfile(args.excel_file):
        sys.exit(f"ERROR: File not found: {args.excel_file}")

    plan = args.plan or os.path.splitext(os.path.basename(args.excel_file))[0]
    base_dir = os.path.dirname(os.path.abspath(args.excel_file))
    plan_dir = os.path.join(args.output, plan)
    os.makedirs(plan_dir, exist_ok=True)

    wb = openpyxl.load_workbook(args.excel_file, data_only=True)
    try:
        print(f"[parse] {args.objects_sheet}...")
        assets, obj_errors = parse_object_repository(wb, args.objects_sheet)
        for e in obj_errors:
            print(f"  [ERROR] {e}")
        if obj_errors:
            sys.exit(1)
        print(f"        {len(assets)} objects loaded")

        print(f"[parse] {args.actions_sheet}...")
        keywords, flows, act_errors = parse_action_logic(wb, args.actions_sheet)
        for e in act_errors:
            print(f"  [ERROR] {e}")
        print(f"        {len(keywords)} action keywords, {len(flows)} flow descriptions")

        print("[validate] Asset paths on disk...")
        val_issues = validate_assets(assets, base_dir)
        print(f"           {len(val_issues)} missing file(s)")

        print(f"[parse] {args.steps_sheet}...")
        suites, step_errors = parse_test_execution(wb, args.steps_sheet)
        for e in step_errors:
            print(f"  [ERROR] {e}")
        if step_errors:
            sys.exit(1)
        print(
            f"        {len(suites)} suite(s), {sum(len(v) for v in suites.values())} step(s)"
        )
    finally:
        wb.close()

    ctx = GenCtx(assets=assets, app_package=args.app_package)
    source_name = os.path.basename(args.excel_file)

    print(f"[generate] Building scripts ({type(gen).__name__})...")
    written: list[tuple[str, str]] = []
    gen_issues: list = []
    for suite_id, steps in suites.items():
        script, issues = gen.generate_suite_script(steps, ctx, source_name, suite_id)
        out_path = write_suite(script, plan_dir, suite_id)
        written.append((suite_id, out_path))
        gen_issues.extend(issues)
        print(f"  [write] {suite_id} -> {out_path}  ({len(issues)} issue(s))")

    if args.report:
        report = write_report(
            generator_name=type(gen).__name__,
            source=args.excel_file,
            output_dir=args.output,
            written=written,
            gen_issues=gen_issues,
            val_issues=val_issues,
            flows=flows,
        )
        print(f"[report] {report}")

    total = len(gen_issues) + len(val_issues)
    if gen_issues or val_issues:
        print(f"\n[issues] {total} problem(s) found:")
        for i in gen_issues:
            hint = diagnostic_hint(i.reason)
            print(
                f"  row {i.excel_row:>4} | {i.suite_id} | step {i.step_no} | {i.reason}"
            )
            print(f"          -> fix: {hint}")
        for v in val_issues:
            print(f"  asset    | {v.component}: {v.path}")
            print("          -> fix: copy or create this image file at the listed path")
    print(
        f"\n{'OK - no issues' if total == 0 else f'DONE - {len(gen_issues)} generation issue(s), {len(val_issues)} missing asset(s)'}"
    )
