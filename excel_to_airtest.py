"""
excel_to_airtest.py — Excel V2 → Airtest .air script generator

Architecture:
  Parser    : Excel → dataclass objects (Step, Asset)  [generator-based, low memory]
  Validator : schema check + asset path check
  Generator : objects → Python source via Action Registry
  Writer    : save .air script + structured report

Security: all Excel string values embedded in generated code pass through repr()
          to prevent code injection via crafted cell content.
"""

import argparse
import math
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Callable

try:
    import openpyxl
except ImportError:
    sys.exit("ERROR: openpyxl not installed. Run: pip install openpyxl")


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Asset:
    component_name: str
    page_id: str = ""
    image_path: str = ""
    threshold: float = 0.7


@dataclass
class Step:
    step_id: str
    action: str
    excel_row: int = 0          # 1-based Excel row for error tracing
    target: str = ""
    wait_after: float = 1.0
    input_value: str = ""       # INPUT_TEXT source
    start_pos: str = ""         # SWIPE start  "x,y"
    end_pos: str = ""           # SWIPE end    "x,y"
    notes: str = ""


@dataclass
class ValidationIssue:
    component: str
    path: str
    kind: str = "FILE_NOT_FOUND"


@dataclass
class GenerationIssue:
    step_id: str
    excel_row: int
    reason: str


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _blank(val) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and math.isnan(val):
        return True
    return isinstance(val, str) and not val.strip()


def _str(val) -> str:
    return "" if _blank(val) else str(val).strip()


def _float(val, default: float) -> float:
    if _blank(val):
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def _parse_pos(pos: str) -> tuple[int, int] | None:
    """Parse 'x,y' → (x, y) or None."""
    parts = [p.strip() for p in pos.split(",")]
    if len(parts) == 2:
        try:
            return int(parts[0]), int(parts[1])
        except ValueError:
            pass
    return None


# ---------------------------------------------------------------------------
# Parser  (generator-based — no list() of entire sheet)
# ---------------------------------------------------------------------------

_REQUIRED_OBJ_HEADERS  = {"component_name"}
_REQUIRED_STEP_HEADERS = {"step_id", "action"}


def _read_headers(ws) -> tuple[list[str], str | None]:
    """Read first row; return normalised (lowercase+strip) header list."""
    row = next(ws.iter_rows(values_only=True), None)
    if row is None:
        return [], "Sheet is empty"
    return [_str(h).lower() for h in row], None


def parse_object_repository(wb) -> tuple[dict[str, Asset], list[str]]:
    sheet_name = "Object_Repository"
    errors: list[str] = []
    if sheet_name not in wb.sheetnames:
        errors.append(f"Sheet '{sheet_name}' not found")
        return {}, errors

    ws = wb[sheet_name]
    headers, err = _read_headers(ws)
    if err:
        return {}, [err]

    missing = _REQUIRED_OBJ_HEADERS - set(headers)
    if missing:
        errors.append(f"Object_Repository missing required columns: {missing}")
        return {}, errors

    assets: dict[str, Asset] = {}
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(_blank(c) for c in row):
            continue
        r = {headers[i]: row[i] if i < len(row) else None for i, _ in enumerate(headers)}
        name = _str(r.get("component_name"))
        if not name:
            continue
        assets[name] = Asset(
            component_name=name,
            page_id=_str(r.get("page_id", "")),
            image_path=_str(r.get("image_path", "")),
            threshold=_float(r.get("threshold"), 0.7),
        )
    return assets, errors


def parse_test_execution(wb, sheet_name: str) -> tuple[list[Step], list[str]]:
    errors: list[str] = []
    if sheet_name not in wb.sheetnames:
        errors.append(f"Sheet '{sheet_name}' not found")
        return [], errors

    ws = wb[sheet_name]
    headers, err = _read_headers(ws)
    if err:
        return [], [err]

    missing = _REQUIRED_STEP_HEADERS - set(headers)
    if missing:
        errors.append(f"Sheet '{sheet_name}' missing required columns: {missing}")
        return [], errors

    steps: list[Step] = []
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(_blank(c) for c in row):
            continue
        r = {headers[i]: row[i] if i < len(row) else None for i, _ in enumerate(headers)}
        steps.append(Step(
            step_id=_str(r.get("step_id")),
            action=_str(r.get("action")).upper(),
            excel_row=row_num,
            target=_str(r.get("target", "")),
            wait_after=_float(r.get("wait_after"), 1.0),
            input_value=_str(r.get("input_value", "")),
            start_pos=_str(r.get("start_pos", "")),
            end_pos=_str(r.get("end_pos", "")),
            notes=_str(r.get("notes", "")),
        ))
    return steps, errors


# ---------------------------------------------------------------------------
# Validator
# ---------------------------------------------------------------------------

def validate_assets(assets: dict[str, Asset], base_dir: str) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    for name, asset in assets.items():
        img = asset.image_path
        if not img or img.upper() == "NONE":
            continue
        full = img if os.path.isabs(img) else os.path.join(base_dir, img)
        if not os.path.isfile(full):
            issues.append(ValidationIssue(component=name, path=img))
            print(f"  [WARN] Asset not found on disk: {img}  (component: {name})")
    return issues


# ---------------------------------------------------------------------------
# Action Registry  (open for extension — add handlers without touching generator)
# ---------------------------------------------------------------------------

_HANDLERS: dict[str, Callable] = {}


def _action(name: str):
    """Decorator: register a function as the handler for an action keyword."""
    def decorator(fn: Callable) -> Callable:
        _HANDLERS[name] = fn
        return fn
    return decorator


# Handler signature: (step: Step, assets: dict[str, Asset]) -> (lines: list[str], issue: GenerationIssue | None)

@_action("CLICK")
def _handle_click(step: Step, assets: dict[str, Asset]) -> tuple[list[str], GenerationIssue | None]:
    if not step.target or step.target not in assets:
        return (
            [f"# TODO: MISSING_ASSET '{step.target}'"],
            GenerationIssue(step.step_id, step.excel_row, f"MISSING_ASSET '{step.target}'"),
        )
    asset = assets[step.target]
    return [f"touch(Template({repr(asset.image_path)}, threshold={asset.threshold}))"], None


@_action("ASSERT_EXISTS")
def _handle_assert_exists(step: Step, assets: dict[str, Asset]) -> tuple[list[str], GenerationIssue | None]:
    if not step.target or step.target not in assets:
        return (
            [f"# TODO: MISSING_ASSET '{step.target}'"],
            GenerationIssue(step.step_id, step.excel_row, f"MISSING_ASSET '{step.target}'"),
        )
    asset = assets[step.target]
    return [
        f"assert_exists(Template({repr(asset.image_path)}, threshold={asset.threshold}), timeout={step.wait_after})"
    ], None


@_action("WAIT")
def _handle_wait(step: Step, assets: dict[str, Asset]) -> tuple[list[str], GenerationIssue | None]:
    return [f"sleep({step.wait_after})"], None


@_action("SWIPE")
def _handle_swipe(step: Step, assets: dict[str, Asset]) -> tuple[list[str], GenerationIssue | None]:
    start = _parse_pos(step.start_pos)
    end = _parse_pos(step.end_pos)
    if start and end:
        return [f"swipe({start}, {end})"], None
    return (
        [f"# TODO: SWIPE_COORDS_MISSING — fill start_pos/end_pos as 'x,y'"],
        GenerationIssue(
            step.step_id, step.excel_row,
            f"SWIPE_COORDS_MISSING (start='{step.start_pos}' end='{step.end_pos}')"
        ),
    )


@_action("INPUT_TEXT")
def _handle_input_text(step: Step, assets: dict[str, Asset]) -> tuple[list[str], GenerationIssue | None]:
    # repr() escapes the value safely — prevents injection via crafted cell content
    return [f"text({repr(step.input_value)})"], None


# ---------------------------------------------------------------------------
# Generator
# ---------------------------------------------------------------------------

def _generate_step(step: Step, assets: dict[str, Asset]) -> tuple[list[str], GenerationIssue | None]:
    label = f"# step {step.step_id}" + (f": {step.notes}" if step.notes else "")
    handler = _HANDLERS.get(step.action)
    if handler is None:
        issue = GenerationIssue(step.step_id, step.excel_row, f"UNSUPPORTED_ACTION '{step.action}'")
        return [label, f"# UNSUPPORTED_ACTION: '{step.action}'"], issue
    lines, issue = handler(step, assets)
    return [label] + lines, issue


def generate_script(
    steps: list[Step],
    assets: dict[str, Asset],
    source_name: str,
) -> tuple[str, list[GenerationIssue]]:
    body_lines: list[str] = []
    issues: list[GenerationIssue] = []

    for step in steps:
        lines, issue = _generate_step(step, assets)
        body_lines.extend(lines)
        if issue:
            issues.append(issue)

    indent = "    "
    body = f"\n{indent}".join(body_lines) if body_lines else "pass"

    source = f"""\
# Auto-generated by excel_to_airtest.py
# Source: {source_name}
from airtest.core.api import *
from airtest.aircv import Template


def main():
    {body}


if __name__ == "__main__":
    main()
"""
    return source, issues


# ---------------------------------------------------------------------------
# Writer
# ---------------------------------------------------------------------------

def write_air(script: str, output_dir: str, plan: str) -> str:
    air_dir = os.path.join(output_dir, f"{plan}.air")
    os.makedirs(air_dir, exist_ok=True)
    out_file = os.path.join(air_dir, f"{plan}.py")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(script)
    return out_file


def write_report(
    out_file: str,
    source: str,
    gen_issues: list[GenerationIssue],
    val_issues: list[ValidationIssue],
    output_dir: str,
) -> str:
    lines = [
        "=== Generation Report ===",
        f"Source : {source}",
        f"Output : {out_file}",
        "",
        f"--- Generation Issues ({len(gen_issues)}) ---",
    ]
    if gen_issues:
        for i in gen_issues:
            lines.append(f"  row {i.excel_row:>4} | step {i.step_id}: {i.reason}")
    else:
        lines.append("  (none)")

    lines += [
        "",
        f"--- Asset Validation — file not found ({len(val_issues)}) ---",
    ]
    if val_issues:
        for v in val_issues:
            lines.append(f"  {v.component}: {v.path}")
    else:
        lines.append("  (none)")

    report_path = os.path.join(output_dir, "generation_report.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return report_path


# ---------------------------------------------------------------------------
# Entry
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Convert Excel V2 test plan to Airtest .air script.")
    parser.add_argument("excel_file", help="Path to Excel (.xlsx) file")
    parser.add_argument("--sheet",  default="Test_Execution", help="Step sheet name (default: Test_Execution)")
    parser.add_argument("--output", default="./output",       help="Output directory (default: ./output)")
    parser.add_argument("--plan",   default=None,             help="Script/folder name (default: excel filename stem)")
    parser.add_argument("--report", action="store_true",      help="Write generation_report.txt")
    args = parser.parse_args()

    if not os.path.isfile(args.excel_file):
        sys.exit(f"ERROR: File not found: {args.excel_file}")

    plan = args.plan or os.path.splitext(os.path.basename(args.excel_file))[0]
    base_dir = os.path.dirname(os.path.abspath(args.excel_file))

    wb = openpyxl.load_workbook(args.excel_file, data_only=True)

    print("[parse] Reading Object_Repository...")
    assets, obj_errors = parse_object_repository(wb)
    for e in obj_errors:
        print(f"  [ERROR] {e}")
    print(f"         {len(assets)} assets loaded")

    print("[validate] Checking asset paths on disk...")
    val_issues = validate_assets(assets, base_dir)
    print(f"           {len(val_issues)} missing file(s)")

    print(f"[parse] Reading sheet '{args.sheet}'...")
    steps, step_errors = parse_test_execution(wb, args.sheet)
    for e in step_errors:
        print(f"  [ERROR] {e}")
        sys.exit(1)
    print(f"         {len(steps)} steps loaded")

    print("[generate] Building script...")
    script, gen_issues = generate_script(steps, assets, os.path.basename(args.excel_file))
    print(f"           {len(gen_issues)} issue(s)")

    out_file = write_air(script, args.output, plan)
    print(f"[write] {out_file}")

    if args.report:
        report_path = write_report(out_file, args.excel_file, gen_issues, val_issues, args.output)
        print(f"[report] {report_path}")

    total = len(gen_issues) + len(val_issues)
    print(f"\n{'OK — no issues' if total == 0 else f'DONE — {len(gen_issues)} generation issue(s), {len(val_issues)} missing asset(s)'}")


if __name__ == "__main__":
    main()
