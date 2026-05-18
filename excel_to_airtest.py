"""
excel_to_airtest.py — Excel V2 → Airtest .air script generator

Architecture:
  Parser    : Excel → dataclass objects (Step, Asset)
  Validator : check asset paths exist on disk
  Generator : objects → Python source code
  Writer    : save .air script + optional report
"""

import argparse
import math
import os
import re
import sys
from dataclasses import dataclass, field

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
    position_hint: str = ""


@dataclass
class Step:
    step_id: str
    action: str
    target: str = ""
    wait_after: float = 1.0
    data: str = ""
    notes: str = ""


@dataclass
class ValidationIssue:
    component: str
    path: str
    kind: str = "FILE_NOT_FOUND"


@dataclass
class GenerationIssue:
    step_id: str
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


def _sheet_to_dicts(wb, sheet_name: str) -> tuple[list[dict], str | None]:
    if sheet_name not in wb.sheetnames:
        return [], f"Sheet '{sheet_name}' not found"
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], None
    headers = [_str(h) for h in rows[0]]
    records = []
    for row in rows[1:]:
        if all(_blank(c) for c in row):
            continue
        records.append({headers[i]: row[i] if i < len(row) else None for i, _ in enumerate(headers)})
    return records, None


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

def parse_object_repository(wb) -> dict[str, Asset]:
    records, err = _sheet_to_dicts(wb, "Object_Repository")
    if err:
        print(f"WARNING: {err} — object lookup disabled")
        return {}
    assets = {}
    for r in records:
        name = _str(r.get("component_name"))
        if not name:
            continue
        assets[name] = Asset(
            component_name=name,
            page_id=_str(r.get("page_id")),
            image_path=_str(r.get("image_path")),
            threshold=_float(r.get("threshold"), 0.7),
            position_hint=_str(r.get("position_hint")),
        )
    return assets


def parse_test_execution(wb, sheet_name: str) -> tuple[list[Step], str | None]:
    records, err = _sheet_to_dicts(wb, sheet_name)
    if err:
        return [], err
    steps = []
    for r in records:
        steps.append(Step(
            step_id=_str(r.get("step_id")),
            action=_str(r.get("action")).upper(),
            target=_str(r.get("target")),
            wait_after=_float(r.get("wait_after"), 1.0),
            data=_str(r.get("data")),
            notes=_str(r.get("notes")),
        ))
    return steps, None


# ---------------------------------------------------------------------------
# Validator
# ---------------------------------------------------------------------------

def validate_assets(assets: dict[str, Asset], base_dir: str) -> list[ValidationIssue]:
    issues = []
    for name, asset in assets.items():
        img = asset.image_path
        if not img or img.upper() == "NONE":
            continue
        full = img if os.path.isabs(img) else os.path.join(base_dir, img)
        if not os.path.isfile(full):
            issues.append(ValidationIssue(component=name, path=img))
            print(f"  [WARN] Asset not found on disk: {img}  (component: {name})")
    return issues


def _parse_swipe_coords(hint: str) -> tuple | None:
    """Parse '(x1,y1)-(x2,y2)' or 'x1,y1,x2,y2' → ((x1,y1),(x2,y2)) or None."""
    m = re.match(r'\((\d+)\s*,\s*(\d+)\)\s*-\s*\((\d+)\s*,\s*(\d+)\)', hint.strip())
    if m:
        return (int(m.group(1)), int(m.group(2))), (int(m.group(3)), int(m.group(4)))
    parts = [p.strip() for p in hint.split(",")]
    if len(parts) == 4:
        try:
            return (int(parts[0]), int(parts[1])), (int(parts[2]), int(parts[3]))
        except ValueError:
            pass
    return None


# ---------------------------------------------------------------------------
# Generator
# ---------------------------------------------------------------------------

def _generate_step(step: Step, assets: dict[str, Asset]) -> tuple[list[str], GenerationIssue | None]:
    label = f"# step {step.step_id}" + (f": {step.notes}" if step.notes else "")
    lines = [label]

    match step.action:
        case "WAIT":
            lines.append(f"sleep({step.wait_after})")
            return lines, None

        case "SWIPE":
            hint = assets[step.target].position_hint if step.target in assets else step.target
            coords = _parse_swipe_coords(hint) if hint else None
            if coords:
                lines.append(f"swipe({coords[0]}, {coords[1]})")
            else:
                issue = GenerationIssue(step.step_id, f"SWIPE_COORDS_MISSING '{step.target}'")
                lines.append(f"# TODO: SWIPE_COORDS_MISSING — add position_hint 'x1,y1,x2,y2' for '{step.target}'")
                return lines, issue
            return lines, None

        case "INPUT_TEXT":
            value = step.data or step.notes
            lines.append(f'text("{value}")')
            return lines, None

        case "CLICK" | "ASSERT_EXISTS":
            if not step.target or step.target not in assets:
                issue = GenerationIssue(step.step_id, f"MISSING_ASSET '{step.target}'")
                lines.append(f"# TODO: MISSING_ASSET '{step.target}'")
                return lines, issue
            asset = assets[step.target]
            img, t = asset.image_path, asset.threshold
            if step.action == "CLICK":
                lines.append(f'touch(Template(r"{img}", threshold={t}))')
            else:
                lines.append(f'assert_exists(Template(r"{img}", threshold={t}), timeout={step.wait_after})')
            return lines, None

        case _:
            issue = GenerationIssue(step.step_id, f"UNSUPPORTED_ACTION '{step.action}'")
            lines.append(f"# UNSUPPORTED_ACTION: '{step.action}'")
            return lines, issue


def generate_script(steps: list[Step], assets: dict[str, Asset], source_name: str) -> tuple[str, list[GenerationIssue]]:
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
        f"--- Skipped / Issues ({len(gen_issues)}) ---",
    ]
    lines += [f"  step {i.step_id}: {i.reason}" for i in gen_issues] or ["  (none)"]
    lines += [
        "",
        f"--- Asset Validation ({len(val_issues)} missing on disk) ---",
    ]
    lines += [f"  {v.component}: {v.path}" for v in val_issues] or ["  (none)"]

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
    assets = parse_object_repository(wb)
    print(f"        {len(assets)} assets loaded")

    print("[validate] Checking asset paths...")
    val_issues = validate_assets(assets, base_dir)
    print(f"           {len(val_issues)} missing file(s)")

    print(f"[parse] Reading sheet '{args.sheet}'...")
    steps, err = parse_test_execution(wb, args.sheet)
    if err:
        sys.exit(f"ERROR: {err}")
    print(f"        {len(steps)} steps loaded")

    print("[generate] Building script...")
    script, gen_issues = generate_script(steps, assets, os.path.basename(args.excel_file))
    print(f"           {len(gen_issues)} issue(s)")

    out_file = write_air(script, args.output, plan)
    print(f"[write] {out_file}")

    if args.report:
        report_path = write_report(out_file, args.excel_file, gen_issues, val_issues, args.output)
        print(f"[report] {report_path}")

    if gen_issues or val_issues:
        print(f"\nSummary: {len(gen_issues)} generation issue(s), {len(val_issues)} missing asset(s)")
    else:
        print("\nOK — no issues")


if __name__ == "__main__":
    main()
