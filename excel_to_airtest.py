# Auto-generated tool: excel_to_airtest.py
import argparse
import math
import os
import re
import sys

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip install openpyxl")
    sys.exit(1)


def is_blank(val):
    if val is None:
        return True
    if isinstance(val, float) and math.isnan(val):
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False


def safe_str(val):
    if is_blank(val):
        return ""
    return str(val).strip()


def safe_float(val, default):
    if is_blank(val):
        return default
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def read_sheet_as_dicts(wb, sheet_name):
    if sheet_name not in wb.sheetnames:
        return None, f"Sheet '{sheet_name}' not found in workbook."
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], None
    headers = [safe_str(h) for h in rows[0]]
    records = []
    for row in rows[1:]:
        if all(is_blank(c) for c in row):
            continue
        record = {}
        for i, h in enumerate(headers):
            record[h] = row[i] if i < len(row) else None
        records.append(record)
    return records, None


def parse_swipe_coords(hint):
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


def validate_assets(repo, base_dir, invalid_assets):
    for name, asset in repo.items():
        img = asset.get("image_path", "")
        if not img or img.upper() == "NONE":
            continue
        full = img if os.path.isabs(img) else os.path.join(base_dir, img)
        if not os.path.isfile(full):
            invalid_assets.append({"component": name, "path": img})
            print(f"WARNING: Asset not found: {img} (component: {name})")


def build_object_repository(wb):
    records, err = read_sheet_as_dicts(wb, "Object_Repository")
    if err:
        print(f"WARNING: {err} — object lookup will be empty.")
        return {}
    repo = {}
    for rec in records:
        name = safe_str(rec.get("component_name", ""))
        if not name:
            continue
        repo[name] = {
            "page_id": safe_str(rec.get("page_id", "")),
            "image_path": safe_str(rec.get("image_path", "")),
            "threshold": safe_float(rec.get("threshold", None), 0.7),
            "position_hint": safe_str(rec.get("position_hint", "")),
        }
    return repo


def generate_step_code(step, repo, skipped, missing_assets):
    step_id = safe_str(step.get("step_id", ""))
    action = safe_str(step.get("action", "")).upper()
    target = safe_str(step.get("target", ""))
    notes = safe_str(step.get("notes", ""))
    wait_after = safe_float(step.get("wait_after", None), 1)

    comment = f"# {step_id}: {notes}" if notes else f"# {step_id}"
    lines = [comment]

    if action == "WAIT":
        lines.append(f"sleep({wait_after})")
        return lines, None

    if action == "SWIPE":
        hint = ""
        if target and target in repo:
            hint = repo[target].get("position_hint", "")
        elif target:
            hint = target
        coords = parse_swipe_coords(hint) if hint else None
        if coords:
            lines.append(f"swipe({coords[0]}, {coords[1]})")
        else:
            skipped.append({"step_id": step_id, "reason": f"SWIPE_COORDS_MISSING '{target}'"})
            lines.append(f"# TODO: SWIPE_COORDS_MISSING — set position_hint 'x1,y1,x2,y2' for '{target}'")
        return lines, None

    if action == "INPUT_TEXT":
        data = safe_str(step.get("data", "")) or notes
        lines.append(f'text("{data}")')
        return lines, None

    if action in ("CLICK", "ASSERT_EXISTS"):
        if not target or target not in repo:
            missing_assets.append(target or "(empty)")
            skipped.append({
                "step_id": step_id,
                "reason": f"MISSING_ASSET '{target}'",
            })
            lines.append(f"# TODO: MISSING_ASSET '{target}'")
            return lines, f"MISSING_ASSET '{target}'"

        asset = repo[target]
        img = asset["image_path"]
        threshold = asset["threshold"]

        if action == "CLICK":
            lines.append(f'touch(Template(r"{img}", threshold={threshold}))')
        else:
            lines.append(
                f'assert_exists(Template(r"{img}", threshold={threshold}), timeout={wait_after})'
            )
        return lines, None

    skipped.append({
        "step_id": step_id,
        "reason": f"UNSUPPORTED_ACTION '{action}'",
    })
    lines.append(f"# UNSUPPORTED_ACTION: '{action}'")
    return lines, f"UNSUPPORTED_ACTION '{action}'"


def generate_script(excel_file, sheet_name, output_dir, plan, write_report):
    wb = openpyxl.load_workbook(excel_file, data_only=True)

    repo = build_object_repository(wb)

    base_dir = os.path.dirname(os.path.abspath(excel_file))
    invalid_assets = []
    validate_assets(repo, base_dir, invalid_assets)

    steps, err = read_sheet_as_dicts(wb, sheet_name)
    if err:
        print(f"ERROR: {err}")
        sys.exit(1)

    skipped = []
    missing_assets = []
    generated_lines = []

    for step in steps:
        lines, issue = generate_step_code(step, repo, skipped, missing_assets)
        generated_lines.extend(lines)

    indent = "    "
    body = ("\n" + indent).join(generated_lines) if generated_lines else "pass"

    excel_basename = os.path.basename(excel_file)
    script_content = f"""# Auto-generated by excel_to_airtest.py
# Source: {excel_basename} | Sheet: {sheet_name}
from airtest.core.api import *
from airtest.aircv import Template
import time


def main():
    {body}


if __name__ == "__main__":
    main()
"""

    plan_dir = os.path.join(output_dir, plan)
    air_dir = os.path.join(plan_dir, f"{sheet_name}.air")
    os.makedirs(air_dir, exist_ok=True)

    out_file = os.path.join(air_dir, f"{sheet_name}.py")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(script_content)

    print(f"Generated: {out_file}")

    if write_report:
        report_lines = []
        report_lines.append("=== Generation Report ===")
        report_lines.append(f"Source:  {excel_file}")
        report_lines.append(f"Sheet:   {sheet_name}")
        report_lines.append(f"Plan:    {plan}")
        report_lines.append("")
        report_lines.append("--- Generated Files ---")
        report_lines.append(out_file)
        report_lines.append("")

        report_lines.append("--- Skipped Steps ---")
        if skipped:
            for s in skipped:
                report_lines.append(f"  {s['step_id']}: {s['reason']}")
        else:
            report_lines.append("  (none)")
        report_lines.append("")

        report_lines.append("--- Missing Assets (not in Object_Repository) ---")
        unique_missing = sorted(set(missing_assets))
        if unique_missing:
            for a in unique_missing:
                report_lines.append(f"  {a}")
        else:
            report_lines.append("  (none)")
        report_lines.append("")

        report_lines.append("--- Invalid Assets (file not found on disk) ---")
        if invalid_assets:
            for a in invalid_assets:
                report_lines.append(f"  {a['component']}: {a['path']}")
        else:
            report_lines.append("  (none)")

        report_path = os.path.join(output_dir, "generation_report.txt")
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(report_lines) + "\n")
        print(f"Report:    {report_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Convert Excel V2 test plan to Airtest Python script."
    )
    parser.add_argument("excel_file", help="Path to Excel (.xlsx) file")
    parser.add_argument(
        "--sheet",
        default="Test_Execution",
        help="Sheet name to read steps from (default: Test_Execution)",
    )
    parser.add_argument(
        "--output",
        default="./output",
        help="Output directory (default: ./output)",
    )
    parser.add_argument(
        "--plan",
        default=None,
        help="Subfolder name under output dir (default: stem of excel filename)",
    )
    parser.add_argument(
        "--report",
        action="store_true",
        help="Write generation_report.txt in output dir",
    )
    args = parser.parse_args()

    if not os.path.isfile(args.excel_file):
        print(f"ERROR: File not found: {args.excel_file}")
        sys.exit(1)

    plan = args.plan if args.plan else os.path.splitext(os.path.basename(args.excel_file))[0]

    generate_script(
        excel_file=args.excel_file,
        sheet_name=args.sheet,
        output_dir=args.output,
        plan=plan,
        write_report=args.report,
    )


if __name__ == "__main__":
    main()
