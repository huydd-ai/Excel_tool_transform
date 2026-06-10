"""
excel_tool/writer.py — File output and generation-report writer.

Two responsibilities:
  1. Write a generated script into a ``.air`` directory on disk.
  2. Produce a structured ``generation_report.txt`` with issues, asset warnings, and
     flow references extracted from Action_Logic.
"""

from __future__ import annotations

import os

from models import AirtestError, FlowDoc, GenerationIssue, ValidationIssue
from parsers import safe_name


def write_suite(script: str, plan_dir: str, suite_id: str) -> str:
    """Write `script` to ``{plan_dir}/{suite_id}.air/{suite_id}.py``.

    Returns the absolute path of the written file.

    Raises:
        AirtestError: if the directory or file cannot be written.
    """
    safe = safe_name(suite_id)
    air_dir = os.path.join(plan_dir, f"{safe}.air")
    out_file = os.path.join(air_dir, f"{safe}.py")
    try:
        os.makedirs(air_dir, exist_ok=True)
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(script)
    except OSError as exc:
        raise AirtestError(f"Failed to write suite to {out_file}: {exc}") from exc
    return out_file


def write_report(
    generator_name: str,
    source: str,
    output_dir: str,
    written: list[tuple[str, str]],
    gen_issues: list[GenerationIssue],
    val_issues: list[ValidationIssue],
    flows: list[FlowDoc],
) -> str:
    """Write ``generation_report.txt`` into *output_dir*.

    Returns the absolute path of the written report file.
    """
    lines: list[str] = [
        "=== Generation Report ===",
        f"Generator : {generator_name}",
        f"Source    : {source}",
        "",
        f"--- Generated Suites ({len(written)}) ---",
    ]
    for suite, path in written:
        lines.append(f"  {suite}  ->  {path}")
    if not written:
        lines.append("  (none)")

    lines += ["", f"--- Generation Issues ({len(gen_issues)}) ---"]
    if gen_issues:
        for i in gen_issues:
            lines.append(
                f"  row {i.excel_row:>4} | suite {i.suite_id} | step {i.step_no}: {i.reason}"
            )
    else:
        lines.append("  (none)")

    lines += ["", f"--- Asset Validation - file not found ({len(val_issues)}) ---"]
    if val_issues:
        for v in val_issues:
            lines.append(f"  {v.component}: {v.path}")
    else:
        lines.append("  (none)")

    lines += [
        "",
        f"--- Action_Logic Flows ({len(flows)}) - reference only, not generated ---",
    ]
    if flows:
        for fl in flows:
            lines.append(f"  {fl.logic_id} | {fl.action_name} ({fl.target_page})")
    else:
        lines.append("  (none)")

    report_path = os.path.join(output_dir, "generation_report.txt")
    try:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
    except OSError as exc:
        raise AirtestError(f"Failed to write report to {report_path}: {exc}") from exc
    return report_path
