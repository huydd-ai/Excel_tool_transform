"""CLI error-path tests for excel_to_airtest.py and rules_to_suites.py."""

import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAIN = os.path.join(ROOT, "excel_to_airtest.py")
RULES = os.path.join(ROOT, "rules_to_suites.py")


def _run(script, *cli_args):
    return subprocess.run(
        [sys.executable, script, *cli_args],
        capture_output=True,
        text=True,
        cwd=ROOT,
    )


class TestMainCliErrors:
    def test_missing_excel_file_exits_nonzero(self, tmp_path):
        r = _run(MAIN, str(tmp_path / "nope.xlsx"))
        assert r.returncode != 0
        assert "File not found" in (r.stdout + r.stderr)

    def test_unknown_project_exits_with_known_list(self, tmp_path):
        r = _run(MAIN, str(tmp_path / "nope.xlsx"), "--project", "does_not_exist")
        assert r.returncode != 0
        out = r.stdout + r.stderr
        assert "Unknown project" in out
        assert "Known:" in out


class TestRulesCliErrors:
    def test_missing_excel_file_exits_nonzero(self, tmp_path):
        r = _run(RULES, str(tmp_path / "nope.xlsx"))
        assert r.returncode != 0
        assert "File not found" in (r.stdout + r.stderr)

    def test_unknown_project_exits_with_known_list(self, tmp_path):
        # Build a minimal valid rules workbook so the file-exists check passes.
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Feature Rules"
        ws.append(["Feature", "Logic Item", "Rule / Condition", "Expected Behavior"])
        ws.append(["Heart System", "Consumption", "Fail", "Deduct"])
        path = tmp_path / "r.xlsx"
        wb.save(path)

        r = _run(RULES, str(path), "--project", "does_not_exist")
        assert r.returncode != 0
        out = r.stdout + r.stderr
        assert "Unknown project" in out
        assert "Known:" in out
