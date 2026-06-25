"""CLI error-path tests for excel_to_airtest.py."""

import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAIN = os.path.join(ROOT, "excel_to_airtest.py")


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
