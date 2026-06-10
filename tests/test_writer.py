"""Tests for writer.py — write_suite / write_report file output + error handling."""

import os

import pytest

from models import AirtestError, FlowDoc, GenerationIssue, ValidationIssue
from writer import write_suite, write_report


class TestWriteSuite:
    def test_creates_air_dir_and_py_file(self, tmp_path):
        out = write_suite("print('hi')\n", str(tmp_path), "MY_SUITE")
        assert os.path.isfile(out)
        assert out.endswith(os.path.join("MY_SUITE.air", "MY_SUITE.py"))
        assert open(out, encoding="utf-8").read() == "print('hi')\n"

    def test_sanitizes_suite_id(self, tmp_path):
        out = write_suite("x=1\n", str(tmp_path), "bad/name:here")
        # safe_name replaces illegal filesystem chars with underscores
        leaf = os.path.basename(out)
        assert "/" not in leaf and ":" not in leaf
        assert leaf.endswith(".py")
        assert os.path.isfile(out)

    def test_raises_airtest_error_on_unwritable_path(self, tmp_path):
        # A file standing where a directory is needed -> os.makedirs raises OSError.
        blocker = tmp_path / "blocker"
        blocker.write_text("not a dir")
        with pytest.raises(AirtestError):
            write_suite("x=1\n", str(blocker), "SUITE")


class TestWriteReport:
    def test_writes_all_sections(self, tmp_path):
        path = write_report(
            generator_name="AirtestGenerator",
            source="plan.xlsx",
            output_dir=str(tmp_path),
            written=[("S1", "/out/S1.air/S1.py")],
            gen_issues=[GenerationIssue("S1", 2, 5, "MISSING_TARGET")],
            val_issues=[ValidationIssue("heart_icon", "./a.png")],
            flows=[FlowDoc("FLOW_1", "Buy Heart", "cmd", "HomePage")],
        )
        assert os.path.isfile(path)
        text = open(path, encoding="utf-8").read()
        assert "Generated Suites (1)" in text
        assert "Generation Issues (1)" in text
        assert "MISSING_TARGET" in text
        assert "heart_icon" in text
        assert "FLOW_1" in text

    def test_empty_sections_render_none(self, tmp_path):
        path = write_report(
            generator_name="G",
            source="s.xlsx",
            output_dir=str(tmp_path),
            written=[],
            gen_issues=[],
            val_issues=[],
            flows=[],
        )
        text = open(path, encoding="utf-8").read()
        assert text.count("(none)") == 4

    def test_raises_airtest_error_on_unwritable_dir(self, tmp_path):
        # output_dir points at a file -> open(report_path) raises OSError.
        blocker = tmp_path / "file_not_dir"
        blocker.write_text("x")
        with pytest.raises(AirtestError):
            write_report(
                generator_name="G",
                source="s.xlsx",
                output_dir=str(blocker),
                written=[],
                gen_issues=[],
                val_issues=[],
                flows=[],
            )
