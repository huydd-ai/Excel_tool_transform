"""Tests for parsers.validate_assets — IMAGE asset disk validation."""

from models import Asset, LOCATOR_IMAGE, LOCATOR_OCR
from parsers import validate_assets


def test_existing_image_yields_no_issue(tmp_path):
    img = tmp_path / "heart.png"
    img.write_bytes(b"\x89PNG\r\n")
    assets = {"heart": Asset("heart", LOCATOR_IMAGE, "heart.png")}
    issues = validate_assets(assets, str(tmp_path))
    assert issues == []


def test_missing_image_yields_issue(tmp_path):
    assets = {"ghost": Asset("ghost", LOCATOR_IMAGE, "ghost.png")}
    issues = validate_assets(assets, str(tmp_path))
    assert len(issues) == 1
    assert issues[0].component == "ghost"
    assert issues[0].path == "ghost.png"
    assert issues[0].kind == "FILE_NOT_FOUND"


def test_non_image_locator_skipped(tmp_path):
    assets = {"label": Asset("label", LOCATOR_OCR, "whatever.png")}
    assert validate_assets(assets, str(tmp_path)) == []


def test_none_and_empty_resource_path_skipped(tmp_path):
    assets = {
        "a": Asset("a", LOCATOR_IMAGE, "NONE"),
        "b": Asset("b", LOCATOR_IMAGE, ""),
    }
    assert validate_assets(assets, str(tmp_path)) == []


def test_absolute_path_honored(tmp_path):
    img = tmp_path / "abs.png"
    img.write_bytes(b"x")
    assets = {"abs": Asset("abs", LOCATOR_IMAGE, str(img))}
    # base_dir is irrelevant when path is absolute
    assert validate_assets(assets, "/some/other/dir") == []
