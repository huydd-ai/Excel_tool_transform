"""Pure helper functions: _safe_name / _blank / _str / _int / _float."""
import math

from excel_to_airtest import _blank, _float, _int, _safe_name, _str


def test_safe_name_preserves_valid_chars():
    assert _safe_name("TC_MISSION_1") == "TC_MISSION_1"
    assert _safe_name("a.b-c_1") == "a.b-c_1"


def test_safe_name_collapses_invalid_chars():
    assert _safe_name("TC/MISSION#1") == "TC_MISSION_1"
    assert _safe_name("a b  c") == "a_b_c"


def test_safe_name_strips_leading_trailing_dots_underscores():
    assert _safe_name("__TC__") == "TC"
    assert _safe_name(".hidden.") == "hidden"


def test_safe_name_falls_back_when_empty():
    assert _safe_name("###") == "unnamed"
    assert _safe_name("") == "unnamed"


def test_blank_detects_none_nan_and_whitespace():
    assert _blank(None)
    assert _blank(math.nan)
    assert _blank("")
    assert _blank("   ")


def test_blank_rejects_real_values():
    assert not _blank("x")
    assert not _blank(0)
    assert not _blank(0.0)


def test_str_trims_and_handles_none():
    assert _str("  x  ") == "x"
    assert _str(None) == ""
    assert _str(42) == "42"


def test_int_defaults_on_blank_or_bad():
    assert _int(None, 9) == 9
    assert _int("", 9) == 9
    assert _int("bad", 7) == 7
    assert _int("3.0", 0) == 3
    assert _int(5, 0) == 5


def test_float_defaults_on_blank_or_bad():
    assert _float(None, 1.5) == 1.5
    assert _float("", 1.5) == 1.5
    assert _float("bad", 2.0) == 2.0
    assert _float("0.85", 0.0) == 0.85
