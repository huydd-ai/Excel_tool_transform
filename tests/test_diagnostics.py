"""_diagnostic_hint returns an actionable hint for each error reason."""

import pytest
from excel_to_airtest import _diagnostic_hint


@pytest.mark.parametrize(
    "reason,expected_fragment",
    [
        ("MISSING_TARGET", "Target_ID"),
        ("UNKNOWN_TARGET 'btn_foo'", "btn_foo"),
        ("UNKNOWN_TARGET 'btn_foo'", "Object_Repository"),
        ("MISSING_RESOURCE_PATH for 'btn_bar'", "btn_bar"),
        ("MISSING_RESOURCE_PATH for 'btn_bar'", "Resource_Path"),
        ("START_APP_NEEDS_PACKAGE", "--app-package"),
        ("STOP_APP_NEEDS_PACKAGE", "--app-package"),
        ("INVALID_PARAMS_JSON", "JSON"),
        ("SWIPE_NEEDS_PARAMS", "from"),
        ("INVALID_SWIPE_PARAMS", "JSON"),
        ("INVALID_SCROLL_DIRECTION", "up/down/left/right"),
        ("INVALID_SLEEP_PARAMS", "seconds"),
        ("READ_TEXT", "v2"),
        ("UNSUPPORTED_LOCATOR 'OCR'", "OCR"),
    ],
)
def test_hint_contains_expected_fragment(reason, expected_fragment):
    hint = _diagnostic_hint(reason)
    assert expected_fragment.lower() in hint.lower(), (
        f"hint for '{reason}' missing '{expected_fragment}': got '{hint}'"
    )


def test_hint_unknown_reason_returns_generic():
    hint = _diagnostic_hint("TOTALLY_UNKNOWN_ERROR_XYZ")
    assert len(hint) > 0
    assert "check" in hint.lower() or "excel" in hint.lower()


def test_hint_unknown_target_with_no_quotes():
    """reason with no quotes should not crash — returns safe fallback."""
    hint = _diagnostic_hint("UNKNOWN_TARGET noquotes")
    assert "Object_Repository" in hint


def test_hint_unknown_target_with_multiple_quoted_names():
    """First quoted name is extracted correctly when multiple exist."""
    hint = _diagnostic_hint("UNKNOWN_TARGET 'btn_a' in 'suite_x'")
    assert "btn_a" in hint


def test_hint_missing_resource_path_with_no_quotes():
    hint = _diagnostic_hint("MISSING_RESOURCE_PATH for noquotes")
    assert "Resource_Path" in hint
