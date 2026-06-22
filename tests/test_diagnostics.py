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
