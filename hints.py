"""
excel_tool/hints.py — Diagnostic hints and column descriptions.

Provides an actionable hint for any generation-issue reason string,
plus schema-level descriptions of Excel column names.

Constants:
    HINT_MAP          — issue-key → human-readable fix suggestion
    VERSION_HINTS     — version-keyed overrides / extensions of HINT_MAP

Functions:
    diagnostic_hint(reason)        — best hint for a generation-issue reason
    get_hints_for_version(version) — HINT_MAP + version overrides
    explain_column(col_name)       — short description of a column
"""

from __future__ import annotations

import re

# ── Hint Map ──────────────────────────────────────────────────────────────────

HINT_MAP: dict[str, str] = {
    # Target resolution
    "MISSING_TARGET": ("fill the Target_ID cell for this step in Test_Execution sheet"),
    # MISSING_RESOURCE_PATH is handled by the heuristic below
    # (embeds the object name for actionable hints)
    # START_APP / STOP_APP
    "START_APP_NEEDS_PACKAGE": ("pass --app-package com.your.app on the CLI"),
    "STOP_APP_NEEDS_PACKAGE": ("pass --app-package com.your.app on the CLI"),
    "INVALID_PARAMS_JSON": ("Params must be valid JSON or a plain number"),
    # SWIPE / SCROLL
    "SWIPE_NEEDS_PARAMS": (
        'provide JSON {"from":[x1,y1],"to":[x2,y2],"duration":0.5} in Params'
    ),
    "INVALID_SWIPE_PARAMS": (
        'Params must be valid JSON: {"from":[x1,y1],"to":[x2,y2],"duration":0.5}'
    ),
    "INVALID_SCROLL_DIRECTION": ("use one of: up/down/left/right in the Params cell"),
    # SLEEP
    "INVALID_SLEEP_PARAMS": (
        'provide seconds as a number (e.g. "2.5") or {"seconds": 2.5}'
    ),
    # Feature gaps
    "READ_TEXT": ("READ_TEXT is a v2 feature — leave as TODO stub for now"),
}

# ── Version-specific Overrides ────────────────────────────────────────────────

VERSION_HINTS: dict[str, dict[str, str]] = {}
"""Per-version hint overrides.

Future schema versions may change column names or introduce new actions.
Add version-keyed overrides here::

    VERSION_HINTS["v2.0"] = {
        "OLD_ACTION": "this action was renamed in v2.0 — use NEW_ACTION instead",
    }
"""

_QUOTED_NAME_RE = re.compile(r"'([^']*)'")

_COLUMN_DESCRIPTIONS: dict[str, str] = {
    "object_id": "unique identifier for a UI element",
    "locator_type": "how to find the element (IMAGE or OCR)",
    "resource_path": "filesystem path to the image template",
    "smart_threshold": "image-matching confidence threshold (0.0–1.0)",
    "timeout": "max wait time in seconds before giving up",
    "suite_id": "test-case identifier (one .air script per suite)",
    "step": "ordinal step number within the suite",
    "action_keyword": "action to perform (TOUCH, WAIT_FOR, …)",
    "target_id": "Object_ID from Object_Repository",
    "params": "JSON or plain-text parameters for the action",
    "expected_result": "what the test expects after this step",
    "logic_id": "identifier for an Action_Logic row (ACT_* / FLOW_*)",
    "action_name": "display name for the action/flow",
    "machine_command": "canonical keyword this row documents",
    "target_page": "associated page-object (informational)",
}

# ── Query Functions ───────────────────────────────────────────────────────────


def diagnostic_hint(reason: str) -> str:
    """Return the most specific actionable hint for a generation-issue *reason*.

    Checks ``HINT_MAP`` keys first (substring match), then falls back to
    heuristic patterns for ``UNKNOWN_TARGET``, ``UNSUPPORTED_LOCATOR``,
    and ``MISSING_RESOURCE_PATH``.
    """
    for key, hint in HINT_MAP.items():
        if key in reason:
            return hint
    if "UNKNOWN_TARGET" in reason:
        m = _QUOTED_NAME_RE.search(reason)
        target = m.group(1) if m else reason
        return f"add row with Object_ID='{target}' to Object_Repository sheet"
    if "UNSUPPORTED_LOCATOR" in reason:
        return (
            "OCR locators are v2 — change Locator_Type to IMAGE or leave as TODO stub"
        )
    if "MISSING_RESOURCE_PATH" in reason:
        m = _QUOTED_NAME_RE.search(reason)
        target = m.group(1) if m else ""
        return f"fill Resource_Path for '{target}' in Object_Repository sheet"
    return "check the Excel cell for this step"


def get_hints_for_version(version: str) -> dict[str, str]:
    """Return the full hint map for a given schema *version*.

    Base is ``HINT_MAP`` merged with any version-specific overrides
    from ``VERSION_HINTS``.
    """
    base = dict(HINT_MAP)
    base.update(VERSION_HINTS.get(version, {}))
    return base


def explain_column(col_name: str) -> str:
    """Return a short human-readable description of an Excel column name."""
    key = col_name.strip().lower().replace(" ", "_")
    return _COLUMN_DESCRIPTIONS.get(key, f"column '{col_name}'")
