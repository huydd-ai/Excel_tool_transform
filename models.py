"""
excel_tool/models.py — Data models for the Excel-to-Airtest codegen engine.

Defines the core data structures and constants shared across parsing,
validation, hinting, and code generation. All models are @dataclass
for immutability, readability, and easy construction.

Constants re-exported here:
    LOCATOR_IMAGE, LOCATOR_OCR  — locator strategy strings
    _SCROLL_PRESETS             — coordinate presets for SWIPE/SCROLL
"""

from __future__ import annotations

from dataclasses import dataclass

# ── Type Aliases ──────────────────────────────────────────────────────────────

HeaderMap = dict[str, int]
"""Mapping from normalised column name → zero-based column index."""

# ── Locator Constants ─────────────────────────────────────────────────────────

LOCATOR_IMAGE = "IMAGE"
"""Locator strategy: template-based image matching."""

LOCATOR_OCR = "OCR"
"""Locator strategy: optical character recognition (stub / project-implemented)."""

# ── Scroll Presets ────────────────────────────────────────────────────────────

_SCROLL_PRESETS: dict[str, tuple[str, str]] = {
    "up": ("(int(w*0.5), int(h*0.7))", "(int(w*0.5), int(h*0.3))"),
    "down": ("(int(w*0.5), int(h*0.3))", "(int(w*0.5), int(h*0.7))"),
    "left": ("(int(w*0.7), int(h*0.5))", "(int(w*0.3), int(h*0.5))"),
    "right": ("(int(w*0.3), int(h*0.5))", "(int(w*0.7), int(h*0.5))"),
}
"""Screen-fraction coordinate pairs for SCROLL actions.

Each entry maps a direction name → (from_coord_expr, to_coord_expr).
w and h are substituted at generation time with the screen dimensions.
"""

# ── Error Hierarchy ───────────────────────────────────────────────────────────


class AirtestError(Exception):
    """Base exception for all excel_tool codegen errors."""


class AirtestParseError(AirtestError):
    """Raised when an Excel workbook cannot be parsed (missing sheet, bad format)."""


# ── Existing Data Models (preserved for backward compatibility) ───────────────


@dataclass
class Asset:
    """A UI element entry from the Object_Repository sheet.

    Maps an Object_ID to its locator strategy and matching parameters.
    """

    object_id: str
    locator_type: str = LOCATOR_IMAGE
    resource_path: str = ""
    threshold: float = 0.8
    timeout: float = 5.0


@dataclass
class Step:
    """A single parsed row from the Test_Execution sheet.

    Represents one action step within a test suite (Suite_ID).
    """

    suite_id: str
    step_no: int
    action: str
    excel_row: int = 0
    target: str = ""
    params: str = ""
    expected: str = ""


@dataclass
class FlowDoc:
    """A reference row from the Action_Logic sheet.

    ACT_* rows document primitive action keywords.
    FLOW_* rows describe composite flows (narrative, not generated).
    """

    logic_id: str
    action_name: str
    command: str
    target_page: str


@dataclass
class ValidationIssue:
    """Describes a validation problem found during asset validation.

    component — name of the component / suite being validated
    path      — filesystem path involved
    kind      — category of the issue (FILE_NOT_FOUND, etc.)
    """

    component: str
    path: str
    kind: str = "FILE_NOT_FOUND"


@dataclass
class GenerationIssue:
    """Describes an issue encountered during code generation.

    tied back to the originating Excel row for traceability.
    """

    suite_id: str
    step_no: int
    excel_row: int
    reason: str


@dataclass
class GenCtx:
    """Context object passed to action handlers during code generation.

    Carries the resolved asset dictionary and the target app-package string.
    """

    assets: dict[str, Asset]
    app_package: str


# ── New Canonical Names ───────────────────────────────────────────────────────


@dataclass
class AirtestStep:
    """A single generated line of Airtest source code.

    Captures the code string together with its indentation level and an
    optional trailing comment for readability in the generated .air file.
    """

    code: str
    indent: int = 0
    comment: str | None = None


# Backward-compatible aliases
ParsedAction = Step
"""Alias: a parsed Excel row is a Step."""

GeneratorConfig = GenCtx
"""Alias: generator configuration / context is a GenCtx."""
