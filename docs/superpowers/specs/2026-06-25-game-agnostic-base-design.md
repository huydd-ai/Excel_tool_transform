# Game-Agnostic Puzzle-Test Base — Design Spec

**Date:** 2026-06-25
**Branch:** `feature/game-agnostic-base`
**Status:** Approved (design); pending implementation plan

## Goal

Turn `excel_tool` into a clean, game-agnostic **base** that generates Airtest `.air` test
scripts for *any* puzzle game from a structured Excel test plan. Remove every reference to the
specific pixon / "Screw Land" game and its hardcoded test plan, de-brand the tooling, and ship a
single generic example so a new user has a working reference.

## Background (verified)

The engine is **already game-agnostic**. The six core files (`excel_to_airtest.py`, `models.py`,
`parsers.py`, `registry.py`, `cli.py`, `writer.py`, plus `hints.py`) contain zero game names. A
"project" is an `AirtestGenerator` subclass registered via `@register_project(name)` and
auto-discovered from `projects/*.py`. All game coupling lives in exactly two shippable artifacts
plus some brand strings:

- `projects/pixon.py` — the concrete pixon plugin (imports `pixon.*`, page objects, app package
  `com.woodpuzzle.pin3d`).
- `rules_to_suites.py` + `rules_parser.py` — a second pipeline that hardcodes one game's entire
  test plan (`FEATURE_RULE_MAP` / `EDGE_CASE_MAP` / `CHECKLIST_MAP` / `_BASE_ASSETS` for
  HeartSystem / RoyalPass / LavaQuest).

**Import graph verified:** no core file imports `rules_to_suites` or `rules_parser`; the only
importers are the two rules test files and `tests/test_cli_errors.py::TestRulesCliErrors`.
`pyproject.toml` exposes only one entry point (`excel-tool = excel_to_airtest:main`) — no rules
script is orphaned by deletion.

Scope is **`excel_tool/` only**. The sibling `pixon/` framework package, `Test/` generated
outputs, top-level `D:\AutoRebase\pyproject.toml`/`README.md` (the `autorebase`/`pixon` package),
and `dagster/` (a separate gitignored repo) are **out of scope and untouched**.

## Decisions

| # | Decision |
|---|----------|
| 1 | Scope = the `excel_tool` generator only. |
| 2 | **Delete** the pixon plugin and the hardcoded rules pipeline entirely (not relocate, not make data-driven). |
| 3 | Refactor in place on branch `feature/game-agnostic-base`. |
| 4 | Ship a generic puzzle example plugin + a neutral on-demand sample workbook (via `--init-excel`). |
| 5 | `game-flow-analyst.md` agent: de-path only (keep the agent, remove the hardcoded absolute path). |
| 6 | Include the `hints.py` dead-key prune (verified safe). |

## Work

### Step 0 — Safety tag (before any deletion)

```
git tag pre-decouple-base
```

Preserves the freshly-hardened rules pipeline (generic prose-rules→suites engine) so it can be
resurrected if a data-driven version is wanted later. Deletion discards a *capability*, not just
game code; the tag is the one-command undo.

### Delete

- `projects/pixon.py`
- `rules_to_suites.py`
- `rules_parser.py`
- `tests/test_rules_to_suites.py`
- `tests/test_rules_autoassets.py`
- `generators/` (vestigial empty package — `generators/__init__.py`, referenced nowhere)

### Prune (keep the generic remainder)

- `tests/test_cli_errors.py` — remove the `TestRulesCliErrors` class, the `RULES = …rules_to_suites.py`
  constant (line 9), and the `and rules_to_suites.py` phrase in the module docstring. Keep
  `TestMainCliErrors`.
- `hints.py` — remove the now-dead rules-only keys `SERVER_VALIDATION`, `MANUAL_TIMER_CHECK`,
  `MANUAL_CHECK`. Verified: referenced only in `test_rules_to_suites.py` (deleted), nowhere in
  surviving source or tests.

### De-brand (string edits, all under `excel_tool/`)

- `excel_to_airtest.py:46` `DESCRIPTION` — "AutomationRebase Excel" → "a structured test-plan Excel".
- `cli.py:103` `build_parser` default `description` — same neutral wording.
- `pyproject.toml:8` `description` — neutral (e.g. "Convert a structured test-plan Excel workbook to
  Airtest .air scripts").
- `README.md` — rewrite: remove the pixon section and `--project pixon` examples; remove the stale
  `generators/pixon_generator.py` references (file no longer exists); rename "AutomationRebase schema"
  to neutral wording; document the `--init-project` / `--init-excel` / `--project example_puzzle`
  workflow.
- `.claude/agents/game-flow-analyst.md` — replace the only hardcoded absolute path
  (`D:\AutoRebase\excel_tool\.claude\agent-memory\…`, line ~247) with a repo-relative path; change the
  `AutomationRebase.xlsx` default (lines 3, 9, 17) to a neutral placeholder. Agent kept.

### Add — starter example

- `projects/example_puzzle.py` — a small, heavily-commented generic `AirtestGenerator` subclass
  registered as `example_puzzle`. Neutral app package `com.example.puzzle`, bare-Airtest `IMPORTS`,
  and one overridden handler shown as the extension pattern. Placed in `projects/` (not `examples/`)
  so it is auto-discovered: `--list-projects` shows it and `--project example_puzzle` runs end-to-end
  (it is also the regression anchor). A user can delete it.
- `templates/excel_template.py` — replace the two game-flavored example cells with neutral
  puzzle-generic equivalents: the `heart_count` OCR row (line 77) → a neutral OCR example
  (e.g. `score_text`); the START_APP params `{"heart":"5","level":"1"}` (line 114) → neutral config
  (e.g. `{}` or a non-game key). Everything else in the template is already generic.

## Acceptance criteria

1. `pytest tests/` is green at the new (lower) test count — rules tests removed, `TestRulesCliErrors`
   pruned. The implementation plan pins the exact count; the spec requires "all green, no skips."
2. `grep -rIE 'pixon|HeartSystem|RoyalPass|LavaQuest|AutomationRebase|D:\\AutoRebase' excel_tool`
   (excluding `graphify-out/`, `.air` outputs, and `docs/`) returns **zero** matches in source.
3. `python -m excel_tool --list-projects` → lists `example_puzzle`, no `pixon`.
4. `python -m excel_tool --init-excel` → writes a `template.xlsx` containing no game terms.
5. `python -m excel_tool --init-project foo` → scaffolds `projects/foo.py` cleanly.
6. End-to-end: generate from the `--init-excel` sample workbook with `--project example_puzzle` →
   produces a valid `.air` script and exits 0.

## Out of scope (noted, not done)

- Removing the module-level backward-compat shims in `excel_to_airtest.py` (`_default_generator`,
  `_discover_projects`, `_resolve_image`, `_HANDLERS`) — these existed for pixon-era external
  subclasses; removing them is a separate breaking change. Future cleanup.
- `agent_step1.py` — orphan core-schema verifier (not CLI-wired). Generic; kept as-is.
- Any change to `pixon/`, `Test/`, `dagster/`, or the top-level `autorebase` package.
