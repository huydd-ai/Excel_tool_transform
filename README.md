# Excel Tool Transform

Convert a structured Excel test plan (the **AutomationRebase** schema) into runnable
[Airtest](https://airtest.netease.com/) `.air` scripts — **one `.air` per `Suite_ID`**.
The same workbook can describe many test cases across many projects; per-project
differences (e.g. the Android package id) come from CLI flags.

## Overview

Reads three sheets from an Excel file:

| Sheet | Purpose |
|-------|---------|
| `Object_Repository` | UI element registry: id, locator type, image path, threshold, timeout |
| `Action_Logic` | Action keyword registry + high-level flow descriptions (reference) |
| `Test_Execution` | Ordered steps grouped by `Suite_ID` |

For each distinct `Suite_ID` in `Test_Execution`, generates a `.air` directory
containing a runnable Airtest Python script.

## Requirements

- Python 3.10+
- `openpyxl`

```bash
pip install -r requirements.txt
```

## Usage

```bash
python excel_to_airtest.py <excel_file> [options]
```

### Quick start for a new project

```bash
# 1. Get a blank Excel template to fill in
python excel_to_airtest.py --init-excel
# → template.xlsx created in current directory

# 2. Scaffold new project config (developer does once)
python excel_to_airtest.py --init-project mygame
# → projects/mygame.py created — edit DEFAULT_APP_PACKAGE and IMPORTS

# 3. Verify project is discovered
python excel_to_airtest.py --list-projects
# → mygame   projects/mygame.py   com.example.mygame

# 4. Generate test scripts
python excel_to_airtest.py MyPlan.xlsx --project mygame --report
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--project <name>`      | auto (if 1 project) | Select project from `projects/` directory |
| `--list-projects`       | —                   | Print all discovered projects and exit |
| `--init-project <name>` | —                   | Scaffold `projects/<name>.py` and exit |
| `--init-excel`          | —                   | Generate `template.xlsx` and exit |
| `--output`              | `./output`          | Output root directory |
| `--plan`                | excel filename stem | Subfolder under output |
| `--app-package`         | project default     | App package id for START_APP / STOP_APP |
| `--report`              | off                 | Write `generation_report.txt` |

### Example

```bash
# Pixon project (pre-configured in projects/pixon.py)
python excel_to_airtest.py AutomationRebase.xlsx --project pixon --report

# Base Airtest (no project selected)
python excel_to_airtest.py Plan.xlsx --app-package com.my.game
```

## Excel Schema

Headers are matched **case-insensitively**.

### Sheet `Object_Repository`

| Object_ID | Locator_Type | Resource_Path | Smart_Threshold | Timeout |
|-----------|--------------|---------------|-----------------|---------|
| `btn_play` | `IMAGE` | `./assets/home/btn_play.png` | `0.85` | `3` |
| `heart_count` | `OCR` | `NONE` | `0.7` | `10` |

- `Locator_Type`: `IMAGE` (template-matching) or `OCR` (text). OCR rows currently emit
  TODO stubs in the generated script — implement per project as needed.
- `Resource_Path` may be relative (resolved against the workbook directory) or absolute.
- `Timeout` is used by `WAIT_FOR` / `ASSERT_VISIBLE` for that target.

### Sheet `Action_Logic` *(optional)*

| Logic_ID | Action_Name | Machine_Command | Target_Page |
|----------|-------------|-----------------|-------------|
| `ACT_001` | `START_APP`      | `START_APP(Logic_ID)`          | |
| `ACT_002` | `CLICK`          | `CLICK(Object_ID)`             | |
| `FLOW_001` | `Display Splash` | *(narrative)*                 | `SplashScreen` |

- `ACT_*` rows declare primitive action keywords (purely informational; the registry
  in code is the source of truth for what gets generated).
- `FLOW_*` rows are high-level flow descriptions — captured in the report for
  reference, **not** generated. Composite-flow generation is on the roadmap.

### Sheet `Test_Execution`

| Suite_ID | Step | Action_Keyword | Target_ID | Params | Expected_Result |
|----------|------|----------------|-----------|--------|-----------------|
| `TC_MISSION_1_CHECKHEART` | 1 | `CLICK`          | `btn_play`     |  | Open game |
| `TC_MISSION_1_CHECKHEART` | 2 | `START_APP`      |                | `{"heart": "5", "coin": "10000", "level": "3"}` | Cold start with config |
| `TC_MISSION_1_CHECKHEART` | 3 | `WAIT_FOR`       | `btn_main_home`|  | Home loaded |
| `TC_MISSION_1_CHECKHEART` | 4 | `ASSERT_VISIBLE` | `heart_count`  |  | Heart visible |

- All steps sharing a `Suite_ID` go into one generated `.air` script.
- `Params` is free-form text; for `START_APP` it should be a JSON object.

### Supported Actions

| Action | Target_ID | Params | Generated Code |
|--------|-----------|--------|----------------|
| `TAP` | IMAGE object | — | `touch(Template(...))` |
| `TOUCH` | IMAGE object | — | `touch(Template(...))` (same as TAP) |
| `CLICK` | IMAGE object | — | `touch(Template(...))` *(deprecated — use TAP)* |
| `WAIT_FOR` | IMAGE object | — | `wait(Template(...), timeout=t)` |
| `ASSERT_VISIBLE` | IMAGE object | — | `assert_exists(Template(...), timeout=t)` |
| `START_APP` | — | JSON config | `start_app("pkg")` |
| `STOP_APP` | — | — | `stop_app("pkg")` |
| `INPUT_TEXT` | — | text string | `text("value")` |
| `READ_TEXT` | — | — | `# TODO stub (v2)` |
| `SWIPE` | — | `{"from":[x1,y1],"to":[x2,y2],"duration":0.5}` | `swipe((x1,y1),(x2,y2),duration=0.5)` |
| `SCROLL` | — | `{"direction":"up"}` or `up`/`down`/`left`/`right` | `swipe` with screen-fraction coords |
| `LONG_PRESS` | IMAGE object | — | `long_click(Template(...))` |
| `SLEEP` | — | `2.5` or `{"seconds":2.5}` | `sleep(2.5)` |
| `BACK` | — | — | `keyevent("BACK")` |
| `HOME` | — | — | `keyevent("HOME")` |
| `SNAPSHOT` | — | filename | `snapshot(filename="name.png")` |

### Error Handling

| Situation | Output |
|-----------|--------|
| `Target_ID` empty | `# TODO: MISSING_TARGET` + logged |
| `Target_ID` not in Object_Repository | `# TODO: UNKNOWN_TARGET '<id>'` + logged |
| Target is `OCR` (not yet supported by handler) | `# TODO: UNSUPPORTED_LOCATOR 'OCR' for '<id>'` + logged |
| `Resource_Path` is empty or `NONE` for IMAGE target | `# TODO: MISSING_RESOURCE_PATH for '<id>'` + logged |
| `START_APP` with no `--app-package` | `# TODO: START_APP_NEEDS_PACKAGE` + logged |
| `START_APP` Params not valid JSON | `# TODO: INVALID_PARAMS_JSON` + logged |
| Unknown `Action_Keyword` | `# UNSUPPORTED_ACTION` + logged |
| `Resource_Path` not on disk | WARNING printed + logged in report |

All Excel string values are passed through Python `repr()` before being embedded in
the generated source, so crafted cell content cannot inject code.

## Report (`generation_report.txt`)

When `--report` is passed, contains:

- List of generated `.air` files (per suite)
- Generation issues with **Excel row number** + suite + step + reason
- Asset validation issues (file-not-found)
- `FLOW_*` entries from `Action_Logic` for reference

## Testing

A `pytest` suite covers helpers, the three parsers, every action handler and
`_resolve_image` branch, the generator dispatch, and an end-to-end CLI run that
asserts generated scripts compile (`py_compile`).

```bash
pip install -r requirements-dev.txt
python -m pytest tests/ -v
```

The injection guard is enforced by `ast`-level checks: tests assert the generated
expression parses to exactly one `Call(...)` with the crafted payload appearing
only as a string-literal `Constant`, not as a syntactic call.

## Extending the tool — one file per project

Project-specific differences (framework imports, page objects, action overrides) live in a **single file under `projects/`**. Files are auto-discovered at startup. Switch projects with `--project <name>`.

### Scaffold a new project

```bash
python excel_to_airtest.py --init-project mygame
# → creates projects/mygame.py with commented stubs
```

Edit the 2–3 lines marked `← change this`. Run `--list-projects` to verify.

### What you can override

Same hooks as before — `IMPORTS`, `MODULE_PROLOGUE`, `DEFAULT_APP_PACKAGE`, `wrap_main_body()`, `@action("NAME")` handlers, `add_arguments()`.

### Example: `projects/pixon.py`

The repo ships one concrete subclass targeting the `pixon` framework. It overrides `TAP`/`TOUCH`/`WAIT_FOR`/`ASSERT_VISIBLE`/`START_APP`/`wrap_main_body` to route through `wrapper.*` and `TestCaseTimer`.

```bash
python excel_to_airtest.py AutomationRebase.xlsx --project pixon --report
```

### Adding a new project

```python
# projects/game2.py
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_to_airtest import AirtestGenerator, action, register_project

@register_project("game2")
class Game2Generator(AirtestGenerator):
    IMPORTS             = "from my_framework import driver, wait, tap"
    DEFAULT_APP_PACKAGE = "com.my.app"

    @action("TAP")
    def handle_tap(self, step, ctx):
        asset, lines, issue = self._resolve_image(step, ctx)
        if asset is None: return lines, issue
        return [f"tap({asset.resource_path!r})"], None

if __name__ == "__main__":
    Game2Generator.main()
```

### generators/ (deprecated)

The old `generators/pixon_generator.py` is now a backward-compatibility shim that imports from `projects/pixon.py`. New projects go in `projects/`.

### Schema column changes

Required headers live at the top of the parser section (`_REQUIRED_*_HEADERS`).
Headers are normalised lowercase, so renaming a workbook column is a one-line change.

## Roadmap

### Approved (v1)
- [x] Excel (AutomationRebase schema) -> Airtest `.air` script generation
- [x] Multi-suite output (one `.air` per `Suite_ID`)
- [x] `Object_Repository` with `Locator_Type` and per-object `Timeout`
- [x] `Action_Logic` registry + flow reference
- [x] `CLICK` / `WAIT_FOR` / `ASSERT_VISIBLE` / `START_APP` / `INPUT_TEXT` / `READ_TEXT` (stub)
- [x] Asset disk validation
- [x] Generation report with Excel row context
- [x] Pytest suite (helpers, parsers, handlers, generator, CLI end-to-end, injection, inheritance)
- [x] Class-based architecture — one subclass file per project (`generators/pixon_generator.py` ships as the canonical example)

### Future (pending approval)
- [ ] OCR locator support (real implementation, not stub)
- [ ] `READ_TEXT` real implementation (assign to variable, use in later assertions)
- [ ] `FLOW_*` composite flow expansion from `Action_Logic`
- [ ] If/else / loop control in step sequence
- [ ] Pre-run asset verification (fail fast before script generation)
- [ ] Screenshot capture (`snapshot`) and retry-on-fail wrapper

## License

MIT
