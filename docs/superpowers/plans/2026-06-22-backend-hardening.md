# Backend Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix 4 backend-quality issues found in senior-backend review: workbook resource leaks, generated-comment injection, fragile hint parsing, and module-level constant placement.

**Architecture:** All fixes are surgical — no new abstractions, no file splits. Each task touches 1–2 files and ships with a new or updated test.

**Tech Stack:** Python 3.10+, openpyxl 3.1+, pytest 8+

---

## Issues Fixed

| # | Severity | File | Issue |
|---|----------|------|-------|
| 1 | MEDIUM | `cli.py`, `rules_parser.py`, `agent_step1.py`, `templates/excel_template.py` | openpyxl workbooks never closed — file handle leak |
| 2 | LOW | `excel_to_airtest.py:234` | `step.expected` embedded raw in generated Python comment — newline can escape comment and inject code |
| 3 | MEDIUM | `hints.py:78,85` | `reason.split("'")[1]` is fragile — fails on reasons with zero or 2+ quoted names |
| 4 | LOW | `hints.py:107` | `_COLUMN_DESCRIPTIONS` dict redefined on every `explain_column()` call |

---

## Task 1: Close openpyxl workbooks — 4 files

**Files:**
- Modify: `cli.py:198–226`
- Modify: `rules_parser.py:95–149`
- Modify: `agent_step1.py:17–54`
- Modify: `templates/excel_template.py:40–52`
- Test: `tests/test_parsers.py` (no new test needed — existing tests exercise these paths; verify they still pass)

- [ ] **Step 1: Wrap cli.py workbook in try/finally**

  In `cli.py`, replace:
  ```python
  wb = openpyxl.load_workbook(args.excel_file, data_only=True)

  print(f"[parse] {args.objects_sheet}...")
  assets, obj_errors = parse_object_repository(wb, args.objects_sheet)
  ...
  suites, step_errors = parse_test_execution(wb, args.steps_sheet)
  ...
  if step_errors:
      sys.exit(1)
  ```

  With:
  ```python
  wb = openpyxl.load_workbook(args.excel_file, data_only=True)
  try:
      print(f"[parse] {args.objects_sheet}...")
      assets, obj_errors = parse_object_repository(wb, args.objects_sheet)
      for e in obj_errors:
          print(f"  [ERROR] {e}")
      if obj_errors:
          sys.exit(1)
      print(f"        {len(assets)} objects loaded")

      print(f"[parse] {args.actions_sheet}...")
      keywords, flows, act_errors = parse_action_logic(wb, args.actions_sheet)
      for e in act_errors:
          print(f"  [ERROR] {e}")
      print(f"        {len(keywords)} action keywords, {len(flows)} flow descriptions")

      print("[validate] Asset paths on disk...")
      val_issues = validate_assets(assets, base_dir)
      print(f"           {len(val_issues)} missing file(s)")

      print(f"[parse] {args.steps_sheet}...")
      suites, step_errors = parse_test_execution(wb, args.steps_sheet)
      for e in step_errors:
          print(f"  [ERROR] {e}")
      if step_errors:
          sys.exit(1)
      print(
          f"        {len(suites)} suite(s), {sum(len(v) for v in suites.values())} step(s)"
      )
  finally:
      wb.close()
  ```

  Note: `sys.exit()` raises `SystemExit`, which triggers `finally` — the workbook closes correctly even when parsing fails.

- [ ] **Step 2: Wrap rules_parser.py workbook in try/finally**

  In `rules_parser.py`, replace:
  ```python
  def read_rules_excel(path: str) -> RulesDoc:
      """Parse Manual_Testing_Guidelines_Rules.xlsx into a RulesDoc."""
      wb = openpyxl.load_workbook(path, data_only=True)
      doc = RulesDoc()

      for sheet_name in wb.sheetnames:
          ...

      return doc
  ```

  With:
  ```python
  def read_rules_excel(path: str) -> RulesDoc:
      """Parse Manual_Testing_Guidelines_Rules.xlsx into a RulesDoc."""
      wb = openpyxl.load_workbook(path, data_only=True)
      doc = RulesDoc()
      try:
          for sheet_name in wb.sheetnames:
              ws = wb[sheet_name]
              hdr = _header_index(ws)
              rows = list(ws.iter_rows(min_row=2, values_only=True))
              if not rows:
                  continue

              if sheet_name == "General Guidelines":
                  doc.guidelines = [
                      (r[0], r[1], r[2]) for r in rows if any(c is not None for c in r)
                  ]

              elif sheet_name == "Feature Rules":
                  for r in rows:
                      if not any(c is not None for c in r):
                          continue
                      doc.feature_rules.append(
                          FeatureRule(
                              feature=_col(r, hdr.get("feature"), 0),
                              logic_item=_col(r, hdr.get("logic item"), 1),
                              condition=_col(r, hdr.get("rule / condition"), 2),
                              expected=_col(r, hdr.get("expected behavior"), 3),
                          )
                      )

              elif sheet_name == "Edge Cases":
                  for r in rows:
                      if not any(c is not None for c in r):
                          continue
                      doc.edge_cases.append(
                          EdgeCase(
                              edge_id=_col(r, hdr.get("id"), 0),
                              scenario=_col(r, hdr.get("scenario"), 1),
                              condition=_col(r, hdr.get("condition"), 2),
                              recovery=_col(r, hdr.get("required handling (recovery)"), 3),
                          )
                      )

              elif sheet_name == "Release Checklist":
                  for r in rows:
                      if not any(c is not None for c in r):
                          continue
                      doc.checklist.append(
                          ChecklistItem(
                              check_id=_col(r, hdr.get("check id"), 0),
                              description=_col(r, hdr.get("verify item"), 1),
                          )
                      )
      finally:
          wb.close()
      return doc
  ```

- [ ] **Step 3: Wrap agent_step1.py workbook in try/finally**

  In `agent_step1.py`, replace:
  ```python
  try:
      wb = openpyxl.load_workbook(excel_path)
  except FileNotFoundError:
      print(f"ERROR: File not found: {excel_path}")
      sys.exit(1)
  except Exception as e:
      print(f"ERROR: Could not open workbook: {e}")
      sys.exit(1)

  print(f"File loaded: {excel_path}")
  print(f"Sheets found: {wb.sheetnames}")
  ...
  print(json.dumps(stats))
  ```

  With:
  ```python
  try:
      wb = openpyxl.load_workbook(excel_path)
  except FileNotFoundError:
      print(f"ERROR: File not found: {excel_path}")
      sys.exit(1)
  except Exception as e:
      print(f"ERROR: Could not open workbook: {e}")
      sys.exit(1)

  try:
      print(f"File loaded: {excel_path}")
      print(f"Sheets found: {wb.sheetnames}")

      missing = [s for s in REQUIRED_SHEETS if s not in wb.sheetnames]
      if missing:
          print(f"ERROR: Missing mandatory sheets: {missing}")
          sys.exit(1)

      stats = {}
      for sheet_name in REQUIRED_SHEETS:
          ws = wb[sheet_name]
          data_rows = 0
          for row in ws.iter_rows(min_row=2):
              if any(cell.value is not None for cell in row):
                  data_rows += 1
          stats[sheet_name] = {
              "max_row": ws.max_row,
              "data_rows": data_rows,
              "max_col": ws.max_column,
          }
          print(
              f"  Sheet '{sheet_name}': max_row={ws.max_row}, data_rows={data_rows}, max_col={ws.max_column}"
          )

      for sheet_name in REQUIRED_SHEETS:
          ws = wb[sheet_name]
          headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
          print(f"  Headers '{sheet_name}': {headers}")

      print("VERIFICATION: All mandatory sheets present.")
      print(json.dumps(stats))
  finally:
      wb.close()
  ```

- [ ] **Step 4: Close workbook in templates/excel_template.py**

  In `templates/excel_template.py`, replace:
  ```python
  wb.save(output_path)
  return os.path.abspath(output_path)
  ```

  With:
  ```python
  wb.save(output_path)
  wb.close()
  return os.path.abspath(output_path)
  ```

- [ ] **Step 5: Run tests**

  ```bash
  pytest tests/ -v
  ```
  Expected: 177 passed.

- [ ] **Step 6: Commit**

  ```bash
  git add cli.py rules_parser.py agent_step1.py templates/excel_template.py
  git commit -m "fix: close openpyxl workbooks to prevent file handle leaks"
  ```

---

## Task 2: Sanitize step.expected in generated comment

**Files:**
- Modify: `excel_to_airtest.py:233–236`
- Modify: `tests/test_handlers.py` (add test for newline injection)

- [ ] **Step 1: Write the failing test**

  Add to `tests/test_handlers.py`:
  ```python
  def test_step_label_strips_newlines_from_expected(make_wb):
      """Newline in expected must not escape the Python comment."""
      gen = AirtestGenerator()
      step = Step(
          suite_id="S1",
          step_no=1,
          action="SLEEP",
          expected='ok\nimport os; os.system("evil")',
      )
      label = gen.step_label(step)
      assert "\n" not in label
      assert "import os" not in label
  ```

  Run: `pytest tests/test_handlers.py::test_step_label_strips_newlines_from_expected -v`
  Expected: FAIL — `assert "\n" not in label` fails because current code embeds raw.

- [ ] **Step 2: Fix step_label in excel_to_airtest.py**

  Replace in `excel_to_airtest.py`:
  ```python
  def step_label(self, step: Step) -> str:
      return f"# step {step.step_no}" + (
          f" - {step.expected}" if step.expected else ""
      )
  ```

  With:
  ```python
  def step_label(self, step: Step) -> str:
      if step.expected:
          safe = step.expected.replace("\r", "").replace("\n", " ")
          return f"# step {step.step_no} - {safe}"
      return f"# step {step.step_no}"
  ```

- [ ] **Step 3: Run test to verify pass**

  ```bash
  pytest tests/test_handlers.py::test_step_label_strips_newlines_from_expected -v
  ```
  Expected: PASS.

- [ ] **Step 4: Run full suite**

  ```bash
  pytest tests/ -v
  ```
  Expected: 178 passed (1 new test added).

- [ ] **Step 5: Commit**

  ```bash
  git add excel_to_airtest.py tests/test_handlers.py
  git commit -m "fix: strip newlines from step.expected in generated comment"
  ```

---

## Task 3: Fix fragile hint parsing with regex

**Files:**
- Modify: `hints.py:1,78–86`
- Modify: `tests/test_diagnostics.py` (add edge-case tests)

- [ ] **Step 1: Write failing tests**

  Add to `tests/test_diagnostics.py`:
  ```python
  def test_hint_unknown_target_with_no_quotes():
      """reason with no quotes should not crash — returns safe fallback."""
      hint = diagnostic_hint("UNKNOWN_TARGET noquotes")
      assert "Object_Repository" in hint

  def test_hint_unknown_target_with_multiple_quoted_names():
      """First quoted name is extracted correctly when multiple exist."""
      hint = diagnostic_hint("UNKNOWN_TARGET 'btn_a' in 'suite_x'")
      assert "btn_a" in hint

  def test_hint_missing_resource_path_with_no_quotes():
      hint = diagnostic_hint("MISSING_RESOURCE_PATH for noquotes")
      assert "Resource_Path" in hint
  ```

  Run: `pytest tests/test_diagnostics.py -v -k "noquotes or multiple_quoted"`
  Expected: FAIL — current `split("'")[1]` raises `IndexError` when no quotes.

- [ ] **Step 2: Fix diagnostic_hint in hints.py**

  Add import at top of `hints.py` (after the existing `from __future__ import annotations`):
  ```python
  import re
  ```

  Add module-level constant after the `VERSION_HINTS` block:
  ```python
  _QUOTED_NAME_RE = re.compile(r"'([^']*)'")
  ```

  Replace in `diagnostic_hint()`:
  ```python
  if "UNKNOWN_TARGET" in reason:
      target = reason.split("'")[1] if "'" in reason else reason
      return f"add row with Object_ID='{target}' to Object_Repository sheet"
  ```
  With:
  ```python
  if "UNKNOWN_TARGET" in reason:
      m = _QUOTED_NAME_RE.search(reason)
      target = m.group(1) if m else reason
      return f"add row with Object_ID='{target}' to Object_Repository sheet"
  ```

  Replace:
  ```python
  if "MISSING_RESOURCE_PATH" in reason:
      target = reason.split("'")[1] if "'" in reason else ""
      return f"fill Resource_Path for '{target}' in Object_Repository sheet"
  ```
  With:
  ```python
  if "MISSING_RESOURCE_PATH" in reason:
      m = _QUOTED_NAME_RE.search(reason)
      target = m.group(1) if m else ""
      return f"fill Resource_Path for '{target}' in Object_Repository sheet"
  ```

- [ ] **Step 3: Run new tests**

  ```bash
  pytest tests/test_diagnostics.py -v
  ```
  Expected: all pass including the 3 new tests.

- [ ] **Step 4: Run full suite**

  ```bash
  pytest tests/ -v
  ```
  Expected: 181 passed.

- [ ] **Step 5: Commit**

  ```bash
  git add hints.py tests/test_diagnostics.py
  git commit -m "fix: use regex for hint parsing, replace fragile split-on-quote"
  ```

---

## Task 4: Hoist _COLUMN_DESCRIPTIONS to module level

**Files:**
- Modify: `hints.py:107–125`

- [ ] **Step 1: Move the dict to module level**

  In `hints.py`, move `_COLUMN_DESCRIPTIONS` OUT of `explain_column()` and place it after `_QUOTED_NAME_RE`.

  Remove from inside the function body:
  ```python
  def explain_column(col_name: str) -> str:
      _COLUMN_DESCRIPTIONS: dict[str, str] = {
          "object_id": "unique identifier for a UI element",
          ...
      }
      key = col_name.strip().lower().replace(" ", "_")
      return _COLUMN_DESCRIPTIONS.get(key, f"column '{col_name}'")
  ```

  Replace with:
  ```python
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


  def explain_column(col_name: str) -> str:
      """Return a short human-readable description of an Excel column name."""
      key = col_name.strip().lower().replace(" ", "_")
      return _COLUMN_DESCRIPTIONS.get(key, f"column '{col_name}'")
  ```

- [ ] **Step 2: Run tests**

  ```bash
  pytest tests/ -v
  ```
  Expected: 181 passed.

- [ ] **Step 3: Commit**

  ```bash
  git add hints.py
  git commit -m "refactor: hoist _COLUMN_DESCRIPTIONS to module level"
  ```

---

## Verification

After all tasks:

```bash
pytest tests/ -v
# Expected: 181 passed (4 new tests added)

python excel_to_airtest.py --list-projects
# Expected: no crash, prints registered projects

python excel_to_airtest.py --init-excel
# Expected: template.xlsx created, no resource leak

python agent_step1.py
# Expected: usage message (no crash), no open file handles
```
