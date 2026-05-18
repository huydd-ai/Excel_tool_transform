# Excel Tool Transform

Convert structured Excel test plans (V2 format) into runnable [Airtest](https://airtest.netease.com/) `.air` scripts.

## Overview

Reads two sheets from an Excel file:
- **`Object_Repository`** — component name → image path, threshold, page, swipe coords
- **`Test_Execution`** — ordered steps with action keywords, targets, and data

Generates a `.air` directory (Airtest-compatible) containing a Python script for each test plan.

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

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--sheet` | `Test_Execution` | Sheet name to read steps from |
| `--output` | `./output` | Output root directory |
| `--plan` | Excel filename stem | Subfolder name under output |
| `--report` | off | Write `generation_report.txt` |

### Example

```bash
python excel_to_airtest.py Testcase.xlsx --output ./out --plan daily_mission --report
```

Output structure:
```
out/
└── daily_mission/
    └── Test_Execution.air/
        └── Test_Execution.py
out/
└── generation_report.txt
```

Run the generated script:
```bash
python -m airtest run out/daily_mission/Test_Execution.air
```

## Excel V2 Format

### Sheet `Object_Repository`

| component_name | page_id | image_path | threshold | position_hint |
|----------------|---------|------------|-----------|---------------|
| store_btn | home | ./assets/home/btn_store.png | 0.7 | |
| swipe_area | game | | | 100,200,300,400 |

`position_hint` format for SWIPE: `x1,y1,x2,y2` or `(x1,y1)-(x2,y2)`

### Sheet `Test_Execution`

| step_id | action | target | wait_after | data | notes |
|---------|--------|--------|------------|------|-------|
| 1 | CLICK | store_btn | 2 | | Open store |
| 2 | ASSERT_EXISTS | store_header | 3 | | |
| 3 | WAIT | | 1 | | |
| 4 | SWIPE | swipe_area | 1 | | Scroll down |
| 5 | INPUT_TEXT | | 0 | hello@example.com | Enter email |

### Supported Actions

| Action | Generated Code |
|--------|----------------|
| `CLICK` | `touch(Template(r"<path>", threshold=<t>))` |
| `ASSERT_EXISTS` | `assert_exists(Template(r"<path>", threshold=<t>), timeout=<wait>)` |
| `WAIT` | `sleep(<wait_after>)` |
| `SWIPE` | `swipe((x1,y1), (x2,y2))` — coords from `position_hint` |
| `INPUT_TEXT` | `text("<data value>")` |

### Error Handling

| Situation | Output |
|-----------|--------|
| Target not in Object_Repository | `# TODO: MISSING_ASSET 'name'` + logged |
| SWIPE with no valid coords | `# TODO: SWIPE_COORDS_MISSING` + logged |
| Unknown action | `# UNSUPPORTED_ACTION: 'X'` + logged |
| image_path not found on disk | WARNING printed + logged in report |

## Report (`generation_report.txt`)

When `--report` is passed, generates:
- Generated files list
- Skipped steps (step_id + reason)
- Missing assets (not in Object_Repository)
- Invalid assets (file not found on disk)

---

## Roadmap

### Approved Features (v1)
- [x] Excel V2 → Airtest `.air` script generation
- [x] Object_Repository + Test_Execution sheet parsing
- [x] CLICK / ASSERT_EXISTS / WAIT / SWIPE / INPUT_TEXT actions
- [x] Asset validation (disk check)
- [x] Generation report

### Future (pending approval)
- [ ] **Logic**: if/else conditions and loop support in steps
- [ ] **Action expansion**: screenshot capture (`snapshot`) and retry-on-fail wrapper
- [ ] **Asset verification**: pre-run check that all referenced images exist before generation

## License

MIT
