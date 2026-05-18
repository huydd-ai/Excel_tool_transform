# Excel Tool Transform

Convert structured Excel test plans (V2 format) into runnable [Airtest](https://airtest.netease.com/) `.air` scripts.

## Overview

Reads two sheets from an Excel file:
- **`Object_Repository`** — component name → image path, threshold, page
- **`Test_Execution`** — ordered steps with action keywords and targets

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

### Sheet `Test_Execution`

| step_id | action | target | wait_after | notes |
|---------|--------|--------|------------|-------|
| 1 | CLICK | store_btn | 2 | Open store |
| 2 | ASSERT_EXISTS | store_header | 3 | |
| 3 | WAIT | | 1 | |

### Supported Actions

| Action | Generated Code |
|--------|----------------|
| `CLICK` | `touch(Template(r"<path>", threshold=<t>))` |
| `ASSERT_EXISTS` | `assert_exists(Template(r"<path>", threshold=<t>), timeout=<wait>)` |
| `WAIT` | `sleep(<wait_after>)` |
| `SWIPE` | `swipe(start, end)  # position_hint: ...` |
| `INPUT_TEXT` | `text("<notes value>")` |

Missing asset → `# TODO: MISSING_ASSET 'name'` inserted inline and logged to report.

## License

MIT
