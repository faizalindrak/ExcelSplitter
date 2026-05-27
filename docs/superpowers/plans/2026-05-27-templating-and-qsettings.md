# Templating Modes and QSettings Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [x]`) syntax for tracking.

**Goal:** Add selectable template modes, manual column mapping, selected-worksheet source templating, and automatic Qt `QSettings` persistence.

**Architecture:** Keep the existing single-file application structure, but add small pure helpers in `main.py` for header extraction, normalization, auto-mapping, and mapping validation. Extend `split_excel_with_template` with explicit `template_mode` and `column_mapping` arguments, then wire the PySide6 UI to those helpers through a mapping card and QSettings-backed persistence.

**Tech Stack:** Python, pandas, openpyxl, PySide6, qfluentwidgets, unittest, PyInstaller.

---

### Task 1: Mapping Helper Tests and Functions

**Files:**
- Modify: `main.py`
- Create: `tests/test_templating_modes.py`

- [x] **Step 1: Write failing tests for header normalization, auto-mapping, and validation**

```python
from main import auto_map_columns, normalize_header, validate_column_mapping


def test_normalize_header_ignores_case_spaces_and_punctuation():
    assert normalize_header(" Employee  Name!") == "employeename"


def test_auto_map_columns_matches_normalized_headers():
    mapping = auto_map_columns(
        ["Employee Name", "Dept"],
        ["employee_name", "Dept", "Amount"],
    )

    assert mapping == {"Employee Name": "employee_name", "Dept": "Dept"}


def test_validate_column_mapping_reports_unmapped_template_columns():
    missing = validate_column_mapping(
        ["Employee Name", "Dept"],
        {"Employee Name": "Name", "Dept": None},
    )

    assert missing == ["Dept"]
```

- [x] **Step 2: Run tests and verify they fail because helpers do not exist**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_templating_modes`

Expected: import failure for `auto_map_columns`, `normalize_header`, or `validate_column_mapping`.

- [x] **Step 3: Add mapping helper functions in `main.py` above split logic**

```python
def normalize_header(value) -> str:
    value = "" if value is None else str(value)
    return re.sub(r"[^a-z0-9]+", "", value.strip().lower())


def auto_map_columns(template_headers, source_headers):
    source_by_key = {}
    for source in source_headers:
        key = normalize_header(source)
        if key and key not in source_by_key:
            source_by_key[key] = source

    return {
        template: source_by_key.get(normalize_header(template))
        for template in template_headers
    }


def validate_column_mapping(template_headers, mapping):
    mapping = mapping or {}
    return [
        template
        for template in template_headers
        if not mapping.get(template)
    ]
```

- [x] **Step 4: Run tests and verify they pass**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_templating_modes`

Expected: all tests pass.

### Task 2: Template Header Reading and Use Template File Splitting

**Files:**
- Modify: `main.py`
- Modify: `tests/test_templating_modes.py`

- [x] **Step 1: Add failing tests for manual mapping with different headers and missing mapping errors**

Create source workbook with headers `Name`, `Dept`, and template workbook with headers `Worker`, `Team`. Call `split_excel_with_template(..., template_mode="template_file", column_mapping={"Worker": "Name", "Team": "Dept"})` and assert the `A.xlsx` output writes `Alice` and `A` under the template headers. Add a second test with `{"Worker": "Name"}` and assert `ValueError` mentions `Team`.

- [x] **Step 2: Run tests and verify failure**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_templating_modes`

Expected: failure because `template_mode` and `column_mapping` are not supported yet.

- [x] **Step 3: Add header readers**

```python
def read_excel_headers(path: Path, sheet_name: str, header_rows: int) -> list[str]:
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_rows - 1, nrows=0)
    return [str(col) for col in df.columns]


def read_template_headers(path: Path, header_rows: int) -> tuple[list[str], int]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        headers, col_idx, empty_streak = [], 1, 0
        first_col = None
        while col_idx <= 500 and empty_streak < 5:
            value = ws.cell(row=header_rows, column=col_idx).value
            if value is None or str(value).strip() == "":
                empty_streak += 1
            else:
                if first_col is None:
                    first_col = col_idx
                headers.append(str(value).strip())
                empty_streak = 0
            col_idx += 1
        return headers, first_col or 1
    finally:
        wb.close()
```

- [x] **Step 4: Extend splitter signature and Use Template File branch**

Add `template_mode="template_file"` and `column_mapping=None` to `split_excel_with_template`. For `template_file`, read template headers, auto-map if `column_mapping` is missing, validate missing mappings, and construct the output DataFrame in template header order before grouping and writing.

- [x] **Step 5: Run tests and verify they pass**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_templating_modes`

Expected: all tests pass.

### Task 3: Use Source as Template Splitting

**Files:**
- Modify: `main.py`
- Modify: `tests/test_templating_modes.py`

- [x] **Step 1: Add failing source-template test**

Create a source workbook with sheets `Data` and `Other`, split by `Dept`, call `split_excel_with_template(..., template_mode="source_template", pdf_engine="none")`, and assert `A.xlsx` contains only `Data`, preserves the header row, and contains only rows whose key is `A`.

- [x] **Step 2: Run tests and verify failure**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_templating_modes`

Expected: failure because `source_template` mode is unsupported.

- [x] **Step 3: Implement source-template branch**

For each group, load the source workbook, keep only `sheet_name`, delete all data rows from bottom to `header_rows + 1` whose original row number is not in the group index, set print titles/area, save to the usual output file name, and export PDF if requested.

- [x] **Step 4: Run tests and verify they pass**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_templating_modes`

Expected: all templating tests pass.

### Task 4: UI Controls and Mapping Card

**Files:**
- Modify: `main.py`
- Modify: `tests/test_ui_smoke.py`

- [x] **Step 1: Add failing UI smoke assertions**

Assert `SplitApp` has `cmb_template_mode`, `mapping_card`, and `btn_auto_map`, and that the template mode combo contains `Use Template File` and `Use Source as Template`.

- [x] **Step 2: Run UI smoke test and verify failure**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke`

Expected: attribute assertion failure.

- [x] **Step 3: Add template mode combo and mapping card**

Add a mode row to the Template card, add a `Column Mapping` accordion card with Auto Map button and dynamic `Template Column -> Source Column` combo rows, and hide/disable mapping for Use Source as Template mode.

- [x] **Step 4: Pass template mode and mapping to worker**

In `on_run_clicked`, skip template path validation for `source_template`, collect mapping from combo rows for `template_file`, block generation with an InfoBar if mappings are missing, and include `template_mode` and `column_mapping` in worker params.

- [x] **Step 5: Run UI smoke and templating tests**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke tests.test_templating_modes`

Expected: all tests pass.

### Task 5: QSettings Persistence

**Files:**
- Modify: `main.py`
- Create or modify: `tests/test_settings.py`

- [x] **Step 1: Write failing QSettings test**

Instantiate `SplitApp(settings=QSettings(temp_file, QSettings.IniFormat))`, set source path, output path, template mode, PDF engine, prefix, suffix, and a mapping. Call `save_settings()`. Instantiate a second app with the same settings and assert the values load without `.ini` files.

- [x] **Step 2: Run settings test and verify failure**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_settings`

Expected: failure because `SplitApp` does not accept injected settings and still exposes `.ini` save/load flow.

- [x] **Step 3: Implement QSettings**

Import `QSettings`, accept optional `settings` in `SplitApp.__init__`, replace Save/Load `.ini` buttons with Reset Settings, load settings after UI construction, save settings on generation and field changes, and serialize mapping as JSON.

- [x] **Step 4: Run settings and UI tests**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_settings tests.test_ui_smoke`

Expected: all tests pass.

### Task 6: Verification and Build

**Files:**
- Modify: `README.md`
- Modify: `dist/ExcelSplitter.exe`

- [x] **Step 1: Update README**

Document the two template options, manual mapping behavior, source worksheet template behavior, and automatic QSettings persistence.

- [x] **Step 2: Run full tests**

Run: `rtk .venv\Scripts\python.exe -m unittest discover`

Expected: all tests pass.

- [x] **Step 3: Run syntax and whitespace checks**

Run: `rtk python -m py_compile main.py`

Expected: exit code 0.

Run: `rtk git diff --check`

Expected: exit code 0.

- [x] **Step 4: Build executable**

Run: `rtk cmd /c build.cmd`

Expected: exit code 0 and `dist\ExcelSplitter.exe` exists.

- [x] **Step 5: Final status**

Run: `rtk git status --short`

Expected: only intentional implementation changes are listed.
