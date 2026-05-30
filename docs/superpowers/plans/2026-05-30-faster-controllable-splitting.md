# Faster & Controllable Splitting Implementation Plan

**Goal:** Improve the split engine across three fronts — backend performance and control, UX/UI feedback, and a new selective-key feature — without changing existing default behavior.

This plan delivers four capabilities:

1. **Backend performance** — eliminate per-key disk reloads of the source/template workbook and reduce per-cell style work, so large jobs run dramatically faster.
2. **Cancellable splits (UX)** — a Cancel button that cooperatively stops a running split and keeps already-generated files.
3. **Selective key generation (new feature)** — load the unique key values up front and let the user generate only a chosen subset.
4. **Quiet logs by default (backend + UX)** — gate the noisy `Debug:` output behind an opt-in "Verbose logging" toggle.

**Architecture:** Keep the single-file `main.py` structure and the pure `mail_merge.py` module. All engine changes are additive keyword arguments on `split_excel_with_template` with backward-compatible defaults, so the positional call signature `(source_path, sheet_name, key_col, template_path, out_dir, header_rows, ...)` used by existing tests is preserved. The UI wires new controls to those arguments and to a cancellable `SplitWorker`, mirroring the existing `MailMergeWorker` cancel pattern.

**Tech Stack:** Python 3.10+, PySide6, qfluentwidgets, pandas, openpyxl, unittest, PyInstaller.

---

## File Structure

- Modify `main.py`:
  - Add a pure helper `read_key_values(...)` to list unique key values.
  - Extend `split_excel_with_template(...)` with `selected_keys`, `stop_requested`, and `verbose` keyword arguments; load workbook bytes once; cache template styles; delete source rows in contiguous runs.
  - Add `cancel()` and `stop_requested` wiring to `SplitWorker`.
  - Add a "Keys" panel (load + checkbox list + select/clear all), a footer Cancel button, a footer key-count label, and a "Verbose logging" checkbox.
  - Persist the verbose toggle in `QSettings`.
- Modify `tests/test_templating_modes.py`: tests for key filtering, cancellation, verbose gating, and `read_key_values`.
- Modify `tests/test_ui_smoke.py`: tests for the keys panel, cancel button, key-count label, and verbose checkbox.
- Modify `tests/test_settings.py`: assert the verbose toggle persists.
- Modify `README.md`: document selective keys, cancel, and verbose logging.
- Rebuild `dist/ExcelSplitter.exe` after verification.

---

## Design Notes

### Backward compatibility
- New `split_excel_with_template` parameters default to `selected_keys=None` (all keys), `stop_requested=None` (never stop), `verbose=False`. With these defaults the function behaves exactly as today, so all 89 existing tests continue to pass.
- `selected_keys` is a `set[str]`. Group keys are matched on `str(key_val)` so numeric/NaN keys compare consistently with what the UI displays.

### Performance
- **Read workbook once:** Before the group loop, read the template (template-file mode) or source (source-template mode) into an in-memory `bytes` buffer once. Per key, reload with `load_workbook(io.BytesIO(buffer))`. This removes K full-file disk reads.
- **Template styles:** In template-file mode, capture the template data-row style objects once per loaded workbook into a `{col_idx: cell}` reference map and assign those style objects to data cells instead of calling `.copy()` for every cell of every row. Style objects are workbook-bound, so capture happens per reloaded workbook (still far cheaper than per-cell copies).
- **Source row deletion:** In source-template mode, delete non-kept rows in contiguous runs (`ws.delete_rows(start, amount)`) from the bottom up, instead of one `delete_rows` call per row.

### Cancellation
- `stop_requested` is checked at the top of each group iteration. On cancel: emit a "Dibatalkan." status, stop the loop, and return the partial `split_results` already produced. Files generated before cancel are kept.
- `SplitWorker.cancel()` sets `self._cancel_requested = True`; the worker passes `stop_requested=lambda: self._cancel_requested`. This mirrors `MailMergeWorker`.

### Selective keys
- `read_key_values(source_path, sheet_name, key_col, source_header_rows)` returns the ordered, de-duplicated list of `str` key values (matching groupby order: first-seen, `sort=False`, `dropna=False` with `NaN` shown as the string `"nan"` consistent with pandas grouping display). It reads only the key column for speed.
- The UI "Keys" panel renders one checkbox per value (all checked by default). On Generate, if the panel has loaded keys and a strict subset is checked, pass that subset as `selected_keys`; if all are checked or none loaded, pass `None` (no filtering).
- Total/progress reflects the filtered count.

### Verbose logging
- All `status_cb("Debug: ...")` calls are wrapped by a local `debug(msg)` that only forwards when `verbose` is true. Non-debug status messages (progress, "Selesai.") are unchanged.

---

## Tasks

### Task 1: Engine — key filtering and cancellation

**Files:** `main.py`, `tests/test_templating_modes.py`

- [ ] **Step 1: Add failing tests**

Append to `tests/test_templating_modes.py` a `SplitControlTests` class:
- `test_selected_keys_limits_generated_files`: source with Dept A/B/C; call split with `selected_keys={"A", "C"}`; assert only `A.xlsx` and `C.xlsx` exist and `B.xlsx` does not; assert returned manifest has 2 results.
- `test_selected_keys_none_generates_all`: same source, `selected_keys=None`; assert all three files exist.
- `test_stop_requested_halts_after_first_key`: use a `stop_requested` callable that returns `True` after its first call; assert fewer files than total are produced and the function returns the partial manifest without raising.

Run: `.venv\Scripts\python.exe -m unittest tests.test_templating_modes` → expect failures (unknown kwargs).

- [ ] **Step 2: Implement in `split_excel_with_template`**

Add params `selected_keys: set | None = None`, `stop_requested=None` (default `lambda: False`). After `groups = df.groupby(...)`:
- Build `group_items = [(k, g) for k, g in groups if selected_keys is None or str(k) in selected_keys]`.
- Set `total = len(group_items)`; iterate `group_items`.
- At the top of each iteration: `if stop_requested(): status_cb("Dibatalkan."); break`.

Run the tests → expect pass.

- [ ] **Step 3: Add `read_key_values` helper + test**

Add `read_key_values(path, sheet_name, key_col, source_header_rows)` returning ordered unique `str` values from the key column (resolve by name or 1-based index, reuse `resolve_header_label`). Add `test_read_key_values_returns_ordered_unique_strings`. Run tests → pass.

### Task 2: Engine — performance (single read, style cache, run deletes)

**Files:** `main.py`, `tests/test_templating_modes.py`

- [ ] **Step 1: Add correctness tests**

Append tests asserting output is unchanged after the optimization:
- `test_source_template_preserves_only_matching_rows_after_optimization`: source with several rows across keys; assert each output workbook contains exactly that key's data rows and the header rows.
- `test_template_file_applies_template_styles_to_data_rows`: template row carries a number format / fill; assert a generated data cell inherits that number format.

Run → these should pass against current code (guards behavior), confirming the optimization keeps them green.

- [ ] **Step 2: Read workbook bytes once**

Add `import io`. Before the loop:
- template-file mode: `template_bytes = template_path.read_bytes()`; per key `wb = load_workbook(io.BytesIO(template_bytes))`.
- source-template mode: `source_bytes = source_path.read_bytes()`; per key `wb = load_workbook(io.BytesIO(source_bytes))`.

- [ ] **Step 3: Cache template row styles**

In template-file mode, after loading each workbook capture `style_src = {col_idx: ws.cell(row=template_row, column=col_idx) for col_idx in template_column_indices}` once, then for data rows assign `current_cell.font = style_src[col].font` etc. (reference assignment, no `.copy()` per cell).

- [ ] **Step 4: Delete source rows in contiguous runs**

Replace the per-row delete loop with a routine that walks rows from `ws.max_row` down to `start_row`, accumulates consecutive non-kept rows, and issues one `ws.delete_rows(run_start, run_len)` per run.

Run: `.venv\Scripts\python.exe -m unittest tests.test_templating_modes` → all pass.

### Task 3: Engine — verbose logging gate

**Files:** `main.py`, `tests/test_templating_modes.py`

- [ ] **Step 1: Add failing test**

`test_debug_messages_suppressed_unless_verbose`: capture `status_cb` messages with `verbose=False` and assert none start with `"Debug:"`; with `verbose=True` assert at least one does.

- [ ] **Step 2: Implement**

Add `verbose: bool = False`. Define `def debug(msg): (status_cb(msg) if verbose else None)`. Replace every `status_cb("Debug: ...")` with `debug("Debug: ...")`. Leave progress/"Selesai."/"Dibatalkan." on `status_cb`.

Run tests → pass.

### Task 4: Worker + footer — cancel and key count

**Files:** `main.py`, `tests/test_ui_smoke.py`

- [ ] **Step 1: Add failing UI tests**

- `test_split_app_has_cancel_button_hidden_until_running`: assert `btn_cancel_split` exists and is hidden initially.
- `test_split_app_has_key_count_label`: assert `lbl_key_count` exists.

- [ ] **Step 2: Make `SplitWorker` cancellable**

Add `self._cancel_requested = False` and `def cancel(self)`. In `run`, pass `stop_requested=lambda: self._cancel_requested` (and the other new params via `self.params.get(...)`).

- [ ] **Step 3: Add footer Cancel button + key-count label**

In `_build_actions_card`: add `btn_cancel_split` (hidden, calls `cancel_split`) and `lbl_key_count`. In `set_busy(True)` show Cancel and disable Generate; in `set_busy(False)` hide Cancel. Add `cancel_split()` calling `self.worker.cancel()` and logging "Membatalkan...". Handle cancelled finish gracefully in `_on_worker_finished`.

Run: `.venv\Scripts\python.exe -m unittest tests.test_ui_smoke` → pass.

### Task 5: UI — selective keys panel

**Files:** `main.py`, `tests/test_ui_smoke.py`

- [ ] **Step 1: Add failing UI tests**

- `test_split_app_has_keys_panel`: assert `keys_card`, `btn_load_keys`, `btn_select_all_keys`, `btn_clear_all_keys` exist.
- `test_selected_keys_subset_passed_to_worker`: load a known source, call `load_keys()`, uncheck one key, and assert `collect_selected_keys()` returns the expected subset (and returns `None` when all checked).

- [ ] **Step 2: Build the Keys panel**

Add `_build_keys_card()` between source and template cards: a "Load Keys" button, Select All / Clear All buttons, a count caption, and a bounded scroll area hosting one `CheckBox` per key. Store checkboxes in `self.key_checkboxes`.

- [ ] **Step 3: Implement handlers**

- `load_keys()`: validate source/sheet/key selected, call `read_key_values(...)`, render checkboxes (all checked), update count, log result; guard errors with `InfoBar`.
- `select_all_keys()` / `clear_all_keys()`.
- `collect_selected_keys()`: return `None` if no keys loaded or all checked; otherwise the set of checked key strings.
- Update `lbl_key_count` to show `"{checked} / {total} keys"`.

- [ ] **Step 4: Wire into `on_run_clicked`**

Add `selected_keys=self.collect_selected_keys()` and `verbose=self.chk_verbose_logging.isChecked()` to worker `params`. If a strict subset is selected, log it.

Run: `.venv\Scripts\python.exe -m unittest tests.test_ui_smoke tests.test_templating_modes` → pass.

### Task 6: UI — verbose toggle + persistence

**Files:** `main.py`, `tests/test_settings.py`, `tests/test_ui_smoke.py`

- [ ] **Step 1: Add failing tests**

- UI smoke: `test_split_app_has_verbose_toggle` asserts `chk_verbose_logging` exists and is unchecked by default.
- Settings: extend the persistence test to set `chk_verbose_logging` true, `save_settings()`, reload in a fresh `SplitApp`, and assert it stays checked.

- [ ] **Step 2: Implement**

Add `chk_verbose_logging` (CheckBox "Verbose logging", default off) to the Log card. Persist via key `"verbose_logging"` in `save_settings`/`load_settings`/`reset_settings`; connect `stateChanged` to `save_settings`.

Run: `.venv\Scripts\python.exe -m unittest tests.test_settings tests.test_ui_smoke` → pass.

### Task 7: Verification, docs, build, commit

**Files:** `README.md`, `dist/ExcelSplitter.exe`

- [ ] **Step 1: Full test suite** — `.venv\Scripts\python.exe -m unittest discover -s tests` → all pass.
- [ ] **Step 2: Syntax/whitespace** — `.venv\Scripts\python.exe -m py_compile main.py` and `git diff --check` → clean.
- [ ] **Step 3: README** — document selective key generation, cancel, and verbose logging under Usage and Advanced Features.
- [ ] **Step 4: Build** — `cmd /c build.cmd` → `dist\ExcelSplitter.exe` exists.
- [ ] **Step 5: Commit** on `feature/faster-controllable-splitting` with focused messages per task or one squashed feature commit.

---

## Risks & Mitigations

- **openpyxl `BytesIO` reload semantics:** loading from an in-memory buffer is equivalent to a path load; correctness tests in Task 2 guard output parity.
- **Style reference sharing:** assigning shared style objects within one workbook is supported by openpyxl and dedupes in the style table; guarded by the style-inheritance test.
- **Contiguous-run deletion off-by-one:** guarded by the row-preservation test; deletion proceeds bottom-up so indices stay valid.
- **Key string matching (NaN/numeric):** UI and engine both compare on `str(key_val)`; `read_key_values` returns the same string forms used for filenames/grouping.
