# Compact Workflow UI Redesign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Redesign the PySide6 UI into a compact workflow dashboard with bounded input controls, a left workflow rail, compact panels, and a persistent action footer.

**Architecture:** Keep the existing single-file `main.py` structure and preserve the current split worker, template mode, mapping, and QSettings behavior. Replace the stacked accordion layout with a two-zone dashboard layout: a narrow workflow rail plus compact panels, with run actions moved into a footer. Add focused UI tests that verify construction, compact widths, and mode-dependent visibility.

**Tech Stack:** Python, PySide6, PySide6-Fluent-Widgets/qfluentwidgets, unittest, PyInstaller.

---

## File Structure

- Modify `main.py`: update imports, add compact UI constants/helpers, replace the root UI layout, add workflow rail/status widgets, adjust mapping rows, and preserve existing event handlers/settings.
- Modify `tests/test_ui_smoke.py`: add smoke assertions for the workflow rail, compact controls, footer, and mode-dependent visibility.
- Leave split logic tests unchanged unless UI behavior requires a compatibility assertion.
- Rebuild `dist/ExcelSplitter.exe` after verification.

---

### Task 1: Add Failing UI Layout Tests

**Files:**
- Modify: `tests/test_ui_smoke.py`

- [x] **Step 1: Add tests for compact dashboard controls**

Add the following methods to `UISmokeTests`:

```python
    def test_split_app_uses_compact_dashboard_layout(self):
        window = main.SplitApp()
        self.addCleanup(window.deleteLater)

        self.assertTrue(hasattr(window, "workflow_rail"))
        self.assertTrue(hasattr(window, "main_panel_layout"))
        self.assertTrue(hasattr(window, "footer_bar"))

        self.assertLessEqual(window.edit_source.maximumWidth(), 520)
        self.assertLessEqual(window.edit_template.maximumWidth(), 520)
        self.assertLessEqual(window.edit_outdir.maximumWidth(), 520)
        self.assertLessEqual(window.edit_lo_path.maximumWidth(), 360)
        self.assertLessEqual(window.edit_prefix.maximumWidth(), 180)
        self.assertLessEqual(window.edit_suffix.maximumWidth(), 180)

    def test_template_mode_hides_template_file_controls_for_source_template(self):
        window = main.SplitApp()
        self.addCleanup(window.deleteLater)

        window.cmb_template_mode.setCurrentIndex(
            window.cmb_template_mode.findText("Use Source as Template")
        )
        window.on_template_mode_changed()

        self.assertFalse(window.edit_template.isVisible())
        self.assertFalse(window.btn_browse_template.isVisible())
        self.assertFalse(window.mapping_card.isVisible())
```

- [x] **Step 2: Run UI smoke tests and verify failure**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke`

Expected: failure because `workflow_rail`, `main_panel_layout`, and `footer_bar` do not exist yet, and template controls are disabled rather than hidden.

### Task 2: Add Compact UI Helpers and Imports

**Files:**
- Modify: `main.py`

- [x] **Step 1: Update QtWidgets imports**

Add `QGridLayout` and `QSizePolicy` to the existing `PySide6.QtWidgets` import:

```python
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QFileDialog, QGridLayout, QSizePolicy
)
```

- [x] **Step 2: Add compact width constants below template mode constants**

```python
PATH_FIELD_WIDTH = 460
SECONDARY_PATH_FIELD_WIDTH = 320
COMBO_FIELD_WIDTH = 190
SMALL_FIELD_WIDTH = 120
NAME_FIELD_WIDTH = 160
```

- [x] **Step 3: Add helper methods inside `SplitApp` before `_build_ui`**

```python
    def _fixed_width(self, widget, width):
        widget.setMinimumWidth(width)
        widget.setMaximumWidth(width)
        return widget

    def _panel(self, title, icon=None):
        card = SimpleCardWidget()
        card.setBorderRadius(8)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(10)

        header = QHBoxLayout()
        if icon:
            icon_widget = ToolButton(icon)
            icon_widget.setFixedSize(20, 20)
            icon_widget.setEnabled(False)
            header.addWidget(icon_widget)
        header.addWidget(SubtitleLabel(title))
        header.addStretch()
        layout.addLayout(header)
        return card, layout

    def _labeled(self, label, widget):
        wrap = QWidget()
        layout = QVBoxLayout(wrap)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        layout.addWidget(CaptionLabel(label))
        layout.addWidget(widget)
        return wrap
```

- [x] **Step 4: Run syntax check**

Run: `rtk python -m py_compile main.py`

Expected: exit code 0.

### Task 3: Replace Root Layout with Workflow Dashboard

**Files:**
- Modify: `main.py`

- [x] **Step 1: Replace `_build_ui` with dashboard shell**

Replace the current `_build_ui` method with:

```python
    def _build_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        toolbar = QHBoxLayout()
        toolbar.setContentsMargins(18, 12, 18, 8)
        toolbar.setSpacing(10)
        toolbar.addWidget(SubtitleLabel("Excel Splitter"))
        self.lbl_workflow_status = CaptionLabel("Ready")
        toolbar.addWidget(self.lbl_workflow_status)
        toolbar.addStretch()
        self.btn_reset_settings = PushButton("Reset Settings")
        self.btn_reset_settings.clicked.connect(self.reset_settings)
        toolbar.addWidget(self.btn_reset_settings)
        root_layout.addLayout(toolbar)

        body = QHBoxLayout()
        body.setContentsMargins(16, 0, 16, 10)
        body.setSpacing(14)
        self.workflow_rail = self._build_workflow_rail()
        body.addWidget(self.workflow_rail)

        scroll = ScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_widget = QWidget()
        self.main_panel_layout = QVBoxLayout(scroll_widget)
        self.main_panel_layout.setContentsMargins(0, 0, 0, 0)
        self.main_panel_layout.setSpacing(10)
        scroll.setWidget(scroll_widget)
        body.addWidget(scroll, 1)
        root_layout.addLayout(body, 1)

        self._build_source_card()
        self._build_template_card()
        self._build_mapping_card()
        self._build_output_card()
        self._build_log_card()
        self.main_panel_layout.addStretch()

        self.footer_bar = QWidget()
        footer = QHBoxLayout(self.footer_bar)
        footer.setContentsMargins(16, 10, 16, 14)
        footer.setSpacing(10)
        self._build_actions_card(footer)
        root_layout.addWidget(self.footer_bar)
```

- [x] **Step 2: Add workflow rail builder**

Add this method after `_build_ui`:

```python
    def _build_workflow_rail(self):
        rail = SimpleCardWidget()
        rail.setBorderRadius(8)
        rail.setFixedWidth(148)
        layout = QVBoxLayout(rail)
        layout.setContentsMargins(12, 14, 12, 14)
        layout.setSpacing(10)

        layout.addWidget(CaptionLabel("Workflow"))
        self.workflow_steps = {}
        for name in ["Source", "Template", "Output", "Run"]:
            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(8)
            dot = CaptionLabel("○")
            label = BodyLabel(name)
            row_layout.addWidget(dot)
            row_layout.addWidget(label)
            row_layout.addStretch()
            layout.addWidget(row)
            self.workflow_steps[name] = dot

        layout.addStretch()
        return rail
```

- [x] **Step 3: Run UI tests and verify failure moves forward**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke`

Expected: remaining failures from card builders still using `self.scroll_layout` or missing `_build_log_card`.

### Task 4: Convert Source, Template, Mapping, Output, and Log Panels

**Files:**
- Modify: `main.py`

- [x] **Step 1: Rewrite `_build_source_card`**

Replace the current `_build_source_card` method with:

```python
    def _build_source_card(self):
        card, layout = self._panel("Source", FIF.DOCUMENT)

        path_row = QHBoxLayout()
        self.edit_source = self._fixed_width(LineEdit(), PATH_FIELD_WIDTH)
        self.edit_source.setPlaceholderText("Source Excel file")
        self.btn_browse_source = ToolButton(FIF.FOLDER)
        self.btn_browse_source.clicked.connect(self.browse_source)
        path_row.addWidget(self._labeled("Workbook", self.edit_source))
        path_row.addWidget(self.btn_browse_source)
        path_row.addStretch()
        layout.addLayout(path_row)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(8)
        self.cmb_sheet = self._fixed_width(ComboBox(), COMBO_FIELD_WIDTH)
        self.cmb_sheet.setPlaceholderText("Sheet")
        self.btn_load_sheets = PushButton("Load Sheets")
        self.btn_load_sheets.clicked.connect(self.load_sheets)
        self.cmb_key = self._fixed_width(ComboBox(), COMBO_FIELD_WIDTH)
        self.cmb_key.setPlaceholderText("Key Column")
        self.btn_load_headers = PushButton("Load Headers")
        self.btn_load_headers.clicked.connect(self.load_headers)
        self.spin_header_rows = self._fixed_width(SpinBox(), SMALL_FIELD_WIDTH)
        self.spin_header_rows.setRange(1, 100)
        self.spin_header_rows.setValue(5)

        grid.addWidget(self._labeled("Sheet", self.cmb_sheet), 0, 0)
        grid.addWidget(self.btn_load_sheets, 0, 1)
        grid.addWidget(self._labeled("Key Column", self.cmb_key), 0, 2)
        grid.addWidget(self.btn_load_headers, 0, 3)
        grid.addWidget(self._labeled("Header Rows", self.spin_header_rows), 0, 4)
        grid.setColumnStretch(5, 1)
        layout.addLayout(grid)

        self.main_panel_layout.addWidget(card)
```

- [x] **Step 2: Rewrite `_build_template_card`**

Replace the current `_build_template_card` method with:

```python
    def _build_template_card(self):
        card, layout = self._panel("Template", FIF.EDIT)

        mode_row = QHBoxLayout()
        self.cmb_template_mode = self._fixed_width(ComboBox(), 220)
        self.cmb_template_mode.addItems([
            TEMPLATE_MODE_LABELS[TEMPLATE_MODE_TEMPLATE_FILE],
            TEMPLATE_MODE_LABELS[TEMPLATE_MODE_SOURCE_TEMPLATE],
        ])
        self.cmb_template_mode.setCurrentIndex(0)
        self.cmb_template_mode.currentTextChanged.connect(self.on_template_mode_changed)
        mode_row.addWidget(self._labeled("Template Option", self.cmb_template_mode))
        mode_row.addStretch()
        layout.addLayout(mode_row)

        self.template_file_row_widget = QWidget()
        row = QHBoxLayout(self.template_file_row_widget)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(8)
        self.edit_template = self._fixed_width(LineEdit(), PATH_FIELD_WIDTH)
        self.edit_template.setPlaceholderText("Template Excel file")
        self.btn_browse_template = ToolButton(FIF.FOLDER)
        self.btn_browse_template.clicked.connect(self.browse_template)
        row.addWidget(self._labeled("Template Workbook", self.edit_template))
        row.addWidget(self.btn_browse_template)
        row.addStretch()
        layout.addWidget(self.template_file_row_widget)

        self.main_panel_layout.addWidget(card)
```

- [x] **Step 3: Rewrite `_build_mapping_card`**

Replace the current `_build_mapping_card` method with:

```python
    def _build_mapping_card(self):
        self.mapping_card, layout = self._panel("Column Mapping", FIF.EDIT)

        row = QHBoxLayout()
        self.btn_auto_map = PushButton("Auto Map")
        self.btn_auto_map.clicked.connect(lambda: self.refresh_template_mapping(auto=True))
        self.lbl_mapping_status = CaptionLabel("Map template columns to source columns.")
        row.addWidget(self.btn_auto_map)
        row.addWidget(self.lbl_mapping_status)
        row.addStretch()
        layout.addLayout(row)

        self.mapping_rows_widget = QWidget()
        self.mapping_rows_layout = QVBoxLayout(self.mapping_rows_widget)
        self.mapping_rows_layout.setContentsMargins(0, 0, 0, 0)
        self.mapping_rows_layout.setSpacing(6)
        layout.addWidget(self.mapping_rows_widget)

        self.main_panel_layout.addWidget(self.mapping_card)
```

- [x] **Step 4: Rewrite `_build_output_card`**

Replace the current `_build_output_card` method with:

```python
    def _build_output_card(self):
        card, layout = self._panel("Output", FIF.FOLDER)

        out_row = QHBoxLayout()
        self.edit_outdir = self._fixed_width(LineEdit(), PATH_FIELD_WIDTH)
        self.edit_outdir.setPlaceholderText("Output folder")
        self.btn_browse_outdir = ToolButton(FIF.FOLDER)
        self.btn_browse_outdir.clicked.connect(self.browse_outdir)
        out_row.addWidget(self._labeled("Folder", self.edit_outdir))
        out_row.addWidget(self.btn_browse_outdir)
        out_row.addStretch()
        layout.addLayout(out_row)

        options = QGridLayout()
        options.setHorizontalSpacing(10)
        options.setVerticalSpacing(8)
        self.edit_prefix = self._fixed_width(LineEdit(), NAME_FIELD_WIDTH)
        self.edit_prefix.setPlaceholderText("Prefix")
        self.edit_suffix = self._fixed_width(LineEdit(), NAME_FIELD_WIDTH)
        self.edit_suffix.setPlaceholderText("Suffix")
        self.cmb_pdf_engine = self._fixed_width(ComboBox(), 180)
        self.cmb_pdf_engine.addItems(["xlwings", "libreoffice", "none"])
        self.cmb_pdf_engine.setCurrentIndex(0)
        self.edit_lo_path = self._fixed_width(LineEdit(), SECONDARY_PATH_FIELD_WIDTH)
        self.edit_lo_path.setPlaceholderText("soffice.exe")
        self.btn_browse_soffice = ToolButton(FIF.FOLDER)
        self.btn_browse_soffice.clicked.connect(self.browse_soffice)

        options.addWidget(self._labeled("Prefix", self.edit_prefix), 0, 0)
        options.addWidget(self._labeled("Suffix", self.edit_suffix), 0, 1)
        options.addWidget(self._labeled("PDF Engine", self.cmb_pdf_engine), 0, 2)
        options.addWidget(self._labeled("LibreOffice", self.edit_lo_path), 0, 3)
        options.addWidget(self.btn_browse_soffice, 0, 4)
        options.setColumnStretch(5, 1)
        layout.addLayout(options)

        self.main_panel_layout.addWidget(card)
```

- [x] **Step 5: Add `_build_log_card`**

Add this new method before `_build_actions_card`:

```python
    def _build_log_card(self):
        card, layout = self._panel("Log")
        self.txt_log = TextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMinimumHeight(130)
        self.txt_log.setMaximumHeight(180)
        layout.addWidget(self.txt_log)
        self.main_panel_layout.addWidget(card)
```

- [x] **Step 6: Rewrite `_build_actions_card` to receive footer layout**

Replace the current `_build_actions_card` method with:

```python
    def _build_actions_card(self, layout):
        self.btn_generate = PrimaryPushButton(FIF.PLAY, "Generate")
        self.btn_generate.setFixedHeight(38)
        self.btn_generate.clicked.connect(self.on_run_clicked)
        self.progress_bar = ProgressBar()
        self.progress_bar.setFixedWidth(240)
        self.progress_bar.setValue(0)
        self.btn_open_output = PushButton(FIF.FOLDER, "Open Output Folder")
        self.btn_open_output.setFixedHeight(36)
        self.btn_open_output.clicked.connect(self.open_output_folder)
        self.btn_open_output.setVisible(False)
        self.btn_debug = PushButton("Debug Excel")
        self.btn_debug.setFixedHeight(36)
        self.btn_debug.clicked.connect(self.debug_excel)

        layout.addWidget(self.btn_generate)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.btn_open_output)
        layout.addWidget(self.btn_debug)
        layout.addStretch()
```

- [x] **Step 7: Run UI smoke tests**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke`

Expected: remaining failures around mode visibility or mapping row widths.

### Task 5: Update Mode Visibility, Mapping Rows, and Workflow Status

**Files:**
- Modify: `main.py`

- [x] **Step 1: Update `on_template_mode_changed`**

Replace the body after `use_template_file = ...` with:

```python
        self.template_file_row_widget.setVisible(use_template_file)
        self.edit_template.setVisible(use_template_file)
        self.btn_browse_template.setVisible(use_template_file)
        self.mapping_card.setVisible(use_template_file)
        if use_template_file:
            self.refresh_template_mapping(auto=True)
        self.update_workflow_status()
```

- [x] **Step 2: Update `render_mapping_rows` for compact rows**

Inside the `for template_header in self.template_headers:` loop, set bounded row/control widths:

```python
            row_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            row.addWidget(self._fixed_width(BodyLabel(template_header), 220))
            row.addWidget(BodyLabel("->"))
            combo = self._fixed_width(ComboBox(), 240)
```

After selecting combo value, add:

```python
            status = CaptionLabel("Mapped" if selected else "Missing")
            row.addWidget(status)
```

Keep the existing `combo.currentTextChanged.connect(self.save_settings)` and `self.mapping_combos[template_header] = combo`.

- [x] **Step 3: Add workflow status updater**

Add this method before `browse_source`:

```python
    def update_workflow_status(self):
        if not hasattr(self, "workflow_steps"):
            return

        states = {
            "Source": bool(self.edit_source.text().strip() and self.cmb_sheet.currentText().strip() and self.cmb_key.currentText().strip()),
            "Template": self.current_template_mode() == TEMPLATE_MODE_SOURCE_TEMPLATE or bool(self.edit_template.text().strip()),
            "Output": bool(self.edit_outdir.text().strip()),
            "Run": not self.is_running,
        }
        for name, ready in states.items():
            self.workflow_steps[name].setText("●" if ready else "○")
        missing = [name for name, ready in states.items() if not ready and name != "Run"]
        self.lbl_workflow_status.setText("Ready" if not missing else "Missing: " + ", ".join(missing))
```

- [x] **Step 4: Call workflow status after relevant mutations**

Add `self.update_workflow_status()` at the end of:

- `load_settings`
- `reset_settings`
- `browse_source`
- `browse_template`
- `browse_outdir`
- `load_sheets`
- `load_headers`
- `set_busy`

- [x] **Step 5: Run UI and settings tests**

Run: `rtk .venv\Scripts\python.exe -m unittest tests.test_ui_smoke tests.test_settings`

Expected: all UI/settings tests pass.

### Task 6: Verification, Build, and Commit

**Files:**
- Modify: `dist/ExcelSplitter.exe`

- [x] **Step 1: Run full tests**

Run: `rtk .venv\Scripts\python.exe -m unittest discover`

Expected: all tests pass.

- [x] **Step 2: Run syntax and whitespace checks**

Run: `rtk python -m py_compile main.py`

Expected: exit code 0.

Run: `rtk git diff --check`

Expected: exit code 0.

- [x] **Step 3: Build executable**

Run: `rtk cmd /c build.cmd`

Expected: exit code 0 and `dist\ExcelSplitter.exe` exists. Existing optional PyInstaller hidden-import warnings may still appear.

- [x] **Step 4: Inspect final status**

Run: `rtk git status --short`

Expected: only intentional UI redesign, tests, plan, and rebuilt executable changes.

- [x] **Step 5: Commit implementation**

Run:

```bash
rtk git add -A
rtk git commit -m "feat: redesign ui as compact workflow dashboard"
```

Expected: commit succeeds.
