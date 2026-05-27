# Templating Modes, Column Mapping, and QSettings Design

## Context

Excel Splitter currently always uses a separate template workbook. It reads the source sheet into a DataFrame, tries to exact-match template headers to source headers, then writes grouped rows into the template active sheet. The PySide6 UI stores configuration through manual Save/Load `.ini` actions.

## Goals

- Add two template modes:
  - **As Is**: use a separate template file.
  - **Source as Template**: use the selected source worksheet as the template.
- Support automatic best-effort column mapping when source and template headers differ.
- Require manual mapping before generation when any required template column is unmapped.
- Replace `.ini`-based config with Qt `QSettings`.

## Non-Goals

- No multi-sheet workbook splitting in this phase.
- No fuzzy matching UI that silently guesses low-confidence mappings without user review.
- No support for preserving VBA/macros beyond what openpyxl already supports.

## User-Facing Behavior

### Template Mode: As Is

The user selects a separate template workbook, as today. The app reads template headers from the configured header row and maps each template output column to a source column.

Mapping is attempted automatically using normalized header names. Normalization trims whitespace, lowercases text, and ignores repeated spaces and common punctuation. Exact normalized matches are selected automatically.

If every template column is mapped, generation can proceed. If one or more template columns are unmapped, the app blocks generation and shows the user which mappings must be completed manually.

### Template Mode: Source as Template

The app uses the selected source worksheet as the output template. For each key value, it creates one output workbook containing a copied version of that worksheet only. Header rows, column widths, row heights, styles, formulas, merged cells, and worksheet layout are preserved as much as openpyxl supports.

Only data rows for the current key remain in the output worksheet. Other source workbook sheets are not included.

Column mapping is not required in this mode because the output columns are the selected source worksheet columns.

### Column Mapping UI

Add a **Column Mapping** card below Template:

- **Auto Map** button.
- A visible mapping list/table with rows like `Template Column -> Source Column`.
- Source column selection per template column.
- Unmapped required template columns are visibly marked.

Generation is blocked if Template Mode is **As Is** and any required template column is unmapped.

### QSettings Configuration

Replace Save/Load `.ini` as the primary config mechanism with Qt `QSettings`.

The app automatically loads settings on startup and saves settings when the user changes relevant fields or starts generation. Stored fields include:

- source path
- sheet name
- key column
- template mode
- template path
- header row count
- output directory
- PDF engine
- LibreOffice path
- filename prefix
- filename suffix
- column mappings for the current source/template/header combination

The Save `.ini` and Load `.ini` toolbar buttons are removed or replaced with a compact **Reset Settings** action. The implementation should prefer removing them unless a clear compatibility need appears during development.

## Internal Design

### Mapping Model

Introduce pure helper functions:

- `normalize_header(value: object) -> str`
- `read_excel_headers(path, sheet_name, header_rows) -> list[str]`
- `read_template_headers(path, header_rows) -> tuple[list[str], int]`
- `auto_map_columns(template_headers, source_headers) -> dict[str, str | None]`
- `validate_column_mapping(template_headers, mapping) -> list[str]`

The splitter should accept an explicit `column_mapping` for As Is mode. Data written to template columns follows template header order:

`template_column -> mapped source column -> group row value`

### Splitter Mode API

Extend `split_excel_with_template` with:

- `template_mode: str = "as_is"`
- `column_mapping: dict[str, str] | None = None`

For `template_mode == "as_is"`, keep the existing template workbook flow but use explicit mapping instead of relying only on exact header intersection.

For `template_mode == "source"`, create output workbooks by copying the selected worksheet and deleting rows that do not belong to the current key group.

### Error Handling

- Invalid source/template paths continue to show blocking errors.
- Missing mappings in As Is mode raise a clear `ValueError` listing unmapped template columns.
- Unsupported template mode raises `ValueError`.
- Source as Template mode should log that separate template path is ignored.

## Testing

Add tests for:

- Header normalization and auto-mapping.
- Validation failure when template columns remain unmapped.
- As Is output uses manual mapping when template/source headers differ.
- Source as Template output preserves only the selected worksheet and only rows for one key.
- QSettings load/save uses Qt settings rather than `.ini` files.
- UI smoke still constructs successfully.

## Acceptance Criteria

- User can select **As Is** or **Source as Template**.
- As Is mode auto-maps matching headers and blocks generation until all required template mappings are filled.
- Source as Template mode splits the selected worksheet directly and does not require a template file.
- App settings persist automatically between launches through `QSettings`.
- Existing tests and new tests pass.
- Windows build succeeds and produces `dist\ExcelSplitter.exe`.
