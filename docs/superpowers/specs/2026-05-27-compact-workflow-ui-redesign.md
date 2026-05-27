# Compact Workflow UI Redesign

## Goal

Redesign ExcelSplitter from stacked accordion cards into a compact workflow dashboard. The UI should feel denser and more professional without becoming crowded. Inputs should have deliberate widths instead of stretching across the full window.

## Selected Direction

Use the approved **Workflow Dashboard** direction:

- A narrow left workflow rail shows `Source`, `Template`, `Output`, and `Run`.
- The main work area uses compact grouped panels with bounded controls.
- The run controls stay visible in a compact footer.
- The log remains available but secondary.

## Layout

The window remains a single PySide6/qfluentwidgets desktop app. The root layout changes from a single vertical scroll of accordion cards to:

- Top toolbar: app title/status, reset settings action.
- Body: left workflow rail plus right content area.
- Footer: generate/progress/open-output/debug actions.

The left rail is informational and navigational-looking, but it does not need full page switching in this phase. It should show clear workflow stages and completion state based on loaded paths/selections where practical.

## Main Panels

### Source Panel

Contains source workbook, sheet, key column, and header rows.

- Source path uses a bounded path field plus folder icon button.
- `Load Sheets` and `Load Headers` remain explicit actions.
- Sheet, key column, and header rows sit on one compact row when width allows.

### Template Panel

Contains template mode and template file controls.

- Template option labels remain exactly `Use Template File` and `Use Source as Template`.
- In `Use Template File`, show template path and browse button.
- In `Use Source as Template`, hide or disable template-file-specific controls and mapping.

### Column Mapping Panel

Visible only for `Use Template File`.

- Show `Auto Map` as the primary action.
- Render mappings in compact rows: template column, source column dropdown, status.
- Avoid full-width dropdowns; use bounded widths and keep rows scannable.
- Missing mappings should be visually obvious and should still block generation.

### Output Panel

Contains output folder, file naming, PDF engine, and LibreOffice path.

- Output folder uses a bounded path field plus folder icon button.
- Prefix and suffix sit together.
- PDF engine and LibreOffice path sit together where practical.
- LibreOffice path can remain visible but visually secondary unless `libreoffice` is selected.

### Footer

Contains:

- Primary `Generate` button.
- Progress bar.
- `Open Output Folder`, visible after successful generation.
- `Debug Excel`.

The footer should keep the run action easy to find without making the log dominate the initial view.

## UX Rules

- Keep controls compact but not cramped.
- Avoid full-width input bars unless the window is too narrow and responsive wrapping requires it.
- Use icon buttons for browse actions and retain clear labels near control groups.
- Minimize explanatory text inside the app. Prefer concise labels, placeholders, and status indicators.
- Preserve all existing behavior: splitting, template modes, manual mapping, QSettings persistence, PDF options, and debug actions.

## Technical Notes

- Keep the implementation in `main.py` for this phase.
- Introduce small UI helper methods for repeated row patterns and fixed control widths.
- Do not change split logic unless a UI integration bug requires it.
- Existing tests for template behavior and settings must continue to pass.

## Verification

Required checks:

- UI smoke test verifies construction and key controls.
- UI layout tests verify compact width constraints and mode-dependent visibility.
- Existing settings tests verify persisted fields still load and save.
- Existing templating tests verify split behavior is unchanged.
- Full unittest discovery passes.
- `main.py` compiles.
- Build command produces `dist\ExcelSplitter.exe`.
