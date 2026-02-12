# EXCELVBA Place Shape Form Code

## Overview
This workbook automation places and refreshes shapes on a worksheet based on a shape code and row data.

The codebase currently contains two VBA files:
- `sheet.bas`: Worksheet event routing (`Worksheet_Change`).
- `module.bas`: Core shape operations (place, delete, placeholder text replacement).

## File-by-File Responsibilities

### `sheet.bas`
Handles user edits and calls core macros.

- `Worksheet_Change(ByVal Target As Range)`
  - Disables events at entry, restores on exit.
  - If change is in column `G`: calls `Handle_L_Column_Change`.
  - If change is in columns `I:N`: calls `Handle_MS_Columns_Change`.

- `Handle_L_Column_Change(ByVal Target As Range)`
  - Runs only for single-cell, non-empty edits.
  - Selects the edited cell.
  - Calls `PlaceShapeFromCode`.

- `Handle_MS_Columns_Change(ByVal Target As Range)`
  - Runs only for single-cell edits.
  - Checks shape code existence in column `G` of the edited row.
  - Selects column `G` cell in that row.
  - Calls `DeleteShapeFromRow`, then `PlaceShapeFromCode`.

### `module.bas`
Implements shape copy/paste, naming, positioning, and text placeholder replacement.

- `PlaceShapeFromCode()`
  1. Validates environment and required sheet `ShapeLibrary`.
  2. Uses active row and reads shape code from column `G`.
  3. Fetches source shape from `ShapeLibrary.Shapes(shapeCode)`.
  4. Deletes existing shape named `<shapeCode>_<row>` on active sheet.
  5. Copies source shape and pastes near column `S`.
  6. Renames pasted shape to `<shapeCode>_<row>`.
  7. Centers shape in cell `S` of the row.
  8. Replaces placeholders `{A}`..`{G}` using row values from columns `I:O`.
  9. Restores original active-cell selection.

- `DeleteShapeFromRow()`
  - Uses active row.
  - Builds target shape name `<shapeCode>_<row>`.
  - Deletes that shape if present.

- `ProcessTextRecursively(shp, keys, values)`
  - Recursively traverses grouped shapes.
  - Replaces all placeholders in `TextFrame2` text.

## Data/Column Mapping
- Shape code source for placement: column `G`.
- Triggered update columns: `I:N`.
- Placeholder source values:
  - `{A}` -> `I`
  - `{B}` -> `J`
  - `{C}` -> `K`
  - `{D}` -> `L`
  - `{E}` -> `M`
  - `{F}` -> `N`
  - `{G}` -> `O`
- Shape destination anchor cell: column `S`.
- Shape naming scheme: `<ShapeCode>_<RowNumber>`.

## Runtime Behavior
- The workflow depends on `ActiveCell` / `Selection` context.
- Macros run in a silent mode pattern:
  - `Application.ScreenUpdating = False`
  - `Application.EnableEvents = False`
  - `Application.DisplayAlerts = False`
  - then restored in `CleanExit`.
- Existing shape of same computed name is deleted before insertion.

## Dependencies and Assumptions
- A worksheet named `ShapeLibrary` must exist.
- Shape names in `ShapeLibrary` must match values entered in column `G`.
- Pasted shapes should support `TextFrame2` for placeholder replacement.

## Known Consistency Risk
There is a likely column mismatch:
- `PlaceShapeFromCode` reads code from column `G`.
- `DeleteShapeFromRow` currently reads code from column `L`.

If `L` does not equal `G`, deletion may target a different/non-existing shape name and leave stale shapes.

## Suggested Next Maintenance Step
Align `DeleteShapeFromRow` to use the same shape-code column as placement (column `G`) unless the design intentionally stores a separate delete key in column `L`.
