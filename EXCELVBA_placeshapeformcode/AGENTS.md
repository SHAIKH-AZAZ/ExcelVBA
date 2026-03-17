# AGENTS.md

## Purpose
Guidelines for safely extending this VBA workbook automation.

## Scope
Applies to:
- `sheet.bas`
- `module.bas`

## Core Contract
- Shape code source column is `G`.
- Placeholder data columns are `I:O` mapped to `{A}:{G}`.
- Shape destination anchor is column `S`.
- Shape name format is `<ShapeCode>_<RowNumber>`.
- `ShapeLibrary` worksheet must exist and contain shapes named by code.

## Safe Edit Rules
- Keep `Worksheet_Change` guarded with `Application.EnableEvents = False/True` and a single safe exit path.
- Preserve single-cell guards unless multi-cell paste behavior is intentionally designed.
- Avoid changing logic that depends on `ActiveCell`/`Selection` without updating all dependent macros.
- Always restore Excel state in cleanup blocks:
  - `ScreenUpdating`
  - `EnableEvents`
  - `DisplayAlerts`
- Keep shape create/delete naming logic identical; mismatches cause stale shapes.

## Error Handling
- Use fail-safe exits (`GoTo CleanExit` / `GoTo SafeExit`) rather than silent fall-through.
- Use targeted `On Error Resume Next` only around expected missing-object operations.
- Return to normal error handling immediately after risky lines.

## Consistency Checks After Any Change
- Edit a value in `G` for one row:
  - shape is created in `S`
  - shape name is `<code>_<row>`
- Edit any value in `I:N` for same row:
  - previous shape is deleted
  - shape is recreated with updated text
- Verify placeholders `{A}`..`{G}` are replaced from `I:O`.

## Known Risk To Address
Current code has a likely mismatch:
- placement reads code from `G`
- deletion reads code from `L`

Unify both to the intended source column unless this split is explicitly required.

## Change Style
- Keep edits small and local.
- Prefer clear procedure names over adding hidden side effects.
- Update `README.md` when column mappings, triggers, or workflow are changed.
