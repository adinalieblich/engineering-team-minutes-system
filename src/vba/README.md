# VBA Modules

## How to use these files

These are exported VBA modules. They are not executable as standalone files. To use them:

1. Open a fresh `.xlsm` workbook in Excel.
2. Press `Alt+F11` to open the VBA editor.
3. Right-click `VBAProject (your file)` → `Import File…`
4. Select `Module1.bas`.
5. Repeat for `Module2-archive.bas` only if you need access to archived/build-time macros.

## Module1.bas — live macros (daily use)

Macros that run in normal day-to-day operation of the workbook.

### Keyboard shortcuts
- `Ctrl+Shift+A` — `AddRow` — adds a data row below the current row, inherits Project (col I), stamps Date Added (col J)
- `Ctrl+Shift+P` — `AddProject` — inserts a project banner below the current row (prompts for project name)

Lifecycle banners (ACTIVE PROJECTS, DLP PROJECTS, etc.) are typed manually — there is no auto-rebuild macro by design.

### Buttons / toolbar
- `ShowActionRegisterView` — filters Meeting Minutes to Type=Action, unhides the Project column so it's visible in the filtered view
- `RestoreFullView` — undoes the above (clears the filter, re-hides Project)
- `ToggleDoneRows` — hides/shows rows with Status=Done on Meeting Minutes
- `ToggleClosedRisks` — same for Risk Register
- `PopulateRiskRegister` — copies new Risk rows from Meeting Minutes into Risk Register sheet (additive only, never overwrites existing entries — matched on compound key Description + Owner)

### One-shot (kept in Module1 for re-runs)
- `WriteODDaysFormula` — writes the Days OD formula to column G
- `FixSummaryOverdueFormula` — rebuilds the Summary tab overdue counts
- `BuildSummaryTab` — full Summary tab rebuild
- `BuildHowToUseTab` — rebuilds the embedded How To Use sheet
- `BuildChangelogTab` — rebuilds the locked Changelog sheet

## Module2-archive.bas — old/build-time macros

These are kept for reference only. Most reference OLD column positions or OLD architecture (separate Action Register sheet, checkboxes, etc.). **Do not run on the current file** without checking each one against the new column structure.

See `docs/Macro-Archive.html` for the full rationale on each macro.

## Column layout (current)

These macros assume:

| Col | Field |
|---|---|
| A | Type |
| B | Description |
| C | Owner |
| D | Due |
| E | Status |
| F | Priority |
| G | Days OD |
| H | Notes |
| I | Project (hidden) |
| J | Date Added (hidden) |

Data starts at **row 16**. Rows 1–14 are the header block. Row 15 is column headers.

## Sanitisation note

These files have been sanitised for public release:
- Owner initials → `TM1` through `TM6` and `ALL`
- All client-identifying names replaced with generic terms
- No real project names or descriptions

The real working version is held privately.
