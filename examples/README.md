# Examples

This folder will contain a sanitised example workbook so that anyone (or any AI) picking up the project can see the system in action without needing access to the real working file.

## Status

**Not yet built.** The sanitised example workbook is on the to-do list. The work involves:

1. Take a copy of the latest working `.xlsm`
2. Strip real meeting content (descriptions, notes)
3. Replace owner initials with the generic `TM1`–`TM6` and `ALL`
4. Replace project names with `Project A`, `Project B`, etc.
5. Replace section names with generic engineering categories
6. Verify no client-identifying content remains
7. Save as `example-workbook.xlsm` here

Anyone picking up Claude Code on this can build this from the VBA modules in `src/vba/` and a blank workbook. The sequence of macros to run is roughly:

1. Run `BuildHowToUseTab` to create the embedded help
2. Run `BuildChangelogTab` to create the locked changelog
3. Manually build the Meeting Minutes header block (rows 1–14) — or write a `BuildHeaderBlock` macro
4. Apply the dropdown validations to columns A, C, E, F
5. Apply the CF rules per `docs/Style-Guide.md`
6. Add 10–20 example rows with dummy data
7. Run `BuildSummaryTab` and `PopulateRiskRegister` to demo the registers

## Why the real file isn't here

It contains live operational data for a working team. The sanitised example will demonstrate the system without the data.
