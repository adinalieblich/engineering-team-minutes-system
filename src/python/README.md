# Python utilities

Read-only verification and inspection scripts for the Excel workbook.

## ⚠ Critical rule

**Never use these scripts (or openpyxl in general) to write Excel files that will be opened in Excel.**

openpyxl writes conditional-formatting fills as `bgColor` which Excel renders as invisible. The resulting file looks corrupt to the user but valid to openpyxl. The only fix is rebuilding the file via VBA in a fresh blank workbook. See `docs/What-AI-Wont-Tell-You.html` — Error 001.

These scripts are safe because they only:
- READ existing files for inspection
- WRITE BRAND NEW files (helper outputs like `Missing_Rows.xlsx`) that are NEVER opened, edited, and re-saved in Excel as the working file

## Setup

```bash
cd src/python
pip install -r requirements.txt
```

## Scripts

### `verify_against_source.py`

Compares a working `.xlsm` against an emergency-save source-of-truth `.xlsx` using compound key matching (Description + Owner).

```bash
python verify_against_source.py working_file.xlsm source_of_truth.xlsx
```

Outputs to `out/`:
- `Missing_Rows.xlsx` — rows in source but missing from working file
- `Updates_Needed.xlsx` — field-level mismatches

### `inspect_xml.py`

Reads internal XML to surface things invisible in Excel's UI:
- CF sqref fragmentation (a top cause of "CF not working on new rows")
- bgColor vs fgColor on FormatConditions (openpyxl-corruption signature)
- `#REF!` errors left behind by historical row/column deletes
- Data validation list ranges
- Merged cell positions

```bash
python inspect_xml.py file.xlsm
```

## Why use Python at all?

Two things:

1. **Inspection of internal XML.** Excel's UI hides sqref fragmentation, dxf details, and a lot of state that matters for diagnosis. openpyxl reads it. This is genuinely useful and not replaceable by VBA.

2. **Bulk comparison and analysis.** Compound-key matching across thousands of rows is faster and more reliable in Python than VBA.

What Python is NOT for in this project:
- Building the working `.xlsm` (use VBA)
- Modifying formatting in the working `.xlsm` (use VBA)
- Anything where the output file will be opened in Excel as the live working file
