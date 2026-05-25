"""
inspect_xml.py

Reads the internal XML of an .xlsm/.xlsx file to surface things that
are invisible in Excel's UI:

  - CF rule sqref fragmentation (a top cause of "CF not working on new rows")
  - bgColor vs fgColor on FormatConditions (openpyxl-corruption signature)
  - #REF! errors left behind by historical row/column deletes
  - Data validation list ranges
  - Merged cell positions

READ-ONLY. Never modify the file. See docs/What-AI-Wont-Tell-You.html — Error 001.

Usage:
  python inspect_xml.py <file.xlsm> [sheet_name]
"""

import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


XL_NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def inspect(filepath, sheet_name=None):
    filepath = Path(filepath)
    if not filepath.exists():
        print(f"File not found: {filepath}")
        sys.exit(1)

    with zipfile.ZipFile(filepath) as zf:
        names = zf.namelist()

        # Find sheet XML files
        sheet_files = [n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        print(f"\nFound {len(sheet_files)} sheets in workbook")

        for sheet_file in sheet_files:
            with zf.open(sheet_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()

            print(f"\n{'='*60}")
            print(f"SHEET FILE: {sheet_file}")
            print(f"{'='*60}")

            # Conditional formatting
            cf_rules = root.findall("x:conditionalFormatting", XL_NS)
            if cf_rules:
                print(f"\n  CF blocks: {len(cf_rules)}")
                for i, cf in enumerate(cf_rules, 1):
                    sqref = cf.get("sqref", "")
                    segments = sqref.split()
                    print(f"    [{i}] sqref segments: {len(segments)}")
                    if len(segments) > 1:
                        print(f"        ⚠ FRAGMENTED — new rows in gaps will have no CF coverage")
                        print(f"        First 3 segments: {' '.join(segments[:3])}")

                    rules = cf.findall("x:cfRule", XL_NS)
                    for rule in rules:
                        rule_type = rule.get("type", "?")
                        formula_el = rule.find("x:formula", XL_NS)
                        formula = formula_el.text if formula_el is not None else ""
                        if "#REF!" in formula:
                            print(f"        ⚠ #REF! in formula: {formula[:60]}")

                        # Check for bgColor vs fgColor in dxfs (handled separately)
            else:
                print("\n  No CF rules in this sheet")

            # Data validations
            validations = root.findall(".//x:dataValidation", XL_NS)
            if validations:
                print(f"\n  Data validations: {len(validations)}")
                for v in validations[:5]:
                    sqref = v.get("sqref", "")
                    formula = v.find("x:formula1", XL_NS)
                    formula_text = formula.text if formula is not None else ""
                    print(f"    sqref={sqref}  formula={formula_text[:60]}")

            # Merged cells
            merges = root.findall(".//x:mergeCell", XL_NS)
            if merges:
                print(f"\n  Merged cells: {len(merges)}")

        # Inspect styles.xml for dxfs (CF fill format definitions)
        if "xl/styles.xml" in names:
            with zf.open("xl/styles.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()

            dxfs = root.find("x:dxfs", XL_NS)
            if dxfs is not None:
                print(f"\n{'='*60}")
                print(f"DXFS (CF fill format definitions)")
                print(f"{'='*60}")
                for i, dxf in enumerate(dxfs.findall("x:dxf", XL_NS), 1):
                    fill = dxf.find("x:fill", XL_NS)
                    if fill is not None:
                        pattern = fill.find("x:patternFill", XL_NS)
                        if pattern is not None:
                            bg = pattern.find("x:bgColor", XL_NS)
                            fg = pattern.find("x:fgColor", XL_NS)
                            bg_val = bg.get("rgb", "") if bg is not None else ""
                            fg_val = fg.get("rgb", "") if fg is not None else ""
                            pattern_type = pattern.get("patternType", "")

                            warn = ""
                            if bg_val and not fg_val:
                                warn = "  ⚠ ONLY bgColor — Excel will render as invisible (openpyxl corruption signature)"
                            print(f"  dxf[{i-1}]: patternType={pattern_type} fg={fg_val} bg={bg_val}{warn}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    inspect(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
