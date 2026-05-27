"""One-shot: append a 'cpi iip' sheet to the snapshot report containing
every row whose dataset name mentions CPI or IIP, drawn from every sheet
in the source file. Section headers separate each origin sheet.

Usage:
    python scripts/extract_cpi_iip.py <report.xlsx>
    python scripts/extract_cpi_iip.py    # uses the snapshot the user named
"""
import sys
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

DEFAULT = Path(
    "reports/comparison_report_20260526_154750_snapshot_20260527_091901.xlsx"
)
NEW_SHEET = "cpi iip"
KEYWORDS = ("CPI", "IIP")

_HEADER_FILL = PatternFill("solid", fgColor="305496")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_SECTION_FILL = PatternFill("solid", fgColor="FFE699")
_SECTION_FONT = Font(bold=True, size=12)


def _matches(name) -> bool:
    if not name:
        return False
    s = str(name).upper()
    return any(k in s for k in KEYWORDS)


def main(path: Path) -> None:
    if not path.exists():
        sys.exit(f"File not found: {path}")

    wb = load_workbook(path)
    src_names = [n for n in wb.sheetnames if n != NEW_SHEET]

    if NEW_SHEET in wb.sheetnames:
        del wb[NEW_SHEET]
    out = wb.create_sheet(NEW_SHEET)

    cur = 1
    grand_total = 0

    for sname in src_names:
        ws = wb[sname]
        if ws.max_row < 2:
            continue
        header = [c.value for c in ws[1]]

        kept = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            # Dataset name is always column A in this report's sheets.
            if _matches(row[0]):
                kept.append(row)

        # Section banner
        out.cell(row=cur, column=1, value=f"=== {sname} ({len(kept)} rows) ===")
        out.cell(row=cur, column=1).font = _SECTION_FONT
        for col in range(1, max(len(header), 1) + 1):
            out.cell(row=cur, column=col).fill = _SECTION_FILL
        cur += 1

        # Header row
        for ci, h in enumerate(header, start=1):
            cell = out.cell(row=cur, column=ci, value=h)
            cell.font = _HEADER_FONT
            cell.fill = _HEADER_FILL
        cur += 1

        # Matching rows
        for row in kept:
            for ci, v in enumerate(row, start=1):
                out.cell(row=cur, column=ci, value=v)
            cur += 1

        cur += 1  # blank line between sections
        grand_total += len(kept)
        print(f"  {sname:>15s}: {len(kept):6d} rows matched")

    # Auto-size first 12 columns roughly so the sheet is readable.
    for col_idx in range(1, 13):
        max_len = 0
        letter = out.cell(row=1, column=col_idx).column_letter
        for row in out.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
            v = row[0]
            if v is None:
                continue
            max_len = max(max_len, min(len(str(v)), 60))
        out.column_dimensions[letter].width = max(10, min(max_len + 2, 60))

    try:
        wb.save(path)
        print(f"\nOK: Wrote sheet '{NEW_SHEET}' to {path}  ({grand_total} total rows)")
    except PermissionError:
        alt = path.with_name(path.stem + "_cpi_iip.xlsx")
        wb.save(alt)
        print(
            f"\nWARN: Source file is locked (open in Excel?). Saved a copy to:\n"
            f"   {alt}  ({grand_total} total rows)"
        )


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT)
