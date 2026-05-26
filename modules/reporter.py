import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .comparator import ComparisonResult

log = logging.getLogger(__name__)

# Palette
_BLUE       = "FF4472C4"
_WHITE      = "FFFFFFFF"
_LIGHT_GRAY = "FFF2F2F2"
_LIGHT_GREEN = "FFC6EFCE"
_LIGHT_RED   = "FFFFC7CE"
_LIGHT_AMBER = "FFFFEB9C"
_DARK_RED    = "FF9C0006"
_DARK_GREEN  = "FF276221"
_DARK_AMBER  = "FF9C5700"


def _hdr(cell, bg=_BLUE) -> None:
    cell.font      = Font(bold=True, color=_WHITE, size=11)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _auto_width(ws) -> None:
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width  = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[letter].width = min(max(width + 2, 10), 70)


class Reporter:
    """Accumulates per-table results and writes a final .xlsx comparison report."""

    def __init__(self) -> None:
        self._entries: List[Dict] = []

    # ── Public API ────────────────────────────────────────────────────────────

    def add_result(self, dataset: str, meta: Dict, result: ComparisonResult) -> None:
        self._entries.append({"dataset": dataset, "meta": meta, "result": result})

    def add_error(self, dataset: str, label: str, error_msg: str) -> None:
        r = ComparisonResult(status="ERROR", error=error_msg)
        empty_meta = {
            "table_name": label, "table_no": "",
            "product": "", "category": "", "release_date": "",
        }
        self._entries.append({"dataset": dataset, "meta": empty_meta, "result": r})

    def generate(self, report_dir: Path = Path("reports")) -> Path:
        report_dir.mkdir(parents=True, exist_ok=True)
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = report_dir / f"comparison_report_{ts}.xlsx"

        wb = openpyxl.Workbook()
        self._write_summary(wb)
        self._write_mismatches(wb)

        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        wb.save(path)
        log.info("Report saved: %s", path)
        print(f"\n✓ Report saved → {path}")
        return path

    # ── Sheet builders ────────────────────────────────────────────────────────

    def _write_summary(self, wb: openpyxl.Workbook) -> None:
        ws = wb.create_sheet("Summary", 0)
        ws.freeze_panes = "A2"

        cols = [
            "Dataset", "Table No.", "Table Name", "Product",
            "Category", "Release Date", "Status",
            "Total Rows", "Mismatches", "Notes",
        ]
        for ci, h in enumerate(cols, 1):
            _hdr(ws.cell(row=1, column=ci, value=h))

        for ri, entry in enumerate(self._entries, 2):
            meta: Dict            = entry["meta"]
            res: ComparisonResult = entry["result"]

            ws.cell(ri, 1,  entry["dataset"])
            ws.cell(ri, 2,  meta.get("table_no",      ""))
            ws.cell(ri, 3,  meta.get("table_name",    ""))
            ws.cell(ri, 4,  meta.get("product",       ""))
            ws.cell(ri, 5,  meta.get("category",      ""))
            ws.cell(ri, 6,  meta.get("release_date",  ""))

            sc = ws.cell(ri, 7, res.status)
            if res.status == "PASS":
                sc.fill = PatternFill("solid", fgColor=_LIGHT_GREEN)
                sc.font = Font(bold=True, color=_DARK_GREEN)
            elif res.status == "FAIL":
                sc.fill = PatternFill("solid", fgColor=_LIGHT_RED)
                sc.font = Font(bold=True, color=_DARK_RED)
            else:
                sc.fill = PatternFill("solid", fgColor=_LIGHT_AMBER)
                sc.font = Font(bold=True, color=_DARK_AMBER)
            sc.alignment = Alignment(horizontal="center")

            ws.cell(ri, 8,  res.total_rows)
            ws.cell(ri, 9,  res.mismatch_count)

            notes: List[str] = []
            if res.error:
                notes.append(f"Error: {res.error}")
            if res.col_mismatch:
                if res.missing_cols:
                    notes.append(f"Missing cols (Excel only): {', '.join(res.missing_cols)}")
                if res.extra_cols:
                    notes.append(f"Extra cols (Web only): {', '.join(res.extra_cols)}")
            if res.excel_extra_rows:
                notes.append(f"Excel has {res.excel_extra_rows} extra row(s)")
            if res.web_extra_rows:
                notes.append(f"Web has {res.web_extra_rows} extra row(s)")
            ws.cell(ri, 10, "; ".join(notes))

            if ri % 2 == 0:
                for ci in [1, 2, 3, 4, 5, 6, 8, 9, 10]:
                    ws.cell(ri, ci).fill = PatternFill("solid", fgColor=_LIGHT_GRAY)

        _auto_width(ws)
        ws.row_dimensions[1].height = 30

    def _write_mismatches(self, wb: openpyxl.Workbook) -> None:
        ws = wb.create_sheet("Mismatches", 1)
        ws.freeze_panes = "A2"

        cols = ["Dataset", "Table No.", "Table Name", "Row #", "Column", "Excel Value", "Web Value"]
        for ci, h in enumerate(cols, 1):
            _hdr(ws.cell(row=1, column=ci, value=h))

        rn = 2
        for entry in self._entries:
            meta: Dict            = entry["meta"]
            res: ComparisonResult = entry["result"]

            for mm in res.mismatches:
                ws.cell(rn, 1, entry["dataset"])
                ws.cell(rn, 2, meta.get("table_no",   ""))
                ws.cell(rn, 3, meta.get("table_name", ""))
                ws.cell(rn, 4, mm.row)
                ws.cell(rn, 5, mm.column)

                ec = ws.cell(rn, 6, mm.excel_value)
                ec.fill = PatternFill("solid", fgColor=_LIGHT_RED)
                ec.font = Font(color=_DARK_RED)

                wc = ws.cell(rn, 7, mm.web_value)
                wc.fill = PatternFill("solid", fgColor=_LIGHT_AMBER)
                wc.font = Font(color=_DARK_AMBER)

                if rn % 2 == 0:
                    for ci in [1, 2, 3, 4, 5]:
                        ws.cell(rn, ci).fill = PatternFill("solid", fgColor=_LIGHT_GRAY)
                rn += 1

        if rn == 2:
            ws.cell(2, 1, "No mismatches found — all data matches.")

        _auto_width(ws)
        ws.row_dimensions[1].height = 30
