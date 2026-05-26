import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional

log = logging.getLogger(__name__)


def _norm(val) -> str:
    """Normalise a cell value: convert to str and strip whitespace."""
    return "" if val is None else str(val).strip()


@dataclass
class Mismatch:
    row: int          # 1-indexed data row number
    column: str       # column header name
    excel_value: str
    web_value: str


@dataclass
class ComparisonResult:
    status: str = "PASS"           # PASS | FAIL | ERROR
    error: Optional[str] = None
    total_rows: int = 0
    total_cols: int = 0
    mismatches: List[Mismatch] = field(default_factory=list)
    col_mismatch: bool = False
    missing_cols: List[str] = field(default_factory=list)   # in Excel, not in Web
    extra_cols: List[str] = field(default_factory=list)     # in Web, not in Excel
    excel_extra_rows: int = 0
    web_extra_rows: int = 0

    @property
    def mismatch_count(self) -> int:
        return len(self.mismatches)


def compare(excel_data: Optional[Dict], web_data: Optional[Dict]) -> ComparisonResult:
    """
    Compare an Excel dict with a web-table dict, both of the form:
        {"headers": [...], "rows": [[...], ...]}

    Returns a ComparisonResult with cell-level mismatch details.
    """
    result = ComparisonResult()

    if excel_data is None:
        result.status = "ERROR"
        result.error = "Excel file could not be read"
        return result

    if web_data is None or not web_data.get("rows"):
        result.status = "ERROR"
        result.error = "Web table data could not be read"
        return result

    excel_hdrs = [_norm(h) for h in excel_data["headers"]]
    web_hdrs   = [_norm(h) for h in web_data["headers"]]

    # ── Column comparison ────────────────────────────────────────────────────
    if excel_hdrs != web_hdrs:
        log.warning("Column mismatch:\n  Excel: %s\n  Web:   %s", excel_hdrs, web_hdrs)
        result.col_mismatch = True
        result.missing_cols = [h for h in excel_hdrs if h not in web_hdrs]
        result.extra_cols   = [h for h in web_hdrs   if h not in excel_hdrs]

    # ── Normalise all cell values ────────────────────────────────────────────
    excel_rows = [[_norm(c) for c in r] for r in excel_data["rows"]]
    web_rows   = [[_norm(c) for c in r] for r in web_data["rows"]]

    result.total_rows = max(len(excel_rows), len(web_rows))
    result.total_cols = len(excel_hdrs)

    if len(excel_rows) > len(web_rows):
        result.excel_extra_rows = len(excel_rows) - len(web_rows)
        log.warning("Excel has %d extra rows vs web table", result.excel_extra_rows)
    elif len(web_rows) > len(excel_rows):
        result.web_extra_rows = len(web_rows) - len(excel_rows)
        log.warning("Web table has %d extra rows vs Excel", result.web_extra_rows)

    # ── Cell-by-cell comparison on shared columns ────────────────────────────
    if result.col_mismatch:
        shared = [h for h in excel_hdrs if h in web_hdrs]
        exc_indices = [excel_hdrs.index(h) for h in shared]
        web_indices = [web_hdrs.index(h)   for h in shared]
    else:
        shared      = excel_hdrs
        exc_indices = list(range(len(excel_hdrs)))
        web_indices = list(range(len(web_hdrs)))

    min_rows = min(len(excel_rows), len(web_rows))
    for row_i in range(min_rows):
        er = excel_rows[row_i]
        wr = web_rows[row_i]
        for j, col_name in enumerate(shared):
            ev = er[exc_indices[j]] if exc_indices[j] < len(er) else ""
            wv = wr[web_indices[j]] if web_indices[j] < len(wr) else ""
            if ev != wv:
                result.mismatches.append(
                    Mismatch(row=row_i + 1, column=col_name, excel_value=ev, web_value=wv)
                )
                log.debug("Mismatch Row %d, Col '%s': Excel=%r  Web=%r", row_i + 1, col_name, ev, wv)

    if result.mismatches or result.col_mismatch or result.excel_extra_rows or result.web_extra_rows:
        result.status = "FAIL"

    log.info("Comparison: %s | %d mismatch(es)", result.status, result.mismatch_count)
    return result
