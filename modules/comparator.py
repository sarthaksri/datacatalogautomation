import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from openpyxl.utils import get_column_letter

log = logging.getLogger(__name__)


def _norm(val) -> str:
    """Normalise a cell value: convert to str and strip whitespace."""
    return "" if val is None else str(val).strip()


def _values_equal(a: str, b: str) -> bool:
    """
    True if two cell values represent the same content.

    Exact string match wins. If both sides parse as numbers, compare
    numerically so trailing-zero formatting differences ("104" vs "104.00",
    "1.5" vs "1.50", "-4.8" vs "-4.80") aren't flagged as mismatches.
    Non-numeric values fall back to plain string equality.
    """
    if a == b:
        return True
    try:
        return float(a.replace(",", "")) == float(b.replace(",", ""))
    except (ValueError, AttributeError):
        return False


def _row_label(row: List[str], col_idx: int) -> str:
    """
    Build a short identifier for a data row, drawn from the cells preceding
    the mismatched column. Empty cells are skipped; we keep at most two
    leading values (typically code + name). Examples:
        ['4', 'Housing, water, electricity, gas and other fuels', '103.22']
            with col_idx=3 → "4 - Housing, water, electricity, gas and other fuels"
        ['', '', 'Rural', ...] with col_idx=4 → "Rural"
    """
    bits = [v for v in row[:col_idx] if v][:2]
    label = " - ".join(bits)
    return label[:120] + ("..." if len(label) > 120 else "")


def _trim_trailing_empty_cols(
    hdrs: List[str], rows: List[List[str]]
) -> Tuple[List[str], List[List[str]]]:
    """
    Drop trailing columns whose header is empty AND every row cell is empty.
    Excel exports stop at the real column count; Datawrapper CSVs often pad
    with trailing empties. Trimming both sides lets a positional compare work.
    """
    n = len(hdrs)
    while n > 0:
        last = n - 1
        if hdrs[last]:
            break
        if any(last < len(r) and r[last] for r in rows):
            break
        n -= 1
    return hdrs[:n], [r[:n] for r in rows]


@dataclass
class Mismatch:
    row: int          # 1-indexed data row number
    column: str       # column header name (or positional label for unnamed cols)
    excel_value: str
    web_value: str
    excel_cell: str = ""   # A1-style reference into the source Excel (e.g. "C7")
    row_label: str = ""    # leading identifier of the row, e.g. "4 - Housing..."


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
    excel_rows = [[_norm(c) for c in r] for r in excel_data["rows"]]
    web_rows   = [[_norm(c) for c in r] for r in web_data["rows"]]

    # 0-indexed row position of the Excel header within the source file —
    # data row i (0-indexed) lives at Excel row (header_row_idx + 2 + i).
    header_row_idx: int = int(excel_data.get("header_row_idx", 0))

    excel_hdrs, excel_rows = _trim_trailing_empty_cols(excel_hdrs, excel_rows)
    web_hdrs,   web_rows   = _trim_trailing_empty_cols(web_hdrs,   web_rows)

    # ── Column comparison ────────────────────────────────────────────────────
    if excel_hdrs != web_hdrs:
        log.warning("Column mismatch:\n  Excel: %s\n  Web:   %s", excel_hdrs, web_hdrs)
        result.col_mismatch = True
        # Report by named headers only; empty headers are merged-cell artefacts.
        ex_named = [h for h in excel_hdrs if h]
        wb_named = [h for h in web_hdrs   if h]
        result.missing_cols = [h for h in ex_named if h not in wb_named]
        result.extra_cols   = [h for h in wb_named if h not in ex_named]

    result.total_rows = max(len(excel_rows), len(web_rows))
    result.total_cols = max(len(excel_hdrs), len(web_hdrs))

    if len(excel_rows) > len(web_rows):
        result.excel_extra_rows = len(excel_rows) - len(web_rows)
        log.warning("Excel has %d extra rows vs web table", result.excel_extra_rows)
    elif len(web_rows) > len(excel_rows):
        result.web_extra_rows = len(web_rows) - len(excel_rows)
        log.warning("Web table has %d extra rows vs Excel", result.web_extra_rows)

    # ── Pick column mapping ──────────────────────────────────────────────────
    # When col counts match, compare positionally — merged-cell headers leave
    # multiple "" header slots that can't be disambiguated by name.
    # When they differ, fall back to name-based matching on non-empty headers.
    if len(excel_hdrs) == len(web_hdrs):
        n_cols = len(excel_hdrs)
        col_labels = [
            excel_hdrs[i] or web_hdrs[i] or f"col_{i + 1}"
            for i in range(n_cols)
        ]
        exc_indices = list(range(n_cols))
        web_indices = list(range(n_cols))
    else:
        shared = [h for h in excel_hdrs if h and h in web_hdrs]
        col_labels  = shared
        exc_indices = [excel_hdrs.index(h) for h in shared]
        web_indices = [web_hdrs.index(h)   for h in shared]

    # ── Cell-by-cell comparison ──────────────────────────────────────────────
    min_rows = min(len(excel_rows), len(web_rows))
    for row_i in range(min_rows):
        er = excel_rows[row_i]
        wr = web_rows[row_i]
        # Excel rows are 1-indexed; data row 0 sits at header_row_idx + 2.
        excel_row_num = header_row_idx + 2 + row_i
        for j, col_name in enumerate(col_labels):
            exc_idx = exc_indices[j]
            ev = er[exc_idx]         if exc_idx        < len(er) else ""
            wv = wr[web_indices[j]]  if web_indices[j] < len(wr) else ""
            if not _values_equal(ev, wv):
                cell_ref = f"{get_column_letter(exc_idx + 1)}{excel_row_num}"
                label    = _row_label(er, exc_idx)
                result.mismatches.append(
                    Mismatch(
                        row=row_i + 1,
                        column=col_name,
                        excel_value=ev,
                        web_value=wv,
                        excel_cell=cell_ref,
                        row_label=label,
                    )
                )
                log.debug(
                    "Mismatch %s [%s] Col %r: Excel=%r  Web=%r",
                    cell_ref, label, col_name, ev, wv,
                )

    if result.mismatches or result.col_mismatch or result.excel_extra_rows or result.web_extra_rows:
        result.status = "FAIL"

    log.info("Comparison: %s | %d mismatch(es)", result.status, result.mismatch_count)
    return result
