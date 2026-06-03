"""Shared header-row picker used by ExcelHandler and TableViewer.

The MoSPI catalogue mixes simple tables (one header row above data) with
NAS-style multi-row headers — a sparse 'sub-title' row sits above the real
column header, e.g.::

    Row 0:  "At Current Prices ₹ crore"  (in one merged cell)
    Row 1:  ""  ...  "At Current Prices"  ""  ...  "At Constant Prices"  ""
    Row 2:  "Sl.No."  "Industry"  "NIC"  "2011-12"  "2012-13"  ...
    Row 3:  1  Food  10  150411  141414  ...

The naive "first row with >1 non-empty cell" picks Row 1 (two sub-title
labels), leaves Row 2 inside the data, and every comparison after that is
column-misaligned. Pick the data row first, then walk back to the row
that actually spans the data's columns.
"""
from typing import List


def _numeric_cell_count(row: List[str], cap: int = 2) -> int:
    n = 0
    for v in row:
        s = str(v).strip()
        if not s:
            continue
        try:
            float(s.replace(",", "").replace("%", ""))
        except ValueError:
            continue
        n += 1
        if n >= cap:
            return n
    return n


def _non_empty_count(row: List[str]) -> int:
    return sum(1 for v in row if str(v).strip())


def is_serial_label_row(row: List[str]) -> bool:
    """True if a row is an NSS-style column-number key row, e.g. ``(1) (2) (3)…``.

    These rows sit directly below the real header on NSS survey tables and
    number every column 1..n. The web (Datawrapper) table treats this row as
    its header, but Excel stores ``(1)`` as the accounting-negative ``-1`` so
    ``pick_header_idx`` sees a fully-numeric row, picks the *named* sub-header
    above it, and leaves this row sitting in the data — a phantom extra row
    that misaligns the comparison. Detect it so callers can drop it.

    A match requires at least three non-empty cells, every one of them a bare
    integer (optionally signed or wrapped in parens), whose absolute values are
    exactly the consecutive sequence 1, 2, 3, … across the row. A genuine data
    row never numbers its own columns left-to-right like that.
    """
    non_empty = [str(v).strip() for v in row if str(v).strip()]
    if len(non_empty) < 3:
        return False
    vals: List[int] = []
    for cell in non_empty:
        t = cell.lstrip("(").rstrip(")").strip().lstrip("+-")
        if not t.isdigit():
            return False
        vals.append(int(t))
    return vals == list(range(1, len(vals) + 1))


def pick_header_idx(rows: List[List[str]], scan_limit: int = 50) -> int:
    """Return the index of the row to treat as the column header.

    Strategy:
      1. Find the first row with >=2 numeric cells — that's the first data row.
      2. Walk back to the nearest row above it whose non-empty count is at
         least half the data row's non-empty count. That's the real header
         (multi-row sub-titles are sparse and get skipped).
      3. If no data row is found, fall back to the first row with >1 non-empty
         cell.

    `scan_limit` caps how deep we look — MoSPI headers always sit in the top
    few dozen rows, even for tables that are thousands of rows long.
    """
    if not rows:
        return 0

    n = min(len(rows), scan_limit)

    first_data = None
    for i in range(n):
        if _numeric_cell_count(rows[i]) >= 2:
            first_data = i
            break

    if first_data is None:
        for i in range(n):
            if _non_empty_count(rows[i]) > 1:
                return i
        return 0

    data_cols = _non_empty_count(rows[first_data])
    threshold = max(2, (data_cols + 1) // 2)

    for i in range(first_data - 1, -1, -1):
        if _non_empty_count(rows[i]) >= threshold:
            return i

    return max(0, first_data - 1)
