import logging
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, List, Optional, Tuple

from openpyxl.utils import get_column_letter

log = logging.getLogger(__name__)

_MONTH_NAMES = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4,  "may": 5,  "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

# Tokens that mean "no value" in MoSPI tables. Excel typically leaves the
# cell blank; the Datawrapper-rendered web table sometimes literally
# writes "(null)" or "NA" instead. Treat all of these as equivalent.
_NULL_TOKENS = {"", "(null)", "null", "none", "n/a", "na", "-", "--",
                "–", "—"}  # en-dash, em-dash


def _is_null(s: str) -> bool:
    return s.strip().lower() in _NULL_TOKENS
_ISO_DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})(?:[ T]\d{2}:\d{2}:\d{2})?$")
_MMM_YY_RE   = re.compile(r"^([A-Za-z]{3})[-\s/]?(\d{2,4})$")


def _to_month_year(s: str) -> Optional[Tuple[int, int]]:
    """
    Parse a string as (year, month) if it looks like a month-year token.
    Recognises 'Apr-12', 'Apr-2012', 'Apr 12' and ISO timestamps like
    '2012-04-11 00:00:00' (any day - we ignore it). Returns None otherwise.

    This is the bridge between pandas' stringified-datetime headers from
    Excel ('2012-04-11 00:00:00') and the human-friendly 'Apr-12' the
    Datawrapper CSV uses for monthly columns on IIP tables.
    """
    s = s.strip()
    m = _ISO_DATE_RE.match(s)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = _MMM_YY_RE.match(s)
    if m:
        mname = m.group(1).lower()
        if mname in _MONTH_NAMES:
            year = int(m.group(2))
            if year < 100:
                year += 2000 if year < 50 else 1900
            return year, _MONTH_NAMES[mname]
    return None


def _decimals(s: str) -> int:
    """Count significant decimal places in a numeric string (ignores trailing zeros)."""
    s = s.replace(",", "").strip()
    if "." not in s:
        return 0
    return len(s.split(".", 1)[1].rstrip("0"))


def _round_half_up(num_str: str, ndigits: int) -> Optional[Decimal]:
    """
    Round a numeric string to `ndigits` decimals using HALF_UP (away from zero).
    Necessary because Python's built-in round() uses banker's rounding,
    which would round '0.79845' to '0.7984' while spreadsheets and the
    MoSPI web tables both round to '0.7985'.
    """
    try:
        d = Decimal(num_str.replace(",", "").strip())
    except (InvalidOperation, AttributeError):
        return None
    quant = Decimal("1") if ndigits <= 0 else Decimal(10) ** -ndigits
    return d.quantize(quant, rounding=ROUND_HALF_UP)


def _norm(val) -> str:
    """Normalise a cell value: convert to str and strip whitespace."""
    return "" if val is None else str(val).strip()


def _looks_like_header_row(row: List[str]) -> bool:
    """True if a row's non-empty cells are mostly non-numeric (labels, months, etc.)."""
    non_empty = [v for v in row if v]
    if not non_empty:
        return False
    non_numeric = 0
    for v in non_empty:
        try:
            float(v.replace(",", "").replace("%", ""))
        except (ValueError, AttributeError):
            non_numeric += 1
    return non_numeric / len(non_empty) > 0.5


def _values_equal(a: str, b: str) -> bool:
    """
    True if two cell values represent the same content.

    Layers, in order:
      1. Exact string match.
      2. Month-year tokens: 'Apr-12' equals '2012-04-11 00:00:00' (the
         day part of pandas-stringified Excel dates is irrelevant for
         the monthly columns MoSPI uses on IIP tables).
      3. Strict numeric equality: '104' equals '104.00', '-4.8' equals
         '-4.80'.
      4. Rounded-precision tolerance: when one side carries higher
         precision than the other (and the lower side is >= 2 dp), they
         agree if the higher rounds to the lower within half its last
         digit. Catches IIP weights where Excel keeps full precision
         like '5.302468' but the web table renders '5.3025'.
      5. Otherwise, not equal.
    """
    if a == b:
        return True

    # Both sides represent "no data" (empty cell, "(null)", "NA", "-", etc.)
    if _is_null(a) and _is_null(b):
        return True

    ma, mb = _to_month_year(a), _to_month_year(b)
    if ma is not None and mb is not None:
        return ma == mb

    try:
        na = float(a.replace(",", ""))
        nb = float(b.replace(",", ""))
    except (ValueError, AttributeError):
        return False
    if na == nb:
        return True

    da, db = _decimals(a), _decimals(b)
    if da == db:
        return False  # same precision, values genuinely differ
    target_dp = min(da, db)
    ra = _round_half_up(a, target_dp)
    rb = _round_half_up(b, target_dp)
    if ra is None or rb is None:
        return False
    return ra == rb


def _headers_equal(a: List[str], b: List[str]) -> bool:
    """Same as `a == b` but using `_values_equal` per slot (so 'Apr-12' == '2012-04-11 00:00:00')."""
    if len(a) != len(b):
        return False
    return all(_values_equal(x, y) for x, y in zip(a, b))


_PLACEHOLDER_HDR_RE = re.compile(r"^\(?[+-]?\d+\)?$")


def _is_placeholder_header(h: str) -> bool:
    """True for a non-informative header slot: empty, or a bare column number.

    NSS tables label their columns ``(1) (2) (3)…``; the web side keeps that
    numbered row as its header while Excel keeps the *named* sub-header
    (``Male/Female/Person`` etc.). When the two tables have the same number of
    columns the comparison is positional, so a numbered placeholder on one side
    against a real name on the other is not a genuine column conflict.
    """
    return not h or bool(_PLACEHOLDER_HDR_RE.match(h))


def _headers_match_positional(a: List[str], b: List[str]) -> bool:
    """Header equality for equal-length (positional) comparison.

    Slots agree when their values are equal *or* either side is a placeholder
    (empty / column-number). Lets a numbered web header line up with a named
    Excel header without raising a false column mismatch.
    """
    if len(a) != len(b):
        return False
    return all(
        _values_equal(x, y) or _is_placeholder_header(x) or _is_placeholder_header(y)
        for x, y in zip(a, b)
    )


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


_YEAR_RANGE_RE = re.compile(r"^(\d{4})\s*[-–/]\s*(\d{2,4})$")


def _canon_key(s: str) -> str:
    """Canonicalise a row-key cell for matching.

    Fiscal-year ranges are written inconsistently across MoSPI tables and
    even within one ('1999-2000', '2003-04', '2003-2004'). The start year
    identifies the range uniquely, so collapse any 'YYYY-YY'/'YYYY-YYYY' to
    just the start year. Everything else is lower-cased and stripped.
    """
    s = s.strip()
    m = _YEAR_RANGE_RE.match(s)
    if m:
        return m.group(1)
    return s.lower()


def _drop_empty_rows(rows: List[List[str]]) -> List[List[str]]:
    return [r for r in rows if any(c for c in r)]


def _key_align(
    excel_rows: List[List[str]], web_rows: List[List[str]],
    ekey: int, wkey: int,
) -> Optional[List[Tuple[Optional[int], Optional[int]]]]:
    """Align rows by their first-column key instead of by position.

    Returns a list of (excel_idx | None, web_idx | None) pairs, or None when
    the first column isn't a reliable key on both sides (then the caller
    falls back to positional comparison). A clean key column turns "web is
    missing one row in the middle" into a single extra-row finding rather
    than a cascade of off-by-one cell mismatches.
    """
    e_keys = [_canon_key(r[ekey]) if ekey < len(r) else "" for r in excel_rows]
    w_keys = [_canon_key(r[wkey]) if wkey < len(r) else "" for r in web_rows]

    def good(keys: List[str]) -> bool:
        non_empty = [k for k in keys if k]
        return (
            len(keys) >= 3
            and len(non_empty) >= 0.9 * len(keys)
            and len(set(non_empty)) == len(non_empty)   # unique
        )

    if not (good(e_keys) and good(w_keys)):
        return None

    w_index = {k: j for j, k in enumerate(w_keys) if k}
    used_w = set()
    pairs: List[Tuple[Optional[int], Optional[int]]] = []
    for i, k in enumerate(e_keys):
        if k and k in w_index:
            j = w_index[k]
            pairs.append((i, j))
            used_w.add(j)
        else:
            pairs.append((i, None))   # row only in Excel
    for j, k in enumerate(w_keys):
        if j not in used_w:
            pairs.append((None, j))   # row only on web
    return pairs


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

    # ── Sub-header sanity check ──────────────────────────────────────────────
    # MoSPI tables often have a sub-header row right under the main header
    # (e.g. "July 24 Index | July 25 Index | Inflation" for the Cereals
    # row of the YoY table). When the Datawrapper CSV behind the iframe
    # is a stale version, that sub-header reads "June 24 / June 25" while
    # the downloaded Excel says "July 24 / July 25" — every data row then
    # appears mismatched because the time periods differ. Catch this and
    # surface a single clear error instead of hundreds of false cell rows.
    if excel_rows and web_rows:
        r0_e = excel_rows[0]
        r0_w = web_rows[0]
        min_len = min(len(r0_e), len(r0_w))
        if (
            min_len >= 4
            and _looks_like_header_row(r0_e)
            and _looks_like_header_row(r0_w)
        ):
            diffs = sum(
                1 for i in range(min_len)
                if not _values_equal(r0_e[i], r0_w[i])
            )
            if diffs >= 2 and diffs / min_len >= 0.25:
                log.warning(
                    "Sub-header row differs in %d/%d cells — skipping per-cell "
                    "comparison. Excel head: %s | Web head: %s",
                    diffs, min_len, r0_e[:6], r0_w[:6],
                )
                result.status = "ERROR"
                result.error = (
                    f"Sub-header content differs ({diffs}/{min_len} cells) — "
                    f"likely a time-period or table-version mismatch between "
                    f"the downloaded Excel and the web table. "
                    f"Excel row 0 starts: {r0_e[:6]}; Web row 0 starts: {r0_w[:6]}"
                )
                result.total_rows = max(len(excel_rows), len(web_rows))
                result.total_cols = max(len(excel_hdrs), len(web_hdrs))
                return result

    # ── Column comparison ────────────────────────────────────────────────────
    # When the column counts match the comparison is positional, so tolerate
    # placeholder ("(1) (2)…") headers on one side against real names on the
    # other. Only different counts or genuinely conflicting names are a problem.
    if len(excel_hdrs) == len(web_hdrs):
        headers_ok = _headers_match_positional(excel_hdrs, web_hdrs)
    else:
        headers_ok = False
    if not headers_ok:
        log.warning("Column mismatch:\n  Excel: %s\n  Web:   %s", excel_hdrs, web_hdrs)
        result.col_mismatch = True
        # Report by named headers only; empty headers are merged-cell artefacts.
        # Use semantic equality so 'Apr-12' isn't listed as missing just because
        # the Excel side stringified it as '2012-04-11 00:00:00'.
        ex_named = [h for h in excel_hdrs if h]
        wb_named = [h for h in web_hdrs   if h]
        result.missing_cols = [h for h in ex_named if not any(_values_equal(h, w) for w in wb_named)]
        result.extra_cols   = [h for h in wb_named if not any(_values_equal(h, e) for e in ex_named)]

    # Drop fully-empty rows on both sides (Datawrapper pads with blank rows;
    # Excel exports sometimes carry blank spacer rows) so they don't skew the
    # row-count diff or the key-uniqueness check below.
    excel_rows = _drop_empty_rows(excel_rows)
    web_rows   = _drop_empty_rows(web_rows)

    result.total_rows = max(len(excel_rows), len(web_rows))
    result.total_cols = max(len(excel_hdrs), len(web_hdrs))

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

    # ── Row alignment ─────────────────────────────────────────────────────────
    # Prefer matching rows by their first-column key (year, state, item…) so a
    # single missing/extra row in the middle is reported once, not as an
    # off-by-one cascade of false cell mismatches. Fall back to positional
    # pairing when the first column isn't a clean key on both sides.
    ekey = exc_indices[0] if exc_indices else 0
    wkey = web_indices[0] if web_indices else 0
    aligned = _key_align(excel_rows, web_rows, ekey, wkey)

    if aligned is None:
        min_rows = min(len(excel_rows), len(web_rows))
        aligned = [(i, i) for i in range(min_rows)]
        if len(excel_rows) > len(web_rows):
            result.excel_extra_rows = len(excel_rows) - len(web_rows)
        elif len(web_rows) > len(excel_rows):
            result.web_extra_rows = len(web_rows) - len(excel_rows)
    else:
        result.excel_extra_rows = sum(1 for e, w in aligned if w is None)
        result.web_extra_rows   = sum(1 for e, w in aligned if e is None)

    if result.excel_extra_rows:
        log.warning("Excel has %d row(s) with no web match", result.excel_extra_rows)
    if result.web_extra_rows:
        log.warning("Web table has %d row(s) with no Excel match", result.web_extra_rows)

    # ── Cell-by-cell comparison over aligned pairs ────────────────────────────
    for ei, wj in aligned:
        if ei is None or wj is None:
            continue  # unmatched row — counted as extra above, no cell diffs
        er = excel_rows[ei]
        wr = web_rows[wj]
        # Excel rows are 1-indexed; data row 0 sits at header_row_idx + 2.
        excel_row_num = header_row_idx + 2 + ei
        for j, col_name in enumerate(col_labels):
            exc_idx = exc_indices[j]
            ev = er[exc_idx]         if exc_idx        < len(er) else ""
            wv = wr[web_indices[j]]  if web_indices[j] < len(wr) else ""
            if not _values_equal(ev, wv):
                cell_ref = f"{get_column_letter(exc_idx + 1)}{excel_row_num}"
                label    = _row_label(er, exc_idx)
                result.mismatches.append(
                    Mismatch(
                        row=ei + 1,
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
