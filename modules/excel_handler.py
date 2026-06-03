import logging
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from modules.header_detect import pick_header_idx, is_serial_label_row

log = logging.getLogger(__name__)


class ExcelHandler:
    """Reads a downloaded Excel file (first sheet only) into a normalised dict."""

    @staticmethod
    def read(file_path: Path) -> Optional[Dict]:
        """
        Returns {"headers": [...], "rows": [[...], ...]} or None on error.
        All values are returned as stripped strings.

        MoSPI Excel files often have a multi-row header: a sparse "sub-title"
        row (e.g. just "At Current Prices" / "At Constant Prices" in two cells)
        sitting above the real column header (Sl.No., Industry, year columns).
        We pick the header by finding the first numeric data row, then walking
        back to the nearest row whose fill ratio matches the data — see
        modules.header_detect.pick_header_idx for the algorithm.
        """
        if not file_path or not file_path.exists():
            log.error("Excel file not found: %s", file_path)
            return None

        df = None
        for engine in (None, "xlrd"):
            try:
                kwargs = {"sheet_name": 0, "header": None, "dtype": str}
                if engine:
                    kwargs["engine"] = engine
                df = pd.read_excel(file_path, **kwargs)
                break
            except Exception as e:
                log.debug("Read failed with engine=%s: %s", engine, e)
        if df is None:
            log.error("Could not read Excel file: %s", file_path)
            return None

        df = df.fillna("").astype(str)
        for col in df.columns:
            df[col] = df[col].str.strip()

        all_rows: List[List[str]] = [list(r) for r in df.values.tolist()]
        header_idx = pick_header_idx(all_rows)

        headers: List[str] = list(all_rows[header_idx]) if all_rows else []
        rows: List[List[str]] = all_rows[header_idx + 1:]
        while rows and not any(c for c in rows[-1]):
            rows.pop()

        # NSS tables carry a "(1) (2) (3)…" column-number row right under the
        # header. Excel stores it as -1,-2,-3… so pick_header_idx couldn't see
        # it as a header and left it as the first data row. Drop it and treat
        # it as part of the header (bump the index so A1 cell refs stay right),
        # matching how the web table consumes that row as its header.
        if rows and is_serial_label_row(rows[0]):
            rows = rows[1:]
            header_idx += 1

        log.info(
            "Excel read: %d rows × %d cols from %s (header row %d)",
            len(rows), len(headers), file_path.name, header_idx,
        )
        # header_row_idx is the 0-indexed row in the source file where headers
        # live. The comparator uses it to compute exact A1 cell references.
        return {"headers": headers, "rows": rows, "header_row_idx": header_idx}
