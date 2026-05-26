import logging
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

log = logging.getLogger(__name__)


class ExcelHandler:
    """Reads a downloaded Excel file (first sheet only) into a normalised dict."""

    @staticmethod
    def read(file_path: Path) -> Optional[Dict]:
        """
        Returns {"headers": [...], "rows": [[...], ...]} or None on error.
        All values are returned as stripped strings.

        MoSPI Excel files begin with a title row plus a blank row before the
        real header. We read with no header and pick the first row with more
        than one non-empty cell — the same heuristic the web-side CSV parser
        uses, so Excel and web align row-for-row.
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

        header_idx = 0
        for i in range(len(df)):
            non_empty = sum(1 for v in df.iloc[i] if v)
            if non_empty > 1:
                header_idx = i
                break

        headers: List[str] = [str(c).strip() for c in df.iloc[header_idx].tolist()]
        rows: List[List[str]] = [list(r) for r in df.iloc[header_idx + 1:].values.tolist()]
        while rows and not any(c for c in rows[-1]):
            rows.pop()

        log.info(
            "Excel read: %d rows × %d cols from %s (header row %d)",
            len(rows), len(headers), file_path.name, header_idx,
        )
        # header_row_idx is the 0-indexed row in the source file where headers
        # live. The comparator uses it to compute exact A1 cell references.
        return {"headers": headers, "rows": rows, "header_row_idx": header_idx}
