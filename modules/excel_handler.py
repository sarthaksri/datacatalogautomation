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
        """
        if not file_path or not file_path.exists():
            log.error("Excel file not found: %s", file_path)
            return None

        # Try openpyxl (xlsx) first, fall back to xlrd (xls)
        for engine in (None, "xlrd"):
            try:
                kwargs = {"sheet_name": 0, "header": 0, "dtype": str}
                if engine:
                    kwargs["engine"] = engine
                df = pd.read_excel(file_path, **kwargs)
                break
            except Exception as e:
                log.debug("Read failed with engine=%s: %s", engine, e)
        else:
            log.error("Could not read Excel file: %s", file_path)
            return None

        df = df.fillna("")
        df.columns = [str(c).strip() for c in df.columns]

        # Strip all cell values
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()

        headers: List[str] = list(df.columns)
        rows: List[List[str]] = [list(r) for r in df.values.tolist()]

        log.info("Excel read: %d rows × %d cols from %s", len(rows), len(headers), file_path.name)
        return {"headers": headers, "rows": rows}
