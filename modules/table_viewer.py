import csv
import json
import logging
import re
import time
from io import StringIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from playwright.sync_api import Page

from modules.header_detect import pick_header_idx

log = logging.getLogger(__name__)


class TableViewer:
    """
    Reads data from the MoSPI 'View Table' page.

    Page DOM:
        <div class="page-body">
            <h5>{table title}</h5>
            <div class="embedLoader-iframe">
                <iframe src="https://datawrapper.dwcdn.net/{publicId}/{version}/" ...>
            </div>
        </div>

    The iframe is a published Datawrapper chart. The chart's full table is
    served as a static CSV at:
        https://datawrapper.dwcdn.net/{publicId}/{publicVersion}/dataset.csv

    Both publicId and publicVersion are baked into the iframe URL itself
    (e.g. /8jBsB/2/), and also into a __DW_SVELTE_PROPS__ JSON blob inside
    the iframe HTML — we use the latter as a fallback when the URL doesn't
    include a version segment.

    Reading the CSV bypasses pagination, the sticky-header twin <table>,
    and the heatmap-style visualization <table> that lives in the same DOM.
    DOM scraping is kept as a last-resort fallback.
    """

    IFRAME_SEL   = "div.embedLoader-iframe iframe"
    BACK_BTN_SEL = "button[title='Previous page']"
    TITLE_SEL    = "div.page-body > h5"

    # Captures (publicId, version) from URLs like
    # https://datawrapper.dwcdn.net/8jBsB/2/  →  ("8jBsB", "2")
    # If the version segment is missing we still capture the id.
    _CHART_ID_VERSION_RE = re.compile(
        r"datawrapper\.dwcdn\.net/([A-Za-z0-9]+)(?:/(\d+))?/?"
    )

    # __DW_SVELTE_PROPS__ extraction: the value is a JSON-encoded string of
    # another JSON object. We pull the inner quoted string and double-decode.
    _SVELTE_PROPS_MARKER = 'window.__DW_SVELTE_PROPS__ = JSON.parse("'

    def __init__(self, page: Page, debug_dump_dir: Optional[Path] = None):
        self.page = page
        self.debug_dump_dir = debug_dump_dir
        self._dumped = False

    # ── Iframe / chart-id discovery ──────────────────────────────────────────

    def wait_for_iframe(self, timeout_ms: int = 30_000) -> bool:
        try:
            self.page.wait_for_selector(self.IFRAME_SEL, timeout=timeout_ms)
            time.sleep(0.3)
            return True
        except Exception as e:
            log.warning("View Table iframe not found: %s", e)
            return False

    def _get_iframe_src(self) -> Optional[str]:
        try:
            return self.page.locator(self.IFRAME_SEL).first.get_attribute("src")
        except Exception as e:
            log.error("Could not read iframe src: %s", e)
            return None

    def _parse_id_and_version(self, src: str) -> Tuple[Optional[str], Optional[str]]:
        m = self._CHART_ID_VERSION_RE.search(src or "")
        if not m:
            return None, None
        return m.group(1), m.group(2)

    def _extract_version_from_iframe_html(self, iframe_src: str) -> Optional[str]:
        """
        Fetch the iframe's HTML and pull publicVersion out of
        window.__DW_SVELTE_PROPS__. Used when the iframe URL doesn't carry
        a version segment.
        """
        try:
            resp = self.page.context.request.get(iframe_src, timeout=30_000)
            if resp.status != 200:
                return None
            html = resp.text()
        except Exception as e:
            log.debug("Could not fetch iframe HTML for version lookup: %s", e)
            return None

        i = html.find(self._SVELTE_PROPS_MARKER)
        if i == -1:
            return None
        # Walk the JSON-string literal until its unescaped closing quote
        j = i + len(self._SVELTE_PROPS_MARKER)
        out = []
        while j < len(html):
            ch = html[j]
            if ch == "\\":
                out.append(html[j:j + 2])
                j += 2
                continue
            if ch == '"':
                break
            out.append(ch)
            j += 1
        try:
            inner = json.loads('"' + "".join(out) + '"')
            props = json.loads(inner)
            version = props.get("chart", {}).get("publicVersion")
            return str(version) if version is not None else None
        except Exception as e:
            log.debug("Could not parse __DW_SVELTE_PROPS__: %s", e)
            return None

    def get_chart_coords(self) -> Tuple[Optional[str], Optional[str]]:
        """Return (publicId, publicVersion) — falls back to JSON if needed."""
        src = self._get_iframe_src()
        if not src:
            return None, None
        chart_id, version = self._parse_id_and_version(src)
        if not chart_id:
            log.error("iframe src does not match Datawrapper pattern: %s", src)
            return None, None
        if not version:
            log.info("Version missing from iframe URL — extracting from JSON props")
            version = self._extract_version_from_iframe_html(src)
        log.info("Datawrapper chart: id=%s version=%s", chart_id, version)
        return chart_id, version

    def get_title(self) -> str:
        try:
            return self.page.locator(self.TITLE_SEL).first.inner_text().strip()
        except Exception:
            return ""

    # ── Data fetch ───────────────────────────────────────────────────────────

    def fetch_dataset_csv(self, chart_id: str, version: Optional[str]) -> Optional[str]:
        """Fetch the raw CSV/TSV string from Datawrapper's CDN."""
        if version:
            urls = [f"https://datawrapper.dwcdn.net/{chart_id}/{version}/dataset.csv"]
        else:
            # Last-ditch attempts when we couldn't find a version anywhere
            urls = [
                f"https://datawrapper.dwcdn.net/{chart_id}/1/dataset.csv",
                f"https://datawrapper.dwcdn.net/{chart_id}/dataset.csv",
            ]
        for url in urls:
            log.info("Fetching Datawrapper dataset: %s", url)
            try:
                resp = self.page.context.request.get(
                    url,
                    headers={
                        "Referer": "https://datawrapper.dwcdn.net/",
                        "Accept": "text/csv, text/plain, */*",
                    },
                    timeout=30_000,
                )
                if resp.status == 200:
                    return resp.text()
                log.warning("Dataset fetch %s → HTTP %d", url, resp.status)
            except Exception as e:
                log.warning("Dataset fetch %s → %s", url, e)
        return None

    @staticmethod
    def _parse_csv(text: str) -> Dict:
        """Parse CSV/TSV text into {'headers': [...], 'rows': [[...], ...]}.

        Datawrapper payloads lead with a single-cell banner row (the chart
        title) and, for NSS/NAS tables, several sparse sub-title rows above
        the real column header. We must NOT let pandas sniff the delimiter
        from that first line: on banner tables it guesses "1 column" and then
        raises "Expected 1 fields, saw N" on the first wide data row, which
        used to drop us into the lossy DOM-scrape fallback (banner mistaken
        for the header, paginated row subset, Datawrapper-formatted display
        values like '6.0' for '6021').

        Parse with the stdlib csv reader instead: it tolerates ragged rows,
        and we choose the delimiter by counting candidates across the whole
        payload (comma for dataset.csv, tab for the .tsv variant). The shared
        header picker in modules.header_detect then locates the real header.
        """
        if not text or not text.strip():
            return {"headers": [], "rows": []}

        delim = "\t" if text.count("\t") > text.count(",") else ","
        try:
            raw = [list(r) for r in csv.reader(StringIO(text), delimiter=delim)]
        except Exception as e:
            log.error("CSV parse failed: %s", e)
            return {"headers": [], "rows": []}

        # Rows are ragged — the banner row has one cell, data rows have many.
        # Pad every row to the widest so positional column alignment holds.
        width = max((len(r) for r in raw), default=0)
        all_rows = [[c.strip() for c in r] + [""] * (width - len(r)) for r in raw]
        if not all_rows:
            return {"headers": [], "rows": []}

        header_idx = pick_header_idx(all_rows)
        headers = list(all_rows[header_idx])
        rows = all_rows[header_idx + 1:]
        # Drop trailing all-empty rows
        while rows and not any(c for c in rows[-1]):
            rows.pop()

        return {"headers": headers, "rows": rows}

    # ── Iframe scraping fallback ─────────────────────────────────────────────

    def _dump_iframe_html(self, frame) -> None:
        if not self.debug_dump_dir or self._dumped:
            return
        try:
            html = frame.locator("body").first.evaluate("el => el.outerHTML")
            self.debug_dump_dir.mkdir(parents=True, exist_ok=True)
            ts = time.strftime("%Y%m%d_%H%M%S")
            path = self.debug_dump_dir / f"iframe_{ts}.html"
            path.write_text(html, encoding="utf-8")
            log.info("Dumped iframe HTML → %s", path)
            self._dumped = True
        except Exception as e:
            log.debug("Iframe HTML dump failed: %s", e)

    def _pick_real_table(self, frame) -> Optional[str]:
        """
        Pick the real data <table> from the iframe.

        Datawrapper renders multiple tables in the same iframe:
          - a 'heatmap' visualization table: many rows, only ONE <td> per row
          - the real interactive table: class contains 'resortable', and
            has a <caption> like "Table with N columns and M rows."

        Strategy: enumerate every <table>; prefer one whose class contains
        'resortable'; otherwise pick the one with the widest data rows
        (max columns) so we skip the 1-column heatmap.
        """
        try:
            tables = frame.locator("table").evaluate_all("""
                els => els.map(el => ({
                    html: el.outerHTML,
                    cls:  el.getAttribute('class') || '',
                    rows: el.querySelectorAll('tr').length,
                    maxCols: Math.max(0, ...Array.from(el.querySelectorAll('tr'))
                        .map(tr => tr.children.length))
                }))
            """)
        except Exception as e:
            log.error("Could not enumerate iframe tables: %s", e)
            return None

        if not tables:
            return None

        # 1. resortable wins
        for t in tables:
            if "resortable" in t["cls"]:
                log.info("Picked 'resortable' table (rows=%d, cols=%d)", t["rows"], t["maxCols"])
                return t["html"]

        # 2. otherwise: widest table that has >= 2 columns
        wide = [t for t in tables if t["maxCols"] >= 2]
        if wide:
            best = max(wide, key=lambda t: (t["maxCols"], t["rows"]))
            log.info("Picked widest table (rows=%d, cols=%d)", best["rows"], best["maxCols"])
            return best["html"]

        log.warning("No suitable iframe table found (only single-column visualisations)")
        return None

    def _scrape_iframe_table(self) -> Dict:
        """Last-resort: parse the rendered iframe <table>."""
        try:
            frame = self.page.frame_locator(self.IFRAME_SEL)
            frame.locator("table").first.wait_for(state="attached", timeout=30_000)
            time.sleep(1.5)
            self._dump_iframe_html(frame)

            html = self._pick_real_table(frame)
            if not html:
                return {"headers": [], "rows": []}

            dfs = pd.read_html(StringIO(html))
            if not dfs:
                return {"headers": [], "rows": []}
            df = dfs[0].fillna("").astype(str)
            for col in df.columns:
                df[col] = df[col].str.strip()
            return {
                "headers": [str(c).strip() for c in df.columns],
                "rows":    [list(r) for r in df.values.tolist()],
            }
        except Exception as e:
            log.error("Iframe scrape failed: %s", e, exc_info=True)
            return {"headers": [], "rows": []}

    # ── Public API ───────────────────────────────────────────────────────────

    def get_all_data(self) -> Dict:
        """Returns {"headers": [...], "rows": [[...], ...]}.

        Tries Datawrapper's static dataset.csv first (full table, no
        pagination). Falls back to scraping the iframe <table> if the CSV
        endpoint can't be reached.
        """
        if not self.wait_for_iframe():
            return {"headers": [], "rows": []}

        chart_id, version = self.get_chart_coords()
        if chart_id:
            csv_text = self.fetch_dataset_csv(chart_id, version)
            if csv_text:
                data = self._parse_csv(csv_text)
                if data["rows"]:
                    log.info(
                        "View Table data (CSV): %d rows × %d cols",
                        len(data["rows"]), len(data["headers"]),
                    )
                    return data

        log.info("Datawrapper CSV unavailable — falling back to DOM scrape")
        return self._scrape_iframe_table()

    # ── Navigation back ──────────────────────────────────────────────────────

    def go_back(self) -> bool:
        try:
            btn = self.page.locator(self.BACK_BTN_SEL).first
            if btn.count() > 0 and btn.is_visible():
                btn.click()
                self.page.wait_for_load_state("networkidle")
                time.sleep(0.5)
                return True
        except Exception:
            pass
        return False
