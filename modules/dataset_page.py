import logging
import re
import time
from pathlib import Path
from typing import Dict, List, Optional

from playwright.sync_api import Download, Page

log = logging.getLogger(__name__)


class DatasetPage:
    """
    Handles the MoSPI catalogue dataset listing page.

    Real DOM structure (per card):
        <div class="catalog-grid">
            <h5 class="title">{table name}</h5>
            <ul class="gridTable">
                <li><p><span>Product:</span> IIP</p></li>
                <li><p><span>Category:</span> Industrial Statistics</p></li>
                <li><p><span>Geography:</span> All India</p></li>
                <li><p><span>Frequency:</span> Monthly</p></li>
                <li><p><span>Reference Period:</span> Mar 2026</p></li>
                <li><p><span>Data Source:</span> ...</p></li>
                <li class="product-desc"><p><span>Description:</span> ...</p></li>
                <li class="product-id"><p><span>Release Date:</span> 28 Apr 2026</p></li>
                <li class="product-id"><p><span>Table No.</span> IIPMFY25045MAR</p></li>
            </ul>
            <div class="grid-action">
                <button>View Table</button>
                <button>View Chart</button>
                <a href="/datacatalogue/.../*.xlsx" download="">...</a>
            </div>
        </div>

    Pagination (MUI):
        <p class="MuiTablePagination-displayedRows">1–10 of 436</p>
        <button aria-label="Go to next page">...</button>
    """

    CARD_SEL          = "div.catalog-grid"
    TITLE_SEL         = "h5.title"
    INFO_ITEM_SEL     = "ul.gridTable > li"
    DOWNLOAD_SEL      = "div.grid-action a[download]"
    VIEW_TABLE_SEL    = "div.grid-action button:has-text('View Table')"

    PAGINATION_TEXT   = "p.MuiTablePagination-displayedRows"
    NEXT_PAGE_BTN     = "button[aria-label='Go to next page']"
    PREV_PAGE_BTN     = "button[aria-label='Go to previous page']"

    # `<li>` prefix → meta-dict field name
    FIELD_PREFIXES = {
        "Product:":      "product",
        "Category:":     "category",
        "Release Date:": "release_date",
        "Table No.":     "table_no",
    }

    def __init__(self, page: Page, download_dir: Path):
        self.page = page
        self.download_dir = download_dir

    # ── Waiting ───────────────────────────────────────────────────────────────

    def wait_for_cards(self) -> None:
        self.page.wait_for_selector(self.CARD_SEL, timeout=30_000)
        time.sleep(0.5)

    # Back-compat alias for main.py
    wait_for_table = wait_for_cards

    # ── Pagination ────────────────────────────────────────────────────────────

    def _pagination_text(self) -> str:
        try:
            return self.page.locator(self.PAGINATION_TEXT).first.inner_text().strip()
        except Exception:
            return ""

    def get_total_records(self) -> int:
        """Parse '1–10 of 436' → 436."""
        m = re.search(r"of\s+([\d,]+)", self._pagination_text())
        return int(m.group(1).replace(",", "")) if m else 0

    def get_total_pages(self, rows_per_page: int = 10) -> int:
        total = self.get_total_records()
        if total <= 0:
            return 1
        return (total + rows_per_page - 1) // rows_per_page

    def _is_disabled(self, locator) -> bool:
        try:
            if locator.count() == 0:
                return True
            cls = locator.first.get_attribute("class") or ""
            return "Mui-disabled" in cls or locator.first.is_disabled()
        except Exception:
            return True

    def has_next_page(self) -> bool:
        return not self._is_disabled(self.page.locator(self.NEXT_PAGE_BTN))

    def go_to_next_page(self) -> bool:
        """Click Next and wait for MUI to actually advance the page.

        The previous version waited on `networkidle` + a fixed 1s sleep,
        which often returned while React was still mid-rerender. The next
        iteration's get_rows() then read either stale cards or zero cards,
        and the outer loop bailed out thinking the dataset was finished.

        Now we capture the pagination text ("1–10 of 436") BEFORE clicking
        and poll until it changes — that's the only reliable signal that
        MUI has finished swapping the page.
        """
        if not self.has_next_page():
            log.info("No next page — at end of listing")
            return False

        before_text = self._pagination_text()
        try:
            self.page.locator(self.NEXT_PAGE_BTN).first.click()
        except Exception as e:
            log.warning("Next-page click failed: %s", e)
            return False

        deadline = time.time() + 20
        while time.time() < deadline:
            after_text = self._pagination_text()
            if after_text and after_text != before_text:
                break
            time.sleep(0.25)
        else:
            log.warning(
                "Pagination text did not change after Next click "
                "(stayed at %r) — page advance failed",
                before_text,
            )
            return False

        try:
            self.page.wait_for_load_state("networkidle", timeout=10_000)
        except Exception:
            pass
        try:
            self.wait_for_cards()
        except Exception as e:
            log.warning("Cards not present after Next: %s", e)
            return False

        log.info("Navigated to next listing page: %s", after_text)
        return True

    # ── Rows / metadata ───────────────────────────────────────────────────────

    def get_row_count(self) -> int:
        return self.page.locator(self.CARD_SEL).count()

    def get_rows(self) -> List[Dict]:
        """Return metadata dicts for every card on the current page."""
        cards = self.page.locator(self.CARD_SEL)
        n = cards.count()
        results: List[Dict] = []

        for i in range(n):
            card = cards.nth(i)

            try:
                title = card.locator(self.TITLE_SEL).first.inner_text().strip()
            except Exception:
                title = ""

            meta: Dict = {
                "row_index":    i,
                "table_name":   title,
                "product":      "",
                "category":     "",
                "release_date": "",
                "table_no":     "",
            }

            try:
                for li in card.locator(self.INFO_ITEM_SEL).all():
                    text = li.inner_text().strip()
                    for prefix, field in self.FIELD_PREFIXES.items():
                        if text.startswith(prefix):
                            meta[field] = text[len(prefix):].strip()
                            break
            except Exception as e:
                log.debug("Card %d: metadata parse failed: %s", i, e)

            results.append(meta)

        log.info("Read metadata for %d cards on this page", len(results))
        return results

    # ── Actions ──────────────────────────────────────────────────────────────

    # Real spreadsheets start with one of these magic byte sequences:
    #   PK\x03\x04  — zip container (.xlsx / .xlsm)
    #   \xD0\xCF\x11\xE0 — OLE2 compound file (.xls)
    # The MoSPI SPA sometimes hands back its index.html shell (starts with
    # '<') instead of the file, which lands on disk as a tiny HTML doc and
    # then fails to parse as "Excel file could not be read".
    _XLSX_MAGIC = b"PK\x03\x04"
    _XLS_MAGIC  = b"\xD0\xCF\x11\xE0"

    @classmethod
    def _looks_like_excel(cls, path: Path) -> bool:
        try:
            with open(path, "rb") as fh:
                head = fh.read(8)
        except OSError:
            return False
        return head.startswith(cls._XLSX_MAGIC) or head.startswith(cls._XLS_MAGIC)

    def download_excel(self, row_index: int, table_name: str, table_no: str) -> Optional[Path]:
        """
        Click the download link and wait for the file to land on disk.
        Uses Playwright's expect_download — the `download` attribute on the
        <a> prevents navigation, so the page stays put.

        The download is occasionally the SPA's HTML shell rather than the
        real spreadsheet (a known MoSPI flake). We validate the saved file's
        magic bytes and retry a few times before giving up, so a single bad
        download no longer turns into an "Excel file could not be read" error.

        After the download completes we re-wait for the cards before
        returning, so the caller can safely interact with View Table next.
        """
        safe = (f"{table_no}_{table_name}"[:80]
                .replace("/", "_").replace("\\", "_")
                .replace(" ", "_").replace(":", "_").replace("?", ""))
        dest = self.download_dir / f"{safe}.xlsx"

        attempts = 3
        for attempt in range(1, attempts + 1):
            card = self.page.locator(self.CARD_SEL).nth(row_index)
            link = card.locator(self.DOWNLOAD_SEL).first

            if link.count() == 0:
                log.error("No download link on card %d", row_index)
                return None

            try:
                log.info("Downloading Excel for [%s] %s (attempt %d/%d)",
                         table_no, table_name, attempt, attempts)
                try:
                    link.scroll_into_view_if_needed(timeout=5_000)
                except Exception:
                    pass

                with self.page.expect_download(timeout=90_000) as dl_info:
                    link.click(timeout=15_000)
                download: Download = dl_info.value
                download.save_as(str(dest))

                # Give the page a moment to settle (React/MUI re-render after click)
                try:
                    self.wait_for_cards()
                except Exception:
                    pass

                if self._looks_like_excel(dest):
                    log.info("Saved to: %s", dest)
                    return dest

                size = dest.stat().st_size if dest.exists() else 0
                log.warning(
                    "Download for [%s] is not a valid spreadsheet (%d bytes, "
                    "likely the HTML shell) — retrying",
                    table_no, size,
                )
                time.sleep(1.5)
            except Exception as e:
                log.error("Download attempt %d failed for card %d: %s",
                          attempt, row_index, e)
                try:
                    self.wait_for_cards()
                except Exception:
                    pass
                time.sleep(1.0)

        log.error("All %d download attempts for [%s] produced a non-Excel file",
                  attempts, table_no)
        return dest if dest.exists() else None

    def click_view_table(self, row_index: int) -> Optional[Page]:
        """
        Click the 'View Table' button. On MoSPI this always loads in the
        same tab at /catalogue-main/catalogue/tableview/{id}, so we drop the
        new-tab detection and just wait for the URL change.

        Always returns None (kept Optional[Page] for caller compatibility).
        """
        self.wait_for_cards()

        card = self.page.locator(self.CARD_SEL).nth(row_index)
        btn = card.locator(self.VIEW_TABLE_SEL).first

        log.info("Clicking View Table (card %d)", row_index)

        try:
            btn.wait_for(state="visible", timeout=15_000)
        except Exception as e:
            raise RuntimeError(f"View Table button not visible on card {row_index}: {e}")

        try:
            btn.scroll_into_view_if_needed(timeout=5_000)
        except Exception:
            pass

        try:
            btn.click(timeout=15_000)
        except Exception as e:
            log.warning("Normal click failed on card %d (%s) — retrying with force=True", row_index, e)
            btn.click(force=True, timeout=10_000)

        try:
            self.page.wait_for_url("**/tableview/**", timeout=30_000)
        except Exception:
            # Some installations may navigate differently; fall back to load-state
            self.page.wait_for_load_state("networkidle", timeout=15_000)

        time.sleep(0.5)
        log.info("View Table loaded: %s", self.page.url)
        return None
