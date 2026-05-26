"""
MoSPI Data Catalogue Automation
================================
Navigates every dataset on the catalogue page, iterates all table rows
(across all pages), downloads each Excel file, reads the View Table data,
and produces a cell-level comparison report.

Usage:
    python main.py
"""

import logging
import sys
import time
from datetime import datetime
from pathlib import Path

import config
from modules.browser import BrowserManager
from modules.catalogue_page import CataloguePage
from modules.comparator import compare
from modules.dataset_page import DatasetPage
from modules.excel_handler import ExcelHandler
from modules.reporter import Reporter
from modules.table_viewer import TableViewer


# ── Logging ──────────────────────────────────────────────────────────────────

def _setup_logging() -> logging.Logger:
    config.LOG_DIR.mkdir(parents=True, exist_ok=True)
    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    fmt = "%(asctime)s [%(levelname)-8s] %(name)s: %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(config.LOG_DIR / f"run_{ts}.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("main")


# ── Helpers ──────────────────────────────────────────────────────────────────

def _make_dataset_page(page, download_dir: Path) -> DatasetPage:
    return DatasetPage(page=page, download_dir=download_dir)


def _make_table_viewer(page) -> TableViewer:
    return TableViewer(page=page, debug_dump_dir=config.SCREENSHOT_DIR)


# ── Core loop ────────────────────────────────────────────────────────────────

def _process_row(
    bm: BrowserManager,
    dp: DatasetPage,
    row_index: int,
    row_meta: dict,
    dataset_url: str,
    reporter: Reporter,
    dataset_name: str,
    log: logging.Logger,
) -> None:
    """
    Process one card: download Excel, open View Table, scrape it, compare.
    Nothing should escape this function; on any error we record the failure
    and force-navigate back to dataset_url so the next card can proceed.
    """
    table_name = row_meta.get("table_name", f"Table_{row_index}")
    table_no   = row_meta.get("table_no",   str(row_index + 1))
    label      = f"[{table_no}] {table_name}"
    log.info("    ▶ %s", label)

    excel_path = None
    web_data   = None

    # ── 1. Download Excel ────────────────────────────────────────────────
    try:
        excel_path = dp.download_excel(row_index, table_name, table_no)
    except Exception as exc:
        log.error("      ✗ Download crashed: %s", exc, exc_info=True)

    time.sleep(0.5)  # let MUI/React settle after the click

    # ── 2. Open View Table and scrape it ─────────────────────────────────
    try:
        dp.click_view_table(row_index)
        tv = _make_table_viewer(bm.page)
        web_data = tv.get_all_data()
    except Exception as exc:
        log.error("      ✗ View Table failed: %s", exc, exc_info=True)
        bm.screenshot(f"err_{dataset_name[:30]}_{table_no}_view", config.SCREENSHOT_DIR)

    # ── 3. Return to the dataset listing (explicit goto = reliable) ──────
    try:
        if "/tableview/" in bm.page.url or bm.page.url != dataset_url:
            bm.page.goto(dataset_url, wait_until="networkidle")
        dp.wait_for_cards()
    except Exception as exc:
        log.error("      ✗ Failed to return to listing: %s", exc)

    # ── 4. Read Excel & compare ──────────────────────────────────────────
    try:
        excel_data = ExcelHandler.read(excel_path) if excel_path else None
        result = compare(excel_data, web_data)
        log.info(
            "      → %s  |  %d mismatch(es)  |  %d rows",
            result.status, result.mismatch_count, result.total_rows,
        )
        reporter.add_result(dataset_name, row_meta, result)
    except Exception as exc:
        log.error("      ✗ Compare/report failed: %s", exc, exc_info=True)
        reporter.add_error(dataset_name, label, f"compare failed: {exc}")


def run() -> None:
    for d in [config.DOWNLOAD_DIR, config.REPORT_DIR, config.LOG_DIR, config.SCREENSHOT_DIR]:
        d.mkdir(parents=True, exist_ok=True)

    log = _setup_logging()
    log.info("=" * 70)
    log.info("MoSPI Data Catalogue Automation — starting")
    log.info("=" * 70)

    reporter = Reporter()

    try:
        with BrowserManager(
            headless=config.HEADLESS,
            download_dir=config.DOWNLOAD_DIR,
            slow_mo=config.SLOW_MO,
        ) as bm:
            bm.page.set_default_timeout(config.BROWSER_TIMEOUT)

            # ── Load catalogue ────────────────────────────────────────────
            catalogue = CataloguePage(
                page=bm.page,
                url=config.CATALOGUE_URL,
                card_selectors=config.CATALOGUE_CARD_SELECTORS,
            )
            catalogue.load()
            datasets = catalogue.get_datasets()

            if not datasets:
                log.error(
                    "No dataset cards found. Check CATALOGUE_CARD_SELECTORS in config.py.\n"
                    "Tip: run the script with HEADLESS=False and inspect the page."
                )
                return

            log.info("Found %d dataset(s) on catalogue page", len(datasets))

            # ── Iterate datasets ──────────────────────────────────────────
            for ds_idx, ds_info in enumerate(datasets):
                ds_name = ds_info["name"]
                log.info("")
                log.info("━" * 70)
                log.info("Dataset %d/%d: %s", ds_idx + 1, len(datasets), ds_name)
                log.info("━" * 70)

                try:
                    # Navigate back to catalogue if needed
                    if config.CATALOGUE_URL not in bm.page.url:
                        catalogue.reload()

                    catalogue.click_dataset_by_index(ds_idx)

                    dp = _make_dataset_page(bm.page, config.DOWNLOAD_DIR)
                    dp.wait_for_table()

                    total_pages = dp.get_total_pages()
                    log.info("  Pages of tables: %d", total_pages)

                    ds_page_num = 1
                    while True:
                        log.info("  — Dataset page %d/%d —", ds_page_num, total_pages)

                        rows = dp.get_rows()
                        if not rows:
                            log.warning("  No rows found, skipping page")
                            break

                        dataset_url = bm.page.url  # remember for recovery

                        for row_idx, row_meta in enumerate(rows):
                            _process_row(
                                bm, dp, row_idx, row_meta,
                                dataset_url, reporter, ds_name, log,
                            )
                            # Re-wait for table after returning from View Table
                            try:
                                dp.wait_for_table()
                            except Exception:
                                pass

                        if dp.has_next_page():
                            dp.go_to_next_page()
                            dp.wait_for_table()
                            ds_page_num += 1
                        else:
                            break

                except Exception as exc:
                    log.error("Error processing dataset '%s': %s", ds_name, exc, exc_info=True)
                    bm.screenshot(f"err_dataset_{ds_idx}", config.SCREENSHOT_DIR)

                finally:
                    # Always try to return to catalogue for the next iteration
                    try:
                        catalogue.reload()
                        time.sleep(1)
                    except Exception:
                        pass

    except Exception as exc:
        log.critical("Fatal error: %s", exc, exc_info=True)

    finally:
        reporter.generate(config.REPORT_DIR)
        log.info("Automation complete.")


if __name__ == "__main__":
    run()
