"""
MoSPI Data Catalogue Automation
================================
Navigates every dataset on the catalogue page, iterates all table rows
(across all pages), downloads each Excel file, reads the View Table data,
and produces a cell-level comparison report.

Usage:
    python main.py             # fresh run, every dataset
    python main.py --continue  # resume the most recent report
    python main.py --IIP       # only datasets whose name contains "IIP"
    python main.py --only CPI  # same idea, long form
    python main.py --from CPI  # CPI and every dataset listed after it
"""

import argparse
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import List, Optional

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
) -> "Path | None":
    """
    Process one card: download Excel, open View Table, scrape it, compare.
    Nothing should escape this function; on any error we record the failure
    and force-navigate back to dataset_url so the next card can proceed.

    Returns the path of the downloaded Excel (or None on download failure)
    so the caller can clean it up once the dataset is finished.
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

    return excel_path


def _pick_report_path(continue_run: bool, log: logging.Logger) -> Path:
    """
    Decide where this run's report goes. With --continue, reuse the newest
    report in REPORT_DIR so its prior entries / Completed sheet are loaded
    by Reporter. Otherwise mint a new timestamped filename.
    """
    if continue_run:
        existing = sorted(
            config.REPORT_DIR.glob("comparison_report_*.xlsx"),
            key=lambda p: p.stat().st_mtime,
        )
        if existing:
            log.info("Continuing previous run from: %s", existing[-1])
            return existing[-1]
        log.info("--continue requested but no prior report found; starting fresh")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return config.REPORT_DIR / f"comparison_report_{ts}.xlsx"


def run(
    continue_run: bool = False,
    rerun_datasets: Optional[List[str]] = None,
    only: Optional[str] = None,
    start_from: Optional[str] = None,
) -> None:
    for d in [config.DOWNLOAD_DIR, config.REPORT_DIR, config.LOG_DIR, config.SCREENSHOT_DIR]:
        d.mkdir(parents=True, exist_ok=True)

    log = _setup_logging()
    log.info("=" * 70)
    log.info("MoSPI Data Catalogue Automation — starting (continue=%s)", continue_run)
    log.info("=" * 70)

    report_path = _pick_report_path(continue_run, log)
    reporter = Reporter(report_path=report_path)

    rerun_datasets = rerun_datasets or []
    for name in rerun_datasets:
        removed = reporter.drop_dataset(name)
        log.info(
            "Dropped %d existing entries for dataset %r - will reprocess from scratch",
            removed, name,
        )
    if rerun_datasets:
        reporter.save()  # persist the drops immediately

    skip_datasets = reporter.completed_datasets()
    skip_tables   = reporter.completed_tables()
    if skip_datasets or skip_tables:
        log.info(
            "Resume state: %d completed dataset(s), %d already-processed table(s)",
            len(skip_datasets), len(skip_tables),
        )

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

            # The catalogue carries an aggregate "ALL DATASETS" card that just
            # re-lists every other dataset's tables. Comparing it duplicates
            # thousands of tables (and inflates the report), so always skip it.
            before = len(datasets)
            datasets = [
                d for d in datasets
                if not d["name"].strip().upper().startswith("ALL DATASETS")
            ]
            if len(datasets) < before:
                log.info("Skipping the aggregate 'ALL DATASETS' card")

            if not datasets:
                log.error(
                    "No dataset cards found. Check CATALOGUE_CARD_SELECTORS in config.py.\n"
                    "Tip: run the script with HEADLESS=False and inspect the page."
                )
                return

            log.info("Found %d dataset(s) on catalogue page", len(datasets))

            if only:
                needle = only.lower()
                filtered = [d for d in datasets if needle in d["name"].lower()]
                if not filtered:
                    log.error(
                        "--only %r matched no datasets. Available: %s",
                        only, [d["name"] for d in datasets],
                    )
                    return
                log.info(
                    "Filter --only %r matched %d/%d dataset(s): %s",
                    only, len(filtered), len(datasets), [d["name"] for d in filtered],
                )
                datasets = filtered

            if start_from:
                needle = start_from.lower()
                start_idx = next(
                    (i for i, d in enumerate(datasets) if needle in d["name"].lower()),
                    None,
                )
                if start_idx is None:
                    log.error(
                        "--from %r matched no datasets. Available: %s",
                        start_from, [d["name"] for d in datasets],
                    )
                    return
                tail = datasets[start_idx:]
                log.info(
                    "Filter --from %r matched at position %d — running %d dataset(s): %s",
                    start_from, start_idx, len(tail), [d["name"] for d in tail],
                )
                datasets = tail

            # ── Iterate datasets ──────────────────────────────────────────
            for ds_idx, ds_info in enumerate(datasets):
                ds_name = ds_info["name"]
                log.info("")
                log.info("━" * 70)
                log.info("Dataset %d/%d: %s", ds_idx + 1, len(datasets), ds_name)
                log.info("━" * 70)

                if ds_name in skip_datasets:
                    log.info("✓ Already fully processed in prior run — skipping")
                    continue

                downloaded: list[Path] = []  # cleaned up after this dataset
                dataset_finished_cleanly = False

                try:
                    # Navigate back to catalogue if needed
                    if config.CATALOGUE_URL not in bm.page.url:
                        catalogue.reload()

                    catalogue.click_dataset_by_index(ds_info["index"])

                    dp = _make_dataset_page(bm.page, config.DOWNLOAD_DIR)
                    dp.wait_for_table()

                    total_pages = dp.get_total_pages()
                    log.info("  Pages of tables: %d", total_pages)

                    ds_page_num = 1
                    pagination_broke = False
                    while True:
                        log.info("  — Dataset page %d/%d —", ds_page_num, total_pages)

                        rows = dp.get_rows()
                        if not rows:
                            log.warning(
                                "  No rows on page %d (of %d expected) — "
                                "pagination likely failed mid-rerender; will retry once",
                                ds_page_num, total_pages,
                            )
                            time.sleep(2)
                            rows = dp.get_rows()
                            if not rows:
                                log.error(
                                    "  Still no rows after retry — stopping; "
                                    "dataset will NOT be marked complete so --continue can resume."
                                )
                                pagination_broke = True
                                break

                        dataset_url = bm.page.url  # remember for recovery

                        for row_idx, row_meta in enumerate(rows):
                            table_no = row_meta.get("table_no", str(row_idx + 1))
                            if (ds_name, table_no) in skip_tables:
                                log.info("    ▶ [%s] — already processed, skipping", table_no)
                                continue

                            excel_path = _process_row(
                                bm, dp, row_idx, row_meta,
                                dataset_url, reporter, ds_name, log,
                            )
                            if excel_path:
                                downloaded.append(excel_path)
                            # Re-wait for table after returning from View Table
                            try:
                                dp.wait_for_table()
                            except Exception:
                                pass

                        # Checkpoint at page boundaries (≤10 tables of work
                        # at risk on crash). Cheap enough; pages have at most
                        # ~10 tables, so this is roughly 10× less frequent
                        # than per-table saves.
                        reporter.save()

                        if dp.has_next_page():
                            if not dp.go_to_next_page():
                                log.error(
                                    "  go_to_next_page() failed at page %d/%d — "
                                    "stopping; dataset will NOT be marked complete.",
                                    ds_page_num, total_pages,
                                )
                                pagination_broke = True
                                break
                            ds_page_num += 1
                        else:
                            break

                    # Only mark complete if we actually visited every expected page.
                    # Catches the case where MUI silently leaves us on the same
                    # page or has_next_page() flakes early.
                    if pagination_broke:
                        dataset_finished_cleanly = False
                    elif ds_page_num < total_pages:
                        log.error(
                            "  Stopped at page %d but expected %d — NOT marking "
                            "dataset complete so --continue can resume.",
                            ds_page_num, total_pages,
                        )
                        dataset_finished_cleanly = False
                    else:
                        dataset_finished_cleanly = True

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

                    # Delete this dataset's Excel downloads — they've already
                    # been compared and the report holds the diffs.
                    removed = 0
                    for p in downloaded:
                        try:
                            p.unlink(missing_ok=True)
                            removed += 1
                        except Exception as exc:
                            log.warning("Could not delete %s: %s", p, exc)
                    if downloaded:
                        log.info("Cleaned up %d/%d Excel file(s) for dataset", removed, len(downloaded))

                    if dataset_finished_cleanly:
                        reporter.mark_dataset_complete(ds_name)
                        log.info("✓ Dataset complete: %s", ds_name)
                    reporter.save()

    except Exception as exc:
        log.critical("Fatal error: %s", exc, exc_info=True)

    finally:
        reporter.generate(config.REPORT_DIR)
        log.info("Automation complete.")


_KNOWN_FLAGS = {"--continue", "--rerun", "--only", "--from", "-h", "--help"}


def _normalize_argv(argv: List[str]) -> List[str]:
    """Translate shorthand flags like --IIP, --CPI, --PLFS to --only IIP.

    Anything matching --<alnum> that isn't a known flag and isn't followed
    by '=' is rewritten. Leaves --only/--rerun/--continue alone.
    """
    out: List[str] = []
    for a in argv:
        if (
            a.startswith("--")
            and "=" not in a
            and a not in _KNOWN_FLAGS
            and re.fullmatch(r"--[A-Za-z][A-Za-z0-9_-]*", a)
        ):
            out.extend(["--only", a[2:]])
        else:
            out.append(a)
    return out


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MoSPI Data Catalogue Automation")
    parser.add_argument(
        "--continue", dest="continue_run", action="store_true",
        help="Resume the most recent report, skipping completed datasets/tables.",
    )
    parser.add_argument(
        "--rerun", dest="rerun_datasets", action="append", default=[],
        metavar="DATASET_NAME",
        help='Drop existing entries for this dataset and reprocess it. '
             'May be passed multiple times to redo several datasets in one run. '
             'Examples: --rerun "INDEX OF INDUSTRIAL PRODUCTION IIP" '
             '--rerun "CONSUMER PRICE INDEX CPI". Use with --continue.',
    )
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--only", dest="only", default=None, metavar="SUBSTRING",
        help='Run only datasets whose name contains SUBSTRING '
             '(case-insensitive). Shorthand: --IIP / --CPI / --PLFS etc. '
             'all map to --only <name>.',
    )
    group.add_argument(
        "--from", dest="start_from", default=None, metavar="SUBSTRING",
        help='Start at the first dataset whose name contains SUBSTRING '
             '(case-insensitive) and run it plus every dataset after it. '
             'Example: --from CPI runs CPI and all datasets listed after CPI.',
    )
    args = parser.parse_args(_normalize_argv(sys.argv[1:]))
    run(
        continue_run=args.continue_run,
        rerun_datasets=args.rerun_datasets,
        only=args.only,
        start_from=args.start_from,
    )
