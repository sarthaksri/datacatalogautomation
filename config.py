from pathlib import Path

BASE_URL = "https://esankhyiki.mospi.gov.in"
CATALOGUE_URL = f"{BASE_URL}/catalogue-main"

# Browser
HEADLESS = False
SLOW_MO = 200        # ms between actions — keeps the site from throttling
BROWSER_TIMEOUT = 60_000  # ms

# Filesystem
DOWNLOAD_DIR = Path("downloads")
REPORT_DIR = Path("reports")
LOG_DIR = Path("logs")
SCREENSHOT_DIR = Path("screenshots")

# ── Selectors ────────────────────────────────────────────────────────────────
# If the site changes structure, update these lists (first match wins).

# Catalogue page: clickable KPI/dataset cards
CATALOGUE_CARD_SELECTORS = [
    "a.card",
    "[class*='card'] a[href]",
    "[class*='dataset'] a[href]",
    "[class*='kpi'] a[href]",
    "a[href*='catalogue']",
    "a[href*='dataset']",
]

# Dataset detail page: table rows and navigation
DATASET_ROW_SEL       = "table tbody tr"
DATASET_HEADER_SEL    = "table thead th, table thead td"
DATASET_NEXT_PAGE_SEL = [
    "li.active + li:not(.disabled) a",
    ".pagination .next a",
    "a[aria-label='Next']",
    "a[aria-label='Next page']",
    "button:has-text('Next')",
    ".page-item:not(.disabled) a:has-text('›')",
    ".page-item:not(.disabled) a:has-text('»')",
]

# Excel download button (searched inside each table row)
DOWNLOAD_BTN_SELECTORS = [
    "a[href*='.xlsx']",
    "a[href*='.xls']",
    "a[href*='download']",
    "a[download]",
    "a:has-text('Download')",
    "a:has-text('Excel')",
    "button:has-text('Download')",
]

# View Table button (searched inside each table row)
VIEW_TABLE_SELECTORS = [
    "a:has-text('View Table')",
    "a:has-text('View')",
    "button:has-text('View Table')",
    "button:has-text('View')",
    "a[href*='table']",
    "a[href*='view']",
]

# View Table page pagination
VIEW_NEXT_PAGE_SEL = [
    "li.active + li:not(.disabled) a",
    ".pagination .next a",
    "a[aria-label='Next']",
    "a[aria-label='Next page']",
    "button:has-text('Next')",
    ".page-item:not(.disabled) a:has-text('›')",
    ".page-item:not(.disabled) a:has-text('»')",
]

# ── Column keyword mapping ────────────────────────────────────────────────────
# Maps our field names to keywords to look for in table headers (case-insensitive).
# Extend if the site uses different header text.
COL_KEYWORDS = {
    "table_no":      ["table no", "tbl no", "sl", "sr", "s.no", "s no", "sno", "serial"],
    "table_name":    ["table name", "title", "name", "description"],
    "product":       ["product"],
    "category":      ["category", "sector", "subject"],
    "release_date":  ["release date", "released", "date"],
}
