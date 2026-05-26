import logging
from pathlib import Path
from playwright.sync_api import sync_playwright, Browser, BrowserContext, Page

log = logging.getLogger(__name__)


class BrowserManager:
    """Playwright browser session as a context manager."""

    def __init__(self, headless: bool = True, download_dir: Path = Path("downloads"), slow_mo: int = 0):
        self.headless = headless
        self.download_dir = download_dir
        self.slow_mo = slow_mo
        self._pw = None
        self._browser: Browser = None
        self.context: BrowserContext = None
        self.page: Page = None

    def __enter__(self) -> "BrowserManager":
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(headless=self.headless, slow_mo=self.slow_mo)
        self.context = self._browser.new_context(
            accept_downloads=True,
            viewport={"width": 1920, "height": 1080},
        )
        self.page = self.context.new_page()
        self.page.set_default_timeout(60_000)
        log.info("Browser started (headless=%s)", self.headless)
        return self

    def __exit__(self, *args) -> None:
        if self.context:
            self.context.close()
        if self._browser:
            self._browser.close()
        if self._pw:
            self._pw.stop()
        log.info("Browser closed")

    def screenshot(self, name: str, screenshot_dir: Path = Path("screenshots")) -> None:
        screenshot_dir.mkdir(parents=True, exist_ok=True)
        path = screenshot_dir / f"{name}.png"
        try:
            self.page.screenshot(path=str(path), full_page=True)
            log.debug("Screenshot saved: %s", path)
        except Exception as e:
            log.warning("Screenshot failed: %s", e)
