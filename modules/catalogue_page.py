import logging
import time
from typing import List, Dict
from playwright.sync_api import Page

log = logging.getLogger(__name__)


class CataloguePage:
    """Handles the /catalogue-main page — finds and clicks dataset KPI cards."""

    def __init__(self, page: Page, url: str, card_selectors: List[str]):
        self.page = page
        self.url = url
        self.card_selectors = card_selectors
        self._found_selector: str = None

    def load(self) -> None:
        log.info("Loading catalogue: %s", self.url)
        self.page.goto(self.url, wait_until="networkidle")
        self.page.wait_for_load_state("domcontentloaded")
        time.sleep(1)

    def get_datasets(self) -> List[Dict]:
        """Return metadata for every KPI card on the page."""
        datasets: List[Dict] = []

        for sel in self.card_selectors:
            cards = self.page.locator(sel).all()
            if not cards:
                continue

            log.info("Found %d dataset card(s) with selector: %s", len(cards), sel)
            self._found_selector = sel

            for i, card in enumerate(cards):
                try:
                    name = card.inner_text().strip().replace("\n", " ")[:120]
                    href = card.get_attribute("href") or ""
                    datasets.append({"index": i, "name": name, "href": href})
                except Exception as e:
                    log.debug("Could not read card %d: %s", i, e)

            if datasets:
                break

        if not datasets:
            log.warning(
                "No dataset cards found. Adjust CATALOGUE_CARD_SELECTORS in config.py. "
                "Current page title: %s",
                self.page.title(),
            )
        return datasets

    def click_dataset_by_index(self, index: int) -> None:
        """Click the dataset card at the given list position (0-based)."""
        for sel in self.card_selectors:
            cards = self.page.locator(sel).all()
            if len(cards) > index:
                label = cards[index].inner_text().strip()[:60]
                log.info("Clicking dataset %d: %s", index, label)
                cards[index].click()
                self.page.wait_for_load_state("networkidle")
                time.sleep(1)
                return
        raise RuntimeError(f"Dataset card at index {index} not found")

    def reload(self) -> None:
        self.load()
