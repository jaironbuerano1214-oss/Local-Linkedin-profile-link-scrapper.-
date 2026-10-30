"""
LinkedIn Profile Finder — Selenium Edition (No API keys required)
==================================================================
Reads names from 'list of names (1).xlsx', uses a real Chrome browser
(via Selenium) to search Google for each person's LinkedIn profile,
filters by DHVSU / Don Honorio / Pampanga State Univ OR Pampanga location
AND IT-related field, then writes the profile URL back to column C.

Setup:
    pip install -r requirements.txt
    (Chrome must be installed — chromedriver is auto-managed)

Usage:
    python linkedin_finder.py              # run all rows
    python linkedin_finder.py --dry-run    # search only, no writes
    python linkedin_finder.py --start-row 50
    python linkedin_finder.py --name "Reyes, John Vincent V."
    python linkedin_finder.py --visible    # show the browser window
"""

import re
import time
import random
import logging
import argparse
from pathlib import Path

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    WebDriverException,
)
from webdriver_manager.chrome import ChromeDriverManager

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("linkedin_finder.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
EXCEL_FILE = "list of names (1).xlsx"
SHEET_NAME = "Sheet1"
NAME_COL   = 1   # Column A
LINK_COL   = 3   # Column C (linkedin profile)

# Seconds to wait between searches (be a good citizen, avoid blocks)
MIN_DELAY = 4.0
MAX_DELAY = 8.0

# ---------------------------------------------------------------------------
# Filter keywords
# ---------------------------------------------------------------------------
SCHOOL_KEYWORDS = [
    "don honorio ventura",
    "dhvsu",
    "don honorio",
    "pampanga state university",
    "psu pampanga",
    "pamantasan ng lungsod ng san fernando",
    "don honorio ventura technological state university",
]

LOCATION_KEYWORDS = [
    "pampanga",
    "san fernando",
    "angeles city",
    "mabalacat",
    "magalang",
    "apalit",
    "bacolor",
    "candaba",
    "floridablanca",
    "guagua",
    "lubao",
    "macabebe",
    "masantol",
    "mexico",
    "minalin",
    "porac",
    "san luis",
    "san simon",
    "santa ana",
    "santa rita",
    "santo tomas",
    "sasmuan",
    "clark",
]

IT_KEYWORDS = [
    "computer science",
    "information technology",
    "information systems",
    "css",
    "bs cs",
    "bscs",
    "bs it",
    "bsit",
    "bs is",
    "bsis",
    "bachelor of science in information technology",
    "bachelor of science in computer science",
    "bachelor of science in information systems",
    "software",
    "developer",
    "programmer",
    "engineer",
    "web developer",
    "data analyst",
    "cybersecurity",
    "network",
    "systems analyst",
    "tech",
    "it specialist",
    "cloud",
    "devops",
    "full stack",
    "frontend",
    "backend",
    "mobile developer",
    "qa engineer",
    "ui/ux",
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.lower().strip())


def contains_any(text: str, keywords: list) -> bool:
    t = normalize(text)
    return any(kw in t for kw in keywords)


def name_to_query(raw_name: str) -> str:
    """'Lastname, Firstname M.' → 'Firstname Lastname'"""
    raw_name = raw_name.strip()
    if "," in raw_name:
        parts = raw_name.split(",", 1)
        last = parts[0].strip()
        first_parts = re.sub(r"\s+[A-Za-z]\.$", "", parts[1].strip()).strip()
        return f"{first_parts} {last}"
    return raw_name


def is_linkedin_url(url: str) -> bool:
    return "linkedin.com/in/" in url.lower()


def passes_filter(title: str, snippet: str) -> bool:
    """
    Returns True if text matches:
      (school OR location) AND IT field
    """
    combined = normalize(f"{title} {snippet}")
    geo_ok = contains_any(combined, SCHOOL_KEYWORDS) or contains_any(combined, LOCATION_KEYWORDS)
    it_ok  = contains_any(combined, IT_KEYWORDS)
    return geo_ok and it_ok


def passes_filter_relaxed(title: str, snippet: str) -> bool:
    """
    Relaxed version for broad Philippines-only queries:
    only requires IT field match (geo is unknown from snippet).
    """
    combined = normalize(f"{title} {snippet}")
    return contains_any(combined, IT_KEYWORDS)


def random_delay():
    t = random.uniform(MIN_DELAY, MAX_DELAY)
    log.debug("Waiting %.1fs...", t)
    time.sleep(t)

# ---------------------------------------------------------------------------
# Browser setup
# ---------------------------------------------------------------------------

def make_driver(headless: bool = True) -> webdriver.Chrome:
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
    options.add_argument("--window-size=1280,900")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    # Mask webdriver property
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return driver


# ---------------------------------------------------------------------------
# Google search via Selenium
# ---------------------------------------------------------------------------

def google_search(driver: webdriver.Chrome, query: str) -> list[dict]:
    """
    Performs a Google search and returns a list of
    {'title': ..., 'link': ..., 'snippet': ...} for results that are
    LinkedIn profile URLs.
    """
    try:
        driver.get("https://www.google.com")
        wait = WebDriverWait(driver, 10)

        # Accept cookies popup if it appears (EU/PH sometimes shows it)
        try:
            accept_btn = driver.find_element(By.XPATH, "//button[.//span[contains(text(),'Accept')]]")
            accept_btn.click()
            time.sleep(1)
        except NoSuchElementException:
            pass

        # Find search box and type query
        search_box = wait.until(EC.presence_of_element_located((By.NAME, "q")))
        search_box.clear()
        search_box.send_keys(query)
        search_box.send_keys(Keys.RETURN)

        # Wait for results
        wait.until(EC.presence_of_element_located((By.ID, "search")))
        time.sleep(1.5)

        results = []
        # Each organic result is typically a div with data-hveid or class 'g'
        result_blocks = driver.find_elements(By.CSS_SELECTOR, "div.g, div[data-hveid]")

        for block in result_blocks:
            try:
                # Get the anchor link
                anchor = block.find_element(By.CSS_SELECTOR, "a[href]")
                link = anchor.get_attribute("href") or ""
                if not is_linkedin_url(link):
                    continue

                # Title
                try:
                    title = block.find_element(By.CSS_SELECTOR, "h3").text
                except NoSuchElementException:
                    title = ""

                # Snippet
                try:
                    snippet = block.find_element(
                        By.CSS_SELECTOR, "div[data-sncf], div.VwiC3b, span.aCOpRe, div[style='-webkit-line-clamp:2']"
                    ).text
                except NoSuchElementException:
                    snippet = ""

                results.append({"title": title, "link": link, "snippet": snippet})
                log.debug("  Result: [%s] %s", title[:60], link)

            except (NoSuchElementException, WebDriverException):
                continue

        return results

    except TimeoutException:
        log.warning("Google search timed out for query: %s", query)
        return []
    except Exception as e:
        log.warning("Google search error: %s", e)
        return []


def check_captcha(driver: webdriver.Chrome) -> bool:
    """Returns True if Google is showing a CAPTCHA page."""
    page = driver.page_source.lower()
    return "captcha" in page or "are you a robot" in page or "unusual traffic" in page


# ---------------------------------------------------------------------------
# Core profile finder
# ---------------------------------------------------------------------------

def find_profile(driver: webdriver.Chrome, name: str) -> str | None:
    """
    Runs 3 progressively broader Google searches for the person,
    filters by school/location + IT field, returns first match URL.
    """
    search_name = name_to_query(name)

    # Short focused queries — Google rejects long boolean strings with no results.
    # Most specific first; last query is broadest (Philippines only, filter relaxed).
    queries = [
        (f'site:linkedin.com/in "{search_name}" Pampanga',                                    True),
        (f'site:linkedin.com/in "{search_name}" DHVSU',                                       True),
        (f'site:linkedin.com/in "{search_name}" "Don Honorio"',                               True),
        (f'site:linkedin.com/in "{search_name}" "Don Honorio Ventura"',                       True),
        (f'site:linkedin.com/in "{search_name}" "Don Honorio Ventura Technological"',         True),
        (f'site:linkedin.com/in "{search_name}" Philippines',                                 False),  # relaxed filter
    ]

    for q, strict in queries:
        log.info("  Query: %s", q)
        results = google_search(driver, q)

        if check_captcha(driver):
            log.warning("⚠ Google CAPTCHA detected! Pausing 60s then continuing...")
            time.sleep(60)
            results = google_search(driver, q)

        random_delay()

        for item in results:
            link    = item["link"]
            title   = item["title"]
            snippet = item["snippet"]

            ok = passes_filter(title, snippet) if strict else passes_filter_relaxed(title, snippet)
            if ok:
                log.info("  ✓ Match: %s", link)
                return link
            else:
                log.info("  ✗ Filtered: %s | %s | %s", link[:60], title[:40], snippet[:60])

    return None


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def run(dry_run: bool = False, start_row: int = 2, end_row: int = None, headless: bool = True):
    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists():
        log.error("Excel file not found: %s", excel_path)
        return

    wb = load_workbook(excel_path)
    ws = wb[SHEET_NAME]

    max_row = ws.max_row
    if end_row is None or end_row > max_row:
        end_row = max_row

    log.info("Processing rows %d – %d | dry_run=%s | headless=%s", start_row, end_row, dry_run, headless)

    driver = make_driver(headless=headless)
    log.info("Chrome started.")

    found_count   = 0
    skipped_count = 0

    try:
        for row_idx in range(start_row, end_row + 1):
            name_cell = ws.cell(row=row_idx, column=NAME_COL)
            link_cell = ws.cell(row=row_idx, column=LINK_COL)

            name = (name_cell.value or "").strip()
            if not name:
                skipped_count += 1
                continue

            existing = (link_cell.value or "").strip()
            if existing:
                log.info("Row %d [%s]: already has link, skipping.", row_idx, name)
                skipped_count += 1
                continue

            log.info("Row %d/%d — %s", row_idx, end_row, name)

            profile_url = find_profile(driver, name)

            if profile_url:
                log.info("  → Saved: %s", profile_url)
                found_count += 1
                if not dry_run:
                    link_cell.value = profile_url
                    wb.save(excel_path)   # save after every found row
            else:
                log.info("  → Not found.")

    except KeyboardInterrupt:
        log.info("Interrupted by user. Saving progress...")
        if not dry_run:
            wb.save(excel_path)
    finally:
        driver.quit()
        log.info("Chrome closed.")

    if not dry_run:
        wb.save(excel_path)
        log.info("Final save: '%s'", excel_path)

    log.info("Done. Found: %d | Skipped: %d", found_count, skipped_count)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="LinkedIn Finder — Selenium, no API keys")
    parser.add_argument("--dry-run",    action="store_true", help="Search only, don't write to Excel.")
    parser.add_argument("--visible",    action="store_true", help="Show the Chrome browser window.")
    parser.add_argument("--start-row",  type=int, default=2,    help="Row to start from (default: 2).")
    parser.add_argument("--end-row",    type=int, default=None,  help="Row to stop at (default: last).")
    parser.add_argument("--name",       type=str, default=None,  help="Test a single name.")
    args = parser.parse_args()

    if args.name:
        driver = make_driver(headless=not args.visible)
        try:
            url = find_profile(driver, args.name)
            print(f"\nResult for '{args.name}': {url or 'Not found / filtered out'}")
        finally:
            driver.quit()
    else:
        run(
            dry_run=args.dry_run,
            start_row=args.start_row,
            end_row=args.end_row,
            headless=not args.visible,
        )
