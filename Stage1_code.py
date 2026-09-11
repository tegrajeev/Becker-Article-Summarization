# -*- coding: utf-8 -*- gbgr
"""
STAGE 1 — INCREMENTAL + MERGED + DELTA
HYBRID VERSION:
- Becker Payer sections -> requests
- Becker Hospital Review finance -> Selenium

What this version does:
1. Reads previous Stage-1 merged CSV
2. Uses latest published_dt in merged CSV as watermark
3. Crawls section pages incrementally
4. Dedupes merged rows by:
      (normalized_title + published_date_YYYYMMDD)
5. If same article appears in multiple sections,
   keeps them in ONE row with comma-separated unique values:
      - sources
      - sections
      - urls
6. Writes:
      - OUT_CSV   = merged master history
      - DELTA_CSV = only truly new rows from this run

Important behavior:
- Next week, if someone runs this again, it will read the same master CSV,
  get the latest published date, and fetch only newer rows.
- Finance archive is fetched with Selenium because requests gets blocked.

ANTI-BOT FIXES (v2):
- UA aligned to actual Chrome version (147)
- Stealth flags added to ChromeOptions
- navigator.webdriver hidden via JS
- Human-like scroll simulation after page load
- Randomized sleeps (wider range, no fixed values)
- Extended blocked-phrase list covers the exact error message seen
- requests warm-up starts from a Google referrer
"""

from __future__ import annotations

import csv
import os
import re
import time
from dataclasses import dataclass
from datetime import datetime, date
from typing import Optional, List, Dict, Tuple
from urllib.parse import urljoin, urlparse, parse_qsl, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup
from dateutil import parser as date_parser

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import undetected_chromedriver as uc
import random

# ============================
# CONFIG
# ============================
BASE_SECTIONS = [
    ("https://www.beckerspayer.com/contracting/", "Beckers Payer", "contracting"),
    ("https://www.beckerspayer.com/payer/", "Beckers Payer", "payer"),
    ("https://www.beckerspayer.com/payer/medicare-advantage/", "Beckers Payer", "medicare_advantage"),
    ("https://www.beckerspayer.com/payer/medicaid/", "Beckers Payer", "medicaid"),
    ("https://www.beckerspayer.com/policy-updates/", "Beckers Payer", "policy_updates"),
    ("https://www.beckerspayer.com/payer/aca/", "Beckers Payer", "aca"),
    ("https://www.beckershospitalreview.com/finance/", "Beckers Hospital Review", "finance"),
]

USE_SELENIUM_FOR_FINANCE = True
USE_SELENIUM_FOR_ALL = True


# ✅ FIX 12: Detect the installed Chrome major version at runtime so the
# requests User-Agent never drifts out of sync after a Chrome auto-update.
def detect_chrome_major_version(default: int = 150) -> int:
    """
    Reads the installed Chrome version from the Windows registry / common
    install paths. Falls back to `default` if detection fails.
    """
    # Windows registry (most reliable)
    try:
        import winreg
        for hive, path in (
            (winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Google\Chrome\BLBeacon"),
        ):
            try:
                with winreg.OpenKey(hive, path) as k:
                    ver, _ = winreg.QueryValueEx(k, "version")
                    return int(str(ver).split(".")[0])
            except OSError:
                continue
    except Exception:
        pass

    # Fallback: query the chrome.exe binary directly
    try:
        import subprocess
        for exe in (
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ):
            if os.path.exists(exe):
                out = subprocess.run(
                    ["powershell", "-NoProfile", "-Command",
                     f"(Get-Item '{exe}').VersionInfo.ProductVersion"],
                    capture_output=True, text=True, timeout=15,
                )
                v = (out.stdout or "").strip()
                if v:
                    return int(v.split(".")[0])
    except Exception:
        pass

    return default


CHROME_MAJOR = detect_chrome_major_version()
print(f"[init] Detected Chrome major version: {CHROME_MAJOR}")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    f"Chrome/{CHROME_MAJOR}.0.0.0 Safari/537.36"
)

HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Referer": "https://www.google.com/",
    "Upgrade-Insecure-Requests": "1",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
}

OUT_DIR = r"C:\Users\NancySamanta\OneDrive - TEG Business Solutions Pvt Ltd\Work Data\Becker Codes_Offline\OUT_PUT_2022tocurr"

OUT_CSV = os.path.join(OUT_DIR, "listings_2022tocurr_merged.csv")
DELTA_CSV = os.path.join(OUT_DIR, "listings_2022tocurr_new_delta.csv")

MAX_PAGES = 300
# ✅ FIX 2: Increased base sleep from 10 → 20 to reduce request rate on flagged IP
SLEEP_SEC = 20
TIMEOUT = 40
SELENIUM_WAIT_SEC = 60

# Used only if master file does not exist yet
DEFAULT_CUTOFF_DATE = date(2026, 1, 1)

# ✅ FIX 9: Stop a section after this many CONSECUTIVE pages where
# every dated article is older than (or equal to) the watermark.
# Listing pages aren't strictly sorted, so a single page's "newest"
# date isn't a reliable stop signal — but several pages in a row
# with nothing new is.
CONSECUTIVE_OLD_PAGES_TO_STOP = 3

# ✅ FIX 3: Extended list of blocked phrases covering the exact error message seen
BLOCKED_PHRASES = [
    "Please enable JS and disable any ad blocker",
    "Slide right to secure your access",
    "Verification Required",
    "unusual activity from your device",       # exact phrase from reported error
    "Automated (bot) activity",                 # exact phrase from reported error
    "JavaScript disabled or not working",       # exact phrase from reported error
    "Rapid taps or clicks",                     # exact phrase from reported error
    "Use of developer or inspection tools",     # exact phrase from reported error
    "restricted access",                        # ✅ FIX 10: seen on finance section block
    "Access denied",
    "Access to this page has been denied",      # ✅ FIX 10
    "Please verify you are a human",
    "Enable JavaScript and cookies to continue",
    "cf-browser-verification",
    "Checking your browser",
    "DDoS protection by",
    "Reference ID",                             # ✅ FIX 10: common on WAF block pages
]

# ✅ FIX 10: If a "successful" page is suspiciously small AND has no article
# cards, treat it as a likely block page even if the wording didn't match
# BLOCKED_PHRASES exactly (WAF wording varies by domain/section).
MIN_VALID_PAGE_LENGTH = 5000



# ============================
# MODEL
# ============================
@dataclass
class Listing:
    title: str
    url: str
    published_dt: Optional[str]
    source: str
    section: str


# ============================
# SESSION
# ============================
session = requests.Session()
session.headers.update(HEADERS)


# ============================
# DRIVER SETUP
# ============================
def make_driver():
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")

    # ✅ FIX 4: Stealth flags to suppress automation signals
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--lang=en-US,en")
    options.add_argument("--window-size=1920,1080")

    # ✅ FIX 12: Do NOT override the User-Agent with a hardcoded version.
    # Chrome's own native UA always matches the installed version exactly,
    # which is what we want — a hardcoded UA drifts out of sync after every
    # Chrome auto-update and becomes a bot-detection signal (mismatch between
    # claimed UA version and actual browser fingerprint).

    # ✅ FIX 13: Pass the DETECTED version (not a hardcoded one, and not nothing).
    # With no version_main, undetected_chromedriver sometimes downloads the
    # newest driver (e.g. 151) even when the installed Chrome is older (150),
    # which fails the other way round. CHROME_MAJOR is detected at startup,
    # so this stays correct across Chrome auto-updates without manual edits.
    driver = uc.Chrome(options=options, version_main=CHROME_MAJOR)

    # ✅ FIX 5: Hide webdriver property so JS fingerprinting can't detect automation
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return driver


# ============================
# HUMAN-LIKE SCROLL
# ============================
def human_scroll(driver):
    """
    ✅ FIX 6: Simulate realistic human scrolling pattern.
    Bots jump straight to scraping; humans scroll, pause, scroll back.
    """
    try:
        scroll_pause = random.uniform(0.6, 1.4)
        screen_height = driver.execute_script("return window.screen.height;")

        # Scroll down in small increments
        for i in range(1, 5):
            driver.execute_script(f"window.scrollTo(0, {int(screen_height * i * 0.25)});")
            time.sleep(scroll_pause)

        # Slight scroll back up — humans do this
        driver.execute_script(f"window.scrollTo(0, {int(screen_height * 0.15)});")
        time.sleep(random.uniform(0.4, 0.9))

    except Exception:
        pass  # Non-critical; continue even if scroll fails


# ============================
# WARM-UP
# ============================
def warm_up_session():
    """
    ✅ FIX 7: Start warm-up from a Google search referrer to mimic organic traffic.
    """
    warm_urls = [
        "https://www.google.com/search?q=becker+hospital+review+finance",  # come from Google first
        "https://www.beckerspayer.com/",
        "https://www.beckerspayer.com/payer/",
        "https://www.beckershospitalreview.com/",
    ]
    for u in warm_urls:
        try:
            r = session.get(u, timeout=TIMEOUT)
            print(f"  [warm-up requests] {u} -> {r.status_code}")
            # ✅ FIX 8: Randomized delay between warm-up requests
            time.sleep(random.uniform(3.0, 6.0))
        except Exception as e:
            print(f"  [warm-up requests failed] {u}: {e}")


def warm_up_selenium(driver):
    """
    Warm up Selenium by visiting home pages before scraping section pages.
    """
    warm_urls = [
        "https://www.google.com/search?q=becker+hospital+review",  # ✅ start from Google
        "https://www.beckerspayer.com/",
        "https://www.beckershospitalreview.com/",
    ]
    for u in warm_urls:
        try:
            driver.get(u)
            WebDriverWait(driver, SELENIUM_WAIT_SEC).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            human_scroll(driver)  # ✅ scroll during warm-up too
            # ✅ FIX 8: Randomized warm-up delay
            time.sleep(random.uniform(4.0, 8.0))
            print(f"  [warm-up selenium] {u} OK")
        except Exception as e:
            print(f"  [warm-up selenium] {u} failed: {e}")


# ============================
# URL NORMALIZE
# ============================
def clean_url(url: str) -> str:
    p = urlparse(url)
    q = [
        (k, v)
        for k, v in parse_qsl(p.query, keep_blank_values=True)
        if not k.lower().startswith("utm_") and k.lower() not in {"oly_enc_id", "origin"}
    ]
    path = p.path
    if path.endswith("/") and path != "/":
        path = path[:-1]
    return urlunparse((p.scheme, p.netloc, path, p.params, urlencode(q, doseq=True), ""))


def normalize_url(base_url: str, href: str) -> str:
    return clean_url(urljoin(base_url, href))


# ============================
# TEXT NORMALIZATION
# ============================
def normalize_title(title: str) -> str:
    if not title:
        return ""
    t = title.strip().lower()
    # Fix mojibake: UTF-8 bytes mis-decoded as latin-1
    # Each replacement uses only ASCII or safe \xNN escapes — no raw unicode symbols
    t = t.replace("\xe2\x80\x98", "'")   # left single quotation mark
    t = t.replace("\xe2\x80\x99", "'")   # right single quotation mark / apostrophe
    t = t.replace("\xe2\x80\x9c", '"')   # left double quotation mark
    t = t.replace("\xe2\x80\x9d", '"')   # right double quotation mark
    t = t.replace("\xe2\x80\x93", "-")   # en dash
    t = t.replace("\xe2\x80\x94", "-")   # em dash
    t = t.replace("\xc2\xa0", " ")       # non-breaking space
    t = re.sub(r"\s+", " ", t).strip()
    return t


def normalize_pub_date(published_dt: Optional[str]) -> str:
    if not published_dt:
        return "UNKNOWN"
    try:
        dt = date_parser.parse(published_dt)
        return dt.date().isoformat()
    except Exception:
        return "UNKNOWN"


# ============================
# DATE PARSING
# ============================
def parse_date_loose(text: str) -> Optional[datetime]:
    if not text:
        return None
    t = re.sub(r"\s+", " ", text).strip()
    t = re.sub(r"(\d{1,2})(st|nd|rd|th)\b", r"\1", t, flags=re.IGNORECASE)
    try:
        return date_parser.parse(t, fuzzy=True)
    except Exception:
        return None


def parse_iso_any(dt_str: str) -> Optional[datetime]:
    if not dt_str:
        return None
    s = dt_str.strip()
    try:
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        return datetime.fromisoformat(s)
    except Exception:
        return None


# ============================
# REQUESTS FETCH
# ============================
def fetch_html(url: str) -> Tuple[Optional[str], Optional[str]]:
    last_err = None
    for attempt in range(5):
        try:
            r = session.get(url, timeout=TIMEOUT, allow_redirects=True)
            if r.status_code == 200:
                return r.text, None
            last_err = f"HTTP {r.status_code}"
        except Exception as e:
            last_err = f"{type(e).__name__}: {e}"

        time.sleep(3 * (attempt + 1))
    return None, last_err


def fetch_page_candidates(base_url: str, page: int) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    if page == 1:
        candidates = [
            base_url,
            f"{base_url.rstrip('/')}/",
            f"{base_url.rstrip('/')}/page/1/",
        ]
    else:
        candidates = [
            f"{base_url.rstrip('/')}/page/{page}/",
        ]

    seen = set()
    candidates = [c for c in candidates if not (c in seen or seen.add(c))]

    last_err = None
    for candidate in candidates:
        html, err = fetch_html(candidate)
        if html:
            return html, None, candidate
        last_err = err

    return None, last_err, candidates[-1] if candidates else None


# ============================
# SELENIUM FINANCE FETCH
# ============================
def is_blocked_page(html: str) -> bool:
    """
    ✅ FIX 10: Returns True if the page looks like a bot-block / WAF challenge page.
    Combines exact phrase matching with a structural fallback (short page,
    no article cards) since WAF wording varies by domain/section.
    """
    if any(phrase in html for phrase in BLOCKED_PHRASES):
        return True

    # Structural fallback: a real listing page has many <article> / <a> tags
    # and is well over MIN_VALID_PAGE_LENGTH bytes. A block page is typically
    # short and contains little real content.
    if len(html) < MIN_VALID_PAGE_LENGTH:
        soup = BeautifulSoup(html, "html.parser")
        if not soup.find_all("article") and len(soup.find_all("a")) < 10:
            return True

    return False


def fetch_finance_page_selenium(driver, page_url: str) -> Tuple[Optional[str], Optional[str]]:
    # ✅ FIX 10: retry with longer backoff if a block is detected,
    # instead of immediately giving up on the first attempt.
    max_attempts = 3
    last_err = None

    for attempt in range(max_attempts):
        try:
            driver.get(page_url)

            WebDriverWait(driver, SELENIUM_WAIT_SEC).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )

            # ✅ FIX 8: Randomized post-load wait instead of fixed 5s
            time.sleep(random.uniform(3.0, 7.0))

            # ✅ FIX 6: Human-like scroll after page load
            human_scroll(driver)

            html = driver.page_source

            # ✅ FIX 10: combined phrase + structural block detection
            if is_blocked_page(html):
                last_err = "Blocked by anti-bot challenge"
                if attempt < max_attempts - 1:
                    backoff = random.uniform(45.0, 90.0) * (attempt + 1)
                    print(f"    [block detected] attempt {attempt + 1}/{max_attempts}, "
                          f"backing off {backoff:.0f}s before retry...")
                    time.sleep(backoff)
                    continue
                return None, last_err

            return html, None

        except Exception as e:
            last_err = f"Selenium {type(e).__name__}: {e}"
            if attempt < max_attempts - 1:
                time.sleep(random.uniform(10.0, 20.0))
                continue
            return None, last_err

    return None, last_err


def fetch_finance_page_candidates_selenium(driver, base_url: str, page: int):
    if page == 1:
        candidates = [
            base_url,
            f"{base_url.rstrip('/')}/",
            f"{base_url.rstrip('/')}/page/1/",
        ]
    else:
        candidates = [
            f"{base_url.rstrip('/')}/page/{page}/",
        ]

    seen = set()
    candidates = [c for c in candidates if not (c in seen or seen.add(c))]

    last_err = None
    for candidate in candidates:
        html, err = fetch_finance_page_selenium(driver, candidate)
        if html:
            return html, None, candidate
        last_err = err

    return None, last_err, candidates[-1] if candidates else None


# ============================
# LISTING PARSE
# ============================
def extract_bhr_cards(soup: BeautifulSoup):
    return soup.find_all("article", class_="bh-card")


def parse_bhr_card(card, base_url: str):
    title_tag = card.find("h3", class_="bh-card__title")
    if not title_tag:
        return None

    a = title_tag.find("a", href=True)
    if not a:
        return None

    title = a.get_text(" ", strip=True)
    url = normalize_url(base_url, a["href"])

    published = None
    t = card.find("time", class_="byline__time")
    if t:
        dt_str = t.get("datetime") or t.get_text(" ", strip=True)
        published = parse_iso_any(dt_str) or parse_date_loose(dt_str)

    return title, url, published


def parse_generic_listing(soup: BeautifulSoup, base_url: str):
    results = []
    seen = set()

    for a in soup.select("h2 a, h3 a, h4 a"):
        href = a.get("href")
        title = a.get_text(" ", strip=True)

        if not href or not title:
            continue

        url = normalize_url(base_url, href)

        if ("beckerspayer.com" not in url) and ("beckershospitalreview.com" not in url):
            continue

        if len(title) < 20:
            continue

        pub = None
        parent = a.parent
        for _ in range(8):
            if parent is None:
                break
            text = parent.get_text(" ", strip=True)
            if re.search(r"\b(20\d{2}|yesterday|\d+\s+hours?\s+ago|\d+\s+days?\s+ago)\b", text, re.IGNORECASE):
                pub = parse_date_loose(text)
                if pub:
                    break
            parent = parent.parent

        dedupe_key = (title.lower(), url)
        if dedupe_key not in seen:
            seen.add(dedupe_key)
            results.append((title, url, pub))

    return results


# ============================
# ARTICLE DATE FALLBACK
# ============================
def extract_article_date(article_url: str) -> Tuple[Optional[datetime], Optional[str]]:
    html, err = fetch_html(article_url)
    if not html:
        return None, err or "Unknown error fetching article"

    soup = BeautifulSoup(html, "html.parser")

    t = soup.find("time")
    if t:
        dt_str = t.get("datetime") or t.get_text(" ", strip=True)
        d = parse_iso_any(dt_str) or parse_date_loose(dt_str)
        if d:
            return d, None

    meta = soup.find("meta", attrs={"property": "article:published_time"})
    if meta and meta.get("content"):
        d = parse_iso_any(meta["content"]) or parse_date_loose(meta["content"])
        if d:
            return d, None

    return None, "Date not found on article page"


# ============================
# MASTER READ HELPERS
# ============================
def read_existing_outcsv(out_csv: str) -> Tuple[Optional[date], set]:
    existing_urls = set()
    max_dt: Optional[datetime] = None

    try:
        with open(out_csv, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for r in reader:
                urls_field = (r.get("urls") or "").strip()
                if urls_field:
                    for u in urls_field.split(","):
                        u = u.strip()
                        if u:
                            existing_urls.add(u)

                s = (r.get("published_dt") or "").strip()
                if not s:
                    continue
                try:
                    dt = date_parser.parse(s)
                    if max_dt is None or dt > max_dt:
                        max_dt = dt
                except Exception:
                    continue

        return (max_dt.date() if max_dt else None), existing_urls

    except Exception:
        return None, set()


def load_existing_merged_rows(out_csv: str) -> Dict[Tuple[str, str], Dict[str, str]]:
    merged: Dict[Tuple[str, str], Dict[str, str]] = {}
    try:
        with open(out_csv, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for r in reader:
                title = (r.get("title") or "").strip()
                published_dt = (r.get("published_dt") or "").strip()
                sources = (r.get("sources") or "").strip()
                sections = (r.get("sections") or "").strip()
                urls = (r.get("urls") or "").strip()

                tkey = normalize_title(title)
                dkey = normalize_pub_date(published_dt if published_dt else None)
                mkey = (tkey, dkey)

                merged[mkey] = {
                    "title": title,
                    "published_dt": published_dt,
                    "sources": sources,
                    "sections": sections,
                    "urls": urls,
                }
    except Exception:
        pass
    return merged


# ============================
# MERGE HELPERS
# ============================
def merge_csv_values(old_val: str, new_val: str) -> str:
    old_items = [x.strip() for x in old_val.split(",") if x.strip()] if old_val else []
    new_items = [x.strip() for x in new_val.split(",") if x.strip()] if new_val else []

    combined = []
    seen = set()

    for item in old_items + new_items:
        if item not in seen:
            seen.add(item)
            combined.append(item)

    return ", ".join(combined)


def add_to_merged(merged: Dict[Tuple[str, str], Dict[str, str]], it: Listing):
    tkey = normalize_title(it.title)
    dkey = normalize_pub_date(it.published_dt)
    mkey = (tkey, dkey)

    if mkey not in merged:
        merged[mkey] = {
            "title": it.title,
            "published_dt": it.published_dt or "",
            "sources": it.source,
            "sections": it.section,
            "urls": it.url,
        }
        return

    merged[mkey]["sources"] = merge_csv_values(merged[mkey]["sources"], it.source)
    merged[mkey]["sections"] = merge_csv_values(merged[mkey]["sections"], it.section)
    merged[mkey]["urls"] = merge_csv_values(merged[mkey]["urls"], it.url)


# ============================
# MAIN
# ============================
def main():
    t0 = time.time()

    os.makedirs(OUT_DIR, exist_ok=True)

    watermark, existing_urls = read_existing_outcsv(OUT_CSV)
    if watermark:
        print(f"[watermark] Latest published_dt in previous OUT_CSV = {watermark}")
    else:
        watermark = DEFAULT_CUTOFF_DATE
        print(f"[watermark] No previous OUT_CSV found -> using default {watermark}")

    merged = load_existing_merged_rows(OUT_CSV)

    driver = None
    try:
        driver = make_driver()
        warm_up_selenium(driver)
        print("[selenium] Driver started and warmed up successfully")
    except Exception as e:
        print(f"[warn] Selenium driver could not be started: {e}")

    delta_rows: List[Dict[str, str]] = []
    section_stats: Dict[str, Dict[str, int]] = {}

    new_kept_total = 0
    new_added_urls = 0

    try:
        for base_url, source, section in BASE_SECTIONS:
            # ✅ FIX 8: Wider randomized inter-section delay
            time.sleep(random.uniform(10.0, 20.0))
            key = f"{source} | {section}"
            section_stats[key] = {
                "pages": 0,
                "found": 0,
                "new_kept": 0,
                "skipped_old": 0,
                "missing_date": 0,
                "page_errors": 0,
                "article_date_errors": 0,
            }

            print(f"\n[crawl] {key}")

            # ============================
            # SELENIUM PATH (finance + USE_SELENIUM_FOR_ALL)
            # ============================
            if USE_SELENIUM_FOR_ALL or (section == "finance" and USE_SELENIUM_FOR_FINANCE):
                page = 1
                stop_section = False
                consecutive_old_pages = 0  # ✅ FIX 9

                while not stop_section and page <= MAX_PAGES:
                    if driver is None:
                        section_stats[key]["page_errors"] += 1
                        print("  [finance error] Selenium driver unavailable")
                        break

                    html, err, used_url = fetch_finance_page_candidates_selenium(driver, base_url, page)

                    if not html:
                        section_stats[key]["page_errors"] += 1
                        print(f"  [finance page error] {used_url} -> {err}")
                        break

                    section_stats[key]["pages"] += 1
                    soup = BeautifulSoup(html, "html.parser")

                    items: List[Tuple[str, str, Optional[datetime]]] = []
                    cards = extract_bhr_cards(soup)

                    if cards:
                        for c in cards:
                            parsed = parse_bhr_card(c, base_url)
                            if parsed:
                                items.append(parsed)
                    else:
                        items.extend(parse_generic_listing(soup, base_url))

                    if not items:
                        # ✅ FIX 11: dump the page HTML for diagnosis instead of
                        # silently breaking — helps tell apart "no cards in HTML
                        # structure changed" vs "this is actually a block/interstitial page"
                        debug_path = os.path.join(
                            OUT_DIR if os.path.isdir(OUT_DIR) else ".",
                            f"debug_{section}_page{page}.html"
                        )
                        try:
                            with open(debug_path, "w", encoding="utf-8") as dbg:
                                dbg.write(html)
                            print(f"  [debug] No items found. Page length={len(html)}. "
                                  f"Saved HTML to: {debug_path}")
                        except Exception as e:
                            print(f"  [debug] No items found. Page length={len(html)}. "
                                  f"Could not save debug HTML: {e}")
                        break

                    section_stats[key]["found"] += len(items)

                    page_dates: List[date] = []
                    page_all_have_dates = True
                    resolved_items: List[Tuple[str, str, Optional[datetime]]] = []

                    for title, url, pub in items:
                        if pub is None:
                            page_all_have_dates = False
                            section_stats[key]["missing_date"] += 1

                        if pub is not None:
                            page_dates.append(pub.date())

                        resolved_items.append((title, url, pub))

                    # ✅ FIX 9b: track how many NEW items this specific page contributed,
                    # independent of whether every item had a parseable date.
                    new_kept_before_page = section_stats[key]["new_kept"]

                    for title, url, pub in resolved_items:
                        it = Listing(
                            source=source,
                            section=section,
                            title=title,
                            url=url,
                            published_dt=pub.isoformat() if pub else None,
                        )

                        if pub is not None and pub.date() <= watermark:
                            section_stats[key]["skipped_old"] += 1
                            add_to_merged(merged, it)
                            continue

                        if url in existing_urls:
                            add_to_merged(merged, it)
                            continue

                        existing_urls.add(url)
                        new_added_urls += 1
                        add_to_merged(merged, it)

                        delta_rows.append({
                            "title": it.title,
                            "published_dt": it.published_dt or "",
                            "sources": it.source,
                            "sections": it.section,
                            "urls": it.url,
                        })

                        section_stats[key]["new_kept"] += 1
                        new_kept_total += 1

                    # ✅ FIX 9b: the real stop signal is "did this page add ANY new rows?"
                    # This is independent of date-parsing (FIX 9 alone failed when
                    # pub was None for some items, which kept resetting the counter
                    # to 0 on every page regardless of how old the content was).
                    page_contributed_new = section_stats[key]["new_kept"] > new_kept_before_page

                    if page_contributed_new:
                        consecutive_old_pages = 0
                    else:
                        consecutive_old_pages += 1

                    # ✅ FIX 9: stop only after several consecutive pages with nothing new
                    if consecutive_old_pages >= CONSECUTIVE_OLD_PAGES_TO_STOP:
                        stop_section = True

                    page += 1
                    # ✅ FIX 8: Wider random sleep between pages
                    time.sleep(SLEEP_SEC + random.uniform(5.0, 12.0))

                s = section_stats[key]
                print(
                    f"  [section summary] pages={s['pages']} found={s['found']} new_kept={s['new_kept']} "
                    f"skipped_old={s['skipped_old']} missing_date={s['missing_date']} "
                    f"page_errors={s['page_errors']} article_date_errors={s['article_date_errors']}"
                )
                continue

            # ============================
            # NON-FINANCE via requests
            # ============================
            page = 1
            stop_section = False
            consecutive_old_pages = 0  # ✅ FIX 9

            while not stop_section and page <= MAX_PAGES:
                html, err, used_url = fetch_page_candidates(base_url, page)

                if not html:
                    section_stats[key]["page_errors"] += 1
                    print(f"  [page error] {used_url} -> {err}")
                    break

                section_stats[key]["pages"] += 1
                soup = BeautifulSoup(html, "html.parser")

                items: List[Tuple[str, str, Optional[datetime]]] = []
                cards = extract_bhr_cards(soup)

                if cards:
                    for c in cards:
                        parsed = parse_bhr_card(c, base_url)
                        if parsed:
                            items.append(parsed)
                else:
                    items.extend(parse_generic_listing(soup, base_url))

                if not items:
                    break

                section_stats[key]["found"] += len(items)

                page_dates: List[date] = []
                page_all_have_dates = True
                resolved_items: List[Tuple[str, str, Optional[datetime]]] = []

                for title, url, pub in items:
                    if pub is None:
                        got, _ = extract_article_date(url)
                        # ✅ FIX 8: Wider random sleep for article date fetches
                        time.sleep(SLEEP_SEC + random.uniform(5.0, 12.0))
                        pub = got
                        if pub is None:
                            page_all_have_dates = False
                            section_stats[key]["missing_date"] += 1
                            section_stats[key]["article_date_errors"] += 1

                    if pub is not None:
                        page_dates.append(pub.date())

                    resolved_items.append((title, url, pub))

                new_kept_before_page = section_stats[key]["new_kept"]

                for title, url, pub in resolved_items:
                    it = Listing(
                        source=source,
                        section=section,
                        title=title,
                        url=url,
                        published_dt=pub.isoformat() if pub else None,
                    )

                    if pub is not None and pub.date() <= watermark:
                        section_stats[key]["skipped_old"] += 1
                        add_to_merged(merged, it)
                        continue

                    if url in existing_urls:
                        add_to_merged(merged, it)
                        continue

                    existing_urls.add(url)
                    new_added_urls += 1
                    add_to_merged(merged, it)

                    delta_rows.append({
                        "title": it.title,
                        "published_dt": it.published_dt or "",
                        "sources": it.source,
                        "sections": it.section,
                        "urls": it.url,
                    })

                    section_stats[key]["new_kept"] += 1
                    new_kept_total += 1

                # ✅ FIX 9b: stop based on whether this page contributed any new rows,
                # independent of whether all items had parseable dates.
                page_contributed_new = section_stats[key]["new_kept"] > new_kept_before_page

                if page_contributed_new:
                    consecutive_old_pages = 0
                else:
                    consecutive_old_pages += 1

                if consecutive_old_pages >= CONSECUTIVE_OLD_PAGES_TO_STOP:
                    stop_section = True

                page += 1
                # ✅ FIX 8: Wider random sleep between pages
                time.sleep(SLEEP_SEC + random.uniform(5.0, 12.0))

            s = section_stats[key]
            print(
                f"  [section summary] pages={s['pages']} found={s['found']} new_kept={s['new_kept']} "
                f"skipped_old={s['skipped_old']} missing_date={s['missing_date']} "
                f"page_errors={s['page_errors']} article_date_errors={s['article_date_errors']}"
            )

    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass

    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["title", "published_dt", "sources", "sections", "urls"])
        writer.writeheader()
        for row in merged.values():
            writer.writerow(row)

    with open(DELTA_CSV, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["title", "published_dt", "sources", "sections", "urls"])
        writer.writeheader()
        for row in delta_rows:
            writer.writerow(row)

    t1 = time.time()
    total_seconds = t1 - t0

    print("\n==============================")
    print("[done] Stage 1 incremental complete")
    print(f"New rows kept in this run: {new_kept_total}")
    print(f"New unique URLs added: {new_added_urls}")
    print(f"Output (merged master) file: {OUT_CSV}")
    print(f"Output (delta new-only) file: {DELTA_CSV}")
    print(f"Runtime: {total_seconds:.2f} sec ({total_seconds/60:.2f} min)")
    print("==============================\n")


if __name__ == "__main__":
    main()

    import pandas as pd

    df = pd.read_csv(OUT_CSV)

    print("Total rows in merged file:", len(df))
    print("Unique titles:", df["title"].nunique())
    print("Unique URLs:", df["urls"].nunique())

    print("\nRows by source (raw merged field):")
    print(
        df.groupby("sources")
          .size()
          .reset_index(name="row_count")
          .sort_values("row_count", ascending=False)
    )

    print("\nRows by section (raw merged field):")
    print(
        df.groupby("sections")
          .size()
          .reset_index(name="row_count")
          .sort_values("row_count", ascending=False)
    )

    print("\nActual section-wise counts:")
    section_count = (
        df.assign(section=df["sections"].fillna("").str.split(", "))
          .explode("section")
          .groupby("section")
          .size()
          .reset_index(name="row_count")
          .sort_values("row_count", ascending=False)
    )
    print(section_count)

    print("\nActual source-section counts:")
    df2 = df.copy()
    df2["sources"] = df2["sources"].fillna("").str.split(", ")
    df2["sections"] = df2["sections"].fillna("").str.split(", ")
    df2 = df2.explode("sources").explode("sections")

    source_section_count = (
        df2.groupby(["sources", "sections"])
           .size()
           .reset_index(name="row_count")
           .sort_values(["sources", "row_count"], ascending=[True, False])
    )
    print(source_section_count)