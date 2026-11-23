"""

REQUIREMENTS
    pip install selenium beautifulsoup4 openpyxl pandas
    (Google Chrome must be installed. Selenium 4.6+ auto-manages the
    matching chromedriver, no separate driver download needed.)


"""

import os
import re
import sys
import time
import random
import threading
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urlparse, parse_qs, unquote

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ----------------------------------------------------------------------
# CONFIG -- edit these two paths, then just run: python fb_scraper.py
# ----------------------------------------------------------------------
INPUT_FILE = "links.txt"
OUTPUT_FILE = "output.xlsx"

LOGIN = False   # set True only if pages need login to show contact info
FB_EMAIL = os.environ.get("FB_EMAIL", "")
FB_PASSWORD = os.environ.get("FB_PASSWORD", "")

HEADLESS = True                # runs in the background, no visible window
PAGE_LOAD_WAIT = 5             # seconds to let JS finish rendering
PAGE_LOAD_TIMEOUT = 25         # max seconds to wait for a page before giving up on it
MIN_DELAY, MAX_DELAY = 4, 9    # randomized delay between pages (be gentle)
MAX_PARALLEL_ALLOWED = 10      # upper bound offered at the prompt
AUTOSAVE_EVERY = 100           # write progress to OUTPUT_FILE every N completed links

COLUMNS = ["Page URL", "Name", "Followers", "Following",
           "Address", "Phone Number", "Email", "Website Link", "Status"]

LABEL_MAP = {
    "address": "Address",
    "mobile": "Phone Number",
    "phone": "Phone Number",
    "call": "Phone Number",
    "email address": "Email",
    "email": "Email",
    "website": "Website Link",
}

print_lock = threading.Lock()
save_lock = threading.Lock()
progress_lock = threading.Lock()
_completed_count = 0


def safe_print(msg):
    with print_lock:
        print(msg, flush=True)


def save_results(results):
    """Thread-safe write of current (possibly partial) results to OUTPUT_FILE."""
    with save_lock:
        df = pd.DataFrame([r if r else {} for r in results], columns=COLUMNS)
        df.to_excel(OUTPUT_FILE, index=False)


def note_completed_and_maybe_autosave(results, total):
    """Call once per finished link. Every AUTOSAVE_EVERY completions,
    saves current progress to OUTPUT_FILE so nothing is lost if the
    run is interrupted or a page crashes the browser."""
    global _completed_count
    with progress_lock:
        _completed_count += 1
        count = _completed_count
    if count % AUTOSAVE_EVERY == 0 or count == total:
        save_results(results)
        safe_print(f"[Autosave] {count}/{total} links done -- progress saved to {OUTPUT_FILE}")


def build_driver():
    options = Options()
    if HEADLESS:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1366,900")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
    return webdriver.Chrome(options=options)


def configure_driver_timeouts(driver):
    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
    driver.set_script_timeout(PAGE_LOAD_TIMEOUT)


def login(driver):
    if not (LOGIN and FB_EMAIL and FB_PASSWORD):
        return
    try:
        driver.get("https://www.facebook.com/login")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "email")))
        driver.find_element(By.ID, "email").send_keys(FB_EMAIL)
        driver.find_element(By.ID, "pass").send_keys(FB_PASSWORD)
        driver.find_element(By.NAME, "login").click()
        time.sleep(6)
    except TimeoutException:
        print("Login page did not load as expected -- continuing without login.")


def extract_raw_url(line):
    """
    Pull the actual URL out of a line that may be a plain URL, or a
    markdown-style link like:
        [https://www.facebook.com/edenbythelake](https://www.facebook.com/edenbythelake/about)
    In the markdown case we take the URL inside the parentheses, since
    that's the one meant to actually be opened.
    """
    line = line.strip()
    if not line:
        return None
    md_match = re.match(r"^\[.*?\]\((https?://[^\)]+)\)$", line)
    if md_match:
        return md_match.group(1)
    # fallback: grab the first http(s) URL found anywhere in the line
    url_match = re.search(r"https?://[^\s\)\]]+", line)
    if url_match:
        return url_match.group(0)
    return None


def normalize_about_url(page_url):
    raw = extract_raw_url(page_url)
    if not raw:
        return None
    raw = raw.rstrip("/")
    # strip a trailing /about (and variants) so we don't end up with /about/about
    raw = re.sub(r"/about(_contact_and_basic_info)?$", "", raw, flags=re.IGNORECASE)
    return raw + "/about"


def clean_outbound_link(url):
    """
    Facebook wraps outbound links (Website, Instagram, etc.) through its
    l.facebook.com/l.php?u=<encoded target>&h=... redirector. Unwrap that
    back to the real destination URL. Non-wrapped URLs pass through as-is.
    """
    if not url:
        return url
    parsed = urlparse(url)
    if "l.facebook.com" in parsed.netloc and parsed.path == "/l.php":
        qs = parse_qs(parsed.query)
        real = qs.get("u", [None])[0]
        if real:
            return unquote(real)
    return url


def parse_about_html(html):
    soup = BeautifulSoup(html, "html.parser")

    data = {
        "Name": "",
        "Followers": "",
        "Following": "",
        "Address": "",
        "Phone Number": "",
        "Email": "",
        "Website Link": "",
    }

    # --- Name ---
    h1 = soup.find("h1")
    if h1:
        data["Name"] = h1.get_text(strip=True)

    full_text = soup.get_text("\n")
    lines = [l.strip() for l in full_text.split("\n") if l.strip()]

    # --- Followers / Following ---
    joined = " ".join(lines)
    m = re.search(r"([\d.,]+\s?[KMB]?)\s*followers", joined, re.IGNORECASE)
    if m:
        data["Followers"] = m.group(1).strip()
    m = re.search(r"([\d.,]+\s?[KMB]?)\s*following\b", joined, re.IGNORECASE)
    if m:
        data["Following"] = m.group(1).strip()

    # --- Value-before-label pass (Address / Phone / Email / Website) ---
    for i, line in enumerate(lines):
        key = line.lower()
        if key in LABEL_MAP and i > 0:
            field = LABEL_MAP[key]
            if not data[field]:
                candidate = lines[i - 1]
                # guard against grabbing another label as the "value"
                if candidate.lower() not in LABEL_MAP:
                    if field == "Website Link":
                        candidate = clean_outbound_link(candidate)
                    data[field] = candidate

    # --- Regex fallbacks in case label-based matching missed something ---
    if not data["Email"]:
        m = re.search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", joined)
        if m:
            data["Email"] = m.group(0)

    if not data["Phone Number"]:
        m = re.search(r"(\+?\d[\d\-\s\(\)]{7,}\d)", joined)
        if m:
            data["Phone Number"] = m.group(0).strip()

    # --- Prefer the real href for the website link ---
    website_section = None
    for h2 in soup.find_all("h2"):
        if "website" in h2.get_text(strip=True).lower():
            website_section = h2.find_parent("section")
            break
    if website_section:
        a = website_section.find("a", href=True)
        if a and a["href"].startswith("http"):
            data["Website Link"] = clean_outbound_link(a["href"])
    elif not data["Website Link"]:
        m = re.search(r"https?://[^\s\"'<>]+", joined)
        if m:
            data["Website Link"] = clean_outbound_link(m.group(0))

    return data


def scrape_page(driver, raw_line, idx, total, worker_id=None):
    about_url = normalize_about_url(raw_line)
    clean_page_url = about_url[:-len("/about")] if about_url else raw_line.strip()
    result = {
        "Page URL": clean_page_url,
        "Name": "", "Followers": "", "Following": "",
        "Address": "", "Phone Number": "", "Email": "", "Website Link": "",
        "Status": "",
    }

    tag = f"W{worker_id}|" if worker_id is not None else ""
    prefix = f"[{tag}{idx}/{total}]"

    if not about_url:
        result["Status"] = "Invalid URL"
        safe_print(f"{prefix} SKIP  (couldn't parse a URL from: {raw_line!r})")
        return result

    safe_print(f"{prefix} Opening : {about_url}")

    try:
        t0 = time.time()
        try:
            driver.get(about_url)
        except TimeoutException:
            elapsed = time.time() - t0
            result["Status"] = "Timed out (likely deleted/unavailable)"
            safe_print(
                f"{prefix} SKIP    : page didn't finish loading in {PAGE_LOAD_TIMEOUT}s "
                f"[{elapsed:.1f}s] -- likely deleted/unavailable, moving on\n" + "-" * 70
            )
            try:
                driver.execute_script("window.stop();")
            except Exception:
                pass
            return result

        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.TAG_NAME, "h1"))
            )
        except TimeoutException:
            safe_print(f"{prefix}   (page took a while to load, continuing anyway)")
        time.sleep(PAGE_LOAD_WAIT)

        html = driver.page_source
        parsed = parse_about_html(html)
        result.update(parsed)
        result["Status"] = "OK" if result["Name"] else "No data found"
        elapsed = time.time() - t0

        name = result["Name"] or "(no name found)"
        lines = [
            f"{prefix} Done    : {name}  [{elapsed:.1f}s]",
            f"{prefix}   Followers : {result['Followers'] or '-'}   "
            f"Following: {result['Following'] or '-'}",
            f"{prefix}   Address   : {result['Address'] or '-'}",
            f"{prefix}   Phone     : {result['Phone Number'] or '-'}",
            f"{prefix}   Email     : {result['Email'] or '-'}",
            f"{prefix}   Website   : {result['Website Link'] or '-'}",
            "-" * 70,
        ]
        safe_print("\n".join(lines))

    except Exception as e:
        result["Status"] = f"Error: {e}"
        safe_print(f"{prefix} ERROR   : {e}\n" + "-" * 70)

    return result


def ask_parallel_count(total_links):
    default = 1
    while True:
        raw = input(
            f"How many links do you want to run in parallel at a time? "
            f"(1-{MAX_PARALLEL_ALLOWED}, default {default}): "
        ).strip()
        if not raw:
            return min(default, total_links)
        if raw.isdigit():
            n = int(raw)
            if 1 <= n <= MAX_PARALLEL_ALLOWED:
                return min(n, total_links)
        print(f"  Please enter a whole number between 1 and {MAX_PARALLEL_ALLOWED}.")


def worker_run(worker_id, links, indices, total, results):
    """One worker: owns a single browser and works through its assigned
    subset of links (given as indices into the master `links` / `results`
    lists), writing each result into `results[idx]` in place."""
    driver = build_driver()
    configure_driver_timeouts(driver)
    login(driver)
    try:
        for n, idx in enumerate(indices, 1):
            results[idx] = scrape_page(
                driver, links[idx], idx + 1, total, worker_id=worker_id
            )
            note_completed_and_maybe_autosave(results, total)
            if n < len(indices):
                time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))
    finally:
        driver.quit()


def main():
    print("=" * 70)
    print(" FACEBOOK PAGE ABOUT SCRAPER")
    print("=" * 70)
    print(f" Input file  : {INPUT_FILE}")
    print(f" Output file : {OUTPUT_FILE}")
    print(f" Headless    : {HEADLESS}")
    print(f" Login       : {LOGIN}")
    print(f" Autosave    : every {AUTOSAVE_EVERY} links")
    print("=" * 70)

    if not os.path.exists(INPUT_FILE):
        print(f"\nERROR: input file not found: {INPUT_FILE}")
        print("Edit INPUT_FILE at the top of this script, or place the file next to it.")
        sys.exit(1)

    with open(INPUT_FILE, "r", encoding="utf-8") as f:
        links = [line.strip() for line in f if line.strip()]

    total = len(links)
    if total == 0:
        print(f"\nERROR: {INPUT_FILE} has no links in it.")
        sys.exit(1)

    print(f"\nLoaded {total} link(s) from {INPUT_FILE}\n")

    num_workers = ask_parallel_count(total)
    print(f"\nRunning with {num_workers} browser(s) in parallel.")
    print("Starting browser(s)...\n")
    print("-" * 70)

    results = [None] * total
    start_time = time.time()

    # Split link indices round-robin across workers so progress prints
    # from each worker stay interleaved evenly rather than batched.
    worker_indices = [list(range(w, total, num_workers)) for w in range(num_workers)]

    with ThreadPoolExecutor(max_workers=num_workers) as executor:
        futures = [
            executor.submit(worker_run, w + 1, links, worker_indices[w], total, results)
            for w in range(num_workers)
            if worker_indices[w]
        ]
        for fut in futures:
            fut.result()  # re-raise any worker exception

    save_results(results)

    filled = sum(1 for r in results if r and r.get("Status") == "OK")
    skipped = sum(1 for r in results if r and r.get("Status", "").startswith("Timed out"))
    total_time = time.time() - start_time

    print("=" * 70)
    print(f" DONE — {filled}/{total} pages scraped successfully"
          + (f"  ({skipped} timed out / unavailable)" if skipped else ""))
    print(f" Total time  : {total_time/60:.1f} min")
    print(f" Output file : {os.path.abspath(OUTPUT_FILE)}")
    print("=" * 70)


if __name__ == "__main__":
    main()