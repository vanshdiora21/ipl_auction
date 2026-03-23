"""
CREX Fantasy Points Tracker — v2
==================================
Scrapes today's IPL match scorecard(s) from crex.com for fantasy points
and logs them into your players.xlsx with today's date.

STRATEGY: Rather than searching per player (which gets blocked), this script:
  1. Finds today's IPL matches on CREX
  2. Scrapes the fantasy points tab of each scorecard
  3. Matches the points back to your player list

HOW TO RUN:
    python3 fantasy_points_tracker.py

REQUIREMENTS (inside your venv):
    pip install selenium openpyxl webdriver-manager beautifulsoup4 requests

YOUR EXCEL FILE:
    Sheet name:  Players
    Column A:    Player Name   (e.g. "Abhishek Sharma")
    Column B:    Team          (e.g. "SRH") — optional but helps matching

python3 -m venv venv

source venv/bin/activate
pip install openpyxl selenium webdriver-manager beautifulsoup4 requests
python3 fantasy_points_tracker.py


"""

import time
import re
import sys
from datetime import datetime
from pathlib import Path

# ─── Dependency checks ───────────────────────────────────────────────────────
try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    sys.exit("pip install openpyxl")

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    sys.exit("pip install requests beautifulsoup4")

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, WebDriverException
except ImportError:
    sys.exit("pip install selenium")

try:
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service
    HAS_WDM = True
except ImportError:
    HAS_WDM = False

# ─── Config ──────────────────────────────────────────────────────────────────
EXCEL_FILE      = "players.xlsx"
DATA_START_ROW  = 2
HEADLESS        = False   # False = visible browser window (more reliable, bypasses blocks)
WAIT_SECS       = 4
IPL_SERIES_ID   = "indian-premier-league-2026-1PW"
CREX_BASE       = "https://crex.com"
# ─────────────────────────────────────────────────────────────────────────────


def make_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--window-size=1280,900")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    )
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--ignore-ssl-errors")

    if HAS_WDM:
        svc = Service(ChromeDriverManager().install())
        d = webdriver.Chrome(service=svc, options=opts)
    else:
        d = webdriver.Chrome(options=opts)

    d.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    return d


def get_todays_match_urls(driver):
    url = f"{CREX_BASE}/series/{IPL_SERIES_ID}/matches"
    print(f"  Opening: {url}")
    driver.get(url)
    time.sleep(WAIT_SECS)

    soup = BeautifulSoup(driver.page_source, "html.parser")
    today = datetime.today()
    # Multiple date format variants to match against page text
    date_variants = [
        today.strftime("%-d %b").lower(),   # "22 mar"
        today.strftime("%b %-d").lower(),   # "mar 22"
        today.strftime("%d %b").lower(),    # "22 mar" (zero-padded)
        today.strftime("%B %-d").lower(),   # "march 22"
    ]

    match_urls = []
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "/scoreboard/" not in href and "scorecard" not in href:
            continue
        card = a.find_parent(class_=re.compile(r"match|card|fixture|row", re.I)) or a
        parent_text = card.get_text(separator=" ", strip=True).lower()
        if any(dv in parent_text for dv in date_variants):
            full = href if href.startswith("http") else CREX_BASE + href
            if full not in match_urls:
                match_urls.append(full)

    # Fallback: grab up to 2 most recent scorecard links
    if not match_urls:
        print("  Could not detect today by date — grabbing latest match links.")
        seen = []
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if "/scoreboard/" in href or "scorecard" in href:
                full = href if href.startswith("http") else CREX_BASE + href
                if full not in seen:
                    seen.append(full)
        match_urls = seen[:2]

    return match_urls


def scrape_fantasy_points_from_scorecard(driver, scorecard_url):
    base_url = scorecard_url.split("?")[0].rstrip("/")
    if not base_url.endswith("/scorecard"):
        base_url = base_url + "/scorecard"

    print(f"  Loading: {base_url}")
    driver.get(base_url)
    time.sleep(WAIT_SECS)

    # Try clicking the Fantasy tab
    try:
        tabs = driver.find_elements(
            By.XPATH,
            "//*[(self::a or self::button or self::span or self::li or self::div) "
            "and contains(translate(normalize-space(text()),'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
            "'abcdefghijklmnopqrstuvwxyz'),'fantasy')]"
        )
        for tab in tabs:
            try:
                if tab.is_displayed():
                    driver.execute_script("arguments[0].click();", tab)
                    time.sleep(WAIT_SECS)
                    break
            except Exception:
                continue
    except Exception:
        pass

    points = _parse_page(driver.page_source)

    # Try /fantasy URL if tab parsing got nothing
    if not points:
        fantasy_url = base_url.replace("/scorecard", "/fantasy")
        print(f"  Trying: {fantasy_url}")
        driver.get(fantasy_url)
        time.sleep(WAIT_SECS)
        points = _parse_page(driver.page_source)

    return points


def _parse_page(html):
    soup = BeautifulSoup(html, "html.parser")
    points = {}

    # Walk every row-like element, look for name + number pairs
    for row in soup.find_all(["tr", "div", "li", "section"]):
        parts = [p.strip() for p in row.get_text(separator="|", strip=True).split("|") if p.strip()]
        nums = [(i, float(p)) for i, p in enumerate(parts) if re.match(r'^\d+(\.\d+)?$', p) and float(p) >= 2]
        names = [(i, p) for i, p in enumerate(parts) if _is_name(p)]
        if nums and names:
            n_idx, name = names[0]
            p_idx, pts = nums[0]
            # Ensure name and number are adjacent-ish
            if abs(n_idx - p_idx) <= 4:
                points[name] = pts

    # Fallback: regex scan entire text for "Name ... XX.X pts" patterns
    if not points:
        text = soup.get_text(separator=" ")
        # Pattern: capitalized words followed by a number
        for m in re.finditer(r'([A-Z][a-z]+(?: [A-Z][a-z]+){1,3})\s+(\d+\.?\d*)', text):
            name, pts = m.group(1), float(m.group(2))
            if _is_name(name) and pts >= 2:
                points[name] = pts

    return points


def _is_name(text):
    if not text or len(text) < 5 or len(text) > 40:
        return False
    if re.match(r'^\d+(\.\d+)?$', text):
        return False
    if not re.search(r'[A-Za-z]{2,}', text):
        return False
    bad = {"batting", "bowling", "player", "team", "runs", "wkts", "over",
           "points", "fantasy", "economy", "strike", "average", "total",
           "wicket", "innings", "catch", "stumping", "direct"}
    return text.lower().strip() not in bad


def normalize(name):
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z\s]', '', name.lower())).strip()


def match_players(your_players, scraped_points):
    results = {}
    scraped_norm = {normalize(k): v for k, v in scraped_points.items()}

    for player in your_players:
        name = player["name"]
        norm = normalize(name)
        parts = norm.split()
        pts = None

        # 1. Exact
        if norm in scraped_norm:
            pts = scraped_norm[norm]
        # 2. Last name match
        elif parts:
            last = parts[-1]
            for sn, sv in scraped_norm.items():
                if last in sn.split():
                    pts = sv
                    break
        # 3. First name match
        if pts is None and len(parts) > 1:
            first = parts[0]
            for sn, sv in scraped_norm.items():
                if sn.split() and sn.split()[0] == first:
                    pts = sv
                    break

        results[name] = str(pts) if pts is not None else "No match today"

    return results


def load_players(wb):
    ws = wb["Players"]
    players = []
    for row in ws.iter_rows(min_row=DATA_START_ROW, values_only=True):
        name = row[0]
        if not name:
            continue
        team = row[1] if len(row) > 1 else ""
        players.append({"name": str(name).strip(), "team": str(team).strip() if team else ""})
    return players


def get_or_create_date_col(ws, today_str):
    for col in range(3, ws.max_column + 2):
        cell = ws.cell(row=1, column=col)
        if cell.value is None:
            cell.value = today_str
            cell.font = Font(bold=True, color="FFFFFF", name="Arial")
            cell.fill = PatternFill("solid", start_color="1F4E79")
            cell.alignment = Alignment(horizontal="center")
            ws.column_dimensions[cell.column_letter].width = 14
            return col
        if str(cell.value) == today_str:
            return col
    return ws.max_column + 1


def write_results(ws, players, results, date_col):
    for row_idx, player in enumerate(players, start=DATA_START_ROW):
        val = results.get(player["name"], "N/A")
        cell = ws.cell(row=row_idx, column=date_col, value=val)
        cell.alignment = Alignment(horizontal="center")
        cell.font = Font(name="Arial", size=10)
        try:
            float(val)
            cell.fill = PatternFill("solid", start_color="E2EFDA")
            cell.font = Font(color="375623", name="Arial", size=10)
        except ValueError:
            if val == "No match today":
                cell.fill = PatternFill("solid", start_color="F2F2F2")
            else:
                cell.fill = PatternFill("solid", start_color="FCE4D6")


def main():
    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists():
        sys.exit(f"'{EXCEL_FILE}' not found. Put it in the same folder as this script.")

    today_str = datetime.today().strftime("%d-%b-%Y")
    print(f"\nCREX Fantasy Points Tracker  —  {today_str}")
    print("=" * 52)

    wb = load_workbook(excel_path)
    if "Players" not in wb.sheetnames:
        sys.exit("No sheet named 'Players' found.")

    players = load_players(wb)
    if not players:
        sys.exit("No players found. Names go from row 2 in Column A.")

    print(f"{len(players)} players loaded.\n")

    print(f"Launching Chrome (visible window — do not close it)...")
    try:
        driver = make_driver()
    except WebDriverException as e:
        sys.exit(f"Chrome failed: {e}")

    all_points = {}
    try:
        print("Finding today's IPL matches...")
        match_urls = get_todays_match_urls(driver)

        if not match_urls:
            print("No matches found for today.")
        else:
            print(f"Found {len(match_urls)} match(es):")
            for u in match_urls:
                print(f"  {u}")
            print()

        for url in match_urls:
            print(f"\nScraping: {url}")
            pts = scrape_fantasy_points_from_scorecard(driver, url)
            print(f"  {len(pts)} player-points found.")
            all_points.update(pts)

    finally:
        driver.quit()

    print(f"\nMatching players...")
    results = match_players(players, all_points)
    found = 0
    for name, pts in results.items():
        has_pts = pts not in ("N/A", "No match today")
        if has_pts:
            found += 1
        icon = "✅" if has_pts else "—"
        print(f"  {icon}  {name}: {pts}")

    ws = wb["Players"]
    date_col = get_or_create_date_col(ws, today_str)
    write_results(ws, players, results, date_col)
    wb.save(excel_path)

    print(f"\nDone! {found}/{len(players)} players had points today.")
    print(f"Saved to '{EXCEL_FILE}' under column '{today_str}'\n")


if __name__ == "__main__":
    main()