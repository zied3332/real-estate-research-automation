from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from bs4 import BeautifulSoup
import pandas as pd
from pathlib import Path
import re
import time
import logging
from datetime import datetime

INPUT_FILE = Path("data/input/case_ids.xlsx")
OUTPUT_DIR = Path("data/output")
SCREENSHOT_DIR = Path("screenshots")
LOG_DIR = Path("logs")

OUTPUT_FILE = OUTPUT_DIR / "cases_results.xlsx"
LOG_FILE = LOG_DIR / "run.log"
PROGRESS_FILE = Path("progress.txt")

START_URL = "https://fjdefile.phila.gov/efsfjd/zk_fjd_public_qry_00.zp_main_idx"

BATCH_SIZE = 10
MAX_RETRIES = 2
WAIT_BETWEEN_CASES = 2
WAIT_BETWEEN_RETRIES = 2
FAILURE_COOLDOWN = 20

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter("%(message)s"))
logging.getLogger("").addHandler(console)


def log(message):
    logging.info(message)


def clean_text(text):
    if not text:
        return ""
    return " ".join(text.replace("\xa0", " ").split())


def extract_money(text):
    match = re.search(r"\$[\d,]+(?:\.\d{2})?", text)
    return match.group(0) if match else ""


def load_case_ids():
    df = pd.read_excel(INPUT_FILE)
    log(f"Columns found: {df.columns.tolist()}")

    column = "Case ID" if "Case ID" in df.columns else df.columns[0]
    case_ids = df[column].dropna().astype(str).str.strip().tolist()

    if not case_ids:
        raise Exception("No Case IDs found in the input Excel file.")

    return case_ids


def load_existing_results():
    if OUTPUT_FILE.exists():
        df = pd.read_excel(OUTPUT_FILE)
        results = df.to_dict("records")
        log(f"Loaded existing output rows: {len(results)}")
        return results

    return []


def load_start_index():
    if PROGRESS_FILE.exists():
        text = PROGRESS_FILE.read_text().strip()
        if text.isdigit():
            return int(text)
    return 0


def save_start_index(index):
    PROGRESS_FILE.write_text(str(index))
    log(f"Progress index saved: {index}")


def save_results(results):
    df = pd.DataFrame(results)
    df.to_excel(OUTPUT_FILE, index=False)
    log(f"Progress saved to: {OUTPUT_FILE}")


def save_debug(page, case_id, label):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_case_id = re.sub(r"[^a-zA-Z0-9_-]", "_", str(case_id))

    screenshot_path = SCREENSHOT_DIR / f"{label}_{safe_case_id}_{timestamp}.png"
    html_path = SCREENSHOT_DIR / f"{label}_{safe_case_id}_{timestamp}.html"

    try:
        page.screenshot(path=str(screenshot_path), full_page=True)
    except Exception:
        pass

    try:
        html_path.write_text(page.content(), encoding="utf-8")
    except Exception:
        pass

    log(f"Debug saved: {screenshot_path}")
    log(f"HTML saved: {html_path}")


def is_main_page(page):
    try:
        url = page.url.lower()

        if "zp_main_idx" in url:
            return True

        html = page.content().lower()
        title = page.title().lower()

        if "civil docket access" in html and "display civil docket report" in html:
            return True

        if "docket access" in title:
            return True

    except Exception:
        return True

    return False


def get_case_description(soup):
    data = {}
    description_anchor = soup.find("a", {"name": "description"})

    if not description_anchor:
        return data

    table = description_anchor.find_next("table")

    if not table:
        return data

    for row in table.find_all("tr"):
        cells = row.find_all("td")

        if len(cells) >= 3:
            label = clean_text(cells[1].get_text()).replace(":", "").lower()
            value = clean_text(cells[2].get_text())

            if label:
                data[label] = value

    return data


def get_parties(soup):
    plaintiff_name = ""
    plaintiff_address = ""
    defendant_name = ""
    defendant_address = ""

    parties_anchor = soup.find("a", {"name": "parties"})

    if not parties_anchor:
        return plaintiff_name, plaintiff_address, defendant_name, defendant_address

    table = parties_anchor.find_next("table")

    if not table:
        return plaintiff_name, plaintiff_address, defendant_name, defendant_address

    rows = table.find_all("tr")

    for i, row in enumerate(rows):
        cells = row.find_all(["td", "th"])
        row_text = [clean_text(cell.get_text(" ")) for cell in cells]

        if len(row_text) >= 5:
            party_type = row_text[3]
            party_name = row_text[4]
            address = ""

            if i + 1 < len(rows):
                next_cells = rows[i + 1].find_all(["td", "th"])

                if len(next_cells) >= 2 and "Address" in clean_text(next_cells[0].get_text()):
                    address = clean_text(next_cells[1].get_text(" "))

            if party_type == "PLAINTIFF":
                plaintiff_name = party_name
                plaintiff_address = address

            elif party_type == "DEFENDANT":
                defendant_name = party_name
                defendant_address = address

    return plaintiff_name, plaintiff_address, defendant_name, defendant_address


def get_docket_data(soup):
    lien_amount = ""
    docket_entry = ""
    pdf_links = []

    dockets_anchor = soup.find("a", {"name": "dockets"})

    if not dockets_anchor:
        return lien_amount, docket_entry, ""

    table = dockets_anchor.find_next("table")

    if not table:
        return lien_amount, docket_entry, ""

    for link in table.find_all("a", href=True):
        href = link.get("href", "")
        text = clean_text(link.get_text())

        if ".pdf" in text.lower() or ".pdf" in href.lower():
            pdf_links.append(text or href)

    for row in table.find_all("tr"):
        cells = row.find_all(["td", "th"])
        texts = [clean_text(cell.get_text(" ")) for cell in cells]

        for text in texts:
            money = extract_money(text)

            if money:
                lien_amount = money

        if len(texts) >= 2 and "Docket Entry" in texts[0]:
            docket_entry = texts[1]

    return lien_amount, docket_entry, ", ".join(pdf_links)


def make_failed_result(case_id, error):
    return {
        "searched_case_id": case_id,
        "status_scrape": "FAILED",
        "error": error,
        "case_id": "",
        "case_caption": "",
        "filing_date": "",
        "court": "",
        "location": "",
        "jury": "",
        "case_type": "",
        "status": "",
        "plaintiff_name": "",
        "plaintiff_address": "",
        "defendant_name": "",
        "defendant_address": "",
        "lien_amount": "",
        "last_docket_entry": "",
        "pdf_links": "",
    }


def open_fresh_search_page(page):
    log("Opening main page...")
    page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(1000)

    log(f"Main page URL: {page.url}")

    try:
        page.get_by_text("Display Civil Docket Report").click(timeout=15000)
    except Exception:
        save_debug(page, "navigation", "display_link_not_found")
        raise Exception("Could not find/click 'Display Civil Docket Report'")

    page.wait_for_selector('input[name="case_id"]', timeout=30000)
    log(f"Search page URL: {page.url}")


def scrape_case_once(page, case_id):
    open_fresh_search_page(page)

    log(f"Filling Case ID: {case_id}")
    page.fill('input[name="case_id"]', case_id)

    log("Submitting search...")
    page.click('input[type="submit"], button[type="submit"]')

    try:
        page.wait_for_selector('a[name="description"]', timeout=45000)

    except PlaywrightTimeoutError:
        log(f"Timeout waiting for result. Current URL: {page.url}")

        if is_main_page(page):
            save_debug(page, case_id, "redirected_main_page")
            raise Exception("Redirected to main page instead of result page")

        save_debug(page, case_id, "result_timeout")
        raise Exception("Result page did not load")

    log(f"Result URL: {page.url}")

    if is_main_page(page):
        save_debug(page, case_id, "main_page_after_submit")
        raise Exception("Search submitted but returned to main page")

    soup = BeautifulSoup(page.content(), "html.parser")
    case_desc = get_case_description(soup)

    if not case_desc.get("case id"):
        save_debug(page, case_id, "empty_result")
        raise Exception("Result loaded but no case data found")

    plaintiff_name, plaintiff_address, defendant_name, defendant_address = get_parties(soup)
    lien_amount, docket_entry, pdf_links = get_docket_data(soup)

    return {
        "searched_case_id": case_id,
        "status_scrape": "SUCCESS",
        "error": "",
        "case_id": case_desc.get("case id", ""),
        "case_caption": case_desc.get("case caption", ""),
        "filing_date": case_desc.get("filing date", ""),
        "court": case_desc.get("court", ""),
        "location": case_desc.get("location", ""),
        "jury": case_desc.get("jury", ""),
        "case_type": case_desc.get("case type", ""),
        "status": case_desc.get("status", ""),
        "plaintiff_name": plaintiff_name,
        "plaintiff_address": plaintiff_address,
        "defendant_name": defendant_name,
        "defendant_address": defendant_address,
        "lien_amount": lien_amount,
        "last_docket_entry": docket_entry,
        "pdf_links": pdf_links,
    }


def scrape_case(page, case_id):
    log(f"\nScraping Case ID: {case_id}")

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            log(f"Attempt {attempt}/{MAX_RETRIES}")
            return scrape_case_once(page, case_id)

        except Exception as e:
            log(f"Attempt failed: {e}")

            if attempt < MAX_RETRIES:
                log(f"Waiting {WAIT_BETWEEN_RETRIES}s before retry...")
                time.sleep(WAIT_BETWEEN_RETRIES)

            else:
                log(f"Failed after {MAX_RETRIES} attempts. Moving to next case.")
                return make_failed_result(case_id, str(e))


def main():
    all_case_ids = load_case_ids()
    start_index = load_start_index()

    batch_case_ids = all_case_ids[start_index:start_index + BATCH_SIZE]

    log(f"Loaded {len(all_case_ids)} Case IDs from {INPUT_FILE}")
    log(f"Starting from index: {start_index}")
    log(f"Processing this batch: {len(batch_case_ids)} cases")

    if not batch_case_ids:
        log("All cases processed.")
        return

    results = load_existing_results()
    consecutive_failures = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=100)
        context = browser.new_context(viewport={"width": 1400, "height": 900})
        page = context.new_page()

        try:
            for batch_index, case_id in enumerate(batch_case_ids, start=1):
                real_index = start_index + batch_index

                log(f"\n===== Batch item {batch_index}/{len(batch_case_ids)} =====")
                log(f"Real index: {real_index}/{len(all_case_ids)}")

                result = scrape_case(page, case_id)
                results.append(result)

                if result["status_scrape"] == "SUCCESS":
                    consecutive_failures = 0
                else:
                    consecutive_failures += 1

                save_results(results)
                save_start_index(real_index)

                if consecutive_failures >= 5:
                    log(f"\n5 failures in a row. Cooling down for {FAILURE_COOLDOWN}s...")
                    time.sleep(FAILURE_COOLDOWN)
                    consecutive_failures = 0

                log(f"Waiting {WAIT_BETWEEN_CASES}s before next case...")
                time.sleep(WAIT_BETWEEN_CASES)

            log(f"\nBatch done. Output saved to: {OUTPUT_FILE}")

        finally:
            log("Batch complete. Closing browser...")
            try:
                page.close()
            except Exception:
                pass

            try:
                context.close()
            except Exception:
                pass

            browser.close()


if __name__ == "__main__":
    main()