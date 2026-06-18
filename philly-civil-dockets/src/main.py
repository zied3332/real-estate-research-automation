from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from bs4 import BeautifulSoup
import pandas as pd
from pathlib import Path
import re

CASE_IDS = [
    "141202168",
    "141202169",
    "141202170",
]

URL = "https://fjdefile.phila.gov/dockets/zk_fjd_public_qry_03.zp_dktrpt_setup_idx?uid=BaiXGnvECIolqsSYDxam&o=HcobAXl4Y!rzHdb"

OUTPUT_DIR = Path("data/output")
SCREENSHOT_DIR = Path("screenshots")

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)


def clean_text(text):
    if not text:
        return ""
    return " ".join(text.replace("\xa0", " ").split())


def extract_money(text):
    match = re.search(r"\$[\d,]+(?:\.\d{2})?", text)
    return match.group(0) if match else ""


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
                next_row = rows[i + 1]
                next_cells = next_row.find_all(["td", "th"])

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
        if ".pdf" in link.get_text().lower():
            pdf_links.append(clean_text(link.get_text()))

    rows = table.find_all("tr")

    for row in rows:
        cells = row.find_all(["td", "th"])
        texts = [clean_text(cell.get_text(" ")) for cell in cells]

        for text in texts:
            money = extract_money(text)
            if money:
                lien_amount = money

        if len(texts) >= 2 and "Docket Entry" in texts[0]:
            docket_entry = texts[1]

    return lien_amount, docket_entry, ", ".join(pdf_links)


def scrape_case(page, case_id):
    print(f"\nScraping Case ID: {case_id}")

    page.goto(URL, wait_until="domcontentloaded", timeout=60000)

    page.wait_for_selector('input[name="case_id"]', timeout=30000)
    page.fill('input[name="case_id"]', case_id)

    page.click('input[type="submit"], button[type="submit"]')

    try:
        page.wait_for_selector('a[name="description"]', timeout=60000)
    except PlaywrightTimeoutError:
        page.screenshot(path=str(SCREENSHOT_DIR / f"error_{case_id}.png"), full_page=True)

        return {
            "searched_case_id": case_id,
            "status_scrape": "FAILED",
            "error": "Result page did not load",
        }

    page.wait_for_timeout(2000)

    soup = BeautifulSoup(page.content(), "html.parser")

    case_desc = get_case_description(soup)
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


with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False,
        slow_mo=250
    )

    context = browser.new_context(
        viewport={"width": 1400, "height": 900}
    )

    page = context.new_page()

    results = []

    try:
        for case_id in CASE_IDS:
            result = scrape_case(page, case_id)
            results.append(result)

            print("Result:", result.get("status_scrape"))
            page.wait_for_timeout(2000)

        df = pd.DataFrame(results)

        output_file = OUTPUT_DIR / "cases_results.xlsx"
        df.to_excel(output_file, index=False)

        print(f"\nSaved Excel file to: {output_file}")

    finally:
        input("\nPress Enter to close browser...")
        browser.close()