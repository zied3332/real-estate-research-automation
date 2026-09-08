from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)

from bs4 import BeautifulSoup

import pandas as pd

from pathlib import Path

import re
import time
import random
import logging
from datetime import datetime, timedelta


# ============================================================
# PATHS
# ============================================================

INPUT_FILE = Path("data/input/case_ids.xlsx")

OUTPUT_DIR = Path("data/output")
SCREENSHOT_DIR = Path("screenshots")
LOG_DIR = Path("logs")

OUTPUT_FILE = OUTPUT_DIR / "cases_results.xlsx"
RETRY_FILE = OUTPUT_DIR / "retry_queue.xlsx"

PROGRESS_FILE = Path("progress.txt")
DONE_FILE = OUTPUT_DIR / "DONE.txt"

LOG_FILE = LOG_DIR / "run.log"


# ============================================================
# WEBSITE
# ============================================================

START_URL = (
    "https://fjdefile.phila.gov/efsfjd/"
    "zk_fjd_public_qry_00.zp_main_idx"
)


# ============================================================
# SCRAPER SETTINGS
# ============================================================

# Number of NEW cases per main.py run
BATCH_SIZE = 10

# Number of old failed cases retried on each run
RETRY_CASES_PER_RUN = 3

# Retries during one immediate scrape attempt
MAX_IMMEDIATE_RETRIES = 3

# Maximum total failures before a case is kept as permanent failure
MAX_TOTAL_RETRY_ATTEMPTS = 8

# Random wait between successful/new cases
MIN_WAIT_BETWEEN_CASES = 3
MAX_WAIT_BETWEEN_CASES = 7

# Base exponential-backoff retry delay
BASE_RETRY_WAIT = 5

# If this many failures happen consecutively,
# pause before continuing.
CONSECUTIVE_FAILURE_LIMIT = 4

FAILURE_COOLDOWN_MIN = 30
FAILURE_COOLDOWN_MAX = 60

# Result page timeout
RESULT_TIMEOUT_MS = 45000

# Navigation timeout
NAVIGATION_TIMEOUT_MS = 60000


# ============================================================
# DIRECTORIES
# ============================================================

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter("%(message)s"))

root_logger = logging.getLogger("")

if not any(
    isinstance(handler, logging.StreamHandler)
    and not isinstance(handler, logging.FileHandler)
    for handler in root_logger.handlers
):
    root_logger.addHandler(console)


def log(message):
    logging.info(message)


# ============================================================
# HELPERS
# ============================================================

def clean_text(text):
    if not text:
        return ""

    return " ".join(
        text.replace("\xa0", " ").split()
    )


def extract_money(text):
    match = re.search(
        r"\$[\d,]+(?:\.\d{2})?",
        text,
    )

    return match.group(0) if match else ""


def random_wait(min_seconds, max_seconds):
    delay = random.uniform(
        min_seconds,
        max_seconds,
    )

    log(
        f"Waiting {delay:.1f}s..."
    )

    time.sleep(delay)


# ============================================================
# INPUT
# ============================================================

def load_case_ids():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input file does not exist: {INPUT_FILE}"
        )

    df = pd.read_excel(INPUT_FILE)

    log(
        f"Columns found: {df.columns.tolist()}"
    )

    if df.empty:
        raise Exception(
            "Input Excel contains no rows."
        )

    column = (
        "Case ID"
        if "Case ID" in df.columns
        else (
            "case_id"
            if "case_id" in df.columns
            else df.columns[0]
        )
    )

    case_ids = (
        df[column]
        .dropna()
        .astype(str)
        .str.strip()
        .tolist()
    )

    if not case_ids:
        raise Exception(
            "No Case IDs found in input Excel."
        )

    return case_ids


# ============================================================
# PROGRESS
# ============================================================

def load_start_index():
    if not PROGRESS_FILE.exists():
        return 0

    text = (
        PROGRESS_FILE
        .read_text()
        .strip()
    )

    if text.isdigit():
        return int(text)

    return 0


def save_start_index(index):
    PROGRESS_FILE.write_text(
        str(index)
    )

    log(
        f"Progress index saved: {index}"
    )


# ============================================================
# RESULTS
# ============================================================

RESULT_COLUMNS = [
    "searched_case_id",
    "status_scrape",
    "error",
    "case_id",
    "case_caption",
    "filing_date",
    "court",
    "location",
    "jury",
    "case_type",
    "status",
    "plaintiff_name",
    "plaintiff_address",
    "defendant_name",
    "defendant_address",
    "lien_amount",
    "last_docket_entry",
    "pdf_links",
]


def load_existing_results():
    if not OUTPUT_FILE.exists():
        return []

    try:
        df = pd.read_excel(
            OUTPUT_FILE
        )

        results = (
            df.to_dict("records")
        )

        log(
            f"Loaded existing result rows: {len(results)}"
        )

        return results

    except Exception as e:
        log(
            f"Could not read existing output: {e}"
        )

        return []


def upsert_result(results, result):
    searched_case_id = str(
        result.get(
            "searched_case_id",
            "",
        )
    )

    for index, existing in enumerate(results):
        existing_id = str(
            existing.get(
                "searched_case_id",
                "",
            )
        )

        if existing_id == searched_case_id:
            results[index] = result
            return results

    results.append(result)

    return results


def save_results(results):
    df = pd.DataFrame(
        results,
        columns=RESULT_COLUMNS,
    )

    df.to_excel(
        OUTPUT_FILE,
        index=False,
    )

    log(
        f"Results saved: {OUTPUT_FILE}"
    )


# ============================================================
# RETRY QUEUE
# ============================================================

RETRY_COLUMNS = [
    "case_id",
    "attempts",
    "last_error",
    "last_attempt",
    "next_retry",
]


def load_retry_queue():
    if not RETRY_FILE.exists():
        return pd.DataFrame(
            columns=RETRY_COLUMNS
        )

    try:
        df = pd.read_excel(
            RETRY_FILE
        )

        for column in RETRY_COLUMNS:
            if column not in df.columns:
                df[column] = ""

        return df[
            RETRY_COLUMNS
        ]

    except Exception as e:
        log(
            f"Could not load retry queue: {e}"
        )

        return pd.DataFrame(
            columns=RETRY_COLUMNS
        )


def save_retry_queue(df):
    df.to_excel(
        RETRY_FILE,
        index=False,
    )

    log(
        f"Retry queue saved. Pending: {len(df)}"
    )


def add_to_retry_queue(
    retry_df,
    case_id,
    error,
):
    case_id = str(case_id)

    matching = (
        retry_df["case_id"]
        .astype(str)
        == case_id
    )

    now = datetime.now()

    if matching.any():
        row_index = retry_df[
            matching
        ].index[0]

        try:
            attempts = int(
                retry_df.at[
                    row_index,
                    "attempts",
                ]
            )
        except Exception:
            attempts = 0

        attempts += 1

    else:
        attempts = 1
        row_index = None

    # Increasing delay:
    # 1 failure -> ~1 minute
    # 2 failures -> ~2 minutes
    # 3 failures -> ~4 minutes
    # etc.
    delay_minutes = min(
        2 ** (attempts - 1),
        60,
    )

    # small random jitter
    delay_minutes += random.uniform(
        0,
        1,
    )

    next_retry = (
        now
        + timedelta(
            minutes=delay_minutes
        )
    )

    data = {
        "case_id": case_id,
        "attempts": attempts,
        "last_error": str(error),
        "last_attempt": now.isoformat(
            timespec="seconds"
        ),
        "next_retry": next_retry.isoformat(
            timespec="seconds"
        ),
    }

    if row_index is not None:

        for key, value in data.items():
            retry_df.at[
                row_index,
                key,
            ] = value

    else:
        retry_df = pd.concat(
            [
                retry_df,
                pd.DataFrame([data]),
            ],
            ignore_index=True,
        )

    return retry_df


def remove_from_retry_queue(
    retry_df,
    case_id,
):
    case_id = str(case_id)

    if retry_df.empty:
        return retry_df

    retry_df = retry_df[
        retry_df["case_id"]
        .astype(str)
        != case_id
    ]

    return retry_df.reset_index(
        drop=True
    )


def get_due_retry_cases(
    retry_df,
    limit,
):
    if retry_df.empty:
        return []

    now = datetime.now()

    due = []

    for _, row in retry_df.iterrows():

        try:
            attempts = int(
                row["attempts"]
            )
        except Exception:
            attempts = 0

        if attempts >= MAX_TOTAL_RETRY_ATTEMPTS:
            continue

        try:
            next_retry = datetime.fromisoformat(
                str(row["next_retry"])
            )
        except Exception:
            next_retry = now

        if next_retry <= now:
            due.append(
                str(row["case_id"])
            )

        if len(due) >= limit:
            break

    return due


# ============================================================
# DEBUG
# ============================================================

def save_debug(
    page,
    case_id,
    label,
):
    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    safe_case_id = re.sub(
        r"[^a-zA-Z0-9_-]",
        "_",
        str(case_id),
    )

    screenshot_path = (
        SCREENSHOT_DIR
        / f"{label}_{safe_case_id}_{timestamp}.png"
    )

    html_path = (
        SCREENSHOT_DIR
        / f"{label}_{safe_case_id}_{timestamp}.html"
    )

    try:
        page.screenshot(
            path=str(
                screenshot_path
            ),
            full_page=True,
        )

    except Exception:
        pass

    try:
        html_path.write_text(
            page.content(),
            encoding="utf-8",
        )

    except Exception:
        pass

    log(
        f"Debug screenshot: {screenshot_path}"
    )

    log(
        f"Debug HTML: {html_path}"
    )


# ============================================================
# PAGE DETECTION
# ============================================================

def is_main_page(page):
    try:

        url = (
            page.url
            .lower()
        )

        if "zp_main_idx" in url:
            return True

        html = (
            page.content()
            .lower()
        )

        title = (
            page.title()
            .lower()
        )

        if (
            "civil docket access" in html
            and
            "display civil docket report" in html
        ):
            return True

        if (
            "docket access" in title
            and
            "zp_dktrpt_search" not in url
        ):
            return True

    except Exception:
        return True

    return False


# ============================================================
# PARSING
# ============================================================

def get_case_description(soup):
    data = {}

    description_anchor = soup.find(
        "a",
        {"name": "description"},
    )

    if not description_anchor:
        return data

    table = (
        description_anchor
        .find_next("table")
    )

    if not table:
        return data

    for row in table.find_all("tr"):

        cells = row.find_all(
            "td"
        )

        if len(cells) >= 3:

            label = (
                clean_text(
                    cells[1].get_text()
                )
                .replace(":", "")
                .lower()
            )

            value = clean_text(
                cells[2].get_text()
            )

            if label:
                data[label] = value

    return data


def get_parties(soup):
    plaintiff_name = ""
    plaintiff_address = ""

    defendant_name = ""
    defendant_address = ""

    parties_anchor = soup.find(
        "a",
        {"name": "parties"},
    )

    if not parties_anchor:
        return (
            plaintiff_name,
            plaintiff_address,
            defendant_name,
            defendant_address,
        )

    table = (
        parties_anchor
        .find_next("table")
    )

    if not table:
        return (
            plaintiff_name,
            plaintiff_address,
            defendant_name,
            defendant_address,
        )

    rows = table.find_all("tr")

    for i, row in enumerate(rows):

        cells = row.find_all(
            ["td", "th"]
        )

        row_text = [
            clean_text(
                cell.get_text(" ")
            )
            for cell in cells
        ]

        if len(row_text) < 5:
            continue

        party_type = (
            row_text[3]
            .strip()
            .upper()
        )

        party_name = (
            row_text[4]
            .strip()
        )

        address = ""

        if i + 1 < len(rows):

            next_cells = (
                rows[i + 1]
                .find_all(
                    ["td", "th"]
                )
            )

            if len(next_cells) >= 2:

                first_text = clean_text(
                    next_cells[0]
                    .get_text(" ")
                )

                if "address" in first_text.lower():

                    address = clean_text(
                        next_cells[1]
                        .get_text(" ")
                    )

        if party_type == "PLAINTIFF":

            if not plaintiff_name:
                plaintiff_name = party_name
                plaintiff_address = address

        elif party_type == "DEFENDANT":

            if not defendant_name:
                defendant_name = party_name
                defendant_address = address

    return (
        plaintiff_name,
        plaintiff_address,
        defendant_name,
        defendant_address,
    )


def get_docket_data(soup):
    lien_amount = ""
    docket_entry = ""

    pdf_links = []

    dockets_anchor = soup.find(
        "a",
        {"name": "dockets"},
    )

    if not dockets_anchor:
        return (
            lien_amount,
            docket_entry,
            "",
        )

    table = (
        dockets_anchor
        .find_next("table")
    )

    if not table:
        return (
            lien_amount,
            docket_entry,
            "",
        )

    for link in table.find_all(
        "a",
        href=True,
    ):

        href = link.get(
            "href",
            "",
        )

        text = clean_text(
            link.get_text()
        )

        if (
            ".pdf" in text.lower()
            or
            ".pdf" in href.lower()
        ):
            pdf_links.append(
                href
            )

    rows = table.find_all(
        "tr"
    )

    # Search every docket row
    for row in rows:

        cells = row.find_all(
            ["td", "th"]
        )

        texts = [
            clean_text(
                cell.get_text(" ")
            )
            for cell in cells
        ]

        for text in texts:

            money = extract_money(
                text
            )

            if money:
                lien_amount = money

        # Keep latest meaningful docket text
        meaningful = [
            text
            for text in texts
            if text
        ]

        if meaningful:
            docket_entry = " | ".join(
                meaningful
            )

    return (
        lien_amount,
        docket_entry,
        ", ".join(pdf_links),
    )


# ============================================================
# RESULT MODELS
# ============================================================

def make_failed_result(
    case_id,
    error,
):
    return {
        "searched_case_id": case_id,
        "status_scrape": "RETRY",
        "error": str(error),

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


# ============================================================
# NAVIGATION
# ============================================================

def open_fresh_search_page(page):

    log(
        "Opening Civil Docket main page..."
    )

    page.goto(
        START_URL,
        wait_until="domcontentloaded",
        timeout=NAVIGATION_TIMEOUT_MS,
    )

    page.wait_for_timeout(
        random.randint(
            1000,
            2000,
        )
    )

    log(
        f"Main page URL: {page.url}"
    )

    try:

        page.get_by_text(
            "Display Civil Docket Report"
        ).click(
            timeout=15000
        )

    except Exception:

        save_debug(
            page,
            "navigation",
            "display_link_not_found",
        )

        raise Exception(
            "Could not open Civil Docket Report search."
        )

    page.wait_for_selector(
        'input[name="case_id"]',
        timeout=30000,
    )

    log(
        f"Search page URL: {page.url}"
    )


# ============================================================
# SCRAPE ONE ATTEMPT
# ============================================================

def scrape_case_once(
    page,
    case_id,
):

    open_fresh_search_page(
        page
    )

    log(
        f"Filling Case ID: {case_id}"
    )

    page.fill(
        'input[name="case_id"]',
        str(case_id),
    )

    page.wait_for_timeout(
        random.randint(
            300,
            900,
        )
    )

    log(
        "Submitting search..."
    )

    page.click(
        'input[type="submit"], '
        'button[type="submit"]'
    )

    try:

        page.wait_for_selector(
            'a[name="description"]',
            timeout=RESULT_TIMEOUT_MS,
        )

    except PlaywrightTimeoutError:

        log(
            f"Timeout. Current URL: {page.url}"
        )

        if is_main_page(page):

            save_debug(
                page,
                case_id,
                "redirected_main_page",
            )

            raise Exception(
                "Redirected to main page"
            )

        save_debug(
            page,
            case_id,
            "result_timeout",
        )

        raise Exception(
            "Result page timeout"
        )

    log(
        f"Result URL: {page.url}"
    )

    if is_main_page(page):

        save_debug(
            page,
            case_id,
            "main_page_after_submit",
        )

        raise Exception(
            "Search returned to main page"
        )

    html = page.content()

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    case_desc = get_case_description(
        soup
    )

    if not case_desc.get(
        "case id"
    ):

        save_debug(
            page,
            case_id,
            "empty_result",
        )

        raise Exception(
            "Result page loaded but case data was empty"
        )

    (
        plaintiff_name,
        plaintiff_address,
        defendant_name,
        defendant_address,
    ) = get_parties(
        soup
    )

    (
        lien_amount,
        docket_entry,
        pdf_links,
    ) = get_docket_data(
        soup
    )

    return {

        "searched_case_id": case_id,

        "status_scrape": "SUCCESS",

        "error": "",

        "case_id": case_desc.get(
            "case id",
            "",
        ),

        "case_caption": case_desc.get(
            "case caption",
            "",
        ),

        "filing_date": case_desc.get(
            "filing date",
            "",
        ),

        "court": case_desc.get(
            "court",
            "",
        ),

        "location": case_desc.get(
            "location",
            "",
        ),

        "jury": case_desc.get(
            "jury",
            "",
        ),

        "case_type": case_desc.get(
            "case type",
            "",
        ),

        "status": case_desc.get(
            "status",
            "",
        ),

        "plaintiff_name":
            plaintiff_name,

        "plaintiff_address":
            plaintiff_address,

        "defendant_name":
            defendant_name,

        "defendant_address":
            defendant_address,

        "lien_amount":
            lien_amount,

        "last_docket_entry":
            docket_entry,

        "pdf_links":
            pdf_links,
    }


# ============================================================
# SCRAPE WITH RETRIES
# ============================================================

def scrape_case(
    page,
    case_id,
):

    log(
        f"\nScraping Case ID: {case_id}"
    )

    last_error = ""

    for attempt in range(
        1,
        MAX_IMMEDIATE_RETRIES + 1,
    ):

        try:

            log(
                f"Attempt "
                f"{attempt}/"
                f"{MAX_IMMEDIATE_RETRIES}"
            )

            return scrape_case_once(
                page,
                case_id,
            )

        except Exception as e:

            last_error = str(e)

            log(
                f"Attempt failed: {last_error}"
            )

            if (
                attempt
                < MAX_IMMEDIATE_RETRIES
            ):

                # 5s, 10s, 20s + jitter
                delay = (
                    BASE_RETRY_WAIT
                    * (2 ** (attempt - 1))
                )

                delay += random.uniform(
                    1,
                    4,
                )

                log(
                    f"Retrying after "
                    f"{delay:.1f}s..."
                )

                time.sleep(
                    delay
                )

    log(
        "Temporary failure. "
        "Adding case to retry queue."
    )

    return make_failed_result(
        case_id,
        last_error,
    )


# ============================================================
# BROWSER
# ============================================================

def create_browser_session(
    playwright,
):

    browser = playwright.chromium.launch(
        headless=False,
        slow_mo=50,
    )

    context = browser.new_context(
        viewport={
            "width": 1400,
            "height": 900,
        },
    )

    page = context.new_page()

    return (
        browser,
        context,
        page,
    )


def close_browser_session(
    browser,
    context,
    page,
):

    try:
        page.close()
    except Exception:
        pass

    try:
        context.close()
    except Exception:
        pass

    try:
        browser.close()
    except Exception:
        pass


# ============================================================
# MAIN
# ============================================================

def main():

    all_case_ids = load_case_ids()

    start_index = load_start_index()

    results = load_existing_results()

    retry_df = load_retry_queue()

    due_retries = get_due_retry_cases(
        retry_df,
        RETRY_CASES_PER_RUN,
    )

    new_case_ids = all_case_ids[
        start_index:
        start_index + BATCH_SIZE
    ]

    log(
        f"\nLoaded {len(all_case_ids)} total Case IDs"
    )

    log(
        f"Starting main queue at: {start_index}"
    )

    log(
        f"New cases this run: {len(new_case_ids)}"
    )

    log(
        f"Retry cases this run: {len(due_retries)}"
    )

    if DONE_FILE.exists():
        DONE_FILE.unlink()

    # Nothing left anywhere
    if (
        not new_case_ids
        and retry_df.empty
    ):

        log(
            "\nAll cases processed successfully."
        )

        DONE_FILE.write_text(
            "All cases processed."
        )

        return

    with sync_playwright() as p:

        (
            browser,
            context,
            page,
        ) = create_browser_session(
            p
        )

        consecutive_failures = 0

        try:

            # =================================================
            # RETRY OLD FAILURES FIRST
            # =================================================

            for retry_case_id in due_retries:

                log(
                    "\n================================"
                )

                log(
                    f"Retry queue case: "
                    f"{retry_case_id}"
                )

                log(
                    "================================"
                )

                result = scrape_case(
                    page,
                    retry_case_id,
                )

                results = upsert_result(
                    results,
                    result,
                )

                if (
                    result["status_scrape"]
                    == "SUCCESS"
                ):

                    log(
                        "Retry succeeded."
                    )

                    retry_df = (
                        remove_from_retry_queue(
                            retry_df,
                            retry_case_id,
                        )
                    )

                    consecutive_failures = 0

                else:

                    retry_df = (
                        add_to_retry_queue(
                            retry_df,
                            retry_case_id,
                            result["error"],
                        )
                    )

                    consecutive_failures += 1

                save_results(
                    results
                )

                save_retry_queue(
                    retry_df
                )

                random_wait(
                    MIN_WAIT_BETWEEN_CASES,
                    MAX_WAIT_BETWEEN_CASES,
                )

            # =================================================
            # PROCESS NEW CASES
            # =================================================

            for batch_index, case_id in enumerate(
                new_case_ids,
                start=1,
            ):

                real_index = (
                    start_index
                    + batch_index
                )

                log(
                    "\n================================"
                )

                log(
                    f"Batch item "
                    f"{batch_index}/"
                    f"{len(new_case_ids)}"
                )

                log(
                    f"Real index: "
                    f"{real_index}/"
                    f"{len(all_case_ids)}"
                )

                log(
                    "================================"
                )

                result = scrape_case(
                    page,
                    case_id,
                )

                results = upsert_result(
                    results,
                    result,
                )

                if (
                    result["status_scrape"]
                    == "SUCCESS"
                ):

                    consecutive_failures = 0

                    retry_df = (
                        remove_from_retry_queue(
                            retry_df,
                            case_id,
                        )
                    )

                else:

                    consecutive_failures += 1

                    retry_df = (
                        add_to_retry_queue(
                            retry_df,
                            case_id,
                            result["error"],
                        )
                    )

                # Save immediately
                save_results(
                    results
                )

                save_retry_queue(
                    retry_df
                )

                # Main queue always progresses,
                # because failures are now safely
                # stored in retry_queue.xlsx.
                save_start_index(
                    real_index
                )

                # Site/session appears unhealthy
                if (
                    consecutive_failures
                    >= CONSECUTIVE_FAILURE_LIMIT
                ):

                    cooldown = random.randint(
                        FAILURE_COOLDOWN_MIN,
                        FAILURE_COOLDOWN_MAX,
                    )

                    log(
                        f"\n"
                        f"{consecutive_failures} "
                        f"consecutive failures."
                    )

                    log(
                        f"Cooling down "
                        f"{cooldown}s and "
                        f"restarting browser session..."
                    )

                    close_browser_session(
                        browser,
                        context,
                        page,
                    )

                    time.sleep(
                        cooldown
                    )

                    (
                        browser,
                        context,
                        page,
                    ) = create_browser_session(
                        p
                    )

                    consecutive_failures = 0

                random_wait(
                    MIN_WAIT_BETWEEN_CASES,
                    MAX_WAIT_BETWEEN_CASES,
                )

            log(
                "\nBatch finished."
            )

            log(
                f"Results: {OUTPUT_FILE}"
            )

            log(
                f"Retry queue: {RETRY_FILE}"
            )

            # Check whether everything is finished
            latest_index = load_start_index()

            retry_df = load_retry_queue()

            if (
                latest_index
                >= len(all_case_ids)
                and retry_df.empty
            ):

                DONE_FILE.write_text(
                    "All cases processed."
                )

                log(
                    "\nAll cases finished."
                )

        finally:

            log(
                "Closing browser..."
            )

            close_browser_session(
                browser,
                context,
                page,
            )


if __name__ == "__main__":
    main()