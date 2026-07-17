"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
collect_partner_inventory.py

Purpose:
Collects every Cisco partner matching the configured
portfolio and competency filters and stores a clean,
de-duplicated inventory for subsequent processing.

Runs unattended on a bi-weekly schedule via GitHub Actions.

Author:
Taksh Jain

Version:
2.1.0
===========================================================
"""

import math
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
import requests

from config import (
    COMPETENCY_IDS,
    CSV_INDEX,
    FILE_ENCODING,
    INVENTORY_COLUMNS,
    LANGUAGE,
    MAX_RETRIES,
    PAGE_SIZE,
    PARTNER_INVENTORY_FILE,
    PARTNER_SEARCH_API,
    PORTFOLIO_IDS,
    PROJECT_NAME,
    RANDOM_SEED,
    REQUEST_DELAY,
    REQUEST_HEADERS,
    REQUEST_TIMEOUT,
    RETRY_DELAY,
    SECONDS_PER_HOUR,
    SECONDS_PER_MINUTE,
    SUMMARY_BORDER,
    TIMESTAMP_FORMAT,
    VERSION,
)
from logger import log

# Column used to identify and order partners in the inventory.
PARTNER_NAME_COLUMN = "partner_name"


# =========================================================
# Value Helpers
# =========================================================

def _format_duration(seconds: float) -> str:
    """Render a duration in seconds as a compact ``h m s`` string."""
    total_seconds = int(round(seconds))
    hours, remainder = divmod(total_seconds, SECONDS_PER_HOUR)
    minutes, secs = divmod(remainder, SECONDS_PER_MINUTE)
    if hours:
        return f"{hours}h {minutes}m {secs}s"
    if minutes:
        return f"{minutes}m {secs}s"
    return f"{secs}s"


# =========================================================
# HTTP Session & Requests
# =========================================================

def create_session() -> requests.Session:
    """Build a reusable session with the standard request headers."""
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)
    return session


def build_payload(page: int) -> Dict[str, Any]:
    """Build the getPfPartners request body for a single page."""
    return {
        "LOCATION_IDS": "",
        "CITY_STATE_IDS": "",
        "GLOBAL_STATUS_IDS": "",
        "PORTFOLIO_IDS": PORTFOLIO_IDS,
        "COMPETENCY_IDS": COMPETENCY_IDS,
        "PARTNER_TYPE_IDS": "",
        "CISCO_POWERED_SERVICE_IDS": "",
        "TECHNOLOGY_IDS": "",
        "SPECIALIZATION_IDS": "",
        "PARTNER_KEYWORD_IDS": "",
        "PARTNER_DESIGNATION_IDS": "",
        "INDUSTRY_IDS": "",
        "COMPANYSIZE_IDS": "",
        "pageSize": PAGE_SIZE,
        "pageCount": page,
        "SORT_BY": "partner_name.sort",
        "SORT_ORDER": "random",
        "LANGUAGE_CD": LANGUAGE,
        "RANDOM_SEED": RANDOM_SEED,
    }


def fetch_total_matches(session: requests.Session) -> int:
    """Fetch the total partner count from the first search page.

    Raises:
        RuntimeError: If the response is malformed or the match count
            is missing or invalid.
    """
    response = session.post(
        PARTNER_SEARCH_API,
        json=build_payload(1),
        timeout=REQUEST_TIMEOUT,
    )
    response.raise_for_status()
    data = response.json()

    if not isinstance(data, dict) or "matches" not in data:
        raise RuntimeError(
            "Partner search response does not contain 'matches'."
        )

    matches = data.get("matches")
    if not isinstance(matches, int) or matches < 0:
        raise RuntimeError(
            f"Partner search returned an invalid match count: {matches!r}"
        )

    return matches


def _extract_page_partners(page_data: Any) -> List[Dict[str, Any]]:
    """Extract inventory rows from a single search page response.

    Raises:
        RuntimeError: If the response is malformed or missing 'data'.
    """
    if not isinstance(page_data, dict):
        raise RuntimeError(
            "Partner search response was not a JSON object."
        )

    records = page_data.get("data")
    if not isinstance(records, list):
        raise RuntimeError(
            "Partner search response is missing the 'data' array."
        )

    partners: List[Dict[str, Any]] = []
    for partner in records:
        if not isinstance(partner, dict):
            continue
        raw = {
            "partner_name": partner.get("partner_name"),
            "be_geo_id": partner.get("be_geo_id"),
            "location_id": partner.get("location_keyword_ids"),
            # Kept for future use if Cisco exposes it.
            "partner_id": partner.get("partner_id"),
        }
        partners.append(
            {
                key: ("" if value is None else str(value).strip())
                for key, value in raw.items()
            }
        )
    return partners


def fetch_page(
    session: requests.Session,
    page: int,
    total_pages: int,
) -> List[Dict[str, Any]]:
    """Fetch and validate one search page, with retries.

    Retries up to ``MAX_RETRIES`` times, waiting ``RETRY_DELAY`` seconds
    between attempts. Every attempt is logged. A page that cannot be
    retrieved is fatal to the run and raises after the final attempt.
    """
    last_error: Exception | None = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.post(
                PARTNER_SEARCH_API,
                json=build_payload(page),
                timeout=REQUEST_TIMEOUT,
            )
            response.raise_for_status()
            return _extract_page_partners(response.json())
        except Exception as exc:  # noqa: BLE001 - retried, then re-raised
            last_error = exc
            log(
                f"Page {page}/{total_pages} - attempt "
                f"{attempt}/{MAX_RETRIES} failed: {exc}",
                "WARNING",
            )
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY)

    raise RuntimeError(
        f"Unable to retrieve page {page}/{total_pages} after "
        f"{MAX_RETRIES} attempts: {last_error}"
    )


# =========================================================
# Persistence
# =========================================================

def save_inventory(partners: List[Dict[str, Any]]) -> pd.DataFrame:
    """Validate, de-duplicate, sort and persist the partner inventory.

    Rows without a partner name are dropped, exact duplicates are
    removed, and the result is sorted by ``partner_name`` before being
    written with the fixed inventory schema.
    """
    frame = pd.DataFrame(partners, columns=INVENTORY_COLUMNS)

    frame = frame[frame[PARTNER_NAME_COLUMN].notna()]

    before_count = len(frame)
    frame = frame.drop_duplicates()
    duplicates_removed = before_count - len(frame)
    log(f"Duplicate Partners Removed : {duplicates_removed}")

    frame = frame.sort_values(
        by=PARTNER_NAME_COLUMN,
        key=lambda column: column.str.casefold(),
        kind="stable",
        na_position="last",
    )
    frame = frame.reset_index(drop=True)

    if frame.empty:
        log(
            "Inventory is empty after collection; writing an empty file.",
            "WARNING",
        )

    output_path = Path(PARTNER_INVENTORY_FILE)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_csv(
        output_path,
        index=CSV_INDEX,
        encoding=FILE_ENCODING,
    )
    return frame


# =========================================================
# Reporting
# =========================================================

def print_summary(
    total_matches: int,
    total_pages: int,
    successful_pages: int,
    failed_pages: int,
    saved_count: int,
    duration: float,
) -> None:
    """Log the final inventory-collection summary block."""
    minutes = duration / SECONDS_PER_MINUTE
    pages_per_minute = (total_pages / minutes) if minutes > 0 else 0.0

    log(SUMMARY_BORDER)
    log("Partner Inventory Collection Complete")
    log(f"Total Partners Found : {total_matches}")
    log(f"Partners Saved       : {saved_count}")
    log(f"Pages Processed      : {total_pages}")
    log(f"Successful Pages     : {successful_pages}")
    log(f"Failed Pages         : {failed_pages}")
    log(f"Duration             : {_format_duration(duration)}")
    log(f"Pages Per Minute     : {pages_per_minute:.1f}")
    log(f"Output File          : {PARTNER_INVENTORY_FILE}")
    log(SUMMARY_BORDER)


# =========================================================
# Orchestration
# =========================================================

def main() -> None:
    """Collect the full partner inventory and persist it to disk."""
    start_time = time.perf_counter()
    start_timestamp = datetime.now().strftime(TIMESTAMP_FORMAT)

    log(SUMMARY_BORDER)
    log(PROJECT_NAME)
    log(f"Version         : {VERSION}")
    log(f"Start Time      : {start_timestamp}")
    log(f"Search API      : {PARTNER_SEARCH_API}")
    log(f"Inventory File  : {PARTNER_INVENTORY_FILE}")
    log(SUMMARY_BORDER)

    all_partners: List[Dict[str, Any]] = []
    successful_pages = 0
    failed_pages = 0

    with create_session() as session:
        total_matches = fetch_total_matches(session)
        total_pages = math.ceil(total_matches / PAGE_SIZE) if PAGE_SIZE else 0

        if total_pages == 0:
            raise RuntimeError(
                "Cisco returned zero pages although partners were found."
            )

        log(f"Total partners found: {total_matches}")
        log(f"Total pages to process: {total_pages}")

        for page in range(1, total_pages + 1):
            start_partner = ((page - 1) * PAGE_SIZE) + 1
            end_partner = min(page * PAGE_SIZE, total_matches)
            log(
                f"Processing page {page}/{total_pages} "
                f"({start_partner}-{end_partner})"
            )

            page_partners = fetch_page(session, page, total_pages)
            all_partners.extend(page_partners)
            successful_pages += 1

            time.sleep(REQUEST_DELAY)

    frame = save_inventory(all_partners)

    duration = time.perf_counter() - start_time
    print_summary(
        total_matches,
        total_pages,
        successful_pages,
        failed_pages,
        len(frame),
        duration,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001 - logged before fatal re-raise
        log("Fatal error during partner inventory collection.", "ERROR")
        log(str(exc), "ERROR")
        raise
