"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
get_partner_details.py

Purpose:
Master data-collection layer. Retrieves detailed overview
information for every Cisco partner listed in the inventory
and writes a clean, well-typed dataset that every downstream
stage (change detection, Excel build, executive email) relies
on.

Runs unattended on a bi-weekly schedule via GitHub Actions.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
import requests

from config import (
    CATEGORY_COMPETENCY,
    CATEGORY_PORTFOLIO,
    CHECKPOINT_INTERVAL,
    CSV_INDEX,
    FAILED_COLUMNS,
    FAILED_PARTNERS_FILE,
    FAILURE_THRESHOLD,
    FILE_ENCODING,
    HTTPS_PREFIX,
    LANGUAGE,
    LOCATION_SEPARATOR,
    MAX_RETRIES,
    OUTPUT_COLUMNS,
    PARTNER_DETAILS_FILE,
    PARTNER_DETAILS_PREVIOUS_FILE,
    PARTNER_INVENTORY_FILE,
    PARTNER_OVERVIEW_API,
    PROJECT_NAME,
    QUALIFICATION_SEPARATOR,
    REQUEST_DELAY,
    REQUEST_HEADERS,
    REQUEST_TIMEOUT,
    REQUIRED_INVENTORY_COLUMNS,
    RETRY_DELAY,
    SAVE_BACKUP,
    SECONDS_PER_HOUR,
    SECONDS_PER_MINUTE,
    SUMMARY_BORDER,
    TIMESTAMP_FORMAT,
    VALID_URL_SCHEMES,
    VERSION,
)
from logger import log


class PartnerDetailError(Exception):
    """Raised when a partner's details cannot be retrieved or parsed."""


# =========================================================
# Value Helpers
# =========================================================

def _clean(value: Any) -> str:
    """Return a trimmed string, treating ``None`` and NaN as empty."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        # ``pd.isna`` raises on unhashable / array-like values; those
        # are never empty, so fall through to string conversion.
        pass
    return str(value).strip()


def _normalize_website(value: Any) -> str:
    """Trim a website value and ensure it carries a usable scheme.

    A non-empty address without a recognized scheme is prefixed with
    ``https://``. Any value that already begins with a known scheme is
    returned untouched.
    """
    website = _clean(value)
    if not website:
        return ""
    if website.lower().startswith(VALID_URL_SCHEMES):
        return website
    return f"{HTTPS_PREFIX}{website}"


def _build_location(city: str, country: str) -> str:
    """Compose a ``City, Country`` label, skipping missing parts."""
    parts = [part for part in (city, country) if part]
    return LOCATION_SEPARATOR.join(parts)


def _extract_qualifications(qualifications: Any) -> Tuple[str, str]:
    """Split partner qualifications into portfolios and competencies.

    Values are de-duplicated, sorted alphabetically (case-insensitive)
    and joined with the qualification separator. Unknown categories are
    ignored so new Cisco categories never break collection.
    """
    portfolios: set = set()
    competencies: set = set()

    if not isinstance(qualifications, list):
        return "", ""

    for item in qualifications:
        if not isinstance(item, dict):
            continue
        category = _clean(item.get("category_display_name")).upper()
        name = _clean(item.get("keyword_display_name"))
        if not name:
            continue
        if category == CATEGORY_PORTFOLIO:
            portfolios.add(name)
        elif category == CATEGORY_COMPETENCY:
            competencies.add(name)

    portfolio_value = QUALIFICATION_SEPARATOR.join(
        sorted(portfolios, key=str.casefold)
    )
    competency_value = QUALIFICATION_SEPARATOR.join(
        sorted(competencies, key=str.casefold)
    )
    return portfolio_value, competency_value


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
# Filesystem & Inventory
# =========================================================

def ensure_directories() -> None:
    """Create the parent directories for every output artifact."""
    for file_path in (
        PARTNER_DETAILS_FILE,
        PARTNER_DETAILS_PREVIOUS_FILE,
        FAILED_PARTNERS_FILE,
    ):
        Path(file_path).parent.mkdir(parents=True, exist_ok=True)


def backup_previous_file() -> None:
    """Copy the current details file aside before it is overwritten."""
    if not SAVE_BACKUP:
        log("Backup disabled in configuration. Skipping backup.")
        return

    current = Path(PARTNER_DETAILS_FILE)
    previous = Path(PARTNER_DETAILS_PREVIOUS_FILE)

    if not current.exists():
        log(
            "No previous partner_details.csv found. Skipping backup.",
            "WARNING",
        )
        return

    shutil.copy2(current, previous)
    log(f"Previous partner_details.csv backed up to {previous}")


def load_inventory() -> pd.DataFrame:
    """Load the partner inventory and validate its required columns."""
    inventory_path = Path(PARTNER_INVENTORY_FILE)
    if not inventory_path.exists():
        raise FileNotFoundError(
            f"Partner inventory not found: {inventory_path}"
        )

    inventory = pd.read_csv(inventory_path)

    missing = [
        column
        for column in REQUIRED_INVENTORY_COLUMNS
        if column not in inventory.columns
    ]
    if missing:
        raise ValueError(
            "Inventory file is missing required columns: "
            f"{', '.join(missing)}"
        )

    log(f"Inventory loaded: {len(inventory)} partners from {inventory_path}")
    return inventory


# =========================================================
# HTTP Session & Requests
# =========================================================

def create_session() -> requests.Session:
    """Build a reusable session with the standard request headers."""
    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)
    return session


def build_payload(row: pd.Series) -> Dict[str, str]:
    """Build the getPartnerOverview request body for one partner."""
    return {
        "BE_GEO_ID": _clean(row.get("be_geo_id")),
        "LOCATION_IDS": _clean(row.get("location_id")),
        "LANGUAGE_CD": LANGUAGE,
    }


def _parse_overview_response(data: Any) -> Dict[str, Any]:
    """Validate an API payload and return the first partner record.

    Raises:
        PartnerDetailError: If the payload is malformed, reports failure,
            or contains no partner records.
    """
    if not isinstance(data, dict):
        raise PartnerDetailError(
            "Unexpected API response: expected a JSON object."
        )
    if not data.get("success"):
        raise PartnerDetailError("API response did not report success.")

    records = data.get("data")
    if not isinstance(records, list) or not records:
        raise PartnerDetailError(
            "API response did not contain any partner records."
        )

    partner = records[0]
    if not isinstance(partner, dict):
        raise PartnerDetailError("API partner record is malformed.")

    return partner


def fetch_partner_details(
    session: requests.Session,
    payload: Dict[str, str],
    partner_name: str,
) -> Dict[str, Any]:
    """Fetch and validate one partner's overview, with retries.

    Retries up to ``MAX_RETRIES`` times, waiting ``RETRY_DELAY`` seconds
    between attempts. Every attempt is logged. After the final attempt a
    :class:`PartnerDetailError` is raised carrying the last error.
    """
    last_error: Exception | None = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.post(
                PARTNER_OVERVIEW_API,
                json=payload,
                timeout=REQUEST_TIMEOUT,
            )
            response.raise_for_status()
            return _parse_overview_response(response.json())
        except Exception as exc:  # noqa: BLE001 - retried, then re-raised
            last_error = exc
            log(
                f"{partner_name} - attempt {attempt}/{MAX_RETRIES} "
                f"failed: {exc}",
                "WARNING",
            )
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY)

    raise PartnerDetailError(
        f"Failed to fetch details for '{partner_name}' after "
        f"{MAX_RETRIES} attempts: {last_error}"
    ) from last_error


# =========================================================
# Data Extraction
# =========================================================

def extract_partner_data(
    partner: Dict[str, Any],
    inventory_row: pd.Series,
) -> Dict[str, Any]:
    """Map a raw partner overview record onto the output schema.

    Identity fields (partner_name, site_id, be_geo_id) fall back to the
    inventory values whenever the API omits them, so known identifiers
    are never left blank.
    """
    city = _clean(partner.get("site_city"))
    country = _clean(partner.get("site_country"))
    portfolios, competencies = _extract_qualifications(
        partner.get("qualifications", [])
    )

    return {
        "partner_name": (
            _clean(partner.get("partner_name"))
            or _clean(inventory_row.get("partner_name"))
        ),
        "site_id": (
            _clean(partner.get("site_id"))
            or _clean(inventory_row.get("site_id"))
        ),
        "be_geo_id": (
            _clean(partner.get("be_geo_id"))
            or _clean(inventory_row.get("be_geo_id"))
        ),
        "partner_since": _clean(partner.get("partner_since")),
        "website": _normalize_website(partner.get("web_addr")),
        "phone": _clean(partner.get("site_phone")),
        "address_line_1": _clean(partner.get("site_addr_1")),
        "address_line_2": _clean(partner.get("site_addr_2")),
        "address_line_3": _clean(partner.get("site_addr_3")),
        "city": city,
        "state": _clean(partner.get("site_state")),
        "country": country,
        "zip_code": _clean(partner.get("site_zip")),
        "coordinates": _clean(partner.get("loc_coordinates")),
        "description": _clean(partner.get("note_text")),
        "location": _build_location(city, country),
        "portfolios": portfolios,
        "competencies": competencies,
    }


def _build_failure_record(
    inventory_row: pd.Series,
    partner_name: str,
    error: Exception,
) -> Dict[str, Any]:
    """Build one row of the failed-partners audit log."""
    return {
        "partner_name": partner_name,
        "site_id": _clean(inventory_row.get("site_id")),
        "be_geo_id": _clean(inventory_row.get("be_geo_id")),
        "location_id": _clean(inventory_row.get("location_id")),
        "attempts": MAX_RETRIES,
        "error_message": str(error),
        "timestamp": datetime.now().strftime(TIMESTAMP_FORMAT),
    }


# =========================================================
# Persistence
# =========================================================

def save_checkpoint(
    results: List[Dict[str, Any]],
    failed: List[Dict[str, Any]],
) -> None:
    """Persist collected and failed partners to disk.

    Both frames are written with an explicit column schema so the files
    are stable even when the in-memory records are empty or partial.
    """
    pd.DataFrame(results, columns=OUTPUT_COLUMNS).to_csv(
        PARTNER_DETAILS_FILE,
        index=CSV_INDEX,
        encoding=FILE_ENCODING,
    )
    pd.DataFrame(failed, columns=FAILED_COLUMNS).to_csv(
        FAILED_PARTNERS_FILE,
        index=CSV_INDEX,
        encoding=FILE_ENCODING,
    )


# =========================================================
# Reporting
# =========================================================

def print_summary(
    total: int,
    results: List[Dict[str, Any]],
    failed: List[Dict[str, Any]],
    duration: float,
) -> None:
    """Log the final run summary block."""
    successful = len(results)
    failure_count = len(failed)
    failure_rate = (failure_count / total) if total else 0.0
    minutes = duration / SECONDS_PER_MINUTE
    partners_per_minute = (total / minutes) if minutes > 0 else 0.0

    log(SUMMARY_BORDER)
    log("Partner Detail Collection Complete")
    log(f"Total Partners      : {total}")
    log(f"Successful          : {successful}")
    log(f"Failed              : {failure_count}")
    log(f"Failure Rate        : {failure_rate:.1%}")
    log(f"Duration            : {_format_duration(duration)}")
    log(f"Partners Per Minute : {partners_per_minute:.1f}")
    log(f"Output File         : {PARTNER_DETAILS_FILE}")
    log(f"Failed File         : {FAILED_PARTNERS_FILE}")
    log(SUMMARY_BORDER)


def _enforce_failure_threshold(
    total: int,
    failed: List[Dict[str, Any]],
) -> None:
    """Raise if the failure rate exceeds the configured threshold."""
    if total == 0:
        return
    failure_rate = len(failed) / total
    if failure_rate > FAILURE_THRESHOLD:
        raise RuntimeError(
            f"{len(failed)} partners ({failure_rate:.1%}) could not be "
            f"processed, exceeding the {FAILURE_THRESHOLD:.0%} threshold."
        )


# =========================================================
# Orchestration
# =========================================================

def main() -> None:
    """Collect details for every inventory partner and persist results."""
    start_time = time.perf_counter()
    start_timestamp = datetime.now().strftime(TIMESTAMP_FORMAT)

    log(SUMMARY_BORDER)
    log(PROJECT_NAME)
    log(f"Version              : {VERSION}")
    log(f"Start Time           : {start_timestamp}")
    log(f"Inventory File       : {PARTNER_INVENTORY_FILE}")
    log(f"Partner Details File : {PARTNER_DETAILS_FILE}")
    log(SUMMARY_BORDER)

    ensure_directories()
    backup_previous_file()
    inventory = load_inventory()

    total = len(inventory)
    if total == 0:
        log("Inventory is empty. Nothing to process.", "WARNING")
        save_checkpoint([], [])
        return

    log(f"Total partners to process: {total}")

    results: List[Dict[str, Any]] = []
    failed: List[Dict[str, Any]] = []

    with create_session() as session:
        for position, (_, row) in enumerate(inventory.iterrows(), start=1):
            partner_name = _clean(row.get("partner_name")) or "Unknown Partner"
            log(f"[{position}/{total}] {partner_name}")

            payload = build_payload(row)

            try:
                partner = fetch_partner_details(session, payload, partner_name)
                results.append(extract_partner_data(partner, row))
            except Exception as exc:  # noqa: BLE001 - recorded, run continues
                failed.append(_build_failure_record(row, partner_name, exc))
                log(
                    f"{partner_name} - permanently failed after "
                    f"{MAX_RETRIES} attempts.",
                    "ERROR",
                )

            if position % CHECKPOINT_INTERVAL == 0:
                save_checkpoint(results, failed)
                log(
                    f"Checkpoint saved ({position}/{total}) - "
                    f"{len(results)} ok, {len(failed)} failed"
                )

            time.sleep(REQUEST_DELAY)

    save_checkpoint(results, failed)

    duration = time.perf_counter() - start_time
    print_summary(total, results, failed, duration)

    _enforce_failure_threshold(total, failed)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001 - logged before fatal re-raise
        log("Fatal error during partner detail collection.", "ERROR")
        log(str(exc), "ERROR")
        raise
