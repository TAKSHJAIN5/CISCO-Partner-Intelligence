"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
detect_changes.py

Purpose:
Compares the latest partner details snapshot with the
previous snapshot and generates a complete change log
covering new partners, removed partners, and changes to
tracked fields.

Runs unattended on a bi-weekly schedule via GitHub Actions.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd

from config import (
    CHANGE_LOG_COLUMNS,
    CHANGE_LOG_FILE,
    CSV_INDEX,
    ENABLE_CHANGE_LOG,
    FILE_ENCODING,
    PARTNER_DETAILS_FILE,
    PARTNER_DETAILS_PREVIOUS_FILE,
    PROJECT_NAME,
    SECONDS_PER_HOUR,
    SECONDS_PER_MINUTE,
    SUMMARY_BORDER,
    TIMESTAMP_FORMAT,
    TRACKED_CHANGE_FIELDS,
    VERSION,
)
from logger import log

# Column used as the unique partner key across both snapshots.
PARTNER_NAME_COLUMN = "partner_name"

# Change log column that stores the type of each change.
CHANGE_TYPE_COLUMN = "change_type"

# Change types that are not tied to a tracked field.
CHANGE_TYPE_NEW_PARTNER = "NEW_PARTNER"
CHANGE_TYPE_REMOVED_PARTNER = "REMOVED_PARTNER"

# Value labels recorded for partner additions and removals.
NEW_PARTNER_NEW_VALUE = "Added"
REMOVED_PARTNER_OLD_VALUE = "Present"
REMOVED_PARTNER_NEW_VALUE = "Removed"

# Columns that must exist in every snapshot before comparison.
REQUIRED_SNAPSHOT_COLUMNS = [PARTNER_NAME_COLUMN, *TRACKED_CHANGE_FIELDS]


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


def _build_change_record(
    timestamp: str,
    partner_name: str,
    change_type: str,
    old_value: Any,
    new_value: Any,
) -> Dict[str, Any]:
    """Build a single change log record in the canonical schema."""
    return {
        "timestamp": timestamp,
        "partner_name": partner_name,
        "change_type": change_type,
        "old_value": old_value,
        "new_value": new_value,
    }


# =========================================================
# Filesystem & Validation
# =========================================================

def _validate_snapshot_exists(file_path: str, label: str) -> None:
    """Raise if a required snapshot file is missing."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(
            f"The {label} snapshot was not found: {path}"
        )


def _validate_columns(
    frame: pd.DataFrame,
    required_columns: List[str],
    label: str,
) -> None:
    """Raise if a snapshot is missing any required column."""
    missing = [
        column
        for column in required_columns
        if column not in frame.columns
    ]
    if missing:
        raise ValueError(
            f"The {label} snapshot is missing required columns: "
            f"{', '.join(missing)}"
        )


# =========================================================
# Loading
# =========================================================

def load_snapshot(file_path: str, label: str) -> pd.DataFrame:
    """Load a partner snapshot, treating missing cells as empty text."""
    frame = pd.read_csv(file_path).fillna("")
    log(f"Loaded {label} snapshot: {len(frame)} partners from {file_path}")
    return frame


def build_lookup(frame: pd.DataFrame) -> Dict[str, Dict[str, Any]]:
    """Index snapshot rows by partner name for O(1) comparison.

    When a partner name appears more than once, the last occurrence
    wins, matching the previous behavior.
    """
    lookup: Dict[str, Dict[str, Any]] = {}
    for record in frame.to_dict("records"):
        partner_name = record.get(PARTNER_NAME_COLUMN, "")
        lookup[partner_name] = record
    return lookup


# =========================================================
# Comparison
# =========================================================

def detect_new_partners(
    current_lookup: Dict[str, Dict[str, Any]],
    previous_lookup: Dict[str, Dict[str, Any]],
    timestamp: str,
) -> List[Dict[str, Any]]:
    """Find partners present in the current snapshot but not the previous."""
    return [
        _build_change_record(
            timestamp,
            partner_name,
            CHANGE_TYPE_NEW_PARTNER,
            "",
            NEW_PARTNER_NEW_VALUE,
        )
        for partner_name in current_lookup
        if partner_name not in previous_lookup
    ]


def detect_removed_partners(
    current_lookup: Dict[str, Dict[str, Any]],
    previous_lookup: Dict[str, Dict[str, Any]],
    timestamp: str,
) -> List[Dict[str, Any]]:
    """Find partners present in the previous snapshot but not the current."""
    return [
        _build_change_record(
            timestamp,
            partner_name,
            CHANGE_TYPE_REMOVED_PARTNER,
            REMOVED_PARTNER_OLD_VALUE,
            REMOVED_PARTNER_NEW_VALUE,
        )
        for partner_name in previous_lookup
        if partner_name not in current_lookup
    ]


def detect_field_changes(
    current_lookup: Dict[str, Dict[str, Any]],
    previous_lookup: Dict[str, Dict[str, Any]],
    timestamp: str,
) -> List[Dict[str, Any]]:
    """Find changed tracked fields for partners in both snapshots.

    Tracked fields and their change types are driven by
    ``TRACKED_CHANGE_FIELDS`` so new fields can be tracked by editing
    configuration alone.
    """
    changes: List[Dict[str, Any]] = []
    common_partners = current_lookup.keys() & previous_lookup.keys()

    for partner_name in common_partners:
        current = current_lookup[partner_name]
        previous = previous_lookup[partner_name]

        for field, change_type in TRACKED_CHANGE_FIELDS.items():
            old_value = previous.get(field, "")
            new_value = current.get(field, "")
            if old_value != new_value:
                changes.append(
                    _build_change_record(
                        timestamp,
                        partner_name,
                        change_type,
                        old_value,
                        new_value,
                    )
                )

    return changes


def detect_changes(
    current_lookup: Dict[str, Dict[str, Any]],
    previous_lookup: Dict[str, Dict[str, Any]],
    timestamp: str,
) -> List[Dict[str, Any]]:
    """Collect every change between the previous and current snapshots."""
    changes: List[Dict[str, Any]] = []
    changes.extend(
        detect_new_partners(current_lookup, previous_lookup, timestamp)
    )
    changes.extend(
        detect_removed_partners(current_lookup, previous_lookup, timestamp)
    )
    changes.extend(
        detect_field_changes(current_lookup, previous_lookup, timestamp)
    )
    return changes


# =========================================================
# Persistence
# =========================================================

def save_change_log(changes: List[Dict[str, Any]]) -> pd.DataFrame:
    """Persist the change log with explicit columns and stable ordering."""
    changes_df = pd.DataFrame(changes, columns=CHANGE_LOG_COLUMNS)

    if not changes_df.empty:
        changes_df = changes_df.sort_values(
            by=[PARTNER_NAME_COLUMN, CHANGE_TYPE_COLUMN],
            key=lambda column: column.astype(str).str.casefold(),
            kind="stable",
        ).reset_index(drop=True)

    output_path = Path(CHANGE_LOG_FILE)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    changes_df.to_csv(
        output_path,
        index=CSV_INDEX,
        encoding=FILE_ENCODING,
    )
    return changes_df


# =========================================================
# Reporting
# =========================================================

def print_summary(
    partners_compared: int,
    changes_df: pd.DataFrame,
    duration: float,
) -> None:
    """Log the final change-detection summary block."""
    if not changes_df.empty:
        counts = changes_df[CHANGE_TYPE_COLUMN].value_counts().to_dict()
        partners_changed = changes_df[PARTNER_NAME_COLUMN].nunique()
    else:
        counts = {}
        partners_changed = 0

    new_count = counts.get(CHANGE_TYPE_NEW_PARTNER, 0)
    removed_count = counts.get(CHANGE_TYPE_REMOVED_PARTNER, 0)
    portfolio_count = counts.get(TRACKED_CHANGE_FIELDS["portfolios"], 0)
    competency_count = counts.get(TRACKED_CHANGE_FIELDS["competencies"], 0)
    location_count = counts.get(TRACKED_CHANGE_FIELDS["location"], 0)

    log(SUMMARY_BORDER)
    log("Change Detection Complete")
    log(f"Partners Compared    : {partners_compared}")
    log(f"New Partners         : {new_count}")
    log(f"Removed Partners     : {removed_count}")
    log(f"Portfolio Changes    : {portfolio_count}")
    log(f"Competency Changes   : {competency_count}")
    log(f"Location Changes     : {location_count}")
    log(f"Partners Changed     : {partners_changed}")
    log(f"Total Change Records : {len(changes_df)}")
    log(f"Duration             : {_format_duration(duration)}")
    log(f"Output File          : {CHANGE_LOG_FILE}")
    log(SUMMARY_BORDER)


# =========================================================
# Orchestration
# =========================================================

def main() -> None:
    """Compare the current and previous snapshots and log the changes."""
    start_time = time.perf_counter()
    start_timestamp = datetime.now().strftime(TIMESTAMP_FORMAT)

    log(SUMMARY_BORDER)
    log(PROJECT_NAME)
    log(f"Version           : {VERSION}")
    log(f"Start Time        : {start_timestamp}")
    log(f"Current Snapshot  : {PARTNER_DETAILS_FILE}")
    log(f"Previous Snapshot : {PARTNER_DETAILS_PREVIOUS_FILE}")
    log(f"Output File       : {CHANGE_LOG_FILE}")
    log(SUMMARY_BORDER)

    if not ENABLE_CHANGE_LOG:
        log("Change detection is disabled in configuration. Skipping.", "WARNING")
        return

    if not Path(PARTNER_DETAILS_PREVIOUS_FILE).exists():
        log(
            "No previous snapshot found. Change detection skipped.",
            "WARNING",
        )
        return

    _validate_snapshot_exists(PARTNER_DETAILS_FILE, "current")

    current_df = load_snapshot(PARTNER_DETAILS_FILE, "current")
    previous_df = load_snapshot(PARTNER_DETAILS_PREVIOUS_FILE, "previous")

    _validate_columns(current_df, REQUIRED_SNAPSHOT_COLUMNS, "current")
    _validate_columns(previous_df, REQUIRED_SNAPSHOT_COLUMNS, "previous")

    current_lookup = build_lookup(current_df)
    previous_lookup = build_lookup(previous_df)

    timestamp = datetime.now().strftime(TIMESTAMP_FORMAT)
    changes = detect_changes(current_lookup, previous_lookup, timestamp)

    changes_df = save_change_log(changes)

    partners_compared = len(set(current_lookup) | set(previous_lookup))
    duration = time.perf_counter() - start_time
    print_summary(partners_compared, changes_df, duration)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001 - logged before fatal re-raise
        log("Fatal error during change detection.", "ERROR")
        log(str(exc), "ERROR")
        raise
