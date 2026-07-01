"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
detect_changes.py

Purpose:
Compares the latest partner details with the previous
snapshot and generates a complete change log.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

from datetime import datetime
from pathlib import Path

import pandas as pd

from config import (
    PARTNER_DETAILS_FILE,
    PARTNER_DETAILS_PREVIOUS_FILE,
    CHANGE_LOG_FILE,
    ENABLE_CHANGE_LOG
)

from logger import log


try:

    if not ENABLE_CHANGE_LOG:

        log(
            "Change detection disabled in configuration.",
            "WARNING"
        )

        raise SystemExit

    # =====================================================
    # Verify Previous Snapshot Exists
    # =====================================================

    if not Path(PARTNER_DETAILS_PREVIOUS_FILE).exists():

        log(
            "No previous snapshot found. "
            "Change detection skipped.",
            "WARNING"
        )

        raise SystemExit

    # =====================================================
    # Load Data
    # =====================================================

    current_df = pd.read_csv(
        PARTNER_DETAILS_FILE
    ).fillna("")

    previous_df = pd.read_csv(
        PARTNER_DETAILS_PREVIOUS_FILE
    ).fillna("")

    current_lookup = {
        row["partner_name"]: row
        for _, row in current_df.iterrows()
    }

    previous_lookup = {
        row["partner_name"]: row
        for _, row in previous_df.iterrows()
    }

    changes = []

    timestamp = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    # =====================================================
    # New Partners
    # =====================================================

    for partner in current_lookup:

        if partner not in previous_lookup:

            changes.append({

                "timestamp": timestamp,

                "partner_name": partner,

                "change_type": "NEW_PARTNER",

                "old_value": "",

                "new_value": "Added"

            })

    # =====================================================
    # Removed Partners
    # =====================================================

    for partner in previous_lookup:

        if partner not in current_lookup:

            changes.append({

                "timestamp": timestamp,

                "partner_name": partner,

                "change_type": "REMOVED_PARTNER",

                "old_value": "Present",

                "new_value": "Removed"

            })

    # =====================================================
    # Existing Partner Changes
    # =====================================================

    common_partners = set(
        current_lookup.keys()
    ).intersection(
        previous_lookup.keys()
    )

    for partner in common_partners:

        current = current_lookup[partner]

        previous = previous_lookup[partner]

        if current["portfolios"] != previous["portfolios"]:

            changes.append({

                "timestamp": timestamp,

                "partner_name": partner,

                "change_type": "PORTFOLIO_CHANGED",

                "old_value": previous["portfolios"],

                "new_value": current["portfolios"]

            })

        if current["competencies"] != previous["competencies"]:

            changes.append({

                "timestamp": timestamp,

                "partner_name": partner,

                "change_type": "COMPETENCY_CHANGED",

                "old_value": previous["competencies"],

                "new_value": current["competencies"]

            })

        if current["location"] != previous["location"]:

            changes.append({

                "timestamp": timestamp,

                "partner_name": partner,

                "change_type": "LOCATION_CHANGED",

                "old_value": previous["location"],

                "new_value": current["location"]

            })

    # =====================================================
    # Save Change Log
    # =====================================================

    changes_df = pd.DataFrame(
        changes,
        columns=[
            "timestamp",
            "partner_name",
            "change_type",
            "old_value",
            "new_value"
        ]
    )

    changes_df.to_csv(
        CHANGE_LOG_FILE,
        index=False
    )

    # =====================================================
    # Summary
    # =====================================================

    new_count = len(
        changes_df[
            changes_df["change_type"] == "NEW_PARTNER"
        ]
    )

    removed_count = len(
        changes_df[
            changes_df["change_type"] == "REMOVED_PARTNER"
        ]
    )

    portfolio_count = len(
        changes_df[
            changes_df["change_type"] == "PORTFOLIO_CHANGED"
        ]
    )

    competency_count = len(
        changes_df[
            changes_df["change_type"] == "COMPETENCY_CHANGED"
        ]
    )

    location_count = len(
        changes_df[
            changes_df["change_type"] == "LOCATION_CHANGED"
        ]
    )

    log("=" * 60)

    log("Change detection completed successfully.")

    log(f"New Partners : {new_count}")

    log(f"Removed Partners : {removed_count}")

    log(f"Portfolio Changes : {portfolio_count}")

    log(f"Competency Changes : {competency_count}")

    log(f"Location Changes : {location_count}")

    log(f"Total Changes : {len(changes_df)}")

    log(f"Output File : {CHANGE_LOG_FILE}")

    log("=" * 60)

except Exception as e:

    log(
        "Fatal error during change detection.",
        "ERROR"
    )

    log(str(e), "ERROR")

    raise