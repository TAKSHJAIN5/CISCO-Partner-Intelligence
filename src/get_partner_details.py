"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
get_partner_details.py

Purpose:
Retrieves detailed information for every Cisco partner
collected in the inventory and stores a clean dataset.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

import shutil
import time
from pathlib import Path

import pandas as pd
import requests

from config import (
    PARTNER_INVENTORY_FILE,
    PARTNER_DETAILS_FILE,
    PARTNER_DETAILS_PREVIOUS_FILE,
    FAILED_PARTNERS_FILE,
    PARTNER_OVERVIEW_API,
    LANGUAGE,
    REQUEST_TIMEOUT,
    MAX_RETRIES,
    RETRY_DELAY,
    REQUEST_DELAY,
    SAVE_BACKUP
)

from logger import log


try:

    # ==================================================
    # Backup Previous File
    # ==================================================

    if SAVE_BACKUP:

        current_file = Path(PARTNER_DETAILS_FILE)
        backup_file = Path(PARTNER_DETAILS_PREVIOUS_FILE)

        if current_file.exists():

            shutil.copy(
                current_file,
                backup_file
            )

            log(
                "Previous partner_details.csv backed up successfully."
            )

        else:

            log(
                "No previous partner_details.csv found. Skipping backup.",
                "WARNING"
            )

    inventory = pd.read_csv(
        PARTNER_INVENTORY_FILE
    )

    results = []

    failed = []

    total = len(inventory)

    log(f"Total Partners To Process : {total}")

    for index, row in inventory.iterrows():

        partner_name = row["partner_name"]

        log(
            f"[{index + 1}/{total}] {partner_name}"
        )

        payload = {
            "BE_GEO_ID": str(
                row["be_geo_id"]
            ),
            "LOCATION_IDS": str(
                row["location_id"]
            ),
            "LANGUAGE_CD": LANGUAGE
        }

        success = False
        for attempt in range(1, MAX_RETRIES + 1):

            try:

                response = requests.post(
                    PARTNER_OVERVIEW_API,
                    json=payload,
                    timeout=REQUEST_TIMEOUT
                )

                response.raise_for_status()

                data = response.json()

                partner = data["data"][0]

                city = partner.get("site_city", "")
                country = partner.get("site_country", "")

                portfolios = []
                competencies = []

                for item in partner.get("qualifications", []):

                    category = item.get(
                        "category_display_name",
                        ""
                    )

                    name = item.get(
                        "keyword_display_name",
                        ""
                    )

                    if category == "PORTFOLIO":
                        portfolios.append(name)

                    elif category == "COMPETENCY":
                        competencies.append(name)

                results.append({

                    "partner_name": partner.get(
                        "partner_name",
                        ""
                    ),

                    "city": city,

                    "country": country,

                    "location": f"{city}, {country}",

                    "portfolios": " | ".join(portfolios),

                    "competencies": " | ".join(competencies)

                })

                success = True
                break

            except Exception as e:

                log(
                    f"{partner_name} - Attempt "
                    f"{attempt}/{MAX_RETRIES} failed",
                    "WARNING"
                )

                log(str(e), "WARNING")

                time.sleep(RETRY_DELAY)

        if not success:

            failed.append({

                "partner_name": partner_name,

                "be_geo_id": row["be_geo_id"],

                "location_id": row["location_id"]

            })

        # -----------------------------------------
        # Checkpoint Save
        # -----------------------------------------

        if (index + 1) % 50 == 0:

            pd.DataFrame(results).to_csv(
                PARTNER_DETAILS_FILE,
                index=False
            )

            pd.DataFrame(failed).to_csv(
                FAILED_PARTNERS_FILE,
                index=False
            )

            log(
                f"Checkpoint Saved "
                f"({index + 1}/{total})"
            )

        time.sleep(REQUEST_DELAY)

    # ==================================================
    # Final Save
    # ==================================================

    pd.DataFrame(results).to_csv(
        PARTNER_DETAILS_FILE,
        index=False
    )

    pd.DataFrame(failed).to_csv(
        FAILED_PARTNERS_FILE,
        index=False
    )

    log("=" * 60)

    log("Partner detail collection completed successfully.")

    log(f"Successful Partners : {len(results)}")

    log(f"Failed Partners : {len(failed)}")

    log(f"Output File : {PARTNER_DETAILS_FILE}")

    log(f"Failed File : {FAILED_PARTNERS_FILE}")

    log("=" * 60)

except Exception as e:

    log(
        "Fatal error during partner detail collection.",
        "ERROR"
    )

    log(str(e), "ERROR")

    raise