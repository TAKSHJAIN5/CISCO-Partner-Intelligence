"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
collect_partner_inventory.py

Purpose:
Collects Cisco partners matching the configured portfolio
and competency filters and stores the inventory for
subsequent processing.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

import math
import time

import pandas as pd
import requests

from config import (
    PARTNER_INVENTORY_FILE,
    PARTNER_SEARCH_API,
    PORTFOLIO_IDS,
    COMPETENCY_IDS,
    PAGE_SIZE,
    LANGUAGE,
    RANDOM_SEED,
    REQUEST_TIMEOUT,
    MAX_RETRIES,
    RETRY_DELAY,
    REQUEST_DELAY
)

from logger import log


payload = {
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
    "pageCount": 1,
    "SORT_BY": "partner_name.sort",
    "SORT_ORDER": "random",
    "LANGUAGE_CD": LANGUAGE,
    "RANDOM_SEED": RANDOM_SEED
}


try:

    log("Starting partner inventory collection.")

    response = requests.post(
        PARTNER_SEARCH_API,
        json=payload,
        timeout=REQUEST_TIMEOUT
    )

    response.raise_for_status()

    data = response.json()

    matches = data["matches"]

    total_pages = math.ceil(matches / PAGE_SIZE)

    log(f"Total Partners Found : {matches}")
    log(f"Total Pages : {total_pages}")

    all_partners = []

    for page in range(1, total_pages + 1):

        log(f"Processing Page {page}/{total_pages}")

        payload["pageCount"] = page

        success = False

        for attempt in range(1, MAX_RETRIES + 1):

            try:

                response = requests.post(
                    PARTNER_SEARCH_API,
                    json=payload,
                    timeout=REQUEST_TIMEOUT
                )

                response.raise_for_status()

                page_data = response.json()

                for partner in page_data["data"]:

                    all_partners.append({
                        "partner_name": partner.get("partner_name"),
                        "be_geo_id": partner.get("be_geo_id"),
                        "location_id": partner.get("location_keyword_ids")
                    })

                success = True
                break

            except Exception as e:

                log(
                    f"Page {page} failed "
                    f"(Attempt {attempt}/{MAX_RETRIES})",
                    "WARNING"
                )

                log(str(e), "WARNING")

                time.sleep(RETRY_DELAY)

        if not success:

    message = (
        f"Unable to retrieve Page {page} "
        f"after {MAX_RETRIES} attempts."
    )

    log(
        message,
        "ERROR"
    )

    raise RuntimeError(message)

        time.sleep(REQUEST_DELAY)

    df = pd.DataFrame(all_partners)
    
from pathlib import Path

Path(
    PARTNER_INVENTORY_FILE
).parent.mkdir(
    parents=True,
    exist_ok=True
)
    
    df.to_csv(
        PARTNER_INVENTORY_FILE,
        index=False
    )

    log("=" * 60)
    log("Partner inventory collection completed successfully.")
    log(f"Rows Saved : {len(df)}")
    log(f"Output File : {PARTNER_INVENTORY_FILE}")
    log("=" * 60)

except Exception as e:

    log(
        "Fatal error during partner inventory collection.",
        "ERROR"
    )

    log(str(e), "ERROR")

    raise
