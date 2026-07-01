"""
===========================================================
Cisco Partner Intelligence Automation Platform

Script:
build_final_excel.py

Purpose:
Builds the final Excel dataset by transforming the raw
partner details into the business-required format.

Author:
Taksh Jain

Version:
1.0.0
===========================================================
"""

import pandas as pd

from config import (
    PARTNER_DETAILS_FILE,
    OUTPUT_EXCEL_FILE
)

from logger import log


try:

    log("Building final Excel dataset...")

    df = pd.read_csv(
        PARTNER_DETAILS_FILE
    )

    final_rows = []

    for _, row in df.iterrows():

        portfolios_text = (
            ""
            if pd.isna(row["portfolios"])
            else str(row["portfolios"])
        )

        portfolio_list = [

            portfolio.strip()

            for portfolio in portfolios_text.split("|")

            if portfolio.strip()

        ]

        preferred_portfolios = []

        non_preferred_portfolios = []

        for portfolio in portfolio_list:

            if "Preferred" in portfolio:

                preferred_portfolios.append(
                    portfolio
                )

            else:

                non_preferred_portfolios.append(
                    portfolio
                )

        preferred_nw = (
            "Preferred Networking Partner"
            in preferred_portfolios
        )

        preferred_sec = (
            "Preferred Security Partner"
            in preferred_portfolios
        )

        preferred_cloud_ai = (
            "Preferred Cloud and AI Infrastructure Partner"
            in preferred_portfolios
        )

        sec_net_spez = (
            preferred_nw
            and preferred_sec
        )

        sec_ai_infra_spez = (
            preferred_sec
            and preferred_cloud_ai
        )

        final_rows.append({

            "Name":
            row["partner_name"],

            "Preferred Portfolios":
            ", ".join(
                preferred_portfolios
            ),

            "Non Preferred Portfolios":
            ", ".join(
                non_preferred_portfolios
            ),

            "Sec Net Spez":
            "Yes" if sec_net_spez else "No",

            "Sec AI Infra Spez":
            "Yes" if sec_ai_infra_spez else "No",

            "Preferred NW":
            "Yes" if preferred_nw else "No",

            "Preferred SEC":
            "Yes" if preferred_sec else "No",

            "Preferred Cloud + AI":
            "Yes" if preferred_cloud_ai else "No",

            "Competencies":
            row["competencies"],

            "Location":
            row["location"]

        })

    final_df = pd.DataFrame(
        final_rows
    )

    final_df.to_excel(

        OUTPUT_EXCEL_FILE,

        index=False

    )

    log("=" * 60)

    log("Excel dataset created successfully.")

    log(f"Rows Exported : {len(final_df)}")

    log(f"Output File : {OUTPUT_EXCEL_FILE}")

    log("=" * 60)

except Exception as e:

    log(
        "Fatal error while building Excel dataset.",
        "ERROR"
    )

    log(
        str(e),
        "ERROR"
    )

    raise