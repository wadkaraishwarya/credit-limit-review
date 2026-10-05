"""
Credit Limit Review Automation

Main executable workflow.

Configuration and file paths: config.py
Reusable processing/reporting logic: functions.py

Run:
    python credit_limit_automation.py
"""

# ==================================================
# CREDIT LIMIT REVIEW WORKFLOW
# ==================================================

# Project configuration and reusable functions
from config import (
    URL, USERNAME, PASSWORD,
    CLIENT_FILE, OUTPUT_FOLDER, OUTPUT_FILE,
    BATCH_SIZE, MAX_WORKERS, BROWSER_START_DELAY,
    DECISION_WEIGHTS,
    REPAYMENT_STRONG, REPAYMENT_GOOD, REPAYMENT_LOW,
    MAX_WEIGHTED_LIMIT_CHANGE,
    INCREASE_SCORE_THRESHOLD, REDUCE_SCORE_THRESHOLD,
    LIMIT_ROUNDING,
)
from functions import (
    clean_text, parse_number, parse_date, month_index,
    calculate_balance_trend, calculate_estimated_limit,
    process_client,
)

import os
import re
import time
import math
import traceback
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ==================================================
# CREATE OUTPUT FOLDER
# ==================================================

os.makedirs(
    OUTPUT_FOLDER,
    exist_ok=True
)

print(
    "Output folder:",
    OUTPUT_FOLDER
)


# ==================================================
# READ CLIENT FILE
# ==================================================

print(
    "\nReading client file..."
)

source_workbook = openpyxl.load_workbook(
    CLIENT_FILE,
    data_only=True
)

source_sheet = (
    source_workbook.active
)

print(
    "Source sheet:",
    source_sheet.title
)


# ==================================================
# FIND COLUMNS
# ==================================================

business_name_column = None
for cell in source_sheet[1]:

    header = clean_text(
        cell.value
    ).lower()


    if header == "business name":

        business_name_column = (
            cell.column
        )
if business_name_column is None:

    raise ValueError(
        "Could not find 'Business Name' column."
    )


# ==================================================
# EXTRACT CLIENTS
# ==================================================

clients = []

seen_clients = set()


for row_number in range(
    2,
    source_sheet.max_row + 1
):

    business_name = clean_text(
        source_sheet.cell(
            row=row_number,
            column=business_name_column
        ).value
    )
    if not business_name:

        continue


    client_key = (
        business_name
        .lower()
        .strip()
    )


    if client_key in seen_clients:

        continue


    seen_clients.add(
        client_key
    )


    clients.append(
        {
            "Client Name":
                business_name,}
    )


print(
    "\n====================================="
)

print(
    "CLIENTS EXTRACTED:",
    len(clients)
)

print(
    "====================================="
)


for index, client in enumerate(
    clients,
    start=1
):

    print(
        index,
        client["Client Name"]
    )


# ==================================================
# PARALLEL BATCH PROCESSING
# ==================================================

all_results = []

total_clients = len(
    clients
)


for batch_start in range(
    0,
    total_clients,
    BATCH_SIZE
):

    batch_end = min(
        batch_start + BATCH_SIZE,
        total_clients
    )


    batch = clients[
        batch_start:batch_end
    ]


    print(
        "\n\n====================================="
    )

    print(
        "PROCESSING BATCH"
    )

    print(
        f"Clients "
        f"{batch_start + 1} "
        f"to {batch_end}"
    )

    print(
        "Clients in batch:",
        len(batch)
    )

    print(
        "Parallel browsers:",
        MAX_WORKERS
    )

    print(
        "Browser start delay:",
        BROWSER_START_DELAY,
        "seconds"
    )

    print(
        "====================================="
    )


    with ThreadPoolExecutor(
        max_workers=MAX_WORKERS
    ) as executor:


        future_to_client = {}


        # ==================================================
        # STAGGER INITIAL BROWSER STARTS
        # ==================================================

        for position, client in enumerate(
            batch
        ):

            future = executor.submit(
                process_client,
                client
            )


            future_to_client[
                future
            ] = client


            # Only delay while filling the initial
            # worker slots.
            #
            # Example with 5 workers:
            #
            # Client 1 = now
            # Client 2 = +3 sec
            # Client 3 = +6 sec
            # Client 4 = +9 sec
            # Client 5 = +12 sec
            #
            # After that, ThreadPoolExecutor manages
            # the remaining clients automatically.

            if (
                position < MAX_WORKERS - 1
                and position < len(batch) - 1
            ):

                print(
                    f"Waiting "
                    f"{BROWSER_START_DELAY} seconds "
                    f"before launching next browser..."
                )


                time.sleep(
                    BROWSER_START_DELAY
                )


        completed = 0


        # ==================================================
        # RECEIVE RESULTS
        # ==================================================

        for future in as_completed(
            future_to_client
        ):

            client = (
                future_to_client[
                    future
                ]
            )


            try:

                result = (
                    future.result()
                )


            except Exception as error:

                result = {
                    "Client Name":
                        client[
                            "Client Name"
                        ],

                    "Current Limit":
                        None,

                    "Balances":
                        [],

                    "Borrowings":
                        [],

                    "Repayments":
                        [],

                    "Disabled Status":
                        [],

                    "FICO":
                        [],

                    "Status":
                        "ERROR",

                    "Failed Step":
                        "Future",

                    "Error":
                        str(error)
                }


            all_results.append(
                result
            )


            completed += 1


            print(
                f"[{completed}/"
                f"{len(batch)}]",
                result["Status"],
                "|",
                result["Client Name"]
            )


            if (
                result["Status"]
                == "ERROR"
            ):

                print(
                    "Failed Step:",
                    result[
                        "Failed Step"
                    ]
                )

                print(
                    "Error:",
                    result[
                        "Error"
                    ]
                )


# ==================================================
# CREATE OUTPUT WORKBOOK
# ==================================================

workbook = Workbook()

workbook.remove(
    workbook.active
)


# ==================================================
# MONTHLY BALANCE SHEET
# ==================================================

balance_sheet = workbook.create_sheet(
    "Monthly Balance"
)

balance_headers = [
    "Client Name",
    "Current Limit",
    "Month",
    "Closing Balance"
]


for col, header in enumerate(
    balance_headers,
    start=1
):

    cell = balance_sheet.cell(
        row=1,
        column=col,
        value=header
    )

    cell.font = Font(
        bold=True
    )


# ==================================================
# BORROWING SHEET
# ==================================================

borrowing_sheet = workbook.create_sheet(
    "Borrowing"
)

borrowing_headers = [
    "Client Name",
    "Current Limit",
    "Date",
    "Description",
    "Reason",
    "Borrowing Amount",
    "Plexe LLC / Non-Plexe"
]


for col, header in enumerate(
    borrowing_headers,
    start=1
):

    cell = borrowing_sheet.cell(
        row=1,
        column=col,
        value=header
    )

    cell.font = Font(
        bold=True
    )


# ==================================================
# REPAYMENT SHEET
# ==================================================

repayment_sheet = workbook.create_sheet(
    "Repayment"
)

repayment_headers = [
    "Client Name",
    "Current Limit",
    "Repayment Type",
    "Repayment %"
]

for col, header in enumerate(
    repayment_headers,
    start=1
):

    cell = repayment_sheet.cell(
        row=1,
        column=col,
        value=header
    )

    cell.font = Font(bold=True)


# ==================================================
# DISABLED STATUS SHEET
# ==================================================

disabled_sheet = workbook.create_sheet(
    "Disabled Status"
)

disabled_headers = [
    "Client Name",
    "Current Limit",
    "Status Text",
    "Disabled?",
    "Disabled Reason"
]

for col, header in enumerate(
    disabled_headers,
    start=1
):

    cell = disabled_sheet.cell(
        row=1,
        column=col,
        value=header
    )

    cell.font = Font(bold=True)


# ==================================================
# FICO SHEET
# ==================================================

fico_sheet = workbook.create_sheet(
    "FICO"
)

fico_headers = [
    "Client Name",
    "Current Limit",
    "Latest FICO",
    "Latest FICO Date"
]

for col, header in enumerate(fico_headers, start=1):
    cell = fico_sheet.cell(row=1, column=col, value=header)
    cell.font = Font(bold=True)


# ==================================================
# LIMIT ADJUSTMENT HISTORY SHEET
# ==================================================

limit_adjustment_sheet = workbook.create_sheet(
    "Limit Adjustment History"
)

limit_adjustment_headers = [
    "Client Name",
    "Current Limit",
    "Latest Adjustment Date",
    "From Limit",
    "To Limit",
    "Latest Limit Change Text",
    "Reason",
    "Result",
    "Full Credit Log Text"
]

for col, header in enumerate(
    limit_adjustment_headers,
    start=1
):

    cell = limit_adjustment_sheet.cell(
        row=1,
        column=col,
        value=header
    )

    cell.font = Font(
        bold=True
    )


# ==================================================
# PROCESS STATUS SHEET
# ==================================================

status_sheet = workbook.create_sheet(
    "Process Status"
)

status_headers = [
    "Client Name",
    "Status",
    "Failed Step",
    "Error"
]


for col, header in enumerate(
    status_headers,
    start=1
):

    cell = status_sheet.cell(
        row=1,
        column=col,
        value=header
    )

    cell.font = Font(
        bold=True
    )


# ==================================================
# DATA STORAGE
# ==================================================

client_balance_data = {}

client_borrowing_data = {}

client_repayment_data = {}

client_disabled_data = {}

client_fico_data = {}


for client in clients:

    client_name = (
        client[
            "Client Name"
        ]
    )

    client_balance_data[
        client_name
    ] = []


    client_borrowing_data[
        client_name
    ] = []

    client_repayment_data[
        client_name
    ] = []

    client_disabled_data[
        client_name
    ] = []

    client_fico_data[
        client_name
    ] = []


# ==================================================
# CURRENT LIMIT LOOKUP - SELENIUM RESULTS
# ==================================================

# all_results has now been populated by process_client().
# Each Current Limit in all_results came from the Selenium XPath.
client_current_limit_data = {
    result.get("Client Name"): result.get("Current Limit")
    for result in all_results
    if result.get("Client Name")
}


# ==================================================
# WRITE RESULTS
# ==================================================

for result in all_results:

    client_name = (
        result[
            "Client Name"
        ]
    )

    current_limit = (
        result[
            "Current Limit"
        ]
    )


    # ==================================================
    # STATUS
    # ==================================================

    status_sheet.append(
        [
            client_name,
            result["Status"],
            result["Failed Step"],
            result["Error"]
        ]
    )


    # ==================================================
    # BALANCE
    # ==================================================

    for item in result[
        "Balances"
    ]:

        balance_sheet.append(
            [
                client_name,
                current_limit,
                item["Month"],
                item["Balance"]
            ]
        )


        client_balance_data[
            client_name
        ].append(
            {
                "Month":
                    item["Month"],

                "Date":
                    item["Date"],

                "Balance":
                    item["Balance"]
            }
        )


    # ==================================================
    # BORROWING
    # ==================================================

    for item in result[
        "Borrowings"
    ]:

        borrowing_sheet.append(
            [
                client_name,
                current_limit,
                item["Date Text"],
                item["Description"],
                item["Reason"],
                item["Amount"],
                item["Borrowing Type"]
            ]
        )


        client_borrowing_data[
            client_name
        ].append(
            {
                "Date Text":
                    item["Date Text"],

                "Date":
                    item["Date"],

                "Description":
                    item["Description"],

                "Reason":
                    item["Reason"],

                "Amount":
                    item["Amount"],

                "Borrowing Type":
                    item[
                        "Borrowing Type"
                    ]
            }
        )


    # ==================================================
    # DISABLED STATUS
    # ==================================================

    for item in result.get(
        "Disabled Status",
        []
    ):

        disabled_sheet.append(
            [
                client_name,
                current_limit,
                item.get("Status Text", ""),
                item.get("Disabled", "UNKNOWN"),
                item.get("Disabled Reason", "")
            ]
        )

        client_disabled_data[
            client_name
        ].append(item)


    # ==================================================
    # FICO
    # ==================================================

    for item in result.get(
        "FICO",
        []
    ):

        fico_sheet.append(
            [
                client_name,
                current_limit,
                item.get("Latest FICO"),
                item.get("Latest FICO Date", "")
            ]
        )

        client_fico_data[
            client_name
        ].append(item)


    # ==================================================
    # LIMIT ADJUSTMENT HISTORY
    # ==================================================

    for item in result.get(
        "Limit Adjustments",
        []
    ):

        limit_adjustment_sheet.append(
            [
                client_name,
                current_limit,
                item.get(
                    "Latest Adjustment Date",
                    "Not Found"
                ),
                item.get(
                    "From Limit"
                ),
                item.get(
                    "To Limit"
                ),
                item.get(
                    "Latest Limit Change Text",
                    "Not Found"
                ),
                item.get(
                    "Reason",
                    "Not Found"
                ),
                item.get(
                    "Result",
                    "NOT FOUND"
                ),
                item.get(
                    "Full Credit Log Text",
                    ""
                )
            ]
        )


    # ==================================================
    # REPAYMENT
    # ==================================================

    for item in result.get(
        "Repayments",
        []
    ):

        repayment_sheet.append(
            [
                client_name,
                current_limit,
                item["Repayment Type"],
                item["Repayment Percent"]
            ]
        )

        client_repayment_data[
            client_name
        ].append(item)


# ==================================================
# LIMIT REVIEW SHEET
# ==================================================

review_sheet = workbook.create_sheet(
    "Limit Review"
)


review_headers = [
    "Client Name",
    "Current Limit",
    "Disabled?",
    "Disabled Status Text",
    "Disabled Reason",
    "Latest FICO",
    "Latest FICO Date",
    "50% of Current Limit",
    "Latest Closing Balance",
    "Lowest Balance - Last 5 Months",
    "Balance Trend - Last 5 Months",
    "Balance Below Current Limit - Last 5 Months",
    "Largest Borrowing - Last 3 Months",
    "Borrowing 50%+ of Current Limit - Last 3 Months",
    "Non-Plexe Borrowing - Last 3 Months",
    "Forecasted Repayment %",
    "Historical Repayment %",
    "Repayment % Used",
    "Repayment Assessment",
    "Balance Level Score",
    "Balance Trend Score",
    "Large Borrowing Score",
    "Non-Plexe Borrowing Score",
    "Repayment Score",
    "Score Breakdown",
    "Weighted Decision Score",
    "Weighted Limit",
    "Estimated New Limit",
    "Recommended Action",
    "Review Reason"
]


for col, header in enumerate(
    review_headers,
    start=1
):

    cell = review_sheet.cell(
        row=1,
        column=col,
        value=header
    )

    cell.font = Font(
        bold=True
    )

# ==================================================
# REVIEW EACH CLIENT
# ==================================================

for client in clients:

    client_name = (
        client[
            "Client Name"
        ]
    )

    # Use the Current Limit previously extracted from Selenium.
    current_limit = client_current_limit_data.get(
        client_name
    )

    # ==================================================
    # DISABLED STATUS
    # ==================================================

    disabled_items = client_disabled_data.get(
        client_name,
        []
    )

    if disabled_items:
        disabled_flag = disabled_items[0].get("Disabled", "UNKNOWN")
        disabled_status_text = disabled_items[0].get("Status Text", "")
        disabled_reason = disabled_items[0].get("Disabled Reason", "")
    else:
        disabled_flag = "UNKNOWN"
        disabled_status_text = ""
        disabled_reason = ""


    # ==================================================
    # LATEST FICO
    # ==================================================

    fico_items = client_fico_data.get(
        client_name,
        []
    )

    if fico_items:
        latest_fico = fico_items[0].get("Latest FICO")
        latest_fico_date = fico_items[0].get("Latest FICO Date", "")
    else:
        latest_fico = None
        latest_fico_date = ""


    # ==================================================
    # 50% LIMIT
    # ==================================================

    if current_limit is not None:

        fifty_percent_limit = (
            current_limit
            * 0.50
        )

    else:

        fifty_percent_limit = None


    # ==================================================
    # BALANCE DATA
    # ==================================================

    balances = (
        client_balance_data.get(
            client_name,
            []
        )
    )


    dated_balances = [
        item
        for item in balances
        if item["Date"] is not None
    ]


    if dated_balances:

        dated_balances.sort(
            key=lambda x: x["Date"],
            reverse=True
        )

        latest_five_balances = (
            dated_balances[:5]
        )

    else:

        latest_five_balances = (
            balances[:5]
        )


    # ==================================================
    # BALANCE TREND
    # ==================================================

    balance_trend = (
        calculate_balance_trend(
            latest_five_balances
        )
    )


    # ==================================================
    # LATEST BALANCE
    # ==================================================

    if latest_five_balances:

        latest_closing_balance = (
            latest_five_balances[
                0
            ]["Balance"]
        )

    else:

        latest_closing_balance = None


    # ==================================================
    # LOWEST BALANCE
    # ==================================================

    valid_balances = [
        item["Balance"]
        for item in latest_five_balances
        if item["Balance"] is not None
    ]


    if valid_balances:

        lowest_last_five = min(
            valid_balances
        )

    else:

        lowest_last_five = None


    # ==================================================
    # BALANCE BELOW LIMIT
    # ==================================================

    balance_below_limit = False


    if current_limit is not None:

        for item in (
            latest_five_balances
        ):

            if (
                item["Balance"] is not None
                and item["Balance"]
                < current_limit
            ):

                balance_below_limit = True


    # ==================================================
    # BORROWING DATA
    # ==================================================

    borrowings = (
        client_borrowing_data.get(
            client_name,
            []
        )
    )


    dated_borrowings = [
        item
        for item in borrowings
        if item["Date"] is not None
    ]


    dated_borrowings.sort(
        key=lambda x: x["Date"],
        reverse=True
    )


    # ==================================================
    # REFERENCE DATE
    # ==================================================

    reference_date = None


    if dated_balances:

        reference_date = (
            dated_balances[
                0
            ]["Date"]
        )


    elif dated_borrowings:

        reference_date = (
            dated_borrowings[
                0
            ]["Date"]
        )


    # ==================================================
    # LAST 3 MONTHS BORROWING
    # ==================================================

    last_three_month_borrowings = []


    if (
        reference_date is not None
        and dated_borrowings
    ):

        reference_month = (
            month_index(
                reference_date
            )
        )


        for item in (
            dated_borrowings
        ):

            difference = (
                reference_month
                - month_index(
                    item["Date"]
                )
            )


            if (
                0 <= difference <= 2
            ):

                last_three_month_borrowings.append(
                    item
                )


    # ==================================================
    # LARGEST BORROWING
    # ==================================================

    borrowing_amounts = [
        item["Amount"]
        for item in last_three_month_borrowings
        if item["Amount"] is not None
    ]


    if borrowing_amounts:

        largest_borrowing = max(
            borrowing_amounts
        )

    else:

        largest_borrowing = None


    # ==================================================
    # BORROWING >= 50%
    # ==================================================

    borrowing_over_50 = False


    if fifty_percent_limit is not None:

        for item in (
            last_three_month_borrowings
        ):

            if (
                item["Amount"] is not None
                and item["Amount"]
                >= fifty_percent_limit
            ):

                borrowing_over_50 = True


    # ==================================================
    # NON-PLEXE BORROWING
    # ==================================================

    non_plexe_borrowing = False


    for item in (
        last_three_month_borrowings
    ):

        description = clean_text(
            item["Description"]
        )


        if (
            description
            and "plexe llc"
            not in description.lower()
        ):

            non_plexe_borrowing = True


    # ==================================================
    # REPAYMENT DATA
    # ==================================================

    repayment_items = client_repayment_data.get(
        client_name,
        []
    )

    forecasted_repayment = None
    historical_repayment = None

    for item in repayment_items:

        if item["Repayment Type"] == "Forecasted Repayment":
            forecasted_repayment = item["Repayment Percent"]

        elif item["Repayment Type"] == "Historical Repayment":
            historical_repayment = item["Repayment Percent"]

    valid_repayments = [
        value
        for value in [
            forecasted_repayment,
            historical_repayment
        ]
        if value is not None
    ]

    if valid_repayments:
        repayment_percent_used = sum(valid_repayments) / len(valid_repayments)
    else:
        repayment_percent_used = None

    if repayment_percent_used is None:
        repayment_signal = 0.0
        repayment_assessment = "NO DATA"

    elif repayment_percent_used >= REPAYMENT_STRONG:
        repayment_signal = 1.0
        repayment_assessment = "STRONG - SUPPORTS KEEP/INCREASE"

    elif repayment_percent_used >= REPAYMENT_GOOD:
        repayment_signal = 0.5
        repayment_assessment = "GOOD - LEANS KEEP"

    elif repayment_percent_used >= REPAYMENT_LOW:
        repayment_signal = 0.0
        repayment_assessment = "MODERATE - NEUTRAL"

    else:
        repayment_signal = -1.0
        repayment_assessment = "LOW - SUPPORTS REDUCTION"


    # ==================================================
    # ESTIMATED BALANCE LIMIT
    # ==================================================

    estimated_balance_limit = (
        calculate_estimated_limit(
            lowest_last_five
        )
    )


    # ==================================================
    # WEIGHTED DECISION FORMULA
    # ==================================================
    # Every constraint produces a signal from -1 (supports reduction)
    # to +1 (supports keeping/increasing). The weights above control how
    # much each signal matters.

    if (
        estimated_balance_limit is None
        or current_limit is None
    ):
        balance_level_signal = 0.0
    elif estimated_balance_limit < current_limit:
        balance_level_signal = -1.0
    elif estimated_balance_limit > current_limit:
        balance_level_signal = 1.0
    else:
        balance_level_signal = 0.0

    if balance_trend == "INCREASING":
        balance_trend_signal = 1.0
    elif balance_trend == "DECREASING":
        balance_trend_signal = -1.0
    else:
        balance_trend_signal = 0.0

    large_borrowing_signal = (
        -1.0 if borrowing_over_50 else 1.0
    )

    non_plexe_signal = (
        -1.0 if non_plexe_borrowing else 1.0
    )

    weighted_components = {
        "balance_level": balance_level_signal,
        "balance_trend": balance_trend_signal,
        "large_borrowing": large_borrowing_signal,
        "non_plexe_borrowing": non_plexe_signal,
        "repayment": repayment_signal
    }

    total_weight = sum(DECISION_WEIGHTS.values())

    if total_weight > 0:
        weighted_decision_score = sum(
            weighted_components[name] * DECISION_WEIGHTS[name]
            for name in DECISION_WEIGHTS
        ) / total_weight
    else:
        weighted_decision_score = 0.0

    if current_limit is not None and current_limit > 0:
        raw_weighted_limit = current_limit * (
            1 + MAX_WEIGHTED_LIMIT_CHANGE * weighted_decision_score
        )

        weighted_limit = max(
            0.0,
            float(
                math.floor(raw_weighted_limit / LIMIT_ROUNDING)
                * LIMIT_ROUNDING
            )
        )
    else:
        weighted_limit = 0.0


    # ==================================================
    # FINAL RECOMMENDATION
    # ==================================================
    # Repayment >= 30% is a strong positive signal. Repayment 25%-30%
    # leans toward keeping the limit, while repayment below 15% is a
    # negative signal. It is not an absolute override; it is combined
    # with the other equally weighted constraints.

    if current_limit is None or current_limit <= 0:
        recommended_action = "KEEP"
        estimated_new_limit = 0.0

    elif weighted_decision_score <= REDUCE_SCORE_THRESHOLD:
        recommended_action = "REDUCE"
        estimated_new_limit = min(current_limit, weighted_limit)

    elif weighted_decision_score >= INCREASE_SCORE_THRESHOLD:
        recommended_action = "INCREASE"
        estimated_new_limit = max(current_limit, weighted_limit)

    else:
        recommended_action = "KEEP"
        estimated_new_limit = current_limit


    # ==================================================
    # REVIEW REASON
    # ==================================================

    reason_parts = []


    if balance_below_limit:

        reason_parts.append(
            "the closing balance was below the current limit "
            "in at least one of the latest 5 months"
        )


    if borrowing_over_50:

        reason_parts.append(
            "borrowing equal to or greater than 50% of the "
            "current limit was identified in the latest 3 months"
        )


    if non_plexe_borrowing:

        reason_parts.append(
            "non-Plexe borrowing was identified "
            "in the latest 3 months"
        )


    if repayment_percent_used is not None:

        if repayment_percent_used >= REPAYMENT_STRONG:
            reason_parts.append(
                f"repayment is strong at {repayment_percent_used:.1f}% "
                "and supports keeping or increasing the limit"
            )

        elif repayment_percent_used >= REPAYMENT_GOOD:
            reason_parts.append(
                f"repayment is {repayment_percent_used:.1f}% and leans "
                "toward keeping the current limit"
            )

        elif repayment_percent_used < REPAYMENT_LOW:
            reason_parts.append(
                f"repayment is low at {repayment_percent_used:.1f}% "
                "and supports reducing the limit"
            )


    # ==================================================
    # REDUCE REASON
    # ==================================================

    if recommended_action == "REDUCE":

        if reason_parts:

            review_reason = (
                "Reduce the limit because "
                + ", ".join(
                    reason_parts
                )
                + f". The suggested new limit is "
                f"${estimated_new_limit:,.0f}."
            )

        else:

            review_reason = (
                "Reduce the limit because the latest "
                "5-month balance history supports a lower limit. "
                f"The suggested new limit is "
                f"${estimated_new_limit:,.0f}."
            )


    # ==================================================
    # INCREASE REASON
    # ==================================================

    elif recommended_action == "INCREASE":

        review_reason = (
            "Increase the limit because the latest 5-month "
            "balance history supports a higher limit, "
            f"the balance trend is {balance_trend.lower()}, "
            "and no significant borrowing risk was identified "
            "in the latest 3 months. "
            f"The suggested new limit is "
            f"${estimated_new_limit:,.0f}."
        )


    # ==================================================
    # KEEP REASON
    # ==================================================

    else:

        if reason_parts:

            review_reason = (
                "Keep the current limit because "
                + ", ".join(
                    reason_parts
                )
                + ", so a limit increase is not recommended "
                "at this time."
            )

        elif (
            balance_trend
            == "DECREASING"
        ):

            review_reason = (
                "Keep the current limit because the balance "
                "trend is decreasing and recent activity does "
                "not support an increase."
            )

        elif (
            balance_trend
            == "INSUFFICIENT DATA"
        ):

            review_reason = (
                "Keep the current limit because there is not "
                "enough recent balance information to support "
                "a limit change."
            )

        else:

            review_reason = (
                "Keep the current limit because the recent "
                "balance and borrowing activity does not "
                "support a limit change."
            )


    # ==================================================
    # WRITE REVIEW
    # ==================================================

    review_sheet.append(
        [
            client_name,
            current_limit,
            disabled_flag,
            disabled_status_text,
            disabled_reason,
            latest_fico,
            latest_fico_date,
            fifty_percent_limit,
            latest_closing_balance,
            lowest_last_five,
            balance_trend,

            (
                "YES"
                if balance_below_limit
                else "NO"
            ),

            largest_borrowing,

            (
                "YES"
                if borrowing_over_50
                else "NO"
            ),

            (
                "YES"
                if non_plexe_borrowing
                else "NO"
            ),

            forecasted_repayment,
            historical_repayment,
            repayment_percent_used,
            repayment_assessment,
            balance_level_signal,
            balance_trend_signal,
            large_borrowing_signal,
            non_plexe_signal,
            repayment_signal,
            (
                f"Balance Level: {balance_level_signal:+g} | "
                f"Balance Trend: {balance_trend_signal:+g} | "
                f"Large Borrowing: {large_borrowing_signal:+g} | "
                f"Non-Plexe Borrowing: {non_plexe_signal:+g} | "
                f"Repayment: {repayment_signal:+g}"
            ),
            weighted_decision_score,
            weighted_limit,
            estimated_new_limit,
            recommended_action,
            review_reason
        ]
    )


# ==================================================
# FORMATTING
# ==================================================

for sheet in [
    balance_sheet,
    borrowing_sheet,
    repayment_sheet,
    disabled_sheet,
    fico_sheet,
    limit_adjustment_sheet,
    status_sheet,
    review_sheet
]:

    sheet.freeze_panes = (
        "A2"
    )

    sheet.auto_filter.ref = (
        sheet.dimensions
    )


    for column_cells in (
        sheet.columns
    ):

        max_length = 0


        column_letter = (
            get_column_letter(
                column_cells[
                    0
                ].column
            )
        )


        for cell in column_cells:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(
                        str(
                            cell.value
                        )
                    )
                )


        sheet.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            60
        )


# ==================================================
# CURRENCY FORMATTING
# ==================================================

for row in balance_sheet.iter_rows(
    min_row=2
):

    row[1].number_format = (
        '$#,##0.00'
    )

    row[3].number_format = (
        '$#,##0.00'
    )


for row in borrowing_sheet.iter_rows(
    min_row=2
):

    row[1].number_format = (
        '$#,##0.00'
    )

    row[5].number_format = (
        '$#,##0.00'
    )


for row in repayment_sheet.iter_rows(
    min_row=2
):

    row[1].number_format = '$#,##0.00'
    row[3].number_format = '0.00"%"'


for row in review_sheet.iter_rows(
    min_row=2
):

    # Currency columns (0-based indexes): Current Limit, 50%, latest balance,
    # lowest balance, largest borrowing, weighted limit, estimated new limit.
    for index in [1, 7, 8, 9, 12, 26, 27]:
        row[index].number_format = '$#,##0.00'

    # Repayment percentages are stored as 30 for 30%, not 0.30.
    for index in [15, 16, 17]:
        row[index].number_format = '0.00"%"'

    # Individual factor scores and final weighted score.
    for index in [19, 20, 21, 22, 23, 25]:
        row[index].number_format = '0.000'

    # Score Breakdown and Review Reason
    row[24].alignment = Alignment(wrap_text=True, vertical="top")
    row[29].alignment = Alignment(wrap_text=True, vertical="top")


# ==================================================
# COLUMN WIDTHS
# ==================================================

limit_adjustment_sheet.column_dimensions["A"].width = 35
limit_adjustment_sheet.column_dimensions["B"].width = 18
limit_adjustment_sheet.column_dimensions["C"].width = 24
limit_adjustment_sheet.column_dimensions["D"].width = 18
limit_adjustment_sheet.column_dimensions["E"].width = 18
limit_adjustment_sheet.column_dimensions["F"].width = 90
limit_adjustment_sheet.column_dimensions["G"].width = 16
limit_adjustment_sheet.column_dimensions["H"].width = 100

for row in limit_adjustment_sheet.iter_rows(
    min_row=2
):
    row[1].number_format = '$#,##0.00'
    row[3].number_format = '$#,##0.00'
    row[4].number_format = '$#,##0.00'
    row[5].alignment = Alignment(
        wrap_text=True,
        vertical="top"
    )
    row[7].alignment = Alignment(
        wrap_text=True,
        vertical="top"
    )


review_sheet.column_dimensions["A"].width = 35
review_sheet.column_dimensions["C"].width = 14
review_sheet.column_dimensions["D"].width = 24
review_sheet.column_dimensions["E"].width = 50
review_sheet.column_dimensions["F"].width = 14
review_sheet.column_dimensions["G"].width = 20
review_sheet.column_dimensions["I"].width = 28
review_sheet.column_dimensions["J"].width = 38
review_sheet.column_dimensions["L"].width = 38
review_sheet.column_dimensions["M"].width = 22
review_sheet.column_dimensions["N"].width = 22
review_sheet.column_dimensions["O"].width = 22
review_sheet.column_dimensions["P"].width = 22
review_sheet.column_dimensions["Q"].width = 38
review_sheet.column_dimensions["Y"].width = 80
review_sheet.column_dimensions["AD"].width = 90


# ==================================================
# COLORS
# ==================================================

red_fill = PatternFill(
    fill_type="solid",
    fgColor="FFC7CE"
)

green_fill = PatternFill(
    fill_type="solid",
    fgColor="C6EFCE"
)

yellow_fill = PatternFill(
    fill_type="solid",
    fgColor="FFEB9C"
)


# ==================================================
# COLOR REVIEW ROWS
# ==================================================

for row_number in range(
    2,
    review_sheet.max_row + 1
):

    action = (
        review_sheet.cell(
            row=row_number,
            column=29
        ).value
    )


    if action == "REDUCE":

        fill = red_fill


    elif action == "INCREASE":

        fill = green_fill


    else:

        fill = yellow_fill


    for column_number in range(
        1,
        31
    ):

        review_sheet.cell(
            row=row_number,
            column=column_number
        ).fill = fill


# ==================================================
# FINAL SAVE
# ==================================================

workbook.save(
    OUTPUT_FILE
)


print(
    "\n\n====================================="
)

print(
    "ALL PROCESSING FINISHED"
)

print(
    "====================================="
)

print(
    "Total Clients:",
    len(clients)
)

print(
    "Batch Size:",
    BATCH_SIZE
)

print(
    "Parallel Browsers:",
    MAX_WORKERS
)

print(
    "Browser Start Delay:",
    BROWSER_START_DELAY,
    "seconds"
)

print(
    "\nFinal Excel:"
)

print(
    OUTPUT_FILE
)

print(
    "\nSheets:"
)

print(
    "1. Monthly Balance"
)

print(
    "2. Borrowing"
)

print(
    "3. Repayment"
)

print(
    "4. Disabled Status (includes Disabled Reason)"
)

print(
    "5. FICO"
)

print(
    "6. Process Status"
)

print(
    "7. Limit Adjustment History"
)

print(
    "8. Limit Review"
)


# ==================================================
# CLIENT HTML REPORT GENERATION
# ==================================================

from config import HTML_REPORT_FOLDER
from functions import (
    ih_clean, ih_number, ih_money, ih_percent, ih_score,
    ih_safe_filename, ih_rows, ih_client_rows,
    ih_fallback_summary, ih_report_summary, ih_json, ih_build_html,
)

# ==================================================
# INTERACTIVE STANDALONE HTML CLIENT REPORTS
# ==================================================
# Creates ONE self-contained .html file per client.
# No PDF files are created.
#
# Features:
# - Works offline after generation
# - No CDN / external JavaScript / external CSS
# - Interactive tabs
# - Interactive balance chart with hover/tap details
# - Interactive decision-score bars
# - Expand/collapse detail sections
# - Print button (browser print only; Python does NOT create a PDF)
# ==================================================

import os
import re
import json
import html
import openpyxl


os.makedirs(
    HTML_REPORT_FOLDER,
    exist_ok=True
)


# ==================================================
# READ COMPLETED EXCEL OUTPUT
# ==================================================

if not os.path.exists(OUTPUT_FILE):
    raise FileNotFoundError(
        "Run the Selenium/Excel workflow first. "
        "Excel output was not found: "
        + OUTPUT_FILE
    )

html_wb = openpyxl.load_workbook(
    OUTPUT_FILE,
    data_only=True
)

for required_sheet in [
    "Monthly Balance",
    "Limit Adjustment History",
    "Limit Review"
]:
    if required_sheet not in html_wb.sheetnames:
        raise ValueError(
            "Required Excel sheet missing: "
            + required_sheet
        )

monthly_data = ih_rows(
    html_wb["Monthly Balance"]
)

adjustment_data = ih_rows(
    html_wb["Limit Adjustment History"]
)

review_data = ih_rows(
    html_wb["Limit Review"]
)


# ==================================================
# CREATE ONE STANDALONE HTML PER CLIENT
# ==================================================

created_html_files = []

for review in review_data:

    client_name = ih_clean(
        review.get("Client Name")
    )

    if not client_name:
        continue

    client_monthly = ih_client_rows(
        monthly_data,
        client_name
    )

    client_adjustments = ih_client_rows(
        adjustment_data,
        client_name
    )

    # Limit Adjustment History contains one latest result per client.
    latest_adjustment = (
        client_adjustments[0]
        if client_adjustments
        else {}
    )

    # Narrative is generated locally from the calculated review values.
    analysis_summary = ih_report_summary(
        review
    )

    safe_client_name = ih_safe_filename(
        client_name
    )

    html_path = os.path.join(
        HTML_REPORT_FOLDER,
        f"{safe_client_name}_Credit_Review.html"
    )

    ih_build_html(
        client_name=client_name,
        review=review,
        adjustment=latest_adjustment,
        monthly_rows=client_monthly,
        narrative=analysis_summary,
        output_path=html_path
    )

    created_html_files.append(
        html_path
    )

    print(
        client_name,
        "| interactive standalone HTML created"
    )


print("\n=====================================")
print("INTERACTIVE HTML REPORTS FINISHED")
print("=====================================")
print("Folder:", HTML_REPORT_FOLDER)
print("HTML files created:", len(created_html_files))
print("PDF generation: DISABLED")
print("Standalone/offline HTML: YES")
