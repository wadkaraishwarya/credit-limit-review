"""
Reusable functions for the credit-limit-review automation.
"""

import os
import re
import time
import math
import json
import html
import traceback
from datetime import datetime

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from config import (
    URL,
    USERNAME,
    PASSWORD,
    DECISION_WEIGHTS,
    REPAYMENT_STRONG,
    REPAYMENT_GOOD,
    REPAYMENT_LOW,
    MAX_WEIGHTED_LIMIT_CHANGE,
    INCREASE_SCORE_THRESHOLD,
    REDUCE_SCORE_THRESHOLD,
    LIMIT_ROUNDING,
)



def clean_text(value):

    if value is None:
        return ""

    return str(value).strip()


def parse_number(value):

    if value is None:
        return None

    if isinstance(
        value,
        (int, float)
    ):
        return float(value)

    text = clean_text(value)

    if not text:
        return None

    negative = (
        text.startswith("(")
        and text.endswith(")")
    )

    text = (
        text
        .replace("$", "")
        .replace(",", "")
        .replace("(", "")
        .replace(")", "")
        .replace("%", "")
        .strip()
    )

    try:

        number = float(text)

        if negative:
            number = -number

        return number

    except Exception:

        return None


def parse_date(value):

    if value is None:
        return None

    if isinstance(
        value,
        datetime
    ):
        return value

    text = clean_text(value)

    if not text:
        return None

    formats = [
        "%b %d, %Y",
        "%b %d %Y",
        "%B %d, %Y",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d-%m-%Y"
    ]

    for fmt in formats:

        try:

            return datetime.strptime(
                text,
                fmt
            )

        except Exception:

            continue

    return None


def month_index(date_value):

    return (
        date_value.year * 12
        + date_value.month
    )


def calculate_balance_trend(
    balance_records
):

    valid_records = [
        item
        for item in balance_records
        if item.get("Balance") is not None
    ]

    if len(valid_records) < 2:

        return "INSUFFICIENT DATA"


    dated_records = [
        item
        for item in valid_records
        if item.get("Date") is not None
    ]


    if len(dated_records) >= 2:

        dated_records.sort(
            key=lambda x: x["Date"],
            reverse=True
        )

        latest_five = (
            dated_records[:5]
        )

    else:

        latest_five = (
            valid_records[:5]
        )


    if len(latest_five) < 2:

        return "INSUFFICIENT DATA"


    latest_balance = (
        latest_five[0]["Balance"]
    )

    oldest_balance = (
        latest_five[-1]["Balance"]
    )


    if oldest_balance == 0:

        if latest_balance > 0:
            return "INCREASING"

        elif latest_balance < 0:
            return "DECREASING"

        else:
            return "STABLE"


    percentage_change = (
        (
            latest_balance
            - oldest_balance
        )
        / abs(oldest_balance)
    )


    if percentage_change > 0.05:

        return "INCREASING"

    elif percentage_change < -0.05:

        return "DECREASING"

    else:

        return "STABLE"


def calculate_estimated_limit(
    lowest_balance
):

    if lowest_balance is None:

        return None


    if lowest_balance <= 0:

        return 0.0


    return float(
        math.floor(
            lowest_balance / 5000
        )
        * 5000
    )


def process_client(
    client_info
):

    CLIENT_NAME = (
        client_info[
            "Client Name"
        ]
    )

    balance_results = []

    borrowing_results = []

    repayment_results = []

    disabled_results = []

    fico_results = []

    limit_adjustment_results = []

    driver = None

    STEP = "Starting"


    try:

        print(
            "\n====================================="
        )

        print(
            "START:",
            CLIENT_NAME
        )

        print(
            "====================================="
        )


        # ==================================================
        # START CHROME
        # ==================================================

        STEP = "Starting Chrome"


        options = (
            webdriver.ChromeOptions()
        )


        options.add_argument(
            "--window-size=1920,1080"
        )

        options.add_argument(
            "--disable-gpu"
        )

        options.add_argument(
            "--no-sandbox"
        )

        options.add_argument(
            "--disable-dev-shm-usage"
        )


        # Optional:
        #
        # options.add_argument(
        #     "--headless=new"
        # )


        # Selenium Manager automatically detects the installed Chrome
        # version and obtains a compatible ChromeDriver.
        driver = webdriver.Chrome(
            options=options
        )


        wait = WebDriverWait(
            driver,
            30
        )


        # ==================================================
        # SAFE CLICK
        # ==================================================

        def safe_click(
            xpath,
            timeout=30
        ):

            element = WebDriverWait(
                driver,
                timeout
            ).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        xpath
                    )
                )
            )


            driver.execute_script(
                """
                arguments[0].scrollIntoView({
                    block: 'center'
                });
                """,
                element
            )


            time.sleep(
                0.5
            )


            try:

                element.click()

            except Exception:

                driver.execute_script(
                    "arguments[0].click();",
                    element
                )


            return element


        # ==================================================
        # LOGIN
        # ==================================================

        STEP = "Login"


        driver.get(
            URL
        )


        username_field = wait.until(
            EC.visibility_of_element_located(
                (
                    By.ID,
                    "InputEmail1"
                )
            )
        )


        username_field.clear()

        username_field.send_keys(
            USERNAME
        )


        password_field = wait.until(
            EC.visibility_of_element_located(
                (
                    By.ID,
                    "InputPassword1"
                )
            )
        )


        password_field.clear()

        password_field.send_keys(
            PASSWORD
        )


        login_button = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/app-root/"
                    "div/div/div/div/"
                    "app-login/main/"
                    "div/form/button"
                )
            )
        )


        login_button.click()

        time.sleep(10)

        wait.until(
            lambda d:
            "/login" not in d.current_url
        )


        print(
            CLIENT_NAME,
            "| Login successful"
        )


        time.sleep(
            3
        )


        # ==================================================
        # SEARCH CLIENT
        # ==================================================

        STEP = "Search Client"


        search_field = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/app-root/"
                    "div[2]/div/div/"
                    "div/div/div[1]/"
                    "div[1]/input"
                )
            )
        )


        search_field.clear()

        search_field.send_keys(
            CLIENT_NAME
        )


        time.sleep(
            3
        )


        # ==================================================
        # SELECT CLIENT
        # ==================================================

        STEP = "Select Client"


        search_result = wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/app-root/"
                    "div[2]/div/div/"
                    "div[1]/div/div[1]/"
                    "div[1]/div/button"
                )
            )
        )


        result_text = clean_text(
            search_result.get_attribute(
                "textContent"
            )
        )


        print(
            CLIENT_NAME,
            "| Search result:",
            result_text
        )


        search_result.click()


        time.sleep(
            3
        )


        # ==================================================
        # CURRENT LIMIT - SELENIUM
        # ==================================================

        current_limit_xpath = (
            "/html/body/app-root/div[2]/div/div/div[2]/app-home/"
            "div/div[2]/div/div/div/aside/div[2]/div[3]/span[2]"
        )

        current_limit = None
        current_limit_text = ""

        try:
            current_limit_element = WebDriverWait(
                driver,
                20
            ).until(
                EC.presence_of_element_located(
                    (By.XPATH, current_limit_xpath)
                )
            )

            current_limit_text = clean_text(
                current_limit_element.text
            )

            current_limit = parse_number(
                current_limit_text
            )

            # This is the single source of truth for Current Limit.
            # All existing downstream calculations/output use CURRENT_LIMIT.
            CURRENT_LIMIT = current_limit

            print(
                "Current Limit extracted from Selenium:",
                current_limit_text,
                "| Parsed:",
                current_limit
            )

        except Exception as e:
            current_limit = None
            current_limit_text = ""
            CURRENT_LIMIT = None

            print(
                "Could not extract Current Limit from Selenium:",
                repr(e)
            )


        # ==================================================
        # DISABLED STATUS
        # ==================================================
        # Capture this immediately after opening the client.
        # The supplied XPath points to the status text shown on the client page.

        STEP = "Capture Disabled Status"

        try:

            disabled_element = WebDriverWait(
                driver,
                100
            ).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/aside/div[1]/div/button/span[2]"
                    )
                )
            )
            print(disabled_element)

            disabled_text = clean_text(
                disabled_element.get_attribute("textContent")
            )

            disabled_flag = (
                "YES"
                if "disabled" in disabled_text.lower()
                else "NO"
            )

            disabled_results.append(
                {
                    "Status Text": disabled_text,
                    "Disabled": disabled_flag
                }
            )

            print(
                CLIENT_NAME,
                "| Disabled status:",
                disabled_text,
                "| Disabled?",
                disabled_flag
            )

        except Exception as disabled_error:

            disabled_results.append(
                {
                    "Status Text": "",
                    "Disabled": "UNKNOWN"
                }
            )

            print(
                CLIENT_NAME,
                "| Disabled status could not be captured:",
                str(disabled_error)
            )


        # ==================================================
        # DISABLED REASON
        # ==================================================

        STEP = "Capture Disabled Reason"

        if disabled_results and disabled_results[0].get("Disabled") == "YES":

            try:
                disabled_reason_first_button_xpath = (
"/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/div/div[2]/div[1]/app-tab-menu/div/div[2]/button[3]"
                )

                disabled_reason_second_button_xpath = (
"/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/div/div[2]/div[2]/app-lazy-credit-tab-host/app-credit/div/app-tab-menu/div/div[2]/button[5]/span"
                )

                disabled_reason_cell_xpath = (
                    '//*[@id="enabled-disabled-reasons-content"]/div/'
                    'app-enabled-disabled-reasons/ngx-datatable/div/div/'
                    'datatable-body/datatable-scroller/div/'
                    'datatable-row-wrapper[1]/datatable-body-row/div/'
                    'datatable-body-cell[5]'
                )

                print(CLIENT_NAME, "| Opening Disabled Reasons...")

                # First click - find the button fresh
                first_button = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, disabled_reason_first_button_xpath)
                    )
                )
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});",
                    first_button
                )
                time.sleep(1)
                driver.execute_script("arguments[0].click();", first_button)
                print(CLIENT_NAME, "| Disabled Reasons first click successful")
                time.sleep(2)

                # Second click - re-find because Angular may refresh the DOM
                second_button = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, disabled_reason_second_button_xpath)
                    )
                )
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});",
                    second_button
                )
                time.sleep(1)
                driver.execute_script("arguments[0].click();", second_button)
                print(CLIENT_NAME, "| Disabled Reasons second click successful")
                time.sleep(2)

                # Wait until the Disabled Reasons content itself is present
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located(
                        (By.ID, "enabled-disabled-reasons-content")
                    )
                )
                print(CLIENT_NAME, "| Disabled Reasons content loaded")

                # Capture first row, column 5
                disabled_reason_element = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located(
                        (By.XPATH, disabled_reason_cell_xpath)
                    )
                )

                disabled_reason = driver.execute_script(
                    "return arguments[0].textContent;",
                    disabled_reason_element
                )
                disabled_reason = clean_text(disabled_reason)

                if not disabled_reason:
                    disabled_reason = clean_text(disabled_reason_element.text)

                if disabled_reason:
                    disabled_results[0]["Disabled Reason"] = disabled_reason
                    print(CLIENT_NAME, "| Disabled reason:", disabled_reason)
                else:
                    disabled_results[0]["Disabled Reason"] = "BLANK"
                    print(CLIENT_NAME, "| Disabled Reason cell found but value is blank")

            except Exception as disabled_reason_error:
                disabled_results[0]["Disabled Reason"] = "NOT CAPTURED"
                print(
                    CLIENT_NAME,
                    "| Disabled reason could not be captured:",
                    str(disabled_reason_error)
                )

        elif disabled_results:
            disabled_results[0]["Disabled Reason"] = ""


        # ==================================================
        # LATEST FICO SCORE + DATE
        # ==================================================
        STEP = "Capture Latest FICO Date -> Click Date -> Capture FICO"
        latest_fico = None
        latest_fico_date_text = ""

        fico_date_click_1_xpath = (
            "/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/div/div[2]/div[1]/app-tab-menu/div/div[2]/button[4]/span"
        )
        fico_date_click_2_xpath = (
            "/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/div/div[2]/div[2]/app-applicants/div/div/ngb-tabset/ul/li[2]/button"
        )

        # This UL contains ALL FICO report-date tabs, including wrapped rows.
        FICO_DATES_UL_XPATH = (
            "/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/div/div[2]/div[2]/app-applicants/div/div/ngb-tabset/div/div[2]/div/div/app-equifax-data-us/div[2]/app-consumer-credit-reports/div/ngb-tabset/ul"
        )

        # Base path for the report panes. The pane number corresponds to
        # the date-tab position (li[1] -> div[1], li[9] -> div[9], etc.).
        FICO_REPORT_PANES_XPATH = (
            "/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/div/div[2]/div[2]/app-applicants/div/div/ngb-tabset/div/div[2]/div/div/app-equifax-data-us/div[2]/app-consumer-credit-reports/div/ngb-tabset/div"
        )

        # Open Applications -> Equifax.
        try:
            b1 = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, fico_date_click_1_xpath))
            )
            driver.execute_script("arguments[0].click();", b1)
            time.sleep(2)

            b2 = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, fico_date_click_2_xpath))
            )
            driver.execute_script("arguments[0].click();", b2)
            time.sleep(3)

            print(CLIENT_NAME, "| Opened Equifax")

        except Exception as open_error:
            print(CLIENT_NAME, "| Could not open Equifax:", repr(open_error))

        # --------------------------------------------------
        # 1. READ EVERY DATE TAB FROM THE EXACT UL
        # 2. FIND THE TRUE MAXIMUM DATE
        # 3. REMEMBER ITS TAB POSITION
        # 4. CLICK THAT TAB
        # 5. READ FICO FROM THE MATCHING REPORT div[n]
        # --------------------------------------------------
        latest_tab_index = None

        try:
            dates_ul = WebDriverWait(driver, 25).until(
                EC.presence_of_element_located(
                    (By.XPATH, FICO_DATES_UL_XPATH)
                )
            )

            WebDriverWait(driver, 20).until(
                lambda d: len(
                    d.find_elements(
                        By.XPATH,
                        FICO_DATES_UL_XPATH + "/li"
                    )
                ) > 0
            )

            time.sleep(2)

            date_pattern = re.compile(
                r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
                r"\s+\d{1,2},\s+\d{4}\b",
                re.IGNORECASE
            )

            date_tabs = []

            # IMPORTANT: direct LI children only.
            tab_elements = dates_ul.find_elements(
                By.XPATH,
                "./li"
            )

            print(
                CLIENT_NAME,
                "| Total FICO date tabs:",
                len(tab_elements)
            )

            for tab_index, li in enumerate(
                tab_elements,
                start=1
            ):
                try:
                    txt = clean_text(
                        li.get_attribute("textContent")
                    )

                    match = date_pattern.search(
                        txt or ""
                    )

                    if not match:
                        continue

                    date_text = clean_text(
                        match.group(0)
                    ).title()

                    date_obj = parse_date(
                        date_text
                    )

                    if date_obj is None:
                        continue

                    date_tabs.append({
                        "date": date_obj,
                        "text": date_text,
                        "index": tab_index
                    })

                    print(
                        CLIENT_NAME,
                        "| FICO DATE FOUND:",
                        date_text,
                        "| TAB:",
                        tab_index
                    )

                except Exception as tab_error:
                    print(
                        CLIENT_NAME,
                        "| Date tab read error:",
                        tab_index,
                        repr(tab_error)
                    )

            if not date_tabs:
                raise Exception(
                    "No FICO dates found in the supplied UL XPath"
                )

            # True chronological latest date, NOT visually last/first.
            latest_item = max(
                date_tabs,
                key=lambda x: x["date"]
            )

            latest_date_obj = latest_item["date"]
            latest_fico_date_text = latest_item["text"]
            latest_tab_index = latest_item["index"]

            print(
                CLIENT_NAME,
                "| TRUE LATEST FICO DATE:",
                latest_fico_date_text,
                "| TAB:",
                latest_tab_index
            )

            # Re-find the exact tab immediately before clicking.
            latest_li_xpath = (
                FICO_DATES_UL_XPATH
                + f"/li[{latest_tab_index}]"
            )

            latest_li = WebDriverWait(
                driver,
                20
            ).until(
                EC.presence_of_element_located(
                    (By.XPATH, latest_li_xpath)
                )
            )

            clickable_items = latest_li.find_elements(
                By.XPATH,
                ".//button | .//a | .//*[@role='tab']"
            )

            latest_tab = (
                clickable_items[0]
                if clickable_items
                else latest_li
            )

            driver.execute_script(
                """
                arguments[0].scrollIntoView({
                    block:'center',
                    inline:'center'
                });
                """,
                latest_tab
            )

            time.sleep(0.5)

            try:
                latest_tab.click()
            except Exception:
                driver.execute_script(
                    "arguments[0].click();",
                    latest_tab
                )

            print(
                CLIENT_NAME,
                "| CLICKED FICO DATE:",
                latest_fico_date_text,
                "| TAB:",
                latest_tab_index
            )

            # Allow Angular/ng-bootstrap to switch the report pane.
            time.sleep(3)

        except Exception as date_error:
            print(
                CLIENT_NAME,
                "| FICO DATE capture/click failed:",
                repr(date_error)
            )
            latest_fico_date_text = ""
            latest_tab_index = None

        # --------------------------------------------------
        # CAPTURE FICO SCORE
        #
        # User-provided example:
        # .../ngb-tabset/div/div[9]/app-consumer-credit-report/
        # div/div[2]/div/table/tr[1]/td[2]
        #
        # Instead of hard-coding div[9], use the index of the
        # latest date tab selected above.
        # --------------------------------------------------
        try:
            if latest_tab_index is None:
                raise Exception(
                    "Latest FICO tab index was not captured"
                )

            latest_fico_xpath = (
                FICO_REPORT_PANES_XPATH
                + f"/div[{latest_tab_index}]"
                + "/app-consumer-credit-report"
                + "/div/div[2]/div/table/tr[1]/td[2]"
            )

            print(
                CLIENT_NAME,
                "| Latest FICO XPath pane:",
                latest_tab_index
            )

            # Wait for the matching report's FICO cell.
            fico_element = WebDriverWait(
                driver,
                25
            ).until(
                EC.presence_of_element_located(
                    (By.XPATH, latest_fico_xpath)
                )
            )

            # Wait until the cell has non-blank text.
            WebDriverWait(
                driver,
                15
            ).until(
                lambda d: clean_text(
                    d.find_element(
                        By.XPATH,
                        latest_fico_xpath
                    ).get_attribute("textContent")
                ) != ""
            )

            # Re-find after waiting in case Angular refreshed the node.
            fico_element = driver.find_element(
                By.XPATH,
                latest_fico_xpath
            )

            raw_fico_text = clean_text(
                fico_element.get_attribute("textContent")
            )

            if not raw_fico_text:
                raw_fico_text = clean_text(
                    fico_element.text
                )

            print(
                CLIENT_NAME,
                "| RAW LATEST FICO:",
                repr(raw_fico_text)
            )

            # FICO values should be 300-850.
            fico_match = re.search(
                r"(?<!\d)([3-8]\d{2})(?!\d)",
                raw_fico_text or ""
            )

            if not fico_match:
                raise Exception(
                    "FICO cell found, but no valid 3-digit FICO "
                    f"was found. Raw value: {raw_fico_text!r}"
                )

            latest_fico = int(
                fico_match.group(1)
            )

            if not (
                300 <= latest_fico <= 850
            ):
                raise Exception(
                    f"Captured FICO outside expected range: {latest_fico}"
                )

            print(
                CLIENT_NAME,
                "| Latest FICO Date:",
                latest_fico_date_text,
                "| Latest FICO:",
                latest_fico
            )

        except Exception as score_error:
            print(
                CLIENT_NAME,
                "| FICO SCORE capture failed:",
                repr(score_error)
            )

            latest_fico = None

            # Do NOT clear latest_fico_date_text.
            # A score failure must not erase a successfully captured date.

        fico_results.append({
            "Latest FICO": latest_fico,
            "Latest FICO Date": latest_fico_date_text
        })


                # ==================================================
        # BUTTON 2
        # ==================================================

        STEP = "Click Button 2"
        safe_click(
"/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/div/div[2]/div[1]/app-tab-menu/div/div[2]/button[2]/span"
        )



        print(
            CLIENT_NAME,
            "| Button 2 clicked"
        )


        time.sleep(
            3
        )


        # ==================================================
        # CASHFLOW
        # ==================================================

        STEP = "Open Cashflow"


        safe_click(
            "/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/div/div[2]/div[2]/app-details/div/div/app-tab-menu/div/div[2]/button[2]/span"
        )


        print(
            CLIENT_NAME,
            "| Cashflow clicked"
        )


        time.sleep(
            3
        )


        # ==================================================
        # READ ALL MONTHLY BALANCES
        # ==================================================

        STEP = "Read Monthly Balance"

        time.sleep(4)

        monthly_rows_xpath = (
            "/html/body/app-root/div[2]/div/div/div[2]/app-home/"
            "div/div[2]/div/div/div/div/div[2]/div[2]/app-details/"
            "div/div/div/div/div/div/div/"
            "app-bank-transactions-monthly-summary/ngx-datatable/"
            "div/div/datatable-body/datatable-scroller/div/"
            "datatable-row-wrapper"
        )

        monthly_rows = wait.until(
            EC.presence_of_all_elements_located(
                (By.XPATH, monthly_rows_xpath)
            )
        )

        print(
            CLIENT_NAME,
            "| Monthly Summary rows:",
            len(monthly_rows)
        )

        for row_number, row in enumerate(monthly_rows, start=1):

            try:
                # New Monthly Summary table:
                # Column 1 = Month, Column 3 = Balance
                month_text = clean_text(
                    row.find_element(
                        By.XPATH,
                        ".//datatable-body-cell[1]"
                    ).get_attribute("textContent")
                )

                balance_text = clean_text(
                    row.find_element(
                        By.XPATH,
                        ".//datatable-body-cell[3]"
                    ).get_attribute("textContent")
                )

                closing_balance = parse_number(balance_text)

                if not month_text:
                    continue

                balance_results.append(
                    {
                        "Month": month_text,
                        "Date": parse_date(month_text),
                        "Balance": closing_balance
                    }
                )

                print(
                    CLIENT_NAME,
                    "| Balance:",
                    month_text,
                    "|",
                    closing_balance
                )

            except Exception as row_error:
                print(
                    CLIENT_NAME,
                    "| Balance row error:",
                    row_number,
                    repr(row_error)
                )

        print(
            CLIENT_NAME,
            "| Total monthly balances captured:",
            len(balance_results)
        )

        # ==================================================
        # CLICK ACTIVITY SUMMARY FIRST
        # ==================================================
        
        print(f" Clicking Activity Summary...")
        
        activity_summary_xpath = (
"/html/body/app-root/div[2]/div/div/div[2]/app-home/div/div[2]/div/div/div/div/div[2]/div[2]/app-details/div/div/app-tab-menu/div/div[2]/button[3]"
        )
        
        activity_summary_button = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, activity_summary_xpath)
            )
        )
        activity_summary_button.click()

        # ==================================================
        # READ REPAYMENT % (SAME CASHFLOW TABLE)
        # ==================================================
        # Row 1, column 12 = Forecasted Repayment
        # Row 2, column 12 = Historical Repayment

        STEP = "Read Repayment Percent"

        repayment_specs = [
            (1, "Forecasted Repayment"),
            (2, "Historical Repayment")
        ]

        for repayment_row_number, repayment_type in repayment_specs:

            try:
                print("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA12222222222222223")
                print(repayment_row_number)
                print("repayment_element")
                repayment_element = wait.until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '//*[@id="v-pills-activity-summary"]'
                            '/div/div/app-cashflow/ngx-datatable/div/div/'
                            'datatable-body/datatable-scroller/div/'
                            f'datatable-row-wrapper[{repayment_row_number}]/'
                            'datatable-body-row/div/datatable-body-cell[10]'
                        )
                    )
                )
                print(repayment_element)
                print("FOUND REPAYMENT ELEMENT |", repayment_type, "| Row:", repayment_row_number)
                repayment_text = clean_text(
                    repayment_element.get_attribute("textContent")
                )

                print("RAW REPAYMENT TEXT |", repayment_type, "|", repr(repayment_text))

                repayment_percent = parse_number(repayment_text)

                print("PARSED REPAYMENT |", repayment_type, "|", repayment_percent)

                repayment_results.append(
                    {
                        "Repayment Type": repayment_type,
                        "Repayment Text": repayment_text,
                        "Repayment Percent": repayment_percent
                    }
                )

                print(
                    CLIENT_NAME,
                    "|",
                    repayment_type + ":",
                    repayment_percent
                )

            except Exception as repayment_error:

                # Repayment should not make the whole client fail.
                repayment_results.append(
                    {
                        "Repayment Type": repayment_type,
                        "Repayment Text": "",
                        "Repayment Percent": None
                    }
                )

                print(
                    CLIENT_NAME,
                    "|",
                    repayment_type,
                    "not available:",
                    str(repayment_error)
                )


        # ==================================================
        # ANALYTICS
        # ==================================================

        STEP = "Open Analytics"


        print(
            CLIENT_NAME,
            "| Looking for Analytics tab..."
        )


        analytics_xpath = (
            "//*[self::button or self::a]"
            "[.//*[normalize-space()='Analytics'] "
            "or normalize-space()='Analytics']"
        )


        analytics_button = WebDriverWait(
            driver,
            20
        ).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    analytics_xpath
                )
            )
        )


        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});",
            analytics_button
        )


        time.sleep(1)


        try:

            WebDriverWait(
                driver,
                10
            ).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        analytics_xpath
                    )
                )
            ).click()

        except Exception:

            driver.execute_script(
                "arguments[0].click();",
                analytics_button
            )


        print(
            CLIENT_NAME,
            "| Analytics clicked"
        )


        time.sleep(
            3
        )


        # ==================================================
        # WAIT ANALYTICS
        # ==================================================

        STEP = "Wait Analytics"


        wait.until(
            EC.presence_of_element_located(
                (
                    By.ID,
                    "v-pills-analytics"
                )
            )
        )


        print(
            CLIENT_NAME,
            "| Analytics loaded"
        )


        # ==================================================
        # BORROWING
        # ==================================================

        STEP = "Open Borrowing"


        safe_click(
            '//*[@id="v-pills-analytics"]'
            '/div/app-analytics/div/'
            'app-tab-menu/div/div[2]/'
            'button[6]/span'
        )


        print(
            CLIENT_NAME,
            "| Borrowing clicked"
        )


        time.sleep(
            3
        )


        # ==================================================
        # BORROWING TABLE
        # ==================================================

        STEP = "Read Borrowing"


        borrowing_table = wait.until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    '//*[@id="v-pills-analytics"]'
                    '//ngx-datatable'
                    '[.//datatable-body]'
                )
            )
        )


        borrowing_rows = (
            borrowing_table.find_elements(
                By.XPATH,
                ".//datatable-body/"
                "datatable-scroller/div/"
                "datatable-row-wrapper"
            )
        )


        print(
            CLIENT_NAME,
            "| Borrowing rows:",
            len(borrowing_rows)
        )


        for row_number, row in enumerate(
            borrowing_rows,
            start=1
        ):

            try:

                borrowing_date_text = clean_text(
                    row.find_element(
                        By.XPATH,
                        ".//datatable-body-cell[1]"
                    ).get_attribute(
                        "textContent"
                    )
                )


                description = clean_text(
                    row.find_element(
                        By.XPATH,
                        ".//datatable-body-cell[2]"
                    ).get_attribute(
                        "textContent"
                    )
                )


                reason = clean_text(
                    row.find_element(
                        By.XPATH,
                        ".//datatable-body-cell[3]"
                    ).get_attribute(
                        "textContent"
                    )
                )


                amount_text = clean_text(
                    row.find_element(
                        By.XPATH,
                        ".//datatable-body-cell[4]"
                    ).get_attribute(
                        "textContent"
                    )
                )


                borrowing_amount = parse_number(
                    amount_text
                )


                borrowing_date = parse_date(
                    borrowing_date_text
                )


                if (
                    not borrowing_date_text
                    and not description
                    and not reason
                    and borrowing_amount is None
                ):

                    continue


                if (
                    "plexe llc"
                    in description.lower()
                ):

                    borrowing_type = (
                        "PLEXE LLC"
                    )

                else:

                    borrowing_type = (
                        "NON-PLEXE"
                    )


                borrowing_results.append(
                    {
                        "Date Text":
                            borrowing_date_text,

                        "Date":
                            borrowing_date,

                        "Description":
                            description,

                        "Reason":
                            reason,

                        "Amount":
                            borrowing_amount,

                        "Borrowing Type":
                            borrowing_type
                    }
                )


                print(
                    CLIENT_NAME,
                    "| Borrowing:",
                    borrowing_date_text,
                    "|",
                    borrowing_amount,
                    "|",
                    borrowing_type
                )


            except Exception as row_error:

                print(
                    CLIENT_NAME,
                    "| Borrowing row error:",
                    row_number,
                    str(row_error)
                )


        # ==================================================
        # LATEST LIMIT ADJUSTMENT / CREDIT LOGS
        # ==================================================

        STEP = "Read Latest Limit Adjustment"

        try:

            print(
                CLIENT_NAME,
                "| Opening Credit / Limit Adjustment logs..."
            )

            # First click
            safe_click(
                "/html/body/app-root/div[2]/div/div/div[2]/app-home/"
                "div/div[2]/div/div/div/div/div[2]/div[1]/"
                "app-tab-menu/div/div[2]/button[3]/span"
            )

            time.sleep(2)

            # Second click - Credit Logs
            safe_click(
                "/html/body/app-root/div[2]/div/div/div[2]/app-home/"
                "div/div[2]/div/div/div/div/div[2]/div[2]/"
                "app-lazy-credit-tab-host/app-credit/div/"
                "app-tab-menu/div/div[2]/button[6]/span"
            )

            time.sleep(3)

            credit_log_xpath = (
                "/html/body/app-root/div[2]/div/div/div[2]/app-home/"
                "div/div[2]/div/div/div/div/div[2]/div[2]/"
                "app-lazy-credit-tab-host/app-credit/div/div/div/"
                "div/div/app-credit-logs/div/div/div/div[2]"
            )

            credit_log_container = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, credit_log_xpath)
                )
            )

            full_credit_log_text = clean_text(
                credit_log_container.get_attribute("textContent")
            )

            # Get smaller descendant blocks as candidate log entries.
            # The full container text is still saved to Excel for validation.
            candidate_elements = credit_log_container.find_elements(
                By.XPATH,
                ".//*"
            )

            candidate_texts = []

            for element in candidate_elements:

                try:
                    text_value = clean_text(
                        element.get_attribute("textContent")
                    )

                    if text_value:
                        candidate_texts.append(text_value)

                except Exception:
                    pass

            # Include the complete container as a fallback candidate.
            if full_credit_log_text:
                candidate_texts.append(full_credit_log_text)

            target_phrase = (
                "change in the Override Credit Limit Cashflow from"
            )

            # --------------------------------------------------
            # IMPORTANT DATE LOGIC
            # --------------------------------------------------
            # The adjustment date is NOT a date found anywhere
            # inside/after the change text.
            #
            # For every occurrence of:
            # "change in the Override Credit Limit Cashflow from"
            #
            # find the FIRST / NEAREST date appearing BEFORE that
            # phrase in the complete Credit Log text.
            #
            # Then select the occurrence with the latest parsed date.
            # --------------------------------------------------

            matching_candidates = []

            date_pattern = re.compile(
                r"\b(?:"
                r"[A-Z][a-z]{2}\s+\d{1,2},\s+\d{4}"
                r"|[A-Z][a-z]+\s+\d{1,2},\s+\d{4}"
                r"|\d{1,2}/\d{1,2}/\d{4}"
                r"|\d{1,2}/\d{1,2}/\d{2}"
                r"|\d{4}-\d{2}-\d{2}"
                r")\b"
            )

            phrase_pattern = re.compile(
                re.escape(target_phrase),
                flags=re.IGNORECASE
            )

            # --------------------------------------------------
            # LATEST REASON
            # --------------------------------------------------
            # Example:
            # Reasonrecent turnover of $13.8k. once improves can
            # increase back to $25k.  Jun 11, 2025
            #
            # Extract text AFTER "Reason" and BEFORE the next date.
            # For multiple Reason entries, use the Reason whose
            # following date is the latest.
            # --------------------------------------------------

            reason_candidates = []

            reason_pattern = re.compile(
                r"Reason\s*:?\s*",
                flags=re.IGNORECASE
            )

            for reason_match in reason_pattern.finditer(
                full_credit_log_text
            ):

                reason_text_start = reason_match.end()

                next_reason_date_match = date_pattern.search(
                    full_credit_log_text,
                    reason_text_start
                )

                if not next_reason_date_match:
                    continue

                reason_text = clean_text(
                    full_credit_log_text[
                        reason_text_start:
                        next_reason_date_match.start()
                    ]
                )

                reason_date_text = (
                    next_reason_date_match.group(0)
                )

                reason_date = parse_date(
                    reason_date_text
                )

                if (
                    reason_text
                    and reason_date is not None
                ):

                    reason_candidates.append(
                        {
                            "Date":
                                reason_date,

                            "Date Text":
                                reason_date_text,

                            "Reason":
                                reason_text
                        }
                    )

            if reason_candidates:

                latest_reason_item = max(
                    reason_candidates,
                    key=lambda x: x["Date"]
                )

                latest_reason = (
                    latest_reason_item["Reason"]
                )

                latest_reason_date = (
                    latest_reason_item["Date Text"]
                )

            else:

                latest_reason = "Not Found"
                latest_reason_date = "Not Found"

            print(
                CLIENT_NAME,
                "| Latest reason:",
                latest_reason,
                "| Reason date:",
                latest_reason_date
            )


            phrase_matches = list(
                phrase_pattern.finditer(
                    full_credit_log_text
                )
            )

            print(
                CLIENT_NAME,
                "| Limit change entries found:",
                len(phrase_matches)
            )

            for phrase_match in phrase_matches:

                phrase_start = phrase_match.start()

                # Everything BEFORE this specific change phrase.
                text_before_phrase = (
                    full_credit_log_text[:phrase_start]
                )

                # Find all dates before the phrase and use the
                # nearest/last one. This is the date directly
                # associated with this change entry.
                dates_before = list(
                    date_pattern.finditer(
                        text_before_phrase
                    )
                )

                if not dates_before:
                    continue

                nearest_date_match = dates_before[-1]

                adjustment_date_text = (
                    nearest_date_match.group(0)
                )

                adjustment_date = parse_date(
                    adjustment_date_text
                )

                if adjustment_date is None:
                    continue

                # Capture this change message from the phrase
                # until the next date/log entry, or a reasonable
                # maximum length as a fallback.
                next_date_match = date_pattern.search(
                    full_credit_log_text,
                    phrase_match.end()
                )

                if next_date_match:
                    change_text = clean_text(
                        full_credit_log_text[
                            phrase_start:
                            next_date_match.start()
                        ]
                    )
                else:
                    change_text = clean_text(
                        full_credit_log_text[
                            phrase_start:
                            phrase_start + 1000
                        ]
                    )

                # Extract FROM and TO limits.
                limit_match = re.search(
                    r"Override\s+Credit\s+Limit\s+Cashflow\s+from\s*"
                    r"\$?\s*([0-9,]+(?:\.\d+)?)\s*"
                    r"(?:to|->|→)\s*"
                    r"\$?\s*([0-9,]+(?:\.\d+)?)",
                    change_text,
                    flags=re.IGNORECASE
                )

                from_limit = None
                to_limit = None

                if limit_match:

                    from_limit = parse_number(
                        limit_match.group(1)
                    )

                    to_limit = parse_number(
                        limit_match.group(2)
                    )

                matching_candidates.append(
                    {
                        "Date":
                            adjustment_date,

                        "Date Text":
                            adjustment_date_text,

                        "From Limit":
                            from_limit,

                        "To Limit":
                            to_limit,

                        "Change Text":
                            change_text,

                        "Phrase Position":
                            phrase_start
                    }
                )

                print(
                    CLIENT_NAME,
                    "| Limit change candidate:",
                    adjustment_date_text,
                    "| From:",
                    from_limit,
                    "| To:",
                    to_limit
                )

            # Remove duplicates caused by nested HTML elements.
            unique_matches = []
            seen_change_keys = set()

            for item in matching_candidates:

                key = (
                    item["Date Text"],
                    item["From Limit"],
                    item["To Limit"]
                )

                if key in seen_change_keys:
                    continue

                seen_change_keys.add(key)
                unique_matches.append(item)

            if unique_matches:

                # Latest adjustment = change occurrence whose
                # NEAREST DATE BEFORE THE PHRASE is the latest date.
                latest_change = max(
                    unique_matches,
                    key=lambda x: x["Date"]
                )

                limit_adjustment_results.append(
                    {
                        "Latest Adjustment Date":
                            latest_change["Date Text"],

                        "From Limit":
                            latest_change["From Limit"],

                        "To Limit":
                            latest_change["To Limit"],

                        "Latest Limit Change Text":
                            latest_change["Change Text"],

                        "Reason":
                            latest_reason,

                        "Full Credit Log Text":
                            full_credit_log_text,

                        "Result":
                            "FOUND"
                    }
                )

                print(
                    CLIENT_NAME,
                    "| Latest limit adjustment:",
                    latest_change["Date Text"],
                    "| From:",
                    latest_change["From Limit"],
                    "| To:",
                    latest_change["To Limit"]
                )

            else:

                limit_adjustment_results.append(
                    {
                        "Latest Adjustment Date":
                            "Not Found",

                        "From Limit":
                            None,

                        "To Limit":
                            None,

                        "Latest Limit Change Text":
                            "Not Found",

                        "Reason":
                            latest_reason,

                        "Full Credit Log Text":
                            full_credit_log_text,

                        "Result":
                            "NOT FOUND"
                    }
                )

                print(
                    CLIENT_NAME,
                    "| Latest limit adjustment: Not Found"
                )

        except Exception as limit_adjustment_error:

            # Limit history is optional and should not fail the whole client.
            limit_adjustment_results.append(
                {
                    "Latest Adjustment Date":
                        "Not Found",

                    "From Limit":
                        None,

                    "To Limit":
                        None,

                    "Latest Limit Change Text":
                        "Not Found",

                    "Reason":
                        "Not Found",

                    "Full Credit Log Text":
                        "",

                    "Result":
                        "NOT FOUND"
                }
            )

            print(
                CLIENT_NAME,
                "| Limit adjustment history not available:",
                repr(limit_adjustment_error)
            )


        STEP = "Completed"


        print(
            "\nSUCCESS:",
            CLIENT_NAME
        )


        return {
            "Client Name":
                CLIENT_NAME,

            "Current Limit":
                CURRENT_LIMIT,

            "Balances":
                balance_results,

            "Borrowings":
                borrowing_results,

            "Repayments":
                repayment_results,

            "Disabled Status":
                disabled_results,

            "FICO":
                fico_results,

            "Limit Adjustments":
                limit_adjustment_results,

            "Status":
                "SUCCESS",

            "Failed Step":
                "",

            "Error":
                ""
        }


    # ==================================================
    # CLIENT ERROR
    # ==================================================

    except Exception as error:

        print(
            "\n====================================="
        )

        print(
            "CLIENT:",
            CLIENT_NAME
        )

        print(
            "FAILED STEP:",
            STEP
        )

        print(
            "ERROR TYPE:",
            type(error).__name__
        )

        print(
            "ERROR:",
            str(error)
        )

        print(
            "TRACEBACK:"
        )

        traceback.print_exc()


        try:

            print(
                "CURRENT URL:",
                driver.current_url
            )

        except Exception:

            print(
                "CURRENT URL unavailable"
            )


        print(
            "=====================================\n"
        )


        # ==================================================
        # SCREENSHOT
        # ==================================================

        if driver is not None:

            screenshot_name = re.sub(
                r"[^A-Za-z0-9_-]+",
                "_",
                CLIENT_NAME
            )


            screenshot_path = os.path.join(
                OUTPUT_FOLDER,
                f"error_{screenshot_name}.png"
            )


            try:

                driver.save_screenshot(
                    screenshot_path
                )

                print(
                    "Screenshot saved:",
                    screenshot_path
                )

            except Exception:

                pass


        return {
            "Client Name":
                CLIENT_NAME,

            "Current Limit":
                CURRENT_LIMIT,

            "Balances":
                balance_results,

            "Borrowings":
                borrowing_results,

            "Repayments":
                repayment_results,

            "Disabled Status":
                disabled_results,

            "FICO":
                fico_results,

            "Limit Adjustments":
                limit_adjustment_results,

            "Status":
                "ERROR",

            "Failed Step":
                STEP,

            "Error":
                str(error)
        }


    # ==================================================
    # ALWAYS CLOSE BROWSER
    # ==================================================

    finally:

        if driver is not None:

            try:

                driver.quit()

                print(
                    CLIENT_NAME,
                    "| Browser closed"
                )

            except Exception:

                pass


def ih_clean(value):
    if value is None:
        return ""
    return str(value).strip()


def ih_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = (
        str(value)
        .replace("$", "")
        .replace(",", "")
        .replace("%", "")
        .strip()
    )
    if not text:
        return None
    try:
        return float(text)
    except Exception:
        return None


def ih_money(value):
    number = ih_number(value)
    return (
        "${:,.0f}".format(number)
        if number is not None
        else "N/A"
    )


def ih_percent(value):
    number = ih_number(value)
    return (
        "{:.1f}%".format(number)
        if number is not None
        else "N/A"
    )


def ih_score(value):
    number = ih_number(value)
    return (
        "{:+.2f}".format(number)
        if number is not None
        else "N/A"
    )


def ih_safe_filename(value):
    text = ih_clean(value)
    text = re.sub(r'[<>:"/\\|?*]+', "_", text)
    text = re.sub(r"\s+", "_", text)
    return text.strip("._") or "Client"


def ih_rows(sheet):
    values = list(sheet.iter_rows(values_only=True))
    if not values:
        return []

    headers = [ih_clean(x) for x in values[0]]
    output = []

    for row in values[1:]:
        if not any(
            x is not None and ih_clean(x)
            for x in row
        ):
            continue

        output.append(
            {
                headers[i]:
                    row[i] if i < len(row) else None
                for i in range(len(headers))
            }
        )

    return output


def ih_client_rows(rows, client_name):
    return [
        row
        for row in rows
        if ih_clean(row.get("Client Name")) == client_name
    ]


def ih_fallback_summary(review):
    return (
        "The rule-based review produced a "
        f"{ih_clean(review.get('Recommended Action')) or 'N/A'} "
        f"recommendation with a final score of "
        f"{ih_score(review.get('Weighted Decision Score'))}. "
        f"The current limit is {ih_money(review.get('Current Limit'))} "
        f"and the calculated new limit is "
        f"{ih_money(review.get('Estimated New Limit'))}. "
        "The dashboard below shows the individual decision components "
        "and supporting balance, borrowing and repayment indicators."
    )


def ih_report_summary(review):
    # Deterministic local narrative only.
    return ih_fallback_summary(review)


def ih_json(value):
    return json.dumps(
        value,
        ensure_ascii=False
    ).replace("</", "<\\/")


def ih_build_html(
    client_name,
    review,
    adjustment,
    monthly_rows,
    narrative,
    output_path
):
    # Client name is added locally while building the standalone HTML.
    esc = lambda x: html.escape(ih_clean(x))

    action = (
        ih_clean(review.get("Recommended Action"))
        or "N/A"
    ).upper()

    action_class = {
        "INCREASE": "good",
        "REDUCE": "bad",
        "KEEP": "warn"
    }.get(action, "warn")

    score_fields = [
        ("Balance Level", "Balance Level Score"),
        ("Balance Trend", "Balance Trend Score"),
        ("Large Borrowing", "Large Borrowing Score"),
        ("Non-Plexe Borrowing", "Non-Plexe Borrowing Score"),
        ("Repayment", "Repayment Score")
    ]

    score_data = []
    for label, field in score_fields:
        score_data.append(
            {
                "label": label,
                "value": ih_number(review.get(field)) or 0
            }
        )

    balance_data = []
    for row in reversed(monthly_rows):
        month = (
            row.get("Month")
            or row.get("Date")
            or row.get("Month Date")
        )
        balance = ih_number(
            row.get("Closing Balance")
        )
        if balance is None:
            balance = ih_number(
                row.get("Balance")
            )
        if month is not None and balance is not None:
            balance_data.append(
                {
                    "label": ih_clean(month),
                    "value": balance
                }
            )

    # Bound the chart size while retaining recent history.
    balance_data = balance_data[-18:]

    latest_reason = (
        ih_clean(adjustment.get("Reason"))
        or "Not Found"
    )

    disabled_reason = (
        ih_clean(review.get("Disabled Reason"))
        or "Not Found"
    )

    data_json = ih_json(
        {
            "balances": balance_data,
            "scores": score_data
        }
    )

    document = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{esc(client_name)} - Credit Limit Review</title>
<style>
:root {{
  --navy:#10233f;
  --blue:#246bfd;
  --cyan:#18a9bd;
  --green:#15966b;
  --amber:#e39a08;
  --red:#d74c4c;
  --purple:#7257c8;
  --bg:#f4f7fb;
  --card:#ffffff;
  --text:#17233a;
  --muted:#667085;
  --line:#e4e9f2;
}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--bg);color:var(--text);font-family:Inter,"Segoe UI",Arial,sans-serif}}
.shell{{max-width:1180px;margin:28px auto;padding:0 22px 50px}}
.hero{{position:relative;overflow:hidden;background:linear-gradient(125deg,#10233f,#1d4f88 60%,#246bfd);
color:#fff;border-radius:26px;padding:34px 38px;box-shadow:0 16px 42px rgba(16,35,63,.18)}}
.hero:after{{content:"";position:absolute;width:300px;height:300px;border-radius:50%;right:-90px;top:-145px;background:rgba(255,255,255,.08)}}
.eyebrow{{font-size:12px;letter-spacing:1.6px;text-transform:uppercase;font-weight:800;opacity:.75}}
h1{{font-size:34px;margin:8px 0 5px;line-height:1.08}}
.client{{font-size:18px;opacity:.9}}
.action{{display:inline-flex;margin-top:18px;padding:9px 16px;border-radius:999px;font-weight:850;letter-spacing:.5px}}
.action.good{{background:var(--green)}} .action.bad{{background:var(--red)}} .action.warn{{background:var(--amber)}}
.toolbar{{display:flex;gap:10px;flex-wrap:wrap;margin:18px 0}}
button{{border:1px solid var(--line);background:#fff;color:var(--text);border-radius:12px;padding:10px 14px;
font:inherit;font-weight:700;cursor:pointer;min-height:44px}}
button:hover{{border-color:#b9c7dc}} button:focus-visible{{outline:3px solid rgba(36,107,253,.25);outline-offset:2px}}
.tabs{{display:flex;gap:8px;flex-wrap:wrap}}
.tab.active{{background:var(--navy);color:#fff;border-color:var(--navy)}}
.grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:15px}}
.kpi{{background:var(--card);border-radius:17px;padding:17px;border-top:4px solid var(--blue);box-shadow:0 7px 22px rgba(31,45,61,.06)}}
.kpi.green{{border-top-color:var(--green)}} .kpi.amber{{border-top-color:var(--amber)}}
.kpi.red{{border-top-color:var(--red)}} .kpi.purple{{border-top-color:var(--purple)}}
.label{{font-size:11px;color:var(--muted);font-weight:800;text-transform:uppercase;letter-spacing:.55px}}
.value{{font-size:24px;font-weight:850;margin-top:8px;color:var(--navy);overflow-wrap:anywhere}}
.sub{{font-size:12px;color:var(--muted);margin-top:5px;line-height:1.35}}
.panel{{margin-top:20px;background:var(--card);border-radius:18px;padding:23px;box-shadow:0 7px 22px rgba(31,45,61,.055)}}
.panel h2{{font-size:18px;margin:0 0 15px;color:var(--navy)}}
.two{{display:grid;grid-template-columns:1fr 1fr;gap:20px}}
.analysis{{white-space:pre-line;line-height:1.7;color:#344054}}
.callout{{background:#f2f6ff;border-left:5px solid var(--blue);border-radius:10px;padding:15px;line-height:1.6}}
.reason{{background:#fff8e8;border-left-color:var(--amber)}}
.chart-wrap{{position:relative;width:100%;overflow-x:auto}}
svg{{width:100%;min-width:560px;height:310px;display:block}}
.chart-tip{{min-height:24px;margin-top:6px;color:var(--muted);font-size:13px;font-weight:650}}
.score-row{{display:grid;grid-template-columns:160px 1fr 55px;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid var(--line)}}
.track{{height:12px;background:#eef2f7;border-radius:999px;position:relative;overflow:hidden}}
.mid{{position:absolute;left:50%;top:0;bottom:0;width:1px;background:#98a2b3}}
.scorebar{{position:absolute;top:0;bottom:0;border-radius:999px}}
table{{width:100%;border-collapse:collapse}}td{{padding:12px;border-bottom:1px solid var(--line);vertical-align:top}}
td:first-child{{color:var(--muted);font-weight:700;width:48%}}
details{{border-top:1px solid var(--line);padding:13px 0}}summary{{cursor:pointer;font-weight:800;color:var(--navy)}}
.tabpage[hidden]{{display:none}}
.footer{{font-size:11px;color:var(--muted);text-align:center;margin-top:24px;line-height:1.5}}
@media(max-width:850px){{.grid{{grid-template-columns:repeat(2,minmax(0,1fr))}}.two{{grid-template-columns:1fr}}}}
@media(max-width:520px){{.grid{{grid-template-columns:1fr}}.hero{{padding:28px 23px}}h1{{font-size:28px}}.shell{{padding:0 14px 35px}}.score-row{{grid-template-columns:110px 1fr 45px}}}}
@media print{{body{{background:#fff}}.toolbar{{display:none}}.shell{{max-width:none;margin:0;padding:0}}.hero,.panel,.kpi{{box-shadow:none}}.tabpage[hidden]{{display:block}}}}
</style>
</head>
<body>
<main class="shell" id="creditDashboard">
  <header class="hero">
    <div class="eyebrow">Credit Risk Monitoring</div>
    <h1>Credit Limit Review</h1>
    <div class="client">{esc(client_name)}</div>
    <div class="action {action_class}">{esc(action)}</div>
  </header>

  <div class="toolbar">
    <div class="tabs" role="tablist" aria-label="Report sections">
      <button class="tab active" type="button" data-tab="overview">Overview</button>
      <button class="tab" type="button" data-tab="balances">Balances</button>
      <button class="tab" type="button" data-tab="decision">Decision</button>
      <button class="tab" type="button" data-tab="history">History</button>
    </div>
    <button type="button" id="printReport">Print report</button>
  </div>

  <section class="tabpage" data-page="overview">
    <div class="grid">
      <article class="kpi"><div class="label">Current Limit</div><div class="value">{ih_money(review.get("Current Limit"))}</div><div class="sub">Existing approved limit</div></article>
      <article class="kpi {action_class}"><div class="label">Calculated New Limit</div><div class="value">{ih_money(review.get("Estimated New Limit"))}</div><div class="sub">Deterministic calculation</div></article>
      <article class="kpi purple"><div class="label">Latest FICO</div><div class="value">{esc(review.get("Latest FICO")) or "N/A"}</div><div class="sub">{esc(review.get("Latest FICO Date"))}</div></article>
      <article class="kpi green"><div class="label">Repayment Used</div><div class="value">{ih_percent(review.get("Repayment % Used"))}</div><div class="sub">{esc(review.get("Repayment Assessment"))}</div></article>
      <article class="kpi"><div class="label">Latest Balance</div><div class="value">{ih_money(review.get("Latest Closing Balance"))}</div><div class="sub">Most recent closing balance</div></article>
      <article class="kpi amber"><div class="label">Largest Borrowing</div><div class="value">{ih_money(review.get("Largest Borrowing - Last 3 Months"))}</div><div class="sub">Last 3 months</div></article>
      <article class="kpi purple"><div class="label">Final Score</div><div class="value">{ih_score(review.get("Weighted Decision Score"))}</div><div class="sub">Increase ≥ +0.20 · Reduce ≤ −0.20</div></article>
      <article class="kpi red"><div class="label">Disabled</div><div class="value">{esc(review.get("Disabled?")) or "N/A"}</div><div class="sub">{esc(review.get("Disabled Status Text"))}</div></article>
    </div>

    <article class="panel">
      <h2>Executive Analysis</h2>
      <div class="analysis">{esc(narrative)}</div>
    </article>

    <div class="two">
      <article class="panel">
        <h2>Repayment Detail</h2>
        <table>
          <tr><td>Forecasted repayment</td><td>{ih_percent(review.get("Forecasted Repayment %"))}</td></tr>
          <tr><td>Historical repayment</td><td>{ih_percent(review.get("Historical Repayment %"))}</td></tr>
          <tr><td>Repayment used</td><td>{ih_percent(review.get("Repayment % Used"))}</td></tr>
          <tr><td>Assessment</td><td>{esc(review.get("Repayment Assessment"))}</td></tr>
        </table>
      </article>
      <article class="panel">
        <h2>Risk Indicators</h2>
        <table>
          <tr><td>Lowest balance – last 5 months</td><td>{ih_money(review.get("Lowest Balance - Last 5 Months"))}</td></tr>
          <tr><td>Balance trend</td><td>{esc(review.get("Balance Trend - Last 5 Months"))}</td></tr>
          <tr><td>Non-Plexe borrowing</td><td>{esc(review.get("Non-Plexe Borrowing - Last 3 Months"))}</td></tr>
          <tr><td>Disabled reason</td><td>{esc(disabled_reason)}</td></tr>
        </table>
      </article>
    </div>
  </section>

  <section class="tabpage" data-page="balances" hidden>
    <article class="panel">
      <h2>Interactive Monthly Balance Trend</h2>
      <div class="chart-wrap"><svg id="balanceChart" viewBox="0 0 900 300" role="img" aria-label="Monthly closing balance chart"></svg></div>
      <div class="chart-tip" id="balanceTip">Select a point to inspect the month and balance.</div>
    </article>
  </section>

  <section class="tabpage" data-page="decision" hidden>
    <div class="two">
      <article class="panel">
        <h2>Decision Components</h2>
        <div id="scoreBars"></div>
      </article>
      <article class="panel">
        <h2>Decision Summary</h2>
        <table>
          <tr><td>Recommended action</td><td><b>{esc(action)}</b></td></tr>
          <tr><td>Final score</td><td>{ih_score(review.get("Weighted Decision Score"))}</td></tr>
          <tr><td>Current limit</td><td>{ih_money(review.get("Current Limit"))}</td></tr>
          <tr><td>Calculated new limit</td><td>{ih_money(review.get("Estimated New Limit"))}</td></tr>
        </table>
        <div class="callout" style="margin-top:16px">{esc(review.get("Review Reason"))}</div>
      </article>
    </div>
  </section>

  <section class="tabpage" data-page="history" hidden>
    <div class="two">
      <article class="panel">
        <h2>Latest Limit Adjustment</h2>
        <table>
          <tr><td>Adjustment date</td><td>{esc(adjustment.get("Latest Adjustment Date")) or "Not Found"}</td></tr>
          <tr><td>From limit</td><td>{ih_money(adjustment.get("From Limit"))}</td></tr>
          <tr><td>To limit</td><td>{ih_money(adjustment.get("To Limit"))}</td></tr>
        </table>
      </article>
      <article class="panel">
        <h2>Latest Reason</h2>
        <div class="callout reason">{esc(latest_reason)}</div>
      </article>
    </div>

    <article class="panel">
      <h2>Additional Detail</h2>
      <details>
        <summary>Latest limit-change text</summary>
        <p>{esc(adjustment.get("Latest Limit Change Text"))}</p>
      </details>
      <details>
        <summary>Rule-based review reason</summary>
        <p>{esc(review.get("Review Reason"))}</p>
      </details>
    </article>
  </section>

  <div class="footer">
    Standalone HTML: all CSS, JavaScript and report data are embedded in this file.
    Client name is inserted locally after the anonymous LLM narrative is generated.
  </div>
</main>

<script>
(() => {{
  const root = document.getElementById("creditDashboard");
  if (!root || root.dataset.ready === "1") return;
  root.dataset.ready = "1";

  const data = {data_json};

  const tabs = [...root.querySelectorAll(".tab")];
  const pages = [...root.querySelectorAll(".tabpage")];

  tabs.forEach(btn => {{
    btn.addEventListener("click", () => {{
      const target = btn.dataset.tab;
      tabs.forEach(x => x.classList.toggle("active", x === btn));
      pages.forEach(page => {{
        page.hidden = page.dataset.page !== target;
      }});
    }});
  }});

  root.querySelector("#printReport").addEventListener("click", () => window.print());

  const svg = root.querySelector("#balanceChart");
  const tip = root.querySelector("#balanceTip");
  const balances = Array.isArray(data.balances) ? data.balances : [];

  const NS = "http://www.w3.org/2000/svg";
  const add = (name, attrs, text) => {{
    const el = document.createElementNS(NS, name);
    Object.entries(attrs || {{}}).forEach(([k,v]) => el.setAttribute(k, String(v)));
    if (text != null) el.textContent = text;
    svg.appendChild(el);
    return el;
  }};

  if (!balances.length) {{
    add("text", {{x:450,y:150,"text-anchor":"middle",fill:"#667085","font-size":"16"}}, "No monthly balance data available");
  }} else {{
    const W=900,H=300,L=72,R=25,T=24,B=55;
    const vals=balances.map(d=>Number(d.value)).filter(Number.isFinite);
    let min=Math.min(...vals), max=Math.max(...vals);
    if (min===max) {{ min-=1; max+=1; }}
    const pad=(max-min)*0.12;
    min-=pad; max+=pad;
    const x=i=>L+(i*(W-L-R)/Math.max(1,balances.length-1));
    const y=v=>T+(max-v)*(H-T-B)/(max-min);

    for(let i=0;i<5;i++) {{
      const yy=T+i*(H-T-B)/4;
      const value=max-i*(max-min)/4;
      add("line",{{x1:L,y1:yy,x2:W-R,y2:yy,stroke:"#e4e9f2","stroke-width":"1"}});
      add("text",{{x:L-10,y:yy+4,"text-anchor":"end",fill:"#667085","font-size":"11"}},
          "$"+Math.round(value).toLocaleString());
    }}

    const points=balances.map((d,i)=>`${{x(i)}},${{y(Number(d.value))}}`).join(" ");
    add("polyline",{{points,fill:"none",stroke:"#246bfd","stroke-width":"3","stroke-linejoin":"round","stroke-linecap":"round"}});

    balances.forEach((d,i)=>{{
      const cx=x(i), cy=y(Number(d.value));
      const c=add("circle",{{cx,cy,r:6,fill:"#fff",stroke:"#246bfd","stroke-width":"3",tabindex:"0"}});
      const show=()=>{{
        tip.textContent=`${{d.label}} — $${{Number(d.value).toLocaleString(undefined,{{maximumFractionDigits:0}})}}`;
        [...svg.querySelectorAll("circle")].forEach(n=>n.setAttribute("r","6"));
        c.setAttribute("r","9");
      }};
      c.addEventListener("click",show);
      c.addEventListener("focus",show);

      if(i===0 || i===balances.length-1 || balances.length<=8) {{
        add("text",{{x:cx,y:H-20,"text-anchor":"middle",fill:"#667085","font-size":"10"}},String(d.label).slice(0,12));
      }}
    }});
  }}

  const scoreRoot = root.querySelector("#scoreBars");
  const scores = Array.isArray(data.scores) ? data.scores : [];
  scores.forEach(d => {{
    const v = Math.max(-1, Math.min(1, Number(d.value) || 0));
    const row=document.createElement("div");
    row.className="score-row";

    const label=document.createElement("span");
    label.textContent=d.label;

    const track=document.createElement("div");
    track.className="track";
    const mid=document.createElement("span");
    mid.className="mid";
    const bar=document.createElement("span");
    bar.className="scorebar";

    const pct=Math.abs(v)*50;
    if(v>=0) {{
      bar.style.left="50%";
      bar.style.width=pct+"%";
      bar.style.background="#15966b";
    }} else {{
      bar.style.left=(50-pct)+"%";
      bar.style.width=pct+"%";
      bar.style.background="#d74c4c";
    }}
    track.append(mid,bar);

    const value=document.createElement("b");
    value.textContent=(v>=0?"+":"")+v.toFixed(1);
    value.style.color=v>0?"#15966b":v<0?"#d74c4c":"#667085";

    row.append(label,track,value);
    scoreRoot.appendChild(row);
  }});
}})();
</script>
</body>
</html>"""

    with open(
        output_path,
        "w",
        encoding="utf-8"
    ) as f:
        f.write(document)
