"""
Configuration for the credit-limit-review automation.

Update the file paths and runtime parameters here instead of editing the notebook.
For security, login credentials are read from environment variables.
"""

import os

# ==================================================
# APPLICATION / LOGIN
# ==================================================

URL = "https://admin.plexe.co/login?password=true"

# Set these Windows environment variables before running:
#   PLEXE_USERNAME
#   PLEXE_PASSWORD
USERNAME = os.getenv("PLEXE_USERNAME", "")
PASSWORD = os.getenv("PLEXE_PASSWORD", "")


# ==================================================
# INPUT / OUTPUT FILES
# ==================================================

CLIENT_FILE = r"C:\Users\aishw\Documents\Client file with Credit limit - Copy.xlsx"

OUTPUT_FOLDER = r"C:\Users\aishw\Documents\Final decision"

OUTPUT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "client_balance_borrowing_review_dashboard.xlsx"
)

HTML_REPORT_FOLDER = os.path.join(
    OUTPUT_FOLDER,
    "Client Reports"
)


# ==================================================
# PARALLEL PROCESSING
# ==================================================

BATCH_SIZE = 50
MAX_WORKERS = 5
BROWSER_START_DELAY = 3


# ==================================================
# CREDIT LIMIT DECISION PARAMETERS
# ==================================================

DECISION_WEIGHTS = {
    "balance_level": 1.0,
    "balance_trend": 1.0,
    "large_borrowing": 1.0,
    "non_plexe_borrowing": 1.0,
    "repayment": 1.0,
}

REPAYMENT_STRONG = 30.0
REPAYMENT_GOOD = 25.0
REPAYMENT_LOW = 15.0

MAX_WEIGHTED_LIMIT_CHANGE = 0.25

INCREASE_SCORE_THRESHOLD = 0.20
REDUCE_SCORE_THRESHOLD = -0.20

LIMIT_ROUNDING = 5000
