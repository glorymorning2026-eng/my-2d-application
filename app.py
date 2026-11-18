# ============================================================
# Myanmar 2D Dealer
# Production Version
# Part 1
# ============================================================

import streamlit as st
import pandas as pd
import re
import io

from datetime import datetime
from itertools import permutations, product

import openpyxl

from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side
)

# ============================================================
# APP CONFIG
# ============================================================

APP_NAME = "Myanmar 2D Dealer"

st.set_page_config(
    page_title=APP_NAME,
    page_icon="📒",
    layout="wide"
)

# ============================================================
# STYLE
# ============================================================

st.markdown("""
<style>

.main-header {
    border-left: 6px solid #D97706;
    padding-left: 15px;
    margin-bottom:20px;
}

.stButton>button {
    border-radius:8px;
    font-weight:bold;
}

div[data-testid="stMetricValue"]{
    color:#D97706;
}

.warning-box{
    background:#FEF3C7;
    padding:10px;
    border-radius:6px;
}

</style>
""", unsafe_allow_html=True)

# ============================================================
# DATA CONSTANTS
# ============================================================

COLUMNS = [
    "Date",
    "Day",
    "Customer",
    "BatchID",
    "Number",
    "Amount"
]

DOUBLE_NUMBERS = [
    "00","11","22","33","44",
    "55","66","77","88","99"
]

NAKHAT_NUMBERS = [
    "07","70",
    "18","81",
    "24","42",
    "35","53",
    "69","96"
]

POWER_NUMBERS = [
    "05","50",
    "16","61",
    "27","72",
    "38","83",
    "49","94"
]

BROTHER_PAIRS = [
    "01","12","23","34","45",
    "56","67","78","89","90"
]

VALID_AMOUNT_MIN = 100
VALID_AMOUNT_MAX = 500000

DEFAULT_AMOUNT = 1000

# ============================================================
# SESSION STATE
# ============================================================

if "data" not in st.session_state:
    st.session_state.data = pd.DataFrame(
        columns=COLUMNS
    )

if "preview_rows" not in st.session_state:
    st.session_state.preview_rows = []

if "parser_warnings" not in st.session_state:
    st.session_state.parser_warnings = []

if "parser_errors" not in st.session_state:
    st.session_state.parser_errors = []

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def normalize_number(value):

    value = str(value).strip()

    if not value.isdigit():
        return None

    num = value.zfill(2)

    if len(num) != 2:
        return None

    if not (0 <= int(num) <= 99):
        return None

    return num


def validate_amount(amount):

    try:
        amount = int(amount)
    except:
        return False

    return (
        VALID_AMOUNT_MIN
        <= amount
        <= VALID_AMOUNT_MAX
    )


def safe_int(value):

    try:
        return int(
            str(value).replace(",", "")
        )
    except:
        return None


def add_row(rows, number, amount):

    number = normalize_number(number)

    if number is None:
        return

    rows.append(
        {
            "Number": number,
            "Amount": int(amount)
        }
    )

# ============================================================
# PARSER RESULT MODEL
# ============================================================

class ParserResult:

    def __init__(self):

        self.rows = []

        self.warnings = []

        self.errors = []

        self.unresolved_lines = []

    def add_warning(self, text):

        self.warnings.append(text)

    def add_error(self, text):

        self.errors.append(text)

# ============================================================
# REGEX DEFINITIONS
# ============================================================

RE_DAYSHIN = re.compile(
    r"([\d,]+)\s*နေ့ရှင်း"
)

RE_BK = re.compile(
    r"(\d)\s*(?:bk|BK|B|ဘရိတ်)"
)

RE_PAAT = re.compile(
    r"(\d)\s*(?:ပတ်|ပါတ်|ပက်|ပါ|ပ)"
)

RE_TOP = re.compile(
    r"(\d)\s*(?:ထိပ်|ထိပ်စီး)"
)

RE_BOTTOM = re.compile(
    r"(\d)\s*(?:ပိတ်|နောက်ပိတ်)"
)

RE_WHEEL = re.compile(
    r"(\d{2,5})\s*(?:ခွေ|အခွေ)"
)

RE_SINGLE_R = re.compile(
    r"(\d{1,2})\s*[Rr]\s*(\d+)"
)

RE_AT = re.compile(
    r"(\d{1,2})@(\d+)"
)

# ============================================================
# CORE NUMBER GENERATORS
# ============================================================

def generate_bk_numbers(target_digit):

    results = []

    for i in range(100):

        num = f"{i:02d}"

        digit_sum = (
            int(num[0]) +
            int(num[1])
        ) % 10

        if digit_sum == int(target_digit):
            results.append(num)

    return results


def generate_paat_numbers(target_digit):

    results = []

    digit = str(target_digit)

    for i in range(100):

        num = f"{i:02d}"

        if digit in num:
            results.append(num)

    return results


def generate_top_numbers(digit):

    return [
        f"{digit}{i}"
        for i in range(10)
    ]


def generate_bottom_numbers(digit):

    return [
        f"{i}{digit}"
        for i in range(10)
    ]


def generate_brother_numbers():

    result = []

    seen = set()

    for num in BROTHER_PAIRS:

        rev = num[::-1]

        if num not in seen:
            seen.add(num)
            result.append(num)

        if rev not in seen:
            seen.add(rev)
            result.append(rev)

    return result

# ============================================================
# PREVIEW DATAFRAME
# ============================================================

def rows_to_dataframe(rows):

    if not rows:
        return pd.DataFrame(
            columns=["Number", "Amount"]
        )

    return pd.DataFrame(rows)

# ============================================================
# PART 1 END
# ============================================================
