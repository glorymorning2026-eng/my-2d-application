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
# PARSER UTILITIES
# ============================================================

def extract_amount(value):

    try:
        amount = int(
            str(value)
            .replace(",", "")
            .replace("tt", "")
            .replace("TT", "")
        )

        if validate_amount(amount):
            return amount

    except:
        pass

    return None


def unique_keep_order(items):

    seen = set()

    result = []

    for item in items:

        if item not in seen:

            seen.add(item)

            result.append(item)

    return result


def add_reverse(rows, number, amount):

    number = normalize_number(number)

    if number is None:
        return

    add_row(
        rows,
        number,
        amount
    )

    reverse = number[::-1]

    if reverse != number:

        add_row(
            rows,
            reverse,
            amount
        )


# ============================================================
# PLAIN NUMBERS
# 44 99 11
# 44/99/11
# 44-99-11
# ============================================================

def parse_plain_numbers(line):

    line = line.strip()

    if not re.fullmatch(
        r"\d{2}(?:[\s,./\-]+\d{2})*",
        line
    ):
        return []

    rows = []

    for token in re.findall(
        r"\d{2}",
        line
    ):

        add_row(
            rows,
            token,
            DEFAULT_AMOUNT
        )

    return rows


# ============================================================
# STAR FORMAT
# 11/99/77*3000
# 33/55*1000
# 77*2500
# ============================================================

def parse_star_amount(line):

    match = re.fullmatch(
        r"([\d\s,./\-]+?)\*(\d{2,6})",
        line.strip()
    )

    if not match:
        return None

    amount = extract_amount(
        match.group(2)
    )

    if amount is None:
        return None

    rows = []

    numbers = re.findall(
        r"\d{1,2}",
        match.group(1)
    )

    for num in numbers:

        add_row(
            rows,
            num,
            amount
        )

    return rows


# ============================================================
# SINGLE NUMBER WITH AMOUNT
# 20-3000
# 20 3000
# 20.3000
# 99=250
# ============================================================

def parse_single_with_amount(line):

    match = re.fullmatch(
        r"(\d{1,2})[\s.\-=/_]+(\d{2,6})",
        line.strip()
    )

    if not match:
        return None

    amount = extract_amount(
        match.group(2)
    )

    if amount is None:
        return None

    return [
        {
            "Number": match.group(1).zfill(2),
            "Amount": amount
        }
    ]


# ============================================================
# SINGLE R
# 20R3000
# ============================================================

def parse_single_r(line):

    match = re.fullmatch(
        r"(\d{1,2})\s*[Rr]\s*(\d+)",
        line.strip()
    )

    if not match:
        return None

    amount = extract_amount(
        match.group(2)
    )

    if amount is None:
        return None

    rows = []

    add_reverse(
        rows,
        match.group(1),
        amount
    )

    return rows


# ============================================================
# MULTI R
# 46/96r1500
# 30-35-50-R1500
# ============================================================

def parse_multi_r(line):

    match = re.fullmatch(
        r"([\d./,\-\s]+?)[\-\s]*[Rr]\s*(\d+)",
        line.strip()
    )

    if not match:
        return None

    amount = extract_amount(
        match.group(2)
    )

    if amount is None:
        return None

    rows = []

    numbers = re.findall(
        r"\d{1,2}",
        match.group(1)
    )

    for number in numbers:

        add_reverse(
            rows,
            number,
            amount
        )

    return rows


# ============================================================
# AT FORMAT
# 20@3000
# ============================================================

def parse_at_format(line):

    match = re.fullmatch(
        r"(\d{1,2})@(\d+)",
        line.strip()
    )

    if not match:
        return None

    amount = extract_amount(
        match.group(2)
    )

    if amount is None:
        return None

    rows = []

    add_reverse(
        rows,
        match.group(1),
        amount
    )

    return rows


# ============================================================
# DUAL AT FORMAT
# 20_4000@1000
# 20.4000@1000
# ============================================================

def parse_at_dual(line):

    match = re.fullmatch(
        r"(\d{1,2})\d+@(\d+)",
        line.strip()
    )

    if not match:
        return None

    number = match.group(1).zfill(2)

    base_amount = extract_amount(
        match.group(2)
    )

    reverse_amount = extract_amount(
        match.group(3)
    )

    if (
        base_amount is None or
        reverse_amount is None
    ):
        return None

    rows = []

    add_row(
        rows,
        number,
        base_amount
    )

    reverse = number[::-1]

    if reverse != number:

        add_row(
            rows,
            reverse,
            reverse_amount
        )

    return rows


# ============================================================
# DOT R
# 85.3000.r.500
# ============================================================

def parse_dot_r(line):

    match = re.fullmatch(
        r"(\d{2})\.(\d+)\.[Rr]\.(\d+)",
        line.strip()
    )

    if not match:
        return None

    number = match.group(1)

    base_amount = extract_amount(
        match.group(2)
    )

    reverse_amount = extract_amount(
        match.group(3)
    )

    if (
        base_amount is None or
        reverse_amount is None
    ):
        return None

    rows = []

    add_row(
        rows,
        number,
        base_amount
    )

    reverse = number[::-1]

    if reverse != number:

        add_row(
            rows,
            reverse,
            reverse_amount
        )

    return rows


# ============================================================
# DOUBLES
# အပူး1000
# အပူ500
# ပူး1000
# ============================================================

def parse_double_group(line):

    if not re.search(
        r"(အပူး|အပူ|ပူး)",
        line
    ):
        return None

    amount = DEFAULT_AMOUNT

    amounts = re.findall(
        r"\d+",
        line
    )

    if amounts:

        parsed = extract_amount(
            amounts[-1]
        )

        if parsed is not None:
            amount = parsed

    rows = []

    for number in DOUBLE_NUMBERS:

        add_row(
            rows,
            number,
            amount
        )

    return rows


# ============================================================
# NAKHAT
# ============================================================

def parse_nakhat(line):

    if "နက္ခတ်" not in line:
        return None

    amount = DEFAULT_AMOUNT

    numbers = re.findall(
        r"\d+",
        line
    )

    if numbers:

        parsed = extract_amount(
            numbers[-1]
        )

        if parsed is not None:
            amount = parsed

    rows = []

    for number in NAKHAT_NUMBERS:

        add_row(
            rows,
            number,
            amount
        )

    return rows


# ============================================================
# POWER
# ပါဝါ1000
# power1000
# ============================================================

def parse_power(line):

    if not re.search(
        r"(ပါဝါ|power)",
        line,
        re.I
    ):
        return None

    amount = DEFAULT_AMOUNT

    numbers = re.findall(
        r"\d+",
        line
    )

    if numbers:

        parsed = extract_amount(
            numbers[-1]
        )

        if parsed is not None:
            amount = parsed

    rows = []

    for number in POWER_NUMBERS:

        add_row(
            rows,
            number,
            amount
        )

    return rows


# ============================================================
# BROTHERS
# ============================================================

def parse_brother(line):

    if not re.search(
        r"(ညီအကို|ညီကို)",
        line
    ):
        return None

    amount = DEFAULT_AMOUNT

    numbers = re.findall(
        r"\d+",
        line
    )

    if numbers:

        parsed = extract_amount(
            numbers[-1]
        )

        if parsed is not None:
            amount = parsed

    rows = []

    for number in generate_brother_numbers():

        add_row(
            rows,
            number,
            amount
        )

    return rows


# ============================================================
# END OF PART 2
# ============================================================

# ============================================================
# PART 1 END
# ============================================================
