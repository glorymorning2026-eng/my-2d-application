import io
import re
from datetime import datetime
from itertools import permutations, product

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# =========================================================
# CONFIG
# =========================================================
st.set_page_config(page_title="Myanmar 2D Dealer", page_icon="📒", layout="wide")

st.markdown(
    """
<style>
.main-header { border-left: 6px solid #D97706; padding-left: 15px; margin-bottom: 20px; }
.stButton>button { border-radius: 8px; font-weight: bold; }
div[data-testid="stMetricValue"] { color: #D97706; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="main-header">
<h1>📒 မြန်မာ 2D Master Dealer App</h1>
<p>Clean &amp; Accurate Parser • Preview &amp; Edit • Warning System • Viber Statement • Excel Export</p>
</div>
""",
    unsafe_allow_html=True,
)

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များ")
max_limit = st.sidebar.number_input(
    "Limit (Max/number)", min_value=0, value=500000, step=50000
)
payout_rate = st.sidebar.number_input(
    "ဆုကြေးနှုန်း (ဆ)", min_value=0, value=80, step=5
)
commission_pct = st.sidebar.number_input(
    "ကော်မရှင် (%)", min_value=0, max_value=100, value=10, step=1
)
winning_number = st.sidebar.text_input("ထွက်မည့် ပေါက်ဂဏန်း", value="")

with st.sidebar.expander("🗓️ နေ့/အချိန်အလိုက် ပေါက်ဂဏန်းများ (၁၀ ကြိမ်စာ)"):
    slot_win_inputs = {}
    for _d in ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"]:
        _c1, _c2 = st.columns(2)
        slot_win_inputs[(_d, "1")] = _c1.text_input(
            f"{_d[:3]} [1]မနက်", value="", key=f"win_{_d}_1", max_chars=2
        )
        slot_win_inputs[(_d, "2")] = _c2.text_input(
            f"{_d[:3]} [2]ညနေ", value="", key=f"win_{_d}_2", max_chars=2
        )

st.sidebar.markdown("---")
st.sidebar.subheader("⏰ Time Lock")
enable_time_lock = st.sidebar.checkbox("Enable Time Lock", value=False)
closing_time = st.sidebar.time_input(
    "Closing Time", value=datetime.strptime("12:00", "%H:%M").time()
)

st.sidebar.markdown("---")
st.sidebar.subheader("🚫 Blocked Numbers")
blocked_input = st.sidebar.text_input("Blocked (comma-separated)", value="")
blocked_numbers = [
    n.strip().zfill(2)
    for n in blocked_input.split(",")
    if n.strip().isdigit() and 0 <= int(n.strip()) <= 99
]

# =========================================================
# SESSION STATE
# =========================================================
COLUMNS = [
    "Date",
    "Day",
    "Session",
    "Customer",
    "BatchID",
    "Category",
    "Description",
    "Number",
    "Amount",
]

for k, v in [
    ("data", pd.DataFrame(columns=COLUMNS)),
    ("pending_parse", None),
    ("confirm_clear", False),
    ("flash_success", None),
    ("flash_warning", None),
]:
    if k not in st.session_state:
        st.session_state[k] = v

if not st.session_state.data.empty and "Session" not in st.session_state.data.columns:
    st.session_state.data["Session"] = "1"

if st.session_state.flash_success:
    st.success(st.session_state.flash_success)
    st.session_state.flash_success = None
if st.session_state.flash_warning:
    st.warning(st.session_state.flash_warning)
    st.session_state.flash_warning = None


# =========================================================
# CLEAN PARSER HELPERS
# =========================================================
MIN_BET_AMOUNT = 100
MAX_BET_AMOUNT = 500000
DOUBLE_NO_INCLUDE_RE = re.compile(r"ပူးမပါ|အပူးမပါ")
DOUBLE_INCLUDE_RE = re.compile(
    r"(?:အပူးအပါ|အပူးပါ|ပူးအပါ|ပူးပါ|ခွေပူး|အခွေပူး|ခေွပူး|အခေွပူး|ခွေအပူး|ခေွအပူး|အခွေအပူး|အခေွအပူး)"
)
KHVE_KEYWORD_RE = re.compile(r"(?:အခွေ|ခွေ|အခေွ|ခေွ)")
REV_MARKER_RE = re.compile(r"[@Rr&]")

KNOWN_KEYWORDS_RE = re.compile(
    r"(?:ပူးမပါ|အပူးမပါ|အခွေပူး|ခွေပူး|အခေွပူး|ခေွပူး|အခွေအပူး|ခွေအပူး|အခေွအပူး|ခေွအပူး|အပူးအပါ|အပူးပါ|ပူးအပါ|ပူးပါ|ညီအကို|ညီကို|စုံပူး|အပူးစုံ|မပူး|အပူးမ|အပူး|ပူး|နက္ခတ်|ပါဝါ|power|အခွေ|ခွေ|အခေွ|ခေွ|ဘရိတ်|bk|ပါတ်|ပတ်|ပက်|ထိပ်စီး|ထိပ်|ထိတ်|နောက်ပိတ်|ပိတ်|နေ့ရှင်း)",
    re.IGNORECASE,
)


def add_bet(rows, number, amount, category="main", desc="", is_default=False):
    n = str(number).strip().zfill(2)
    amt = int(amount)
    if len(n) == 2 and n.isdigit() and 0 <= int(n) <= 99 and amt >= MIN_BET_AMOUNT:
        rows.append({
            "Number": n,
            "Amount": amt,
            "_default": is_default,
            "Category": category,
            "Description": desc or n,
        })
        return True
    return False


def strip_irrelevant_text(text):
    """
    Removes irrelevant words/names/noise (like 'tt', 'နေ့ရှင်း', customer names,
    non-rule Myanmar/English words) while preserving digits, rule keywords,
    and key markers (@, R, r, &, =, *).
    """
    s = re.sub(r"\btt\b", " ", text, flags=re.IGNORECASE)
    s = re.sub(r"(?:tt|ks|mmk|ကျပ်)+", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"နေ့ရှင်း", " ", s)

    protected = []

    def _keep(m):
        protected.append(m.group(0))
        return f" __KW{len(protected) - 1}__ "

    s_prot = KNOWN_KEYWORDS_RE.sub(_keep, s)

    # Remove standalone English words (e.g. names) except single-letter B/R/r
    s_prot = re.sub(r"\b(?![BRr]\b)[A-Za-z]{2,}\b", " ", s_prot)
    # Remove non-rule Myanmar character runs
    s_prot = re.sub(r"[\u1000-\u109F]+", " ", s_prot)

    for idx, kw in enumerate(protected):
        s_prot = s_prot.replace(f"__KW{idx}__", f" {kw} ")

    return s_prot


def clean_noise(text):
    """Strip irrelevant words/names and normalize punctuation separators to spaces."""
    s = strip_irrelevant_text(text.strip())
    s = re.sub(r"[./,;_]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def extract_numbers_and_amounts(text):
    """
    Priority rule:
    - 2-digit numbers (00-99) are treated as bet numbers.
    - Numbers >= 100 are treated as stake amounts.
    Returns (two_digit_nums, amounts_list).
    """
    nums = []
    amounts = []
    for tok in re.findall(r"\d+", text):
        val = int(tok)
        if len(tok) <= 2 and 0 <= val <= 99:
            nums.append(tok.zfill(2))
        elif val >= MIN_BET_AMOUNT:
            amounts.append(val)
    return nums, amounts


def _bk_numbers(digit):
    d = int(digit)
    return [f"{i:02d}" for i in range(100) if (i // 10 + i % 10) % 10 == d]


def _pat_numbers(digit):
    return [f"{i:02d}" for i in range(100) if digit in f"{i:02d}"]


# =========================================================
# LINE PARSER (CLEAR & DETERMINISTIC)
# =========================================================
def parse_line(raw_line):
    rows = []
    warns = []

    line = clean_noise(raw_line)
    if not line:
        return rows, warns

    # 1. ထိပ် / ထိတ်
    m_thate = re.search(r"(\d)\s*(?:ထိပ်စီး|ထိပ်|ထိတ်)\s*(.*)", line, re.IGNORECASE)
    if m_thate:
        dig = m_thate.group(1)
        tail = m_thate.group(2)
        rev_parts = REV_MARKER_RE.split(tail, maxsplit=1)
        _, front_amts = extract_numbers_and_amounts(rev_parts[0])
        amt_front = front_amts[-1] if front_amts else 1000

        for i in range(10):
            add_bet(
                rows,
                f"{dig}{i}",
                amt_front,
                category="r_pat",
                desc=f"{dig} ထိပ် (၁၀ ကွက် × {amt_front:,})",
            )

        if len(rev_parts) == 2:
            _, back_amts = extract_numbers_and_amounts(rev_parts[1])
            amt_back = back_amts[-1] if back_amts else amt_front
            for i in range(10):
                add_bet(
                    rows,
                    f"{i}{dig}",
                    amt_back,
                    category="r_pat",
                    desc=f"{dig} ပိတ် ပြောင်းပြန် (၁၀ ကွက် × {amt_back:,})",
                )
        return rows, warns

    # 2. နောက်ပိတ် / ပိတ်
    m_pate = re.search(r"(\d)\s*(?:နောက်ပိတ်|ပိတ်)\s*(.*)", line, re.IGNORECASE)
    if m_pate:
        dig = m_pate.group(1)
        tail = m_pate.group(2)
        rev_parts = REV_MARKER_RE.split(tail, maxsplit=1)
        _, back_amts = extract_numbers_and_amounts(rev_parts[0])
        amt_back = back_amts[-1] if back_amts else 1000

        for i in range(10):
            add_bet(
                rows,
                f"{i}{dig}",
                amt_back,
                category="r_pat",
                desc=f"{dig} ပိတ် (၁၀ ကွက် × {amt_back:,})",
            )

        if len(rev_parts) == 2:
            _, front_amts = extract_numbers_and_amounts(rev_parts[1])
            amt_front = front_amts[-1] if front_amts else amt_back
            for i in range(10):
                add_bet(
                    rows,
                    f"{dig}{i}",
                    amt_front,
                    category="r_pat",
                    desc=f"{dig} ထိပ် ပြောင်းပြန် (၁၀ ကွက် × {amt_front:,})",
                )
        return rows, warns

    # 3. ညီအကို / ညီကို
    if re.search(r"ညီအကို|ညီကို", line):
        _, amts = extract_numbers_and_amounts(line)
        amt = amts[-1] if amts else 1000
        pairs = ["01", "12", "23", "34", "45", "56", "67", "78", "89", "90"]
        added = set()
        for n in pairs:
            for item in (n, n[::-1]):
                if item not in added:
                    added.add(item)
                    add_bet(
                        rows,
                        item,
                        amt,
                        category="r_pat",
                        desc=f"ညီအကို (၂၀ ကွက် × {amt:,})",
                    )
        return rows, warns

    # 4A. စုံပူး (00, 22, 44, 66, 88)
    if re.search(r"စုံပူး|အပူးစုံ", line):
        _, amts = extract_numbers_and_amounts(line)
        amt = amts[-1] if amts else 1000
        for n in ["00", "22", "44", "66", "88"]:
            add_bet(
                rows, n, amt, category="r_pat", desc=f"စုံပူး (၅ ကွက် × {amt:,})"
            )
        return rows, warns

    # 4B. မပူး (11, 33, 55, 77, 99)
    if re.search(r"မပူး|အပူးမ", line):
        _, amts = extract_numbers_and_amounts(line)
        amt = amts[-1] if amts else 1000
        for n in ["11", "33", "55", "77", "99"]:
            add_bet(
                rows, n, amt, category="r_pat", desc=f"မပူး (၅ ကွက် × {amt:,})"
            )
        return rows, warns

    # 4C. အပူး / ပူး (standalone doubles, excluding ခွေ/ခေွ)
    if (
        re.search(r"အပူ|ပူး", line)
        and not DOUBLE_NO_INCLUDE_RE.search(line)
        and not KHVE_KEYWORD_RE.search(line)
    ):
        _, amts = extract_numbers_and_amounts(line)
        amt = amts[-1] if amts else 1000
        for i in range(10):
            add_bet(
                rows,
                f"{i}{i}",
                amt,
                category="r_pat",
                desc=f"အပူး (၁၀ ကွက် × {amt:,})",
            )
        return rows, warns

    # 5. နက္ခတ်
    if "နက္ခတ်" in line:
        _, amts = extract_numbers_and_amounts(line)
        amt = amts[-1] if amts else 1000
        for n in ["07", "70", "18", "81", "24", "42", "35", "53", "69", "96"]:
            add_bet(
                rows, n, amt, category="r_pat", desc=f"နက္ခတ် (၁၀ ကွက် × {amt:,})"
            )
        return rows, warns

    # 6. ပါဝါ / power
    if re.search(r"ပါဝါ|power", line, re.IGNORECASE):
        _, amts = extract_numbers_and_amounts(line)
        amt = amts[-1] if amts else 1000
        for n in ["05", "50", "16", "61", "27", "72", "38", "83", "49", "94"]:
            add_bet(
                rows, n, amt, category="r_pat", desc=f"ပါဝါ (၁၀ ကွက် × {amt:,})"
            )
        return rows, warns

    # 7. ခွေ / အခွေ / ခေွ / အခေွ (with optional ပူး / အပူးပါ / ပူးအပါ / အပူးအပါ; excluded if ပူးမပါ or omitted)
    m_khve = re.search(
        r"(\d{2,6})\s*(?:အခွေ|ခွေ|အခေွ|ခေွ)(.*)", line
    )
    if m_khve:
        digits = m_khve.group(1)
        tail = m_khve.group(2)
        _, amts = extract_numbers_and_amounts(tail)
        if not amts:
            _, amts = extract_numbers_and_amounts(line)
        amt = amts[-1] if amts else 1000

        explicit_no_double = bool(DOUBLE_NO_INCLUDE_RE.search(line))
        inc_double = (not explicit_no_double) and bool(
            DOUBLE_INCLUDE_RE.search(line)
        )

        combos = (
            product(dict.fromkeys(digits), repeat=2)
            if inc_double
            else permutations(digits, 2)
        )
        added = set()
        for p in combos:
            n = "".join(p)
            if n not in added:
                added.add(n)
                add_bet(
                    rows,
                    n,
                    amt,
                    category="r_pat",
                    desc=f"{digits} ခွေ{' (ပူးပါ)' if inc_double else ' (ပူးမပါ)'} ({amt:,})",
                )
        return rows, warns

    # 8. BK / ဘရိတ် / B
    m_bk = re.search(
        r"((?:\d\s*)+)\s*(?:ဘရိတ်|bk|B)(?![a-zA-Z])(.*)", line, re.IGNORECASE
    )
    if m_bk:
        bk_digits_raw = m_bk.group(1)
        tail = m_bk.group(2)
        _, amts = extract_numbers_and_amounts(tail)
        if not amts:
            _, amts = extract_numbers_and_amounts(line)
        amt = amts[-1] if amts else 1000

        bk_digits = []
        for d in re.findall(r"\d", bk_digits_raw):
            if d not in bk_digits:
                bk_digits.append(d)

        label = " / ".join(f"{d} bk" for d in bk_digits)
        desc_str = f"{label} (ဘရိတ်ဂဏန်း {len(bk_digits)} ခု × {amt:,})"
        for d in bk_digits:
            for n in _bk_numbers(d):
                add_bet(rows, n, amt, category="r_pat", desc=desc_str)
        return rows, warns

    # 9. ပတ် / ပါတ် / ပက် / ပါ / ပ
    m_pat = re.search(
        r"((?:\d\s*)+)\s*(?:ပါတ်|ပတ်|ပက်|ပါ|ပ)(?!ိတ်|ူး)(.*)", line
    )
    if m_pat:
        pat_digits_raw = m_pat.group(1)
        tail = m_pat.group(2)
        _, amts = extract_numbers_and_amounts(tail)
        if not amts:
            _, amts = extract_numbers_and_amounts(line)
        amt = amts[-1] if amts else 1000
        inc_double = (not DOUBLE_NO_INCLUDE_RE.search(line)) and bool(
            DOUBLE_INCLUDE_RE.search(line)
        )

        pat_digits = []
        for d in re.findall(r"\d", pat_digits_raw):
            if d not in pat_digits:
                pat_digits.append(d)

        for d in pat_digits:
            desc_str = (
                f"{d} ပတ် အပူးပါ ({amt:,})"
                if inc_double
                else f"{d} ပတ် (၁၉ ကွက် × {amt:,})"
            )
            for n in _pat_numbers(d):
                add_bet(rows, n, amt, category="r_pat", desc=desc_str)
            if inc_double:
                dbl = f"{d}{d}"
                add_bet(
                    rows,
                    dbl,
                    amt,
                    category="r_pat",
                    desc=f"{d} ပတ် အပူးထပ်ဆောင်း ({dbl} = {amt:,})",
                )
        return rows, warns

    # 10. Reverse markers (R / r / @ / &)
    if REV_MARKER_RE.search(line):
        parts = REV_MARKER_RE.split(line, maxsplit=1)
        left_nums, left_amts = extract_numbers_and_amounts(parts[0])
        _, right_amts = extract_numbers_and_amounts(parts[1])

        if not left_nums and left_amts:
            first_str = str(left_amts[0])
            if len(first_str) >= 2:
                left_nums = [first_str[:2]]
                left_amts = [int(first_str)]
                warns.append(f"Assumed '{line}' → {left_nums[0]}={left_amts[0]}")

        rev_amt = (
            right_amts[-1]
            if right_amts
            else (left_amts[-1] if left_amts else None)
        )
        base_amt = left_amts[-1] if left_amts else rev_amt

        if left_nums and base_amt is not None and rev_amt is not None:
            num_label = " / ".join(left_nums)
            total_k = sum(2 if n != n[::-1] else 1 for n in left_nums)
            desc_str = (
                f"{num_label} R ({total_k} ကွက် × {rev_amt:,})"
                if base_amt == rev_amt
                else f"{num_label} R ({base_amt:,} / {rev_amt:,})"
            )
            for n in left_nums:
                add_bet(rows, n, base_amt, category="r_pat", desc=desc_str)
                if n != n[::-1]:
                    add_bet(
                        rows, n[::-1], rev_amt, category="r_pat", desc=desc_str
                    )
            return rows, warns

        warns.append(f"Could not parse reverse line: '{raw_line.strip()}'")
        return rows, warns

    # 11. Priority rule: 2-digit bet numbers + stake amount (>= 100)
    nums, amts = extract_numbers_and_amounts(line)
    if nums and amts:
        amt = amts[-1]
        num_label = ", ".join(nums)
        desc_str = f"{num_label} (တစ်ကွက် {amt:,})" if len(nums) > 1 else num_label
        for n in nums:
            add_bet(rows, n, amt, category="main", desc=desc_str)
        return rows, warns

    # Standalone 2-digit numbers without amount
    if nums and not amts and re.fullmatch(r"[\d\s\-]+", line):
        num_label = ", ".join(nums)
        for n in nums:
            add_bet(
                rows,
                n,
                1000,
                category="main",
                desc=f"{num_label} (ထပ်ဆောင်း)",
                is_default=True,
            )
        return rows, warns

    warns.append(f"Unrecognized: '{raw_line.strip()}'")
    return rows, warns


# =========================================================
# MIXED LINE SPLITTER & MAIN PARSER
# =========================================================
def split_mixed_line(raw_line):
    line = raw_line.strip()
    if "နေ့ရှင်း" in line:
        return [line]

    has_kw = bool(
        re.search(
            r"(?:bk|ဘရိတ်|ပတ်|ပါတ်|ပက်|ခွေ|ခေွ|ထိပ်|ထိတ်|ပိတ်|စုံပူး|မပူး|အပူး|နက္ခတ်|ပါဝါ|power|ညီအကို|ညီကို)",
            line,
            re.IGNORECASE,
        )
    )
    if has_kw and ("=" in line or "*" in line):
        raw_parts = re.split(r"[/*]+", line)
        expanded = []
        last_amt = None
        for part in raw_parts:
            p = part.strip()
            if not p:
                continue
            m_eq = re.fullmatch(r"(\d{1,2})\s*=\s*(\d+)", p)
            if m_eq:
                expanded.append(f"{m_eq.group(1)} {m_eq.group(2)}")
                last_amt = m_eq.group(2)
                continue
            m_star = re.fullmatch(r"(\d{1,2})\s*\*\s*(\d+)", p)
            if m_star:
                expanded.append(f"{m_star.group(1)} {m_star.group(2)}")
                last_amt = m_star.group(2)
                continue
            if re.fullmatch(r"\d\s*(?:bk|ဘရိတ်|B)", p, re.IGNORECASE):
                expanded.append(p)
                continue
            _, p_amts = extract_numbers_and_amounts(clean_noise(p))
            if p_amts:
                last_amt = str(p_amts[-1])
            expanded.append(p)

        if last_amt:
            return [
                f"{item} {last_amt}"
                if re.fullmatch(r"\d\s*(?:bk|ဘရိတ်|B)", item, re.IGNORECASE)
                else item
                for item in expanded
            ]
        return expanded

    return [line]


def detect_inline_day_session(line, fallback_day, fallback_sess):
    """Detects explicit standalone day/session words only when not colliding with names/codes."""
    day_out = fallback_day
    sess_out = fallback_sess

    day_map = [
        ("MONDAY", r"(?:^|\b)monday(?:\b|$)|တနင်္လာ|တနင်္လါ"),
        ("TUESDAY", r"(?:^|\b)tuesday(?:\b|$)|အင်္ဂါ"),
        ("WEDNESDAY", r"(?:^|\b)wednesday(?:\b|$)|ဗုဒ္ဓဟူး"),
        ("THURSDAY", r"(?:^|\b)thursday(?:\b|$)|ကြာသပတေး"),
        ("FRIDAY", r"(?:^|\b)friday(?:\b|$)|သောကြာ"),
    ]
    for d_code, pat in day_map:
        if re.search(pat, line, re.IGNORECASE):
            day_out = d_code
            break

    if re.search(r"ညနေ|ညနေပိုင်း|\[2\]", line, re.IGNORECASE):
        sess_out = "2"
    elif re.search(r"မနက်|မနက်ပိုင်း|\[1\]", line, re.IGNORECASE):
        sess_out = "1"

    return day_out, sess_out


def parse_bets_locally(text, default_day="MONDAY", default_sess="1"):
    rows, warnings, errors = [], [], []
    pending = []

    raw_lines = (
        text.replace("၊", "\n").replace(";", "\n").replace("\r", "\n").split("\n")
    )
    expanded_lines = []
    for rl in raw_lines:
        if rl.strip():
            expanded_lines.extend(split_mixed_line(rl))

    cur_day = default_day
    cur_sess = default_sess

    for ln, raw in enumerate(expanded_lines, 1):
        line = raw.strip()
        if not line:
            continue

        cur_day, cur_sess = detect_inline_day_session(line, cur_day, cur_sess)

        m_dc = re.search(r"([\d,]+)\s*နေ့ရှင်း", line)
        if m_dc:
            amt = int(m_dc.group(1).replace(",", ""))
            if not pending:
                add_bet(rows, "11", amt, category="main", desc=f"{amt:,} နေ့ရှင်း")
                warnings.append(
                    f"Line {ln}: နေ့ရှင်း {amt:,} (no preceding numbers → 11)"
                )
            else:
                per_amt = amt // len(pending)
                rem = amt % len(pending)
                nums_txt = ", ".join(x["Number"] for x in pending)
                for idx, r in enumerate(pending):
                    r["Amount"] = per_amt + (rem if idx == 0 else 0)
                    r["_default"] = False
                    r["Description"] = f"{amt:,} နေ့ရှင်း ({nums_txt})"
                warnings.append(
                    f"Line {ln}: နေ့ရှင်း {amt:,} → {len(pending)} number(s)"
                )
                pending.clear()
            continue

        try:
            new_rows, new_warns = parse_line(line)
            for r in new_rows:
                r["Day"] = cur_day
                r["Session"] = cur_sess
                if r.get("_default"):
                    pending.append(r)
            rows.extend(new_rows)
            warnings.extend(f"Line {ln}: {w}" for w in new_warns)
        except Exception as e:
            errors.append(f"Line {ln}: {e}")

    for r in rows:
        r.pop("_default", None)

    return {"rows": rows, "warnings": warnings, "errors": errors}


# =========================================================
# PROCESS & ADD TO LEDGER
# =========================================================
def process_and_add_bets(extracted, date_str, day, session, customer):
    valid, blocked = [], []
    invalid, total = 0, 0

    existing = st.session_state.data[
        st.session_state.data["Customer"] == customer
    ]["BatchID"]
    next_n = 1
    if not existing.empty:
        nums = [
            int(m.group(1))
            for b in existing.unique()
            if (m := re.search(r"\((\d+)\)", str(b)))
        ]
        if nums:
            next_n = max(nums) + 1
    batch_id = f"{customer}({next_n})"

    for row in extracted:
        try:
            amt = int(row.get("Amount", 0))
            num = str(row.get("Number", "")).zfill(2)
            cat = str(row.get("Category", "main") or "main")
            desc = str(row.get("Description", num) or num)
            r_day = str(day).strip().upper()
            r_sess = _normalize_session_code(session)
        except (ValueError, TypeError):
            invalid += 1
            continue
        if (
            len(num) != 2
            or not num.isdigit()
            or amt < MIN_BET_AMOUNT
            or amt > MAX_BET_AMOUNT
        ):
            invalid += 1
            continue
        if num in blocked_numbers:
            blocked.append(num)
            continue
        valid.append({
            "Date": date_str,
            "Day": r_day,
            "Session": r_sess,
            "Customer": customer,
            "BatchID": batch_id,
            "Category": cat,
            "Description": desc,
            "Number": num,
            "Amount": amt,
        })
        total += amt

    return valid, blocked, invalid, total


# =========================================================
# WEEKLY CUSTOMER STATEMENT BUILDER (MON - FRI, 10 SESSIONS)
# =========================================================
WEEKDAYS_MON_FRI = [
    ("MONDAY", "တနင်္လာ"),
    ("TUESDAY", "အင်္ဂါ"),
    ("WEDNESDAY", "ဗုဒ္ဓဟူး"),
    ("THURSDAY", "ကြာသပတေး"),
    ("FRIDAY", "သောကြာ"),
]


def _normalize_session_code(val):
    s = str(val).strip().lower()
    if s in ("2", "[2]", "ညနေ", "ညနေပိုင်း", "evening", "pm"):
        return "2"
    return "1"


def _resolve_slot_win_num(day_key, sess_key, default_win, slot_wins=None):
    if slot_wins:
        raw_s = str(slot_wins.get((day_key, sess_key), "")).strip()
        if raw_s.isdigit() and 0 <= int(raw_s) <= 99:
            return raw_s.zfill(2)
    return default_win


def build_customer_weekly_report(
    df, cust_name, payout, default_win_in, slot_wins=None
):
    default_win = (
        str(default_win_in).strip().zfill(2)
        if str(default_win_in).strip().isdigit()
        else ""
    )
    cdf = df[df["Customer"] == cust_name].copy()
    if not cdf.empty:
        cdf["_day_norm"] = cdf["Day"].astype(str).str.strip().str.upper()
        cdf["_sess_norm"] = cdf["Session"].apply(_normalize_session_code)

    lines = [
        f"ဖောက်သည် အမည်- {cust_name}",
        "=" * 32,
    ]

    total_bets_amt = 0.0
    total_win_payout = 0.0
    slot_rows = []

    for day_en, day_mm in WEEKDAYS_MON_FRI:
        lines.append(f"\n{day_en} ({day_mm})")
        ddf = (
            cdf[cdf["_day_norm"] == day_en]
            if not cdf.empty
            else cdf.iloc[0:0]
        )

        for sess_code, sess_mm in [("1", "မနက်ပိုင်း"), ("2", "ညနေပိုင်း")]:
            sdf = (
                ddf[ddf["_sess_norm"] == sess_code]
                if not ddf.empty
                else ddf.iloc[0:0]
            )
            bet_sum = float(sdf["Amount"].sum()) if not sdf.empty else 0.0
            w_num = _resolve_slot_win_num(
                day_en, sess_code, default_win, slot_wins
            )
            win_stake = (
                float(sdf[sdf["Number"] == w_num]["Amount"].sum())
                if (not sdf.empty and w_num)
                else 0.0
            )
            win_pay = win_stake * payout
            res_val = bet_sum - win_pay

            total_bets_amt += bet_sum
            total_win_payout += win_pay

            slot_rows.append({
                "Day": f"{day_en} ({day_mm})",
                "Session": f"[{sess_code}]({sess_mm})",
                "Bets": bet_sum,
                "WinPayout": win_pay,
                "Result": res_val,
            })

            lines.append(f"[{sess_code}]({sess_mm}) :")
            lines.append(
                f"({bet_sum:,.0f}) - ({win_pay:,.0f})=({res_val:,.0f})"
            )

    total_net_res = total_bets_amt - total_win_payout

    if total_net_res > 0:
        outcome_lbl = f"ဒိုင်အနိုင် / ထိုးသူပေးရန် (+{total_net_res:,.0f} ကျပ်)"
    elif total_net_res < 0:
        outcome_lbl = f"ထိုးသူအနိုင် / ဒိုင်ပေးရန် ({total_net_res:,.0f} ကျပ်)"
    else:
        outcome_lbl = "အညီအမျှ (0 ကျပ်)"

    lines.extend([
        "\n" + "=" * 32,
        f"📌 *{cust_name}* — တစ်ပတ်စာ (၁၀ ကြိမ်စာ) ချုပ်",
        f"• ၁၀ ကြိမ်စာ စုစုပေါင်း ထိုးငွေ : {total_bets_amt:,.0f} ကျပ်",
        f"• ၁၀ ကြိမ်စာ စုစုပေါင်း ပေါက်ငွေ : {total_win_payout:,.0f} ကျပ်",
        f"• ဖော်မြူလာ : ({total_bets_amt:,.0f}) - ({total_win_payout:,.0f})=({total_net_res:,.0f})",
        f"• အရှုံး / အနိုင် ရလဒ် : *{outcome_lbl}*",
    ])

    return "\n".join(lines), {
        "customer": cust_name,
        "bets_total": total_bets_amt,
        "win_payout": total_win_payout,
        "net_result": total_net_res,
        "outcome_label": outcome_lbl,
        "slots": slot_rows,
    }


def build_all_customers_weekly_report(
    df, payout, default_win_in, slot_wins=None
):
    custs = sorted(df["Customer"].unique())
    blocks = []
    for c_name in custs:
        txt, _ = build_customer_weekly_report(
            df, c_name, payout, default_win_in, slot_wins
        )
        blocks.append(txt)
    return ("\n\n" + "#" * 36 + "\n\n").join(blocks)


# =========================================================
# EXCEL EXPORT
# =========================================================
def generate_styled_excel(
    df, comm_pct, payout, win_num_in, limit, slot_wins=None
):
    out = io.BytesIO()
    wb = openpyxl.Workbook()

    hdr_fill = PatternFill("solid", fgColor="1E293B")
    accent = PatternFill("solid", fgColor="FEF3C7")
    zebra = PatternFill("solid", fgColor="F8FAFC")
    profit = PatternFill("solid", fgColor="DCFCE7")
    loss = PatternFill("solid", fgColor="FEE2E2")
    f_title = Font(name="Pyidaungsu", size=14, bold=True, color="1E293B")
    f_hdr = Font(name="Pyidaungsu", size=11, bold=True, color="FFFFFF")
    f_body = Font(name="Pyidaungsu", size=11, color="0F172A")
    f_bold = Font(name="Pyidaungsu", size=11, bold=True, color="0F172A")
    bd = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    def sheet_fmt(ws, widths):
        ws.views.sheetView[0].showGridLines = True
        for c, w in widths.items():
            ws.column_dimensions[c].width = w
        ws.freeze_panes = "A4"

    def hdr(ws, row, headers):
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=i, value=h)
            cell.fill, cell.font = hdr_fill, f_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = bd

    win = str(win_num_in).strip().zfill(2) if win_num_in else ""
    tot = float(df["Amount"].sum())
    comm = tot * comm_pct / 100
    net = tot - comm
    win_tot = float(df[df["Number"] == win]["Amount"].sum()) if win else 0
    payout_amt = win_tot * payout
    pl = net - payout_amt

    ws = wb.active
    ws.title = "စာရင်းချုပ်အကျဉ်း"
    ws["A1"] = "မြန်မာနှစ်လုံးထီ (2D) ငွေစာရင်းချုပ်"
    ws["A1"].font = f_title
    ws["A2"] = f"ထုတ်ပြန်သည့်ရက်စွဲ - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A2"].font = f_body

    summ = [
        ("အကြောင်းအရာ", "ပမာဏ", "မှတ်ချက်"),
        ("ရောင်းရငွေ", tot, f"အကွက် {len(df):,}"),
        (f"ကော်မရှင် ({comm_pct}%)", comm, ""),
        ("Net Sales", "=B5-B6", ""),
        ("ထွက်ဂဏန်း", win or "N/A", f"{payout} ဆ"),
        ("ပေါက်ဂဏန်းတန်ဖိုး", win_tot, ""),
        ("ဆုကြေး", f"=B9*{payout}", ""),
        ("အသားတင်", "=B7-B10", "(+) / (-)"),
    ]
    for i, vals in enumerate(summ):
        r = i + 4
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.border = bd
            if i == 0:
                c.fill, c.font = hdr_fill, f_hdr
                c.alignment = Alignment(horizontal="center")
            else:
                c.font = f_body
                if i == 3:
                    c.fill = accent
                if i == 7:
                    c.fill = profit if pl >= 0 else loss
                if j == 2 and i != 4:
                    c.number_format = "#,##0"
                    c.alignment = Alignment(horizontal="right")
    sheet_fmt(ws, {"A": 42, "B": 24, "C": 38})

    ws = wb.create_sheet("နေ့အလိုက်_အပတ်စဉ်")
    ws["A1"] = "နေ့အလိုက် စာရင်းချုပ်"
    ws["A1"].font = f_title
    dg = (
        df.groupby(["Date", "Day"])
        .agg(Bets=("Number", "count"), Sales=("Amount", "sum"))
        .reset_index()
    )
    hdr(
        ws,
        3,
        ["ရက်စွဲ", "နေ့", "အကွက်", "ထိုးကြေး", f"ကော်မရှင် ({comm_pct}%)", "Net"],
    )
    for r, row in enumerate(dg.itertuples(), 4):
        vals = [
            row.Date,
            str(row.Day),
            int(row.Bets),
            float(row.Sales),
            f"=D{r}*{comm_pct/100}",
            f"=D{r}-E{r}",
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, bd
            if ci >= 3:
                c.number_format = "#,##0"
            if r % 2:
                c.fill = zebra
    sheet_fmt(ws, {"A": 18, "B": 16, "C": 14, "D": 22, "E": 20, "F": 22})

    ws = wb.create_sheet("ဖောက်သည်အလိုက်")
    ws["A1"] = "ဖောက်သည်အလိုက် ရှင်းတမ်း"
    ws["A1"].font = f_title
    cg = (
        df.groupby("Customer")["Amount"]
        .sum()
        .reset_index()
        .sort_values("Amount", ascending=False)
    )
    hdr(
        ws,
        3,
        [
            "ဖောက်သည်",
            "ထိုးကြေး",
            f"ကော်မရှင် ({comm_pct}%)",
            "Net",
            f"ဆုကြေး ({payout}ဆ)",
            "အသားတင်",
        ],
    )
    for r, row in enumerate(cg.itertuples(), 4):
        cwin = 0
        if win:
            cwin = (
                float(
                    df[(df["Customer"] == row.Customer) & (df["Number"] == win)][
                        "Amount"
                    ].sum()
                )
                * payout
            )
        vals = [
            row.Customer,
            float(row.Amount),
            f"=B{r}*{comm_pct/100}",
            f"=B{r}-C{r}",
            cwin,
            f"=D{r}-E{r}",
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, bd
            if ci > 1:
                c.number_format = "#,##0"
            if r % 2:
                c.fill = zebra
    sheet_fmt(ws, {"A": 26, "B": 22, "C": 20, "D": 24, "E": 24, "F": 24})

    # Single-sheet Weekly Customer Statement (Monday -> Friday, 10 Sessions + Summary)
    ws_wk = wb.create_sheet("ဖောက်သည်_တစ်ပတ်စာ")
    ws_wk["A1"] = "ဖောက်သည် တစ်ပတ်စာ (၁၀ ကြိမ်စာ) စာရင်းရှင်းတမ်း (Mon - Fri)"
    ws_wk["A1"].font = f_title
    ws_wk["A2"] = f"ထုတ်ပြန်သည့်ရက်စွဲ - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws_wk["A2"].font = f_body
    hdr(
        ws_wk,
        3,
        [
            "ဖောက်သည်",
            "နေ့",
            "အချိန်ပိုင်း",
            "စုစုပေါင်းထိုးငွေ",
            "ပေါက်ဂဏန်းပမာဏ",
            "ရလဒ် (-)",
        ],
    )
    cur_r = 4
    df_wk = df.copy()
    df_wk["_day_norm"] = df_wk["Day"].astype(str).str.strip().str.upper()
    df_wk["_sess_norm"] = df_wk["Session"].apply(_normalize_session_code)

    for cust_name in sorted(df_wk["Customer"].unique()):
        cdf = df_wk[df_wk["Customer"] == cust_name]
        w_bets = 0.0
        w_pay = 0.0
        for day_en, day_mm in WEEKDAYS_MON_FRI:
            ddf = cdf[cdf["_day_norm"] == day_en]
            for sess_code, sess_mm in [
                ("1", "[1](မနက်ပိုင်း)"),
                ("2", "[2](ညနေပိုင်း)"),
            ]:
                sdf = ddf[ddf["_sess_norm"] == sess_code]
                b_amt = float(sdf["Amount"].sum()) if not sdf.empty else 0.0
                wn = _resolve_slot_win_num(day_en, sess_code, win, slot_wins)
                wstk = (
                    float(sdf[sdf["Number"] == wn]["Amount"].sum())
                    if (not sdf.empty and wn)
                    else 0.0
                )
                p_amt = wstk * payout
                r_amt = b_amt - p_amt
                w_bets += b_amt
                w_pay += p_amt

                row_vals = [
                    cust_name,
                    f"{day_en} ({day_mm})",
                    sess_mm,
                    b_amt,
                    p_amt,
                    r_amt,
                ]
                for ci, val in enumerate(row_vals, 1):
                    cell = ws_wk.cell(row=cur_r, column=ci, value=val)
                    cell.font, cell.border = f_body, bd
                    if ci >= 4:
                        cell.number_format = "#,##0"
                        cell.alignment = Alignment(horizontal="right")
                cur_r += 1

        w_res = w_bets - w_pay
        sum_vals = [
            f"{cust_name} (၁၀ ကြိမ်စာ ချုပ်)",
            "Mon - Fri",
            "အနိုင်" if w_res > 0 else ("အရှုံး" if w_res < 0 else "ညီမျှ"),
            w_bets,
            w_pay,
            w_res,
        ]
        for ci, val in enumerate(sum_vals, 1):
            cell = ws_wk.cell(row=cur_r, column=ci, value=val)
            cell.font, cell.border = f_bold, bd
            cell.fill = accent if ci < 6 else (profit if w_res >= 0 else loss)
            if ci >= 4:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right")
        cur_r += 1

    sheet_fmt(
        ws_wk,
        {"A": 28, "B": 24, "C": 20, "D": 22, "E": 22, "F": 24},
    )

    ws = wb.create_sheet("ဂဏန်း_Limit")
    ws["A1"] = "ဂဏန်းအလိုက် Limit"
    ws["A1"].font = f_title
    ng = (
        df.groupby("Number")["Amount"]
        .sum()
        .reset_index()
        .sort_values("Amount", ascending=False)
    )
    hdr(ws, 3, ["2D", "စုစုပေါင်း", "ဒိုင်လက်ခံ", "ကျော်လွန်"])
    for r, row in enumerate(ng.itertuples(), 4):
        nc = ws.cell(row=r, column=1, value=str(row.Number).zfill(2))
        nc.number_format, nc.alignment = "@", Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=float(row.Amount)).number_format = (
            "#,##0"
        )
        ws.cell(row=r, column=3, value=f"=MIN(B{r},{limit})").number_format = (
            "#,##0"
        )
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = "#,##0"
        for ci in range(1, 5):
            ws.cell(row=r, column=ci).font = f_body
            ws.cell(row=r, column=ci).border = bd
        if float(row.Amount) > limit:
            ws.cell(row=r, column=4).fill = loss
            ws.cell(row=r, column=4).font = f_bold
    sheet_fmt(ws, {"A": 16, "B": 24, "C": 24, "D": 26})

    ws = wb.create_sheet("အသေးစိတ်")
    ws["A1"] = "အသေးစိတ် မှတ်တမ်း"
    ws["A1"].font = f_title
    dg = (
        df.groupby(["Date", "Day", "Session", "Customer", "BatchID", "Number"])[
            "Amount"
        ]
        .sum()
        .reset_index()
    )
    hdr(ws, 3, ["ရက်စွဲ", "နေ့", "အချိန်", "ဖောက်သည်", "Batch", "2D", "ထိုးငွေ"])
    for r, row in enumerate(dg.itertuples(), 4):
        vals = [
            row.Date,
            row.Day,
            f"[{row.Session}]",
            row.Customer,
            row.BatchID,
            str(row.Number).zfill(2),
            float(row.Amount),
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, bd
            if ci == 6:
                c.number_format = "@"
            if ci == 7:
                c.number_format = "#,##0"
            if r % 2:
                c.fill = zebra
    sheet_fmt(
        ws,
        {"A": 16, "B": 16, "C": 12, "D": 22, "E": 22, "F": 16, "G": 22},
    )

    wb.save(out)
    return out.getvalue()


# =========================================================
# ENTRY SECTION
# =========================================================
st.subheader("📝 စာရင်းသွင်းရန်")

is_closed = False
if enable_time_lock and datetime.now().time() >= closing_time:
    is_closed = True
    st.error(
        f"⚠️ စာရင်းပိတ်ချိန် ({closing_time.strftime('%H:%M')}) ကျော်လွန်ပါပြီ။"
    )

c1, c2, c3 = st.columns(3)
with c1:
    entry_date = st.date_input("📆 ရက်စွဲ", value=datetime.now().date())
with c2:
    day_opts = [
        "MONDAY",
        "TUESDAY",
        "WEDNESDAY",
        "THURSDAY",
        "FRIDAY",
        "WEEKEND",
    ]
    selected_day = st.selectbox("📅 နေ့", day_opts, index=0)
with c3:
    selected_session = st.selectbox(
        "🕒 အချိန်ပိုင်း",
        ["1", "2"],
        index=0,
        format_func=lambda x: (
            "[1] မနက်ပိုင်း" if x == "1" else "[2] ညနေပိုင်း"
        ),
    )
date_str = entry_date.strftime("%Y-%m-%d")

col_a, col_b = st.columns([1, 2])
with col_a:
    cust = st.text_input("ဖောက်သည် အမည်", placeholder="ဥပမာ - စည်သူ")
with col_b:
    text_in = st.text_area(
        "ဖောက်သည် ပို့လိုက်သော စာသား",
        placeholder="5016ခွေပူး500\n1ဘရိတ်750\n4ဘရိတ်.500\n10.40.90.68.31.83.56.r500\n99.94.14.19.48.89r300",
        height=180,
    )

if st.button("🔍 Parse & Preview", type="primary"):
    if is_closed:
        st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပါသည်။")
    elif not cust.strip() or not text_in.strip():
        st.warning("ဖောက်သည်အမည်နှင့် စာသားကို ဖြည့်ပါ။")
    else:
        res = parse_bets_locally(
            text_in, default_day=selected_day, default_sess=selected_session
        )
        st.session_state.pending_parse = {
            "customer": cust.strip(),
            "date": date_str,
            "day": selected_day,
            "session": selected_session,
            "rows": res["rows"],
            "warnings": res["warnings"],
            "errors": res["errors"],
        }
        st.rerun()

# =========================================================
# PREVIEW SECTION
# =========================================================
p = st.session_state.pending_parse
if p:
    st.markdown("---")
    sess_lbl = "[1] မနက်ပိုင်း" if p.get("session", "1") == "1" else "[2] ညနေပိုင်း"
    st.subheader(
        f"👁️ Preview — {p['customer']} ({p['date']} / {p['day']} / {sess_lbl})"
    )

    for e in p["errors"]:
        st.error(e)
    for w in p["warnings"]:
        st.warning(w)

    if p["rows"]:
        df_prev = pd.DataFrame(p["rows"])[
            ["Number", "Amount", "Category", "Description", "Day", "Session"]
        ]
        total_bets = len(df_prev)
        total_amt = df_prev["Amount"].sum()

        st.info(
            f"**စုစုပေါင်း အကွက်ရေ:** {total_bets} ကွက် | **စုစုပေါင်း ပမာဏ:** {total_amt:,.0f} ကျပ်"
        )

        if winning_number.strip():
            win_check = winning_number.strip().zfill(2)
            win_matches = df_prev[df_prev["Number"] == win_check]
            if not win_matches.empty:
                st.warning(
                    f"🏆 **Winning Number ({win_check}) ပါဝင်နေပါသည်။** ထိုးငွေ: {win_matches['Amount'].sum():,.0f} ကျပ်"
                )
            else:
                st.success(f"✅ **Winning Number ({win_check}) မပါဝင်ပါ။**")

        edited = st.data_editor(
            df_prev[["Number", "Amount", "Day", "Session"]],
            hide_index=True,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "Number": st.column_config.TextColumn("2D", max_chars=2),
                "Amount": st.column_config.NumberColumn(
                    "Amount",
                    min_value=MIN_BET_AMOUNT,
                    max_value=MAX_BET_AMOUNT,
                    step=100,
                ),
                "Day": st.column_config.TextColumn("Day"),
                "Session": st.column_config.TextColumn("Session (1/2)"),
            },
            key="parse_preview",
        )

        cs, cc = st.columns(2)
        with cs:
            if st.button("✅ Save to Ledger", type="primary", use_container_width=True):
                rows_fixed = []
                orig_meta = df_prev[["Category", "Description"]].to_dict(
                    "records"
                )
                for idx, row in edited.iterrows():
                    try:
                        num = str(row.get("Number", "") or "").strip()
                        if not num or not num.isdigit() or len(num) > 2:
                            continue
                        amt_val = row.get("Amount")
                        if pd.isna(amt_val):
                            continue
                        amt = int(amt_val)
                        if amt < MIN_BET_AMOUNT:
                            continue
                        meta = (
                            orig_meta[idx]
                            if idx < len(orig_meta)
                            else {"Category": "main", "Description": num.zfill(2)}
                        )
                        r_day = str(row.get("Day") or p["day"]).strip().upper()
                        r_sess = _normalize_session_code(
                            row.get("Session") or p.get("session", "1")
                        )
                        rows_fixed.append({
                            "Number": num.zfill(2),
                            "Amount": amt,
                            "Category": meta["Category"],
                            "Description": meta["Description"],
                            "Day": r_day,
                            "Session": r_sess,
                        })
                    except (ValueError, TypeError):
                        continue
                v, blk, inv, tot = process_and_add_bets(
                    rows_fixed,
                    p["date"],
                    p["day"],
                    p.get("session", "1"),
                    p["customer"],
                )
                if v:
                    st.session_state.data = pd.concat(
                        [st.session_state.data, pd.DataFrame(v)], ignore_index=True
                    )
                    msg = f"🎉 {len(v)} ကွက် / {tot:,} ကျပ် သိမ်းဆည်းပြီး"
                    if blk:
                        msg += f" | ပိတ်ဂဏန်း ဖယ်: {', '.join(sorted(set(blk)))}"
                    if inv:
                        msg += f" | မမှန်ကန်သော {inv} ကွက် ဖယ်ထားသည်"
                    st.session_state.flash_success = msg
                    st.session_state.pending_parse = None
                    st.rerun()
                else:
                    st.session_state.flash_warning = "သိမ်းဆည်းနိုင်သော ကွက် မရှိပါ။"
                    st.rerun()
        with cc:
            if st.button("❌ Cancel", use_container_width=True):
                st.session_state.pending_parse = None
                st.rerun()
    else:
        st.error("Parse မှ ရလဒ် မထွက်ပါ။ စာသားပုံစံကို စစ်ဆေးပါ။")
        if st.button("❌ Cancel"):
            st.session_state.pending_parse = None
            st.rerun()

# =========================================================
# DASHBOARD
# =========================================================
st.markdown("---")

if not st.session_state.data.empty:
    df = st.session_state.data.copy()
    if "Session" not in df.columns:
        df["Session"] = "1"
    tot = float(df["Amount"].sum())
    comm = tot * commission_pct / 100
    net = tot - comm
    win = str(winning_number).strip().zfill(2) if winning_number.strip() else ""
    matched = df[df["Number"] == win] if win else df.iloc[0:0]
    win_tot = float(matched["Amount"].sum())
    payout = win_tot * payout_rate
    pl = net - payout

    st.subheader("📊 အဓိက စာရင်းချုပ်")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("ရောင်းရငွေ", f"{tot:,.0f}")
    m2.metric(f"ကော်မရှင် {commission_pct}%", f"{comm:,.0f}")
    m3.metric("Net Sales", f"{net:,.0f}")
    m4.metric("အသားတင် P/L", f"{pl:,.0f}")

    st.markdown("---")
    st.subheader("🏆 ပေါက်ဂဏန်း တိုက်စစ်")
    if win:
        if not matched.empty:
            st.success(f"ပေါက်ဂဏန်း ({win}) ကံထူးရှင်များ:")
            st.dataframe(matched, use_container_width=True)
            cc1, cc2 = st.columns(2)
            cc1.metric("ဆုကြေး", f"{payout:,.0f}")
            cc2.metric("P/L", f"{pl:,.0f}")
        else:
            st.info(
                f"ပေါက်ဂဏန်း ({win}) ထိုးထားသူ မရှိ။ ဒိုင်အမြတ် = {net:,.0f}"
            )
    else:
        st.info("Sidebar တွင် ပေါက်ဂဏန်း ထည့်ပါ။")

    st.markdown("---")
    with st.expander("✏️ Edit / Delete Bet Records"):
        edited_ledger = st.data_editor(
            st.session_state.data,
            num_rows="dynamic",
            use_container_width=True,
            key="edit_ledger",
        )
        ce1, ce2 = st.columns(2)
        with ce1:
            if st.button("💾 Apply Changes to Ledger"):
                st.session_state.data = edited_ledger.reset_index(drop=True)
                st.session_state.flash_success = "ပြင်ဆင်မှု သိမ်းပြီး"
                st.rerun()
        with ce2:
            st.caption(f"စုစုပေါင်း Record: {len(st.session_state.data)}")

    st.markdown("---")
    st.subheader("📨 Viber Statements")
    col_rep1, col_rep2 = st.columns(2)

    with col_rep1:
        st.markdown("##### 📒 ဒိုင်ချုပ်")
        dsum = df.groupby(["Date", "Day"])["Amount"].sum().reset_index()
        st.dataframe(dsum, use_container_width=True)
        lines = ["*မြန်မာ 2D စာရင်းချုပ်အစီရင်ခံစာ*", "=" * 28]
        for r in dsum.itertuples():
            lines.append(f"{r.Day} ({r.Date}): {r.Amount:,.0f} ကျပ်")
        lines.extend([
            "-" * 28,
            f"ရောင်းရငွေ = {tot:,.0f}",
            f"ကော်မရှင် ({commission_pct}%) = -{comm:,.0f}",
            f"Net = {net:,.0f}",
        ])
        if win:
            lines.append(f"ပေါက်ဂဏန်း ({win}) = {win_tot:,.0f}")
            lines.append(f"ဆုကြေး ({payout_rate}ဆ) = -{payout:,.0f}")
        lines.append(f"အသားတင် = {pl:,.0f}")
        st.text_area("Dealer Viber Text", "\n".join(lines), height=220)

    with col_rep2:
        st.markdown("##### 👤 ဖောက်သည် ရှင်းတမ်း (အပိုင်းအလိုက် အသေးစိတ်)")
        custs = sorted(df["Customer"].unique())
        sel = st.selectbox("ဖောက်သည်", custs)
        if sel:
            cdf = df[df["Customer"] == sel]
            lines = [
                f"📊 *{sel}* ၏ စုစုပေါင်း ထိုးငွေအသေးစိတ်ရှင်းတမ်း",
                "=" * 32,
            ]
            batch_totals = []

            for b in cdf["BatchID"].unique():
                bdf = cdf[cdf["BatchID"] == b]
                bt = float(bdf["Amount"].sum())
                batch_totals.append((b, bt))
                lines.append(f"--- *{b}* ---")

                main_df = bdf[bdf["Category"] == "main"]
                if not main_df.empty:
                    lines.append("🔹 *နေ့ရှင်းနှင့် အဓိကထိုးငွေများ:*")
                    for desc_lbl, grp_d in main_df.groupby(
                        "Description", sort=False
                    ):
                        sub_amt = grp_d["Amount"].sum()
                        lines.append(f"  • {desc_lbl} = {sub_amt:,.0f} ကျပ်")

                rpat_df = bdf[bdf["Category"] == "r_pat"]
                if not rpat_df.empty:
                    lines.append("\n🔹 *ပတ်လည်နှင့် ရိုဂဏန်းများ (R & Pat):*")
                    for desc_lbl, grp_d in rpat_df.groupby(
                        "Description", sort=False
                    ):
                        sub_amt = grp_d["Amount"].sum()
                        lines.append(f"  • {desc_lbl} = {sub_amt:,.0f} ကျပ်")

                lines.append(f"\n*{b} စုစုပေါင်း* = {bt:,.0f} ကျပ်\n")

            ctot = float(cdf["Amount"].sum())
            ccomm = ctot * commission_pct / 100
            cnet = ctot - ccomm
            cwin = 0
            if win:
                cwin = (
                    float(
                        cdf[cdf["Number"] == win]["Amount"].sum() * payout_rate
                    )
                )
            bal = cnet - cwin
            formula = " + ".join([f"[{b}={a:,.0f} ကျပ်]" for b, a in batch_totals])
            lines.extend([
                "=" * 32,
                f"{formula} = *({ctot:,.0f} ကျပ်)*",
                f"ကော်မရှင် ({commission_pct}%) = -{ccomm:,.0f} ကျပ်",
                f"စုစုပေါင်း ရှင်းတမ်း (Net Total) = *{cnet:,.0f} ကျပ်*",
            ])
            if win:
                lines.append(f"ဆုကြေး ({win}) = +{cwin:,.0f} ကျပ်")
                lines.append(f"နောက်ဆုံး လက်ကျန်ရှင်းတမ်း = {bal:,.0f} ကျပ်")
            st.text_area(f"Viber ({sel})", "\n".join(lines), height=360)

    # Single-sheet Weekly Statement per Customer (Mon - Fri, 10 Sessions + Summary)
    st.markdown("---")
    st.subheader(
        "🗓️ ဖောက်သည် တစ်ပတ်စာ (၁၀ ကြိမ်စာ) စာရင်းရှင်းတမ်း (Mon - Fri | [1]မနက် - [2]ညနေ)"
    )
    cust_wk_list = sorted(df["Customer"].unique())
    sel_wk_cust = st.selectbox(
        "တစ်ပတ်စာ ရှင်းတမ်းထုတ်မည့် ဖောက်သည်",
        cust_wk_list,
        key="weekly_stmt_cust",
    )
    if sel_wk_cust:
        wk_text, wk_meta = build_customer_weekly_report(
            df, sel_wk_cust, payout_rate, winning_number, slot_win_inputs
        )
        wm1, wm2, wm3 = st.columns(3)
        wm1.metric(
            "၁၀ ကြိမ်စာ စုစုပေါင်းထိုးငွေ", f"{wk_meta['bets_total']:,.0f} ကျပ်"
        )
        wm2.metric(
            "၁၀ ကြိမ်စာ စုစုပေါင်းပေါက်ငွေ",
            f"{wk_meta['win_payout']:,.0f} ကျပ်",
        )
        wm3.metric("အရှုံး / အနိုင် ရလဒ်", wk_meta["outcome_label"])

        st.dataframe(pd.DataFrame(wk_meta["slots"]), use_container_width=True)
        st.text_area(
            f"ထိုးသူထံ ပြန်ပို့ရန် စာသား ({sel_wk_cust})",
            wk_text,
            height=440,
            key="weekly_stmt_area",
        )

    with st.expander("📋 ဖောက်သည်အားလုံး၏ တစ်ပတ်စာ (၁၀ ကြိမ်စာ) ရှင်းတမ်းစာသားများ"):
        all_wk_text = build_all_customers_weekly_report(
            df, payout_rate, winning_number, slot_win_inputs
        )
        st.text_area(
            "All Customers Weekly Statements",
            all_wk_text,
            height=460,
            key="all_weekly_stmt_area",
        )

    st.markdown("---")
    cl, cr = st.columns(2)

    with cl:
        st.subheader("🔢 ဂဏန်း Limit")
        grp = (
            df.groupby("Number")["Amount"]
            .sum()
            .reset_index()
            .sort_values("Amount", ascending=False)
        )
        grp["Allowed"] = grp["Amount"].apply(lambda x: min(x, max_limit))
        grp["Excess"] = grp["Amount"] - grp["Allowed"]
        st.dataframe(grp, use_container_width=True)

    with cr:
        st.subheader("👥 Batch အလိုက်")
        cons = (
            df.groupby(["Date", "Day", "Session", "Customer", "BatchID", "Number"])[
                "Amount"
            ]
            .sum()
            .reset_index()
        )
        st.dataframe(cons, use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Export & Management")
    ce, cx = st.columns([2, 1])

    with ce:
        xls = generate_styled_excel(
            df,
            commission_pct,
            payout_rate,
            winning_number,
            max_limit,
            slot_win_inputs,
        )
        st.download_button(
            "📥 Excel Download",
            data=xls,
            file_name=f"2D_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    with cx:
        if not st.session_state.confirm_clear:
            if st.button("🗑️ Clear All"):
                st.session_state.confirm_clear = True
                st.rerun()
        else:
            st.warning("သေချာပါသလား? ပြန်မရနိုင်ပါ။")
            c1x, c2x = st.columns(2)
            with c1x:
                if st.button("✅ Yes"):
                    st.session_state.data = pd.DataFrame(columns=COLUMNS)
                    st.session_state.confirm_clear = False
                    st.rerun()
            with c2x:
                if st.button("❌ No"):
                    st.session_state.confirm_clear = False
                    st.rerun()

else:
    st.info("စာရင်းမရှိသေးပါ။ အထက်တွင် စတင်သွင်းပါ။")
