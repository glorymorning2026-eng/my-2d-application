import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from itertools import permutations, product

# =========================================================
# APP CONFIGURATION
# =========================================================
st.set_page_config(page_title="Myanmar 2D Dealer", page_icon="📒", layout="wide")

st.markdown("""
<style>
.main-header { border-left: 6px solid #D97706; padding-left: 15px; margin-bottom: 20px; }
.stButton>button { border-radius: 8px; font-weight: bold; }
div[data-testid="stMetricValue"] { color: #D97706; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
<h1>📒 မြန်မာနှစ်လုံးထီ (2D) Dealer App</h1>
<p>Smart Parser • Preview Before Save • Warning System • Viber Statement • Excel Export</p>
</div>
""", unsafe_allow_html=True)

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များ")

max_limit = st.sidebar.number_input("Limit (Max per number)", min_value=0, value=500000, step=50000)
payout_rate = st.sidebar.number_input("ဆုကြေးနှုန်း (ဆ)", min_value=0, value=80, step=5)
commission_pct = st.sidebar.number_input("ကော်မရှင် (%)", min_value=0, max_value=100, value=10, step=1)
winning_number = st.sidebar.text_input("ထွက်မည့် ပေါက်ဂဏန်း", value="")

st.sidebar.markdown("---")
st.sidebar.subheader("⏰ Time Lock")
enable_time_lock = st.sidebar.checkbox("Enable Time Lock", value=False)
closing_time = st.sidebar.time_input("Closing Time", value=datetime.strptime("12:00", "%H:%M").time())

st.sidebar.markdown("---")
st.sidebar.subheader("🚫 Blocked Numbers")
blocked_input = st.sidebar.text_input("Blocked (comma-separated)", value="")
blocked_numbers = [
    n.strip().zfill(2) for n in blocked_input.split(",")
    if n.strip().isdigit() and 0 <= int(n.strip()) <= 99
]

# =========================================================
# SESSION STATE
# =========================================================
COLUMNS = ["Date", "Day", "Customer", "BatchID", "Number", "Amount"]

if "data" not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=COLUMNS)
if "pending_parse" not in st.session_state:
    st.session_state.pending_parse = None

# =========================================================
# PARSER CORE
# =========================================================
def add_number(rows, number, amount, is_default=False):
    number = str(number).zfill(2)
    if len(number) == 2 and number.isdigit() and 0 <= int(number) <= 99:
        rows.append({"Number": number, "Amount": int(amount), "_default": is_default})
        return True
    return False


def get_last_amount(line, default=1000):
    m = re.search(r"(\d[\d,]*)(?:tt)?\s*$", line, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1).replace(",", "")), False
        except ValueError:
            pass
    return default, True


def extract_numbers(text):
    return re.findall(r"\d{1,2}", text)


def _bk_for_digit(digit):
    out = []
    for i in range(100):
        n = f"{i:02d}"
        if (int(n[0]) + int(n[1])) % 10 == int(digit):
            out.append(n)
    return out


def _pat_for_digit(digit):
    return [f"{i:02d}" for i in range(100) if digit in f"{i:02d}"]


def parse_segment(seg, default_bk_amount=None):
    """Parse a single '/'-separated segment. Returns (rows, warnings)."""
    rows, warnings = [], []
    seg = seg.strip()
    if not seg:
        return rows, warnings

    # BK
    bk = re.fullmatch(r"(\d)\s*(ဘရိတ်|B|bk)\s*(\d+)?", seg, re.IGNORECASE)
    if bk:
        digit = bk.group(1)
        if bk.group(3):
            amount, is_default = int(bk.group(3)), False
        elif default_bk_amount is not None:
            amount, is_default = default_bk_amount, False
        else:
            amount, is_default = 1000, True
            warnings.append(f"BK {digit} amount not specified, default 1000")
        for n in _bk_for_digit(digit):
            add_number(rows, n, amount, is_default)
        return rows, warnings

    # ပါတ် / ပတ် / ပက်
    pat = re.fullmatch(r"(\d)\s*(ပါတ်|ပတ်|ပက်|ပါ|ပ)(?!ူး)\s*(\d+)?(?:tt)?", seg)
    if pat:
        digit = pat.group(1)
        if pat.group(3):
            amount, is_default = int(pat.group(3)), False
        else:
            amount, is_default = 1000, True
            warnings.append(f"ပတ် {digit} amount not specified, default 1000")
        for n in _pat_for_digit(digit):
            add_number(rows, n, amount, is_default)
        return rows, warnings

    # ထိပ်
    th = re.fullmatch(r"(\d)\s*(ထိပ်စီး|ထိပ်)\s*(\d+)?", seg)
    if th:
        digit = th.group(1)
        amount = int(th.group(3)) if th.group(3) else 1000
        is_default = not bool(th.group(3))
        for i in range(10):
            add_number(rows, f"{digit}{i}", amount, is_default)
        return rows, warnings

    # ပိတ်
    pt = re.fullmatch(r"(\d)\s*(နောက်ပိတ်|ပိတ်)\s*(\d+)?", seg)
    if pt:
        digit = pt.group(1)
        amount = int(pt.group(3)) if pt.group(3) else 1000
        is_default = not bool(pt.group(3))
        for i in range(10):
            add_number(rows, f"{i}{digit}", amount, is_default)
        return rows, warnings

    # number=amount (full match)
    eq = re.fullmatch(r"(\d{1,2})\s*[=]\s*(\d{3,6})", seg)
    if eq:
        add_number(rows, eq.group(1), int(eq.group(2)), False)
        return rows, warnings

    # number*amount (full match)
    star = re.fullmatch(r"(\d{1,2})\s*\*\s*(\d{3,6})", seg)
    if star:
        add_number(rows, star.group(1), int(star.group(2)), False)
        return rows, warnings

    # numberRamount
    r = re.fullmatch(r"(\d{1,2})\s*[Rr]\s*(\d+)", seg)
    if r:
        n = r.group(1).zfill(2)
        amt = int(r.group(2))
        add_number(rows, n, amt, False)
        if n != n[::-1]:
            add_number(rows, n[::-1], amt, False)
        return rows, warnings

    # partial num=amount + rest (non-greedy amount)
    eq_partial = re.match(r"(\d{1,2})\s*=\s*(\d{3,6}?)(.+)", seg)
    if eq_partial:
        num = eq_partial.group(1)
        amt = int(eq_partial.group(2))
        rest = eq_partial.group(3)
        rest_rows, rest_warns = parse_segment(rest, default_bk_amount)
        if rest_rows:
            add_number(rows, num, amt, False)
            rows.extend(rest_rows)
            warnings.extend(rest_warns)
            return rows, warnings

    # concatenated num+amount (e.g. "773000" -> 77, 3000)
    if re.fullmatch(r"\d{5,8}", seg):
        num = seg[:2]
        amt_str = seg[2:]
        if amt_str.isdigit():
            amt = int(amt_str)
            if 100 <= amt <= 500000:
                add_number(rows, num, amt, False)
                warnings.append(f"Ambiguous '{seg}' interpreted as {num}={amt}")
                return rows, warnings

    # pure numbers (no amount)
    if re.fullmatch(r"\d{2}(?:[\s,./-]+\d{2})*", seg):
        for n in re.findall(r"\d{2}", seg):
            add_number(rows, n, 1000, True)
        return rows, warnings

    # single 1-2 digit number
    if re.fullmatch(r"\d{1,2}", seg):
        add_number(rows, seg, 1000, True)
        return rows, warnings

    warnings.append(f"Unrecognized segment: '{seg}'")
    return rows, warnings


def is_complex_line(line):
    if "/" not in line:
        return False
    segs = [s.strip() for s in line.split("/") if s.strip()]
    if len(segs) <= 1:
        return False
    if all(re.fullmatch(r"\d\s*(bk|ဘရိတ်|B)\s*(\d+)?", s, re.I) for s in segs):
        return False
    for s in segs:
        if re.fullmatch(r"\d{1,2}", s):
            continue
        if re.fullmatch(r"\d{1,2}\s*[=*]\s*\d{3,6}", s):
            continue
        if re.fullmatch(r"\d\s*(bk|ဘရိတ်|B)\s*(\d+)?", s, re.I):
            continue
        if re.fullmatch(r"\d\s*(ပါတ်|ပတ်|ပက်|ပါ|ပ)(?!ူး)(\d+)?(?:tt)?", s):
            continue
        if re.fullmatch(r"\d\s*(ထိပ်စီး|ထိပ်|နောက်ပိတ်|ပိတ်)(\d+)?", s):
            continue
        if re.fullmatch(r"\d{1,2}\s*[Rr]\s*\d+", s):
            continue
        return True
    return False


def parse_complex_line(line):
    segs = [s.strip() for s in line.split("/") if s.strip()]
    bk_amts = []
    for s in segs:
        m = re.fullmatch(r"\d\s*(bk|ဘရိတ်|B)\s*(\d+)", s, re.I)
        if m:
            bk_amts.append(int(m.group(2)))
    default_bk = bk_amts[0] if bk_amts else None
    rows, warnings = [], []
    for s in segs:
        r, w = parse_segment(s, default_bk_amount=default_bk)
        rows.extend(r)
        warnings.extend(w)
    return rows, warnings


def parse_single_line(line):
    """Try full-line patterns. Returns (rows, warnings)."""
    rows, warnings = [], []

    if re.search(r"ညီအကို|ညီကို", line):
        amount, is_def = get_last_amount(line, 1000)
        pairs = ["01", "12", "23", "34", "45", "56", "67", "78", "89", "90"]
        added = set()
        for n in pairs:
            for item in (n, n[::-1]):
                if item not in added:
                    added.add(item)
                    add_number(rows, item, amount, is_def)
        return rows, warnings

    if re.search(r"အပူ|ပူး", line) and not re.search(r"ခွေ", line):
        amount, is_def = get_last_amount(line, 1000)
        for i in range(10):
            add_number(rows, f"{i}{i}", amount, is_def)
        return rows, warnings

    if "နက္ခတ်" in line:
        amount, is_def = get_last_amount(line, 1000)
        for n in ["07", "70", "18", "81", "24", "42", "35", "53", "69", "96"]:
            add_number(rows, n, amount, is_def)
        return rows, warnings

    if re.search(r"ပါဝါ|power", line, re.I):
        amount, is_def = get_last_amount(line, 1000)
        for n in ["05", "50", "16", "61", "27", "72", "38", "83", "49", "94"]:
            add_number(rows, n, amount, is_def)
        return rows, warnings

    # BK
    bks = re.findall(r"(\d)\s*(ဘရိတ်|B|bk)", line, re.IGNORECASE)
    if bks:
        amount, is_def = get_last_amount(line, 1000)
        for b in bks:
            for n in _bk_for_digit(b[0]):
                add_number(rows, n, amount, is_def)
        return rows, warnings

    # ပါတ်
    pa = re.search(r"(\d)\s*(ပါတ်|ပတ်|ပက်|ပါ|ပ)(?!ူး)", line)
    if pa:
        digit = pa.group(1)
        amount, is_def = get_last_amount(line, 1000)
        for i in range(100):
            n = f"{i:02d}"
            if digit in n:
                add_number(rows, n, amount, is_def)
        return rows, warnings

    # ထိပ်
    th = re.search(r"(\d)\s*(ထိပ်စီး|ထိပ်)", line)
    if th:
        digit = th.group(1)
        amount, is_def = get_last_amount(line, 1000)
        for i in range(10):
            add_number(rows, f"{digit}{i}", amount, is_def)
        return rows, warnings

    # ပိတ်
    pt = re.search(r"(\d)\s*(နောက်ပိတ်|ပိတ်)", line)
    if pt:
        digit = pt.group(1)
        amount, is_def = get_last_amount(line, 1000)
        for i in range(10):
            add_number(rows, f"{i}{digit}", amount, is_def)
        return rows, warnings

    # ခွေ
    kh = re.search(r"(\d{2,5})\s*(ခွေ|အခွေ)", line)
    if kh:
        digits = kh.group(1)
        amount, is_def = get_last_amount(line, 1000)
        inc_dbl = bool(re.search(r"အပူးပါ|ပူးပါ|အပူးအပါ", line))
        added = set()
        combos = product(set(digits), repeat=2) if inc_dbl else permutations(digits, 2)
        for pair in combos:
            n = "".join(pair)
            if n not in added:
                added.add(n)
                add_number(rows, n, amount, is_def)
        return rows, warnings

    # num_amount@rev_amount
    at2 = re.fullmatch(r"(\d{1,2})[_.](\d+)@(\d+)", line)
    if at2:
        n = at2.group(1).zfill(2)
        add_number(rows, n, int(at2.group(2)), False)
        if n != n[::-1]:
            add_number(rows, n[::-1], int(at2.group(3)), False)
        return rows, warnings

    # num@amount
    at1 = re.fullmatch(r"(\d{1,2})@(\d+)", line)
    if at1:
        n = at1.group(1).zfill(2)
        amt = int(at1.group(2))
        add_number(rows, n, amt, False)
        if n != n[::-1]:
            add_number(rows, n[::-1], amt, False)
        return rows, warnings

    # complex R
    cr = re.fullmatch(r"([\d./,\-\s]+?)[\-=\s](\d{3,6})\s*([Rrရ@&အာ])\s*(\d+)", line)
    if cr:
        nums = extract_numbers(cr.group(1))
        base = int(cr.group(2))
        rev = int(cr.group(4))
        for item in nums:
            n = item.zfill(2)
            add_number(rows, n, base, False)
            if n != n[::-1]:
                add_number(rows, n[::-1], rev, False)
        return rows, warnings

    # multi R
    mr = re.fullmatch(r"([\d./,\-\s]+?)\s*[Rr]\s*(\d+)", line)
    if mr:
        amt = int(mr.group(2))
        for item in extract_numbers(mr.group(1)):
            n = item.zfill(2)
            add_number(rows, n, amt, False)
            if n != n[::-1]:
                add_number(rows, n[::-1], amt, False)
        return rows, warnings

    # dot R
    dr = re.fullmatch(r"([\d./,\-\s]+?)\.?\s*[Rr]\.?\s*(\d+)", line)
    if dr:
        rev = int(dr.group(2))
        tokens = [t for t in re.split(r"[./,\-\s]+", dr.group(1)) if t.isdigit()]
        nums, base = [], None
        for t in tokens:
            if len(t) <= 2:
                nums.append(t.zfill(2))
            else:
                base = int(t)
        if base is None:
            base = rev
        for n in nums:
            add_number(rows, n, base, False)
            if n != n[::-1]:
                add_number(rows, n[::-1], rev, False)
        return rows, warnings

    # slash/dash amount
    sa = re.fullmatch(r"([\d./,\-\s]+?)\s*[\*=\s]\s*(\d{3,6})", line)
    if sa:
        amt = int(sa.group(2))
        for item in extract_numbers(sa.group(1)):
            add_number(rows, item, amt, False)
        return rows, warnings

    # multi amount
    ma = re.fullmatch(r"([\d./,\-\s]+?)\s*[./,\-\s]\s*(\d{3,6})", line)
    if ma:
        amt = int(ma.group(2))
        for item in extract_numbers(ma.group(1)):
            add_number(rows, item, amt, False)
        return rows, warnings

    # single R
    sr = re.fullmatch(r"(\d{1,2})\s*[Rr]\s*(\d+)", line)
    if sr:
        n = sr.group(1).zfill(2)
        amt = int(sr.group(2))
        add_number(rows, n, amt, False)
        if n != n[::-1]:
            add_number(rows, n[::-1], amt, False)
        return rows, warnings

    # single amount
    sn = re.fullmatch(r"(\d{1,2})[\s.\-](\d{3,6})", line)
    if sn:
        add_number(rows, sn.group(1), int(sn.group(2)), False)
        return rows, warnings

    # concatenated num+amount (like "772500")
    if re.fullmatch(r"\d{5,8}", line):
        num = line[:2]
        amt_str = line[2:]
        if amt_str.isdigit():
            amt = int(amt_str)
            if 100 <= amt <= 500000:
                add_number(rows, num, amt, False)
                warnings.append(f"Ambiguous '{line}' interpreted as {num}={amt}")
                return rows, warnings

    # numbers only
    if re.fullmatch(r"\d{2}(?:[\s,./-]+\d{2})*", line) and not re.search(r"[A-Za-z]", line):
        for item in re.findall(r"\d{2}", line):
            add_number(rows, item, 1000, True)
        return rows, warnings

    # fallback pairs
    pairs = re.findall(r"(\d{1,2})\s*[-=/_\.\s]\s*(\d{3,6})", line)
    if pairs:
        for num, amt in pairs:
            add_number(rows, num, int(amt), False)
        return rows, warnings

    warnings.append(f"Unrecognized line: '{line}'")
    return rows, warnings


def parse_bets_locally(text):
    """Main parser entry. Returns dict with rows, warnings, errors."""
    warnings, errors = [], []
    all_rows = []
    pending = []  # row refs affected by upcoming နေ့ရှင်း

    lines = (
        text.replace("၊", "\n").replace(";", "\n").replace("\r", "\n").split("\n")
    )

    for ln, raw in enumerate(lines, 1):
        line = raw.strip()
        if not line:
            continue

        # နေ့ရှင်း
        m = re.search(r"([\d,]+)\s*နေ့ရှင်း", line)
        if m:
            amt = int(m.group(1).replace(",", ""))
            if not pending:
                warnings.append(f"Line {ln}: နေ့ရှင်း without preceding numbers")
            else:
                for r in pending:
                    r["Amount"] = amt
                    r["_default"] = False
                warnings.append(
                    f"Line {ln}: နေ့ရှင်း {amt:,} applied to {len(pending)} numbers"
                )
                pending.clear()
            continue

        try:
            if is_complex_line(line):
                new_rows, new_warns = parse_complex_line(line)
            else:
                new_rows, new_warns = parse_single_line(line)

            for r in new_rows:
                if r.get("_default"):
                    pending.append(r)
            all_rows.extend(new_rows)
            for w in new_warns:
                warnings.append(f"Line {ln}: {w}")
        except Exception as e:
            errors.append(f"Line {ln}: Parse error: {e}")

    for r in all_rows:
        r.pop("_default", None)

    return {"rows": all_rows, "warnings": warnings, "errors": errors}


# =========================================================
# PROCESS AND ADD TO LEDGER
# =========================================================
def process_and_add_bets(extracted_data, entry_date_str, selected_day, customer):
    valid_rows, blocked_found = [], []
    invalid_amt, total_added = 0, 0

    existing = st.session_state.data[st.session_state.data["Customer"] == customer]["BatchID"]
    next_num = 1
    if not existing.empty:
        nums = []
        for b in existing.unique():
            m = re.search(r"\((\d+)\)", str(b))
            if m:
                nums.append(int(m.group(1)))
        if nums:
            next_num = max(nums) + 1
    batch_id = f"{customer}({next_num})"

    for row in extracted_data:
        try:
            amount = int(row.get("Amount", 0))
            number = str(row.get("Number", "")).zfill(2)
        except (ValueError, TypeError):
            invalid_amt += 1
            continue
        if len(number) != 2 or not number.isdigit():
            invalid_amt += 1
            continue
        if amount < 100 or amount > 500000:
            invalid_amt += 1
            continue
        if number in blocked_numbers:
            blocked_found.append(number)
            continue
        valid_rows.append({
            "Date": entry_date_str, "Day": selected_day,
            "Customer": customer, "BatchID": batch_id,
            "Number": number, "Amount": amount,
        })
        total_added += amount

    return valid_rows, blocked_found, invalid_amt, total_added


# =========================================================
# EXCEL EXPORT
# =========================================================
def generate_styled_excel(df, commission_pct, payout_rate, winning_number, max_limit):
    out = io.BytesIO()
    wb = openpyxl.Workbook()

    header_fill = PatternFill("solid", start_color="1E293B", end_color="1E293B")
    accent_fill = PatternFill("solid", start_color="FEF3C7", end_color="FEF3C7")
    zebra_fill = PatternFill("solid", start_color="F8FAFC", end_color="F8FAFC")
    profit_fill = PatternFill("solid", start_color="DCFCE7", end_color="DCFCE7")
    loss_fill = PatternFill("solid", start_color="FEE2E2", end_color="FEE2E2")

    f_title = Font(name="Pyidaungsu", size=14, bold=True, color="1E293B")
    f_header = Font(name="Pyidaungsu", size=11, bold=True, color="FFFFFF")
    f_body = Font(name="Pyidaungsu", size=11, color="0F172A")
    f_bold = Font(name="Pyidaungsu", size=11, bold=True, color="0F172A")
    border = Border(
        left=Side("thin", color="CBD5E1"), right=Side("thin", color="CBD5E1"),
        top=Side("thin", color="CBD5E1"), bottom=Side("thin", color="CBD5E1"),
    )

    def fmt_sheet(ws, widths):
        ws.views.sheetView[0].showGridLines = True
        for c, w in widths.items():
            ws.column_dimensions[c].width = w
        ws.freeze_panes = "A4"

    def style_hdr(ws, row, headers):
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=row, column=i, value=h)
            cell.fill, cell.font = header_fill, f_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

    win_num = str(winning_number).strip().zfill(2) if winning_number else ""
    total_sales = float(df["Amount"].sum())
    comm_amt = total_sales * commission_pct / 100
    net_sales = total_sales - comm_amt
    win_total = float(df[df["Number"] == win_num]["Amount"].sum()) if win_num else 0
    payout = win_total * payout_rate
    net_pl = net_sales - payout

    # Sheet 1
    ws = wb.active
    ws.title = "စာရင်းချုပ်အကျဉ်း"
    ws["A1"] = "မြန်မာနှစ်လုံးထီ (2D) ငွေစာရင်းချုပ်"
    ws["A1"].font = f_title
    ws["A2"] = f"ထုတ်ပြန်သည့်ရက်စွဲ - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A2"].font = f_body

    summary = [
        ("အကြောင်းအရာ", "ပမာဏ", "မှတ်ချက်"),
        ("စုစုပေါင်း ရောင်းရငွေ", total_sales, f"အကွက်ရေ {len(df):,}"),
        (f"ကော်မရှင် ({commission_pct}%)", comm_amt, ""),
        ("ဒိုင်ရငွေ (Net)", "=B5-B6", ""),
        ("ထွက်ဂဏန်း", win_num or "N/A", f"{payout_rate} ဆ"),
        ("ပေါက်ဂဏန်းထိုးငွေ", win_total, ""),
        ("ပေးရမည့်ဆုကြေး", f"=B9*{payout_rate}", ""),
        ("အသားတင် အနိုင်/အရှုံး", "=B7-B10", "(+) / (-)"),
    ]
    for i, vals in enumerate(summary):
        r = i + 4
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v)
            c.border = border
            if i == 0:
                c.fill, c.font = header_fill, f_header
                c.alignment = Alignment(horizontal="center")
            else:
                c.font = f_body
                if i == 3:
                    c.fill = accent_fill
                if i == 7:
                    c.fill = profit_fill if net_pl >= 0 else loss_fill
                if j == 2 and i != 4:
                    c.number_format = "#,##0"
                    c.alignment = Alignment(horizontal="right")
    fmt_sheet(ws, {"A": 42, "B": 24, "C": 38})

    # Sheet 2
    ws = wb.create_sheet("နေ့အလိုက်_အပတ်စဉ်")
    ws["A1"] = "နေ့အလိုက် စာရင်းချုပ်"
    ws["A1"].font = f_title
    dg = df.groupby(["Date", "Day"]).agg(
        Bets=("Number", "count"), Sales=("Amount", "sum")
    ).reset_index()
    style_hdr(ws, 3, ["ရက်စွဲ", "နေ့", "အကွက်ရေ", "ထိုးကြေး", f"ကော်မရှင် ({commission_pct}%)", "ဒိုင်ရငွေ"])
    for r, row in enumerate(dg.itertuples(), 4):
        vals = [row.Date, str(row.Day), int(row.Bets), float(row.Sales),
                f"=D{r}*{commission_pct/100}", f"=D{r}-E{r}"]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, border
            if ci >= 3:
                c.number_format = "#,##0"
            if r % 2:
                c.fill = zebra_fill
    fmt_sheet(ws, {"A": 18, "B": 16, "C": 14, "D": 22, "E": 20, "F": 22})

    # Sheet 3
    ws = wb.create_sheet("ဖောက်သည်အလိုက်")
    ws["A1"] = "ဖောက်သည်အလိုက် ရှင်းတမ်း"
    ws["A1"].font = f_title
    cg = df.groupby("Customer")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
    style_hdr(ws, 3, ["ဖောက်သည်", "ထိုးကြေး", f"ကော်မရှင် ({commission_pct}%)", "ဒိုင်ရငွေ", f"ဆုကြေး ({payout_rate}ဆ)", "အသားတင်"])
    for r, row in enumerate(cg.itertuples(), 4):
        cwin = 0
        if win_num:
            cwin = float(df[(df["Customer"] == row.Customer) & (df["Number"] == win_num)]["Amount"].sum()) * payout_rate
        vals = [row.Customer, float(row.Amount), f"=B{r}*{commission_pct/100}",
                f"=B{r}-C{r}", cwin, f"=D{r}-E{r}"]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, border
            if ci > 1:
                c.number_format = "#,##0"
            if r % 2:
                c.fill = zebra_fill
    fmt_sheet(ws, {"A": 26, "B": 22, "C": 20, "D": 24, "E": 24, "F": 24})

    # Sheet 4
    ws = wb.create_sheet("ဂဏန်း_Limit")
    ws["A1"] = "ဂဏန်းအလိုက် Limit"
    ws["A1"].font = f_title
    ng = df.groupby("Number")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
    style_hdr(ws, 3, ["2D", "စုစုပေါင်း", "ဒိုင်လက်ခံ", "ကျော်လွန်"])
    for r, row in enumerate(ng.itertuples(), 4):
        nc = ws.cell(row=r, column=1, value=str(row.Number).zfill(2))
        nc.number_format, nc.alignment = "@", Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=float(row.Amount)).number_format = "#,##0"
        ws.cell(row=r, column=3, value=f"=MIN(B{r},{max_limit})").number_format = "#,##0"
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = "#,##0"
        for ci in range(1, 5):
            ws.cell(row=r, column=ci).font = f_body
            ws.cell(row=r, column=ci).border = border
        if float(row.Amount) > max_limit:
            ws.cell(row=r, column=4).fill = loss_fill
            ws.cell(row=r, column=4).font = f_bold
    fmt_sheet(ws, {"A": 16, "B": 24, "C": 24, "D": 26})

    # Sheet 5
    ws = wb.create_sheet("အသေးစိတ်")
    ws["A1"] = "အသေးစိတ် မှတ်တမ်း"
    ws["A1"].font = f_title
    dg = df.groupby(["Date", "Day", "Customer", "BatchID", "Number"])["Amount"].sum().reset_index()
    style_hdr(ws, 3, ["ရက်စွဲ", "နေ့", "ဖောက်သည်", "Batch", "2D", "ထိုးငွေ"])
    for r, row in enumerate(dg.itertuples(), 4):
        vals = [row.Date, row.Day, row.Customer, row.BatchID, str(row.Number).zfill(2), float(row.Amount)]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, border
            if ci == 5:
                c.number_format = "@"
            if ci == 6:
                c.number_format = "#,##0"
            if r % 2:
                c.fill = zebra_fill
    fmt_sheet(ws, {"A": 16, "B": 16, "C": 22, "D": 22, "E": 16, "F": 22})

    wb.save(out)
    return out.getvalue()


# =========================================================
# PARSER TEST PANEL
# =========================================================
with st.sidebar.expander("🧪 Parser Tests"):
    if st.button("Run Parser Tests"):
        tests = [
            ("T1 နေ့ရှင်း", "44 99 11\n11000နေ့ရှင်း", 33000),
            ("T2 slash *", "11/99/77*3000", 9000),
            ("T3 r-reverse", "46/96r1500", 6000),
            ("T4 bk multiple", "5bk/7bk250", 5500),
            ("T5 ပက်", "7ပက်2500", 47500),
            ("T6 multi R", "30-35-50-R1500", 9000),
            ("T7 single *", "77*2500", 2500),
            ("T9 leading zero", "01/02/03*1000", 3000),
            ("T10 duplicate", "11*1000\n11*1000", 2000),
        ]
        for name, text, exp in tests:
            res = parse_bets_locally(text)
            actual = sum(r["Amount"] for r in res["rows"])
            mark = "✅" if actual == exp else "❌"
            st.write(f"{mark} {name}: got `{actual:,}` (exp `{exp:,}`)")

# =========================================================
# ENTRY SECTION
# =========================================================
st.subheader("📝 စာရင်းသွင်းရန်")

is_closed = False
if enable_time_lock and datetime.now().time() >= closing_time:
    is_closed = True
    st.error(f"⚠️ စာရင်းပိတ်ချိန် ({closing_time.strftime('%H:%M')}) ကျော်လွန်သွားပါပြီ။")

col_d1, col_d2 = st.columns(2)
with col_d1:
    entry_date = st.date_input("📆 ရက်စွဲ", value=datetime.now().date())
with col_d2:
    day_options = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "WEEKEND"]
    default_day = min(entry_date.weekday(), 5)
    selected_day = st.selectbox("📅 နေ့", day_options, index=default_day)
entry_date_str = entry_date.strftime("%Y-%m-%d")

col1, col2 = st.columns([1, 2])
with col1:
    cust_manual = st.text_input("ဖောက်သည် အမည်", placeholder="ဥပမာ - သိမ့်")
with col2:
    smart_input_text = st.text_area(
        "ဖောက်သည် ပို့လိုက်သော စာသား",
        placeholder="44 99 11\n11000နေ့ရှင်း\n11/99/77*3000\n46/96r1500\n5bk/7bk250\n7ပက်2500",
        height=180,
    )

if st.button("🔍 Parse & Preview", type="primary"):
    if is_closed:
        st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပါသည်။")
    elif not cust_manual.strip() or not smart_input_text.strip():
        st.warning("ဖောက်သည်အမည်နှင့် စာသားကို ဖြည့်ပါ။")
    else:
        res = parse_bets_locally(smart_input_text)
        st.session_state.pending_parse = {
            "customer": cust_manual.strip(),
            "date": entry_date_str,
            "day": selected_day,
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
    st.subheader(f"👁️ Preview — {p['customer']} ({p['date']} / {p['day']})")

    for e in p["errors"]:
        st.error(e)
    for w in p["warnings"]:
        st.warning(w)

    if p["rows"]:
        df_prev = pd.DataFrame(p["rows"])
        df_prev = df_prev[["Number", "Amount"]]
        st.caption(f"စုစုပေါင်း {len(df_prev)} ကွက် | {df_prev['Amount'].sum():,.0f} ကျပ်")

        edited = st.data_editor(
            df_prev,
            column_config={
                "Number": st.column_config.TextColumn("2D", max_chars=2),
                "Amount": st.column_config.NumberColumn(
                    "Amount", min_value=100, max_value=500000, step=100
                ),
            },
            hide_index=True, use_container_width=True, key="parse_preview",
        )

        c_save, c_cancel = st.columns(2)
        with c_save:
            if st.button("✅ Save to Ledger", type="primary"):
                edited_rows = []
                for _, row in edited.iterrows():
                    edited_rows.append({
                        "Number": str(row["Number"]).zfill(2),
                        "Amount": int(row["Amount"]),
                    })
                vrows, blk, inv, tot = process_and_add_bets(
                    edited_rows, p["date"], p["day"], p["customer"]
                )
                if vrows:
                    st.session_state.data = pd.concat(
                        [st.session_state.data, pd.DataFrame(vrows)], ignore_index=True
                    )
                    st.success(f"🎉 {len(vrows)} ကွက် / {tot:,} ကျပ် သိမ်းဆည်းပြီး")
                    if blk:
                        st.warning("ပိတ်ဂဏန်း ဖယ်ထုတ်: " + ", ".join(sorted(set(blk))))
                    if inv:
                        st.warning(f"မမှန်ကန်သော {inv} ကွက် ဖယ်ထုတ်ထားသည်။")
                    st.session_state.pending_parse = None
                    st.rerun()
                else:
                    st.warning("သိမ်းဆည်းနိုင်သော အကွက် မရှိပါ။")
        with c_cancel:
            if st.button("❌ Cancel"):
                st.session_state.pending_parse = None
                st.rerun()

# =========================================================
# DASHBOARD
# =========================================================
st.markdown("---")

if not st.session_state.data.empty:
    df = st.session_state.data.copy()
    total_sales = float(df["Amount"].sum())
    comm_amt = total_sales * commission_pct / 100
    net_sales = total_sales - comm_amt
    win_num = str(winning_number).strip().zfill(2) if winning_number.strip() else ""
    matched = df[df["Number"] == win_num] if win_num else df.iloc[0:0]
    win_total = float(matched["Amount"].sum())
    payout = win_total * payout_rate
    net_pl = net_sales - payout

    st.subheader("📊 အဓိက စာရင်းချုပ်")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("ရောင်းရငွေ", f"{total_sales:,.0f}")
    m2.metric(f"ကော်မရှင် {commission_pct}%", f"{comm_amt:,.0f}")
    m3.metric("Net Sales", f"{net_sales:,.0f}")
    m4.metric("အသားတင်", f"{net_pl:,.0f}")

    st.markdown("---")
    st.subheader("🏆 ပေါက်ဂဏန်း တိုက်စစ်ခြင်း")
    if win_num:
        if not matched.empty:
            st.success(f"ပေါက်ဂဏန်း ({win_num}) ကံထူးရှင်များ:")
            st.dataframe(matched, use_container_width=True)
            c1, c2 = st.columns(2)
            c1.metric("ဆုကြေးစုစုပေါင်း", f"{payout:,.0f}")
            c2.metric("အသားတင် P/L", f"{net_pl:,.0f}")
        else:
            st.info(f"ပေါက်ဂဏန်း ({win_num}) ထိုးထားသူ မရှိ။ ဒိုင်အမြတ် = {net_sales:,.0f}")
