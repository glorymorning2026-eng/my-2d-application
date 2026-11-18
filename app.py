import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from itertools import permutations, product

# =========================================================
# CONFIG
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
<h1>📒 မြန်မာ 2D Dealer App</h1>
<p>Smart Parser • Preview & Edit • Warning System • Viber Statement • Excel Export</p>
</div>
""", unsafe_allow_html=True)

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များ")
max_limit = st.sidebar.number_input("Limit (Max/number)", min_value=0, value=500000, step=50000)
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

for k, v in [
    ("data", pd.DataFrame(columns=COLUMNS)),
    ("pending_parse", None),
    ("confirm_clear", False),
    ("flash_success", None),
    ("flash_warning", None),
]:
    if k not in st.session_state:
        st.session_state[k] = v

# Flash messages (survive rerun)
if st.session_state.flash_success:
    st.success(st.session_state.flash_success)
    st.session_state.flash_success = None
if st.session_state.flash_warning:
    st.warning(st.session_state.flash_warning)
    st.session_state.flash_warning = None

# =========================================================
# PARSER HELPERS
# =========================================================
def add_number(rows, number, amount, is_default=False):
    n = str(number).zfill(2)
    if len(n) == 2 and n.isdigit() and 0 <= int(n) <= 99:
        rows.append({"Number": n, "Amount": int(amount), "_default": is_default})
        return True
    return False


def _last_amount(line, default=1000):
    m = re.search(r"(\d[\d,]*)(?:tt)?\s*$", line, re.IGNORECASE)
    return int(m.group(1).replace(",", "")) if m else default


def extract_numbers(text):
    return re.findall(r"\d{1,2}", text)


def _bk_for_digit(digit):
    d = int(digit)
    return [f"{i:02d}" for i in range(100) if (i // 10 + i % 10) % 10 == d]


def _pat_for_digit(digit):
    return [f"{i:02d}" for i in range(100) if digit in f"{i:02d}"]

# =========================================================
# SEGMENT PARSER (recursive, handles / and space splits)
# =========================================================
def parse_segment(seg, default_bk=None, depth=0):
    """Parse one segment. Returns (rows, warns, matched)."""
    rows, warns = [], []
    seg = seg.strip().strip("*/").strip()
    if not seg or depth > 6:
        return rows, warns, False

    # BK
    bk = re.fullmatch(r"(\d)\s*(bk|ဘရိတ်|B)(?![a-zA-Z])\s*(\d+)?(?:tt)?", seg, re.I)
    if bk:
        amt = int(bk.group(3)) if bk.group(3) else (default_bk if default_bk else 1000)
        for n in _bk_for_digit(bk.group(1)):
            add_number(rows, n, amt)
        return rows, warns, True

    # ပါတ်
    pat = re.fullmatch(r"(\d)\s*(ပါတ်|ပတ်|ပက်|ပါ|ပ)(?!ူး)\s*(\d+)?(?:tt)?", seg)
    if pat:
        amt = int(pat.group(3)) if pat.group(3) else 1000
        for n in _pat_for_digit(pat.group(1)):
            add_number(rows, n, amt)
        return rows, warns, True

    # ထိပ်
    th = re.fullmatch(r"(\d)\s*(ထိပ်စီး|ထိပ်)\s*(\d+)?(?:tt)?", seg)
    if th:
        amt = int(th.group(3)) if th.group(3) else 1000
        for i in range(10):
            add_number(rows, f"{th.group(1)}{i}", amt)
        return rows, warns, True

    # ပိတ်
    pt = re.fullmatch(r"(\d)\s*(နောက်ပိတ်|ပိတ်)\s*(\d+)?(?:tt)?", seg)
    if pt:
        amt = int(pt.group(3)) if pt.group(3) else 1000
        for i in range(10):
            add_number(rows, f"{i}{pt.group(1)}", amt)
        return rows, warns, True

    # num=amount (full)
    eq = re.fullmatch(r"(\d{1,2})\s*=\s*(\d{3,6})", seg)
    if eq:
        add_number(rows, eq.group(1), int(eq.group(2)))
        return rows, warns, True

    # num*amount (full)
    st_ = re.fullmatch(r"(\d{1,2})\s*\*\s*(\d{3,6})", seg)
    if st_:
        add_number(rows, st_.group(1), int(st_.group(2)))
        return rows, warns, True

    # numRamount
    r = re.fullmatch(r"(\d{1,2})\s*[Rr]\s*(\d+)", seg)
    if r:
        n, amt = r.group(1).zfill(2), int(r.group(2))
        add_number(rows, n, amt)
        if n != n[::-1]:
            add_number(rows, n[::-1], amt)
        return rows, warns, True

    # num=amount + rest (greedy amount → correct)
    pm = re.fullmatch(r"(\d{1,2})\s*=\s*(\d{3,6})(.+)", seg)
    if pm:
        add_number(rows, pm.group(1), int(pm.group(2)))
        r2, w2, _ = parse_segment(pm.group(3), default_bk, depth + 1)
        rows.extend(r2)
        warns.extend(w2)
        return rows, warns, True

    # concatenated num+amount
    if re.fullmatch(r"\d{5,8}", seg):
        num, amt = seg[:2], int(seg[2:])
        if 100 <= amt <= 500000:
            add_number(rows, num, amt)
            warns.append(f"Ambiguous '{seg}' → {num}={amt}")
            return rows, warns, True

    # pure 2D
    if re.fullmatch(r"\d{1,2}", seg):
        add_number(rows, seg, 1000, is_default=True)
        return rows, warns, True

    # Split by / or whitespace (recursive)
    if "/" in seg or " " in seg:
        parts = [p.strip() for p in re.split(r"[/\s]+", seg) if p.strip()]
        if len(parts) > 1:
            for part in parts:
                r3, w3, m3 = parse_segment(part, default_bk, depth + 1)
                if m3:
                    rows.extend(r3)
                    warns.extend(w3)
                else:
                    warns.append(f"Unrecognized: '{part}'")
            if rows:
                return rows, warns, True

    return rows, warns, False

# =========================================================
# LINE PARSERS
# =========================================================
def is_complex_line(line):
    if "/" not in line:
        return False
    segs = [s.strip() for s in line.split("/") if s.strip()]
    if len(segs) <= 1:
        return False
    if all(re.fullmatch(r"\d{1,2}", s) for s in segs):
        return False
    return True


def parse_complex_line(line):
    segs = [s.strip() for s in line.split("/") if s.strip()]
    bk_amts = []
    for s in segs:
        m = re.fullmatch(r"\d\s*(bk|ဘရိတ်|B)\s*(\d+)", s, re.I)
        if m:
            bk_amts.append(int(m.group(2)))
    default_bk = bk_amts[0] if bk_amts else None
    rows, warns = [], []
    if len(set(bk_amts)) > 1:
        warns.append(f"Multiple BK amounts {bk_amts}, using {default_bk} as default")
    for s in segs:
        r, w, matched = parse_segment(s, default_bk)
        if not matched:
            warns.append(f"Unrecognized segment: '{s}'")
        rows.extend(r)
        warns.extend(w)
    return rows, warns


def parse_single_line(line):
    rows, warns = [], []

    # Whole-line keyword groups
    if re.search(r"ညီအကို|ညီကို", line):
        amt = _last_amount(line, 1000)
        pairs = ["01", "12", "23", "34", "45", "56", "67", "78", "89", "90"]
        added = set()
        for n in pairs:
            for item in (n, n[::-1]):
                if item not in added:
                    added.add(item)
                    add_number(rows, item, amt)
        return rows, warns

    if re.search(r"အပူ|ပူး", line) and "ခွေ" not in line:
        amt = _last_amount(line, 1000)
        for i in range(10):
            add_number(rows, f"{i}{i}", amt)
        return rows, warns

    if "နက္ခတ်" in line:
        amt = _last_amount(line, 1000)
        for n in ["07", "70", "18", "81", "24", "42", "35", "53", "69", "96"]:
            add_number(rows, n, amt)
        return rows, warns

    if re.search(r"ပါဝါ|power", line, re.I):
        amt = _last_amount(line, 1000)
        for n in ["05", "50", "16", "61", "27", "72", "38", "83", "49", "94"]:
            add_number(rows, n, amt)
        return rows, warns

    # ခွေ
    kh = re.search(r"(\d{2,5})\s*(ခွေ|အခွေ)", line)
    if kh:
        digits = kh.group(1)
        amt = _last_amount(line, 1000)
        inc = bool(re.search(r"အပူးပါ|ပူးပါ|အပူးအပါ", line))
        combos = product(set(digits), repeat=2) if inc else permutations(digits, 2)
        added = set()
        for p in combos:
            n = "".join(p)
            if n not in added:
                added.add(n)
                add_number(rows, n, amt)
        return rows, warns

    # Delegate to segment parser
    r, w, matched = parse_segment(line)
    if matched:
        return r, w

    warns.append(f"Unrecognized: '{line}'")
    return rows, warns

# =========================================================
# MAIN PARSER
# =========================================================
def parse_bets_locally(text):
    rows, warnings, errors = [], [], []
    pending = []

    lines = text.replace("၊", "\n").replace(";", "\n").replace("\r", "\n").split("\n")

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
                warnings.append(f"Line {ln}: နေ့ရှင်း {amt:,} → {len(pending)} number(s)")
                pending.clear()
            continue

        # List-with-trailing-amount pre-check (preserves T2: 11/99/77*3000)
        list_amt = re.fullmatch(
            r"((?:\d{1,2}[\s/,\-\.]+)*\d{1,2})\s*[\*=]?\s*(\d{3,6})", line
        )
        if list_amt:
            amt = int(list_amt.group(2))
            for n in re.findall(r"\d{1,2}", list_amt.group(1)):
                add_number(rows, n, amt)
            continue

        try:
            if is_complex_line(line):
                new_rows, new_warns = parse_complex_line(line)
            else:
                new_rows, new_warns = parse_single_line(line)
            for r in new_rows:
                if r.get("_default"):
                    pending.append(r)
            rows.extend(new_rows)
            warnings.extend([f"Line {ln}: {w}" for w in new_warns])
        except Exception as e:
            errors.append(f"Line {ln}: {e}")

    for r in rows:
        r.pop("_default", None)

    return {"rows": rows, "warnings": warnings, "errors": errors}

# =========================================================
# PROCESS & ADD
# =========================================================
def process_and_add_bets(extracted, date_str, day, customer):
    valid, blocked = [], []
    invalid, total = 0, 0

    existing = st.session_state.data[st.session_state.data["Customer"] == customer]["BatchID"]
    next_n = 1
    if not existing.empty:
        nums = [int(m.group(1)) for b in existing.unique()
                if (m := re.search(r"\((\d+)\)", str(b)))]
        if nums:
            next_n = max(nums) + 1
    batch_id = f"{customer}({next_n})"

    for row in extracted:
        try:
            amt = int(row.get("Amount", 0))
            num = str(row.get("Number", "")).zfill(2)
        except (ValueError, TypeError):
            invalid += 1
            continue
        if len(num) != 2 or not num.isdigit():
            invalid += 1
            continue
        if amt < 100 or amt > 500000:
            invalid += 1
            continue
        if num in blocked_numbers:
            blocked.append(num)
            continue
        valid.append({"Date": date_str, "Day": day, "Customer": customer,
                      "BatchID": batch_id, "Number": num, "Amount": amt})
        total += amt

    return valid, blocked, invalid, total

# =========================================================
# EXCEL EXPORT
# =========================================================
def generate_styled_excel(df, comm_pct, payout, win_num_in, limit):
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
    bd = Border(left=Side(style="thin", color="CBD5E1"),
                right=Side(style="thin", color="CBD5E1"),
                top=Side(style="thin", color="CBD5E1"),
                bottom=Side(style="thin", color="CBD5E1"))

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

    # Sheet 1
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
                if i == 3: c.fill = accent
                if i == 7: c.fill = profit if pl >= 0 else loss
                if j == 2 and i != 4:
                    c.number_format = "#,##0"
                    c.alignment = Alignment(horizontal="right")
    sheet_fmt(ws, {"A": 42, "B": 24, "C": 38})

    # Sheet 2
    ws = wb.create_sheet("နေ့အလိုက်_အပတ်စဉ်")
    ws["A1"] = "နေ့အလိုက် စာရင်းချုပ်"
    ws["A1"].font = f_title
    dg = df.groupby(["Date", "Day"]).agg(Bets=("Number", "count"), Sales=("Amount", "sum")).reset_index()
    hdr(ws, 3, ["ရက်စွဲ", "နေ့", "အကွက်", "ထိုးကြေး", f"ကော်မရှင် ({comm_pct}%)", "Net"])
    for r, row in enumerate(dg.itertuples(), 4):
        vals = [row.Date, str(row.Day), int(row.Bets), float(row.Sales),
                f"=D{r}*{comm_pct/100}", f"=D{r}-E{r}"]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, bd
            if ci >= 3: c.number_format = "#,##0"
            if r % 2: c.fill = zebra
    sheet_fmt(ws, {"A": 18, "B": 16, "C": 14, "D": 22, "E": 20, "F": 22})

    # Sheet 3
    ws = wb.create_sheet("ဖောက်သည်အလိုက်")
    ws["A1"] = "ဖောက်သည်အလိုက် ရှင်းတမ်း"
    ws["A1"].font = f_title
    cg = df.groupby("Customer")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
    hdr(ws, 3, ["ဖောက်သည်", "ထိုးကြေး", f"ကော်မရှင် ({comm_pct}%)", "Net", f"ဆုကြေး ({payout}ဆ)", "အသားတင်"])
    for r, row in enumerate(cg.itertuples(), 4):
        cwin = 0
        if win:
            cwin = float(df[(df["Customer"] == row.Customer) & (df["Number"] == win)]["Amount"].sum()) * payout
        vals = [row.Customer, float(row.Amount), f"=B{r}*{comm_pct/100}",
                f"=B{r}-C{r}", cwin, f"=D{r}-E{r}"]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, bd
            if ci > 1: c.number_format = "#,##0"
            if r % 2: c.fill = zebra
    sheet_fmt(ws, {"A": 26, "B": 22, "C": 20, "D": 24, "E": 24, "F": 24})

    # Sheet 4
    ws = wb.create_sheet("ဂဏန်း_Limit")
    ws["A1"] = "ဂဏန်းအလိုက် Limit"
    ws["A1"].font = f_title
    ng = df.groupby("Number")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
    hdr(ws, 3, ["2D", "စုစုပေါင်း", "ဒိုင်လက်ခံ", "ကျော်လွန်"])
    for r, row in enumerate(ng.itertuples(), 4):
        nc = ws.cell(row=r, column=1, value=str(row.Number).zfill(2))
        nc.number_format, nc.alignment = "@", Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=float(row.Amount)).number_format = "#,##0"
        ws.cell(row=r, column=3, value=f"=MIN(B{r},{limit})").number_format = "#,##0"
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = "#,##0"
        for ci in range(1, 5):
            ws.cell(row=r, column=ci).font = f_body
            ws.cell(row=r, column=ci).border = bd
        if float(row.Amount) > limit:
            ws.cell(row=r, column=4).fill = loss
            ws.cell(row=r, column=4).font = f_bold
    sheet_fmt(ws, {"A": 16, "B": 24, "C": 24, "D": 26})

    # Sheet 5
    ws = wb.create_sheet("အသေးစိတ်")
    ws["A1"] = "အသေးစိတ် မှတ်တမ်း"
    ws["A1"].font = f_title
    dg = df.groupby(["Date", "Day", "Customer", "BatchID", "Number"])["Amount"].sum().reset_index()
    hdr(ws, 3, ["ရက်စွဲ", "နေ့", "ဖောက်သည်", "Batch", "2D", "ထိုးငွေ"])
    for r, row in enumerate(dg.itertuples(), 4):
        vals = [row.Date, row.Day, row.Customer, row.BatchID, str(row.Number).zfill(2), float(row.Amount)]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font, c.border = f_body, bd
            if ci == 5: c.number_format = "@"
            if ci == 6: c.number_format = "#,##0"
            if r % 2: c.fill = zebra
    sheet_fmt(ws, {"A": 16, "B": 16, "C": 22, "D": 22, "E": 16, "F": 22})

    wb.save(out)
    return out.getvalue()

# =========================================================
# PARSER TEST RUNNER (sidebar)
# =========================================================
with st.sidebar.expander("🧪 Parser Test Runner"):
    if st.button("Run All Tests"):
        tests = [
            ("T1 နေ့ရှင်း", "44 99 11\n11000နေ့ရှင်း", 33000),
            ("T2 slash*", "11/99/77*3000\n33/55*1000", 11000),
            ("T3 r-reverse", "46/96r1500", 6000),
            ("T4 bk-multi", "5bk/7bk250", 5000),
            ("T5 ပက်", "7ပက်2500", 47500),
            ("T6 multi-R", "30-35-50-R1500", 9000),
            ("T7 single*", "77*2500", 2500),
            ("T9 leading-0", "01/02/03*1000", 3000),
            ("T10 dup", "11*1000\n11*1000", 2000),
        ]
        for name, txt, exp in tests:
            res = parse_bets_locally(txt)
            actual = sum(r["Amount"] for r in res["rows"])
            mark = "✅" if actual == exp else "❌"
            st.write(f"{mark} {name}: `{actual:,}` (exp `{exp:,}`)")

# =========================================================
# ENTRY SECTION
# =========================================================
st.subheader("📝 စာရင်းသွင်းရန်")

is_closed = False
if enable_time_lock and datetime.now().time() >= closing_time:
    is_closed = True
    st.error(f"⚠️ စာရင်းပိတ်ချိန် ({closing_time.strftime('%H:%M')}) ကျော်လွန်ပါပြီ။")

c1, c2 = st.columns(2)
with c1:
    entry_date = st.date_input("📆 ရက်စွဲ", value=datetime.now().date())
with c2:
    day_opts = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "WEEKEND"]
    selected_day = st.selectbox("📅 နေ့", day_opts, index=min(entry_date.weekday(), 5))
date_str = entry_date.strftime("%Y-%m-%d")

col_a, col_b = st.columns([1, 2])
with col_a:
    cust = st.text_input("ဖောက်သည် အမည်", placeholder="ဥပမာ - သိမ့်")
with col_b:
    text_in = st.text_area(
        "ဖောက်သည် ပို့လိုက်သော စာသား",
        placeholder="44 99 11\n11000နေ့ရှင်း\n11/99/77*3000\n46/96r1500\n5bk/7bk250\n7ပက်2500",
        height=180,
    )

if st.button("🔍 Parse & Preview", type="primary"):
    if is_closed:
        st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပါသည်။")
    elif not cust.strip() or not text_in.strip():
        st.warning("ဖောက်သည်အမည်နှင့် စာသားကို ဖြည့်ပါ။")
    else:
        res = parse_bets_locally(text_in)
        st.session_state.pending_parse = {
            "customer": cust.strip(), "date": date_str, "day": selected_day,
            "rows": res["rows"], "warnings": res["warnings"], "errors": res["errors"],
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
        df_prev = pd.DataFrame(p["rows"])[["Number", "Amount"]]
        st.caption(f"ကွက် {len(df_prev)} ခု | စုစုပေါင်း **{df_prev['Amount'].sum():,.0f}** ကျပ်")

        edited = st.data_editor(
            df_prev, hide_index=True, use_container_width=True, num_rows="dynamic",
            column_config={
                "Number": st.column_config.TextColumn("2D", max_chars=2),
                "Amount": st.column_config.NumberColumn("Amount", min_value=100, max_value=500000, step=100),
            }, key="parse_preview",
        )

        cs, cc = st.columns(2)
        with cs:
            if st.button("✅ Save to Ledger", type="primary"):
                rows_fixed = []
                for _, row in edited.iterrows():
                    try:
                        num = str(row.get("Number", "") or "").strip()
                        if not num or not num.isdigit() or len(num) > 2:
                            continue
                        amt_val = row.get("Amount")
                        if pd.isna(amt_val):
                            continue
                        amt = int(amt_val)
                        if amt <= 0:
                            continue
                        rows_fixed.append({"Number": num.zfill(2), "Amount": amt})
                    except (ValueError, TypeError):
                        continue
                v, blk, inv, tot = process_and_add_bets(
                    rows_fixed, p["date"], p["day"], p["customer"]
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
            if st.button("❌ Cancel"):
                st.session_state.pending_parse = None
                st.rerun()
    else:
        st.warning("Parse မှ ရလဒ် မထွက်ပါ။ စာသားပုံစံကို စစ်ဆေးပါ။")
        if st.button("❌ Cancel"):
            st.session_state.pending_parse = None
            st.rerun()

# =========================================================
# DASHBOARD
# =========================================================
st.markdown("---")

if not st.session_state.data.empty:
    df = st.session_state.data.copy()
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
            st.info(f"ပေါက်ဂဏန်း ({win}) ထိုးထားသူ မရှိ။ ဒိုင်အမြတ် = {net:,.0f}")
    else:
        st.info("Sidebar တွင် ပေါက်ဂဏန်း ထည့်ပါ။")

    # Edit/Delete
    st.markdown("---")
    with st.expander("✏️ Edit / Delete Bet Records"):
        edited_ledger = st.data_editor(
            st.session_state.data, num_rows="dynamic", use_container_width=True, key="edit_ledger",
        )
        ce1, ce2 = st.columns(2)
        with ce1:
            if st.button("💾 Apply Changes to Ledger"):
                st.session_state.data = edited_ledger.reset_index(drop=True)
                st.session_state.flash_success = "ပြင်ဆင်မှု သိမ်းပြီး"
                st.rerun()
        with ce2:
            st.caption(f"စုစုပေါင်း Record: {len(st.session_state.data)}")

    # Viber statements
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
        st.markdown("##### 👤 ဖောက်သည် ရှင်းတမ်း")
        custs = sorted(df["Customer"].unique())
        sel = st.selectbox("ဖောက်သည်", custs)
        if sel:
            cdf = df[df["Customer"] == sel]
            lines = [f"*{sel}* ၏ 2D ရှင်းတမ်း", "=" * 28]
            batch_totals = []
            for b in cdf["BatchID"].unique():
                bdf = cdf[cdf["BatchID"] == b]
                bt = float(bdf["Amount"].sum())
                batch_totals.append((b, bt))
                lines.append(f"--- *{b}* ---")
                for _, r in bdf.iterrows():
                    lines.append(f"{r['Number']} = {r['Amount']:,} ကျပ်")
                lines.append(f"*{b} Total* = {bt:,.0f}\n")
            ctot = float(cdf["Amount"].sum())
            ccomm = ctot * commission_pct / 100
            cnet = ctot - ccomm
            cwin = 0
            if win:
                cwin = float(cdf[cdf["Number"] == win]["Amount"].sum()) * payout_rate
            bal = cnet - cwin
            formula = " + ".join([f"{b}={a:,.0f}" for b, a in batch_totals])
            lines.extend([
                "=" * 28,
                f"{formula} = *({ctot:,.0f} ကျပ်)*",
                f"ကော်မရှင် ({commission_pct}%) = -{ccomm:,.0f}",
                f"Net = {cnet:,.0f}",
            ])
            if win:
                lines.append(f"ဆုကြေး ({win}) = +{cwin:,.0f}")
                lines.append(f"စုစုပေါင်း = {bal:,.0f}")
            st.text_area(f"Viber ({sel})", "\n".join(lines), height=320)

    # Number limit
    st.markdown("---")
    cl, cr = st.columns(2)

    with cl:
        st.subheader("🔢 ဂဏန်း Limit")
        grp = df.groupby("Number")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
        grp["Allowed"] = grp["Amount"].apply(lambda x: min(x, max_limit))
        grp["Excess"] = grp["Amount"] - grp["Allowed"]
        st.dataframe(grp, use_container_width=True)

    with cr:
        st.subheader("👥 Batch အလိုက်")
        cons = df.groupby(["Date", "Day", "Customer", "BatchID", "Number"])["Amount"].sum().reset_index()
        st.dataframe(cons, use_container_width=True)

    # Export & Clear
    st.markdown("---")
    st.subheader("📥 Export & Management")
    ce, cx = st.columns([2, 1])

    with ce:
        xls = generate_styled_excel(df, commission_pct, payout_rate, winning_number, max_limit)
        st.download_button(
            "📥 Excel Download", data=xls,
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
