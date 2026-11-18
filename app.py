
import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from itertools import permutations, product

# =========================================================
# 1. APP CONFIGURATION
# =========================================================
st.set_page_config(
    page_title="Myanmar 2D Master Dealer App",
    page_icon="📒",
    layout="wide"
)

st.markdown("""
<style>
.main-header {
    border-left: 6px solid #D97706;
    padding-left: 15px;
    margin-bottom: 20px;
}
.stButton>button {
    border-radius: 8px;
    font-weight: bold;
}
div[data-testid="stMetricValue"] {
    color: #D97706;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
<h1>📒 မြန်မာနှစ်လုံးထီ (2D) အဆင့်မြင့် အလိုအလျောက် စာရင်းချုပ်စနစ်</h1>
<p>Master Smart Ledger, Ultimate Precision Parser & Custom Formula Viber Statement Generator (API-Free)</p>
</div>
""", unsafe_allow_html=True)

# =========================================================
# 2. SIDEBAR SETTINGS
# =========================================================
st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များနှင့် ကန့်သတ်ချက်များ")

max_limit = st.sidebar.number_input(
    "ဂဏန်းတစ်လုံးလျှင် အများဆုံးလက်ခံငွေ (Limit)",
    min_value=0,
    value=500000,
    step=50000
)

payout_rate = st.sidebar.number_input(
    "ဆုကြေးနှုန်း (ဆ)",
    min_value=0,
    value=80,
    step=5
)

commission_pct = st.sidebar.number_input(
    "ကော်မရှင် (%)",
    min_value=0,
    max_value=100,
    value=10,
    step=1
)

winning_number = st.sidebar.text_input(
    "ထွက်မည့် ပေါက်ဂဏန်း",
    value=""
)

st.sidebar.markdown("---")
st.sidebar.subheader("⏰ အချိန်ပိတ်ရက် ထိန်းချုပ်မှု")

enable_time_lock = st.sidebar.checkbox(
    "အချိန်ပိတ်ရန် ဖွင့်မည်",
    value=False
)

closing_time = st.sidebar.time_input(
    "ပိတ်မည့် အချိန်",
    value=datetime.strptime("12:00", "%H:%M").time()
)

st.sidebar.markdown("---")
st.sidebar.subheader("🚫 ပိတ်ဂဏန်းများ")

blocked_input = st.sidebar.text_input(
    "မခံမည့် ဂဏန်းများ (ဥပမာ: 39, 88)",
    value=""
)

blocked_numbers = [
    num.strip().zfill(2)
    for num in blocked_input.split(",")
    if num.strip().isdigit() and 0 <= int(num.strip()) <= 99
]

# =========================================================
# 3. SESSION STATE
# =========================================================
COLUMNS = ["Date", "Day", "Customer", "BatchID", "Number", "Amount"]

if "data" not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=COLUMNS)

# =========================================================
# 4. PARSER HELPERS
# =========================================================
def get_last_amount(line, default=1000):
    """
    စာကြောင်း၏ နောက်ဆုံးတွင်ပါသော ထိုးငွေကို ရှာဖွေသည်။
    ဥပမာ 7ပက်2500 / 9ပတ်150tt
    """
    match = re.search(r"(\d[\d,]*)(?:tt)?\s*$", line, re.IGNORECASE)

    if match:
        value = match.group(1).replace(",", "")
        try:
            return int(value)
        except ValueError:
            pass

    return default


def add_number(rows, number, amount):
    """ဂဏန်းနှင့် ငွေပမာဏကို စစ်ဆေးပြီး ထည့်သွင်းသည်။"""
    number = str(number).zfill(2)

    if len(number) == 2 and number.isdigit():
        if 0 <= int(number) <= 99:
            rows.append({
                "Number": number,
                "Amount": int(amount)
            })


def extract_numbers(text):
    """ဂဏန်းစာရင်းမှ 1 လုံး/2 လုံး ဂဏန်းများကို ရှာဖွေသည်။"""
    return re.findall(r"\d{1,2}", text)


# =========================================================
# 5. ULTIMATE 2D PARSER
# =========================================================
def parse_bets_locally(text):
    rows = []

    lines = (
        text.replace("၊", "\n")
        .replace(";", "\n")
        .replace("\r", "\n")
        .split("\n")
    )

    for original_line in lines:
        line = original_line.strip()

        if not line:
            continue

        # -------------------------------------------------
        # 1. နေ့ရှင်း
        # ဥပမာ 11000 နေ့ရှင်း / 11,000 နေ့ရှင်း
        # -------------------------------------------------
        match = re.search(r"([\d,]+)\s*နေ့ရှင်း", line)

        if match:
            amount = int(match.group(1).replace(",", ""))
            add_number(rows, "11", amount)
            continue

        # -------------------------------------------------
        # 2. ညီအကို / ညီကို
        # -------------------------------------------------
        if re.search(r"ညီအကို|ညီကို", line):
            amount = get_last_amount(line, 1000)

            pairs = [
                "01", "12", "23", "34", "45",
                "56", "67", "78", "89", "90"
            ]

            added = set()

            for number in pairs:
                for item in (number, number[::-1]):
                    if item not in added:
                        added.add(item)
                        add_number(rows, item, amount)

            continue

        # -------------------------------------------------
        # 3. အပူ / ပူး
        # -------------------------------------------------
        if re.search(r"အပူ|ပူး", line) and not re.search(r"ခွေ", line):
            amount = get_last_amount(line, 1000)

            for i in range(10):
                add_number(rows, f"{i}{i}", amount)

            continue

        # -------------------------------------------------
        # 4. နက္ခတ်
        # -------------------------------------------------
        if "နက္ခတ်" in line:
            amount = get_last_amount(line, 1000)

            for number in [
                "07", "70", "18", "81", "24",
                "42", "35", "53", "69", "96"
            ]:
                add_number(rows, number, amount)

            continue

        # -------------------------------------------------
        # 5. ပါဝါ / POWER
        # -------------------------------------------------
        if re.search(r"ပါဝါ|power", line, re.IGNORECASE):
            amount = get_last_amount(line, 1000)

            for number in [
                "05", "50", "16", "61", "27",
                "72", "38", "83", "49", "94"
            ]:
                add_number(rows, number, amount)

            continue

        # -------------------------------------------------
        # 6. ဘရိတ် / B / BK
        # ဥပမာ 5bk / 7bk250 / 4 B 1500
        # -------------------------------------------------
        break_matches = re.findall(
            r"(\d)\s*(ဘရိတ်|B|bk)",
            line,
            re.IGNORECASE
        )

        if break_matches:
            amount = get_last_amount(line, 1000)

            for b_match in break_matches:
                digit = int(b_match[0])

                for i in range(100):
                    number = f"{i:02d}"
                    digit_sum = (
                        int(number[0]) + int(number[1])
                    ) % 10

                    if digit_sum == digit:
                        add_number(rows, number, amount)

            continue

        # -------------------------------------------------
        # 7. ပတ် / ပါတ် / ပက် / ပါ / ပ
        # ဥပမာ 9ပတ်150tt / 7ပက်2500
        # -------------------------------------------------
        paat_match = re.search(
            r"(\d)\s*(ပါတ်|ပတ်|ပက်|ပါ|ပ)(?!ူး)",
            line
        )

        if paat_match:
            digit = paat_match.group(1)
            amount = get_last_amount(line, 1000)

            for i in range(100):
                number = f"{i:02d}"

                if digit in number:
                    add_number(rows, number, amount)

            continue

        # -------------------------------------------------
        # 8. ထိပ်စီး / ထိပ်
        # ဥပမာ 5ထိပ်စီး 20000
        # -------------------------------------------------
        thate_match = re.search(
            r"(\d)\s*(ထိပ်စီး|ထိပ်)",
            line
        )

        if thate_match:
            digit = thate_match.group(1)
            amount = get_last_amount(line, 1000)

            for i in range(10):
                add_number(rows, f"{digit}{i}", amount)

            continue

        # -------------------------------------------------
        # 9. နောက်ပိတ် / ပိတ်
        # -------------------------------------------------
        pate_match = re.search(
            r"(\d)\s*(နောက်ပိတ်|ပိတ်)",
            line
        )

        if pate_match:
            digit = pate_match.group(1)
            amount = get_last_amount(line, 1000)

            for i in range(10):
                add_number(rows, f"{i}{digit}", amount)

            continue

        # -------------------------------------------------
        # 10. ခွေ / အခွေ
        # ဥပမာ 123 ခွေ 1000
        # အပူးပါပါက doubles များပါ ထည့်မည်
        # -------------------------------------------------
        khve_match = re.search(
            r"(\d{2,5})\s*(ခွေ|အခွေ)",
            line
        )

        if khve_match:
            digits = khve_match.group(1)
            amount = get_last_amount(line, 1000)

            include_doubles = bool(
                re.search(r"အပူးပါ|ပူးပါ|အပူးအပါ", line)
            )

            if include_doubles:
                combinations = product(set(digits), repeat=2)
            else:
                combinations = permutations(digits, 2)

            added = set()

            for pair in combinations:
                number = "".join(pair)

                if number not in added:
                    added.add(number)
                    add_number(rows, number, amount)

            continue

        # -------------------------------------------------
        # 11. @ နှင့် _ ပုံစံ
        # ဥပမာ 20_4000@1000
        # -------------------------------------------------
        at_match = re.fullmatch(
            r"(\d{1,2})[_.](\d+)@(\d+)",
            line
        )

        if at_match:
            number = at_match.group(1).zfill(2)
            base_amount = int(at_match.group(2))
            reverse_amount = int(at_match.group(3))

            add_number(rows, number, base_amount)
            add_number(rows, number[::-1], reverse_amount)
            continue

        # -------------------------------------------------
        # 12. ရိုးရိုး @ ပုံစံ
        # ဥပမာ 20@3000
        # -------------------------------------------------
        at_simple = re.fullmatch(
            r"(\d{1,2})@(\d+)",
            line
        )

        if at_simple:
            number = at_simple.group(1).zfill(2)
            amount = int(at_simple.group(2))

            add_number(rows, number, amount)
            add_number(rows, number[::-1], amount)
            continue

        # -------------------------------------------------
        # 13. Complex R ပုံစံ
        # ဥပမာ 20-30-40-3000R500
        # -------------------------------------------------
        complex_r = re.fullmatch(
            r"([\d./,\-\s]+?)[\-=\s](\d{3,6})"
            r"\s*([Rrရ@&အာ])\s*(\d+)",
            line
        )

        if complex_r:
            nums_part = complex_r.group(1)
            base_amount = int(complex_r.group(2))
            reverse_amount = int(complex_r.group(4))

            numbers = extract_numbers(nums_part)

            for item in numbers:
                number = item.zfill(2)
                add_number(rows, number, base_amount)
                add_number(rows, number[::-1], reverse_amount)

            continue

        # -------------------------------------------------
        # 14. Multi R ပုံစံ
        # ဥပမာ 46/96r1500
        # ဥပမာ 30-35-50-R1500
        # -------------------------------------------------
        multi_r = re.fullmatch(
            r"([\d./,\-\s]+?)\s*[Rr]\s*(\d+)",
            line
        )

        if multi_r:
            nums_part = multi_r.group(1)
            amount = int(multi_r.group(2))

            numbers = extract_numbers(nums_part)

            for item in numbers:
                number = item.zfill(2)

                add_number(rows, number, amount)

                reverse = number[::-1]

                if reverse != number:
                    add_number(rows, reverse, amount)

            continue

        # -------------------------------------------------
        # 15. Dot R ပုံစံ
        # ဥပမာ 85.3000.r.500
        # -------------------------------------------------
        dot_r = re.fullmatch(
            r"([\d./,\-\s]+?)\.?\s*[Rr]\.?\s*(\d+)",
            line
        )

        if dot_r:
            nums_part = dot_r.group(1)
            reverse_amount = int(dot_r.group(2))

            tokens = [
                item for item in
                re.split(r"[./,\-\s]+", nums_part)
                if item.isdigit()
            ]

            numbers = []
            base_amount = None

            for token in tokens:
                if len(token) <= 2:
                    numbers.append(token.zfill(2))
                else:
                    base_amount = int(token)

            if base_amount is None:
                base_amount = reverse_amount

            for number in numbers:
                add_number(rows, number, base_amount)
                add_number(rows, number[::-1], reverse_amount)

            continue

        # -------------------------------------------------
        # 16. Slash / Dash / * / = ပုံစံ
        # ဥပမာ 11/99/77*3000
        # ဥပမာ 33/55*1000
        # ဥပမာ 99=250
        # -------------------------------------------------
        slash_amt = re.fullmatch(
            r"([\d./,\-\s]+?)\s*[\*=\s]\s*(\d{3,6})",
            line
        )

        if slash_amt:
            nums_part = slash_amt.group(1)
            amount = int(slash_amt.group(2))

            for item in extract_numbers(nums_part):
                add_number(rows, item, amount)

            continue

        # -------------------------------------------------
        # 17. ဂဏန်းများစွာ + နောက်ဆုံးထိုးငွေ
        # ဥပမာ 82.45.10.20.40.50.58.250
        # -------------------------------------------------
        multi_amount = re.fullmatch(
            r"([\d./,\-\s]+?)\s*[./,\-\s]\s*(\d{3,6})",
            line
        )

        if multi_amount:
            nums_part = multi_amount.group(1)
            amount = int(multi_amount.group(2))

            for item in extract_numbers(nums_part):
                add_number(rows, item, amount)

            continue

        # -------------------------------------------------
        # 18. တစ်ကွက် R
        # ဥပမာ 20R3000
        # -------------------------------------------------
        single_r = re.fullmatch(
            r"(\d{1,2})\s*[Rr]\s*(\d+)",
            line
        )

        if single_r:
            number = single_r.group(1).zfill(2)
            amount = int(single_r.group(2))

            add_number(rows, number, amount)

            if number != number[::-1]:
                add_number(rows, number[::-1], amount)

            continue

        # -------------------------------------------------
        # 19. တစ်ကွက် ရိုးရိုးထိုးငွေ
        # ဥပမာ 20-3000 / 20 3000
        # -------------------------------------------------
        single_amount = re.fullmatch(
            r"(\d{1,2})[\s.\-](\d{3,6})",
            line
        )

        if single_amount:
            number = single_amount.group(1).zfill(2)
            amount = int(single_amount.group(2))

            add_number(rows, number, amount)
            continue

        # -------------------------------------------------
        # 20. ဂဏန်းများသာ ပါသောပုံစံ
        # ဥပမာ 44 99 11
        # -------------------------------------------------
        if (
            re.fullmatch(r"\d{2}(?:[\s,./-]+\d{2})*", line)
            and not re.search(r"[A-Za-z]", line)
        ):
            numbers = re.findall(r"\d{2}", line)

            for number in numbers:
                add_number(rows, number, 1000)

            continue

        # -------------------------------------------------
        # 21. Fallback
        # ဥပမာ 20-3000 / 35=1500
        # -------------------------------------------------
        pairs = re.findall(
            r"(\d{1,2})\s*[-=/_\.\s]\s*(\d{3,6})",
            line
        )

        if pairs:
            for number, amount in pairs:
                add_number(rows, number, int(amount))

            continue

    return rows


# =========================================================
# 6. PROCESS AND ADD BETS
# =========================================================
def process_and_add_bets(
    extracted_data,
    entry_date_str,
    selected_day,
    default_customer
):
    valid_rows = []
    blocked_found = []
    invalid_amt_found = 0
    total_added_amount = 0

    existing_batches = st.session_state.data[
        st.session_state.data["Customer"] == default_customer
    ]["BatchID"]

    next_batch_num = 1

    if not existing_batches.empty:
        batch_numbers = []

        for batch in existing_batches.unique():
            match = re.search(r"\((\d+)\)", str(batch))

            if match:
                batch_numbers.append(int(match.group(1)))

        if batch_numbers:
            next_batch_num = max(batch_numbers) + 1

    batch_id = f"{default_customer}({next_batch_num})"

    for row in extracted_data:
        amount = int(row.get("Amount", 0))
        number = str(row.get("Number", "")).zfill(2)

        if amount < 100 or amount > 500000:
            invalid_amt_found += 1
            continue

        if number in blocked_numbers:
            blocked_found.append(number)
            continue

        valid_rows.append({
            "Date": entry_date_str,
            "Day": selected_day,
            "Customer": default_customer,
            "BatchID": batch_id,
            "Number": number,
            "Amount": amount
        })

        total_added_amount += amount

    return (
        valid_rows,
        blocked_found,
        invalid_amt_found,
        total_added_amount
    )


# =========================================================
# 7. EXCEL REPORT GENERATOR - 5 SHEETS
# =========================================================
def generate_styled_excel(
    df,
    commission_pct,
    payout_rate,
    winning_number,
    max_limit
):
    output = io.BytesIO()
    wb = openpyxl.Workbook()

    header_fill = PatternFill(
        start_color="1E293B",
        end_color="1E293B",
        fill_type="solid"
    )

    accent_fill = PatternFill(
        start_color="FEF3C7",
        end_color="FEF3C7",
        fill_type="solid"
    )

    zebra_fill = PatternFill(
        start_color="F8FAFC",
        end_color="F8FAFC",
        fill_type="solid"
    )

    profit_fill = PatternFill(
        start_color="DCFCE7",
        end_color="DCFCE7",
        fill_type="solid"
    )

    loss_fill = PatternFill(
        start_color="FEE2E2",
        end_color="FEE2E2",
        fill_type="solid"
    )

    font_title = Font(
        name="Pyidaungsu",
        size=14,
        bold=True,
        color="1E293B"
    )

    font_header = Font(
        name="Pyidaungsu",
        size=11,
        bold=True,
        color="FFFFFF"
    )

    font_body = Font(
        name="Pyidaungsu",
        size=11,
        color="0F172A"
    )

    font_bold = Font(
        name="Pyidaungsu",
        size=11,
        bold=True,
        color="0F172A"
    )

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    def format_sheet(ws, widths):
        ws.views.sheetView[0].showGridLines = True

        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        ws.freeze_panes = "A4"

    def style_header(ws, row, headers):
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = font_header
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )
            cell.border = thin_border

    win_num = (
        str(winning_number).strip().zfill(2)
        if winning_number else ""
    )

    total_sales = float(df["Amount"].sum())
    commission_amount = total_sales * commission_pct / 100
    net_sales = total_sales - commission_amount

    win_total = (
        float(df[df["Number"] == win_num]["Amount"].sum())
        if win_num else 0
    )

    payout_total = win_total * payout_rate
    net_pl = net_sales - payout_total

    # -----------------------------------------------------
    # TAB 1: SUMMARY
    # -----------------------------------------------------
    ws = wb.active
    ws.title = "စာရင်းချုပ်အကျဉ်း"

    ws["A1"] = "မြန်မာနှစ်လုံးထီ (2D) ငွေစာရင်းချုပ် အစီရင်ခံစာ"
    ws["A1"].font = font_title

    ws["A2"] = (
        f"ထုတ်ပြန်သည့်ရက်စွဲ - "
        f"{datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )
    ws["A2"].font = font_body

    summary_rows = [
        ("အကြောင်းအရာ", "ပမာဏ / တန်ဖိုး", "မှတ်ချက်"),
        (
            "စုစုပေါင်း ရောင်းရငွေ (Total Sales)",
            total_sales,
            f"စုစုပေါင်း အကွက်ရေ {len(df):,} ကွက်"
        ),
        (
            f"ကော်မရှင် နှုတ်ငွေ ({commission_pct}%)",
            commission_amount,
            "ဒိုင်ကော်မရှင် နှုတ်ယူငွေ"
        ),
        (
            "ကော်မရှင်နှုတ်ပြီး ဒိုင်ရငွေ (Net Sales)",
            "=B5-B6",
            "အသားတင် ရငွေ"
        ),
        (
            "ထွက်ဂဏန်း (Winning Number)",
            win_num if win_num else "မသတ်မှတ်ရသေး",
            f"ဆုကြေးနှုန်း {payout_rate} ဆ"
        ),
        (
            "ပေါက်ဂဏန်းပေါ် ထိုးငွေစုစုပေါင်း",
            win_total,
            "ပေါက်သီး တန်ဖိုး"
        ),
        (
            "ပေးရမည့် ဆုကြေးငွေစုစုပေါင်း",
            f"=B9*{payout_rate}",
            "ကံထူးရှင်များသို့ ပေးလျော်ငွေ"
        ),
        (
            "ဒိုင် အသားတင် အနိုင် / အရှုံး",
            "=B7-B10",
            "အပေါင်း (+) အမြတ် / အနှုတ် (-) အရှုံး"
        )
    ]

    for idx, values in enumerate(summary_rows):
        row_num = idx + 4

        for col_num, value in enumerate(values, 1):
            cell = ws.cell(
                row=row_num,
                column=col_num,
                value=value
            )

            cell.border = thin_border

            if idx == 0:
                cell.fill = header_fill
                cell.font = font_header
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.font = font_body

                if idx == 3:
                    cell.fill = accent_fill

                if idx == 7:
                    cell.fill = profit_fill if net_pl >= 0 else loss_fill

                if col_num == 2 and idx != 4:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right")

    format_sheet(ws, {
        "A": 42,
        "B": 24,
        "C": 38
    })

    # -----------------------------------------------------
    # TAB 2: DAILY SUMMARY
    # -----------------------------------------------------
    ws_day = wb.create_sheet("နေ့အလိုက်_အပတ်စဉ်")

    ws_day["A1"] = "နေ့အလိုက် နှင့် တစ်ပတ်စာ စာရင်းချုပ်"
    ws_day["A1"].font = font_title

    day_group = (
        df.groupby(["Date", "Day"])
        .agg(
            Total_Bets=("Number", "count"),
            Gross_Sales=("Amount", "sum")
        )
        .reset_index()
    )

    style_header(ws_day, 3, [
        "ရက်စွဲ",
        "နေ့",
        "အကွက်ရေ",
        "စုစုပေါင်းထိုးကြေး",
        f"ကော်မရှင် ({commission_pct}%)",
        "ဒိုင်ရငွေ (Net)"
    ])

    for r, row in enumerate(day_group.itertuples(), 4):
        values = [
            row.Date,
            str(row.Day),
            int(row.Total_Bets),
            float(row.Gross_Sales),
            f"=D{r}*{commission_pct / 100}",
            f"=D{r}-E{r}"
        ]

        for c, value in enumerate(values, 1):
            cell = ws_day.cell(row=r, column=c, value=value)
            cell.font = font_body
            cell.border = thin_border

            if c >= 3:
                cell.number_format = "#,##0"

            if r % 2 == 1:
                cell.fill = zebra_fill

    format_sheet(ws_day, {
        "A": 18,
        "B": 16,
        "C": 14,
        "D": 22,
        "E": 20,
        "F": 22
    })

    # -----------------------------------------------------
    # TAB 3: CUSTOMER SUMMARY
    # -----------------------------------------------------
    ws_cust = wb.create_sheet("ဖောက်သည်အလိုက် ရှင်းတမ်း")

    ws_cust["A1"] = "ဖောက်သည်တစ်ဦးချင်းစီ၏ ရှင်းတမ်း"
    ws_cust["A1"].font = font_title

    cust_group = (
        df.groupby("Customer")["Amount"]
        .sum()
        .reset_index()
        .sort_values(by="Amount", ascending=False)
    )

    style_header(ws_cust, 3, [
        "ဖောက်သည် အမည်",
        "စုစုပေါင်းထိုးကြေး",
        f"ကော်မရှင် ({commission_pct}%)",
        "ဒိုင်ရငွေ",
        f"ဆုကြေး ({payout_rate}ဆ)",
        "အသားတင်ရှင်းတမ်း"
    ])

    for r, row in enumerate(cust_group.itertuples(), 4):
        customer = row.Customer

        customer_win = 0

        if win_num:
            customer_win = float(
                df[
                    (df["Customer"] == customer) &
                    (df["Number"] == win_num)
                ]["Amount"].sum()
            ) * payout_rate

        values = [
            customer,
            float(row.Amount),
            f"=B{r}*{commission_pct / 100}",
            f"=B{r}-C{r}",
            customer_win,
            f"=D{r}-E{r}"
        ]

        for c, value in enumerate(values, 1):
            cell = ws_cust.cell(row=r, column=c, value=value)
            cell.font = font_body
            cell.border = thin_border

            if c > 1:
                cell.number_format = "#,##0"

            if r % 2 == 1:
                cell.fill = zebra_fill

    format_sheet(ws_cust, {
        "A": 26,
        "B": 22,
        "C": 20,
        "D": 24,
        "E": 24,
        "F": 24
    })

    # -----------------------------------------------------
    # TAB 4: NUMBER LIMIT
    # -----------------------------------------------------
    ws_num = wb.create_sheet("ဂဏန်းပေါင်းချုပ်_Limit")

    ws_num["A1"] = "ဂဏန်းအလိုက် စုစုပေါင်းနှင့် Limit ထိန်းချုပ်မှု"
    ws_num["A1"].font = font_title

    num_group = (
        df.groupby("Number")["Amount"]
        .sum()
        .reset_index()
        .sort_values(by="Amount", ascending=False)
    )

    style_header(ws_num, 3, [
        "2D ဂဏန်း",
        "စုစုပေါင်း ထိုးငွေ",
        "ဒိုင်လက်ခံငွေ (Allowed)",
        "ကျော်လွန်ငွေ (Excess)"
    ])

    for r, row in enumerate(num_group.itertuples(), 4):
        number_cell = ws_num.cell(
            row=r,
            column=1,
            value=str(row.Number).zfill(2)
        )

        number_cell.number_format = "@"
        number_cell.alignment = Alignment(horizontal="center")

        ws_num.cell(
            row=r,
            column=2,
            value=float(row.Amount)
        ).number_format = "#,##0"

        ws_num.cell(
            row=r,
            column=3,
            value=f"=MIN(B{r},{max_limit})"
        ).number_format = "#,##0"

        ws_num.cell(
            row=r,
            column=4,
            value=f"=B{r}-C{r}"
        ).number_format = "#,##0"

        for c in range(1, 5):
            cell = ws_num.cell(row=r, column=c)
            cell.font = font_body
            cell.border = thin_border

        if float(row.Amount) > max_limit:
            ws_num.cell(row=r, column=4).fill = loss_fill
            ws_num.cell(row=r, column=4).font = font_bold

    format_sheet(ws_num, {
        "A": 16,
        "B": 24,
        "C": 24,
        "D": 26
    })

    # -----------------------------------------------------
    # TAB 5: DETAIL RECORDS
    # -----------------------------------------------------
    ws_det = wb.create_sheet("အသေးစိတ်စာရင်း")

    ws_det["A1"] = "အသေးစိတ် စာရင်းမှတ်တမ်း"
    ws_det["A1"].font = font_title

    detail_group = (
        df.groupby([
            "Date", "Day", "Customer",
            "BatchID", "Number"
        ])["Amount"]
        .sum()
        .reset_index()
    )

    style_header(ws_det, 3, [
        "ရက်စွဲ",
        "နေ့",
        "ဖောက်သည်",
        "အကြိမ်ရေ (Batch)",
        "2D ဂဏန်း",
        "ထိုးငွေ (MMK)"
    ])

    for r, row in enumerate(detail_group.itertuples(), 4):
        values = [
            row.Date,
            row.Day,
            row.Customer,
            row.BatchID,
            str(row.Number).zfill(2),
            float(row.Amount)
        ]

        for c, value in enumerate(values, 1):
            cell = ws_det.cell(row=r, column=c, value=value)
            cell.font = font_body
            cell.border = thin_border

            if c == 5:
                cell.number_format = "@"

            if c == 6:
                cell.number_format = "#,##0"

            if r % 2 == 1:
                cell.fill = zebra_fill

    format_sheet(ws_det, {
        "A": 16,
        "B": 16,
        "C": 22,
        "D": 22,
        "E": 16,
        "F": 22
    })

    wb.save(output)

    return output.getvalue()


# =========================================================
# 8. ENTRY SECTION
# =========================================================
st.subheader("📝 စာရင်းသွင်းရန် နည်းလမ်းများ")

is_closed = False

if enable_time_lock:
    current_time = datetime.now().time()

    if current_time >= closing_time:
        is_closed = True
        st.error(
            f"⚠️ စာရင်းပိတ်ချိန် "
            f"({closing_time.strftime('%H:%M')}) "
            "ကျော်လွန်သွားပြီဖြစ်၍ စာရင်းအသစ် လက်မခံတော့ပါ။"
        )

col_d1, col_d2 = st.columns(2)

with col_d1:
    entry_date = st.date_input(
        "📆 စာရင်းသွင်းမည့် ရက်စွဲ",
        value=datetime.now().date()
    )

with col_d2:
    day_options = [
        "MONDAY",
        "TUESDAY",
        "WEDNESDAY",
        "THURSDAY",
        "FRIDAY",
        "WEEKEND"
    ]

    default_day_idx = min(entry_date.weekday(), 5)

    selected_day = st.selectbox(
        "📅 စာရင်းသွင်းမည့် နေ့",
        day_options,
        index=default_day_idx
    )

entry_date_str = entry_date.strftime("%Y-%m-%d")

col1, col2 = st.columns([1, 2])

with col1:
    cust_manual = st.text_input(
        "ဖောက်သည် အမည်",
        placeholder="ဥပမာ - သိမ့်"
    )

with col2:
    smart_input_text = st.text_area(
        "ဖောက်သည် ပို့လိုက်သော စာရင်းစာသားကို ထည့်ပါ",
        placeholder=(
            "85.3000.r.500\n"
            "20@3000\n"
            "82.45.10.20.40.50.58.r.250\n"
            "5ထိပ်စီး 20000\n"
            "4 B 1500\n"
            "11000နေ့ရှင်း\n"
            "11/99/77*3000\n"
            "46/96r1500\n"
            "7ပက်2500\n"
            "44 99 11"
        ),
        height=180
    )

if st.button("➕ စာရင်းသွင်းမည်", type="primary"):
    if is_closed:
        st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")

    elif not cust_manual.strip() or not smart_input_text.strip():
        st.warning("ဖောက်သည် အမည်နှင့် ထိုးမည့် စာသားကို ဖြည့်သွင်းပါ။")

    else:
        try:
            clean_customer = cust_manual.strip()

            extracted_data = parse_bets_locally(smart_input_text)

            (
                valid_rows,
                blocked_found,
                invalid_amt,
                total_amt
            ) = process_and_add_bets(
                extracted_data,
                entry_date_str,
                selected_day,
                clean_customer
            )

            if valid_rows:
                st.session_state.data = pd.concat(
                    [
                        st.session_state.data,
                        pd.DataFrame(valid_rows)
                    ],
                    ignore_index=True
                )

                st.success(
                    f"🎉 စာရင်းသွင်းပြီးပါပြီ! "
                    f"({clean_customer}) "
                    f"စုစုပေါင်းထိုးငွေ = {total_amt:,} ကျပ် "
                    f"({len(valid_rows)} ကွက်)"
                )

                if blocked_found:
                    st.warning(
                        "ပိတ်ဂဏန်းများ ဖယ်ထုတ်ထားပါသည်: "
                        + ", ".join(sorted(set(blocked_found)))
                    )

                if invalid_amt:
                    st.warning(
                        f"ငွေပမာဏ မမှန်ကန်သော အကွက် {invalid_amt} ကွက် "
                        "ဖယ်ထုတ်ထားပါသည်။"
                    )

                st.rerun()

            else:
                st.warning(
                    "ဂဏန်းနှင့် ငွေပမာဏကို ရှာမတွေ့ပါ။ "
                    "စာသားပုံစံ မှန်/မမှန် စစ်ဆေးပါ။"
                )

                if blocked_found:
                    st.warning(
                        "ပိတ်ဂဏန်းများ: "
                        + ", ".join(sorted(set(blocked_found)))
                    )

        except Exception as error:
            st.error(f"စာရင်းတွက်ချက်ရာတွင် အမှားဖြစ်ပေါ်ခဲ့သည်: {error}")


# =========================================================
# 9. DASHBOARD AND REPORTS
# =========================================================
st.markdown("---")

if not st.session_state.data.empty:

    df = st.session_state.data.copy()

    total_sales = float(df["Amount"].sum())
    commission_amount = total_sales * commission_pct / 100
    net_sales = total_sales - commission_amount

    win_num = (
        str(winning_number).strip().zfill(2)
        if winning_number.strip()
        else ""
    )

    winning_matched = (
        df[df["Number"] == win_num]
        if win_num
        else df.iloc[0:0]
    )

    winning_total = float(winning_matched["Amount"].sum())
    payout_amount = winning_total * payout_rate
    net_pl = net_sales - payout_amount

    # -----------------------------------------------------
    # DASHBOARD METRICS
    # -----------------------------------------------------
    st.subheader("📊 အဓိက စာရင်းချုပ်")

    m1, m2, m3, m4 = st.columns(4)

    m1.metric(
        "စုစုပေါင်း ရောင်းရငွေ",
        f"{total_sales:,.0f} MMK"
    )

    m2.metric(
        f"ကော်မရှင် ({commission_pct}%)",
        f"{commission_amount:,.0f} MMK"
    )

    m3.metric(
        "ကော်မရှင်နှုတ်ပြီး ဒိုင်ရငွေ",
        f"{net_sales:,.0f} MMK"
    )

    m4.metric(
        "အသားတင် အနိုင် / အရှုံး",
        f"{net_pl:,.0f} MMK"
    )

    st.markdown("---")

    # -----------------------------------------------------
    # WINNING NUMBER CHECK
    # -----------------------------------------------------
    st.subheader("🏆 ပေါက်ဂဏန်း ကံထူးသူများ စာရင်း တိုက်စစ်ခြင်း")

    if winning_number.strip():

        if not winning_matched.empty:

            st.success(
                f"ပေါက်ဂဏန်း ({win_num}) အတွက် ကံထူးရှင်များ ရှိပါသည်။"
            )

            st.dataframe(
                winning_matched,
                use_container_width=True
            )

            w_col1, w_col2 = st.columns(2)

            w_col1.metric(
                "ပေးရမည့် ဆုကြေးငွေစုစုပေါင်း",
                f"{payout_amount:,.0f} MMK"
            )

            w_col2.metric(
                "ဒိုင် အမြတ် / အရှုံး",
                f"{net_pl:,.0f} MMK"
            )

        else:
            st.info(
                f"ပေါက်ဂဏန်း ({win_num}) ထိုးထားသူ မရှိပါ။ "
                f"ဒိုင်အမြတ် = {net_sales:,.0f} MMK"
            )

    else:
        st.info(
            "Sidebar တွင် ထွက်မည့် ပေါက်ဂဏန်း ထည့်သွင်းပါက "
            "ကံထူးရှင်စာရင်းနှင့် ဆုကြေးများကို တိုက်စစ်နိုင်ပါသည်။"
        )

    st.markdown("---")

    # -----------------------------------------------------
    # VIBER STATEMENTS
    # -----------------------------------------------------
    st.subheader(
        "📨 နေ့အလိုက် / တစ်ပတ်စာ စာရင်းချုပ်နှင့် "
        "ဖောက်သည် Viber Statement Generator"
    )

    col_rep1, col_rep2 = st.columns(2)

    # -----------------------------------------------------
    # DEALER SUMMARY VIBER
    # -----------------------------------------------------
    with col_rep1:

        st.markdown("##### 📒 တစ်နေ့တာ / တစ်ပတ်စာ စုစုပေါင်း ရှင်းတမ်း")

        day_summary = (
            df.groupby(["Date", "Day"])["Amount"]
            .sum()
            .reset_index()
        )

        st.dataframe(
            day_summary,
            use_container_width=True
        )

        report_lines = [
            "*မြန်မာ 2D စာရင်းချုပ်အစီရင်ခံစာ*",
            "=" * 28
        ]

        for row in day_summary.itertuples():
            report_lines.append(
                f"{row.Day} ({row.Date}) : "
                f"{row.Amount:,.0f} ကျပ်"
            )

        report_lines.extend([
            "-" * 28,
            f"စုစုပေါင်းရောင်းရငွေ = {total_sales:,.0f} ကျပ်",
            f"ကော်မရှင် ({commission_pct}%) = -{commission_amount:,.0f} ကျပ်",
            f"ဒိုင်ရငွေ (Net Sales) = {net_sales:,.0f} ကျပ်"
        ])

        if win_num:
            report_lines.append(
                f"ပေါက်ဂဏန်း ({win_num}) တန်ဖိုး = "
                f"{winning_total:,.0f} ကျပ်"
            )

            report_lines.append(
                f"ပေးရမည့်ဆုကြေး ({payout_rate}ဆ) = "
                f"-{payout_amount:,.0f} ကျပ်"
            )

        status = "အမြတ် (+)" if net_pl >= 0 else "အရှုံး (-)"

        report_lines.append(
            f"အသားတင် {status} = {net_pl:,.0f} ကျပ်"
        )

        st.text_area(
            "ဒိုင်ချုပ် Viber Text (Copy ယူရန်)",
            "\n".join(report_lines),
            height=220
        )

    # -----------------------------------------------------
    # CUSTOMER VIBER STATEMENT
    # -----------------------------------------------------
    with col_rep2:

        st.markdown(
            "##### 👤 ဖောက်သည်တစ်ဦးချင်းစီ၏ "
            "အသေးစိတ် Batch များနှင့် ပေါင်းချုပ် Statement"
        )

        customers = sorted(df["Customer"].unique())

        selected_customer = st.selectbox(
            "ဖောက်သည် ရွေးချယ်ပါ",
            customers
        )

        if selected_customer:

            customer_df = df[
                df["Customer"] == selected_customer
            ]

            batches = customer_df["BatchID"].unique()

            viber_lines = [
                f"*{selected_customer}* ၏ 2D အသေးစိတ်ရှင်းတမ်း",
                "=" * 28
            ]

            batch_totals = []

            for batch_id in batches:

                batch_df = customer_df[
                    customer_df["BatchID"] == batch_id
                ]

                batch_total = float(batch_df["Amount"].sum())

                batch_totals.append((batch_id, batch_total))

                viber_lines.append(f"--- *{batch_id}* ---")

                for _, row in batch_df.iterrows():
                    viber_lines.append(
                        f"{row['Number']} = "
                        f"{row['Amount']:,} ကျပ်"
                    )

                viber_lines.append(
                    f"*{batch_id} စုစုပေါင်း* = "
                    f"{batch_total:,.0f} ကျပ်\n"
                )

            customer_total = float(customer_df["Amount"].sum())

            customer_commission = (
                customer_total * commission_pct / 100
            )

            customer_net = customer_total - customer_commission

            customer_win_amount = 0

            if win_num:
                customer_win_amount = float(
                    customer_df[
                        customer_df["Number"] == win_num
                    ]["Amount"].sum()
                ) * payout_rate

            customer_balance = customer_net - customer_win_amount

            formula_items = [
                f"[{batch}={amount:,.0f} ကျပ်]"
                for batch, amount in batch_totals
            ]

            formula_string = (
                " + ".join(formula_items)
                + f" = *({customer_total:,.0f} ကျပ်)*"
            )

            viber_lines.extend([
                "=" * 28,
                formula_string,
                f"ကော်မရှင် ({commission_pct}%) = "
                f"-{customer_commission:,.0f} ကျပ်",
                f"ပေးချေရန် Net = {customer_net:,.0f} ကျပ်"
            ])

            if win_num:
                viber_lines.append(
                    f"ပေါက်ဂဏန်းဆုကြေး ({win_num}) = "
                    f"+{customer_win_amount:,.0f} ကျပ်"
                )

                balance_label = (
                    "ပေးရန် (+) / ရရန် (-)"
                    if customer_balance >= 0
                    else "ရရန်"
                )

                viber_lines.append(
                    f"ရှင်းတမ်းစုစုပေါင်း ({balance_label}) = "
                    f"{customer_balance:,.0f} ကျပ်"
                )

            st.text_area(
                f"Viber သို့ ပို့ရန် ({selected_customer} ရှင်းတမ်း)",
                "\n".join(viber_lines),
                height=320
            )

    st.markdown("---")

    # -----------------------------------------------------
    # NUMBER LIMIT AND CONSOLIDATED VIEW
    # -----------------------------------------------------
    col_left, col_right = st.columns(2)

    with col_left:

        st.subheader(
            "🔢 ဂဏန်းအလိုက် စုစုပေါင်းနှင့် ကန့်သတ်ချက်"
        )

        grouped = (
            df.groupby("Number")["Amount"]
            .sum()
            .reset_index()
            .sort_values(by="Amount", ascending=False)
        )

        grouped["Allowed"] = grouped["Amount"].apply(
            lambda amount: min(amount, max_limit)
        )

        grouped["Excess"] = (
            grouped["Amount"] - grouped["Allowed"]
        )

        st.dataframe(
            grouped,
            use_container_width=True
        )

    with col_right:

        st.subheader(
            "👥 ဖောက်သည်နှင့် Batch အလိုက် ပေါင်းချုပ်စာရင်း"
        )

        consolidated = (
            df.groupby([
                "Date",
                "Day",
                "Customer",
                "BatchID",
                "Number"
            ])["Amount"]
            .sum()
            .reset_index()
        )

        st.dataframe(
            consolidated,
            use_container_width=True
        )

    st.markdown("---")

    # -----------------------------------------------------
    # EXCEL DOWNLOAD AND CLEAR DATA
    # -----------------------------------------------------
    st.subheader("📥 Excel Report နှင့် စာရင်းထိန်းချုပ်မှု")

    col_export, col_clear = st.columns([2, 1])

    with col_export:

        excel_bytes = generate_styled_excel(
            df,
            commission_pct,
            payout_rate,
            winning_number,
            max_limit
        )

        st.download_button(
            label="📥 Excel ဖိုင် Download ဆွဲရန်",
            data=excel_bytes,
            file_name=(
                "2D_Complete_Report_"
                f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary"
        )

    with col_clear:

        if st.button("🗑️ စာရင်း အစအဆုံး ရှင်းလင်းရန်"):

            st.session_state.data = pd.DataFrame(
                columns=COLUMNS
            )

            st.rerun()

else:

    st.info(
        "ယခုထိ ထည့်ထားသော စာရင်း မရှိသေးပါ။ "
        "အထက်ပါ စာရင်းသွင်းကွက်တွင် ရိုက်ထည့်၍ "
        "စာရင်းစတင်သွင်းနိုင်ပါသည်။"
    )
