
import streamlit as st
import pandas as pd
import re
import io
import openpyxl

from datetime import datetime
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)

# ==================================================
# 1. PAGE CONFIGURATION
# ==================================================

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


# ==================================================
# 2. SIDEBAR SETTINGS
# ==================================================

st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များနှင့် ကန့်သတ်ချက်များ")

max_limit = st.sidebar.number_input(
    "ဂဏန်းတစ်လုံးလျှင် အများဆုံးလက်ခံငွေ (Limit)",
    value=500000,
    step=50000
)

payout_rate = st.sidebar.number_input(
    "ဆုကြေးနှုန်း (ဆ)",
    value=80,
    step=5
)

commission_pct = st.sidebar.number_input(
    "ကော်မရှင် (%)",
    value=0,
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
    n.strip().zfill(2)
    for n in blocked_input.split(",")
    if n.strip().isdigit()
]


# ==================================================
# 3. SESSION STATE
# ==================================================

if "data" not in st.session_state:
    st.session_state.data = pd.DataFrame(
        columns=[
            "Date", "Day", "Customer",
            "BatchID", "Number", "Amount"
        ]
    )


# ==================================================
# 4. SMART BET PARSER
# ==================================================

def parse_bets_locally(text):

    rows = []

    lines = (
        text.replace("၊", "\n")
        .replace(";", "\n")
        .split("\n")
    )

    for line in lines:

        line = line.strip()

        if not line:
            continue

        # 1. Day Clearing: 11000နေ့ရှင်း

        match = re.search(
            r'([\d,]+)\s*နေ့ရှင်း',
            line
        )

        if match:
            amount = int(match.group(1).replace(",", ""))
            rows.append({
                "Number": "11",
                "Amount": amount
            })
            continue

        # 2. Break: 5bk / 7bk250

        break_matches = re.findall(
            r'(\d)\s*(ဘရိတ်|B|bk)',
            line,
            re.IGNORECASE
        )

        if break_matches:

            amount_match = re.search(r'(\d+)$', line)
            amount = int(amount_match.group(1)) if amount_match else 250

            for b_match in break_matches:

                digit_sum = int(b_match[0])

                for i in range(100):

                    number = f"{i:02d}"

                    if (
                        int(number[0]) +
                        int(number[1])
                    ) % 10 == digit_sum:

                        rows.append({
                            "Number": number,
                            "Amount": amount
                        })

            continue

        # 3. Paat: 9ပတ်150 / 7ပက်2500

        paat_match = re.search(
            r'(\d)\s*(ပတ်|ပါတ်|ပက်|ပ)',
            line
        )

        if paat_match:

            digit = paat_match.group(1)

            amount_match = re.search(
                r'(\d+)(?:tt)?$',
                line,
                re.IGNORECASE
            )

            amount = (
                int(amount_match.group(1))
                if amount_match else 1000
            )

            for i in range(100):

                number = f"{i:02d}"

                if digit in number:

                    rows.append({
                        "Number": number,
                        "Amount": amount
                    })

            continue

        # 4. Multi R: 46/96r1500

        match = re.search(
            r'^([\d./,\-\s]+?)\s*[Rr]\s*(\d+)$',
            line
        )

        if match:

            numbers = re.findall(
                r'\d{1,2}',
                match.group(1)
            )

            amount = int(match.group(2))

            for n in numbers:

                number = n.zfill(2)

                rows.append({
                    "Number": number,
                    "Amount": amount
                })

                reverse = number[::-1]

                if reverse != number:
                    rows.append({
                        "Number": reverse,
                        "Amount": amount
                    })

            continue

        # 5. Slash / Star / Equals:
        # 11/99/77*3000 or 99=250

        match = re.search(
            r'^([\d./,\-\s]+?)\s*[*=]\s*(\d+)$',
            line
        )

        if match:

            numbers = re.findall(
                r'\d{1,2}',
                match.group(1)
            )

            amount = int(match.group(2))

            for n in numbers:

                rows.append({
                    "Number": n.zfill(2),
                    "Amount": amount
                })

            continue

        # 6. Advanced @: 20@3000

        match = re.search(
            r'^(\d{1,2})\s*@\s*(\d+)$',
            line
        )

        if match:

            number = match.group(1).zfill(2)
            amount = int(match.group(2))

            for n in set([number, number[::-1]]):

                rows.append({
                    "Number": n,
                    "Amount": amount
                })

            continue

        # 7. Dot R: 85.3000.r.500

        match = re.search(
            r'^(\d{1,2})[.-](\d+)[.-][Rr][.-](\d+)$',
            line
        )

        if match:

            number = match.group(1).zfill(2)
            base_amount = int(match.group(2))
            reverse_amount = int(match.group(3))

            rows.append({
                "Number": number,
                "Amount": base_amount
            })

            rows.append({
                "Number": number[::-1],
                "Amount": reverse_amount
            })

            continue

        # 8. Spaced Numbers: 44 99 11

        amount_match = re.search(
            r'(\d{3,6})$',
            line
        )

        clean_line = line

        if amount_match:
            amount = int(amount_match.group(1))
            clean_line = line[:amount_match.start()]
        else:
            amount = 1000

        if (
            re.fullmatch(r'[\d\s]+', clean_line)
            and len(clean_line.strip()) <= 15
            and re.search(r'\d{1,2}', clean_line)
        ):

            numbers = re.findall(
                r'\d{1,2}',
                clean_line
            )

            for n in numbers:

                rows.append({
                    "Number": n.zfill(2),
                    "Amount": amount
                })

            continue

        # 9. General Number-Amount Pairs

        pairs = re.findall(
            r'(\d{1,2})\s*[-=/_.]\s*(\d+)',
            line
        )

        if pairs:

            for number, amount in pairs:

                number = number.zfill(2)
                amount = int(amount)

                if (
                    0 <= int(number) <= 99
                    and 100 <= amount <= 500000
                ):

                    rows.append({
                        "Number": number,
                        "Amount": amount
                    })

    return rows


# ==================================================
# 5. PROCESS BETS
# ==================================================

def process_and_add_bets(
    extracted_data,
    entry_date_str,
    selected_day,
    customer
):

    valid_rows = []
    blocked_found = []
    invalid_amt_found = 0
    limit_found = []

    existing = st.session_state.data[
        st.session_state.data["Customer"] == customer
    ]["BatchID"]

    batch_numbers = []

    for batch in existing.unique():

        match = re.search(r'\((\d+)\)', str(batch))

        if match:
            batch_numbers.append(int(match.group(1)))

    next_batch = max(batch_numbers, default=0) + 1
    batch_id = f"{customer}({next_batch})"

    # Existing accepted amounts per number
    current_totals = {}

    for _, row in st.session_state.data.iterrows():

        number = str(row["Number"]).zfill(2)

        current_totals[number] = (
            current_totals.get(number, 0)
            + int(row["Amount"])
        )

    for row in extracted_data:

        amount = int(row.get("Amount", 0))
        number = str(row.get("Number", "")).zfill(2)

        if not (100 <= amount <= 500000):
            invalid_amt_found += 1
            continue

        if number in blocked_numbers:
            blocked_found.append(number)
            continue

        current_amount = current_totals.get(number, 0)

        if current_amount + amount > max_limit:
            limit_found.append(number)
            continue

        valid_rows.append({
            "Date": entry_date_str,
            "Day": selected_day,
            "Customer": customer,
            "BatchID": batch_id,
            "Number": number,
            "Amount": amount
        })

        current_totals[number] = current_amount + amount

    total_amount = sum(
        row["Amount"] for row in valid_rows
    )

    return (
        valid_rows,
        blocked_found,
        invalid_amt_found,
        total_amount,
        limit_found
    )


# ==================================================
# 6. EXCEL REPORT GENERATOR
# ==================================================

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

    ws = wb.active
    ws.title = "စာရင်းချုပ်အကျဉ်း"

    ws["A1"] = "မြန်မာနှစ်လုံးထီ (2D) ငွေစာရင်းချုပ်"
    ws["A1"].font = font_title

    ws["A2"] = (
        "ထုတ်ပြန်သည့်ရက်စွဲ - "
        + datetime.now().strftime("%Y-%m-%d %H:%M")
    )

    ws["A2"].font = font_body

    total_sales = float(df["Amount"].sum())

    commission = total_sales * commission_pct / 100

    net_sales = total_sales - commission

    win_number = (
        str(winning_number).strip().zfill(2)
        if winning_number else ""
    )

    winning_bets = (
        float(df[df["Number"] == win_number]["Amount"].sum())
        if win_number else 0
    )

    payout = winning_bets * payout_rate

    net_profit = net_sales - payout

    summary_rows = [
        ("အကြောင်းအရာ", "ပမာဏ / တန်ဖိုး", "မှတ်ချက်"),
        ("စုစုပေါင်း ရောင်းရငွေ", total_sales, f"{len(df):,} ကွက်"),
        (f"ကော်မရှင် ({commission_pct}%)", commission, "ကော်မရှင်နှုတ်ငွေ"),
        ("ကော်မရှင်နှုတ်ပြီး ဒိုင်ရငွေ", "=B5-B6", "Net Sales"),
        ("ထွက်ဂဏန်း", win_number or "မသတ်မှတ်ရသေး", f"ဆုကြေး {payout_rate} ဆ"),
        ("ပေါက်ဂဏန်းပေါ် ထိုးငွေ", winning_bets, "Winning Bets"),
        ("ပေးရမည့် ဆုကြေးငွေ", f"=B9*{payout_rate}", "Total Payout"),
        ("ဒိုင် အသားတင် အမြတ် / အရှုံး", "=B7-B10", "Net Profit/Loss")
    ]

    for idx, values in enumerate(summary_rows, start=4):

        for col, value in enumerate(values, start=1):

            cell = ws.cell(
                row=idx,
                column=col,
                value=value
            )

            cell.border = thin_border

            if idx == 4:
                cell.fill = header_fill
                cell.font = font_header
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            else:
                cell.font = (
                    font_bold if idx in (7, 11)
                    else font_body
                )

                if idx == 7:
                    cell.fill = accent_fill

                if idx == 11:
                    cell.fill = (
                        profit_fill if net_profit >= 0
                        else loss_fill
                    )

                if col == 2 and idx not in (8,):
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(
                        horizontal="right"
                    )

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 38

    # Detailed transaction sheet

    detail = wb.create_sheet("အသေးစိတ်စာရင်း")

    for row in openpyxl.utils.dataframe.dataframe_to_rows(
        df,
        index=False,
        header=True
    ):
        detail.append(row)

    for cell in detail[1]:
        cell.fill = header_fill
        cell.font = font_header
        cell.border = thin_border

    for row in detail.iter_rows(min_row=2):

        for cell in row:
            cell.font = font_body
            cell.border = thin_border

    for column in detail.columns:

        letter = column[0].column_letter

        detail.column_dimensions[letter].width = 20

    wb.save(output)

    return output.getvalue()


# ==================================================
# 7. BET ENTRY INTERFACE
# ==================================================

st.subheader(
    "📝 စာရင်းသွင်းရန် နည်းလမ်းများ "
    "(Ultimate Precision Tracker API-Free)"
)

col_d1, col_d2 = st.columns(2)

with col_d1:

    entry_date = st.date_input(
        "📆 စာရင်းသွင်းမည့် ရက်စွဲ",
        value=datetime.now().date()
    )

with col_d2:

    default_day_idx = min(entry_date.weekday(), 5)

    day_options = [
        "MONDAY", "TUESDAY", "WEDNESDAY",
        "THURSDAY", "FRIDAY", "WEEKEND"
    ]

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
        placeholder="ဥပမာ - မကြီးခင်လေး"
    )

with col2:

    smart_input_text = st.text_area(
        "ဖောက်သည် ပို့လိုက်သော စာရင်းစာသား",
        placeholder=(
            "44 99 11\n"
            "11000နေ့ရှင်း\n"
            "11/99/77*3000\n"
            "46/96r1500\n"
            "7ပက်2500"
        ),
        height=130
    )

if st.button("➕ စာရင်းသွင်းမည်", type="primary"):

    if not cust_manual or not smart_input_text:

        st.warning("ဖောက်သည်အမည်နှင့် ထိုးမည့်စာသား ဖြည့်ပါ။")

    elif (
        enable_time_lock
        and datetime.now().time() >= closing_time
    ):

        st.error("⛔ စာရင်းပိတ်ချိန် ကျော်လွန်သွားပါပြီ။")

    else:

        try:

            customer = cust_manual.strip()

            extracted = parse_bets_locally(
                smart_input_text
            )

            (
                valid_rows,
                blocked_found,
                invalid_amt,
                total_amt,
                limit_found
            ) = process_and_add_bets(
                extracted,
                entry_date_str,
                selected_day,
                customer
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
                    f"🎉 စာရင်းသွင်းပြီးပါပြီ!\n"
                    f"ဖောက်သည်: {customer}\n"
                    f"စုစုပေါင်း: {total_amt:,} ကျပ်\n"
                    f"အကွက်ရေ: {len(valid_rows)}"
                )

                st.rerun()

            else:

                st.warning(
                    "စာရင်းသွင်းနိုင်သော ဂဏန်းမတွေ့ပါ။"
                )

            if blocked_found:
                st.warning(
                    f"ပိတ်ဂဏန်းများ: {blocked_found}"
                )

            if limit_found:
                st.warning(
                    f"Limit ကျော်နေသော ဂဏန်းများ: {limit_found}"
                )

            if invalid_amt:
                st.warning(
                    f"ငွေပမာဏမမှန်သော အရေအတွက်: {invalid_amt}"
                )

        except Exception as e:

            st.error(
                f"စာရင်းသွင်းရာတွင် Error ဖြစ်နေပါသည်: {e}"
            )


# ==================================================
# 8. SALES SUMMARY & VIBER STATEMENT
# ==================================================

st.markdown("---")

if not st.session_state.data.empty:

    df = st.session_state.data.copy()

    total_sales = df["Amount"].sum()

    commission_amt = (
        total_sales * commission_pct / 100
    )

    net_sales = total_sales - commission_amt

    st.subheader(
        "📊 နေ့အလိုက် / တစ်ပတ်စာ စာရင်းချုပ် "
        "နှင့် ဖောက်သည် Viber Statement Generator"
    )

    m1, m2, m3 = st.columns(3)

    m1.metric(
        "စုစုပေါင်း ရောင်းရငွေ",
        f"{total_sales:,.0f} ကျပ်"
    )

    m2.metric(
        "ကော်မရှင်",
        f"{commission_amt:,.0f} ကျပ်"
    )

    m3.metric(
        "ကော်မရှင်နှုတ်ပြီး",
        f"{net_sales:,.0f} ကျပ်"
    )

    col_rep1, col_rep2 = st.columns(2)

    with col_rep1:

        st.markdown("##### 📊 အသေးစိတ်စာရင်း")

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        excel_file = generate_styled_excel(
            df,
            commission_pct,
            payout_rate,
            winning_number,
            max_limit
        )

        st.download_button(
            label="📥 Excel စာရင်းချုပ် Download",
            data=excel_file,
            file_name="Myanmar_2D_Master_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_rep2:

        st.markdown(
            "##### 👤 ဖောက်သည်တစ်ဦးချင်းစီ၏ Statement"
        )

        customers = df["Customer"].unique()

        selected_customer = st.selectbox(
            "ဖောက်သည် ရွေးချယ်ပါ",
            customers
        )

        if selected_customer:

            cust_df = df[
                df["Customer"] == selected_customer
            ]

            batches = cust_df["BatchID"].unique()

            viber_lines = [
                f"📒 {selected_customer} ၏ 2D အသေးစိတ်ရှင်းတမ်း",
                "=" * 28
            ]

            batch_totals = []

            for batch_id in batches:

                batch_df = cust_df[
                    cust_df["BatchID"] == batch_id
                ]

                batch_total = batch_df["Amount"].sum()

                batch_totals.append(
                    (batch_id, batch_total)
                )

                viber_lines.append(
                    f"--- {batch_id} ---"
                )

                for _, row in batch_df.iterrows():

                    viber_lines.append(
                        f"• {row['Number']} = "
                        f"{row['Amount']:,} ကျပ်"
                    )

                viber_lines.append(
                    f"{batch_id} စုစုပေါင်း = "
                    f"{batch_total:,.0f} ကျပ်\n"
                )

            customer_total = cust_df["Amount"].sum()

            formula = " + ".join(
                [
                    f"[{batch}={amount:,.0f} ကျပ်]"
                    for batch, amount in batch_totals
                ]
            )

            viber_lines.extend([
                "=" * 28,
                formula,
                f"စုစုပေါင်း = {customer_total:,.0f} ကျပ်",
                f"ကော်မရှင် ({commission_pct}%) = "
                f"{customer_total * commission_pct / 100:,.0f} ကျပ်",
                f"Net Sales = "
                f"{customer_total * (1 - commission_pct / 100):,.0f} ကျပ်",
                f"ပေါက်ဂဏန်းဆုကြေးနှုန်း = {payout_rate} ဆ",
                "=" * 28
            ])

            st.text_area(
                f"Viber သို့ ပို့ရန် ({selected_customer})",
                "\n".join(viber_lines),
                height=350
            )

    st.markdown("---")

    if st.button("🗑️ စာရင်းအားလုံး ဖျက်မည်"):

        st.session_state.data = pd.DataFrame(
            columns=[
                "Date", "Day", "Customer",
                "BatchID", "Number", "Amount"
            ]
        )

        st.rerun()

else:

    st.info("လက်ရှိတွင် စာရင်းသွင်းထားသော Data မရှိသေးပါ။")
