import streamlit as st
import pandas as pd
import json
import io
from PIL import Image, ImageOps
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from google import genai

st.set_page_config(page_title="Myanmar 2D Dealer App", page_icon="📒", layout="wide")

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

st.markdown('<div class="main-header"><h1>📒 မြန်မာနှစ်လုံးထီ (2D) အလိုအလျောက် စာရင်းချုပ်စနစ်</h1><p>Smart Ledger, Winning Checker & Advanced Viber Statement Generator</p></div>', unsafe_allow_html=True)

st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များနှင့် ကန့်သတ်ချက်များ")
max_limit = st.sidebar.number_input("ဂဏန်းတစ်လုံးလျှင် အများဆုံးလက်ခံငွေ (Limit)", value=50000, step=5000)
payout_rate = st.sidebar.number_input("ဆုကြေးနှုန်း (ဆ)", value=80, step=5)
commission_pct = st.sidebar.number_input("ကော်မရှင် (%)", value=10, step=1)
winning_number = st.sidebar.text_input("ထွက်မည့် ပေါက်ဂဏန်း", value="")

st.sidebar.markdown("---")
st.sidebar.subheader("⏰ အချိန်ပိတ်ရက် ထိန်းချုပ်မှု")
enable_time_lock = st.sidebar.checkbox("အချိန်ပိတ်ရန် ဖွင့်မည်", value=False)
closing_time = st.sidebar.time_input("ပိတ်မည့် အချိန်", value=datetime.strptime("12:00", "%H:%M").time())

st.sidebar.markdown("---")
st.sidebar.subheader("🚫 ပိတ်ဂဏန်းများ (Comma ခြားရန်)")
blocked_input = st.sidebar.text_input("မခံမည့် ဂဏန်းများ (ဥပမာ: 39, 88)", value="")
blocked_numbers = [num.strip() for num in blocked_input.split(",") if num.strip()]

st.sidebar.markdown("---")
st.sidebar.subheader("🤖 AI စနစ် ဆက်တင် (Auto-Fallback)")
available_models = [
    "gemini-4-hydrogen",
    "gemini-3.8-pro",
    "gemini-3.8-flash-lite",
    "gemini-3.8-flash",
    "Custom (ကိုယ်တိုင်ရိုက်ထည့်ရန်)"
]
selected_model_opt = st.sidebar.selectbox("ဦးစားပေး AI Model ရွေးချယ်ရန်", available_models, index=0)

if selected_model_opt == "Custom (ကိုယ်တိုင်ရိုက်ထည့်ရန်)":
    model_choice = st.sidebar.text_input("Model အမည် ရိုက်ထည့်ပါ", value="gemini-4-hydrogen")
else:
    model_choice = selected_model_opt

enable_fallback = st.sidebar.checkbox("Quota ပြည့်ပါက အခြား Model ဖြင့် အလိုအလျောက် ပြောင်းဖတ်မည်", value=True)
gemini_api_key = st.sidebar.text_input("Gemini API Key ထည့်ရန်", type="password", value="")

if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=["Date", "Day", "Customer", "Number", "Amount"])

def normalize_image(uploaded_file):
    img = Image.open(uploaded_file)
    img = ImageOps.exif_transpose(img)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    max_edge = 1600
    if max(img.size) > max_edge:
        img.thumbnail((max_edge, max_edge), Image.Resampling.LANCZOS)
    return img

SYSTEM_RULES = """
You are an expert AI assistant for a Myanmar 2D lottery business managing 100 numbers (00 to 99).
Extract and expand ALL bets into individual 2-digit numbers ("00" to "99") and their exact amounts in MMK based on these strict rules:
1. Reversed ('r', 'ရ', '@', '&', 'အာ'): E.g., '39@1000', '65r1000', '39ရ1000' means both the number and its reverse with the specified amount each. (If double like 33r1000, count once).
2. Khve / Permutation (အခွေ / ခွေ):
   - E.g., '269ခွေ 500' or '269အခွေ 500' means all 2-digit permutations WITHOUT doubles (26, 29, 62, 69, 92, 96 - 6 numbers) with 500 MMK each.
   - E.g., '269ခွေအပူးပါ 500' or '269အပူးအပါ 500' means permutations PLUS doubles (26, 29, 62, 69, 92, 96, 22, 66, 99 - 9 numbers) with 500 MMK each.
3. Paat-Thie (ပါတ်သီး variations: 'ပါတ်', 'ပတ်', 'ပါ', 'ပ'): E.g., '2ပါ 500' or '3ပါတ် 250' means all 19 numbers containing that digit.
4. Break (ဘရိတ် variations: 'ဘရိတ်', 'B', 'bk'): E.g., '2ဘရိတ် 300', '2B 300', '2bk 300' means all 10 numbers whose digits sum ends in 2 (02, 11, 20, 39, 93, 48, 84, 57, 75, 66).
5. Ah-Pu (အပူ): E.g., 'အပူ 500' means all 10 double numbers (00, 11, 22, 33, 44, 55, 66, 77, 88, 99).
6. Nakhat (နက္ခတ်): 07, 70, 18, 81, 24, 42, 35, 53, 69, 96.
7. Direct pairs with amounts: E.g., '12=1000', '48_3000', '50.25' (50 -> 250 or exact amount written).
Ensure each extracted bet amount is strictly between 100 and 50000 MMK.
Return the result strictly as a raw JSON list of objects with keys: "Customer", "Number" (2-digit string "00"-"99"), and "Amount" (integer). No markdown formatting.
"""

def call_gemini_with_fallback(client, primary_model, contents):
    fallback_candidates = [
        primary_model,
        "gemini-4-hydrogen",
        "gemini-3.8-pro",
        "gemini-3.8-flash-lite",
        "gemini-3.8-flash"
    ]
    # Remove duplicates while preserving order
    models_to_try = list(dict.fromkeys([m.strip() for m in fallback_candidates if m.strip()]))
    if not enable_fallback:
        models_to_try = [primary_model.strip()]

    last_error = None
    for m_name in models_to_try:
        try:
            response = client.models.generate_content(model=m_name, contents=contents)
            return response, m_name
        except Exception as e:
            last_error = e
            err_str = str(e)
            if any(code in err_str for code in ["429", "503", "404", "RESOURCE_EXHAUSTED", "UNAVAILABLE", "NOT_FOUND"]):
                continue
            else:
                raise e
    raise last_error

def process_and_add_bets(extracted_data, entry_date_str, selected_day, default_customer):
    valid_rows = []
    blocked_found = []
    invalid_amt_found = 0
    total_added_amount = 0
    
    for row in extracted_data:
        amt_val = int(row.get("Amount", 0))
        num_str = str(row.get("Number", "")).zfill(2)
        cust_name = row.get("Customer") if row.get("Customer") and row.get("Customer") != "Unknown" else default_customer
        
        if amt_val < 100 or amt_val > 50000:
            invalid_amt_found += 1
            continue
        if num_str in blocked_numbers:
            blocked_found.append(num_str)
        else:
            valid_rows.append({
                "Date": entry_date_str,
                "Day": selected_day,
                "Customer": cust_name,
                "Number": num_str,
                "Amount": amt_val
            })
            total_added_amount += amt_val
            
    return valid_rows, blocked_found, invalid_amt_found, total_added_amount

def generate_styled_excel(df, commission_pct, payout_rate, winning_number, max_limit):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    accent_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    profit_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    loss_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    
    font_title = Font(name="Pyidaungsu", size=14, bold=True, color="1E293B")
    font_header = Font(name="Pyidaungsu", size=11, bold=True, color="FFFFFF")
    font_body = Font(name="Pyidaungsu", size=11, color="0F172A")
    font_bold = Font(name="Pyidaungsu", size=11, bold=True, color="0F172A")
    
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    def format_sheet(ws, col_widths):
        ws.views.sheetView[0].showGridLines = True
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

    win_num_clean = str(winning_number).strip().zfill(2) if winning_number else ""

    # TAB 1: စာရင်းချုပ်အကျဉ်း
    ws_sum = wb.active
    ws_sum.title = "စာရင်းချုပ်အကျဉ်း"
    ws_sum["A1"] = "မြန်မာနှစ်လုံးထီ (2D) ငွေစာရင်းချုပ် အစီရင်ခံစာ"
    ws_sum["A1"].font = font_title
    ws_sum["A2"] = f"ထုတ်ပြန်သည့်ရက်စွဲ - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws_sum["A2"].font = font_body

    total_sales = float(df["Amount"].sum())
    comm_amount = total_sales * (commission_pct / 100.0)
    net_sales = total_sales - comm_amount
    
    win_bet_total = float(df[df["Number"] == win_num_clean]["Amount"].sum()) if win_num_clean else 0.0
    payout_total = win_bet_total * payout_rate
    net_pl = net_sales - payout_total

    summary_rows = [
        ("အကြောင်းအရာ", "ပမာဏ / တန်ဖိုး", "မှတ်ချက်"),
        ("စုစုပေါင်း ရောင်းရငွေ (Total Sales)", total_sales, f"စုစုပေါင်း အကွက်ရေ {len(df):,} ကွက်"),
        (f"ကော်မရှင် နှုတ်ငွေ ({commission_pct}%)", comm_amount, "ဒိုင်ကော်မရှင် နှုတ်ယူငွေ"),
        ("ကော်မရှင်နှုတ်ပြီး ဒိုင်ရငွေ (Net Sales)", "=B5-B6", "အသားတင် ရငွေ"),
        ("ထွက်ဂဏန်း (Winning Number)", win_num_clean if win_num_clean else "မသတ်မှတ်ရသေး", f"ဆုကြေးနှုန်း {payout_rate} ဆ"),
        ("ပေါက်ဂဏန်းပေါ် ထိုးငွေစုစုပေါင်း", win_bet_total, "ပေါက်သီး တန်ဖိုး"),
        ("ပေးရမည့် ဆုကြေးငွေစုစုပေါင်း (Total Payout)", f"=B9*{payout_rate}", "ကံထူးရှင်များသို့ ပေးလျော်ငွေ"),
        ("ဒိုင် အသားတင် အနိုင် / အရှုံး (Net Profit/Loss)", "=B7-B10", "အပေါင်း (+) ဖြစ်ပါက အမြတ် / အနှုတ် (-) ဖြစ်ပါက အရှုံး")
    ]

    start_row = 4
    for idx, (label, val, note) in enumerate(summary_rows):
        r = start_row + idx
        c1 = ws_sum.cell(row=r, column=1, value=label)
        c2 = ws_sum.cell(row=r, column=2, value=val)
        c3 = ws_sum.cell(row=r, column=3, value=note)
        
        for c in (c1, c2, c3):
            c.border = thin_border
            if idx == 0:
                c.fill = header_fill
                c.font = font_header
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.font = font_bold if idx in (3, 7) else font_body
                if idx == 3:
                    c.fill = accent_fill
                elif idx == 7:
                    c.fill = profit_fill if net_pl >= 0 else loss_fill
        if idx > 0 and idx != 4:
            c2.number_format = "#,##0"
            c2.alignment = Alignment(horizontal="right")
        elif idx == 4:
            c2.alignment = Alignment(horizontal="center")

    format_sheet(ws_sum, {"A": 42, "B": 24, "C": 38})

    # TAB 2: နေ့အလိုက်_အပတ်စဉ်
    ws_day = wb.create_sheet(title="နေ့အလိုက်_အပတ်စဉ်")
    ws_day["A1"] = "နေ့အလိုက် နှင့် တစ်ပတ်စာ စာရင်းချုပ် (Daily & Weekly Breakdown)"
    ws_day["A1"].font = font_title

    day_order = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "WEEKEND"]
    day_grp = df.groupby(["Date", "Day"]).agg(
        Total_Bets=("Number", "count"),
        Gross_Sales=("Amount", "sum")
    ).reset_index()
    day_grp["Day"] = pd.Categorical(day_grp["Day"], categories=day_order, ordered=True)
    day_grp = day_grp.sort_values(["Date", "Day"])

    headers_day = ["ရက်စွဲ (Date)", "နေ့ (Day)", "အကွက်ရေ", "စုစုပေါင်းထိုးကြေး", f"ကော်မရှင် ({commission_pct}%)", "ဒိုင်ရငွေ (Net)"]
    for col_idx, h in enumerate(headers_day, 1):
        cell = ws_day.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    for r_idx, row in enumerate(day_grp.itertuples(), 4):
        ws_day.cell(row=r_idx, column=1, value=row.Date).alignment = Alignment(horizontal="center")
        ws_day.cell(row=r_idx, column=2, value=str(row.Day)).alignment = Alignment(horizontal="center")
        ws_day.cell(row=r_idx, column=3, value=int(row.Total_Bets)).number_format = "#,##0"
        ws_day.cell(row=r_idx, column=4, value=float(row.Gross_Sales)).number_format = "#,##0"
        ws_day.cell(row=r_idx, column=5, value=f"=D{r_idx}*{commission_pct/100.0}").number_format = "#,##0"
        ws_day.cell(row=r_idx, column=6, value=f"=D{r_idx}-E{r_idx}").number_format = "#,##0"
        for c_i in range(1, 7):
            c = ws_day.cell(row=r_idx, column=c_i)
            c.font = font_body
            c.border = thin_border
            if r_idx % 2 == 1:
                c.fill = zebra_fill

    last_day_row = 3 + len(day_grp)
    tot_r = last_day_row + 1
    ws_day.cell(row=tot_r, column=1, value="စုစုပေါင်း (Total)").font = font_bold
    ws_day.cell(row=tot_r, column=3, value=f"=SUM(C4:C{last_day_row})").number_format = "#,##0"
    ws_day.cell(row=tot_r, column=4, value=f"=SUM(D4:D{last_day_row})").number_format = "#,##0"
    ws_day.cell(row=tot_r, column=5, value=f"=SUM(E4:E{last_day_row})").number_format = "#,##0"
    ws_day.cell(row=tot_r, column=6, value=f"=SUM(F4:F{last_day_row})").number_format = "#,##0"
    for c_i in range(1, 7):
        c = ws_day.cell(row=tot_r, column=c_i)
        c.font = font_bold
        c.fill = accent_fill
        c.border = thin_border

    format_sheet(ws_day, {"A": 18, "B": 16, "C": 14, "D": 22, "E": 20, "F": 22})

    # TAB 3: ဖောက်သည်အလိုက် ရှင်းတမ်း
    ws_cust = wb.create_sheet(title="ဖောက်သည်အလိုက် ရှင်းတမ်း")
    ws_cust["A1"] = "ဖောက်သည်တစ်ဦးချင်းစီ၏ ထိုးကြေး၊ ကော်မရှင်နှင့် ဆုကြေးရှင်းတမ်း"
    ws_cust["A1"].font = font_title

    cust_summary_df = df.groupby("Customer")["Amount"].sum().reset_index().sort_values(by="Amount", ascending=False)
    
    headers_cust = ["ဖောက်သည် အမည်", "စုစုပေါင်းထိုးကြေး", f"ကော်မရှင် ({commission_pct}%)", "ကော်မရှင်နှုတ်ပြီး ဒိုင်ရငွေ", f"ပေါက်ဂဏန်းဆုကြေး ({payout_rate}ဆ)", "အသားတင် ရှင်းတမ်း (Net)"]
    for col_idx, h in enumerate(headers_cust, 1):
        cell = ws_cust.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    for r_idx, row in enumerate(cust_summary_df.itertuples(), 4):
        c_name = row.Customer
        cust_win_amt = 0.0
        if win_num_clean:
            match_subset = df[(df["Customer"] == c_name) & (df["Number"] == win_num_clean)]
            cust_win_amt = float(match_subset["Amount"].sum() * payout_rate)

        ws_cust.cell(row=r_idx, column=1, value=c_name)
        ws_cust.cell(row=r_idx, column=2, value=float(row.Amount)).number_format = "#,##0"
        ws_cust.cell(row=r_idx, column=3, value=f"=B{r_idx}*{commission_pct/100.0}").number_format = "#,##0"
        ws_cust.cell(row=r_idx, column=4, value=f"=B{r_idx}-C{r_idx}").number_format = "#,##0"
        ws_cust.cell(row=r_idx, column=5, value=cust_win_amt).number_format = "#,##0"
        ws_cust.cell(row=r_idx, column=6, value=f"=D{r_idx}-E{r_idx}").number_format = "#,##0"

        for c_i in range(1, 7):
            c = ws_cust.cell(row=r_idx, column=c_i)
            c.font = font_body
            c.border = thin_border
            if r_idx % 2 == 1:
                c.fill = zebra_fill

    last_cust_row = 3 + len(cust_summary_df)
    tot_c_r = last_cust_row + 1
    ws_cust.cell(row=tot_c_r, column=1, value="စုစုပေါင်း (Total)").font = font_bold
    ws_cust.cell(row=tot_c_r, column=2, value=f"=SUM(B4:B{last_cust_row})").number_format = "#,##0"
    ws_cust.cell(row=tot_c_r, column=3, value=f"=SUM(C4:C{last_cust_row})").number_format = "#,##0"
    ws_cust.cell(row=tot_c_r, column=4, value=f"=SUM(D4:D{last_cust_row})").number_format = "#,##0"
    ws_cust.cell(row=tot_c_r, column=5, value=f"=SUM(E4:E{last_cust_row})").number_format = "#,##0"
    ws_cust.cell(row=tot_c_r, column=6, value=f"=SUM(F4:F{last_cust_row})").number_format = "#,##0"
    for c_i in range(1, 7):
        c = ws_cust.cell(row=tot_c_r, column=c_i)
        c.font = font_bold
        c.fill = accent_fill
        c.border = thin_border

    format_sheet(ws_cust, {"A": 26, "B": 22, "C": 20, "D": 24, "E": 24, "F": 24})

    # TAB 4: လစဉ်စာရင်းချုပ်
    ws_mon = wb.create_sheet(title="လစဉ်စာရင်းချုပ်")
    ws_mon["A1"] = "လစဉ် စာရင်းချုပ်အကျဉ်း (Monthly Summary)"
    ws_mon["A1"].font = font_title

    df_mon = df.copy()
    df_mon["Month"] = df_mon["Date"].astype(str).str.slice(0, 7)
    mon_grp = df_mon.groupby("Month").agg(
        Total_Bets=("Number", "count"),
        Gross_Sales=("Amount", "sum")
    ).reset_index()

    headers_mon = ["လ (Year-Month)", "စုစုပေါင်း အကွက်ရေ", "စုစုပေါင်း ရောင်းရငွေ", f"ကော်မရှင် ({commission_pct}%)", "ဒိုင်အသားတင်ရငွေ"]
    for col_idx, h in enumerate(headers_mon, 1):
        cell = ws_mon.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    for r_idx, row in enumerate(mon_grp.itertuples(), 4):
        ws_mon.cell(row=r_idx, column=1, value=row.Month).alignment = Alignment(horizontal="center")
        ws_mon.cell(row=r_idx, column=2, value=int(row.Total_Bets)).number_format = "#,##0"
        ws_mon.cell(row=r_idx, column=3, value=float(row.Gross_Sales)).number_format = "#,##0"
        ws_mon.cell(row=r_idx, column=4, value=f"=C{r_idx}*{commission_pct/100.0}").number_format = "#,##0"
        ws_mon.cell(row=r_idx, column=5, value=f"=C{r_idx}-D{r_idx}").number_format = "#,##0"
        for c_i in range(1, 6):
            c = ws_mon.cell(row=r_idx, column=c_i)
            c.font = font_body
            c.border = thin_border

    format_sheet(ws_mon, {"A": 20, "B": 20, "C": 24, "D": 22, "E": 24})

    # TAB 5: ဂဏန်းပေါင်းချုပ်_Limit
    ws_num = wb.create_sheet(title="ဂဏန်းပေါင်းချုပ်_Limit")
    ws_num["A1"] = "ဂဏန်းအလိုက် စုစုပေါင်းထိုးကြေးနှင့် Limit ထိန်းချုပ်မှု"
    ws_num["A1"].font = font_title

    num_grp = df.groupby("Number")["Amount"].sum().reset_index().sort_values(by="Amount", ascending=False)
    headers_num = ["2D ဂဏန်း", "စုစုပေါင်း ထိုးငွေ", "ဒိုင်လက်ခံငွေ (Allowed)", "ကျော်လွန်ငွေ (Excess/ကာရန်)"]
    for col_idx, h in enumerate(headers_num, 1):
        cell = ws_num.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    for r_idx, row in enumerate(num_grp.itertuples(), 4):
        c_num = ws_num.cell(row=r_idx, column=1, value=str(row.Number).zfill(2))
        c_num.alignment = Alignment(horizontal="center")
        c_num.number_format = "@"
        ws_num.cell(row=r_idx, column=2, value=float(row.Amount)).number_format = "#,##0"
        ws_num.cell(row=r_idx, column=3, value=f"=MIN(B{r_idx},{max_limit})").number_format = "#,##0"
        ws_num.cell(row=r_idx, column=4, value=f"=B{r_idx}-C{r_idx}").number_format = "#,##0"
        for c_i in range(1, 5):
            c = ws_num.cell(row=r_idx, column=c_i)
            c.font = font_body
            c.border = thin_border
            if float(row.Amount) > max_limit and c_i == 4:
                c.fill = loss_fill
                c.font = font_bold

    format_sheet(ws_num, {"A": 16, "B": 24, "C": 24, "D": 26})

    # TAB 6: အသေးစိတ်စာရင်း
    ws_det = wb.create_sheet(title="အသေးစိတ်စာရင်း")
    ws_det["A1"] = "ဖောက်သည်နှင့် ဂဏန်းအလိုက် အသေးစိတ် စာရင်းမှတ်တမ်း"
    ws_det["A1"].font = font_title

    det_grp = df.groupby(["Date", "Day", "Customer", "Number"])["Amount"].sum().reset_index()
    headers_det = ["ရက်စွဲ (Date)", "နေ့ (Day)", "ဖောက်သည် အမည်", "2D ဂဏန်း", "ထိုးငွေ (MMK)"]
    for col_idx, h in enumerate(headers_det, 1):
        cell = ws_det.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border

    for r_idx, row in enumerate(det_grp.itertuples(), 4):
        ws_det.cell(row=r_idx, column=1, value=row.Date).alignment = Alignment(horizontal="center")
        ws_det.cell(row=r_idx, column=2, value=row.Day).alignment = Alignment(horizontal="center")
        ws_det.cell(row=r_idx, column=3, value=row.Customer)
        c_n = ws_det.cell(row=r_idx, column=4, value=str(row.Number).zfill(2))
        c_n.alignment = Alignment(horizontal="center")
        c_n.number_format = "@"
        ws_det.cell(row=r_idx, column=5, value=float(row.Amount)).number_format = "#,##0"
        for c_i in range(1, 6):
            c = ws_det.cell(row=r_idx, column=c_i)
            c.font = font_body
            c.border = thin_border
            if r_idx % 2 == 1:
                c.fill = zebra_fill

    format_sheet(ws_det, {"A": 16, "B": 16, "C": 26, "D": 16, "E": 22})

    wb.save(output)
    return output.getvalue()

st.subheader("📝 စာရင်းသွင်းရန် နည်းလမ်းများ")

is_closed = False
if enable_time_lock:
    current_time = datetime.now().time()
    if current_time >= closing_time:
        is_closed = True
        st.error(f"⚠️ သတိပေးချက် - စာရင်းပိတ်ချိန် ({closing_time.strftime('%H:%M')}) ကျော်လွန်သွားပြီဖြစ်၍ စာရင်းအသစ် လက်မခံတော့ပါ။")

col_d1, col_d2 = st.columns(2)
with col_d1:
    entry_date = st.date_input("📆 စာရင်းသွင်းမည့် ရက်စွဲ (Date)", value=datetime.now().date())
with col_d2:
    default_day_idx = min(entry_date.weekday(), 5)
    day_options = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "WEEKEND"]
    selected_day = st.selectbox("📅 စာရင်းသွင်းမည့် နေ့ (Day)", day_options, index=default_day_idx)

entry_date_str = entry_date.strftime("%Y-%m-%d")

tab1, tab2, tab3 = st.tabs([
    "✍️ ဂဏန်းနှင့်ငွေ ရောနှောရိုက်ထည့်ရန် (Smart Input)", 
    "💬 ဖောက်သည်ပို့သော စာသား (SMS) ဖြင့် AI ဖတ်ရန်", 
    "📸 စကရင်ရှော့ ပုံတင်၍ AI ဖတ်ရန်"
])

with tab1:
    st.markdown("#### 💡 ထိုးမည့်ဂဏန်းနှင့် ငွေပမာဏကို တစ်ခါတည်း ရောနှောရိုက်ထည့်ပါ")
    col1, col2 = st.columns([1, 2])
    with col1:
        cust_manual = st.text_input("ဖောက်သည် အမည်", key="c_man", placeholder="ဥပမာ - Sandar")
    with col2:
        smart_input_text = st.text_input("ထိုးသည့် 2D ဂဏန်းနှင့် ငွေပမာဏ (ရောနှောလျက်)", key="s_input", placeholder="12@1000 10@1000 48_3000")

    if st.button("➕ အလိုအလျောက် တွက်ချက်၍ စာရင်းသွင်းမည်", type="primary"):
        if is_closed:
            st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
        elif not cust_manual or not smart_input_text:
            st.warning("ဖောက်သည် အမည်နှင့် ထိုးမည့် စာသားကို ဖြည့်သွင်းပါ။")
        elif not gemini_api_key:
            st.error("ကျေးဇူးပြု၍ ဘယ်ဘက်မီနူး (Sidebar) တွင် Gemini API Key ထည့်သွင်းပေးပါ။")
        else:
            with st.spinner("စာရင်းများကို တွက်ချက်နေပါသည်..."):
                try:
                    client = genai.Client(api_key=gemini_api_key)
                    prompt = f"{SYSTEM_RULES}\nCustomer Name: '{cust_manual}'\nInput text: {smart_input_text}"
                    response, used_model = call_gemini_with_fallback(client, model_choice, prompt)
                    
                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                    extracted_data = json.loads(clean_text)
                    valid_rows, blocked_found, invalid_amt, total_amt = process_and_add_bets(extracted_data, entry_date_str, selected_day, cust_manual)
                    
                    if valid_rows:
                        st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame(valid_rows)], ignore_index=True)
                        st.success(f"🎉 စာရင်းသွင်းပြီးပါပြီ! ({selected_day} - {cust_manual}) စုစုပေါင်းထိုးငွေ = {total_amt:,} ကျပ် ({len(valid_rows)} ကွက်) | Model: {used_model}")
                        if blocked_found:
                            st.warning(f"🚫 ပိတ်ဂဏန်းများ ဖယ်ထုတ်ထားပါသည်: {', '.join(set(blocked_found))}")
                        st.rerun()
                    else:
                        st.warning("သတ်မှတ်ချက်နှင့် ကိုက်ညီသော စာရင်းမတွေ့ပါ။")
                except Exception as e:
                    st.error(f"စာရင်းဖတ်ရာတွင် အမှားအယွင်း ရှိသွားပါသည်: {e}")

with tab2:
    st.markdown("#### 💬 ဖောက်သည်ပေးပို့ထားသော စာသား (Viber / Messenger SMS) ကို ကူးထည့်ပါ")
    sms_customer = st.text_input("ဖောက်သည် အမည် (စာသားအတွက်)", key="sms_cust", placeholder="ဥပမာ - မောင်မောင်")
    sms_text_input = st.text_area("ဖောက်သည် ပို့လိုက်သော စာသားမက်ဆေ့ဂျ် အပြည့်အစုံကို ကူးထည့်ပါ (Paste)", height=130)
    
    if st.button("🤖 စာသားကို AI ဖြင့် ဖြန့်ကျက်ဖတ်မည်", type="primary"):
        if is_closed:
            st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
        elif not gemini_api_key:
            st.error("ကျေးဇူးပြု၍ ဘယ်ဘက်မီနူး (Sidebar) တွင် Gemini API Key ထည့်သွင်းပေးပါ။")
        elif not sms_customer or not sms_text_input:
            st.error("ဖောက်သည် အမည်နှင့် စာသားကို ဖြည့်သွင်းပေးပါ။")
        else:
            with st.spinner("စာသားမက်ဆေ့ဂျ်များကို တွက်ချက်နေပါသည်..."):
                try:
                    client = genai.Client(api_key=gemini_api_key)
                    prompt = f"{SYSTEM_RULES}\nCustomer Name: '{sms_customer}'\nMessage text:\n{sms_text_input}"
                    response, used_model = call_gemini_with_fallback(client, model_choice, prompt)
                    
                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                    extracted_data = json.loads(clean_text)
                    valid_rows, blocked_found, invalid_amt, total_amt = process_and_add_bets(extracted_data, entry_date_str, selected_day, sms_customer)
                    
                    if valid_rows:
                        st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame(valid_rows)], ignore_index=True)
                        st.success(f"🎉 စာသားဖတ်ရှုပြီးပါပြီ! ({selected_day}) စုစုပေါင်းထိုးငွေ = {total_amt:,} ကျပ် ({len(valid_rows)} ကွက်) | Model: {used_model}")
                        if blocked_found:
                            st.warning(f"🚫 ပိတ်ဂဏန်းများ ဖယ်ထုတ်ထားပါသည်: {', '.join(set(blocked_found))}")
                        st.rerun()
                    else:
                        st.warning("သတ်မှတ်ချက်နှင့် ကိုက်ညီသော စာရင်းမတွေ့ပါ။")
                except Exception as e:
                    st.error(f"စာရင်းဖတ်ရာတွင် အမှားအယွင်း ရှိသွားပါသည်: {e}")

with tab3:
    st.markdown("#### 📸 စကရင်ရှော့ (သို့) ဖုန်းဓာတ်ပုံတင်၍ စာရင်းဖတ်ရန်")
    img_cust_fallback = st.text_input("ဖောက်သည် အမည် (ပုံထဲတွင် အမည်မပါပါက ထည့်ရန်)", value="Screenshot Customer", key="img_cust")
    uploaded_file = st.file_uploader("ဖောက်သည် ပို့ထားသော စကရင်ရှော့ ပုံရွေးချယ်ပါ", type=["png", "jpg", "jpeg"])
    
    if uploaded_file is not None:
        if uploaded_file.size > 8 * 1024 * 1024:
            st.error("⚠️ ဖိုင်အရွယ်အစား 8MB ထက် ကြီးနေပါသည်။ ပိုမိုသေးငယ်သော ပုံကို ပြန်လည်ရွေးချယ်ပေးပါ။")
        else:
            try:
                normalized_img = normalize_image(uploaded_file)
                col_img, col_meta = st.columns([1, 1])
                with col_img:
                    st.image(normalized_img, caption="စစ်ဆေးပြီး ပုံအကြိုကြည့်ရှုခြင်း (Ready)", use_container_width=True)
                with col_meta:
                    st.markdown("**ဖိုင်အချက်အလက် (Metadata):**")
                    st.write(f"• ဖိုင်အမည်: `{uploaded_file.name}`")
                    st.write(f"• ချိန်ညှိပြီး အရွယ်အစား: `{normalized_img.size[0]} x {normalized_img.size[1]} px`")
                    st.write(f"• သတ်မှတ်ရက်စွဲ: `{entry_date_str} ({selected_day})`")
                    
                    if st.button("🤖 ဓာတ်ပုံထဲမှ စာရင်းများကို AI ဖြင့် ဖတ်မည်", type="primary"):
                        if is_closed:
                            st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
                        elif not gemini_api_key:
                            st.error("ကျေးဇူးပြု၍ ဘယ်ဘက်မီနူး (Sidebar) တွင် Gemini API Key ထည့်သွင်းပေးပါ။")
                        else:
                            with st.spinner("ပုံထဲမှ စာရင်းများကို AI ဖြင့် ဖတ်ရှုတွက်ချက်နေပါသည်..."):
                                try:
                                    client = genai.Client(api_key=gemini_api_key)
                                    prompt = f"{SYSTEM_RULES}\nExtract customer name if visible in the screenshot, otherwise use '{img_cust_fallback}'."
                                    
                                    response, used_model = call_gemini_with_fallback(
                                        client,
                                        model_choice,
                                        [normalized_img, prompt]
                                    )
                                    
                                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                                    extracted_data = json.loads(clean_text)
                                    valid_rows, blocked_found, invalid_amt, total_amt = process_and_add_bets(extracted_data, entry_date_str, selected_day, img_cust_fallback)
                                    
                                    if valid_rows:
                                        st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame(valid_rows)], ignore_index=True)
                                        st.success(f"🎉 ပုံထဲမှ စာရင်းများကို အောင်မြင်စွာ ဖတ်ရှုပြီးပါပြီ! ({selected_day}) စုစုပေါင်းထိုးငွေ = {total_amt:,} ကျပ် ({len(valid_rows)} ကွက်) | Model: {used_model}")
                                        if blocked_found:
                                            st.warning(f"🚫 ပိတ်ဂဏန်းများ ဖယ်ထုတ်ထားပါသည်: {', '.join(set(blocked_found))}")
                                        st.rerun()
                                    else:
                                        st.warning("ပုံထဲတွင် သတ်မှတ်ချက်နှင့် ကိုက်ညီသော စာရင်းမတွေ့ပါ။ ပုံကြည်လင်မှု ရှိ/မရှိ ပြန်စစ်ပေးပါ။")
                                except Exception as e:
                                    st.error(f"စာရင်းဖတ်ရာတွင် အမှားအယွင်း ရှိသွားပါသည်: {e}")
            except Exception as img_err:
                st.error(f"ပုံဖိုင် ဖွင့်ရာတွင် အမှားရှိနေပါသည်: {img_err}")
    else:
        st.caption("ပုံမတက်သေးပါ — အပေါ်ပါခလုတ်မှ စကရင်ရှော့ပုံကို ရွေးချယ်နိုင်ပါသည်။")

st.markdown("---")
if not st.session_state.data.empty:
    df = st.session_state.data
    
    total_sales = df["Amount"].sum()
    commission_amt = total_sales * (commission_pct / 100)
    net_sales = total_sales - commission_amt
    total_bets = len(df)
    
    win_num_clean = str(winning_number).strip().zfill(2) if winning_number else ""
    payout_amount = 0
    winning_total_amt = 0
    if winning_number:
        winning_matched = df[df["Number"] == win_num_clean]
        winning_total_amt = winning_matched["Amount"].sum()
        payout_amount = winning_total_amt * payout_rate
    net_pl = net_sales - payout_amount

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("💰 စုစုပေါင်း ရောင်းရငွေ", f"{total_sales:,.0f} MMK")
    m2.metric(f"💸 ကော်မရှင် ({commission_pct}%)", f"{commission_amt:,.0f} MMK")
    m3.metric("📉 ကော်မရှင်နှုတ်ပြီး ဒိုင်ရငွေ", f"{net_sales:,.0f} MMK")
    m4.metric("📊 အသားတင် အနိုင် / အရှုံး", f"{net_pl:,.0f} MMK")
    
    st.markdown("---")
    
    # --- SECTION: Winning Checker & Statement Generator ---
    st.subheader("🎯 ပေါက်ဂဏန်း ကံထူးသူများ စာရင်း တိုက်စစ်ခြင်း")
    if winning_number:
        if not winning_matched.empty:
            st.success(f"🎉 ပေါက်ဂဏန်း ({win_num_clean}) အတွက် ကံထူးရှင်များ ရှိပါသည်။")
            st.dataframe(winning_matched, use_container_width=True)
            w_col1, w_col2 = st.columns(2)
            w_col1.metric("ပေးရမည့် ဆုကြေးငွေစုစုပေါင်း", f"{payout_amount:,.0f} MMK")
            w_col2.metric("ဒိုင် အမြတ် / အရှုံး (Net P/L)", f"{net_pl:,.0f} MMK")
        else:
            st.info(f"ပေါက်ဂဏန်း ({win_num_clean}) ထိုးထားသူ မရှိပါ။ ဒိုင်အမြတ် = {net_sales:,.0f} MMK")
    else:
        st.info("💡 ဘယ်ဘက် Sidebar တွင် 'ထွက်မည့် ပေါက်ဂဏန်း' ထည့်သွင်းပေးခြင်းဖြင့် ကံထူးရှင်စာရင်းနှင့် ဆုကြေးများကို အလိုအလျောက် တိုက်စစ်နိုင်ပါသည်။")

    st.markdown("---")
    st.subheader("📊 နေ့အလိုက် / တစ်ပတ်စာ စာရင်းချုပ်နှင့် ဖောက်သည် Viber Statement Generator")
    
    col_rep1, col_rep2 = st.columns([1, 1])
    
    with col_rep1:
        st.markdown("##### 📅 တစ်နေ့တာ / တစ်ပတ်စာ စုစုပေါင်း ရှင်းတမ်း")
        day_summary = df.groupby(["Date", "Day"])["Amount"].sum().reset_index()
        st.dataframe(day_summary, use_container_width=True)
        
        report_lines = ["📒 *မြန်မာ 2D စာရင်းချုပ်အစီရင်ခံစာ*", "=" * 28]
        for r in day_summary.itertuples():
            report_lines.append(f"📅 {r.Day} ({r.Date}) : {r.Amount:,.0f} ကျပ်")
        report_lines.append("-" * 28)
        report_lines.append(f"💰 စုစုပေါင်းရောင်းရငွေ = {total_sales:,.0f} ကျပ်")
        report_lines.append(f"💸 ကော်မရှင် ({commission_pct}%) = -{commission_amt:,.0f} ကျပ်")
        report_lines.append(f"✅ ဒိုင်ရငွေ (Net Sales) = {net_sales:,.0f} ကျပ်")
        if winning_number:
            report_lines.append(f"🎯 ပေါက်ဂဏန်း ({win_num_clean}) တန်ဖိုး = {winning_total_amt:,.0f} ကျပ်")
            report_lines.append(f"🏆 ပေးရမည့်ဆုငွေ ({payout_rate}ဆ) = -{payout_amount:,.0f} ကျပ်")
        status_label = "အမြတ် (+)" if net_pl >= 0 else "အရှုံး (-)"
        report_lines.append(f"📊 အသားတင် {status_label} = {net_pl:,.0f} ကျပ်")
        
        st.text_area("ဒိုင်ချုပ် Viber Text (Copy ယူရန်)", "\n".join(report_lines), height=140)

    with col_rep2:
        st.markdown("##### 👤 ဖောက်သည်တစ်ဦးချင်းစီအတွက် Viber Statement ထုတ်ရန်")
        unique_customers = df["Customer"].unique()
        selected_cust_stmt = st.selectbox("ဖောက်သည် ရွေးချယ်ပါ", unique_customers)
        
        if selected_cust_stmt:
            c_df = df[df["Customer"] == selected_cust_stmt]
            c_total = c_df["Amount"].sum()
            c_comm = c_total * (commission_pct / 100.0)
            c_net = c_total - c_comm
            
            c_win_amt = 0.0
            if winning_number:
                c_win_match = c_df[c_df["Number"] == win_num_clean]
                c_win_amt = float(c_win_match["Amount"].sum() * payout_rate)
            
            c_balance = c_net - c_win_amt

            cust_viber_lines = [
                f"📒 *{selected_cust_stmt}* ၏ 2D ထိုးငွေနှင့် ရှင်းတမ်း",
                "-" * 25
            ]
            for idx, row in c_df.iterrows():
                cust_viber_lines.append(f"• {row['Number']} = {row['Amount']:,} ကျပ် ({row['Day']})")
            cust_viber_lines.append("-" * 25)
            cust_viber_lines.append(f"စုစုပေါင်းထိုးကြေး = {c_total:,.0f} ကျပ်")
            cust_viber_lines.append(f"ကော်မရှင် ({commission_pct}%) = -{c_comm:,.0f} ကျပ်")
            cust_viber_lines.append(f"ပေးချေရန် Net = {c_net:,.0f} ကျပ်")
            if winning_number:
                cust_viber_lines.append(f"ပေါက်ဂဏန်းဆုကြေး ({win_num_clean}) = +{c_win_amt:,.0f} ကျပ်")
                bal_label = "ပေးရန် (+) / ရရန် (-)" if c_balance >= 0 else "ရရန်"
                cust_viber_lines.append(f"ရှင်းတမ်းစုစုပေါင်း ({bal_label}) = {c_balance:,.0f} ကျပ်")

            st.text_area(f"Viber သို့ ပို့ရန် ({selected_cust_stmt} ရှင်းတမ်း)", "\n".join(cust_viber_lines), height=170)

    st.markdown("---")
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("📊 ဂဏန်းအလိုက် စုစုပေါင်းနှင့် ကန့်သတ်ချက် (Limit Control)")
        grouped = df.groupby("Number")["Amount"].sum().reset_index().sort_values(by="Amount", ascending=False)
        grouped["Allowed"] = grouped["Amount"].apply(lambda x: min(x, max_limit))
        grouped["Excess"] = grouped["Amount"] - grouped["Allowed"]
        st.dataframe(grouped, use_container_width=True)
        
    with col_right:
        st.subheader("📋 ဖောက်သည်နှင့် ဂဏန်းအလိုက် ပေါင်းချုပ်စာရင်း")
        consolidated_view = df.groupby(["Date", "Day", "Customer", "Number"])["Amount"].sum().reset_index()
        st.dataframe(consolidated_view, use_container_width=True)
    
    st.markdown("---")
    c_exp, c_clr = st.columns([2, 1])
    with c_exp:
        excel_bytes = generate_styled_excel(df, commission_pct, payout_rate, winning_number, max_limit)
        st.download_button(
            label="📥 ဖောက်သည်ရှင်းတမ်း Tab အပါအဝင် Excel ဖိုင် Download ဆွဲရန်",
            data=excel_bytes,
            file_name=f"2D_Complete_Report_With_Customer_Statement_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    with c_clr:
        if st.button("🗑️ စာရင်း အစအဆုံး ရှင်းလင်းရန်"):
            st.session_state.data = pd.DataFrame(columns=["Date", "Day", "Customer", "Number", "Amount"])
            st.rerun()
else:
    st.info("📒 ယခုထိ ထည့်ထားသော စာရင်း မရှိသေးပါ။ အပေါ်ပါ နည်းလမ်း (၃) မျိုးမှ တစ်ခုခုဖြင့် စာရင်းစတင်သွင်းနိုင်ပါသည်။")
