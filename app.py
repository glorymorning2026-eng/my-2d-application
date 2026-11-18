import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from itertools import permutations, product

# စာမျက်နှာ အပြင်အဆင် သတ်မှတ်ခြင်း
st.set_page_config(page_title="Myanmar 2D Master Dealer App", page_icon="📒", layout="wide")

# စတိုင်နှင့် အရောင်များ သတ်မှတ်ခြင်း
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

st.markdown('<div class="main-header"><h1>📒 မြန်မာနှစ်လုံးထီ (2D) အဆင့်မြင့် အလိုအလျောက် စာရင်းချုပ်စနစ်</h1><p>Master Smart Ledger, Ultimate Precision Parser & Custom Formula Viber Statement Generator (API-Free)</p></div>', unsafe_allow_html=True)

# --- ဘယ်ဘက် Sidebar ဆက်တင်များ ---
st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များနှင့် ကန့်သတ်ချက်များ")
max_limit = st.sidebar.number_input("ဂဏန်းတစ်လုံးလျှင် အများဆုံးလက်ခံငွေ (Limit)", value=500000, step=50000)
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
blocked_numbers = [num.strip().zfill(2) for num in blocked_input.split(",") if num.strip()]

# Session State တွင် စာရင်းဒေတာ သိမ်းဆည်းရန် ပြင်ဆင်ခြင်း
if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=["Date", "Day", "Customer", "BatchID", "Number", "Amount"])

# --- 2D စည်းမျဉ်းအားလုံးကို တိကျစွာ ခွဲခြမ်းစိတ်ဖြာတွက်ချက်ပေးမည့် Engine ---
def parse_bets_locally(text):
    rows = []
    # ၊ နှင့် ; တို့ကို လိုင်းသစ်အဖြစ် ပြောင်းလဲခြင်း
    lines = text.replace("၊", "\n").replace(";", "\n").split("\n")
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # ၁။ ညီအကို / ညီကို (Sequential Pairs)
        if re.search(r'ညီအကို|ညီကို', line):
            amt_match = re.search(r'(\d+)', line)
            amt = int(amt_match.group(1)) if amt_match else 1000
            seq_pairs = ["01", "12", "23", "34", "45", "56", "67", "78", "89", "90"]
            added_set = set()
            for p in seq_pairs:
                rev = p[::-1]
                for n in [p, rev]:
                    if n not in added_set:
                        added_set.add(n)
                        rows.append({"Number": n, "Amount": amt})
            continue

        # ၂။ အပူ / ပူး (Doubles: 00-99)
        if re.search(r'အပူ|ပူး', line) and not re.search(r'ခွေ', line):
            amt_match = re.search(r'(\d+)', line)
            amt = int(amt_match.group(1)) if amt_match else 1000
            doubles = [f"{i}{i}" for i in range(10)]
            for d in doubles:
                rows.append({"Number": d, "Amount": amt})
            continue

        # ၃။ နက္ခတ် (Nakhat)
        if re.search(r'နက္ခတ်', line):
            amt_match = re.search(r'(\d+)', line)
            amt = int(amt_match.group(1)) if amt_match else 1000
            nakhat_list = ["07", "70", "18", "81", "24", "42", "35", "53", "69", "96"]
            for n in nakhat_list:
                rows.append({"Number": n, "Amount": amt})
            continue

        # ၄။ ပါဝါ (Power)
        if re.search(r'ပါဝါ|power', line, re.IGNORECASE):
            amt_match = re.search(r'(\d+)', line)
            amt = int(amt_match.group(1)) if amt_match else 1000
            power_list = ["05", "50", "16", "61", "27", "72", "38", "83", "49", "94"]
            for p in power_list:
                rows.append({"Number": p, "Amount": amt})
            continue

        # ၅။ ဘရိတ် (Break / B / bk) - ဥပမာ: 4 B 1500
        break_match = re.search(r'(\d)\s*(ဘရိတ်|B|bk)', line, re.IGNORECASE)
        if break_match:
            b_val = int(break_match.group(1))
            amt_match = re.search(r'(\d+)$', line)
            amt = int(amt_match.group(1)) if amt_match else 1000
            for i in range(100):
                num_str = f"{i:02d}"
                d_sum = (int(num_str[0]) + int(num_str[1])) % 10
                if d_sum == b_val:
                    rows.append({"Number": num_str, "Amount": amt})
            continue

        # ၆။ ပါတ်သီး / ပတ် (Paat-Thie)
        paat_match = re.search(r'(\d)\s*(ပါတ်|ပတ်|ပါ|ပ)', line)
        if paat_match:
            p_dig = paat_match.group(1)
            amt_match = re.search(r'(\d+)$', line)
            amt = int(amt_match.group(1)) if amt_match else 1000
            for i in range(100):
                num_str = f"{i:02d}"
                if p_dig in num_str:
                    rows.append({"Number": num_str, "Amount": amt})
            continue

        # ၇။ ထိပ်စီး / ထိပ် (Thate) - ဥပမာ: 5ထိပ်စီး 20000
        thate_match = re.search(r'(\d)\s*(ထိပ်စီး|ထိပ်)', line)
        if thate_match:
            t_dig = thate_match.group(1)
            amt_match = re.search(r'(\d+)$', line)
            amt = int(amt_match.group(1)) if amt_match else 1000
            for i in range(10):
                num_str = f"{t_dig}{i}"
                rows.append({"Number": num_str, "Amount": amt})
            continue

        # ၈။ နောက်ပိတ် / ပိတ် (Pate)
        pate_match = re.search(r'(\d)\s*(နောက်ပိတ်|ပိတ်)', line)
        if pate_match:
            pt_dig = pate_match.group(1)
            amt_match = re.search(r'(\d+)$', line)
            amt = int(amt_match.group(1)) if amt_match else 1000
            for i in range(10):
                num_str = f"{i}{pt_dig}"
                rows.append({"Number": num_str, "Amount": amt})
            continue

        # ၉။ အခွေ / ခွေ (Khve) - အပူးပါ / မပါ ခွဲခြားတွက်ချက်ခြင်း
        khve_match = re.search(r'([\d]{2,5})\s*(ခွေ|အခွေ)', line)
        if khve_match:
            digits = khve_match.group(1)
            include_doubles = bool(re.search(r'အပူးပါ|ပူးပါ|အပူးအပါ', line))
            amt_match = re.search(r'(\d+)$', line)
            amt = int(amt_match.group(1)) if amt_match else 1000
            
            perms = list(permutations(digits, 2))
            if include_doubles:
                perms = list(product(set(digits), repeat=2))
            
            added_nums = set()
            for p in perms:
                num_str = "".join(p)
                if num_str not in added_nums:
                    added_nums.add(num_str)
                    rows.append({"Number": num_str, "Amount": amt})
            continue

        # ၁၀။ @ နှင့် _ ပါသော အထူးပုံစံ (ဥပမာ: 20_4000@1000 သို့မဟုတ် 89_3000@2000)
        at_match = re.search(r'^(\d{1,2})[\_\.](\d+)@(\d+)$', line)
        if at_match:
            n1 = at_match.group(1).zfill(2)
            base_amt = int(at_match.group(2))
            at_amt = int(at_match.group(3))
            rows.append({"Number": n1, "Amount": base_amt})
            n2 = n1[::-1]
            rows.append({"Number": n2, "Amount": at_amt})
            continue

        # ရိုးရိုး @ ပုံစံ (ဥပမာ: 20@3000)
        at_simple = re.search(r'^(\d{1,2})@(\d+)$', line)
        if at_simple:
            n1 = at_simple.group(1).zfill(2)
            amt = int(at_simple.group(2))
            rows.append({"Number": n1, "Amount": amt})
            continue

        # ၁၁။ Slash/Dash နှင့် r တွဲလျက်ပါသော ပုံစံ (ဥပမာ: 51/52/58-50000r10000)
        complex_slash_r = re.search(r'^([\d\.\-\/\,\s]+)[\-\=](\d+)([Rrရ@&အာ])(\d+)$', line)
        if complex_slash_r:
            nums_part = complex_slash_r.group(1)
            base_amt = int(complex_slash_r.group(2))
            rev_amt = int(complex_slash_r.group(4))
            raw_nums = re.findall(r'\d{1,2}', nums_part)
            for n_str in raw_nums:
                n1 = n_str.zfill(2)
                rows.append({"Number": n1, "Amount": base_amt})
                n2 = n1[::-1]
                rows.append({"Number": n2, "Amount": rev_amt})
            continue

        # ၁၂။ အစက်များနှင့် .r. ပါဝင်သော ပုံစံ (ဥပမာ: 85.3000.r.500, 45.43.40.r.250, 82.45.10.20.40.50.58.r.250)
        match_dot_r = re.search(r'^([\d\.\-\/\,\s]+?)\.?\s*[Rr]\.?\s*(\d+)$', line)
        if match_dot_r:
            nums_part = match_dot_r.group(1)
            rev_amt = int(match_dot_r.group(2))
            
            tokens = [t for t in re.split(r'[\.\-\/\,\s]+', nums_part) if t.isdigit()]
            raw_nums = []
            base_amt = None
            
            # ဂဏန်းနှင့် မူလထိုးငွေ (ဥပမာ 3000) ခွဲထုတ်ခြင်း
            for t in tokens:
                if len(t) <= 2:
                    raw_nums.append(t.zfill(2))
                else:
                    base_amt = int(t)
            
            # အကယ်၍ မူလထိုးငွေ သီးသန့်မပါခဲ့လျှင် r နောက်မှ ငွေပမာဏကို အတူတူယူမည်
            if base_amt is None:
                base_amt = rev_amt
                
            for n1 in raw_nums:
                rows.append({"Number": n1, "Amount": base_amt})
                n2 = n1[::-1]
                rows.append({"Number": n2, "Amount": rev_amt})
            continue

        # ၁၃။ ဂဏန်းအများကြီးနှင့် နောက်ဆုံးတွင် ထိုးငွေတစ်ခုတည်းပါသော ပုံစံ (ဥပမာ: 66.77.88.1000, 22.77.2500)
        match_multi_amt = re.search(r'^([\d\.\-\/\,\s]+?)[\.\-\/\,\s](\d{3,6})$', line)
        if match_multi_amt:
            nums_part = match_multi_amt.group(1)
            amt = int(match_multi_amt.group(2))
            raw_nums = re.findall(r'\d{1,2}', nums_part)
            valid_nums = [n.zfill(2) for n in raw_nums if len(n) <= 2]
            for n1 in valid_nums:
                rows.append({"Number": n1, "Amount": amt})
            continue

        # ၁၄။ တစ်ကွက်တည်း ပြောင်းပြန် / R (ဥပမာ: 20R2500 သို့မဟုတ် 45 R 10000)
        single_r = re.search(r'^(\d{1,2})\s*[Rr]\s*(\d+)$', line)
        if single_r:
            n1 = single_r.group(1).zfill(2)
            amt = int(single_r.group(2))
            rows.append({"Number": n1, "Amount": amt})
            n2 = n1[::-1]
            if n1 != n2:
                rows.append({"Number": n2, "Amount": amt})
            continue

        # ၁၅။ တစ်ကွက်တည်း ရိုးရိုးထိုးငွေ (ဥပမာ: 44.500 သို့မဟုတ် 66.2000)
        single_amt = re.search(r'^(\d{1,2})[\.\-\s](\d{3,6})$', line)
        if single_amt:
            n1 = single_amt.group(1).zfill(2)
            amt = int(single_amt.group(2))
            rows.append({"Number": n1, "Amount": amt})
            continue

        # ၁၆။ အခြား ခွဲခြားသင်္ကေတများပါသော အတွဲများ (Fallback)
        pairs = re.findall(r'(\d{1,2})\s*[\-\=\/\_\.\s]\s*(\d+)', line)
        if pairs:
            for num_str, amt_str in pairs:
                num = num_str.zfill(2)
                amt = int(amt_str)
                if 0 <= int(num) <= 99 and 100 <= amt <= 500000:
                    rows.append({"Number": num, "Amount": amt})
            continue
            
    return rows

# --- စာရင်းများကို စစ်ဆေးပြီး BatchID ဖြင့် သိမ်းဆည်းပေးမည့် Function ---
def process_and_add_bets(extracted_data, entry_date_str, selected_day, default_customer):
    valid_rows = []
    blocked_found = []
    invalid_amt_found = 0
    total_added_amount = 0
    
    # ဖောက်သည်တစ်ဦးချင်းစီအတွက် Batch နံပါတ် အစဉ်လိုက် တပ်ပေးခြင်း (ဥပမာ: သိမ့်(1), သိမ့်(2))
    existing_batches = st.session_state.data[st.session_state.data["Customer"] == default_customer]["BatchID"]
    next_batch_num = 1
    if not existing_batches.empty:
        nums = []
        for b in existing_batches.unique():
            m = re.search(r'\((\d+)\)', str(b))
            if m:
                nums.append(int(m.group(1)))
        if nums:
            next_batch_num = max(nums) + 1
            
    batch_id_str = f"{default_customer}({next_batch_num})"

    for row in extracted_data:
        amt_val = int(row.get("Amount", 0))
        num_str = str(row.get("Number", "")).zfill(2)
        
        if amt_val < 100 or amt_val > 500000:
            invalid_amt_found += 1
            continue
        if num_str in blocked_numbers:
            blocked_found.append(num_str)
        else:
            valid_rows.append({
                "Date": entry_date_str,
                "Day": selected_day,
                "Customer": default_customer,
                "BatchID": batch_id_str,
                "Number": num_str,
                "Amount": amt_val
            })
            total_added_amount += amt_val
            
    return valid_rows, blocked_found, invalid_amt_found, total_added_amount

# --- Multi-Tab Excel အစီရင်ခံစာ ထုတ်ပေးမည့် Function ---
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
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1")
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
                c.fill = header_fill; c.font = font_header; c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.font = font_bold if idx in (3, 7) else font_body
                if idx == 3: c.fill = accent_fill
                elif idx == 7: c.fill = profit_fill if net_pl >= 0 else loss_fill
        if idx > 0 and idx != 4:
            c2.number_format = "#,##0"; c2.alignment = Alignment(horizontal="right")
        elif idx == 4:
            c2.alignment = Alignment(horizontal="center")

    format_sheet(ws_sum, {"A": 42, "B": 24, "C": 38})

    # TAB 2: နေ့အလိုက်_အပတ်စဉ်
    ws_day = wb.create_sheet(title="နေ့အလိုက်_အပတ်စဉ်")
    ws_day["A1"] = "နေ့အလိုက် နှင့် တစ်ပတ်စာ စာရင်းချုပ်"
    ws_day["A1"].font = font_title
    day_grp = df.groupby(["Date", "Day"]).agg(Total_Bets=("Number", "count"), Gross_Sales=("Amount", "sum")).reset_index()
    
    headers_day = ["ရက်စွဲ (Date)", "နေ့ (Day)", "အကွက်ရေ", "စုစုပေါင်းထိုးကြေး", f"ကော်မရှင် ({commission_pct}%)", "ဒိုင်ရငွေ (Net)"]
    for col_idx, h in enumerate(headers_day, 1):
        cell = ws_day.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill; cell.font = font_header; cell.alignment = Alignment(horizontal="center"); cell.border = thin_border

    for r_idx, row in enumerate(day_grp.itertuples(), 4):
        ws_day.cell(row=r_idx, column=1, value=row.Date).alignment = Alignment(horizontal="center")
        ws_day.cell(row=r_idx, column=2, value=str(row.Day)).alignment = Alignment(horizontal="center")
        ws_day.cell(row=r_idx, column=3, value=int(row.Total_Bets)).number_format = "#,##0"
        ws_day.cell(row=r_idx, column=4, value=float(row.Gross_Sales)).number_format = "#,##0"
        ws_day.cell(row=r_idx, column=5, value=f"=D{r_idx}*{commission_pct/100.0}").number_format = "#,##0"
        ws_day.cell(row=r_idx, column=6, value=f"=D{r_idx}-E{r_idx}").number_format = "#,##0"
        for c_i in range(1, 7):
            c = ws_day.cell(row=r_idx, column=c_i); c.font = font_body; c.border = thin_border
            if r_idx % 2 == 1: c.fill = zebra_fill
    format_sheet(ws_day, {"A": 18, "B": 16, "C": 14, "D": 22, "E": 20, "F": 22})

    # TAB 3: ဖောက်သည်အလိုက် ရှင်းတမ်း
    ws_cust = wb.create_sheet(title="ဖောက်သည်အလိုက် ရှင်းတမ်း")
    ws_cust["A1"] = "ဖောက်သည်တစ်ဦးချင်းစီ၏ ရှင်းတမ်း"
    ws_cust["A1"].font = font_title
    cust_summary_df = df.groupby("Customer")["Amount"].sum().reset_index().sort_values(by="Amount", ascending=False)
    
    headers_cust = ["ဖောက်သည် အမည်", "စုစုပေါင်းထိုးကြေး", f"ကော်မရှင် ({commission_pct}%)", "ဒိုင်ရငွေ", f"ဆုကြေး ({payout_rate}ဆ)", "အသားတင်ရှင်းတမ်း"]
    for col_idx, h in enumerate(headers_cust, 1):
        cell = ws_cust.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill; cell.font = font_header; cell.alignment = Alignment(horizontal="center"); cell.border = thin_border

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
            c = ws_cust.cell(row=r_idx, column=c_i); c.font = font_body; c.border = thin_border
            if r_idx % 2 == 1: c.fill = zebra_fill
    format_sheet(ws_cust, {"A": 26, "B": 22, "C": 20, "D": 24, "E": 24, "F": 24})

    # TAB 4: ဂဏန်းပေါင်းချုပ်_Limit
    ws_num = wb.create_sheet(title="ဂဏန်းပေါင်းချုပ်_Limit")
    ws_num["A1"] = "ဂဏန်းအလိုက် စုစုပေါင်းနှင့် Limit ထိန်းချုပ်မှု"
    ws_num["A1"].font = font_title
    num_grp = df.groupby("Number")["Amount"].sum().reset_index().sort_values(by="Amount", ascending=False)
    headers_num = ["2D ဂဏန်း", "စုစုပေါင်း ထိုးငွေ", "ဒိုင်လက်ခံငွေ (Allowed)", "ကျော်လွန်ငွေ (Excess)"]
    for col_idx, h in enumerate(headers_num, 1):
        cell = ws_num.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill; cell.font = font_header; cell.alignment = Alignment(horizontal="center"); cell.border = thin_border

    for r_idx, row in enumerate(num_grp.itertuples(), 4):
        c_num = ws_num.cell(row=r_idx, column=1, value=str(row.Number).zfill(2))
        c_num.alignment = Alignment(horizontal="center"); c_num.number_format = "@"
        ws_num.cell(row=r_idx, column=2, value=float(row.Amount)).number_format = "#,##0"
        ws_num.cell(row=r_idx, column=3, value=f"=MIN(B{r_idx},{max_limit})").number_format = "#,##0"
        ws_num.cell(row=r_idx, column=4, value=f"=B{r_idx}-C{r_idx}").number_format = "#,##0"
        for c_i in range(1, 5):
            c = ws_num.cell(row=r_idx, column=c_i); c.font = font_body; c.border = thin_border
            if float(row.Amount) > max_limit and c_i == 4:
                c.fill = loss_fill; c.font = font_bold
    format_sheet(ws_num, {"A": 16, "B": 24, "C": 24, "D": 26})

    # TAB 5: အသေးစိတ်စာရင်း
    ws_det = wb.create_sheet(title="အသေးစိတ်စာရင်း")
    ws_det["A1"] = "အသေးစိတ် စာရင်းမှတ်တမ်း"
    ws_det["A1"].font = font_title
    det_grp = df.groupby(["Date", "Day", "Customer", "BatchID", "Number"])["Amount"].sum().reset_index()
    headers_det = ["ရက်စွဲ (Date)", "နေ့ (Day)", "ဖောက်သည်", "အကြိမ်ရေ (Batch)", "2D ဂဏန်း", "ထိုးငွေ (MMK)"]
    for col_idx, h in enumerate(headers_det, 1):
        cell = ws_det.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill; cell.font = font_header; cell.alignment = Alignment(horizontal="center"); cell.border = thin_border

    for r_idx, row in enumerate(det_grp.itertuples(), 4):
        ws_det.cell(row=r_idx, column=1, value=row.Date).alignment = Alignment(horizontal="center")
        ws_det.cell(row=r_idx, column=2, value=row.Day).alignment = Alignment(horizontal="center")
        ws_det.cell(row=r_idx, column=3, value=row.Customer)
        ws_det.cell(row=r_idx, column=4, value=row.BatchID).alignment = Alignment(horizontal="center")
        c_n = ws_det.cell(row=r_idx, column=5, value=str(row.Number).zfill(2))
        c_n.alignment = Alignment(horizontal="center"); c_n.number_format = "@"
        ws_det.cell(row=r_idx, column=6, value=float(row.Amount)).number_format = "#,##0"
        for c_i in range(1, 7):
            c = ws_det.cell(row=r_idx, column=c_i); c.font = font_body; c.border = thin_border
            if r_idx % 2 == 1: c.fill = zebra_fill
    format_sheet(ws_det, {"A": 16, "B": 16, "C": 22, "D": 22, "E": 16, "F": 22})

    wb.save(output)
    return output.getvalue()

# --- ပင်မ စာရင်းသွင်းရာ အပိုင်း ---
st.subheader("📝 စာရင်းသွင်းရန် နည်းလမ်းများ (Ultimate Precision Tracker API-Free)")

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

col1, col2 = st.columns([1, 2])
with col1:
    cust_manual = st.text_input("ဖောက်သည် အမည်", placeholder="ဥပမာ - သိမ့်")
with col2:
    smart_input_text = st.text_area("ဖောက်သည် ပို့လိုက်သော စာရင်းစာသားကို ထည့်ပါ", placeholder="85.3000.r.500\n20@3000\n82.45.10.20.40.50.58.r.250\n5ထိပ်စီး 20000\n4 B 1500", height=130)

if st.button("➕ စာရင်းသွင်းမည်", type="primary"):
    if is_closed:
        st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
    elif not cust_manual or not smart_input_text:
        st.warning("ဖောက်သည် အမည်နှင့် ထိုးမည့် စာသားကို ဖြည့်သွင်းပါ။")
    else:
        try:
            clean_cust = cust_manual.strip()
            extracted_data = parse_bets_locally(smart_input_text)
            valid_rows, blocked_found, invalid_amt, total_amt = process_and_add_bets(extracted_data, entry_date_str, selected_day, clean_cust)
            
            if valid_rows:
                st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame(valid_rows)], ignore_index=True)
                st.success(f"🎉 စာရင်းသွင်းပြီးပါပြီ! ({clean_cust}) စုစုပေါင်းထိုးငွေ = {total_amt:,} ကျပ် ({len(valid_rows)} ကွက်)")
                if blocked_found:
                    st.warning(f"🚫 ပိတ်ဂဏန်းများ ဖယ်ထုတ်ထားပါသည်: {', '.join(set(blocked_found))}")
                st.rerun()
            else:
                st.warning("စနစ်က ဂဏန်းနှင့် ငွေပမာဏကို ရှာမတွေ့ပါ။ စာသားပုံစံ မှန်/မမှန် စစ်ဆေးပါ။")
        except Exception as e:
            st.error(f"စာရင်းတွက်ချက်ရာတွင် အမှားအယွင်း ဖြစ်ပေါ်ခဲ့ပါသည်: {e}")

# --- စာရင်းချုပ် ပြသခြင်းနှင့် Statement ထုတ်ပေးခြင်း ---
st.markdown("---")
if not st.session_state.data.empty:
    df = st.session_state.data
    
    total_sales = df["Amount"].sum()
    commission_amt = total_sales * (commission_pct / 100.0)
    net_sales = total_sales - commission_amt
    
    win_num_clean = str(winning_number).strip().zfill(2) if winning_number else ""
    payout_amount = 0
    winning_total_amt = 0
    if winning_number:
        winning_matched = df[df["Number"] == win_num_clean]
        winning_total_amt = winning_matched["Amount"].sum()
        payout_amount = winning_total_amt * payout_rate
    net_pl = net_sales - payout_amount

    # အပေါ်ပိုင်း အကျဉ်းချုပ် ကိန်းဂဏန်းများ
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("💰 စုစုပေါင်း ရောင်းရငွေ", f"{total_sales:,.0f} MMK")
    m2.metric(f"💸 ကော်မရှင် ({commission_pct}%)", f"{commission_amt:,.0f} MMK")
    m3.metric("📉 ကော်မရှင်နှုတ်ပြီး ဒိုင်ရငွေ", f"{net_sales:,.0f} MMK")
    m4.metric("📊 အသားတင် အနိုင် / အရှုံး", f"{net_pl:,.0f} MMK")
    
    st.markdown("---")
    
    # ပေါက်ဂဏန်း တိုက်စစ်ခြင်း
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
    
    # ဘယ်ဘက်အခြမ်း - ဒိုင်ချုပ် Viber Report
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
            report_lines.append(f"🏆 ပေးရမည့်ဆုကြေး ({payout_rate}ဆ) = -{payout_amount:,.0f} ကျပ်")
        status_label = "အမြတ် (+)" if net_pl >= 0 else "အရှုံး (-)"
        report_lines.append(f"📊 အသားတင် {status_label} = {net_pl:,.0f} ကျပ်")
        
        st.text_area("ဒိုင်ချုပ် Viber Text (Copy ယူရန်)", "\n".join(report_lines), height=180)

    # ညာဘက်အခြမ်း - ဖောက်သည်တစ်ဦးချင်းစီ၏ အထူးဖော်မြူလာပါ Viber Statement
    with col_rep2:
        st.markdown("##### 👤 ဖောက်သည်တစ်ဦးချင်းစီ၏ အသေးစိတ် Batch များနှင့် ပေါင်းချုပ် Statement")
        unique_customers = df["Customer"].unique()
        selected_cust_stmt = st.selectbox("ဖောက်သည် ရွေးချယ်ပါ", unique_customers)
        
        if selected_cust_stmt:
            cust_df = df[df["Customer"] == selected_cust_stmt]
            cust_batches = cust_df["BatchID"].unique()
            
            cust_viber_lines = [f"📒 *{selected_cust_stmt}* ၏ 2D အသေးစိတ်ရှင်းတမ်း", "=" * 28]
            batch_totals_for_sum = []
            
            for b_id in cust_batches:
                b_df = cust_df[cust_df["BatchID"] == b_id]
                b_total = b_df["Amount"].sum()
                batch_totals_for_sum.append((b_id, b_total))
                
                cust_viber_lines.append(f"--- *{b_id}* ---")
                for idx, row in b_df.iterrows():
                    cust_viber_lines.append(f"• {row['Number']} = {row['Amount']:,} ကျပ်")
                cust_viber_lines.append(f"*{b_id} စုစုပေါင်း* = {b_total:,.0f} ကျပ်\n")
                
            cust_overall_total = cust_df["Amount"].sum()
            cust_comm = cust_overall_total * (commission_pct / 100.0)
            cust_net_pay = cust_overall_total - cust_comm
            
            cust_win_amt = 0.0
            if winning_number:
                cust_win_match = cust_df[cust_df["Number"] == win_num_clean]
                cust_win_amt = float(cust_win_match["Amount"].sum() * payout_rate)
            
            cust_final_balance = cust_net_pay - cust_win_amt

            # အထူးတောင်းဆိုထားသော ပုံစံ: [သိမ့်(1)=သိမ့်(1)၏Total] + [သိမ့်(2)=သိမ့်(2)၏Total] = *(Total ကျပ်)*
            formula_items = [f"[{b[0]}={b[1]:,.0f} ကျပ်]" for b in batch_totals_for_sum]
            formula_string = " + ".join(formula_items) + f" = *({cust_overall_total:,.0f} ကျပ်)*"

            cust_viber_lines.append("=" * 28)
            cust_viber_lines.append(formula_string)
            cust_viber_lines.append(f"ကော်မရှင် ({commission_pct}%) = -{cust_comm:,.0f} ကျပ်")
            cust_viber_lines.append(f"ပေးချေရန် Net = {cust_net_pay:,.0f} ကျပ်")
            if winning_number:
                cust_viber_lines.append(f"ပေါက်ဂဏန်းဆုကြေး ({win_num_clean}) = +{cust_win_amt:,.0f} ကျပ်")
                bal_label = "ပေးရန် (+) / ရရန် (-)" if cust_final_balance >= 0 else "ရရန်"
                cust_viber_lines.append(f"ရှင်းတမ်းစုစုပေါင်း ({bal_label}) = {cust_final_balance:,.0f} ကျပ်")

            st.text_area(f"Viber သို့ ပို့ရန် ({selected_cust_stmt} ရှင်းတမ်းအပြည့်အစုံ)", "\n".join(cust_viber_lines), height=280)

    st.markdown("---")
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("📊 ဂဏန်းအလိုက် စုစုပေါင်းနှင့် ကန့်သတ်ချက် (Limit Control)")
        grouped = df.groupby("Number")["Amount"].sum().reset_index().sort_values(by="Amount", ascending=False)
        grouped["Allowed"] = grouped["Amount"].apply(lambda x: min(x, max_limit))
        grouped["Excess"] = grouped["Amount"] - grouped["Allowed"]
        st.dataframe(grouped, use_container_width=True)
        
    with col_right:
        st.subheader("📋 ဖောက်သည်နှင့် Batch အလိုက် ပေါင်းချုပ်စာရင်း")
        consolidated_view = df.groupby(["Date", "Day", "Customer", "BatchID", "Number"])["Amount"].sum().reset_index()
        st.dataframe(consolidated_view, use_container_width=True)
    
    st.markdown("---")
    c_exp, c_clr = st.columns([2, 1])
    with c_exp:
        excel_bytes = generate_styled_excel(df, commission_pct, payout_rate, winning_number, max_limit)
        st.download_button(
            label="📥 ဖောက်သည်ရှင်းတမ်း Tab အပါအဝင် Excel ဖိုင် Download ဆွဲရန်",
            data=excel_bytes,
            file_name=f"2D_Complete_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    with c_clr:
        if st.button("🗑️ စာရင်း အစအဆုံး ရှင်းလင်းရန်"):
            st.session_state.data = pd.DataFrame(columns=["Date", "Day", "Customer", "BatchID", "Number", "Amount"])
            st.rerun()
else:
    st.info("📒 ယခုထိ ထည့်ထားသော စာရင်း မရှိသေးပါ။ အထက်ပါ စာရင်းသွင်းကွက်တွင် ရိုက်ထည့်၍ စာရင်းစတင်သွင်းနိုင်ပါသည်။")
