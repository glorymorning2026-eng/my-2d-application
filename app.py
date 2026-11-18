import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from itertools import permutations, product
​st.set_page_config(page_title="Myanmar 2D Master Dealer App", page_icon="📒", layout="wide")
​st.markdown("""
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
​st.markdown('<div class="main-header"><h1>📒 မြန်မာနှစ်လုံးထီ (2D) အဆင့်မြင့် အလိုအလျောက် စာရင်းချုပ်စနစ်</h1><p>Master Smart Ledger, Ultimate Precision Parser & Custom Formula Viber Statement Generator (API-Free)</p></div>', unsafe_allow_html=True)
​st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များနှင့် ကန့်သတ်ချက်များ")
max_limit = st.sidebar.number_input("ဂဏန်းတစ်လုံးလျှင် အများဆုံးလက်ခံငွေ (Limit)", value=500000, step=50000)
payout_rate = st.sidebar.number_input("ဆုကြေးနှုန်း (ဆ)", value=80, step=5)
commission_pct = st.sidebar.number_input("ကော်မရှင် (%)", value=0, step=1)
winning_number = st.sidebar.text_input("ထွက်မည့် ပေါက်ဂဏန်း", value="")
​st.sidebar.markdown("---")
st.sidebar.subheader("⏰ အချိန်ပိတ်ရက် ထိန်းချုပ်မှု")
enable_time_lock = st.sidebar.checkbox("အချိန်ပိတ်ရန် ဖွင့်မည်", value=False)
closing_time = st.sidebar.time_input("ပိတ်မည့် အချိန်", value=datetime.strptime("12:00", "%H:%M").time())
​st.sidebar.markdown("---")
st.sidebar.subheader("🚫 ပိတ်ဂဏန်းများ (Comma ခြားရန်)")
blocked_input = st.sidebar.text_input("မခံမည့် ဂဏန်းများ (ဥပမာ: 39, 88)", value="")
blocked_numbers = [num.strip().zfill(2) for num in blocked_input.split(",") if num.strip()]
​if 'data' not in st.session_state:
st.session_state.data = pd.DataFrame(columns=["Date", "Day", "Customer", "BatchID", "Number", "Amount"])
​--- Comprehensive Unified Precision Parsing Engine ---
​def parse_bets_locally(text):
rows = []
lines = text.replace("၊", "\n").replace(";", "\n").split("\n")
​for line in lines:
line = line.strip()
if not line:
continue
​# 1. Day Clearing (ဥပမာ - 11000 နေ့ရှင်း)
day_clear_match = re.search(r'([\d,]+)\s*နေ့ရှင်း', line)
if day_clear_match:
amt_str = day_clear_match.group(1).replace(",", "")
amt = int(amt_str)
rows.append({"Number": "11", "Amount": amt})
continue
​# 2. Break (ဥပမာ - 5bk / 7bk250)
break_matches = re.findall(r'(\d)\s*(ဘရိတ်|B|bk)', line, re.IGNORECASE)
if break_matches:
amt_match = re.search(r'(\d+)$', line)
amt = int(amt_match.group(1)) if amt_match else 250
for b_match in break_matches:
b_val = int(b_match[0])
for i in range(100):
num_str = f"{i:02d}"
d_sum = (int(num_str[0]) + int(num_str[1])) % 10
if d_sum == b_val:
rows.append({"Number": num_str, "Amount": amt})
continue
​# 3. Paat / Pat (ဥပမာ - 9ပတ်150tt သို့မဟုတ် 7ပက်2500)
paat_match = re.search(r'(\d)\s*(ပတ်|ပါတ်|ပက်|ပ)', line)
if paat_match:
p_dig = paat_match.group(1)
amt_match = re.search(r'(\d+)(?:tt)?$', line)
amt = int(amt_match.group(1)) if amt_match else 1000
for i in range(100):
num_str = f"{i:02d}"
if p_dig in num_str:
rows.append({"Number": num_str, "Amount": amt})
continue
​# 4. Multi R / Dash R (ဥပမာ - 46/96r1500 သို့မဟုတ် 30-35-50-R1500)
match_multi_r = re.search(r'^([\d.-/,\s]+?)[Rr]\s*(\d+)$', line)
if match_multi_r:
nums_part = match_multi_r.group(1)
amt = int(match_multi_r.group(2))
raw_nums = re.findall(r'\d{1,2}', nums_part)
for n_str in raw_nums:
n1 = n_str.zfill(2)
rows.append({"Number": n1, "Amount": amt})
n2 = n1[::-1]
if n1 != n2:
rows.append({"Number": n2, "Amount": amt})
continue
​# 5. Slash / Star / Equals with Amount (ဥပမာ - 11/99/77*3000 သို့မဟုတ် 99=250)
slash_amt_match = re.search(r'^([\d.-/,\s]+)*=\s$', line)
if slash_amt_match:
nums_part = slash_amt_match.group(1)
amt = int(slash_amt_match.group(2))
raw_nums = re.findall(r'\d{1,2}', nums_part)
for n_str in raw_nums:
rows.append({"Number": n_str.zfill(2), "Amount": amt})
continue
​# 6. Advanced @ Format (ဥပမာ - 20@3000)
at_simple = re.search(r'^(\d{1,2})\s*@\s*(\d+)$', line)
if at_simple:
n1 = at_simple.group(1).zfill(2)
amt = int(at_simple.group(2))
rows.append({"Number": n1, "Amount": amt})
n2 = n1[::-1]
rows.append({"Number": n2, "Amount": amt})
continue
​# 7. Dot R Format (ဥပမာ - 85.3000.r.500)
match_dot_r = re.search(r'^([\d.-/,\s]+?).?\s*[Rr].?\s*(\d+)$', line)
if match_dot_r:
nums_part = match_dot_r.group(1)
rev_amt = int(match_dot_r.group(2))
tokens = [t for t in re.split(r'[.-/,\s]+', nums_part) if t.isdigit()]
raw_nums = []
base_amt = None
for t in tokens:
if len(t) <= 2:
raw_nums.append(t.zfill(2))
else:
base_amt = int(t)
if base_amt is None:
base_amt = rev_amt
for n1 in raw_nums:
rows.append({"Number": n1, "Amount": base_amt})
n2 = n1[::-1]
rows.append({"Number": n2, "Amount": rev_amt})
continue
​# 8. Single Spaced Numbers (ဥပမာ - 44 99 11)
single_spaced = re.findall(r'(\d{2})', line)
if single_spaced and len(line) <= 15 and not any(k in line for k in ['ပတ်', 'ဘရိတ်', 'R', 'r']):
amt_match = re.search(r'(\d{3,6})$', line)
amt = int(amt_match.group(1)) if amt_match else 1000
for n_str in single_spaced:
rows.append({"Number": n_str.zfill(2), "Amount": amt})
continue
​# 9. Fallback General Pairs
pairs = re.findall(r'(\d{1,2})\s*[-=/_.\s]\s*(\d+)', line)
if pairs:
for num_str, amt_str in pairs:
num = num_str.zfill(2)
amt = int(amt_str)
if 0 <= int(num) <= 99 and 100 <= amt <= 500000:
rows.append({"Number": num, "Amount": amt})
continue
​return rows
​def process_and_add_bets(extracted_data, entry_date_str, selected_day, default_customer):
valid_rows = []
blocked_found = []
invalid_amt_found = 0
total_added_amount = 0
​existing_batches = st.session_state.data[st.session_state.data["Customer"] == default_customer]["BatchID"]
next_batch_num = 1
if not existing_batches.empty:
nums = []
for b in existing_batches.unique():
m = re.search(r'((\d+))', str(b))
if m:
nums.append(int(m.group(1)))
if nums:
next_batch_num = max(nums) + 1
​batch_id_str = f"{default_customer}({next_batch_num})"
​for row in extracted_data:
amt_val = int(row.get("Amount", 0))
num_str = str(row.get("Number", "")).zfill(2)
​if amt_val < 100 or amt_val > 500000:
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
​return valid_rows, blocked_found, invalid_amt_found, total_added_amount
​def generate_styled_excel(df, commission_pct, payout_rate, winning_number, max_limit):
output = io.BytesIO()
wb = openpyxl.Workbook()
​header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
accent_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
profit_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
loss_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
​font_title = Font(name="Pyidaungsu", size=14, bold=True, color="1E293B")
font_header = Font(name="Pyidaungsu", size=11, bold=True, color="FFFFFF")
font_body = Font(name="Pyidaungsu", size=11, color="0F172A")
font_bold = Font(name="Pyidaungsu", size=11, bold=True, color="0F172A")
​thin_border = Border(
left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1")
)
​def format_sheet(ws, col_widths):
ws.views.sheetView[0].showGridLines = True
for col_letter, width in col_widths.items():
ws.column_dimensions[col_letter].width = width
​win_num_clean = str(winning_number).strip().zfill(2) if winning_number else ""
​ws_sum = wb.active
ws_sum.title = "စာရင်းချုပ်အကျဉ်း"
ws_sum["A1"] = "မြန်မာနှစ်လုံးထီ (2D) ငွေစာရင်းချုပ် အစီရင်ခံစာ"
ws_sum["A1"].font = font_title
ws_sum["A2"] = f"ထုတ်ပြန်သည့်ရက်စွဲ - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
ws_sum["A2"].font = font_body
​total_sales = float(df["Amount"].sum())
comm_amount = total_sales * (commission_pct / 100.0)
net_sales = total_sales - comm_amount
​win_bet_total = float(df[df["Number"] == win_num_clean]["Amount"].sum()) if win_num_clean else 0.0
payout_total = win_bet_total * payout_rate
net_pl = net_sales - payout_total
​summary_rows = [
("အကြောင်းအရာ", "ပမာဏ / တန်ဖိုး", "မှတ်ချက်"),
("စုစုပေါင်း ရောင်းရငွေ (Total Sales)", total_sales, f"စုစုပေါင်း အကွက်ရေ {len(df):,} ကွက်"),
(f"ကော်မရှင် နှုတ်ငွေ ({commission_pct}%)", comm_amount, "ဒိုင်ကော်မရှင် နှုတ်ယူငွေ"),
("ကော်မရှင်နှုတ်ပြီး ဒိုင်ရငွေ (Net Sales)", "=B5-B6", "အသားတင် ရငွေ"),
("ထွက်ဂဏန်း (Winning Number)", win_num_clean if win_num_clean else "မသတ်မှတ်ရသေး", f"ဆုကြေးနှုန်း {payout_rate} ဆ"),
("ပေါက်ဂဏန်းပေါ် ထိုးငွေစုစုပေါင်း", win_bet_total, "ပေါက်သီး တန်ဖိုး"),
("ပေးရမည့် ဆုကြေးငွေစုစုပေါင်း (Total Payout)", f"=B9*{payout_rate}", "ကံထူးရှင်များသို့ ပေးလျော်ငွေ"),
("ဒိုင် အသားတင် အနိုင် / အရှုံး (Net Profit/Loss)", "=B7-B10", "အပေါင်း (+) ဖြစ်ပါက အမြတ် / အနှုတ် (-) ဖြစ်ပါက အရှုံး")
]
​start_row = 4
for idx, (label, val, note) in enumerate(summary_rows):
r = start_row + idx
c1 = ws_sum.cell(row=r, column=1, value=label)
c2 = ws_sum.cell(row=r, column=2, value=val)
c3 = ws_sum.cell(row=r, column=3, value=note)
​for c in (c1, c2, c3):
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
​format_sheet(ws_sum, {"A": 42, "B": 24, "C": 38})
wb.save(output)
return output.getvalue()
​st.subheader("📝 စာရင်းသွင်းရန် နည်းလမ်းများ (Ultimate Precision Tracker API-Free)")
​col_d1, col_d2 = st.columns(2)
with col_d1:
entry_date = st.date_input("📆 စာရင်းသွင်းမည့် ရက်စွဲ (Date)", value=datetime.now().date())
with col_d2:
default_day_idx = min(entry_date.weekday(), 5)
day_options = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "WEEKEND"]
selected_day = st.selectbox("📅 စာရင်းသွင်းမည့် နေ့ (Day)", day_options, index=default_day_idx)
​entry_date_str = entry_date.strftime("%Y-%m-%d")
​col1, col2 = st.columns([1, 2])
with col1:
cust_manual = st.text_input("ဖောက်သည် အမည်", placeholder="ဥပမာ - မကြီးခင်လေး")
with col2:
smart_input_text = st.text_area("ဖောက်သည် ပို့လိုက်သော စာရင်းစာသားကို ထည့်ပါ", placeholder="44 99 11\n11000နေ့ရှင်း\n11/99/77*3000\n46/96r1500\n7ပက်2500", height=130)
​if st.button("➕ စာရင်းသွင်းမည်", type="primary"):
if not cust_manual or not smart_input_text:
st.warning("ဖောက်သည် အမည်နှင့် ထိုးမည့် စာသားကို ဖြည့်သွင်းပါ။")
else:
try:
clean_cust = cust_manual.strip()
extracted_data = parse_bets_locally(smart_input_text)
valid_rows, blocked_found, invalid_amt, total_amt = process_and_add_bets(extracted_data, entry_date_str, selected_day, clean_cust)
​if valid_rows:
st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame(valid_rows)], ignore_index=True)
st.success(f"🎉 စာရင်းသွင်းပြီးပါပြီ! ({clean_cust}) စုစုပေါင်းထိုးငွေ = {total_amt:,} ကျပ် ({len(valid_rows)} ကွက်)")
st.rerun()
else:
st.warning("စနစ်က ဂဏန်းနှင့် ငွေပမာဏကို ရှာမတွေ့ပါ။ စာသားပုံစံ မှန်/မမှန် စစ်ဆေးပါ။")
except Exception as e:
st.error(f"စာရင်းတွက်ချက်ရာတွင် အမှားအယွင်း ဖြစ်ပေါ်ခဲ့ပါသည်: {e}")
​st.markdown("---")
if not st.session_state.data.empty:
df = st.session_state.data
total_sales = df["Amount"].sum()
commission_amt = total_sales * (commission_pct / 100.0)
net_sales = total_sales - commission_amt
​st.subheader("📊 နေ့အလိုက် / တစ်ပတ်စာ စာရင်းချုပ်နှင့် ဖောက်သည် Viber Statement Generator")
​col_rep1, col_rep2 = st.columns([1, 1])
with col_rep2:
st.markdown("##### 👤 ဖောက်သည်တစ်ဦးချင်းစီ၏ အသေးစိတ် Batch များနှင့် ပေါင်းချုပ် Statement")
unique_customers = df["Customer"].unique()
selected_cust_stmt = st.selectbox("ဖောက်သည် ရွေးချယ်ပါ", unique_customers)
​if selected_cust_stmt:
cust_df = df[df["Customer"] == selected_cust_stmt]
cust_batches = cust_df["BatchID"].unique()
​cust_viber_lines = [f"📒 {selected_cust_stmt} ၏ 2D အသေးစိတ်ရှင်းတမ်း", "=" * 28]
batch_totals_for_sum = []
​for b_id in cust_batches:
b_df = cust_df[cust_df["BatchID"] == b_id]
b_total = b_df["Amount"].sum()
batch_totals_for_sum.append((b_id, b_total))
​cust_viber_lines.append(f"--- {b_id} ---")
for idx, row in b_df.iterrows():
cust_viber_lines.append(f"• {row['Number']} = {row['Amount']:,} ကျပ်")
cust_viber_lines.append(f"{b_id} စုစုပေါင်း = {b_total:,.0f} ကျပ်\n")
​cust_overall_total = cust_df["Amount"].sum()
formula_items = [f"[{b[0]}={b[1]:,.0f} ကျပ်]" for b in batch_totals_for_sum]
formula_string = " + ".join(formula_items) + f" = ({cust_overall_total:,.0f} ကျပ်)"
​cust_viber_lines.append("=" * 28)
cust_viber_lines.append(formula_string)
cust_viber_lines.append(f"ကော်မရှင် ({commission_pct}%) = -0 ကျပ်")
cust_viber_lines.append(f"ပေးချေရန် Net = {cust_overall_total:,.0f} ကျပ်")
cust_viber_lines.append(f"ပေါက်ဂဏန်းဆုကြေး (60) = +0 ကျပ်")
cust_viber_lines.append(f"ရှင်းတမ်းစုစုပေါင်း (ပေးရန် (+) / ရရန် (-)) = {cust_overall_total:,.0f} ကျပ်")
​st.text_area(f"Viber သို့ ပို့ရန် ({selected_cust_stmt} ရှင်းတမ်းအပြည့်အစုံ)", "\n".join(cust_viber_lines), height=280)
