import streamlit as st
import pandas as pd
import json
from PIL import Image, ImageOps
from datetime import datetime
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

st.markdown('<div class="main-header"><h1>📒 မြန်မာနှစ်လုံးထီ (2D) အလိုအလျောက် စာရင်းချုပ်စနစ်</h1><p>Smart Ledger & Weekly Report Automation</p></div>', unsafe_allow_html=True)

# Sidebar Settings
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
st.sidebar.subheader("🤖 AI စနစ် ဆက်တင်")
gemini_api_key = st.sidebar.text_input("Gemini API Key ထည့်ရန်", type="password", value="")

if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=["Day", "Customer", "Number", "Amount"])

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

def process_and_add_bets(extracted_data, selected_day, default_customer):
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
            valid_rows.append({"Day": selected_day, "Customer": cust_name, "Number": num_str, "Amount": amt_val})
            total_added_amount += amt_val
            
    return valid_rows, blocked_found, invalid_amt_found, total_added_amount

st.subheader("📝 စာရင်းသွင်းရန် နည်းလမ်းများ")

is_closed = False
if enable_time_lock:
    current_time = datetime.now().time()
    if current_time >= closing_time:
        is_closed = True
        st.error(f"⚠️ သတိပေးချက် - စာရင်းပိတ်ချိန် ({closing_time.strftime('%H:%M')}) ကျော်လွန်သွားပြီဖြစ်၍ စာရင်းအသစ် လက်မခံတော့ပါ။")

selected_day = st.selectbox("📅 စာရင်းသွင်းမည့် ရက် (Day)", ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "WEEKEND"])

tab1, tab2, tab3 = st.tabs([
    "✍️ ဂဏန်းနှင့်ငွေ ရောနှောရိုက်ထည့်ရန် (Smart Input)", 
    "💬 ဖောက်သည်ပို့သော စာသား (SMS) ဖြင့် AI ဖတ်ရန်", 
    "📸 စကရင်ရှော့ ပုံတင်၍ AI ဖတ်ရန်"
])

with tab1:
    st.markdown("#### 💡 ထိုးမည့်ဂဏန်းနှင့် ငွေပမာဏကို တစ်ခါတည်း ရောနှောရိုက်ထည့်ပါ")
    col1, col2 = st.columns([1, 2])
    with col1:
        cust_manual = st.text_input("ဖောက်သည် အမည်", key="c_man", placeholder="ဥပမာ - Hnin")
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
                    response = client.models.generate_content(model='gemini-2.0-flash', contents=prompt)
                    
                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                    extracted_data = json.loads(clean_text)
                    valid_rows, blocked_found, invalid_amt, total_amt = process_and_add_bets(extracted_data, selected_day, cust_manual)
                    
                    if valid_rows:
                        st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame(valid_rows)], ignore_index=True)
                        st.success(f"🎉 စာရင်းသွင်းပြီးပါပြီ! ({selected_day} - {cust_manual}) စုစုပေါင်းထိုးငွေ = {total_amt:,} ကျပ် ({len(valid_rows)} ကွက်)")
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
                    response = client.models.generate_content(model='gemini-2.0-flash', contents=prompt)
                    
                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                    extracted_data = json.loads(clean_text)
                    valid_rows, blocked_found, invalid_amt, total_amt = process_and_add_bets(extracted_data, selected_day, sms_customer)
                    
                    if valid_rows:
                        st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame(valid_rows)], ignore_index=True)
                        st.success(f"🎉 စာသားဖတ်ရှုပြီးပါပြီ! ({selected_day}) စုစုပေါင်းထိုးငွေ = {total_amt:,} ကျပ် ({len(valid_rows)} ကွက်)")
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
                    st.write(f"• သတ်မှတ်ရက်: `{selected_day}`")
                    
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
                                    
                                    response = client.models.generate_content(
                                        model='gemini-2.0-flash',
                                        contents=[normalized_img, prompt]
                                    )
                                    
                                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                                    extracted_data = json.loads(clean_text)
                                    valid_rows, blocked_found, invalid_amt, total_amt = process_and_add_bets(extracted_data, selected_day, img_cust_fallback)
                                    
                                    if valid_rows:
                                        st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame(valid_rows)], ignore_index=True)
                                        st.success(f"🎉 ပုံထဲမှ စာရင်းများကို အောင်မြင်စွာ ဖတ်ရှုပြီးပါပြီ! ({selected_day}) စုစုပေါင်းထိုးငွေ = {total_amt:,} ကျပ် ({len(valid_rows)} ကွက်)")
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
    unique_customers = df["Customer"].nunique()
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("💰 စုစုပေါင်း ရောင်းရငွေ", f"{total_sales:,} MMK")
    m2.metric(f"💸 ကော်မရှင် ({commission_pct}%)", f"{commission_amt:,.0f} MMK")
    m3.metric("📉 ကော်မရှင်နှုတ်ပြီး ဒိုင်ရငွေ", f"{net_sales:,.0f} MMK")
    m4.metric("🔢 စုစုပေါင်း အကွက်ရေ", f"{total_bets:,} ကွက်")
    
    st.markdown("---")
    
    st.subheader("📊 နေ့အလိုက် / တစ်ပတ်စာ စာရင်းချုပ်အကျဉ်း (Weekly Report Summary)")
    day_summary = df.groupby("Day")["Amount"].sum().reset_index()
    st.dataframe(day_summary, use_container_width=True)
    
    payout_amount = 0
    if winning_number:
        winning_matched = df[df["Number"] == str(winning_number).strip().zfill(2)]
        winning_total_amt = winning_matched["Amount"].sum()
        payout_amount = winning_total_amt * payout_rate
        net_pl = net_sales - payout_amount
        
        st.markdown(f"### 🎯 ပေါက်ဂဏန်း တိုက်စစ်ချက် ({winning_number.zfill(2)})")
        w1, w2, w3 = st.columns(3)
        w1.metric("ပေါက်ဂဏန်းပေါ် ထိုးငွေစုစုပေါင်း", f"{winning_total_amt:,} MMK")
        w2.metric("ပေးရမည့် ဆုငွေ", f"{payout_amount:,} MMK")
        w3.metric("ဒိုင် အနိုင် / အရှုံး (Net Result)", f"{net_pl:,} MMK")

    st.markdown("---")
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("📊 ဂဏန်းအလိုက် စုစုပေါင်းနှင့် ကန့်သတ်ချက် (Limit Control)")
        grouped = df.groupby("Number")["Amount"].sum().reset_index().sort_values(by="Amount", ascending=False)
        grouped["Allowed"] = grouped["Amount"].apply(lambda x: min(x, max_limit))
        grouped["Excess"] = grouped["Amount"] - grouped["Allowed"]
        st.dataframe(grouped, use_container_width=True)
        
    with col_right:
        st.subheader("📋 ထိုးထားသော စာရင်းများအားလုံး")
        st.dataframe(df, use_container_width=True)
    
    c_exp, c_clr = st.columns([2, 1])
    with c_exp:
        export_df = df.groupby(["Day", "Customer", "Number"])["Amount"].sum().reset_index()
        csv_data = export_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 နေ့အလိုက် ပြည့်စုံသော CSV ဖိုင် Download ဆွဲရန်",
            data=csv_data,
            file_name=f"2D_Weekly_Ledger_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime='text/csv'
        )
    with c_clr:
        if st.button("🗑️ စာရင်း အစအဆုံး ရှင်းလင်းရန်"):
            st.session_state.data = pd.DataFrame(columns=["Day", "Customer", "Number", "Amount"])
            st.rerun()
else:
    st.info("📒 ယခုထိ ထည့်ထားသော စာရင်း မရှိသေးပါ။ အပေါ်ပါ နည်းလမ်း (၃) မျိုးမှ တစ်ခုခုဖြင့် စာရင်းစတင်သွင်းနိုင်ပါသည်။")
