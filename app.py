import streamlit as st
import pandas as pd
import json
from PIL import Image
from datetime import datetime
from google import genai

st.set_page_config(page_title="Myanmar 2D Dealer App", layout="wide")

st.title("🇲🇲 မြန်မာနှစ်လုံးထီ (2D) အလိုအလျောက် စာရင်းချုပ်စနစ်")
st.markdown("---")

# Sidebar Settings
st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များနှင့် ကန့်သတ်ချက်များ")
max_limit = st.sidebar.number_input("ဂဏန်းတစ်လုံးလျှင် အများဆုံးလက်ခံငွေ (Limit)", value=50000, step=5000)
payout_rate = st.sidebar.number_input("ဆုကြေးနှုန်း (ဆ)", value=80, step=5)
winning_number = st.sidebar.text_input("ထွက်မည့် ပေါက်ဂဏန်း", value="")

# Closing Time Control
st.sidebar.markdown("---")
st.sidebar.subheader("⏰ အချိန်ပိတ်ရက် ထိန်းချုပ်မှု")
enable_time_lock = st.sidebar.checkbox("အချိန်ပိတ်ရန် ဖွင့်မည်", value=False)
closing_time = st.sidebar.time_input("ပိတ်မည့် အချိန်", value=datetime.strptime("12:00", "%H:%M").time())

# Blocked Numbers Control
st.sidebar.markdown("---")
st.sidebar.subheader("🚫 ပိတ်ဂဏန်းများ (Comma ခြားရန်)")
blocked_input = st.sidebar.text_input("မခံမည့် ဂဏန်းများ (ဥပမာ: 39, 88)", value="")
blocked_numbers = [num.strip() for num in blocked_input.split(",") if num.strip()]

# Gemini API Key Input in Sidebar
st.sidebar.markdown("---")
st.sidebar.subheader("🤖 AI စနစ် ဆက်တင်")
gemini_api_key = st.sidebar.text_input("Gemini API Key ထည့်ရန်", type="password", value="")

if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=["Customer", "Number", "Amount"])

st.subheader("📝 စာရင်းသွင်းရန် နည်းလမ်းများ")

# Time check warning
is_closed = False
if enable_time_lock:
    current_time = datetime.now().time()
    if current_time >= closing_time:
        is_closed = True
        st.error(f"⚠️ သတိပေးချက် - စာရင်းပိတ်ချိန် ({closing_time.strftime('%H:%M')}) ကျော်လွန်သွားပြီဖြစ်၍ စာရင်းအသစ် လက်မခံတော့ပါ။")

tab1, tab2, tab3 = st.tabs(["✍️ ဂဏန်းနှင့်ငွေ ရောနှောရိုက်ထည့်ရန် (Smart Input)", "💬 ဖောက်သည်ပို့သော စာသား (SMS) ဖြင့် AI ဖတ်ရန်", "📸 စကရင်ရှော့ ပုံတင်၍ AI ဖတ်ရန်"])

with tab1:
    st.markdown("#### 💡 ထိုးမည့်ဂဏန်းနှင့် ငွေပမာဏကို တစ်ခါတည်း ရောနှောရိုက်ထည့်ပါ")
    st.markdown("*(ဥပမာ: `12@1000 10@1000 48_3000 39r1000 269ခွေအပူးပါ 500` စသဖြင့် ရေးနိုင်ပါသည်)*")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        cust_manual = st.text_input("ဖောက်သည် အမည်", key="c_man")
    with col2:
        smart_input_text = st.text_input("ထိုးသည့် 2D ဂဏန်းနှင့် ငွေပမာဏ (ရောနှောလျက်)", key="s_input", placeholder="12@1000 10@1000 48_3000")

    # Smart parsing simulation / preview for manual input
    if smart_input_text:
        # Simple local preview calculation helper or trigger
        pass

    if st.button("➕ အလိုအလျောက် တွက်ချက်၍ စာရင်းသွင်းမည်"):
        if is_closed:
            st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
        elif not cust_manual or not smart_input_text:
            st.warning("ဖောက်သည် အမည်နှင့် ထိုးမည့် စာသားကို ဖြည့်သွင်းပါ။")
        elif not gemini_api_key:
            st.error("ကျေးဇူးပြု၍ ဘယ်ဘက်မီနူး (Sidebar) တွင် Gemini API Key ထည့်သွင်းပေးပါ။")
        else:
            with st.spinner("ဂဏန်းများနှင့် ငွေပမာဏများကို အလိုအလျောက် တွက်ချက်နေပါပြီ..."):
                try:
                    client = genai.Client(api_key=gemini_api_key)
                    
                    prompt = f"""
                    You are an expert AI assistant for a Myanmar 2D lottery business managing 100 numbers (00 to 99).
                    Analyze the following mixed betting input where numbers and amounts are written together (e.g., "12@1000 10@1000 48_3000", "39r1000", "269ခွေအပူးပါ 500", "2ပါ 500", etc.).
                    Extract and expand ALL bets into individual 2-digit numbers ("00" to "99") and their exact amounts in MMK based on these rules:
                    1. Reversed ('r', 'ရ', '@', '&', 'အာ'): E.g., '39@1000' or '65r1000' means both numbers and their reverses with the specified amount.
                    2. Khve / Permutation (အခွေ / ခွေ): 
                       - E.g., '269ခွေ 500' means 6 permutations (26, 29, 62, 69, 92, 96) with 500 MMK each.
                       - E.g., '269ခွေအပူးပါ 500' means 6 permutations plus doubles (22, 66, 99 - total 9 numbers) with 500 MMK each.
                    3. Paat-Thie (ပါတ်သီး variations like 'ပါတ်', 'ပတ်', 'ပါ', 'ပ'): E.g., '2ပါ 500' means all 19 numbers containing '2'.
                    4. Break (ဘရိတ် variations like 'ဘရိတ်', 'B', 'bk'): E.g., '2ဘရိတ် 300' means all 10 numbers.
                    5. Ah-Pu (အပူ): E.g., 'အပူ 500' means all 10 double numbers.
                    6. Nakhat (နက္ခတ်): 07, 70, 18, 81, 24, 42, 35, 53, 69, 96.
                    7. Direct pairs with amounts: E.g., '12@1000' -> Number '12', Amount 1000. '48_3000' -> Number '48', Amount 3000. Ensure each bet amount is between 100 and 50000 MMK.

                    Customer Name is: "{cust_manual}"
                    Input text to analyze:
                    {smart_input_text}

                    Return the result strictly as a JSON list of objects, where each object has:
                    - "Customer": string ("{cust_manual}")
                    - "Number": string (2D 2-digit number from "00" to "99")
                    - "Amount": integer (Bet amount for that specific number in MMK, strictly between 100 and 50000)
                    Do not include any extra text, markdown formatting like ```json, just the raw JSON list.
                    """
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt
                    )
                    
                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                    extracted_data = json.loads(clean_text)
                    
                    valid_rows = []
                    blocked_found = []
                    invalid_amt_found = 0
                    total_added_amount = 0
                    
                    for row in extracted_data:
                        amt_val = int(row["Amount"])
                        if amt_val < 100 or amt_val > 50000:
                            invalid_amt_found += 1
                            continue
                        if row["Number"] in blocked_numbers:
                            blocked_found.append(row["Number"])
                        else:
                            valid_rows.append(row)
                            total_added_amount += amt_val
                    
                    if valid_rows:
                        new_df = pd.DataFrame(valid_rows)
                        st.session_state.data = pd.concat([st.session_state.data, new_df], ignore_index=True)
                        st.success(f"🎉 အောင်မြင်ပါပြီ! ဖောက်သည် ({cust_manual}) ၏ စုစုပေါင်းထိုးငွေ = {total_added_amount:,} ကျပ် (အကွက်ပေါင်း {len(valid_rows)} ကွက်) ထည့်သွင်းပြီးပါပြီ။")
                        if invalid_amt_found > 0:
                            st.info(f"ℹ️ မှတ်ချက်။ ။ ငွေပမာဏ ၁၀၀ မှ ၅၀,၀၀၀ အတွင်း မဟုတ်သော ကွက် {invalid_amt_found} ကွက်ကို ဖယ်ထုတ်ထားပါသည်။")
                        st.rerun()
                    else:
                        st.warning("ဖတ်ရရှိသော စာရင်းများတွင် သတ်မှတ်ချက်နှင့် မကိုက်ညီသည်များ (သို့မဟုတ် ပိတ်ဂဏန်းများ) သာ ပါဝင်နေပါသည်။")
                        
                except Exception as e:
                    st.error(f"စာရင်းဖတ်ရာတွင် အမှားအယွင်း ရှိသွားပါသည်: {e}")

with tab2:
    st.markdown("#### 💬 ဖောက်သည်ပေးပို့ထားသော စာသား (Viber / Messenger SMS) ကို ဤနေရာတွင် ကူးထည့်ပါ")
    sms_customer = st.text_input("ဖောက်သည် အမည် (စာသားအတွက်)", key="sms_cust")
    sms_text_input = st.text_area("ဖောက်သည် ပို့လိုက်သော စာသားမက်ဆေ့ဂျ် အပြည့်အစုံကို ကူးထည့်ပါ (Paste)", height=120)
    
    if st.button("🤖 စာသားကို AI ဖြင့် အလိုအလျောက် ဖြန့်ကျက်ဖတ်မည်"):
        if is_closed:
            st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
        elif not gemini_api_key:
            st.error("ကျေးဇူးပြု၍ ဘယ်ဘက်မီနူး (Sidebar) တွင် Gemini API Key ထည့်သွင်းပေးပါ။")
        elif not sms_customer or not sms_text_input:
            st.error("ဖောက်သည် အမည်နှင့် စာသားကို ဖြည့်သွင်းပေးပါ။")
        else:
            with st.spinner("AI က စာသားမက်ဆေ့ဂျ်များကို တွက်ချက်နေပါပြီ..."):
                try:
                    client = genai.Client(api_key=gemini_api_key)
                    
                    prompt = f"""
                    You are an expert AI assistant for a Myanmar 2D lottery business managing 100 numbers (00 to 99).
                    Analyze the following customer betting message text containing mixed entries and amounts.
                    Extract and expand ALL bets into individual 2-digit numbers ("00" to "99") and their exact amounts in MMK based on standard rules (Reversed, Khve, Paat-Thie, Break, Ah-Pu, Nakhat, Direct pairs).

                    Customer Name is: "{sms_customer}"
                    Message text to analyze:
                    {sms_text_input}

                    Return the result strictly as a JSON list of objects, where each object has:
                    - "Customer": string ("{sms_customer}")
                    - "Number": string (2D 2-digit number from "00" to "99")
                    - "Amount": integer (Bet amount for that specific number in MMK, strictly between 100 and 50000)
                    Do not include any extra text, markdown formatting like ```json, just the raw JSON list.
                    """
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt
                    )
                    
                    clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                    extracted_data = json.loads(clean_text)
                    
                    valid_rows = []
                    blocked_found = []
                    invalid_amt_found = 0
                    total_added_amount = 0
                    
                    for row in extracted_data:
                        amt_val = int(row["Amount"])
                        if amt_val < 100 or amt_val > 50000:
                            invalid_amt_found += 1
                            continue
                        if row["Number"] in blocked_numbers:
                            blocked_found.append(row["Number"])
                        else:
                            valid_rows.append(row)
                            total_added_amount += amt_val
                    
                    if valid_rows:
                        new_df = pd.DataFrame(valid_rows)
                        st.session_state.data = pd.concat([st.session_state.data, new_df], ignore_index=True)
                        st.success(f"🎉 စာသားမက်ဆေ့ဂျ်ထဲမှ စာရင်းများကို အောင်မြင်စွာ ဖတ်ရှုပြီးပါပြီ! စုစုပေါင်းထိုးငွေ = {total_added_amount:,} ကျပ်")
                        st.rerun()
                    else:
                        st.warning("ဖတ်ရရှိသော စာရင်းများတွင် သတ်မှတ်ချက်နှင့် မကိုက်ညီသည်များ ပါဝင်နေပါသည်။")
                        
                except Exception as e:
                    st.error(f"စာရင်းဖတ်ရာတွင် အမှားအယွင်း ရှိသွားပါသည်: {e}")

with tab3:
    uploaded_file = st.file_uploader("ဖောက်သည် ပို့ထားသော စကရင်ရှော့ ပုံတင်ရန်", type=["png", "jpg", "jpeg"])
    
    if uploaded_file is not None:
        img = Image.open(uploaded_file)
        st.image(img, caption="တင်ထားသော စကရင်ရှော့ပုံ", width=300)
        
        if st.button("🤖 ဓာတ်ပုံထဲမှ စာရင်းများကို AI ဖြင့် ဖတ်မည်"):
            if is_closed:
                st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
            elif not gemini_api_key:
                st.error("ကျေးဇူးပြု၍ ဘယ်ဘက်မီနူး (Sidebar) တွင် Gemini API Key ထည့်သွင်းပေးပါ။")
            else:
                with st.spinner("AI က ပုံထဲပါသော စာရင်းများကို တွက်ချက်နေပါပြီ..."):
                    try:
                        client = genai.Client(api_key=gemini_api_key)
                        
                        prompt = """
                        You are an expert AI assistant for a Myanmar 2D lottery business managing 100 numbers (00 to 99). 
                        Analyze the provided image of customer betting messages/slips containing mixed Myanmar 2D shorthand terms and amounts. 
                        Extract and expand ALL bets into individual 2-digit numbers ("00" to "99") and their exact amounts in MMK (between 100 and 50000).

                        Return the result strictly as a JSON list of objects, where each object has:
                        - "Customer": string (Name of customer if visible, otherwise "Unknown")
                        - "Number": string (2D 2-digit number from "00" to "99")
                        - "Amount": integer (Bet amount between 100 and 50000 in MMK)
                        Do not include any extra text, markdown formatting like ```json, just the raw JSON list.
                        """
                        
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[img, prompt]
                        )
                        
                        clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                        extracted_data = json.loads(clean_text)
                        
                        valid_rows = []
                        blocked_found = []
                        total_added_amount = 0
                        for row in extracted_data:
                            amt_val = int(row["Amount"])
                            if amt_val < 100 or amt_val > 50000:
                                continue
                            if row["Number"] in blocked_numbers:
                                blocked_found.append(row["Number"])
                            else:
                                valid_rows.append(row)
                                total_added_amount += amt_val
                        
                        if valid_rows:
                            new_df = pd.DataFrame(valid_rows)
                            st.session_state.data = pd.concat([st.session_state.data, new_df], ignore_index=True)
                            st.success(f"🎉 အောင်မြင်ပါပြီ! စုစုပေါင်းထိုးငွေ = {total_added_amount:,} ကျပ် ထည့်သွင်းပြီးပါပြီ။")
                            st.rerun()
                        else:
                            st.warning("ဖတ်ရရှိသော စာရင်းများတွင် ပိတ်ထားသောဂဏန်းများ သို့မဟုတ် သတ်မှတ်ငွေပမာဏပြင်ပ ရောက်နေသည်များသာ ပါဝင်နေပါသည်။")
                        
                    except Exception as e:
                        st.error(f"စာရင်းဖတ်ရာတွင် အမှားအယွင်း ရှိသွားပါသည်: {e}")

st.markdown("---")
if not st.session_state.data.empty:
    df = st.session_state.data
    
    # Customer Statement Section
    st.subheader("👤 ဖောက်သည်အလိုက် စာရင်းချုပ်ထုတ်ရန်")
    selected_customer = st.selectbox("ဖောက်သည် ရွေးချယ်ပါ", df["Customer"].unique())
    if selected_customer:
        cust_df = df[df["Customer"] == selected_customer]
        cust_total = cust_df["Amount"].sum()
        cust_text = f"✨ *{selected_customer}* ၏ ထိုးငွေစာရင်း\n"
        for idx, row in cust_df.iterrows():
            cust_text += f"• ဂဏန်း {row['Number']} - {row['Amount']:,} ကျပ်\n"
        cust_text += f"\nစုစုပေါင်းထိုးငွေ = {cust_total:,} ကျပ်"
        
        st.text_area("Viber / Telegram သို့ ကူးယူရန် စာသား (Copy Statement)", cust_text, height=150)

    grouped = df.groupby("Number")["Amount"].sum().reset_index()
    grouped["Allowed"] = grouped["Amount"].apply(lambda x: min(x, max_limit))
    grouped["Excess"] = grouped["Amount"] - grouped["Allowed"]
    
    total_sales = df["Amount"].sum()
    st.metric("စုစုပေါင်း ရောင်းရငွေ", f"{total_sales:,} MMK")
    
    st.subheader("📊 ဂဏန်းအလိုက် စုစုပေါင်းနှင့် ကန့်သတ်ချက် (Limit Control)")
    st.dataframe(grouped, use_container_width=True)
    
    st.subheader("📋 ထိုးထားသော စာရင်းများအားလုံး")
    st.dataframe(df, use_container_width=True)
    
    # Export to CSV Button
    csv_data = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 စာရင်းအားလုံးကို CSV ဖိုင်ဖြင့် Download ဆွဲရန်",
        data=csv_data,
        file_name=f"2D_Sales_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime='text/css'
    )
    
    if winning_number:
        st.subheader(f"🎯 ပေါက်ဂဏန်းတိုက်စစ်ခြင်း ({winning_number})")
        winners = df[df["Number"] == str(winning_number).zfill(2)]
        if not winners.empty:
            st.success("ကံထူးရှင်များ ရှိပါသည်။")
            st.dataframe(winners, use_container_width=True)
            payout = winners["Amount"].sum() * payout_rate
            profit = total_sales - payout
            st.metric("ပေးရမည့် ဆုငွေ", f"{payout:,} MMK")
            st.metric("ဒိုင် အမြတ် / အရှုံး", f"{profit:,} MMK")
        else:
            st.info("ကံထူးရှင် မရှိပါ။")

    if st.button("🗑️ စာရင်း အစအဆုံး ရှင်းလင်းရန်"):
        st.session_state.data = pd.DataFrame(columns=["Customer", "Number", "Amount"])
        st.rerun()
else:
    st.info("ယခုထိ ထည့်ထားသော စာရင်း မရှိသေးပါ။")
