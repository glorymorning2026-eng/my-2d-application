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
st.sidebar.subheader("🤖 AI Interactions API ဆက်တင်")
gemini_api_key = st.sidebar.text_input("Gemini API Key ထည့်ရန်", type="password", value="")

if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=["Customer", "Number", "Amount"])

# Helper function to expand shorthand manually if needed
def expand_shorthand(customer, text_input, amount):
    text = text_input.strip().lower()
    entries = []
    
    if amount < 100 or amount > 50000:
        return "invalid_amount"

    if 'r' in text or 'ရ' in text:
        num_part = ''.join(filter(str.isdigit, text))
        if len(num_part) == 2:
            n1, n2 = num_part, num_part[::-1]
            entries.append({"Customer": customer, "Number": n1, "Amount": amount})
            if n1 != n2:
                entries.append({"Customer": customer, "Number": n2, "Amount": amount})
            return entries

    if 'ပါတ်' in text or 'ပတ်' in text:
        p_num = ''.join(filter(str.isdigit, text))
        if len(p_num) == 1:
            for i in range(10):
                for n_str in [f"{p_num}{i}", f"{i}{p_num}"]:
                    if n_str not in [e["Number"] for e in entries]:
                        entries.append({"Customer": customer, "Number": n_str, "Amount": amount})
            return entries

    if 'ပူ' in text:
        for i in range(10):
            entries.append({"Customer": customer, "Number": f"{i}{i}", "Amount": amount})
        return entries

    if 'နက္ခတ်' in text or 'နက္ခ' in text:
        nakhat_pairs = ["07", "70", "18", "81", "24", "42", "35", "53", "69", "96"]
        for n_str in nakhat_pairs:
            entries.append({"Customer": customer, "Number": n_str, "Amount": amount})
        return entries

    num_clean = ''.join(filter(str.isdigit, text))
    if len(num_clean) == 2:
        entries.append({"Customer": customer, "Number": num_clean, "Amount": amount})
        
    return entries

st.subheader("📝 စာရင်းသွင်းရန် နည်းလမ်းများ")

# Time check warning
is_closed = False
if enable_time_lock:
    current_time = datetime.now().time()
    if current_time >= closing_time:
        is_closed = True
        st.error(f"⚠️ သတိပေးချက် - စာရင်းပိတ်ချိန် ({closing_time.strftime('%H:%M')}) ကျော်လွန်သွားပြီဖြစ်၍ စာရင်းအသစ် လက်မခံတော့ပါ။")

tab1, tab2, tab3 = st.tabs(["✍️ လက်ဖြင့် / အတိုကောက် ရိုက်ထည့်ရန်", "💬 ဖောက်သည်ပို့သော စာသား (SMS) ဖြင့် AI ဖတ်ရန်", "📸 စကရင်ရှော့ ပုံတင်၍ AI ဖတ်ရန်"])

with tab1:
    col1, col2, col3 = st.columns(3)
    with col1:
        cust = st.text_input("ဖောက်သည် အမည်", key="c1")
    with col2:
        num_text = st.text_input("ထိုးသည့် 2D ဂဏန်း (သို့ အတိုကောက်)", key="n1")
    with col3:
        amt = st.number_input("ထိုးငွေ ပမာဏ (၁၀၀ မှ ၅၀,၀၀၀ ထိ)", min_value=0, step=1000, key="a1")

    if st.button("➕ စာရင်းသွင်းမည်"):
        if is_closed:
            st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
        elif cust and num_text and amt > 0:
            if amt < 100 or amt > 50000:
                st.error("❌ ထိုးငွေပမာဏသည် တစ်ကွက်လျှင် အနည်းဆုံး ၁၀၀ ကျပ်မှ အများဆုံး ၅၀,၀၀၀ ကျပ် အတွင်းသာ ဖြစ်ရပါမည်။")
            else:
                expanded_rows = expand_shorthand(cust, num_text, amt)
                if expanded_rows:
                    valid_rows = []
                    blocked_found = []
                    for row in expanded_rows:
                        if row["Number"] in blocked_numbers:
                            blocked_found.append(row["Number"])
                        else:
                            valid_rows.append(row)
                    
                    if valid_rows:
                        new_df = pd.DataFrame(valid_rows)
                        st.session_state.data = pd.concat([st.session_state.data, new_df], ignore_index=True)
                        st.success(f"✅ စာရင်းသွင်းပြီးပါပြီ - ဖောက်သည် ({cust}) / စုစုပေါင်း အကွက် {len(valid_rows)} ကွက်")
                    if blocked_found:
                        st.warning(f"🚫 ပိတ်ထားသောဂဏန်းများ ပါဝင်နေသဖြင့် ထည့်မပေးပါ။ (ပိတ်ဂဏန်း: {', '.join(set(blocked_found))})")
                else:
                    st.warning("ထည့်သွင်းထားသော ဂဏန်းပုံစံ မမှန်ကန်ပါ။")
        else:
            st.warning("အချက်အလက် အပြည့်အစုံ ဖြည့်ပါ။")

with tab2:
    st.markdown("#### 💬 ဖောက်သည်ပေးပို့ထားသော စာသား (Viber / Messenger SMS) ကို ဤနေရာတွင် ကူးထည့်ပါ")
    st.markdown("*(ဥပမာ: 3ပါတ် 500, 34=300, 39=900, 65r1000 စသည်ဖြင့် ရောနှောပါဝင်နိုင်သည်)*")
    
    sms_customer = st.text_input("ဖောက်သည် အမည် (စာသားအတွက်)", key="sms_cust")
    sms_text_input = st.text_area("ဖောက်သည် ပို့လိုက်သော စာသားမက်ဆေ့ဂျ် အပြည့်အစုံကို ကူးထည့်ပါ (Paste)", height=120)
    
    if st.button("🤖 စာသားကို AI (Interactions API) ဖြင့် အလိုအလျောက် ဖြန့်ကျက်ဖတ်မည်"):
        if is_closed:
            st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
        elif not gemini_api_key:
            st.error("ကျေးဇူးပြု၍ ဘယ်ဘက်မီနူး (Sidebar) တွင် Gemini API Key ထည့်သွင်းပေးပါ။")
        elif not sms_customer or not sms_text_input:
            st.error("ဖောက်သည် အမည်နှင့် စာသားကို ဖြည့်သွင်းပေးပါ။")
        else:
            with st.spinner("AI က Interactions API ကိုသုံး၍ ရောနှောနေသော စာသားများကို တွက်ချက်နေပါပြီ..."):
                try:
                    client = genai.Client(api_key=gemini_api_key)
                    
                    prompt = f"""
                    You are an expert AI assistant for a Myanmar 2D lottery business managing 100 numbers (00 to 99).
                    Analyze the following customer betting message text which may contain mixed entries, multiple items separated by commas, spaces, or lines (e.g., "3ပါတ် 500, 34=300, 39=900, 65r1000", etc.).
                    Extract and expand ALL bets into individual 2-digit numbers ("00" to "99") and their exact amounts in MMK based on these rules:
                    1. Reversed ('r' or 'ရ'): E.g., '39 r 1000' or '65r1000' means both numbers and their reverses with the specified amount each.
                    2. Paat-Thie (ပါတ်သီး): E.g., '3 ပါတ် 500' means all 19 numbers containing '3' with 500 MMK each.
                    3. Ah-Pu (အပူ): E.g., 'အပူ 500' means all 10 double numbers with 500 MMK each.
                    4. Nakhat (နက္ခတ်): 07, 70, 18, 81, 24, 42, 35, 53, 69, 96 with specified amount each.
                    5. Direct numbers with mixed amounts: E.g., '34=300' or '39-900' or '12ခွက် 1000'. Ensure each extracted bet amount is between 100 and 50000 MMK.

                    Customer Name is: "{sms_customer}"
                    Message text to analyze:
                    {sms_text_input}

                    Return the result strictly as a JSON list of objects, where each object has:
                    - "Customer": string ("{sms_customer}")
                    - "Number": string (2D 2-digit number from "00" to "99")
                    - "Amount": integer (Bet amount for that specific number in MMK, strictly between 100 and 50000)
                    Do not include any extra text, markdown formatting like ```json, just the raw JSON list.
                    """
                    
                    # Using Interactions API (client.interactions.create) with latest recommended model
                    interaction = client.interactions.create(
                        model='gemini-4-hydrogen',
                        input=prompt
                    )
                    
                    clean_text = interaction.output_text.strip().replace("```json", "").replace("```", "").strip()
                    extracted_data = json.loads(clean_text)
                    
                    valid_rows = []
                    blocked_found = []
                    invalid_amt_found = 0
                    
                    for row in extracted_data:
                        amt_val = int(row["Amount"])
                        if amt_val < 100 or amt_val > 50000:
                            invalid_amt_found += 1
                            continue
                        if row["Number"] in blocked_numbers:
                            blocked_found.append(row["Number"])
                        else:
                            valid_rows.append(row)
                    
                    if valid_rows:
                        new_df = pd.DataFrame(valid_rows)
                        st.session_state.data = pd.concat([st.session_state.data, new_df], ignore_index=True)
                        st.success("🎉 စာသားမက်ဆေ့ဂျ်ထဲမှ စာရင်းများကို အောင်မြင်စွာ ဖတ်ရှုပြီး ဇယားထဲ ထည့်သွင်းပြီးပါပြီ!")
                        if invalid_amt_found > 0:
                            st.info(f"ℹ️ မှတ်ချက်။ ။ ငွေပမာဏ ၁၀၀ မှ ၅၀,၀၀၀ အတွင်း မဟုတ်သော ကွက် {invalid_amt_found} ကွက်ကို ဖယ်ထုတ်ထားပါသည်။")
                        st.rerun()
                    else:
                        st.warning("ဖတ်ရရှိသော စာရင်းများတွင် သတ်မှတ်ချက်နှင့် မကိုက်ညီသည်များ (သို့မဟုတ် ပိတ်ဂဏန်းများ) သာ ပါဝင်နေပါသည်။")
                        
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
                with st.spinner("AI က ပုံထဲပါသော ရောနှောစာရင်းများကို တွက်ချက်ဖတ်ရှုနေပါပြီ..."):
                    try:
                        client = genai.Client(api_key=gemini_api_key)
                        
                        prompt = """
                        You are an expert AI assistant for a Myanmar 2D lottery business managing 100 numbers (00 to 99). 
                        Analyze the provided image of customer betting messages/slips containing mixed Myanmar 2D shorthand terms and amounts. 
                        Extract and expand ALL bets into individual 2-digit numbers ("00" to "99") and their exact amounts in MMK (must be between 100 and 50000) based on these rules:
                        1. Reversed ('r' or 'ရ'): E.g., '39 r 1000' means both numbers with 1000 MMK each.
                        2. Paat-Thie (ပါတ်သီး): E.g., '3 ပါတ် 500' means all 19 numbers with 500 MMK each.
                        3. Ah-Pu (အပူ): E.g., 'အပူ 500' means all 10 double numbers with 500 MMK each.
                        4. Nakhat (နက္ခတ်): 07, 70, 18, 81, 24, 42, 35, 53, 69, 96 with specified amount each.
                        5. Direct mixed numbers with amounts.

                        Return the result strictly as a JSON list of objects, where each object has:
                        - "Customer": string (Name of customer if visible, otherwise "Unknown")
                        - "Number": string (2D 2-digit number from "00" to "99")
                        - "Amount": integer (Bet amount between 100 and 50000 in MMK)
                        Do not include any extra text, markdown formatting like ```json, just the raw JSON list.
                        """
                        
                        # Note: For image processing, client.interactions.create supports contents/files depending on SDK version, 
                        # falling back or matching the standard interactions pattern.
                        interaction = client.interactions.create(
                            model='gemini-4-hydrogen',
                            input=[img, prompt]
                        )
                        
                        clean_text = interaction.output_text.strip().replace("```json", "").replace("```", "").strip()
                        extracted_data = json.loads(clean_text)
                        
                        valid_rows = []
                        blocked_found = []
                        for row in extracted_data:
                            amt_val = int(row["Amount"])
                            if amt_val < 100 or amt_val > 50000:
                                continue
                            if row["Number"] in blocked_numbers:
                                blocked_found.append(row["Number"])
                            else:
                                valid_rows.append(row)
                        
                        if valid_rows:
                            new_df = pd.DataFrame(valid_rows)
                            st.session_state.data = pd.concat([st.session_state.data, new_df], ignore_index=True)
                            st.success("🎉 ပုံထဲမှ စာရင်းများကို အောင်မြင်စွာ ဖတ်ရှုပြီး ဇယားထဲ ထည့်သွင်းပြီးပါပြီ!")
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
        mime='text/csv'
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
    st.info("ယခုထိုက် ထည့်ထားသော စာရင်း မရှိသေးပါ။")
