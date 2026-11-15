import streamlit as st
import pandas as pd
import json
from PIL import Image
from datetime import datetime
from google import genai

st.set_page_config(page_title="Myanmar 2D Dealer App", layout="wide")

st.title("🇲🇲 မြန်မာနှစ်လုံးထီ (2D) အလိုအလျောက် စာရင်းချုပ်စနစ် (အပြည့်အစုံ)")
st.markdown("---")

# Sidebar Settings
st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များနှင့် ကန့်သတ်ချက်များ")
max_limit = st.sidebar.number_input("ဂဏန်းတစ်လုံးလျှင် အများဆုံးလက်ခံငွေ", value=50000, step=5000)
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
st.sidebar.subheader("🤖 AI ဓာတ်ပုံဖတ်စနစ် ဆက်တင်")
gemini_api_key = st.sidebar.text_input("Gemini API Key ထည့်ရန်", type="password", value="")

if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=["Customer", "Number", "Amount"])

# Helper function to expand Myanmar 2D shorthand terms manually
def expand_shorthand(customer, text_input, amount):
    text = text_input.strip().lower()
    entries = []
    
    # Check for 'r' or 'ရ' (Reversed) e.g., "39r"
    if 'r' in text or 'ရ' in text:
        num_part = ''.join(filter(str.isdigit, text))
        if len(num_part) == 2:
            n1, n2 = num_part, num_part[::-1]
            entries.append({"Customer": customer, "Number": n1, "Amount": amount})
            if n1 != n2:
                entries.append({"Customer": customer, "Number": n2, "Amount": amount})
            return entries

    # Check for Paat-Thie (ပါတ်)
    if 'ပါတ်' in text or 'ပတ်' in text:
        p_num = ''.join(filter(str.isdigit, text))
        if len(p_num) == 1:
            for i in range(10):
                for n_str in [f"{p_num}{i}", f"{i}{p_num}"]:
                    if n_str not in [e["Number"] for e in entries]:
                        entries.append({"Customer": customer, "Number": n_str, "Amount": amount})
            return entries

    # Check for Ah-Pu (အပူ)
    if 'ပူ' in text:
        for i in range(10):
            entries.append({"Customer": customer, "Number": f"{i}{i}", "Amount": amount})
        return entries

    # Check for Nakhat (နက္ခတ်)
    if 'နက္ခတ်' in text or 'နက္ခ' in text:
        nakhat_pairs = ["07", "70", "18", "81", "24", "42", "35", "53", "69", "96"]
        for n_str in nakhat_pairs:
            entries.append({"Customer": customer, "Number": n_str, "Amount": amount})
        return entries

    # Direct 2-digit number
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

tab1, tab2 = st.tabs(["✍️ လက်ဖြင့် / အတိုကောက် ရိုက်ထည့်ရန်", "📸 ဓာတ်ပုံ/စကရင်ရှော့ တင်၍ AI ဖြင့် ဖတ်ရန်"])

with tab1:
    col1, col2, col3 = st.columns(3)
    with col1:
        cust = st.text_input("ဖောက်သည် အမည်", key="c1")
    with col2:
        num_text = st.text_input("ထိုးသည့် 2D ဂဏန်း (သို့ အတိုကောက် 39r, 5ပတ်, နက္ခတ်)", key="n1")
    with col3:
        amt = st.number_input("ထိုးငွေ ပမာဏ (တစ်ကွက်စာ)", min_value=0, step=1000, key="a1")

    if st.button("➕ စာရင်းသွင်းမည်"):
        if is_closed:
            st.error("စာရင်းပိတ်ချိန် ကျော်လွန်နေပြီဖြစ်ပါသည်။")
        elif cust and num_text and amt > 0:
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
                with st.spinner("AI က မြန်မာ့ 2D ဝေါဟာရများနှင့် နက္ခတ်များကို တွက်ချက်ဖတ်ရှုနေပါပြီ..."):
                    try:
                        client = genai.Client(api_key=gemini_api_key)
                        
                        prompt = """
                        You are an expert AI assistant for a Myanmar 2D lottery business managing 100 numbers (00 to 99). 
                        Analyze the provided image of customer betting messages/slips which may contain Myanmar 2D shorthand terms. 
                        Extract and expand ALL bets into individual 2-digit numbers ("00" to "99") and their exact amounts in MMK based on these rules:
                        1. Reversed ('r' or 'ရ'): E.g., '39 r 1000' means both '39' and '93' with 1000 MMK each.
                        2. Paat-Thie (ပါတ်သီး): E.g., '5 ပါတ် 1000' means all 19 numbers containing '5' with 1000 MMK each.
                        3. Ah-Pu (အပူ): E.g., 'အပူ 500' means all 10 double numbers (00 to 99) with 500 MMK each.
                        4. Nakhat (နက္ခတ်): 07, 70, 18, 81, 24, 42, 35, 53, 69, 96 with specified amount each.
                        5. Direct numbers: E.g., '65r1000' -> '65' (1000) and '56' (1000).

                        Return the result strictly as a JSON list of objects, where each object has:
                        - "Customer": string (Name of customer if visible, otherwise "Unknown")
                        - "Number": string (2D 2-digit number from "00" to "99")
                        - "Amount": integer (Bet amount for that specific number in MMK)
                        Do not include any extra text, markdown formatting like ```json, just the raw JSON list.
                        """
                        
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[img, prompt]
                        )
                        
                        clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                        extracted_data = json.loads(clean_text)
                        
                        # Filter out blocked numbers
                        filtered_data = [row for row in extracted_data if row["Number"] not in blocked_numbers]
                        
                        if filtered_data:
                            new_df = pd.DataFrame(filtered_data)
                            st.session_state.data = pd.concat([st.session_state.data, new_df], ignore_index=True)
                            st.success("🎉 ပုံထဲမှ စာရင်းများကို အောင်မြင်စွာ ဖတ်ရှုပြီး ဇယားထဲ ထည့်သွင်းပြီးပါပြီ!")
                            st.rerun()
                        else:
                            st.warning("ဖတ်ရရှိသော စာရင်းများထဲတွင် ပိတ်ထားသောဂဏန်းများသာ ပါဝင်နေပါသည်။")
                        
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
    st.info("ယခုထိ ထည့်ထားသော စာရင်း မရှိသေးပါ။")
