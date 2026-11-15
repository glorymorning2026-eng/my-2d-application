import streamlit as st
import pandas as pd
import json
from PIL import Image
from google import genai

st.set_page_config(page_title="Myanmar 2D Dealer App", layout="wide")

st.title("🇲🇲 မြန်မာနှစ်လုံးထီ (2D) အလိုအလျောက် စာရင်းချုပ်စနစ်")
st.markdown("---")

# Sidebar Settings
st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များ")
max_limit = st.sidebar.number_input("ဂဏန်းတစ်လုံးလျှင် အများဆုံးလက်ခံငွေ", value=50000, step=5000)
payout_rate = st.sidebar.number_input("ဆုကြေးနှုန်း (ဆ)", value=80, step=5)
winning_number = st.sidebar.text_input("ထွက်မည့် ပေါက်ဂဏန်း", value="")

# Gemini API Key Input in Sidebar
st.sidebar.markdown("---")
st.sidebar.header("🤖 AI ဓာတ်ပုံဖတ်စနစ် ဆက်တင်")
gemini_api_key = st.sidebar.text_input("Gemini API Key ထည့်ရန်", type="password", value="")

if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=["Customer", "Number", "Amount"])

st.subheader("📝 စာရင်းသွင်းရန် နည်းလမ်းများ")

tab1, tab2 = st.tabs(["✍️ လက်ဖြင့် ရိုက်ထည့်ရန်", "📸 ဓာတ်ပုံ/စကရင်ရှော့ တင်၍ AI ဖြင့် ဖတ်ရန်"])

with tab1:
    col1, col2, col3 = st.columns(3)
    with col1:
        cust = st.text_input("ဖောက်သည် အမည်", key="c1")
    with col2:
        num = st.text_input("ထိုးသည့် 2D ဂဏန်း", key="n1")
    with col3:
        amt = st.number_input("ထိုးငွေ ပမာဏ", min_value=0, step=1000, key="a1")

    if st.button("➕ စာရင်းသွင်းမည်"):
        if cust and num and amt > 0:
            new_row = {"Customer": cust, "Number": str(num).zfill(2), "Amount": amt}
            st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
            st.success(f"ထည့်သွင်းပြီးပါပြီ - {cust} ({num} / {amt:,} ကျပ်)")
        else:
            st.warning("အချက်အလက် အပြည့်အစုံ ဖြည့်ပါ။")

with tab2:
    uploaded_file = st.file_uploader("ဖောက်သည် ပို့ထားသော စကရင်ရှော့ ပုံတင်ရန်", type=["png", "jpg", "jpeg"])
    
    if uploaded_file is not None:
        img = Image.open(uploaded_file)
        st.image(img, caption="တင်ထားသော စကရင်ရှော့ပုံ", width=300)
        
        if st.button("🤖 ဓာတ်ပုံထဲမှ စာရင်းများကို AI ဖြင့် ဖတ်မည်"):
            if not gemini_api_key:
                st.error("ကျေးဇူးပြု၍ ဘယ်ဘက်မီနူး (Sidebar) တွင် Gemini API Key ထည့်သွင်းပေးပါ။")
            else:
                with st.spinner("AI က ဓာတ်ပုံထဲပါသော နံပါတ်နှင့် ငွေပမာဏများကို ဖတ်ရှုနေပါပြီ..."):
                    try:
                        client = genai.Client(api_key=gemini_api_key)
                        
                        prompt = """
                        Analyze this image of a 2D lottery slip or customer message containing numbers like '65r1000' (number 65, amount 1000), 'အပူ10ကွက်1000' or similar entries. Extract all individual betting entries.
                        Return the result strictly as a JSON list of objects, where each object has:
                        - "Customer": string (Name of customer if available, otherwise "Unknown")
                        - "Number": string (2D 2-digit number, e.g., "65", "63", "05")
                        - "Amount": integer (Bet amount in MMK, e.g., 1000, 750)
                        Do not include any extra text, markdown formatting like ```json, just the raw JSON list.
                        """
                        
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[img, prompt]
                        )
                        
                        clean_text = response.text.strip().replace("```json", "").replace("```", "").strip()
                        extracted_data = json.loads(clean_text)
                        
                        new_df = pd.DataFrame(extracted_data)
                        st.session_state.data = pd.concat([st.session_state.data, new_df], ignore_index=True)
                        
                        st.success("🎉 ပုံထဲမှ စာရင်းများကို အောင်မြင်စွာ ဖတ်ရှုပြီး ဇယားထဲ ထည့်သွင်းပြီးပါပြီ!")
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"စာရင်းဖတ်ရာတွင် အမှားအယွင်း ရှိသွားပါသည်: {e}")

st.markdown("---")
if not st.session_state.data.empty:
    df = st.session_state.data
    grouped = df.groupby("Number")["Amount"].sum().reset_index()
    grouped["Allowed"] = grouped["Amount"].apply(lambda x: min(x, max_limit))
    grouped["Excess"] = grouped["Amount"] - grouped["Allowed"]
    
    total_sales = df["Amount"].sum()
    st.metric("စုစုပေါင်း ရောင်းရငွေ", f"{total_sales:,} MMK")
    
    st.subheader("📊 ဂဏန်းအလိုက် စုစုပေါင်းနှင့် ကန့်သတ်ချက် (Limit Control)")
    st.dataframe(grouped, use_container_width=True)
    
    st.subheader("📋 ထိုးထားသော စာရင်းများအားလုံး")
    st.dataframe(df, use_container_width=True)
    
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
