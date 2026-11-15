import streamlit as st
import pandas as pd

st.set_page_config(page_title="Myanmar 2D Dealer App", layout="wide")

st.title("🇲🇲 မြန်မာနှစ်လုံးထီ (2D) အလိုအလျောက် စာရင်းချုပ်စနစ်")
st.markdown("---")

st.sidebar.header("⚙️ ဒိုင် ဆက်တင်များ")
max_limit = st.sidebar.number_input("ဂဏန်းတစ်လုံးလျှင် အများဆုံးလက်ခံငွေ", value=50000, step=5000)
payout_rate = st.sidebar.number_input("ဆုကြေးနှုန်း (ဆ)", value=80, step=5)
winning_number = st.sidebar.text_input("ထွက်မည့် ပေါက်ဂဏန်း", value="")

if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame(columns=["Customer", "Number", "Amount"])

st.subheader("📝 စာရင်းသွင်းရန်")
col1, col2, col3 = st.columns(3)
with col1:
    cust = st.text_input("ဖောက်သည် အမည်")
with col2:
    num = st.text_input("ထိုးသည့် 2D ဂဏန်း")
with col3:
    amt = st.number_input("ထိုးငွေ ပမာဏ", min_value=0, step=1000)

if st.button("➕ စာရင်းသွင်းမည်"):
    if cust and num and amt > 0:
        new_row = {"Customer": cust, "Number": str(num).zfill(2), "Amount": amt}
        st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([new_row])], ignore_index=True)
        st.success(f"ထည့်သွင်းပြီးပါပြီ - {cust} ({num} / {amt:,} ကျပ်)")
    else:
        st.warning("အချက်အလက် အပြည့်အစုံ ဖြည့်ပါ။")

st.markdown("---")
if not st.session_state.data.empty:
    df = st.session_state.data
    grouped = df.groupby("Number")["Amount"].sum().reset_index()
    grouped["Allowed"] = grouped["Amount"].apply(lambda x: min(x, max_limit))
    grouped["Excess"] = grouped["Amount"] - grouped["Allowed"]
    
    total_sales = df["Amount"].sum()
    st.metric("စုစုပေါင်း ရောင်းရငွေ", f"{total_sales:,} MMK")
    
    st.subheader("📊 ဂဏန်းအလိုက် စုစုပေါင်းနှင့် ကန့်သတ်ချက်")
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
