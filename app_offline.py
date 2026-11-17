import re
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "2d_ledger.sqlite3"

st.set_page_config(page_title="Myanmar 2D Offline Ledger", page_icon="", layout="wide")
st.markdown("""
<style>
.main-header {border-left:6px solid #D97706;padding-left:15px;margin-bottom:20px}
.stButton>button {border-radius:8px;font-weight:700}
div[data-testid="stMetricValue"] {color:#D97706}
</style>
""", unsafe_allow_html=True)
st.markdown('<div class="main-header"><h1> Myanmar 2D Offline Ledger</h1><p>API Key မလိုပါ  Data အားလုံးကို ဤစက်ထဲတွင်သာ သိမ်းဆည်းသည်</p></div>', unsafe_allow_html=True)


def db():
    con = sqlite3.connect(DB_PATH)
    con.execute("""CREATE TABLE IF NOT EXISTS bets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer TEXT NOT NULL,
        number TEXT NOT NULL,
        amount INTEGER NOT NULL,
        created_at TEXT NOT NULL
    )""")
    con.commit()
    return con


def myanmar_digits(text):
    return str(text).translate(str.maketrans("၀၁၂၃၄၅၆၇၈၉", "0123456789"))


def number(v):
    digits = re.sub(r"\D", "", myanmar_digits(v or ""))
    if not digits:
        return None
    n = int(digits)
    return f"{n:02d}" if 0 <= n <= 99 else None


def amount(v):
    digits = re.sub(r"\D", "", myanmar_digits(v or ""))
    return int(digits) if digits else None


def add_pair(rows, customer, n, amt):
    n = number(n)
    amt = amount(amt)
    if n and amt is not None:
        rows.append({"Customer": customer, "Number": n, "Amount": amt})


def parse_local(text, customer):
    """Local rule parser; no network, API key, or cloud service is used."""
    text = myanmar_digits(text or "")
    rows = []
    # 12=1000, 48_3000, 39@500, 65r1000, and 12 1000
    direct = re.compile(r"(?<!\d)(\d{1,2})\s*(=|_|@|ရ|r|\s)\s*(\d{2,7})(?!\d)", re.I)
    consumed = []
    for m in direct.finditer(text):
        add_pair(rows, customer, m.group(1), m.group(3))
        consumed.append(m.span())

    # 269ခွေ 500 / 269ခွေအပူးပါ 500
    for m in re.finditer(r"(\d{2,3})\s*(ခွေ|အခွေ)\s*(အပူးပါ|အပူးအပါ)?\s*(\d{2,7})", text, re.I):
        digits = list(dict.fromkeys(m.group(1)))
        for a in digits:
            for b in digits:
                if a != b or m.group(3):
                    add_pair(rows, customer, a + b, m.group(4))

    # 2ပါ 500 / 3ပါတ် 250 / 2ပ 500  all 19 numbers containing digit
    for m in re.finditer(r"([0-9])\s*(ပါတ်|ပတ်|ပါ|ပ)\s*(\d{2,7})", text, re.I):
        d, amt = m.group(1), m.group(3)
        for n in range(100):
            s = f"{n:02d}"
            if d in s:
                add_pair(rows, customer, s, amt)

    # 2ဘရိတ် 300 / 2B 300 / 2bk 300
    for m in re.finditer(r"([0-9])\s*(ဘရိတ်|bk|b)\s*(\d{2,7})", text, re.I):
        d, amt = int(m.group(1)), m.group(3)
        for n in range(100):
            s = f"{n:02d}"
            if (int(s[0]) + int(s[1])) % 10 == d:
                add_pair(rows, customer, s, amt)

    # အပူ 500  all doubles
    for m in re.finditer(r"အပူ\s*(\d{2,7})", text, re.I):
        for d in range(10):
            add_pair(rows, customer, f"{d}{d}", m.group(1))

    # နက္ခတ် 500 (optional amount)
    nakhat = ["07", "70", "18", "81", "24", "42", "35", "53", "69", "96"]
    for m in re.finditer(r"နက္ခတ်\s*(\d{2,7})?", text, re.I):
        amt = m.group(1) or "0"
        if int(amt) >= 100:
            for n in nakhat:
                add_pair(rows, customer, n, amt)

    # De-duplicate only exact repeated records from overlapping rules.
    unique = {(r["Customer"], r["Number"], r["Amount"]): r for r in rows}
    return list(unique.values())


def load_df():
    con = db()
    df = pd.read_sql_query("SELECT id, customer AS Customer, number AS Number, amount AS Amount, created_at AS CreatedAt FROM bets ORDER BY id DESC", con)
    con.close()
    return df


def save_rows(rows, blocked, max_limit, enforce):
    con = db()
    existing = pd.read_sql_query("SELECT number, SUM(amount) total FROM bets GROUP BY number", con)
    totals = dict(zip(existing.number, existing.total))
    accepted, blocked_out, invalid, limited = [], [], [], []
    for r in rows:
        n, a = number(r.get("Number")), amount(r.get("Amount"))
        if not n or a is None or a < 100 or a > 50000:
            invalid.append(r); continue
        if n in blocked:
            blocked_out.append(n); continue
        if enforce and totals.get(n, 0) + a > max_limit:
            limited.append(n); continue
        totals[n] = totals.get(n, 0) + a
        accepted.append((str(r.get("Customer") or "Unknown"), n, a, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    if accepted:
        con.executemany("INSERT INTO bets(customer,number,amount,created_at) VALUES(?,?,?,?)", accepted)
        con.commit()
    con.close()
    return accepted, blocked_out, invalid, limited


def display_result(result):
    accepted, blocked, invalid, limited = result
    if accepted:
        st.success(f" {len(accepted)} ကွက် / {sum(x[2] for x in accepted):,} ကျပ် ထည့်ပြီးပါပြီ။")
    else:
        st.warning("ထည့်သွင်းနိုင်သော စာရင်းမတွေ့ပါ။")
    if blocked: st.warning(" ပိတ်ဂဏန်း: " + ", ".join(sorted(set(blocked))))
    if limited: st.error(" Limit ကျော်သဖြင့် မထည့်ထားသောဂဏန်း: " + ", ".join(sorted(set(limited))))
    if invalid: st.info(f" မမှန်ကန်သော row {len(invalid)} ခု ဖယ်ထုတ်ထားပါသည်။")


# Settings
st.sidebar.header(" ဆက်တင်များ")
max_limit = int(st.sidebar.number_input("ဂဏန်းတစ်လုံးအတွက် အများဆုံးငွေ", min_value=100, value=50000, step=5000))
enforce_limit = st.sidebar.checkbox("Limit ကျော်လျှင် မထည့်ရန်", True)
payout_rate = float(st.sidebar.number_input("ဆုကြေးနှုန်း (ဆ)", min_value=0.0, value=80.0, step=5.0))
winning = st.sidebar.text_input("ပေါက်ဂဏန်း (00-99)")
blocked = {number(x) for x in st.sidebar.text_input("ပိတ်ဂဏန်းများ (39,88)").split(",") if number(x)}

# Import/export
with st.expander(" Local Backup / Restore", expanded=False):
    df_now = load_df()
    st.caption(f"Database: {DB_PATH} | စာရင်း {len(df_now)} ကြောင်း")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(" CSV Backup", df_now.to_csv(index=False).encode("utf-8-sig"), "2d_offline_backup.csv", "text/csv")
    with c2:
        restore = st.file_uploader("CSV ပြန်သွင်းရန်", type="csv")
        if restore and st.button(" Restore", key="restore"):
            try:
                old = pd.read_csv(restore)
                con = db()
                for r in old.to_dict("records"):
                    if number(r.get("Number")) and amount(r.get("Amount")) is not None:
                        con.execute("INSERT INTO bets(customer,number,amount,created_at) VALUES(?,?,?,?)", (str(r.get("Customer", "Unknown")), number(r["Number"]), amount(r["Amount"]), str(r.get("CreatedAt", datetime.now()))))
                con.commit(); con.close(); st.success("Restore ပြီးပါပြီ။"); st.rerun()
            except Exception as exc: st.error(f"Restore မအောင်မြင်ပါ: {exc}")

# Input tabs
t1, t2, t3 = st.tabs([" Local Smart Input", " Local SMS Parser", " Optional Local OCR"])
with t1:
    customer = st.text_input("ဖောက်သည်အမည်", key="local_customer")
    text = st.text_input("Input", placeholder="12@1000 48_3000 269ခွေ 500 2ပါ 500")
    if st.button(" Local ဖြင့် စာရင်းသွင်းမည်", type="primary", key="local_add"):
        if not customer.strip() or not text.strip(): st.warning("အမည်နှင့် input ဖြည့်ပါ။")
        else: display_result(save_rows(parse_local(text, customer.strip()), blocked, max_limit, enforce_limit))
with t2:
    customer2 = st.text_input("ဖောက်သည်အမည်", key="sms_customer")
    sms = st.text_area("SMS / Viber စာသား", height=150)
    if st.button(" Local စာသားဖတ်၍ ထည့်မည်", type="primary", key="sms_add"):
        if not customer2.strip() or not sms.strip(): st.warning("အမည်နှင့် စာသားဖြည့်ပါ။")
        else: display_result(save_rows(parse_local(sms, customer2.strip()), blocked, max_limit, enforce_limit))
with t3:
    st.info("ဒီနေရာမှာ cloud AI မသုံးပါ။ OCR engine ကို စက်ထဲမှာ သီးခြားတပ်ဆင်ထားမှ ပုံထဲကစာကို ဖတ်နိုင်ပါမည်။")
    img_customer = st.text_input("ပုံထဲမှာအမည်မပါလျှင် သုံးမည့်အမည်", value="Screenshot Customer")
    image_file = st.file_uploader("Screenshot", type=["png", "jpg", "jpeg"], key="offline_image")
    if image_file:
        image = ImageOps.exif_transpose(Image.open(image_file))
        st.image(image, use_container_width=True)
        if st.button(" Local OCR ဖြင့်ဖတ်မည်", key="ocr"):
            try:
                import pytesseract
                text_ocr = pytesseract.image_to_string(image, lang="eng")
                st.text_area("OCR ရလဒ် (စစ်ဆေးပြင်ဆင်နိုင်သည်)", text_ocr, height=120)
                display_result(save_rows(parse_local(text_ocr, img_customer), blocked, max_limit, enforce_limit))
            except ImportError:
                st.error("pytesseract မတပ်ရသေးပါ။ `pip install pytesseract` နှင့် Tesseract OCR ကို local စက်တွင် တပ်ပါ။")
            except Exception as exc: st.error(f"OCR အမှား: {exc}")

# Dashboard
df = load_df()
st.markdown("---")
if df.empty:
    st.info(" စာရင်းမရှိသေးပါ။")
else:
    total = int(df.Amount.sum())
    a, b, c = st.columns(3)
    a.metric(" စုစုပေါင်းရောင်းရငွေ", f"{total:,} MMK")
    b.metric(" အကွက်ရေ", f"{len(df):,}")
    c.metric(" ဖောက်သည်", f"{df.Customer.nunique():,}")

    st.subheader(" Customer Statement")
    selected = st.selectbox("ဖောက်သည်", sorted(df.Customer.unique()))
    selected_df = df[df.Customer == selected]
    statement = [f" {selected} ၏ 2D စာရင်း", "-" * 25] + [f" {r.Number} = {r.Amount:,} ကျပ်" for r in selected_df.itertuples()] + ["-" * 25, f"စုစုပေါင်း = {int(selected_df.Amount.sum()):,} ကျပ်"]
    st.text_area("Copy / Paste", "\n".join(statement), height=160)

    col1, col2 = st.columns(2)
    with col1:
        grouped = df.groupby("Number", as_index=False).Amount.sum().sort_values("Amount", ascending=False)
        grouped["Remaining"] = (max_limit - grouped.Amount).clip(lower=0)
        grouped["Status"] = grouped.Amount.apply(lambda x: "OVER" if x > max_limit else "OK")
        st.subheader(" Limit Control")
        st.dataframe(grouped, use_container_width=True, hide_index=True)
    with col2:
        st.subheader(" စာရင်းအားလုံး")
        st.dataframe(df, use_container_width=True, hide_index=True)

    win = number(winning)
    if winning:
        st.subheader(f" ပေါက်ဂဏန်းတိုက်စစ်ခြင်း ({win or 'မမှန်'})")
        if win:
            winners = df[df.Number == win]
            payout = int(winners.Amount.sum() * payout_rate)
            st.dataframe(winners, use_container_width=True, hide_index=True)
            st.metric("ဆုငွေ", f"{payout:,} MMK")
            st.metric("ဒိုင်အမြတ် / အရှုံး", f"{total-payout:,} MMK")

    st.download_button(" CSV Download", df.to_csv(index=False).encode("utf-8-sig"), "2d_offline_ledger.csv", "text/csv")
    if st.button(" Database အားလုံးဖျက်မည်", key="clear"):
        st.session_state.ask_clear = True
    if st.session_state.get("ask_clear"):
        st.warning("စာရင်းအားလုံးကို ဖျက်မည်မှာ သေချာပါသလား?")
        y, n = st.columns(2)
        if y.button("ဖျက်မည်", key="yes"):
            con = db(); con.execute("DELETE FROM bets"); con.commit(); con.close(); st.session_state.ask_clear = False; st.rerun()
        if n.button("မဖျက်တော့ပါ", key="no"):
            st.session_state.ask_clear = False; st.rerun()
