# ตาราง
import streamlit as st
import pandas as pd
import os

FILENAME = "table/funding_source.xlsx"


@st.cache_data
def load_data():
    if os.path.exists(FILENAME):
        return pd.read_excel(FILENAME, dtype=str)
    return pd.DataFrame()

st.set_page_config(page_title="ตารางรหัสงบประมาณ", layout="wide")
st.title("📊 ตารางรหัสงบประมาณ")

df = load_data()

if df.empty:
    st.info("ยังไม่มีข้อมูลให้แสดง")
else:
    # 🔍 ช่องค้นหา
    search_text = st.text_input("🔎 ค้นหารหัสงบประมาณหรือคำที่เกี่ยวข้อง", "")

    # กรองข้อมูล
    if search_text:
        filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(search_text, case=False, na=False).any(), axis=1)]
    else:
        filtered_df = df

    st.markdown(f"📌 พบทั้งหมด {len(filtered_df):,} รายการที่ตรงกับการค้นหา")
    st.dataframe(filtered_df, use_container_width=True)
