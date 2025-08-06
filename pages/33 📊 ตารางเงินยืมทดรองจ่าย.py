import streamlit as st
import pandas as pd
import os

FILENAME = "table/expend_data.xlsx"

@st.cache_data
def load_data():
    if os.path.exists(FILENAME):
        return pd.read_excel(FILENAME, sheet_name="Sheet1", dtype=str)
    else:
        return pd.DataFrame()
    
st.set_page_config(page_title="ตารางรายจ่าย(เงินยืมทดรองจ่าย)", layout="wide")
st.title("📊 ตารางข้อมูลรายจ่าย (เงินยืมทดรองจ่าย)")

df = load_data()

if df.empty:
    st.info("ยังไม่มีข้อมูลให้แสดง")
else:
    df.columns = df.columns.str.strip()

    if "ประเภทการจ่ายเงิน" in df.columns and "จำนวนเงิน" in df.columns:
        df["ประเภทการจ่ายเงิน"] = df["ประเภทการจ่ายเงิน"].str.strip()
        df["จำนวนเงิน"] = pd.to_numeric(df["จำนวนเงิน"], errors="coerce").fillna(0)

        # ✅ เงื่อนไขหลัก: เงินยืมทดรองจ่าย และ จำนวนเงิน > 0
        filtered_df = df[
            (df["ประเภทการจ่ายเงิน"] == "เงินยืมทดรองจ่าย") &
            (df["จำนวนเงิน"] > 0)
        ]

        # 🔍 ช่องค้นหาเพิ่มเติม
        search_text = st.text_input("🔎 ค้นหา (เช่น รหัสโครงการ, รหัสค่าใช้จ่าย, รายการ ฯลฯ):")

        if search_text:
            filtered_df = filtered_df[filtered_df.apply(
                lambda row: row.astype(str).str.contains(search_text, case=False, na=False).any(),
                axis=1
            )]

        st.markdown(f"📌 พบทั้งหมด {len(filtered_df):,} รายการที่ตรงกับเงื่อนไข")
        st.dataframe(filtered_df, use_container_width=True)
    else:
        st.warning("ไม่พบคอลัมน์ 'ประเภทการจ่ายเงิน' หรือ 'จำนวนเงิน' ในข้อมูล")
