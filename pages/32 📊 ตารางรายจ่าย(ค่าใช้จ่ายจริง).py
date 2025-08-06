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
    
st.set_page_config(page_title="ตารางรายจ่าย(ค่าใช้จ่ายจริง)", layout="wide")
st.title("📊 ตารางข้อมูลรายจ่าย(ค่าใช้จ่ายจริง)")

df = load_data()

if df.empty:
    st.info("ยังไม่มีข้อมูลให้แสดง")
else:
    # ลบช่องว่างหัวท้ายของชื่อคอลัมน์
    df.columns = df.columns.str.strip()

    # กรองเฉพาะแถวที่มี "ประเภทการจ่ายเงิน" = "ค่าใช้จ่ายจริง"
    if "ประเภทการจ่ายเงิน" in df.columns:
        df["ประเภทการจ่ายเงิน"] = df["ประเภทการจ่ายเงิน"].str.strip()
        filtered_df = df[df["ประเภทการจ่ายเงิน"] == "ค่าใช้จ่ายจริง"]
    else:
        st.warning("ไม่พบคอลัมน์ 'ประเภทการจ่ายเงิน' ในข้อมูล")
        filtered_df = df

    # ช่องค้นหา
    search_text = st.text_input("🔎 ค้นหารหัสหรือคำที่เกี่ยวข้อง", "")

    # กรองด้วยข้อความค้นหา (ถ้ามี)
    if search_text:
        filtered_df = filtered_df[
            filtered_df.apply(lambda row: row.astype(str).str.contains(search_text, case=False, na=False).any(), axis=1)
        ]

    st.markdown(f"📌 พบทั้งหมด {len(filtered_df):,} รายการที่ตรงกับการค้นหา")
    st.dataframe(filtered_df, use_container_width=True)
