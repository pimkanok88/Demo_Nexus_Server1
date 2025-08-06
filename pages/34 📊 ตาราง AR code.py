import streamlit as st
import pandas as pd
import os

AR_FILE = "table/ar_code.xlsx"
SPEND_FILE = "table/unique_spend_code.csv"

@st.cache_data
def load_ar_data():
    if os.path.exists(AR_FILE):
        return pd.read_excel(AR_FILE, dtype=str)
    return pd.DataFrame()

@st.cache_data
def load_spend_data():
    if os.path.exists(SPEND_FILE):
        return pd.read_csv(SPEND_FILE, dtype=str).fillna("")
    return pd.DataFrame(columns=["รหัสค่าใช้จ่าย", "หมวดรายจ่าย", "รายการ", "ประเภทค่าใช้จ่าย"])

st.set_page_config(page_title="ตาราง AR code", layout="wide")
st.title("📊 ตาราง AR code พร้อมรายละเอียด")

ar_df = load_ar_data()
spend_df = load_spend_data()

if ar_df.empty:
    st.info("ยังไม่มีข้อมูลให้แสดง")
else:
    # ตัวกรองรหัสโครงการวิจัย
    project_codes = sorted(ar_df["รหัสโครงการวิจัย"].dropna().unique())
    selected_code = st.selectbox("🔍 เลือกรหัสโครงการวิจัย", options=["(ทั้งหมด)"] + project_codes)

    if selected_code != "(ทั้งหมด)":
        filtered_df = ar_df[ar_df["รหัสโครงการวิจัย"] == selected_code]
    else:
        filtered_df = ar_df.copy()

    # เชื่อมรายละเอียดรหัสค่าใช้จ่าย
    merged_df = filtered_df.merge(spend_df, how="left", on="รหัสค่าใช้จ่าย")

    # สรุปตามรหัสโครงการวิจัย และแยก ar_code พร้อมรหัสค่าใช้จ่าย
    grouped = merged_df.groupby(["รหัสโครงการวิจัย", "ar_code"]).agg({
        "รหัสค่าใช้จ่าย": lambda x: ", ".join(sorted(x.dropna().unique()))
    }).reset_index()



    current_project = None
    for _, row in grouped.iterrows():
        project = row["รหัสโครงการวิจัย"]
        if project != current_project:
            st.markdown(f"### รหัสโครงการวิจัย: `{project}`")
            current_project = project
        st.markdown(f"- AR Code `{row['ar_code']}`: รหัสค่าใช้จ่าย: {row['รหัสค่าใช้จ่าย']}")

    st.markdown("---")
    st.markdown("### 📋 รายการรายละเอียดทั้งหมด")
    st.dataframe(merged_df, use_container_width=True)
