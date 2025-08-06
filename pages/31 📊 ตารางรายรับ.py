import streamlit as st
import pandas as pd
import os

FILENAME = "table/income_data.xlsx"

@st.cache_data
def load_data():
    if os.path.exists(FILENAME):
        return pd.read_excel(FILENAME, dtype=str)
    else:
        return pd.DataFrame()

st.set_page_config(page_title="ตารางรายรับ", layout="wide")
st.title("📊 ตารางข้อมูลรายรับ")

df = load_data()

if df.empty:
    st.info("ยังไม่มีข้อมูลให้แสดง")
else:
    # ช่องค้นหา
    search_text = st.text_input("🔎 ค้นหารหัสหรือคำที่เกี่ยวข้อง", "")

    # กรองข้อมูล
    if search_text:
        filtered_df = df[df.apply(lambda row: row.astype(str).str.contains(search_text, case=False, na=False).any(), axis=1)]
    else:
        filtered_df = df

    st.markdown(f"📌 พบทั้งหมด {len(filtered_df):,} รายการที่ตรงกับการค้นหา")

    # ใช้ pandas Styler เพื่อ wrap text ใน header
    styled_df = filtered_df.style.set_table_styles(
        [{
            'selector': 'th',
            'props': [
                ('white-space', 'normal'),   # ให้ข้อความ header ขึ้นบรรทัดใหม่ได้
                ('max-width', '150px'),      # กำหนดความกว้างสูงสุดของ header cell ปรับได้ตามชอบ
                ('text-align', 'center'),
                ('vertical-align', 'middle'),
            ]
        }]
    )

    st.dataframe(styled_df, use_container_width=True)
