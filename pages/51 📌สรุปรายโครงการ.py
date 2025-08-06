import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

INCOME_FILE = "table/income_data.xlsx"
EXPEND_FILE = "table/expend_data.xlsx"

@st.cache_data
def load_data():
    income_df = pd.read_excel(INCOME_FILE, dtype=str)
    expend_df = pd.read_excel(EXPEND_FILE, dtype=str)
    return income_df, expend_df

st.set_page_config(page_title="📊 สรุปรายรับ-รายจ่ายรายโครงการ", layout="wide")
st.title("📊 สรุปรายรับ-รายจ่ายรายโครงการ")

# โหลดข้อมูล
income_df, expend_df = load_data()

# เตรียมข้อมูล
income_df["จำนวนเงิน"] = pd.to_numeric(income_df["จำนวนเงิน"], errors='coerce').fillna(0)
income_df["งวด"] = pd.to_numeric(income_df["งวด"], errors='coerce').fillna(0).astype(int)
expend_df["จำนวนเงิน"] = pd.to_numeric(expend_df["จำนวนเงิน"], errors='coerce').fillna(0)
expend_df["งวด"] = pd.to_numeric(expend_df["งวด"], errors='coerce').fillna(0).astype(int)

for df in [income_df, expend_df]:
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].str.strip()

# ตัวกรองรหัสโครงการ
project_codes = income_df["รหัสโครงการวิจัย"].dropna().unique()
selected_code = st.selectbox("📌 เลือกรหัสโครงการวิจัย:", sorted(project_codes))

# กรองข้อมูลตามรหัสโครงการ
income_proj = income_df[income_df["รหัสโครงการวิจัย"] == selected_code].copy()
expend_proj = expend_df[expend_df["รหัสโครงการวิจัย"] == selected_code].copy()

# กรองเฉพาะค่าใช้จ่ายจริง
expend_proj = expend_proj[expend_proj["ประเภทการจ่ายเงิน"] == "ค่าใช้จ่ายจริง"]

# รวมรายรับ-รายจ่าย
merged = pd.merge(
    income_proj[[
        "วันที่กรอกข้อมูล", "ประเภททุน", "ar_code", "รหัสค่าใช้จ่าย", "จำนวนเงิน", "งวด"
    ]].rename(columns={"จำนวนเงิน": "รายรับ"}),
    expend_proj[[
        "ประเภททุน", "ar_code", "รหัสค่าใช้จ่าย", "จำนวนเงิน", "งวด", "วันที่เบิกจ่าย"
    ]].rename(columns={"จำนวนเงิน": "รายจ่าย"}),
    on=["ประเภททุน", "ar_code", "รหัสค่าใช้จ่าย", "งวด"],
    how="outer"
)

merged["รายรับ"] = merged["รายรับ"].fillna(0)
merged["รายจ่าย"] = merged["รายจ่าย"].fillna(0)
merged["คงเหลือ"] = merged["รายรับ"] - merged["รายจ่าย"]

# คำนวณสรุปภาพรวม
total_income_all = merged["รายรับ"].sum()
total_expend_all = merged["รายจ่าย"].sum()
total_balance_all = total_income_all - total_expend_all

# สร้าง Pie Chart เฉพาะ รายจ่าย + คงเหลือ
pie_total = pd.DataFrame({
    "ประเภท": ["รายจ่าย", "คงเหลือ"],
    "จำนวนเงิน": [total_expend_all, total_balance_all]
})

fig_total = px.pie(
    pie_total,
    names="ประเภท",
    values="จำนวนเงิน",
    hole=0,
    color="ประเภท",
    color_discrete_map={
        "รายจ่าย": "#e98888",
        "คงเหลือ": "#9febc5"
    }
)
fig_total.update_traces(textposition="inside", textinfo="percent+label")

# แสดงรายละเอียดโครงการ + Pie Chart
proj_info = income_proj.iloc[0] if not income_proj.empty else None
if proj_info is not None:
    contract_code = proj_info.get("รหัสสัญญา", "-")
    contract_date_str = proj_info.get("วันที่เซนสัญญา", "")
    project_months = int(float(proj_info.get("ระยะเวลาดำเนินโครงการ (เดือน)", 0)))

    try:
        contract_date = pd.to_datetime(contract_date_str, errors="coerce")
        today = datetime.today()
        elapsed_days = (today - contract_date).days
        elapsed_months = elapsed_days // 30
        remain_days = max(project_months * 30 - elapsed_days, 0)
        remain_months = remain_days // 30
        remain_day_remain = remain_days % 30
    except:
        contract_date = None
        elapsed_months = remain_months = remain_day_remain = elapsed_days = 0

    col1, col2 = st.columns([1.2, 1])

    with col1:
        st.markdown(
            """
            <div style="height: 100%; display: flex; flex-direction: column; justify-content: center;">
            <h3>📁 รายละเอียดโครงการ</h3>
            <p>- <b>รหัสโครงการวิจัย:</b> {selected_code}</p>
            <p>- <b>รหัสสัญญา:</b> {contract_code}</p>
            <p>- <b>วันที่เซนสัญญา:</b> {contract_date}</p>
            <p>- <b>ระยะเวลาดำเนินการ:</b> {project_months} เดือน</p>
            <p>- <b>ดำเนินการแล้ว:</b> {elapsed_months} เดือน {elapsed_days_mod} วัน</p>
            <p>- <b>เหลืออีก:</b> {remain_months} เดือน {remain_day_remain} วัน</p>
            </div>
            """.format(
                selected_code=selected_code,
                contract_code=contract_code,
                contract_date=contract_date.strftime('%d/%m/%Y') if contract_date else '-',
                project_months=project_months,
                elapsed_months=elapsed_months,
                elapsed_days_mod=elapsed_days % 30,
                remain_months=remain_months,
                remain_day_remain=remain_day_remain
            ),
            unsafe_allow_html=True
        )

    with col2:
        st.plotly_chart(fig_total, use_container_width=True)

# แสดงข้อมูลรายงวด
st.markdown("### 📑 รายละเอียดรายรับ-รายจ่ายแต่ละงวด")

for period in sorted(merged["งวด"].unique()):
    period_data = merged[merged["งวด"] == period].copy()
    st.markdown(f"#### 🔸 งวดที่ {period}")

    st.dataframe(period_data[[
        "วันที่กรอกข้อมูล", "ประเภททุน", "ar_code", "รหัสค่าใช้จ่าย",
        "รายรับ", "วันที่เบิกจ่าย", "รายจ่าย", "คงเหลือ"
    ]].fillna("").style.format({
        "รายรับ": "{:,.2f}", "รายจ่าย": "{:,.2f}", "คงเหลือ": "{:,.2f}"
    }))

    total_income = period_data["รายรับ"].sum()
    total_expend = period_data["รายจ่าย"].sum()
    balance = total_income - total_expend

    overall_income = merged["รายรับ"].sum()
    percent_income = (total_income / overall_income) * 100 if overall_income > 0 else 0
    percent_expend = (total_expend / overall_income) * 100 if overall_income > 0 else 0
    percent_balance = (balance / overall_income) * 100 if overall_income > 0 else 0

    st.markdown(f"""
    🔹 **สรุปงวดที่ {period}**
    - รายรับรวม: {total_income:,.2f} บาท ({percent_income:.2f}% ของรายรับรวมทั้งหมด)
    - รายจ่ายรวม: {total_expend:,.2f} บาท ({percent_expend:.2f}% ของรายรับรวมทั้งหมด)
    - คงเหลือ: {balance:,.2f} บาท ({percent_balance:.2f}% ของรายรับรวมทั้งหมด)
    """)

# สรุปรวมทั้งหมด
st.markdown("### 🧾 สรุปภาพรวมทั้งหมด")
st.markdown(f"""
- รายรับรวมทั้งหมด: {total_income_all:,.2f} บาท (100%)
- รายจ่ายรวมทั้งหมด: {total_expend_all:,.2f} บาท ({(total_expend_all/total_income_all)*100:.2f}%)
- คงเหลือรวม: {total_balance_all:,.2f} บาท ({(total_balance_all/total_income_all)*100:.2f}%)
""")
