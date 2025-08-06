import streamlit as st
import pandas as pd
from datetime import datetime
import uuid

today = pd.to_datetime(datetime.today().date())
RESERVE_FILE = "table/reserve_payment.xlsx"

@st.cache_data(ttl=1)
def load_reserve_data():
    try:
        df = pd.read_excel(RESERVE_FILE, dtype=str).fillna("")
        df.columns = df.columns.str.strip()

        df["จำนวนเงิน"] = pd.to_numeric(df["จำนวนเงิน"].str.replace(",", ""), errors="coerce")
        df["เงินที่คืน"] = pd.to_numeric(df.get("เงินที่คืน", 0), errors="coerce").fillna(0.0)
        df["วันที่ยืม"] = pd.to_datetime(df["วันที่ยืม"], errors='coerce')
        df["วันที่ต้องคืน"] = pd.to_datetime(df["วันที่ต้องคืน"], errors='coerce')
        df["วันที่คืนเงิน"] = pd.to_datetime(df.get("วันที่คืนเงิน", None), errors='coerce')

        return df
    except Exception as e:
        st.error(f"ไม่สามารถโหลดข้อมูลได้: {e}")
        return pd.DataFrame()
    
def add_remaining_column(df):
    df = df.copy()
    df["จำนวนเงิน"] = pd.to_numeric(df["จำนวนเงิน"], errors="coerce")
    df["เงินที่คืน"] = pd.to_numeric(df["เงินที่คืน"], errors="coerce").fillna(0.0)
    df["คงเหลือ"] = df["จำนวนเงิน"] - df["เงินที่คืน"]
    df["คงเหลือ"] = df["คงเหลือ"].apply(lambda x: max(x, 0))
    return df

def save_reserve_data(df):
    try:
        df = add_remaining_column(df)
        df.to_excel(RESERVE_FILE, index=False)
        st.success("✅ บันทึกข้อมูลเรียบร้อยแล้ว")
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการบันทึกข้อมูล: {e}")


def process_summary(df):
    df = df.copy()
    # แปลงเป็นตัวเลข
    df["จำนวนเงิน"] = pd.to_numeric(df["จำนวนเงิน"], errors="coerce").fillna(0)
    df["เงินที่คืน"] = pd.to_numeric(df["เงินที่คืน"], errors="coerce").fillna(0)
    
    # รวมเงินที่คืนทั้งหมด ต่อกลุ่ม key ที่เหมือนกัน
    group_cols = ['รหัสโครงการวิจัย', 'ar_code', 'รหัสค่าใช้จ่าย', 'วันที่ยืม', 'จำนวนเงิน', 'วันที่ต้องคืน']
    
    refund_sum = df.groupby(group_cols)["เงินที่คืน"].sum().reset_index()
    
    # คำนวณคงเหลือ = จำนวนเงิน - รวมเงินที่คืน
    refund_sum["คงเหลือ"] = refund_sum["จำนวนเงิน"] - refund_sum["เงินที่คืน"]
    refund_sum["คงเหลือ"] = refund_sum["คงเหลือ"].apply(lambda x: max(x, 0))
    
    # กำหนดสถานะ
    today = pd.to_datetime(datetime.today().date())
    refund_sum["วันที่ต้องคืน"] = pd.to_datetime(refund_sum["วันที่ต้องคืน"], errors='coerce')

    def get_status(row):
        remain = row["คงเหลือ"]
        due_date = row["วันที่ต้องคืน"]
        if remain != 0 and pd.notna(due_date) and due_date > today:
            return "ยังไม่คืน"
        elif remain != 0 and pd.notna(due_date) and due_date < today:
            return "เลยกำหนด"
        elif remain == 0:
            return "ปิดบัญชี"
        else:
            return "ไม่ทราบสถานะ"

    refund_sum["สถานะ"] = refund_sum.apply(get_status, axis=1)
    
    return refund_sum


def reset_form():
    tmp_rows = st.session_state.get("__tmp_new_rows__", None)
    keys_to_keep = {"__tmp_new_rows__", "refund_key"}
    keys_to_delete = [k for k in st.session_state.keys() if k not in keys_to_keep]

    for k in keys_to_delete:
        del st.session_state[k]

    st.session_state["just_reset"] = True

def highlight_status(row):
    status = row["สถานะ"]
    color = ""
    if status == "ยังไม่คืน":
        color = "background-color: #fff3cd"
    elif status == "เลยกำหนด":
        color = "background-color: #f8d7da"
    elif status == "ปิดบัญชี":
        color = "background-color: #d4edda"
    return [""] * (len(row) - 1) + [color]

# ====== ฟังก์ชันกรองแถวคืนเงินล่าสุด ======
def filter_latest_return(df):
    df = df.copy()
    df["วันที่คืนเงิน_filled"] = df["วันที่คืนเงิน"].fillna(pd.Timestamp("1900-01-01"))

    idx = df.groupby(
        ['รหัสโครงการวิจัย', 'ar_code', 'รหัสค่าใช้จ่าย', 'วันที่ยืม']
    )["วันที่คืนเงิน_filled"].idxmax()

    return df.loc[idx].drop(columns=["วันที่คืนเงิน_filled"])


# ====== UI =======
st.set_page_config(page_title="อัปเดตการคืนเงิน", layout="wide")
st.title("📌สรุปเงินยืมทดรองจ่าย")

reserve_df = load_reserve_data()

if reserve_df.empty:
    st.warning("ไม่พบข้อมูลต้นทาง")
    st.stop()

reserve_df = reserve_df[reserve_df["จำนวนเงิน"] > 0]

project_codes = sorted(reserve_df["รหัสโครงการวิจัย"].dropna().unique())
selected_project = st.selectbox("เลือกรหัสโครงการวิจัย", project_codes)

filtered_df = reserve_df[reserve_df["รหัสโครงการวิจัย"] == selected_project].reset_index(drop=True)

sum_df = process_summary(filtered_df)
filtered_latest_df = sum_df

st.markdown(f"### 📊 รายการเงินยืมทดรองจ่ายของโครงการวิจัย: `{selected_project}`")

col_refresh, _ = st.columns([1, 5])
with col_refresh:
    if st.button("🔄 รีเฟรชตารางข้อมูล"):
        st.rerun()

styled_df = filtered_latest_df[[
    "รหัสโครงการวิจัย", "ar_code", "รหัสค่าใช้จ่าย", "วันที่ยืม", "จำนวนเงิน",
    "วันที่ต้องคืน", "เงินที่คืน", "คงเหลือ", "สถานะ"
]].style\
  .apply(highlight_status, axis=1)\
  .format({"จำนวนเงิน": "{:,.2f}", "เงินที่คืน": "{:,.2f}", "คงเหลือ": "{:,.2f}"})

st.dataframe(styled_df, use_container_width=True)

st.markdown("---")
st.markdown("### ➕ เพิ่มข้อมูลการคืนเงิน")

df_not_zero = filtered_latest_df[filtered_latest_df["คงเหลือ"] != 0]

has_ar_code = filtered_df["ar_code"].replace("", pd.NA).dropna().nunique() > 0

if has_ar_code:
    st.markdown("#### เลือกรหัส ar_code")
    ar_codes = df_not_zero["ar_code"].dropna().unique()
    selected_ar = st.selectbox("เลือกรหัส ar_code", ar_codes)

    sub_df = df_not_zero[df_not_zero["ar_code"] == selected_ar]
    spend_codes = sub_df["รหัสค่าใช้จ่าย"].dropna().unique()
    selected_spend = st.selectbox("เลือกรหัสค่าใช้จ่าย", spend_codes)
    rows = sub_df[sub_df["รหัสค่าใช้จ่าย"] == selected_spend]
else:
    st.markdown("#### ไม่พบ ar_code → เลือกรหัสค่าใช้จ่ายแทน")
    spend_codes = df_not_zero["รหัสค่าใช้จ่าย"].dropna().unique()
    selected_spend = st.selectbox("เลือกรหัสค่าใช้จ่าย", spend_codes)
    rows = df_not_zero[df_not_zero["รหัสค่าใช้จ่าย"] == selected_spend]

if rows.empty:
    st.warning("ไม่พบข้อมูลรายการที่เกี่ยวข้อง")
else:
    row_idx = rows.index[0]

    # ✅ เตรียม key สำหรับ reset ช่องกรอกจำนวนเงิน
    if "refund_key" not in st.session_state:
        st.session_state["refund_key"] = str(uuid.uuid4())

    with st.form("refund_form"):
        return_date = st.date_input("📆 วันที่คืนเงิน", datetime.today())

        return_amt = st.number_input(
            "💰 จำนวนที่คืน", 
            min_value=0.0, 
            step=100.0, 
            format="%.2f", 
            key=st.session_state["refund_key"]
        )
        submitted = st.form_submit_button("📂 บันทึก")

        if submitted:
            selected_row = sum_df.loc[row_idx]
            actual_idx = filtered_df.index[row_idx]

            # ข้อมูลแถวเก่า
            old_row = reserve_df.loc[actual_idx].copy()

            old_refund = pd.to_numeric(old_row["เงินที่คืน"], errors="coerce")
            if pd.isna(old_refund):
                old_refund = 0.0

            new_refund = old_refund + return_amt
            new_return_date = pd.to_datetime(return_date)

            # สร้างแถวใหม่ แทนการแก้ไขแถวเก่า
            new_row = old_row.copy()
            new_row["เงินที่คืน"] = new_refund
            new_row["วันที่คืนเงิน"] = new_return_date

            # เพิ่มแถวใหม่ลง DataFrame
            reserve_df = pd.concat([reserve_df, pd.DataFrame([new_row])], ignore_index=True)

            save_reserve_data(reserve_df)

            # เปลี่ยน key ใหม่เพื่อ reset ช่องกรอกจำนวนเงิน
            st.session_state["refund_key"] = str(uuid.uuid4())

            st.success("✅ เพิ่มข้อมูลเรียบร้อยแล้ว")
            reset_form()
            st.rerun()
