import streamlit as st
import pandas as pd
import os
import re

SPEND_LOOKUP_FILE = "table/unique_spend_code.csv"
AR_FILE = "table/ar_code.xlsx"

@st.cache_data
def load_spend_lookup():
    if os.path.exists(SPEND_LOOKUP_FILE):
        return pd.read_csv(SPEND_LOOKUP_FILE, dtype=str).fillna("")
    return pd.DataFrame(columns=["รหัสค่าใช้จ่าย", "หมวดรายจ่าย", "รายการ", "ประเภทค่าใช้จ่าย"])


def load_ar_data():
    if os.path.exists(AR_FILE):
        return pd.read_excel(AR_FILE, dtype=str).fillna("")
    return pd.DataFrame(columns=["รหัสโครงการวิจัย", "ar_code", "รหัสค่าใช้จ่าย"])

def save_ar_data(new_rows):
    old = load_ar_data()
    combined = pd.concat([old, new_rows], ignore_index=True)
    combined.to_excel(AR_FILE, index=False)


def reset_form():
    # รีเซตทุกฟิลด์ที่เราควบคุม
    st.session_state["ar_sets"] = [0]
    st.session_state.pop("project_code", None)
    st.session_state.pop("new_rows", None)
    st.session_state["reset_project_code"] = True
    for i in range(50):
        st.session_state.pop(f"ar_code_{i}", None)
        st.session_state.pop(f"spend_codes_{i}", None)
    st.session_state["just_reset"] = True



# --- UI ---
st.set_page_config(page_title="เพิ่ม AR Code หลายชุด", layout="wide")
st.title("📌 เพิ่ม AR Code สำหรับโครงการวิจัย")

# 🔔 จัดการ flag reset และแสดงข้อมูลล่าสุด
if st.session_state.get("just_reset", False):
    if "__tmp_new_rows__" in st.session_state:
        st.session_state["new_rows"] = st.session_state["__tmp_new_rows__"]
        del st.session_state["__tmp_new_rows__"]
    st.success("🔄 ฟอร์มถูกรีเซตเรียบร้อยแล้ว")
    del st.session_state["just_reset"]

if "saved_successfully" in st.session_state:
    st.success(f"✅ บันทึก AR Codes สำหรับโครงการ {st.session_state['saved_successfully']} แล้ว")
    del st.session_state["saved_successfully"]

spend_df = load_spend_lookup()

# --- รหัสโครงการหลัก ---
if "reset_project_code" in st.session_state and st.session_state["reset_project_code"]:
    default_code = ""
    del st.session_state["reset_project_code"]
else:
    default_code = st.session_state.get("project_code", "")


pattern_project_code = r"^E\d{4}_\d{3}$"
project_code = st.text_input("รหัสโครงการวิจัย (เช่น E2568_001)", value=default_code, key="project_code")
if not re.match(pattern_project_code, st.session_state.project_code):
    st.warning("รหัสโครงการวิจัยต้องอยู่ในรูปแบบ EXXXX_XXX (ตัวอย่าง: E2568_001)")


# --- รายการฟอร์ม AR code ---
if "ar_sets" not in st.session_state:
    st.session_state.ar_sets = [0]

for i in st.session_state.ar_sets:
    with st.container():
        st.subheader(f"🔁 AR Code ชุดที่ {i+1}")
        ar_code = st.text_input("🏷️ AR Code (รูปแบบ ARCXXX เช่น ARC001)", key=f"ar_code_{i}")
        selected_spend = st.multiselect(
            "🧾 เลือกรหัสค่าใช้จ่าย",
            spend_df["รหัสค่าใช้จ่าย"].unique().tolist(),
            key=f"spend_codes_{i}",
            default=[]  # ✅ เพิ่มบรรทัดนี้
        )

        if selected_spend:
            st.markdown("📄 รายละเอียดรหัสค่าใช้จ่าย")
            detail_rows = spend_df[spend_df["รหัสค่าใช้จ่าย"].isin(selected_spend)]
            st.dataframe(detail_rows, use_container_width=True)

        if i != 0:
            if st.button(f"➖ ลบ AR Code ชุดที่ {i+1}", key=f"remove_{i}"):
                st.session_state.ar_sets.remove(i)
                st.session_state.pop(f"ar_code_{i}", None)
                st.session_state.pop(f"spend_codes_{i}", None)
                st.rerun()

# --- เพิ่มชุดใหม่ ---
if st.button("➕ เพิ่ม AR Code ใหม่"):
    if st.session_state.ar_sets:
        new_idx = max(st.session_state.ar_sets) + 1
    else:
        new_idx = 1
    st.session_state.ar_sets.append(new_idx)
    st.rerun()

# --- บันทึกทั้งหมด ---
if st.button("💾 บันทึก AR Codes ทั้งหมด"):
    if not project_code.strip():
        st.warning("กรุณากรอกรหัสโครงการวิจัย")
    else:
        all_rows = []
        for i in st.session_state.ar_sets:
            ar_code = st.session_state.get(f"ar_code_{i}", "").strip()
            spend_codes = st.session_state.get(f"spend_codes_{i}", [])
            if ar_code and spend_codes:
                for code in spend_codes:
                    all_rows.append({
                        "รหัสโครงการวิจัย": project_code.strip(),
                        "ar_code": ar_code,
                        "รหัสค่าใช้จ่าย": code
                    })

        if all_rows:
            df = pd.DataFrame(all_rows)
            save_ar_data(df)

            # เตรียมแสดงข้อมูลล่าสุดในรอบถัดไป
            st.session_state["saved_successfully"] = project_code.strip()
            st.session_state["__tmp_new_rows__"] = all_rows
            reset_form()
            st.rerun()
            # st.cache_data.clear()
        else:
            st.warning("กรุณากรอก AR code และเลือกรหัสค่าใช้จ่ายอย่างน้อย 1 ชุด")

# --- แสดงข้อมูลที่เพิ่งเพิ่มล่าสุด ---
if "new_rows" in st.session_state:
    new_rows_df = pd.DataFrame(st.session_state["new_rows"])
    if not new_rows_df.empty:
        st.subheader("📋 ข้อมูลที่เพิ่งเพิ่มล่าสุด")
        st.dataframe(new_rows_df.astype(str), use_container_width=True)

for key in list(st.session_state.keys()):
    del st.session_state[key]
