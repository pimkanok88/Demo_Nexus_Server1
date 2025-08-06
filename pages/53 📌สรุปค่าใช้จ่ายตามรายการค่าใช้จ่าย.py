import streamlit as st
import pandas as pd
import io

INCOME_FILE = "table/income_data.xlsx"
EXPEND_FILE = "table/expend_data.xlsx"

st.set_page_config(page_title="ตารางสรุปงบประมาณ 2 ระดับ", layout="wide")

@st.cache_data
def load_data():
    income_df = pd.read_excel(INCOME_FILE)
    expend_df = pd.read_excel(EXPEND_FILE)
    return income_df, expend_df

income_df, expend_df = load_data()

# รวมชื่อคอลัมน์ให้ตรงกับไฟล์จริงถ้าจำเป็น
income_df = income_df.rename(columns={'รายการงบ': 'รายการ'})
expend_df = expend_df.rename(columns={'รายการรายจ่าย': 'รายการ'})

# รวมยอดจัดสรรและเบิกจ่ายแยกตามงวดและรายการ
income_grouped = income_df.groupby(['งวด', 'รายการ']).agg({'จำนวนเงิน': 'sum'}).reset_index()
income_grouped['ประเภท'] = 'จัดสรร'

expend_grouped = expend_df.groupby(['งวด', 'รายการ']).agg({'จำนวนเงิน': 'sum'}).reset_index()
expend_grouped['ประเภท'] = 'เบิกจ่าย'

combined = pd.concat([income_grouped, expend_grouped], ignore_index=True)
combined = combined.rename(columns={'จำนวนเงิน': 'ยอดรวม'})

# สร้าง Pivot table: index = ประเภท, columns = (งวด, รายการ)
pivot = combined.pivot_table(
    index='ประเภท',
    columns=['งวด', 'รายการ'],
    values='ยอดรวม',
    aggfunc='sum',
    fill_value=0
)

pivot = pivot.T.groupby(level=[0,1]).sum().T

# เพิ่มแถว "คงเหลือ" = จัดสรร - เบิกจ่าย
allocated = pivot.loc['จัดสรร']
spent = pivot.loc['เบิกจ่าย']
remain = allocated - spent
remain.name = 'คงเหลือ'
pivot = pd.concat([pivot, remain.to_frame().T])

# สร้างคอลัมน์ "รวม" ต่อแต่ละงวด (sum ตามหมวด)
sum_dfs = []
for period in pivot.columns.get_level_values(0).unique():
    df_slice = pivot.loc[:, period]
    sum_col = df_slice.sum(axis=1)
    sum_df = pd.DataFrame(sum_col)
    sum_df.columns = pd.MultiIndex.from_tuples([(period, 'รวม')])
    sum_dfs.append(sum_df)

pivot_with_sum = pd.concat([pivot] + sum_dfs, axis=1)
pivot_with_sum = pivot_with_sum.sort_index(axis=1, level=[0,1])

pivot_with_sum.columns.names = ['งวด', 'หมวด']

# ตั้งชื่อ columns level 0 เป็น "งวดที่ x"
cols = pivot_with_sum.columns.to_frame(index=False)
cols['งวด'] = cols['งวด'].apply(lambda x: f"งวดที่ {x}")
pivot_with_sum.columns = pd.MultiIndex.from_frame(cols)

# ฟังก์ชันใส่สีพื้นหลังคอลัมน์สลับตามงวด
def highlight_cols_auto(col):
    colors = ['#f0f8ff', '#faebd7', '#e6e6fa']
    period = col.name[0]
    try:
        num = int(str(period).split()[-1])
    except:
        num = 0
    color = colors[(num - 1) % len(colors)]
    return [f'background-color: {color}'] * len(col)

styled_df = pivot_with_sum.style.apply(highlight_cols_auto, axis=0)

header_style = [
    {'selector': 'th.col_heading.level0', 'props': [('background-color', '#a2d5f2'), ('color', 'black'), ('text-align', 'center'), ('font-weight', 'bold')]},
    {'selector': 'th.col_heading.level1', 'props': [('background-color', '#d3e0ea'), ('color', 'black'), ('text-align', 'center'), ('font-weight', 'bold')]},
    {'selector': 'th.row_heading', 'props': [('background-color', '#f7cac9'), ('color', 'black'), ('text-align', 'center'), ('font-weight', 'bold')]}
]

styled_df = styled_df.set_table_styles(header_style)

st.subheader("📊 ตารางสรุปงบประมาณ 2 ระดับ")

st.dataframe(
    styled_df.format("{:,.0f}"),
    use_container_width=True
)


# สร้าง Excel ไฟล์ในหน่วยความจำ
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    pivot_with_sum.to_excel(writer, sheet_name='สรุปงบประมาณ')
output.seek(0)

st.download_button(
    label="⬇️ ดาวน์โหลด Excel (MultiIndex columns)",
    data=output,
    file_name="สรุปงบประมาณ_2ระดับ.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
