import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="กราฟทุน", layout="wide")

TOTAL_BUDGET = 5_000_000
DATA_FILE = "table/income_data.xlsx"

@st.cache_data
def load_data():
    return pd.read_excel(DATA_FILE)

df = load_data()
st.set_page_config(page_title="สรุปการใช้ทุน)", layout="wide")
st.title("📋 📊 สรุปการใช้ทุน (เฉพาะทุนภายใน)")


if df.empty:
    st.error("❌ ไม่มีข้อมูลในไฟล์ income_data.xlsx กรุณากรอกข้อมูลก่อน")
else:
    required_cols = ['รหัสโครงการวิจัย', 'หมวดรายจ่าย', 'จำนวนเงิน', 'ประเภททุน']
    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"❌ ไม่พบคอลัมน์: {', '.join(missing_cols)} กรุณาตรวจสอบไฟล์")
    else:
        internal_df = df[df['ประเภททุน'] == 'ทุนภายใน'].copy()

        if internal_df.empty:
            st.warning("⚠️ ไม่มีข้อมูลสำหรับประเภท 'ทุนภายใน'")
        else:
            summary = internal_df.groupby('รหัสโครงการวิจัย')['จำนวนเงิน'].sum().reset_index()
            summary['สัดส่วนจากเงินทุนทั้งหมด'] = (summary['จำนวนเงิน'] / TOTAL_BUDGET) * 100

            formatted_summary = summary.copy()
            formatted_summary['จำนวนเงิน'] = formatted_summary['จำนวนเงิน'].apply(lambda x: f"{x:,.0f}")
            formatted_summary['สัดส่วนจากเงินทุนทั้งหมด'] = formatted_summary['สัดส่วนจากเงินทุนทั้งหมด'].apply(lambda x: f"{x:.2f} %")
            

            pie_data = summary[['รหัสโครงการวิจัย', 'จำนวนเงิน']].copy()
            total_used = pie_data['จำนวนเงิน'].sum()
            remaining = TOTAL_BUDGET - total_used

            # เพิ่มแถวคงเหลือ
            pie_data = pd.concat([
                pie_data,
                pd.DataFrame([{
                    'รหัสโครงการวิจัย': 'คงเหลือ',
                    'จำนวนเงิน': remaining
                }])
            ], ignore_index=True)
            # ✅ เรียงชื่อโครงการ A-Z แล้วตามด้วย 'คงเหลือ'
            project_names = sorted(pie_data[pie_data['รหัสโครงการวิจัย'] != 'คงเหลือ']['รหัสโครงการวิจัย'].unique().tolist())
            project_names.append('คงเหลือ')
            pie_data['รหัสโครงการวิจัย'] = pd.Categorical(
                pie_data['รหัสโครงการวิจัย'],
                categories=project_names,
                ordered=True
            )
            fig = px.pie(
                pie_data,
                names='รหัสโครงการวิจัย',
                values='จำนวนเงิน',
                hole=0,
                color='รหัสโครงการวิจัย',
                color_discrete_map={'คงเหลือ': '#d3d3d3'  # บังคับให้สี "คงเหลือ" เป็นสีเทา
                                    },
                color_discrete_sequence=px.colors.qualitative.Vivid, ## Plotly, Pastel, Dark24, Vivid, Bold, Prism, Safe
                category_orders={
                    'รหัสโครงการวิจัย': project_names  # ✅ บังคับลำดับแสดงในกราฟ
                    }
            )

            fig.update_layout(
                height=600,  # ความสูง
                font=dict(family="Tahoma", size=20, color="black"),
                legend_title_text="",  # ❌ ซ่อนคำว่า "รหัสโครงการวิจัย"
                legend_font_size=20,
                legend=dict(
                    orientation="v",   # แสดงแนวนอน (horizontal) / v
                    yanchor="top",  # 'auto', 'top', 'middle', 'bottom'
                    y=1,            # เลื่อน legend ลงด้านล่างใต้กราฟ
                    xanchor="left",  # left right center
                    x=0            # ให้อยู่ตรงกลาง
            ))
            fig.update_traces(
                hoverlabel=dict(
                    font_size=16,
                    font_family="Tahoma"
                )
            )
            # st.dataframe(formatted_summary)
            # st.plotly_chart(fig, use_container_width=True)

            col1, col2 = st.columns(2)
            col1, col2 = st.columns([2, 1])

        with col1:
            st.subheader("📊 กราฟแสดงสัดส่วนการใช้เงินทุนวิจัย")
            st.plotly_chart(fig, use_container_width=False)

        with col2:
            st.subheader("📊 รายละเอียดการใช้ทุน")
            st.dataframe(formatted_summary, use_container_width=True)

        st.markdown(f"### 💰 ใช้ไป (ทุนภายใน): `{total_used:,.0f}` บาท / คงเหลือ: `{remaining:,.0f}` บาท")
