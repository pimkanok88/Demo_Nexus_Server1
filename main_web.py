import streamlit as st
import pandas as pd
from datetime import datetime
import os
from openpyxl import load_workbook

st.set_page_config(page_title="ระบบบันทึกค่าใช้จ่าย", layout="wide")
st.title("DEMO 📊 ระบบบันทึกรายรับ-รายจ่ายโครงการวิจัย")

# ใช้ CSS ปรับขนาด container
st.markdown("""
    <style>
        .main .block-container {
            padding-top: 1rem;
            padding-right: 2rem;
            padding-left: 2rem;
            padding-bottom: 1rem;
            max-width: 100% !important;
        }

        .stDataFrame > div {
            max-width: 100% !important;
        }
    </style>
""", unsafe_allow_html=True)