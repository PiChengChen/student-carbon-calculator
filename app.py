import streamlit as st
from openpyxl import load_workbook
from datetime import datetime
import pandas as pd

# 載入 Excel 工作表
excel_path = '碳存摺.xlsx'
wb = load_workbook(excel_path)
ws = wb.active

# Streamlit 介面設定
st.set_page_config(page_title="碳排放計算器", layout="wide")
st.title("NCKU 環工系 碳排放計算器")

st.markdown("請輸入您今天的活動資料，系統將為您計算碳排放量（kg CO2）")

# --- 衣 ---
st.header("衣")
cloth1 = st.number_input("您今天買了幾件系服或社服？", min_value=0.0, step=1.0)
cloth2 = st.number_input("您今天買了幾件營隊或系隊隊服？", min_value=0.0, step=1.0)
sum_cloth = cloth1 * 4.3 + cloth2 * 5.5

# --- 住 ---
st.header("住")
bill_period = st.selectbox("您拿到幾月的帳單？", ['1-2', '3-4', '5-6', '7-8', '9-10', '11-12'])
living1 = st.number_input("水（度）", min_value=0.0, step=1.0)
living2 = st.number_input("電（度）", min_value=0.0, step=1.0)
living3 = st.number_input("瓦斯（單位）", min_value=0.0, step=1.0)
people = st.number_input("同居人數（包含自己）", min_value=1, step=1)
sum_living = (living1 * 0.156 + living2 * 0.494 + living3 * 2.63) / people

# --- 行 ---
st.header("行")
moving10 = st.number_input("汽機車共乘人數（包含自己）", min_value=1, step=1)
moving1 = st.number_input("電動機車（km）", min_value=0.0, step=1.0)
moving2 = st.number_input("燃油機車（km）", min_value=0.0, step=1.0)
moving3 = st.number_input("電動汽車（km）", min_value=0.0, step=1.0)
moving4 = st.number_input("燃油汽車（km）", min_value=0.0, step=1.0)
moving5 = st.number_input("油電混合（km）", min_value=0.0, step=1.0)
moving6 = st.number_input("高鐵（km）", min_value=0.0, step=1.0)
moving7 = st.number_input("火車（km）", min_value=0.0, step=1.0)
moving8 = st.number_input("捷運（km）", min_value=0.0, step=1.0)
moving9 = st.number_input("公車（km）", min_value=0.0, step=1.0)

sum_moving = (moving1 * 0.025 + moving2 * 0.06 + moving3 * 0.06 +
              moving4 * 0.24 + moving5 * 0.088) / moving10 + \
             (moving6 * 0.032 + moving7 * 0.06 + moving8 * 0.04 + moving9 * 0.04)

# --- 育 ---
st.header("育")
educ1 = st.number_input("47110 教室上課幾節？", min_value=0.0, step=1.0)
educ2 = st.number_input("47112 教室上課幾節？", min_value=0.0, step=1.0)
educ3 = st.number_input("47114 教室上課幾節？", min_value=0.0, step=1.0)
educ4 = st.number_input("47118 教室上課幾節？", min_value=0.0, step=1.0)
educ5 = st.number_input("47111 教室上課幾節？", min_value=0.0, step=1.0)
lab = st.selectbox("您今天在哪位老師的實驗室？", ['無', '吳義林', '朱信', '蔡俊鴻', '周佩欣', '林財富', '黃良銘',
                                                '陳婉如', '張智華', '林心恬', '陳必晟', '林聖倫',
                                                '劉守恆', '侯文哲', '黃榮振', '吳哲宏'])
educ6 = st.number_input("在該實驗室待了幾小時？", min_value=0.0, step=0.5)

lab_factors = {
    '吳義林': 0.5559, '朱信': 1.6325, '蔡俊鴻': 0.6126, '周佩欣': 1.0882,
    '林財富': 1.1179, '黃良銘': 0.6358, '陳婉如': 0.8461, '張智華': 0.2388,
    '林心恬': 0.1062, '陳必晟': 0.0728, '林聖倫': 1.0374, '劉守恆': 0.7006,
    '侯文哲': 1.9656, '黃榮振': 0.5044, '吳哲宏': 0.9688, '無': 0
}
sum_educ = educ1 * 0.035074 + educ2 * 0.030628 + educ3 * 0.042978 + \
           educ4 * 0.074594 + educ5 * 0.091884 + educ6 * lab_factors[lab]

# --- 結果顯示與儲存 ---
if st.button("計算碳排"):
    sum_total = sum_cloth + sum_living + sum_moving + sum_educ
    st.success(f"您的碳排放為：{sum_total:.2f} kg CO2")
    
    # 寫入 Excel
    year, month, day = datetime.now().year, datetime.now().month, datetime.now().day
    ws.append([year, month, day, sum_cloth, 0, sum_moving, sum_educ, sum_total])
    wb.save(excel_path)
    st.info("已儲存至碳存摺.xlsx")
