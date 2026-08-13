import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import zipfile
import xml.etree.ElementTree as ET

st.set_page_config(page_title="智能出貨單據與物流轉換工具", page_icon="📦", layout="centered")

st.title("📦 智能出貨單據與物流轉換系統 (3合1)")

# --- 1. 物流參數輸入 ---
st.subheader("⚙️ 1. 基本出貨物流參數")
col_a, col_b = st.columns(2)
with col_a:
    tracking_no = st.text_input("托運單號", value="SF1536155531299")
with col_b:
    doc_date = st.text_input("托運日期", value="2026.08.04")

# 多箱明細狀態初始化
if 'boxes' not in st.session_state:
    st.session_state.boxes = [{"dim": "42X34X17CM", "nw": 11.28, "gw": 12.28}]

st.markdown("#### 📦 多箱材積與重量明細")
col_btn1, col_btn2 = st.columns(2)
with col_btn1:
    if st.button("➕ 新增箱數"):
        st.session_state.boxes.append({"dim": "30X30X30CM", "nw": 10.0, "gw": 11.0})
        st.rerun()
with col_btn2:
    if len(st.session_state.boxes) > 1 and st.button("➖ 刪除最後一箱"):
        st.session_state.boxes.pop()
        st.rerun()

box_rows = ""
total_gw = 0

# 使用安全迴圈渲染每一箱輸入框
for i, b in enumerate(st.session_state.boxes):
    st.markdown(f"**── 第 {i+1} 箱 ──**")
    bc1, bc2, bc3 = st.columns(3)
    with bc1:
        st.session_state.boxes[i]['dim'] = st.text_input(f"尺寸 (Box {i+1})", value=b['dim'], key=f"dim_box_{i}")
    with bc2:
        st.session_state.boxes[i]['nw'] = st.number_input(f"淨重 KG (Box {i+1})", value=float(b['nw']), step=0.1, key=f"nw_box_{i}")
    with bc3:
        st.session_state.boxes[i]['gw'] = st.number_input(f"毛重 KG (Box {i+1})", value=float(b['gw']), step=0.1, key=f"gw_box_{i}")
    
    total_gw += st.session_state.boxes[i]['gw']
    box_rows += f"<tr><td>Box {i+1}</td><td>{st.session_state.boxes[i]['dim']}</td><td>{st.session_state.boxes[i]['nw']} KG</td><td>{st.session_state.boxes[i]['gw']} KG</td></tr>"

st.info(f"📊 **總計：** 共 {len(st.session_state.boxes)} 件 | 總毛重: **{total_gw:.2f} KG**")

# --- 2. 檔案上傳與預覽 ---
st.markdown("---")
st.subheader("📁 2. 上傳相關檔案")
so_file = st.file_uploader("上傳內部銷貨單 Excel", type=["xlsx", "xls"])

items_data = [{"品號": "D13750", "品名與規格": "美國壓縮彈簧 D13750", "數量": 50, "單價": 60.5, "金額": 3025}]
table_rows = "".join([f"<tr><td>{i+1}</td><td>{item['品號']}<br>{item['品名與規格']}</td><td>{item['數量']}</td><td>{item['單價']}</td><td>{item['金額']}</td></tr>" for i, item in enumerate(items_data)])

# --- 3. 3頁整合 HTML ---
html_code = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    body {{ background: #eee; font-family: sans-serif; }}
    .page {{ background: white; width: 750px; margin: 20px auto; padding: 40px; box-shadow: 0 0 10px rgba(0,0,0,0.1); page-break-after: always; }}
    .print-btn {{ position: fixed; top: 20px; right: 20px; padding: 15px 30px; background: #1a365d; color: white; border: none; border-radius: 5px; cursor: pointer; z-index: 1000; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
    td, th {{ border: 1px solid #000; padding: 8px; vertical-align: top; font-size: 10pt; }}
    th {{ background: #1a365d; color: white; }}
    .highlight {{ background-color: #ffe6cc; }}
    @media print {{ .print-btn {{ display: none; }} .page {{ box-shadow: none; margin: 0; }} }}
</style>
</head>
<body>
    <button class="print-btn" onclick="window.print()">🖨️ 列印 3 頁單據</button>

    <!-- P1: Invoice & Packing List -->
    <div class="page">
        <h2>INVOICE & PACKING LIST</h2>
        <table>
            <tr><td><b>Ship To:</b> 天津元象國際貿易有限公司</td><td><b>Date:</b> {doc_date}</td></tr>
            <tr><td><b>Tracking No:</b> {tracking_no}</td><td><b>Total G.W.:</b> {total_gw:.2f} KG</td></tr>
        </table>
        <table>
            <tr><th>箱號</th><th>尺寸</th><th>淨重</th><th>毛重</th></tr>
            {box_rows}
        </table>
        <table>
            <tr><th>項次</th><th>品名與規格</th><th>數量</th><th>單價</th><th>金額</th></tr>
            {table_rows}
        </table>
    </div>

    <!-- P2: 出口正式報單申請書 -->
    <div class="page">
        <h2 style="text-align:center;">出口正式報單申請書</h2>
        <table>
            <tr><td style="width: 50%;">托運單號：{tracking_no}</td><td class="highlight">托運日期： {doc_date}</td></tr>
            <tr><td>★統一編號：82850850</td><td>★取貨地址：桃園市蘆竹區安中街20巷13號4樓</td></tr>
            <tr><td>★出口公司：美加卓赫股份有限公司</td><td>★台灣連絡人：陳憲輝</td></tr>
            <tr><td>★緊急聯絡人員及手機：陳憲輝 0930-906-963</td><td>★公司電話：(02) 8201-4393#19</td></tr>
            <tr><td>★出口報單國外買方名稱：天津元象國際貿易有限公司</td><td>★送件目的地：中國天津市濱海新區新北路4668號</td></tr>
            <tr><td class="highlight">★托寄物中文品名：彈簧</td><td class="highlight">★箱數： {len(st.session_state.boxes)} 件</td></tr>
            <tr><td>★托寄物指定稅則：7320.90.00.00.0</td><td class="highlight">★重量(公斤)：{total_gw:.2f}KG</td></tr>
        </table>
    </div>

    <!-- P3: 個案委任書 -->
    <div class="page" style="height: 800px; display: flex; flex-direction: column; justify-content: space-between;">
        <div>
            <h2>個案委任書</h2>
            <p style="margin-top: 30px; line-height: 1.8;">茲委任<b>台灣順豐速運股份有限公司</b>辦理貴公司委託快遞出口貨物之報關、通關及相關查驗事宜。</p>
        </div>
        <div style="text-align: right; font-size: 14pt; padding-bottom: 50px;">
            中華民國 {int(doc_date.split('.')[0])-1911} 年 {doc_date.split('.')[1]} 月 {doc_date.split('.')[2]} 日
        </div>
    </div>
</body>
</html>
"""
components.html(html_code, height=1200)
