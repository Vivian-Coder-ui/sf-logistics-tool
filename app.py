import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import zipfile
import xml.etree.ElementTree as ET

st.set_page_config(page_title="智能出貨單據與物流轉換工具", page_icon="📦", layout="centered")

st.title("📦 智能出貨單據與物流轉換工具")
st.markdown("請在下方填寫物流參數並上傳內部銷貨單 Excel，系統將自動轉換並提供精美預覽與一鍵列印！")

# 1. 頂部物流參數輸入區
st.subheader("⚙️ 1. 輸入出貨物流參數")
col1, col2 = st.columns(2)
with col1:
    tracking_no = st.text_input("提單號碼 (Tracking No.)", value="SF157xxxxxxxx")
    gross_weight = st.text_input("毛重 (Gross Weight)", value="19.3 KG")
with col2:
    dimensions = st.text_input("材積尺寸 (Dimensions)", value="42X34X17CM")
    doc_date = st.text_input("單據日期", value="2026-08-12")

st.markdown("---")

# 2. 檔案上傳區
st.subheader("📁 2. 上傳內部銷貨單 Excel")
uploaded_file = st.file_uploader("請上傳銷貨單檔案 (.xlsx)", type=["xlsx", "xls"])

items_data = []
header_data = {}

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        fixed_io = io.BytesIO()
        
        # 自動修復 openpyxl 常見的 NamedCellStyle 錯誤
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as zin:
                with zipfile.ZipFile(fixed_io, 'w') as zout:
                    for item in zin.infolist():
                        buffer = zin.read(item.filename)
                        if item.filename == 'xl/styles.xml':
                            root = ET.fromstring(buffer)
                            for elem in root.iter():
                                if elem.tag.endswith('cellStyle') and ('name' not in elem.attrib or not elem.attrib['name']):
                                    elem.attrib['name'] = 'Normal'
                            buffer = ET.tostring(root)
                        zout.writestr(item, buffer)
            fixed_io.seek(0)
            excel_to_read = fixed_io
        except Exception:
            excel_to_read = io.BytesIO(file_bytes)

        # 讀取 ERP 銷貨單結構
        df_head = pd.read_excel(excel_to_read, sheet_name='單頭資料', header=2)
        df_body = pd.read_excel(excel_to_read, sheet_name='單身資料', header=2)
        
        header_data = df_head.iloc[0]
        items_df = df_body[['品號', '品名', '規格', '數量', '單價', '金額']].dropna(subset=['品號'])
        
        for _, row in items_df.iterrows():
            item_name = str(row.get('品名', '')).strip()
            spec = str(row.get('規格', '')).strip()
            full_desc = f"{item_name} {spec}".strip() if spec and spec != 'nan' else item_name
            items_data.append({
                "品號": str(row.get('品號', '')),
                "品名與規格": full_desc,
                "數量": int(row.get('數量', 0)),
                "單價": float(row.get('單價', 0)),
                "金額": float(row.get('金額', 0))
            })
            
        st.success(f"✅ 成功讀取銷貨單！共 {len(items_data)} 筆品項。")
        
    except Exception as e:
        st.error(f"❌ 讀取 Excel 發生錯誤：{e}")
else:
    # 預設範例資料（方便未上傳時預覽）
    header_data = {
        '銷貨單號': '20260722002',
        '客戶全名': '天津元象國際貿易有限公司',
        '送貨地址(一)': '中國天津市濱海新區新北路4668號濱海創新創業園4棟'
    }
    items_data = [
        {"品號": "D13750", "品名與規格": "美國壓縮彈簧 D13750", "數量": 50, "單價": 60.5, "金額": 3025},
        {"品號": "SB20210540", "品名與規格": "Belle Disc Springs For Spindle 二代碟簧", "數量": 100, "單價": 7.2, "金額": 720}
    ]

# 3. 組合 HTML 表格內容
table_rows_html = ""
grand_total = 0
for idx, item in enumerate(items_data):
    subtotal = item['數量'] * item['單價']
    grand_total += subtotal
    table_rows_html += f"""
    <tr>
        <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: center;">{idx+1}</td>
        <td style="padding: 8px; border: 1px solid #cbd5e1;"><strong>{item['品號']}</strong><br>{item['品名與規格']}</td>
        <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">{item['數量']:,}</td>
        <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">{item['單價']:,.2f}</td>
        <td style="padding: 8px; border: 1px solid #cbd5e1; text-align: right;">{item['金額']:,.2f}</td>
    </tr>
    """

# 4. 建立如同你之前採購單一樣精美的 HTML 預覽版面
html_code = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    body {{
        background: #f8fafc;
        color: #333;
        font-family: Arial, sans-serif;
        margin: 0;
        padding: 10px;
    }}
    .container {{
        max-width: 750px;
        margin: auto;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 30px;
        background: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }}
    .print-btn {{
        background-color: #1a365d;
        color: white;
        border: none;
        padding: 12px 24px;
        font-size: 14pt;
        font-weight: bold;
        border-radius: 6px;
        cursor: pointer;
        display: block;
        margin: 0 auto 25px auto;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }}
    .print-btn:hover {{ background-color: #2a4365; }}
    h2 {{ color: #1a365d; margin-bottom: 0px; text-align: center; }}
    .subtitle {{ color: #666; margin-top: 5px; font-size: 11pt; text-align: center; }}
    hr {{ border: 1px solid #1a365d; margin: 15px 0; }}
    .grid {{ width: 100%; margin-top: 10px; border-collapse: collapse; }}
    .box {{ background: #f8fafc; padding: 12px; border-radius: 5px; border: 1px solid #e2e8f0; font-size: 10pt; line-height: 1.5; }}
    table.items {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
    table.items th, table.items td {{ border: 1px solid #cbd5e1; padding: 8px; font-size: 10pt; }}
    table.items th {{ background-color: #1a365d; color: white; text-align: center; }}
    .text-right {{ text-align: right; }}

    @media print {{
        body {{ background: white; padding: 0; }}
        .container {{ border: none; box-shadow: none; padding: 0; max-width: 100%; }}
        .print-btn {{ display: none; }}
    }}
</style>
</head>
<body>
    <div class="container">
        <button class="print-btn" onclick="window.print()">🖨️ 點此列印 / 另存為 PDF 檔</button>

        <h2>INVOICE & PACKING LIST</h2>
        <div class="subtitle">商業發票與裝箱單</div>
        <hr>
        
        <table class="grid">
            <tr>
                <td class="box" style="width: 50%; vertical-align: top;">
                    <strong>【客戶資訊 (Ship to)】</strong><br>
                    {header_data.get('客戶全名', '')}<br>
                    地址：{header_data.get('送貨地址(一)', '')}
                </td>
                <td class="box" style="width: 50%; vertical-align: top;">
                    <strong>【出貨資訊】</strong><br>
                    Invoice No: {header_data.get('銷貨單號', '')}<br>
                    Date: {doc_date}<br>
                    Tracking No: {tracking_no}<br>
                    Gross Weight: {gross_weight}<br>
                    Dimensions: {dimensions}
                </td>
            </tr>
        </table>

        <table class="items">
            <thead>
                <tr>
                    <th style="width: 10%;">項次</th>
                    <th style="width: 45%;">品名與規格</th>
                    <th class="text-right" style="width: 15%;">數量</th>
                    <th class="text-right" style="width: 15%;">單價</th>
                    <th class="text-right" style="width: 15%;">金額</th>
                </tr>
            </thead>
            <tbody>
                {table_rows_html}
            </tbody>
        </table>

        <div style="text-align: right; font-size: 12pt; font-weight: bold; margin-top: 15px;">
            Total Amount: {grand_total:,.2f}
        </div>
    </div>
</body>
</html>
"""

st.markdown("---")
st.subheader("📋 3. 正式單據預覽與一鍵列印")
components.html(html_code, height=950, scrolling=True)
