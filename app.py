import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import zipfile
import xml.etree.ElementTree as ET

st.set_page_config(page_title="智能出貨單據與物流轉換工具", page_icon="📦", layout="centered")

st.title("📦 智能出貨單據與物流轉換工具")
st.markdown("請在下方填寫多箱物流明細並上傳內部銷貨單 Excel，系統將自動計算總重並提供精美預覽與一鍵列印！")

# 1. 頂部基本資訊
st.subheader("⚙️ 1. 基本出貨參數")
col_a, col_b = st.columns(2)
with col_a:
    tracking_no = st.text_input("主提單號碼 (Master Tracking No.)", value="SF157xxxxxxxx")
with col_b:
    doc_date = st.text_input("單據日期", value="2026-08-12")

currency = "RMB (人民幣)"

# 2. 多箱明細動態輸入區
st.subheader("📦 2. 多箱材積與重量明細 (Multi-Box Details)")
st.markdown("如果本次出貨有多個箱子，請在下方新增並填寫每一箱的資料：")

# 使用 session_state 來管理多箱資料
if 'boxes' not in st.session_state:
    st.session_state.boxes = [
        {"box_no": 1, "dimensions": "42X34X17CM", "net_w": 18.3, "gross_w": 19.3}
    ]

# 新增與刪除按鈕
col_btn1, col_btn2 = st.columns(2)
with col_btn1:
    if st.button("➕ 新增一箱"):
        next_no = len(st.session_state.boxes) + 1
        st.session_state.boxes.append({"box_no": next_no, "dimensions": "30X30X30CM", "net_w": 10.0, "gross_w": 11.0})
        st.rerun()
with col_btn2:
    if len(st.session_state.boxes) > 1 and st.button("➖ 刪除最後一箱"):
        st.session_state.boxes.pop()
        st.rerun()

# 渲染每一箱的輸入欄位
box_details_html = ""
total_net_weight = 0.0
total_gross_weight = 0.0

for i, box in enumerate(st.session_state.boxes):
    st.markdown(f"**── 第 {i+1} 箱 ──**")
    b_col1, b_col2, b_col3 = st.columns(3)
    with b_col1:
        box['dimensions'] = st.text_input(f"材積尺寸 (Box {i+1})", value=box['dimensions'], key=f"dim_{i}")
    with b_col2:
        box['net_w'] = st.number_input(f"淨重 KG (Box {i+1})", value=float(box['net_w']), step=0.1, key=f"nw_{i}")
    with b_col3:
        box['gross_w'] = st.number_input(f"毛重 KG (Box {i+1})", value=float(box['gross_w']), step=0.1, key=f"gw_{i}")
    
    total_net_weight += box['net_w']
    total_gross_weight += box['gross_w']
    
    box_details_html += f"""
    <tr>
        <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center;">Box {i+1}</td>
        <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center;">{box['dimensions']}</td>
        <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: right;">{box['net_w']:.1f} KG</td>
        <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: right;">{box['gross_w']:.1f} KG</td>
    </tr>
    """

st.info(f"📊 **總計：** 共 {len(st.session_state.boxes)} 箱 | 總淨重 (Total N.W.): **{total_net_weight:.1f} KG** | 總毛重 (Total G.W.): **{total_gross_weight:.1f} KG**")

st.markdown("---")

# 3. 檔案上傳區
st.subheader("📁 3. 上傳內部銷貨單 Excel")
uploaded_file = st.file_uploader("請上傳銷貨單檔案 (.xlsx)", type=["xlsx", "xls"])

items_data = []
header_data = {}

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.read()
        fixed_io = io.BytesIO()
        
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
    header_data = {
        '銷貨單號': '20260722002',
        '客戶全名': '天津元象國際貿易有限公司',
        '送貨地址(一)': '中國天津市濱海新區新北路4668號濱海創新創業園4棟'
    }
    items_data = [
        {"品號": "D13750", "品名與規格": "美國壓縮彈簧 D13750", "數量": 50, "單價": 60.5, "金額": 3025},
        {"品號": "SB20210540", "品名與規格": "Belle Disc Springs For Spindle 二代碟簧", "數量": 100, "單價": 7.2, "金額": 720}
    ]

# 4. 組合商品表格
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

# 5. 組合 HTML 預覽（包含多箱明細表格）
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
    table.items, table.boxes {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
    table.items th, table.items td, table.boxes th, table.boxes td {{ border: 1px solid #cbd5e1; padding: 8px; font-size: 10pt; }}
    table.items th, table.boxes th {{ background-color: #1a365d; color: white; text-align: center; }}
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
        <div class="subtitle">商業發票與裝箱單 (Currency: RMB)</div>
        <hr>
        
        <table class="grid">
            <tr>
                <td class="box" style="width: 50%; vertical-align: top;">
                    <strong>【客戶資訊 (Ship to)】</strong><br>
                    {header_data.get('客戶全名', '')}<br>
                    地址：{header_data.get('送貨地址(一)', '')}
                </td>
                <td class="box" style="width: 50%; vertical-align: top;">
                    <strong>【出貨與物流資訊】</strong><br>
                    Invoice No: {header_data.get('銷貨單號', '')}<br>
                    Date: {doc_date}<br>
                    Currency: <b>RMB (人民幣)</b><br>
                    Tracking No: {tracking_no}<br>
                    Total Boxes: <b>{len(st.session_state.boxes)} Box(es)</b><br>
                    Total N.W.: <b>{total_net_weight:.1f} KG</b><br>
                    Total G.W.: <b>{total_gross_weight:.1f} KG</b>
                </td>
            </tr>
        </table>

        <!-- 多箱明細小表格 -->
        <div style="margin-top: 15px; font-size: 10pt; font-weight: bold; color: #1a365d;">【各箱重量與尺寸明細 (Packing Breakdown)】</div>
        <table class="boxes">
            <thead>
                <tr>
                    <th style="width: 25%;">箱號</th>
                    <th style="width: 35%;">尺寸 (Dimensions)</th>
                    <th style="width: 20%;" class="text-right">淨重 (N.W.)</th>
                    <th style="width: 20%;" class="text-right">毛重 (G.W.)</th>
                </tr>
            </thead>
            <tbody>
                {box_details_html}
            </tbody>
        </table>

        <div style="margin-top: 20px; font-size: 10pt; font-weight: bold; color: #1a365d;">【商品明細 (Items)】</div>
        <table class="items">
            <thead>
                <tr>
                    <th style="width: 10%;">項次</th>
                    <th style="width: 45%;">品名與規格</th>
                    <th class="text-right" style="width: 15%;">數量</th>
                    <th class="text-right" style="width: 15%;">單價 (RMB)</th>
                    <th class="text-right" style="width: 15%;">金額 (RMB)</th>
                </tr>
            </thead>
            <tbody>
                {table_rows_html}
            </tbody>
        </table>

        <div style="text-align: right; font-size: 12pt; font-weight: bold; margin-top: 15px;">
            Total Amount (RMB): RMB {grand_total:,.2f}
        </div>
    </div>
</body>
</html>
"""

st.markdown("---")
st.subheader("📋 4. 正式單據預覽與一鍵列印")
components.html(html_code, height=1050, scrolling=True)
