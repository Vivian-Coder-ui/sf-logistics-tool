import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import zipfile
import xml.etree.ElementTree as ET

st.set_page_config(page_title="智能出貨單據與物流轉換工具", page_icon="📦", layout="centered")

st.title("📦 智能出貨單據與物流轉換系統 (3合1)")
st.markdown("請在下方填寫物流參數、上傳**內部銷貨單**與**出口正式報單申請書**，系統將自動同步並生成 3 頁完整單據供一鍵列印！")

# 1. 頂部基本資訊與多箱設定
st.subheader("⚙️ 1. 基本出貨與物流參數")
col_a, col_b = st.columns(2)
with col_a:
    tracking_no = st.text_input("托運/提單號碼 (Tracking No.)", value="SF1536155531299")
with col_b:
    doc_date = st.text_input("單據日期 (例: 2026.08.04)", value="2026.08.04")

# 多箱明細動態輸入區
st.markdown("#### 📦 多箱材積與重量明細 (Multi-Box Details)")
if 'boxes' not in st.session_state:
    st.session_state.boxes = [
        {"box_no": 1, "dimensions": "42X34X17CM", "net_w": 11.28, "gross_w": 12.28}
    ]

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
        <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: right;">{box['net_w']:.2f} KG</td>
        <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: right;">{box['gross_w']:.2f} KG</td>
    </tr>
    """

st.info(f"📊 **總計：** 共 {len(st.session_state.boxes)} 件 (箱) | 總淨重: **{total_net_weight:.2f} KG** | 總毛重: **{total_gross_weight:.2f} KG**")

st.markdown("---")

# 2. 檔案上傳區
st.subheader("📁 2. 上傳相關 Excel 檔案")
col_up1, col_up2 = st.columns(2)
with col_up1:
    so_file = st.file_uploader("上傳內部銷貨單 Excel", type=["xlsx", "xls"], key="so")
with col_up2:
    declaration_file = st.file_uploader("上傳出口正式報單申請書 Excel", type=["xlsx", "xls"], key="dec")

# 處理銷貨單
items_data = []
header_data = {}

if so_file is not None:
    try:
        file_bytes = so_file.read()
        fixed_io = io.BytesIO()
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
        
        df_head = pd.read_excel(fixed_io, sheet_name='單頭資料', header=2)
        df_body = pd.read_excel(fixed_io, sheet_name='單身資料', header=2)
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
    except Exception as e:
        st.error(f"❌ 讀取銷貨單發生錯誤：{e}")

if not items_data:
    header_data = {
        '銷貨單號': '20260722002',
        '客戶全名': '天津元象國際貿易有限公司',
        '送貨地址(一)': '中國天津市濱海新區新北路4668號濱海創新創業園4棟'
    }
    items_data = [
        {"品號": "D13750", "品名與規格": "美國壓縮彈簧 D13750", "數量": 50, "單價": 60.5, "金額": 3025}
    ]

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

# 3. 組合完整 3 頁預覽 HTML
html_code = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    body {{
        background: #f1f5f9;
        color: #333;
        font-family: Arial, sans-serif;
        margin: 0;
        padding: 20px;
    }}
    .page {{
        max-width: 750px;
        margin: 30px auto;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 40px;
        background: white;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        page-break-after: always;
    }}
    .print-bar {{
        max-width: 750px;
        margin: auto;
        text-align: center;
        margin-bottom: 20px;
    }}
    .print-btn {{
        background-color: #1a365d;
        color: white;
        border: none;
        padding: 14px 28px;
        font-size: 14pt;
        font-weight: bold;
        border-radius: 6px;
        cursor: pointer;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }}
    .print-btn:hover {{ background-color: #2a4365; }}
    h2 {{ color: #1a365d; margin-bottom: 0px; text-align: center; }}
    .subtitle {{ color: #666; margin-top: 5px; font-size: 11pt; text-align: center; }}
    hr {{ border: 1px solid #1a365d; margin: 15px 0; }}
    .grid {{ width: 100%; margin-top: 10px; border-collapse: collapse; }}
    .box {{ background: #f8fafc; padding: 12px; border-radius: 5px; border: 1px solid #e2e8f0; font-size: 10pt; line-height: 1.5; }}
    table.items, table.boxes, table.dec-table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
    table.items th, table.items td, table.boxes th, table.boxes td, table.dec-table td {{ border: 1px solid #cbd5e1; padding: 8px; font-size: 10pt; }}
    table.items th, table.boxes th {{ background-color: #1a365d; color: white; text-align: center; }}
    .text-right {{ text-align: right; }}
    .highlight {{ background-color: #fed7aa; }} /* 橘色底模擬 */

    @media print {{
        body {{ background: white; padding: 0; }}
        .page {{ border: none; box-shadow: none; padding: 0; max-width: 100%; margin: 0; }}
        .print-bar {{ display: none; }}
    }}
</style>
</head>
<body>
    <div class="print-bar">
        <button class="print-btn" onclick="window.print()">🖨️ 點此列印 / 一鍵另存 3 頁完整 PDF</button>
    </div>

    <!-- ── 第 1 頁：Invoice & Packing List ── -->
    <div class="page">
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
                    Total N.W.: <b>{total_net_weight:.2f} KG</b><br>
                    Total G.W.: <b>{total_gross_weight:.2f} KG</b>
                </td>
            </tr>
        </table>

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

    <!-- ── 第 2 頁：出口正式報單申請書 ── -->
    <div class="page">
        <h2>出口正式報單申請書</h2>
        <div class="subtitle">FORM FOR FORMAL EXPORT DECLARATION</div>
        <hr>
        <table class="dec-table">
            <tr>
                <td style="width: 50%;">托運單號：{tracking_no}</td>
                <td style="width: 50%;" class="highlight">托運日期： {doc_date}</td>
            </tr>
            <tr>
                <td>★統一編號：82850850</td>
                <td>★取貨地址：桃園市蘆竹區安中街20巷13號4樓</td>
            </tr>
            <tr>
                <td>★出口公司：美加卓赫股份有限公司</td>
                <td>★台灣連絡人：陳憲輝</td>
            </tr>
            <tr>
                <td>★緊急聯絡人員及手機(必填)：陳憲輝 0930-906-963</td>
                <td>★公司電話(含分機)：(02) 8201-4393#19</td>
            </tr>
            <tr>
                <td>★出口報單國外買方名稱(付款)：天津元象國際貿易有限公司</td>
                <td>★送件目的地：中國天津市濱海新區新北路4668號濱海創新創業園4棟</td>
            </tr>
            <tr>
                <td>★國外清關聯系人: 魏吉勇先生 Tel:13820625153</td>
                <td>★國外清關聯系電話：13820625153 / 8622-6535-9011</td>
            </tr>
            <tr>
                <td class="highlight">★托寄物中文品名：彈簧</td>
                <td class="highlight">★箱數： {len(st.session_state.boxes)} 件</td>
            </tr>
            <tr>
                <td>★托寄物指定稅則：7320.90.00.00.0</td>
                <td class="highlight">★重量(公斤)：{total_gross_weight:.2f}KG</td>
            </tr>
        </table>
        <div style="margin-top: 15px; font-size: 9pt; color: #666; line-height: 1.6;">
            <b>【報關申報須知】</b><br>
            1. 貿易條件：FOB (不含運保費)<br>
            2. 商標：無商標<br>
            3. 報關類別：02銷售 / 02銷售<br>
            4. 申報視窗：X8@sf-express.com
        </div>
    </div>

    <!-- ── 第 3 頁：個案委任書 ── -->
    <div class="page" style="display: flex; flex-direction: column; justify-content: space-between; height: 850px;">
        <div>
            <h2>個案委任書 (Declaration)</h2>
            <div class="subtitle">快遞貨物進出口報關個案委任書</div>
            <hr>
            <div style="font-size: 11pt; line-height: 1.8; margin-top: 20px;">
                <p>茲委任<b>台灣順豐速運股份有限公司</b>辦理貴公司委託快遞出口貨物之報關、通關及相關查驗事宜，並同意遵守相關法令規定。</p>
                <p><b>委任人 (出口公司)：</b>美加卓赫股份有限公司</p>
                <p><b>統一編號：</b>82850850</p>
                <p><b>負責人：</b>（請加大小章）</p>
            </div>
        </div>
        <div style="text-align: right; font-size: 14pt; font-weight: bold; padding-bottom: 50px;">
            中  華  民  國 &nbsp;&nbsp;<u>&nbsp;{int(doc_date.split('.')[0]) - 1911}&nbsp;</u>&nbsp;&nbsp;年 &nbsp;&nbsp;<u>&nbsp;{doc_date.split('.')[1]}&nbsp;</u>&nbsp;&nbsp;月 &nbsp;&nbsp;<u>&nbsp;{doc_date.split('.')[2]}&nbsp;</u>&nbsp;&nbsp;日
        </div>
    </div>
</body>
</html>
"""

st.markdown("---")
st.subheader("📋 3. 完整 3 頁單據即時預覽與列印")
components.html(html_code, height=1200, scrolling=True)
