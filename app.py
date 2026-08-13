import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import zipfile
import xml.etree.ElementTree as ET

st.set_page_config(page_title="智能出貨單據與物流轉換工具", page_icon="📦", layout="centered")

st.title("📦 智能出貨單據與物流轉換系統 (完整 3 合 1)")

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
total_nw = 0

for i, b in enumerate(st.session_state.boxes):
    st.markdown(f"**── 第 {i+1} 箱 ──**")
    bc1, bc2, bc3 = st.columns(3)
    with bc1:
        st.session_state.boxes[i]['dim'] = st.text_input(f"尺寸 (Box {i+1})", value=b['dim'], key=f"dim_box_{i}")
    with bc2:
        st.session_state.boxes[i]['nw'] = st.number_input(f"淨重 KG (Box {i+1})", value=float(b['nw']), step=0.1, key=f"nw_box_{i}")
    with bc3:
        st.session_state.boxes[i]['gw'] = st.number_input(f"毛重 KG (Box {i+1})", value=float(b['gw']), step=0.1, key=f"gw_box_{i}")
    
    total_nw += st.session_state.boxes[i]['nw']
    total_gw += st.session_state.boxes[i]['gw']
    box_rows += f"<tr><td>Box {i+1}</td><td>{st.session_state.boxes[i]['dim']}</td><td>{st.session_state.boxes[i]['nw']} KG</td><td>{st.session_state.boxes[i]['gw']} KG</td></tr>"

st.info(f"📊 **總計：** 共 {len(st.session_state.boxes)} 件 | 總淨重: **{total_nw:.2f} KG** | 總毛重: **{total_gw:.2f} KG**")

# --- 2. 檔案上傳與預覽 ---
st.markdown("---")
st.subheader("📁 2. 上傳內部銷貨單 Excel")
so_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx", "xls"])

items_data = [{"品號": "D13750", "品名與規格": "美國壓縮彈簧 D13750", "數量": 50, "單價": 60.5, "金額": 3025}]
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
        items_df = df_body[['品號', '品名', '規格', '數量', '單價', '金額']].dropna(subset=['品號'])
        items_data = []
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
        st.error(f"讀取 Excel 發生錯誤: {e}")

table_rows = "".join([f"<tr><td>{i+1}</td><td>{item['品號']}<br>{item['品名與規格']}</td><td style='text-align:right;'>{item['數量']:,}</td><td style='text-align:right;'>{item['單價']:,.2f}</td><td style='text-align:right;'>{item['金額']:,.2f}</td></tr>" for i, item in enumerate(items_data)])
grand_total = sum([item['金額'] for item in items_data])

# --- 3. 3頁整合 HTML (完整內容) ---
html_code = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    body {{ background: #f1f5f9; font-family: Arial, sans-serif; margin: 0; padding: 20px; }}
    .page {{ background: white; width: 750px; margin: 20px auto; padding: 40px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); page-break-after: always; }}
    .print-btn {{ position: fixed; top: 20px; right: 20px; padding: 12px 24px; background: #1a365d; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 12pt; font-weight: bold; z-index: 1000; box-shadow: 0 2px 4px rgba(0,0,0,0.2); }}
    .print-btn:hover {{ background: #2a4365; }}
    h2 {{ color: #1a365d; text-align: center; margin-bottom: 5px; }}
    .subtitle {{ text-align: center; color: #666; font-size: 10pt; margin-bottom: 15px; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
    td, th {{ border: 1px solid #1a365d; padding: 6px 8px; vertical-align: top; font-size: 9pt; line-height: 1.4; }}
    th {{ background: #1a365d; color: white; text-align: center; }}
    .highlight {{ background-color: #ffe6cc; }}
    .section-title {{ font-weight: bold; background: #e2e8f0; padding: 6px; margin-top: 10px; font-size: 9pt; border: 1px solid #1a365d; }}
    @media print {{ body {{ background: white; padding: 0; }} .print-btn {{ display: none; }} .page {{ box-shadow: none; margin: 0; width: 100%; padding: 0; }} }}
</style>
</head>
<body>
    <button class="print-btn" onclick="window.print()">🖨️ 一鍵列印 / 存為完整 3 頁 PDF</button>

    <!-- 第 1 頁：Invoice & Packing List -->
    <div class="page">
        <h2>INVOICE & PACKING LIST</h2>
        <div class="subtitle">商業發票與裝箱單 (Currency: RMB)</div>
        <hr style="border: 1px solid #1a365d; margin-bottom: 15px;">
        <table>
            <tr>
                <td style="width: 50%;"><b>【客戶資訊 (Ship to)】</b><br>天津元象國際貿易有限公司<br>地址：中國天津市濱海新區新北路4668號</td>
                <td style="width: 50%;"><b>【出貨資訊】</b><br>Date: {doc_date}<br>Tracking No: {tracking_no}<br>Total N.W.: {total_nw:.2f} KG<br>Total G.W.: {total_gw:.2f} KG</td>
            </tr>
        </table>
        <br><b>【各箱重量與尺寸明細】</b>
        <table>
            <tr><th>箱號</th><th>尺寸 (Dimensions)</th><th>淨重 (N.W.)</th><th>毛重 (G.W.)</th></tr>
            {box_rows}
        </table>
        <br><b>【商品明細】</b>
        <table>
            <tr><th>項次</th><th>品名與規格</th><th style="text-align:right;">數量</th><th style="text-align:right;">單價 (RMB)</th><th style="text-align:right;">金額 (RMB)</th></tr>
            {table_rows}
        </table>
        <div style="text-align: right; font-weight: bold; margin-top: 15px; font-size: 11pt;">
            Total Amount (RMB): RMB {grand_total:,.2f}
        </div>
    </div>

    <!-- 第 2 頁：出口正式報單申請書 (完整 Excel 原貌) -->
    <div class="page">
        <h2>出口正式報單申請書</h2>
        <div class="subtitle">正式報關受理窗口: X8@sf-express.com 傳真: 02-27128032</div>
        <table>
            <tr><td style="width: 50%;">托運單號：{tracking_no}</td><td style="width: 50%;" class="highlight">托運日期： {doc_date}</td></tr>
            <tr><td>★統一編號：82850850</td><td>★取貨地址：桃園市蘆竹區安中街20巷13號4樓</td></tr>
            <tr><td>★出口公司：美加卓赫股份有限公司</td><td>★台灣連絡人：陳憲輝</td></tr>
            <tr><td>★緊急聯絡人員及手機(必填)：陳憲輝 0930-906-963</td><td>★公司電話(含分機)：(02) 8201-4393#19</td></tr>
            <tr><td>★出口報單國外買方名稱(付款)：天津元象國際貿易有限公司</td><td>★送件目的地：中國天津市濱海新區新北路4668號</td></tr>
            <tr><td>★國外清關聯系人: 魏吉勇先生 Tel:13820625153</td><td>★國外清關聯系電話：13820625153 / 8622-6535-9011</td></tr>
            <tr><td class="highlight">★托寄物中文品名：彈簧</td><td class="highlight">★箱數： {len(st.session_state.boxes)} 件</td></tr>
            <tr><td>★托寄物指定稅則：7320.90.00.00.0</td><td class="highlight">★重量(公斤)：{total_gw:.2f}KG</td></tr>
        </table>
        <div class="section-title">★ 請勾選欲申請報關方式及報關類別</div>
        <table>
            <tr><td><b>一、台灣出口正式報單：</b> ▓ 空運(NTD 500元)</td><td><b>二、大陸進口正式報關：</b> ▓ D類正式進口</td></tr>
            <tr><td><b>貿易條件：</b> ▓ FOB (不含運保費)</td><td><b>商標：</b> ▓ 無商標</td></tr>
            <tr><td colspan="2"><b>報關類別：</b> ▓ 02銷售</td></tr>
        </table>
    </div>

    <!-- 第 3 頁：個案委任書 -->
    <div class="page" style="height: 800px; display: flex; flex-direction: column; justify-content: space-between;">
        <div>
            <h2>個案委任書 (Declaration)</h2>
            <div class="subtitle">快遞貨物進出口報關個案委任書</div>
            <hr style="border: 1px solid #1a365d; margin-bottom: 20px;">
            <p style="line-height: 1.8; font-size: 11pt;">
                茲委任<b>台灣順豐速運股份有限公司</b>辦理貴公司委託快遞出口貨物之報關、通關及相關查驗事宜，並同意遵守相關法令規定。
            </p>
            <br>
            <p style="line-height: 1.8; font-size: 11pt;">
                <b>委任人 (出口公司)：</b>美加卓赫股份有限公司<br>
                <b>統一編號：</b>82850850<br>
                <b>負責人 / 蓋章：</b>____________________
            </p>
        </div>
        <div style="text-align: right; font-size: 13pt; font-weight: bold; padding-bottom: 40px;">
            中華民國 {int(doc_date.split('.')[0])-1911} 年 {doc_date.split('.')[1]} 月 {doc_date.split('.')[2]} 日
        </div>
    </div>
</body>
</html>
"""
components.html(html_code, height=1200, scrolling=True)
