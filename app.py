import streamlit as st
import pandas as pd
from weasyprint import HTML

# 頁面標題與設定
st.set_page_config(page_title="智能出貨單據轉換工具", page_icon="📦", layout="centered")

st.title("📦 智能出貨單據與物流轉換工具")
st.markdown("上傳內部銷貨單 Excel，自動轉換為 **Invoice & Packing List PDF**，並補齊提單號、重量與材積。")

# 1. 檔案上傳
uploaded_file = st.file_uploader("請上傳內部銷貨單 Excel (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 讀取 ERP 銷貨單結構 (對應 單頭資料 與 單身資料)
        df_head = pd.read_excel(uploaded_file, sheet_name='單頭資料', header=2)
        df_body = pd.read_excel(uploaded_file, sheet_name='單身資料', header=2)
        
        header_data = df_head.iloc[0]
        items = df_body[['品號', '品名', '規格', '數量', '單價', '金額']].dropna(subset=['品號'])
        
        st.success("✅ 銷貨單讀取成功！")
        
        # 顯示訂單摘要
        st.subheader("📋 訂單資訊摘要")
        st.write(f"**銷貨單號：** {header_data.get('銷貨單號', 'N/A')}")
        st.write(f"**客戶名稱：** {header_data.get('客戶全名', 'N/A')}")
        st.write(f"**送貨地址：** {header_data.get('送貨地址(一)', 'N/A')}")
        
        # 2. 填寫物流參數（發票/箱單缺少的動態變數）
        st.subheader("⚙️ 請填寫出貨物流參數")
        col1, col2 = st.columns(2)
        with col1:
            tracking_no = st.text_input("提單號碼 (Tracking No.)", value="SF157xxxxxxxx")
            gross_weight = st.text_input("毛重 (Gross Weight, KG)", value="19.3 KG")
        with col2:
            dimensions = st.text_input("材積尺寸 (Dimensions)", value="42X34X17CM")
            doc_date = st.date_input("單據日期")
            
        # 3. 生成 PDF 按鈕
        if st.button("🚀 生成 Invoice & Packing List PDF"):
            # 建立專業排版的 HTML
            html_invoice = f"""
            <!DOCTYPE html>
            <html>
            <head>
            <meta charset="utf-8">
            <style>
                @page {{ size: A4; margin: 15mm; }}
                body {{ font-family: sans-serif; font-size: 10pt; color: #333; }}
                .header {{ font-size: 18pt; font-weight: bold; text-align: center; margin-bottom: 20px; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
                th, td {{ border: 1px solid #444; padding: 6px; text-align: center; }}
                th {{ background-color: #f2f2f2; }}
                .info {{ margin-bottom: 15px; line-height: 1.6; }}
            </style>
            </head>
            <body>
                <div class="header">INVOICE & PACKING LIST</div>
                <div class="info">
                    <p><b>Invoice No:</b> {header_data.get('銷貨單號', '')}</p>
                    <p><b>Date:</b> {doc_date}</p>
                    <p><b>Ship to:</b> {header_data.get('客戶全名', '')}</p>
                    <p><b>Address:</b> {header_data.get('送貨地址(一)', '')}</p>
                </div>
                <table>
                    <tr><th>Item</th><th>Description</th><th>Quantity</th><th>Unit Price</th><th>Amount</th></tr>
                    {"".join([f"<tr><td>{row['品號']}</td><td>{row['品名']}</td><td>{row['數量']}</td><td>{row['單價']}</td><td>{row['金額']}</td></tr>" for _, row in items.iterrows()])}
                </table>
                <div style="margin-top:20px; line-height: 1.8;">
                    <p><b>Tracking No:</b> {tracking_no}</p>
                    <p><b>Gross Weight:</b> {gross_weight}</p>
                    <p><b>Dimensions:</b> {dimensions}</p>
                </div>
            </body>
            </html>
            """
            
            output_filename = "Invoice_PackingList.pdf"
            HTML(string=html_invoice).write_pdf(output_filename)
            
            # 提供下載按鈕
            with open(output_filename, "rb") as f:
                st.download_button(
                    label="📥 下載產出的 Invoice & Packing List PDF",
                    data=f,
                    file_name="Invoice_PackingList.pdf",
                    mime="application/pdf"
                )
            st.success("🎉 PDF 檔案已成功生成，請點擊上方按鈕下載！")
            
    except Exception as e:
        st.error(f"讀取 Excel 發生錯誤，請確認上傳的銷貨單格式是否正確：{e}")
