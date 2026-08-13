import streamlit as st
import pandas as pd
from fpdf import FPDF
import tempfile
import os

st.set_page_config(page_title="智能出貨單據轉換工具", page_icon="📦", layout="centered")

st.title("📦 智能出貨單據與物流轉換工具")
st.markdown("上傳內部銷貨單 Excel，手動填入物流參數，一鍵生成 **Invoice & Packing List PDF**！")

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
        
        # 2. 讓使用者直接在網頁上填寫變數
        st.subheader("⚙️ 請手動輸入本次出貨物流參數")
        col1, col2 = st.columns(2)
        with col1:
            tracking_no = st.text_input("提單號碼 (Tracking No.)", value="SF157xxxxxxxx")
            gross_weight = st.text_input("毛重 (Gross Weight)", value="19.3 KG")
        with col2:
            dimensions = st.text_input("材積尺寸 (Dimensions)", value="42X34X17CM")
            doc_date = st.text_input("單據日期", value="2026-08-12")
            
        # 3. 生成 PDF 按鈕
        if st.button("🚀 生成 Invoice & Packing List PDF"):
            # 使用 FPDF 輕量生成 PDF
            class PDF(FPDF):
                def header(self):
                    self.set_font("helvetica", "B", 16)
                    self.cell(0, 10, "INVOICE & PACKING LIST", align="C", new_x="LMARGIN", new_y="NEXT")
                    self.ln(5)

                def footer(self):
                    self.set_y(-15)
                    self.set_font("helvetica", "I", 8)
                    self.cell(0, 10, f"Page {self.page_no()}", align="C")

            pdf = PDF()
            pdf.add_page()
            
            # 設定字型 (使用內建 Helvetica)
            pdf.set_font("helvetica", "", 10)
            
            # 基本資訊
            pdf.cell(0, 6, f"Invoice No: {header_data.get('銷貨單號', '')}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 6, f"Date: {doc_date}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 6, f"Ship to: {str(header_data.get('客戶全名', ''))}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 6, f"Address: {str(header_data.get('送貨地址(一)', ''))}", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(5)
            
            # 表格標題
            pdf.set_font("helvetica", "B", 10)
            pdf.cell(35, 8, "Item", border=1, align="C")
            pdf.cell(85, 8, "Description", border=1, align="C")
            pdf.cell(20, 8, "Qty", border=1, align="C")
            pdf.cell(25, 8, "Price", border=1, align="C")
            pdf.cell(25, 8, "Amount", border=1, align="C", new_x="LMARGIN", new_y="NEXT")
            
            # 表格內容
            pdf.set_font("helvetica", "", 9)
            for _, row in items.iterrows():
                pdf.cell(35, 7, str(row['品號']), border=1, align="C")
                pdf.cell(85, 7, str(row['品名'])[:40], border=1, align="L")
                pdf.cell(20, 7, str(row['數量']), border=1, align="C")
                pdf.cell(25, 7, str(row['單價']), border=1, align="C")
                pdf.cell(25, 7, str(row['金額']), border=1, align="C", new_x="LMARGIN", new_y="NEXT")
                
            pdf.ln(10)
            
            # 物流與重量資訊
            pdf.set_font("helvetica", "B", 10)
            pdf.cell(0, 6, f"Tracking No: {tracking_no}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 6, f"Gross Weight: {gross_weight}", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 6, f"Dimensions: {dimensions}", new_x="LMARGIN", new_y="NEXT")
            
            # 輸出暫存檔案
            output_filename = "Invoice_PackingList.pdf"
            pdf.output(output_filename)
            
            # 提供下載按鈕
            with open(output_filename, "rb") as f:
                st.download_button(
                    label="📥 下載產出的 Invoice & Packing List PDF",
                    data=f,
                    file_name="Invoice_PackingList.pdf",
                    mime="application/pdf"
                )
            st.success("🎉 PDF 檔案已成功生成！")
            
    except Exception as e:
        st.error(f"讀取 Excel 發生錯誤：{e}")
