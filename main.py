import streamlit as st
import io
from docx import Document

st.set_page_config(page_title="SOP 範本下載器")
st.title("🎁 專屬 SOP 範本生成器")
st.info("點擊下方按鈕，系統會自動幫你製作並下載排版完美的 `sop_template.docx`！")

# 建立 Word 檔案與排版
doc = Document()
doc.add_heading('📦 出貨標準作業程序 (Shipment SOP)', 0)
doc.add_paragraph('規範主題：出貨作業流程與規範 / Shipment Process and Regulations')

doc.add_heading('📌 1. 基本出貨資訊 (General Shipping Info)', level=2)
t1 = doc.add_table(rows=4, cols=2)
t1.style = 'Table Grid'
t1.cell(0,0).text, t1.cell(0,1).text = '建立日期 (Date)：', '{{ DATE }}'
t1.cell(1,0).text, t1.cell(1,1).text = '客戶編號 (Customer Code)：', '{{ CUSTOMER_ID }}'
t1.cell(2,0).text, t1.cell(2,1).text = '出貨方式與帳號 (Ship Via & Account)：', '{{ SHIPPING_METHOD_INFO }}'
t1.cell(3,0).text, t1.cell(3,1).text = '出貨必備文件 (Required Docs)：', '{{ REQUIRED_DOCS }}'

doc.add_heading('⚠️ 2. 出貨注意事項 (Shipping Notes)', level=2)
doc.add_paragraph('{{ SHIPPING_NOTES }}')

doc.add_heading('🏷️ 3. 標籤格式與黏貼規範 (Label Format & Guidelines)', level=2)
doc.add_paragraph('規範要求：每件貨物在外箱上必須黏貼符合客戶及物流商要求的標籤。標籤資訊必須清晰、無摺疊、無遮擋。\nGuideline: Each cargo carton must be affixed with a label that meets customer and logistics provider requirements. Label information must be clear, unwrinkled, and unobstructed.')
doc.add_paragraph('客戶要求標籤包含項目： {{ LABEL_FORMATS }}\n標籤圖示說明 (Label Reference)：\n{{ LABEL_IMAGE }}')

doc.add_heading('👥 4. 職務代理與聯繫窗口 (Staff & Contact Matrix)', level=2)
t2 = doc.add_table(rows=4, cols=2)
t2.style = 'Table Grid'
t2.cell(0,0).text, t2.cell(0,1).text = '負責角色 (Role)', '聯絡人姓名 (Name)'
t2.cell(1,0).text, t2.cell(1,1).text = '主要負責同事 (Primary Owner)', '{{ MAIN_CONTACT }}'
t2.cell(2,0).text, t2.cell(2,1).text = '職務代理人 (Backup Personnel)', '{{ BACKUP_CONTACT }}'
t2.cell(3,0).text, t2.cell(3,1).text = '對應業務群 (Corresponding Sales)', '{{ RESPONSIBLE_SALES }}'

# 存檔至記憶體
bio = io.BytesIO()
doc.save(bio)

# 產生下載按鈕
st.download_button(
    label="📥 點擊這裡！立刻下載排版完美的 sop_template.docx",
    data=bio.getvalue(),
    file_name="sop_template.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
