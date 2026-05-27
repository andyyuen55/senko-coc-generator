import streamlit as st
import sqlite3
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import io
from docx.oxml import OxmlElement  # ✨ 新增：用來操作 Word 底層畫線的工具
from docx.oxml.ns import qn        # ✨ 新增：用來操作 Word 底層畫線的工具

# ==========================================
# 0. 專屬格子畫框線工具 (精準畫線，不破壞標題)
# ==========================================
def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    # 針對格子的上下左右畫上單條實線 (黑色)
    for edge in ('top', 'left', 'bottom', 'right'):
        existing = tcBorders.find(qn(f'w:{edge}'))
        if existing is not None:
            tcBorders.remove(existing)
        edge_elm = OxmlElement(f'w:{edge}')
        edge_elm.set(qn('w:val'), 'single')
        edge_elm.set(qn('w:sz'), '4') # 線條粗細
        edge_elm.set(qn('w:space'), '0')
        edge_elm.set(qn('w:color'), '000000') # 黑色
        tcBorders.append(edge_elm)

# ==========================================
# 1. 資料庫初始化與升級
# ==========================================
def init_db():
    conn = sqlite3.connect('senko_products.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS main_products 
                 (senko_pn TEXT PRIMARY KEY, description TEXT, country TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS sub_items 
                 (item_pn TEXT PRIMARY KEY, parent_senko_pn TEXT, sub_desc TEXT, sub_country TEXT)''')
    try:
        c.execute("ALTER TABLE sub_items ADD COLUMN sub_country TEXT")
    except sqlite3.OperationalError:
        pass 
    conn.commit()
    return conn

conn = init_db()

# ==========================================
# 2. 網頁介面設定
# ==========================================
st.set_page_config(page_title="SENKO 報表生成系統", layout="wide")
st.title("📦 產品資料庫 & 自動化 COC 報表生成")

tab1, tab2 = st.tabs(["🗂️ 資料庫管理", "📄 生成 COC 報表"])

# --- 分頁 1：資料庫管理 ---
with tab1:
    st.subheader("➕ 新增產品資料")
    col1, col2 = st.columns(2)
    
    with col1:
        senko_pn = st.text_input("總型號 (Senko PN)", key="add_senko")
        description = st.text_area("產品描述 (Description)", key="add_desc")
        country = st.text_input("總型號原產地 (Country of Origin)", value="CHINA", key="add_country")
        
    with col2:
        st.info("💡 格式：`子項目型號 | 描述 | 產地` (可直接從 Excel 複製 3 個欄位貼上)")
        sub_items_input = st.text_area(
            "輸入子項目", 
            placeholder="例如:\nMT-01-SL12-2-NB | MT 12F SM LL | CHINA\nMT-FB-12-1 | MT Ferrule Boot | TAIWAN, CHINA",
            height=150
        )

    if st.button("💾 新增至資料庫"):
        if senko_pn:
            try:
                c = conn.cursor()
                c.execute("INSERT OR REPLACE INTO main_products VALUES (?, ?, ?)", (senko_pn, description, country))
                
                if sub_items_input:
                    lines = sub_items_input.strip().split('\n')
                    for line in lines:
                        if line.strip():
                            parts = []
                            if '\t' in line: parts = line.split('\t')
                            elif '|' in line: parts = line.split('|')
                            elif ',' in line: parts = line.split(',')
                            else: parts = [line]
                            
                            while len(parts) < 3: parts.append("")
                            item_pn, sub_desc, sub_country = [p.strip() for p in parts[:3]]
                            
                            c.execute("INSERT OR REPLACE INTO sub_items (item_pn, parent_senko_pn, sub_desc, sub_country) VALUES (?, ?, ?, ?)", 
                                      (item_pn, senko_pn, sub_desc, sub_country))
                conn.commit()
                st.success(f"✅ 成功新增總型號：{senko_pn}")
            except Exception as e:
                st.error(f"儲存失敗: {e}")
        else:
            st.warning("請至少輸入總型號！")

    st.divider()
    
    st.subheader("📊 目前資料庫概覽")
    df_main = pd.read_sql_query("SELECT * FROM main_products", conn)
    st.table(df_main) 

    st.subheader("📝 修改或刪除產品資料")
    c = conn.cursor()
    c.execute("SELECT senko_pn FROM main_products")
    all_senko = ["(請選擇要修改的型號)"] + [row[0] for row in c.fetchall()]
    
    edit_target = st.selectbox("選擇總型號", all_senko)
    
    if edit_target != "(請選擇要修改的型號)":
        c.execute("SELECT * FROM main_products WHERE senko_pn=?", (edit_target,))
        m_data = c.fetchone()
        
        c.execute("SELECT item_pn, sub_desc, sub_country FROM sub_items WHERE parent_senko_pn=?", (edit_target,))
        s_data = c.fetchall()
        
        s_text = "\n".join([f"{row[0]} | {row[1]} | {row[2]}" for row in s_data])
        
        with st.form("edit_form"):
            st.write(f"正在編輯： **{edit_target}**")
            edit_desc = st.text_area("修改總描述", m_data[1])
            edit_country = st.text_input("修改總產地", m_data[2])
            edit_sub = st.text_area("修改子項目 (格式：型號 | 描述 | 產地)", s_text, height=150)
            
            col_save, col_del = st.columns(2)
            submit_edit = col_save.form_submit_button("💾 儲存修改")
            submit_del = col_del.form_submit_button("❌ 刪除此產品 (包含所有子項目)")
            
            if submit_edit:
                c.execute("UPDATE main_products SET description=?, country=? WHERE senko_pn=?", (edit_desc, edit_country, edit_target))
                c.execute("DELETE FROM sub_items WHERE parent_senko_pn=?", (edit_target,))
                if edit_sub:
                    for line in edit_sub.strip().split('\n'):
                        if line.strip():
                            parts = []
                            if '\t' in line: parts = line.split('\t')
                            elif '|' in line: parts = line.split('|')
                            elif ',' in line: parts = line.split(',')
                            else: parts = [line]
                            while len(parts) < 3: parts.append("")
                            item_pn, sub_desc, sub_country = [p.strip() for p in parts[:3]]
                            c.execute("INSERT INTO sub_items (item_pn, parent_senko_pn, sub_desc, sub_country) VALUES (?, ?, ?, ?)", 
                                      (item_pn, edit_target, sub_desc, sub_country))
                conn.commit()
                st.success("✅ 更新成功！請重新整理頁面查看最新資料。")
                
            if submit_del:
                c.execute("DELETE FROM main_products WHERE senko_pn=?", (edit_target,))
                c.execute("DELETE FROM sub_items WHERE parent_senko_pn=?", (edit_target,))
                conn.commit()
                st.success("🗑️ 產品已刪除！請重新整理頁面。")


# --- 分頁 2：生成 COC 報表 ---
with tab2:
    st.subheader("🚀 一鍵生成 Word 文件 (支援多型號合併)")
    
    col3, col4 = st.columns(2)
    with col3:
        target_senko_input = st.text_area("🔍 輸入要生成的總型號 (Senko PN) [可直接貼上多行]", height=150)
    with col4:
        po_input = st.text_area("📋 貼上 PO Number", height=68)
        invoice_input = st.text_area("📋 貼上 INVOICE NO (必填！)", height=68)
        
    if not invoice_input.strip():
        st.error("⚠️ 錯,少了很多發票號碼！請務必填寫 INVOICE NO 才能生成文件。")
        can_generate = False
    else:
        can_generate = True

    if st.button("生成 Word 檔案", disabled=not can_generate):
        if target_senko_input.strip():
            po_str = ", ".join([p.strip() for p in po_input.split('\n') if p.strip()])
            inv_str = ", ".join([i.strip() for i in invoice_input.split('\n') if i.strip()])
            today_str = datetime.now().strftime("%d-%b-%Y").upper()

            senko_list = [s.strip() for s in target_senko_input.strip().split('\n') if s.strip()]
            
            items_list = []
            not_found_list = []
            c = conn.cursor()

            for target_senko in senko_list:
                c.execute("SELECT * FROM main_products WHERE senko_pn=?", (target_senko,))
                main_data = c.fetchone()
                
                if main_data:
                    c.execute("SELECT * FROM sub_items WHERE parent_senko_pn=?", (target_senko,))
                    sub_data = c.fetchall()
                    
                    m_senko = main_data[0] if main_data[0] else ""
                    m_desc = main_data[1] if main_data[1] else ""
                    m_country = main_data[2] if main_data[2] else ""
                    
                    if sub_data:
                        first_sub = sub_data[0]
                        items_list.append({
                            "senko_pn": m_senko,
                            "item_pn": first_sub[0] if first_sub[0] else "",
                            "description": first_sub[2] if first_sub[2] else m_desc,
                            "country": first_sub[3] if first_sub[3] else m_country
                        })
                        for sub in sub_data[1:]:
                            items_list.append({
                                "senko_pn": "",  
                                "item_pn": sub[0] if sub[0] else "",
                                "description": sub[2] if sub[2] else m_desc,
                                "country": sub[3] if sub[3] else m_country
                            })
                    else:
                        items_list.append({
                            "senko_pn": m_senko,
                            "item_pn": "", 
                            "description": m_desc,
                            "country": m_country
                        })
                else:
                    not_found_list.append(target_senko)

            if not items_list:
                st.error(f"❌ 找不到你輸入的型號，請確認是否已新增至「資料庫管理」！")
            else:
                if not_found_list:
                    st.warning(f"⚠️ 以下型號在資料庫中找不到，已自動忽略：{', '.join(not_found_list)}")
                
                try:
                    doc = DocxTemplate("template.docx")
                    context = {"PO_NUMBER": po_str, "INVOICE_NO": inv_str, "DATE": today_str}
                    doc.render(context)
                    
                    target_table = doc.tables[0] 
                    # ⚠️ 注意：已經把 target_table.style = 'Table Grid' 拿掉，保護原本標題
                    
                    for item in items_list:
                        row_cells = target_table.add_row().cells
                        row_cells[0].text = str(item["senko_pn"])
                        row_cells[1].text = str(item["item_pn"])
                        row_cells[2].text = str(item["description"])
                        row_cells[3].text = str(item["country"])
                        
                        # ✨ 在這裡呼叫畫框線工具，單獨幫這 4 個格子畫上框線
                        for cell in row_cells:
                            set_cell_border(cell)
                    
                    bio = io.BytesIO()
                    doc.save(bio)
                    
                    st.success("🎉 檔案合併生成成功！請點擊下載。")
                    
                    if len(senko_list) == 1 and not not_found_list:
                        file_n = f"COC_{senko_list[0]}_{today_str}.docx"
                    else:
                        file_n = f"COC_Multiple_Items_{today_str}.docx"

                    st.download_button(
                        label="📥 下載 COC Word 檔",
                        data=bio.getvalue(),
                        file_name=file_n,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"生成失敗: {e}")
        else:
            st.warning("請輸入至少一個總型號！")
