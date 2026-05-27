import streamlit as st
import pandas as pd
import gspread
import json
from docxtpl import DocxTemplate
from datetime import datetime
import io
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# 0. 專屬格子畫框線工具 
# ==========================================
def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right'):
        existing = tcBorders.find(qn(f'w:{edge}'))
        if existing is not None: tcBorders.remove(existing)
        edge_elm = OxmlElement(f'w:{edge}')
        edge_elm.set(qn('w:val'), 'single')
        edge_elm.set(qn('w:sz'), '4') 
        edge_elm.set(qn('w:space'), '0')
        edge_elm.set(qn('w:color'), '000000') 
        tcBorders.append(edge_elm)

# ==========================================
# 1. 連線 Google Sheets 
# ==========================================
@st.cache_resource
def init_gsheets():
    try:
        creds_dict = json.loads(st.secrets["gcp_service_account"])
        gc = gspread.service_account_from_dict(creds_dict)
        sh = gc.open_by_url(st.secrets["spreadsheet_url"])
        return sh
    except Exception as e:
        st.error(f"連線失敗，請檢查 Secrets 設定: {e}")
        st.stop()

sh = init_gsheets()
ws_main = sh.worksheet("main_products")
ws_sub = sh.worksheet("sub_items")

def get_main_df():
    records = ws_main.get_all_records()
    return pd.DataFrame(records) if records else pd.DataFrame(columns=["senko_pn", "description", "country"])

def get_sub_df():
    records = ws_sub.get_all_records()
    return pd.DataFrame(records) if records else pd.DataFrame(columns=["item_pn", "parent_senko_pn", "sub_desc", "sub_country"])

def write_df(ws, df):
    ws.clear()
    df = df.fillna("")
    ws.update(values=[df.columns.values.tolist()] + df.values.tolist(), range_name="A1")

# ==========================================
# 2. 網頁介面設定
# ==========================================
st.set_page_config(page_title="SENKO 報表生成系統", layout="wide")
st.title("📦 產品資料庫 & 自動化 COC 報表生成")

tab1, tab2 = st.tabs(["🗂️ 資料庫管理", "📄 生成 COC 報表"])

# --- 分頁 1：資料庫管理 ---
with tab1:
    st.subheader("➕ 新增/修改產品資料")
    col1, col2 = st.columns(2)
    
    with col1:
        senko_pn = st.text_input("總型號 (Senko PN)")
        description = st.text_area("產品描述 (Description)")
        country = st.text_input("總型號原產地 (Country of Origin)", value="CHINA")
        
        # ✨ 新增：防呆確認框
        st.markdown("<br>", unsafe_allow_html=True)
        overwrite_confirm = st.checkbox("⚠️ 我確認這是一筆舊資料，我要「修改/覆蓋」它")
        
    with col2:
        st.info("💡 格式：`子項目型號 | 描述 | 產地` (可直接從 Excel 貼上)")
        sub_items_input = st.text_area("輸入子項目", height=150)

    if st.button("💾 儲存至雲端資料庫"):
        if senko_pn:
            with st.spinner('正在檢查並同步至 Google Sheets...'):
                df_m = get_main_df()
                df_s = get_sub_df()
                
                # ✨ 新增：檢查是否重複
                is_exist = not df_m.empty and senko_pn in df_m["senko_pn"].values
                
                # 如果重複了，且沒有勾選確認框，就擋下來報錯！
                if is_exist and not overwrite_confirm:
                    st.error(f"🛑 發現重複！資料庫中已經有「{senko_pn}」的資料了。\n如果您是想要更新這筆資料，請勾選上方的「⚠️ 我確認這是一筆舊資料...」再按儲存。")
                else:
                    # 處理 Main (新增或覆蓋)
                    if is_exist:
                        df_m.loc[df_m["senko_pn"] == senko_pn, ["description", "country"]] = [description, country]
                    else:
                        new_m = pd.DataFrame([{"senko_pn": senko_pn, "description": description, "country": country}])
                        df_m = pd.concat([df_m, new_m], ignore_index=True)
                    
                    # 處理 Sub (先刪除舊的，再加入新的)
                    if not df_s.empty:
                        df_s = df_s[df_s["parent_senko_pn"] != senko_pn]
                    
                    new_subs = []
                    if sub_items_input:
                        for line in sub_items_input.strip().split('\n'):
                            if line.strip():
                                parts = []
                                if '\t' in line: parts = line.split('\t')
                                elif '|' in line: parts = line.split('|')
                                elif ',' in line: parts = line.split(',')
                                else: parts = [line]
                                while len(parts) < 3: parts.append("")
                                ipn, sdesc, scountry = [p.strip() for p in parts[:3]]
                                new_subs.append({"item_pn": ipn, "parent_senko_pn": senko_pn, "sub_desc": sdesc, "sub_country": scountry})
                    
                    if new_subs:
                        df_s = pd.concat([df_s, pd.DataFrame(new_subs)], ignore_index=True)
                    
                    # 寫回 Google Sheets
                    write_df(ws_main, df_m)
                    write_df(ws_sub, df_s)
                    
                    if is_exist:
                        st.success(f"🔄 成功更新舊型號：{senko_pn}")
                    else:
                        st.success(f"✅ 成功新增新型號：{senko_pn}")
        else:
            st.warning("請至少輸入總型號！")

    st.divider()
    
    st.subheader("📊 目前 Google 試算表資料概覽")
    st.write("💡 提示：你現在可以直接打開 Google 試算表進行修改，網頁重整後會自動讀取最新資料！")
    df_main = get_main_df()
    st.table(df_main) 

    st.subheader("🗑️ 刪除產品資料")
    all_senko = ["(請選擇要刪除的型號)"] + df_main["senko_pn"].tolist() if not df_main.empty else ["(無資料)"]
    del_target = st.selectbox("選擇要刪除的總型號", all_senko)
    if st.button("❌ 刪除此產品") and del_target != "(請選擇要刪除的型號)":
        with st.spinner('刪除中...'):
            df_m = get_main_df()
            df_s = get_sub_df()
            df_m = df_m[df_m["senko_pn"] != del_target]
            df_s = df_s[df_s["parent_senko_pn"] != del_target]
            write_df(ws_main, df_m)
            write_df(ws_sub, df_s)
            st.success("🗑️ 產品已刪除！")
            st.rerun()

# --- 分頁 2：生成 COC 報表 ---
with tab2:
    st.subheader("🚀 一鍵生成 Word 文件 (支援多型號合併)")
    
    col3, col4 = st.columns(2)
    with col3:
        target_senko_input = st.text_area("🔍 輸入要生成的總型號 (Senko PN) [可多行]", height=150)
    with col4:
        po_input = st.text_area("📋 貼上 PO Number", height=68)
        invoice_input = st.text_area("📋 貼上 INVOICE NO (必填！)", height=68)
        
    can_generate = bool(invoice_input.strip())
    if not can_generate: st.error("⚠️ 請務必填寫 INVOICE NO 才能生成文件。")

    if st.button("生成 Word 檔案", disabled=not can_generate):
        if target_senko_input.strip():
            with st.spinner('正在從 Google Sheets 撈取資料並生成...'):
                po_str = ", ".join([p.strip() for p in po_input.split('\n') if p.strip()])
                inv_str = ", ".join([i.strip() for i in invoice_input.split('\n') if i.strip()])
                today_str = datetime.now().strftime("%d-%b-%Y").upper()

                # ✨ 升級 1：將輸入的搜尋型號「去除頭尾空白」並「強制轉大寫」
                senko_list = [str(s).strip().upper() for s in target_senko_input.strip().split('\n') if str(s).strip()]
                
                df_m = get_main_df()
                df_s = get_sub_df()
                
                items_list = []
                not_found_list = []

                # 準備乾淨的比對資料庫 (不改變原始資料，只在背景做轉換比對)
                if not df_m.empty:
                    db_m_clean = df_m["senko_pn"].astype(str).str.strip().str.upper()
                else:
                    db_m_clean = pd.Series([], dtype=str)
                    
                if not df_s.empty and "parent_senko_pn" in df_s.columns:
                    db_s_clean = df_s["parent_senko_pn"].astype(str).str.strip().str.upper()
                else:
                    db_s_clean = pd.Series([], dtype=str)

                for target_senko in senko_list:
                    # ✨ 升級 2：使用乾淨無空白的字串進行比對
                    main_match = df_m[db_m_clean == target_senko]
                    
                    if not main_match.empty:
                        m_row = main_match.iloc[0]
                        m_senko = str(m_row["senko_pn"]) # 寫入 Word 時，依然保留你最原本輸入的樣子
                        m_desc = str(m_row["description"])
                        m_country = str(m_row["country"])
                        
                        sub_match = df_s[db_s_clean == target_senko]
                        
                        if not sub_match.empty:
                            for i, (_, s_row) in enumerate(sub_match.iterrows()):
                                items_list.append({
                                    "senko_pn": m_senko if i == 0 else "",
                                    "item_pn": str(s_row["item_pn"]),
                                    "description": str(s_row["sub_desc"]) if s_row["sub_desc"] else m_desc,
                                    "country": str(s_row["sub_country"]) if s_row["sub_country"] else m_country
                                })
                        else:
                            items_list.append({
                                "senko_pn": m_senko, "item_pn": "", "description": m_desc, "country": m_country
                            })
                    else:
                        # 紀錄使用者原本輸入的字樣，方便除錯
                        original_input_senko = target_senko_input.strip().split('\n')[senko_list.index(target_senko)]
                        not_found_list.append(original_input_senko)

                if not_found_list:
                    st.error(f"🛑 生成失敗！系統在資料庫中找不到以下型號：\n\n**{', '.join(not_found_list)}**\n\n👉 **請先前往「🗂️ 資料庫管理」分頁新增這些產品資料，然後再回來重新生成。**")
                else:
                    try:
                        doc = DocxTemplate("template.docx")
                        doc.render({"PO_NUMBER": po_str, "INVOICE_NO": inv_str, "DATE": today_str})
                        
                        target_table = doc.tables[0] 
                        for item in items_list:
                            row_cells = target_table.add_row().cells
                            row_cells[0].text = item["senko_pn"]
                            row_cells[1].text = item["item_pn"]
                            row_cells[2].text = item["description"]
                            row_cells[3].text = item["country"]
                            for cell in row_cells:
                                set_cell_border(cell)
                        
                        bio = io.BytesIO()
                        doc.save(bio)
                        
                        file_n = f"COC_{senko_list[0]}_{today_str}.docx" if len(senko_list) == 1 else f"COC_Multiple_Items_{today_str}.docx"

                        st.success("🎉 檔案生成成功！")
                        st.download_button(label="📥 下載 COC Word 檔", data=bio.getvalue(), file_name=file_n, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    except Exception as e:
                        st.error(f"生成失敗: {e}")
        else:
            st.warning("請輸入至少一個總型號！")
