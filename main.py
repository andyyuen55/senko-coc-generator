import streamlit as st
import pandas as pd
import gspread
import json
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from datetime import datetime
import io
import os

# ==========================================
# 0. 專屬格子畫框線工具 
# ==========================================
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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
ws_sop = sh.worksheet("shipment_sops")  # ✨ 新增：SOP 工作表

def get_main_df():
    records = ws_main.get_all_records()
    return pd.DataFrame(records) if records else pd.DataFrame(columns=["senko_pn", "description", "country"])

def get_sub_df():
    records = ws_sub.get_all_records()
    return pd.DataFrame(records) if records else pd.DataFrame(columns=["item_pn", "parent_senko_pn", "sub_desc", "sub_country"])

def get_sop_df():
    records = ws_sop.get_all_records()
    return pd.DataFrame(records) if records else pd.DataFrame(columns=["customer_id", "required_docs", "label_formats", "shipping_notes", "shipping_method_info", "main_contact", "backup_contact", "responsible_sales"])

def write_df(ws, df):
    ws.clear()
    df = df.fillna("")
    ws.update(values=[df.columns.values.tolist()] + df.values.tolist(), range_name="A1")

# ==========================================
# 2. 網頁介面設定
# ==========================================
st.set_page_config(page_title="SENKO 綜合管理系統", layout="wide")
st.title("📦 SENKO 產品資料庫 & 自動化報表系統")

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "auto_senko" not in st.session_state:
    st.session_state.auto_senko = ""
if "auto_po" not in st.session_state:
    st.session_state.auto_po = ""

# 建立三個分頁
tab1, tab2, tab3 = st.tabs(["🗂️ 產品資料庫管理", "📄 生成 COC 報表", "📦 生成 SHIPMENT SOP"])

# ==========================================
# --- 分頁 1：產品資料庫管理 ---
# ==========================================
with tab1:
    if not st.session_state.logged_in:
        st.subheader("🔒 管理員登入")
        col_pwd, col_empty = st.columns([1, 2])
        with col_pwd:
            pwd_input = st.text_input("管理員密碼", type="password")
            if st.button("🔑 登入系統"):
                if pwd_input == st.secrets["admin_password"]:
                    st.session_state.logged_in = True
                    st.success("登入成功！")
                    st.rerun()
                else:
                    st.error("密碼錯誤！")
    else:
        col_title, col_logout = st.columns([4, 1])
        with col_title: st.subheader("➕ 新增/修改產品資料")
        with col_logout:
            if st.button("🚪 登出"):
                st.session_state.logged_in = False
                st.rerun()
                
        col1, col2 = st.columns(2)
        with col1:
            senko_pn = st.text_input("總型號 (Senko PN)")
            description = st.text_area("產品描述 (Description)")
            country = st.text_input("總型號原產地 (Country of Origin)", value="CHINA")
            overwrite_confirm = st.checkbox("⚠️ 我確認這是一筆舊資料，我要「修改/覆蓋」它", key="over_prod")
            
        with col2:
            sub_items_input = st.text_area("輸入子項目 (型號 | 描述 | 產地)", height=150)

        if st.button("💾 儲存產品至雲端"):
            if senko_pn:
                df_m = get_main_df()
                df_s = get_sub_df()
                is_exist = not df_m.empty and senko_pn in df_m["senko_pn"].values
                
                if is_exist and not overwrite_confirm:
                    st.error(f"🛑 發現重複！資料庫中已經有「{senko_pn}」了。請勾選確認框以進行更新。")
                else:
                    if is_exist:
                        df_m.loc[df_m["senko_pn"] == senko_pn, ["description", "country"]] = [description, country]
                    else:
                        df_m = pd.concat([df_m, pd.DataFrame([{"senko_pn": senko_pn, "description": description, "country": country}])], ignore_index=True)
                    
                    if not df_s.empty: df_s = df_s[df_s["parent_senko_pn"] != senko_pn]
                    new_subs = []
                    if sub_items_input:
                        for line in sub_items_input.strip().split('\n'):
                            if line.strip():
                                parts = [p.strip() for p in (line.split('\t') if '\t' in line else line.split('|') if '|' in line else [line])]
                                while len(parts) < 3: parts.append("")
                                new_subs.append({"item_pn": parts[0], "parent_senko_pn": senko_pn, "sub_desc": parts[1], "sub_country": parts[2]})
                    if new_subs: df_s = pd.concat([df_s, pd.DataFrame(new_subs)], ignore_index=True)
                    
                    write_df(ws_main, df_m)
                    write_df(ws_sub, df_s)
                    st.success("✅ 產品資料同步成功！")
            else: st.warning("請輸入總型號！")

        st.divider()
        df_main = get_main_df()
        with st.expander("👀 點擊這裡展開 / 隱藏所有產品清單"):
            search_term = st.text_input("🔍 快速搜尋總型號:", "")
            if search_term.strip():
                st.table(df_main[df_main["senko_pn"].astype(str).str.contains(search_term.strip(), case=False, na=False)])
            else: st.table(df_main)

# ==========================================
# --- 分頁 2：生成 COC 報表 ---
# ==========================================
with tab2:
    st.subheader("🚀 一鍵生成 Word 文件 (支援多型號合併)")
    with st.expander("✨ 智能 Excel 解析器 (可直接從出貨單複製貼上)", expanded=False):
        raw_excel = st.text_area("📋 貼上 Excel 內容", height=120)
        if st.button("🛠️ 自動清洗並填入下方欄位"):
            if raw_excel.strip():
                try:
                    df_paste = pd.read_csv(io.StringIO(raw_excel), sep='\t', dtype=str)
                    po_col = next((c for c in df_paste.columns if "PO" in str(c).upper()), None)
                    senko_col = next((c for c in df_paste.columns if "SENKO" in str(c).upper() or "P/N" in str(c).upper()), None)
                    if po_col and senko_col:
                        pos = df_paste[po_col].dropna().astype(str).str.strip().str.replace(r'\.0$', '', regex=True).unique()
                        senkos = df_paste[senko_col].dropna().astype(str).str.replace(r'\*[^+]*', '', regex=True).str.strip().unique()
                        st.session_state.auto_po = "\n".join(pos)
                        st.session_state.auto_senko = "\n".join(senkos)
                        st.success("✅ 擷取成功！")
                        st.rerun()
                    else: st.error("⚠️ 找不到對應的標題列。")
                except Exception as e: st.error(f"解析失敗: {e}")

    col3, col4 = st.columns(2)
    with col3: target_senko_input = st.text_area("🔍 輸入總型號 [可多行]", key="auto_senko", height=150)
    with col4:
        po_input = st.text_area("📋 貼上 PO Number", key="auto_po", height=68)
        invoice_input = st.text_area("📋 貼上 INVOICE NO (必填！)", height=68)
        
    if st.button("生成 COC Word 檔案", disabled=not bool(invoice_input.strip())):
        if target_senko_input.strip():
            with st.spinner('正在生成...'):
                po_str, inv_str, today_str = ", ".join([p.strip() for p in po_input.split('\n') if p.strip()]), ", ".join([i.strip() for i in invoice_input.split('\n') if i.strip()]), datetime.now().strftime("%d-%b-%Y").upper()
                senko_list = [str(s).strip().upper() for s in target_senko_input.strip().split('\n') if str(s).strip()]
                df_m, df_s, items_list, not_found_list = get_main_df(), get_sub_df(), [], []
                
                db_m_clean = df_m["senko_pn"].astype(str).str.strip().str.upper() if not df_m.empty else pd.Series([], dtype=str)
                db_s_clean = df_s["parent_senko_pn"].astype(str).str.strip().str.upper() if not df_s.empty else pd.Series([], dtype=str)

                for target_senko in senko_list:
                    main_match = df_m[db_m_clean == target_senko]
                    if not main_match.empty:
                        m_row = main_match.iloc[0]
                        sub_match = df_s[db_s_clean == target_senko]
                        if not sub_match.empty:
                            for idx, (_, s_row) in enumerate(sub_match.iterrows()):
                                items_list.append({"senko_pn": str(m_row["senko_pn"]) if idx == 0 else "", "item_pn": str(s_row["item_pn"]), "description": str(s_row["sub_desc"]) if s_row["sub_desc"] else str(m_row["description"]), "country": str(s_row["sub_country"]) if s_row["sub_country"] else str(m_row["country"])})
                        else: items_list.append({"senko_pn": str(m_row["senko_pn"]), "item_pn": "", "description": str(m_row["description"]), "country": str(m_row["country"])})
                    else: not_found_list.append(target_senko)

                if not_found_list: st.error(f"🛑 找不到型號：{', '.join(not_found_list)}，請先至分頁 1 新增。")
                else:
                    try:
                        doc = DocxTemplate("template.docx")
                        doc.render({"PO_NUMBER": po_str, "INVOICE_NO": inv_str, "DATE": today_str})
                        t_table = doc.tables[0]
                        for item in items_list:
                            row_cells = t_table.add_row().cells
                            for k, v in enumerate([item["senko_pn"], item["item_pn"], item["description"], item["country"]]): row_cells[k].text = v
                            for cell in row_cells: set_cell_border(cell)
                        bio = io.BytesIO()
                        doc.save(bio)
                        st.success("🎉 COC 生成成功！")
                        st.download_button("📥 下載 COC Word 檔", data=bio.getvalue(), file_name=f"COC_{senko_list[0]}_{today_str}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    except Exception as e: st.error(f"生成失敗: {e}")

# ==========================================
# --- ✨ 分頁 3：生成 SHIPMENT SOP ---
# ==========================================
with tab3:
    st.subheader("📦 客戶出貨標準作業程序 (SHIPMENT SOP) 管理")
    
    # 讀取現有 SOP 資料庫
    df_sop = get_sop_df()
    
    # 客戶 ID 輸入（具備自動抓取舊資料功能）
    c_id = st.text_input("1. 🔑 請輸入 CUSTOMER ID (大寫)", "").strip().upper()
    
    # 預設值初始化
    exist_row = df_sop[df_sop["customer_id"].astype(str).str.upper() == c_id].iloc[0] if (not df_sop.empty and c_id in df_sop["customer_id"].astype(str).str.upper().values) else None
    
    if exist_row is not None:
        st.success(f"ℹ️ 偵測到已存在「{c_id}」的舊資料，下方已自動帶出，修改後點儲存即可覆蓋。")
        old_docs = [d.strip() for d in str(exist_row["required_docs"]).split(",") if d.strip()]
        old_labels = [l.strip() for l in str(exist_row["label_formats"]).split(",") if l.strip()]
        old_notes = str(exist_row["shipping_notes"])
        old_sales = str(exist_row["responsible_sales"])
        old_main_c = str(exist_row["main_contact"])
        old_backup_c = str(exist_row["backup_contact"])
    else:
        old_docs, old_labels, old_notes, old_sales, old_main_c, old_backup_c = [], [], "", "", "", ""

    st.divider()
    
    col_sop1, col_sop2 = st.columns(2)
    
    with col_sop1:
        st.markdown("**2. 出貨需要的相關文件 (可多選)**")
        doc_options = ["INVOICE", "PI", "PACKING LIST", "CO", "NCV"]
        selected_docs = [doc for doc in doc_options if st.checkbox(doc, value=(doc in old_docs), key=f"doc_{doc}")]
        
        st.markdown("<br>**3. LABEL FORMAT 外箱標籤格式 (可多選)**", unsafe_allow_html=True)
        label_options = ["DESC", "PN", "QTY", "CUSTOMER PO", "CUSTOMER PN", "COO", "CARTON NO"]
        selected_labels = [lbl for lbl in label_options if st.checkbox(lbl, value=(lbl in old_labels), key=f"lbl_{lbl}")]
        
        # 圖片上傳功能
        uploaded_label_img = st.file_uploader("🖼️ 可選：附上 LABEL 圖片作說明 (支援 PNG / JPG)", type=["png", "jpg", "jpeg"])
        
    with col_sop2:
        shipping_notes = st.text_area("4. 📝 出貨注意事項 (可直接貼上)", value=old_notes, height=120)
        
        st.markdown("**5. 出貨方式設定**")
        ship_method = st.selectbox("選擇出貨管道", ["FEDEX", "UPS", "DHL", "SF", "FORWARDER (貨代形式/自取)"])
        
        # 根據出貨方式動態切換輸入框
        if ship_method != "FORWARDER (貨代形式/自取)":
            acc_no = st.text_input("運費付款 ACCOUNT", "")
            tax_no = st.text_input("TAX 付款 ACCOUNT", "")
            method_info_str = f"出貨管道: {ship_method}\n運費帳號: {acc_no}\n稅金帳號: {tax_no}"
        else:
            fw_name = st.text_input("貨代聯絡人姓名", "")
            fw_email = st.text_input("貨代 EMAIL", "")
            fw_tel = st.text_input("貨代 TEL", "")
            method_info_str = f"出貨管道: 貨代自取\n聯絡人: {fw_name}\nEMAIL: {fw_email}\nTEL: {fw_tel}"

    st.divider()
    st.markdown("**6. 內部與客戶對接窗口人員**")
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1: main_contact = st.text_input("CUSTOMER 主要負責同事名稱", value=old_main_c)
    with col_p2: backup_contact = st.text_input("BACKUP 同事名稱", value=old_backup_c)
    with col_p3: responsible_sales = st.text_input("CUSTOMER 主要負責 SALES", value=old_sales)

    # 按鈕區
    st.markdown("<br>", unsafe_allow_html=True)
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        # 7. 儲存/建立資料庫功能
        if st.button("💾 儲存/更新此客戶 SOP 至雲端資料庫"):
            if c_id:
                with st.spinner('正在同步至 Google Sheets...'):
                    # 重新讀取確保最新
                    df_sop_current = get_sop_df()
                    
                    new_sop_data = {
                        "customer_id": c_id,
                        "required_docs": ", ".join(selected_docs),
                        "label_formats": ", ".join(selected_labels),
                        "shipping_notes": shipping_notes,
                        "shipping_method_info": method_info_str,
                        "main_contact": main_contact,
                        "backup_contact": backup_contact,
                        "responsible_sales": responsible_sales
                    }
                    
                    # 判斷是新增還是更新
                    if not df_sop_current.empty and c_id in df_sop_current["customer_id"].astype(str).str.upper().values:
                        df_sop_current.loc[df_sop_current["customer_id"].astype(str).str.upper() == c_id, list(new_sop_data.keys())] = list(new_sop_data.values())
                        st.toast("🔄 舊資料更新成功！")
                    else:
                        df_sop_current = pd.concat([df_sop_current, pd.DataFrame([new_sop_data])], ignore_index=True)
                        st.toast("✅ 新資料建立成功！")
                        
                    write_df(ws_sop, df_sop_current)
                    st.success(f"🎉 客戶「{c_id}」的出貨 SOP 已成功安全儲存至 Google 雲端！")
            else: st.error("請先輸入 CUSTOMER ID 才能儲存！")
            
    with col_btn2:
        # 8. 最終輸出 WORD 功能
        if st.button("🖨️ 導出此客戶之 Word 規範檔"):
            if c_id:
                try:
                    doc_sop = DocxTemplate("sop_template.docx")
                    today_sop_str = datetime.now().strftime("%d-%b-%Y").upper()
                    
                    # 準備圖片物件 (如果使用者有上傳圖片)
                    img_obj = ""
                    if uploaded_label_img is not None:
                        # 讀取圖片並設定寬度為 4 英吋 (自動縮小防破版)
                        img_obj = InlineImage(doc_sop, io.BytesIO(uploaded_label_img.read()), width=Inches(4.0))
                    
                    # 組合轉譯包
                    sop_context = {
                        "CUSTOMER_ID": c_id,
                        "DATE": today_sop_str,
                        "REQUIRED_DOCS": ", ".join(selected_docs) if selected_docs else "無特別要求",
                        "LABEL_FORMATS": ", ".join(selected_labels) if selected_labels else "無特別要求",
                        "SHIPPING_NOTES": shipping_notes if shipping_notes.strip() else "無",
                        "SHIPPING_METHOD_INFO": method_info_str,
                        "MAIN_CONTACT": main_contact if main_contact.strip() else "未指定",
                        "BACKUP_CONTACT": backup_contact if backup_contact.strip() else "未指定",
                        "RESPONSIBLE_SALES": responsible_sales if responsible_sales.strip() else "未指定",
                        "LABEL_IMAGE": img_obj  # 塞入圖片物件
                    }
                    
                    doc_sop.render(sop_context)
                    bio_sop = io.BytesIO()
                    doc_sop.save(bio_sop)
                    
                    # 檔名格式：CUSTOMER ID SOP + 建立日期
                    file_name_sop = f"{c_id} SOP {today_sop_str}.docx"
                    
                    st.success(f"🎉 {file_name_sop} 生成成功！")
                    st.download_button(
                        label="📥 下載 SHIPMENT SOP Word 檔",
                        data=bio_sop.getvalue(),
                        file_name=file_name_sop,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                except Exception as e:
                    st.error(f"Word 導出失敗，請確認是否已將 sop_template.docx 上傳至 GitHub。錯誤訊息: {e}")
            else: st.error("請先輸入 CUSTOMER ID 才能導出！")
