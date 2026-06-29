import streamlit as st
import pandas as pd
import gspread
import json
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from datetime import datetime
import io
import os
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
ws_sop = sh.worksheet("shipment_sops")

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

if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "auto_senko" not in st.session_state: st.session_state.auto_senko = ""
if "auto_po" not in st.session_state: st.session_state.auto_po = ""

tab1, tab2, tab3 = st.tabs(["🗂️ 產品資料庫管理", "📄 生成 COC 報表", "📦 生成 SHIPMENT SOP"])

# --- 分頁 1：產品資料庫管理 ---
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
                else: st.error("密碼錯誤！")
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
                df_m, df_s = get_main_df(), get_sub_df()
                is_exist = not df_m.empty and senko_pn in df_m["senko_pn"].values
                if is_exist and not overwrite_confirm:
                    st.error(f"🛑 發現重複！")
                else:
                    if is_exist: df_m.loc[df_m["senko_pn"] == senko_pn, ["description", "country"]] = [description, country]
                    else: df_m = pd.concat([df_m, pd.DataFrame([{"senko_pn": senko_pn, "description": description, "country": country}])], ignore_index=True)
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
            if search_term.strip(): st.table(df_main[df_main["senko_pn"].astype(str).str.contains(search_term.strip(), case=False, na=False)])
            else: st.table(df_main)

# --- 分頁 2：生成 COC 報表 ---
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

# --- 分頁 3：生成 SHIPMENT SOP ---
with tab3:
    st.subheader("📦 客戶出貨標準作業程序 (SHIPMENT SOP) 管理")
    df_sop = get_sop_df()
    c_id = st.text_input("1. 🔑 請輸入 CUSTOMER ID (大寫)", "").strip().upper()
    
    ship_type_radio = "單一快遞"
    old_ship_method = "FEDEX"
    old_acc_no, old_tax_no, old_fw_name, old_fw_email, old_fw_tel, old_custom_text = "", "", "", "", "", ""
    
    # 預設為 (未指定) 狀態
    old_sales, old_main_c, old_backup_c = "(未指定)", "(未指定)", "(未指定)"
    
    exist_row = df_sop[df_sop["customer_id"].astype(str).str.upper() == c_id].iloc[0] if (not df_sop.empty and c_id in df_sop["customer_id"].astype(str).str.upper().values) else None
    
    if exist_row is not None:
        st.success(f"ℹ️ 偵測到已存在「{c_id}」的舊資料，下方已自動帶出，修改後點儲存即可覆蓋。")
        old_docs = [d.strip() for d in str(exist_row["required_docs"]).split(",") if d.strip()]
        old_labels = [l.strip() for l in str(exist_row["label_formats"]).split(",") if l.strip()]
        old_notes = str(exist_row["shipping_notes"])
        
        # 讀取人員，若是空的就填上 "(未指定)"
        m_c = str(exist_row["main_contact"]).strip()
        b_c = str(exist_row["backup_contact"]).strip()
        s_c = str(exist_row["responsible_sales"]).strip()
        old_main_c = m_c if m_c else "(未指定)"
        old_backup_c = b_c if b_c else "(未指定)"
        old_sales = s_c if s_c else "(未指定)"
        
        old_method_str = str(exist_row["shipping_method_info"]).strip()
        if old_method_str.startswith("出貨管道: "):
            lines = old_method_str.split('\n')
            parsed_method = lines[0].replace("出貨管道: ", "").strip()
            if parsed_method == "貨代自取":
                ship_type_radio = "貨代自取"
            else:
                ship_type_radio = "單一快遞"
                old_ship_method = parsed_method
                
            for line in lines[1:]:
                if line.startswith("運費帳號: "): old_acc_no = line.replace("運費帳號: ", "").strip()
                elif line.startswith("稅金帳號: "): old_tax_no = line.replace("稅金帳號: ", "").strip()
                elif line.startswith("聯絡人: "): old_fw_name = line.replace("聯絡人: ", "").strip()
                elif line.startswith("EMAIL: "): old_fw_email = line.replace("EMAIL: ", "").strip()
                elif line.startswith("TEL: "): old_fw_tel = line.replace("TEL: ", "").strip()
        elif old_method_str != "":
            ship_type_radio = "多種條件 / 待備貨後決定 (自訂)"
            old_custom_text = old_method_str
    else:
        old_docs, old_labels, old_notes = [], [], ""

    st.divider()
    col_sop1, col_sop2 = st.columns(2)
    
    with col_sop1:
        st.markdown("**2. 出貨需要的相關文件 (可多選)**")
        doc_options = ["INVOICE", "PI", "PACKING LIST", "CO", "NCV"]
        selected_docs = [doc for doc in doc_options if st.checkbox(doc, value=(doc in old_docs), key=f"doc_{doc}")]
        
        st.markdown("<br>**3. LABEL FORMAT 外箱標籤格式 (可多選)**", unsafe_allow_html=True)
        label_options = ["DESC", "PN", "QTY", "CUSTOMER PO", "CUSTOMER PN", "COO", "CARTON NO"]
        selected_labels = [lbl for lbl in label_options if st.checkbox(lbl, value=(lbl in old_labels), key=f"lbl_{lbl}")]
        
        uploaded_label_imgs = st.file_uploader("🖼️ 上傳 LABEL 標籤參考圖 (支援一次拖拉多張 / PNG / JPG)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
        if uploaded_label_imgs:
            for img in uploaded_label_imgs:
                st.image(img, caption=img.name, width=200) 
        
    with col_sop2:
        shipping_notes = st.text_area("4. 📝 出貨注意事項 (可直接貼上)", value=old_notes, height=120)
        
        uploaded_notes_imgs = st.file_uploader("🖼️ 上傳注意事項輔助圖 (支援一次拖拉多張 / 例如打板方式)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
        if uploaded_notes_imgs:
            for img in uploaded_notes_imgs:
                st.image(img, caption=img.name, width=200)
            
        st.markdown("**5. 出貨方式設定**")
        
        ship_type = st.radio("選擇出貨情境", ["單一快遞", "貨代自取", "多種條件 / 待備貨後決定 (自訂)"], 
                             index=["單一快遞", "貨代自取", "多種條件 / 待備貨後決定 (自訂)"].index(ship_type_radio), horizontal=True)
        
        if ship_type == "單一快遞":
            method_list = ["FEDEX", "UPS", "DHL", "SF"]
            default_index = method_list.index(old_ship_method) if old_ship_method in method_list else 0
            ship_method = st.selectbox("快遞公司", method_list, index=default_index)
            acc_no = st.text_input("運費付款 ACCOUNT", value=old_acc_no)
            tax_no = st.text_input("TAX 付款 ACCOUNT", value=old_tax_no)
            method_info_str = f"出貨管道: {ship_method}\n運費帳號: {acc_no}\n稅金帳號: {tax_no}"
            
        elif ship_type == "貨代自取":
            fw_name = st.text_input("貨代聯絡人姓名", value=old_fw_name)
            fw_email = st.text_input("貨代 EMAIL", value=old_fw_email)
            fw_tel = st.text_input("貨代 TEL", value=old_fw_tel)
            method_info_str = f"出貨管道: 貨代自取\n聯絡人: {fw_name}\nEMAIL: {fw_email}\nTEL: {fw_tel}"
            
        else:
            custom_note = st.text_area("✍️ 請描述詳細出貨條件 (例如: 45kg以下用Fedex帳號1234，超重待通知 / 待提供重量尺寸後決定)", value=old_custom_text, height=150)
            method_info_str = custom_note

    st.divider()
    st.markdown("**6. 內部與客戶對接窗口人員 (支援自動記憶)**")
    
    # ✨ 核心升級：自動收集曾輸入過的所有人員名單
    all_staff = set()
    if not df_sop.empty:
        for col in ["main_contact", "backup_contact", "responsible_sales"]:
            if col in df_sop.columns:
                for val in df_sop[col].dropna().astype(str).unique():
                    if val.strip() and val.strip() != "(未指定)":
                        all_staff.add(val.strip())
                        
    staff_options = ["(未指定)"] + sorted(list(all_staff)) + ["➕ 手動輸入新人員..."]
    
    col_p1, col_p2, col_p3 = st.columns(3)
    
    # 共用的選單產生器函數
    def render_staff_select(label, old_val, key_prefix):
        default_idx = staff_options.index(old_val) if old_val in staff_options else 0
        if old_val and old_val not in staff_options:
            default_idx = len(staff_options) - 1 # 指向「手動輸入新人員...」
            
        selected = st.selectbox(label, staff_options, index=default_idx, key=f"sel_{key_prefix}")
        
        if selected == "➕ 手動輸入新人員...":
            # 如果選了新增，就彈出一個文字框給他自己打字
            return st.text_input(f"✍️ 輸入新的 {label}", value=old_val if old_val not in staff_options else "", key=f"txt_{key_prefix}")
        elif selected == "(未指定)":
            return ""
        else:
            return selected

    with col_p1: main_contact = render_staff_select("CUSTOMER 主要負責同事", old_main_c, "main")
    with col_p2: backup_contact = render_staff_select("BACKUP 同事名稱", old_backup_c, "backup")
    with col_p3: responsible_sales = render_staff_select("CUSTOMER 主要負責 SALES", old_sales, "sales")

    st.markdown("<br>", unsafe_allow_html=True)
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("💾 儲存/更新此客戶 SOP 至雲端資料庫"):
            if c_id:
                with st.spinner('正在同步...'):
                    df_sop_current = get_sop_df()
                    new_sop_data = {"customer_id": c_id, "required_docs": ", ".join(selected_docs), "label_formats": ", ".join(selected_labels), "shipping_notes": shipping_notes, "shipping_method_info": method_info_str, "main_contact": main_contact, "backup_contact": backup_contact, "responsible_sales": responsible_sales}
                    if not df_sop_current.empty and c_id in df_sop_current["customer_id"].astype(str).str.upper().values:
                        df_sop_current.loc[df_sop_current["customer_id"].astype(str).str.upper() == c_id, list(new_sop_data.keys())] = list(new_sop_data.values())
                        st.toast("🔄 舊資料更新成功！")
                    else:
                        df_sop_current = pd.concat([df_sop_current, pd.DataFrame([new_sop_data])], ignore_index=True)
                        st.toast("✅ 新資料建立成功！")
                    write_df(ws_sop, df_sop_current)
                    st.success(f"🎉 「{c_id}」出貨 SOP 已儲存！")
            else: st.error("請先輸入 CUSTOMER ID！")
            
    with col_btn2:
        if st.button("🖨️ 導出此客戶之 Word 規範檔"):
            if c_id:
                try:
                    doc_sop = DocxTemplate("sop_template.docx")
                    today_sop_str = datetime.now().strftime("%d-%b-%Y").upper()
                    
                    label_img_objs = []
                    if uploaded_label_imgs:
                        for img in uploaded_label_imgs:
                            label_img_objs.append(InlineImage(doc_sop, io.BytesIO(img.getvalue()), width=Inches(3.8)))
                    
                    notes_img_objs = []
                    if uploaded_notes_imgs:
                        for img in uploaded_notes_imgs:
                            notes_img_objs.append(InlineImage(doc_sop, io.BytesIO(img.getvalue()), width=Inches(3.8)))
                    
                    if selected_labels:
                        vert_labels = []
                        for lbl in selected_labels:
                            display_lbl = lbl.replace("PN", "P/N").replace("DESC", "Desc")
                            vert_labels.append(f"{display_lbl}:")
                        formatted_labels = "\n".join(vert_labels)
                    else:
                        formatted_labels = "無特別要求"
                    
                    sop_context = {
                        "CUSTOMER_ID": c_id,
                        "DATE": today_sop_str,
                        "REQUIRED_DOCS": ", ".join(selected_docs) if selected_docs else "無特別要求",
                        "LABEL_FORMATS": formatted_labels, 
                        "SHIPPING_NOTES": shipping_notes if shipping_notes.strip() else "無",
                        "SHIPPING_METHOD_INFO": method_info_str,
                        "MAIN_CONTACT": main_contact if main_contact.strip() else "未指定",
                        "BACKUP_CONTACT": backup_contact if backup_contact.strip() else "未指定",
                        "RESPONSIBLE_SALES": responsible_sales if responsible_sales.strip() else "未指定",
                        "LABEL_IMAGES": label_img_objs,  
                        "NOTES_IMAGES": notes_img_objs   
                    }
                    
                    doc_sop.render(sop_context)
                    bio_sop = io.BytesIO()
                    doc_sop.save(bio_sop)
                    
                    file_name_sop = f"{c_id} SOP {today_sop_str}.docx"
                    st.success(f"🎉 {file_name_sop} 生成成功！")
                    st.download_button("📥 下載 SHIPMENT SOP Word 檔", data=bio_sop.getvalue(), file_name=file_name_sop, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e: st.error(f"Word 導出失敗: {e}")
            else: st.error("請先輸入 CUSTOMER ID！")
