import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account
import io
import time
import base64
import re
from PIL import Image
import urllib.parse  
import requests  
from fpdf import FPDF  

# ==========================================
# 0. 基本設定與 CSS 美化
# ==========================================
st.set_page_config(page_title="學習扶助系統", layout="wide", page_icon="📚")

st.markdown("""
    <style>
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }
    .sys-header { display: flex; flex-direction: column; align-items: center; margin-top: 0.5rem; margin-bottom: 0.5rem; }
    .sys-logo { border-radius: 15px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); margin-bottom: 12px; }
    .sys-title { font-size: 2.2rem !important; font-weight: 700 !important; color: #2c3e50; margin: 0 !important; padding: 0 !important; }
    .section-heading { text-align: center; font-size: 1.3rem; font-weight: 600; color: #34495e; margin-top: 1.5rem; margin-bottom: 0.8rem; }
    .news-card { background-color: #fffdf5; padding: 1.2rem 1.5rem; border-radius: 8px; border-left: 5px solid #ffc107; margin-bottom: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
    </style>
""", unsafe_allow_html=True)

SHEET_URL = "https://docs.google.com/spreadsheets/d/1D1YXEMnsa58p0RKHgaOyDFo3290d03hz_JY_bJrbcpM/edit"
FOLDER_ID = "0AJdsyYPdSxydUk9PVA"
GAS_MAIL_URL = "https://script.google.com/macros/s/AKfycbxIQXJ-QENFEZn6IgyQLdESmdPMMj4NTKF1O4FlE4WhAPeBanbYFMoprVfiqo4c2Wt4/exec"

# ==========================================
# 1. 核心工具函數
# ==========================================
def apply_bg_color(subject):
    bg_color = "#FFFFFF"
    if subject == "數學": bg_color = "#E6F3FF"
    elif subject == "英語": bg_color = "#FFE6E6"
    st.markdown(f"""
        <style>
        .stApp {{ background-color: {bg_color} !important; transition: background-color 0.4s ease; }}
        </style>
    """, unsafe_allow_html=True)

def get_drive_service():
    creds_info = st.secrets["connections"]["gsheets"]
    creds = service_account.Credentials.from_service_account_info(creds_info, scopes=["https://www.googleapis.com/auth/drive"])
    return build('drive', 'v3', credentials=creds)

def upload_to_drive(file, filename):
    try:
        service = get_drive_service()
        file_metadata = {'name': filename, 'parents': [FOLDER_ID]}
        media = MediaIoBaseUpload(io.BytesIO(file.getvalue()), mimetype=file.type, resumable=True)
        uploaded_file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
        return uploaded_file.get('webViewLink')
    except Exception as e:
        st.error(f"❌ 雲端硬碟上傳失敗：{e}")
        return None

def process_image_to_base64(uploaded_file):
    try:
        img = Image.open(uploaded_file)
        if img.mode != 'RGB': img = img.convert('RGB')
        img.thumbnail((150, 150)) 
        buffered = io.BytesIO()
        img.save(buffered, format="JPEG", quality=85)
        img_str = base64.b64encode(buffered.getvalue()).decode()
        return f"data:image/jpeg;base64,{img_str}"
    except Exception as e:
        st.error(f"圖片處理失敗: {e}")
        return ""

def auto_create_sheet(sheet_url, new_sheet_name):
    """檢查並自動新增 Google Sheets 分頁"""
    try:
        sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheet_url)
        if not sheet_id_match: return False
        sheet_id = sheet_id_match.group(1)
        
        creds_info = st.secrets["connections"]["gsheets"]
        creds = service_account.Credentials.from_service_account_info(
            creds_info, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )
        service = build('sheets', 'v4', credentials=creds)
        
        sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        sheet_names = [s['properties']['title'] for s in sheets]
        
        if new_sheet_name not in sheet_names:
            request_body = {
                'requests': [{'addSheet': {'properties': {'title': new_sheet_name}}}]
            }
            service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=request_body).execute()
            time.sleep(2) # 稍微等待 Google 建立完成
        return True
    except Exception as e:
        st.error(f"自動建立資料表 {new_sheet_name} 失敗：{e}")
        return False

class CustomPDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('CustomFont', size=9)
        self.cell(0, 10, f"共 {{nb}} 頁，本頁是第 {self.page_no()} 頁", align='C')

def generate_report_pdf(df, title_name, sys_name):
    pdf = CustomPDF()
    try: pdf.add_font('CustomFont', '', 'font.ttf')
    except:
        st.error("⚠️ 找不到字型檔 font.ttf，請確認檔案位置。")
        return None
    pdf.add_page()
    pdf.alias_nb_pages()

    pdf.set_font('CustomFont', size=18)
    pdf.cell(0, 15, f"{sys_name}", ln=True, align='C')
    pdf.set_font('CustomFont', size=14)
    pdf.cell(0, 10, f"報表名稱：{title_name}", ln=True, align='C')
    
    today_str = pd.Timestamp.today().strftime("%Y/%m/%d")
    pdf.set_font('CustomFont', size=10)
    pdf.cell(0, 10, f"統計日期：{today_str}", ln=True, align='C')
    pdf.ln(2)

    pdf.set_font('CustomFont', size=9)
    col_widths = [pdf.get_string_width(str(col)) + 4 for col in df.columns]
    for _, row in df.iterrows():
        for i, item in enumerate(row):
            w = pdf.get_string_width(str(item)) + 4
            if w > col_widths[i]: col_widths[i] = w

    row_height = pdf.font_size + 2 
    total_table_width = sum(col_widths)
    start_x = (pdf.w - total_table_width) / 2
    
    pdf.set_fill_color(235, 235, 235) 
    pdf.set_x(start_x)
    for i, col in enumerate(df.columns):
        pdf.cell(col_widths[i], row_height, str(col), border=1, fill=True, align='C')
    pdf.ln()

    for _, row in df.iterrows():
        if pdf.get_y() > 260: pdf.add_page()
        pdf.set_x(start_x)
        for i, item in enumerate(row):
            pdf.cell(col_widths[i], row_height, str(item), border=1, align='C')
        pdf.ln()

    pdf.ln(10)
    pdf.set_font('CustomFont', size=11)
    pdf.cell(0, 10, "承辦人簽章：__________________________", ln=True, align='R')
    pdf.set_font('CustomFont', size=9)
    pdf.cell(0, 5, f"列印日期：{today_str}", ln=True, align='R')
    
    return pdf.output()

def get_auto_school_year():
    """根據系統當前時間，自動判斷並回傳所屬的學年度（不分上下學期）"""
    today = pd.Timestamp.today()
    tw_yr = today.year - 1911
    # 台灣學制：8月1日以後為新學年度
    if today.month >= 8:
        return str(tw_yr)
    else:
        return str(tw_yr - 1)

conn = st.connection("gsheets", type=GSheetsConnection)

# ==========================================
# 2. 讀取系統設定與狀態初始化
# ==========================================
try:
    df_settings = conn.read(spreadsheet=SHEET_URL, worksheet="Settings", ttl=600).dropna(how="all")
    sys_name = df_settings.loc[df_settings['設定項'] == 'SchoolName', '設定值'].values[0] if 'SchoolName' in df_settings['設定項'].values else "🏫 學習扶助系統"
    sys_logo = df_settings.loc[df_settings['設定項'] == 'LogoLink', '設定值'].values[0] if 'LogoLink' in df_settings['設定項'].values else ""
    sys_year = df_settings.loc[df_settings['設定項'] == 'SchoolYear', '設定值'].values[0] if 'SchoolYear' in df_settings['設定項'].values else "114上" 
    sys_date = df_settings.loc[df_settings['設定項'] == 'DiagDate', '設定值'].values[0] if 'DiagDate' in df_settings['設定項'].values else pd.Timestamp.today().strftime("%Y/%m/%d")
except Exception:
    sys_name = "🏫 學習扶助系統"
    sys_logo = ""
    sys_year = "114上"
    sys_date = pd.Timestamp.today().strftime("%Y/%m/%d")

try:
    df_news = conn.read(spreadsheet=SHEET_URL, worksheet="News", ttl=600).dropna(how="all")
except Exception:
    df_news = pd.DataFrame(columns=["日期", "標題", "內容"])

if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "user_name" not in st.session_state: st.session_state.user_name = ""
if "user_role" not in st.session_state: st.session_state.user_role = "教師"
if "edit_record_idx" not in st.session_state: st.session_state.edit_record_idx = None
if "record_form_id" not in st.session_state: st.session_state.record_form_id = 0

# ==========================================
# 畫面區塊 1：登入介面
# ==========================================
if not st.session_state.logged_in:
    apply_bg_color("純白") 
    col_space_left, main_col, col_space_right = st.columns([1, 8, 1])
    with main_col:
        clean_sys_name = sys_name.replace("🏫", "").strip() 
        if sys_logo and sys_logo.startswith("data:image"):
            st.markdown(f"""
                <div class="sys-header">
                    <img src="{sys_logo}" width="100" class="sys-logo">
                    <h1 class="sys-title">{clean_sys_name}</h1>
                </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='sys-header'><h1 class='sys-title'>{clean_sys_name}</h1></div>", unsafe_allow_html=True)
        
        st.markdown("<div class='section-heading'>📢 最新消息</div>", unsafe_allow_html=True)
        _, col_news_center, _ = st.columns([1, 2, 1]) 
        with col_news_center:
            if df_news.empty:
                st.info("目前尚無公告。")
            else:
                for _, n in df_news.sort_values("日期", ascending=False).head(5).iterrows():
                    st.markdown(f"""
                        <div class="news-card">
                            <strong style="font-size: 1.05em; color: #2c3e50;">【{n.get('日期', '')}】 {n.get('標題', '')}</strong><br>
                            <span style="color: #555; font-size: 0.95em;">{n.get('內容', '')}</span>
                        </div>
                    """, unsafe_allow_html=True)
                
        st.markdown("<div class='section-heading'>🔐 使用者登入</div>", unsafe_allow_html=True)
        _, col_login_form, _ = st.columns([1, 1, 1])
        with col_login_form:
            with st.form("login_form"):
                login_account = st.text_input("帳號")
                login_password = st.text_input("密碼", type="password")
                submit_login = st.form_submit_button("登入系統", use_container_width=True)
                
                if submit_login:
                    try:
                        df_users = conn.read(spreadsheet=SHEET_URL, worksheet="Users", ttl=0).dropna(how="all")
                        df_users['密碼'] = df_users['密碼'].astype(str)
                        user_match = df_users[(df_users['帳號'] == login_account) & (df_users['密碼'] == str(login_password))]
                        
                        if not user_match.empty:
                            user_status = user_match.iloc[0].get('狀態', '啟用')
                            if user_status == '停用':
                                st.error("❌ 此帳號已被停用，請聯絡系統管理員。")
                            else:
                                st.session_state.logged_in = True
                                st.session_state.user_name = user_match.iloc[0]['姓名']
                                st.session_state.user_role = user_match.iloc[0].get('權限', '教師')
                                st.balloons()
                                st.toast(f"歡迎回來，{st.session_state.user_name}！", icon="👋")
                                time.sleep(1)
                                st.rerun()
                        else:
                            st.error("❌ 帳號或密碼錯誤")
                    except Exception as e:
                        st.error(f"連線失敗：{e}")

# ==========================================
# 畫面區塊 2：系統主畫面
# ==========================================
else:
    col_logo, col_title, col_btn = st.columns([1, 8, 2])
    with col_logo:
        if sys_logo and sys_logo.startswith("data:image"): st.image(sys_logo, width=60)
    with col_title:
        clean_sys_name = sys_name.replace("🏫", "").strip() 
        st.title(clean_sys_name)
        st.write(f"👤 登入者：**{st.session_state.user_name}** | 🛡️ 權限：**{st.session_state.user_role}**")
    with col_btn:
        st.write("") 
        if st.button("🚪 登出系統", use_container_width=True):
            st.cache_data.clear()
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun() 

    st.markdown("---")

    try:
        df_users = conn.read(spreadsheet=SHEET_URL, worksheet="Users", ttl=600).dropna(how="all")
        if '狀態' not in df_users.columns: df_users['狀態'] = '啟用'
            
        df_students = conn.read(spreadsheet=SHEET_URL, worksheet="Students", ttl=600).dropna(how="all")
        df_records = conn.read(spreadsheet=SHEET_URL, worksheet="Records", ttl=600).dropna(how="all")

        tabs_list = ["✍️ 新增輔導紀錄", "🗂️ 個人紀錄回顧", "🕰️ 歷史紀錄查詢", "⚙️ 個人設定"]
        if st.session_state.user_role == "管理者":
            tabs_list.extend(["📊 學生報表", "👨‍🏫 教師報表", "⚙️ 系統設定"])
        
        tabs = st.tabs(tabs_list)

        # ==================================================
        # 分頁 1：✍️ 新增輔導紀錄
        # ==================================================
        with tabs[0]:
            teacher_students = []
            for _, row in df_students.iterrows():
                for sub in ['國語', '英語', '數學']:
                    if str(row.get(f'{sub}_教學者', '')).strip() == st.session_state.user_name:
                        teacher_students.append(f"{row['學生姓名']} - {sub}")
            
            if not teacher_students:
                st.info("您目前沒有被指派負責的學生。")
                apply_bg_color("預設") 
            else:
                st.markdown("**✅ 請勾選本次輔導的對象 (可多選)：**")
                selected_st_subs = []
                chk_cols = st.columns(4)
                for i, ts in enumerate(teacher_students):
                    if chk_cols[i % 4].checkbox(ts, key=f"add_chk_{ts}_{st.session_state.record_form_id}"):
                        selected_st_subs.append(ts)
                
                if st.session_state.edit_record_idx is None:
                    if selected_st_subs: apply_bg_color(selected_st_subs[0].split(" - ")[1])
                    else: apply_bg_color("預設")
                        
                with st.form(key="record_form", clear_on_submit=True):
                    c1, c2 = st.columns(2)
                    with c1:
                        counsel_term = get_auto_school_year()
                        st.info(f"⏳ **(系統帶入) 教學輔導學年度：** {counsel_term}")
                        st.info(f"📋 **(系統帶入) 診斷日期：** {sys_date}")
                        counsel_date = st.date_input("🗓️ 實際輔導日期", value=pd.Timestamp.today().date())
                    with c2:
                        summary = st.text_input("教學摘要 (代碼)")
                        strategy = st.multiselect("輔導策略", ["演示教學", "學習單、試題習寫", "實物操作", "數位學習平台"])
                        status = st.radio("學習狀況", ["精熟", "未精熟，已解說", "未精熟，改日再測"])
                    p_link = st.text_input("佐證網址 (選填)")
                    u_files = st.file_uploader("上傳佐證檔案 (最多三個)", type=["png", "jpg", "jpeg", "pdf", "docx"], accept_multiple_files=True)
                    
                    if st.form_submit_button("💾 儲存紀錄"):
                        if not selected_st_subs: st.warning("⚠️ 請至少勾選一位輔導對象！")
                        elif not summary: st.warning("⚠️ 教學摘要為必填！")
                        elif not p_link and not u_files: st.warning("⚠️ 必須填寫「佐證網址」或上傳「佐證檔案」！")
                        else:
                            with st.spinner('上傳資料中...'):
                                p_entries = []
                                if p_link: p_entries.append(f"連結: {p_link}")
                                if u_files:
                                    for f in u_files[:3]:
                                        drive_link = upload_to_drive(f, f.name)
                                        if drive_link: p_entries.append(f"檔案: {drive_link}")
                                
                                new_records_list = []
                                for st_sub in selected_st_subs:
                                    sel_stu, sel_sub = st_sub.split(" - ")
                                    stu_score = 0
                                    stu_data = df_students[df_students['學生姓名'] == sel_stu]
                                    if not stu_data.empty:
                                        score_col_name = f"{sel_sub}_成績"
                                        if score_col_name in stu_data.columns:
                                            val = stu_data.iloc[0].get(score_col_name, 0)
                                            if pd.notna(val): stu_score = val
                                            
                                    new_records_list.append({
                                        "學年度": counsel_term, "授課科目": sel_sub, "學生姓名": sel_stu,
                                        "篩選測驗成績": stu_score, "教學內容摘要": summary, "教學輔導策略": ", ".join(strategy),
                                        "學習狀況": status, "診斷日期": sys_date, "輔導日期": counsel_date.strftime("%Y/%m/%d"),
                                        "佐證資料連結": " / ".join(p_entries)
                                    })
                                
                                new_rec_df = pd.DataFrame(new_records_list)
                                updated_df = pd.concat([df_records, new_rec_df], ignore_index=True)
                                conn.update(spreadsheet=SHEET_URL, worksheet="Records", data=updated_df)
                                st.session_state.record_form_id += 1 
                                st.cache_data.clear()
                                st.toast(f'✅ 已成功儲存 {len(new_records_list)} 筆紀錄！', icon='🎉')
                                time.sleep(1.5) ; st.rerun()

        # ==================================================
        # 分頁 2：🗂️ 個人紀錄回顧
        # ==================================================
        with tabs[1]:
            if st.session_state.edit_record_idx is None:
                st.subheader("🗂️ 您的教學紀錄")
                my_records = df_records[df_records['學生姓名'].isin([s.split(" - ")[0] for s in teacher_students])]
                if my_records.empty: st.info("尚無相關紀錄")
                else:
                    for idx, row in my_records.sort_values("輔導日期", ascending=False).iterrows():
                        with st.expander(f"📅 {row.get('輔導日期', row.get('診斷日期', ''))} | {row['學生姓名']} ({row['授課科目']}) - {row['教學內容摘要']}"):
                            st.markdown(f"- **教學輔導時間：** {row.get('學年度', '未註記')}\n- **輔導日期：** {row.get('輔導日期', '未註記')}\n- **診斷日期：** {row['診斷日期']}\n- **篩選成績：** {row['篩選測驗成績']}\n- **策略：** {row['教學輔導策略']}\n- **狀況：** {row['學習狀況']}")
                            p_data = str(row['佐證資料連結'])
                            if p_data and p_data != "nan":
                                st.markdown("---")
                                btns = []
                                for p in p_data.split(" / "):
                                    if "連結: " in p: btns.append(f'🔗 <a href="{p.replace("連結: ", "")}" target="_blank">[網址]</a>')
                                    if "檔案: " in p: btns.append(f'📁 <a href="{p.replace("檔案: ", "")}" target="_blank">[檔案]</a>')
                                st.markdown(" ".join(btns), unsafe_allow_html=True)
                            
                            st.markdown("---")
                            col_space, col_edit, col_del = st.columns([6, 1, 1])
                            with col_edit:
                                if st.button("✏️ 編輯", key=f"btn_edit_{idx}"):
                                    st.session_state.edit_record_idx = idx ; st.rerun()
                            with col_del:
                                if st.button("🗑️ 刪除", key=f"btn_del_{idx}"):
                                    df_records = df_records.drop(idx)
                                    conn.update(spreadsheet=SHEET_URL, worksheet="Records", data=df_records)
                                    st.cache_data.clear() ; st.toast("✅ 紀錄已成功刪除！", icon="🗑️")
                                    time.sleep(1) ; st.rerun()
            else:
                edit_idx = st.session_state.edit_record_idx
                if edit_idx not in df_records.index:
                    st.session_state.edit_record_idx = None ; st.rerun()
                edit_row = df_records.loc[edit_idx]
                edit_sub = edit_row['授課科目']
                apply_bg_color(edit_sub)
                
                st.subheader(f"✏️ 編輯紀錄：{edit_row['學生姓名']} - {edit_sub}")
                with st.form("edit_record_form", clear_on_submit=False):
                    c1, c2 = st.columns(2)
                    with c1:
                        old_term = str(edit_row.get('學年度', get_auto_school_year()))
                        st.info(f"⏳ **(原紀錄) 學年度：** {old_term}")
                        new_counsel_term = old_term
                        
                        st.info(f"📊 **篩選測驗成績：** {edit_row['篩選測驗成績']}")
                        st.info(f"📋 **診斷日期：** {edit_row['診斷日期']}")
                        try: default_c_date = pd.to_datetime(edit_row.get('輔導日期')).date()
                        except: default_c_date = pd.Timestamp.today().date()
                        new_counsel_date = st.date_input("🗓️ 實際輔導日期", value=default_c_date)
                    with c2:
                        new_summary = st.text_input("教學摘要 (代碼)", value=str(edit_row['教學內容摘要']))
                        old_strat_str = str(edit_row['教學輔導策略'])
                        old_strats = [s.strip() for s in old_strat_str.split(",")] if old_strat_str and old_strat_str != "nan" else []
                        available_strats = ["演示教學", "學習單、試題習寫", "實物操作", "數位學習平台"]
                        valid_old_strats = [s for s in old_strats if s in available_strats]
                        new_strategy = st.multiselect("輔導策略", available_strats, default=valid_old_strats)
                        old_status = str(edit_row['學習狀況'])
                        status_options = ["精熟", "未精熟，已解說", "未精熟，改日再測"]
                        default_status_idx = status_options.index(old_status) if old_status in status_options else 0
                        new_status = st.radio("學習狀況", status_options, index=default_status_idx)
                    
                    old_p_data = str(edit_row['佐證資料連結'])
                    st.markdown(f"📂 **目前佐證資料：** `{old_p_data if old_p_data and old_p_data != 'nan' else '無'}`")
                    st.info("💡 【提醒您】：如果不更動佐證資料，下方兩個欄位請保持空白即可，系統會為您保留舊有資料。")
                    new_p_link = st.text_input("新佐證網址 (若不更改請留白)")
                    new_u_files = st.file_uploader("上傳新佐證檔案 (若不更改請留白，最多三個)", type=["png", "jpg", "jpeg", "pdf", "docx"], accept_multiple_files=True)
                    
                    c_save, c_cancel = st.columns([1, 1])
                    with c_save: save_btn = st.form_submit_button("💾 儲存修改", use_container_width=True)
                    with c_cancel: cancel_btn = st.form_submit_button("❌ 取消編輯", use_container_width=True)
                        
                    if cancel_btn: st.session_state.edit_record_idx = None ; st.rerun()
                        
                    if save_btn:
                        if not new_summary: st.warning("⚠️ 教學摘要為必填！")
                        else:
                            with st.spinner("更新資料中..."):
                                final_p_entries = []
                                if new_p_link or new_u_files:
                                    if new_p_link: final_p_entries.append(f"連結: {new_p_link}")
                                    if new_u_files:
                                        for f in new_u_files[:3]:
                                            drive_link = upload_to_drive(f, f.name)
                                            if drive_link: final_p_entries.append(f"檔案: {drive_link}")
                                else:
                                    if old_p_data and old_p_data != 'nan': final_p_entries = old_p_data.split(" / ")
                                
                                df_records.at[edit_idx, '學年度'] = new_counsel_term
                                df_records.at[edit_idx, '教學內容摘要'] = new_summary
                                df_records.at[edit_idx, '教學輔導策略'] = ", ".join(new_strategy)
                                df_records.at[edit_idx, '學習狀況'] = new_status
                                df_records.at[edit_idx, '輔導日期'] = new_counsel_date.strftime("%Y/%m/%d")
                                df_records.at[edit_idx, '佐證資料連結'] = " / ".join(final_p_entries)
                                conn.update(spreadsheet=SHEET_URL, worksheet="Records", data=df_records)
                                st.cache_data.clear()
                                st.session_state.edit_record_idx = None
                                st.toast('✅ 紀錄已成功更新！', icon='🎉')
                                time.sleep(1.5) ; st.rerun()

        # ==================================================
        # 分頁 3：🕰️ 歷史紀錄查詢
        # ==================================================
        with tabs[2]:
            st.subheader("🕰️ 歷史紀錄查詢")
            st.write("系統會自動讀取已封存的歷史紀錄，請從下拉式選單選擇要查詢的資料表。")
            
            record_sheets = []
            try:
                sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', SHEET_URL)
                if sheet_id_match:
                    sheet_id = sheet_id_match.group(1)
                    creds_info = st.secrets["connections"]["gsheets"]
                    creds = service_account.Credentials.from_service_account_info(
                        creds_info, 
                        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
                    )
                    service = build('sheets', 'v4', credentials=creds)
                    sheet_metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
                    sheets = sheet_metadata.get('sheets', '')
                    
                    sheet_names = [s['properties']['title'] for s in sheets]
                    record_sheets = [name for name in sheet_names if name.startswith('Records_')]
            except Exception as e:
                st.warning(f"⚠️ 無法自動抓取歷史資料表清單。錯誤：{e}")
                
            if not record_sheets:
                st.info("📭 目前系統中找不到任何已封存的歷史資料表。請確認是否曾經執行過「學期轉換與資料備份」。")
            else:
                with st.form("history_search_form"):
                    selected_archive = st.selectbox("🔍 請選擇要查詢的歷史學年度資料表", record_sheets)
                    submit_search = st.form_submit_button("查詢歷史紀錄")
                    
                if submit_search:
                    with st.spinner(f"正在搜尋 {selected_archive} ..."):
                        try:
                            df_history = conn.read(spreadsheet=SHEET_URL, worksheet=selected_archive, ttl=0).dropna(how="all")
                            if df_history.empty:
                                st.info(f"📭 `{selected_archive}` 資料表存在，但目前沒有任何紀錄。")
                            else:
                                st.success(f"✅ 成功讀取 {selected_archive} 歷史資料！共找到 {len(df_history)} 筆紀錄。")
                                st.dataframe(df_history, use_container_width=True)
                        except Exception as e:
                            st.error(f"❌ 找不到名為 `{selected_archive}` 的歷史資料表，或讀取時發生錯誤。")

        # ==================================================
        # 分頁 4：⚙️ 個人帳號設定
        # ==================================================
        with tabs[3]:
            st.subheader("⚙️ 個人帳號設定")
            with st.form("change_pwd_form"):
                old_pwd = st.text_input("目前的密碼", type="password")
                new_pwd = st.text_input("新密碼", type="password")
                new_pwd_confirm = st.text_input("再次確認新密碼", type="password")
                if st.form_submit_button("💾 儲存新密碼"):
                    current_user_data = df_users[df_users['姓名'] == st.session_state.user_name]
                    if current_user_data.empty: st.error("找不到使用者資料")
                    elif str(current_user_data.iloc[0]['密碼']) != old_pwd: st.error("❌ 目前的密碼輸入錯誤！")
                    elif new_pwd != new_pwd_confirm: st.error("❌ 兩次輸入的新密碼不一致！")
                    elif len(new_pwd) < 4: st.error("❌ 為了安全，密碼請至少輸入4個字元")
                    else:
                        idx = df_users.index[df_users['姓名'] == st.session_state.user_name].tolist()[0]
                        df_users.at[idx, '密碼'] = new_pwd
                        conn.update(spreadsheet=SHEET_URL, worksheet="Users", data=df_users)
                        st.cache_data.clear()
                        st.success("✅ 密碼修改成功！請重新登入。")
                        time.sleep(2)
                        for key in list(st.session_state.keys()): del st.session_state[key]
                        st.rerun()

        # ==================================================
        # 管理者專屬分頁
        # ==================================================
        if st.session_state.user_role == "管理者":
            report_data = []
            for _, stu in df_students.iterrows():
                s_name = stu['學生姓名']
                s_class = str(stu.get('班別', '?'))
                for sub in ['國語', '英語', '數學']:
                    teacher = str(stu.get(f'{sub}_教學者', '')).strip()
                    if teacher and teacher != "nan":
                        count = len(df_records[(df_records['學生姓名'] == s_name) & (df_records['授課科目'] == sub)])
                        report_data.append({
                            "班別": s_class, "學生姓名": s_name, "科目": sub,
                            "指導教師": teacher, "紀錄筆數": count,
                            "狀態": "✅ 已填寫" if count > 0 else "❌ 未填寫"
                        })
            df_report = pd.DataFrame(report_data)
            
            with tabs[4]:
                col_h1, col_dl_csv, col_dl_pdf = st.columns([3.5, 1, 1])
                with col_h1: st.subheader("📊 學生輔導狀況報表")
                with col_dl_csv: 
                    csv_stu = df_report.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(label="📥 匯出 CSV", data=csv_stu, file_name="學生輔導狀況.csv", mime="text/csv", use_container_width=True)
                with col_dl_pdf:
                    pdf_data = generate_report_pdf(df_report, "學生輔導狀況報表", clean_sys_name)
                    if pdf_data: st.download_button(label="📥 匯出 PDF", data=bytes(pdf_data), file_name="學生輔導狀況.pdf", mime="application/pdf", use_container_width=True)
                
                classes = sorted(df_report['班別'].unique()) if not df_report.empty else []
                for cls in classes:
                    with st.expander(f"🏫 【{cls}】 學生列表"):
                        df_cls = df_report[df_report['班別'] == cls]
                        c1, c2, c3, c4, c5 = st.columns([1.5, 1, 1.5, 1.5, 1])
                        c1.markdown("**🧑‍🎓 姓名**") ; c2.markdown("**📖 科目**") ; c3.markdown("**🧑‍🏫 指導教師**")
                        c4.markdown("**📋 狀態**") ; c5.markdown("**🔍 操作**")
                        st.markdown('<div style="border-bottom: 2px solid #ccc; margin-bottom: 5px;"></div>', unsafe_allow_html=True)
                        for _, r in df_cls.iterrows():
                            c1, c2, c3, c4, c5 = st.columns([1.5, 1, 1.5, 1.5, 1])
                            c1.write(f"**{r['學生姓名']}**") ; c2.write(r['科目']) ; c3.write(r['指導教師'])
                            c4.markdown(f"{r['狀態']} ({r['紀錄筆數']}筆)" if r['紀錄筆數']>0 else "<span style='color:red;'>❌ 未填寫</span>", unsafe_allow_html=True)
                            is_expanded = False
                            with c5:
                                if r['紀錄筆數'] > 0: is_expanded = st.toggle("查看", key=f"tgl_stu_{cls}_{r['學生姓名']}_{r['科目']}")
                            if is_expanded:
                                with st.container():
                                    target_recs = df_records[(df_records['學生姓名'] == r['學生姓名']) & (df_records['授課科目'] == r['科目'])]
                                    for _, row in target_recs.sort_values("輔導日期", ascending=False).iterrows():
                                        with st.expander(f"📅 {row.get('輔導日期', row.get('診斷日期', ''))} | {row['教學內容摘要']}", expanded=True):
                                            st.markdown(f"- **策略：** {row['教學輔導策略']}\n- **狀況：** {row['學習狀況']}")
                            st.markdown('<div style="border-bottom: 1px dashed #eee; margin-bottom: 2px;"></div>', unsafe_allow_html=True)

            with tabs[5]:
                col_h2, col_dl_csv2, col_dl_pdf2 = st.columns([3.5, 1, 1])
                with col_h2: st.subheader("👨‍🏫 教師填寫進度報表")
                with col_dl_csv2: 
                    csv_tch = df_report.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(label="📥 匯出 CSV", data=csv_tch, file_name="教師填寫進度.csv", mime="text/csv", use_container_width=True, key="dl_tch_csv")
                with col_dl_pdf2:
                    df_teacher_pdf = df_report.copy().sort_values(by="指導教師")[['指導教師', '班別', '學生姓名', '科目', '紀錄筆數', '狀態']]
                    pdf_data2 = generate_report_pdf(df_teacher_pdf, "教師填寫進度報表", clean_sys_name)
                    if pdf_data2: st.download_button(label="📥 匯出 PDF", data=bytes(pdf_data2), file_name="教師填寫進度.pdf", mime="application/pdf", use_container_width=True, key="dl_tch_pdf")
                
                show_unassigned = st.checkbox("👁️ 顯示未指派學生的教師", value=True)
                teachers_list = df_users[df_users['權限'] == '教師']['姓名'].dropna().unique().tolist()
                for t_name in teachers_list:
                    df_tch = df_report[df_report['指導教師'] == t_name]
                    if df_tch.empty:
                        if show_unassigned:
                            with st.expander(f"🧑‍🏫 【{t_name} 老師】 - 無指派學生"): st.info("無指派對象")
                    else:
                        t_email = ""
                        if 'Email' in df_users.columns:
                            email_match = df_users[df_users['姓名'] == t_name]['Email'].values
                            if len(email_match) > 0 and pd.notna(email_match[0]): t_email = str(email_match[0]).strip()
                        incomplete_count = len(df_tch[df_tch['狀態'] == "❌ 未填寫"])
                        with st.expander(f"🧑‍🏫 【{t_name} 老師】 - 負責 {len(df_tch)} 項輔導"):
                            if incomplete_count > 0 and t_email:
                                subject = f"[{sys_name}] 輔導紀錄填寫提醒"
                                body = f"{t_name} 老師您好：\n\n系統顯示您尚有 {incomplete_count} 筆學生輔導紀錄尚未完成，請盡速登入系統進行填報。\n\n感謝您的辛勞與配合！"
                                if st.button(f"✉️ 寄發 Email 提醒 ({incomplete_count}筆未填)", key=f"btn_mail_{t_name}"):
                                    with st.spinner(f"正在傳送郵件..."):
                                        try:
                                            resp = requests.post(GAS_MAIL_URL, json={"email": t_email, "subject": subject, "body": body})
                                            if resp.status_code == 200 and resp.json().get("status") == "success": st.success(f"✅ 已寄出！")
                                            else: st.error(f"❌ 失敗：{resp.text}")
                                        except Exception as e: st.error(f"❌ 錯誤：{e}")
                            c1, c2, c3, c4, c5 = st.columns([1.5, 1.5, 1, 1.5, 1])
                            c1.markdown("**🏫 班別**") ; c2.markdown("**🧑‍🎓 姓名**") ; c3.markdown("**📖 科目**")
                            c4.markdown("**📋 狀態**") ; c5.markdown("**🔍 操作**")
                            st.markdown('<div style="border-bottom: 2px solid #ccc; margin-bottom: 5px;"></div>', unsafe_allow_html=True)
                            for _, r in df_tch.iterrows():
                                c1, c2, c3, c4, c5 = st.columns([1.5, 1.5, 1, 1.5, 1])
                                c1.write(r['班別']) ; c2.write(f"**{r['學生姓名']}**") ; c3.write(r['科目'])
                                c4.markdown(f"{r['狀態']} ({r['紀錄筆數']}筆)" if r['紀錄筆數']>0 else "<span style='color:red;'>❌ 未填寫</span>", unsafe_allow_html=True)
                                is_expanded = False
                                with c5:
                                    if r['紀錄筆數'] > 0: is_expanded = st.toggle("查看", key=f"tgl_tch_{t_name}_{r['班別']}_{r['學生姓名']}_{r['科目']}")
                                if is_expanded:
                                    with st.container():
                                        target_recs = df_records[(df_records['學生姓名'] == r['學生姓名']) & (df_records['授課科目'] == r['科目'])]
                                        for _, row in target_recs.sort_values("輔導日期", ascending=False).iterrows():
                                            with st.expander(f"📅 {row.get('輔導日期', row.get('診斷日期', ''))} | {row['教學內容摘要']}", expanded=True):
                                                st.markdown(f"- **策略：** {row['教學輔導策略']}\n- **狀況：** {row['學習狀況']}")
                                st.markdown('<div style="border-bottom: 1px dashed #eee; margin-bottom: 2px;"></div>', unsafe_allow_html=True)

            with tabs[6]:
                st.subheader("⚙️ 系統基本設定")
                with st.form("settings_form"):
                    new_sys_name = st.text_input("學校名稱 / 系統標題", value=sys_name)
                    new_logo_file = st.file_uploader("上傳學校 Logo", type=["png", "jpg", "jpeg"])
                    try: default_date = pd.to_datetime(sys_date).date()
                    except: default_date = pd.Timestamp.today().date()
                    new_sys_date = st.date_input("設定統一診斷日期", value=default_date)
                    if st.form_submit_button("💾 儲存系統設定"):
                        final_logo_url = sys_logo
                        if new_logo_file: final_logo_url = process_image_to_base64(new_logo_file)
                        new_settings_df = pd.DataFrame([
                            {"設定項": "SchoolName", "設定值": new_sys_name},
                            {"設定項": "LogoLink", "設定值": final_logo_url},
                            {"設定項": "SchoolYear", "設定值": sys_year}, 
                            {"設定項": "DiagDate", "設定值": new_sys_date.strftime("%Y/%m/%d")}
                        ])
                        conn.update(spreadsheet=SHEET_URL, worksheet="Settings", data=new_settings_df)
                        st.cache_data.clear() ; st.toast('已儲存！') ; time.sleep(1) ; st.rerun()
                
                # ==================================================
                # --- 學期轉換與備份管理 ---
                # ==================================================
                st.markdown("---")
                st.subheader("📦 學期轉換與資料備份管理")
                
                st.write("**1. 一鍵自動備份與封存**")
                st.info(f"💡 系統將自動建立 `{sys_year}` 分頁，一鍵備份「使用者、學生、輔導紀錄」，並**清空**當前輔導紀錄以供新學期使用。")
                if st.button("📦 備份資料並清空本學期輔導紀錄"):
                    with st.spinner("正在自動建立資料表並備份資料..."):
                        tables_to_backup = {
                            "Users": df_users,
                            "Students": df_students,
                            "Records": df_records
                        }
                        all_success = True
                        for base_name, df_data in tables_to_backup.items():
                            archive_name = f"{base_name}_{sys_year}"
                            if auto_create_sheet(SHEET_URL, archive_name):
                                try:
                                    try:
                                        df_archive = conn.read(spreadsheet=SHEET_URL, worksheet=archive_name, ttl=0).dropna(how="all")
                                        df_archive_new = pd.concat([df_archive, df_data], ignore_index=True).drop_duplicates()
                                    except:
                                        df_archive_new = df_data
                                        
                                    conn.update(spreadsheet=SHEET_URL, worksheet=archive_name, data=df_archive_new)
                                except Exception as e:
                                    st.error(f"❌ {archive_name} 備份失敗：{e}")
                                    all_success = False
                            else:
                                all_success = False
                        
                        if all_success:
                            empty_records_df = pd.DataFrame(columns=df_records.columns)
                            conn.update(spreadsheet=SHEET_URL, worksheet="Records", data=empty_records_df)
                            st.cache_data.clear()
                            st.success(f"✅ 資料已成功備份至 `{sys_year}` 歷史資料表，且輔導紀錄已清空！")
                            time.sleep(2)
                            st.rerun()

                st.write("**2. 下載名單範例檔**")
                st.info("💡 提供新學期名單的標準 CSV 範例檔供您下載填寫。")
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    empty_users_df = pd.DataFrame(columns=["帳號", "密碼", "姓名", "權限", "Email", "狀態"]) if df_users.empty else pd.DataFrame(columns=df_users.columns)
                    st.download_button("📥 下載【使用者名單】範例檔", data=empty_users_df.to_csv(index=False).encode('utf-8-sig'), file_name="使用者名單_範例.csv", mime="text/csv", use_container_width=True)
                with col_dl2:
                    empty_students_df = pd.DataFrame(columns=["班別", "學生姓名", "國語_教學者", "國語_成績", "英語_教學者", "英語_成績", "數學_教學者", "數學_成績"]) if df_students.empty else pd.DataFrame(columns=df_students.columns)
                    st.download_button("📥 下載【學生名單】範例檔", data=empty_students_df.to_csv(index=False).encode('utf-8-sig'), file_name="學生名單_範例.csv", mime="text/csv", use_container_width=True)

                st.write("**3. 匯入新學期名單**")
                st.info("💡 請上傳填妥的 CSV 檔案。上傳後將直接覆蓋目前的系統名單，請確保已完成上方的一鍵備份。")
                col_up1, col_up2 = st.columns(2)
                with col_up1:
                    uploaded_users = st.file_uploader("上傳新【使用者】名單 (CSV)", type=["csv"])
                    if uploaded_users and st.button("💾 更新使用者資料"):
                        with st.spinner("更新使用者名單中..."):
                            try:
                                new_users_df = pd.read_csv(uploaded_users)
                                conn.update(spreadsheet=SHEET_URL, worksheet="Users", data=new_users_df)
                                st.cache_data.clear()
                                st.success("✅ 使用者名單已更新！")
                                time.sleep(2); st.rerun()
                            except Exception as e:
                                st.error(f"❌ 更新失敗：{e}")
                                
                with col_up2:
                    uploaded_students = st.file_uploader("上傳新【學生】名單 (CSV)", type=["csv"])
                    if uploaded_students and st.button("💾 更新學生資料"):
                        with st.spinner("更新學生名單中..."):
                            try:
                                new_stu_df = pd.read_csv(uploaded_students)
                                conn.update(spreadsheet=SHEET_URL, worksheet="Students", data=new_stu_df)
                                st.cache_data.clear()
                                st.success("✅ 學生名單已更新！")
                                time.sleep(2); st.rerun()
                            except Exception as e:
                                st.error(f"❌ 更新失敗：{e}")

                # ==================================================
                st.markdown("---")
                
                st.subheader("👥 使用者帳號管理")
                with st.form("add_user_form", clear_on_submit=True):
                    st.write("➕ **新增使用者**")
                    c_acc, c_pwd, c_name = st.columns(3)
                    with c_acc: n_acc = st.text_input("帳號 (必填)")
                    with c_pwd: n_pwd = st.text_input("密碼 (必填)")
                    with c_name: n_name = st.text_input("姓名 (必填)")
                    
                    c_role, c_email, c_space = st.columns(3)
                    with c_role: n_role = st.selectbox("權限", ["教師", "管理者"])
                    with c_email: n_email = st.text_input("Email (選填)")
                    
                    if st.form_submit_button("新增帳號"):
                        if not n_acc or not n_pwd or not n_name:
                            st.warning("⚠️ 帳號、密碼、姓名皆為必填！")
                        elif n_acc in df_users['帳號'].values:
                            st.error("❌ 此帳號已存在，請更換帳號名稱！")
                        else:
                            new_user = pd.DataFrame([{
                                "帳號": n_acc, "密碼": n_pwd, "姓名": n_name,
                                "權限": n_role, "Email": n_email, "狀態": "啟用"
                            }])
                            updated_users_df = pd.concat([df_users, new_user], ignore_index=True)
                            conn.update(spreadsheet=SHEET_URL, worksheet="Users", data=updated_users_df)
                            st.cache_data.clear() ; st.toast("✅ 使用者新增成功！") ; time.sleep(1) ; st.rerun()

                st.write("✏️ **編輯或停用使用者**")
                
                user_options = df_users['姓名'].astype(str) + " (" + df_users['帳號'].astype(str) + ")"
                selected_user_str = st.selectbox("🔍 搜尋或選擇要編輯的使用者", user_options, index=0)
                
                if selected_user_str:
                    selected_idx = user_options[user_options == selected_user_str].index[0]
                    selected_user = df_users.loc[selected_idx]
                    
                    with st.form("edit_user_form"):
                        st.info(f"正在編輯：**{selected_user['姓名']}** 的資料")
                        e_c1, e_c2, e_c3 = st.columns(3)
                        with e_c1: 
                            e_acc = st.text_input("帳號", value=selected_user['帳號'])
                        with e_c2: 
                            e_pwd = st.text_input("密碼", value=str(selected_user['密碼']))
                        with e_c3: 
                            e_name = st.text_input("姓名", value=selected_user['姓名'])
                        
                        e_c4, e_c5, e_c6 = st.columns(3)
                        with e_c4: 
                            roles = ["教師", "管理者"]
                            cur_role = selected_user.get('權限', '教師')
                            default_r_idx = roles.index(cur_role) if cur_role in roles else 0
                            e_role = st.selectbox("權限", roles, index=default_r_idx)
                        with e_c5: 
                            e_email = st.text_input("Email", value=selected_user.get('Email', ''))
                        with e_c6:
                            statuses = ["啟用", "停用"]
                            cur_status = selected_user.get('狀態', '啟用')
                            default_s_idx = statuses.index(cur_status) if cur_status in statuses else 0
                            e_status = st.selectbox("狀態", statuses, index=default_s_idx)
                            
                        if st.form_submit_button("💾 儲存修改"):
                            if not e_acc or not e_pwd or not e_name:
                                st.warning("⚠️ 帳號、密碼、姓名皆不可為空！")
                            else:
                                other_users = df_users.drop(selected_idx)
                                if e_acc in other_users['帳號'].values:
                                    st.error("❌ 此帳號已存在，請更換帳號名稱！")
                                elif selected_user['姓名'] == st.session_state.user_name and e_status == '停用':
                                    st.warning("⚠️ 為了安全起見，您無法停用您自己正在使用的帳號！")
                                else:
                                    df_users.at[selected_idx, '帳號'] = e_acc
                                    df_users.at[selected_idx, '密碼'] = e_pwd
                                    df_users.at[selected_idx, '姓名'] = e_name
                                    df_users.at[selected_idx, '權限'] = e_role
                                    df_users.at[selected_idx, 'Email'] = e_email
                                    df_users.at[selected_idx, '狀態'] = e_status
                                    
                                    conn.update(spreadsheet=SHEET_URL, worksheet="Users", data=df_users)
                                    st.cache_data.clear()
                                    st.success(f"✅ 【{e_name}】的資料更新成功！")
                                    time.sleep(1.5)
                                    st.rerun()

                st.markdown("---")
                st.subheader("📢 最新消息公告管理")
                with st.form("add_news_form", clear_on_submit=True):
                    c_date, c_title = st.columns([1, 3])
                    with c_date: n_date = st.date_input("發布日期")
                    with c_title: n_title = st.text_input("公告標題")
                    n_content = st.text_area("公告內容")
                    if st.form_submit_button("➕ 新增公告"):
                        if not n_title: st.warning("⚠️ 標題為必填！")
                        else:
                            new_news = pd.DataFrame([{"日期": n_date.strftime("%Y/%m/%d"), "標題": n_title, "內容": n_content}])
                            updated_news_df = pd.concat([df_news, new_news], ignore_index=True)
                            conn.update(spreadsheet=SHEET_URL, worksheet="News", data=updated_news_df)
                            st.cache_data.clear() ; st.rerun()
                if not df_news.empty:
                    for i, row in df_news.sort_values("日期", ascending=False).iterrows():
                        with st.expander(f"【{row.get('日期','')}】 {row.get('標題','')}"):
                            st.write(f"{row.get('內容','')}")
                            if st.button("🗑️ 刪除", key=f"del_news_{i}"):
                                updated_news_df = df_news.drop(i)
                                conn.update(spreadsheet=SHEET_URL, worksheet="News", data=updated_news_df)
                                st.cache_data.clear() ; st.rerun()

    except Exception as e:
        st.error(f"系統錯誤：{e}")