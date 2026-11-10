import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account
import io
import time
import threading
import html as _html
import base64
import re
import random
import string
from PIL import Image
import urllib.parse  
import requests  
from fpdf import FPDF  

# ==========================================
# 0. 基本設定與 CSS 美化 (依據最新版面需求調整)
# ==========================================
st.set_page_config(page_title="學習扶助教學輔導登錄系統", layout="wide", page_icon="🏫")

st.markdown("""
    <style>
    /* 全域間距調整 */
    .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; max-width: 1100px; }
    
    /* 登入首頁：方案 C－簡約清爽資訊風 */
    .login-home-marker { height: 0; margin: 0; padding: 0; }
    .login-brand-row {
        display: flex;
        align-items: center;
        gap: 20px;
        margin: 0 0 18px 0;
    }
    .login-school-logo {
        display: block;
        width: 88px;
        height: 88px;
        flex: 0 0 88px;
        object-fit: contain;
    }
    .login-brand {
        flex: 1 1 auto;
        min-width: 0;
        padding: 8px 0 16px 0;
        border-bottom: 2px solid #c6a15b;
    }
    .login-school-name {
        color: #16325c;
        font-size: 1.35rem;
        font-weight: 750;
        letter-spacing: 1px;
        line-height: 1.35;
        margin: 0 0 3px 0;
    }
    .login-school-en {
        color: #64748b;
        font-size: 0.72rem;
        letter-spacing: 1.1px;
        margin: 0 0 10px 0;
    }
    .login-system-name {
        color: #16325c;
        font-size: 1.45rem;
        font-weight: 700;
        line-height: 1.45;
        margin: 0;
    }
    .login-system-en {
        color: #64748b;
        font-size: 0.82rem;
        margin-top: 3px;
    }
    .login-news-panel { margin-top: 10px; }
    .login-news-heading {
        color: #16325c;
        font-size: 1.12rem;
        font-weight: 750;
        padding: 0 0 10px 0;
        margin: 0 0 12px 0;
        border-bottom: 1px solid #d9dee7;
    }
    .login-news-list {
        max-height: 690px;
        overflow-y: auto;
        padding-right: 5px;
    }
    .news-item-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-left: 3px solid #c6a15b;
        border-radius: 5px;
        padding: 13px 15px;
        margin: 0 0 10px 0;
    }
    .news-item-title-wrap {
        display: flex;
        flex-wrap: wrap;
        align-items: baseline;
        gap: 5px 12px;
        margin-bottom: 7px;
    }
    .news-item-title {
        color: #16325c;
        font-size: 0.98rem;
        font-weight: 700;
        line-height: 1.5;
        margin: 0;
    }
    .news-item-date {
        color: #7b8492;
        font-size: 0.78rem;
        white-space: nowrap;
    }
    .news-item-body {
        color: #475569;
        font-size: 0.88rem;
        line-height: 1.65;
        margin: 0;
        white-space: pre-wrap;
        overflow-wrap: anywhere;
    }
    .login-news-empty {
        color: #64748b;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 5px;
        padding: 18px;
        font-size: 0.9rem;
    }

    .login-header { text-align: center; margin-bottom: 24px; }
    .login-icon { font-size: 29px; color: #16325c; margin-bottom: 6px; }
    .login-title { font-size: 1.35rem; font-weight: 700; color: #16325c; margin: 0; }
    .login-subtitle { font-size: 0.88rem; color: #64748b; margin-top: 5px; }

    /* 登入頁窄螢幕調整；只在未登入首頁生效 */
    @media (max-width: 760px) {
        .stApp:has(.login-home-marker) div[data-testid="stHorizontalBlock"]:has(.login-news-panel) {
            flex-direction: column !important;
            gap: 1rem !important;
        }
        .stApp:has(.login-home-marker) div[data-testid="stHorizontalBlock"]:has(.login-news-panel) > div[data-testid="column"] {
            width: 100% !important;
            flex: 1 1 100% !important;
        }
        .login-brand-row { gap: 12px; }
        .login-school-logo { width: 64px; height: 64px; flex-basis: 64px; }
        .login-system-name { font-size: 1.22rem; }
        .login-news-list { max-height: 420px; }
    }
    
    /* 頁尾 */
    .sys-footer { text-align: center; color: #64748b; font-size: 0.82rem; margin-top: 40px; margin-bottom: 20px; line-height: 1.6; }
    .sys-footer a { color: #1b3e6f; text-decoration: underline; font-weight: 600; }
    
    /* Streamlit 按鈕外觀客製 */
    div.stButton > button[kind="primary"] {
        background-color: #1b3e6f !important;
        color: white !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 0.5rem 1rem !important;
    }
    </style>
""", unsafe_allow_html=True)

SHEET_URL = "https://docs.google.com/spreadsheets/d/1D1YXEMnsa58p0RKHgaOyDFo3290d03hz_JY_bJrbcpM/edit"
FOLDER_ID = "0AJdsyYPdSxydUk9PVA"
GAS_MAIL_URL = "https://script.google.com/macros/s/AKfycbxIQXJ-QENFEZn6IgyQLdESmdPMMj4NTKF1O4FlE4WhAPeBanbYFMoprVfiqo4c2Wt4/exec"

# 預警系統定義之 9 大科目
WARN_SUBJECTS = ["國語", "本土語", "英語", "數學", "社會", "自然", "健體", "藝文", "綜合"]
WARN_STUDENT_COLS = ["班別", "學生姓名"] + [col for sub in WARN_SUBJECTS for col in (sub, f"{sub}教學者")]

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

# Streamlit 1.37+ 支援 fragment（局部重跑）；舊版自動退回為一般函式
_fragment = getattr(st, "fragment", None) or getattr(st, "experimental_fragment", None) or (lambda f: f)

def safe_html_text(value):
    """公告文字轉成可安全嵌入 HTML 的單行字串：
    跳脫特殊字元，並把換行改成 <br>，避免空白行或縮排讓 Markdown 把 HTML 當成程式碼區塊"""
    try:
        if value is None or pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        value = value.strftime("%Y/%m/%d")
    text = _html.escape(str(value).strip())
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")

def flash(message, icon=None):
    """先記下訊息，rerun 之後再顯示（取代 toast + time.sleep 的人為等待）"""
    st.session_state["_flash"] = (message, icon)

def show_flash():
    data = st.session_state.pop("_flash", None)
    if data:
        st.toast(data[0], icon=data[1])

def get_drive_service():
    # 每個使用者連線只建立一次 Drive 服務物件（原本每上傳一個檔案就重建一次）
    cached = st.session_state.get("_drive_service")
    if cached is not None:
        return cached
    if "connections" not in st.secrets or "gsheets" not in st.secrets["connections"]:
        st.error("❌ 尚未於 st.secrets 配置 Google Credentials 連線設定。")
        return None
    creds_info = st.secrets["connections"]["gsheets"]
    creds = service_account.Credentials.from_service_account_info(creds_info, scopes=["https://www.googleapis.com/auth/drive"])
    service = build('drive', 'v3', credentials=creds, cache_discovery=False)
    st.session_state["_drive_service"] = service
    return service

def _compress_image_if_needed(file, filename):
    """大張圖片先縮小再上傳（長邊 1600px、JPEG），可大幅縮短上傳時間；其他檔案維持原樣"""
    raw = file.getvalue()
    mime = file.type
    try:
        if mime in ("image/jpeg", "image/png") and len(raw) > 400 * 1024:
            img = Image.open(io.BytesIO(raw))
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.thumbnail((1600, 1600))
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=85, optimize=True)
            small = buf.getvalue()
            if len(small) < len(raw):
                base = filename.rsplit(".", 1)[0] if "." in filename else filename
                return small, "image/jpeg", base + ".jpg"
    except Exception:
        pass
    return raw, mime, filename

def upload_to_drive(file, filename):
    try:
        service = get_drive_service()
        if not service: return None
        data, mime, filename = _compress_image_if_needed(file, filename)
        file_metadata = {'name': filename, 'parents': [FOLDER_ID]}
        # 小檔案用一般上傳即可，省掉 resumable 多出來的來回請求
        media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=len(data) > 5 * 1024 * 1024)
        uploaded_file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
        return uploaded_file.get('webViewLink')
    except Exception as e:
        st.session_state.pop("_drive_service", None)  # 發生錯誤時，下次重新建立連線
        st.error(f"❌ 雲端硬碟上傳失敗：{e}")
        return None

@st.cache_data(ttl=300, show_spinner=False)
def list_archive_sheet_names(sheet_id):
    """列出歷史封存分頁名稱（快取 5 分鐘，避免每次互動都重新呼叫 Sheets API）"""
    creds_info = st.secrets["connections"]["gsheets"]
    creds = service_account.Credentials.from_service_account_info(
        creds_info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
    meta = service.spreadsheets().get(spreadsheetId=sheet_id, fields="sheets.properties.title").execute()
    names = [s['properties']['title'] for s in meta.get('sheets', [])]
    return [n for n in names if n.startswith('Records_') or n.startswith('Warning_Records_')]

@_fragment
def render_target_checkboxes(teacher_students, form_id):
    """勾選輔導對象：放進 fragment，勾選時只重跑這一小塊，不必重跑整支程式"""
    st.markdown("**✅ 請勾選本次輔導的對象 (可多選)：**")
    selected = []
    chk_cols = st.columns(4)
    for i, ts in enumerate(teacher_students):
        if chk_cols[i % 4].checkbox(ts, key=f"add_chk_{ts}_{form_id}"):
            selected.append(ts)
    if st.session_state.edit_record_idx is None:
        if selected: apply_bg_color(selected[0].split(" - ")[1])
        else: apply_bg_color("預設")

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
        
        if "connections" not in st.secrets or "gsheets" not in st.secrets["connections"]:
            return False
            
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
            time.sleep(1.5)
            list_archive_sheet_names.clear()
        return True
    except Exception as e:
        st.error(f"自動建立資料表 {new_sheet_name} 失敗：{e}")
        return False

class CustomPDF(FPDF):
    def footer(self):
        self.set_y(-15)
        try:
            self.set_font('CustomFont', size=9)
            self.cell(0, 10, f"共 {{nb}} 頁，本頁是第 {self.page_no()} 頁", align='C')
        except:
            pass

@st.cache_data(show_spinner=False, max_entries=30, ttl=3600)
def generate_report_pdf(df, title_name, sys_name):
    pdf = CustomPDF()
    try: 
        pdf.add_font('CustomFont', '', 'font.ttf')
        pdf.set_font('CustomFont', size=9)
    except Exception:
        st.error("⚠️ 找不到字型檔 font.ttf，請確認檔案位置。")
        return None
    
    pdf.add_page()
    pdf.alias_nb_pages()

    pdf.set_font('CustomFont', size=18)
    pdf.cell(0, 15, f"{sys_name}", new_x="LMARGIN", new_y="NEXT", align='C')
    pdf.set_font('CustomFont', size=14)
    pdf.cell(0, 10, f"報表名稱：{title_name}", new_x="LMARGIN", new_y="NEXT", align='C')
    
    today_str = pd.Timestamp.today().strftime("%Y/%m/%d")
    pdf.set_font('CustomFont', size=10)
    pdf.cell(0, 10, f"統計日期：{today_str}", new_x="LMARGIN", new_y="NEXT", align='C')
    pdf.ln(2)

    pdf.set_font('CustomFont', size=9)
    col_widths = [pdf.get_string_width(str(col)) + 4 for col in df.columns]
    for _, row in df.iterrows():
        for i, item in enumerate(row):
            w = pdf.get_string_width(str(item)) + 4
            if w > col_widths[i]: col_widths[i] = w

    row_height = pdf.font_size + 4 
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
    pdf.cell(0, 10, "承辦人簽章：__________________________", new_x="LMARGIN", new_y="NEXT", align='R')
    pdf.set_font('CustomFont', size=9)
    pdf.cell(0, 5, f"列印日期：{today_str}", new_x="LMARGIN", new_y="NEXT", align='R')
    
    return bytes(pdf.output())

def get_auto_school_year():
    """根據系統當前時間，自動判斷並回傳所屬的學年度"""
    today = pd.Timestamp.today()
    tw_yr = today.year - 1911
    if today.month >= 8:
        return str(tw_yr)
    else:
        return str(tw_yr - 1)

def get_archive_term_name(diag_date_str):
    """根據統一診斷日期，自動判讀並產生學期備份名稱"""
    try:
        d = pd.to_datetime(diag_date_str)
        tw_yr = d.year - 1911
        term = '上' if d.month < 8 else '下'
        return f"{tw_yr}{term}"
    except Exception:
        return "新學期"

# 連線初始化防呆
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"❌ 無法連線至 Google Sheets，請檢查 Secrets 配置。錯誤資訊：{e}")
    st.stop()

# ==========================================
# 資料存取層：伺服器端共用快取 + 寫入鎖
#  - 讀取：同一張工作表在 ttl 內只向 Google 讀一次，所有使用者共用
#  - 寫入：只讓「被寫入的那一張表」快取失效（原本是整批清空快取，導致全部資料表重讀）
# ==========================================
@st.cache_resource
def _get_data_store():
    return {"data": {}, "ver": {}, "lock": threading.Lock()}

def read_sheet(worksheet, ttl=600):
    store = _get_data_store()
    now = time.time()
    if ttl > 0:
        entry = store["data"].get(worksheet)
        if entry and now - entry[0] < ttl:
            return entry[1].copy()
    ver_before = store["ver"].get(worksheet, 0)
    df = conn.read(spreadsheet=SHEET_URL, worksheet=worksheet, ttl=0)
    # 讀取期間若有人剛寫入，就不要把可能過期的資料放進快取
    if store["ver"].get(worksheet, 0) == ver_before:
        store["data"][worksheet] = (now, df)
    return df.copy()

def write_sheet(worksheet, df):
    store = _get_data_store()
    with store["lock"]:
        conn.update(spreadsheet=SHEET_URL, worksheet=worksheet, data=df)
        store["ver"][worksheet] = store["ver"].get(worksheet, 0) + 1
        store["data"].pop(worksheet, None)

def refresh_all_data():
    store = _get_data_store()
    with store["lock"]:
        store["data"].clear()

# ==========================================
# 2. 讀取系統設定與狀態初始化
# ==========================================
try:
    df_settings = read_sheet("Settings", ttl=600).dropna(how="all")
    sys_name = df_settings.loc[df_settings['設定項'] == 'SchoolName', '設定值'].values[0] if 'SchoolName' in df_settings['設定項'].values else "文昌國小"
    sys_logo = df_settings.loc[df_settings['設定項'] == 'LogoLink', '設定值'].values[0] if 'LogoLink' in df_settings['設定項'].values else ""
    sys_year = df_settings.loc[df_settings['設定項'] == 'SchoolYear', '設定值'].values[0] if 'SchoolYear' in df_settings['設定項'].values else "114上" 
    sys_date = df_settings.loc[df_settings['設定項'] == 'DiagDate', '設定值'].values[0] if 'DiagDate' in df_settings['設定項'].values else pd.Timestamp.today().strftime("%Y/%m/%d")
except Exception:
    sys_name = "文昌國小"
    sys_logo = ""
    sys_year = "114上"
    sys_date = pd.Timestamp.today().strftime("%Y/%m/%d")

try:
    df_news = read_sheet("News", ttl=600).dropna(how="all")
except Exception:
    df_news = pd.DataFrame(columns=["日期", "標題", "內容"])

if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "user_name" not in st.session_state: st.session_state.user_name = ""
if "user_role" not in st.session_state: st.session_state.user_role = "教師"
if "edit_record_idx" not in st.session_state: st.session_state.edit_record_idx = None
if "edit_record_type" not in st.session_state: st.session_state.edit_record_type = "扶助"
if "record_form_id" not in st.session_state: st.session_state.record_form_id = 0

# 忘記密碼狀態管理
if "reset_mode" not in st.session_state: st.session_state.reset_mode = False
if "reset_code" not in st.session_state: st.session_state.reset_code = ""
if "reset_account" not in st.session_state: st.session_state.reset_account = ""

show_flash()

# ==========================================
# 畫面區塊 1：登入與忘記密碼介面 (優化首頁視覺與版面)
# ==========================================
if not st.session_state.logged_in:
    apply_bg_color("純白") 
    
    # 方案 C：白色校務識別頁首，直接顯示 Settings 中原有的校徽資料
    st.markdown('<div class="login-home-marker"></div>', unsafe_allow_html=True)
    logo_html = ""
    if isinstance(sys_logo, str) and sys_logo.startswith("data:image"):
        logo_html = f'<img class="login-school-logo" src="{sys_logo}" alt="文昌國小校徽">'
    st.markdown(f"""
        <div class="login-brand-row">
            {logo_html}
            <div class="login-brand">
                <div class="login-school-name">{sys_name}</div>
                <div class="login-school-en">WUN CHANG ELEMENTARY SCHOOL</div>
                <div class="login-system-name">學習扶助教學輔導登錄系統</div>
                <div class="login-system-en">Learning Support &amp; Teaching Guidance System</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    # 公告與登入採左右分欄；完整呈現 News 工作表內所有公告，不再只顯示前三筆
    col_news, col_login = st.columns([1.15, 0.85], gap="large")
    with col_news:
        news_html = [
            '<div class="login-news-panel">',
            '<div class="login-news-heading">最新消息 <span style="font-size:0.72rem;font-weight:500;color:#8490a0;letter-spacing:1px;">NEWS &amp; ANNOUNCEMENTS</span></div>',
            '<div class="login-news-list">'
        ]
        if df_news.empty:
            news_html.append('<div class="login-news-empty">目前尚無公告。</div>')
        else:
            for _, n in df_news.sort_values("日期", ascending=False).iterrows():
                news_html.append(
                    '<div class="news-item-card">'
                    '<div class="news-item-title-wrap">'
                    f'<span class="news-item-title">{safe_html_text(n.get("標題", ""))}</span>'
                    f'<span class="news-item-date">{safe_html_text(n.get("日期", ""))}</span>'
                    '</div>'
                    f'<div class="news-item-body">{safe_html_text(n.get("內容", ""))}</div>'
                    '</div>'
                )
        news_html.extend(['</div>', '</div>'])
        st.markdown("".join(news_html), unsafe_allow_html=True)

    with col_login:
        
        # 模式 1：標準登入
        if not st.session_state.reset_mode:
            st.markdown("""
                <div class="login-header">
                    <div class="login-icon">🔒</div>
                    <h3 class="login-title">使用者登入</h3>
                    <div class="login-subtitle">請輸入您的帳號與密碼</div>
                </div>
            """, unsafe_allow_html=True)
            
            do_login = False
            matched_user_name = ""
            matched_user_role = "教師"
            
            with st.form("login_form"):
                login_account = st.text_input("帳號", placeholder="請輸入使用者帳號")
                login_password = st.text_input("密碼", type="password", placeholder="請輸入密碼")
                submit_login = st.form_submit_button("登入系統", use_container_width=True, type="primary")
                
                if submit_login:
                    try:
                        df_users = read_sheet("Users", ttl=120).dropna(how="all")
                        df_users['密碼'] = df_users['密碼'].astype(str)
                        user_match = df_users[(df_users['帳號'] == login_account) & (df_users['密碼'] == str(login_password))]
                        
                        if not user_match.empty:
                            user_status = user_match.iloc[0].get('狀態', '啟用')
                            if user_status == '停用':
                                st.error("❌ 此帳號已被停用，請聯絡系統管理員。")
                            else:
                                do_login = True
                                matched_user_name = user_match.iloc[0]['姓名']
                                matched_user_role = user_match.iloc[0].get('權限', '教師')
                        else:
                            st.error("❌ 帳號或密碼錯誤")
                    except Exception as e:
                        st.error(f"連線失敗：{e}")

            # 忘記密碼觸發按鈕
            c_left, c_center, c_right = st.columns([1, 2, 1])
            with c_center:
                if st.button("🔑 忘記密碼？點此寄送驗證碼重置", use_container_width=True):
                    st.session_state.reset_mode = True
                    st.rerun()

            if do_login:
                st.session_state.logged_in = True
                st.session_state.user_name = matched_user_name
                st.session_state.user_role = matched_user_role
                flash(f"歡迎回來，{matched_user_name}！", icon="👋")
                st.rerun()
                
        # 模式 2：忘記密碼與驗證碼重置
        else:
            st.markdown("""
                <div class="login-header">
                    <div class="login-icon">🔑</div>
                    <h3 class="login-title">重置密碼</h3>
                    <div class="login-subtitle">請輸入您的帳號或 Email 獲取驗證碼</div>
                </div>
            """, unsafe_allow_html=True)
            
            if not st.session_state.reset_code:
                # 步驟 1：輸入帳號/Email 寄發驗證碼
                with st.form("send_code_form"):
                    input_acc_or_email = st.text_input("帳號或註冊 Email", placeholder="例如: teacher01 或 teacher@school.edu.tw")
                    btn_send_code = st.form_submit_button("✉️ 寄送驗證碼", use_container_width=True, type="primary")
                    
                    if btn_send_code:
                        if not input_acc_or_email:
                            st.warning("⚠️ 請輸入帳號或 Email！")
                        else:
                            try:
                                df_users = read_sheet("Users", ttl=120).dropna(how="all")
                                u_match = df_users[(df_users['帳號'] == input_acc_or_email) | (df_users['Email'] == input_acc_or_email)]
                                
                                if u_match.empty:
                                    st.error("❌ 找不到匹配的使用者帳號或 Email。")
                                else:
                                    user_info = u_match.iloc[0]
                                    target_email = str(user_info.get('Email', '')).strip()
                                    
                                    if not target_email or target_email == 'nan':
                                        st.error("❌ 此帳號未設定 Email，無法寄送驗證碼，請洽系統管理者。")
                                    else:
                                        v_code = ''.join(random.choices(string.digits, k=6))
                                        st.session_state.reset_code = v_code
                                        st.session_state.reset_account = user_info['帳號']
                                        
                                        subject = f"[{sys_name}] 密碼重置認證碼"
                                        body = f"{user_info['姓名']} 老師您好：\n\n您正在進行密碼重置，您的驗證碼為：【 {v_code} 】\n請在頁面輸入此驗證碼以完成密碼變更。\n\n若非本人操作請忽略此信件。"
                                        
                                        with st.spinner("正在傳送驗證碼至信箱..."):
                                            resp = requests.post(GAS_MAIL_URL, json={"email": target_email, "subject": subject, "body": body})
                                            if resp.status_code == 200 and resp.json().get("status") == "success":
                                                flash(f"✅ 驗證碼已成功寄送至 {target_email[:3]}***{target_email[target_email.find('@'):]} ！")
                                                st.rerun()
                                            else:
                                                st.error(f"❌ 郵件寄送失敗，請稍後再試。")
                            except Exception as e:
                                st.error(f"系統錯誤：{e}")
            else:
                # 步驟 2：輸入驗證碼與新密碼
                st.info(f"驗證碼已寄出，請檢查您的電子郵件。")
                with st.form("verify_code_form"):
                    user_code = st.text_input("輸入 6 位數驗證碼", max_chars=6)
                    new_pwd = st.text_input("設定新密碼", type="password")
                    new_pwd_confirm = st.text_input("再次確認新密碼", type="password")
                    btn_submit_reset = st.form_submit_button("💾 重置密碼並登入", use_container_width=True, type="primary")
                    
                    if btn_submit_reset:
                        if user_code != st.session_state.reset_code:
                            st.error("❌ 驗證碼錯誤，請重新確認！")
                        elif not new_pwd or len(new_pwd) < 4:
                            st.warning("⚠️ 新密碼至少需要 4 個字元！")
                        elif new_pwd != new_pwd_confirm:
                            st.error("❌ 兩次輸入的新密碼不一致！")
                        else:
                            try:
                                df_users = read_sheet("Users", ttl=0).dropna(how="all")
                                u_idx = df_users.index[df_users['帳號'] == st.session_state.reset_account].tolist()[0]
                                df_users.at[u_idx, '密碼'] = new_pwd
                                write_sheet("Users", df_users)
                                
                                flash("🎉 密碼已成功變更！請回登入頁面進行登入。")
                                st.session_state.reset_mode = False
                                st.session_state.reset_code = ""
                                st.session_state.reset_account = ""
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ 密碼更新失敗：{e}")
                                
            # 返回登入按鈕
            if st.button("⬅️ 返回登入頁面", use_container_width=True):
                st.session_state.reset_mode = False
                st.session_state.reset_code = ""
                st.session_state.reset_account = ""
                st.rerun()
                

    # 取得超級管理者的 Email 並在頁尾加上 mailto 連結
    admin_email = ""
    try:
        df_users_footer = read_sheet("Users", ttl=600).dropna(how="all")
        admin_match = df_users_footer[df_users_footer['權限'].isin(['超級管理者', '管理者'])]
        if not admin_match.empty and 'Email' in admin_match.columns:
            valid_emails = admin_match['Email'].dropna().tolist()
            if valid_emails:
                admin_email = str(valid_emails[0]).strip()
    except Exception:
        admin_email = ""
        
    footer_unit_html = f'<a href="mailto:{admin_email}">系統管理單位</a>' if admin_email else "系統管理單位"

    # 頁尾資訊
    st.markdown(f"""
        <div class="sys-footer">
            臺中市北屯區文昌國民小學<br>
            學習扶助教學輔導登錄平台 | {footer_unit_html}
        </div>
    """, unsafe_allow_html=True)

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
        if st.button("🔄 重新載入資料", use_container_width=True, help="若您剛在 Google 試算表直接修改了資料，可按此立即更新"):
            refresh_all_data()
            st.rerun()
        if st.button("🚪 登出系統", use_container_width=True):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun() 

    st.markdown("---")

    try:
        df_users = read_sheet("Users", ttl=600).dropna(how="all")
        if '狀態' not in df_users.columns: df_users['狀態'] = '啟用'
            
        df_students = read_sheet("Students", ttl=600).dropna(how="all")
        df_records = read_sheet("Records", ttl=600).dropna(how="all")
        
        try:
            df_warn_students = read_sheet("Warning_Students", ttl=600).dropna(how="all")
            for col in WARN_STUDENT_COLS:
                if col not in df_warn_students.columns:
                    df_warn_students[col] = ""
        except Exception:
            df_warn_students = pd.DataFrame(columns=WARN_STUDENT_COLS)
            
        try:
            df_warn_records = read_sheet("Warning_Records", ttl=600).dropna(how="all")
            required_warn_cols = ["學年度", "預警科目", "學生姓名", "輔導內容摘要", "教學輔導策略", "學習狀況", "填報日期", "輔導日期", "指導教師", "佐證資料連結"]
            if list(df_warn_records.columns) != required_warn_cols:
                df_warn_records = pd.DataFrame(columns=required_warn_cols)
        except Exception:
            df_warn_records = pd.DataFrame(columns=["學年度", "預警科目", "學生姓名", "輔導內容摘要", "教學輔導策略", "學習狀況", "填報日期", "輔導日期", "指導教師", "佐證資料連結"])

        # ==================================================
        # 動態分頁選單生成 (策略一：依據細分角色授權)
        # ==================================================
        cur_role = st.session_state.user_role
        tabs_list = ["✍️ 新增輔導紀錄", "🗂️ 個人紀錄回顧", "🕰️ 歷史紀錄查詢", "⚙️ 個人設定"]
        
        has_report_access = cur_role in ["管理者", "超級管理者", "扶助承辦人", "預警承辦人"]
        has_sys_settings_access = cur_role in ["管理者", "超級管理者"]
        
        if has_report_access:
            tabs_list.extend(["📊 學生報表", "👨‍🏫 教師報表"])
        if has_sys_settings_access:
            tabs_list.append("⚙️ 系統設定")
        
        # 改用單選導覽：只執行目前所選功能的程式碼（st.tabs 每次都會把所有分頁全部執行一遍）
        active_tab = st.radio("功能選單", tabs_list, horizontal=True, key="main_nav", label_visibility="collapsed")
        st.markdown("---")

        # ==================================================
        # 分頁 1：✍️ 新增輔導紀錄
        # ==================================================
        if active_tab == tabs_list[0]:
            sys_mode = st.radio("🔍 請選擇欲填報的輔導系統：", ["📚 學習扶助系統", "⚠️ 學習預警系統"], horizontal=True)
            st.markdown("---")
            
            if sys_mode == "📚 學習扶助系統":
                st.subheader("📚 新增學習扶助紀錄")
                teacher_students = []
                for _, row in df_students.iterrows():
                    for sub in ['國語', '英語', '數學']:
                        if str(row.get(f'{sub}_教學者', '')).strip() == st.session_state.user_name:
                            teacher_students.append(f"{row['學生姓名']} - {sub}")
                
                if not teacher_students:
                    st.info("您目前沒有被指派負責的扶助學生。")
                    apply_bg_color("預設") 
                else:
                    render_target_checkboxes(teacher_students, st.session_state.record_form_id)
                    selected_st_subs = [ts for ts in teacher_students if st.session_state.get(f"add_chk_{ts}_{st.session_state.record_form_id}", False)]
                            
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
                                    write_sheet("Records", updated_df)
                                    st.session_state.record_form_id += 1 
                                    flash(f'✅ 已成功儲存 {len(new_records_list)} 筆扶助紀錄！', icon='🎉')
                                    st.rerun()
                                    
            else:
                apply_bg_color("預設")
                st.subheader("⚠️ 新增學習預警紀錄")
                
                my_warn_subjects = []
                for sub in WARN_SUBJECTS:
                    teacher_col = f"{sub}教學者"
                    if teacher_col in df_warn_students.columns:
                        if (df_warn_students[teacher_col].astype(str).str.strip() == st.session_state.user_name).any():
                            my_warn_subjects.append(sub)
                
                if not my_warn_subjects:
                    st.info("目前沒有分配給您的預警學生名單。")
                else:
                    selected_warn_subject = st.selectbox("📌 步驟一：請選擇本次輔導的預警科目", my_warn_subjects)
                    
                    teacher_col = f"{selected_warn_subject}教學者"
                    df_target_students = df_warn_students[df_warn_students[teacher_col].astype(str).str.strip() == st.session_state.user_name]
                    student_options = df_target_students.apply(lambda row: f"{row['學生姓名']} ({row['班別']})", axis=1).tolist()
                    
                    with st.form("warning_record_form", clear_on_submit=True):
                        st.markdown(f"**📌 步驟二：請選擇參與輔導的【{selected_warn_subject}】預警學生 (可多選)：**")
                        selected_warn_students = st.multiselect("勾選學生", student_options)
                        
                        c1, c2 = st.columns(2)
                        with c1:
                            counsel_term = get_auto_school_year()
                            st.info(f"⏳ **(系統帶入) 教學輔導學年度：** {counsel_term}")
                            counsel_date = st.date_input("🗓️ 實際輔導日期", value=pd.Timestamp.today().date())
                        with c2:
                            summary = st.text_input("輔導內容摘要 / 觀察紀錄")
                            strategy = st.multiselect("輔導策略", ["個別指導", "小組合作", "學習單練習", "線上資源輔助", "親師聯繫", "其他"])
                            status = st.radio("學習狀況", ["有進步", "持平", "需持續觀察", "退步"])
                        
                        p_link = st.text_input("佐證網址 (選填)")
                        u_files = st.file_uploader("上傳佐證檔案 (最多三個)", type=["png", "jpg", "jpeg", "pdf", "docx"], accept_multiple_files=True)
                        
                        if st.form_submit_button("💾 批次儲存預警紀錄"):
                            if not selected_warn_students: st.warning("⚠️ 請至少選擇一位學生！")
                            elif not summary: st.warning("⚠️ 輔導內容摘要為必填！")
                            else:
                                with st.spinner('上傳資料中...'):
                                    p_entries = []
                                    if p_link: p_entries.append(f"連結: {p_link}")
                                    if u_files:
                                        for f in u_files[:3]:
                                            drive_link = upload_to_drive(f, f.name)
                                            if drive_link: p_entries.append(f"檔案: {drive_link}")
                                    
                                    new_warn_records = []
                                    for stu_info in selected_warn_students:
                                        name = stu_info.split(" (")[0]
                                        new_warn_records.append({
                                            "學年度": counsel_term, 
                                            "預警科目": selected_warn_subject, 
                                            "學生姓名": name,
                                            "輔導內容摘要": summary, 
                                            "教學輔導策略": ", ".join(strategy),
                                            "學習狀況": status, 
                                            "填報日期": sys_date, 
                                            "輔導日期": counsel_date.strftime("%Y/%m/%d"),
                                            "指導教師": st.session_state.user_name,
                                            "佐證資料連結": " / ".join(p_entries)
                                        })
                                    
                                    new_warn_df = pd.DataFrame(new_warn_records)
                                    updated_warn_df = pd.concat([df_warn_records, new_warn_df], ignore_index=True)
                                    write_sheet("Warning_Records", updated_warn_df)
                                    flash(f'✅ 已成功批次儲存 {len(new_warn_records)} 筆預警紀錄！', icon='🎉')
                                    st.rerun()

        # ==================================================
        # 分頁 2：🗂️ 個人紀錄回顧
        # ==================================================
        if active_tab == tabs_list[1]:
            if st.session_state.edit_record_idx is None:
                rev_mode = st.radio("🔍 選擇查詢範圍：", ["📚 學習扶助系統", "⚠️ 學習預警系統"], horizontal=True, key="rev_mode")
                st.markdown("---")
                
                if rev_mode == "📚 學習扶助系統":
                    st.subheader("🗂️ 您的學習扶助紀錄")
                    teacher_students_list = []
                    for _, row in df_students.iterrows():
                        for sub in ['國語', '英語', '數學']:
                            if str(row.get(f'{sub}_教學者', '')).strip() == st.session_state.user_name:
                                teacher_students_list.append(row['學生姓名'])
                    
                    my_records = df_records[df_records['學生姓名'].isin(teacher_students_list)]
                    if my_records.empty: st.info("尚無相關紀錄")
                    else:
                        _lim = st.session_state.get("_lim_sup", 30)
                        _recs_sorted = my_records.sort_values("輔導日期", ascending=False)
                        for idx, row in _recs_sorted.head(_lim).iterrows():
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
                                    if st.button("✏️ 編輯", key=f"btn_edit_sup_{idx}"):
                                        st.session_state.edit_record_idx = idx
                                        st.session_state.edit_record_type = "扶助"
                                        st.rerun()
                                with col_del:
                                    if st.button("🗑️ 刪除", key=f"btn_del_sup_{idx}"):
                                        df_records = df_records.drop(idx)
                                        write_sheet("Records", df_records)
                                        flash("✅ 紀錄已成功刪除！", icon="🗑️")
                                        st.rerun()
                                        
                        if len(_recs_sorted) > _lim:
                            if st.button(f"⬇️ 顯示更多（目前顯示 {_lim} / 共 {len(_recs_sorted)} 筆）", key="more_sup"):
                                st.session_state["_lim_sup"] = _lim + 30
                                st.rerun()
                else:
                    st.subheader("🗂️ 您的學習預警紀錄")
                    my_warn_records = df_warn_records[df_warn_records['指導教師'] == st.session_state.user_name]
                    if my_warn_records.empty: st.info("尚無相關紀錄")
                    else:
                        _lim = st.session_state.get("_lim_warn", 30)
                        _recs_sorted = my_warn_records.sort_values("輔導日期", ascending=False)
                        for idx, row in _recs_sorted.head(_lim).iterrows():
                            with st.expander(f"📅 {row.get('輔導日期', '')} | {row['學生姓名']} ({row['預警科目']}) - {row['輔導內容摘要']}"):
                                st.markdown(f"- **學年度：** {row.get('學年度', '未註記')}\n- **輔導日期：** {row.get('輔導日期', '未註記')}\n- **策略：** {row['教學輔導策略']}\n- **狀況：** {row['學習狀況']}")
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
                                    if st.button("✏️ 編輯", key=f"btn_edit_warn_{idx}"):
                                        st.session_state.edit_record_idx = idx
                                        st.session_state.edit_record_type = "預警"
                                        st.rerun()
                                with col_del:
                                    if st.button("🗑️ 刪除", key=f"btn_del_warn_{idx}"):
                                        df_warn_records = df_warn_records.drop(idx)
                                        write_sheet("Warning_Records", df_warn_records)
                                        flash("✅ 預警紀錄已刪除！", icon="🗑️")
                                        st.rerun()

                        if len(_recs_sorted) > _lim:
                            if st.button(f"⬇️ 顯示更多（目前顯示 {_lim} / 共 {len(_recs_sorted)} 筆）", key="more_warn"):
                                st.session_state["_lim_warn"] = _lim + 30
                                st.rerun()
            else:
                edit_idx = st.session_state.edit_record_idx
                rec_type = st.session_state.edit_record_type
                
                if rec_type == "扶助":
                    if edit_idx not in df_records.index:
                        st.session_state.edit_record_idx = None ; st.rerun()
                    edit_row = df_records.loc[edit_idx]
                    edit_sub = edit_row['授課科目']
                    apply_bg_color(edit_sub)
                    
                    st.subheader(f"✏️ 編輯扶助紀錄：{edit_row['學生姓名']} - {edit_sub}")
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
                                    write_sheet("Records", df_records)
                                    st.session_state.edit_record_idx = None
                                    flash('✅ 扶助紀錄更新成功！', icon='🎉')
                                    st.rerun()
                elif rec_type == "預警":
                    if edit_idx not in df_warn_records.index:
                        st.session_state.edit_record_idx = None ; st.rerun()
                    edit_row = df_warn_records.loc[edit_idx]
                    edit_sub = edit_row['預警科目']
                    apply_bg_color("預設")
                    
                    st.subheader(f"✏️ 編輯預警紀錄：{edit_row['學生姓名']} - {edit_sub}")
                    with st.form("edit_warn_record_form", clear_on_submit=False):
                        c1, c2 = st.columns(2)
                        with c1:
                            old_term = str(edit_row.get('學年度', get_auto_school_year()))
                            st.info(f"⏳ **(原紀錄) 學年度：** {old_term}")
                            try: default_c_date = pd.to_datetime(edit_row.get('輔導日期')).date()
                            except: default_c_date = pd.Timestamp.today().date()
                            new_counsel_date = st.date_input("🗓️ 實際輔導日期", value=default_c_date)
                        with c2:
                            new_summary = st.text_input("輔導內容摘要", value=str(edit_row['輔導內容摘要']))
                            old_strat_str = str(edit_row['教學輔導策略'])
                            old_strats = [s.strip() for s in old_strat_str.split(",")] if old_strat_str and old_strat_str != "nan" else []
                            available_strats = ["個別指導", "小組合作", "學習單練習", "線上資源輔助", "親師聯繫", "其他"]
                            valid_old_strats = [s for s in old_strats if s in available_strats]
                            new_strategy = st.multiselect("輔導策略", available_strats, default=valid_old_strats)
                            old_status = str(edit_row['學習狀況'])
                            status_options = ["有進步", "持平", "需持續觀察", "退步"]
                            default_status_idx = status_options.index(old_status) if old_status in status_options else 0
                            new_status = st.radio("學習狀況", status_options, index=default_status_idx)
                        
                        old_p_data = str(edit_row['佐證資料連結'])
                        st.markdown(f"📂 **目前佐證資料：** `{old_p_data if old_p_data and old_p_data != 'nan' else '無'}`")
                        new_p_link = st.text_input("新佐證網址 (若不更改請留白)")
                        new_u_files = st.file_uploader("上傳新佐證檔案 (若不更改請留白，最多三個)", type=["png", "jpg", "jpeg", "pdf", "docx"], accept_multiple_files=True)
                        
                        c_save, c_cancel = st.columns([1, 1])
                        with c_save: save_btn = st.form_submit_button("💾 儲存修改", use_container_width=True)
                        with c_cancel: cancel_btn = st.form_submit_button("❌ 取消編輯", use_container_width=True)
                            
                        if cancel_btn: st.session_state.edit_record_idx = None ; st.rerun()
                        if save_btn:
                            if not new_summary: st.warning("⚠️ 輔導摘要為必填！")
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
                                    
                                    df_warn_records.at[edit_idx, '輔導內容摘要'] = new_summary
                                    df_warn_records.at[edit_idx, '教學輔導策略'] = ", ".join(new_strategy)
                                    df_warn_records.at[edit_idx, '學習狀況'] = new_status
                                    df_warn_records.at[edit_idx, '輔導日期'] = new_counsel_date.strftime("%Y/%m/%d")
                                    df_warn_records.at[edit_idx, '佐證資料連結'] = " / ".join(final_p_entries)
                                    write_sheet("Warning_Records", df_warn_records)
                                    st.session_state.edit_record_idx = None
                                    flash('✅ 預警紀錄更新成功！', icon='🎉')
                                    st.rerun()

        # ==================================================
        # 分頁 3：🕰️ 歷史紀錄查詢
        # ==================================================
        if active_tab == tabs_list[2]:
            st.subheader("🕰️ 歷史紀錄查詢")
            st.write("系統會自動讀取已封存的歷史紀錄，請從下拉式選單選擇要查詢的資料表。")
            
            record_sheets = []
            try:
                sheet_id_match = re.search(r'/d/([a-zA-Z0-9-_]+)', SHEET_URL)
                if sheet_id_match and "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                    record_sheets = list_archive_sheet_names(sheet_id_match.group(1))
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
                            df_history = read_sheet(selected_archive, ttl=300).dropna(how="all")
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
        if active_tab == tabs_list[3]:
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
                        write_sheet("Users", df_users)
                        for key in list(st.session_state.keys()): del st.session_state[key]
                        flash("✅ 密碼修改成功！請重新登入。")
                        st.rerun()

        # ==================================================
        # 授權管理分頁 (依據角色開放)
        # ==================================================
        tab_idx_counter = 4
        
        # --- 📊 學生報表分頁 ---
        if has_report_access:
            if active_tab == tabs_list[tab_idx_counter]:
                st.subheader("📊 學生輔導狀況報表")
                
                allowed_modes = []
                if cur_role in ["管理者", "超級管理者", "扶助承辦人"]:
                    allowed_modes.append("📚 學習扶助系統")
                if cur_role in ["管理者", "超級管理者", "預警承辦人"]:
                    allowed_modes.append("⚠️ 學習預警系統")
                
                if not allowed_modes:
                    st.warning("⚠️ 您目前沒有可存取的學生報表類型。")
                else:
                    report_mode = st.radio("選擇產出報表類型：", allowed_modes, horizontal=True, key="rep_stu")
                    
                    if report_mode == "📚 學習扶助系統":
                        report_data = []
                        _cnt_sup = df_records.groupby(['學生姓名', '授課科目']).size().to_dict()
                        for _, stu in df_students.iterrows():
                            s_name = stu['學生姓名']
                            s_class = str(stu.get('班別', '?'))
                            for sub in ['國語', '英語', '數學']:
                                teacher = str(stu.get(f'{sub}_教學者', '')).strip()
                                if teacher and teacher != "nan":
                                    count = _cnt_sup.get((s_name, sub), 0)
                                    report_data.append({
                                        "班別": s_class, "學生姓名": s_name, "科目": sub,
                                        "指導教師": teacher, "紀錄筆數": count,
                                        "狀態": "✅ 已填寫" if count > 0 else "❌ 未填寫"
                                    })
                        df_report = pd.DataFrame(report_data)
                        
                        col_h1, col_dl_csv, col_dl_pdf = st.columns([3.5, 1, 1])
                        with col_dl_csv: 
                            csv_stu = df_report.to_csv(index=False).encode('utf-8-sig')
                            st.download_button(label="📥 匯出 CSV", data=csv_stu, file_name="學習扶助_學生狀況.csv", mime="text/csv", use_container_width=True)
                        with col_dl_pdf:
                            pdf_data = generate_report_pdf(df_report, "學習扶助_學生輔導狀況", clean_sys_name)
                            if pdf_data: st.download_button(label="📥 匯出 PDF", data=pdf_data, file_name="學習扶助_學生狀況.pdf", mime="application/pdf", use_container_width=True)
                        
                        classes = sorted(df_report['班別'].unique()) if not df_report.empty else []
                        for cls in classes:
                            with st.expander(f"🏫 【{cls}】 扶助學生列表"):
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
                                        if r['紀錄筆數'] > 0: is_expanded = st.toggle("查看", key=f"tgl_stu_sup_{cls}_{r['學生姓名']}_{r['科目']}")
                                    if is_expanded:
                                        with st.container():
                                            target_recs = df_records[(df_records['學生姓名'] == r['學生姓名']) & (df_records['授課科目'] == r['科目'])]
                                            for _, row in target_recs.sort_values("輔導日期", ascending=False).iterrows():
                                                with st.expander(f"📅 {row.get('輔導日期', row.get('診斷日期', ''))} | {row['教學內容摘要']}", expanded=True):
                                                    st.markdown(f"- **策略：** {row['教學輔導策略']}\n- **狀況：** {row['學習狀況']}")
                                    st.markdown('<div style="border-bottom: 1px dashed #eee; margin-bottom: 2px;"></div>', unsafe_allow_html=True)
                                    
                    elif report_mode == "⚠️ 學習預警系統":
                        report_data = []
                        _cnt_warn = df_warn_records.groupby(['學生姓名', '預警科目', '指導教師']).size().to_dict()
                        for _, stu in df_warn_students.iterrows():
                            s_name = stu['學生姓名']
                            s_class = str(stu.get('班別', '?'))
                            for sub in WARN_SUBJECTS:
                                teacher_col = f"{sub}教學者"
                                teacher = str(stu.get(teacher_col, '')).strip()
                                if teacher and teacher != "nan":
                                    count = _cnt_warn.get((s_name, sub, teacher), 0)
                                    report_data.append({
                                        "班別": s_class, "學生姓名": s_name, "預警科目": sub,
                                        "指導教師": teacher, "紀錄筆數": count,
                                        "狀態": "✅ 已填寫" if count > 0 else "❌ 未填寫"
                                    })
                        df_report_warn = pd.DataFrame(report_data)
                        
                        col_h1, col_dl_csv, col_dl_pdf = st.columns([3.5, 1, 1])
                        with col_dl_csv:
                            if not df_report_warn.empty:
                                csv_stu = df_report_warn.to_csv(index=False).encode('utf-8-sig')
                                st.download_button(label="📥 匯出 CSV", data=csv_stu, file_name="學習預警_學生狀況.csv", mime="text/csv", use_container_width=True)
                        with col_dl_pdf:
                            if not df_report_warn.empty:
                                pdf_data = generate_report_pdf(df_report_warn, "學習預警_學生輔導狀況", clean_sys_name)
                                if pdf_data: st.download_button(label="📥 匯出 PDF", data=pdf_data, file_name="學習預警_學生狀況.pdf", mime="application/pdf", use_container_width=True)

                        classes = sorted(df_report_warn['班別'].unique()) if not df_report_warn.empty else []
                        for cls in classes:
                            with st.expander(f"🏫 【{cls}】 預警學生列表"):
                                df_cls = df_report_warn[df_report_warn['班別'] == cls]
                                c1, c2, c3, c4, c5 = st.columns([1.5, 1, 1.5, 1.5, 1])
                                c1.markdown("**🧑‍🎓 姓名**") ; c2.markdown("**📖 科目**") ; c3.markdown("**🧑‍🏫 指導教師**")
                                c4.markdown("**📋 狀態**") ; c5.markdown("**🔍 操作**")
                                st.markdown('<div style="border-bottom: 2px solid #ccc; margin-bottom: 5px;"></div>', unsafe_allow_html=True)
                                for _, r in df_cls.iterrows():
                                    c1, c2, c3, c4, c5 = st.columns([1.5, 1, 1.5, 1.5, 1])
                                    c1.write(f"**{r['學生姓名']}**") ; c2.write(r['預警科目']) ; c3.write(r['指導教師'])
                                    c4.markdown(f"{r['狀態']} ({r['紀錄筆數']}筆)" if r['紀錄筆數']>0 else "<span style='color:red;'>❌ 未填寫</span>", unsafe_allow_html=True)
                                    is_expanded = False
                                    with c5:
                                        if r['紀錄筆數'] > 0: is_expanded = st.toggle("查看", key=f"tgl_stu_warn_{cls}_{r['學生姓名']}_{r['預警科目']}")
                                    if is_expanded:
                                        with st.container():
                                            target_recs = df_warn_records[(df_warn_records['學生姓名'] == r['學生姓名']) & (df_warn_records['預警科目'] == r['預警科目'])]
                                            for _, row in target_recs.sort_values("輔導日期", ascending=False).iterrows():
                                                with st.expander(f"📅 {row.get('輔導日期', '')} | {row['輔導內容摘要']}", expanded=True):
                                                    st.markdown(f"- **策略：** {row['教學輔導策略']}\n- **狀況：** {row['學習狀況']}")
                                    st.markdown('<div style="border-bottom: 1px dashed #eee; margin-bottom: 2px;"></div>', unsafe_allow_html=True)
            tab_idx_counter += 1

        # --- 👨‍🏫 教師報表分頁 ---
        if has_report_access:
            if active_tab == tabs_list[tab_idx_counter]:
                st.subheader("👨‍🏫 教師填寫進度報表")
                
                allowed_modes_tch = []
                if cur_role in ["管理者", "超級管理者", "扶助承辦人"]:
                    allowed_modes_tch.append("📚 學習扶助系統")
                if cur_role in ["管理者", "超級管理者", "預警承辦人"]:
                    allowed_modes_tch.append("⚠️ 學習預警系統")
                
                if not allowed_modes_tch:
                    st.warning("⚠️ 您目前沒有可存取的教師進度報表類型。")
                else:
                    report_mode_tch = st.radio("選擇產出報表類型：", allowed_modes_tch, horizontal=True, key="rep_tch")
                    
                    if report_mode_tch == "📚 學習扶助系統":
                        report_data = []
                        _cnt_sup = df_records.groupby(['學生姓名', '授課科目']).size().to_dict()
                        for _, stu in df_students.iterrows():
                            s_name = stu['學生姓名']
                            s_class = str(stu.get('班別', '?'))
                            for sub in ['國語', '英語', '數學']:
                                teacher = str(stu.get(f'{sub}_教學者', '')).strip()
                                if teacher and teacher != "nan":
                                    count = _cnt_sup.get((s_name, sub), 0)
                                    report_data.append({
                                        "班別": s_class, "學生姓名": s_name, "科目": sub,
                                        "指導教師": teacher, "紀錄筆數": count,
                                        "狀態": "✅ 已填寫" if count > 0 else "❌ 未填寫"
                                    })
                        df_report = pd.DataFrame(report_data)
                        
                        col_h1, col_dl_csv2, col_dl_pdf2 = st.columns([3.5, 1, 1])
                        with col_dl_csv2: 
                            csv_tch = df_report.to_csv(index=False).encode('utf-8-sig')
                            st.download_button(label="📥 匯出 CSV", data=csv_tch, file_name="學習扶助_教師進度.csv", mime="text/csv", use_container_width=True)
                        with col_dl_pdf2:
                            df_teacher_pdf = df_report.copy().sort_values(by="指導教師")[['指導教師', '班別', '學生姓名', '科目', '紀錄筆數', '狀態']]
                            pdf_data2 = generate_report_pdf(df_teacher_pdf, "學習扶助_教師填寫進度", clean_sys_name)
                            if pdf_data2: st.download_button(label="📥 匯出 PDF", data=pdf_data2, file_name="學習扶助_教師進度.pdf", mime="application/pdf", use_container_width=True)
                        
                        show_unassigned = st.checkbox("👁️ 顯示未指派學生的教師", value=True, key="chk_tch_sup")
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
                                        subject = f"[{sys_name}] 扶助紀錄填寫提醒"
                                        body = f"{t_name} 老師您好：\n\n系統顯示您尚有 {incomplete_count} 筆學生「學習扶助」紀錄尚未完成，請盡速登入系統進行填報。\n\n感謝您的辛勞與配合！"
                                        if st.button(f"✉️ 寄發 Email 提醒 ({incomplete_count}筆未填)", key=f"btn_mail_sup_{t_name}"):
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
                                            if r['紀錄筆數'] > 0: is_expanded = st.toggle("查看", key=f"tgl_tch_sup_{t_name}_{r['班別']}_{r['學生姓名']}_{r['科目']}")
                                        if is_expanded:
                                            with st.container():
                                                target_recs = df_records[(df_records['學生姓名'] == r['學生姓名']) & (df_records['授課科目'] == r['科目'])]
                                                for _, row in target_recs.sort_values("輔導日期", ascending=False).iterrows():
                                                    with st.expander(f"📅 {row.get('輔導日期', row.get('診斷日期', ''))} | {row['教學內容摘要']}", expanded=True):
                                                        st.markdown(f"- **策略：** {row['教學輔導策略']}\n- **狀況：** {row['學習狀況']}")
                                        st.markdown('<div style="border-bottom: 1px dashed #eee; margin-bottom: 2px;"></div>', unsafe_allow_html=True)
                                        
                    elif report_mode_tch == "⚠️ 學習預警系統":
                        report_data = []
                        _cnt_warn = df_warn_records.groupby(['學生姓名', '預警科目', '指導教師']).size().to_dict()
                        for _, stu in df_warn_students.iterrows():
                            s_name = stu['學生姓名']
                            s_class = str(stu.get('班別', '?'))
                            for sub in WARN_SUBJECTS:
                                teacher_col = f"{sub}教學者"
                                teacher = str(stu.get(teacher_col, '')).strip()
                                if teacher and teacher != "nan":
                                    count = _cnt_warn.get((s_name, sub, teacher), 0)
                                    report_data.append({
                                        "班別": s_class, "學生姓名": s_name, "科目": sub,
                                        "指導教師": teacher, "紀錄筆數": count,
                                        "狀態": "✅ 已填寫" if count > 0 else "❌ 未填寫"
                                    })
                        df_report_warn = pd.DataFrame(report_data)
                        
                        col_h1, col_dl_csv2, col_dl_pdf2 = st.columns([3.5, 1, 1])
                        with col_dl_csv2:
                            if not df_report_warn.empty:
                                csv_tch = df_report_warn.to_csv(index=False).encode('utf-8-sig')
                                st.download_button(label="📥 匯出 CSV", data=csv_tch, file_name="學習預警_教師進度.csv", mime="text/csv", use_container_width=True)
                        with col_dl_pdf2:
                            if not df_report_warn.empty:
                                df_teacher_pdf = df_report_warn.copy().sort_values(by="指導教師")[['指導教師', '班別', '學生姓名', '科目', '紀錄筆數', '狀態']]
                                pdf_data2 = generate_report_pdf(df_teacher_pdf, "學習預警_教師填寫進度", clean_sys_name)
                                if pdf_data2: st.download_button(label="📥 匯出 PDF", data=pdf_data2, file_name="學習預警_教師進度.pdf", mime="application/pdf", use_container_width=True)
                        
                        show_unassigned = st.checkbox("👁️ 顯示未指派學生的教師", value=True, key="chk_tch_warn")
                        teachers_list = df_users[df_users['權限'] == '教師']['姓名'].dropna().unique().tolist()
                        for t_name in teachers_list:
                            df_tch = df_report_warn[df_report_warn['指導教師'] == t_name]
                            if df_tch.empty:
                                if show_unassigned:
                                    with st.expander(f"🧑‍🏫 【{t_name} 老師】 - 無指派學生"): st.info("無指派對象")
                            else:
                                t_email = ""
                                if 'Email' in df_users.columns:
                                    email_match = df_users[df_users['姓名'] == t_name]['Email'].values
                                    if len(email_match) > 0 and pd.notna(email_match[0]): t_email = str(email_match[0]).strip()
                                incomplete_count = len(df_tch[df_tch['狀態'] == "❌ 未填寫"])
                                with st.expander(f"🧑‍🏫 【{t_name} 老師】 - 負責 {len(df_tch)} 項預警輔導"):
                                    if incomplete_count > 0 and t_email:
                                        subject = f"[{sys_name}] 預警紀錄填寫提醒"
                                        body = f"{t_name} 老師您好：\n\n系統顯示您尚有 {incomplete_count} 筆學生「學習預警」紀錄尚未完成，請盡速登入系統進行填報。\n\n感謝您的辛勞與配合！"
                                        if st.button(f"✉️ 寄發 Email 提醒 ({incomplete_count}筆未填)", key=f"btn_mail_warn_{t_name}"):
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
                                            if r['紀錄筆數'] > 0: is_expanded = st.toggle("查看", key=f"tgl_tch_warn_{t_name}_{r['班別']}_{r['學生姓名']}_{r['科目']}")
                                        if is_expanded:
                                            with st.container():
                                                target_recs = df_warn_records[(df_warn_records['學生姓名'] == r['學生姓名']) & (df_warn_records['預警科目'] == r['科目'])]
                                                for _, row in target_recs.sort_values("輔導日期", ascending=False).iterrows():
                                                    with st.expander(f"📅 {row.get('輔導日期', '')} | {row['輔導內容摘要']}", expanded=True):
                                                        st.markdown(f"- **策略：** {row['教學輔導策略']}\n- **狀況：** {row['學習狀況']}")
                                        st.markdown('<div style="border-bottom: 1px dashed #eee; margin-bottom: 2px;"></div>', unsafe_allow_html=True)
            tab_idx_counter += 1

        # --- ⚙️ 系統設定分頁 ---
        if has_sys_settings_access:
            if active_tab == tabs_list[tab_idx_counter]:
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
                        write_sheet("Settings", new_settings_df)
                        flash('已儲存！') ; st.rerun()
                
                # ==================================================
                # --- 學期轉換與備份管理 ---
                # ==================================================
                st.markdown("---")
                st.subheader("📦 學期轉換與資料備份管理")
                
                st.write("**1. 一鍵自動備份與封存**")
                
                dynamic_sys_year = get_archive_term_name(sys_date)
                
                st.info(f"💡 系統將自動建立 `{dynamic_sys_year}` 分頁，一鍵備份「使用者、學生、預警名單、扶助紀錄、預警紀錄」，並**清空**當前紀錄以供新學期使用。")
                
                if st.button("📦 備份資料並清空本學期所有輔導紀錄"):
                    with st.spinner("正在自動建立資料表並備份資料..."):
                        tables_to_backup = {
                            "Users": df_users,
                            "Students": df_students,
                            "Warning_Students": df_warn_students,
                            "Records": df_records,
                            "Warning_Records": df_warn_records
                        }
                        all_success = True
                        for base_name, df_data in tables_to_backup.items():
                            archive_name = f"{base_name}_{dynamic_sys_year}"
                            if auto_create_sheet(SHEET_URL, archive_name):
                                try:
                                    try:
                                        df_archive = read_sheet(archive_name, ttl=0).dropna(how="all")
                                        df_archive_new = pd.concat([df_archive, df_data], ignore_index=True).drop_duplicates()
                                    except:
                                        df_archive_new = df_data
                                        
                                    write_sheet(archive_name, df_archive_new)
                                except Exception as e:
                                    st.error(f"❌ {archive_name} 備份失敗：{e}")
                                    all_success = False
                            else:
                                all_success = False
                        
                        if all_success:
                            empty_records_df = pd.DataFrame(columns=df_records.columns)
                            write_sheet("Records", empty_records_df)
                            
                            empty_warn_records_df = pd.DataFrame(columns=df_warn_records.columns)
                            write_sheet("Warning_Records", empty_warn_records_df)
                            
                            flash(f"✅ 資料已成功備份至 `{dynamic_sys_year}` 歷史資料表，且本學期所有輔導紀錄已清空！")
                            st.rerun()

                st.write("**2. 下載名單範例檔**")
                col_dl1, col_dl2, col_dl3 = st.columns(3)
                with col_dl1:
                    empty_users_df = pd.DataFrame(columns=["帳號", "密碼", "姓名", "權限", "Email", "狀態"]) if df_users.empty else pd.DataFrame(columns=df_users.columns)
                    st.download_button("📥 下載【使用者名單】範例", data=empty_users_df.to_csv(index=False).encode('utf-8-sig'), file_name="使用者名單_範例.csv", mime="text/csv", use_container_width=True)
                with col_dl2:
                    empty_students_df = pd.DataFrame(columns=["班別", "學生姓名", "國語_教學者", "國語_成績", "英語_教學者", "英語_成績", "數學_教學者", "數學_成績"]) if df_students.empty else pd.DataFrame(columns=df_students.columns)
                    st.download_button("📥 下載【扶助學生】範例", data=empty_students_df.to_csv(index=False).encode('utf-8-sig'), file_name="扶助學生_範例.csv", mime="text/csv", use_container_width=True)
                with col_dl3:
                    empty_warn_stu_df = pd.DataFrame(columns=WARN_STUDENT_COLS)
                    st.download_button("📥 下載【預警學生】範例", data=empty_warn_stu_df.to_csv(index=False).encode('utf-8-sig'), file_name="預警學生_範例.csv", mime="text/csv", use_container_width=True)

                st.write("**3. 匯入新學期名單**")
                col_up1, col_up2, col_up3 = st.columns(3)
                with col_up1:
                    uploaded_users = st.file_uploader("上傳【使用者名單】", type=["csv"])
                    if uploaded_users and st.button("💾 更新使用者"):
                        with st.spinner("更新中..."):
                            new_users_df = pd.read_csv(uploaded_users)
                            write_sheet("Users", new_users_df)
                            flash("✅ 完成！") ; st.rerun()
                with col_up2:
                    uploaded_students = st.file_uploader("上傳【扶助學生】", type=["csv"])
                    if uploaded_students and st.button("💾 更新扶助學生"):
                        with st.spinner("更新中..."):
                            new_stu_df = pd.read_csv(uploaded_students)
                            write_sheet("Students", new_stu_df)
                            flash("✅ 完成！") ; st.rerun()
                with col_up3:
                    uploaded_warn_students = st.file_uploader("上傳【預警學生】", type=["csv"])
                    if uploaded_warn_students and st.button("💾 更新預警學生"):
                        with st.spinner("更新中..."):
                            new_w_stu_df = pd.read_csv(uploaded_warn_students)
                            write_sheet("Warning_Students", new_w_stu_df)
                            flash("✅ 完成！") ; st.rerun()

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
                    with c_role: n_role = st.selectbox("權限", ["教師", "扶助承辦人", "預警承辦人", "管理者"])
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
                            write_sheet("Users", updated_users_df)
                            flash("✅ 使用者新增成功！") ; st.rerun()

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
                            roles = ["教師", "扶助承辦人", "預警承辦人", "管理者", "超級管理者"]
                            cur_role_val = selected_user.get('權限', '教師')
                            default_r_idx = roles.index(cur_role_val) if cur_role_val in roles else 0
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
                                    
                                    write_sheet("Users", df_users)
                                    flash(f"✅ 【{e_name}】的資料更新成功！")
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
                            write_sheet("News", updated_news_df)
                            st.rerun()
                if not df_news.empty:
                    for i, row in df_news.sort_values("日期", ascending=False).iterrows():
                        with st.expander(f"【{row.get('日期','')}】 {row.get('標題','')}"):
                            st.write(f"{row.get('內容','')}")
                            if st.button("🗑️ 刪除", key=f"del_news_{i}"):
                                updated_news_df = df_news.drop(i)
                                write_sheet("News", updated_news_df)
                                st.rerun()

    except Exception as e:
        st.error(f"系統錯誤：{e}")