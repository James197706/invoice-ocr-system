"""
發票智能辨識系統 — 雲端版 v3.1（Security Patch）
Invoice OCR System for Finance Teams (Cloud Edition)

部署平台：Streamlit Cloud
功能：Claude AI 視覺辨識發票 → 匯出格式化 Excel

設定說明（Streamlit Cloud Secrets）：
  ANTHROPIC_API_KEY = "sk-ant-api..."
  APP_PASSWORD      = "your-password"   ← 必填，未設定系統將拒絕啟動
  COMPANY_NAME      = "貴公司名稱"          # 選填
  ADMIN_NOTE        = "IT 部門聯絡資訊"     # 選填，登入頁顯示
  DAILY_API_LIMIT   = 200                   # 選填，每日辨識張數上限（全站共用，預設 200）

Security Changelog (v3.1):
  - [P0] 移除硬編碼預設密碼，未設定 Secrets 直接拒絕服務
  - [P0] AI 回應欄位全面 html.escape()，防止 AI-Driven XSS
  - [P1] API 每日限額改為全站共享計數器（修正 per-session 繞過）
  - [P1] 暴力破解防護加入全站追蹤，新 session 無法重置計數
  - [P2] 執行緒鎖修正 _check_and_increment_api_count() 競態條件
  - [P2] 新增檔案大小上限（20MB）與 magic bytes 格式驗證
  - [P3] 新增 Session 閒置逾時（120 分鐘自動登出）
  - [P3] 密碼比對改用 hmac.compare_digest()（timing-safe）
  - [P3] 例外錯誤訊息不直接暴露給使用者，僅記錄至伺服器 log
"""

import streamlit as st
import anthropic
import base64
import html as html_module      # [SEC] XSS 防護
import hmac                     # [SEC] timing-safe 密碼比對
import json
import pandas as pd
import re
import threading                # [SEC] 執行緒鎖
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, date, timedelta
from io import BytesIO
from typing import Optional

try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

from excel_exporter import create_invoice_excel

# ─────────────────────────────────────────
# 版本常數
# ─────────────────────────────────────────
APP_VERSION = "3.2"

# ─────────────────────────────────────────
# 共用選項列表（DRY：單一定義，多處共用）
# ─────────────────────────────────────────
DEPT_OPTIONS = ["", "總務部", "業務部", "行銷部", "研發部", "財務部",
                "人資部", "資訊部", "管理部", "其他"]

ACCT_OPTIONS = ["", "辦公費", "文具用品", "交通費", "差旅費", "餐費/交際費",
                "廣告費", "租金費用", "水電費", "電話費", "郵寄費",
                "修繕費", "採購費用", "軟體授權費", "訓練費", "顧問費", "雜項費用"]

TYPE_OPTIONS = ["台灣二聯發票", "台灣三聯發票", "一般收據", "自製發票",
                "國外發票", "電子發票", "其他"]

PAY_OPTIONS  = ["", "現金", "轉帳", "信用卡", "支票", "其他"]

CURR_OPTIONS = ["TWD", "USD", "EUR", "JPY", "CNY", "其他"]

USAGE_GUIDE = [
    {
        "title": "1. 登入與開始使用",
        "body": [
            "輸入系統密碼後登入；若閒置超過 120 分鐘，系統會自動登出。",
            "首次使用前，請先確認左側欄位顯示「API 金鑰已設定」。",
            "建議先在側邊欄設定預設部門、費用科目與專案代碼，後續辨識結果會自動帶入空白欄位。",
        ],
    },
    {
        "title": "2. 上傳辨識",
        "body": [
            "支援 PNG、JPG、JPEG、PDF，單檔上限 20MB；PDF 會自動拆頁逐張辨識。",
            "可一次多選多張檔案，系統會並發辨識以縮短等待時間。",
            "若今日剩餘額度不足，系統會先處理可用額度內的張數，並提示剩餘檔案稍後再試。",
        ],
    },
    {
        "title": "3. 審核與修正",
        "body": [
            "辨識完成後切換到「審核資料」，逐筆確認發票日期、統編、金額與品項明細。",
            "低信心或辨識失敗的單據會顯示提醒，請優先人工覆核。",
            "若要刪除資料，系統會要求二次確認，避免誤刪。",
        ],
    },
    {
        "title": "4. 匯出與交付",
        "body": [
            "在「匯出 Excel」輸入報表標題、期間與製表人後，即可產生 Excel。",
            "輸出內容包含發票彙總、品項明細、統計分析三個工作表。",
            "若公司已有 ERP/會計系統，建議將匯出檔作為過渡流程，後續可再規劃 API 串接。",
        ],
    },
]

BEST_PRACTICES = [
    "同一批次盡量上傳同月份、同部門的單據，能降低審核切換成本。",
    "先確認發票影像方向正確、內容清晰，能明顯提升 OCR 成功率。",
    "財務關帳前，建議先篩查失敗單據與低信心單據，再進行整批匯出。",
    "若同供應商單據很多，可固定使用一致的費用科目與專案代碼，提升作業一致性。",
]

PLATFORM_RELEASE_NOTES = [
    {
        "version": "v3.2",
        "date": "2026-03-17",
        "items": [
            "平台內新增「使用說明」與「版本更新」頁籤，使用者可直接在系統內查閱操作指引與更新內容。",
            "新增辨識前剩餘額度預檢，避免超過今日可用額度的檔案進入辨識流程後才失敗。",
            "修正預設部門、費用科目、專案代碼帶入邏輯，空白欄位現在會正確套用。",
            "強化 AI JSON 解析，兼容多段文字與 code fence 格式回應。",
            "統一平台與後端的單檔上限為 20MB，避免操作說明與實際限制不一致。",
        ],
    },
    {
        "version": "v3.1",
        "date": "2026-03-17",
        "items": [
            "移除硬編碼預設密碼，未設定 APP_PASSWORD 時直接拒絕啟動。",
            "新增 AI 回應 XSS 防護、全站 API 用量統計、暴力破解防護與 session 閒置逾時。",
            "補上檔案大小限制、magic bytes 驗證與例外訊息遮罩。",
        ],
    },
    {
        "version": "v3.0",
        "date": "2026-03-12",
        "items": [
            "加入批次並發辨識、Anthropic Client 快取與用量進度條。",
            "改善刪除/清除二次確認與用量接近上限提醒。",
        ],
    },
]

# ─────────────────────────────────────────
# 頁面設定
# ─────────────────────────────────────────
st.set_page_config(
    page_title="發票智能辨識系統",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────
# 全域 CSS
# ─────────────────────────────────────────
st.markdown("""
<style>
/* ── 登入頁面 ── */
.login-wrapper {
    max-width: 420px;
    margin: 4rem auto;
    padding: 2.5rem;
    background: white;
    border-radius: 16px;
    box-shadow: 0 4px 32px rgba(30,58,138,.12);
    border: 1px solid #e2e8f0;
}
.login-logo { text-align:center; font-size:3.5rem; margin-bottom:.5rem; }
.login-title {
    text-align:center;
    font-size:1.4rem;
    font-weight:700;
    color:#1e3a8a;
    margin-bottom:.3rem;
}
.login-sub { text-align:center; color:#64748b; font-size:.88rem; margin-bottom:1.8rem; }

/* ── 主介面 ── */
.main-header {
    background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
    padding: 1.3rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    color: white;
}
.main-header h1 { margin:0; font-size:1.7rem; }
.main-header p  { margin:.3rem 0 0; opacity:.85; font-size:.9rem; }

.status-box { padding:.8rem 1.2rem; border-radius:8px; margin:.5rem 0; }
.status-success { background:#d1fae5; border-left:4px solid #10b981; color:#065f46; }
.status-error   { background:#fee2e2; border-left:4px solid #ef4444; color:#7f1d1d; }
.status-info    { background:#dbeafe; border-left:4px solid #3b82f6; color:#1e40af; }
.status-warn    { background:#fff3cd; border-left:4px solid #f59e0b; color:#78350f; }

div[data-testid="stExpander"] { border:1px solid #e2e8f0; border-radius:8px; }
.stButton>button { border-radius:8px; font-weight:600; }
.stDownloadButton>button {
    background:linear-gradient(135deg,#059669,#10b981);
    color:white; border:none;
    border-radius:8px; font-weight:600;
    padding:.6rem 1.5rem; width:100%;
}

/* 側邊欄使用者資訊 */
.user-badge {
    background:#eff6ff;
    border:1px solid #bfdbfe;
    border-radius:8px;
    padding:.7rem 1rem;
    margin-bottom:.5rem;
    font-size:.88rem;
    color:#1e40af;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────
# Secrets 讀取（雲端部署用）
# ─────────────────────────────────────────
def _get_secret(key: str, default: str = "") -> str:
    """安全讀取 Streamlit Secrets，找不到時回傳預設值"""
    try:
        return st.secrets[key]
    except (KeyError, FileNotFoundError):
        return default


API_KEY       = _get_secret("ANTHROPIC_API_KEY")
APP_PASSWORD  = _get_secret("APP_PASSWORD")        # [SEC P0] 無預設值，未設定 → 系統拒絕啟動
COMPANY_NAME  = _get_secret("COMPANY_NAME", "")
ADMIN_NOTE    = _get_secret("ADMIN_NOTE", "")
DAILY_LIMIT   = int(_get_secret("DAILY_API_LIMIT", "200"))

# [SEC P0] 對 Secrets 內容預先 escape，防止 XSS
COMPANY_NAME_SAFE = html_module.escape(COMPANY_NAME)
ADMIN_NOTE_SAFE   = html_module.escape(ADMIN_NOTE)


# ─────────────────────────────────────────
# Anthropic Client 快取
# ─────────────────────────────────────────
@st.cache_resource
def _get_anthropic_client() -> anthropic.Anthropic:
    """快取 Anthropic Client，整個應用生命週期只建立一次"""
    return anthropic.Anthropic(api_key=API_KEY)


# ─────────────────────────────────────────
# [SEC P1+P2] 全站共享狀態（執行緒安全）
# ─────────────────────────────────────────
@st.cache_resource
def _get_global_state() -> dict:
    """
    全站共享可變狀態，使用 threading.Lock 確保執行緒安全。
    修正問題：
      - API 計數從 per-session 改為全站共用（P1）
      - _check_and_increment 的 TOCTOU 競態條件（P2）
      - 登入失敗計數不再隨新 session 重置（P1）
    """
    return {
        "api_daily":      {},   # {日期字串: 呼叫次數}
        "auth_failures":  [],   # [datetime, ...] 近期登入失敗時間戳
        "lock":           threading.Lock(),
    }


# ─────────────────────────────────────────
# 每日 API 用量 Hard Limit（全站 + 執行緒安全）
# ─────────────────────────────────────────
def _get_today_key() -> str:
    return f"api_{date.today().isoformat()}"


def _check_and_increment_api_count() -> bool:
    """
    [SEC P1+P2] 全站共享、執行緒安全的 API 呼叫計數。
    用 Lock 消除競態條件；用 cache_resource 確保跨 session 共用。
    """
    gs = _get_global_state()
    key = _get_today_key()
    with gs["lock"]:
        gs["api_daily"].setdefault(key, 0)
        if gs["api_daily"][key] >= DAILY_LIMIT:
            return False
        gs["api_daily"][key] += 1
        return True


def _get_today_api_count() -> int:
    gs = _get_global_state()
    key = _get_today_key()
    with gs["lock"]:
        return gs["api_daily"].get(key, 0)


# ─────────────────────────────────────────
# Session State 初始化
# ─────────────────────────────────────────
def _init_state():
    defaults = {
        "authenticated":       False,
        "invoices":            [],
        "login_error":         False,
        "login_attempts":      0,
        "login_locked_at":     None,
        "confirm_clear":       False,
        "last_active_at":      None,      # [SEC P3] Session 逾時追蹤
        "session_timeout_msg": False,     # [SEC P3] 逾時提示旗標
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ─────────────────────────────────────────
# [SEC P3] Session 閒置逾時
# ─────────────────────────────────────────
SESSION_TIMEOUT_MINUTES = 120   # 2 小時無操作自動登出


def _check_and_refresh_session():
    """
    每次頁面重繪時呼叫：檢查閒置時間是否超過上限。
    超過則強制登出，下次渲染會導向登入頁。
    """
    if not st.session_state.authenticated:
        return
    last = st.session_state.last_active_at
    if last is not None:
        elapsed = (datetime.now() - last).total_seconds() / 60
        if elapsed > SESSION_TIMEOUT_MINUTES:
            st.session_state.authenticated       = False
            st.session_state.invoices            = []
            st.session_state.session_timeout_msg = True
            return
    st.session_state.last_active_at = datetime.now()

_check_and_refresh_session()


# ─────────────────────────────────────────
# [SEC P1] 全站暴力破解防護
# ─────────────────────────────────────────
GLOBAL_LOCKOUT_THRESHOLD = 30    # 全站 30 次失敗 → 觸發全站冷卻
GLOBAL_LOCKOUT_WINDOW    = 300   # 滑動視窗：5 分鐘（秒）

# Per-session 限制（防止單一 session 快速嘗試）
MAX_LOGIN_ATTEMPTS = 5
LOCKOUT_SECONDS    = 60


def _record_global_failure():
    """記錄一次全站登入失敗（執行緒安全）"""
    gs = _get_global_state()
    with gs["lock"]:
        gs["auth_failures"].append(datetime.now())
        # 只保留視窗內的記錄，避免記憶體無限成長
        cutoff = datetime.now() - timedelta(seconds=GLOBAL_LOCKOUT_WINDOW)
        gs["auth_failures"] = [t for t in gs["auth_failures"] if t > cutoff]


def _is_globally_locked() -> tuple[bool, int]:
    """
    [SEC P1] 檢查全站是否因過多失敗而鎖定。
    即使攻擊者不斷開新 session，全站計數仍持續累積。
    回傳 (是否鎖定, 剩餘秒數)。
    """
    gs = _get_global_state()
    with gs["lock"]:
        cutoff = datetime.now() - timedelta(seconds=GLOBAL_LOCKOUT_WINDOW)
        recent = [t for t in gs["auth_failures"] if t > cutoff]
        gs["auth_failures"] = recent
        if len(recent) >= GLOBAL_LOCKOUT_THRESHOLD:
            oldest = min(recent)
            remaining = GLOBAL_LOCKOUT_WINDOW - int((datetime.now() - oldest).total_seconds())
            return True, max(0, remaining)
    return False, 0


def _is_session_locked() -> tuple[bool, int]:
    """
    Per-session 鎖定（連續錯誤 5 次，鎖定 60 秒）。
    與全站鎖定搭配，形成雙層防護。
    """
    locked_at = st.session_state.login_locked_at
    if locked_at is None:
        return False, 0
    elapsed = (datetime.now() - locked_at).seconds
    if elapsed < LOCKOUT_SECONDS:
        return True, LOCKOUT_SECONDS - elapsed
    st.session_state.login_locked_at = None
    st.session_state.login_attempts  = 0
    return False, 0


def _record_session_failure():
    st.session_state.login_attempts += 1
    if st.session_state.login_attempts >= MAX_LOGIN_ATTEMPTS:
        st.session_state.login_locked_at = datetime.now()


def _reset_login_state():
    st.session_state.login_attempts  = 0
    st.session_state.login_locked_at = None


# ─────────────────────────────────────────
# [SEC P2] 檔案大小與格式驗證（Magic Bytes）
# ─────────────────────────────────────────
MAX_FILE_BYTES = 20 * 1024 * 1024   # 20 MB

_MAGIC_BYTES: dict[str, list[tuple[int, bytes]]] = {
    "pdf":  [(0, b"%PDF")],
    "png":  [(0, b"\x89PNG\r\n\x1a\n")],
    "jpg":  [(0, b"\xff\xd8")],
    "jpeg": [(0, b"\xff\xd8")],
    "webp": [(0, b"RIFF"), (8, b"WEBP")],
    "gif":  [(0, b"GIF8")],
}


def _validate_upload(raw: bytes, filename: str) -> Optional[str]:
    """
    驗證上傳檔案。
    - 超過 20MB → 拒絕
    - Magic bytes 不符 → 拒絕（防止副檔名偽裝）
    回傳 None 代表合法；回傳字串代表錯誤訊息。
    """
    if len(raw) > MAX_FILE_BYTES:
        mb = len(raw) / 1024 / 1024
        return f"檔案超過 20MB 限制（{mb:.1f} MB），請分批上傳"
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    checks = _MAGIC_BYTES.get(ext)
    if checks:
        for offset, magic in checks:
            if raw[offset : offset + len(magic)] != magic:
                return f"檔案格式驗證失敗（.{ext} 與實際內容不符），請確認檔案完整性"
    return None


# ─────────────────────────────────────────
# 登入頁面
# ─────────────────────────────────────────
def login_page():
    # 隱藏側邊欄
    st.markdown("""
    <style>
    [data-testid="stSidebar"] { display: none; }
    [data-testid="collapsedControl"] { display: none; }
    </style>
    """, unsafe_allow_html=True)

    # [SEC P0] 未設定密碼 → 系統安全錯誤，拒絕服務
    if not APP_PASSWORD:
        st.error("🔒 系統安全錯誤：尚未設定 APP_PASSWORD")
        st.info("請管理員至 Streamlit Cloud → App settings → Secrets 設定 APP_PASSWORD。")
        st.stop()
        return

    # [SEC P3] Session 逾時提示
    if st.session_state.get("session_timeout_msg"):
        st.warning("⏱️ 閒置超過 120 分鐘，已自動登出。請重新登入。")
        st.session_state.session_timeout_msg = False

    col_l, col_m, col_r = st.columns([1, 1.2, 1])
    with col_m:
        company_sub = f"<br>{COMPANY_NAME_SAFE}" if COMPANY_NAME_SAFE else ""
        st.markdown(f"""
        <div class="login-wrapper">
            <div class="login-logo">🧾</div>
            <div class="login-title">發票智能辨識系統</div>
            <div class="login-sub">Finance Invoice OCR System{company_sub}</div>
        </div>
        """, unsafe_allow_html=True)

        # 先檢查全站鎖定（P1），再檢查 per-session 鎖定
        globally_locked, global_remaining = _is_globally_locked()
        session_locked, session_remaining  = _is_session_locked()

        if globally_locked:
            st.markdown(
                f'<div class="status-box status-error">'
                f'🔒 系統偵測到異常登入活動，暫時鎖定 <b>{global_remaining}</b> 秒。'
                f'<br><small>如有疑問請聯絡管理員。</small></div>',
                unsafe_allow_html=True,
            )
        elif session_locked:
            st.markdown(
                f'<div class="status-box status-error">'
                f'🔒 登入已暫時鎖定，請 <b>{session_remaining}</b> 秒後再試。'
                f'<br><small>連續錯誤 {MAX_LOGIN_ATTEMPTS} 次將觸發鎖定保護。</small></div>',
                unsafe_allow_html=True,
            )
        else:
            with st.form("login_form"):
                password = st.text_input(
                    "存取密碼",
                    type="password",
                    placeholder="請輸入系統密碼",
                )
                submitted = st.form_submit_button("🔐 登入系統", use_container_width=True, type="primary")

            if submitted:
                # [SEC P3] timing-safe 密碼比對（防 timing attack）
                pw_ok = hmac.compare_digest(
                    password.encode("utf-8"),
                    APP_PASSWORD.encode("utf-8"),
                )
                if pw_ok:
                    st.session_state.authenticated       = True
                    st.session_state.login_error         = False
                    st.session_state.last_active_at      = datetime.now()
                    _reset_login_state()
                    st.rerun()
                else:
                    _record_session_failure()
                    _record_global_failure()    # [SEC P1] 全站計數
                    st.session_state.login_error = True

            if st.session_state.login_error:
                locked_now, _ = _is_session_locked()
                if locked_now:
                    st.markdown(
                        f'<div class="status-box status-error">'
                        f'🔒 連續錯誤 {MAX_LOGIN_ATTEMPTS} 次，帳號暫時鎖定 {LOCKOUT_SECONDS} 秒。</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    remaining_attempts = MAX_LOGIN_ATTEMPTS - st.session_state.login_attempts
                    st.markdown(
                        f'<div class="status-box status-error">'
                        f'❌ 密碼錯誤，請重新輸入。（還有 {remaining_attempts} 次機會）</div>',
                        unsafe_allow_html=True,
                    )

        if ADMIN_NOTE_SAFE:
            st.markdown(
                f'<div class="status-box status-info" style="margin-top:1rem">'
                f'ℹ️ {ADMIN_NOTE_SAFE}</div>',  # [SEC P0] 已 escape
                unsafe_allow_html=True,
            )


# ─────────────────────────────────────────
# Claude AI 發票辨識
# ─────────────────────────────────────────
SYSTEM_PROMPT = """你是一位專業的財務發票辨識 AI 助手，專門協助台灣財務人員處理各類發票。
請精準辨識發票圖片上的所有欄位，以結構化 JSON 格式回傳。

規則：
- 若某欄位無法辨識或不存在，填入空字串 ""
- 金額去除千分位逗號，只保留數字（12,500 → 12500）
- 日期統一為 YYYY/MM/DD（民國年請轉換：113/01/15 → 2024/01/15）
- 只回傳 JSON，不要包含 markdown 或其他說明文字
"""

EXTRACT_PROMPT = """請辨識這張發票圖片，並提取以下欄位，以 JSON 格式回傳：

{
  "invoice_type": "台灣二聯發票|台灣三聯發票|一般收據|自製發票|國外發票|電子發票|其他",
  "invoice_number": "發票號碼",
  "invoice_date": "YYYY/MM/DD",
  "seller_name": "賣方名稱",
  "seller_tax_id": "賣方統一編號（8位）",
  "buyer_name": "買方名稱",
  "buyer_tax_id": "買方統一編號（8位）",
  "currency": "TWD|USD|EUR|JPY|其他",
  "items": [
    {
      "description": "品項名稱",
      "quantity": "數量",
      "unit": "單位",
      "unit_price": "單價",
      "amount": "小計"
    }
  ],
  "subtotal": "未稅金額",
  "tax_rate": "稅率（如5代表5%）",
  "tax_amount": "稅額",
  "total_amount": "含稅總金額",
  "payment_method": "現金|轉帳|信用卡|支票|其他",
  "notes": "備註",
  "confidence": 0~100,
  "issues": "辨識困難說明"
}"""


def encode_image(image_bytes: bytes) -> str:
    return base64.standard_b64encode(image_bytes).decode("utf-8")


def _extract_response_text(response) -> str:
    parts = []
    for block in getattr(response, "content", []) or []:
        text = getattr(block, "text", "")
        if text:
            parts.append(text)
    return "\n".join(parts).strip()


def _parse_ai_json(raw_text: str) -> dict:
    cleaned = raw_text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)

    json_match = re.search(r'\{[\s\S]*\}', cleaned)
    payload = json_match.group() if json_match else cleaned
    return json.loads(payload)


def pdf_to_images(pdf_bytes: bytes) -> list:
    if not HAS_PYMUPDF:
        st.warning("⚠️ 未安裝 PyMuPDF，無法處理 PDF，請改上傳圖片格式。")
        return []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        images = []
        for page in doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            images.append(pix.tobytes("png"))
        doc.close()
        return images
    except Exception as e:
        print(f"[SECURITY LOG] pdf_to_images error: {type(e).__name__}: {e}")
        raise ValueError("PDF 解析失敗，可能為加密、損毀或不支援的檔案") from e


def recognize_invoice(image_bytes: bytes, filename: str = "") -> dict:
    """
    呼叫 Claude API 辨識發票。
    使用快取的 Client 物件，並在呼叫前檢查全站每日用量上限。
    """
    if not API_KEY:
        return _error_record(filename, "系統尚未設定 API 金鑰，請聯絡管理員", "")

    if not _check_and_increment_api_count():
        return _error_record(
            filename,
            f"已達今日辨識上限（{DAILY_LIMIT} 張），請明日再試或聯絡管理員調整上限",
            ""
        )

    client = _get_anthropic_client()

    ext = filename.lower().split(".")[-1] if "." in filename else "png"
    media_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg",
                 "png": "image/png", "gif": "image/gif", "webp": "image/webp"}
    media_type = media_map.get(ext, "image/png")

    try:
        response = client.messages.create(
            model="claude-opus-4-5-20251101",
            max_tokens=2000,
            system=SYSTEM_PROMPT,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": media_type,
                            "data": encode_image(image_bytes),
                        },
                    },
                    {"type": "text", "text": EXTRACT_PROMPT},
                ],
            }],
        )

        raw = _extract_response_text(response)
        data = _parse_ai_json(raw)
        data.setdefault("items", [])
        data["source_file"]   = filename
        data["recognized_at"] = datetime.now().strftime("%Y/%m/%d %H:%M")
        data["status"]        = "✅ 完成"
        return data

    except json.JSONDecodeError:
        return _error_record(filename, "AI 回應格式異常，請重試", "")
    except KeyError:
        return _error_record(filename, "回傳資料結構不符，請重試", "")
    except anthropic.AuthenticationError:
        return _error_record(filename, "API 金鑰無效，請聯絡管理員重新設定", "")
    except anthropic.RateLimitError:
        return _error_record(filename, "API 呼叫頻率超限，請稍後再試", "")
    except Exception as e:
        # [SEC P3] 例外細節只記錄到伺服器 log，不暴露給使用者
        print(f"[SECURITY LOG] recognize_invoice error ({filename}): {type(e).__name__}: {e}")
        return _error_record(filename, "辨識失敗，請稍後再試或聯絡管理員", "")


def _error_record(filename: str, issue: str, notes: str) -> dict:
    return {
        "source_file": filename, "recognized_at": datetime.now().strftime("%Y/%m/%d %H:%M"),
        "status": "❌ 失敗", "invoice_type": "", "invoice_number": "",
        "invoice_date": "", "seller_name": "", "seller_tax_id": "",
        "buyer_name": "", "buyer_tax_id": "", "currency": "TWD", "items": [],
        "subtotal": "", "tax_rate": "5", "tax_amount": "", "total_amount": "",
        "payment_method": "", "notes": notes, "confidence": 0, "issues": issue,
    }


# ─────────────────────────────────────────
# 並發辨識（ThreadPoolExecutor）
# ─────────────────────────────────────────
def recognize_batch(tasks: list[tuple[bytes, str]],
                    progress_bar, status_msg) -> list[dict]:
    """並發辨識多張發票（最多 5 個同時進行）"""
    results_map: dict[int, dict] = {}
    total = len(tasks)
    completed_count = 0
    max_workers = min(5, total)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_idx = {
            executor.submit(recognize_invoice, img_b, fname): idx
            for idx, (img_b, fname) in enumerate(tasks)
        }

        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            try:
                results_map[idx] = future.result()
            except Exception as e:
                _, fname = tasks[idx]
                print(f"[SECURITY LOG] recognize_batch exception ({fname}): {type(e).__name__}: {e}")
                results_map[idx] = _error_record(fname, "辨識失敗，請稍後再試", "")

            completed_count += 1
            progress_bar.progress(completed_count / total)
            status_msg.text(f"⏳ 辨識中 ({completed_count}/{total})…")

    return [results_map[i] for i in range(total)]


# ─────────────────────────────────────────
# 輔助函式
# ─────────────────────────────────────────
def _opt_idx(options: list, val: str, default: int = 0) -> int:
    try:
        return options.index(val) if val in options else default
    except ValueError:
        return default


def _safe_amount_for_metrics(value) -> float:
    try:
        return float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return 0.0


def _currency_totals(invoices: list[dict]) -> dict[str, float]:
    totals: dict[str, float] = {}
    for inv in invoices:
        currency = (str(inv.get("currency", "TWD")).strip() or "TWD").upper()
        totals[currency] = totals.get(currency, 0.0) + _safe_amount_for_metrics(inv.get("total_amount", 0))
    return totals


def _apply_accounting_defaults(record: dict, dept: str, acct: str, proj: str) -> None:
    if not record.get("department"):
        record["department"] = dept
    if not record.get("account_category"):
        record["account_category"] = acct
    if not record.get("project_code"):
        record["project_code"] = proj


def render_usage_guide() -> None:
    st.markdown("### 📘 平台使用說明")
    st.caption("給財務人員與業務同仁的快速上手指南，建議第一次使用時先看一遍。")

    for section in USAGE_GUIDE:
        with st.expander(section["title"], expanded=(section["title"] == USAGE_GUIDE[0]["title"])):
            for item in section["body"]:
                st.markdown(f"- {item}")

    st.markdown("### ✅ 使用建議")
    for tip in BEST_PRACTICES:
        st.markdown(f"- {tip}")

    st.info(
        "若辨識後仍需大量人工修正，通常代表上傳影像品質、票據格式差異或會計欄位規則需要進一步優化。"
    )


def render_release_notes() -> None:
    st.markdown("### 🆕 版本更新")
    st.caption("平台內直接查看版本差異，方便財務、IT 與管理者同步變更內容。")

    for entry in PLATFORM_RELEASE_NOTES:
        with st.expander(f"{entry['version']}｜{entry['date']}", expanded=(entry["version"] == APP_VERSION)):
            for item in entry["items"]:
                st.markdown(f"- {item}")


# ─────────────────────────────────────────
# 主應用介面
# ─────────────────────────────────────────
def main_app():
    # ── 側邊欄 ────────────────────────────────
    with st.sidebar:
        st.markdown(
            f'<div class="user-badge">🏢 <b>{COMPANY_NAME_SAFE or "發票辨識系統"}</b><br>'
            f'<span style="color:#64748b;font-size:.82rem;">雲端版 v{APP_VERSION}｜Claude AI</span></div>',
            unsafe_allow_html=True,
        )

        st.markdown("### ⚙️ 會計預設設定")
        default_dept = st.selectbox("預設部門", DEPT_OPTIONS)
        default_acct = st.selectbox("預設費用科目", ACCT_OPTIONS)
        default_proj = st.text_input("預設專案代碼", placeholder="P2024-001")

        st.markdown("---")
        st.markdown("### 📈 本次統計")
        count = len(st.session_state.invoices)
        totals_by_currency = _currency_totals(st.session_state.invoices)
        twd_total = totals_by_currency.get("TWD", 0.0)
        st.metric("已辨識", f"{count} 張")
        st.metric("TWD 合計", f"NT$ {twd_total:,.0f}")
        non_twd_totals = {k: v for k, v in totals_by_currency.items() if k != "TWD" and v}
        if non_twd_totals:
            extras = "｜".join(f"{ccy} {amt:,.2f}" for ccy, amt in sorted(non_twd_totals.items()))
            st.caption(f"其他幣別另計：{extras}")

        # ── 每日 API 用量（全站共用）────────────
        st.markdown("---")
        today_used = _get_today_api_count()
        pct = today_used / DAILY_LIMIT if DAILY_LIMIT else 0
        st.markdown("### 🔢 今日用量（全站）")
        st.progress(min(pct, 1.0))
        st.caption(f"{today_used} / {DAILY_LIMIT} 張（{pct*100:.0f}%）")
        if pct >= 0.9:
            st.markdown(
                '<div class="status-box status-warn" style="font-size:.82rem">'
                '⚠️ 今日用量接近上限，請聯絡管理員</div>',
                unsafe_allow_html=True,
            )

        st.markdown("---")
        col_a, col_b = st.columns(2)
        with col_a:
            if not st.session_state.confirm_clear:
                if st.button("🗑️ 清除", use_container_width=True):
                    st.session_state.confirm_clear = True
                    st.rerun()
            else:
                st.warning("確定清除所有資料？")
                cc1, cc2 = st.columns(2)
                with cc1:
                    if st.button("✅ 確定", use_container_width=True):
                        st.session_state.invoices = []
                        st.session_state.confirm_clear = False
                        st.rerun()
                with cc2:
                    if st.button("❌ 取消", use_container_width=True):
                        st.session_state.confirm_clear = False
                        st.rerun()
        with col_b:
            if st.button("🚪 登出", use_container_width=True):
                st.session_state.authenticated = False
                st.session_state.invoices = []
                st.rerun()

        st.markdown("---")
        if API_KEY:
            st.markdown(
                '<div class="status-box status-success" style="font-size:.82rem">✅ API 金鑰已設定</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="status-box status-error" style="font-size:.82rem">'
                '⚠️ 未設定 API 金鑰<br>請聯絡管理員設定 Secrets</div>',
                unsafe_allow_html=True,
            )

    # ── 頁面標題 ──────────────────────────────
    company_display = f" ｜ {COMPANY_NAME_SAFE}" if COMPANY_NAME_SAFE else ""
    st.markdown(f"""
    <div class="main-header">
        <h1>🧾 發票智能辨識系統</h1>
        <p>Finance Invoice OCR System{company_display}
           ｜ 支援台灣統一發票、一般收據、國外發票、電子發票 ｜ v{APP_VERSION}</p>
    </div>
    """, unsafe_allow_html=True)

    tab_upload, tab_review, tab_export, tab_help, tab_updates = st.tabs(
        ["📤 上傳辨識", "📋 審核資料", "📥 匯出 Excel", "📘 使用說明", "🆕 版本更新"]
    )

    # ══════════════════════════════════════════
    # TAB 1：上傳辨識
    # ══════════════════════════════════════════
    with tab_upload:
        col_up, col_prev = st.columns([1, 1], gap="large")

        with col_up:
            st.markdown("### 📁 上傳發票")
            if not API_KEY:
                st.markdown(
                    '<div class="status-box status-error">'
                    '⚠️ 系統尚未設定 API 金鑰，無法辨識。請聯絡管理員。</div>',
                    unsafe_allow_html=True,
                )

            if _get_today_api_count() >= DAILY_LIMIT:
                st.markdown(
                    f'<div class="status-box status-error">'
                    f'🚫 今日辨識張數已達上限（{DAILY_LIMIT} 張），請明日再試或聯絡管理員。</div>',
                    unsafe_allow_html=True,
                )

            uploaded_files = st.file_uploader(
                "支援 PNG、JPG、JPEG、PDF（可多選，單檔上限 20MB）",
                type=["png", "jpg", "jpeg", "pdf"],
                accept_multiple_files=True,
            )

            if uploaded_files:
                st.success(f"✅ 已選取 {len(uploaded_files)} 個檔案")

            api_ok = bool(API_KEY) and (_get_today_api_count() < DAILY_LIMIT)

            if st.button(
                "🚀 開始 AI 辨識（並發模式）",
                disabled=(not uploaded_files or not api_ok),
                type="primary",
                use_container_width=True,
            ):
                progress = st.progress(0)
                status_msg = st.empty()

                tasks: list[tuple[bytes, str]] = []
                skipped = []

                for f in uploaded_files:
                    raw = f.read()

                    # [SEC P2] 大小與 magic bytes 驗證
                    err = _validate_upload(raw, f.name)
                    if err:
                        skipped.append(f"{f.name}：{err}")
                        continue

                    if f.name.lower().endswith(".pdf"):
                        try:
                            imgs = pdf_to_images(raw)
                        except ValueError as e:
                            skipped.append(f"{f.name}：{e}")
                            continue
                        for i, img_bytes in enumerate(imgs):
                            tasks.append((img_bytes, f"{f.name}_p{i+1}.png"))
                    else:
                        tasks.append((raw, f.name))

                remaining_quota = max(DAILY_LIMIT - _get_today_api_count(), 0)
                if remaining_quota <= 0:
                    st.error(f"🚫 今日可用辨識額度已用完（上限 {DAILY_LIMIT} 張）。")
                    tasks = []
                elif len(tasks) > remaining_quota:
                    overflow = len(tasks) - remaining_quota
                    tasks = tasks[:remaining_quota]
                    st.warning(
                        f"⚠️ 今日剩餘額度僅 {remaining_quota} 張，已先處理前 {remaining_quota} 張，"
                        f"其餘 {overflow} 張請稍後再試。"
                    )

                if skipped:
                    for msg in skipped:
                        st.warning(f"⚠️ 跳過：{msg}")

                if tasks:
                    results = recognize_batch(tasks, progress, status_msg)

                    for result in results:
                        if result:
                            _apply_accounting_defaults(result, default_dept, default_acct, default_proj)

                    st.session_state.invoices.extend([r for r in results if r])
                    progress.empty()
                    status_msg.empty()

                    success_n = sum(1 for r in results if "完成" in r.get("status", ""))
                    fail_n = len(results) - success_n
                    msg = f"✅ 辨識完成：成功 <b>{success_n}</b> 張"
                    if fail_n:
                        msg += f"，失敗 <b>{fail_n}</b> 張（請在審核頁確認原因）"
                    msg += "。請切換至「審核資料」頁籤確認。"
                    st.markdown(
                        f'<div class="status-box status-success">{msg}</div>',
                        unsafe_allow_html=True,
                    )

        with col_prev:
            st.markdown("### 🖼️ 圖片預覽")
            if uploaded_files:
                idx = st.selectbox(
                    "選擇預覽",
                    range(len(uploaded_files)),
                    format_func=lambda i: uploaded_files[i].name,
                )
                f = uploaded_files[idx]
                f.seek(0)
                raw = f.read()
                if f.name.lower().endswith(".pdf") and HAS_PYMUPDF:
                    imgs = pdf_to_images(raw)
                    if imgs:
                        st.image(imgs[0], caption=f"{f.name}（第1頁）", use_container_width=True)
                else:
                    st.image(raw, caption=f.name, use_container_width=True)
            else:
                st.markdown("""
                <div style="border:2px dashed #cbd5e1;border-radius:12px;
                            padding:3.5rem 2rem;text-align:center;color:#94a3b8;">
                    <div style="font-size:3rem">🧾</div>
                    <div>請先上傳發票圖片或 PDF</div>
                </div>
                """, unsafe_allow_html=True)

    # ══════════════════════════════════════════
    # TAB 2：審核資料
    # ══════════════════════════════════════════
    with tab_review:
        if not st.session_state.invoices:
            st.info("📂 尚無辨識資料。請先至「上傳辨識」頁籤上傳發票。")
        else:
            st.markdown(f"### 📋 已辨識發票（共 {len(st.session_state.invoices)} 張）")
            st.caption("展開各項目可直接編輯欄位，修改後資料將自動保留至匯出。")

            del_idx = None
            del_confirm_idx = st.session_state.get("del_confirm_idx", None)

            for i, inv in enumerate(st.session_state.invoices):
                conf = int(inv.get("confidence", 0) or 0)
                icon = "🟢" if conf >= 80 else "🟡" if conf >= 50 else "🔴"
                title = (
                    f"{icon} #{i+1} ｜ "
                    f"{inv.get('seller_name', '（未知廠商）')[:16]} ｜ "
                    f"{inv.get('invoice_date', '')} ｜ "
                    f"NT$ {inv.get('total_amount', 0):>8} ｜ 信心度 {conf}%"
                )

                with st.expander(title, expanded=(i == len(st.session_state.invoices) - 1)):
                    c1, c2 = st.columns(2)

                    with c1:
                        st.markdown("**📌 基本資訊**")
                        inv["invoice_type"] = st.selectbox(
                            "發票類型", TYPE_OPTIONS,
                            index=_opt_idx(TYPE_OPTIONS, inv.get("invoice_type"), 6),
                            key=f"type_{i}")
                        inv["invoice_number"] = st.text_input("發票號碼", inv.get("invoice_number",""), key=f"num_{i}")
                        inv["invoice_date"]   = st.text_input("日期 (YYYY/MM/DD)", inv.get("invoice_date",""), key=f"date_{i}")
                        inv["seller_name"]    = st.text_input("賣方名稱", inv.get("seller_name",""), key=f"sell_{i}")
                        inv["seller_tax_id"]  = st.text_input("賣方統編", inv.get("seller_tax_id",""), key=f"stax_{i}")
                        inv["buyer_name"]     = st.text_input("買方名稱", inv.get("buyer_name",""), key=f"buy_{i}")
                        inv["buyer_tax_id"]   = st.text_input("買方統編", inv.get("buyer_tax_id",""), key=f"btax_{i}")

                    with c2:
                        st.markdown("**💰 金額與會計**")
                        inv["currency"] = st.selectbox(
                            "幣別", CURR_OPTIONS,
                            index=_opt_idx(CURR_OPTIONS, inv.get("currency","TWD")),
                            key=f"cur_{i}")
                        inv["subtotal"]      = st.text_input("未稅金額", str(inv.get("subtotal","")), key=f"sub_{i}")
                        inv["tax_rate"]      = st.text_input("稅率 (%)", str(inv.get("tax_rate","5")), key=f"tr_{i}")
                        inv["tax_amount"]    = st.text_input("稅額", str(inv.get("tax_amount","")), key=f"tax_{i}")
                        inv["total_amount"]  = st.text_input("含稅總金額 ★", str(inv.get("total_amount","")), key=f"tot_{i}")
                        inv["payment_method"] = st.selectbox(
                            "付款方式", PAY_OPTIONS,
                            index=_opt_idx(PAY_OPTIONS, inv.get("payment_method","")),
                            key=f"pay_{i}")

                        st.markdown("**📊 會計分類**")
                        acct_val = inv.get("account_category","")
                        acct_opts = ACCT_OPTIONS.copy()
                        if acct_val and acct_val not in acct_opts:
                            acct_opts.append(acct_val)
                        inv["account_category"] = st.selectbox(
                            "費用科目", acct_opts,
                            index=_opt_idx(acct_opts, acct_val),
                            key=f"acct_{i}")
                        inv["department"] = st.selectbox(
                            "部門", DEPT_OPTIONS,
                            index=_opt_idx(DEPT_OPTIONS, inv.get("department","")),
                            key=f"dept_{i}")
                        inv["project_code"] = st.text_input("專案代碼", inv.get("project_code",""), key=f"proj_{i}")

                    # 品項明細
                    st.markdown("**📦 品項明細**")
                    items = inv.get("items") or []
                    if items:
                        item_df = pd.DataFrame(items)
                        for col in ["description","quantity","unit","unit_price","amount"]:
                            if col not in item_df.columns:
                                item_df[col] = ""
                        edited = st.data_editor(
                            item_df[["description","quantity","unit","unit_price","amount"]],
                            column_config={
                                "description": st.column_config.TextColumn("品項說明", width="large"),
                                "quantity":    st.column_config.NumberColumn("數量",   width="small"),
                                "unit":        st.column_config.TextColumn("單位",     width="small"),
                                "unit_price":  st.column_config.NumberColumn("單價",   width="medium"),
                                "amount":      st.column_config.NumberColumn("小計",   width="medium"),
                            },
                            use_container_width=True, num_rows="dynamic", key=f"items_{i}",
                        )
                        inv["items"] = edited.to_dict("records")
                    else:
                        st.caption("（此發票無品項明細記錄）")

                    inv["notes"] = st.text_area("備註", inv.get("notes",""), height=55, key=f"note_{i}")

                    # [SEC P0] AI 回應的 issues 欄位：html.escape() 防 XSS
                    if inv.get("issues"):
                        safe_issues = html_module.escape(str(inv["issues"]))
                        st.markdown(
                            f'<div class="status-box status-warn">⚠️ 辨識備註：{safe_issues}</div>',
                            unsafe_allow_html=True)

                    # ── 刪除：二次確認 ────────────────────────
                    st.markdown("---")
                    if del_confirm_idx == i:
                        safe_seller = html_module.escape(str(inv.get('seller_name', '此筆')))
                        st.warning(f"⚠️ 確定要刪除「{safe_seller}」的發票記錄？刪除後無法復原。")
                        dc1, dc2, dc3 = st.columns([1, 1, 3])
                        with dc1:
                            if st.button("🗑️ 確定刪除", key=f"del_confirm_{i}", type="primary"):
                                del_idx = i
                                st.session_state.del_confirm_idx = None
                        with dc2:
                            if st.button("取消", key=f"del_cancel_{i}"):
                                st.session_state.del_confirm_idx = None
                                st.rerun()
                    else:
                        if st.button(f"🗑️ 刪除此筆", key=f"del_{i}"):
                            st.session_state.del_confirm_idx = i
                            st.rerun()

            if del_idx is not None:
                st.session_state.invoices.pop(del_idx)
                st.rerun()

    # ══════════════════════════════════════════
    # TAB 3：匯出 Excel
    # ══════════════════════════════════════════
    with tab_export:
        if not st.session_state.invoices:
            st.info("📂 尚無資料，請先辨識發票。")
        else:
            st.markdown("### 📥 匯出 Excel 報表")

            c1, c2 = st.columns(2)
            with c1:
                now_ym = datetime.now().strftime("%Y%m")
                title   = st.text_input("報表標題", value=f"發票明細報表_{now_ym}")
                company = st.text_input("公司名稱", value=COMPANY_NAME)
            with c2:
                p_start   = st.text_input("期間（起）", placeholder="2024/01/01")
                p_end     = st.text_input("期間（迄）", placeholder="2024/01/31")
                preparer  = st.text_input("製表人", value="")

            st.markdown("---")

            st.markdown("#### 📊 資料預覽")
            preview = [{
                "發票號碼":   inv.get("invoice_number",""),
                "日期":       inv.get("invoice_date",""),
                "廠商名稱":   inv.get("seller_name",""),
                "費用科目":   inv.get("account_category",""),
                "部門":       inv.get("department",""),
                "幣別":       inv.get("currency","TWD"),
                "含稅總額":   inv.get("total_amount",""),
                "付款方式":   inv.get("payment_method",""),
                "信心度":     str(inv.get("confidence","")) + "%",
            } for inv in st.session_state.invoices]
            st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)

            if st.button("📊 產生 Excel 報表", type="primary", use_container_width=True):
                with st.spinner("產生中，請稍候…"):
                    excel_bytes = create_invoice_excel(
                        invoices=st.session_state.invoices,
                        report_title=title,
                        company_name=company,
                        period_start=p_start,
                        period_end=p_end,
                        preparer=preparer,
                    )

                fname = f"{title}.xlsx"
                st.download_button(
                    label="⬇️ 下載 Excel 報表",
                    data=excel_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.success(f"✅ Excel 報表已產生（{len(st.session_state.invoices)} 張發票，{len(excel_bytes):,} bytes）")

    with tab_help:
        render_usage_guide()

    with tab_updates:
        render_release_notes()


# ─────────────────────────────────────────
# 程式進入點
# ─────────────────────────────────────────
if not st.session_state.authenticated:
    login_page()
else:
    main_app()
