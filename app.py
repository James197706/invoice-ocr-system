"""
發票智能辨識系統 — 雲端版
Invoice OCR System for Finance Teams (Cloud Edition)

部署平台：Streamlit Cloud
功能：Claude AI 視覺辨識發票 → 匯出格式化 Excel

設定說明（Streamlit Cloud Secrets）：
  ANTHROPIC_API_KEY = "sk-ant-api..."
  APP_PASSWORD      = "your-password"
  COMPANY_NAME      = "貴公司名稱"          # 選填
  ADMIN_NOTE        = "IT 部門聯絡資訊"     # 選填，登入頁顯示
"""

import streamlit as st
import anthropic
import base64
import json
import pandas as pd
import re
from datetime import datetime
from io import BytesIO

try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

from excel_exporter import create_invoice_excel

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


API_KEY      = _get_secret("ANTHROPIC_API_KEY")
APP_PASSWORD = _get_secret("APP_PASSWORD", "invoice2024")   # 預設密碼（請務必在 Secrets 更改）
COMPANY_NAME = _get_secret("COMPANY_NAME", "")
ADMIN_NOTE   = _get_secret("ADMIN_NOTE", "")


# ─────────────────────────────────────────
# Session State 初始化
# ─────────────────────────────────────────
def _init_state():
    defaults = {
        "authenticated": False,
        "invoices": [],
        "login_error": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


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

    col_l, col_m, col_r = st.columns([1, 1.2, 1])
    with col_m:
        st.markdown(f"""
        <div class="login-wrapper">
            <div class="login-logo">🧾</div>
            <div class="login-title">發票智能辨識系統</div>
            <div class="login-sub">Finance Invoice OCR System{f'<br>{COMPANY_NAME}' if COMPANY_NAME else ''}</div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("login_form"):
            password = st.text_input(
                "存取密碼",
                type="password",
                placeholder="請輸入系統密碼",
            )
            submitted = st.form_submit_button("🔐 登入系統", use_container_width=True, type="primary")

        if submitted:
            if password == APP_PASSWORD:
                st.session_state.authenticated = True
                st.session_state.login_error = False
                st.rerun()
            else:
                st.session_state.login_error = True

        if st.session_state.login_error:
            st.markdown(
                '<div class="status-box status-error">❌ 密碼錯誤，請重新輸入。</div>',
                unsafe_allow_html=True,
            )

        if ADMIN_NOTE:
            st.markdown(
                f'<div class="status-box status-info" style="margin-top:1rem">ℹ️ {ADMIN_NOTE}</div>',
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


def pdf_to_images(pdf_bytes: bytes) -> list:
    if not HAS_PYMUPDF:
        st.warning("⚠️ 未安裝 PyMuPDF，無法處理 PDF，請改上傳圖片格式。")
        return []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        images.append(pix.tobytes("png"))
    doc.close()
    return images


def recognize_invoice(image_bytes: bytes, filename: str = "") -> dict:
    """呼叫 Claude API 辨識發票"""
    if not API_KEY:
        st.error("❌ 系統尚未設定 API 金鑰，請聯絡系統管理員。")
        return {}

    client = anthropic.Anthropic(api_key=API_KEY)

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

        raw = response.content[0].text.strip()
        json_match = re.search(r'\{[\s\S]*\}', raw)
        data = json.loads(json_match.group() if json_match else raw)
        data.setdefault("items", [])
        data["source_file"]    = filename
        data["recognized_at"]  = datetime.now().strftime("%Y/%m/%d %H:%M")
        data["status"]         = "✅ 完成"
        return data

    except (json.JSONDecodeError, KeyError):
        return _error_record(filename, "AI 回應格式異常", "")
    except anthropic.AuthenticationError:
        st.error("❌ API 金鑰無效，請聯絡系統管理員重新設定。")
        return {}
    except anthropic.RateLimitError:
        st.error("⚠️ API 呼叫頻率超限，請稍候再試。")
        return {}
    except Exception as e:
        return _error_record(filename, str(e), "")


def _error_record(filename, issue, notes):
    return {
        "source_file": filename, "recognized_at": datetime.now().strftime("%Y/%m/%d %H:%M"),
        "status": "❌ 失敗", "invoice_type": "", "invoice_number": "",
        "invoice_date": "", "seller_name": "", "seller_tax_id": "",
        "buyer_name": "", "buyer_tax_id": "", "currency": "TWD", "items": [],
        "subtotal": "", "tax_rate": "5", "tax_amount": "", "total_amount": "",
        "payment_method": "", "notes": notes, "confidence": 0, "issues": issue,
    }


# ─────────────────────────────────────────
# 主應用介面
# ─────────────────────────────────────────
def main_app():
    # ── 側邊欄 ────────────────────────────────
    with st.sidebar:
        st.markdown(
            f'<div class="user-badge">🏢 <b>{COMPANY_NAME or "發票辨識系統"}</b><br>'
            f'<span style="color:#64748b;font-size:.82rem;">雲端版 v2.0｜Claude AI</span></div>',
            unsafe_allow_html=True,
        )

        st.markdown("### ⚙️ 會計預設設定")
        dept_options = ["", "總務部", "業務部", "行銷部", "研發部", "財務部",
                        "人資部", "資訊部", "管理部", "其他"]
        default_dept = st.selectbox("預設部門", dept_options)

        acct_options = ["", "辦公費", "文具用品", "交通費", "差旅費", "餐費/交際費",
                        "廣告費", "租金費用", "水電費", "電話費", "郵寄費",
                        "修繕費", "採購費用", "軟體授權費", "訓練費", "顧問費", "雜項費用"]
        default_acct = st.selectbox("預設費用科目", acct_options)

        default_proj = st.text_input("預設專案代碼", placeholder="P2024-001")

        st.markdown("---")
        st.markdown("### 📈 本次統計")
        count = len(st.session_state.invoices)
        total = sum(
            float(str(inv.get("total_amount", 0)).replace(",", "") or 0)
            for inv in st.session_state.invoices
            if str(inv.get("total_amount", "")).replace(".", "").replace(",", "").isdigit()
        )
        st.metric("已辨識", f"{count} 張")
        st.metric("合計金額", f"NT$ {total:,.0f}")

        st.markdown("---")
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("🗑️ 清除", use_container_width=True):
                st.session_state.invoices = []
                st.rerun()
        with col_b:
            if st.button("🚪 登出", use_container_width=True):
                st.session_state.authenticated = False
                st.session_state.invoices = []
                st.rerun()

        # API 金鑰狀態提示
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
    st.markdown(f"""
    <div class="main-header">
        <h1>🧾 發票智能辨識系統</h1>
        <p>Finance Invoice OCR System{f' ｜ {COMPANY_NAME}' if COMPANY_NAME else ''}
           ｜ 支援台灣統一發票、一般收據、國外發票、電子發票</p>
    </div>
    """, unsafe_allow_html=True)

    tab_upload, tab_review, tab_export = st.tabs(
        ["📤 上傳辨識", "📋 審核資料", "📥 匯出 Excel"]
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

            uploaded_files = st.file_uploader(
                "支援 PNG、JPG、JPEG、PDF（可多選）",
                type=["png", "jpg", "jpeg", "pdf"],
                accept_multiple_files=True,
            )

            if uploaded_files:
                st.success(f"✅ 已選取 {len(uploaded_files)} 個檔案")

            if st.button(
                "🚀 開始 AI 辨識",
                disabled=(not uploaded_files or not API_KEY),
                type="primary",
                use_container_width=True,
            ):
                progress = st.progress(0)
                status_msg = st.empty()

                tasks = []
                for f in uploaded_files:
                    raw = f.read()
                    if f.name.lower().endswith(".pdf"):
                        imgs = pdf_to_images(raw)
                        for i, img_bytes in enumerate(imgs):
                            tasks.append((img_bytes, f"{f.name}_p{i+1}.png"))
                    else:
                        tasks.append((raw, f.name))

                results = []
                for idx, (img_b, fname) in enumerate(tasks):
                    status_msg.text(f"⏳ 辨識中 ({idx+1}/{len(tasks)})：{fname}")
                    progress.progress((idx + 1) / len(tasks))

                    result = recognize_invoice(img_b, fname)
                    if result:
                        result.setdefault("department",       default_dept)
                        result.setdefault("project_code",     default_proj)
                        result.setdefault("account_category", default_acct)
                        results.append(result)

                st.session_state.invoices.extend(results)
                progress.empty()
                status_msg.empty()

                if results:
                    success_n = sum(1 for r in results if "完成" in r.get("status", ""))
                    fail_n = len(results) - success_n
                    msg = f"✅ 辨識完成：成功 <b>{success_n}</b> 張"
                    if fail_n:
                        msg += f"，失敗 <b>{fail_n}</b> 張（請在審核頁確認）"
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

            TYPE_OPTIONS = ["台灣二聯發票", "台灣三聯發票", "一般收據", "自製發票",
                            "國外發票", "電子發票", "其他"]
            PAY_OPTIONS  = ["", "現金", "轉帳", "信用卡", "支票", "其他"]
            DEPT_OPTIONS = ["", "總務部", "業務部", "行銷部", "研發部", "財務部",
                            "人資部", "資訊部", "管理部", "其他"]
            ACCT_OPTIONS = ["", "辦公費", "文具用品", "交通費", "差旅費", "餐費/交際費",
                            "廣告費", "租金費用", "水電費", "電話費", "郵寄費",
                            "修繕費", "採購費用", "軟體授權費", "訓練費", "顧問費", "雜項費用"]
            CURR_OPTIONS = ["TWD", "USD", "EUR", "JPY", "CNY", "其他"]

            def _opt_idx(options, val, default=0):
                try:
                    return options.index(val) if val in options else default
                except ValueError:
                    return default

            del_idx = None
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
                        # 支援自行輸入科目
                        acct_val = inv.get("account_category","")
                        if acct_val not in ACCT_OPTIONS:
                            ACCT_OPTIONS.append(acct_val)
                        inv["account_category"] = st.selectbox(
                            "費用科目", ACCT_OPTIONS,
                            index=_opt_idx(ACCT_OPTIONS, acct_val),
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

                    if inv.get("issues"):
                        st.markdown(
                            f'<div class="status-box status-warn">⚠️ 辨識備註：{inv["issues"]}</div>',
                            unsafe_allow_html=True)

                    if st.button(f"🗑️ 刪除此筆", key=f"del_{i}"):
                        del_idx = i

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

            # 預覽
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

            # 產生並提供下載
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


# ─────────────────────────────────────────
# 程式進入點
# ─────────────────────────────────────────
if not st.session_state.authenticated:
    login_page()
else:
    main_app()
