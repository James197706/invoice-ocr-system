# 發票智能辨識系統

以 Streamlit 建立的內部財務工具，支援上傳發票圖片或 PDF，呼叫 Anthropic Claude 進行 OCR 與欄位抽取，人工審核後匯出為 Excel 報表。

## 功能特色

- 密碼登入保護
- 支援 PNG、JPG、JPEG、PDF 上傳
- PDF 自動轉圖片後逐頁辨識
- 支援掃描型 PDF 同頁多張發票自動拆分辨識
- 使用 Claude 視覺模型抽取發票欄位
- 可在 Web 介面直接修正欄位與品項明細
- 匯出三個工作表的 Excel 報表
- 平台內建使用說明與版本更新頁籤
- 提供每日 API 用量上限與基本暴力破解防護

## 技術棧

- Python 3.9+
- Streamlit
- Anthropic Python SDK
- Pandas
- OpenPyXL
- PyMuPDF

## 專案結構

```text
invoice-ocr-system/
├── app.py                              # Streamlit 主程式
├── excel_exporter.py                   # Excel 匯出模組
├── requirements.txt                    # Python 相依套件
├── config.toml                         # 目前的 Streamlit 設定檔
├── CHANGELOG.md
├── ADR-001-streamlit-vs-flask.md
└── ADR-002-claude-opus-model-selection.md
```

## 需要的 Secrets

請建立 `.streamlit/secrets.toml`：

```toml
ANTHROPIC_API_KEY = "sk-ant-api-..."
APP_PASSWORD = "your-password"
COMPANY_NAME = "貴公司名稱"
ADMIN_NOTE = "如需協助請聯絡 IT"
DAILY_API_LIMIT = 200
```

說明：

- `ANTHROPIC_API_KEY`: 必填，Claude API 金鑰
- `APP_PASSWORD`: 必填，登入密碼
- `COMPANY_NAME`: 選填，顯示在登入頁與頁首
- `ADMIN_NOTE`: 選填，顯示在登入頁
- `DAILY_API_LIMIT`: 選填，每日辨識上限，預設 `200`

## 本機啟動

1. 建立虛擬環境

```bash
python3 -m venv .venv
source .venv/bin/activate
```

2. 安裝套件

```bash
pip install -r requirements.txt
```

3. 建立 Streamlit secrets 目錄與檔案

```bash
mkdir -p .streamlit
```

接著新增 `.streamlit/secrets.toml`，內容可參考上方範例。

4. 啟動服務

```bash
streamlit run app.py
```

5. 開啟瀏覽器

預設會在本機開啟：

```text
http://localhost:8501
```

## 使用流程

1. 輸入 `APP_PASSWORD` 登入
2. 在「上傳辨識」頁籤上傳發票圖片或 PDF
3. 點選「開始 AI 辨識」
4. 到「審核資料」頁籤修正欄位、金額與品項
5. 到「匯出 Excel」頁籤產生並下載報表
6. 可於「使用說明」與「版本更新」頁籤查看平台操作方式與近期異動

## Excel 輸出內容

- `發票彙總`
- `品項明細`
- `統計分析`

## 部署

此專案目前同時支援：

- 本機 / 內部環境：使用 `.streamlit/secrets.toml`
- Google Cloud Run：使用環境變數與 Secret Manager

### Cloud Run 部署方式

1. 先建立兩個 Secret Manager secrets：

- `anthropic-api-key`
- `invoice-app-password`

2. 建立或確認 Artifact Registry repository 存在，例如：

```bash
gcloud artifacts repositories create cloud-run-source-deploy \
  --repository-format=docker \
  --location=asia-east1
```

3. 使用 Cloud Build 建置並部署：

```bash
gcloud builds submit \
  --config cloudbuild.yaml \
  --substitutions=_SERVICE_NAME=invoice-ocr-system,_REGION=asia-east1
```

如需自訂映像標籤，可額外加入 `_IMAGE_TAG`，例如：

```bash
gcloud builds submit \
  --config cloudbuild.yaml \
  --substitutions=_SERVICE_NAME=invoice-ocr-system,_REGION=asia-east1,_IMAGE_TAG=release-20260320
```

4. Cloud Run 執行時會使用：

- Secret Manager:
  - `ANTHROPIC_API_KEY`
  - `APP_PASSWORD`
- 一般環境變數:
  - `COMPANY_NAME`
  - `ADMIN_NOTE`
  - `DAILY_API_LIMIT`

### Cloud Run 驗證重點

- 容器是否成功在 `$PORT` 啟動 Streamlit
- `ANTHROPIC_API_KEY` 與 `APP_PASSWORD` 是否已正確掛入
- PDF 上傳、OCR、審核資料、Excel 匯出是否可正常完成

## 目前限制

- 資料保存在 `st.session_state`，重新整理、登出或 session 結束後會消失
- 沒有資料庫，因此不保留歷史紀錄
- 沒有自動化測試
- 多幣別資料目前仍以同一份總額呈現，匯總邏輯仍可再強化

## 開發建議

- 新增測試，至少覆蓋 OCR 結果正規化、Excel 匯出與上傳驗證
- 將模型名稱、併發數、逾時與限制值改成可設定
- 補上資料持久化與操作審計紀錄
- 修正多幣別總額統計與 Streamlit 設定檔位置
