# ADR-001：選用 Streamlit 而非 Flask / FastAPI

## 狀態：已採用（2026-03-11）

## 背景
需要快速建立一套供財務部門使用的發票辨識 Web 應用，核心功能為圖片上傳、AI 辨識、資料審核、Excel 匯出。使用族群為財務人員（非技術背景），且希望盡快上線驗證。

## 決策
選用 **Streamlit** 作為前後端框架，人員%��於 Streamlit Cloud 免費方案。

## 理由
- **零前端工程**：Streamlit 用純 Python 即可完成互動式 UI，不需要 HTML/CSS/JavaScript，符合快速上線需求。
- **免費部署**：Streamlit Cloud 免費方案支援公開 GitHub 倉庫直接部署，零維運成本。
- **內建元件齊全**：`st.file_uploader`、`st.data_editor`、`st.progress` 等元件直接滿足發票審核流程需求，無需另行開發。
- **Secrets 管理**：`st.secrets` 直接整合 API Key 管理，不需要額外搭建環境變數服務。
- **快速迭代**：修改程式碼 → push 到 GitHub → 自動重新部署，符合財務工具輕量維護的需求。

## 代價
- **無水平擴展**：Streamlit 不適合高併發場景（同時大量用戶），若將來需支援全公司 100+ 人同時使用，需評估遷移。
- **Session 為主的狀態管理**：辨識結果存在 `st.session_state`，重新整理或關閉瀏覽器後資料消失，不適合需要持久化儲存的場景。
- **無法自訂 URL 路由**：Streamlit 不支援傳統的多頁路由，所有功能集中在單一頁面的 Tab 中。
- **效能上限**：單機部署，CPU/RAM 受 Streamlit Cloud 免費方案限制。

## 排除的替代方案
- **Flask + 自定義前端**：需要額外開發 HTML/CSS/JS 前端，開發週期長 2-3 倍，不符合快速上線需求。
- **FastAPI + React**：技術棧複雜，需要前後端分離部署，成本與複雜度遠超此工具的需求規模。
- **Google Sheets + Apps Script**：財務人員最熟悉，但 Claude Vision API 整合困難，且 PDF 處理能力有限。

## 未來遷移觸發條件
若出現以下情況，應評估遷移至 FastAPI + React 架構：
- 同時在線用戶超過 20 人且出現明顯延遲
- 需要持久化儲存歷史辨識記錄（加入資料庫）
- 需要更細緻的用戶權限管理（部門隔離）
