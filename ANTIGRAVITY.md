# AntiGravity 專案規範 - docx2jpg (Word 轉 JPG 工具)

此檔案為 `docx2jpg` 專案專屬之 AI 代理人開發規範，後續開發時請務必遵循。

## 專案概述
- **專案名稱**：docx2jpg
- **專案用途**：自動將 Word 檔案 (`.docx`) 轉換為高品質 JPG 圖片的 Python 自動化工具。
- **部署狀態**：目前無部署需求。
- **儲存庫屬性**：公開 (Public) GitHub Repository。

---

## 核心技術堆疊
- **開發語言**：Python 3.13+
- **核心依賴**：
  - `pywin32` (Windows COM Automation) 用於控制 Word 轉存 PDF。
  - `PyMuPDF` (`fitz`) 用於高品質 PDF 轉 JPG 圖片。

---

## 安全規範
- **隱私至上**：此專案用於轉檔行政或教務檔案，**嚴禁將學生的真實姓名、真實個資、身分證字號**或任何敏感個資寫入程式碼、測試檔案或 Markdown 檔案。
- **金鑰防護**：嚴禁將任何 API Key、Token、密碼等敏感憑證 commit 至 GitHub。

---

## 每日工作流程

### 開工步驟 (說「開工」時執行)
1. 讀取此 `ANTIGRAVITY.md` 檔案確認專案規範。
2. 讀取 2ndbrain 中的專案駕駛艙 `docx2jpg-專案駕駛艙.md` 掌握進度。
3. 執行 `git status` 與檢視最近一筆 commit 掌握目前分支狀態。
4. 向使用者回報當前狀態，並建議下一步的具體工作。
5. **未經使用者確認，不擅自進行 pull/commit/push**。

### 收工步驟 (說「收工」時執行)
1. **隱私與敏感資訊檢查**：確認專案中無 API key、Token 與學生真實個資。
2. **更新專案筆記**：在 2ndbrain 中的專案駕駛艙記錄完成事項、踩坑紀錄與下一步規劃。
3. **ANTIGRAVITY.md 維護**：非規則改變時，不隨意修改此規範檔。
4. **檢查 git 異動**：執行 `git status` 及 `git diff` 以確認程式碼變更。
5. **精準暫存**：只 stage 本次任務相關的檔案，嚴禁無差別使用 `git add .`。
6. **提交流程**：取得使用者確認後，進行 commit 與 push，並回報同步結果。
