# docx2jpg 專案開發筆記

## 📌 當前進度與待辦事項
- [x] 初始化專案結構 (.gitignore, README.md, ANTIGRAVITY.md)
- [x] 實作 `convert.py` 以 COM Automation 整合 PyMuPDF 轉換 Word 為 JPG
- [x] 成功將段考日程表 `.docx` 轉換為同名的高解析度 `.jpg`
- [x] 初始化 Git 倉庫與建立 GitHub 公開 repository
- [x] 撰寫專案駕駛艙筆記並整合至 2ndbrain

## 📝 踩坑紀錄與解決方案
- **COM 串接**：Windows 環境下 Word COM 組件可用，但環境變數 `GITHUB_TOKEN` 的無效值可能干擾 `gh` CLI，需在執行 `gh` 命令前暫時清除 `GITHUB_TOKEN`。
- **解析度優化**：預設的 PDF 轉圖解析度較低，在 `PyMuPDF` 渲染時設定 Matrix 縮放倍率 `zoom = 3.0` 可以使文字邊緣極其銳利清晰。

## 🎯 下一步規劃
- [x] 專案已順利交付，當前無後續待辦事項。

