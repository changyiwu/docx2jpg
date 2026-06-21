# docx2jpg

一個將 Word 文件 (`.docx`) 轉換為高品質 JPG 圖片的 Python 自動化工具。

## 功能特色
- **高品質轉換**：採用 3x 放大倍率（約 216 DPI）渲染頁面，確保轉出的圖片文字清晰。
- **多頁面處理**：自動偵測文件頁數，單頁自動命名為 `檔名.jpg`，多頁則自動命名為 `檔名_page1.jpg`、`檔名_page2.jpg` 等。
- **低依賴且高整合**：利用 Windows 平台的 COM 自動化技術直接控制 Microsoft Word 進行 PDF 轉存，再利用內建 `PyMuPDF` 解析 PDF 轉出圖片，不需額外安裝複雜的 Poppler 等第三方二進位元件。

## 環境要求
- Windows 作業系統
- 安裝有 Microsoft Word (Office)
- Python 3.x

## 安裝依賴
```bash
pip install pywin32 pymupdf
```

## 使用方法
修改 `convert.py` 中的 `target_file` 路徑，或匯入 `convert_docx_to_jpg` 函數來進行轉換：
```python
from convert import convert_docx_to_jpg

convert_docx_to_jpg(r"C:\path\to\your\document.docx")
```
直接執行：
```bash
python convert.py
```
