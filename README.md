# Taipei Water Report Generate

## 專案概述
本專案的目標是自動生成臺北水務相關的報告，包括管道定位資料回饋表、測量結果文件以及相關圖片與數據報告。透過 Excel 檔案的讀取與處理，專案能夠生成 Word 文件並轉換為 PDF，最終輸出標準化的報告。

## 主要功能
1. **Excel 資料處理**：
   - 讀取 Excel 檔案，解析管道測量數據。
   - 根據施測方式與點位類別，將數據分為模擬數據（simulated data）與保留數據（reserved data）。

2. **文件生成**：
   - 產生包含基本資訊的首頁文件。
   - 生成管線數據回饋表（Word 與 PDF）。
   - 若有設施物數據，則生成對應的文件。
   - 合併 PDF 文件以輸出最終的證明文件。

3. **照片處理**：
   - 自動分組處理「測量照」與「讀數照」，並生成對應的 Word 與 PDF 文件。

4. **平面圖文件生成**：
   - 產生包含圖片的平面圖文件。
   - 產生包含測量數據的平面圖文件。
   - 合併所有平面圖內容並轉換為 PDF。

5. **清理暫存檔案**：
   - 刪除處理過程中的臨時文件，保持資料夾整潔。

## 主要檔案與功能
| 檔案名稱 | 主要功能 |
|----------|--------|
| `main.py` | 專案主要執行流程，負責讀取 Excel、處理數據、生成報告與圖片處理。 |
| `doc_generator.py` | 產生 Word 文件，轉換 PDF，並負責表格與格式化內容。 |
| `excel_processor.py` | 讀取與解析 Excel，並處理數據分類。 |
| `photo_processor.py` | 負責照片分組與插入 Word 文件。 |
| `utils.py` | 包含輔助函數，如文件格式調整、資料拆分、清理暫存文件等。 |
| `requirements.txt` | 專案所需 Python 套件清單。 |

## 依賴套件
專案所需的 Python 套件如下（可透過 `pip install -r requirements.txt` 安裝）：
- `docxtpl` - 模板式 Word 文件生成。
- `python-docx` - 讀寫 Word 文件。
- `docx2pdf` - Word 轉 PDF。
- `PyPDF2` - PDF 合併與處理。
- `pandas` - Excel 數據處理。
- `openpyxl` - 讀取與寫入 Excel 檔案。

## 使用方式
### 1. 安裝環境與套件
確保已安裝 Python（建議 3.6 以上版本），並執行以下指令安裝所需套件：
```sh
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
.venv\Scripts\activate  # Windows
pip install -r requirements.txt
```

### 2. 執行主程式
在終端機或命令提示字元中執行：
```sh
python main.py
```
系統會要求選擇包含 Excel 檔案的資料夾。

### 3. Excel 資料處理
- 程式將讀取選擇的 Excel 檔案並解析管道測量數據。
- 根據數據類型（模擬或保留）進行分類與處理。

### 4. 生成報告與圖片處理
- 程式會自動產生 Word 文件，並轉換為 PDF 報告。
- 若有相關照片，將進行分類與嵌入 Word 文件。

### 5. 輸出結果
- 最終報告將存放於 `output/{案號}/` 目錄下。
- 產生的主要輸出文件包括：
  - `附件1-證明資料.pdf`
  - `附件2-測量照片.pdf`
  - `附件3-記錄器資料.pdf`
  - `附件4-測量結果.pdf`

## 注意事項
- 確保 Excel 資料夾內 **僅有一個 Excel 檔案**，否則程式會終止。
- 若有使用 Windows 轉換 PDF，須確保電腦安裝 Microsoft Word。
- 產生的報告存放於 `output/{案號}/` 目錄下。