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
   - 使用者需自行加入照片資料夾，結構如下：
     ```
     TEMPLATE/
     ├── 附件1模板/
     ├── 附件2模板/
     ├── 附件3模板/
     ├── 附件4模板/
     ├── 照片/
         ├── 西區/
             ├── 監工/
                 ├── a/
                     ├──1.jpg 
                     ├──2.jpg
                 ├── b/
                 ├── ...
         ├── 營業處/
             ├── 1.jpg
             ├── 2.jpg
             ├── 3.jpg
             ├── 4.jpg
     ```

4. **平面圖文件生成**：
   - 產生包含圖片的平面圖文件。
   - 產生包含測量數據的平面圖文件。
   - 合併所有平面圖內容並轉換為 PDF。

5. **清理暫存檔案**：
   - 刪除處理過程中的臨時文件，保持資料夾整潔。

## 使用方式
### Windows
```sh
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

### Linux/macOS
```sh
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

## 主要檔案與功能
| 檔案名稱 | 主要功能 |
|----------|--------|
| `main.py` | 專案主要執行流程，負責讀取 Excel、處理數據、生成報告與圖片處理。 |
| `doc_generator.py` | 產生 Word 文件，轉換 PDF，並負責表格與格式化內容。 |
| `excel_processor.py` | 讀取與解析 Excel，並處理數據分類。 |
| `photo_processor.py` | 負責照片分組與插入 Word 文件。 |
| `utils.py` | 包含輔助函數，如文件格式調整、資料拆分、清理暫存文件等。 |
| `requirements.txt` | 專案所需 Python 套件清單。 |

## 輸出結果
- **附件1-證明資料.pdf**：包含 Excel 資料解析後的完整報告。
- **附件2-測量照片.pdf**：處理後的測量照。
- **附件3-記錄器資料.pdf**：讀數照的整理報告。
- **附件4-測量結果.pdf**：包含所有平面圖與數據的最終報告。
- **所有 Word 文件** 也會存放在 `output/{案號}/` 目錄中，以便後續修改。