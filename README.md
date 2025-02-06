# 管道定位資料回饋表生成工具

這個專案程式碼主要用於生成模擬的管道定位資料，並依據模板文件建立一個格式化的 Word 文件，之後再轉換成 PDF 文件。程式中會隨機生成 100 筆包含座標、地盤高程、埋管深度等資訊的數據，並動態建立一個包含這些數據的表格。主要使用的 Python 套件有：

- [docxtpl](https://pypi.org/project/docxtpl/)：用於根據模板生成 Word 文件。
- [python-docx](https://pypi.org/project/python-docx/)：處理 Word 文件的操作（docxtpl 底層會用到）。
- [docx2pdf](https://pypi.org/project/docx2pdf/)：用於將 Word 文件轉換為 PDF 文件。

## 環境需求

- Python 3.6 或以上版本
- Windows 系統需要安裝 Microsoft Word（用於 docx2pdf 轉換 PDF），macOS 可利用內建功能轉換 PDF

## 建立虛擬環境

建議使用虛擬環境來管理專案相依套件，以下提供在不同作業系統上的建立方法：

### Windows

1. 開啟命令提示字元（Command Prompt）或 PowerShell。
2. 在專案根目錄下執行：
    ```bash
    python -m venv venv
    ```
3. 啟動虛擬環境：
    ```bash
    venv\Scripts\activate
    ```

### macOS / Linux

1. 開啟終端機（Terminal）。
2. 在專案根目錄下執行：
    ```bash
    python3 -m venv venv
    ```
3. 啟動虛擬環境：
    ```bash
    source venv/bin/activate
    ```

## 安裝必要套件

啟動虛擬環境後，請執行下列命令來安裝所有必要的 Python 套件：

```bash
pip install docxtpl python-docx docx2pdf
