import random
from docxtpl import DocxTemplate, Subdoc
from docx.oxml import OxmlElement
from docx.oxml.ns import qn  # 處理 XML namespace
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT
from docx2pdf import convert  # 轉換成 PDF 用

# 讀取模板文件（請將此檔案名稱修改為你的模板檔案）
doc = DocxTemplate("附件1_定位資料回饋表_管道模板.docx")

# ------------------------------
# 生成 100 筆模擬資料
# ------------------------------
simulated_data = []
# 基本數值參考（你可以根據實際情況調整）
base_coordinate_x = 308468.8260
base_coordinate_y = 2774793.9340
base_ground_elevation = 7.878
base_pipe_burial_depth = 1.34

for i in range(1, 101):
    number = i
    type_str = "01管道點{}-實測".format(i)
    coordinate_x = base_coordinate_x + random.uniform(-50, 50)
    coordinate_y = base_coordinate_y + random.uniform(-50, 50)
    ground_elevation = base_ground_elevation + random.uniform(-0.5, 0.5)
    pipe_burial_depth = base_pipe_burial_depth + random.uniform(-0.2, 0.2)
    pipe_top_coordinate_z = ground_elevation - pipe_burial_depth
    simulated_data.append({
         "Number": number,
         "Type": type_str,
         "Coordinate_X": round(coordinate_x, 4),
         "Coordinate_Y": round(coordinate_y, 4),
         "Ground_Elevation": round(ground_elevation, 3),
         "Pipe_Burial_Depth": round(pipe_burial_depth, 2),
         "Pipe_Top_Coordinate_Z": round(pipe_top_coordinate_z, 4)
    })

# ------------------------------
# 建立子文件 (Subdoc) 與 table
# ------------------------------
subdoc = doc.new_subdoc()
num_cols = 7
table = subdoc.add_table(rows=1, cols=num_cols)

# ------------------------------
# 設定表頭內容與格式
# ------------------------------
hdr_cells = table.rows[0].cells
headers = ["編號", "種類", "座標X", "座標Y", "地盤高程", "埋管深度", "管頂座標z"]

for i, text in enumerate(headers):
    paragraph = hdr_cells[i].paragraphs[0]
    # 清除預設縮排
    paragraph.paragraph_format.left_indent = 0
    paragraph.paragraph_format.first_line_indent = 0
    run = paragraph.add_run(text)
    run.font.name = "標楷體"
    # 設定東亞字型
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "標楷體")
    # 水平置中
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # 垂直置中：設定 cell 的垂直對齊方式
    hdr_cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

# 調整表頭列的高度（設定為 30 pt）
header_row = table.rows[0]
header_row.height = Pt(30)
header_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

# ------------------------------
# 動態加入資料列（使用模擬資料 simulated_data）
# ------------------------------
for item in simulated_data:
    row_cells = table.add_row().cells
    row_cells[0].text = str(item["Number"])
    row_cells[1].text = item["Type"]
    row_cells[2].text = str(item["Coordinate_X"])
    row_cells[3].text = str(item["Coordinate_Y"])
    row_cells[4].text = str(item["Ground_Elevation"])
    row_cells[5].text = str(item["Pipe_Burial_Depth"])
    row_cells[6].text = str(item["Pipe_Top_Coordinate_Z"])
    
    # 對每個 cell 的段落清除縮排並設定字型
    for cell in row_cells:
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.left_indent = 0
            paragraph.paragraph_format.first_line_indent = 0
            for run in paragraph.runs:
                run.font.name = "標楷體"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "標楷體")

# ------------------------------
# 固定整個 table 的總寬度及設定自訂欄位寬度
# ------------------------------

# 1. 固定 table 總寬度 (10000 dxa)
tbl = table._element
tblPr_list = tbl.xpath("./w:tblPr")
if tblPr_list:
    tblPr = tblPr_list[0]
else:
    tblPr = OxmlElement("w:tblPr")
    tbl.insert(0, tblPr)

tblW_list = tblPr.xpath("./w:tblW")
if tblW_list:
    tblW = tblW_list[0]
else:
    tblW = OxmlElement("w:tblW")
    tblPr.append(tblW)
tblW.set(qn("w:w"), "10000")
tblW.set(qn("w:type"), "dxa")

# 2. 自訂每個欄位的寬度（單位 dxa）
# 以下根據公分換算得到：
# 1.16 公分 ≒ 658 dxa, 3.1 公分 ≒ 1756 dxa, 2.32 公分 ≒ 1316 dxa,
# 2.52 公分 ≒ 1429 dxa, 1.93 公分 ≒ 1094 dxa, 1.93 公分 ≒ 1094 dxa, 2.13 公分 ≒ 1208 dxa
column_widths = [658, 1756, 1316, 1429, 1094, 1094, 1208]

def set_cell_width(cell, width):
    """設定單一 cell 的寬度 (w:tcW)"""
    tc = cell._element
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(width))
    tcW.set(qn("w:type"), "dxa")

for row in table.rows:
    for idx, cell in enumerate(row.cells):
        set_cell_width(cell, column_widths[idx])

# ------------------------------
# 設定表格邊線（採用粗線）
# ------------------------------
tbl_borders = OxmlElement("w:tblBorders")
for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
    border = OxmlElement("w:" + border_name)
    # 設定邊線為實線，寬度調為 8 dxa（較粗），色彩為黑色
    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), "8")
    border.set(qn("w:space"), "0")
    border.set(qn("w:color"), "000000")
    tbl_borders.append(border)
tblPr.append(tbl_borders)

# ------------------------------
# 渲染模板並儲存 Word 文件
# ------------------------------
word_filename = "generated_doc.docx"
doc.render({"table": subdoc,"number":"1234567890-1"})
doc.save(word_filename)

# ------------------------------
# 將 Word 文件轉換為 PDF
# ------------------------------
# 注意：docx2pdf 在 Windows 上需要安裝 Microsoft Word，
# 在 macOS 上則可透過內建的功能轉換 PDF。
pdf_filename = "generated_doc.pdf"
convert(word_filename, pdf_filename)
