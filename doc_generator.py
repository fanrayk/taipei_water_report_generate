# doc_generator.py
import os
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT
from docx2pdf import convert
from PyPDF2 import PdfMerger
from utils import set_cell_width, chunk_list


def generate_records_doc(record, context_number, output_folder):
    template_path = os.path.join(
        "template", "附件1模板", "附件1_自主查核表_首頁模板.docx"
    )
    doc = DocxTemplate(template_path)
    doc.render(record)
    # 在檔名前加上 context_number
    docx_filename = os.path.join(output_folder, "temp_自主查核表首頁.docx")
    doc.save(docx_filename)
    pdf_path = os.path.join(output_folder, "temp_自主查核表首頁.pdf")
    convert(docx_filename, pdf_path)
    print("Records PDF 已產生：", pdf_path)
    return docx_filename,pdf_path


def generate_pipeline_doc(simulated_data, context_number, output_folder):
    template_path = os.path.join(
        "template", "附件1模板", "附件1_定位資料回饋表_管道模板.docx"
    )
    doc = DocxTemplate(template_path)
    subdoc = doc.new_subdoc()
    num_cols = 7
    table = subdoc.add_table(rows=1, cols=num_cols)
    headers = ["編號", "種類", "座標X", "座標Y", "地盤高程", "埋管深度", "管頂座標z"]
    for i, cell in enumerate(table.rows[0].cells):
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.left_indent = 0
        paragraph.paragraph_format.first_line_indent = 0
        run = paragraph.add_run(headers[i])
        run.font.name = "標楷體"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "標楷體")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    header_row = table.rows[0]
    header_row.height = Pt(30)
    header_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    for item in simulated_data:
        row_cells = table.add_row().cells
        row_cells[0].text = str(item["Number"])
        row_cells[1].text = str(item["Type"])
        row_cells[2].text = str(item["Coordinate_X"])
        row_cells[3].text = str(item["Coordinate_Y"])
        row_cells[4].text = str(item["Ground_Elevation"])
        row_cells[5].text = str(item["Pipe_Burial_Depth"])
        row_cells[6].text = str(item["Pipe_Top_Coordinate_Z"])
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.left_indent = 0
                paragraph.paragraph_format.first_line_indent = 0
                for run in paragraph.runs:
                    run.font.name = "標楷體"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "標楷體")
    tbl = table._element
    tblPr_list = tbl.xpath("./w:tblPr")
    tblPr = tblPr_list[0] if tblPr_list else OxmlElement("w:tblPr")
    if not tblPr_list:
        tbl.insert(0, tblPr)
    tblW_list = tblPr.xpath("./w:tblW")
    tblW = tblW_list[0] if tblW_list else OxmlElement("w:tblW")
    if not tblW_list:
        tblPr.append(tblW)
    tblW.set(qn("w:w"), "10000")
    tblW.set(qn("w:type"), "dxa")
    column_widths = [658, 1756, 1316, 1429, 1094, 1094, 1208]
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            set_cell_width(cell, column_widths[idx])
    tbl_borders = OxmlElement("w:tblBorders")
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement("w:" + border_name)
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "8")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        tbl_borders.append(border)
    tblPr.append(tbl_borders)
    context = {"table": subdoc, "case_number": context_number}
    doc.render(context)
    word_filename = os.path.join(output_folder, "temp_管線.docx")
    doc.save(word_filename)
    pdf_filename = os.path.join(output_folder, "temp_管線.pdf")
    convert(word_filename, pdf_filename)
    print("管線 PDF 已產生：", pdf_filename)
    return word_filename,pdf_filename


def generate_reserved_doc(reserved_data, context_number, output_folder):
    template_path = os.path.join(
        "template", "附件1模板", "附件1_定位資料回饋表_設施物模板.docx"
    )
    doc = DocxTemplate(template_path)
    subdoc = doc.new_subdoc()
    num_cols = 7
    table = subdoc.add_table(rows=1, cols=num_cols)
    reserved_headers = [
        "編號",
        "種類",
        "座標X",
        "座標Y",
        "地盤高程",
        "埋管深度",
        "管頂座標z",
    ]
    for i, cell in enumerate(table.rows[0].cells):
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.left_indent = 0
        paragraph.paragraph_format.first_line_indent = 0
        run = paragraph.add_run(reserved_headers[i])
        run.font.name = "標楷體"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "標楷體")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    header_row = table.rows[0]
    header_row.height = Pt(30)
    header_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    for item in reserved_data:
        row_cells = table.add_row().cells
        row_cells[0].text = str(item["Number"])
        row_cells[1].text = str(item["Type"])
        row_cells[2].text = str(item["Coordinate_X"])
        row_cells[3].text = str(item["Coordinate_Y"])
        row_cells[4].text = str(item["Ground_Elevation"])
        row_cells[5].text = str(item["Pipe_Burial_Depth"])
        row_cells[6].text = str(item["Pipe_Top_Coordinate_Z"])
        for cell in row_cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.left_indent = 0
                paragraph.paragraph_format.first_line_indent = 0
                for run in paragraph.runs:
                    run.font.name = "標楷體"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "標楷體")
    tbl = table._element
    tblPr_list = tbl.xpath("./w:tblPr")
    tblPr = tblPr_list[0] if tblPr_list else OxmlElement("w:tblPr")
    if not tblPr_list:
        tbl.insert(0, tblPr)
    tblW_list = tblPr.xpath("./w:tblW")
    tblW = tblW_list[0] if tblW_list else OxmlElement("w:tblW")
    if not tblW_list:
        tblPr.append(tblW)
    tblW.set(qn("w:w"), "10000")
    tblW.set(qn("w:type"), "dxa")
    column_widths = [658, 1756, 1316, 1429, 1094, 1094, 1208]
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            set_cell_width(cell, column_widths[idx])
    tbl_borders = OxmlElement("w:tblBorders")
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement("w:" + border_name)
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "8")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        tbl_borders.append(border)
    tblPr.append(tbl_borders)
    context = {"table": subdoc, "cast_number": context_number}
    doc.render(context)
    docx_filename = os.path.join(output_folder, "設施物.docx")
    doc.save(docx_filename)
    pdf_filename = os.path.join(output_folder, "設施物.pdf")
    convert(docx_filename, pdf_filename)
    print("設施物 PDF 已產生：", pdf_filename)
    return docx_filename,pdf_filename


def merge_pdf_files(pdf_files, merged_pdf_filename):
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)
    merger.write(merged_pdf_filename)
    merger.close()
    print("PDF 合併完成，最終 PDF 檔案：", merged_pdf_filename)


def generate_image_doc(folder_path, context_number, output_folder):
    template_path = os.path.join("template", "附件4模板", "附件4模板.docx")
    plane_folder = os.path.join(folder_path, "平面圖")
    if not os.path.exists(plane_folder):
        print(f"找不到資料夾：{plane_folder}")
        return None
    valid_exts = (".png", ".jpg", ".jpeg", ".bmp", ".gif")
    image_files = []
    for f in os.listdir(plane_folder):
        if f.lower().endswith(valid_exts):
            name_no_ext, _ = os.path.splitext(f)
            if name_no_ext.isdigit():
                image_files.append(os.path.join(plane_folder, f))
    if not image_files:
        print("在『平面圖』資料夾中找不到符合的照片。")
        return None

    def extract_number(filename):
        base = os.path.basename(filename)
        try:
            return int(os.path.splitext(base)[0])
        except:
            return 0

    image_files.sort(key=extract_number)
    image_groups = [image_files[i : i + 2] for i in range(0, len(image_files), 2)]
    print(f"【平面圖】共找到 {len(image_files)} 張照片，分成 {len(image_groups)} 組。")
    if not os.path.exists(template_path):
        print(f"找不到模板檔案：{template_path}")
        return None
    temp_pages = []
    for idx, group in enumerate(image_groups, start=1):
        doc = DocxTemplate(template_path)
        subdoc = doc.new_subdoc()
        img_table = subdoc.add_table(rows=2, cols=1)
        for i in range(2):
            cell = img_table.cell(i, 0)
            paragraph = cell.paragraphs[0]
            if i < len(group):
                run = paragraph.add_run()
                run.add_picture(group[i], width=Cm(12), height=Cm(10))
            else:
                paragraph.add_run("")
        context = {"case_number": context_number, "table": subdoc}
        doc.render(context)
        temp_page = os.path.join(output_folder, f"temp_image_page_{idx}.docx")
        doc.save(temp_page)
        temp_pages.append(temp_page)
    merged_doc = Document(temp_pages[0])
    for temp_file in temp_pages[1:]:
        merged_doc.add_page_break()
        temp_doc = Document(temp_file)
        for element in temp_doc.element.body:
            merged_doc.element.body.append(element)
    image_docx_path = os.path.join(output_folder, "temp_平面圖_圖片.docx")
    merged_doc.save(image_docx_path)
    print("【平面圖 - 圖片部分】已儲存:", image_docx_path)
    return image_docx_path


def generate_data_doc(
    simulated_data, reserved_data, context_number, output_folder, max_rows_per_page=10
):
    template_path = os.path.join("template", "附件4模板", "附件4模板.docx")
    if not os.path.exists(template_path):
        print(f"找不到模板檔案：{template_path}")
        return None
    combined_data = simulated_data + reserved_data
    data_chunks = chunk_list(combined_data, max_rows_per_page)
    print(
        f"【平面圖 - 資料部分】共分成 {len(data_chunks)} 個資料區塊（每頁最多 {max_rows_per_page} 行）。"
    )
    temp_pages = []
    for idx, chunk in enumerate(data_chunks, start=1):
        doc = DocxTemplate(template_path)
        subdoc = doc.new_subdoc()
        table = subdoc.add_table(rows=1, cols=7)
        headers = [
            "編號",
            "種類",
            "座標X",
            "座標Y",
            "地盤高程",
            "埋管深度",
            "管頂座標z",
        ]
        for j, cell in enumerate(table.rows[0].cells):
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.left_indent = 0
            paragraph.paragraph_format.first_line_indent = 0
            run = paragraph.add_run(headers[j])
            run.font.name = "標楷體"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "標楷體")
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        header_row = table.rows[0]
        header_row.height = Pt(30)
        header_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for row_data in chunk:
            row_cells = table.add_row().cells
            row_cells[0].text = str(row_data.get("Number", ""))
            row_cells[1].text = str(row_data.get("Type", ""))
            row_cells[2].text = str(row_data.get("Coordinate_X", ""))
            row_cells[3].text = str(row_data.get("Coordinate_Y", ""))
            row_cells[4].text = str(row_data.get("Ground_Elevation", ""))
            row_cells[5].text = str(row_data.get("Pipe_Burial_Depth", ""))
            row_cells[6].text = str(row_data.get("Pipe_Top_Coordinate_Z", ""))
        tbl = table._element
        tblPr_list = tbl.xpath("./w:tblPr")
        tblPr = tblPr_list[0] if tblPr_list else OxmlElement("w:tblPr")
        if not tblPr_list:
            tbl.insert(0, tblPr)
        tblW_list = tblPr.xpath("./w:tblW")
        tblW = tblW_list[0] if tblW_list else OxmlElement("w:tblW")
        if not tblW_list:
            tblPr.append(tblW)
        tblW.set(qn("w:w"), "10000")
        tblW.set(qn("w:type"), "dxa")
        tbl_borders = OxmlElement("w:tblBorders")
        for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
            border = OxmlElement("w:" + border_name)
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "8")
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), "000000")
            tbl_borders.append(border)
        tblPr.append(tbl_borders)
        column_widths = [658, 1756, 1316, 1429, 1094, 1094, 1208]
        for row in table.rows:
            for idx2, cell in enumerate(row.cells):
                set_cell_width(cell, column_widths[idx2])
        context = {"case_number": context_number, "table": subdoc}
        doc.render(context)
        temp_page = os.path.join(output_folder, f"temp_data_page_{idx}.docx")
        doc.save(temp_page)
        temp_pages.append(temp_page)
    merged_doc = Document(temp_pages[0])
    for temp_file in temp_pages[1:]:
        merged_doc.add_page_break()
        temp_doc = Document(temp_file)
        for element in temp_doc.element.body:
            merged_doc.element.body.append(element)
    data_docx_path = os.path.join(output_folder, "temp_平面圖_資料.docx")
    merged_doc.save(data_docx_path)
    print("【平面圖 - 資料部分】已儲存:", data_docx_path)
    return data_docx_path


def merge_docs(doc_paths, output_folder,filename):
    # 以第一個文件作為基礎
    merged_doc = Document(doc_paths[0])
    # 從第二個開始依序加入，每個檔案前加入分頁符
    for doc_path in doc_paths[1:]:
        merged_doc.add_page_break()
        temp_doc = Document(doc_path)
        for element in temp_doc.element.body:
            merged_doc.element.body.append(element)
    final_docx = os.path.join(output_folder, filename)
    merged_doc.save(final_docx)
    print("【平面圖】最終合併檔案已儲存:", final_docx)
    return final_docx

