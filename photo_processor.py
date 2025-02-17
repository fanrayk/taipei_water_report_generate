# photo_processor.py
import os
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from PyPDF2 import PdfMerger
from docx2pdf import convert

def photo_grouping_measured(folder_path, context_number, output_folder):
    photo_folder = os.path.join(folder_path, "測量照")
    if not os.path.exists(photo_folder):
        print(f"找不到資料夾：{photo_folder}")
        return
    valid_exts = (".png", ".jpg", ".jpeg", ".bmp", ".gif")
    image_files = [os.path.join(photo_folder, f) for f in os.listdir(photo_folder)
                   if f.lower().endswith(valid_exts) and f.startswith("pt")]
    if not image_files:
        print("在『測量照』資料夾中找不到符合的照片。")
        return

    def extract_number(filename):
        base = os.path.basename(filename)
        num_str = "".join(filter(str.isdigit, base))
        try:
            return int(num_str)
        except:
            return 0
    image_files.sort(key=extract_number)
    groups = [image_files[i:i+8] for i in range(0, len(image_files), 8)]
    print(f"【測量照】共找到 {len(image_files)} 張照片，分成 {len(groups)} 組。")
    template_path = os.path.join("template", "附件2模板", "附件2模板.docx")
    if not os.path.exists(template_path):
        print(f"找不到模板檔案：{template_path}")
        return

    temp_pdf_files = []
    for idx, group in enumerate(groups, start=0):
        doc = DocxTemplate(template_path)
        context = {"case_number": context_number}
        for i in range(8):
            photo_key = f"image_{i+1}"
            point_key = f"point_{i+1}"
            if i < len(group):
                context[photo_key] = InlineImage(doc, group[i], width=Cm(8.09), height=Cm(5))
                context[point_key] = f"編號：pt{i+1+(idx*8)}"
            else:
                context[photo_key] = ""
                context[point_key] = ""
        doc.render(context)
        temp_docx = os.path.join(output_folder, f"temp_photos_{idx}.docx")
        doc.save(temp_docx)
        # 將每個 docx 直接轉成 PDF
        temp_pdf = os.path.join(output_folder, f"temp_photos_{idx}.pdf")
        convert(temp_docx, temp_pdf)
        temp_pdf_files.append(temp_pdf)

    # 合併所有 PDF 檔案
    merger = PdfMerger()
    for pdf in temp_pdf_files:
        merger.append(pdf)
    merged_pdf_filename = os.path.join(output_folder, f"{context_number}-附件2-測量照片.pdf")
    merger.write(merged_pdf_filename)
    merger.close()
    print(f"已儲存合併後的 PDF 檔案: {merged_pdf_filename}")


def photo_grouping_app(folder_path, context_number, output_folder):
    photo_folder = os.path.join(folder_path, "讀數照")
    if not os.path.exists(photo_folder):
        print(f"找不到資料夾：{photo_folder}")
        return
    valid_exts = (".png", ".jpg", ".jpeg", ".bmp", ".gif")
    image_files = [os.path.join(photo_folder, f) for f in os.listdir(photo_folder)
                   if f.lower().endswith(valid_exts) and f.startswith("app_pt")]
    if not image_files:
        print("在『讀數照』資料夾中找不到符合的照片。")
        return

    def extract_number(filename):
        base = os.path.basename(filename)
        num_str = "".join(filter(str.isdigit, base))
        try:
            return int(num_str)
        except:
            return 0
    image_files.sort(key=extract_number)
    groups = [image_files[i:i+8] for i in range(0, len(image_files), 8)]
    print(f"【讀數照】共找到 {len(image_files)} 張照片，分成 {len(groups)} 組。")
    template_path = os.path.join("template", "附件3模板", "附件3模板.docx")
    if not os.path.exists(template_path):
        print(f"找不到模板檔案：{template_path}")
        return

    temp_pdf_files = []
    for idx, group in enumerate(groups, start=0):
        doc = DocxTemplate(template_path)
        context = {"case_number": context_number}
        for i in range(8):
            photo_key = f"image_{i+1}"
            point_key = f"point_{i+1}"
            if i < len(group):
                context[photo_key] = InlineImage(doc, group[i], width=Cm(8.09), height=Cm(5))
                context[point_key] = f"編號：pt{i+1+(idx*8)}"
            else:
                context[photo_key] = ""
                context[point_key] = ""
        doc.render(context)
        temp_docx = os.path.join(output_folder, f"temp_app_photos_{idx}.docx")
        doc.save(temp_docx)
        # 將每個 docx 直接轉成 PDF
        temp_pdf = os.path.join(output_folder, f"temp_app_photos_{idx}.pdf")
        convert(temp_docx, temp_pdf)
        temp_pdf_files.append(temp_pdf)

    # 合併所有 PDF 檔案
    merger = PdfMerger()
    for pdf in temp_pdf_files:
        merger.append(pdf)
    merged_pdf_filename = os.path.join(output_folder, f"{context_number}-附件3-記錄器資料.pdf")
    merger.write(merged_pdf_filename)
    merger.close()
    print(f"【讀數照】已儲存合併後的 PDF 檔案: {merged_pdf_filename}")
