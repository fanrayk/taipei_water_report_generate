# photo_processor.py
import os
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Cm

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
        import os
        base = os.path.basename(filename)
        num_str = "".join(filter(str.isdigit, base))
        try:
            return int(num_str)
        except:
            return 0
    image_files.sort(key=extract_number)
    groups = [image_files[i:i+6] for i in range(0, len(image_files), 6)]
    print(f"【測量照】共找到 {len(image_files)} 張照片，分成 {len(groups)} 組。")
    template_path = os.path.join("template", "附件2模板", "附件2模板.docx")
    if not os.path.exists(template_path):
        print(f"找不到模板檔案：{template_path}")
        return
    temp_files = []
    for idx, group in enumerate(groups, start=1):
        doc = DocxTemplate(template_path)
        context = {"case_number": context_number}
        for i in range(6):
            photo_key = f"image_{i+1}"
            point_key = f"point_{i+1}"
            if i < len(group):
                context[photo_key] = InlineImage(doc, group[i], width=Cm(8.09), height=Cm(5.38))
                context[point_key] = f"編號：pt{i+1}"
            else:
                context[photo_key] = ""
                context[point_key] = ""
        doc.render(context)
        temp_docx = os.path.join(output_folder, f"temp_photos_{idx}.docx")
        doc.save(temp_docx)
        temp_files.append(temp_docx)
    merged_doc = Document(temp_files[0])
    for temp_file in temp_files[1:]:
        merged_doc.add_page_break()
        temp_doc = Document(temp_file)
        for element in temp_doc.element.body:
            merged_doc.element.body.append(element)
    merged_docx_filename = os.path.join(output_folder, f"{context_number}-附件2-測量照片.docx")
    merged_doc.save(merged_docx_filename)
    print(f"已儲存合併後的 Docx 檔案: {merged_docx_filename}")
    from docx2pdf import convert
    pdf_path = os.path.join(output_folder, f"{context_number}-附件2-測量照片.pdf")
    convert(merged_docx_filename, pdf_path)
    print(f"已儲存合併後的 PDF 檔案: {pdf_path}")

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
        import os
        base = os.path.basename(filename)
        num_str = "".join(filter(str.isdigit, base))
        try:
            return int(num_str)
        except:
            return 0
    image_files.sort(key=extract_number)
    groups = [image_files[i:i+6] for i in range(0, len(image_files), 6)]
    print(f"【讀數照】共找到 {len(image_files)} 張照片，分成 {len(groups)} 組。")
    template_path = os.path.join("template", "附件3模板", "附件3模板.docx")
    if not os.path.exists(template_path):
        print(f"找不到模板檔案：{template_path}")
        return
    temp_files = []
    for idx, group in enumerate(groups, start=1):
        doc = DocxTemplate(template_path)
        context = {"case_number": context_number}
        for i in range(6):
            photo_key = f"image_{i+1}"
            point_key = f"point_{i+1}"
            if i < len(group):
                context[photo_key] = InlineImage(doc, group[i], width=Cm(8.09), height=Cm(5.38))
                context[point_key] = f"編號：pt{i+1}"
            else:
                context[photo_key] = ""
                context[point_key] = ""
        doc.render(context)
        temp_docx = os.path.join(output_folder, f"temp_app_photos_{idx}.docx")
        doc.save(temp_docx)
        temp_files.append(temp_docx)
    merged_doc = Document(temp_files[0])
    for temp_file in temp_files[1:]:
        merged_doc.add_page_break()
        temp_doc = Document(temp_file)
        for element in temp_doc.element.body:
            merged_doc.element.body.append(element)
    merged_docx_filename = os.path.join(output_folder, f"{context_number}-附件3-記錄器資料.docx")
    merged_doc.save(merged_docx_filename)
    from docx2pdf import convert
    pdf_path = os.path.join(output_folder, f"{context_number}-附件3-記錄器資料.pdf")
    convert(merged_docx_filename, pdf_path)
    print(f"【讀數照】已儲存合併後的 PDF 檔案: {pdf_path}")
