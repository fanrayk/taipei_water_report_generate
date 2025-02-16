# main.py
import os
from excel_processor import (
    select_folder_and_excel,
    process_excel_pandas,
    process_excel_openpyxl,
    create_output_folder,
)
from doc_generator import (
    generate_records_doc,
    generate_pipeline_doc,
    generate_reserved_doc,
    merge_pdf_files,
    generate_image_doc,
    generate_data_doc,
    merge_docs,
    merge_pdfs
)
from photo_processor import photo_grouping_measured, photo_grouping_app
from utils import cleanup_temp_files
from docx2pdf import convert


def main():
    # 1. 選取資料夾與 Excel 檔案
    folder_path, excel_file_path = select_folder_and_excel()

    # 2. 讀取 Excel 資料 (pandas)
    df_renamed = process_excel_pandas(excel_file_path)
    if df_renamed.empty:
        print("Excel 資料讀取失敗，程式結束。")
        exit()
    context_number = df_renamed["case_number"].iloc[0]
    survey_point_count = df_renamed["survey_point_count"].iloc[0]

    # 3. 建立輸出資料夾
    output_folder = create_output_folder(context_number)

    # 4. 利用 openpyxl 分離資料
    simulated_data, reserved_data = process_excel_openpyxl(
        excel_file_path, survey_point_count
    )

    # 5. 產生首頁文件
    records_list = df_renamed.to_dict(orient="records")
    record = records_list[0] if records_list else {}
    records_doc, records_pdf = generate_records_doc(record, output_folder)

    # 6. 產生管線文件及（如有）設施物文件
    pipeline_doc, pipeline_pdf = generate_pipeline_doc(
        simulated_data, context_number, output_folder
    )
    pdf_list = [records_pdf, pipeline_pdf]
    if reserved_data:
        reserved_doc, reserved_pdf = generate_reserved_doc(
            reserved_data, context_number, output_folder
        )
        pdf_list.append(reserved_pdf)
    merged_pdf_filename = os.path.join(
        output_folder, f"{context_number}-附件1-證明資料.pdf"
    )
    if reserved_data:
        merge_docs(
            [records_doc, reserved_doc, pipeline_doc],
            output_folder,
            f"{context_number}-附件1-證明資料.docx",
        )
    else:
        merge_docs(
            [records_doc, pipeline_doc],
            output_folder,
            f"{context_number}-附件1-證明資料.docx",
        )

    merge_pdf_files(pdf_list, merged_pdf_filename)

    # 7. 照片分組處理
    print("========== 現在開始『測量照』照片分組處理 ==========")
    photo_grouping_measured(folder_path, context_number, output_folder)
    print("========== 『讀數照』照片分組處理 ==========")
    photo_grouping_app(folder_path, context_number, output_folder)

    print("========== 現在開始產生平面圖 - 圖片部分 ==========")
    image_docx = generate_image_doc(folder_path, context_number, output_folder)
    image_pdf = os.path.join(output_folder, f"temp_{context_number}-圖片部分.pdf")
    convert(image_docx, image_pdf)

    print("========== 現在開始產生平面圖 - 資料部分 ==========")
    data_docx = generate_data_doc(
        simulated_data,
        reserved_data,
        context_number,
        output_folder,
        max_rows_per_page=50,
    )
    data_pdf = os.path.join(output_folder, f"temp_{context_number}-資料部分.pdf")
    convert(data_docx, data_pdf)

    if image_pdf and data_pdf:
        final_pdf = os.path.join(output_folder, f"{context_number}-附件4-測量結果.pdf")
        merge_pdfs([image_pdf, data_pdf], final_pdf)
        print("【平面圖】最終 PDF 已儲存:", final_pdf)

    # 9. 刪除暫存檔案
    cleanup_temp_files(output_folder, "temp*")

    print("========== 全部流程完成 ==========")


if __name__ == "__main__":
    main()
