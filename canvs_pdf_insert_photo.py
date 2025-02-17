import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PyPDF2 import PdfReader, PdfWriter

# --- 步驟 1：使用 ReportLab 生成 overlay PDF（存放在記憶體中） ---
packet = io.BytesIO()  # 使用 BytesIO 在記憶體中建立暫存區
c = canvas.Canvas(packet, pagesize=A4)
page_width, page_height = A4

# 設定圖片資訊與位置（根據需求調整）
image_path = "example\\照片\\測試\\陳立軒.png"
ratio = 0.6
img_width = 177 * ratio  # 圖片寬度
img_height = 52 * ratio  # 圖片高度
angle = 5  # 旋轉角度（度）
center_x = 180  # 圖片放置中心 x 座標
center_y = 100  # 圖片放置中心 y 座標

# 保存狀態，移動原點到圖片中心，旋轉，並繪製圖片
c.saveState()
c.translate(center_x, center_y)
c.rotate(angle)
c.drawImage(
    image_path,
    -img_width / 2,
    -img_height / 2,
    width=img_width,
    height=img_height,
    mask="auto",
)
c.restoreState()

c.save()  # 儲存 overlay PDF 到 BytesIO 中

# 將記憶體中的 PDF 指標移回起始位置
packet.seek(0)
overlay_pdf = PdfReader(packet)

# --- 步驟 2：讀取原始 PDF 並將 overlay 疊加上去 ---
with open("output\\114000042\\114000042-附件1-證明資料.pdf", "rb") as f_old:
    original_pdf = PdfReader(f_old)
    output = PdfWriter()

    # 假設只在第一頁疊加 overlay
    for i, page in enumerate(original_pdf.pages):
        if i == 0:
            page.merge_page(overlay_pdf.pages[0])
        output.add_page(page)

    # 輸出合併後的新 PDF
    with open("merged.pdf", "wb") as f_out:
        output.write(f_out)

print("PDF 合併完成！")
