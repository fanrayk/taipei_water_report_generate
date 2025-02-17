import fitz  # pip install pymupdf

# 打開既有的 PDF 文件
doc = fitz.open("output\\114000042\\114000042-附件1-證明資料.pdf")

# 選擇要插入圖片的頁面（例如第一頁，索引為0）
page = doc[0]
x_start=90
y_start=700

ratio=0.6
ratio1=1

x1_start=320
y1_start=680
# 定義圖片插入的位置和大小：這裡定義了一個矩形區域
# 格式為 fitz.Rect(x0, y0, x1, y1)，單位通常是點（1/72 英寸）
rect = fitz.Rect(x_start, y_start, x_start+(180*ratio), y_start+(57*ratio))
rect1=fitz.Rect(x1_start, y1_start, x1_start+(272*ratio), y1_start+(181*ratio))
# 在指定區域插入圖片
page.insert_image(rect, filename="template\\照片\\測試\\張正侑.png")
page.insert_image(rect1, filename="template\\照片\\測試\\區處章.png")

# 保存修改後的 PDF 到一個新文件
doc.save("output.pdf")
doc.close()
