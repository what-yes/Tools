# 批量压缩pdf
import fitz
import os


def covert2pic(zoom):
    if os.path.exists('.pdf'):       # 临时文件，需为空
         os.removedirs('.pdf')
    os.mkdir('.pdf')
    for pg in range(totaling):
        page = doc[pg]
        zoom = int(zoom)            #值越大，分辨率越高，文件越清晰
        rotate = int(0)
        print(page)
        trans = fitz.Matrix(zoom / 100.0, zoom / 100.0).preRotate(rotate)
        pm = page.getPixmap(matrix=trans, alpha=False)
      
        lurl='.pdf/%s.jpg' % str(pg+1)
        pm.writePNG(lurl)
    doc.close()

def pic2pdf(obj):
    doc = fitz.open()
    for pg in range(totaling):
        img = '.pdf/%s.jpg' % str(pg+1)
        imgdoc = fitz.open(img)                 # 打开图片
        pdfbytes = imgdoc.convertToPDF()        # 使用图片创建单页的 PDF
        os.remove(img)  
        imgpdf = fitz.open("pdf", pdfbytes)
        doc.insertPDF(imgpdf)                   # 将当前页插入文档
    if os.path.exists(obj):         # 若文件存在先删除
        os.remove(obj)
    doc.save(obj)                   # 保存pdf文件
    doc.close()


def pdfz(sor, obj, zoom):    
    covert2pic(zoom)
    pic2pdf(obj)

def compress(sor):
    global doc
    global totaling
    index = sor.find(".pdf")
    obj = sor[:index] + "_compress" + sor[index:]
    doc = fitz.open(sor) 
    totaling = doc.pageCount
    
    zoom = 200                     # 清晰度调节，缩放比率
    pdfz(sor, obj, zoom)
    os.removedirs('.pdf') 

if __name__  == "__main__":
    dir = "E:\Junior Year\软件工程\试卷\\test"
    for pdf_file in os.listdir(dir):
        if(pdf_file.endswith(".pdf")):
            compress(os.path.join(dir,pdf_file))

