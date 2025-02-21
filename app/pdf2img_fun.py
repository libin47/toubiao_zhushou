# -*- coding: utf-8 -*-
import fitz
import os
import json
import pymupdf
from app.pdf2img_config import Pdf2imgConfig


def pdf_image(pdfPath, imgPath, config:Pdf2imgConfig, process=None, count=100 ):
    rotation_angle = 0
    zoom_x = config.quality.get()
    zoom_y = config.quality.get()
    # 打开PDF文件
    pdf = pymupdf.open(pdfPath)
    if process:
        step = count/pdf.page_count
        all = 0
    # 逐页读取PDF
    for pg in range(0, pdf.page_count):
        page = pdf[pg]
        # 设置缩放和旋转系数
        trans = fitz.Matrix(zoom_x, zoom_y)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        # 开始写图像
        pm.save(imgPath + str(pg) + ".png")
        if process:
            if all+step < count:
                all = all+step
                process.step(step)
            else:
                process.step(count-all)
            process.update()
    pdf.close()


def pdf2image(config, process, tip):
    files = config.file.get().split(";")
    # 进度条
    all = len(files)
    if all==0 or (all==1 and files[0]==""):
        return "没有文件"
    if all==1:
        procount = [100]
    else:
        procount = [100//all for i in range(all-1)]
        procount.append(100-sum(procount))
    index = 0
    for file in files:
        # process.configure(mask="正在处理文件:%s"%file)
        # process.update()
        tip.set("正在处理文件:%s"%file.split("/")[-1])

        try:
            os.mkdir(file[:-4])
        except:
            pass
        try:
            pdf_image(file, file[:-4]+"/", config, process, procount[index])
        except:
            return "发生错误！"
        index += 1
        # 在资源管理器中定位到文件
        dir = file[:-4].replace("/", '\\')
        os.system(r"explorer.exe /select, %s" % dir)
    # process.configure(text="完成！")
    # process.update_idletasks()
    tip.set("完成！")
    return True

