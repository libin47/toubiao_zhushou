import io
import os
import re
import unicodedata
import yaml
from PIL import Image
from docx import Document
from docx.enum.table import WD_TABLE_DIRECTION, WD_ROW_HEIGHT_RULE, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm
from docx.text.paragraph import Paragraph

from app.hidden_clean_config import HiddenCleanerConfig
# config = yaml.safe_load(open('config.yml', 'r', encoding='utf-8'))


def add_xml(p, name, value=""):
    w = OxmlElement(name)
    # value为空时，不添加子属性
    if value=="":
        pass
    # value为字符串或字典时时，添加子属性
    elif type(value)==dict:
        for key, val in value.items():
            w.set(qn(key), str(val))
    else:
        w.set(qn('w:val'), str(value))
    p.append(w)
    return w

def get_picture():
    pass

def convert_image_to_grayscale(image_bytes, config: HiddenCleanerConfig):
    # 将字节数据转换为PIL图像
    image = Image.open(io.BytesIO(image_bytes))
    # 缩放
    width, height = image.size
    if config.base.图片缩放.get() and width>config.base.图片最大宽度.get():
        new_width = config.base.图片最大宽度.get()
        new_height = int(new_width * (height / width))
        image = image.resize((new_width, new_height))
    # 转换为灰度图
    grayscale_image = image.convert("L")
    # 保存为字节数据
    output = io.BytesIO()
    grayscale_image.save(output, format="PNG")
    return output.getvalue()


def check_docx(file):
    # 检查docx文件是否存在
    if not os.path.exists(file):
        return False
    # 检查docx文件是否为空
    if os.path.getsize(file) == 0:
        return False
    return True


def is_image(paragraph):
    for run in paragraph.runs:
        if run._element.xpath('.//a:blip'):
            return True
    return False

def check_char(char, config:HiddenCleanerConfig):
    """
    检查字符是否是特殊字符
    """
    if char in config.base.特殊字符_保留.get():
        return True
    if char in config.base.特殊字符_删除.get():
        return False
    # 基本希腊字母
    if 0x0020<= ord(char) <= 0x007e:
        return True
    # 中日韩部首补充
    if 0x2E80<= ord(char) <= 0x2EFF:
        return True
    # 康熙部首
    if 0x2F00<= ord(char) <= 0x2FDF:
        return True
    # 中日韩符号和标点
    if 0x3000<= ord(char) <= 0x303F:
        return True
    # 中日韩统一表意文字扩展区A
    if 0x3400<= ord(char) <= 0x4DBF:
        return True
    # 中日韩统一表意文字
    if 0x4E00<= ord(char) <= 0x9FFF:
        return True
    # 中日韩兼容表意文字
    if 0xF900<= ord(char) <= 0xFAFF:
        return True
    # 半角及全角形式
    if 0xFF00 <= ord(char) <= 0xFFEF:
        return True
    # 中日韩统一表意文字扩展区B
    if 0x20000 <= ord(char) <= 0x2A6DF:
        return True
    # 中日韩统一表意文字扩展区C D E F I
    if 0x2A700 <= ord(char) <= 0x2EE5F:
        return True
    # 中日韩统一表意文字扩展区G H
    if 0x30000 <= ord(char) <= 0x323AF:
        return True
    return False

def set_super_char(sentence, config:HiddenCleanerConfig):
    # 特殊字符
    if config.base.删除特殊字符.get():
        result = ""
        for char in sentence:
            if check_char(char, config):
                result += char
        sentence = result
    return sentence

def set_char(sentence, config:HiddenCleanerConfig):
    # 英文标点to中文标点
    if config.base.标点转中文.get():
        for i in range(len(config.base.中英文标点字典[0])):
            sentence.replace(config.base.中英文标点字典[0][i], config.base.中英文标点字典[1][i])
    # 半角2全角
    if config.base.半角转为全角.get():
        full_width_text = ""
        for char in sentence:
            if 0x0020 <= ord(char) <= 0x007e:  # 半角字符范围
                full_width_char = unicodedata.lookup('FULLWIDTH ' + unicodedata.name(char))
                full_width_text += full_width_char
            else:
                full_width_text += char
        sentence = full_width_text
    return sentence


def set_row(row, config: HiddenCleanerConfig):
    # 删除边框样式
    # for ele in row._element.xpath('.//w:tblBorders'):
    #     ele.getparent().remove(ele)
    # 设置行高
    if config.table.style.行高方式.get()=="自适应":
        row.height_rule = WD_ROW_HEIGHT_RULE.AUTO
    elif config.table.style.行高方式.get()=="固定":
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row.height = Cm(config.table.style.行高.get())
    elif config.table.style.行高方式.get()=="最小值":
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.height = Cm(config.table.style.行高.get())
    else:
        print("行高方式错误，请检查配置文件: %s"%config.table.style.行高方式.get())


def set_cell(cell, config:HiddenCleanerConfig):
    # 删除边框样式
    # for ele in cell._element.xpath('.//w:tcBorders'):
    #     ele.getparent().remove(ele)
    # 垂直对齐
    if config.table.style.垂直对齐.get() == "居中":
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    elif config.table.style.垂直对齐.get() == "底部对齐":
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
    elif config.table.style.垂直对齐.get() == "顶部对齐":
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    else:
        print("垂直对齐方式错误，请检查配置文件: %s"%config.table.style.垂直对齐.get())


def set_table(table, config:HiddenCleanerConfig):
    # 表格对齐
    if config.table.style.对齐.get() == "左对齐":
        table.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    elif config.table.style.对齐.get() == "右对齐":
        table.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    elif config.table.style.对齐.get() == "居中":
        table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    else:
        print("表格对齐方式错误，请检查配置文件: %s"%config.table.style.对齐.get())
    # 方向
    if config.table.style.表格方向.get() == "从左到右":
        table.table_direction = WD_TABLE_DIRECTION.LTR
    elif config.table.style.表格方向.get() == "从右到左":
        table.table_direction = WD_TABLE_DIRECTION.RTL
    else:
        print("表格方向错误，请检查配置文件: %s"%config.table.style.表格方向.get())
    # 自动调整列宽
    table.autofit = config.table.style.自动调整列宽.get()
    # 边框
    for ele in table._element.xpath('.//w:tcBorders'):
        ele.getparent().remove(ele)
    for ele in table._element.xpath('.//w:tblBorders'):
        ele.getparent().remove(ele)
    border = add_xml(table._element.tblPr, 'w:tblBorders', "")
    if config.table.style.边框颜色.get()=="黑色":
        color = RGBColor(0,0,0)
    elif config.table.style.边框颜色.get()=="红色":
        color = RGBColor(255,0,0)
    else:
        color = RGBColor(0,0,0)
    add_xml(border, 'w:top', {"w:val": config.table.style.边框.get(), "w:sz": config.table.style.边框粗细.get(), "w:space": "0", "w:color": color})
    add_xml(border, 'w:left', {"w:val": config.table.style.边框.get(), "w:sz": config.table.style.边框粗细.get(), "w:space": "0", "w:color": color})
    add_xml(border, 'w:bottom', {"w:val": config.table.style.边框.get(), "w:sz": config.table.style.边框粗细.get(), "w:space": "0", "w:color": color})
    add_xml(border, 'w:right', {"w:val": config.table.style.边框.get(), "w:sz": config.table.style.边框粗细.get(), "w:space": "0", "w:color": color})
    add_xml(border, 'w:insideH', {"w:val": config.table.style.边框.get(), "w:sz": config.table.style.边框粗细.get(), "w:space": "0", "w:color": color})
    add_xml(border, 'w:insideV', {"w:val": config.table.style.边框.get(), "w:sz": config.table.style.边框粗细.get(), "w:space": "0", "w:color": color})

def set_table_paragraph(paragraph, config:HiddenCleanerConfig):
    if config.extend.清除w14样式.get():
        remove_w14_styles(paragraph)
    # 设置段落格式
    paragraph_format = paragraph.paragraph_format
    image_paragraph = is_image(paragraph)
    if image_paragraph:
        # 首行缩进2字符
        paragraph_format.first_line_indent = 0
        paragraph_format.element.pPr.ind.set(qn("w:firstLineChars"), str(int(config.table.image.首行缩进.get()) * 100))
        # 对齐方式
        if config.table.image.对齐方式.get() == "左对齐":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # 对齐方向
        elif config.table.image.对齐方式.get()  == "右对齐":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  # 对齐方向
        elif config.table.image.对齐方式.get()  == "居中":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 对齐方向
    else:
        # 首行缩进2字符
        paragraph_format.first_line_indent = 0
        paragraph_format.element.pPr.ind.set(qn("w:firstLineChars"), str(int(config.table.paragraph.首行缩进.get())*100))
        # 对齐方式
        if config.table.paragraph.对齐方式.get() =="左对齐":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # 对齐方向
        elif config.table.paragraph.对齐方式.get()=="右对齐":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  # 对齐方向
        elif config.table.paragraph.对齐方式.get()=="居中":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 对齐方向
    # 对齐到网络
    add_xml(paragraph_format.element.pPr, 'w:snapToGrid', str(config.table.paragraph.对齐网络.get()))
    # 右对齐网络
    add_xml(paragraph_format.element.pPr, 'w:adjustRightInd', str(config.table.paragraph.右对齐网络.get()))
    # 行距/行高
    if config.table.paragraph.行距方式.get() == "倍率":
        if config.table.paragraph.行距.get() == 1.5:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        elif config.table.paragraph.行距.get() == 2:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        elif config.table.paragraph.行距.get() == 1:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        elif config.table.paragraph.行距.get() == 0:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        else:
            paragraph_format.line_spacing = config.table.paragraph.行距.get()
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    elif config.table.paragraph.行距方式.get() == "固定":
        paragraph_format.line_spacing = Pt(config.table.paragraph.行距.get())
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    # 孤行控制
    paragraph_format.widow_control = True if config.table.paragraph.孤行控制.get() else False

    # 默认值
    paragraph_format.left_indent = Pt(0) # 左缩进
    paragraph_format.right_indent = Pt(0) # 右缩进
    paragraph_format.space_before = Pt(0) # 段前间距
    paragraph_format.space_after = Pt(0)  # 段后间距
    paragraph_format.keep_together = False # 避免段落被拆分到下一页
    paragraph_format.keep_with_next = False # 段落和下一段保持同一页
    paragraph_format.page_break_before = False # 段落应显示在页面顶部
    paragraph_format.tab_stops.clear_all() # 清除所有制表位
    # 删除空行
    if config.base.删除空行.get() and not image_paragraph:
        if not paragraph.text.strip():
            p = paragraph._element
            p.getparent().remove(p)


def set_table_font(run, config:HiddenCleanerConfig):
    if run._element.xpath('.//a:blip'):
        set_image(run, config)
        return
    # 设置字体样式
    # 字体
    fontname = config.table.font.字体.get()
    run.font.name = fontname
    run._element.rPr.rFonts.set(qn('w:ascii'), fontname)
    run._element.rPr.rFonts.set(qn('w:hAnsi'), fontname)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), fontname)
    # 字号
    run.font.size = Pt(config.table.font.字号.get())
    # 颜色
    if config.table.font.颜色.get()=="黑色":
        run.font.color.rgb = RGBColor(0,0,0)
    elif config.table.font.颜色.get()=="红色":
        run.font.color.rgb = RGBColor(255,0,0)
    # 对齐到网络
    run.font.snap_to_grid =  True if config.table.font.对齐到网络.get() else False
    # 字符缩放
    add_xml(run._element.rPr, "w:w", config.table.font.字符缩放.get())
    # 字符间距
    add_xml(run._element.rPr, "w:spacing", 0)
    # 字符间距调整
    add_xml(run._element.rPr, "w:kern", 0)
    # 其他
    run.font.italic = False # 斜体
    run.font.bold = False # 加粗
    run.font.underline = False # 下划线
    run.font.all_caps = False # 全大写
    run.font.double_strike = False # 双删除线
    run.font.strike = False # 删除线
    run.font.subscript = False # 下标
    run.font.superscript = False # 上标
    run.font.outline = False # 描边
    run.font.shadow = False # 阴影
    run.font.small_caps = False # 小大写
    run.font.emboss = False # 浮雕
    run.font.complex_script = False # 复杂语种
    run.font.cs_bold = False
    run.font.cs_italic = False
    if config.table.font.高亮.get():
        run.font.highlight_color = RGBColor(255, 255, 0)  # 设置高亮颜色为黄色
    else:
        run.font.highlight_color = None
    if config.base.删除空格.get():
        run.text = run.text.replace(' ', '')
    if config.base.删除制表符.get():
        run.text = run.text.replace('\t', '')  # 删除制表符
    if config.base.删除表格特殊字符.get():
        run.text = set_super_char(run.text, config)
    run.text = set_char(run.text, config)

def remove_w14_styles(paragraph):
    for ele in paragraph._element.xpath('.//w14:*'):
        ele.getparent().remove(ele)


def set_paragraph(paragraph, config:HiddenCleanerConfig):
    if config.extend.清除w14样式.get():
        remove_w14_styles(paragraph)
    # 设置段落格式
    paragraph_format = paragraph.paragraph_format
    image_paragraph = is_image(paragraph)
    if image_paragraph:
        # 首行缩进2字符
        paragraph_format.first_line_indent = 0
        paragraph_format.element.pPr.ind.set(qn("w:firstLineChars"), str(int(config.main.image.首行缩进.get()) * 100))
        # 对齐方式
        if config.main.image.对齐方式.get() == "左对齐":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # 对齐方向
        elif config.main.image.对齐方式.get()  == "右对齐":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  # 对齐方向
        elif config.main.image.对齐方式.get()  == "居中":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 对齐方向
        # 行距
        行距方式 = config.main.image.行距方式.get()
        行距 = config.main.image.行距.get()
    else:
        # 首行缩进2字符
        paragraph_format.first_line_indent = 0
        paragraph_format.element.pPr.ind.set(qn("w:firstLineChars"), str(int(config.main.paragraph.首行缩进.get())*100))
        # 对齐方式
        if config.main.paragraph.对齐方式.get() =="左对齐":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # 对齐方向
        elif config.main.paragraph.对齐方式.get()=="右对齐":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  # 对齐方向
        elif config.main.paragraph.对齐方式.get()=="居中":
            paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 对齐方向
        # 行距
        行距方式 = config.main.paragraph.行距方式.get()
        行距 = config.main.paragraph.行距.get()
    # 对齐到网络
    add_xml(paragraph_format.element.pPr, 'w:snapToGrid', str(config.main.paragraph.对齐网络.get()))
    # 右对齐网络
    add_xml(paragraph_format.element.pPr, 'w:adjustRightInd', str(config.main.paragraph.右对齐网络.get()))
    # 行距/行高
    if 行距方式 == "倍率":
        if 行距 == 1.5:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        elif 行距 == 2:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.DOUBLE
        elif 行距 == 1:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        elif 行距 == 0:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        else:
            paragraph_format.line_spacing = 行距
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    elif 行距方式 == "固定":
        paragraph_format.line_spacing = Pt(行距)
        paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    # 孤行控制
    paragraph_format.widow_control = True if config.main.paragraph.孤行控制.get() else False
    # 默认值
    paragraph_format.left_indent = Pt(0) # 左缩进
    paragraph_format.right_indent = Pt(0) # 右缩进
    paragraph_format.space_before = Pt(0) # 段前间距
    paragraph_format.space_after = Pt(0)  # 段后间距
    paragraph_format.keep_together = False # 避免段落被拆分到下一页
    paragraph_format.keep_with_next = False # 段落和下一段保持同一页
    paragraph_format.page_break_before = False # 段落应显示在页面顶部
    paragraph_format.tab_stops.clear_all() # 清除所有制表位
    # 删除空行
    if config.base.删除空行.get() and not image_paragraph:
        if not paragraph.text.strip():
            p = paragraph._element
            p.getparent().remove(p)



def set_font(run, config:HiddenCleanerConfig):
    if run._element.xpath('.//a:blip'):
        set_image(run, config)
        return
    # 设置字体样式
    # 字体
    fontname = config.main.font.字体.get()
    run.font.name = fontname
    run._element.rPr.rFonts.set(qn('w:ascii'), fontname)
    run._element.rPr.rFonts.set(qn('w:hAnsi'), fontname)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), fontname)
    # 字号
    run.font.size = Pt(config.main.font.字号.get())
    # 颜色
    if config.main.font.颜色.get()=="黑色":
        run.font.color.rgb = RGBColor(0,0,0)
    elif config.main.font.颜色.get()=="红色":
        run.font.color.rgb = RGBColor(255,0,0)
    # 对齐到网络
    run.font.snap_to_grid =  True if config.main.font.对齐到网络.get() else False
    # 字符缩放
    add_xml(run._element.rPr, "w:w", config.main.font.字符缩放.get())
    # 字符间距
    add_xml(run._element.rPr, "w:spacing", 0)
    # 字符间距调整
    add_xml(run._element.rPr, "w:kern", 0)
    # 其他
    run.font.italic = False # 斜体
    run.font.bold = False # 加粗
    run.font.underline = False # 下划线
    run.font.all_caps = False # 全大写
    run.font.double_strike = False # 双删除线
    run.font.strike = False # 删除线
    run.font.subscript = False # 下标
    run.font.superscript = False # 上标
    run.font.outline = False # 描边
    run.font.shadow = False # 阴影
    run.font.small_caps = False # 小大写
    run.font.emboss = False # 浮雕
    run.font.complex_script = False # 复杂语种
    run.font.cs_bold = False
    run.font.cs_italic = False
    if config.main.font.高亮.get():
        run.font.highlight_color = RGBColor(255, 255, 0)  # 设置高亮颜色为黄色
    else:
        run.font.highlight_color = None
    if config.base.删除空格.get():
        run.text = run.text.replace(' ', '')
    if config.base.删除制表符.get():
        run.text = run.text.replace('\t', '')  # 删除制表符
    if config.base.删除特殊字符.get():
        run.text = set_super_char(run.text, config)
    run.text = set_char(run.text, config)


def set_image(run, config:HiddenCleanerConfig):
    if config.base.删除图片.get():
        run._element.getparent().remove(run._element)
    else:
        if config.base.图片灰度化:
            blip_elements = run._element.xpath('.//a:blip')
            for blip in blip_elements:
                # 获取图片数据
                embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                part = run.part.related_parts[embed]
                try:
                    new_image = convert_image_to_grayscale(part.image.blob, config)
                except:
                    print("图片处理错误，尝试使用源码转换")
                    new_image = convert_image_to_grayscale(part.blob, config)
                part._blob = new_image

def set_section(section, config:HiddenCleanerConfig):
    # 设置页面布局
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_height = Cm(29.7)  # A4纸的高度
    section.page_width = Cm(21.0)  # A4纸的宽度
    # 设置页边距
    section.top_margin = Cm(config.page.页边距.上.get())  # 上边距
    section.bottom_margin = Cm(config.page.页边距.下.get())  # 下边距
    section.left_margin = Cm(config.page.页边距.左.get())  # 左边距
    section.right_margin = Cm(config.page.页边距.右.get())  # 右边距
    # 页眉页脚
    if config.core.删除页眉页脚:
        header = section.header
        header.is_linked_to_previous = True
        footer = section.footer
        footer.is_linked_to_previous = True

def set_core(doc, config:HiddenCleanerConfig):
    core_properties = doc.core_properties
    if config.core.删除文档属性.get():
        core_properties.author = ""
        core_properties.category = ""
        core_properties.comments = ""


def get_bodys(doc):
    body = doc._body._body
    ps = body.xpath('//w:p')
    for p in ps:
        yield Paragraph(p, doc._body)



def set_docx_one(config:HiddenCleanerConfig, file, process=None, count=None, tip=None):
    try:
        doc = Document(file)
    except:
        return "文件读取失败！"

    nowstep = 0
    mainstep = count*0.1
    # 作者等
    if process:
        tip.set("正在处理文件:%s  处理文档信息"%file.split("/")[-1])
    set_core(doc, config)
    if process:
        process.step(mainstep)
        process.update()
        nowstep += mainstep
    # 页面布局
    sections = doc.sections
    if process:
        tip.set("正在处理文件:%s  处理页面布局"%file.split("/")[-1])
    mainstep = count*0.1
    s_step = mainstep/len(sections) if len(sections)>0 else mainstep
    for section in sections:
        set_section(section, config)
        if process:
            process.step(s_step)
            process.update()
            nowstep += s_step
    # 封面或目录
    if config.extend.封面目录处理.get():
        paragraphs = list(get_bodys(doc))
    else:
        paragraphs = doc.paragraphs
    if process:
        tip.set("正在处理文件:%s  处理正文"%file.split("/")[-1])
    mainstep = count * 0.6
    if len(paragraphs) > 0:
        m_step = mainstep/len(paragraphs)
    else:
        process.step(mainstep)
        process.update()
        nowstep += mainstep
    # 正文
    for paragraph in paragraphs:
        if config.extend.删除自动编号.get():
            for pPr in paragraph._element.iterchildren(qn('w:pPr')):
                pPr.getparent().remove(pPr)
        set_paragraph(paragraph, config)
        for run in paragraph.runs:
            set_font(run, config)

        # 如果删除图片有可能导致有空行
        if config.base.删除图片.get() and config.base.删除空行.get() and paragraph.text.strip() == "":
            paragraph._element.getparent().remove(paragraph._element)
        if process:
            process.step(m_step)
            process.update()
            nowstep += m_step
    # 表格
    if process:
        tip.set("正在处理文件:%s  处理表格"%file.split("/")[-1])
    mainstep = count * 0.2
    if len(doc.tables) > 0:
        t_step = mainstep/len(doc.tables)
    else:
        process.step(mainstep)
        process.update()
        nowstep += mainstep
    for table in doc.tables:
        set_table(table, config)
        for row in table.rows:
            set_row(row, config)
            for cell in row.cells:
                set_cell(cell, config)
                for paragraph in cell.paragraphs:
                    set_table_paragraph(paragraph, config)
                    for run in paragraph.runs:
                        set_table_font(run, config)
        if process:
            process.step(t_step)
            process.update()
            nowstep += t_step
    # 保存
    if process:
        tip.set("正在处理文件:%s  文件保存" % file.split("/")[-1])
    try:
        doc.save(file.replace('.docx', '_new.docx'))
    except:
        return "文件保存失败！"
    # 在资源管理器中定位到文件
    dir = file.replace('.docx', '_new.docx').replace("/", '\\')
    os.system(r"explorer.exe /select, %s" % dir)
    return True

def set_docx(config:HiddenCleanerConfig, process, tip):
    files = config.file.get().split(";")
    if len(files)==0 or (len(files)==1 and files[0]==""):
        return "没有文件"

    all = len(files)
    if all==1:
        procount = [100]
    else:
        procount = [100//all for _ in range(all-1)]
        procount.append(100-sum(procount))

    index = 0
    for file in files:
        tip.set("正在处理文件:%s"%file.split("/")[-1])
        if check_docx(file):
            result = set_docx_one(config, file, process, procount[index], tip)
            if type(result)==str:
                tip.set("保存失败，请在wps/office中关闭文件并重试！")
                return result
        else:
            return "文件不存在！"
    tip.set("完成！")
    return True




if __name__ == '__main__':
    pass