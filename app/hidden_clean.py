import os.path
# import tkinter as tk
import tkinter.filedialog as tkf
# from tktooltip import ToolTip
import time
from ttkbootstrap.dialogs.dialogs import Messagebox
from ttkbootstrap.constants import *
import ttkbootstrap as ttk
from ttkbootstrap.tooltip import ToolTip
from app.hidden_clean_fun import set_docx
from app.hidden_clean_config import HiddenCleanerConfig

def tooltip(widget, text):
    """
    给组件添加提示
    """
    ToolTip(widget, text=text)
    return widget


class HiddenCleaner(ttk.Frame):
    def __init__(self, notebook:ttk.Notebook):
        super().__init__(notebook)
        self.config = HiddenCleanerConfig()
        self.tip = ttk.StringVar()
        self.tip.set("请选择文件")

        self._create()


    def _disable_start(self):
        self.start_button.config(state=DISABLED, text="处理中")
        self.start_button.update_idletasks()

    def _enable_start(self):
        self.start_button.config(state=NORMAL, text="启动")
        self.start_button.update_idletasks()


    def _create(self):
        # 选择文件
        file = ttk.Labelframe(self, text="需处理文件", bootstyle="info")
        file.pack(fill=ttk.X, padx=10, pady=5)
        ttk.Label(file, text="选择文件").pack(side=LEFT, padx=5, pady=10)
        ttk.Entry(file, textvariable=self.config.file, width=50).pack(side=LEFT, padx=5, pady=10)
        ttk.Button(file, bootstyle=INFO, text="选择", command=lambda: self._select_file()).pack(side=LEFT, padx=5, pady=10)
        self.start_button = ttk.Button(file, bootstyle=SUCCESS, text="启动", command=lambda: self.start())
        self.label = ttk.Label(self, textvariable=self.tip, bootstyle="info")
        self.label.pack(fill=ttk.X, padx=10, pady=5)
        self.start_button.pack(side=LEFT, padx=5, pady=10)
        self.process = ttk.Progressbar(self, maximum=100, value=0, bootstyle="info")
        self.process.pack(fill=ttk.X, padx=10, pady=5)
        # 添加配置卡
        self.confignote = ttk.Notebook(self, bootstyle="dark")
        self.confignote.pack(fill=ttk.X, expand=True)
        self.confignote.add(self._config_page(), text="文档配置")
        self.confignote.add(self._config_base(), text="基础设置")
        self.confignote.add(self._config_extend(), text="高级设置")
        self.confignote.add(self._config_main(), text="正文设置")
        self.confignote.add(self._config_table(), text="表格设置")

    def _config_page(self):
        frame = ttk.Frame(self)
        # 文档属性
        doc = ttk.Labelframe(frame, text="文档属性", bootstyle="dark")
        doc.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Checkbutton(doc, text="删除文档属性", variable=self.config.core.删除文档属性), "删除文档属性，包括作者、标题、主题等").grid(row=1, column=1 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(doc, text="删除页眉页脚", variable=self.config.core.删除页眉页脚), "删除页眉页脚").grid(row=1, column=2 , padx=10, pady=10)
        # 页面设置
        page = ttk.Labelframe(frame, text="页面设置", bootstyle="dark")
        page.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Label(page, text="页面距:上"),"上页边距，单位为cm").grid(row=1, column=1, padx=10, pady=5)
        ttk.Entry(page, textvariable=self.config.page.页边距.上).grid(row=1, column=2, padx=10, pady=10)
        tooltip(ttk.Label(page, text="页面距:下"), "下页边距，单位为cm").grid(row=1, column=3, padx=10, pady=5)
        ttk.Entry(page, textvariable=self.config.page.页边距.下).grid(row=1, column=4, padx=10, pady=10)
        tooltip(ttk.Label(page, text="页面距:左"), "左页边距，单位为cm").grid(row=2, column=1, padx=10, pady=5)
        ttk.Entry(page, textvariable=self.config.page.页边距.左).grid(row=2, column=2, padx=10, pady=10)
        tooltip(ttk.Label(page, text="页面距:右"), "右页边距，单位为cm").grid(row=2, column=3, padx=10, pady=5)
        ttk.Entry(page, textvariable=self.config.page.页边距.右).grid(row=2, column=4, padx=10, pady=10)
        return frame

    def _config_base(self):
        frame = ttk.Frame(self)
        # 基础配置
        base = ttk.Labelframe(frame, text="基础配置", bootstyle="dark")
        base.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Checkbutton(base, text="删除空行", variable=self.config.base.删除空行), "删除文中的所有空行").grid(row=1, column=1 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(base, text="删除空格", variable=self.config.base.删除空格), "删除文中的所有空格").grid(row=1, column=2 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(base, text="删除制表符", variable=self.config.base.删除制表符), "删除文中所有制表符").grid(row=1, column=3 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(base, text="标点转中文", variable=self.config.base.标点转中文), "将英文标点转为中文，此选项会先于[半角转全角]生效").grid(row=1, column=4 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(base, text="半角转为全角", variable=self.config.base.半角转为全角), "将英文字符标点及数字等ASCII编号字符转为全角字符").grid(row=2, column=1 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(base, text="删除图片", variable=self.config.base.删除图片, command=self._image_show), "删除文中所有图片").grid(row=2, column=2 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(base, text="删除特殊字符", variable=self.config.base.删除特殊字符, command=self._char_show),"删除文章中所有特殊字符，特殊字符指Unicode编码不在中日韩统一表意文字（及扩展区）、基本希腊字母、中日韩部首补充、康熙部首、中日韩符号和标点、半角及全角形式区域的所有字符，参考https://symbl.cc/cn/unicode-table/。如有必要，自行调整需保留或删除的字符。").grid(row=2, column=3 , padx=10, pady=10)
        # 图片处理
        self.image = ttk.LabelFrame(frame, text="图片处理")
        self.image.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Checkbutton(self.image, text="图片灰度化", variable=self.config.base.图片灰度化),"将文中所有图片转为灰度图").grid(row=1, column=1 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(self.image, text="图片压缩，允许最大宽度：", variable=self.config.base.图片缩放),"启用图片缩放，以将大于此宽度的图片同比缩放到此宽度，小于此宽度的图片不会进行缩放处理，在后方输入框输入启动压缩的最小宽度").grid(row=1, column=2 , padx=10, pady=10)
        ttk.Entry(self.image, textvariable=self.config.base.图片最大宽度).grid(row=1, column=4 , padx=10, pady=10)
        # 特殊字符
        self.char = ttk.LabelFrame(frame, text="特殊字符")
        tooltip(ttk.Checkbutton(self.char, text="删除表格中的特殊字符", variable=self.config.base.删除表格特殊字符),"表格中的特殊字符是否同样进行处理").grid(row=1, column=1 , padx=10, pady=10)
        tooltip(ttk.Label(self.char, text="需保留字符"),"不在默认保留范围内但需保留的字符，优先于[需删除字符]生效").grid(row=2, column=1 , padx=10, pady=10)
        ttk.Entry(self.char, textvariable=self.config.base.特殊字符_保留).grid(row=2, column=2 , padx=10, pady=10)
        tooltip(ttk.Label(self.char, text="需删除字符"),"在默认保留范围内但需删除的的字符").grid(row=3, column=1 , padx=10, pady=10)
        ttk.Entry(self.char, textvariable=self.config.base.特殊字符_删除).grid(row=3, column=2 , padx=10, pady=10)
        return frame

    def _config_extend(self):
        frame = ttk.Frame(self)
        # 高级设置
        extend = ttk.Labelframe(frame, text="高级设置", bootstyle="dark")
        extend.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Checkbutton(extend, text="封面目录处理", variable=self.config.extend.封面目录处理), "对封面和目录页同样进行处理").grid(row=1, column=1 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(extend, text="清除w14样式", variable=self.config.extend.清除w14样式), "新版的office中会使用w14样式，启用此选项删除所有w14样式以确保处理无误").grid(row=1, column=2 , padx=10, pady=10)
        tooltip(ttk.Checkbutton(extend, text="删除自动编号", variable=self.config.extend.删除自动编号), "删除自动编号，实际上会删除所有原来的段落属性").grid(row=1, column=3 , padx=10, pady=10)
        return frame


    def _config_main(self):
        frame = ttk.Frame(self)
        # 段落
        para = ttk.Labelframe(frame, text="段落", bootstyle="dark")
        para.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Checkbutton(para, text="孤行控制", variable=self.config.main.paragraph.孤行控制),
                "孤行控制指段首第一行在页尾或段落最后一页在页首时，强制将此孤儿寡母行放在段落页").grid(row=1, column=1, padx=10, pady=10)
        tooltip(ttk.Checkbutton(para, text="对齐网络", variable=self.config.main.paragraph.对齐网络),
                "如果定义了文档网络，则与网络对齐。简单来说，行距大于1时，字体位于行距高度的上方，开启此项或可使字体居中于行距高度，具体要看文档网络如何定义").grid(row=1, column=2, padx=10, pady=10)
        tooltip(ttk.Checkbutton(para, text="右对齐网络", variable=self.config.main.paragraph.右对齐网络),
                "如果定义了文档网络，则自动调整右缩进。如果开启了列网络，文档自动右缩进以对齐网络，一般默认应该没有开启列网络只有行网络").grid(row=1, column=3, padx=10, pady=10)
        tooltip(ttk.Label(para, text="首行缩进（字符）"),"首行缩进，单位为字符，0即为关闭").grid(row=2, column=1, padx=10, pady=10)
        ttk.Entry(para, textvariable=self.config.main.paragraph.首行缩进).grid(row=2, column=2, padx=10, pady=10)
        tooltip(ttk.Label(para, text="对齐方式"),"左对齐/右对齐/居中").grid(row=2, column=3, padx=10, pady=10)
        ttk.Combobox(para, textvariable=self.config.main.paragraph.对齐方式, value=("居中", "左对齐", "右对齐"), state="readonly").grid(row=2, column=4, padx=10, pady=10)
        tooltip(ttk.Label(para, text="行距方式"), "行距方式，倍率或固定").grid(row=3, column=1, padx=10, pady=10)
        ttk.Combobox(para, textvariable=self.config.main.paragraph.行距方式, value=("固定", "倍率"), state="readonly").grid(row=3, column=2, padx=10, pady=10)
        tooltip(ttk.Label(para, text="行距"), "倍率时为倍数，固定时为磅").grid(row=3, column=3, padx=10, pady=10)
        ttk.Entry(para, textvariable=self.config.main.paragraph.行距).grid(row=3, column=4, padx=10, pady=10)
        # 字体
        font = ttk.Labelframe(frame, text="字体", bootstyle="dark")
        font.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Label(font, text="字体"), "输入后会在系统库里查找，没有找到的话我也不确定会发生什么").grid(row=1, column=1, padx=10, pady=10)
        ttk.Entry(font, textvariable=self.config.main.font.字体).grid(row=1, column=2, padx=10, pady=10)
        tooltip(ttk.Label(font, text="字号"), "单位为磅，小四=12，四号=14").grid(row=1, column=3, padx=10, pady=10)
        ttk.Entry(font, textvariable=self.config.main.font.字号).grid(row=1, column=4, padx=10, pady=10)
        tooltip(ttk.Label(font, text="颜色"), "").grid(row=2, column=1, padx=10, pady=10)
        ttk.Combobox(font, textvariable=self.config.main.font.颜色, value=("黑色", "红色"), state="readonly").grid(row=2, column=2, padx=10, pady=10)
        tooltip(ttk.Checkbutton(font, text="高亮", variable=self.config.main.font.高亮), "是否高亮显示").grid(row=2, column=3, padx=10, pady=10)
        tooltip(ttk.Checkbutton(font, text="对齐到网络", variable=self.config.main.font.对齐到网络), "是否对齐到网络").grid(row=2, column=4, padx=10, pady=10)
        tooltip(ttk.Label(font, text="字体间距"), "字体间距").grid(row=3, column=1, padx=10, pady=10)
        ttk.Combobox(font, textvariable=self.config.main.font.字体间距, value=("标准", "紧密", "较宽"), state="readonly").grid(row=3, column=2, padx=10, pady=10)
        tooltip(ttk.Label(font, text="字符缩放"), "字符缩放,百分比").grid(row=3, column=3, padx=10, pady=10)
        ttk.Entry(font, textvariable=self.config.main.font.字符缩放).grid(row=3, column=4, padx=10, pady=10)
        # 图片
        image = ttk.Labelframe(frame, text="图片段落", bootstyle="dark")
        image.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Label(image, text="首行缩进（字符）"),"首行缩进，单位为字符，0即为关闭").grid(row=1, column=1, padx=10, pady=10)
        ttk.Entry(image, textvariable=self.config.main.image.首行缩进).grid(row=1, column=2, padx=10, pady=10)
        tooltip(ttk.Label(image, text="对齐方式"),"左对齐/右对齐/居中").grid(row=1, column=3, padx=10, pady=10)
        ttk.Combobox(image, textvariable=self.config.main.image.对齐方式, value=("居中", "左对齐", "右对齐"), state="readonly").grid(row=1, column=4, padx=10, pady=10)
        tooltip(ttk.Label(image, text="行距方式"), "行距方式，建议设为倍率，固定不会根据图片大小调整行高，会导致图片上部被隐藏").grid(row=2, column=1, padx=10, pady=10)
        ttk.Combobox(image, textvariable=self.config.main.image.行距方式, value=("固定", "倍率"),
                     state="readonly").grid(row=2, column=2, padx=10, pady=10)
        tooltip(ttk.Label(image, text="行距"), "倍率时为倍数，固定时为磅，建议倍率").grid(row=2, column=3, padx=10, pady=10)
        ttk.Entry(image, textvariable=self.config.main.image.行距).grid(row=2, column=4, padx=10, pady=10)
        return frame

    def _config_table(self):
        frame = ttk.Frame(self)
        # 表格
        table = ttk.Labelframe(frame, text="表格", bootstyle="dark")
        table.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Label(table, text="表格对齐"), "整个表格在页面中的对齐方式").grid(row=1, column=1, padx=10, pady=5)
        ttk.Combobox(table, textvariable=self.config.table.style.对齐, value=("居中", "左对齐", "右对齐"),state="readonly").grid(row=1, column=2, padx=10, pady=5)
        tooltip(ttk.Label(table, text="表格方向"), "决定左到右还是右到左").grid(row=1, column=3, padx=10, pady=5)
        ttk.Combobox(table, textvariable=self.config.table.style.表格方向, value=("从左到右", "从右到左"),state="readonly").grid(row=1, column=4, padx=10, pady=5)
        tooltip(ttk.Label(table, text="表格内垂直对齐方式"), "决定表格中文字的位置").grid(row=2, column=1, padx=10, pady=5)
        ttk.Combobox(table, textvariable=self.config.table.style.垂直对齐, value=("居中", "底部对齐", "顶部对齐"),state="readonly").grid(row=2, column=2, padx=10, pady=5)
        tooltip(ttk.Checkbutton(table, text="自动调整列宽", variable=self.config.table.style.自动调整列宽), "根据表格中的内容自动调整列宽").grid(row=2, column=3, padx=10, pady=5)

        tooltip(ttk.Label(table, text="表格行高方式"), "行高调整方式").grid(row=3, column=1, padx=10, pady=5)
        ttk.Combobox(table, textvariable=self.config.table.style.行高方式, value=("自适应", "固定", "最小"),state="readonly").grid(row=3, column=2, padx=10, pady=5)
        tooltip(ttk.Label(table, text="表格行高"), "单位应该是磅，只有固定时才生效").grid(row=3, column=3, padx=10, pady=5)
        ttk.Entry(table, textvariable=self.config.table.style.行高).grid(row=3, column=4, padx=10, pady=5)
        tooltip(ttk.Label(table, text="边框类型"), "边框类型").grid(row=4, column=1, padx=10, pady=5)
        ttk.Combobox(table, textvariable=self.config.table.style.边框, value=("single", "double"),state="readonly").grid(row=4, column=2, padx=10, pady=5)
        tooltip(ttk.Label(table, text="边框颜色"), "边框颜色").grid(row=4, column=3, padx=10, pady=5)
        ttk.Combobox(table, textvariable=self.config.table.style.边框颜色, value=("黑色", "红色"),state="readonly").grid(row=4, column=4, padx=10, pady=5)
        tooltip(ttk.Label(table, text="边框粗细"), "边框粗细，单位为0.1磅，即5=0.5磅，15=1.5磅").grid(row=5, column=1, padx=10, pady=5)
        ttk.Entry(table, textvariable=self.config.table.style.边框粗细).grid(row=5, column=2, padx=10, pady=5)
        # 段落
        para = ttk.Labelframe(frame, text="段落", bootstyle="dark")
        para.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Checkbutton(para, text="孤行控制", variable=self.config.table.paragraph.孤行控制),
                "孤行控制指段首第一行在页尾或段落最后一页在页首时，强制将此孤儿寡母行放在段落页").grid(row=1, column=1, padx=10, pady=5)
        tooltip(ttk.Checkbutton(para, text="对齐网络", variable=self.config.table.paragraph.对齐网络),
                "如果定义了文档网络，则与网络对齐。简单来说，行距大于1时，字体位于行距高度的上方，开启此项或可使字体居中于行距高度，具体要看文档网络如何定义").grid(
            row=1, column=2, padx=10, pady=5)
        tooltip(ttk.Checkbutton(para, text="右对齐网络", variable=self.config.table.paragraph.右对齐网络),
                "如果定义了文档网络，则自动调整右缩进。如果开启了列网络，文档自动右缩进以对齐网络，一般默认应该没有开启列网络只有行网络").grid(
            row=1, column=3, padx=10, pady=5)
        tooltip(ttk.Label(para, text="首行缩进（字符）"), "首行缩进，单位为字符，0即为关闭").grid(row=2, column=1,padx=10, pady=5)
        ttk.Entry(para, textvariable=self.config.table.paragraph.首行缩进).grid(row=2, column=2, padx=10, pady=5)
        tooltip(ttk.Label(para, text="对齐方式"), "左对齐/右对齐/居中").grid(row=2, column=3, padx=10, pady=5)
        ttk.Combobox(para, textvariable=self.config.table.paragraph.对齐方式, value=("居中", "左对齐", "右对齐"),
                     state="readonly").grid(row=2, column=4, padx=10, pady=5)
        tooltip(ttk.Label(para, text="行距方式"), "行距方式，倍率或固定").grid(row=3, column=1, padx=10, pady=5)
        ttk.Combobox(para, textvariable=self.config.table.paragraph.行距方式, value=("固定", "倍率"),
                     state="readonly").grid(row=3, column=2, padx=10, pady=5)
        tooltip(ttk.Label(para, text="行距"), "倍率时为倍数，固定时为磅").grid(row=3, column=3, padx=10, pady=5)
        ttk.Entry(para, textvariable=self.config.table.paragraph.行距).grid(row=3, column=4, padx=10, pady=5)
        # 字体
        font = ttk.Labelframe(frame, text="字体", bootstyle="dark")
        font.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Label(font, text="字体"), "输入后会在系统库里查找，没有找到的话我也不确定会发生什么").grid(row=1,column=1,padx=10,pady=5)
        ttk.Entry(font, textvariable=self.config.table.font.字体).grid(row=1, column=2, padx=10, pady=5)
        tooltip(ttk.Label(font, text="字号"), "单位为磅，小四=12，四号=14").grid(row=1, column=3, padx=10, pady=5)
        ttk.Entry(font, textvariable=self.config.table.font.字号).grid(row=1, column=4, padx=10, pady=5)
        tooltip(ttk.Label(font, text="颜色"), "").grid(row=2, column=1, padx=10, pady=5)
        ttk.Combobox(font, textvariable=self.config.table.font.颜色, value=("黑色", "红色"), state="readonly").grid(
            row=2, column=2, padx=10, pady=5)
        tooltip(ttk.Checkbutton(font, text="高亮", variable=self.config.table.font.高亮), "是否高亮显示").grid(row=2,column=3,padx=10,pady=5)
        tooltip(ttk.Checkbutton(font, text="对齐到网络", variable=self.config.table.font.对齐到网络),
                "是否对齐到网络").grid(row=2, column=4, padx=10, pady=5)
        tooltip(ttk.Label(font, text="字体间距"), "字体间距").grid(row=3, column=1, padx=10, pady=5)
        ttk.Combobox(font, textvariable=self.config.table.font.字体间距, value=("标准", "紧密", "较宽"),
                     state="readonly").grid(row=3, column=2, padx=10, pady=5)
        tooltip(ttk.Label(font, text="字符缩放"), "字符缩放,百分比").grid(row=3, column=3, padx=10, pady=5)
        ttk.Entry(font, textvariable=self.config.table.font.字符缩放).grid(row=3, column=4, padx=10, pady=5)
        # 图片
        image = ttk.Labelframe(frame, text="图片段落", bootstyle="dark")
        image.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Label(image, text="首行缩进（字符）"), "首行缩进，单位为字符，0即为关闭").grid(row=1, column=1,
                                                                                               padx=10, pady=5)
        ttk.Entry(image, textvariable=self.config.table.image.首行缩进).grid(row=1, column=2, padx=10, pady=5)
        tooltip(ttk.Label(image, text="对齐方式"), "左对齐/右对齐/居中").grid(row=1, column=3, padx=10, pady=5)
        ttk.Combobox(image, textvariable=self.config.table.image.对齐方式, value=("居中", "左对齐", "右对齐"),
                     state="readonly").grid(row=1, column=4, padx=10, pady=5)
        tooltip(ttk.Label(image, text="行距方式"), "行距方式，建议设为倍率，固定不会根据图片大小调整行高，会导致图片上部被隐藏").grid(row=2, column=1, padx=10, pady=5)
        ttk.Combobox(image, textvariable=self.config.table.image.行距方式, value=("固定", "倍率"),
                     state="readonly").grid(row=2, column=2, padx=10, pady=5)
        tooltip(ttk.Label(image, text="行距"), "倍率时为倍数，固定时为磅，建议倍率").grid(row=2, column=3, padx=10, pady=5)
        ttk.Entry(image, textvariable=self.config.table.image.行距).grid(row=2, column=4, padx=10, pady=5)
        return frame

    def _image_show(self):
        if not self.config.base.删除图片.get():
            self.image.pack(fill=ttk.X, padx=10, pady=5)
        else:
            self.image.pack_forget()

    def _char_show(self):
        if self.config.base.删除特殊字符.get():
            self.char.pack(fill=ttk.X, padx=10, pady=5)
        else:
            self.char.pack_forget()


    def _select_file(self):
        path = tkf.askopenfilename(filetypes =[("DOCX", ".docx")], multiple=True)
        self.config.file.set(";".join(path))

    def export_config(self):
        self.config.save()


    def start(self):
        self._disable_start()
        result = set_docx(self.config, self.process, self.tip)
        if type(result)==bool and result:
            Messagebox.ok("处理完成！", "成功")
        else:
            Messagebox.show_error(result, "错误" )
        self._enable_start()