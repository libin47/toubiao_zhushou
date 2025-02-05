# -*- coding: utf-8 -*-
import json

import ttkbootstrap as ttk
from app.utils import convert_to_json, json_to_obj

class ConfigBase(object):
    """
    基础配置类
    """
    def __init__(self):
        self.删除空行 = ttk.BooleanVar()
        self.删除空格 = ttk.IntVar()  # 0不删除 1多个连续空格只保留一个 2全部删除
        self.删除制表符 = ttk.BooleanVar()
        self.删除图片 = ttk.BooleanVar()
        self.图片灰度化 = ttk.BooleanVar()  # 仅不删除图片时有效
        self.图片缩放 = ttk.BooleanVar()  # 仅不删除图片时有效
        self.图片最大宽度 = ttk.IntVar()  # 仅在开启图片缩放时生效
        self.图片嵌入 = ttk.BooleanVar()
        self.标点转中文 = ttk.BooleanVar()  # ＃将英文标点转为中文标点，先于转全角生效
        self.半角转为全角 = ttk.BooleanVar()
        self.删除特殊字符 = ttk.BooleanVar()
        self.删除表格特殊字符 = ttk.BooleanVar()
        self.特殊字符_删除 = ttk.StringVar()
        self.特殊字符_保留 = ttk.StringVar()
        self.中英文标点字典 = [",:;<>?\\|[]!$", "，：；《》？、|【】！￥"]
        self.default()# 英文、数字、英文标点转为全角字符，注：中文字符均为全角没有半角

    def default(self):
        self.删除空行.set(True)
        self.删除空格.set(2)
        self.删除制表符.set(True)
        self.删除图片.set(False)
        self.图片灰度化.set(True)
        self.图片缩放.set(True)
        self.图片最大宽度.set(720)
        self.图片嵌入.set(True)
        self.标点转中文.set(True)
        self.半角转为全角.set(False)
        self.删除特殊字符.set(True)
        self.删除表格特殊字符.set(False)
        self.特殊字符_删除.set("")
        self.特殊字符_保留.set("")

class ConfigExtend(object):
    """
    扩展配置类
    """
    def __init__(self):
        self.清除w14样式 = ttk.BooleanVar()
        self.删除自动编号 = ttk.BooleanVar()  # 意味着将会删除所有段落属性
        self.封面目录处理 = ttk.BooleanVar()
        self.同步office样式 = ttk.BooleanVar()
        self.设置标题编号 = ttk.BooleanVar()
        self.删除原有标题编号 = ttk.BooleanVar()
        self.原有标题编号样式 = ttk.StringVar()
        self.标题编号ID = ttk.IntVar()
        self.一级编号 = ttk.StringVar()
        self.二级编号 = ttk.StringVar()
        self.三级编号 = ttk.StringVar()
        self.四级编号 = ttk.StringVar()
        self.五级编号 = ttk.StringVar()
        self.六级编号 = ttk.StringVar()
        self.七级编号 = ttk.StringVar()
        self.八级编号 = ttk.StringVar()
        self.九级编号 = ttk.StringVar()
        self.一级编号Fmt = ttk.StringVar()
        self.二级编号Fmt = ttk.StringVar()
        self.三级编号Fmt = ttk.StringVar()
        self.四级编号Fmt = ttk.StringVar()
        self.五级编号Fmt = ttk.StringVar()
        self.六级编号Fmt = ttk.StringVar()
        self.七级编号Fmt = ttk.StringVar()
        self.八级编号Fmt = ttk.StringVar()
        self.九级编号Fmt = ttk.StringVar()
        self.一级编号Lgl = ttk.BooleanVar()
        self.二级编号Lgl = ttk.BooleanVar()
        self.三级编号Lgl = ttk.BooleanVar()
        self.四级编号Lgl = ttk.BooleanVar()
        self.五级编号Lgl = ttk.BooleanVar()
        self.六级编号Lgl = ttk.BooleanVar()
        self.七级编号Lgl = ttk.BooleanVar()
        self.八级编号Lgl = ttk.BooleanVar()
        self.九级编号Lgl = ttk.BooleanVar()

        self.清除w14样式.set(True)  # microsoft word2010及更高版本使用w14命名空间提供更多拓展样式，可能会覆盖旧版样式。启用此选项以清除所有新版样式
        self.删除自动编号.set(True)
        self.封面目录处理.set(True)
        self.同步office样式.set(True)
        self.设置标题编号.set(True)
        self.删除原有标题编号.set(True)
        self.原有标题编号样式.set(r"^\d+(\.\d+)*;^第\d+章;^第\d+节")
        self.标题编号ID.set(1)
        self.一级编号.set("第%1章 ")
        self.二级编号.set("%1.%2 ")
        self.三级编号.set("%1.%2.%3 ")
        self.四级编号.set("%1.%2.%3.%4 ")
        self.五级编号.set("%1.%2.%3.%4.%5 ")
        self.六级编号.set("%1.%2.%3.%4.%5.%6 ")
        self.七级编号.set("%1.%2.%3.%4.%5.%6.%7 ")
        self.八级编号.set("%1.%2.%3.%4.%5.%6.%7.%8 ")
        self.九级编号.set("%1.%2.%3.%4.%5.%6.%7.%8.%9 ")
        self.一级编号Fmt.set("一 二 三")
        self.二级编号Fmt.set("1 2 3")
        self.三级编号Fmt.set("1 2 3")
        self.四级编号Fmt.set("1 2 3")
        self.五级编号Fmt.set("1 2 3")
        self.六级编号Fmt.set("1 2 3")
        self.七级编号Fmt.set("1 2 3")
        self.八级编号Fmt.set("1 2 3")
        self.九级编号Fmt.set("1 2 3")
        self.一级编号Lgl.set(False)
        self.二级编号Lgl.set(True)
        self.三级编号Lgl.set(True)
        self.四级编号Lgl.set(True)
        self.五级编号Lgl.set(True)
        self.六级编号Lgl.set(True)
        self.七级编号Lgl.set(True)
        self.八级编号Lgl.set(True)
        self.九级编号Lgl.set(True)

class ConfigPageDistance(object):
    """
    页边距配置类
    """
    def __init__(self):
        self.上 = ttk.DoubleVar()
        self.下 = ttk.DoubleVar()
        self.左 = ttk.DoubleVar()
        self.右 = ttk.DoubleVar()
        self.上.set(2.5)
        self.下.set(2.0)
        self.左.set(2.0)
        self.右.set(2.0)

class ConfigPage(object):
    """
    页面配置类
    """
    def __init__(self):
        self.页边距 = ConfigPageDistance()

class ConfigCore(object):
    """
    核心配置类
    """
    def __init__(self):
        self.删除文档属性 = ttk.BooleanVar()
        self.删除页眉页脚 = ttk.BooleanVar()

        self.删除文档属性.set(True)
        self.删除页眉页脚.set(True)

class ConfigFont(object):
    def __init__(self):
        self.字体 = ttk.StringVar()
        self.字号 = ttk.IntVar()
        self.颜色 = ttk.StringVar()
        self.高亮 = ttk.BooleanVar()
        self.字体间距 = ttk.StringVar()
        self.字符缩放 = ttk.IntVar()
        self.对齐到网络 = ttk.BooleanVar()
        self.字体.set("宋体")
        self.字号.set(14) # 小四=12 四号=14
        self.颜色.set("黑色")
        self.高亮.set(False)
        self.字体间距.set("标准") # 标准2 紧密0 较宽4
        self.字符缩放.set(100) # 单位为%
        self.对齐到网络.set(True)

class ConfigParagraph(object):
    def __init__(self):
        self.首行缩进 = ttk.IntVar()  # 单位是字符
        self.行距方式 = ttk.StringVar()  # 倍率 固定
        self.行距 = ttk.DoubleVar()  # 30
        self.孤行控制 = ttk.BooleanVar()  # 0
        self.对齐方式 = ttk.StringVar()  # 居中、左对齐、右对齐
        self.对齐网络 = ttk.BooleanVar()  # 0
        self.右对齐网络 = ttk.BooleanVar()  # 0

        self.首行缩进.set(2)
        self.行距方式.set("倍率")
        self.行距.set(1.5)
        self.孤行控制.set(False)
        self.对齐方式.set("左对齐")
        self.对齐网络.set(False)
        self.右对齐网络.set(False)

class ConfigImage(object):
    def __init__(self):
        self.首行缩进 = ttk.BooleanVar()
        self.对齐方式 = ttk.StringVar()
        self.行距方式 = ttk.StringVar()
        self.行距 = ttk.DoubleVar()
        self.首行缩进.set(False)
        self.对齐方式.set("居中")
        self.行距方式.set("倍率")
        self.行距.set(1.5)


class ConfigTableBase(object):
    def __init__(self):
        self.对齐 = ttk.StringVar()  # 居中、左对齐、右对齐 # 指整个表格相对页面
        self.表格方向 = ttk.StringVar()  # 从左到右
        self.垂直对齐 = ttk.StringVar()  # 居中、底部对齐、顶部对齐 # 指表格内文字相对单元格，水平方向由段落里的对齐方式控制
        self.行高方式 = ttk.StringVar()  # 自适应 固定 最小
        self.行高 = ttk.DoubleVar()  # 10
        self.自动调整列宽 = ttk.BooleanVar()  # 1
        self.边框 = ttk.StringVar()  # single
        self.边框颜色 = ttk.StringVar()  # [255, 0, 0]  # RGB颜色，黑色为[0,0,0]
        self.边框粗细 = ttk.IntVar()  # 5  # 5=0.5磅 15=1.5磅

        self.对齐.set("居中")
        self.表格方向.set("从左到右")
        self.垂直对齐.set("居中")
        self.行高方式.set("自适应")
        self.行高.set(10)
        self.自动调整列宽.set(True)
        self.边框.set("single")
        self.边框颜色.set("黑色")
        self.边框粗细.set(5)


class ConfigTableParagraph(ConfigParagraph):
    def __init__(self):
        super().__init__()
        self.首行缩进.set(0)
        self.行距方式.set("倍率")
        self.行距.set(1.5)
        self.孤行控制.set(False)
        self.对齐方式.set("居中")
        self.对齐网络.set(False)
        self.右对齐网络.set(False)


class ConfigTableFont(ConfigFont):
    def __init__(self):
        super().__init__()
        self.字体.set("宋体")
        self.字号.set(14) # 小四=12 四号=14
        self.颜色.set("黑色")
        self.高亮.set(False)
        self.字体间距.set("标准") # 标准2 紧密0 较宽4
        self.字符缩放.set(100) # 单位为%
        self.对齐到网络.set(True)

class ConfigTableImage(ConfigImage):
    def __init__(self):
        super().__init__()
        self.首行缩进.set(False)
        self.对齐方式.set("居中")
        self.行距方式.set("倍率")
        self.行距.set(1.5)

class ConfigTable(object):
    """
    表格配置
    """
    def __init__(self):
        self.style = ConfigTableBase()
        self.font = ConfigTableFont()
        self.paragraph = ConfigTableParagraph()
        self.image = ConfigTableImage()



class ConfigMain(object):
    """
    正文配置
    """
    def __init__(self):
        self.font = ConfigFont()
        self.paragraph = ConfigParagraph()
        self.image = ConfigImage()


class HiddenCleanerConfig(object):
    def __init__(self):
        self.file = ttk.StringVar()
        self.base = ConfigBase()
        self.page = ConfigPage()
        self.extend = ConfigExtend()
        self.core = ConfigCore()
        self.table = ConfigTable()
        self.main = ConfigMain()

    def export(self):
        r = convert_to_json(self.__dict__)
        return r

    def import_config(self, data):
        json_to_obj(self, data)

