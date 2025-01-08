import os.path
# import tkinter as tk
import tkinter.filedialog as tkf
from ttkbootstrap.tooltip import ToolTip
import time
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs.dialogs import Messagebox
import ttkbootstrap as ttk
import threading as mt
from app.pdf2img_config import Pdf2imgConfig
from app.pdf2img_fun import pdf2image

def tooltip(widget, text):
    """
    给组件添加提示
    """
    ToolTip(widget, text=text)
    return widget


class Pdf2img(ttk.Frame):
    def __init__(self, notebook:ttk.Notebook):
        super().__init__(notebook)
        self.notebook = notebook
        self.config = Pdf2imgConfig()
        self.tip = ttk.StringVar()
        self.tip.set("请选择文件")
        self._create()


    def _create(self):
        # 选择文件
        file = ttk.Labelframe(self, text="需处理文件", bootstyle="info")
        file.pack(fill=ttk.X, padx=10, pady=5)
        ttk.Label(file, text="选择文件").pack(side=LEFT, padx=5, pady=10)
        ttk.Entry(file, textvariable=self.config.file, width=50).pack(side=LEFT, padx=5, pady=10)
        ttk.Button(file, bootstyle=INFO, text="选择", command=lambda: self._select_file()).pack(side=LEFT, padx=5, pady=10)
        self.start_button = ttk.Button(file, bootstyle=SUCCESS, text="启动", command=lambda: self.start())
        self.start_button.pack(side=LEFT, padx=5, pady=10)
        self.label = ttk.Label(self, textvariable=self.tip, bootstyle="info")
        self.label.pack(fill=ttk.X, padx=10, pady=5)
        self.progress = ttk.Progressbar(self, maximum=100, value=0, bootstyle="info")
        self.progress.pack(fill=ttk.X, padx=10, pady=5)
        # 添加配置卡
        cfg = ttk.Labelframe(self, text="设置", bootstyle="dark")
        cfg.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Label(cfg, text="输出图片质量"), "图片的缩放系数，越大分辨率越大").pack(side=LEFT, padx=5, pady=10)
        ttk.Entry(cfg, textvariable=self.config.quality, width=10).pack(side=LEFT, padx=5, pady=10)


    def _select_file(self):
        path = tkf.askopenfilename(filetypes =[("PDF", ".pdf")], multiple=True)
        self.config.file.set(";".join(path))

    def _disable_start(self):
        self.start_button.config(state=DISABLED, text="处理中")
        self.start_button.update_idletasks()

    def _enable_start(self):
        self.start_button.config(state=NORMAL, text="启动")
        self.start_button.update_idletasks()


    def start(self):
        self._disable_start()
        result = pdf2image(self.config, self.progress, self.tip)
        if type(result)==bool and result:
            Messagebox.ok("处理完成！", "成功")
        else:
            Messagebox.show_error(result, "错误" )
        self._enable_start()