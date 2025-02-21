import tkinter.filedialog as tkf
import tkinter as tk
from ttkbootstrap.dialogs.dialogs import Messagebox
from ttkbootstrap.constants import DISABLED, NORMAL, RIGHT, LEFT, INFO, SUCCESS, WARNING, DANGER, TOP, LIGHT
import ttkbootstrap as ttk
from app.compare_fun import compare
from app.compare_config import CompareConfig
from app.utils import tooltip


class Compare(ttk.Frame):
    def __init__(self, notebook:ttk.Notebook):
        super().__init__(notebook)
        self.config = CompareConfig()
        self.tip = ttk.StringVar()
        self.tip.set("请选择文件")

        self._create()







    def _create(self):
        # 选择招标文件
        self.zb_frame = self._create_zb()

        # 选择投标文件
        self.tb_frame = self._create_tb()

        # 参数设置
        self.config_frame = self._create_config()

        # 启动按钮
        self.start_button = ttk.Button(self, bootstyle=SUCCESS, text="启动", command=lambda: self.start())
        self.start_button.pack(fill=ttk.X, padx=10, pady=5)

        self.process = ttk.Progressbar(self, maximum=100, value=0, bootstyle="info")
        self.process.pack(fill=ttk.X, padx=10, pady=5)

        self.label = ttk.Label(self, textvariable=self.tip, bootstyle="info")
        self.label.pack(fill=ttk.X, padx=10, pady=5)

        # tbfile = ttk.Labelframe(self, text="投标文件", bootstyle="info")
        # tbfile.pack(fill=ttk.X, padx=10, pady=5)
        # ttk.Label(tbfile, text="选择文件").pack(side=LEFT, padx=5, pady=10)
        # ttk.Entry(tbfile, textvariable=self.config.tbfile, width=50).pack(side=LEFT, padx=5, pady=10)
        # ttk.Button(tbfile, bootstyle=INFO, text="选择", command=lambda: self._select_file()).pack(side=LEFT, padx=5, pady=10)
        #
        # self.start_button = ttk.Button(self, bootstyle=SUCCESS, text="启动", command=lambda: self.start())
        # self.label = ttk.Label(self, textvariable=self.tip, bootstyle="info")
        # self.label.pack(fill=ttk.X, padx=10, pady=5)
        # self.start_button.pack(side=LEFT, padx=5, pady=10)
        # self.process = ttk.Progressbar(self, maximum=100, value=0, bootstyle="info")
        # self.process.pack(fill=ttk.X, padx=10, pady=5)


    def _create_zb(self):
        zbfile = ttk.Labelframe(self, text="招标文件", bootstyle="info")
        zbfile.pack(fill=ttk.X, padx=10, pady=5)
        ttk.Button(zbfile, bootstyle=INFO, text="添加文件", command=lambda: self._add_zb_file()).grid(row=1, column=1, columnspan=2, padx=10, pady=5)
        return zbfile

    def _update_zb(self):
        for sth in list(self.zb_frame.children.values())[1:]:
            sth.destroy()
        line = 2
        for file in self.config.zbfiles.get().split(";"):
            if file != "":
                ttk.Label(self.zb_frame, text=file).grid(row=line, column=1, padx=10, pady=5)
                ttk.Button(self.zb_frame, bootstyle=DANGER, text="删除", command=lambda i=line: self._del_zb_file(i-2)).grid(row=line, column=2, padx=10, pady=5)
                line += 1

    def _add_zb_file(self):
        path = tkf.askopenfilename(filetypes =[("DOCX", ".docx")], multiple=True)
        old = self.config.zbfiles.get()
        # 有旧的时候增加并去重，没有旧的时候直接添加
        # 以纯文本形式保存，以英文分号分割多个
        if old == "":
            self.config.zbfiles.set(";".join(path))
        else:
            files = old.split(";")
            files.extend(path)
            files = list(set(files))
            self.config.zbfiles.set(";".join(files))
        self._update_zb()

    def _del_zb_file(self, index):
        files = self.config.zbfiles.get().split(";")
        del files[index]
        self.config.zbfiles.set(";".join(files))
        self._update_zb()


    def _create_tb(self):
        tbfile = ttk.Labelframe(self, text="投标文件", bootstyle="info")
        tbfile.pack(fill=ttk.X, padx=10, pady=5)
        ttk.Button(tbfile, bootstyle=INFO, text="添加文件", command=lambda: self._add_tb_file()).grid(row=1, column=1, columnspan=2, padx=10, pady=5)
        return tbfile

    def _update_tb(self):
        for sth in list(self.tb_frame.children.values())[1:]:
            sth.destroy()
        line = 2
        for file in self.config.tbfiles.get().split(";"):
            if file != "":
                # TODO: 添加修改按钮、添加颜色选择
                ttk.Label(self.tb_frame, text=file).grid(row=line, column=1, padx=10, pady=5)
                ttk.Button(self.tb_frame, bootstyle=DANGER, text="删除", command=lambda i=line: self._del_tb_file(i-2)).grid(row=line, column=2, padx=10, pady=5)
                color_bt = tk.Button(self.tb_frame, text="颜色")
                color_bt["bg"] = "#%02x%02x%02x" % (int(self.config.tbcolor[line-2][0]), int(self.config.tbcolor[line-2][1]), int(self.config.tbcolor[line-2][2]))
                color_bt["fg"] = "black"
                color_bt.grid(row=line, column=3, padx=10, pady=5)
                line += 1

    def _add_tb_file(self):
        path = tkf.askopenfilename(filetypes =[("DOCX", ".docx")], multiple=True)
        old = self.config.tbfiles.get()
        if old == "":
            if len(path)>9:
                Messagebox.show_info("最多支持9个投标文件")
                path = path[:9]
            self.config.tbfiles.set(";".join(path))
        else:
            files = old.split(";")
            files.extend(path)
            files = list(set(files))
            if len(files)>9:
                Messagebox.show_info("最多支持9个投标文件")
                files = files[:9]
            self.config.tbfiles.set(";".join(files))
        self._update_tb()

    def _del_tb_file(self, index):
        files = self.config.tbfiles.get().split(";")
        del files[index]
        self.config.tbfiles.set(";".join(files))
        self._update_tb()

    def _create_config(self):
        tbfile = ttk.Labelframe(self, text="参数配置", bootstyle="info")
        tbfile.pack(fill=ttk.X, padx=10, pady=5)
        tooltip(ttk.Label(tbfile, text="监测字句最小长度"), "只有长度大于此的句子才会被检查").grid(row=1, column=1, padx=10, pady=3)
        ttk.Entry(tbfile, textvariable=self.config.splitnum).grid(row=1, column=2, padx=10, pady=3)
        tooltip(ttk.Label(tbfile, text="分割符"), "以此分开的连续字符为一个句子").grid(row=1, column=3, padx=10, pady=3)
        ttk.Entry(tbfile, textvariable=self.config.splitword).grid(row=1, column=4, padx=10, pady=3)
        tooltip(ttk.Label(tbfile, text="标记阈值"), "小于1则为百分比，大于1则按具体重复字句数量").grid(row=2, column=1, padx=10, pady=3)
        ttk.Entry(tbfile, textvariable=self.config.repeatnum).grid(row=2, column=2, padx=10, pady=3)
        return tbfile

    def _disable_start(self):
        self.start_button.config(state=DISABLED, text="处理中")
        self.start_button.update_idletasks()

    def _enable_start(self):
        self.start_button.config(state=NORMAL, text="启动")
        self.start_button.update_idletasks()

    def start(self):
        self._disable_start()
        result = compare(self.config, self.process, self.tip)
        if type(result)==bool and result:
            Messagebox.ok("处理完成！", "成功")
        else:
            Messagebox.show_error(result, "错误" )
        self._enable_start()
