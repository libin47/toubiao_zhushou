
# import sv_ttk
import ttkbootstrap as ttk
from app.hidden_clean import HiddenCleaner
from app.pdf2img import Pdf2img
import tkinter.filedialog as tkf
import json
from ttkbootstrap.dialogs.dialogs import Messagebox


class Application(ttk.Frame):
    def __init__(self, root=None):
        super().__init__(root)
        self.root = root
        self.pack()
        self.root.title("投标助手")
        # self.root.geometry("400x300")
        # self.root.resizable(False, False)
        # self.root.iconbitmap("ico.ico")

        self.data_clean = {"file": ""}

        self._create()
        self.root.after(100, self.refresh_data)

    def refresh_data(self):
        self.root.update()
        self.root.after(100, self.refresh_data)



    def _create(self):
        self._create_menu()
        self._create_notebooks()

    def _create_menu(self):
        """
        创建菜单栏
        """
        menu_bar = ttk.Menu(self.root)
        self.root.config(menu=menu_bar)
        # Create a File menu
        config_menu = ttk.Menu(menu_bar, tearoff=0)
        about_menu = ttk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="配置", menu=config_menu)
        menu_bar.add_cascade(label="关于", menu=about_menu)
        # 菜单具体内容
        config_menu.add_command(label="配置文件导入", command=self._import_config)
        config_menu.add_command(label="配置文件导出", command=self._export_config)
        config_menu.add_separator()
        config_menu.add_command(label="退出", command=self.root.quit)
        about_menu.add_command(label="关于", command=self._about)

    def _create_notebooks(self):
        """
        创建多标签页
        """
        self.notebook = ttk.Notebook(self, bootstyle="info")
        self.notebook.pack(fill=ttk.BOTH, expand=True)
        self.HiddenCleaner = HiddenCleaner(self.notebook)
        self.notebook.add(self.HiddenCleaner, text="暗标格式刷")
        self.notebook.add(self._create_clone_tab(), text="标书查重")
        self.Pdf2Img = Pdf2img(self.notebook)
        self.notebook.add(self.Pdf2Img, text="pdf转图片")



    def _create_clone_tab(self):
        config_tab = ttk.Frame(self.notebook)
        return config_tab

    def _create_pad2img_tab(self):
        pad2img_tab = ttk.Frame(self.notebook)
        return pad2img_tab

    def _export_config(self):
        file_path = tkf.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            hd = self.HiddenCleaner.config.export()
            result = {
                "hidden": hd
            }
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(json.dumps(result, ensure_ascii=False, indent=4))
            Messagebox.ok("配置文件导出成功！", "成功")
        else:
            Messagebox.show_error( "未选择文件路径！", "警告")

    def _import_config(self):
        file_path = tkf.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            self.HiddenCleaner.config.import_config(config["hidden"])
            Messagebox.ok("配置文件导入成功！", "成功")
        else:
            Messagebox.show_error( "未选择文件路径！", "警告")

    def _about(self):
        Messagebox.show_info("版本：V0.1.20250106", "关于本软件") # TODO: 图表


if __name__ == '__main__':
    ttk.utility.enable_high_dpi_awareness()
    # root = tk.Tk()
    root = ttk.Window(themename="flatly", iconphoto="ico.ico")
    app = Application(root)
    # sv_ttk.use_light_theme()
    root.mainloop()