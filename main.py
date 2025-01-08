from app import Application
import ttkbootstrap as ttk


if __name__ == '__main__':
    ttk.utility.enable_high_dpi_awareness()
    # root = tk.Tk()
    root = ttk.Window(themename="flatly", iconphoto="ico.ico")
    app = Application(root)
    # sv_ttk.use_light_theme()
    root.mainloop()