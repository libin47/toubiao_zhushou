# -*- coding: utf-8 -*-
import ttkbootstrap as ttk

class Pdf2imgConfig(object):
    def __init__(self):
        self.file = ttk.StringVar()
        self.file.set('')
        self.quality = ttk.IntVar()
        self.quality.set(2)