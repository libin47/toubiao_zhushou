# -*- coding: utf-8 -*-
import ttkbootstrap as ttk

class CompareConfig(object):
    def __init__(self):
        self.zbfiles = ttk.StringVar()
        self.tbfiles = ttk.StringVar()
        self.tbcolor = [[155,0,0], [0,155,0],[0,0,155],[100,100,0],[100,0,100],[0,100,100],[155,100,0],[0,100,155],[0,155,100]]

        self.colors = []

        self.splitnum = ttk.IntVar()
        self.splitnum.set(10)

        self.splitword = ttk.StringVar()
        self.splitword.set('。|:|：|,|，')

        self.repeatnum = ttk.DoubleVar()
        self.repeatnum.set(0.8)
