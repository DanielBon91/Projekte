import tkinter.ttk as ttk
from tkinter.ttk import Style
class CustomTreeView(ttk.Treeview):

    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)

        Style().configure("Treeview", rowheight=30, font=("Calibri", 11))
