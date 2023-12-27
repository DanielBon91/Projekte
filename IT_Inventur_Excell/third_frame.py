import customtkinter as ctk
import custom_treeview as ctv
from table1 import Table1
from table2 import Table2
from table3 import Table3

class ThirdFrame(ctk.CTkFrame):

    def __init__(self, master):
        super().__init__(master, fg_color="transparent")

        self.grid_columnconfigure((0, 1, 2), weight=1)

        self.third_frame_lager_button = ctk.CTkButton(self, text="Lager", fg_color="gray25",
                                                                height=50, corner_radius=0, hover_color="gray45",
                                                                font=ctk.CTkFont(size=21, weight="bold"), command=self.table_1)
        self.third_frame_inventar_button = ctk.CTkButton(self, text="Inventarisierung", height=50, corner_radius=0,
                                                                font=ctk.CTkFont(size=21, weight="bold"),
                                                                command=self.table_2, fg_color="gray25", hover_color="gray45")
        self.third_frame_mitarbeter_button = ctk.CTkButton(self, text="Mitarbeiter", height=50, corner_radius=0,
                                                                font=ctk.CTkFont(size=21, weight="bold"),
                                                                command=self.table_3, fg_color="gray25", hover_color="gray45")
        self.third_frame_lager_button.grid(row=0, column=0, pady=(10, 0), padx=(8, 1), sticky="we")
        self.third_frame_inventar_button.grid(row=0, column=1, pady=(10, 0), padx=1, sticky="we")
        self.third_frame_mitarbeter_button.grid(row=0, column=2, pady=(10, 0), padx=(1, 8), sticky="we")

        self.table_1_menu = Table1(self)
        self.table_2_menu = Table2(self)
        self.table_3_menu = Table3(self)

    def select_table(self, table):
        self.third_frame_lager_button.configure(text_color=("black", "white"),
                                                fg_color=("gray75", "gray25") if table == "Table_1" else "transparent")
        self.third_frame_inventar_button.configure(text_color=("black", "white"),
                                                fg_color=("gray75", "gray25") if table == "Table_2" else "transparent")
        self.third_frame_mitarbeter_button.configure(text_color=("black", "white"),
                                                fg_color=("gray75", "gray25") if table == "Table_3" else "transparent")
        if table=="Table_1":
            self.table_1_menu.first_table_function()
        else:
            self.table_1_menu.grid_forget()
        if table=="Table_2":
            self.table_2_menu.second_table_function()
        else:
            self.table_2_menu.grid_forget()
        if table=="Table_3":
            self.table_3_menu.third_table_funktion()
        else:
            self.table_3_menu.grid_forget()


    def table_1(self):
        self.select_table("Table_1")
    def table_2(self):
        self.select_table("Table_2")
    def table_3(self):
        self.select_table("Table_3")
