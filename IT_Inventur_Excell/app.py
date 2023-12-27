import configparser

import customtkinter as ctk
from first_frame import FirstFrame
from second_frame import SecondFrame
from third_frame import ThirdFrame
from fourth_frame import FourthFrame
from PIL import Image
config = configparser.ConfigParser()
config.read("config.ini", encoding='utf-8')


class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.screen_width = self.winfo_screenwidth()
        self.screen_height = self.winfo_screenheight()

        x = (self.screen_width / 2) - (1410 / 2)
        y = (self.screen_height / 2) - (1200 / 2)
        self.geometry(f"{1500}x{1200}+{int(x)}+{int(y)}")

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)
        self.title("Argen IT Invertory")

        img_dir = config['directory']['images_dir']

        self.image_directory = ctk.CTkImage(dark_image=Image.open(fr"{img_dir}//ArgenLogo_Weiss.png"),
                                                      light_image=Image.open(fr"{img_dir}//ArgenLogo_Schwarz.png"),
                                                      size=(300, 50))

        self.image_warenannahme = ctk.CTkImage(dark_image=Image.open(fr"{img_dir}//box_dark.png"),
                                                      light_image=Image.open(fr"{img_dir}//box_light.png"),
                                                      size=(30, 30))

        self.image_warenuebergabe = ctk.CTkImage(dark_image=Image.open(fr"{img_dir}//ubergabe_dark.png"),
                                                      light_image=Image.open(fr"{img_dir}//ubergabe_light.png"),
                                                      size=(35, 35))

        self.image_liste = ctk.CTkImage(dark_image=Image.open(fr"{img_dir}//list_dark.png"),
                                                      light_image=Image.open(fr"{img_dir}//list_light.png"), size=(30, 30))

        self.image_mitarbeiter = ctk.CTkImage(dark_image=Image.open(fr"{img_dir}//arbeiter_dark.png"),
                                                      light_image=Image.open(fr"{img_dir}//arbeiter_light.png"),
                                                      size=(30, 30))

        self.navi_frame = ctk.CTkFrame(self, corner_radius=5)
        self.navi_frame.grid(row=0, column=0, sticky="ns")
        self.navi_frame.grid_columnconfigure(0, minsize=350)

        self.navi_frame.grid_rowconfigure(5, weight=1)
        self.navi_frame.grid_rowconfigure(6, weight=0)

        self.image_label = ctk.CTkLabel(self.navi_frame, text="", image=self.image_directory)
        self.image_label.grid(row=0, column=0, padx=15, pady=55)
        self.navi_button_1 = ctk.CTkButton(self.navi_frame, text="Warenannahme",
                                                     text_color=("gray10", "gray90"),
                                                     height=120, width=250, fg_color="transparent",
                                                     image=self.image_warenannahme, hover_color=("gray70", "gray30"),
                                                     corner_radius=0, font=ctk.CTkFont(size=24),
                                                     command=self.first_frame_navi_button)
        self.navi_button_1.grid(row=1, column=0, sticky="we")
        self.navi_button_2 = ctk.CTkButton(self.navi_frame, text="Warenübergabe",
                                                     text_color=("gray10", "gray90"),
                                                     height=120, image=self.image_warenuebergabe,
                                                     fg_color="transparent", hover_color=("gray70", "gray30"),
                                                     corner_radius=0, border_spacing=10, font=ctk.CTkFont(size=24),
                                                     command=self.second_frame_navi_button)
        self.navi_button_2.grid(row=2, column=0, sticky="we")
        self.navi_button_3 = ctk.CTkButton(self.navi_frame, text="Liste", height=120, fg_color="transparent",
                                                     text_color=("gray10", "gray90"), image=self.image_liste,
                                                     hover_color=("gray70", "gray30"),
                                                     corner_radius=0, font=ctk.CTkFont(size=24),
                                                     command=self.three_frame_navi_button)
        self.navi_button_3.grid(row=3, column=0, sticky="we")
        self.navi_button_4 = ctk.CTkButton(self.navi_frame, text="Mitarbeiter",
                                                     text_color=("gray10", "gray90"), image=self.image_mitarbeiter,
                                                     corner_radius=0, height=120, fg_color="transparent",
                                                     hover_color=("gray70", "gray30"), font=ctk.CTkFont(size=24),
                                                     command=self.four_frame_navi_button)
        self.navi_button_4.grid(row=4, column=0, sticky="we")

        self.change_mode = ctk.CTkOptionMenu(self.navi_frame, corner_radius=0, values=["Dark", "Light"],
                                             command=self.change_mode).grid(row=5, column=0, pady=20, sticky="s")

        self.first_menu = FirstFrame(self)
        self.second_menu = SecondFrame(self)
        self.third_menu = ThirdFrame(self)
        self.fourth_menu = FourthFrame(self)

        self.select_frame_by_name("Button_1")
    def first_frame_navi_button(self):
        self.select_frame_by_name("Button_1")
    def second_frame_navi_button(self):
        self.select_frame_by_name("Button_2")
    def three_frame_navi_button(self):
        self.select_frame_by_name("Button_3")
    def four_frame_navi_button(self):
        self.select_frame_by_name("Button_4")
    def select_frame_by_name(self, name):
        self.navi_button_1.configure(fg_color=("gray75", "gray25") if name == "Button_1" else "transparent")
        self.navi_button_2.configure(fg_color=("gray75", "gray25") if name == "Button_2" else "transparent")
        self.navi_button_3.configure(fg_color=("gray75", "gray25") if name == "Button_3" else "transparent")
        self.navi_button_4.configure(fg_color=("gray75", "gray25") if name == "Button_4" else "transparent")

        if name == "Button_1":
            self.first_menu.grid(row=0, column=1, sticky="nsew")
        else:
            self.first_menu.grid_forget()
        if name == "Button_2":
            self.second_menu.grid(row=0, column=1, sticky="nsew")
            if len(self.second_menu.empty_table.get_children()) == 0:
                self.second_menu.second_frame_lager_tabelle()
        else:
            self.second_menu.grid_forget()
        if name == "Button_3":
            self.third_menu.grid(row=0, column=1, sticky="nsew")
            self.third_menu.select_table("Table_1")
        else:
            self.third_menu.grid_forget()
        if name == "Button_4":
            self.fourth_menu.grid(row=0, column=1, sticky="nsew")
        else:
            self.fourth_menu.grid_forget()
    def change_mode(self, mode):
        ctk.set_appearance_mode(mode)




