import customtkinter
import tkinter as tk
from tkinter import *
from PIL import Image, ImageTk
from tkinter import ttk
import openpyxl
wb=openpyxl.load_workbook("tabledata.xlsx")
sh1=wb['borrow']
def historyin(main_frame):
    # ===================header1====================
    header = customtkinter.CTkLabel(main_frame, text="History", font=("Poppins", 29, 'bold'),
                                    bg_color='#673F80', fg_color='#673F80', width=199, height=55)

    header.place(x=488, y=21)
    frameline = customtkinter.CTkFrame(main_frame, border_width=1, corner_radius=1, width=1060, height=1,
                                       fg_color="white", bg_color="#673F80")
    frameline.place(x=33, y=78)

    def tablefun():
        def load_data():
            path = "tabledata.xlsx"
            workbook = openpyxl.load_workbook(path)
            sh1 = workbook['borrow']
            list_value = list(sh1.values)
            for col_name in list_value[0]:
                Tview.heading(col_name, text=col_name)
            for value_tuple in list_value[1:]:

                if value_tuple[9] == "yes":
                    Tview.insert('', tk.END, values=value_tuple)

        secondframe = customtkinter.CTkFrame(main_frame,fg_color="#412948",corner_radius=10)
        secondframe.place(x=33, y=82)
        scrool = customtkinter.CTkScrollbar(secondframe)
        scrool.pack(side="right", fill="y")
        style = ttk.Style()

        style.theme_use("default")
        style.configure("Treeview",
                        font=("Calibre", 15),
                        background="#412948",
                        foreground="white",
                        rowheight=56,

                        fieldbackground="#412948")
        style.map('Treeview',
                  background=[('selected', '#A569BD')])

        cols = ["No.", "First Name", "Last Name", "ID", "Grade", "Subject", "Type", "Time", "return on", "returned"]
        Tview = ttk.Treeview(secondframe, show="headings", yscrollcommand=scrool.set, columns=cols, height=11)
        Tview.column("No.", width=38)
        Tview.column("First Name", width=136)
        Tview.column("Last Name", width=136)
        Tview.column("ID", width=66)
        Tview.column("Grade", width=96)
        Tview.column("Subject", width=120)
        Tview.column("Type", width=87)
        Tview.column("Time", width=155)
        Tview.column("return on", width=155)
        Tview.column("returned", width=52)
        Tview.pack()
        scrool.configure(command=Tview.yview)
        load_data()
        def line(x,y):
            frameline = customtkinter.CTkFrame(secondframe, border_width=1, corner_radius=1, width=1030, height=1,
                                               fg_color="white", bg_color="#673F80")
            frameline.place(x=x, y=y)
        line(3,75)
        line(3, 75+56)
        line(3,75+56+56)
        line(3,75+56+56+56)
        line(3,75+56+56+56+56)
        line(3,75+56+56+56+56+56)
        line(3,75+56+56+56+56+56+56)
        line(3,75+56+56+56+56+56+56+56)
        line(3,75+56+56+56+56+56+56+56+56)
        line(3,75+56+56+56+56+56+56+56+56+56)
        line(3,75+56+56+56+56+56+56+56+56+56+56)

    tablefun()