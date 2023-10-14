#importing the gui library
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog
from  Converter import XlsxToXml as Converter

global_excel_file = ""
global_output_folder = ""
#declaring the necessary functions above their calling

def select_folder():
    path = filedialog.askdirectory()
    if path:
        print(path)
        global global_output_folder
        global_output_folder = path

        global Folderlabel
        Folderlabel.configure(text = global_output_folder)

def select_file():
    path = filedialog.askopenfilename(filetypes=[("Excel Workbook", "*.xlsx"), 
                                                ("Excel Marco-Enabled Workbook", "*.xlsm"),
                                                ("Excel 97-2003 Workbook", "*.xls")])

    if path:
        print(path)
        global global_excel_file
        global_excel_file = path

        global Filelabel
        Filelabel.configure(text = global_excel_file)

def convert_command():
    print("starting conversion")
    if global_excel_file != "" and global_output_folder != "":
        print(type(global_excel_file), type(global_output_folder))
        Converter.start(global_excel_file,global_output_folder)

#creating a window variable
window = ctk.CTk()
window.title("ExcelToXml")

#window.iconbitmap("my_icon.ico")
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


width = 960
height = 720

screen_width = window.winfo_screenwidth()  # Width of the screen
screen_height = window.winfo_screenheight() # Height of the screen

x = (screen_width/2) - (width/2)
y = (screen_height/2) - (height/2)

frame = ctk.CTkFrame(master=window,
                    width=960,
                    height=720,
                    corner_radius=10)

frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

Uframe = ctk.CTkFrame(master=frame,
                    width=960,
                    height=108,
                    corner_radius=10)

Uframe.place(relx=0.5, rely=0.2, anchor=tk.CENTER)

Mframe = ctk.CTkFrame(master=frame,
                    width=960,
                    height=108,
                    corner_radius=10)

Mframe.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

Folderlabel = ctk.CTkLabel(master=Mframe,
                        text="Select an output folder",
                        width=720,
                        height=25,
                        corner_radius=8)
Folderlabel.place(relx=0.4, rely=0.5, anchor=tk.CENTER)
Folderlabel.configure(font=("Arial", 16), fg_color="black")

Filelabel = ctk.CTkLabel(master=Uframe,
                            text="Select a file",
                            width=720,
                            height=25,
                            corner_radius=8)
Filelabel.place(relx=0.4, rely=0.5, anchor=tk.CENTER)
Filelabel.configure(font=("Arial", 16), fg_color="black")

FolderButton = ctk.CTkButton(
                        master=Mframe,
                        command=select_folder,
                        corner_radius=10,
                        text="SelectFolder")
FolderButton.place(relx=0.9,rely=0.5,anchor = tk.CENTER)


FileButton = ctk.CTkButton(
                        master=Uframe,
                        command=select_file,
                        corner_radius=10,
                        text="SelectFile")
FileButton.place(relx=0.9,rely=0.5,anchor = tk.CENTER)

convert_button = ctk.CTkButton(
                            master=frame,
                            corner_radius=10,
                            command=convert_command,
                            text="Convert",
                            width=200,
                            height=50,)
convert_button.place(relx=0.5,rely=0.9,anchor = tk.S)
convert_button.configure(font=("Arial", 20,"bold"))


window.geometry('%dx%d+%d+%d' % (width, height, x, y))

#intializing the window
window.mainloop()
