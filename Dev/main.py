#importing the gui library
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog
from  Converter import XlsxToXml as Converter
import re

global_excel_file = ""
global_output_folder = ""
#declaring the necessary functions above their calling

def is_allowed_filename_in_windows(filename):
    # Check if the filename is empty.
    if not filename:
        return False
    
    # Check if the filename is one of the reserved names.
    reserved_names = ['CON', 'PRN', 'AUX', 'NUL', 'COM0', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT0', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9']
    if filename in reserved_names:
        return False
    
    # Check if the filename contains any invalid characters.
    invalid_chars = re.compile(r'[<>:"/\|?*]')
    if invalid_chars.search(filename):
        return False

    # Check if the filename ends with a space or a period.
    if filename.endswith(' ') or filename.endswith('.'):
        return False

    # If the filename passes all of the above checks, it is a valid and allowed filename in Windows.
    return True

def check_and_add_xml_extension(filename):
    if not filename.endswith(".xml"):
        print("no xml file extension adding one")
        filename += ".xml"
    return filename


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

    file = OutputEntry.get()

    if not is_allowed_filename_in_windows(file):
        print("not allowed")
        return


    file = check_and_add_xml_extension(file)

    if global_excel_file == "" or global_output_folder == "":
        print("not allowed")
        return
    
    
    print(type(global_excel_file), type(global_output_folder),type(file))
    Converter.start(global_excel_file,global_output_folder,file)

#creating a window variable
window = ctk.CTk()
window.iconbitmap("D:\myprograms\Python\ExcelWork\ExcelToXml\Resources\Icons\\favicon.ico")
window.title("ExcelToXml")
window.resizable(False, False)


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

Dframe = ctk.CTkFrame(master=frame,
                    width=960,
                    height=108,
                    corner_radius=10)

Dframe.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

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

OutputEntry = ctk.CTkEntry(master=Dframe,
                            placeholder_text="Enter a filename",
                            width=720,
                            height=25,
                            corner_radius=10)
OutputEntry.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

OutputEntry.configure(font=("Arial", 16), fg_color="black")


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

