import os.path
import datetime
import ctypes
import pythoncom
from win32com.client import Dispatch

from mail_adresses import *
from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox as mb
from tkinter.ttk import *
from main import *
from config import BODY, COPY



def add_file():
    full_path = fd.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if full_path:
        file_name_with_extens = full_path.split("/")[-1]
        file_name = " ".join(file_name_with_extens.split(".")[:-1])
        if file_name not in listbox.get(0, END):
            listbox.insert(END, file_name)
            dict_of_files[file_name] = full_path
            print(f"{dict_of_files=}")
        else:
            mb.showinfo("Выбрать файл", "Файл уже в списке")
    else:
        mb.showinfo("Выбрать файл", "Файл не выбран")

def create_and_send():
    subject = theme_ent.get()
    body = body_text.get(1.0, END)
    create_from_dict(dict_of_files)
    user = r_var.get()
    print(f"USER = {user}")
    send_files(subject, body, user)
    mb.showinfo("Выполнено", "Выгрузки успешно отправлены")

def delete_from_listbox(event):
    select = listbox.curselection()
    del dict_of_files[listbox.get(select)]
    print(f"{dict_of_files=}")
    listbox.delete(select)

    
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('favicon') # favicon

master = Tk()
master.title("KonturMailey")
master.iconbitmap('favicon.ico')
width = master.winfo_screenwidth() // 2
height = master.winfo_screenheight() // 2
master.geometry(f'+{width-master.winfo_width()}+{height-master.winfo_height()}')
master.resizable(width=False, height=False)

letter_frame = LabelFrame(master)
letter_frame.pack(pady=2, padx=5, ipadx=5, ipady=2)
theme_LabelFrame = LabelFrame(letter_frame, text="Тема письма:")
theme_LabelFrame.pack(anchor=W, pady=2, padx=10, ipadx=5, ipady=2)
theme_ent = Entry(theme_LabelFrame, font="Calibri 9")
subject = f"Новые тендеры {datetime.date.today().strftime('%d.%m')}"
theme_ent.insert(0, subject)
theme_ent.pack()
rb_label_frame = LabelFrame(letter_frame, text="Пользователь")
r_var = IntVar()
r_var.set(3)
user_rb1 = Radiobutton(rb_label_frame, variable=r_var, value=3,  text="Мелюхина")
user_rb2 = Radiobutton(rb_label_frame, variable=r_var, value=1, text="Анохин")
user_rb1.pack(anchor=W)
user_rb2.pack(anchor=W)
rb_label_frame.pack(anchor=W, pady=2, padx=10, ipadx=5, ipady=2)

body_LabelFrame = LabelFrame(letter_frame, text="Текст письма:")
body_LabelFrame.pack(anchor=W, pady=2, padx=10, ipadx=5, ipady=2)
body_text = Text(body_LabelFrame, width=40, height=6, wrap=NONE, font="Calibri 10")
body_text.insert(1.0, BODY)
body_text.pack()

dict_of_files = {}
list_of_files_labelFrame = LabelFrame(master, text="Список выгрузок Контура для обработки:")
list_of_files_labelFrame.pack(pady=2, padx=5, ipadx=5, ipady=2)
listbox = Listbox(list_of_files_labelFrame, height=3, width=50)
listbox.pack()
listbox.bind("<Delete>", delete_from_listbox)
open_but = Button(list_of_files_labelFrame, text="+ Добавить файл", command=add_file)
open_but.pack()

send_button = Button(master, text="Обработать выгрузки", command=create_and_send)
send_button.pack(pady=5)

master.mainloop()







