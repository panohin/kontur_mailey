import os.path
import pythoncom
from win32com.client import Dispatch
from mail_adresses import *
from config import *
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox as mb

def create_from_dict(dict_of_files):
    for file_name in dict_of_files.keys():
        path = dict_of_files[file_name]
        create(file_name, path)
def send_files(subject, body, user):
    files = os.listdir(os.getcwd() + "\\result\\")
    for file in files:
        print(f"ФАЙЛ {file}")
        send(file, subject, body, user)
def create(file_name, path):
    try:
        excel = Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(path)
        work_folder = os.getcwd() + "\\result\\"
        print(f"{work_folder=}")

        #new_window = Toplevel()
        #Label(new_window, text="Получатель:", width=15, borderwidth = 3).grid(row=0, column=0)
        #Label(new_window, text="Шаблон:", width=40, borderwidth = 3).grid(row=0, column=1)
        #Label(new_window, text="Статус:", width=15, borderwidth = 3).grid(row=0, column=2)
        #status_label = Label(new_window, text='asd').grid(row=0, column=3)
        
        nmbr = 1
        for name_of_group_of_sheet in group_of_sheet.keys():
            #Label(new_window, text=f"{name_of_group_of_sheet}").grid(row=nmbr, column=0)
            print(f"{name_of_group_of_sheet=}")
            for name_of_sheet in group_of_sheet[name_of_group_of_sheet]:
                #Label(new_window, text=f"{name_of_sheet}").grid(row=nmbr, column=1)
                print(f"{name_of_sheet=}")
                if not os.path.exists(work_folder + f"{name_of_group_of_sheet}.xlsx"):
                    #status_label.config(text=f"create new workbook {name_of_group_of_sheet}")
                    new_workbook = excel.Workbooks.Add()
                    new_workbook.SaveAs(work_folder + f"{name_of_group_of_sheet}.xlsx")
                    new_workbook.Close()
                    print("create new workbook")
                else:
                    #status_label['text'] = f"workbook {name_of_group_of_sheet} already exists"
                    print(f"workbook {name_of_group_of_sheet} already exists")
                try:
                    if workbook.Worksheets(name_of_sheet):
                        new_workbook = excel.Workbooks.Open(work_folder + f"{name_of_group_of_sheet}.xlsx")
                        worksheet = workbook.Worksheets(name_of_sheet)
                        worksheet.Copy(Before=new_workbook.Worksheets(1))
                        new_workbook.Save()
                        new_workbook.Close()
                        #status_label['text'] = f"Sheet {name_of_sheet} added to {name_of_group_of_sheet} workbook"
                        print(f"Sheet {name_of_sheet} added to {name_of_group_of_sheet} workbook")
                        nmbr += 1
                except pythoncom.com_error:
                    #status_label['text'] = f"Sheet {name_of_sheet} is not in sourse workbook"
                    print(f"Sheet {name_of_sheet} is not in sourse workbook")
            
        workbook.Close()
        excel.Quit()
        #new_window.mainloop()
    except Exception as e:
        print(e)
        print("Ошибка в создании файлов")
        mb.showerror("Ошибка", "Ошибка отправки.\nПри повторении обратитесь к разработчику\n+7 495 155 1717 доб.477\nАнохин Павел")
        exit()
def send(file, subject, body, user):       
    try:
        excel = Dispatch("Excel.Application")
        outlook = Dispatch('outlook.application')
        work_folder = os.getcwd() + "\\result\\"
        workbook_to_send = excel.Workbooks.Open(work_folder + f"{file}")
        mail = outlook.CreateItem(0)
        a = "".join(file.split(".")[:-1])
        print(a, len(workbook_to_send.Sheets), mail_adresses.get(a))
        if len(workbook_to_send.Sheets) > user and mail_adresses.get(a) is not None:
            mail.To = mail_adresses.get(a)
            mail.CC = COPY
            #mail.To = "anokhin.p@a1tis.ru"
            mail.Subject = subject
            mail.Body = body
            attachment  = work_folder + f"{file}"
            mail.Attachments.Add(attachment)
            mail.Send()
            print(f"{file} sent")
            #Label(new_window, text="Отправлено").grid(row=nmbr, column=2)
        workbook_to_send.Close()
        os.remove(work_folder + f"{file}")
    except Exception as e:
        print(e)
        workbook_to_send.Close()
        mb.showerror("Ошибка", "Ошибка отправки.\nПри повторении обратитесь к разработчику\n+7 495 155 1717 доб.477\nАнохин Павел")
        exit()
        



