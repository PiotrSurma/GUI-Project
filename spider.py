from tkinter import *
import os
import shutil
import glob
import xlwt
import zipfile


def open_window():
    top = Toplevel()
    top.title("Operation on Files")
    top.geometry("500x500")
    top.configure(background = "Turquoise")
    top.iconbitmap("spider.ico")
    photo = PhotoImage(file = "images.PNG")
    labelphoto = Label(top, image = photo)
    labelphoto.pack()
    photo1 = PhotoImage(file= 'excel.PNG')
    labelphoto1 = Label(top, image = photo1)
    labelphoto1.pack()
    butt1 = Button(top, text = "Data Copying", command = data_copying(), \
    font = ('comicsans', 12)).pack()
    butt2 = Button(top, text = "Excel List", command = excel_list(), \
    font = ('comicsans', 12)).pack()
    butt3 = Button(top, text = "Creating Archive", command = zippek(), \
    font = ('comicsans', 12)).pack()
    top.mainloop()


window = Tk()
welcome_text = Label(window, text = "Welcome to SPIDER ver. 1.0",\
font = ("comicsans",20), background = "turquoise").pack()
button = Button(window, text = "Operation on Files", command = open_window).pack()
window.geometry("1000x1000")
window.title("Spider")
window.iconbitmap("spider.ico")
window.configure(background = "Turquoise")


def data_copying():
    src_dir = r"C:\Users\Piotr Surma\Desktop\Spychacz"
    dst_dir = r"C:\Users\Piotr Surma\Desktop\Hari pota"
    sourceFiles = os.listdir(src_dir)
    for files in sourceFiles:
        name = os.path.join(src_dir, files)
        if os.path.isfile(name):
            shutil.copy(name, dst_dir)


def excel_list():
    src_dir = r"C:\Users\Piotr Surma\Desktop\Spychacz"
    os.chdir(src_dir)
    lista_plikow1=[]
    for files in glob.iglob(os.path.join(src_dir, "**.**")):
        lista_plikow1.append(files)
        list1= lista_plikow1
        book = xlwt.Workbook(encoding = "utf-8")
        sheet1 = book.add_sheet("Sheet 1")
        sheet1.write(0, 0, "Files:")
        i=2
        for n in list1:
            sheet1.write(i, 0, n)
            i = i+1
        book.save("Lista.xls")

def zippek():
    src_dir = r"C:\Users\Piotr Surma\Desktop\Spychacz"
    os.chdir(src_dir)
    for dirname, subdirs, files in os.walk(src_dir):
        zippo = zipfile.ZipFile("Archiwum.zip", "w", zipfile.ZIP_DEFLATED)
        for filename in files:
            zippo.write(os.path.join(filename))
    zippo.close()

window.mainloop()

