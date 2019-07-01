from tkinter import *
import os
import shutil
import glob
import xlwt
import zipfile
from pygame import mixer

# ---------------  Function that generates 2nd Window from the button in 1st Window. -----------------

def open_window():
    window1 = Toplevel()
    window1.title("Operation on Files")
    window1.geometry("500x500")
    window1.configure(background="Turquoise")
    window1.iconbitmap("spider.ico")
    photo = PhotoImage(file="images.PNG")
    labelphoto = Label(window1,image=photo).pack()
    photo1=PhotoImage(file="excel.PNG")
    labelphoto1 = Label(window1,image=photo1).pack()
    photo2 = PhotoImage(file = "zip.PNG")
    labelphoto2 = Label(window1,image=photo2).pack()
    butt1 = Button(window1,text="Data Copying",command=data_copying,font=('comicsans',12)).pack()
    butt2 = Button(window1,text="Excel List",command=excel_list,font=('comicsans',12)).pack()
    butt3 = Button(window1,text="Creating Archive",command=zipp,font=('comicsans',12)).pack()
    window1.mainloop()

# ----------------------------------Generating primary window. -------------------------------------------

window = Tk()
welcome_text = Label(window, text="Welcome to SPIDER ver. 1.0", font=("comicsans",20), background = "turquoise").pack()
button = Button(window, text="Operation on Files", command=open_window).pack()
window.geometry("1000x1000")
window.title("Spider")
window.iconbitmap("spider.ico")
window.configure(background="Turquoise")

# -------------------------- Functions providing operation on files. -----------------------------

def data_copying():
    src_dir = r"C:\Users\Piotr Surma\Desktop\Programming\Spychacz"
    dst_dir = r"C:\Users\Piotr Surma\Desktop\Programming\Hari pota"
    sourceFiles = os.listdir(src_dir)
    for files in sourceFiles:
        name = os.path.join(src_dir, files)
        if os.path.isfile(name):
            shutil.copy(name, dst_dir)


def excel_list():
    src_dir = r"C:\Users\Piotr Surma\Desktop\Programming\Spychacz"
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


def zipp():
    src_dir = r"C:\Users\Piotr Surma\Desktop\Programming\Spychacz"
    os.chdir(src_dir)
    for dirname, subdirs, files in os.walk(src_dir):
        zippo = zipfile.ZipFile("Archiwum.zip", "w", zipfile.ZIP_DEFLATED)
        for filename in files:
            zippo.write(os.path.join(filename))
    zippo.close()

# ------------------------- Initializing Music ------------------------------------

mixer.init()
def playing_music():
    mixer.music.load("welcome.mp3")
    mixer.music.play()

def stoping_music():
    mixer.music.stop()

def set_vol(val):
    volume = int(val)/100
    mixer.music.set_volume(volume)

playPhoto = PhotoImage(file = "music.PNG")
playBtn = Button(window, image = playPhoto, command=playing_music).pack()
butt4 = Button(window, text = "Stop Music", command= stoping_music, font=('comicsans', 8)).pack()

scale = Scale(window, from_=0, to=100, orient = HORIZONTAL, command = set_vol).pack()


window.mainloop()

