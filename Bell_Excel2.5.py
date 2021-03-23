
import os
import sys
import tkinter as tk
from tkinter import filedialog
from tkinter import *
from PIL import ImageTk, Image
import pandas as pd



##########generate filepath
'getcwd:      ', os.getcwd()
osfile, maindir =('__file__:    ', __file__)
filename = os.path.basename(sys.argv[0])
bulkpath = maindir.replace(filename,"BulkFile.xlsx")
seperatedpath = maindir.replace(filename,"Excels")




#########globals
root = tk.Tk()
tfont = ("", 11)
tbg = "black"
tfg = "white"
ltfg = "white"
btfont = ("", 13)
btfg = "cyan"
bellfiles = []




#############configs window
root.title('Bell Excel')
root.resizable(width=False,height=False,)
root.iconbitmap("imgs/Bell_Excel_App_Icon.ico")



#############generates backround
canvas = tk.Canvas(root, height=475, width= 425)
canvas.pack()
backroundimage = tk.PhotoImage(file="imgs/tielvapor.png")
backround = tk.Label(canvas, image=backroundimage)
backround.place(relwidth=1, relheight=1)




########generates listbox
list_frame = tk.Frame(root,background=tbg)
list_frame.place(relwidth=0.8, relheight=0.365, relx=0.1, rely=0.15)
scrollbar = Scrollbar(list_frame, orient="vertical",bg=tbg)
xl_list = Listbox(list_frame,width=50,height=10,yscrollcommand=scrollbar.set,bg=tbg,fg=ltfg,font=tfont)
scrollbar.config(command=xl_list.yview)
scrollbar.pack(side="right", fill="y")
xl_list.configure(justify=CENTER)
xl_list.pack()
#464
########generates emblem
ico_frame = tk.Frame(root,background=tbg)
ico_frame.place(relwidth=0.8, relheight=0.365, relx=0.1, rely=0.515)
pilframeimage = Image.open("imgs/Bell_Virtual_Emblem.png")
resized_pilframeimage = pilframeimage.resize((150, 150), Image.ANTIALIAS)
frameimage = ImageTk.PhotoImage(resized_pilframeimage)
framelabel = tk.Label(ico_frame, image=frameimage)
framelabel.place(relwidth=0.4, relheight=0.8, relx=0.3, rely=0.1)






##########add file button
def add_file():
    filepath = filedialog.askopenfilename(
        initialdir="/", title="Select file", filetypes=[("Excel files", ".xlsx .xls")])
    bellfiles.append(filepath)
    path ,filename = os.path.split(filepath)
    xl_list.insert(END, filename)


add_file_button = Button(root, text="Choose an Excel File", command=add_file, bg=tbg,fg=tfg,font=tfont)
add_file_button.place(relx=.3, rely=0, relwidth=.4, relheight=.05)





###### add dir button

def read_xl_dir(folderpath):
    for root, dirs, files in os.walk(folderpath):
        for f in files:
            filepath = os.path.join(root, f)
            path, filetext = os.path.split(filepath)
            name, extension = os.path.splitext(filepath)
            if extension == '.xlsx':
                bellfiles.append(filepath)
                xl_list.insert(END,filetext)
def add_dir():
    fpath = filedialog.askdirectory()
    read_xl_dir(fpath)


add_dir_button = Button(root, text="Choose a Folder", command=add_dir, bg=tbg,fg=tfg,font=tfont)
add_dir_button.place(relx=.3, rely=.05, relwidth=.4, relheight=.05)





############# data processing button

def process_files():
    if bellfiles == '':
        return
    for file in bellfiles:
        compact_xl(file)
        cleardata()

def cleardata():
    bellfiles.clear()
    xl_list.delete(0,END)


process_file_button = Button(root, text="Process Files", command=process_files, bg=tbg,fg=btfg,font=btfont)
process_file_button.place(relx=.23, rely=.92, relwidth=.5, relheight=.08)





#########pandas function
def compact_xl(filepath):
    excelframe = pd.read_excel(filepath)
    dataframes = [excelframe]
    compactframes = pd.concat(dataframes)
    compactframes.to_excel(bulkpath)





#####Run app
root.mainloop()