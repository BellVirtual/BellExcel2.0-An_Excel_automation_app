import tkinter as tk
from tkinter import *
from tkinter import filedialog
import os
import pandas as pd
from PIL import ImageTk, Image



##########generate filepath
'getcwd:      ', os.getcwd()
osfile, maindir =('__file__:    ', __file__)
mainpath = maindir.replace('Bell_Excel_2.0.py',"BulkFile.xlsx")




#########globals
root = tk.Tk()
root.resizable(width=False,height=False,)
tfont = ("Courier", 10)
tbg = "black"
tfg = "green2"

#############generates window
canvas = tk.Canvas(root, height=475, width= 425)
canvas.pack()
backroundimage = tk.PhotoImage(file="tielvapor.png")
backround = tk.Label(canvas, image=backroundimage)
backround.place(relwidth=1, relheight=1)



########generates frame
frame = tk.Frame(root)
frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)
pilframeimage = Image.open("BellVirtual_Icon.png")
resized_pilframeimage = pilframeimage.resize((100, 100), Image.ANTIALIAS)
frameimage = ImageTk.PhotoImage(resized_pilframeimage)
framelabel = tk.Label(frame, image=frameimage)
framelabel.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)







##########add file name button
bellfiles = []
labelnames = []

def add_file():
    filepath = filedialog.askopenfilename(
        initialdir="/", title="Select file", filetypes=[("Excel files", ".xlsx .xls")])
    bellfiles.append(filepath)
    for file in bellfiles:
        filepath, filename = os.path.split(file)
        filetag = tk.Label(frame, text=filename, bg=tbg, fg= tfg, font=tfont)
        filetag.config(font=("Courier", 10))
        filetag.pack()


add_file_button = Button(root, text="Choose Excel File", command=add_file, bg=tbg,fg=tfg,font=tfont)
add_file_button.place(relx=.33, rely=0, relwidth=.4, relheight=.08)






############# data processing button
def clearframe():
    for widget in frame.winfo_children():
        if widget != framelabel:
            widget.destroy()
def process_files():
    for file in bellfiles:
        excelframe = pd.read_excel(file)
        dataframes = [excelframe]
        compactframes = pd.concat(dataframes)
        compactframes.to_excel(mainpath)
        clearframe()

add_file_button = Button(root, text="Process files", command=process_files, bg=tbg,fg=tfg,font=tfont)
add_file_button.place(relx=.25, rely=.92, relwidth=.5, relheight=.08)






#####Run app
root.mainloop()

