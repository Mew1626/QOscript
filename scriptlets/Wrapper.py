import os
import datetime
from tkinter import *
from tkinter import filedialog
from tkinter import ttk

root = Tk()
reports = StringVar()
under10min = StringVar()
monthselection = 0

warantycsv = ""
purchasecsv = ""
evalcsv = ""
loncsv = ""

def wrap():
    path = os.path.dirname(os.path.realpath(__file__))
    os.chdir(path)
    os.chdir("../")
    excelpath = os.getcwd()
    os.system('C:/Anaconda3/python.exe "' + os.path.join(path, "QO1.py") + '" "' + os.path.join(excelpath, "Quality_Objective_1_2019.xlsx") + '" ' + reports.get() + " " + under10min.get() + " " + str(month.current()+2)) 
    os.system('C:/Anaconda3/python.exe "' + os.path.join(path, "QO2.py") + '" "' + os.path.join(excelpath, "Quality_Objective_2_2019.xlsx") + '" "' + purchasecsv + '" "' + warantycsv+ '"') 
    os.system('C:/Anaconda3/python.exe "' + os.path.join(path, "QO3.py") + '" "' + os.path.join(excelpath, "Quality_Objective_3_2019.xlsx") + '" "' + evalcsv + '"') 
    os.system('C:/Anaconda3/python.exe "' + os.path.join(path, "QO4.py") + '" "' + os.path.join(excelpath, "Quality_Objective_4_2019.xlsx") + '" "' + loncsv + '"') 

def browse(iden):
    global warantycsv, purchasecsv, evalcsv, loncsv
    root.filename = filedialog.askopenfilename(initialdir = "C:\\Users\\iScree Laptop\\Downloads",title = "Select file",filetypes = (("csv file","*.csv"),("all files","*.*")))
    if iden == 0: 
        purchasecsv = root.filename
        pur.set(root.filename)
    elif iden == 1: 
        warantycsv = root.filename
        war.set(root.filename)
    elif iden == 2: 
        evalcsv = root.filename
        eva.set(root.filename)
    elif iden == 3: 
        loncsv = root.filename
        lon.set(root.filename)

def gui():

    root.title("Quality Objectives")

    mainframe = ttk.Frame(root, padding="3 3 12 12")
    mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    reports_entry = ttk.Entry(mainframe, width=7, textvariable=reports)
    under10min_entry = ttk.Entry(mainframe, width=7, textvariable=under10min)
    reports_entry.grid(column=1, row= 2, sticky= (W, E))
    under10min_entry.grid(column=2, row= 2, sticky= (W, E))

    ttk.Label(mainframe, text="Reports").grid(column=1, row= 1, sticky= (W, E))
    ttk.Label(mainframe, text="Under 10 Min").grid(column=2, row= 1, sticky= (W, E))

    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
    ttk.Label(mainframe, text="Month:").grid(column=1, row= 3, sticky= (W, E))
    monthvar = StringVar()
    global month
    month = ttk.Combobox(mainframe, textvariable=monthvar)
    month['values'] = months
    month.grid(column=2, row= 3, sticky= (W, E))
    month.current(datetime.date.today().month - 2)

    ttk.Button(mainframe, text="Execute", command=wrap).grid(column=3, row= 2, sticky= (W, E))

    global pur, war, eva, lon 
    pur = StringVar()
    war = StringVar()
    eva = StringVar()
    lon = StringVar()

    ttk.Label(mainframe, text="Purchase CSV").grid(column=1, row= 4, sticky= (W, E))
    ttk.Entry(mainframe, textvariable=pur).grid(column=2, row=4, sticky=(W, E))
    ttk.Button(mainframe, text="Browse", command=lambda:browse(0)).grid(column=3, row= 4, sticky= (W, E))
    
    ttk.Label(mainframe, text="Waranty CSV").grid(column=1, row= 5, sticky= (W, E))
    ttk.Entry(mainframe, textvariable=war).grid(column=2, row=5, sticky=(W, E))
    ttk.Button(mainframe, text="Browse", command=lambda:browse(1)).grid(column=3, row= 5, sticky= (W, E))
    
    ttk.Label(mainframe, text="Eval CSV").grid(column=1, row= 6, sticky= (W, E))
    ttk.Entry(mainframe, textvariable=eva).grid(column=2, row=6, sticky=(W, E))
    ttk.Button(mainframe, text="Browse", command=lambda:browse(2)).grid(column=3, row= 6, sticky= (W, E))
    
    ttk.Label(mainframe, text="Lon CSV").grid(column=1, row= 7, sticky= (W, E))
    ttk.Entry(mainframe, textvariable=lon).grid(column=2, row=7, sticky=(W, E))
    ttk.Button(mainframe, text="Browse", command=lambda:browse(3)).grid(column=3, row= 7, sticky= (W, E))

    for child in mainframe.winfo_children(): 
        child.grid_configure(padx=5, pady=5)
    root.mainloop()

def main():
    gui()

main()