import pandas as pd
import numpy
from pyx import *
from tkinter import *
from tkinter import filedialog
import threading

def ConvertToVectorial(rezolutie, spatiu, fis):
    df = pd.DataFrame(pd.read_excel(fis))
    c1 = list(df.iloc[:,1])
    c2 = list(df.iloc[:,2])
    c3 = list(df.iloc[:,3])
    c4 = list(df.iloc[:,4])
    data = [(c1[i], c2[i], c3[i], c4[i]) for i in range(len(c1)) if type(c1[i]) == int and type(c2[i]) == int and type(c3[i]) == int]
 
    unit.set(defaultunit="mm")
    c = canvas.canvas()    
    startx = spatiu/rezolutie
    print(len(data))
    for el in data:
        print(el)
        starty = spatiu/rezolutie
        crtx = el[2]/rezolutie
        crty = el[1]/rezolutie
        for i in range(el[3]):
            rect = path.path(path.moveto(startx, starty), path.lineto(startx, starty+crty), path.lineto(startx + crtx, starty+crty), path.lineto(startx + crtx, starty), path.closepath())
            c.stroke(rect)
            c.text(startx + crtx/2, starty + crty/2, str(el[0]))
            starty += crty
            starty += spatiu/rezolutie      
        startx += crtx
        startx += spatiu/rezolutie    
    if eps.get() == 1:      
        c.writeEPSfile(fis)
    if pdf.get() == 1:
        c.writePDFfile(fis)
    if svg.get() == 1:
        c.writeSVGfile(fis)        

def file_add():
    fname = filedialog.askopenfilename(initialdir = "./", title = "Selectati fisier Excel", filetypes = (("Fisiere EXCEL", "*.xlsx*"), ("Toate fisierele", "*.*")))
    root.path.delete(0, END)
    root.path.insert(0, fname)

def slow_process():
    data = root.path.get()
    if len(data):
        rezolutie = float(root.rez.get())
        distanta = float(root.dist.get())
        ConvertToVectorial(rezolutie, distanta, data)

def process():
    threading.Thread(target=slow_process).start()

root = Tk()
root.title("Conversie documente EXCEL la format grafic vectorial - Narcis CNC Design")
root.geometry('640x130')
root.resizable(False, False)

Label(root, text='Fisier EXCEL: ').grid(column=0, row=0)
root.path = Entry(root, width = 80)
root.path.grid(column=1, row=0)
Button(root, text='Deschide', command=file_add).grid(column=2, row=0)

pdf = IntVar()
svg = IntVar()
eps = IntVar()
pdf.set(1)
svg.set(1)
eps.set(1)
Checkbutton(root, text='PDF', variable=pdf).grid(column=0, row=1)
Checkbutton(root, text='SVG', variable=svg).grid(column=1, row=1)
Checkbutton(root, text='EPS', variable=eps).grid(column=2, row=1)

Label(root, text='Rezolutie: ').grid(column=0, row=2)
root.rez = Entry(root, width = 80)
root.rez.grid(column=1, row=2)
root.rez.delete(0, END)
root.rez.insert(0, "100")

Label(root, text='Distanta: ').grid(column=0, row=3)
root.dist = Entry(root, width = 80)
root.dist.grid(column=1, row=3)
root.dist.delete(0, END)
root.dist.insert(0, "20")

Button(root, text='Proceseaza', command=process).grid(column=1, row=4)

root.mainloop()
    



