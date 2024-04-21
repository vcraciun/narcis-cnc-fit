import pandas as pd
from pyx import *
from pyx.trafo import rotate, scale
from tkinter import *
from tkinter import filedialog
import threading
import os
from copy import deepcopy
from collections import namedtuple
import sys

Rectangle = namedtuple('Rectangle', ['x', 'y', 'w', 'h', 'name'])

class CNCConvert:
    def __init__(self):
        self.root = Tk()
        self.root.title("Conversie documente EXCEL la format grafic vectorial - Narcis CNC Design")
        self.root.geometry('640x110')
        self.root.resizable(False, False)

        Label(self.root, text='Fisier EXCEL: ').grid(column=0, row=0)
        self.root.path = Entry(self.root, width = 80)
        self.root.path.grid(column=1, row=0)
        Button(self.root, text='Deschide', command=self.file_add).grid(column=2, row=0)

        self.pdf = IntVar()
        self.svg = IntVar()
        self.eps = IntVar()
        self.pdf.set(1)
        self.svg.set(1)
        self.eps.set(1)
        Checkbutton(self.root, text='PDF', variable=self.pdf).grid(column=0, row=1)
        Checkbutton(self.root, text='SVG', variable=self.svg).grid(column=1, row=1)
        Checkbutton(self.root, text='EPS', variable=self.eps).grid(column=2, row=1)

        Label(self.root, text='Prefix: ').grid(column=0, row=2)
        self.root.rez = Entry(self.root, width = 80)
        self.root.rez.grid(column=1, row=2)
        self.root.rez.delete(0, END)

        Button(self.root, text='Proceseaza', command=self.process).grid(column=1, row=3)

    def run(self):
        self.root.mainloop()

    def phspprg(self, width, rectangles, sorting="width"):
        if sorting not in ["width", "height" ]:
            raise ValueError("The algorithm only supports sorting by width or height but {} was given.".format(sorting))
        if sorting == "width":
            wh = 0
        else:
            wh = 1
        result = [None] * len(rectangles)
        remaining = deepcopy(rectangles)
        for idx, r in enumerate(remaining):
            if r[0] > r[1]:
                remaining[idx][0], remaining[idx][1] = remaining[idx][1], remaining[idx][0]
        sorted_indices = sorted(range(len(remaining)), key=lambda x: -remaining[x][wh])
        sorted_rect = [remaining[idx] for idx in sorted_indices]
        x, y, w, h, H = 0, 0, 0, 0, 0
        while sorted_indices:
            idx = sorted_indices.pop(0)
            r = remaining[idx]
            if r[1] > width:
                result[idx] = Rectangle(x, y, r[0], r[1], r[2])
                x, y, w, h, H = r[0], H, width - r[0], r[1], H + r[1]
            else:
                result[idx] = Rectangle(x, y, r[1], r[0], r[2])
                x, y, w, h, H = r[1], H, width - r[1], r[0], H + r[0]
            self.recursive_packing(x, y, w, h, 1, remaining, sorted_indices, result)
            x, y = 0, H
        return H, result

    def recursive_packing(self, x, y, w, h, D, remaining, indices, result):
        priority = 6
        for idx in indices:
            for j in range(0, D + 1):
                if priority > 1 and remaining[idx][(0 + j) % 2] == w and remaining[idx][(1 + j) % 2] == h:
                    priority, orientation, best = 1, j, idx
                    break
                elif priority > 2 and remaining[idx][(0 + j) % 2] == w and remaining[idx][(1 + j) % 2] < h:
                    priority, orientation, best = 2, j, idx
                elif priority > 3 and remaining[idx][(0 + j) % 2] < w and remaining[idx][(1 + j) % 2] == h:
                    priority, orientation, best = 3, j, idx
                elif priority > 4 and remaining[idx][(0 + j) % 2] < w and remaining[idx][(1 + j) % 2] < h:
                    priority, orientation, best = 4, j, idx
                elif priority > 5:
                    priority, orientation, best = 5, j, idx
        if priority < 5:        
            if orientation == 0:
                omega, d = remaining[best][0], remaining[best][1]
            else:
                omega, d = remaining[best][1], remaining[best][0]
            result[best] = Rectangle(x, y, omega, d, remaining[best][2])
            indices.remove(best)
            if priority == 2:
                self.recursive_packing(x, y + d, w, h - d, D, remaining, indices, result)
            elif priority == 3:
                self.recursive_packing(x + omega, y, w - omega, h, D, remaining, indices, result)
            elif priority == 4:
                min_w = sys.maxsize
                min_h = sys.maxsize
                for idx in indices:
                    min_w = min(min_w, remaining[idx][0])
                    min_h = min(min_h, remaining[idx][1])
                # Because we can rotate:
                min_w = min(min_h, min_w)
                min_h = min_w
                if w - omega < min_w:
                    self.recursive_packing(x, y + d, w, h - d, D, remaining, indices, result)
                elif h - d < min_h:
                    self.recursive_packing(x + omega, y, w - omega, h, D, remaining, indices, result)
                elif omega < min_w:
                    self.recursive_packing(x + omega, y, w - omega, d, D, remaining, indices, result)
                    self.recursive_packing(x, y + d, w, h - d, D, remaining, indices, result)
                else:
                    self.recursive_packing(x, y + d, omega, h - d, D, remaining, indices, result)
                    self.recursive_packing(x + omega, y, w - omega, h, D, remaining, indices, result)

    def ConvertToVectorial(self, prefix, fis):
        df = pd.DataFrame(pd.read_excel(fis))
        c1 = list(df.iloc[:,1])
        c2 = list(df.iloc[:,2])
        c3 = list(df.iloc[:,3])
        c4 = list(df.iloc[:,4])
        initial_data = [(c1[i], c2[i], c3[i], c4[i]) for i in range(len(c1)) if type(c1[i]) == int and type(c2[i]) == int and type(c3[i]) == int]
        
        data = []
        for item in initial_data:
            for i in range(item[3]):
                data += [[item[1], item[2], item[0]]]

        width = 4500
        height, rectangles = self.phspprg(width, data)

        unit.set(defaultunit="mm")
        c = canvas.canvas()    
        for sq in rectangles:
            rect = path.path(path.moveto(sq.x, sq.y), path.lineto(sq.x, sq.y+sq.h), path.lineto(sq.x + sq.w, sq.y+sq.h), path.lineto(sq.x + sq.w, sq.y), path.closepath())
            name = prefix + str(sq.name) + '-' + str(sq.w) + 'x' + str(sq.h)
            if sq.w < 300:
                c.text(sq.x + sq.w/2, sq.y + sq.h/2, name, [text.halign.center, text.vshift.mathaxis, scale(7), rotate(90)])
            else:
                c.text(sq.x + sq.w/2, sq.y + sq.h/2, name, [text.halign.center, text.vshift.mathaxis, scale(7)])
            c.stroke(rect)        
        if self.eps.get() == 1:      
            c.writeEPSfile(fis)
        if self.pdf.get() == 1:
            c.writePDFfile(fis)
        if self.svg.get() == 1:
            c.writeSVGfile(fis)        

    def file_add(self):
        fname = filedialog.askopenfilename(initialdir = "./", title = "Selectati fisier Excel", filetypes = (("Fisiere EXCEL", "*.xlsx*"), ("Toate fisierele", "*.*")))
        self.root.path.delete(0, END)
        self.root.path.insert(0, fname)
        _, name = os.path.split(fname)
        prefix = name[:5] + '-'
        self.root.rez.insert(0, prefix)

    def slow_process(self):
        excel_file = self.root.path.get()
        if len(excel_file):        
            self.ConvertToVectorial(self.root.rez.get(), excel_file)

    def process(self):
        threading.Thread(target=self.slow_process).start()

def main():
    converter = CNCConvert()
    converter.run()    

if __name__ == "__main__":
    main()
