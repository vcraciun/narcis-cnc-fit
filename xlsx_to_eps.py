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
import json

Rectangle = namedtuple('Rectangle', ['x', 'y', 'w', 'h', 'name'])

class CNCConvert:
    def __init__(self):
        self.root = Tk()
        self.root.title("Conversie documente EXCEL la format grafic vectorial - Narcis CNC Design")
        self.root.geometry('640x290')
        self.root.resizable(False, False)

        config = {}
        if os.path.isfile("config.json"):
            config = json.load(open("config.json", 'r'))

        Label(self.root, text='Fisier EXCEL: ').grid(column=0, row=0)
        self.path = Entry(self.root, width = 80)
        self.path.grid(column=1, row=0, columnspan=2)
        Button(self.root, text='Deschide', command=self.file_add).grid(column=3, row=0)

        Label(self.root, text='Fisiere Output: ', font='Helvetica 16 bold').grid(column=1, row=1, sticky="W")
        self.pdf = IntVar()
        self.svg = IntVar()
        self.eps = IntVar()
        if len(config) == 0:
            self.pdf.set(1)
            self.svg.set(0)
            self.eps.set(0)
        else:
            self.pdf.set(config["pdf"])
            self.svg.set(config["svg"])
            self.eps.set(config["eps"])
        Checkbutton(self.root, text='PDF', variable=self.pdf).grid(column=1, row=2, sticky="W")
        Checkbutton(self.root, text='SVG', variable=self.svg).grid(column=1, row=3, sticky="W")
        Checkbutton(self.root, text='EPS', variable=self.eps).grid(column=1, row=4, sticky="W")

        Label(self.root, text='Formatul aranjarii: ', font='Helvetica 16 bold').grid(column=2, row=1, sticky="W")
        self.format = IntVar()
        if len(config) == 0:
            self.format.set(1)
        else:
            self.format.set(config["format"])
        Radiobutton(self.root, text="Format liniar de-a lungul axei X", variable=self.format, value=1).grid(column=2, row=2, sticky="W")
        Radiobutton(self.root, text="Format Rectangular XY CU rotatii", variable=self.format, value=2).grid(column=2, row=3, sticky="W")
        Radiobutton(self.root, text="Format Rectangular XY FARA rotatii", variable=self.format, value=3).grid(column=2, row=4, sticky="W")
        
        Label(self.root, text='Etichete: ', font='Helvetica 16 bold').grid(column=1, row=5, sticky="W")
        self.etichete = IntVar()
        if len(config) == 0:
            self.etichete.set(1)
        else:
            self.etichete.set(config["etichete"])
        Checkbutton(self.root, text='Afiseaza', variable=self.etichete).grid(column=1, row=6, sticky="W")

        Label(self.root, text='Mareste de: ').grid(column=1, row=7, sticky="W")
        self.scale2 = Entry(self.root, width = 15)
        self.scale2.grid(column=1, row=7, columnspan=1, sticky="E")
        self.scale2.delete(0, END)            
        if len(config) == 0:
            self.scale2.insert(0, "7")
        else:
            self.scale2.insert(0, config["scale_fonts"])
        
        Label(self.root, text='Prefix: ').grid(column=1, row=8, sticky="W")
        self.prefix = Entry(self.root, width = 15)
        self.prefix.grid(column=1, row=8, sticky="E")
        self.prefix.delete(0, END)
        if len(config) == 0:
            self.prefix.delete(0, END)

        self.scale1 = IntVar()
        if len(config) == 0:
            self.scale1.set(10)
        else:
            self.scale1.set(config["scale_squares"])

        Label(self.root, text='Scala Forme Geometrice: ', font='Helvetica 16 bold').grid(column=2, row=5, sticky="W")        
        Radiobutton(self.root, text="1 %", variable=self.scale1, value=100).grid(column=2, row=6, sticky="W")
        Radiobutton(self.root, text="10 %", variable=self.scale1, value=10).grid(column=2, row=7, sticky="W")
        Radiobutton(self.root, text="20 %", variable=self.scale1, value=5).grid(column=2, row=8, sticky="W")                
        Radiobutton(self.root, text="30 %", variable=self.scale1, value=3).grid(column=2, row=6, sticky="E")
        Radiobutton(self.root, text="50 %", variable=self.scale1, value=2).grid(column=2, row=7, sticky="E")
        Radiobutton(self.root, text="100 %", variable=self.scale1, value=1).grid(column=2, row=8, sticky="E")                
        
        Button(self.root, text='Proceseaza', command=self.process).grid(column=2, row=9, sticky="W")

        self.status=StringVar()        
        Label(self.root, bd=1, relief=SUNKEN, width=70, anchor=W,textvariable=self.status,font=('arial',12,'normal')).grid(column=0, row=10, columnspan=4, sticky="W")
        self.status.set('Pregatit ... ')

    def SaveConfig(self):
        data = {    
            "pdf": self.pdf.get(),
            "svg": self.svg.get(),
            "eps": self.eps.get(),
            "format": self.format.get(),
            "etichete": self.etichete.get(),
            "scale_squares": self.scale1.get(),
            "scale_fonts": self.scale2.get()
        }
        with open("config.json", 'w') as f:
            json.dump(data, f)
        self.status.set('Am salvat noua configuratie ...')

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
    
    def phsppog(self, width, rectangles, sorting="width"):
        if sorting not in ["width", "height" ]:
            raise ValueError("The algorithm only supports sorting by width or height but {} was given.".format(sorting))
        if sorting == "width":
            wh = 0
        else:
            wh = 1
        result = [None] * len(rectangles)
        remaining = deepcopy(rectangles)
        sorted_indices = sorted(range(len(remaining)), key=lambda x: -remaining[x][wh])
        sorted_rect = [remaining[idx] for idx in sorted_indices]
        x, y, w, h, H = 0, 0, 0, 0, 0
        while sorted_indices:
            idx = sorted_indices.pop(0)
            r = remaining[idx]
            result[idx] = Rectangle(x, y, r[0], r[1], r[2])
            x, y, w, h, H = r[0], H, width - r[0], r[1], H + r[1]
            self.recursive_packing(x, y, w, h, 0, remaining, sorted_indices, result)
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
        scale_sq = 1 / self.scale1.get()
        scale_fn = float(self.scale2.get())

        mode = self.format.get()        
        if mode == 2:        
            height, rectangles = self.phspprg(4500*scale_sq, self.data)            
            print("Inaltimea aranjarii este: ", height)
            self.status.set('Aranjez chenarele dreptuncghiular CU rotiri ... ')
        elif mode == 3:
            height, rectangles = self.phsppog(4500*scale_sq, self.data)
            print("Inaltimea aranjarii este: ", height)
            self.status.set('Aranjez chenarele dreptuncghiular FARA rotiri ... ')
        else:
            rectangles = []
            x = 20            
            self.status.set('Aranjez chenarele de-a lungul axei X ... ')
            for rect in self.initial_data:
                y = 20
                for i in range(rect[3]):
                    rectangles += [Rectangle(x, y, rect[2], rect[1], rect[0])]    
                    y += 20 + rect[1]
                x += 20 + rect[2]                                
            self.status.set('Aranjez chenarele de-a lungul axei X : GATA')

        unit.set(defaultunit="mm")
        c = canvas.canvas()    
        for sq in rectangles:
            rect = path.path(path.moveto(sq.x*scale_sq, sq.y * scale_sq), path.lineto(sq.x * scale_sq, (sq.y+sq.h)*scale_sq), path.lineto((sq.x + sq.w)*scale_sq, (sq.y+sq.h)*scale_sq), path.lineto((sq.x + sq.w)*scale_sq, sq.y*scale_sq), path.closepath())
            if self.etichete.get() == 1:
                name = prefix + str(sq.name)
                if sq.w < 300:
                    c.text((sq.x + sq.w/2)*scale_sq, (sq.y + sq.h/2)*scale_sq, name, [text.halign.center, text.vshift.mathaxis, scale(scale_fn), rotate(90)])
                else:
                    c.text((sq.x + sq.w/2)*scale_sq, (sq.y + sq.h/2)*scale_sq, name, [text.halign.center, text.vshift.mathaxis, scale(scale_fn)])
            c.stroke(rect)                 

        if self.eps.get() == 1:      
            self.status.set('Scriu: EPS')
            c.writeEPSfile(fis)
        if self.pdf.get() == 1:
            self.status.set('Scriu: PDF')
            c.writePDFfile(fis)
        if self.svg.get() == 1:
            self.status.set('Scriu: SVG')
            c.writeSVGfile(fis)        
        self.status.set('Am Terminat !!!')

    def file_add(self):
        fname = filedialog.askopenfilename(initialdir = "./", title = "Selectati fisier Excel", filetypes = (("Fisiere EXCEL", "*.xlsx*"), ("Toate fisierele", "*.*")))
        self.path.delete(0, END)
        self.path.insert(0, fname)
        _, name = os.path.split(fname)
        prefix = name[:5] + '-'
        self.prefix.delete(0, END)
        self.prefix.insert(0, prefix)

        df = pd.DataFrame(pd.read_excel(fname))
        c1 = list(df.iloc[:,1])
        c2 = list(df.iloc[:,2])
        c3 = list(df.iloc[:,3])
        c4 = list(df.iloc[:,4])
        self.initial_data = [(c1[i], c2[i], c3[i], c4[i]) for i in range(len(c1)) if type(c1[i]) == int and type(c2[i]) == int and type(c3[i]) == int]
        
        self.data = []
        for item in self.initial_data:
            for i in range(item[3]):
                self.data += [[item[1], item[2], item[0]]]

        self.status.set(f'Fisier incarcat: [{len(self.initial_data)} unice] / [{len(self.data)} total]')

    def slow_process(self):
        excel_file = self.path.get()
        if len(excel_file):        
            self.ConvertToVectorial(self.prefix.get(), excel_file)

    def process(self):
        self.SaveConfig()
        threading.Thread(target=self.slow_process).start()

def main():
    converter = CNCConvert()
    converter.run()    

if __name__ == "__main__":
    main()
