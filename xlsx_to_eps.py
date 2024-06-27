import pandas as pd
from pyx import *
from pyx.trafo import rotate, scale
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
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
        self.root.geometry('640x410')
        self.root.resizable(False, False)

        self.tab_parent = ttk.Notebook(self.root)
        self.tab1 = ttk.Frame(self.tab_parent)
        self.tab2 = ttk.Frame(self.tab_parent)
        self.tab_parent.add(self.tab1, text = "Control")
        self.tab_parent.add(self.tab2, text = "Profile")
        self.tab_parent.pack(expand=1, fill='both')

        self.profiles = {"Fara Profil": {}}

        self.config = {}
        if os.path.isfile("config.json"):
            self.config = json.load(open("config.json", 'r'))     

        if os.path.isdir("profile"):
            self.LoadProfiles()        

        #---------------------------------------------- TAB 1 -------------------------------------------------------
        Label(self.tab1, text='Fisier EXCEL: ').grid(column=0, row=0)
        self.path = Entry(self.tab1, width = 80)
        self.path.grid(column=1, row=0, columnspan=2)
        Button(self.tab1, text='Deschide', command=self.file_add).grid(column=3, row=0)

        Label(self.tab1, text='Fisiere Output: ', font='Helvetica 16 bold').grid(column=1, row=1, sticky="W")
        self.pdf = IntVar()
        self.svg = IntVar()
        self.eps = IntVar()
        if len(self.config) == 0:
            self.pdf.set(1)
            self.svg.set(0)
            self.eps.set(0)
        else:
            self.pdf.set(self.config["pdf"])
            self.svg.set(self.config["svg"])
            self.eps.set(self.config["eps"])
        Checkbutton(self.tab1, text='PDF', variable=self.pdf).grid(column=1, row=2, sticky="W")
        Checkbutton(self.tab1, text='SVG', variable=self.svg).grid(column=1, row=3, sticky="W")
        Checkbutton(self.tab1, text='EPS', variable=self.eps).grid(column=1, row=4, sticky="W")

        Label(self.tab1, text='Formatul aranjarii: ', font='Helvetica 16 bold').grid(column=2, row=1, sticky="W")
        self.format = IntVar()
        if len(self.config) == 0:
            self.format.set(1)
        else:
            self.format.set(self.config["format"])
        Radiobutton(self.tab1, text="Format liniar de-a lungul axei X", variable=self.format, value=1).grid(column=2, row=2, sticky="W")
        Radiobutton(self.tab1, text="Format Rectangular XY CU rotatii", variable=self.format, value=2).grid(column=2, row=3, sticky="W")
        Radiobutton(self.tab1, text="Format Rectangular XY FARA rotatii", variable=self.format, value=3).grid(column=2, row=4, sticky="W")
        
        Label(self.tab1, text='Etichete: ', font='Helvetica 16 bold').grid(column=1, row=5, sticky="W")
        self.etichete = IntVar()
        if len(self.config) == 0:
            self.etichete.set(1)
        else:
            self.etichete.set(self.config["etichete"])
        Checkbutton(self.tab1, text='Afiseaza', variable=self.etichete).grid(column=1, row=6, sticky="W")

        Label(self.tab1, text='Mareste de: ').grid(column=1, row=7, sticky="W")
        self.scale2 = Entry(self.tab1, width = 15)
        self.scale2.grid(column=1, row=7, columnspan=1, sticky="E")
        self.scale2.delete(0, END)            
        if len(self.config) == 0:
            self.scale2.insert(0, "0.7")
        else:
            self.scale2.insert(0, self.config["scale_fonts"])
        
        Label(self.tab1, text='Prefix: ').grid(column=1, row=8, sticky="W")
        self.prefix = Entry(self.tab1, width = 15)
        self.prefix.grid(column=1, row=8, sticky="E")
        self.prefix.delete(0, END)
        if len(self.config) == 0:
            self.prefix.delete(0, END)

        self.scale1 = IntVar()
        if len(self.config) == 0:
            self.scale1.set(10)
        else:
            self.scale1.set(self.config["scale_squares"])

        Label(self.tab1, text='Scala Forme Geometrice: ', font='Helvetica 16 bold').grid(column=2, row=5, sticky="W")        
        Radiobutton(self.tab1, text="1 %", variable=self.scale1, value=100, command=self.cmd_100).grid(column=2, row=6, sticky="W")
        Radiobutton(self.tab1, text="10 %", variable=self.scale1, value=10, command=self.cmd_10).grid(column=2, row=7, sticky="W")
        Radiobutton(self.tab1, text="20 %", variable=self.scale1, value=5, command=self.cmd_5).grid(column=2, row=8, sticky="W")                
        Radiobutton(self.tab1, text="30 %", variable=self.scale1, value=3, command=self.cmd_3).grid(column=2, row=6, sticky="E")
        Radiobutton(self.tab1, text="50 %", variable=self.scale1, value=2, command=self.cmd_2).grid(column=2, row=7, sticky="E")
        Radiobutton(self.tab1, text="100 %", variable=self.scale1, value=1, command=self.cmd_1).grid(column=2, row=8, sticky="E")                
        
        Label(self.tab1, text='Spatiu: ').grid(column=1, row=9, sticky="W")
        self.spatiu = Entry(self.tab1, width = 15)
        self.spatiu.grid(column=1, row=9, sticky="E")
        self.spatiu.delete(0, END)
        if len(self.config) == 0:
            self.spatiu.delete(0, END)
            self.spatiu.insert(0, "20")
        else:
            self.spatiu.delete(0, END)
            if "spatiu" in self.config:
                self.spatiu.insert(0, self.config["spatiu"])
            else:
                self.spatiu.insert(0, "20")      

        Label(self.tab1, text='Profil: ', font='Helvetica 16 bold').grid(column=1, row=10, sticky="W")
        self.clicked = StringVar()
        self.clicked.set(list(self.profiles.keys())[0])
        self.drop = OptionMenu(self.tab1, self.clicked, *list(self.profiles.keys()))
        self.drop.grid(column=1, row=11, sticky="W")

        Button(self.tab1, text='Proceseaza', command=self.process).grid(column=2, row=12, sticky="W")

        self.status=StringVar()        
        Label(self.tab1, bd=2, relief=SUNKEN, width=70, anchor=W,textvariable=self.status,font=('arial',12,'normal')).grid(pady = 10, column=0, row=13, columnspan=4, sticky="W")
        self.status.set('Pregatit ... ')
        
        #---------------------------------------------- TAB 2 -------------------------------------------------------
        Label(self.tab2, text='Selectie: ', font='Helvetica 16 bold').grid(column=0, row=0, sticky="W")
        self.clicked2 = StringVar()
        self.clicked2.set(list(self.profiles.keys())[0])
        self.drop2 = OptionMenu(self.tab2, self.clicked2, *list(self.profiles.keys()), command=self.initialize_profile)
        self.drop2.grid(column=0, row=1, sticky="W")

        Label(self.tab2, text='Config: ', font='Helvetica 16 bold').grid(column=0, row=2, sticky="W")
        Label(self.tab2, text='Nume: ').grid(column=0, row=3, sticky="E")
        self.profil_name = Entry(self.tab2, width = 87)
        self.profil_name.grid(column=1, row=3)
        Label(self.tab2, text='Suprapunere: ').grid(column=0, row=4, sticky="E")
        self.clicked3 = StringVar()
        self.optiuni_eroare = ["ignora", "linie"]
        self.clicked3.set(self.optiuni_eroare[0])
        self.drop3 = OptionMenu(self.tab2, self.clicked3, *self.optiuni_eroare)
        self.drop3.grid(column=1, row=4, sticky="W")
        Label(self.tab2, text='Chenare: ').grid(column=0, row=5, sticky="E")
        self.distante = Entry(self.tab2, width = 87)
        self.distante.grid(column=1, row=5)
        Button(self.tab2, text='Salveaza', command=self.salveaza_profil).grid(column=0, row=6, sticky="W")
        Button(self.tab2, text='Adauga', command=self.adauga_profil).grid(column=1, row=6, sticky="W")        

    def LoadProfiles(self):
        files = [os.path.join("profile", ff) for ff in os.listdir("profile") if ff.endswith(".json")]
        for fl in files:
            path, name = os.path.split(fl)
            fname, fext = os.path.splitext(name)
            ff = open(fl, 'r')
            profil = json.load(ff)
            ff.close()
            self.profiles[fname] = profil

    def cmd_100(self):
        self.scale2.delete(0, END)            
        if "_100" in self.config:
            self.scale2.insert(0, self.config["_100"])    
        else:
            self.scale2.insert(0, "0.07")
            self.config["_100"] = "0.07"

    def cmd_10(self):
        self.scale2.delete(0, END)            
        if "_10" in self.config:
            self.scale2.insert(0, self.config["_10"])    
        else:
            self.scale2.insert(0, "0.7")
            self.config["_100"] = "0.7"

    def cmd_5(self):
        self.scale2.delete(0, END)            
        if "_5" in self.config:
            self.scale2.insert(0, self.config["_5"])    
        else:
            self.scale2.insert(0, "1.4")
            self.config["_100"] = "1.4"

    def cmd_3(self):
        self.scale2.delete(0, END)            
        if "_3" in self.config:
            self.scale2.insert(0, self.config["_3"])    
        else:
            self.scale2.insert(0, "2.5")
            self.config["_100"] = "2.5"

    def cmd_2(self):
        self.scale2.delete(0, END)            
        if "_2" in self.config:
            self.scale2.insert(0, self.config["_2"])    
        else:
            self.scale2.insert(0, "3.5")
            self.config["_100"] = "3.5"

    def cmd_1(self):
        self.scale2.delete(0, END)                    
        if "_1" in self.config:
            self.scale2.insert(0, self.config["_1"])    
        else:
            self.scale2.insert(0, "7")   
            self.config["_100"] = "7"                                 

    def SaveConfig(self):
        data = {    
            "pdf": self.pdf.get(),
            "svg": self.svg.get(),
            "eps": self.eps.get(),
            "format": self.format.get(),
            "etichete": self.etichete.get(),
            "scale_squares": self.scale1.get(),
            "scale_fonts": self.scale2.get(),
            "spatiu": self.spatiu.get(),
            "_100": "0.07",
            "_10": "0.7",
            "_5": "1.4",
            "_3": "2.5",
            "_2": "3.5",
            "_1": "7"
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
        spatiu = int(self.spatiu.get())
        profil = self.profiles[self.clicked.get()]
        supra = 0 if profil['suprapunere'] == 'ignora' else 1
        chenare = profil['chenare']
        print(f"profil utilizat: {self.clicked.get()}, suprapunere: {supra}, chenare: {chenare}")

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
            top_rectangles = []
            x = spatiu          
            self.status.set('Aranjez chenarele de-a lungul axei X ... ')
            for rect in self.initial_data:
                y = spatiu
                for i in range(rect[3]):
                    rectangles += [Rectangle(x, y, rect[2], rect[1], rect[0])]    
                    top_rectangles += [Rectangle(x, y, rect[2], rect[1], rect[0])]    
                    for offset in chenare:                        
                        if supra == 0 or (supra == 1 and offset < rect[2]-2*offset and offset < rect[1]-2*offset):
                            rectangles += [Rectangle(x+offset, y+offset, rect[2]-2*offset, rect[1]-2*offset, rect[0])]    
                        else:
                            if offset >= rect[2]-2*offset:
                                rectangles += [Rectangle(x+rect[2]//2, y+offset, 1, rect[1]-2*offset, rect[0])]    
                            if offset >= rect[1]-2*offset:
                                rectangles += [Rectangle(x+offset, y+rect[1]//2, rect[2]-2*offset, 1, rect[0])]
    
                    y += spatiu + rect[1]
                x += spatiu + rect[2]                                
            self.status.set('Aranjez chenarele de-a lungul axei X : GATA')

        unit.set(defaultunit="mm")
        
        c = canvas.canvas()    
        for sq in rectangles:
            rect = path.path(path.moveto(sq.x*scale_sq, sq.y * scale_sq), path.lineto(sq.x * scale_sq, (sq.y+sq.h)*scale_sq), path.lineto((sq.x + sq.w)*scale_sq, (sq.y+sq.h)*scale_sq), path.lineto((sq.x + sq.w)*scale_sq, sq.y*scale_sq), path.closepath())
            c.stroke(rect)                 

        for sq in top_rectangles:
            if self.etichete.get() == 1:
                name = prefix + str(sq.name)
                if sq.w < 300:
                    c.text((sq.x + sq.w/2)*scale_sq, (sq.y + sq.h/2)*scale_sq, name, [text.halign.center, text.vshift.mathaxis, scale(scale_fn), rotate(90)])
                else:
                    c.text((sq.x + sq.w/2)*scale_sq, (sq.y + sq.h/2)*scale_sq, name, [text.halign.center, text.vshift.mathaxis, scale(scale_fn)])

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

    def UpdateOptionMenus(self, new_name):
        self.clicked2.set(new_name)
        self.drop2.destroy()
        self.drop.destroy()
        self.drop2 = OptionMenu(self.tab2, self.clicked2, *list(self.profiles.keys()), command=self.initialize_profile)        
        self.drop = OptionMenu(self.tab1, self.clicked, *list(self.profiles.keys()))        
        self.drop2.grid(column=0, row=1, sticky="W")
        self.drop.grid(column=1, row=11, sticky="W")

    def salveaza_profil(self):        
        new_name = self.profil_name.get()        
        existing_name = self.clicked2.get()
        chenare = self.distante.get()
        
        self.profiles[new_name] = {
            "suprapunere": self.clicked3.get(),
            "chenare": [float(item) for item in chenare.split(',')]
        }

        if new_name != existing_name:
            os.unlink(os.path.join("profile", existing_name + '.json'))
            self.profiles.pop(existing_name)

        json.dump(self.profiles[new_name], open(os.path.join("profile", new_name + '.json'), 'w'))    
        self.LoadProfiles()
        self.UpdateOptionMenus(new_name)

    def adauga_profil(self):
        new_name = self.profil_name.get()
        chenare = self.distante.get()

        self.profiles[new_name] = {
            "suprapunere": self.clicked3.get(),
            "chenare": [float(item) for item in chenare.split(',')]
        }

        json.dump(self.profiles[new_name], open(os.path.join("profile", new_name + '.json'), 'w')) 
        self.LoadProfiles()
        self.UpdateOptionMenus(new_name)

    def initialize_profile(self, caption):
        if "Fara Profil" in caption:
            self.profil_name.delete(0, END)                
            self.distante.delete(0, END)
            self.clicked3.set("ignora")
            return
        data = self.profiles[caption]
        self.profil_name.delete(0, END)            
        self.profil_name.insert(0, caption)
        self.clicked3.set(data["suprapunere"])
        self.distante.delete(0, END)
        self.distante.insert(0, ",".join([str(item) for item in data["chenare"]]))

def main():
    converter = CNCConvert()
    converter.run()    

if __name__ == "__main__":
    main()
