#Versione 0.80 30/3/2024

import customtkinter
from CTkTable import *
from CTkTableRowSelector import *
from tkinter import filedialog as fd
import tkinter as tk
from tkinter.messagebox import askquestion
import pandas as pd
import pyexcel as p
import pyexcel_xls
import pyexcel_xlsx
import os


customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
        def __init__(self):
                super().__init__()

                # configure window
                self.title("Convertitore IBS")
                self.geometry(f"{590}x{250}")
                #self.iconbitmap(r"C:\Users\throa\Desktop\Programma paolo\ww2.ico") #TODO icona

                # configure grid layout (4x4)
                self.grid_columnconfigure(1, weight=1)
                self.grid_columnconfigure((2, 3), weight=0)
                self.grid_rowconfigure((0, 1, 2), weight=1)

                # create sidebar frame with widgets
                self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
                self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
                self.sidebar_frame.grid_rowconfigure(4, weight=1)

                self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text="Home", command=self.Home)
                self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
                self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text="Filtri", command=self.Filtri)
                self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
                
                self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="Zoom", anchor="w")
                self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
                self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["100%", "110%", "120%", "130%", "140%"],
                                                                command=self.change_scaling_event)
                self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

                # frame della home
                self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
                self.home_frame.grid(row=0, column=1, padx=(0, 0), pady=(0, 0), sticky="nsew")

                (self.percorso, self.destinazione)=self.load_defaults()

                self.pathLabel = customtkinter.CTkLabel(self.home_frame, text="Scegli il file da convertire", font=("Futura Std Book",15), anchor=customtkinter.W)
                self.pathLabel.grid(row=0, column=0, padx=17, pady=10)
                self.pathButton = customtkinter.CTkButton(self.home_frame, text="Sfoglia", command=self.get_path)
                self.pathButton.grid(row=0, column=1, padx=17, pady=10)
                self.pathLabel = customtkinter.CTkLabel(self.home_frame, text="Scegli dove salvarlo", font=("Futura Std Book",15))
                self.pathLabel.grid(row=1, column=0, padx=17, pady=10)
                self.salvaButton = customtkinter.CTkButton(self.home_frame, text="Sfoglia", command=self.get_directory)
                self.salvaButton.grid(row=1, column=1, padx=17, pady=10)
                self.aumentaprezzoLabel = customtkinter.CTkLabel(self.home_frame, text="Aumenta il prezzo?", font=("Futura Std Book",15))
                self.aumentaprezzoLabel.grid(row=3, column=0, padx=17, pady=10)
                self.percentualeaumento = customtkinter.CTkEntry(self.home_frame, placeholder_text="Inserisci una percentuale")
                self.percentualeaumento.grid(row=3, column=1, sticky="nsew")
                self.spacer1 = customtkinter.CTkButton(self.home_frame, text="Converti", state="disabled", fg_color="transparent",font=("Futura Std Book",15), text_color_disabled="#242424")
                self.spacer1.grid(row=4, column=0, columnspan=3, padx=17, pady=10)
                
                self.convertiButton = customtkinter.CTkButton(self.home_frame, text="Converti", command=self.convertitore_ibs)
                self.convertiButton.grid(row=5, column=1, padx=17, pady=10)         
                

                # frame dei filtri
                self.filtri_frame = customtkinter.CTkFrame(self, height=480, corner_radius=0, fg_color="transparent")
                self.filtri_frame.grid(row=0, column=1, padx=(0, 0), pady=(0, 0), sticky="nsew")

                self.frame_right = customtkinter.CTkScrollableFrame(master=self)
                self.frame_right.grid(row=1, column=1,pady=20)

                self.frame_bottom = customtkinter.CTkFrame(master=self)
                self.frame_bottom.grid(row=2, column=1,padx=20,pady=0)
                
                self.frame_bikinibottom = customtkinter.CTkFrame(master=self)
                self.frame_bikinibottom.grid(row=3, column=1,padx=20, pady=20)

                self.frame_top = customtkinter.CTkFrame(master=self)
                self.frame_top.grid(row=0, column=1)

                heading=[["Articolo","Nome","Casa Discografica","Genere","Prezzo MM","Prezzo IBS"]]
                self.tableh = CTkTable(master=self.frame_top, column=6, values=heading)
                self.tableh.pack(expand=True, fill="both")

                self.table = CTkTable(master=self.frame_right, column=6, values=self.get_values())
                self.table.pack(expand=True, fill="both")
                self.row_selector = CTkTableRowSelector(self.table, selected_row_color="dark blue", can_select_headers=True, max_selection=1)
                                
                self.textbox1=customtkinter.CTkEntry(master=self.frame_bottom)
                self.textbox1.grid(row=0,column=0)
                self.textbox2=customtkinter.CTkEntry(master=self.frame_bottom)
                self.textbox2.grid(row=0,column=1)
                self.textbox3=customtkinter.CTkEntry(master=self.frame_bottom)
                self.textbox3.grid(row=0,column=2)
                self.textbox4=customtkinter.CTkEntry(master=self.frame_bottom)
                self.textbox4.grid(row=0,column=3)
                self.textbox5=customtkinter.CTkEntry(master=self.frame_bottom)
                self.textbox5.grid(row=0,column=4)
                self.textbox6=customtkinter.CTkEntry(master=self.frame_bottom)
                self.textbox6.grid(row=0,column=5)

                self.bottoneupdate=customtkinter.CTkButton(master=self.frame_bikinibottom, text="Aggiorna",width=285,command=self.aggiorna)
                self.bottoneupdate.grid(row=0,column=1)
                self.bottoneaggiungi=customtkinter.CTkButton(master=self.frame_bikinibottom, text="Aggiungi filtro",width=285,command=self.add_value)
                self.bottoneaggiungi.grid(row=0,column=2)
                self.bottonecancella=customtkinter.CTkButton(master=self.frame_bikinibottom, text="Cancella filtro",width=285,command=self.cancella_filtro)
                self.bottonecancella.grid(row=0,column=3)

                # partenza dalla home
                self.filtri_frame.grid_forget()
                self.frame_bikinibottom.grid_forget()
                self.frame_top.grid_forget()
                self.frame_right.grid_forget()
                self.frame_bottom.grid_forget()

        
        def load_defaults(self):
                try:
                        f = open("default.txt", "r")
                        testo=f.read().splitlines()
                        perc=testo[0]
                        dest=testo[1]
                        
                except:
                        f = open("default.txt", "w")
                        f.close
                        f = open("default.txt", "r")
                        perc=""
                        dest=""
                f.close()
                return(perc,dest)
        
        def set_defaultpath(self):
                f = open("default.txt", "r")
                old=f.readlines()
                f.close()
                try:
                        old.pop(0)
                except:
                        print('')
                old.insert(0,self.percorso+"\n")
                f = open("default.txt", "w")
                f.writelines(old)
                f.close()

        def set_defaultdest(self):
                f = open("default.txt", "r")
                old=f.readlines()
                f.close()
                try:
                        old.pop(1)
                except:
                        print('')        
                old.insert(1,self.destinazione)
                f = open("default.txt", "w")
                f.writelines(old)
                f.close()

        def get_directory(self):
                self.destinazione=fd.askdirectory()
                print(self.destinazione)
                self.set_defaultdest()
        
        def get_path(self):
                self.percorso=fd.askopenfilename()
                print(self.percorso)
                self.set_defaultpath()

        def get_values(self):
                value=[]
                try:
                        f = open("filtri.txt", "r")
                        value1=f.readlines()
                        for v in value1:
                                v=str.split(v, ",")
                                value.append(v)
                except FileNotFoundError:
                        f = open("filtri.txt", "w")
                        f.close()
                        f = open("filtri.txt", "r")
                f.close()
                return value

        def add_value(self):
                f = open("filtri.txt", "a")
                f.write("%s,%s,%s,%s,%s,%s\n" %(self.textbox1.get(),self.textbox2.get(),self.textbox3.get(),self.textbox4.get(),self.textbox5.get(),self.textbox6.get()))
                f.close()
                self.textbox1.delete(0,tk.END)
                self.textbox2.delete(0,tk.END)                                        
                self.textbox3.delete(0,tk.END)                                        
                self.textbox4.delete(0,tk.END)                                        
                self.textbox5.delete(0,tk.END)
                self.textbox6.delete(0,tk.END)
                self.table.add_row("")
                self.table.update_values(self.get_values())

        def Home(self):
                self.geometry(f"{590}x{250}")
                self.filtri_frame.grid_forget()
                self.frame_bikinibottom.grid_forget()
                self.frame_top.grid_forget()
                self.frame_right.grid_forget()
                self.frame_bottom.grid_forget()
                self.home_frame.grid(row=0, column=1, padx=(0, 0), pady=(0, 0), sticky="nsew")
        
        def Filtri(self):
                self.geometry(f"{1037}x{453}")
                self.home_frame.grid_forget()
                self.frame_top.grid(row=0, column=1, sticky="nwe")
                self.frame_right.grid(row=1, column=1, sticky="nwe")
                self.frame_bottom.grid(row=2, column=1, sticky="nwe",pady=20)
                self.frame_bikinibottom.grid(row=3, column=1, sticky="nwe",pady=20)         

        def aggiorna(self):
                if self.textbox1.get()!="":
                        p.save_book_as(file_name=self.percorso, dest_file_name='file.xlsx')
                        p.free_resources()

                        xl = pd.ExcelFile('file.xlsx')
                        df = xl.parse()
                        for index, row in df.iterrows():
                                if self.textbox1.get()==str(row[0]):
                                        self.textbox2.delete(0,tk.END)                                        
                                        self.textbox3.delete(0,tk.END)                                        
                                        self.textbox4.delete(0,tk.END)                                        
                                        self.textbox5.delete(0,tk.END)
                                        
                                        self.textbox2.insert(0,row[4])
                                        self.textbox3.insert(0,row[13])
                                        self.textbox4.insert(0,row[11])
                                        self.textbox5.insert(0,row[28])
                        xl.close()
                        os.remove("file.xlsx")

        def cancella_filtro(self):
                rigaselezionata = self.row_selector.get()[0]
                rigapulita = str(rigaselezionata).replace("[","").replace("]","").replace(" '","").replace("'","").replace("\\n","")
                messaggio = "Cancellerai il filtro " + str(rigaselezionata) + ", sei sicuro?"
                answer = askquestion("Conferma", messaggio)
                if answer == 'yes':
                        with open("filtri.txt", "r") as f:
                                lines = f.readlines()
                        f.close
                        with open("filtri.txt", "w") as f:
                                for line in lines:
                                        if line.strip("\n") != rigapulita:
                                                f.write(line)
                        f.close
                        
                        self.table = CTkTable(master=self.frame, column=6, values=self.get_values())
                        self.table.grid(row=1,column=0)
                        self.row_selector = CTkTableRowSelector(self.table, selected_row_color="dark blue", can_select_headers=True, max_selection=1)
                


        def change_scaling_event(self, new_scaling: str):
                new_scaling_float = int(new_scaling.replace("%", "")) / 100
                customtkinter.set_widget_scaling(new_scaling_float)
                customtkinter.set_window_scaling(new_scaling_float)
        
        def convertitore_ibs(self):
                p.save_book_as(file_name=self.percorso, dest_file_name='file.xlsx')
                p.free_resources()

                xl = pd.ExcelFile('file.xlsx')
                df = xl.parse()
                
                if self.percentualeaumento.get() == "":
                        percentuale=0
                else:
                        percentuale=int(self.percentualeaumento.get())
                
                percentuale=(percentuale/100)+1

                for i, row in df.iterrows():
                        df.at[i,'codart'] = a="{:05d}".format(row[0])
                        df.at[i,'ean'] = a="{:013d}".format(row[1])
                        df.at[i,'prezzov'] = (round(float(row[28])*percentuale)-0.10)
                        df.at[i,'codgenere'] = a ="{:03d}".format(row[11])
                        if row[33]==0:
                                df.drop(i, inplace = True)

                #CODICE FILTRI
                codice=[]
                genere=[]
                prezzoc=[]
                prezzoorigc=[]
                prezzog=[]
                prezzoorigg=[]
                filtriattuali=self.get_values()
                for filtroattuale in filtriattuali:
                        if filtroattuale[0]=="":
                                genere.append("{:03d}".format(int(filtroattuale[3])))
                                prezzoorigg.append(filtroattuale[4])
                                prezzog.append(filtroattuale[5].removesuffix("\n"))
                        else:
                                codice.append(filtroattuale[0])
                                prezzoorigc.append(filtroattuale[4])                     
                                prezzoc.append(filtroattuale[5].removesuffix("\n"))

                for i, row in df.iterrows():
                        for g in genere:
                                if g==df.at[i,'codgenere']:
                                        indice=genere.index(g)
                                        if "%" in prezzog[indice]:                                                
                                                percent=(float(prezzog[indice].removesuffix("%"))/100)+1
                                                df.at[i,'prezzov'] = (round(float(prezzoorigg[indice])*percent)-0.10)
                                        else:
                                                df.at[i,'prezzov'] = (round(float(prezzog[indice]))-0.10)
                        for c in codice:
                                if c==df.at[i,'codart']:
                                        indice=codice.index(c)
                                        if "%" in prezzoc[indice]:
                                                percent=(float(prezzoc[indice].removesuffix("%"))/100)+1
                                                df.at[i,'prezzov'] = (round(float(prezzoorigc[indice])*percent)-0.10)
                                        else:
                                                df.at[i,'prezzov'] = (round(float(prezzoc[indice]))-0.10)


                df.drop(["eanc", "titolo", "titolopos", "codpos", "codsup", "anno", "codcant", "desccant", "codinterp", "codgenere", "codlinea", "codfor", "catalogo", "codiva", "codum", "confezione", "ricarica", "fuoricat", "datains", "datamod", "datasospo", "datafass", "indtipocp", "videof", "audiof", "codlineafor", "prezzoa", "iva", "ricaricar", "marginer", "datainizio", "datafine"], axis=1, inplace= True)

                ean = []
                pic = []
                undici = []
                void = []

                for index, row in df.iterrows():
                        ean.append('EAN')
                        pic.append('PIC')
                        undici.append(11)
                        void.append('')

                df["product-id-type"] = ean
                df["state"] = undici
                df["logistic-class"] = pic
                df["discount-price"] = void
                df["update-delete"] = void
                df["leadtime-to-ship"] = void

                df = df.rename(columns={'codart':'sku','ean':'product-id', 'prezzov':'price', 'esist':'quantity'})
                df=df[['sku', 'product-id', 'product-id-type', 'price', 'quantity', 'state', 'logistic-class', 'discount-price', 'update-delete', 'leadtime-to-ship']]
                
                file_name = self.destinazione+"\OFFERTE.xlsx"
                df.to_excel(file_name, index=False)

                xl.close()
                os.remove("file.xlsx")
 

if __name__ == "__main__":
    app = App()
    app.mainloop()