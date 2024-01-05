#Versione 0.6 5/1/2024

import customtkinter
from tkinter import filedialog as fd
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
                self.iconbitmap(r"C:\Users\throa\Desktop\Programma paolo\ww2.ico") #TODO icona

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

                self.percorso=""
                self.destinazione=""
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


                # partenza dalla home
                self.filtri_frame.grid_forget()

        def get_directory(self):
               self.destinazione=fd.askdirectory()
               print(self.destinazione)

        def get_path(self):
                self.percorso=fd.askopenfilename()
                print(self.percorso)

        def Home(self):
                self.geometry(f"{590}x{250}")
                self.filtri_frame.grid_forget()
                self.home_frame.grid(row=0, column=1, padx=(0, 0), pady=(0, 0), sticky="nsew")
        
        def Filtri(self):
                self.geometry(f"{750}x{750}")
                self.home_frame.grid_forget()
                self.filtri_frame.grid(row=0, column=1, padx=(0, 0), pady=(0, 0), sticky="nsew")                

        def change_scaling_event(self, new_scaling: str):
                new_scaling_float = int(new_scaling.replace("%", "")) / 100
                customtkinter.set_widget_scaling(new_scaling_float)
                customtkinter.set_window_scaling(new_scaling_float)
        
        def convertitore_ibs(self):
                p.save_book_as(file_name=self.percorso, dest_file_name='file.xlsx')

                xl = pd.ExcelFile('file.xlsx')
                df = xl.parse()

                if self.percentualeaumento.get() == "":
                        percentuale=0
                else:
                        percentuale=int(self.percentualeaumento.get())
                
                percentuale=(percentuale/100)+1

                i=0
                for index, row in df.iterrows():
                        df.at[i,'codart'] = a="{:05d}".format(row[0])
                        df.at[i,'ean'] = a="{:013d}".format(row[1])
                        df.at[i,'prezzov'] = (round(row[28]*percentuale)-0.01)
                        if row[33]==0:
                                df.drop(i, inplace = True)
                        i=i+1

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

                #os.remove("file.xlsx")
 

if __name__ == "__main__":
    app = App()
    app.mainloop()