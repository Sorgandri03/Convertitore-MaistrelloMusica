import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import askquestion
import customtkinter
from CTkTable import *
from CTkTableRowSelector import *

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)
        
        self.frame = customtkinter.CTkFrame(master=self)
        self.frame.grid(row=0, column=0, sticky="nswe",padx=20,pady=20)

        self.table = CTkTable(master=self.frame, column=6, values=self.get_values())
        self.table.grid(row=1,column=0)
        self.row_selector = CTkTableRowSelector(self.table, selected_row_color="dark blue", can_select_headers=True, max_selection=1)

        self.bottonecancella=customtkinter.CTkButton(master=self.frame, text="Cancella filtro",width=285,command=self.cancella_filtro)
        self.bottonecancella.grid(row=2,column=0)

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
        
    def cancella_filtro(self):
        rigaselezionata = self.row_selector.get()[0]
        rigapulita = str(rigaselezionata).replace("[","").replace("]","").replace(" '","").replace("'","").replace("\\n","")
        messaggio = "Cancellerai il filtro " + str(rigaselezionata) + ", sei sicuro?"
        answer = askquestion("Conferma", messaggio)
        if answer == 'yes':
                self.table.destroy()
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
                
                
        

app = App()
app.mainloop()