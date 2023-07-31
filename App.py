from Tabela import appGUI
from CincoDias import cinco
from QuinzeDias import quinze
from TrintaDias import trinta
from tkinter import *
class mainApp():
    def __init__(self):
        self.mainroot  = Tk()
        self.mainroot.title("Relatório de Operações")
        self.mainroot.geometry("500x130")
        self.mainroot.resizable(False,False)
        #Criando instruções na tela
        titleMain = Label(self.mainroot, text="Escolha o tipo de tabela:",font=28)
        titleMain.place(x=25,y=10)

        #Criando opções
        btnFive = Button(self.mainroot,text="5 Dias",width=10,font=24,command=self.selectFive)
        btnFive.place(x=50,y=50)

        btnFifteen = Button(self.mainroot,text="15 Dias",width=10,font=24,command=self.selectFifteen)
        btnFifteen.place(x=200,y=50)

        btnThirty = Button(self.mainroot,text="30 Dias",width=10,font=24,command=self.selectThirty)
        btnThirty.place(x=350,y=50)
        self.mainroot.mainloop()
    def selectFive(self):
        self.mainroot.destroy() 
        name = "ROP Cinco Dias"
        table = appGUI(cinco,mainApp,name)
    def selectFifteen(self):
        self.mainroot.destroy() 
        name = "ROP Quinze Dias"
        table = appGUI(quinze,mainApp,name)
    def selectThirty(self):
        self.mainroot.destroy()
        name = "ROP Trinta Dias"
        table = appGUI(trinta,mainApp,name)       
main = mainApp()
main