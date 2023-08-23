import os
import win32com.client as win32
from tkinter import *
from tkinter import ttk
from PIL import ImageTk, Image
from tkinter import messagebox
from openpyxl import Workbook
class appGUI:
    def __init__(self,objs,backwin,nmwin, **kwargs):
        ropObj = objs.copy()
        self.backwin = backwin
        self.nmwin = nmwin
        self.root = Tk()
        #Criando tela da janela
        self.window = Frame()
        self.root.title(self.nmwin)
        self.root.geometry("1000x430")
        self.root.resizable(False,False)
        self.FontSize = 20
        self.root.bind("<Double-1>",self.onClick)
        #Criando Tabela na tela
        self.table = ttk.Treeview(self.window, height=15)
        style = ttk.Style(self.table)
        style.configure('Treeview', rowheight=30)
        self.table['columns'] = ['Objeto de Inspeção', 'Ação','Verificado','Obs.']
        self.table.column('#0', width=0)
        self.table.column('Objeto de Inspeção', anchor=W, width=350)        
        self.table.column('Ação', anchor=CENTER, width=300)
        self.table.column('Verificado', anchor=CENTER,width=100)       
        self.table.column('Obs.', anchor=W ,width=200)
        self.table.heading('Objeto de Inspeção',text='Objeto de Inspeção')       
        self.table.heading('Ação',text='Ação')        
        self.table.heading('Verificado',text='Verificado')        
        self.table.heading('Obs.',text='Obs.')
        #preenchendo a tabela com os dados json
        for i in ropObj:
            self.table.insert(parent='',index='end',values=(i["nome"],i["acao"],"Não","..."))
        #modificando alguns estilos da tabela e barra de rolagem lateral
        ttk.Style(self.root).theme_use('classic')
        sb = Scrollbar(self.window, orient=VERTICAL)
        sb.pack(side=RIGHT, fill=Y)
        self.table.config(yscrollcommand=sb.set)
        sb.config(command=self.table.yview)
        self.table.pack(pady=20)
        self.window.pack()
        #Criando opções de botões para salvar,data,elaborador etc
        btn = Button(self.root,text="Salvar arquivo",command=self.saveXL)
        btn.place(x=885,y=390)
        self.datalabel = ttk.Label(self.root,text="DATA: ")
        self.datalabel.place(x=560,y=362)
        validatecom = (self.root.register(self.validateDate),'%P')
        validatecommonth = (self.root.register(self.validatemonth),'%P')
        validatecomyear = (self.root.register(self.validateYear),'%P')
        self.dayentry = ttk.Entry(self.root,validate="key",validatecommand=validatecom, width=2)
        self.dayentry.place(x=610,y=360)
        self.monthentry = ttk.Entry(self.root,validate="key",validatecommand=validatecommonth, width=2)
        self.monthentry.place(x=640,y=360)   
        self.yearentry = ttk.Entry(self.root,validate="key",validatecommand=validatecomyear, width=4)
        self.yearentry.place(x=670,y=360)
        bar1 = ttk.Label(self.root,text="/",font=self.FontSize)
        bar2 = ttk.Label(self.root,text="/",font=self.FontSize)
        bar1.place(x=630,y=358)
        bar2.place(x=660,y=358)
        self.elblabel = ttk.Label(self.root,text="ELABORADOR: ")
        self.elblabel.place(x=520,y=392)
        self.elbentry = ttk.Entry(self.root, width=25)
        self.elbentry.place(x=610,y=390)
        bckbtn = Button(self.root,text="Voltar",command=self.goBack)
        bckbtn.place(x=5,y=360)
        #loop da janela
        self.root.mainloop()
    def onClick(self,event):
        #identificar região do click
        clickedReg = self.table.identify_region(event.x,event.y) 
           
        #Só precisamos da regiao "cell"
        if clickedReg not in ("cell"):
            return
        #identificar coluna
        column = self.table.identify_column(event.x)
        selectID = self.table.focus()
        selectvalue = self.table.item(selectID)
        #Coluna de observação
        selecttext = selectvalue.get("values")[3]
        columnBox = self.table.bbox(selectID, column)
        if column == "#4":
            entryEdit = ttk.Entry(self.table, width=columnBox[2])
            entryEdit.place(x=columnBox[0],y=columnBox[1],w=columnBox[2],h=columnBox[3])
            entryEdit.insert(0,selecttext)
            entryEdit.focus()
            entryEdit.select_range(0,1000)
            entryEdit.bind("<FocusOut>", self.enterPressed)
            entryEdit.selID = selectID
            entryEdit.bind('<Return>', self.enterPressed)
        #coluna de Verificado
        if column == "#3":
            current = self.table.item(selectID).get("values")
            if current[2] == "Não":
                current[2] = "Sim"
                self.table.item(selectID, values=current)
            else:
                current[2] = "Não"
                self.table.item(selectID, values=current)
    def validateDate(self,length):
        #validar quantidade de caracteres no campo de dia 
        if len(length) <= 2:
            return True
        else:
            self.dayentry.tk_focusNext().focus()
            return False
    def validatemonth(self,length):
        #validar quantidade de caracteres no campo de mes
        if len(length) <= 2:
            return True
        else:
            self.monthentry.tk_focusNext().focus()
            return False
    def validateYear(self,length):
        #validar quantidade de caracteres no campo de ano
        if len(length) <= 4:
            return True
        else:
            self.yearentry.tk_focusNext().focus()
            return False
    def enterPressed(self, event):
        #confirmando o texto ao apertar enter
        newText = event.widget.get()
        current = self.table.item(event.widget.selID).get("values")
        current[3] = newText
        self.table.item(event.widget.selID, values=current)
        event.widget.destroy()       
    def saveXL(self):
        #salvando arquivo
        popup = messagebox.askyesno(title="Atenção!", message="Salvar e enviar o arquivo?")
        if popup:
            wb = Workbook()
            ws = wb.active
            elb = self.elbentry.get() 
            day = self.dayentry.get() 
            month = self.monthentry.get() 
            year = self.yearentry.get() 
            iddaterow = "Data: %s / %s / %s  " % (day,month,year)
            idelbrow = "Elaborador: %s" % elb
            ws.append({'A': "Data: %s / %s / %s" % (day,month,year)})
            ws.append({'A': "Elaborador: %s" % elb})

            for rowid in self.table.get_children():
                row = self.table.item(rowid)['values']
                ws.append(row)
            ws.column_dimensions["A"].width=50
            ws.column_dimensions["B"].width=20
            filedt = "%s-%s-%s__%s" % (day,month,year, elb)
            filename ="%s__%s.xlsx" % (self.nmwin,filedt)
            wb.save(filename=(filename) )
            self.sendEmail(filedt,filename)
        else:
            return
    def sendEmail(self,fdt,flnm):
        #enviando arquivo ao outlook
        olApp = win32.Dispatch('Outlook.Application')
        olNS = olApp.GetNameSpace('MAPI')
        mailItem = olApp.CreateItem(0)
        mailItem.Subject = "Relatório de Operações - %s" % fdt
        mailItem.BodyFormat = 1
        mailItem.Body = "Segue em anexo"
        mailItem.To = 'Email@Email.com' 
        mailItem.Attachments.Add(os.path.join(os.getcwd(), '%s' % flnm))
        mailItem.Display()
        mailItem.Save()
    def goBack(self):
        #botao de voltar
        self.root.destroy()
        app = self.backwin()
        app
