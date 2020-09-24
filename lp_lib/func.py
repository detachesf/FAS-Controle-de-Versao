# -*- coding: cp860 -*-

from tkinter.messagebox import showerror
from tkinter import Toplevel, ttk, E, W, CENTER, LEFT, SOLID, Label
import threading
from time import sleep

try:
    from xlrd import open_workbook
except:
    showerror('Erro', 'M¢dulo xlrd n„o instalado')

dados= '''            
Vers„o do programa: 2.0.10
Atualiza‡„o do programa: 18/07/2016
Fun‡”es adicionais: painelLT69, linhaInicialETitulos
'''


def painelLT69(ListaPontos):
    array_ID = [col[0] for col in ListaPontos[0]] # Separar ID SAGE da lista de pontos (LP em 0 e contadores k nas outras posi‡”es)
    array_checar = [tag[-11:] for tag in array_ID] # Separar as £ltimas 11 posi‡”es (Painel e c¢digo de ponto)
    pos11dupl = [dupl11 for dupl11 in set(array_checar) if array_checar.count(dupl11)>1 and dupl11[:3]=='2UA' ] # Separar as £ltimas 11 posi‡”es duplicadas come‡adas com 2UA (referente a painel de 69kV)
    array_ID_dupl = [dupl[0] for dupl in ListaPontos[0] if dupl[0][-11:] in pos11dupl] #Separa array com IDs duplicados
    
    tratar_par = []
    for i11 in pos11dupl: #Passar array de 11 duplicados
        par = []
        for idupl in array_ID_dupl: #Passar array de IDs a serem tratados
            if idupl[-11:] == i11:
                par.append(idupl) #Separar IDs duplicados em pares para tratamento
        if len(par)>0: tratar_par.append(par) #Gravar par em array para tratamento
    for tratar in tratar_par:
        novoID = tratar[0][0:8]+'/'+tratar[1][6:8]+tratar[0][8:]
        pos0 = array_ID.index(tratar[0]) #Pegar posi‡„o do primero ID a ser Tratado
        ListaPontos[0].pop(pos0) #Excluir ponto referente ao primeiro ID tratar[0]
        array_ID.pop(pos0) #Excluir ponto referente ao primeiro ID tratar[0] para os indices da ListaPontos e array_ID continuem coerentes
        
        pos1 = array_ID.index(tratar[1]) #Pegar posi‡„o do segundo ID a ser Tratado
        ListaPontos[0][pos1][0] = novoID #Substituir ID do ponto referente ao segundo ID tratar[1]
        
        ListaPontos[1][0] -= 1 #Modificar k_lt
    
    return ListaPontos

def linhaInicialETitulos(arquivo, nomeAba):
    """
    Rotina para encontrar a primeira linha v lida de uma lista de pontos e t¡tulos do cabe‡alho da Lista de Pontos
    @param arquivo: Arquivo Excel com a Lista de Pontos
    @param nomeAba: Nome da aba que est  a Lista de Pontos    
    @return: [linhaInicial, {dicion rio de t¡tulos}] - Retorna array com a linha inicial (primeira linha ‚ a  0 - zero), 
    na primeira posi‡„o e dicion rio com t¡tulos do cabe‡alho na segunda posi‡„o. Caso n„o seja encontrada referˆncia ao 
    'D (SAGE)' o valor de retorno de linhaInicial ser  -1 (menos um).  
    """   
    
    arq_conf = open_workbook(arquivo)  
    sheet = arq_conf.sheet_by_name(nomeAba) 
    
    titulos = {}
    for li in range(1, 10):                                         #Varrer as linhas de 2 a 10
        for i in range(sheet.ncols):                                #Varrer as colunas da linha
            texto_coluna = sheet.cell_value(li,i).upper().strip()   #Pegar texto da c‚lula
            if texto_coluna == '':                                  #Gravar £ltima posi‡„o com valor vazio
                titulos[texto_coluna] = i
            elif texto_coluna not in titulos:                       #Iserir chave se n„o existir no dicion rio
                titulos[texto_coluna] = i
        if 'ID (SAGE)' in titulos: break                     #Se foi passado pela linha com chave "ID (SAGE)" parar de varrer linhas 
    
    li += 1                                                         #Seleciona linha ap¢s o t¡tulo
    if 'ID (SAGE)' in titulos:                               #Verifica se foi encontrado chave "ID (SAGE)"
        while True:
            if sheet.cell_value(li,titulos['ID (SAGE)']):              #Verifica se a c‚lula est  preenchida com algum valor
                break                                                   #Parar de procurar linha preenchida
            else:
                li += 1                                                 #Selecionar linha seguinte
    else:
        li = -1            
    return [li,titulos]

def processing(function, args):

    def check(): # Checar momento de fechar janela
        if not thread_f.isAlive(): window.destroy()
        frame.after(500, check)

    window = Toplevel()
    window.resizable(0, 0)
    window.overrideredirect(1)
    window.attributes('-alpha', 0.7)

    w = 150
    h = 60
    ws = window.winfo_screenwidth()
    hs = window.winfo_screenheight()
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)
    window.geometry('{}x{}+{}+{}'.format(w, h, int(x), int(y)))

    frame = ttk.Frame(window, height=h, width=w)
    label = ttk.Label(frame, text = 'Processando', anchor=CENTER)
    label.grid(row=1, column=1, pady=3, sticky=W+E)
    progress = ttk.Progressbar(frame, orient='horizontal', mode='indeterminate', length=w-10)
    progress.grid(row=2, column=1, pady=3, sticky=W+E)
    progress.start(12)
    frame.pack()

    thread_f = threading.Thread(target=function, kwargs=args)
    thread_f.start()

    frame.after(500, check)
    frame.mainloop()

class ToolTip(object):
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0
    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, _cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 27
        y = y + cy + self.widget.winfo_rooty() +27
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))

        label = Label(tw, text=self.text, justify=LEFT,
                         background="#ffffe0", relief=SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)
    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def createToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        sleep(3)
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

