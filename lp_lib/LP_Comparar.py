# -*- coding: cp860 -*-

dados = '''
Vers„o 2.0.6
Atualiza‡„o do programa: 27/07/2015
Monta janela Comparar
'''

import tkinter
from tkinter.messagebox import showerror
from tkinter.ttk import Combobox
from tkinter.filedialog import askopenfilename
from os import path, getcwd

try:
    from xlrd import open_workbook
except:
    showerror('Erro', 'Modulo xlrd n„o instalado')

try:
    from lp_lib.Checar_LP import checar
except:
    showerror('Erro', 'M¢dulo Checar_LP n„o instalado')

try:
    from lp_lib.func import processing
except:
    showerror('Erro', 'M¢dulo func n„o instalado')


class JanelaComp:
    def __init__(self, toplevel, relatorio_tela):

        self.toplevel = toplevel
        self.relatorioJanelaPrincipal = relatorio_tela
        self.pwd = getcwd
        # Armazena caminho dos arquivos utilizados   
        self.LPBase = 'Escolher arquivo...'
        self.Checar = 'Escolher arquivo...'

        # Tamanhos de frame
        frmlargura = 350
        frmaltura = 90

        # FRAME "Arquivo LP Base"-------------------------------------------------------------------        
        self.frm11 = tkinter.LabelFrame(toplevel, text='Arquivo LP Base', height=frmaltura, width=frmlargura)
        self.frm11.grid(row=1, column=1, padx=3, pady=3)
        self.frm11.grid_propagate(0)

        self.lbLPBase = tkinter.Label(self.frm11, text=self.LPBase, width=int((frmlargura / 8) - 10), anchor=tkinter.W)
        self.lbLPBase.grid(row=1, column=1, sticky=tkinter.W, padx=2, pady=5)

        self.cbBase = Combobox(self.frm11)
        self.cbBase.grid(row=2, column=1, sticky=tkinter.W, padx=3)

        self.btEscolheArquivoBase = tkinter.Button(self.frm11, text='Selecionar', bg='#E0E0E0', width=9,
                                           command=self.btEscolheArquivoBaseClick)
        self.btEscolheArquivoBase.grid(row=1, rowspan=2, column=2)

        # FRAME "Arquivo LP_Config"-------------------------------------------------------------------        
        self.frm21 = tkinter.LabelFrame(toplevel, text='Arquivo LP a ser checado', height=frmaltura, width=frmlargura)
        self.frm21.grid(row=2, column=1, padx=3, pady=3)
        self.frm21.grid_propagate(0)

        self.lbChecar = tkinter.Label(self.frm21, text=self.Checar, width=int((frmlargura / 8) - 10), anchor=tkinter.W)
        self.lbChecar.grid(row=1, column=1, sticky=tkinter.W, padx=2, pady=5)

        self.cbChecar = Combobox(self.frm21)
        self.cbChecar.grid(row=2, column=1, sticky=tkinter.W, padx=3)

        self.btEscolheArquivoChecar = tkinter.Button(self.frm21, text='Selecionar', bg='#E0E0E0', width=9,
                                             command=self.btEscolheArquivoChecarClick)
        self.btEscolheArquivoChecar.grid(row=1, rowspan=2, column=2, sticky=tkinter.E)

        # FRAME Botoes ------------------------------------------------------------------------------------
        self.frm31 = tkinter.Frame(toplevel, height=frmaltura, width=frmlargura)
        self.frm31.grid(row=3, column=1, padx=3, pady=3)
        self.frm31.grid_propagate(0)

        self.btChecar = tkinter.Button(self.frm31, text='Checar', bg='#E0E0E0', width=int(frmlargura / 20), height=2,
                               state=tkinter.DISABLED, command=self.btChecarClick)
        self.btChecar.pack(pady=6)

    def btEscolheArquivoBaseClick(self):
        temp = askopenfilename(filetypes=[('Arquivo do Excel', 'xls'), ('Arquivo do Excel', 'xlsx')],
                               initialdir=self.pwd)
        if temp:
            self.pwd = path.dirname(temp)
            self.LPBase = temp
            self.lbLPBase['text'] = path.basename(temp)

            try:
                book = open_workbook(temp)  # Abrir arquivo base
            except:
                aviso = 'Arquivo \"' + temp + '\" n„o encontrado'
                showerror('Erro', aviso)
            array_cbBase = []
            for plan_index in range(book.nsheets):
                sheet = book.sheet_by_index(plan_index)  # Abrir planilhas
                array_cbBase.append(sheet.name)
            self.cbBase['values'] = tuple(array_cbBase)
            self.cbBase.current(0)
            if self.lbChecar['text'] != 'Escolher arquivo...':
                self.btChecar.config(state=tkinter.NORMAL)
        self.btEscolheArquivoChecar.focus_set()

    def btEscolheArquivoChecarClick(self):
        temp = askopenfilename(filetypes=[('Arquivo do Excel', 'xls'), ('Arquivo do Excel', 'xlsx')],
                               initialdir=self.pwd)
        if (temp != ''):
            self.pwd = path.dirname(temp)
            self.Checar = temp
            self.lbChecar['text'] = path.basename(temp)

            try:
                book = open_workbook(temp)  # Abrir arquivo base
            except:
                aviso = 'Arquivo \"' + temp + '\" n„o encontrado'
                showerror('Erro', aviso)
            array_cbChecar = []
            for plan_index in range(book.nsheets):
                sheet = book.sheet_by_index(plan_index)  # Abrir planilhas
                array_cbChecar.append(sheet.name)
            self.cbChecar['values'] = tuple(array_cbChecar)
            self.cbChecar.current(0)
            if self.lbLPBase['text'] != 'Escolher arquivo...':
                self.btChecar.config(state=tkinter.NORMAL)
        self.btChecar.focus_set()

    def btChecarClick(self):

        book = open_workbook(self.LPBase)  # Abrir arquivo de LP Base
        sheet = book.sheet_by_name(self.cbBase.get())  # Abrir planilha
        array_base = []

        try:
            for index_linha in range(6, sheet.nrows):  # Ler c‚lulas da linha 7 ao final
                if sheet.cell(index_linha, 9).value != '' and sheet.cell(index_linha, 9).value != 'CGS' and sheet.cell(
                        index_linha, 9).value != 'PDS' and sheet.cell(index_linha, 9).value != 'PAS':
                    # 0 - ID SAGE
                    array_base.append([sheet.cell(index_linha, 9).value,
                                       # 1 - OCR
                                       sheet.cell(index_linha, 10).value,
                                       # 2 - DESCRI€ŽO
                                       sheet.cell(index_linha, 11).value.strip(),
                                       # 3 - TIPO
                                       sheet.cell(index_linha, 12).value,
                                       # 4 - COMANDO
                                       sheet.cell(index_linha, 13).value,
                                       # 5 - MEDI€ŽO
                                       sheet.cell(index_linha, 14).value,
                                       # 6 - TELA
                                       sheet.cell(index_linha, 15).value,
                                       # 7 - LISTA DE ALARMES
                                       sheet.cell(index_linha, 16).value,
                                       # 8 - SOE
                                       sheet.cell(index_linha, 17).value])
                    # 9 - ENDERE€O N3
                    # sheet.cell(index_linha,34).value])
                    # array_base.append([sheet.cell(index_linha,34).value])
        except:
            showerror('Erro', 'O programa n„o reconhece o arquivo base como v lida')

            # checar(LP_Editado=self.Checar,planilha=self.cbChecar.get(), relatorio=self.relatorioJanelaPrincipal, array_base=array_base)
        # self.toplevel.destroy() #Fechar Janela
        try:
            processing(checar, {'LP_Editado': self.Checar, 'planilha': self.cbChecar.get(),
                                'relatorio': self.relatorioJanelaPrincipal, 'array_base': array_base})
            self.toplevel.destroy()  # Fechar Janela
        except:
            showerror('Erro', 'Erro inesperado ao tentar checar lista de pontos.')
