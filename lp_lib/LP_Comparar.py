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
from traceback import print_exc
from sys import stdout

try:
    from openpyxl import load_workbook,cell
except:
    showerror('Erro', 'M¢dulo openpyxl n„o instalado')

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
try:
    from lp_lib.func import linhaInicialETitulos
except:
    showerror('Erro', 'Arquivo "func.pyc" deve estar no diret¢rio "lp_lib"')


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
                book = load_workbook(temp)  # Abrir arquivo base
            except:
                aviso = 'Arquivo \"' + temp + '\" n„o encontrado ou n„o suportado, utilizar planilha formato .xlsx'
                print_exc(file= stdout)
                showerror('Erro', aviso)
            array_cbBase = []
            for nome_aba in book.sheetnames:
                array_cbBase.append(nome_aba) #Abrir planilhas
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
                book = load_workbook(temp)  # Abrir arquivo base
            except:
                aviso = 'Arquivo \"' + temp + '\" n„o encontrado ou n„o suportado, utilizar planilha formato .xlsx'
                showerror('Erro', aviso)
            array_cbChecar = []
            for nome_aba in book.sheetnames:
                array_cbChecar.append(nome_aba)
            self.cbChecar['values'] = tuple(array_cbChecar)
            self.cbChecar.current(0)
            if self.lbLPBase['text'] != 'Escolher arquivo...':
                self.btChecar.config(state=tkinter.NORMAL)
        self.btChecar.focus_set()

    def btChecarClick(self):

        book = load_workbook(self.LPBase, data_only=True)  # Abrir arquivo de LP Base
        sheet = book[self.cbBase.get()]  # Abrir planilha
        array_base = []

        try:
            # Lˆ planilha e recebe a linha onde come‡a a LP (aqui usando linha inicial e n„o o dicion rio de t¡tulos)
            li, titulo_dic = linhaInicialETitulos(self.LPBase, self.cbBase.get())
            if li < 0:  # Se for um n£mero negativo ent„o n„o foi encontrado "ID (SAGE)" na lista
                raise NameError('Arquivo especificado n„o possui coluna com t¡tulo "ID (SAGE)".')
            for index_linha in range(li, sheet.max_row+1):  # Ler c‚lulas da linha 7 ao final
                if sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['ID (SAGE)']).value != '' and \
                        sheet.cell(row=index_linha,
                                   column=titulo_dic['CHESF - N‹VEL 2']['ID (SAGE)']).value != 'CGS' and \
                        sheet.cell(row=index_linha,
                                   column=titulo_dic['CHESF - N‹VEL 2']['ID (SAGE)']).value != 'PDS' and \
                        sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['ID (SAGE)']).value != 'PAS':
                    try:  # Caso a descri‡„o do campo 6 seja "TELA"
                        # 0 - ID SAGE
                        array_coletado = [
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['ID (SAGE)']).value),
                            # N2
                            # 1 - OCR
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['OCR (SAGE)']).value),
                            # 1 - DESCRI€ŽO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - N‹VEL 2']['DESCRI€ŽO']).value).strip(),
                            # 2 - TIPO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['TIPO']).value),
                            # 3 - COMANDO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['COMANDO']).value),
                            # 4 - MEDI€ŽO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['MEDI€ŽO']).value),
                            # 5 - TELA
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['TELA']).value),
                            # 6 - LISTA DE ALARMES
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - N‹VEL 2']['LISTA DE ALARMES']).value),
                            # 7 - SOE
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['SOE']).value),
                            # TELEASSIST‰NCIA N3
                            # 8 - OCR
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['OCR (SAGE)']).value),
                            # 9 - COMANDO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['COMANDO']).value),
                            # 10 - MEDI€ŽO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['MEDI€ŽO']).value),
                            # 11 - LISTA DE ALARMES
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['LISTA DE ALARME']).value),
                            # 12 - SOE
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['SOE']).value),
                            # 13 - OBSERVA€ŽO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['OBSERVA€ŽO']).value),
                            # 15 - AGRUPAMENTO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['AGRUPAMENTO']).value),
                            # N3
                            # 16 - OCR (SAGE)
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['OCR (SAGE)']).value),
                            # 17 - COMANDO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['COMANDO']).value),
                            # 18 - MEDI€ŽO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['MEDI€ŽO']).value),
                            # 19 - LISTA DE ALARMES
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - N‹VEL 3']['LISTA DE ALARME']).value),
                            # 20 - SOE
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['SOE']).value),
                            # 21 - OBSERVA€ŽO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['OBSERVA€ŽO']).value),
                            # 22 - AGRUPAMETO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['AGRUPAMENTO']).value),
                            # ONS
                            # 23 - ITEM
                            str(sheet.cell(row=index_linha, column=titulo_dic['ONS']['ITEM']).value),
                            # 24 - DESCRI€ŽO
                            str(sheet.cell(row=index_linha, column=titulo_dic['ONS']['DESCRI€ŽO']).value),
                            # LIMITES OPERACIONAIS
                            # 25 - LIU
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LIU']).value),
                            # 26 - LIE
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LIE']).value),
                            # 27 - LIA
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LIA']).value),
                            # 28 - LSA
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LSA']).value),
                            # 29 - LSE
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LSE']).value),
                            # 30 - LSU
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LSU']).value),
                            # 31 - BNDMO
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['BNDMO']).value),
                            # 32 - OBSERVA€™ES
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['LIMITES OPERACIONAIS']['OBSERVA€™ES']).value),
                            # 33 - ENDERE€O N3 Teleassistˆncia
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['ENDERE€O']).value),
                            # 34 - ENDERE€O N3
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['ENDERE€O']).value)]
                    except:  # Caso a descri‡„o do campo 6 seja "ANUNCIADOR"
                        # 0 - ID SAGE
                        array_coletado = [
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['ID (SAGE)']).value),
                            # N2
                            # 1 - OCR
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['OCR (SAGE)']).value),
                            # 2 - DESCRI€ŽO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - N‹VEL 2']['DESCRI€ŽO']).value).strip(),
                            # 3 - TIPO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['TIPO']).value),
                            # 4 - COMANDO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['COMANDO']).value),
                            # 5 - MEDI€ŽO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['MEDI€ŽO']).value),
                            # 6 - TELA
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['ANUNCIADOR']).value),
                            # 7 - LISTA DE ALARMES
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - N‹VEL 2']['LISTA DE ALARMES']).value),
                            # 8 - SOE
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 2']['SOE']).value),
                            # TELEASSIST‰NCIA N3
                            # 9 - OCR
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['OCR (SAGE)']).value),
                            # 10 - COMANDO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['COMANDO']).value),
                            # 11 - MEDI€ŽO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['MEDI€ŽO']).value),
                            # 12 - LISTA DE ALARMES
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['LISTA DE ALARME']).value),
                            # 13 - SOE
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['SOE']).value),
                            # 14 - OBSERVA€ŽO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['OBSERVA€ŽO']).value),
                            # 15 - AGRUPAMENTO
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['AGRUPAMENTO']).value),
                            # N3
                            # 16 - OCR (SAGE)
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['OCR (SAGE)']).value),
                            # 17 - COMANDO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['COMANDO']).value),
                            # 18 - MEDI€ŽO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['MEDI€ŽO']).value),
                            # 19 - LISTA DE ALARMES
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - N‹VEL 3']['LISTA DE ALARME']).value),
                            # 20 - SOE
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['SOE']).value),
                            # 21 - OBSERVA€ŽO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['OBSERVA€ŽO']).value),
                            # 22 - AGRUPAMETO
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['AGRUPAMENTO']).value),
                            # ONS
                            # 23 - ITEM
                            str(sheet.cell(row=index_linha, column=titulo_dic['ONS']['ITEM']).value),
                            # 24 - DESCRI€ŽO
                            str(sheet.cell(row=index_linha, column=titulo_dic['ONS']['DESCRI€ŽO']).value),
                            # LIMITES OPERACIONAIS
                            # 25 - LIU
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LIU']).value),
                            # 26 - LIE
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LIE']).value),
                            # 27 - LIA
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LIA']).value),
                            # 28 - LSA
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LSA']).value),
                            # 29 - LSE
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LSE']).value),
                            # 30 - LSU
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['LSU']).value),
                            # 31 - BNDMO
                            str(sheet.cell(row=index_linha, column=titulo_dic['LIMITES OPERACIONAIS']['BNDMO']).value),
                            # 32 - OBSERVA€™ES
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['LIMITES OPERACIONAIS']['OBSERVA€™ES']).value),
                            # 33 - ENDERE€O N3 Teleassistˆncia
                            str(sheet.cell(row=index_linha,
                                           column=titulo_dic['CHESF - TELEASSIST‰NCIA N3']['ENDERE€O']).value),
                            # 34 - ENDERE€O N3
                            str(sheet.cell(row=index_linha, column=titulo_dic['CHESF - N‹VEL 3']['ENDERE€O']).value)]
                    for i in range(0, len(array_coletado)):
                        if array_coletado[i] == 'None':
                            array_coletado[i] = ''
                    array_base.append(array_coletado)
        except:
            print_exc(file=stdout)
            showerror('Erro', 'O programa n„o reconhece o arquivo base como v lida')

            # checar(LP_Editado=self.Checar,planilha=self.cbChecar.get(), relatorio=self.relatorioJanelaPrincipal, array_base=array_base)
        # self.toplevel.destroy() #Fechar Janela
        try:
            processing(checar, {'LP_Editado': self.Checar, 'planilha': self.cbChecar.get(),
                                'relatorio': self.relatorioJanelaPrincipal, 'array_base': array_base})
            self.toplevel.destroy()  # Fechar Janela
        except:
            showerror('Erro', 'Erro inesperado ao tentar checar lista de pontos.')
