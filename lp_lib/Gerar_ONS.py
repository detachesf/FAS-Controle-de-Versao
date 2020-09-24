# -*- coding: cp860 -*-
dados = '''
AtualizaáÑo do programa: 03/09/2014
Funcionalidades relacionadas a geraáÑo da planilha de pontos ONS a partir de uma lista de pontos padrÑo.
'''

from lp_lib.func import linhaInicialETitulos
from os import path, startfile, getcwd 
from tkinter.messagebox import showerror, askyesno
from tkinter import Button, Label, LabelFrame, N, S, E, W, END, NORMAL, DISABLED
from tkinter.ttk import Combobox
from tkinter.filedialog import askopenfilename

try:
    from xlrd import open_workbook
except:
    showerror('Erro', u'M¢dulo xlrd nÑo instalado')
    
posComentarioPonto  = 1     #: PosiáÑo do coment†rio
posIDponto          = 2     #: PosiáÑo do ID do Ponto
posEnderecoPonto    = 4    #: PosiáÑo do endereáo do Ponto
posONSItem          = 5    #: PosiáÑo do item ONS ao qual ele est† associado
posTipoPonto        = 3    #: PosiáÑo do Tipo do Ponto

from lp_lib.gerarPlanilhaONS import gerarPlanilhaONS


class JanelaGerarONS:
    """
    Classe que implementa a janela que Ç exibida ao clicar no menu 'Ferramentas -> Gerar Planilha ONS...' na janela princial
    """
    def __init__(self,toplevel, Relatorio):
        """
        Construtor da Classe.
        @param Relatorio: Ç o ListBox da janela principal que ir† ser utilizado para exbir informaáîes sobre a geraáÑo da planilha ONS.
        @type Relatorio: Tkinter.ListBox
        @param toplevel: Ç a I{window} que exibir† a janela gerada por esta classe.
        @type toplevel: Tkinter.Toplevel  
        """
        self.escreveRelatorio = Relatorio
        self.toplevel = toplevel
        frmLargura = 330
        frmAltura = 150
        
        #gerarONS_tela = Tk()
        #gerarONS_tela.title('Gerar Planilha ONS')

        #Frames principais
        #self.frame = Frame(gerarONS_tela)
        #self.frame.pack()
        self.frameopcoes = LabelFrame(toplevel, text=u'Arquivo de ParametrizaáÑo', height = frmAltura, width = frmLargura)
        self.frameopcoes.grid(row=0, column=0, padx=3, pady=3)
        self.frameopcoes.grid_propagate(0) 
        
        self.nomeCaminhoArquivoGerarONS = Label(self.frameopcoes, text='Escolha arquivo...')
        self.nomeCaminhoArquivoGerarONS.grid(row=0, column=0)#, columnspan=2)
        self.nomeCaminhoArquivoGerarONS.grid_propagate(0)        
        
        from test.test_typechecks import Integer
        self.botaoEscolheArquivoONS = Button(self.frameopcoes, text='Selecionar Arquivo',bg='#E0E0E0', width= int(frmLargura/8), command=self.btEscolheArquivoONSClick)
        self.botaoEscolheArquivoONS.grid(row=1, column=0,sticky=N+S,padx=3, pady=3)
        self.botaoEscolheArquivoONS.grid_propagate(0)
        
        self.comboONS = Combobox(self.frameopcoes)
        self.comboONS.grid(row=2, column=0, sticky=N+E+S+W, pady=3, padx=3)
        self.comboONS.grid_propagate(0)
        
        from ctypes.wintypes import INT
        self.botaogerarPlanilhaONS = Button(self.frameopcoes, text='Gerar ONS',bg='#E0E0E0', width=int(frmLargura/8), state = DISABLED,command=self.btGerarPlanilhaONS)
        self.botaogerarPlanilhaONS.grid(row=3, column=0,sticky=N+S,padx=3, pady=3)
        self.comboONS.grid_propagate(0)


    def btEscolheArquivoONSClick(self):
        temp = askopenfilename(filetypes=[('Arquivo do Excel','xls'),('Arquivo do Excel 2007','xlsx')])
        if temp:
            self.CaminhoArquivoGerarONS = temp
            self.nomeCaminhoArquivoGerarONS['text'] = path.basename(temp)
            try:
                book = open_workbook(temp) #Abrir arquivo de a ser checado
            except:
                aviso = 'Arquivo \"'+temp+u'\" nÑo encontrado'
                showerror('Erro',aviso)
            array_combo = []
            for plan_index in range(book.nsheets):
                sheet = book.sheet_by_index(plan_index) #Abrir planilhas
                array_combo.append(sheet.name)
                #self.Lb.insert(END,str(plan_index)+' '+sheet.name)
            self.comboONS['values'] = tuple(array_combo)
            self.comboONS.current(0)
            self.botaogerarPlanilhaONS.config(state = NORMAL)
            self.botaogerarPlanilhaONS.focus_set()
              
                
    def btGerarPlanilhaONS(self):
        gerarONS(self.CaminhoArquivoGerarONS,self.comboONS.get(),self.escreveRelatorio)
        self.toplevel.destroy() #Fechar Janela


def geraListaDeEventos(listaPontos):
    """
    FunáÑo que recebe uma lista de pontos e verifica quais os eventos contidos nela.
    @param listaPontos: Lista de pontos a ser avaliada.
    @return: Retorna dois arrays. O Primeiro com a lista de eventos identificados na lista de pontos e o segundo com o c†odigo da instalaáÑo tambÇm identificada na lista de pontos.
    """    
    listaEventos = []
    for ponto in listaPontos:
        if (ponto[posIDponto][5].isdigit()) and (ponto[posIDponto][6].isalnum()) and ((ponto[posIDponto][7].isdigit())):
            if not ((('0'+ponto[posIDponto][5:8]) in listaEventos) or (('1'+ponto[posIDponto][5:8]) in listaEventos)):
                if (ponto[posIDponto][4] != '3') and (not (ponto[posIDponto][6] =='T' and ponto[posIDponto][4] =='1')):
                    if ponto[posIDponto][6] =='E': listaEventos.append('0'+ponto[posIDponto][5:8])
                    else: listaEventos.append(ponto[posIDponto][4:8])
    CodSubestacao = listaPontos[0][posIDponto][:3]        
    
    return listaEventos, CodSubestacao


def geraListaEventosOrganizada(listaEventos):
    """
    FunáÑo que gera lista com os eventos que estÑo contidos numa lista de pontos, de maneira organizada e por tipo.
    @param listaEventos: ista de pontos que ser† avaliada.
    @return: Retorna lista no formato [TT,LT,BR,BT,RE,BC,TR,CS,CE] onde cada item da lista Ç uma segunda lista com os evento associados.  
    """
    TT = []   # Trafo Terra
    LT = []   # Linha TransmissÑo
    BR = []   # Barra
    BT = []   # Bay Transferància
    RE = []   # Reator
    BC = []   # Banco Capacitor
    TR = []   # Transformador
    CS = []   # Compensador S°ncrono
    CE = []   # Compensador Est†tico    
    
    for evento in listaEventos:
        if (evento[2] in ('A')):
            TT.append(evento)
        elif (evento[2] in ('C','F','L','M','N','P','S','V','O','Z')):
            LT.append(evento)
        elif (evento[2] in ('B')):
            BR.append(evento) 
        elif (evento[2] in ('D')):
            BT.append(evento)               
        elif (evento[2] in ('E')):
            RE.append(evento)
        elif (evento[2] in ('H')):
            BC.append(evento)  
        elif (evento[2] in ('T')):
            TR.append(evento)    
        elif (evento[2] in ('K')):
            CS.append(evento)  
        elif (evento[2] in ('Q')):
            CE.append(evento)             
    return [TT,LT,BR,BT,RE,BC,TR,CS,CE]


def geraPontoONSParaPlanilha(ponto):
    """
    FunáÑo que recebe um ponto e o formata removendo apenas os dados de interesse para gravar na Planilha da ONS
    @param ponto: ponto avaliado
    @return: Retorna array de 7 posicoes contendo os dados selecionados
    """    

    ptONS = []
    # Colunas Atendimento (SIM, NéO, N. APLIC.)
    ptONS.append('x')
    ptONS.append('')
    ptONS.append('')
    # Descricao do ponto
    ptONS.append(ponto[posComentarioPonto])
    # Endereco do ponto
    ptONS.append(ponto[posEnderecoPonto])
    # Tipo do ponto
    tipoPonto = ponto[posTipoPonto].strip().replace(' ','')
    if   (tipoPonto == '')   : ptONS.append('')
    elif (tipoPonto == 'AN') : ptONS.append(u'MediáÑo')
    elif (tipoPonto == 'DD') : ptONS.append(u'Duplo')
    elif (tipoPonto == 'DS') : ptONS.append(u'Simples')
    else : ptONS.append(tipoPonto) 
       
    # Observacao do ponto
    ptONS.append(ponto[posIDponto])
    return ptONS    


def geraListaLT(codBay,listaPontosLT):
    '''
    Gera lista preparada dos pontos a serem preenchidos na palanilha ONS para Linhas de TransmissÑo. Essa lista contÇm apenas os campos que sÑo necess†rios para a planilha no formato padrÑo CHESF. (***colocar aqui o formato da lista**)
    @param codBay: C¢digo do bay avaliado
    @param listaPontosLT: Lista de pontos das Linhas de TransmissÑo.
    @return: Retorna lista com dois elementos [listaLT,pontosFalhosLT] onde: "listaLT" Ç uma lista com os campos para a planilha ONS e "pontosFalhosLT" Ç uma lista com todos os pontos que apresentaram algum tipo de problema.
    '''
    # itens do PR 2.6
    itensONS_LT = ['7311c4','7311c5','7311c6',
                   '7312a','7312d',
                   '8214a1','8214a2','8214a3','8214a5','8214b','8214c1','8214c2','8214c3','8214c4','8214c5',
                   '8218a1','8218a2','8218a3','8218b1','8218b2','8218b3','8218b4']

    # guarda pontos que nÑo se encaixam na planilha correspondente
    pontosFalhosLT = []   
    listaPontosONS_LT = []  

    # escreve primeira posiáÑo com o c¢digo do Bay  
    listaPontosONS_LT.append(codBay)
   
    # varre a lista 
    for ponto in listaPontosLT:  
        itensDoPonto = ponto[posONSItem].strip().replace('.','').replace('(','').replace(')','').replace(']','')
        itensDoPonto = itensDoPonto.split('[')   
         
        for itemONS in itensDoPonto:                                    # loop para verificar todos os itens aos quais o ponto esta associado    
            if (itemONS in itensONS_LT):                                # verifica se o item est† na lista de itens esperados para LTs
                # verifica se esse item j† existe na lsita de pontos varridos para o bay
                existe = False
                for subLista in listaPontosONS_LT:                      # verifica se j† foi encontrado algum ponto do mesmo item entre os que j† foram verificados...
                    # verifica se j† foi encontrado algum ponto do mesmo item ...
                    if (subLista[0] == itemONS):
                            # ... se foi encontrado, verifica se um item que admite apenas um ponto associado (agrupamentos)
                            if (itemONS in ['8214b']): 
                                # se s¢ admitir um ponto asociado, grava como erro
                                pontosFalhosLT.append([ponto[posIDponto],
                                                       ponto[posComentarioPonto],
                                                       itemONS,
                                                       'LINHA TRANSMISSéO: Duplicidade de ponto associado ao mesmo item ONS'])
                            else: 
                                # adiciona mais um ponto ao item ONS associado
                                subLista[1].append(geraPontoONSParaPlanilha(ponto))
                                existe = True
                            break
                # se nÑo existe entÑo cria
                if (existe==False):
                    listaPontosONS_LT.append([itemONS,[geraPontoONSParaPlanilha(ponto)]])
            else: # item nÑo est† na lista de itens esperados para um LT
                if(itemONS in ['8214b1','8214b2','8214b3','8214b4','8214b5','8214b6','8214b7','8214b8','8214b9','8214b10','8214b11','8214b12','8214b13','8214b14']):
                    pontosFalhosLT.append([ponto[posIDponto],
                                           ponto[posComentarioPonto],
                                           itemONS,
                                           'LINHA TRANSMISSéO: O item desse ponto indica que ele deve fazer parte de um agrupamento (GRUPO B - agrupamento de eventos)'])                     
                else:    
                    pontosFalhosLT.append([ponto[posIDponto],
                                           ponto[posComentarioPonto],
                                           itemONS,
                                           'LINHA TRANSMISSéO: Item ONS associado ao ponto nÑo esperado para este tipo de evento.'])      
    
    return [listaPontosONS_LT,pontosFalhosLT]


def geraListaBT(codBay,listaPontosBT):
    '''
    Gera lista preparada dos pontos a serem preenchidos na palanilha ONS para um Bay de Transferància. Essa lista contÇm apenas os campos que sÑo necess†rios para a planilha no formato padrÑo CHESF. (***colocar aqui o formato da lista**)
    @param codBay: C¢digo do bay avaliado
    @param listaPontosBT: Lista de pontos dos Bay de Transferància.
    @return: Retorna lista com dois elementos [listaBT,pontosFalhosBT] onde: "listaBT" Ç uma lista com os campos para a planilha ONS e "pontosFalhosBT" Ç uma lista com todos os pontos que apresentaram algum tipo de problema.
    '''
    # itens do PR 2.6
    itensONS_BT = ['7312a','7312d',
                   '8218a1','8218a2','8218a3','8218b1','8218b2','8218b3','8218b4']

    # guarda pontos que nÑo se encaixam na planilha correspondente
    pontosFalhosBT = []   
    listaPontosONS_BT = []  

    # escreve primeira posiáÑo com o c¢digo do Bay  
    listaPontosONS_BT.append(codBay)
   
    # varre a lista 
    for ponto in listaPontosBT:  
        itensDoPonto = ponto[posONSItem].strip().replace('.','').replace('(','').replace(')','').replace(']','')
        itensDoPonto = itensDoPonto.split('[')   
         
        for itemONS in itensDoPonto:                                    # loop para verificar todos os itens aos quais o ponto esta associado    
            if (itemONS in itensONS_BT):                                # verifica se o item est† na lista de itens esperados para LTs
                # verifica se esse item j† existe na lsita de pontos varridos para o bay
                existe = False
                for subLista in listaPontosONS_BT:                      # verifica se j† foi encontrado algum ponto do mesmo item entre os que j† foram verificados...
                    # verifica se j† foi encontrado algum ponto do mesmo item ...
                    if (subLista[0] == itemONS):
                        # adiciona mais um ponto ao item ONS associado
                        subLista[1].append(geraPontoONSParaPlanilha(ponto))
                        existe = True
                        break
                # se nÑo existe entÑo cria
                if (existe==False):
                    listaPontosONS_BT.append([itemONS,[geraPontoONSParaPlanilha(ponto)]])
            else: # item nÑo est† na lista de itens esperados para um BT
                pontosFalhosBT.append([ponto[posIDponto],
                                       ponto[posComentarioPonto],
                                       itemONS,
                                      'BAY TRANSFERâNCIA / DISJUNTOR CENTRAL: Item ONS associado ao ponto nÑo esperado para este tipo de evento.'])     
    
       
    return [listaPontosONS_BT,pontosFalhosBT]


def geraListaTR(codBay,listaPontosTR):
    '''
    Gera lista preparada dos pontos a serem preenchidos na palanilha ONS para uma ConexÑo de Trafo. Essa lista contÇm apenas os campos que sÑo necess†rios para a planilha no formato padrÑo CHESF. (***colocar aqui o formato da lista**)
    @param codBay: C¢digo do bay avaliado
    @param listaPontosTR: Lista de pontos ConexÑo de Trafo.
    @return: Retorna lista com dois elementos [listaTR,pontosFalhosTR] onde: "listaTR" Ç uma lista com os campos para a planilha ONS e "pontosFalhosTR" Ç uma lista com todos os pontos que apresentaram algum tipo de problema.
    '''    
        
    # itens do PR 2.6
    itensONS_TR = ['7311c8','7311c13','7311c14',
                   '7312a','7312d','7312f','7312h',
                   '8211a1','8211b1','8211b2','8211b3',
                   '8218a1','8218a2','8218a3','8218b1','8218b2','8218b3','8218b4']
    # guarda pontos que nÑo se encaixam na planilha correspondente
    pontosFalhosTR = []   
    listaPontosONS_TR = []  

    # escreve primeira posiáÑo com o c¢digo do Bay  
    listaPontosONS_TR.append(codBay)
   
    # varre a lista 
    for ponto in listaPontosTR: 
        itensDoPonto = ponto[posONSItem].strip().replace('.','').replace('(','').replace(')','').replace(']','')
        itensDoPonto = itensDoPonto.split('[')    
           
        for itemONS in itensDoPonto:    
            #verifica se o item est† na lista de item para TRAFOS
            if (itemONS in itensONS_TR):
                # verifica se esse item j† existe
                existe = False
                for subLista in listaPontosONS_TR:
                    # verifica se j† foi encontrado algum ponto do mesmo item ...
                    if (subLista[0] == itemONS):
                            # ... se foi encontrado, verifica se um item que admite apenas um ponto associado (agrupamentos)
                            if (itemONS in ['8211b1','8211b1i','8211b1ii','8211b2','8211b2i','8211b2ii','8211b3','8211b3i','8211b3ii','8211b3iii','8211b3iv','8211b3v']): 
                                # se s¢ admitir um ponto asociado, grava como erro
                                pontosFalhosTR.append([ponto[posIDponto],
                                                       ponto[posComentarioPonto],
                                                       itemONS,
                                                       'TRANSFORMADOR: Duplicidade de ponto associado ao mesmo item ONS'])
                            else: 
                                # adiciona mais um ponto ao item ONS associado
                                subLista[1].append(geraPontoONSParaPlanilha(ponto))
                                existe = True
                            break
                # se nÑo existe entÑo cria
                if (existe==False):
                    listaPontosONS_TR.append([itemONS,[geraPontoONSParaPlanilha(ponto)]])
            else:
                if(itemONS in ['8211b1i','8211b1ii','8211b2i','8211b2ii','8211b3i','8211b3ii','8211b3iii','8211b3iv','8211b3v']):
                    pontosFalhosTR.append([ponto[posIDponto],
                                           ponto[posComentarioPonto],
                                           itemONS,
                                           'TRANSFORMADOR: O item desse ponto indica que ele deve fazer parte de um agrupamento (GRUPO B - agrupamento de eventos)'])                     
                else:    
                    pontosFalhosTR.append([ponto[posIDponto],
                                           ponto[posComentarioPonto],
                                           itemONS,
                                           'TRANSFORMADOR: Item ONS associado ao ponto nÑo aplic†vel ao tipo de evento'])                    
                      
    return [listaPontosONS_TR, pontosFalhosTR]
        
            
def geraListaRE(codBay,listaPontosRE):
    '''
    Gera lista preparada dos pontos a serem preenchidos na palanilha ONS para uma ConexÑo de Reator. Essa lista contÇm apenas os campos que sÑo necess†rios para a planilha no formato padrÑo CHESF. (***colocar aqui o formato da lista**)
    @param codBay: C¢digo do bay avaliado
    @param listaPontosRE: Lista de pontos ConexÑo de Reator.
    @return: Retorna lista com dois elementos [listaRE,pontosFalhosRE] onde: "listaRE" Ç uma lista com os campos para a planilha ONS e "pontosFalhosRE" Ç uma lista com todos os pontos que apresentaram algum tipo de problema.
    '''

    # itens do PR 2.6
    itensONS_RE = ['7312a','7312d','7312h',
                   '8212a1','8212b1','8212b2',
                   '8218a1','8218a2','8218a3','8218b1','8218b2','8218b3','8218b4']

    # guarda pontos que nÑo se encaixam na planilha correspondente
    pontosFalhos_RE = []   
    listaPontosONS_RE = []  

    # escreve primeira posiáÑo com o c¢digo do Bay  
    listaPontosONS_RE.append(codBay)
   
    # varre a lista 
    for ponto in listaPontosRE:  
        itensDoPonto = ponto[posONSItem].strip().replace('.','').replace('(','').replace(')','').replace(']','')
        itensDoPonto = itensDoPonto.split('[')   
         
        for itemONS in itensDoPonto:                                    # loop para verificar todos os itens aos quais o ponto esta associado    
            if (itemONS in itensONS_RE):                                # verifica se o item est† na lista de itens esperados para LTs
                # verifica se esse item j† existe na lsita de pontos varridos para o bay
                existe = False
                for subLista in listaPontosONS_RE:                      # verifica se j† foi encontrado algum ponto do mesmo item entre os que j† foram verificados...
                    # verifica se j† foi encontrado algum ponto do mesmo item ...
                    if (subLista[0] == itemONS):
                            # ... se foi encontrado, verifica se um item que admite apenas um ponto associado (agrupamentos)
                            if (itemONS in ['8212b1','8212b2']): 
                                # se s¢ admitir um ponto asociado, grava como erro
                                pontosFalhos_RE.append([ponto[posIDponto],
                                                       ponto[posComentarioPonto],
                                                       itemONS,
                                                       'REATOR: Duplicidade de ponto associado ao mesmo item ONS'])
                            else: 
                                # adiciona mais um ponto ao item ONS associado
                                subLista[1].append(geraPontoONSParaPlanilha(ponto))
                                existe = True
                            break
                # se nÑo existe entÑo cria
                if (existe==False):
                    listaPontosONS_RE.append([itemONS,[geraPontoONSParaPlanilha(ponto)]])
            else: # item nÑo est† na lista de itens esperados para um LT
                if(itemONS in ['8212b1i','8212b1ii','8212b2i','8212b2ii','8212b2iii','8212b2iv']):
                    pontosFalhos_RE.append([ponto[posIDponto],
                                           ponto[posComentarioPonto],
                                           itemONS,
                                           'REATOR: O item desse ponto indica que ele deve fazer parte de um agrupamento (GRUPO B - agrupamento de eventos)'])                     
                else:    
                    pontosFalhos_RE.append([ponto[posIDponto],
                                           ponto[posComentarioPonto],
                                           itemONS,
                                           'REATOR: Item ONS associado ao ponto nÑo esperado para este tipo de evento.'])        
                                                                                                                                    
    return [listaPontosONS_RE,pontosFalhos_RE]


def geraListaBA(codBay,listaPontosBA):
    '''
    Gera lista preparada dos pontos a serem preenchidos na palanilha ONS para um Barramento. Essa lista contÇm apenas os campos que sÑo necess†rios para a planilha no formato padrÑo CHESF. (***colocar aqui o formato da lista**)
    @param codBay: C¢digo do bay avaliado
    @param listaPontosBA: Lista de pontos dos Barramentos.
    @return: Retorna lista com dois elementos [listaBA,pontosFalhosBA] onde: "listaBA" Ç uma lista com os campos para a planilha ONS e "pontosFalhosBA" Ç uma lista com todos os pontos que apresentaram algum tipo de problema.
    '''
# itens do PR 2.6
    itensONS_BA = ['7311c1','7421a',
                   '8215a1','8215a2','8215b1',
                   '8219a1']

    # guarda pontos que nÑo se encaixam na planilha correspondente
    pontosFalhos_BA = []   
    listaPontosONS_BA = []  

    # escreve primeira posiáÑo com o c¢digo do Bay  
    listaPontosONS_BA.append(codBay)
   
    # varre a lista 
    for ponto in listaPontosBA:  
        itensDoPonto = ponto[posONSItem].strip().replace('.','').replace('(','').replace(')','').replace(']','')
        itensDoPonto = itensDoPonto.split('[')   
         
        for itemONS in itensDoPonto:                                    # loop para verificar todos os itens aos quais o ponto esta associado    
            if (itemONS in itensONS_BA):                                # verifica se o item est† na lista de itens esperados para LTs
                # verifica se esse item j† existe na lsita de pontos varridos para o bay
                existe = False
                for subLista in listaPontosONS_BA:                      # verifica se j† foi encontrado algum ponto do mesmo item entre os que j† foram verificados...
                    # verifica se j† foi encontrado algum ponto do mesmo item ...
                    if (subLista[0] == itemONS):
                           # ... se foi encontrado, verifica se um item que admite apenas um ponto associado (agrupamentos)
                            if (itemONS in ['8215b1']): 
                                # se s¢ admitir um ponto asociado, grava como erro
                                pontosFalhos_BA.append([ponto[posIDponto],
                                                       ponto[posComentarioPonto],
                                                       itemONS,
                                                       'BARRAS: Duplicidade de ponto associado ao mesmo item ONS - (GRUPO B - agrupamento de eventos)'])
                            else: 
                                # adiciona mais um ponto ao item ONS associado
                                subLista[1].append(geraPontoONSParaPlanilha(ponto))
                                existe = True
                            break
                # se nÑo existe entÑo cria
                if (existe==False):
                    listaPontosONS_BA.append([itemONS,[geraPontoONSParaPlanilha(ponto)]])
            else: # item nÑo est† na lista de itens esperados para um LT
                pontosFalhos_BA.append([ponto[posIDponto],
                                        ponto[posComentarioPonto],
                                        itemONS,
                                        'BARRAMENTO: Item ONS associado ao ponto nÑo esperado para este tipo de evento.'])            
                
    return [listaPontosONS_BA,pontosFalhos_BA]


def geraListaBC(codBay, listaPontosBC):
    '''
    Gera lista preparada dos pontos a serem preenchidos na palanilha ONS para um Banco de Capacitor. Essa lista contÇm apenas os campos que sÑo necess†rios para a planilha no formato padrÑo CHESF. (***colocar aqui o formato da lista**)
    @param codBay: C¢digo do bay avaliado
    @param listaPontosBC: Lista de pontos dos Banco de Capacitor.
    @return: Retorna lista com dois elementos [listaBC,pontosFalhosBC] onde: "listaBC" Ç uma lista com os campos para a planilha ONS e "pontosFalhosBC" Ç uma lista com todos os pontos que apresentaram algum tipo de problema.
    '''
    # itens do PR 2.6
    itensONS_BC = ['7312a','7312d',
                   '8213a1','8213a2','8213b',
                   '82111a1','82111a2','82111b',
                   '8218a1','8218a2','8218a3','8218b1','8218b2','8218b3','8218b4']

    # guarda pontos que nÑo se encaixam na planilha correspondente
    pontosFalhos_BC = []   
    listaPontosONS_BC = []  

    # escreve primeira posiáÑo com o c¢digo do Bay  
    listaPontosONS_BC.append(codBay)
   
    # varre a lista 
    for ponto in listaPontosBC:  
        itensDoPonto = ponto[posONSItem].strip().replace('.','').replace('(','').replace(')','').replace(']','')
        itensDoPonto = itensDoPonto.split('[')   
         
        for itemONS in itensDoPonto:                                    # loop para verificar todos os itens aos quais o ponto esta associado    
            if (itemONS in itensONS_BC):                                # verifica se o item est† na lista de itens esperados para LTs
                # verifica se esse item j† existe na lsita de pontos varridos para o bay
                existe = False
                for subLista in listaPontosONS_BC:                      # verifica se j† foi encontrado algum ponto do mesmo item entre os que j† foram verificados...
                    # verifica se j† foi encontrado algum ponto do mesmo item ...
                    if (subLista[0] == itemONS):
                            # ... se foi encontrado, verifica se um item que admite apenas um ponto associado (agrupamentos)
                            if (itemONS in ['8213b','82111b']): 
                                # se s¢ admitir um ponto asociado, grava como erro
                                pontosFalhos_BC.append([ponto[posIDponto],
                                                       ponto[posComentarioPonto],
                                                       itemONS,
                                                       'BANCO CAPACITOR: Duplicidade de ponto associado ao mesmo item ONS'])
                            else: 
                                # adiciona mais um ponto ao item ONS associado
                                subLista[1].append(geraPontoONSParaPlanilha(ponto))
                                existe = True
                            break
                # se nÑo existe entÑo cria
                if (existe==False):
                    listaPontosONS_BC.append([itemONS,[geraPontoONSParaPlanilha(ponto)]])
            else: # item nÑo est† na lista de itens esperados para um LT
                if(itemONS in ['8213b1','8213b2','82111b1','82111b2','82111b3','82111b4']):
                    pontosFalhos_BC.append([ponto[posIDponto],
                                           ponto[posComentarioPonto],
                                           itemONS,
                                           'BANCO CAPACITOR: O item desse ponto indica que ele deve fazer parte de um agrupamento (GRUPO B - agrupamento de eventos)'])                     
                else:    
                    pontosFalhos_BC.append([ponto[posIDponto],
                                           ponto[posComentarioPonto],
                                           itemONS,
                                           'BANCO CAPACITOR: Item ONS associado ao ponto nÑo esperado para este tipo de evento.'])         
                                                                                                                                      
    return [listaPontosONS_BC,pontosFalhos_BC]


def gerarONS(planilhaLP, nomeAba, relatorio):
    ''' 
    Là lista de pontos padrÑo e, a partir dela, gera o sub-conjunto de pontos que estÑo selecionados para formar a planilha da ONS.
    @param planilhaLP: lista de pontos que ser† avaliada.
    @param nomeAba: Nome da aba (planilha da pasta de trabalho do arquivo em Excel), que possui a lista de pontos a ser utilizada como base para gerar planilha ONS..
    @param relatorio: TKinter.ListBox que ser† utilizado para gerar relat¢rio na tela.             
    '''
    
    #dicionario de titulos a ser utilizado durante a varredura do cabeáalho da planilha com a lista de pontos
    titulo_dic = {} 
    colunasONS = [u'PROJETO.COMENTÜRIO',
                    #u'PROJETO.CONTEMPLADO',               
                    #u'CHESF-NãVEL1.TIPODORELê',                         
                    #u'CHESF-NãVEL1.UA-PAINEL',  
                    #u'CHESF-NãVEL1.BI',  
                    #u'CHESF-NãVEL1.BO',  
                    #u'CHESF-NãVEL1.IDPROTOCOLO',  
                    #u'CHESF-NãVEL1.UTIL.',  
                    u'CHESF-NãVEL2.ID(SAGE)',
                    #u'CHESF-NãVEL2.OCR(SAGE)',   
                    #u'CHESF-NãVEL2.DESCRIÄéO',   
                    u'CHESF-NãVEL2.TIPO',
                    #u'CHESF-NãVEL2.COMANDO',   
                    #u'CHESF-NãVEL2.MEDIÄéO',   
                    #u'CHESF-NãVEL2.ANUNCIADOR',   
                    #u'CHESF-NãVEL2.LISTADEALARMES',
                    #u'CHESF-NãVEL2.SOE',  
                    #u'CHESF-NãVEL2.OBSERVAÄéO',  
                    #u'CHESF-NãVEL2.AGRUPAMENTO',
                    #u'CHESF-TELEASSISTâNCIAN3.OCR(SAGE)',                                 
                    #u'CHESF-TELEASSISTâNCIAN3.COMANDO', 
                    #u'CHESF-TELEASSISTâNCIAN3.MEDIÄéO', 
                    #u'CHESF-TELEASSISTâNCIAN3.LISTADEALARME', 
                    #u'CHESF-TELEASSISTâNCIAN3.SOE', 
                    #u'CHESF-TELEASSISTâNCIAN3.OBSERVAÄéO',
                    u'CHESF-TELEASSISTâNCIAN3.ENDEREÄO',      
                    #u'CHESF-TELEASSISTâNCIAN3.AGRUPAMENTO',                                                                                                                    
                    #u'CHESF-NãVEL3.OCR(SAGE)',                                 
                    #u'CHESF-NãVEL3.COMANDO', 
                    #u'CHESF-NãVEL3.MEDIÄéO', 
                    #u'CHESF-NãVEL3.LISTADEALARME', 
                    #u'CHESF-NãVEL3.SOE', 
                    #u'CHESF-NãVEL3.OBSERVAÄéO',
                    #u'CHESF-NãVEL3.ENDEREÄO',      
                    #u'CHESF-NãVEL3.AGRUPAMENTO',    
                    u'ONS.ITEM',
                    u'ONS.DESCRIÄéO']    
      
    #Le planilha com lista de pontos -----------------------------------------------------------------
    try:        
        linhaInicial = linhaInicialETitulos(planilhaLP, nomeAba)[0]             # Là planilha e recebe a linha onde comeáa a LP (aqui est† usando somente a linha inicial e nÑo o dicion†rio de t°tuloas
        if (linhaInicial <0):                                                   # Se for um n£mero negativo entÑo nÑo foi encontrado "ID (SAGE)" na lista
            raise NameError(u'Arquivo especificado nÑo possui coluna com t°tulo "ID (SAGE).')                  
        
        arq_conf = open_workbook(planilhaLP)  
        sheet = arq_conf.sheet_by_name(nomeAba) 
        # linhaInicial = li
        # index_linha = 6     # primeira linha log ap¢s o cabeáalho da lista de pontos
        listaPontos = []      # ir† armazenar os pontos da lista de pontos

        #Leitura dos cabeáalhos da lista de pontos -----------------------------------------------
        grupo_coluna_anterior = ''
        for i in range(sheet.ncols):
            #verifica grupo da coluna a ser lida
            grupo_coluna = sheet.cell_value((linhaInicial-3),i).upper().strip().replace(' ','')
            if (grupo_coluna!=''): grupo_coluna_anterior = grupo_coluna
            
            #verifica coluna a ser lida
            if (grupo_coluna_anterior == 'ONS'): 
                texto_coluna = sheet.cell_value((linhaInicial-1),i).upper().strip()
            elif (grupo_coluna_anterior == 'CHESF-NãVEL1'): 
                if (sheet.cell_value((linhaInicial-1),i).upper().strip() == ''): 
                    # se estiver em branco o valor pode estar na linha de cima
                    texto_coluna = sheet.cell_value((linhaInicial-2),i).upper().strip()
                else: texto_coluna = sheet.cell_value((linhaInicial-1),i).upper().strip()    
            else:               
                texto_coluna = sheet.cell_value((linhaInicial-2),i).upper().strip()
            
            if (texto_coluna ==''):
                #guarda a posicao da ultima coluna em branco
                titulo_dic[texto_coluna] = i
            else: titulo_dic[grupo_coluna_anterior.replace(' ','')+'.'+texto_coluna.replace(' ','')] = i 
             
        if u'CHESF-NãVEL2.ID(SAGE)' not in titulo_dic.keys():
            NameError('Arquivo indicado nÑo corresponde a uma Lista de Pontos v†lida. Coluna "CHESF-NãVEL2.ID(SAGE)" nÑo encontrada.')
        else:         
            #Leitura de todos os pontos da planilha --------------------------------------------------
            erroLeituraColunas = False        
            for index_linha in range(linhaInicial,sheet.nrows):
                ponto = []    #Armazena o ponto lido na linha
                ponto.append(sheet.cell(index_linha,0))
                try:
                    for colunaLida in colunasONS:
                        ponto.append(sheet.cell(index_linha,titulo_dic[colunaLida]).value)

                    listaPontos.append(ponto)                              
                except:
                    showerror('Erro',u'NÑo foram encontrados todos os campos obrigat¢rios com seus t°tulos, conforme Lista de Pontos padrÑo Chesf. Verificar existància da coluna %s na planilha da lista de pontos.' % colunaLida)
                    erroLeituraColunas = True 
                    break
                #Grava o ponto lido 
    
            if (erroLeituraColunas == False):
                # Cria lista com pontos que estÑo marcados para fazerem parte da palnilha ONS ------------               
                listaPontosONS = []
                for ponto2 in listaPontos:    
                        if (ponto2[posONSItem]!=''):
                            listaPontosONS.append(ponto2)
                            
                if (len(listaPontosONS)==0):
                    raise NameError(u'NÑo foram encontrados referàncias dos pontos para ONS na coluna ONS/PROC. DE REDE/ITEM da lista de pontos. Verificar se esta coluna foi alterada/modificada.')
                
                #Descobre quais sÑo os eventos que estÑo listados entre os pontos da LP ONS junto com o c¢digo da subestaáÑo            
                eventos,CodigoSubestacao = geraListaDeEventos(listaPontosONS)            
                
                #Conjunto de listas que armazenarao pontos de forma organizada a serem gravadas na planilha ONS 
                listaEventos = geraListaEventosOrganizada(eventos)
                
                listaTTs = []
                listaTTsFalhas = []     
                listaLTs = []
                listaLTsFalhas = [] 
                listaTRs = [] 
                listaTRsFalhas = []  
                listaBTs = []        
                listaBTsFalhas = []
                listaREs = []
                listaREsFalhas = []  
                listaBAs = []
                listaBAsFalhas = []
                listaBCs = []
                listaBCsFalhas = []
              
                                           
                for event in listaEventos: # varre os eventos da lista [TT,LT,BA,BT,RE,BC,TR,CS]
                    # TT - Trafo Terra ---------------------------------------------------------------------------
                    if (listaEventos.index(event) == 0):
                        for k in event: # varre os bays dentro do evento 
                            print(k)
                    # LT - Linha de TransmissÑo ------------------------------------------------------------------
                    elif (listaEventos.index(event) == 1):
                        for k in event: # varre os bays dentro do evento 
                            #separa todos os pontos pertencentes a esse bay dentro da lista de pontos ONS
                            pontosBay = []
                            for ponto in listaPontosONS:
                                if ponto[posIDponto].find(k[1:])>0:
                                    pontosBay.append(ponto)
                            lista = geraListaLT(k,pontosBay)        
                            listaLTs.append(lista[0])
                            listaLTsFalhas.append(lista[1])  
                    # BA - Barramento -----------------------------------------------------------------------------
                    elif (listaEventos.index(event) == 2):
                        for k in event: # varre os bays dentro do evento 
                            #separa todos os pontos pertencentes a esse bay dentro da lista de pontos ONS
                            pontosBay = []
                            for ponto in listaPontosONS:
                                if ponto[posIDponto].find(k[1:])>0:
                                    pontosBay.append(ponto)
                            lista = geraListaBA(k,pontosBay)        
                            listaBAs.append(lista[0])
                            listaBAsFalhas.append(lista[1])                 
                    # BT - Bay de Transferància -------------------------------------------------------------------
                    elif (listaEventos.index(event) == 3):
                        for k in event: # varre os bays dentro do evento 
                            #separa todos os pontos pertencentes a esse bay dentro da lista de pontos ONS
                            pontosBay = []
                            for ponto in listaPontosONS:
                                if ponto[posIDponto].find(k[1:])>0:
                                    pontosBay.append(ponto)
                            lista = geraListaBT(k,pontosBay)        
                            listaBTs.append(lista[0])
                            listaBTsFalhas.append(lista[1])     
                    # RE - Reator ----------------------------------------------------------------------------------
                    elif (listaEventos.index(event) == 4):
                        for k in event: # varre os bays dentro do evento 
                            #separa todos os pontos pertencentes a esse bay dentro da lista de pontos ONS
                            pontosBay = []
                            for ponto in listaPontosONS:
                                if ponto[posIDponto].find(k[1:])>0:
                                    pontosBay.append(ponto)
                            lista = geraListaRE(k,pontosBay)        
                            listaREs.append(lista[0])
                            listaREsFalhas.append(lista[1])     
                    # BC - Banco de Capacitor -----------------------------------------------------------------------
                    elif (listaEventos.index(event) == 5):
                        for k in event: # varre os bays dentro do evento 
                            #separa todos os pontos pertencentes a esse bay dentro da lista de pontos ONS
                            pontosBay = []
                            for ponto in listaPontosONS:
                                if ponto[posIDponto].find(k[1:])>0:
                                    pontosBay.append(ponto)
                            lista = geraListaBC(k,pontosBay)        
                            listaBCs.append(lista[0])
                            listaBCsFalhas.append(lista[1])                                              
                    # TR - Transformador -----------------------------------------------------------------------------
                    elif (listaEventos.index(event) == 6):
                        for k in event: # varre os bays dentro do evento 
                            #separa todos os pontos pertencentes a esse bay dentro da lista de pontos ONS
                            pontosBay = []
                            for ponto in listaPontosONS:
                                if ponto[posIDponto].find(k[2:])>0:
                                    pontosBay.append(ponto)      
                            lista = geraListaTR(k,pontosBay) 
                            listaTRs.append(lista[0])
                            listaTRsFalhas.append(lista[1])  
                    # TT - Trafo Terra -------------------------------------------------------------------------------
                    '''
                    elif (listaEventos.index(event) == 7):
                        for k in event: # varre os bays dentro do evento 
                            #separa todos os pontos pertencentes a esse bay dentro da lista de pontos ONS
                            pontosBay = []
                            for ponto in listaPontosONS:
                                if ponto[posIDponto].find(k[2:])>0:
                                    pontosBay.append(ponto)      
                            lista = geraListaTT(k,pontosBay) 
                            listaTTs.append(lista[0])
                            listaTTsFalhas.append(lista[1])                                           
                     '''
                listaPlanilhaONS = [listaTTs, listaLTs, listaTRs, listaBTs, listaREs, listaBAs, listaBCs] 
                listaFalhas = [listaTTsFalhas, listaLTsFalhas, listaTRsFalhas,listaBTsFalhas, listaREsFalhas, listaBAsFalhas, listaBCsFalhas]
                arquivoGerado = gerarPlanilhaONS(CodigoSubestacao,listaPlanilhaONS, listaFalhas)
                
                # Gera relat¢rio na tela
                relatorio.delete(0,END)
                relatorio.insert(END,u'Relat¢rio ONS Gerado. Eventos encontrados:')       
                for eventos in listaEventos:
                    for bay in eventos:
                        texto = bay
                        if (listaEventos.index(eventos) in [0,7]):
                            texto = '      '+texto + u' (ONS ainda nÑo implementado)'                              
                        relatorio.insert(END,texto.rjust(10,' '))        
                relatorio.insert(END,u'')
                totalInconsistencias = 0
                for k in listaFalhas:
                    for j in k:
                        totalInconsistencias = totalInconsistencias + len(j)
                relatorio.insert(END,u'Total de inconsistàncias encontradas: %d'%(totalInconsistencias))    
                               
                abrirarquivo = askyesno('Aviso', u'Arquivo \"'+arquivoGerado+'\" gerado em ' + getcwd()+'\n\n Deseja abrir o arquivo gerado agora?')
                if abrirarquivo : startfile(getcwd() + '\\' + arquivoGerado)                      
                
                arquivoLog = open('logONS.txt','w')
                arquivoLog.write('--------------- Pontos Inconsistentes -------------------\n\n') 
                arquivoLog.write(u'ID SAGE          \tCOMENTÜRIO         \tITEM PR 2.7\tDESCRIÄéO INCONSISTâNCIA\n')    
                for k in listaLTsFalhas:
                    for j in k: 
                        arquivoLog.write(j[0]+'\t'+j[1]+'     \t'+j[2]+'\n')
                for k in listaTRsFalhas:
                    for j in k: 
                        arquivoLog.write(j[0]+'\t'+j[1]+'     \t'+j[2]+'\n')                       
                        
                arquivoLog.close()    
   
    except NameError as e:
        showerror('Erro', e)
    except:
        showerror('Erro',u'Arquivo especificado nÑo foi encontrado ou foi informada uma planilha com formataáÑo diferente da planilha PadrÑo Chesf.')
    
