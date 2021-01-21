# -*- coding: cp860 -*-

dados = '''
Vers„o 2.0.4
Atualiza‡„o do programa: 02/12/2014
Gerar LP baseado nos m¢dulos LP e gerarPlanlhaLP
'''

from tkinter.messagebox import showerror, askyesno
import os
from tkinter import END
import pickle

try:
    from xlrd import open_workbook
except:
    showerror('Erro','Modulo xlrd n„o instalado')
try:
    from lp_lib.LP import gerarlp
except:
    showerror('Erro','Arquivo "LP.py" deve estar no mesmo diret¢rio "lp_lib"')

try:
    from lp_lib.gerarPlanilhaLP import gerarPlanilha
except:
    showerror('Erro','Arquivo "gerarPlanilhaLP.py" deve estar no mesmo diret¢rio "lp_lib"')


def gerar(LP_Padrao,relatorio,LP_Config):

    try:
        arq_conf = open_workbook(LP_Config) #Abrir arquivo de configura‡„o
    except:
        showerror('Erro','Arquivo de parametriza‡„o n„o encontrado')
        print(LP_Config)
    try:
        sheet = arq_conf.sheet_by_index(0)           #Abrir planilha "Configura‡”es" do arquivo LP_config.xls
        Codigo_SE = sheet.cell(4,1).value.upper()    #Ler defini‡„o do c¢digo da SE
    except:
        showerror('Erro','Arquivo indicado n„o corresponde a arquivo de parametriza‡„o v lido')


    nome_arq_saida = './LP_gerada_%s.xlsx'%(Codigo_SE)       #Nome do arquivo de sa¡da
    seq_arq = 0                               #Sequˆncia do n£mero de arquivo
    while os.path.exists(nome_arq_saida):   #Enquanto existir na pasta um arquivo com o nome definido
        seq_arq += 1                          #Adicionar um a sequˆncia do n£mero do arquivo
        nome_arq_saida = nome_arq_saida[0:11]+'_'+Codigo_SE+'_'+str(seq_arq)+'.xlsx' #Definir novo nome de arquivo (Ex './LP_gerada_JRM_1.xls)
    #arq_LP.save(nome_arq_saida[2:])         #Gravar o nome do arquivo excluindo './' do nome
 
    
    saida = gerarlp(LP_Padrao,LP_Config)

    arq_LP = gerarPlanilha(nome_arq_saida)                    # Gera um arquivo Excel com uma planilha com formata‡„o da Lista de Pontos Padr„o
    planilha_LP = arq_LP.worksheets()[0]

    linha = 6
    for dado in saida[0]: # Passa por todas as linhas do array de sa¡da gravando pontos no Excel
        tag = dado[0]
        ocr = dado[1].value
        descr = dado[2]
        if dado[9].value == 'Em caso de religamento Monopolar' or\
           dado[9].value == 'Para arranjos disjuntor e meio.' or\
           str(dado[9].value).find('{PNL}') > -1 or\
           dado[9].value == 'Apenas para Banco Monof sico.' or\
           dado[9].value == 'Apenas para Trafo Trif sico.' or\
           dado[9].value == 'Sele‡„o de sincronismo para Barra Dupla' or\
           dado[9].value == '\"n\" - N£mero do Canal Ÿptico.' or\
           str(dado[9].value).startswith('Para sistemas #PASS') or\
           str(dado[9].value).find('(SAGE, UTR-, PCPG)') > -1 or\
           str(dado[9].value).find('{DISP}') > -1 or\
           str(dado[9].value).find('UCn') > -1 or\
           str(dado[9].value).find('"n"-') > -1 or\
           str(dado[9].value).find('n-') > -1 :
            obs = ''
        else:
            obs = dado[9].value

        planilha_LP.write(linha,0,linha-5)                # escreve na coluna "ITEM" na planilha
        planilha_LP.write(linha,9,tag)                    # escreve na coluna "TAG" na planilha
        planilha_LP.write(linha,10,ocr)                   # escreve na coluna "OCR" na planilha
        planilha_LP.write(linha,11,descr)                 # escreve na coluna "DESCRI€ŽO" na planilha
        planilha_LP.write(linha,12,dado[3].value)         # escreve na coluna "TIPO" na planilha
        planilha_LP.write(linha,13,dado[4].value)         # escreve na coluna "COMANDO" na planilha
        planilha_LP.write(linha,14,dado[5].value)         # escreve na coluna "MEDI€ŽO" na planilha
        planilha_LP.write(linha,15,dado[6].value)         # escreve na coluna "TELA" na planilha
        planilha_LP.write(linha,16,dado[7].value)         # escreve na coluna "LISTA DE ALARMES" na planilha
        planilha_LP.write(linha,17,dado[8].value)         # escreve na coluna "SOE" na planilha
        planilha_LP.write(linha,18,obs)                   # escreve na coluna "OBSERVA€ŽO" na planilha
        for colunaxls,campodados in zip(range(19,46),range(10,37)):
            planilha_LP.write(linha,colunaxls,dado[campodados].value)
        linha += 1                                          # incrementa a linha

    arq_LP.close()
    
    #----------Relat¢rio de Gera‡„o de Pontos----------#
    total = 0
    nome_arq_log = './log_{}_GER.txt'.format(Codigo_SE)            # Nome do arquivo de log
    seq_arq = 0                                                    # Sequˆncia do n£mero de arquivo
    while os.path.exists(nome_arq_log):                            # Enquanto existir na pasta um arquivo com o nome definido
        seq_arq += 1                                               # Adicionar um a sequˆncia do n£mero do arquivo
        nome_arq_log = 'log_{}_GER_{}.txt'.format(Codigo_SE, seq_arq) #Definir novo nome de arquivo
    arq_log = open(nome_arq_log,'w')

    relatorio.insert(END,'-----Pontos Gerados-----')
    arq_log.write('-----Pontos Gerados-----\n\n')
    relatorio.insert(END,'')
    for k,evento in ([7,'SAGE/REDE'],
                     [0,'LT'],
                     [3,'Trafo'],
                     [4,'Vao de Transf'],
                     [6,'T_Terra'],
                     [9,'B_CAP'],
                     [1,'Disjuntor'],
                     [2,'Secc'],
                     [10,'BCS'],
                     [5,'Reator'],
                     [11,'SAs'],
                     [8,'BARRA'],
                     [12,'ECE'],
                     [13, 'CS'],
                     [14, 'Prep. Reen.'],
                     [15, 'Sistema Regulacao'],
                     [16, 'Painel de Interface']):

        if saida[1][k]>0: 
            relatorio.insert(END,evento.ljust(30,'_')+str(saida[1][k]).rjust(3)+' pontos')
            arq_log.write(evento.ljust(30,'_')+str(saida[1][k]).rjust(3)+' pontos\n')
        total += saida[1][k]

    relatorio.insert(END,'')
    arq_log.write('\n')
    relatorio.insert(END,'Total: '+str(total))
    arq_log.write('Total: '+str(total))
    arq_log.close()
    #showinfo('Aviso','Arquivo \"'+nome_arq_saida[2:]+'\" gerado em '+os.getcwd().decode('latin_1'))
    
    abrirarquivo = askyesno('Aviso', 'Arquivo \"'+nome_arq_saida[2:]+'\" gerado em ' + os.getcwd()+'\n\n Deseja abrir o arquivo gerado agora?')
    if abrirarquivo : os.startfile(os.getcwd() + '\\' + nome_arq_saida[2:])
      
    nomearquivo = nome_arq_saida[2:]
        
    conf = {'arquivo':nomearquivo}
    pickle.dump(conf, open('fas.p','wb'),-1) #-1 para gravar em Bin rio
    
