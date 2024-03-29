# -*- coding: cp860 -*-

import FASgtkui
import os
import pickle
import gi
import time
dados = '''
Vers�o 2.0.13
Atualiza��o do programa: 02/12/2014
Gerar LP baseado nos m�dulos LP e gerarPlanlhaLP
'''

gi.require_version("Gtk", "3.0")
from gi.repository import GObject

try:
    from bs4 import BeautifulSoup
except:
    FASgtkui.mensagem_erro('Erro', 'M�dulo BeautifulSoup n�o instalado')
try:
    from lp_lib.LP import gerarlp
except:
    FASgtkui.mensagem_erro('Erro', 'Arquivo "LP.py" deve estar no mesmo diret�rio "lp_lib"')

try:
    from lp_lib.gerarPlanilhaLP import gerarPlanilha
except:
    FASgtkui.mensagem_erro('Erro', 'Arquivo "gerarPlanilhaLP.py" deve estar no mesmo diret�rio "lp_lib"')


def gerar(LP_Padrao, Arq_Conf, Diretorio_Padrao):
    try:
        arq_conf = BeautifulSoup(open(Arq_Conf, 'r', encoding='utf-8'), 'html.parser')  # Abrir arquivo de cofigura��o
    except:
        GObject.idle_add(FASgtkui.mensagem_erro,'Erro', 'Arquivo de parametriza��o n�o encontrado')
        time.sleep(1)
        while FASgtkui.mensagem_erro_dialog.get_visible() == True:
            time.sleep(1)
    try:
        Codigo_SE = arq_conf.eventos['codigo_se']  # Ler defini��o do c�digo da SE
    except:
        GObject.idle_add(FASgtkui.mensagem_erro,'Erro',
                               'Arquivo indicado n�o corresponde a arquivo de parametriza��o v�lido, c�digo da SE n�o encontrado')
        time.sleep(1)
        while FASgtkui.mensagem_erro_dialog.get_visible() == True:
            time.sleep(1)
    nome_arq_saida = 'LP_gerada_%s.xlsx' % (Codigo_SE)  # Nome do arquivo de sa�da
    seq_arq = 0  # Sequ�ncia do n�mero de arquivo

    while os.path.exists(
            Diretorio_Padrao + '\\' + nome_arq_saida):  # Enquanto existir na pasta um arquivo com o nome definido
        seq_arq += 1  # Adicionar um a sequ�ncia do n�mero do arquivo
        nome_arq_saida = 'LP_gerada' + '_' + Codigo_SE + '_' + str(
            seq_arq) + '.xlsx'  # Definir novo nome de arquivo (Ex './LP_gerada_JRM_1.xls)

    nome_arq_saida = Diretorio_Padrao + '\\' + nome_arq_saida
    # arq_LP.save(nome_arq_saida[2:])         #Gravar o nome do arquivo excluindo './' do nome

    saida = gerarlp(LP_Padrao, Arq_Conf)

    arq_LP = gerarPlanilha(
        nome_arq_saida)  # Gera um arquivo Excel com uma planilha com formata��o da Lista de Pontos Padr�o
    planilha_LP = arq_LP.worksheets()[0]

    linha = 6
    for dado in saida[0]:  # Passa por todas as linhas do array de sa�da gravando pontos no Excel
        tag = dado[0]
        ocr = dado[1].value
        descr = dado[2]
        if dado[9].value == 'Em caso de religamento Monopolar' or \
                dado[9].value == 'Para arranjos disjuntor e meio.' or \
                str(dado[9].value).find('{PNL}') > -1 or \
                dado[9].value == 'Apenas para Banco Monof�sico.' or \
                dado[9].value == 'Apenas para Trafo Trif�sico.' or \
                dado[9].value == 'Sele��o de sincronismo para Barra Dupla' or \
                dado[9].value == '\"n\" - N�mero do Canal �ptico.' or \
                str(dado[9].value).startswith('Para sistemas #PASS') or \
                str(dado[9].value).find('(SAGE, UTR-, PCPG)') > -1 or \
                str(dado[9].value).find('{DISP}') > -1 or \
                str(dado[9].value).find('UCn') > -1 or \
                str(dado[9].value).find('"n"-') > -1 or \
                str(dado[9].value).find('n-') > -1:
            obs = ''
        else:
            obs = dado[9].value

        planilha_LP.write(linha, 0, linha - 5)  # escreve na coluna "ITEM" na planilha
        planilha_LP.write(linha, 9, tag)  # escreve na coluna "TAG" na planilha
        planilha_LP.write(linha, 10, ocr)  # escreve na coluna "OCR" na planilha
        planilha_LP.write(linha, 11, descr)  # escreve na coluna "DESCRI��O" na planilha
        planilha_LP.write(linha, 12, dado[3].value)  # escreve na coluna "TIPO" na planilha
        planilha_LP.write(linha, 13, dado[4].value)  # escreve na coluna "COMANDO" na planilha
        planilha_LP.write(linha, 14, dado[5].value)  # escreve na coluna "MEDI��O" na planilha
        planilha_LP.write(linha, 15, dado[6].value)  # escreve na coluna "TELA" na planilha
        planilha_LP.write(linha, 16, dado[7].value)  # escreve na coluna "LISTA DE ALARMES" na planilha
        planilha_LP.write(linha, 17, dado[8].value)  # escreve na coluna "SOE" na planilha
        planilha_LP.write(linha, 18, obs)  # escreve na coluna "OBSERVA��O" na planilha
        for colunaxls, campodados in zip(range(19, 46), range(10, 37)):
            planilha_LP.write(linha, colunaxls, dado[campodados].value)
        linha += 1  # incrementa a linha

    arq_LP.close()

    # ----------Relat�rio de Gera��o de Pontos----------#
    total = 0
    nome_arq_log = 'log_{}_GER.txt'.format(Codigo_SE)  # Nome do arquivo de log
    seq_arq = 0  # Sequ�ncia do n�mero de arquivo
    while os.path.exists(
            Diretorio_Padrao + '\\' + nome_arq_log):  # Enquanto existir na pasta um arquivo com o nome definido
        seq_arq += 1  # Adicionar um a sequ�ncia do n�mero do arquivo
        nome_arq_log = 'log_{}_GER_{}.txt'.format(Codigo_SE, seq_arq)  # Definir novo nome de arquivo
    nome_arq_log = Diretorio_Padrao + '\\' + nome_arq_log
    arq_log = open(nome_arq_log, 'w')
    arq_log.write('-----Pontos Gerados-----\n\n')
    for k, evento in ([7, 'SAGE/REDE'],
                      [0, 'LT'],
                      [3, 'Trafo'],
                      [4, 'Vao de Transf'],
                      [6, 'T_Terra'],
                      [9, 'B_CAP'],
                      [1, 'Disjuntor'],
                      [2, 'Secc'],
                      [10, 'BCS'],
                      [5, 'Reator'],
                      [11, 'SAs'],
                      [8, 'BARRA'],
                      [12, 'ECE'],
                      [13, 'CS'],
                      [14, 'Prep. Reen.'],
                      [15, 'Sistema Regulacao'],
                      [16, 'Painel de Interface']):
        if saida[1][k] > 0:
            arq_log.write(evento.ljust(30, '_') + str(saida[1][k]).rjust(3) + ' pontos\n')
        total += saida[1][k]
    arq_log.write('\n')
    arq_log.write('Total: ' + str(total))
    arq_log.close()
    GObject.idle_add(FASgtkui.dialogo_abrir_arquivo_gerado, nome_arq_saida.rsplit('\\', 1)[1], Diretorio_Padrao)

    nomearquivo = nome_arq_saida[2:]
    conf = {'arquivo': nomearquivo}
    pickle.dump(conf, open('fas.p', 'wb'), -1)  # -1 para gravar em Bin�rio
