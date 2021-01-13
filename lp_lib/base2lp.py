# -*- coding: cp860 -*-
from tkinter.messagebox import askyesno
from os import listdir, path, getcwd, startfile
from FASgtkui import mensagem_aviso, mensagem_erro, pergunta_sim_nao

dados = '''
Vers„o 2.0.12
Atualiza‡„o do programa: 10/11/2020
Gera‡„o de LP no padr„o Chesf tendo como base os arquivos .dat de configura‡„o de base SAGE
'''

try:
    from lp_lib.gerarPlanilhaLP import gerarPlanilha
except:
    mensagem_erro('Erro', 'Arquivo "gerarPlanilhaLP.py" deve estar no diret¢rio "lp_lib"')


def base2lp(diretorio):
    nome_arq_saida = './LP_da_Base.xlsx'  # Nome do arquivo de sa¡da
    seq_arq = 0  # Sequˆncia do n£mero de arquivo
    while path.exists(nome_arq_saida):  # Enquanto existir na pasta um arquivo com o nome definido
        seq_arq += 1  # Adicionar um a sequˆncia do n£mero do arquivo
        nome_arq_saida = nome_arq_saida[0:12] + '_' + str(seq_arq) + '.xlsx'  # Definir novo nome de arquivo

    arq_lp = gerarPlanilha(nome_arq_saida)  # Gera um arquivo Excel com uma planilha com formata‡„o da LP Padr„o
    planilha_lp = arq_lp.worksheets()[0]
    planilha_relatorio = arq_lp.add_worksheet('RELATORIO')

    # ***** Captar texto de arquivos de telas para verificar preenchimento da coluna anunciador *****
    try:
        nome_arq_tela = '{}/ihm/VTelasBotoes.led'.format(diretorio.rsplit('/', 2)[0])
        arq_telas = open(nome_arq_tela)
        telas = []
        for linha in arq_telas.readlines():
            if 'TELA' in linha:
                telas.append(linha.split('\"')[1].split()[1])
        arq_telas.close()
    except:
        mensagem_aviso('Aviso', 'O arquivo {}/ihm/VTelasBotoes.led n„o pode ser carregado. A coluna "ANUNCIADOR" da \
                    tabela gerada n„o ser  preenchida.'.format(diretorio.rsplit('/', 2)[0]))
        telas = []

    texto_telas = ''
    for tela in telas:
        try:
            arq_txt = open('{}/telas/{}'.format(diretorio.rsplit('/', 2)[0], tela), 'r')
            texto_telas += arq_txt.read()
            arq_txt.close()
        except:
            mensagem_erro('Erro', 'O arquivo {}/{} n„o pode ser carregado'.format(diretorio.rsplit('/', 2)[0], tela))

    if 'pds.dat' not in listdir(diretorio):
        mensagem_erro('Erro', 'Arquivo "pds.dat" n„o encontrado no diret¢rio {}'.format(diretorio))
    elif 'pdd.dat' not in listdir(diretorio):
        mensagem_erro('Erro', 'Arquivo "pdd.dat" n„o encontrado no diret¢rio {}'.format(diretorio))
    elif 'pdf.dat' not in listdir(diretorio):
        mensagem_erro('Erro', 'Arquivo "pdf.dat" n„o encontrado no diret¢rio {}'.format(diretorio))
    elif 'pas.dat' not in listdir(diretorio):
        mensagem_erro('Erro', 'Arquivo "pas.dat" n„o encontrado no diret¢rio {}'.format(diretorio))
    elif 'pad.dat' not in listdir(diretorio):
        mensagem_erro('Erro', 'Arquivo "pad.dat" n„o encontrado no diret¢rio {}'.format(diretorio))
    elif 'paf.dat' not in listdir(diretorio):
        mensagem_erro('Erro', 'Arquivo "paf.dat" n„o encontrado no diret¢rio {}'.format(diretorio))
    elif 'cgs.dat' not in listdir(diretorio):
        mensagem_erro('Erro', 'Arquivo "cgs.dat" n„o encontrado no diret¢rio {}'.format(diretorio))
    elif 'cgf.dat' not in listdir(diretorio):
        mensagem_erro('Erro', 'Arquivo "cgf.dat" n„o encontrado no diret¢rio {}'.format(diretorio))
    elif 'tac.dat' not in listdir(diretorio):
        mensagem_erro('Erro', 'Arquivo "tac.dat" n„o encontrado no diret¢rio {}'.format(diretorio))
    else:
        # ********** PDS **********
        def pdsfunc(caminho):
            # Se n„o acabar com ".dat" insere "pds.dat", sen„o j  ‚ resultado do include e vem com .dat no caminho
            arq = open('{}\\pds.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0  # Para verificar fim de arquivo
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)  # Retornar ao in¡cio do arquivo
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    if len(
                            linha.strip().split()) <= 2:  # Caso o nome do arquivo de include n„o seja separado por espa‡o em branco
                        include = '{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))
                    elif len(
                            linha.strip().split()) > 2:  # Caso o nomde do arquivo de include tenha separa‡„o por espa‡o em branco Ex: "BT 14D1"
                        include = '{}\\{} {}'.format(caminho, linha.strip().split()[1].replace('/', '\\'),
                                                     linha.strip().split()[2].replace('/', '\\'))
                    dic.update(pdsfunc(include))
                elif linha.strip() == 'PDS':
                    ID = OCR = NOME = TIPO = TAC = ''
                    ALRIN = SOE = 'X'
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'PDS' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'OCR':
                        OCR = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'NOME':
                        NOME = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'TIPO':
                        TIPO = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'ALRIN':
                        ALRIN = ('' if linha.split('=')[1].strip() == 'SIM' else 'X')
                    elif linha.split('=')[0].strip() == 'SOEIN':
                        SOE = ('' if linha.split('=')[1].strip() == 'SIM' else 'X')
                    elif linha.split('=')[0].strip() == 'TAC':
                        TAC = linha.split('=')[1].strip()
                    if k == kf:  # Fim do arquivo
                        dic[ID] = [OCR, NOME, TIPO, ALRIN, SOE, TAC]
                elif novo:
                    dic[ID] = [OCR, NOME, TIPO, ALRIN, SOE, TAC]
                    novo = False
            arq.close()
            return dic

        # ********** PDD **********
        def pddfunc(caminho):
            arq = open('{}\\pdd.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    dic.update(pddfunc('{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))))
                elif linha.strip() == 'PDD':
                    ID = PDS = ''
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'PDD' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'PDS':
                        PDS = linha.split('=')[1].strip()
                    if k == kf:
                        dic[PDS] = [ID]
                elif novo:
                    dic[PDS] = [ID]
                    novo = False
            arq.close()
            return dic

        # ********** PDF **********
        def pdffunc(caminho):
            arq = open('{}\\pdf.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    if len(
                            linha.strip().split()) <= 2:  # Caso o nome do arquivo de include n„o seja separado por espa‡o em branco
                        include = '{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))
                    elif len(
                            linha.strip().split()) > 2:  # Caso o nomde do arquivo de include tenha separa‡„o por espa‡o em branco Ex: "BT 14D1"
                        include = '{}\\{} {}'.format(caminho, linha.strip().split()[1].replace('/', '\\'),
                                                     linha.strip().split()[2].replace('/', '\\'))
                    dic.update(pdffunc(include))
                elif linha.strip() == 'PDF':
                    ID = PNT = ''
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'PDF' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'PNT':
                        PNT = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'ORDEM':
                        ORDEM = linha.split('=')[1].strip()
                    if k == kf:
                        if 'ORDEM' not in locals(): ORDEM = ''
                        dic[PNT] = [ID, ORDEM]
                elif novo:
                    if 'ORDEM' not in locals(): ORDEM = ''
                    dic[PNT] = [ID, ORDEM]
                    novo = False
            arq.close()
            return dic

        # ********** PTS **********
        def ptsfunc(caminho):
            arq = open('{}\\pts.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    dic.update(ptsfunc('{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))))
                elif linha.strip() == 'PTS':
                    ID = OCR = NOME = TIPO = LSA = LSE = LSU = TAC = ''
                    ALRIN = 'X'
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'PTS' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'OCR':
                        OCR = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'NOME':
                        NOME = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'TIPO':
                        TIPO = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'ALRIN':
                        ALRIN = ('' if linha.split('=')[1].strip() == 'SIM' else 'X')
                    elif linha.split('=')[0].strip() == 'LSA':
                        LSA = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'LSE':
                        LSE = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'LSU':
                        LSU = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'TAC':
                        TAC = linha.split('=')[1].strip()
                    if k == kf: dic[ID] = [OCR, NOME, TIPO, ALRIN, LSA, LSE, LSU, TAC]
                elif novo:
                    dic.update({ID: [OCR, NOME, TIPO, ALRIN, LSA, LSE, LSU, TAC]})
                    novo = False
            arq.close()
            return dic

        # ********** PTD **********
        def ptdfunc(caminho):
            arq = open('{}\\ptd.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    dic.update(ptdfunc('{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))))
                elif linha.strip() == 'PTD':
                    ID = PTS = ''
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'PTD' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'PTS':
                        PTS = linha.split('=')[1].strip()
                    if k == kf:
                        dic[PTS] = [ID]
                elif novo:
                    dic[PTS] = [ID]
                    novo = False
            arq.close()
            return dic

        # ********** PTF **********
        def ptffunc(caminho):
            arq = open('{}\\ptf.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    dic.update(ptffunc('{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))))
                elif linha.strip() == 'PTF':
                    ID = PNT = ''
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'PTF' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'PNT':
                        PNT = linha.split('=')[1].strip()
                    if k == kf: dic[PNT] = [ID]
                elif novo:
                    dic[PNT] = [ID]
                    novo = False
            arq.close()
            return dic

        # ********** PAS **********
        def pasfunc(caminho):
            arq = open('{}\\pas.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    if len(
                            linha.strip().split()) <= 2:  # Caso o nome do arquivo de include n„o seja separado por espa‡o em branco
                        include = '{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))
                    elif len(
                            linha.strip().split()) > 2:  # Caso o nomde do arquivo de include tenha separa‡„o por espa‡o em branco Ex: "BT 14D1"
                        include = '{}\\{} {}'.format(caminho, linha.strip().split()[1].replace('/', '\\'),
                                                     linha.strip().split()[2].replace('/', '\\'))
                    dic.update(pasfunc(include))
                elif linha.strip() == 'PAS':
                    ID = OCR = NOME = TIPO = LIU = LIE = LIA = LSA = LSE = LSU = BNDMO = TAC = ''
                    ALRIN = 'X'
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'PAS' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'OCR':
                        OCR = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'NOME':
                        NOME = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'TIPO':
                        TIPO = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'ALRIN':
                        ALRIN = ('' if linha.split('=')[1].strip() == 'SIM' else 'X')
                    elif linha.split('=')[0].strip() == 'LIU':
                        LIU = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'LIE':
                        LIE = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'LIA':
                        LIA = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'LSA':
                        LSA = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'LSE':
                        LSE = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'LSU':
                        LSU = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'BNDMO':
                        BNDMO = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'TAC':
                        TAC = linha.split('=')[1].strip()
                    if k == kf:
                        dic[ID] = [OCR, NOME, TIPO, ALRIN, LIU, LIE, LIA, LSA, LSE, LSU, BNDMO, TAC]
                elif novo:
                    dic[ID] = [OCR, NOME, TIPO, ALRIN, LIU, LIE, LIA, LSA, LSE, LSU, BNDMO, TAC]
                    novo = False
            arq.close()
            return dic

        # ********** PAD **********
        def padfunc(caminho):
            arq = open('{}\\pad.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    dic.update(padfunc('{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))))
                elif linha.strip() == 'PAD':
                    ID = PAS = ''
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'PAD' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'PAS':
                        PAS = linha.split('=')[1].strip()
                    if k == kf: dic[PAS] = [ID]
                elif novo:
                    dic[PAS] = [ID]
                    novo = False
            arq.close()
            return dic

        # ********** PAF **********
        def paffunc(caminho):
            arq = open('{}\\paf.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    if len(
                            linha.strip().split()) <= 2:  # Caso o nome do arquivo de include n„o seja separado por espa‡o em branco
                        include = '{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))
                    elif len(
                            linha.strip().split()) > 2:  # Caso o nomde do arquivo de include tenha separa‡„o por espa‡o em branco Ex: "BT 14D1"
                        include = '{}\\{} {}'.format(caminho, linha.strip().split()[1].replace('/', '\\'),
                                                     linha.strip().split()[2].replace('/', '\\'))
                    dic.update(paffunc(include))
                elif linha.strip() == 'PAF':
                    ID = PNT = ''
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'PAF' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'PNT':
                        PNT = linha.split('=')[1].strip()
                    if k == kf: dic[PNT] = [ID]
                elif novo:
                    dic[PNT] = [ID]
                    novo = False
            arq.close()
            return dic

        # ********** CGS **********
        def cgsfunc(caminho):
            arq = open('{}\\cgs.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1

                if linha.strip().startswith('#include'):
                    if len(
                            linha.strip().split()) <= 2:  # Caso o nome do arquivo de include n„o seja separado por espa‡o em branco
                        include = '{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))
                    elif len(
                            linha.strip().split()) > 2:  # Caso o nomde do arquivo de include tenha separa‡„o por espa‡o em branco Ex: "BT 14D1"
                        include = '{}\\{} {}'.format(caminho, linha.strip().split()[1].replace('/', '\\'),
                                                     linha.strip().split()[2].replace('/', '\\'))
                    dic.update(cgsfunc(include))
                elif linha.strip() == 'CGS':
                    ID = NOME = TAC = PAC = TIPOE = ''
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'CGS' and linha != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'NOME':
                        NOME = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'PAC':
                        PAC = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'TAC':
                        TAC = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'TIPOE':
                        TIPOE = linha.split('=')[1].strip()
                    if k == kf: dic[ID] = [NOME, PAC, TAC, TIPOE]
                elif novo:
                    dic.update({ID: [NOME, PAC, TAC, TIPOE]})
                    novo = False
            arq.close()
            return dic

        # ********** CGF **********
        def cgffunc(caminho):
            arq = open('{}\\cgf.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    if len(
                            linha.strip().split()) <= 2:  # Caso o nome do arquivo de include n„o seja separado por espa‡o em branco
                        include = '{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))
                    elif len(
                            linha.strip().split()) > 2:  # Caso o nomde do arquivo de include tenha separa‡„o por espa‡o em branco Ex: "BT 14D1"
                        include = '{}\\{} {}'.format(caminho, linha.strip().split()[1].replace('/', '\\'),
                                                     linha.strip().split()[2].replace('/', '\\'))
                    dic.update(cgffunc(include))
                elif linha.strip() == 'CGF':
                    ID = PNT = ''
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'CGF' and linha != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'CGS':
                        CGS = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'KCONV':
                        if linha.split('=')[1].strip() == 'CGS':
                            CGS = linha.split('=')[2].strip()
                    elif linha.split('=')[0].strip() == 'NV2':
                        NV2 = linha.split('=')[1].strip()
                    if k == kf:
                        dic[CGS] = [ID, NV2]
                elif novo:
                    dic[CGS] = [ID, NV2]
                    novo = False
            arq.close()
            return dic

        # ********** TAC **********
        def tacfunc(caminho):
            arq = open('{}\\tac.dat'.format(caminho)) if not caminho.endswith('.dat') else open(caminho)
            dic = {}
            novo = False
            kf = k = 0
            for linha in arq.readlines():
                kf += 1
            arq.seek(0)
            for linha in arq.readlines():
                k += 1
                if linha.strip().startswith('#include'):
                    if len(
                            linha.strip().split()) <= 2:  # Caso o nome do arquivo de include n„o seja separado por espa‡o em branco
                        include = '{}\\{}'.format(caminho, linha.strip().split()[1].replace('/', '\\'))
                    elif len(
                            linha.strip().split()) > 2:  # Caso o nomde do arquivo de include tenha separa‡„o por espa‡o em branco Ex: "BT 14D1"
                        include = '{}\\{} {}'.format(caminho, linha.strip().split()[1].replace('/', '\\'),
                                                     linha.strip().split()[2].replace('/', '\\'))
                    dic.update(tacfunc(include))
                elif linha.strip() == 'TAC':
                    ID = LSC = ''
                    novo = True
                elif novo and (linha.strip() != '' and linha.strip() != 'TAC' and linha[0] != ';'):
                    if linha.split('=')[0].strip() == 'ID':
                        ID = linha.split('=')[1].strip()
                    elif linha.split('=')[0].strip() == 'LSC':
                        LSC = linha.split('=')[1].strip()
                    if k == kf:
                        dic[ID] = [LSC]
                elif novo:
                    dic[ID] = [LSC]
                    novo = False
            arq.close()
            return dic

        pds_dic = pdsfunc(diretorio)
        pdd_dic = pddfunc(diretorio)
        pdf_dic = pdffunc(diretorio)
        pas_dic = pasfunc(diretorio)
        pad_dic = padfunc(diretorio)
        paf_dic = paffunc(diretorio)
        cgs_dic = cgsfunc(diretorio)
        cgf_dic = cgffunc(diretorio)
        tac_dic = tacfunc(diretorio)
        try:
            pts_dic = ptsfunc(diretorio)
        except:
            mensagem_aviso('Notifica‡„o', 'Arquivo "pts.dat" n„o foi processado')
        try:
            ptd_dic = ptdfunc(diretorio)
        except:
            mensagem_aviso('Notifica‡„o', 'Arquivo "ptd.dat" n„o foi processado')
            ptd_dic = ''
        try:
            ptf_dic = ptffunc(diretorio)
        except:
            mensagem_aviso('Notifica‡„o', 'Arquivo "ptf.dat" n„o foi processado')

        # ********** Gravar Pontos Digitais Excel **********
        pdig = []
        linha_rel = 1
        planilha_relatorio.write(0, 0, 'ID n„o encontrados em PDF')
        for id_tag in pds_dic:
            try:
                pdf_id = pdf_dic[id_tag]
            except:
                pdf_id = ['']
                if 'CALC' not in pds_dic[id_tag][5] and 'LOCAL' not in pds_dic[id_tag][5]:
                    planilha_relatorio.write(linha_rel, 0, id_tag)
                    linha_rel += 1
            pdig.append([id_tag] + pds_dic[id_tag] + pdf_id)

        linha = 6
        for dado in pdig:  # Passa por todas as linhas do array de pontos digitais gravando pontos no Excel
            tac = dado[6]
            lsc = tac_dic.get(tac, ['?'])[0]
            id_protocolo = dado[7]
            tag = dado[0]
            ocr = dado[1]
            descr = dado[2]
            tipo = dado[3]
            anunciador = ('X' if tag in texto_telas else '')
            alarme = dado[4]
            soe = dado[5]
            obs = ''

            id_pdd = pdd_dic.get(tag, ['', ''])
            end = pdf_dic.get(id_pdd[0], ['', ''])[1]

            planilha_lp.write(linha, 0, linha - 5)  # escreve na coluna "ITEM"
            planilha_lp.write(linha, 2, tac)  # escreve na coluna "TAC"
            planilha_lp.write(linha, 3, lsc)  # escreve na coluna "IED"
            planilha_lp.write(linha, 7, id_protocolo)  # escreve na coluna "ID PROTOCOLO"
            planilha_lp.write(linha, 9, tag)  # escreve na coluna "ID (SAGE)"
            planilha_lp.write(linha, 10, ocr)  # escreve na coluna "OCR"
            planilha_lp.write(linha, 11, descr)  # escreve na coluna "DESCRI€ŽO"
            planilha_lp.write(linha, 12, tipo)  # escreve na coluna "TIPO"
            planilha_lp.write(linha, 15, anunciador)  # escreve na coluna "ANUNCIADOR"
            planilha_lp.write(linha, 16, alarme)  # escreve na coluna "LISTA DE ALARMES"
            planilha_lp.write(linha, 17, soe)  # escreve na coluna "SOE"
            planilha_lp.write(linha, 18, obs)  # escreve na coluna "OBSERVA€ŽO"
            planilha_lp.write(linha, 34, end)  # escreve na coluna "ENDERECO"
            linha += 1  # incrementa a linha

        # ********** Gravar Pontos Anal¢gicos Excel **********
        pana = []
        linha_rel = 1
        planilha_relatorio.write(0, 2, 'ID n„o encontrados em PAF')
        for id_tag in pas_dic:
            try:
                paf_id = paf_dic[id_tag]
            except:
                paf_id = ['']
                if 'CALC' not in pas_dic[id_tag][11] and 'LOCAL' not in pas_dic[id_tag][11]:
                    planilha_relatorio.write(linha_rel, 2, id_tag)
                    linha_rel += 1
            pana.append([id_tag] + pas_dic[id_tag] + paf_id)

        med_dic = {'FR': 'Hz', 'KV': 'kV', 'AM': 'A', 'DI': 'km', 'MV': 'MVAR', 'MW': 'MW', 'TM': 'ø C'}
        for dado in pana:  # Passa por todas as linhas do array de pontos anal¢gicos gravando pontos no Excel
            tac = dado[12]
            lsc = tac_dic.get(tac, ['?'])[0]
            id_protocolo = dado[13]
            tag = dado[0]
            ocr = dado[1]
            descr = dado[2]
            tipo = dado[3]
            medicao = med_dic.get(tipo[:2], '')
            anunciador = ('X' if tag in texto_telas else '')
            alarme = dado[4]
            obs = ''
            soe = 'X'

            id_pad = pad_dic.get(tag, ['', ''])
            end = paf_dic.get(id_pad[0], ['', ''])[0]

            liu = dado[5]
            lie = dado[6]
            lia = dado[7]
            lsa = dado[8]
            lse = dado[9]
            lsu = dado[10]
            bndmo = dado[11]

            planilha_lp.write(linha, 0, linha - 5)  # escreve na coluna "ITEM"
            planilha_lp.write(linha, 2, tac)  # escreve na coluna "TAC"
            planilha_lp.write(linha, 3, lsc)  # escreve na coluna "IED"
            planilha_lp.write(linha, 7, id_protocolo)  # escreve na coluna "ID PROTOCOLO"
            planilha_lp.write(linha, 9, tag)  # escreve na coluna "ID (SAGE)"
            planilha_lp.write(linha, 10, ocr)  # escreve na coluna "OCR"
            planilha_lp.write(linha, 11, descr)  # escreve na coluna "DESCRI€ŽO"
            planilha_lp.write(linha, 12, tipo)  # escreve na coluna "TIPO"
            planilha_lp.write(linha, 13, '')  # escreve na coluna "COMANDO"
            planilha_lp.write(linha, 14, medicao)  # escreve na coluna "MEDI€ŽO"
            planilha_lp.write(linha, 15, anunciador)  # escreve na coluna "ANUNCIADOR"
            planilha_lp.write(linha, 16, alarme)  # escreve na coluna "LISTA DE ALARMES"
            planilha_lp.write(linha, 17, soe)  # escreve na coluna "SOE"
            planilha_lp.write(linha, 18, obs)  # escreve na coluna "OBSERVA€ŽO"
            planilha_lp.write(linha, 34, end)  # escreve na coluna "ENDERECO"
            planilha_lp.write(linha, 38, liu)  # escreve na coluna "LIU"
            planilha_lp.write(linha, 39, lie)  # escreve na coluna "LIE"
            planilha_lp.write(linha, 40, lia)  # escreve na coluna "LIA"
            planilha_lp.write(linha, 41, lsa)  # escreve na coluna "LSA"
            planilha_lp.write(linha, 42, lse)  # escreve na coluna "LSE"
            planilha_lp.write(linha, 43, lsu)  # escreve na coluna "LSU"
            planilha_lp.write(linha, 44, bndmo)  # escreve na coluna "BNDMO"
            linha += 1  # incrementa a linha  

        # ********** Gravar Pontos Totalizadores Excel **********
        ptot = []
        linha_rel = 1
        planilha_relatorio.write(0, 4, 'ID n„o encontrados em PTF')
        try:
            for id_tag in pts_dic:
                try:
                    ptf_id = ptf_dic[id_tag]
                except:
                    ptf_id = ['']
                    if 'CALC' not in pts_dic[id_tag][7] and 'LOCAL' not in pts_dic[id_tag][7]:
                        planilha_relatorio.write(linha_rel, 4, id_tag)
                        linha_rel += 1
                ptot.append([id_tag] + pts_dic[id_tag] + ptf_id)
        except:
            mensagem_aviso('Aten‡„o', 'N„o foram gravados pontos totalizadores')
        for dado in ptot:  # Passa por todas as linhas do array de pontos anal¢gicos gravando pontos no Excel
            tac = dado[8]
            lsc = tac_dic.get(tac, ['?'])[0]
            id_protocolo = dado[9]
            tag = dado[0]
            ocr = dado[1]
            descr = dado[2]
            tipo = dado[3]
            anunciador = ('X' if tag in texto_telas else '')
            alarme = dado[4]
            obs = ''
            soe = 'X'
            if ptd_dic != '':
                id_ptd = ptd_dic.get(tag, ['', ''])
                end = ptf_dic.get(id_ptd[0], ['', ''])[0]
            else:
                id_ptd = ''
                end = ''
            lsa = dado[5]
            lse = dado[6]
            lsu = dado[7]

            planilha_lp.write(linha, 0, linha - 5)  # escreve na coluna "ITEM"
            planilha_lp.write(linha, 2, tac)  # escreve na coluna "TAC"
            planilha_lp.write(linha, 3, lsc)  # escreve na coluna "IED"
            planilha_lp.write(linha, 7, id_protocolo)  # escreve na coluna "ID PROTOCOLO"
            planilha_lp.write(linha, 9, tag)  # escreve na coluna "ID (SAGE)"
            planilha_lp.write(linha, 10, ocr)  # escreve na coluna "OCR"
            planilha_lp.write(linha, 11, descr)  # escreve na coluna "DESCRI€ŽO"
            planilha_lp.write(linha, 12, tipo)  # escreve na coluna "TIPO"
            planilha_lp.write(linha, 13, '')  # escreve na coluna "COMANDO"
            planilha_lp.write(linha, 15, anunciador)  # escreve na coluna "ANUNCIADOR"
            planilha_lp.write(linha, 16, alarme)  # escreve na coluna "LISTA DE ALARMES"
            planilha_lp.write(linha, 17, soe)  # escreve na coluna "SOE"
            planilha_lp.write(linha, 18, obs)  # escreve na coluna "OBSERVA€ŽO"
            planilha_lp.write(linha, 34, end)  # escreve na coluna "ENDERECO"
            planilha_lp.write(linha, 41, lsa)  # escreve na coluna "LSA"
            planilha_lp.write(linha, 42, lse)  # escreve na coluna "LSE"
            planilha_lp.write(linha, 43, lsu)  # escreve na coluna "LSU"
            linha += 1  # incrementa a linha

        # ********** Gravar Pontos Comandos Excel **********
        pcmd = []
        linha_rel = 1
        planilha_relatorio.write(0, 6, 'ID n„o encontrados em CGF')
        for id_tag in cgs_dic:
            try:
                cgf_id = [cgf_dic[id_tag][0]]
            except:
                cgf_id = ['']
                planilha_relatorio.write(linha_rel, 6, id_tag)
                linha_rel += 1

            try:
                cgf_com = cgf_dic[id_tag][1]
            except:
                cgf_com = ''
            if 'CSIM' in cgf_com:
                cgf_com = ['CS']
            elif 'CDUP' in cgf_com:
                cgf_com = ['CD']
            else:
                cgf_com = ['CD']

            pcmd.append([id_tag] + cgs_dic[id_tag] + cgf_id + cgf_com)

        for dado in pcmd:  # Passa por todas as linhas do array de pontos anal¢gicos gravando pontos no Excel
            tac = dado[3]
            lsc = tac_dic.get(tac, ['?'])[0]
            id_protocolo = dado[5]
            tag = dado[0]
            descr = dado[1]
            pac = dado[2]
            tipo = dado[4]
            comando = dado[6]

            if (tag == 'LOCAL') or ('COR' not in tag) or (len(tag) == len(pac)):
                planilha_lp.write(linha, 0, linha - 5)  # escreve na coluna "ITEM"
                planilha_lp.write(linha, 2, tac)  # escreve na coluna "TAC"
                planilha_lp.write(linha, 3, lsc)  # escreve na coluna "IED"
                planilha_lp.write(linha, 7, id_protocolo)  # escreve na coluna "ID PROTOCOLO"
                planilha_lp.write(linha, 9, tag)  # escreve na coluna "ID (SAGE)"
                planilha_lp.write(linha, 11, descr)  # escreve na coluna "DESCRI€ŽO"
                planilha_lp.write(linha, 12, tipo)  # escreve na coluna "TIPO"
                planilha_lp.write(linha, 13, comando)  # escreve na coluna "COMANDO"
                linha += 1  # incrementa a linha        

        arq_lp.close()

        abrirarquivo = pergunta_sim_nao('Aviso', 'Arquivo \"' + nome_arq_saida[
                                                        2:] + '\" gerado em ' + getcwd() + '\n\n Deseja abrir o arquivo gerado agora?')
        if abrirarquivo: startfile(getcwd() + '\\' + nome_arq_saida[2:])


if __name__ == "__main__":
    from tkinter.filedialog import askdirectory

    diretorio = askdirectory(title='Selecione o diret¢rio que est„o os arquivos .dat')
    if diretorio:
        base2lp(diretorio)
