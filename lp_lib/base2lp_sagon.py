# -*- coding: cp860 -*-
from FASgtkui import mensagem_erro, dialogo_abrir_arquivo_gerado, mensagem_erro_dialog
import sagon.sagist as sg
import sagon.xlsage as xs
import sagon.datapi as dt
import optparse
import os
import gi
gi.require_version("Gtk", "3.0")
from gi.repository import GObject, Gtk
import time
GObject.threads_init()

dados = '''
Vers„o 2.0.12
Atualiza‡„o do programa: 10/11/2020
Gera‡„o de LP no padr„o Chesf tendo como base os arquivos .dat de configura‡„o de base SAGE
'''

try:
    from lp_lib.gerarPlanilhaLP import gerarPlanilha
except:
    mensagem_erro('Erro', 'Arquivo "gerarPlanilhaLP.py" deve estar no diret¢rio "lp_lib"')
try:
    import xlsxwriter
except:
    mensagem_erro('Erro', 'M¢dulo XlsxWriter n„o instalado')

def base2xls(base_path='', Diretorio_Padrao = '', **kwargs):

    label_progressbar: Gtk.Label = kwargs.get('label_progressbar',None)
    progress_bar: Gtk.ProgressBar = kwargs.get('progress_bar', None)
    janela_carregando: Gtk.Window = kwargs.get('janela_carregando', None)
    # ***** Captar texto de arquivos de telas para verificar preenchimento da coluna anunciador *****
    telas = kwargs.get('telas', None)
    include_indice = kwargs.get('include_indice', None)
    parcela_rfc = kwargs.get('parcela_rfc')
    parcela_rca = kwargs.get('parcela_rca')
    texto_telas = ''
    for tela in telas:
        try:
            arq_txt = open('{}\\telas\\{}'.format(base_path.rsplit('\\', include_indice)[0], tela), 'r')
            texto_telas += arq_txt.read()
            arq_txt.close()
        except:
            GObject.idle_add(mensagem_erro,'Erro','O arquivo {}\\{} n„o pode ser carregado'.format(base_path.rsplit('\\', 2)[0], tela))
            time.sleep(1)
            while mensagem_erro_dialog.get_visible() == True:
                time.sleep(1)
    #dt.print_msg(__name__, 'tudo OK',dt.MSG_INFO, **kwargs)
    med_dic = {'FR': 'Hz', 'KV': 'kV', 'AM': 'A', 'DI': 'km', 'MV': 'MVAR', 'MW': 'MW', 'TM': 'ø C'}
    include_cmts = False
    label_progressbar.set_text('Carregando Base')
    base = sg.load_base(source_path=base_path, **kwargs)
    label_progressbar.set_text('Carregando dicion rios')
    ids_pdd = sg.create_dict_pdd(base['pdd'])

    ids_pad = sg.create_dict_pad(base['pad'])
    ids_ptd = sg.create_dict_ptd(base['ptd'])
    ids_pdf = sg.create_dict_xxf(base['pdf'])
    ids_cgf = sg.create_dict_cgf(base['cgf'])

    ids_paf = sg.create_dict_xxf(base['paf'])
    arquivos_fisicos ={'pdf':ids_pdf,'cgf':ids_cgf,'paf':ids_paf}
    if base.get('ptf'):
        idf_ptf = sg.create_dict_xxf(base['ptf'])
        arquivos_fisicos['ptf'] = idf_ptf
    saida_array = []
    def grava_ponto(CONTEMPLADO ='', TIPO_RELE='', ID_PROTOCOLO='',ID_SAGE='',
                    OCR_SAGE='',DESCRICAO ='',TIPO='',COMANDO='',MEDICAO='',TELA='',LISTA_DE_ALARMES='',
                    SOE = '', OBSERVACAO='', ENDERECO='', LIU = '', LIE = '',LIA = '',LSA='',LSE='',LSU='',BNDMO=''):
        saida_array.append([CONTEMPLADO, TIPO_RELE, ID_PROTOCOLO,ID_SAGE,
                    OCR_SAGE, DESCRICAO, TIPO, COMANDO, MEDICAO , TELA,LISTA_DE_ALARMES,
                    SOE, OBSERVACAO, ENDERECO, LIU, LIE, LIA, LSA, LSE,LSU,BNDMO])

    #Loop principal
    for dat_type in ['pds','pas','pts','cgs']:
        if dat_type == 'pds' or dat_type == 'cgs':
            xxd = ids_pdd
        elif dat_type == 'pas':
            xxd = ids_pad
        elif dat_type == 'pts':
            xxd = ids_ptd
        dat = base[dat_type]
        GObject.idle_add(janela_carregando.set_title, 'Lendo {0}...'.format(dat_type.upper()))
        key_number = 1
        total_keys = len(dat.keys())
        for key in dt.list_keys(dat):
            GObject.idle_add(label_progressbar.set_text,'Processando arquivo ({0} de {1}): {2}'.format(key_number, total_keys, key))
            interation = 0
            total = len(dat[key])
            if total == 0:
                total = 1
            GObject.idle_add(atualiza_progresso, progress_bar, interation, total) #Esse m‚todo direciona a execu‡„o da fun‡„o para a thread encarregada da interface gr fica
            #print(base['pdd'])
            for dat_item in dat[key]:

                CONTEMPLADO = TIPO_RELE = ID_PROTOCOLO = ID_SAGE = OCR_SAGE = DESCRICAO = TIPO = COMANDO = MEDICAO = TELA = LISTA_DE_ALARMES = SOE = OBSERVACAO = ENDERECO = LIU = LIE = LIA = LSA = LSE = LSU = BNDMO = ''
                # loop de itera‡„o sobre os itens
                if (not dt.is_comment(dat_item)):
                    if dt.is_commented_point(dat_item) and include_cmts:
                        dat_item = dt.clean_commented_point(dat_item)
                    if dat_item.get('ID', '') == '':
                        # ponto mal formado (sem id)
                        interation += 1
                        GObject.idle_add(atualiza_progresso, progress_bar, interation, total)
                        continue

                    dat_typef = dat_type[:2] + 'f'
                    dat_conf = sg.get_aconf_from_base(dat_type, item_id=dat_item.get('ID', ''), base_item=base, sitem = dat_item,
                                                      slocation = key,xxf=arquivos_fisicos[dat_typef], **kwargs)


                    # checa se ‚ roteamento de controle e pula se for o caso
                    if (dat_type == 'cgs') and dat_conf.get('cgf'):
                        if (len(str(dat_conf.get('cgf').get('items')[0].get('ID')).split('-')) == 3):
                            interation += 1
                            GObject.idle_add(atualiza_progresso, progress_bar, interation, total)
                            continue


                    ID_SAGE = dat_item.get('ID')
                    TELA = ('X' if 'WHERE id = ' + ID_SAGE in texto_telas else '')
                    DESCRICAO = dat_item.get('NOME')
                    CONTEMPLADO = dat_item.get('TAC')
                    tempo_antes_endn3 = time.time()
                    ENDERECO = sg.get_endN3_dist(dat_type, ID_SAGE, base=base, xxd=xxd, xxf =arquivos_fisicos[dat_typef], **kwargs)
                    tempo_depois_endn3 = time.time()
                    OBSERVACAO = dat_item.get('OBSRV')

                    # extrai metacampos de cmt
                    if '|' in dat_item.get('CMT', ''):
                        try:
                            testado, vao, ied, origem = str(dat_item.get('CMT')).split('|')
                            TIPO_RELE = ied
                        except:
                            pass

                    if TIPO_RELE == '' and dat_conf.get('lsc'):
                        try:
                            TIPO_RELE = dat_conf.get('lsc').get('item').get('ID')
                        except:
                            pass

                    if dat_type != 'cgs':
                        OCR_SAGE = dat_item.get('OCR')
                        # ws['K' + str(row)].value = dat_item.get('OCR')
                        ALRIN = dat_item.get('ALRIN')
                        if str(ALRIN).upper() == 'NAO':
                            # escreve na coluna "LISTA DE ALARMES"
                            # ws['Q' + str(row)].value = 'X'
                            LISTA_DE_ALARMES = 'X'
                    if dat_type == 'pds':
                        # campos de pds
                        SOEIN = dat_item.get('SOEIN')

                        if str(SOEIN).upper() == 'NAO':
                            # escreve na coluna "SOE"
                            # ws['R' + str(row)].value = 'X'
                            SOE = 'X'
                        # escreve na coluna "TIPO"
                        # ws['M' + str(row)].value = dat_item.get('TIPO')
                        TIPO = dat_item.get('TIPO')
                    elif dat_type == 'pts':
                        LSA = dat_item.get('LSA')
                        LSE = dat_item.get('LSE')
                        LSU = dat_item.get('LSU')
                        # ws['AP' + str(row)].value = dat_item.get('LSA')  # escreve na coluna "LSA"
                        # ws['AQ' + str(row)].value = dat_item.get('LSE')  # escreve na coluna "LSE"
                        # ws['AR' + str(row)].value = dat_item.get('LSU')  # escreve na coluna "LSU"
                    elif dat_type == 'pas':
                        # campos de pas
                        tipo = dat_item.get('TIPO')
                        try:
                            medicao = med_dic.get(tipo[:2], '')
                        except:
                            medicao = ''
                        TIPO = dat_item.get('TIPO')
                        MEDICAO = medicao
                        LIU = dat_item.get('LIU')
                        LIE = dat_item.get('LIE')
                        LIA = dat_item.get('LIA')
                        LSA = dat_item.get('LSA')
                        LSE = dat_item.get('LSE')
                        LSU = dat_item.get('LSU')
                        BNDMO = dat_item.get('BNDMO')

                    elif dat_type == 'cgs':
                        # campos de cgs
                        # ws['M' + str(row)].value = dat_item.get('TIPOE')   # escreve na coluna "TIPO"
                        TIPO = dat_item.get('TIPOE')
                        PAC = dat_item.get('PAC')
                        if 'CSIM' in PAC:
                            COMANDO = 'CS'
                        elif 'CDUP' in PAC:
                            COMANDO = 'CD'
                        else:
                            COMANDO = 'CD'
                    # se for um filtro composto
                    tempo_antes_filtros = time.time()
                    if 'rfc' in list(dat_conf.keys()):
                        OBSERVACAO = 'RFC Parcelas: '
                        DESCRICAO = dat_item.get('NOME')
                        for rfc in dat_conf.get('rfc').get('items'):
                            if OBSERVACAO != 'RFC Parcelas: ':
                                OBSERVACAO = OBSERVACAO + ' ; '
                            ID_PROTOCOLO = rfc.get('PNT', '')
                            if parcela_rfc:
                                OBSERVACAO = OBSERVACAO + rfc.get('PARC','')
                            else:
                                OBSERVACAO = ''
                                break
                        grava_ponto(CONTEMPLADO, TIPO_RELE, ID_PROTOCOLO, ID_SAGE, OCR_SAGE, DESCRICAO, TIPO,
                                        COMANDO, MEDICAO, TELA, LISTA_DE_ALARMES,
                                        SOE, OBSERVACAO, ENDERECO, LIU, LIE, LIA, LSA, LSE, LSU, BNDMO)
                        interation += 1
                        GObject.idle_add(atualiza_progresso, progress_bar, interation, total)
                        continue
                            # row += 1
                        # row -= 1
                    # se for um filtro simples
                    if 'rfi' in list(dat_conf.keys()):
                        i = 0
                        ID_PROTOCOLO =''
                        for rfi in dat_conf.get('rfi').get('items'):
                            ID_PROTOCOLO = ID_PROTOCOLO + rfi.get('PNT', '') + ' ; '
                            DESCRICAO = dat_item.get('NOME')
                            if str(rfi.get('TIPOP')).upper() == 'PDF':
                                # preenche campos pdf da planilha FILTRO
                                ID_SAGE = dat_conf.get('pdf').get('items')[i].get('PNT')
                                TELA = ('X' if 'WHERE id = ' + ID_SAGE in texto_telas else '')

                            elif str(rfi.get('TIPOP')).upper() == 'PAF':
                                ID_SAGE = dat_conf.get('paf').get('items')[i].get('PNT')
                                TELA = ('X' if 'WHERE id = ' + ID_SAGE in texto_telas else '')

                            elif str(rfi.get('TIPOP')).upper() == 'PTF':
                                # preenche campos pdf da planilha FILS
                                ID_SAGE = dat_conf.get('ptf').get('items')[i].get('PNT')
                                TELA = ('X' if 'WHERE id = ' + ID_SAGE in texto_telas else '')

                            if '|' in dat_conf['rfi']['items'][i].get('CMT', ''):
                                try:
                                    testado, vao, ied, origem = str(dat_conf['rfi']['items'][i].get('CMT', '')).split(
                                        '|')
                                    # escreve na coluna "IED"
                                    # ws['D' + str(row)].value = ied
                                    TIPO_RELE = ied
                                except:
                                    pass
                            OBSERVACAO = 'FILTRO SIMPLES'
                            i += 1
                        grava_ponto(CONTEMPLADO, TIPO_RELE, ID_PROTOCOLO, ID_SAGE, OCR_SAGE, DESCRICAO, TIPO,
                                    COMANDO, MEDICAO, TELA, LISTA_DE_ALARMES,
                                    SOE, OBSERVACAO, ENDERECO, LIU, LIE, LIA, LSA, LSE, LSU, BNDMO)
                        interation += 1
                        GObject.idle_add(atualiza_progresso, progress_bar, interation, total)
                        continue
                    tempo_depois_filtros = time.time()
                    # caso seja um ponto f¡sico aquisitado, preencher conf f¡sica
                    dat_typef = dat_type[:2] + 'f'
                    tempo_antes_confisica = time.time()
                    if dat_typef in list(dat_conf.keys()):
                        # n„o for um filtro
                        if len(dat_conf[dat_typef]['items']) == 1:
                            #print('dat_type- {} dat_item- {} aconf- {}'.format(dat_type,dat_item['ID'],dat_conf))
                            if sg.is_61850(dat_type, item_id=dat_item['ID'], aconf=dat_conf):
                                ID_PROTOCOLO = xs.expand_address(dat_type=dat_typef, aconf=dat_conf)
                                if ID_PROTOCOLO == False:
                                    i=2
                                    while ID_PROTOCOLO == False or i<10:
                                        ID_alterado = dat_item.get('ID') + '_' + str(i)
                                        dat_conf = sg.get_aconf_from_base(dat_type, item_id=dat_item.get('ID', ''),
                                                                      base_item=base, sitem=dat_item,
                                                                      slocation=key, xxf=arquivos_fisicos[dat_typef],
                                                                    item_id_alterado=ID_alterado,
                                                                      **kwargs)
                                        ID_PROTOCOLO = xs.expand_address(dat_type=dat_typef, aconf=dat_conf)
                                        if ID_PROTOCOLO != False:
                                            break
                                        print(ID_alterado)
                                        print(dat_conf)
                                        i+=1
                                    if ID_PROTOCOLO == False:
                                        ID_PROTOCOLO=''
                                        print('nao achou')

                            else:
                                # ws['H' + str(row)].value = dat_conf[dat_typef]['items'][0].get('ID')  # escreve na coluna "ID PROTOCOLO"
                                ID_PROTOCOLO = dat_conf[dat_typef]['items'][0].get('ID')
                    tempo_depois_confisica = time.time()
                    # caso seja um ponto calculado, preencher planilha CALC
                    if 'rca' in list(dat_conf):
                        OBSERVACAO ='Parcelas: '
                        DESCRICAO = dat_item.get('NOME')  # escreve na coluna "DESCRI€ŽO"
                        TIPO_RELE = 'CALC'  # escreve na coluna "IED"
                        CONTEMPLADO = 'CALC'  # escreve na coluna "TAC"
                        for i in range(0, len(dat_conf['rca']['items'])):
                            ID_SAGE = dat_conf['rca']['items'][i].get('PNT')
                            TELA = ('X' if 'WHERE id = ' + ID_SAGE in texto_telas else '')
                            if parcela_rca:
                                OBSERVACAO = OBSERVACAO + str(dat_conf['rca']['items'][i].get('PARC')) + ' ; '
                            else:
                                OBSERVACAO = ''
                                break
                        grava_ponto(CONTEMPLADO=CONTEMPLADO, ID_SAGE=ID_SAGE, OBSERVACAO=OBSERVACAO,
                                        TIPO_RELE=TIPO_RELE, DESCRICAO=DESCRICAO, OCR_SAGE=OCR_SAGE, TELA=TELA)
                            # row +=1
                        interation += 1
                        GObject.idle_add(atualiza_progresso, progress_bar, interation, total)
                        continue
                        # row-=1
                    grava_ponto(CONTEMPLADO, TIPO_RELE, ID_PROTOCOLO, ID_SAGE, OCR_SAGE, DESCRICAO, TIPO,
                                COMANDO, MEDICAO, TELA, LISTA_DE_ALARMES,
                                SOE, OBSERVACAO, ENDERECO, LIU, LIE, LIA, LSA, LSE, LSU, BNDMO)
                    # row+=1
                interation += 1
                GObject.idle_add(atualiza_progresso, progress_bar, interation, total)


                # fim do loop de itera‡„o sobre os itens
            key_number += 1
    # FIM DA LEITURA DA PLANILHA
    nome_arq_saida = 'LP_da_Base'  # Nome do arquivo de sa¡da
    seq_arq = 0  # Sequˆncia do n£mero de arquivo
    while os.path.exists(
            Diretorio_Padrao + '\\' + nome_arq_saida + '.xlsx'):  # Enquanto existir na pasta um arquivo com o nome definido
        seq_arq += 1  # Adicionar um a sequˆncia do n£mero do arquivo
        nome_arq_saida = 'LP_da_Base_' + str(seq_arq)  # Definir novo nome de arquivo
    nome_arq_saida = Diretorio_Padrao + '\\' + nome_arq_saida + '.xlsx'
    arq_lp = gerarPlanilha(nome_arq_saida)  # Gera um arquivo Excel com uma planilha com formata‡„o da LP Padr„o
    planilha_lp = arq_lp.worksheets()[0]
    linha =6
    for dado in saida_array:

        planilha_lp.write(linha, 0, linha - 5)  # escreve na coluna "ITEM"
        planilha_lp.write(linha, 2, dado[0])  # escreve na coluna "TAC"
        planilha_lp.write(linha, 3, dado[1])  # escreve na coluna "IED"
        planilha_lp.write(linha, 7, dado[2])  # escreve na coluna "ID PROTOCOLO"
        planilha_lp.write(linha, 9, dado[3])  # escreve na coluna "ID (SAGE)"
        planilha_lp.write(linha, 10, dado[4])  # escreve na coluna "OCR"
        planilha_lp.write(linha, 11, dado[5])  # escreve na coluna "DESCRI€ŽO"
        planilha_lp.write(linha, 12, dado[6])  # escreve na coluna "TIPO"
        planilha_lp.write(linha, 13, dado[7])  # escreve na coluna "COMANDO"
        planilha_lp.write(linha, 14, dado[8])  # escreve na coluna "MEDI€ŽO"
        planilha_lp.write(linha, 15, dado[9])  # escreve na coluna "ANUNCIADOR"
        planilha_lp.write(linha, 16, dado[10])  # escreve na coluna "LISTA DE ALARMES"
        planilha_lp.write(linha, 17, dado[11])  # escreve na coluna "SOE"
        planilha_lp.write(linha, 18, dado[12])  # escreve na coluna "OBSERVA€ŽO"
        planilha_lp.write(linha, 34, dado[13])  # escreve na coluna "ENDERECO"
        planilha_lp.write(linha, 38, dado[14])  # escreve na coluna "LIU"
        planilha_lp.write(linha, 39, dado[15])  # escreve na coluna "LIE"
        planilha_lp.write(linha, 40, dado[16])  # escreve na coluna "LIA"
        planilha_lp.write(linha, 41, dado[17])  # escreve na coluna "LSA"
        planilha_lp.write(linha, 42, dado[18])  # escreve na coluna "LSE"
        planilha_lp.write(linha, 43, dado[19])  # escreve na coluna "LSU"
        planilha_lp.write(linha, 44, dado[20])  # escreve na coluna "BNDMO"
        linha += 1  # incrementa a linha
    arq_lp.close()
    janela_carregando.hide()
    GObject.idle_add(atualiza_progresso, progress_bar, 0, total)
    GObject.idle_add(janela_carregando.set_title, 'Processando...')
    GObject.idle_add(dialogo_abrir_arquivo_gerado,nome_arq_saida.rsplit('\\', 1)[1],Diretorio_Padrao)

    #abrirarquivo = pergunta_sim_nao('Aviso', 'Arquivo \"' + nome_arq_saida.rsplit('\\', 1)[
    #    1] + '\" gerado em ' + Diretorio_Padrao + '\n\n Deseja abrir o arquivo gerado agora?')
    #if abrirarquivo: os.startfile(nome_arq_saida)
def atualiza_progresso(progress_bar, interation, total):
    progress_bar.set_fraction(interation / float(total))
    progress_bar.set_text('{:.2f} % Completo'.format(100 * interation / float(total)))
    progress_bar.set_pulse_step(interation / float(total))


def main():
    parser = optparse.OptionParser()
    parser.add_option('-f','--file', dest='filename', default='config.xlsx', help='Nome do xls de destino', metavar='FILE')
    parser.add_option('-q','--quiet', dest='verbose', default=True, action='store_false',
                      help='n„o imprime mensagens de progresso do script')
    parser.add_option('-i','--ignore_cmts', action='store_false', dest='include_cmts',
                      default=True, help='ignora linhas comentadas da base')
    parser.add_option('-m','--model_file', default='modelo.xlsx', dest='model_file',
                      help='arquivo modelo xls')
    (options, args) = parser.parse_args()
    print('options:', str(options))
    print('arguments:',args)
    if len(args) !=1:
        base_path=''
    else:
        base_path=str(args[0])

    base2xls(base_path=base_path, filename=options.filename, model_wb=options.model_file, include_cmts=options.include_cmts, verbose=options.verbose)

if __name__ == '__main__':
    main()
else:
    print('base2xls carregado como m¢dulo')
