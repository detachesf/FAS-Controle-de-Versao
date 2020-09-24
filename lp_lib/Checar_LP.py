# -*- coding: cp860 -*-

dados = '''
Vers„o 2.0.9
Atualiza‡„o do programa: 18/04/2016
Funcionalidade de checagem entre LPs ou de LP com LP gerada por arquivo LP_Config. 
'''

from tkinter.messagebox import showerror, showwarning, askyesno
import pickle
import os.path
from os import startfile
from tkinter import END
from operator import itemgetter
from difflib import get_close_matches

try:
    from xlrd import open_workbook
except:
    showerror('Erro', 'M¢dulo xlrd n„o instalado')

try:
    from lp_lib.LP import gerarlp
except:
    showerror('Erro', 'Arquivo "LP.py" deve estar no mesmo diret¢rio "lp_lib"')
try:
    import xlsxwriter
except:
    showerror('Erro', 'M¢dulo XlsxWriter n„o instalado')
try:
    from lp_lib.func import linhaInicialETitulos
except:
    showerror('Erro', 'Arquivo "func.pyc" deve estar no diret¢rio "lp_lib"')


def checar(LP_Padrao='', LP_Editado='', planilha='', relatorio='', LP_Config='',
           array_base=''):  # array_base s¢ ser  preenchido em caso de compara‡„o

    """

    :type array_base: object
    """
    gerararquivo = True

    # ----------Declara‡„o de Vari veis----------#
    array_padrao = []
    array_validar = []
    diferenca_array = []
    pfalta_array = []
    array_validar_endereco = []
    endduplicado_array = []
    sugestao_ID_array = []
    k_inc = 0
    k_falta = 0
    k_enddupl = 0

    # ----------Ler LP Validar----------#

    LP_Validar = LP_Editado  # Ler defini‡„o do arquivo de LP padr„o
    Nome_Planilha = planilha

    try:
        book = open_workbook(LP_Validar)  # Abrir arquivo de LP a ser validada
    except:
        showerror('Erro', 'Arquivo ' + LP_Validar + ' n„o encontrado')

    sheet = book.sheet_by_name(Nome_Planilha)  # Abrir planilhas
    try:
        # Lˆ planilha e recebe a linha onde come‡a a LP (aqui usando linha inicial e n„o o dicion rio de t¡tulos)
        li, titulo_dic = linhaInicialETitulos(LP_Validar, Nome_Planilha)
        if li < 0:  # Se for um n£mero negativo ent„o n„o foi encontrado "ID (SAGE)" na lista
            raise NameError('Arquivo especificado n„o possui coluna com t¡tulo "ID (SAGE)".')

        for index_linha in range(li, sheet.nrows):  # Ler colulas da linha selecionada ao final
            if sheet.cell(index_linha, titulo_dic['ID (SAGE)']).value != '' and \
                            sheet.cell(index_linha, titulo_dic['ID (SAGE)']).value != 'CGS' and \
                            sheet.cell(index_linha, titulo_dic['ID (SAGE)']).value != 'PDS' and \
                            sheet.cell(index_linha, titulo_dic['ID (SAGE)']).value != 'PAS':
                try:  # Caso a descri‡„o do campo 6 seja "TELA"
                    # 0 - ID SAGE
                    array_validar.append([sheet.cell(index_linha, titulo_dic['ID (SAGE)']).value,
                                          # 1 - OCR
                                          sheet.cell(index_linha, titulo_dic['OCR (SAGE)']).value,
                                          # 2 - DESCRI€ŽO
                                          sheet.cell(index_linha, titulo_dic['DESCRI€ŽO']).value.strip(),
                                          # 3 - TIPO
                                          sheet.cell(index_linha, titulo_dic['TIPO']).value,
                                          # 4 - COMANDO
                                          sheet.cell(index_linha, titulo_dic['COMANDO']).value,
                                          # 5 - MEDI€ŽO
                                          sheet.cell(index_linha, titulo_dic['MEDI€ŽO']).value,
                                          # 6 - TELA
                                          sheet.cell(index_linha, titulo_dic['TELA']).value,
                                          # 7 - LISTA DE ALARMES
                                          sheet.cell(index_linha, titulo_dic['LISTA DE ALARMES']).value,
                                          # 8 - SOE
                                          sheet.cell(index_linha, titulo_dic['SOE']).value,
                                          # ENDERE€O N3
                                          sheet.cell(index_linha, 34).value])
                except:  # Caso a descri‡„o do campo 6 seja "ANUNCIADOR"
                    # 0 - ID SAGE
                    array_validar.append([sheet.cell(index_linha, titulo_dic['ID (SAGE)']).value,
                                          # 1 - OCR
                                          sheet.cell(index_linha, titulo_dic['OCR (SAGE)']).value,
                                          # 2 - DESCRI€ŽO
                                          sheet.cell(index_linha, titulo_dic['DESCRI€ŽO']).value.strip(),
                                          # 3 - TIPO
                                          sheet.cell(index_linha, titulo_dic['TIPO']).value,
                                          # 4 - COMANDO
                                          sheet.cell(index_linha, titulo_dic['COMANDO']).value,
                                          # 5 - MEDI€ŽO
                                          sheet.cell(index_linha, titulo_dic['MEDI€ŽO']).value,
                                          # 6 - TELA / ANUNCIADOR
                                          sheet.cell(index_linha, titulo_dic['ANUNCIADOR']).value,
                                          # 7 - LISTA DE ALARMES
                                          sheet.cell(index_linha, titulo_dic['LISTA DE ALARMES']).value,
                                          # 8 - SOE
                                          sheet.cell(index_linha, titulo_dic['SOE']).value,
                                          # ENDERE€O N3
                                          sheet.cell(index_linha, 34).value])

    except:
        showerror('Erro', 'O programa n„o reconhece o arquivo a ser checado como v lido')
        gerararquivo = False

    if array_base:
        array_padrao = array_base
    else:
        for pad in gerarlp(LP_Padrao, LP_Config)[0]:  # usar fun‡„o gerarlp para criar array_padrao
            # ID SAGE
            array_padrao.append([pad[0],
                                 # OCR
                                 pad[1].value,
                                 # DESCRI€ŽO
                                 pad[2],
                                 # TIPO
                                 pad[3].value,
                                 # COMANDO
                                 pad[4].value,
                                 # MEDI€ŽO
                                 pad[5].value,
                                 # ANUNCIADOR
                                 pad[6].value,
                                 # LISTA DE ALARMES
                                 pad[7].value,
                                 # SOE
                                 pad[8].value])

            # -----------------------------------------------------------------------------------------------------------------------------------------
    COD_SE = array_validar[0][0].split(':')[0]
    nome_arq_saida = './Relatorio_{}.xlsx'.format(COD_SE)  # Nome do arquivo de sa¡da
    seq_arq = 0  # Sequˆncia do n£mero de arquivo
    while os.path.exists(nome_arq_saida):  # Enquanto existir na pasta um arquivo com o nome definido
        seq_arq += 1  # Adicionar um a sequˆncia do n£mero do arquivo
        nome_arq_saida = '{}_{}_{}.xlsx'.format(nome_arq_saida[0:11], COD_SE, str(
            seq_arq))  # Definir novo nome de arquivo (Ex './LP_gerada.'+'_'+'1'+'.xlsx')
    arq_Relatorio = xlsxwriter.Workbook(nome_arq_saida[2:])

    ### Formata‡„o da c‚lula T¡tulo ###
    formatCelTitulo = arq_Relatorio.add_format({
        'bold': True,
        'font_name': 'Arial',
        'font_size': 9,
        'rotation': 90,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': 'silver',
    })

    ### Formata‡„o da c‚lula Errada###
    formatCelErro = arq_Relatorio.add_format({
        # 'bold': True,
        # 'font_name':'Arial',
        # 'font_size':12,
        'rotation': 0,
        'align': 'left',
        'valign': 'vcenter',
        'bg_color': 'red',
    })

    ### Formata‡„o da c‚lula Sugerida###
    formatCelSur = arq_Relatorio.add_format({
        # 'bold': True,
        # 'font_name':'Arial',
        # 'font_size':12,
        'rotation': 0,
        'align': 'left',
        'valign': 'vcenter',
        'bg_color': 'yellow',
    })

    # ----------Validar pontos da LP que est  sendo verificada----------#
    dic_padrao = {'{}_{}'.format(dic[0], dic[4].strip()): dic for dic in
                  array_padrao}  # ID_COMANDO : LINHA COMPLETA DE REGISTRO
    dic_validar = {'{}_{}'.format(dic[0], dic[4].strip()): dic[:-1] for dic in
                   array_validar}  # ID_COMANDO : LINHA DE REGISTRO SEM N3
    dic_faltando = {x: dic_padrao[x] for x in dic_padrao if x not in dic_validar}

    try:
        sugestao1_dic = {
            '{}_{}_{}_{}'.format(dic[0].split(':')[1], dic[0].split(':')[2], dic[2].strip(), dic[4].strip()): (
                dic[0], dic[4]) for dic in
            dic_faltando.values()}  # Chave: [VŽO]_[IED]_[DESCRI€ŽO]_[COMANDO], Valor: [ID SAGE]
        sugestao2_dic = {'{}_{}_{}'.format(dic[0].split(':')[1], dic[2].strip(), dic[4].strip()): (dic[0], dic[4]) for
                         dic in dic_faltando.values()}  # Chave: [VŽO]_[DESCRI€ŽO]_[COMANDO], Valor: [ID SAGE]
    except:
        showwarning('Impossibilidade de Sugest„o de ID',
                    'N„o ser  poss¡vel realizar sugest„o de ponto.\nProvavelmente existem ID de pontos fora do padr„o')

    # sugestao1_dic = {'{}_{}_{}_{}'.format(dic[0].split(':')[1], dic[0].split(':')[2], dic[2].strip(),dic[4].strip()) : (dic[0],dic[4]) for dic in array_padrao} # Chave: [VŽO]_[IED]_[DESCRI€ŽO]_[COMANDO], Valor: [ID SAGE]
    # sugestao2_dic = {'{}_{}_{}'.format(dic[0].split(':')[1], dic[2].strip(),dic[4].strip()) : (dic[0],dic[4]) for dic in array_padrao} # Chave: [VŽO]_[DESCRI€ŽO]_[COMANDO], Valor: [ID SAGE]
    # array_validar_ID_COM = [(col[0],col[4]) for col in array_validar]

    array_validar_semN3 = [arr[:-1] for arr in array_validar]
    for validar in array_validar_semN3:
        if validar not in array_padrao:
            diferenca_array.append(validar)
            k_inc += 1

    array_padrao_ID_COM = [(col[0], col[4]) for col in array_padrao]

    for diferenca in diferenca_array:
        try:
            posicao = array_padrao_ID_COM.index((diferenca[0], diferenca[4]))
            campos_corretos = []
            for i in range(9):
                if array_padrao[posicao][i] != diferenca[i]:
                    diferenca[i] = '*' + diferenca[i]
                    if array_padrao[posicao][i].strip() != 'X' or array_padrao[posicao][i].strip() != '':
                        campos_corretos.append(array_padrao[posicao][i])
            diferenca.append(' <<>> '.join(campos_corretos))
        except:  # entra aqui se "array_padrao_ID.index(diferenca[0])" levantar exce‡„o por n„o conter "diferenca[0]" no array "array_padrao_ID"
            # Sugerir ID baseado no equipamento e descri‡„o do ponto
            try:
                vao_dif = diferenca[0].split(':')[1]  # V„o/Equipamento do ponto que n„o foi achado ID(SAGE)
                ied_dif = diferenca[0].split(':')[2]  # IED do ponto que n„o foi achado ID(SAGE)
                dsc_dif = diferenca[2].strip()  # Descri‡„o do ponto que n„o foi achado ID(SAGE)
                cmd_dif = diferenca[4].strip()  # Campo Comando do ponto que n„o foi achado ID(SAGE)

                # Tentar sugest„o usando VŽO_EQUIP_DESC_COMANDO, se n„o conseguir tentar com VŽO_DESC_COMANDO
                sugestao_ID = sugestao1_dic.get('{}_{}_{}_{}'.format(vao_dif, ied_dif, dsc_dif, cmd_dif),
                                                sugestao2_dic.get('{}_{}_{}'.format(vao_dif, dsc_dif, cmd_dif), ''))

                # Se n„o conseguiu sugest„o_ID ainda, tentar por similaridade da descri‡„o nos pontos faltantes
                if not sugestao_ID:
                    dic_vao = {}
                    for reg in dic_faltando.values():  # Passar todos os registros faltantes
                        vao = reg[0].split(':')[1]  # V„o/Equipamento
                        if vao not in dic_vao:  # Se ainda n„o existir o dicion rio do V„o/Equipamento
                            dic_vao[vao] = {reg[2]: (
                                reg[0], reg[4])}  # Criar dicion rio do V„o/Equipamento com Descri‡„o como chave
                        else:  # Se existir o dicion rio do V„o/Equipamento
                            dic_vao[vao][reg[2]] = (
                                reg[0], reg[4])  # Gravar mais um registro no dicion rio do V„o/Equipamento

                    # Procura descri‡„o semelhante no dic_vao nos registros faltantes do vao_dif
                    dsc_match = get_close_matches(dsc_dif, dic_vao[vao_dif])[0]
                    # Procura ID dic_vao nos registros faltantes do vao_dif de acrodo com dsc_match
                    sugestao_ID = dic_vao[vao_dif][dsc_match]

                sugestao_ID_array.append(sugestao_ID)
            except:
                sugestao_ID = ''

            diferenca[0] = '*' + diferenca[0]  # Marcar ID como n„o encontrado
            if diferenca[4] not in ['', 'CS', 'CD', 'SP']:
                diferenca[4] = '*' + diferenca[4]  # Marcar Comando inv lido
            if sugestao_ID:
                diferenca.append('Poss¡vel ID -> {}'.format(sugestao_ID[0]))
            else:
                diferenca.append('')

    planilha_problema = arq_Relatorio.add_worksheet('Problema')  # Criar Planilha "Problema"

    largura = [22, 18, 40, 8, 5, 5, 5, 5, 5, 50]
    for i in range(0, 10):  # Ajuste da largura das colunas
        planilha_problema.set_column(i, i, largura[i])

    array_titulo = ['ID (SAGE)',
                    'OCR (SAGE)',
                    'DESCRI€ŽO',
                    'TIPO',
                    'COMANDO',
                    'MEDI€ŽO',
                    'ANUNCIADOR',
                    'LISTA DE ALARMES',
                    'SOE',
                    'OBSERVA€™ES']

    for titulo in array_titulo:  # Gravar t¡tulo
        planilha_problema.write(0, array_titulo.index(titulo), titulo, formatCelTitulo)

    linha = 1
    msgerroNumero = False
    for dado in diferenca_array:  # Passa por todas as linhas do array de sa¡da
        for i in range(10):
            try:
                if dado[i].startswith('*'):  # testa se o campo est  marcado como "incoerente"
                    planilha_problema.write(linha, i, dado[i][1:],
                                            formatCelErro)  # se est  "incoerente" grava na planilha usando uma formata‡„o diferente
                else:
                    planilha_problema.write(linha, i, dado[
                        i])  # se est  "incoerente" grava na planilha usando uma formata‡„o default
            except:
                msgerroNumero = True
        linha += 1

    if msgerroNumero:
        gerararquivo = False
        showerror('Erro',
                  'Verifique preenchimento de campos no Arquivo LP a ser checado. Nem um dos campos deve ser preenchido apenas com n£meros')

    # ----------Pontos Faltantes----------#
    for pfaltando in sorted(dic_faltando.items(), key=itemgetter(0)):
        pfalta_array.append(pfaltando[1])
        k_falta += 1

    planilha_Pfaltantes = arq_Relatorio.add_worksheet('Pontos_faltantes')  # Criar Planilha "Pontos_faltantes"

    for i in range(0, 10):  # Ajuste da largura das colunas
        planilha_Pfaltantes.set_column(i, i, largura[i])

    for titulo in array_titulo:  # Gravar t¡tulo
        planilha_Pfaltantes.write(0, array_titulo.index(titulo), titulo, formatCelTitulo)

    linha = 1
    for dado in pfalta_array:  # Passa por todas as linhas do array de sa¡da
        for i in range(9):
            if (dado[i], dado[4]) in sugestao_ID_array:  # testa se o campo est  entre IDs sugeridos
                planilha_Pfaltantes.write(linha, i, dado[i], formatCelSur)
            else:
                planilha_Pfaltantes.write(linha, i, dado[i])
        linha += 1

    # ----------Verificar Endere‡o N3 da LP padr„o que n„o est„o na LP que est  sendo verificada----------#
    for endereco in array_validar_endereco:
        if array_validar_endereco.count(endereco) > 1:
            if endereco not in endduplicado_array:
                endduplicado_array.append(endereco)
                k_enddupl += array_validar_endereco.count(endereco)

    planilha_EndDupl = arq_Relatorio.add_worksheet('End. N3 duplicados')  # Criar Planilha "End. N3 duplicados"

    array_titulo = ['Endere‡o',
                    'Ocorrˆncia']

    coluna = 0
    for titulo in array_titulo:  # Gravar t¡tulo
        planilha_EndDupl.write(0, coluna, titulo)
        coluna += 1

    linha = 1
    for dado in endduplicado_array:  # Passa por todas as linhas do array de sa¡da
        planilha_EndDupl.write(linha, 0, str(dado)[1:-3])
        planilha_EndDupl.write(linha, 1, str(array_validar_endereco.count(dado)))
        linha += 1

    # ----------Planilha Resuno----------#
    planilha_resumo = arq_Relatorio.add_worksheet('Resumo')  # Criar Planilha "Resumo"

    texto_resumo = ['-----Pontos com problemas-----',
                    '',
                    'Quantidade: {} pontos'.format(k_inc),
                    'Percentual: {:2.2f}%'.format(float(len(diferenca_array)) * 100 / len(array_validar)),
                    '',
                    '-----Pontos faltantes-----',
                    '',
                    'Quantidade: {} pontos'.format(k_falta),
                    'Percentual: {:2.2f}%'.format(float(len(pfalta_array)) * 100 / len(array_validar)),
                    '',
                    '-----Endere‡o para N3 Duplicado-----',
                    '',
                    'Quantidade: {} pontos'.format(k_enddupl),
                    'Percentual: {:2.2f}%'.format(k_enddupl * 100 / len(array_validar))]

    planilha_resumo.set_column(0, 0, 35)
    for linha, texto in enumerate(texto_resumo):
        planilha_resumo.write(linha, 0, texto)

    # ----------Gravar arquivo Excel----------#
    if gerararquivo:

        arq_Relatorio.close()
        try:
            for texto in texto_resumo:
                relatorio.insert(END, texto)
        except:
            pass

        abrirarquivo = askyesno('Aviso', 'Arquivo \"' + nome_arq_saida[
                                                        2:] + '\" gerado em ' + os.getcwd() + '\n\n Deseja abrir o arquivo gerado agora?')
        if abrirarquivo: startfile(os.getcwd() + '\\' + nome_arq_saida[2:])

        nomearquivo = nome_arq_saida[2:]

        conf = {'arquivo': nomearquivo}
        pickle.dump(conf, open('fas.p', 'wb'), -1)  # -1 para gravar em Bin rio
