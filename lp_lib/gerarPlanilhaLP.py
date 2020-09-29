# -*- coding: cp860 -*-

from tkinter.messagebox import showerror

try:
    import xlsxwriter
except:
    showerror('Erro','M¢dulo XlsxWriter nÑo instalado')


dados= '''                  
VersÑo do programa: 2.0.5   
AtualizaáÑo do programa: 04/12/2014
GeraáÑo de planilha formatada segundo definido em LP padrÑo.
'''

def gerarPlanilha(nome_arquivo):
    
    # Criaáao do arquivo ------------------------------------------------------ 
    arquivo = xlsxwriter.Workbook(nome_arquivo)
    planilha = arquivo.add_worksheet('PADRéO')
    planilha.set_zoom(80)       
   
    # Ajuste da largura das colunas -----------------------------------------------
    largura = [5, 30, 8, 9, 9, 9, 9, 9, 9, 
               30, 18, 35, 8, 5, 5, 4, 4, 4, 32, 32, 
                   18, 5, 5, 4, 4, 4, 4, 4, 
                   18, 5, 5, 4, 4, 33, 9, 24, 
                   9, 30, 9, 9, 9, 9, 9, 9, 9, 9, 30]

    for i in range (0,47):
        planilha.set_column(i, i,largura[i])

    # Oculta linhas iniciais
    planilha.set_row(0, options={'hidden': True})
    planilha.set_row(1, options={'hidden': True})
    planilha.set_row(2, options={'hidden': True})
    planilha.set_row(5, 115)    
    planilha.freeze_panes(6,0)
    planilha
    
    # Campo ITEM ------------------------------------------------------------------
    formato1 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 90,
    'align':'center',
    'valign':'vcenter',
    'bg_color': 'silver',
    'border' : 1,
    })
    planilha.merge_range('A4:A6', 'ITEM', formato1)
    
    # Campo PROJETO ---------------------------------------------------------------
    formato2 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': '#87CEEB',
    'border' : 1,
    })
    planilha.merge_range('B4:C4', 'PROJETO', formato2)
    
    col =1 
    
    formato3 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 90,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#87CEEB',
    'border' : 1,
    })
    for i in ('COMENTÜRIO','CONTEMPLADO'):
            planilha.merge_range(4,col,5,col,i,formato3)
            col+=1
    
    # Campo CHESF - NãVEL 1 -------------------------------------------------------
    formato4 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': '#FFFF99',
    'border' : 1,
    })
    planilha.merge_range('D4:I4', 'CHESF - NãVEL 1', formato4)
    
    formato5 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 90,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#FFFF99',
    'border' : 1,
    })
    for i in ('TIPO DO RELê','UA - PAINEL','BI','BO','ID PROTOCOLO','UTIL'):
        planilha.merge_range(4,col,5,col,i, formato5)
        col+=1
    
    # Campo CHESF - NãVEL 2 -------------------------------------------------------
    formato6 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': '#7DF9FF',
    'border' : 1,
    })
    planilha.merge_range('J4:T4', 'CHESF - NãVEL 2', formato6)
    
    formato7 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 90,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#7DF9FF',
    'border' : 1,
    })
    for i in ('ID (SAGE)','OCR (SAGE)','DESCRIÄéO','TIPO','COMANDO','MEDIÄéO','TELA','LISTA DE ALARMES','SOE','OBSERVAÄôES','AGRUPAMENTO'):
        planilha.merge_range(4,col,5,col,i, formato7)
        col+=1
    
    # Campo CHESF - TELEASSISTâNCIA N3 --------------------------------------------
    formato8 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': '#32B141',
    'border' : 1,
    })
    planilha.merge_range('U4:AB4', 'CHESF - TELEASSISTâNCIA N3', formato8)
    
    formato9 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 90,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#32B141',
    'border' : 1,
    })
    for i in ('OCR (SAGE)','COMANDO','MEDIÄéO','LISTA DE ALARME','SOE','OBSERVAÄéO','ENDEREÄO','AGRUPAMENTO'):
        planilha.merge_range(4,col,5,col,i, formato9)
        col+=1
        
    # Campo CHESF - NãVEL 3 ------------------------------------------------------
    formato10 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': 'FFFF99',
    'border' : 1,
    })
    planilha.merge_range('AC4:AJ4', 'CHESF - NãVEL 3', formato10)
    
    formato11 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 90,
    'align':'center',
    'valign':'vcenter',
    'bg_color': 'FFFF99',
    'border' : 1,
    })
    for i in ('OCR (SAGE)','COMANDO','MEDIÄéO','LISTA DE ALARME','SOE','OBSERVAÄO','ENDEREÄO','AGRUPAMENTO'):
        planilha.merge_range(4,col,5,col,i, formato11)
        col+=1
    # Campo ONS ------------------------------------------------------------------
    formato12 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': '7DF9FF',
    'border' : 1,
    })
    planilha.merge_range('AK4:AL4', 'ONS', formato12)
    planilha.merge_range('AK5:AL5', 'PROC DE REDE', formato12)
    
    formato13 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 90,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '7DF9FF',
    'border' : 1,
    })
    for i in ('ITEM','DESCRIÄéO'):
        planilha.write(5,col,i, formato13)
        col+=1
    # Campo LIMITES OPERACIONAIS --------------------------------------------------
    formato14 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': '87CEEB',
    'border' : 1,
    })
    planilha.merge_range('AM4:AS4', 'LIMITES OPERACIONAIS', formato14)
    
    formato15 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 90,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '87CEEB',
    'border' : 1,
    })
    for i in ('LIU','LIE','LIA','LSA','LSE','LSU','BNDMO'):
        planilha.merge_range(4,col,5,col,i, formato15)
        col+=1
    # Campo OBSERVAÄôES --------------------------------------------------
    formato16 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': 'FFCC99',
    'border' : 1,
    })
    planilha.write(3,45,'', formato16)
    
    formato17 = arquivo.add_format({
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 90,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#FFCC99',
    'border' : 1,
    })
    planilha.merge_range(4,col,5,col,'OBSERVAÄôES', formato17)
    col+=1  
    
    return arquivo




'''
black #000000
blue #0000FF
brown #800000
cyan #00FFFF
sky blue 10 #87CEEB
gray #808080
green #008000
lime #00FF00
magenta #FF00FF
navy #000080
orange #FF6600
pink #FF00FF
purple #800080
red #FF0000
silver #C0C0C0
white #FFFFFF
yellow #FFFF00

'''

