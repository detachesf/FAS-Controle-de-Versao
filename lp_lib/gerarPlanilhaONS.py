# -*- coding: cp860 -*-
dados= u'''
AtualizaáÑo do programa: 03/09/2014
Este m¢dulo implementa as funcionalidades relacionadas a geraáÑo da planilha de pontos ONS a partir de uma lista de pontos padrÑo.
'''

import os
from tkinter.messagebox import showerror

try:
    import xlsxwriter
except:
    showerror('Erro',u'M¢dulo XlsxWriter nÑo instalado')

kCol = 0.9  #Constante para transformar medida de coluna em Excel em medida de coluna no xlwt
kLin = 20   #Constante para transformar medida de linha em Excel em medida de coluna no xlwt


# armazena largura das colunas do cabecalho da planilha (valores abaixo sao os valores
# a serem exibidos na planilha em EXCEL
largura = [ kCol*20.57, kCol*90, kCol*6, kCol*6, kCol*12.29, kCol*76, kCol*12, kCol*12.43, kCol*84.86]


def gerarPlanilhaONS(Codigo_SE, listaEventosEPontos, listaDeFalhas):
    ''' 
    Gera as tabelas ONS para os bays da LP avaliada.
    @param Codigo_SE: Codigo da Subestacao
    @param listaEventosEPontos: Lista de Eventos e Pontos a serem utilizados para gerar Planilha ONS
    @param listaDeFalhas: Lista de pontos que aparecem com algum tipo de falha ou discordancia  
    @return: retorna o nome do arquivo de saida que contem a Planilha ONS          
    '''


    def gravaSEP_ONS (planilhaAserGravadaSeq, linhaTabelaSeq, baySeq):
       
       # verifica se h† algum ponto SEP (8219a1)
       for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8219a1'): # verifica se h† algum ponto SEP (8219a1)
                # 01 - 8.2.1.9.a.1 - SEP Sistemas Especiais de ProteáÑo ------------------------------------------------------------
                linhaTabelaSeq = linhaTabelaSeq + 1 
                planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,4,u'8.2.1.9 - SEP: Sistemas Especiais de ProteáÑo', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,5,linhaTabelaSeq+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,6,linhaTabelaSeq+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
                linhaTabelaSeq = linhaTabelaSeq + 1 
                planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,1,linhaTabelaSeq+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,2,linhaTabelaSeq,4,u'Atendido', style_hor_neg_cent_cinza)
                linhaTabelaSeq = linhaTabelaSeq + 1                
                planilhaAserGravadaSeq.write(linhaTabelaSeq,2,u'Sim', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.write(linhaTabelaSeq,3,u'NÑo', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.write(linhaTabelaSeq,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.write(linhaTabelaSeq,5,u'PONTO', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.write(linhaTabelaSeq,6,u'Endereáo', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.write(linhaTabelaSeq,7,u'Tipo ponto', style_hor_neg_cent_cinza)
                planilhaAserGravadaSeq.write(linhaTabelaSeq,8,u'ObservaáÑo', style_hor_neg_cent_cinza)                
                linhaTabelaSeq = linhaTabelaSeq + 1 
                
                # Escrita dos itens do requisito de rede 8.2.1.9.a.1 ---------------------------------------------------------- 
                planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Todos os disparos e alarmes;', style_hor_neg_esq)
                for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)                 
                    
                pos = 1
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                     
                planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.9.a.1', style_hor_neg_cent) 
                linhaTabelaSeq = linhaTabelaSeq + pos
                break           
        
       return linhaTabelaSeq

    def gravaSeqEventosDisjuntor(planilhaAsergravadaSeq,linhaTabelaSeq, nomeDisjuntor):
        ''' 
        Grava no arquivo Excel uma tabela com pontos da sequencia de eventos de um Disjuntor
        @param linhaTabelaSeq: linha a partir da qual ser† gravada a tabela
        @param planilhaAsergravadaSeq: planilha da pasta de trabalho do arquivo em Excel (xlsxwriter.worksheet)
        @param nomeDisjuntor: nome do disjuntor 
        @return: retorna a £ltima pr¢xima linha que poder† ser utilizada para gravar uma nova tabela na "planilhaAserGravadaSeq"          
        '''        
        # 04 - 8.2.1.8 - SEQUENCIA DE EVENTOS DISJUNTOR LADO DE ALTA ************************************************************
        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,4,u'8.2.1.8 - SEQUâNCIA DE EVENTOS DISJUNTOR - ' + nomeDisjuntor, style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,5,linhaTabelaSeq+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,6,linhaTabelaSeq+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1 
        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,1,linhaTabelaSeq+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,2,linhaTabelaSeq,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1
        planilhaAsergravadaSeq.write(linhaTabelaSeq,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.write(linhaTabelaSeq,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.write(linhaTabelaSeq,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.write(linhaTabelaSeq,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.write(linhaTabelaSeq,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.write(linhaTabelaSeq,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAsergravadaSeq.write(linhaTabelaSeq,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1       
        
        # Escrita dos itens do requisito de rede 8.2.1.8.a.1 ---------------------------------------------------------- 
        planilhaAsergravadaSeq.write(linhaTabelaSeq,1,u'Alarme de mudanáa de posiáÑo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAsergravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)        
    
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8218a1'):
                #totalPontos = len(conjuntoPontos[1])    # s¢ haver† um ponto associado aqui
                for item in conjuntoPontos[1]:    
                    if (item[6].find(nomeDisjuntor)>=0):   # verifica se h† referància ao disjuntor  no Id do ponto    
                        planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,1,item[3], style_hor_norm_dir)              
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_cent)                     
                            else: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_esq)                                   
                        pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+pos-1,0,u'8.2.1.8.a.1', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos           
                  
        # Escrita dos itens do requisito de rede 8.2.1.8.a.2 --------------------------------------------------------------------
        planilhaAsergravadaSeq.write(linhaTabelaSeq,1,u'Disparo da proteáÑo de falha do disjuntor;', style_hor_neg_esq)
        for col in range(2,9): planilhaAsergravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)        
    
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8218a2'):
                #totalPontos = len(conjuntoPontos[1])    # s¢ haver† um ponto associado aqui
                for item in conjuntoPontos[1]:    
                    if (item[6].find(nomeDisjuntor)>=0):   # verifica se h† referància ao disjuntor no Id do ponto    
                        planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,1,item[3], style_hor_norm_dir)              
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_cent)                     
                            else: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_esq)                                   
                        pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+pos-1,0,u'8.2.1.8.a.2', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos                                              
         
        # Escrita dos itens do requisito de rede 8.2.1.8.a.3 --------------------------------------------------------------------                 
        planilhaAsergravadaSeq.write(linhaTabelaSeq,1,u'Disparo do relà de bloqueio;', style_hor_neg_esq)
        for col in range(2,9): planilhaAsergravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)        
    
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8218a3'):
                #totalPontos = len(conjuntoPontos[1])    # s¢ haver† um ponto associado aqui
                for item in conjuntoPontos[1]:    
                    if (item[6].find(nomeDisjuntor)>=0):   # verifica se h† referància ao disjuntor no Id do ponto    
                        planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,1,item[3], style_hor_norm_dir)              
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_cent)                     
                            else: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_esq)                                   
                        pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+pos-1,0,u'8.2.1.8.a.3', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos                                         
        
        # Escrita dos itens do requisito de rede 8.2.1.8.b.1 --------------------------------------------------------------------                
        planilhaAsergravadaSeq.write(linhaTabelaSeq,1,u'Disparo da proteáÑo de discordÉncia de polos;', style_hor_neg_esq)
        for col in range(2,9): planilhaAsergravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)        
    
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8218b1'):
                #totalPontos = len(conjuntoPontos[1])    # s¢ haver† um ponto associado aqui
                for item in conjuntoPontos[1]:    
                    if (item[6].find(nomeDisjuntor)>=0):   # verifica se h† referància ao disjuntor no Id do ponto    
                        planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,1,item[3], style_hor_norm_dir)              
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_cent)                     
                            else: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_esq)                                   
                        pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+pos-1,0,u'8.2.1.8.b.1', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos 
    
        # Escrita dos itens do requisito de rede 8.2.1.8.b.2 --------------------------------------------------------------------        
        planilhaAsergravadaSeq.write(linhaTabelaSeq,1,u'Alarme de fechamento bloqueado;', style_hor_neg_esq)
        for col in range(2,9): planilhaAsergravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)        
    
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8218b2'):
                #totalPontos = len(conjuntoPontos[1])    # s¢ haver† um ponto associado aqui
                for item in conjuntoPontos[1]:    
                    if (item[6].find(nomeDisjuntor)>=0):   # verifica se h† referància ao disjuntor no Id do ponto    
                        planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,1,item[3], style_hor_norm_dir)              
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_cent)                     
                            else: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_esq)                                   
                        pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+pos-1,0,u'8.2.1.8.b.2', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos        
                        
        # Escrita dos itens do requisito de rede 8.2.1.8.b.3 -------------------------------------------- -----------------------               
        planilhaAsergravadaSeq.write(linhaTabelaSeq,1,u'Alarme de abertura bloqueada;', style_hor_neg_esq)
        for col in range(2,9): planilhaAsergravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)        
    
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8218b3'):
                #totalPontos = len(conjuntoPontos[1])    # s¢ haver† um ponto associado aqui
                for item in conjuntoPontos[1]:    
                    if (item[6].find(nomeDisjuntor)>=0):   # verifica se h† referància ao disjuntor no Id do ponto    
                        planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,1,item[3], style_hor_norm_dir)              
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_cent)                     
                            else: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_esq)                                   
                        pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+pos-1,0,u'8.2.1.8.b.3', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos                                           
        
        # Escrita dos itens do requisito de rede 8.2.1.8.b.4 ----------------------------------------------------------                
        planilhaAsergravadaSeq.write(linhaTabelaSeq,1,u'Alarme de sobrecarga do disjuntor central;', style_hor_neg_esq)
        for col in range(2,9): planilhaAsergravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)        
    
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8218b4'):
                #totalPontos = len(conjuntoPontos[1])    # s¢ haver† um ponto associado aqui
                for item in conjuntoPontos[1]:    
                    if (item[6].find(nomeDisjuntor)>=0):   # verifica se h† referància ao disjuntor no Id do ponto    
                        planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,1,item[3], style_hor_norm_dir)              
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_cent)                     
                            else: planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col,item[col-2], style_hor_norm_esq)                                   
                        pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAsergravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAsergravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+pos-1,0,u'8.2.1.8.b.4', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos 
        return linhaTabelaSeq
    
    def gravaLT_ONS(linhaTabelaSeq, planilhaAserGravadaSeq, baySeq):
        ''' 
        Grava no arquivo Excel uma tabela com pontos da LT
        @param linhaTabelaSeq: linha a partir da qual ser† gravada a tabela
        @param planilhaAserGravadaSeq: planilha da pasta de trabalho do arquivo em Excel (xlsxwriter.worksheet)
        @param baySeq: array com dados do bay a ser gravado   
        @return: retorna a £ltima pr¢xima linha que poder† ser utilizada para gravar uma nova tabela na "planilhaAserGravadaSeq"          
        '''
        #==========================================================================================================
        # FORMATACAO DA PLANILHA ONS LINHA DE TRANSMISSAO
        #==========================================================================================================
        # Escrita do cabecalho no arquivo----------------------------------------------------------------
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,8,u'LT - '+ bay[0] , style_hor_neg_cent_azul)            
                    
        # 01 - 7.3.1.1 - MEDICOES ANALOGICAS ------------------------------------------------------------
        linhaTabelaSeq = linhaTabelaSeq + 1
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,4,u'7.3.1.1 - MEDIÄôES ANALüGICAS', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,5,linhaTabelaSeq+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,6,linhaTabelaSeq+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1 
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,1,linhaTabelaSeq+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,2,linhaTabelaSeq,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1                
        planilhaAserGravadaSeq.write(linhaTabelaSeq,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        
        # Escrita dos itens do requisito de rede 7.3.1.1.c.4 -------------------------------------------------------------------
        linhaTabelaSeq = linhaTabelaSeq + 1
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Potància trif†sica ativa em MW e reativa em Mvar nos terminais de todas as LT;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7311c4'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'7.3.1.1.c.4', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos                   
                                                   
        # Escrita dos itens do requisito de rede 7.3.1.1.c.5 --------------------------------------------------------------------
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Corrente em uma das fases em Ampere nos terminais de todas as LT;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7311c5'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'7.3.1.1.c.5', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos                                      
                    
        # Escrita dos itens do requisito de rede 7.3.1.1.c.6 --------------------------------------------------------------------
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Uma mediáÑo do m¢dulo de tensÑo fase-fase em kV de cada terminal de LT;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7311c6'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'7.3.1.1.c.6', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos + 1           
        
                                    
        # 02 - 7.3.1.2 - SINALIZACAO DE ESTADO ----------------------------------------------------------
        # Escrita do cabecalho --------------------------------------------------------------------------
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,4,u'7.3.1.2 - SINALIZAÄéO DE ESTADO', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,5,linhaTabelaSeq+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,6,linhaTabelaSeq+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1 
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,1,linhaTabelaSeq+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,2,linhaTabelaSeq,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1
        planilhaAserGravadaSeq.write(linhaTabelaSeq,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        
        # Escrita dos itens do requisito de rede 7.3.1.2.a ----------------------------------------------------------------------
        linhaTabelaSeq = linhaTabelaSeq + 1
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Referente a todos os disjuntores e chaves utilizados nos barramentos e nas conexîes de equipamentos da rede de operaáÑo, a° inclu°das as chaves de by-pass. Este requisito Ç aplic†vel tanto a sistemas de geraáÑo e transmissÑo em corrente alternada quanto a sistemas de transmissÑo em CC (incluindo filtros), sendo que para os disjuntores Ç necess†rio que a sinalizaáÑo seja acompanhada do selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312a'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[6].split(':')[1], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'7.3.1.2.a', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos        
                                
        # Escrita dos itens do requisito de rede 7.3.1.2.d ----------------------------------------------------------------------  
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'RelÇs de bloqueio com selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312d'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'7.3.1.2.d', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos + 1                       
            
        # 03 - 8.2.1.4 - SEQUENCIA DE EVENTOS LINHA DE TRANSMISSAO
        # Escrita do cabecalho --------------------------------------------------------------------------------------------------
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,4,u'8.2.1.4 - SEQUâNCIA DE EVENTOS LINHA DE TRANSMISSéO', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,5,linhaTabelaSeq+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,6,linhaTabelaSeq+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1 
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,1,linhaTabelaSeq+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,2,linhaTabelaSeq,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1
        planilhaAserGravadaSeq.write(linhaTabelaSeq,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
            
        # Escrita dos itens do requisito de rede 8.2.1.4.a.1 --------------------------------------------
        linhaTabelaSeq = linhaTabelaSeq + 1
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Disparo por sobretensÑo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214a1'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.a.1', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos               
             
        # Escrita dos itens do requisito de rede 8.2.1.4.a.2 --------------------------------------------  
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'AtuaáÑo da l¢gica de bloqueio por oscilaáÑo de potància;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214a2'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.a.2', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos
          
        # Escrita dos itens do requisito de rede 8.2.1.4.a.3 --------------------------------------------        
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Disparo da proteáÑo para perda de sincronismo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214a3'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.a.3', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos        
                               
        # Escrita dos itens do requisito de rede 8.2.1.4.a.4 --------------------------------------------                
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'AtuaáÑo do relÇ de bloqueio de recepáÑo permanente de transferància de disparo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214a4'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.a.4', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos
    
        # Escrita dos itens do requisito de rede 8.2.1.4.a.5 --------------------------------------------                
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Disparo do relÇ de bloqueio de linha subterrÉnea;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214a5'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.a.5', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos                                     
            
        # Escrita dos itens do requisito de rede 8.2.1.4.b ----------------------------------------------        
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+14,0,'8.2.1.4.b', style_hor_neg_cent) 
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'AtuaáÑo da proteáÑo da linha de transmissÑo - outras funáîes (Agrupamento dos eventos abaixo relacionados):', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq, col,'', style_hor_norm_esq)  

        planilhaAserGravadaSeq.write(linhaTabelaSeq+1,1,u'Disparo da proteáÑo principal de fase;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+2,1,u'Disparo da proteáÑo alternada de fase;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+3,1,u'Disparo da proteáÑo principal de neutro;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+4,1,u'Disparo da proteáÑo alternada de neutro;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+5,1,u'TransmissÑo de sinal de desbloqueio/bloqueio ou sinal permissivo da teleproteáÑo;', style_hor_norm_dir)        
        planilhaAserGravadaSeq.write(linhaTabelaSeq+6,1,u'TransmissÑo de sinal de transferància de disparo da teleproteáÑo;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+7,1,u'RecepáÑo de sinal de desbloqueio/bloqueio ou sinal permissivo da teleproteáÑo;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+8,1,u'Disparo por recepáÑo de sinal de transferància de disparo da teleproteáÑo;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+9,1,u'AtuaáÑo da l¢gica de bloqueio por perda de potencial;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+10,1,u'Disparo da 2¶ zona da proteáÑo de distÉncia;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+11,1,u'Disparo da 3¶ zona da proteáÑo de distÉncia;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+12,1,u'Disparo da 4¶ zona da proteáÑo de distÉncia;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+13,1,u'Disparo da proteáÑo de sobrecorrente direcional de neutro temporizada;', style_hor_norm_dir)
        planilhaAserGravadaSeq.write(linhaTabelaSeq+14,1,u'Disparo da proteáÑo de sobrecorrente direcional de neutro instantÉnea.', style_hor_norm_dir)

        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214b'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.merge_range(linhaTabelaSeq+pos,col,linhaTabelaSeq+14,col, item[col-2], style_hor_norm_cent)                    
                        else:                  planilhaAserGravadaSeq.merge_range(linhaTabelaSeq+pos,col,linhaTabelaSeq+14,col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(2,9):
                planilhaAserGravadaSeq.merge_range(linhaTabelaSeq+pos,col, linhaTabelaSeq+14, col,'', style_hor_norm_cent_verm) 
            pos = pos +1 
                                                  
        linhaTabelaSeq = linhaTabelaSeq + 14 
                 
        # Escrita dos itens do requisito de rede 8.2.1.4.c.1 --------------------------------------------
        linhaTabelaSeq = linhaTabelaSeq + 1    
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Partida da proteáÑo principal de fase (por fase), nos casos em que o disparo da proteáÑo de fase nÑo indique a(s) fase(s) defeituosas;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214c1'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.c.1', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos                                 
                        
        # Escrita dos itens do requisito de rede 8.2.1.4.c.2 --------------------------------------------                    
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Partida da proteáÑo alternada de fase (por fase), nos casos em que o disparo da proteáÑo de fase nÑo indique a(s) fase(s) defeituosas;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214c2'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.c.2', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos                    
                   
        # Escrita dos itens do requisito de rede 8.2.1.4.c.3 --------------------------------------------                
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Partida da proteáÑo principal de neutro (por fase), nos casos em que o disparo da proteáÑo nÑo indique a fase defeituosa;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214c3'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.c.3', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos              
        
        # Escrita dos itens do requisito de rede 8.2.1.4.c.4 --------------------------------------------        
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Partida da proteáÑo alternada de neutro (por fase), nos casos em que o disparo da proteáÑo nÑo indique a fase defeituosa;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214c4'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.c.4', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos
                             
        # Escrita dos itens do requisito de rede 8.2.1.4.c.5 --------------------------------------------              
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Partida do religamento autom†tico.;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8214c5'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.4.c.5', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos + 1                                         
             
        # 04 - 8.2.1.8 - SEQUENCIA DE EVENTOS DISJUNTOR
        nomeDisjuntor = bay[0]
        nomeDisjuntor = nomeDisjuntor.replace('0','1')
        linhaTabelaSeq = gravaSeqEventosDisjuntor(planilhaAserGravadaSeq, linhaTabelaSeq, nomeDisjuntor)                    
        linhaTabelaSeq = linhaTabelaSeq + 1
        
        return linhaTabelaSeq 
    
    def gravaTR_ONS(linhaTabela,  planilhaAserGravada, bay):
        ''' 
        Grava no arquivo Excel uma tabela com pontos de um Transformador
        @param linhaTabela: linha a partir da qual ser† gravada a tabela
        @param planilhaAserGravada: planilha da pasta de trabalho do arquivo em Excel (xlsxwriter.worksheet)
        @param bay: nome do bay 
        @return: retorna a £ltima pr¢xima linha que poder† ser utilizada para gravar uma nova tabela na "planilhaAserGravada"          
        '''    
        
        def disjuntores(abay):
            ''' 
            Grava no arquivo Excel uma tabela com pontos da sequencia de eventos de um Transformador
            @param abay: lista de pontos do bay em questao
            @return: retorna nomes dos disjuntores do lado de alta e lado de baixa          
            '''    
            listaDisjuntores=[]
            for conjuntoPts in abay:
                try:
                    for item in conjuntoPts[1]:
                        evento = item[6].split(':')[1]
                        if (evento[0]=='1'): # Ç um disjuntor
                            if (evento not in listaDisjuntores): #testa se nÑo est† na lista
                                listaDisjuntores.append(evento)
                except:
                   pass  
                
            if (listaDisjuntores[0][1]>listaDisjuntores[1][1]): 
                disLadoAlta  = listaDisjuntores[0]
                disjLadoBaixa= listaDisjuntores[1]
            else:
                disLadoAlta  = listaDisjuntores[1]
                disjLadoBaixa= listaDisjuntores[0]               
                    
            return disLadoAlta, disjLadoBaixa
        
        #============================================================================================================================================
        # FORMATACAO DA PLANILHA ONS TRANSFORMADORES
        #============================================================================================================================================
        # segregaáÑo dos disjuntores de Alta e Baixa do transformado
        disjuntorAlta,disjuntorBaixa = disjuntores(bay)
        # Escrita do cabecalho no arquivo
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,8,u'CT - '+ bay[0] , style_hor_neg_cent_azul)            
        # 01 - 7.3.1.1 - MEDICOES ANALOGICAS **********************************************************************************************
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,4,u'7.3.1.1 - MEDIÄôES ANALüGICAS', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,5,linhaTabela+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,6,linhaTabela+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1 
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,1,linhaTabela+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,2,linhaTabela,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
       
        # Escrita dos itens do requisito de rede 7.3.1.1.c.8 --------------------------------------------------------------------
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,1,u'Potància trif†sica ativa em MW e reativa em MVAr e corrente em Ampäre do prim†rio e secund†rio dos transformadores;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7311c8'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    if   (item[6].find('AMPA:P')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Corrente fase A Prim†rio'   , style_hor_norm_dir) 
                    elif (item[6].find('AMPB:P')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Corrente fase B Prim†rio'   , style_hor_norm_dir)
                    elif (item[6].find('AMPC:P')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Corrente fase C Prim†rio'   , style_hor_norm_dir)
                    elif (item[6].find('AMPA:S')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Corrente fase A Secund†rio' , style_hor_norm_dir)
                    elif (item[6].find('AMPB:S')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Corrente fase B Secund†rio' , style_hor_norm_dir)
                    elif (item[6].find('AMPC:S')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Corrente fase C Secund†rio' , style_hor_norm_dir) 
                    elif (item[6].find('AMPA:T')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Corrente fase A Terci†rio'  , style_hor_norm_dir)  
                    elif (item[6].find('AMPB:T')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Corrente fase B Terci†rio'  , style_hor_norm_dir)  
                    elif (item[6].find('AMPC:T')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Corrente fase C Terci†rio'  , style_hor_norm_dir)                          
                    elif (item[6].find('MW:P')>0):   planilhaAserGravada.write(linhaTabela+pos,1,u'Potància Ativa Prim†rio'    , style_hor_norm_dir)
                    elif (item[6].find('MVAR:P')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Potància Reativa Prim†rio'  , style_hor_norm_dir)
                    elif (item[6].find('MW:S')>0):   planilhaAserGravada.write(linhaTabela+pos,1,u'Potància Ativa Secund†rio'  , style_hor_norm_dir)
                    elif (item[6].find('MVAR:S')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Potància Reativa Secund†rio', style_hor_norm_dir)
                    elif (item[6].find('MW:T')>0):   planilhaAserGravada.write(linhaTabela+pos,1,u'Potància Ativa Terci†rio'   , style_hor_norm_dir)
                    elif (item[6].find('MVAR:T')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'Potància Reativa Terci†rio' , style_hor_norm_dir)                                                                                                                                                                      
                    else: planilhaAserGravada.write(linhaTabela+pos,1,item[3], style_hor_norm_dir)
                    # molduras    
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm)   
            pos = pos +1    
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.1.c.8', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos
        
        # Escrita dos itens do requisito de rede 7.3.1.1.c.13 -------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'PosiáÑo de tap de transformadores equipados com comutadores sob carga, desde que tecnicamente vi†vel; Nos casos em que se constate este tipo de inviabilidade esta dever† ser eliminada quando da substituiáÑo do transformador;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7311c13'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    if   (item[6].find('TAPA')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'PosiáÑo de TAP Fase A', style_hor_norm_dir) 
                    elif (item[6].find('TAPB')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'PosiáÑo de TAP Fase B', style_hor_norm_dir)
                    elif (item[6].find('TAPC')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'PosiáÑo de TAP Fase C', style_hor_norm_dir)
                    elif (item[6].find('TAP')>0):  planilhaAserGravada.write(linhaTabela+pos,1,u'PosiáÑo de TAP'       , style_hor_norm_dir)
                    else: planilhaAserGravada.write(linhaTabela+pos,1,item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm)   
            pos = pos +1    
        
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.1.c.13', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos

        # Escrita dos itens do requisito de rede 7.3.1.1.c.14 -------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'1 (uma) mediáÑo do m¢dulo de tensÑo fase-fase em kV para os transformadores, excetuando-se aqueles na fronteira da rede de operaáÑo. Esta mediáÑo deve ser no lado ligado Ö barra de menor potància de curto-circuito, geralmente o de menor tensÑo, caso o ONS nÑo explicite que seja no outro lado do transformador.', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7311c14'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    if   (item[6].find('KVAB:P')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'TensÑo fases AB Prim†rio'  , style_hor_norm_dir) 
                    elif (item[6].find('KVBC:P')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'TensÑo fases BC Prim†rio'  , style_hor_norm_dir)
                    elif (item[6].find('KVCA:P')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'TensÑo fases CA Prim†rio'  , style_hor_norm_dir)
                    elif (item[6].find('KVAB:S')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'TensÑo fases AB Secund†rio', style_hor_norm_dir) 
                    elif (item[6].find('KVBC:S')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'TensÑo fases BC Secund†rio', style_hor_norm_dir)
                    elif (item[6].find('KVCA:S')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'TensÑo fases CA Secund†rio', style_hor_norm_dir)
                    elif (item[6].find('KVAB:T')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'TensÑo fases AB Terci†rio' , style_hor_norm_dir) 
                    elif (item[6].find('KVBC:T')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'TensÑo fases BC Terci†rio' , style_hor_norm_dir)
                    elif (item[6].find('KVCA:T')>0): planilhaAserGravada.write(linhaTabela+pos,1,u'TensÑo fases CA Terci†rio' , style_hor_norm_dir)                                                                                                                                                                                             
                    else: planilhaAserGravada.write(linhaTabela+pos,1,item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1  
        
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.1.c.14', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos             
            
        # 01 - 7.3.1.2 - SINALIZACAO DE ESTADO ********************************************************************************************
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,4,u'7.3.1.2 - SINALIZAÄéO DE ESTADO', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,5,linhaTabela+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,6,linhaTabela+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1 
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,1,linhaTabela+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,2,linhaTabela,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        
        # Escrita dos itens do requisito de rede 7.3.1.2.a ----------------------------------------------------------------------
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,1,u'Referente a todos os disjuntores e chaves utilizados nos barramentos e nas conexîes de equipamentos da rede de operaáÑo, a° inclu°das as chaves de by-pass. Este requisito Ç aplic†vel tanto a sistemas de geraáÑo e transmissÑo em corrente alternada quanto a sistemas de transmissÑo em CC (incluindo filtros), sendo que para os disjuntores Ç necess†rio que a sinalizaáÑo seja acompanhada do selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312a'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[6].split(':')[1], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm)      
            pos = pos +1     
        
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.2.a', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos        
        
        # Escrita dos itens do requisito de rede 7.3.1.2.d --------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'RelÇs de bloqueio com selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312d'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[6].split(':')[1], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.2.d', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos          
                
        # Escrita dos itens do requisito de rede 7.3.1.2.f --------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'Estado dos comutadores sob carga (em autom†tico/manual/remoto);', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312f'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.2.f', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos         
               
        # Escrita dos itens do requisito de rede 7.3.1.2.h --------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'Alarmes de temperatura de enrolamento e ¢leo de trafos e reatores.', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312h'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm)  
            pos = pos +1                      

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.2.h', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos + 1      
                             
        # 01 - 8.2.1.8 - SEQUâNCIA DE EVENTOS TRANSFORMADOR *******************************************************************************
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,4,u'8.2.1.1 - SEQUâNCIA DE EVENTOS TRANSFORMADOR', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,5,linhaTabela+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,6,linhaTabela+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1 
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,1,linhaTabela+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,2,linhaTabela,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1        
        
        # Escrita dos itens do requisito de rede 8.2.1.1.a.1 --------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'Disparo do RelÇ de Bloqueio;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8211a1'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'8.2.1.1.a.1', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos  
                   
        # Escrita dos itens do requisito de rede 8.2.1.1.b.1 --------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'AtuaáÑo da proteáÑo do transformador - FunáÑo Sobrecorrente (Agrupamento dos eventos abaixo relacionados):', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)        
        planilhaAserGravada.write(linhaTabela+1,1,u'AtuaáÑo da proteáÑo de sobrecorrente do comutador sob carga;', style_hor_norm_dir)
        planilhaAserGravada.write(linhaTabela+2,1,u'Disparo da proteáÑo de sobrecorrente de fase e neutro (por enrolamento);', style_hor_norm_dir)        
          
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8211b1'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+1,col,item[col-2], style_hor_norm_cent)                     
                        else: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+1,col,item[col-2], style_hor_norm_esq)                                   
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(2,9):
                planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+1,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos+1,0,u'8.2.1.1.b.1', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos +1               
        
        # Escrita dos itens do requisito de rede 8.2.1.1.b.2 --------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'AtuaáÑo da proteáÑo do transformador - FunáÑo Sobretemperatura (Agrupamento dos eventos abaixo relacionados):', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)        
        planilhaAserGravada.write(linhaTabela+1,1,u'Disparo por sobretemperatura do ¢leo;', style_hor_norm_dir)
        planilhaAserGravada.write(linhaTabela+2,1,u'Disparo por sobretemperatura do enrolamento;', style_hor_norm_dir)         
          
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8211b2'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+1,col,item[col-2], style_hor_norm_cent)                     
                        else: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+1,col,item[col-2], style_hor_norm_esq)                                   
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(2,9):
                planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+1,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos+1,0,u'8.2.1.1.b.2', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos +1  
        
        # Escrita dos itens do requisito de rede 8.2.1.1.b.3 --------------------------------------------        
        planilhaAserGravada.write(linhaTabela,1,u'AtuaáÑo da proteáÑo do transformador - Outras funáîes (Agrupamento dos eventos abaixo relacionados):', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)        
        planilhaAserGravada.write(linhaTabela+1,1,u'Disparo da proteáÑo de g†s;', style_hor_norm_dir)
        planilhaAserGravada.write(linhaTabela+2,1,u'Disparo da proteáÑo de sobretensÑo de sequància zero para o enrolamento terci†rio em ligaáÑo delta;', style_hor_norm_dir) 
        planilhaAserGravada.write(linhaTabela+3,1,u'Disparo da proteáÑo de g†s do comutador de derivaáîes;', style_hor_norm_dir) 
        planilhaAserGravada.write(linhaTabela+4,1,u'Disparo da proteáÑo diferencial (por fase);', style_hor_norm_dir)         
          
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8211b3'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+3,col,item[col-2], style_hor_norm_cent)                     
                        else: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+3,col,item[col-2], style_hor_norm_esq)                                   
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(2,9):
                planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+3,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos+3,0,u'8.2.1.1.b.3', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos + 4                     
            
        # 04 - 8.2.1.8 - SEQUENCIA DE EVENTOS DISJUNTOR LADO DE ALTA ************************************************************
        linhaTabela = gravaSeqEventosDisjuntor(planilhaAserGravada, linhaTabela, disjuntorAlta)
        linhaTabela = linhaTabela + 1
        # 04 - 8.2.1.8 - SEQUENCIA DE EVENTOS DISJUNTOR LADO DE BAIXA ***********************************************************

        linhaTabela = gravaSeqEventosDisjuntor(planilhaAserGravada, linhaTabela, disjuntorBaixa)        
        linhaTabela = linhaTabela + 1      
               
        return linhaTabela
    
    def gravaBT_ONS(linhaTabela,  planilhaAserGravada, bay):
        ''' 
        Grava no arquivo Excel uma tabela com pontos de um Bay de Transferencia
        @param linhaTabela: linha a partir da qual ser† gravada a tabela
        @param planilhaAserGravada: planilha da pasta de trabalho do arquivo em Excel (xlsxwriter.worksheet)
        @param bay: nome do bay 
        @return: retorna a £ltima pr¢xima linha que poder† ser utilizada para gravar uma nova tabela na "planilhaAserGravada"          
        '''         
        # Escrita do cabecalho no arquivo----------------------------------------------------------------
        if (bay[0][1]=='5'):
            planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,8,u'IB - '+ bay[0] , style_hor_neg_cent_azul)   
        else: 
            planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,8,u'BT - '+ bay[0] , style_hor_neg_cent_azul)   
                         
        # 01 - 7.3.1.2 - SINALIZACAO DE ESTADO ------------------------------------------------------------
        linhaTabela = linhaTabela + 1            
        planilhaAserGravada.merge_range(linhaTabela,0, linhaTabela,  4,u'7.3.1.2 - SINALIZAÄéO DE ESTADO', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,5, linhaTabela+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,6, linhaTabela+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1 
        planilhaAserGravada.merge_range(linhaTabela,0, linhaTabela+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,1, linhaTabela+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,2, linhaTabela  ,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        
        # Escrita dos itens do requisito de rede 7.3.1.2.a --------------------------------------------
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,1,u'Referente a todos os disjuntores e chaves utilizados nos barramentos e nas conexîes de equipamentos da rede de operaáÑo, a° inclu°das as chaves de by-pass. Este requisito Ç aplic†vel tanto a sistemas de geraáÑo e transmissÑo em corrente alternada quanto a sistemas de transmissÑo em CC (incluindo filtros), sendo que para os disjuntores Ç necess†rio que a sinalizaáÑo seja acompanhada do selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312a'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[6].split(':')[1], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm)      
            pos = pos +1     
        
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.2.a', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos        
        
        # Escrita dos itens do requisito de rede 7.3.1.2.d --------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'RelÇs de bloqueio com selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312d'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[6].split(':')[1], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.2.d', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos  
        
        # 04 - 8.2.1.8 - SEQUENCIA DE EVENTOS DISJUNTOR
        linhaTabela = linhaTabela + 1 
        linhaTabela = gravaSeqEventosDisjuntor(planilhaAserGravada, linhaTabela, bay[0])        
        linhaTabela = linhaTabela + 1 
                
        return linhaTabela

    def gravaRE_ONS(linhaTabela,  planilhaAserGravada, bay):
        ''' 
        Grava no arquivo Excel uma tabela com pontos dao Reator
        @param linhaTabela: linha a partir da qual ser† gravada a tabela
        @param planilhaAserGravada: planilha da pasta de trabalho do arquivo em Excel (xlsxwriter.worksheet)
        @param bay: nome do bay 
        @return: retorna a £ltima pr¢xima linha que poder† ser utilizada para gravar uma nova tabela na "planilhaAserGravada"          
        ''' 
        #==========================================================================================================
        # FORMATACAO DA PLANILHA ONS REATORES
        #==========================================================================================================
                # Escrita do cabecalho no arquivo
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,8,u'RE - '+ bay[0] , style_hor_neg_cent_azul)    
        # 01 - 7.3.1.2 - SINALIZACAO DE ESTADO ********************************************************************************************
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,4,u'7.3.1.2 - SINALIZAÄéO DE ESTADO', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,5,linhaTabela+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,6,linhaTabela+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1 
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,1,linhaTabela+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,2,linhaTabela,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        
        # Escrita dos itens do requisito de rede 7.3.1.2.a ----------------------------------------------------------------------
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,1,u'Referente a todos os disjuntores e chaves utilizados nos barramentos e nas conexîes de equipamentos da rede de operaáÑo, a° inclu°das as chaves de by-pass. Este requisito Ç aplic†vel tanto a sistemas de geraáÑo e transmissÑo em corrente alternada quanto a sistemas de transmissÑo em CC (incluindo filtros), sendo que para os disjuntores Ç necess†rio que a sinalizaáÑo seja acompanhada do selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312a'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[6].split(':')[1], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm)      
            pos = pos +1     
        
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.2.a', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos        
        
        # Escrita dos itens do requisito de rede 7.3.1.2.d --------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'RelÇs de bloqueio com selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312d'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[6].split(':')[1], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.2.d', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos           
 
         # Escrita dos itens do requisito de rede 7.3.1.2.h --------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'Alarmes de temperatura de enrolamento e ¢leo de trafos e reatores.', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312h'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm)  
            pos = pos +1                      
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.2.h', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos + 1             

        # 01 - 8.2.1.8 - SEQUâNCIA DE EVENTOS REATOR *******************************************************************************
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,4,u'8.2.1.1 - SEQUâNCIA DE EVENTOS REATOR', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,5,linhaTabela+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,6,linhaTabela+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1 
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,1,linhaTabela+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,2,linhaTabela,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1   

        # Escrita dos itens do requisito de rede 8.2.1.2.a.1 --------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'Disparo do RelÇ de Bloqueio;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8212a1'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'8.2.1.2.a.1', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos  

        # Escrita dos itens do requisito de rede 8.2.1.2.b.1 --------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'AtuaáÑo da proteáÑo do reator - FunáÑo sobretemperatura (Agrupamento dos eventos abaixo relacionados):', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)        
        planilhaAserGravada.write(linhaTabela+1,1,u'Disparo da proteáÑo de sobretemperatura do ¢leo;', style_hor_norm_dir)
        planilhaAserGravada.write(linhaTabela+2,1,u'Disparo da proteáÑo de sobretemperatura do enrolamento;', style_hor_norm_dir)        
          
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8212b1'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+1,col,item[col-2], style_hor_norm_cent)                     
                        else: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+1,col,item[col-2], style_hor_norm_esq)                                   
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(2,9):
                planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+1,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos+1,0,u'8.2.1.2.b.1', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos +1               

        # Escrita dos itens do requisito de rede 8.2.1.2.b.2 --------------------------------------------        
        planilhaAserGravada.write(linhaTabela,1,u'AtuaáÑo da proteáÑo do reator - Outras funáîes (Agrupamento dos eventos abaixo relacionados):', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)        
        planilhaAserGravada.write(linhaTabela+1,1,u'Disparo da proteáÑo de g†s;', style_hor_norm_dir)
        planilhaAserGravada.write(linhaTabela+2,1,u'Disparo da v†lvula de al°vio de pressÑo', style_hor_norm_dir) 
        planilhaAserGravada.write(linhaTabela+3,1,u'Disparo da proteáÑo diferencial (por fase);', style_hor_norm_dir) 
        planilhaAserGravada.write(linhaTabela+4,1,u'Disparo da proteáÑo de sobrecorrente de fase e neutro;', style_hor_norm_dir)         
          
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8212b2'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+3,col,item[col-2], style_hor_norm_cent)                     
                        else: planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+3,col,item[col-2], style_hor_norm_esq)                                   
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(2,9):
                planilhaAserGravada.merge_range(linhaTabela+pos,col,linhaTabela+pos+3,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos+3,0,u'8.2.1.2.b.2', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos + 4             
        
        # 04 - 8.2.1.8 - SEQUENCIA DE EVENTOS DISJUNTOR
        nomeDisjuntor = bay[0]
        nomeDisjuntor = nomeDisjuntor.replace('0','1')
        linhaTabela = gravaSeqEventosDisjuntor(planilhaAserGravada, linhaTabela, nomeDisjuntor)                    
        linhaTabela = linhaTabela + 1        
        
        return  linhaTabela   
    
    def gravaBA_ONS(linhaTabela,  planilhaAserGravada, bay):
        ''' 
        Grava no arquivo Excel uma tabela com pontos da Barra
        @param linhaTabela: linha a partir da qual ser† gravada a tabela
        @param planilhaAserGravada: planilha da pasta de trabalho do arquivo em Excel (xlsxwriter.worksheet)
        @param bay: nome do bay 
        @return: retorna a £ltima pr¢xima linha que poder† ser utilizada para gravar uma nova tabela na "planilhaAserGravada"          
        ''' 
        #==========================================================================================================
        # FORMATACAO DA PLANILHA ONS BARRAMENTO
        #==========================================================================================================
        # Escrita do cabecalho no arquivo----------------------------------------------------------------
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,8,u'BA - '+ bay[0] , style_hor_neg_cent_azul)            
                    
        # 01 - 7.3.1.1 - MEDICOES ANALOGICAS ------------------------------------------------------------
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,4,u'7.3.1.1 - MEDIÄôES ANALüGICAS', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,5,linhaTabela+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,6,linhaTabela+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1 
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,1,linhaTabela+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,2,linhaTabela,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1                
        planilhaAserGravada.write(linhaTabela,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        
        # Escrita dos itens do requisito de rede 7.3.1.1.c.1 -------------------------------------------------------------------
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,1,u'1 (uma) mediáÑo do m¢dulo de tensÑo fase-fase em kV de cada secáÑo de barramento que possa formar um n¢ elÇtrico. Caso venha a ser adotado o arranjo em anel, uma mediáÑo do m¢dulo de tensÑo fase-fase em kV nos terminais de cada equipamento que a ele se conectem (linhas de transmissÑo (LT), transformadores, etc.);', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7311c1'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.3.1.1.c.1', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos        

        # Escrita dos itens do requisito de rede 7.4.2.1.a -------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'FreqÅància em Hz em barramentos designados pelo ONS;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)  
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7421a'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'7.4.2.1.a', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos 
 
         # 01 - 8.2.1.5 - SEQUâNCIA DE EVENTOS BARRAS *******************************************************************************
        linhaTabela = linhaTabela + 1   
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela,4,u'8.2.1.1 - SEQUâNCIA DE EVENTOS BARRAS', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,5,linhaTabela+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,6,linhaTabela+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1 
        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,1,linhaTabela+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravada.merge_range(linhaTabela,2,linhaTabela,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1
        planilhaAserGravada.write(linhaTabela,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravada.write(linhaTabela,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
        linhaTabela = linhaTabela + 1   

        # Escrita dos itens do requisito de rede 8.2.1.5.a.1 --------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'Disparo da proteáÑo de sobretensÑo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8215a1'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'8.2.1.5.a.1', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos  
        
        # Escrita dos itens do requisito de rede 8.2.1.5.a.2 --------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'Disparo do RelÇ de Bloqueio;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8215a2'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravada.write(linhaTabela+pos,1, item[3], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravada.write(linhaTabela+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravada.write(linhaTabela+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'8.2.1.5.a.2', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos  
 
        # Escrita dos itens do requisito de rede 8.2.1.5.b.1 --------------------------------------------------------------------
        planilhaAserGravada.write(linhaTabela,1,u'AtuaáÑo da proteáÑo diferencial do barramento (Agrupamento dos eventos abaixo relacionados):', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)        
        planilhaAserGravada.write(linhaTabela+1,1,u'AtuaáÑo da proteáÑo diferencial (por fase);', style_hor_norm_dir)    
          
        for col in range(2,9): planilhaAserGravada.write(linhaTabela,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '8215b1'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravada.write(linhaTabela+pos,col,item[col-2], style_hor_norm_cent)                     
                        else:                  planilhaAserGravada.write(linhaTabela+pos,col,item[col-2], style_hor_norm_esq)                                   
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(2,9):
                planilhaAserGravada.write(linhaTabela+pos,col,'', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravada.merge_range(linhaTabela,0,linhaTabela+totalPontos,0,u'8.2.1.5.b.1', style_hor_neg_cent) 
        linhaTabela = linhaTabela + pos        
        
        # SEP ----------------------------------------------------------------------------------------------------
        linhaTabela = gravaSEP_ONS(planilhaAserGravada,linhaTabela,bay)
        
        linhaTabela = linhaTabela + 1
        return linhaTabela

    def gravaBC_ONS(linhaTabelaSeq, planilhaAserGravadaSeq, baySeq ):
        ''' 
        Grava no arquivo Excel uma tabela com pontos de um banco de capacitor
        @param linhaTabelaSeq: linha a partir da qual ser† gravada a tabela
        @param planilhaAserGravadaSeq: planilha da pasta de trabalho do arquivo em Excel (xlsxwriter.worksheet)
        @param baySeq: nome do bay 
        @return: retorna a £ltima pr¢xima linha que poder† ser utilizada para gravar uma nova tabela na "planilhaAserGravadaSeq"          
        ''' 
        #==========================================================================================================
        # FORMATACAO DA PLANILHA ONS BANCO CAPACITORES
        #==========================================================================================================
        # Escrita do cabecalho no arquivo----------------------------------------------------------------
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,8,u'BC - '+ bay[0] , style_hor_neg_cent_azul)            
                    
        # 01 - 7.3.1.2 - SINALIZAÄéO DE ESTADO ------------------------------------------------------------
        linhaTabelaSeq = linhaTabelaSeq + 1
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,4,u'7.3.1.2 - SINALIZAÄéO DE ESTADO', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,5,linhaTabelaSeq+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,6,linhaTabelaSeq+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1 
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,1,linhaTabelaSeq+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,2,linhaTabelaSeq,4,u'Atendido', style_hor_neg_cent_cinza)
        linhaTabelaSeq = linhaTabelaSeq + 1                
        planilhaAserGravadaSeq.write(linhaTabelaSeq,2,u'Sim', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,3,u'NÑo', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,5,u'PONTO', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,6,u'Endereáo', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,7,u'Tipo ponto', style_hor_neg_cent_cinza)
        planilhaAserGravadaSeq.write(linhaTabelaSeq,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
 
        # Escrita dos itens do requisito de rede 7.3.1.2.a ----------------------------------------------------------------------
        linhaTabelaSeq = linhaTabelaSeq + 1
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Referente a todos os disjuntores e chaves utilizados nos barramentos e nas conexîes de equipamentos da rede de operaáÑo, a° inclu°das as chaves de by-pass. Este requisito Ç aplic†vel tanto a sistemas de geraáÑo e transmissÑo em corrente alternada quanto a sistemas de transmissÑo em CC (incluindo filtros), sendo que para os disjuntores Ç necess†rio que a sinalizaáÑo seja acompanhada do selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312a'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[6].split(':')[1], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm)      
            pos = pos +1     
        
        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'7.3.1.2.a', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos        
        
        # Escrita dos itens do requisito de rede 7.3.1.2.d --------------------------------------------
        planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'RelÇs de bloqueio com selo de tempo;', style_hor_neg_esq)
        for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)      
        pos = 1 
        totalPontos = 1   
        for conjuntoPontos in bay:
            if (conjuntoPontos[0] == '7312d'):
                totalPontos = len(conjuntoPontos[1])    
                for item in conjuntoPontos[1]:                    
                    # itens
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[6].split(':')[1], style_hor_norm_dir)
                    # molduras
                    for col in range(2,9):
                        if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                        else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                    pos = pos + 1 
                break 
        if (pos==1):    # nÑo houve pontos para esse item
            for col in range(1,9):
                planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
            pos = pos +1                       

        planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'7.3.1.2.d', style_hor_neg_cent) 
        linhaTabelaSeq = linhaTabelaSeq + pos
  
        # verifica se h† algum ponto pertencente ao BCSÇrie ou BCShunt ----------------------------------------------------------
        BCShunt = False
        BCSerie = False
        for teste in bay:
            if (teste[0] in ['8213a1','8213a2','8213b']):
                BCShunt = True
                break
            elif (teste[0] in ['82111a1','82111a2','82111b']):
                BCSerie = True
                break
            
        # 01 - 8.2.1.8 - SEQUâNCIA DE EVENTOS BANCO CAPACITOR SHUNT ************************************************************************            
        if (BCShunt):    
            linhaTabelaSeq = linhaTabelaSeq + 1                      
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,4,u'8.2.1.1 - SEQUâNCIA DE EVENTOS BANCO CAPACITOR', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,5,linhaTabelaSeq+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,6,linhaTabelaSeq+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
            linhaTabelaSeq = linhaTabelaSeq + 1 
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,1,linhaTabelaSeq+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,2,linhaTabelaSeq,4,u'Atendido', style_hor_neg_cent_cinza)
            linhaTabelaSeq = linhaTabelaSeq + 1
            planilhaAserGravadaSeq.write(linhaTabelaSeq,2,u'Sim', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,3,u'NÑo', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,5,u'PONTO', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,6,u'Endereáo', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,7,u'Tipo ponto', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
            linhaTabelaSeq = linhaTabelaSeq + 1  
              
            # Escrita dos itens do requisito de rede 8.2.1.3.a.1 --------------------------------------------------------------------
            planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Disparo da proteáÑo de sobretensÑo;', style_hor_neg_esq)
            for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)      
            pos = 1 
            totalPontos = 1   
            for conjuntoPontos in bay:
                if (conjuntoPontos[0] == '8213a1'):
                    totalPontos = len(conjuntoPontos[1])    
                    for item in conjuntoPontos[1]:                    
                        # itens
                        planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                        # molduras
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                            else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                        pos = pos + 1 
                    break 
            if (pos==1):    # nÑo houve pontos para esse item
                for col in range(1,9):
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
                pos = pos +1                       
    
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.3.a.1', style_hor_neg_cent) 
            linhaTabelaSeq = linhaTabelaSeq + pos  
     
            # Escrita dos itens do requisito de rede 8.2.1.3.a.2 --------------------------------------------------------------------
            planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Disparo dos relÇs de bloqueio;', style_hor_neg_esq)
            for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)      
            pos = 1 
            totalPontos = 1   
            for conjuntoPontos in bay:
                if (conjuntoPontos[0] == '8213a2'):
                    totalPontos = len(conjuntoPontos[1])    
                    for item in conjuntoPontos[1]:                    
                        # itens
                        planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                        # molduras
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                            else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                        pos = pos + 1 
                    break 
            if (pos==1):    # nÑo houve pontos para esse item
                for col in range(1,9):
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
                pos = pos +1                       
    
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.3.a.2', style_hor_neg_cent) 
            linhaTabelaSeq = linhaTabelaSeq + pos   
       
            # Escrita dos itens do requisito de rede 8.2.1.3.b --------------------------------------------------------------------
            planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'AtuaáÑo da proteáÑo dos bancos de capacitores - FunáÑo Sobretemperatura (Agrupamento dos eventos abaixo relacionados):', style_hor_neg_esq)
            for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)        
            planilhaAserGravadaSeq.write(linhaTabelaSeq+1,1,u'Disparo da proteáÑo de desequil°brio de neutro;', style_hor_norm_dir)
            planilhaAserGravadaSeq.write(linhaTabelaSeq+2,1,u'Disparo da proteáÑo de sobrecorrente de fase e neutro;', style_hor_norm_dir)         
              
            for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)      
            pos = 1 
            totalPontos = 1   
            for conjuntoPontos in bay:
                if (conjuntoPontos[0] == '8213b'):
                    totalPontos = len(conjuntoPontos[1])    
                    for item in conjuntoPontos[1]:                    
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAserGravadaSeq.merge_range(linhaTabelaSeq+pos,col,linhaTabelaSeq+pos+1,col,item[col-2], style_hor_norm_cent)                     
                            else: planilhaAserGravadaSeq.merge_range(linhaTabelaSeq+pos,col,linhaTabelaSeq+pos+1,col,item[col-2], style_hor_norm_esq)                                   
                        pos = pos + 1 
                    break 
            if (pos==1):    # nÑo houve pontos para esse item
                for col in range(2,9):
                    planilhaAserGravadaSeq.merge_range(linhaTabelaSeq+pos,col,linhaTabelaSeq+pos+1,col, '', style_hor_norm_cent_verm) 
                pos = pos +1                       
    
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos+1,0,u'8.2.1.3.b', style_hor_neg_cent) 
            linhaTabelaSeq = linhaTabelaSeq + pos +1 
            
            # 04 - 8.2.1.8 - SEQUENCIA DE EVENTOS DISJUNTOR
            linhaTabelaSeq = linhaTabelaSeq + 1 
            nomeDisjuntor = bay[0]
            nomeDisjuntor = nomeDisjuntor.replace('0','1')
            linhaTabelaSeq = gravaSeqEventosDisjuntor(planilhaAserGravadaSeq, linhaTabelaSeq, nomeDisjuntor)                    
 
            
        elif(BCSerie):       
            # 01 - 8.2.1.8 - SEQUâNCIA DE EVENTOS BANCO CAPACITOR  SêRIE ************************************************************************
            linhaTabelaSeq = linhaTabelaSeq + 1         
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq,4,u'8.2.1.1 - SEQUâNCIA DE EVENTOS BANCO CAPACITOR SêRIE', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,5,linhaTabelaSeq+1,5,u'NãVEL 0', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,6,linhaTabelaSeq+1,8,u'NãVEL 3', style_hor_neg_cent_cinza)
            linhaTabelaSeq = linhaTabelaSeq + 1 
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+1,0,u'Item do SUBMODULO 2.7', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,1,linhaTabelaSeq+1,1,u'DescriáÑo 2.7', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,2,linhaTabelaSeq,4,u'Atendido', style_hor_neg_cent_cinza)
            linhaTabelaSeq = linhaTabelaSeq + 1
            planilhaAserGravadaSeq.write(linhaTabelaSeq,2,u'Sim', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,3,u'NÑo', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,4,u'NÑo se aplica', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,5,u'PONTO', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,6,u'Endereáo', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,7,u'Tipo ponto', style_hor_neg_cent_cinza)
            planilhaAserGravadaSeq.write(linhaTabelaSeq,8,u'ObservaáÑo', style_hor_neg_cent_cinza)
            linhaTabelaSeq = linhaTabelaSeq + 1  
              
            # Escrita dos itens do requisito de rede 8.2.1.11.a.1 --------------------------------------------------------------------
            planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Disparo da proteáÑo de sobretensÑo;', style_hor_neg_esq)
            for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)      
            pos = 1 
            totalPontos = 1   
            for conjuntoPontos in bay:
                if (conjuntoPontos[0] == '82111a1'):
                    totalPontos = len(conjuntoPontos[1])    
                    for item in conjuntoPontos[1]:                    
                        # itens
                        planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                        # molduras
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                            else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                        pos = pos + 1 
                    break 
            if (pos==1):    # nÑo houve pontos para esse item
                for col in range(1,9):
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
                pos = pos +1                       
    
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.11.a.1', style_hor_neg_cent) 
            linhaTabelaSeq = linhaTabelaSeq + pos  
     
            # Escrita dos itens do requisito de rede 8.2.1.11.a.2 --------------------------------------------------------------------
            planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'Disparo dos relÇs de bloqueio;', style_hor_neg_esq)
            for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq,col,'', style_hor_norm_esq)      
            pos = 1 
            totalPontos = 1   
            for conjuntoPontos in bay:
                if (conjuntoPontos[0] == '82111a2'):
                    totalPontos = len(conjuntoPontos[1])    
                    for item in conjuntoPontos[1]:                    
                        # itens
                        planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,1, item[3], style_hor_norm_dir)
                        # molduras
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, item[col-2], style_hor_norm_cent)                    
                            else: planilhaAserGravadaSeq.write(linhaTabelaSeq+pos, col, item[col-2], style_hor_norm_esq)                                    
                        pos = pos + 1 
                    break 
            if (pos==1):    # nÑo houve pontos para esse item
                for col in range(1,9):
                    planilhaAserGravadaSeq.write(linhaTabelaSeq+pos,col, '', style_hor_norm_cent_verm) 
                pos = pos +1                       
    
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+totalPontos,0,u'8.2.1.11.a.2', style_hor_neg_cent) 
            linhaTabelaSeq = linhaTabelaSeq + pos   
    
            # Escrita dos itens do requisito de rede 8.2.1.11.b ----------------------------------------------        
            planilhaAserGravadaSeq.merge_range(linhaTabelaSeq,0,linhaTabelaSeq+4,0,'8.2.1.11.b', style_hor_neg_cent) 
            planilhaAserGravadaSeq.write(linhaTabelaSeq,1,u'AtuaáÑo da proteáÑo dos bancos de capacitores - outras funáîes (Agrupamento dos eventos abaixo relacionados):', style_hor_neg_esq)
            for col in range(2,9): planilhaAserGravadaSeq.write(linhaTabelaSeq, col,'', style_hor_norm_esq)  
    
            planilhaAserGravadaSeq.write(linhaTabelaSeq+1,1,u'Disparo da proteáÑo de sub-harmìnicas;', style_hor_norm_dir)
            planilhaAserGravadaSeq.write(linhaTabelaSeq+2,1,u'Disparo da proteáÑo do centelhador;', style_hor_norm_dir)
            planilhaAserGravadaSeq.write(linhaTabelaSeq+3,1,u'Disparo da proteáÑo de desbalanáo de tensÑo;', style_hor_norm_dir)
            planilhaAserGravadaSeq.write(linhaTabelaSeq+4,1,u'Disparo da proteáÑo de fuga para a plataforma;', style_hor_norm_dir)
    
            pos = 1 
            totalPontos = 1   
            for conjuntoPontos in bay:
                if (conjuntoPontos[0] == '82111b'):
                    totalPontos = len(conjuntoPontos[1])    
                    for item in conjuntoPontos[1]:                    
                        # molduras
                        for col in range(2,9):
                            if col in [2,3,4,6,7]: planilhaAserGravadaSeq.merge_range(linhaTabelaSeq+pos,col,linhaTabelaSeq+pos+3,col, item[col-2], style_hor_norm_cent)                    
                            else:                  planilhaAserGravadaSeq.merge_range(linhaTabelaSeq+pos,col,linhaTabelaSeq+pos+3,col, item[col-2], style_hor_norm_esq)                                    
                        pos = pos + 1 
                    break 
            if (pos==1):    # nÑo houve pontos para esse item
                for col in range(2,9):
                    planilhaAserGravadaSeq.merge_range(linhaTabelaSeq+pos,col, linhaTabelaSeq+pos+3, col,'', style_hor_norm_cent_verm) 
                pos = pos +1 
                                                      
            linhaTabelaSeq = linhaTabelaSeq + 5   
            
        linhaTabelaSeq = linhaTabelaSeq + 1  
        return  linhaTabelaSeq

    
    nome_arq_saida = './PlanilhaONS_%s.xlsx'%(Codigo_SE)     #Nome do arquivo de sa°da
    seq_arq = 0                                             #Sequància do n£mero de arquivo
    while os.path.exists(nome_arq_saida):                   #Enquanto existir na pasta um arquivo com o nome definido
        seq_arq+=1                                          #Adicionar um a sequància do n£mero do arquivo
        nome_arq_saida = nome_arq_saida[0:13]+'_'+Codigo_SE+'_'+str(seq_arq)+'.xlsx' #Definir novo nome de arquivo (Ex './LP_gerada_JRM_1.xls)
 
    # Gera o arquivo ----------------------------------------------------------
    arquivoONS = xlsxwriter.Workbook(nome_arq_saida)
    planilhaONS = arquivoONS.add_worksheet('ONS 2.7')
    planilhaONS.set_zoom(60)     
 
    # =============================================================================
    # Definicao de estilos de formatacao 
    # =============================================================================
    style_hor_norm_dir = arquivoONS.add_format({
    'text_wrap':True,                                            
    'bold': False,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'right',
    'valign':'vcenter',
    'bg_color': '#FFFFFF',
    'border' : 1,
    })
    
    style_hor_norm_esq = arquivoONS.add_format({
    'text_wrap':True,
    'bold': False,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': '#FFFFFF',
    'border' : 1,
    })
    
    style_hor_norm_cent = arquivoONS.add_format({
    'text_wrap':True,    
    'bold': False,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#FFFFFF',
    'border' : 1,
    })   
    
    style_hor_neg_cent = arquivoONS.add_format({
    'text_wrap':True,
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#FFFFFF',
    'border' : 1,
    })    
    
    style_hor_neg_esq = arquivoONS.add_format({
    'text_wrap':True,
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'left',
    'valign':'vcenter',
    'bg_color': '#FFFFFF',
    'border' : 1,
    })
    
    style_hor_neg_cent_cinza = arquivoONS.add_format({
    'text_wrap':True,
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#A9A9A9',
    'border' : 1,
    })      
    
    style_hor_neg_cent_azul = arquivoONS.add_format({
    'text_wrap':True,
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#87CEEB',
    'border' : 1,
    }) 
       
    style_hor_norm_cent_verm = arquivoONS.add_format({
    'text_wrap':True,
    'bold': True,
    'font_name':'Arial',
    'font_size':12,
    'rotation': 0,
    'align':'center',
    'valign':'vcenter',
    'bg_color': '#FF0000',
    'border' : 1,
    })        
       
    # INICIO DE ESCRITA DA PLANILHA ONS************************************************************************
    planilhaONS.write(0,0,"SE")
    planilhaONS.write(1,0,"EVENTO")
    planilhaONS.write(2,0,"DATA PROJETO")    
    linhaAtual = 4
       
    # Ajuste da largura das colunas/linhas ----------------------------------------
    for i in range (0,9):
        planilhaONS.set_column(i, i,int(largura[i]))
            
    #linha = linhaAtual   
    for evento in listaEventosEPontos:
        # =================================================================================================================
        # LINHA DE TRANSMISSAO
        # =================================================================================================================         
        if (listaEventosEPontos.index(evento) == 1):         
            for bay in evento:    
                linhaAtual = gravaLT_ONS(linhaAtual, planilhaONS, bay)
              
        # =================================================================================================================
        # TRANSFORMADORES
        # =================================================================================================================         
        elif (listaEventosEPontos.index(evento) == 2):
            for bay in evento:
                linhaAtual = gravaTR_ONS(linhaAtual, planilhaONS, bay)
                                                   
        # =================================================================================================================
        # BAY DE TRANSFERâNCIA
        # =================================================================================================================         
        if (listaEventosEPontos.index(evento) == 3):         
            for bay in evento:    
                linhaAtual = gravaBT_ONS(linhaAtual, planilhaONS, bay)  
        
        # =================================================================================================================
        # REATOR
        # =================================================================================================================         
        if (listaEventosEPontos.index(evento) == 4):                      
            for bay in evento:    
                linhaAtual = gravaRE_ONS(linhaAtual, planilhaONS, bay)  
                            
        # =================================================================================================================
        # BARRAS
        # =================================================================================================================         
        if (listaEventosEPontos.index(evento) == 5):                      
            for bay in evento:    
                linhaAtual = gravaBA_ONS(linhaAtual, planilhaONS, bay)
             
        # =================================================================================================================
        # BANCO DE CAPACITORES
        # =================================================================================================================         
        if (listaEventosEPontos.index(evento) == 6):                      
            for bay in evento:   
                linhaAtual = gravaBC_ONS(linhaAtual, planilhaONS, bay)
                
    '''                
    # GravaáÑo da aba de log de geraáÑo -----------------------------------------------------------------------------------
    planilhaFalhas = arquivo.add_sheet('LOG')
    planilhaFalhas.write_merge(0,0,0,2,u'PONTOS INCONSISTENTES')
    planilhaFalhas.write(1,0,u"ID SAGE")
    planilhaFalhas.write(1,1,u"ITEM PR 2.7")
    planilhaFalhas.write(1,2,u"DESCRIÄéO INCONSISTâNCIA")
    linha = 2
    for evento in listaDeFalhas:
        for falhas in evento:
            for itemDaFalha in falhas:
                planilhaFalhas.write(linha,0,itemDaFalha[0])
                planilhaFalhas.write(linha,1,itemDaFalha[1])
                planilhaFalhas.write(linha,2,itemDaFalha[2])
                linha = linha + 1                                    
    
    
    nome_arq_saida = './PlanilhaONS_%s.xls'%(Codigo_SE)       #Nome do arquivo de sa°da
    seq_arq = 0                               #Sequància do n£mero de arquivo
    while os.path.exists(nome_arq_saida):   #Enquanto existir na pasta um arquivo com o nome definido
        seq_arq+=1                          #Adicionar um a sequància do n£mero do arquivo
        nome_arq_saida = nome_arq_saida[0:13]+'_'+Codigo_SE+'_'+str(seq_arq)+'.xls' #Definir novo nome de arquivo (Ex './LP_gerada_JRM_1.xls)
    #arq_LP.save(nome_arq_saida[2:])         #Gravar o nome do arquivo excluindo './' do nome    
            
    arquivo.save(nome_arq_saida)         #Gravar o nome do arquivo excluindo do nome
    '''
    
    # INICIO DE ESCRITA DA PLANILHA RELATORIO ************************************************************************
    planilhaRelatorioONS = arquivoONS.add_worksheet('RELATORIO')    
    planilhaRelatorioONS.set_zoom(100)     

    planilhaRelatorioONS.write(0,0,u'ID SAGE')
    planilhaRelatorioONS.write(0,1,u'COMENTÜRIO')
    planilhaRelatorioONS.write(0,2,'ITEM PR 2.7')
    planilhaRelatorioONS.write(0,3,'DESCRIÄéO INCONSISTâNCIA')            
    linhaAtualRelatorio = 1
      
    for evento in listaDeFalhas:
        for bay in evento:
            for ponto in bay:
                planilhaRelatorioONS.write(linhaAtualRelatorio,0,ponto[0])
                planilhaRelatorioONS.write(linhaAtualRelatorio,1,ponto[1])
                planilhaRelatorioONS.write(linhaAtualRelatorio,2,ponto[2])
                planilhaRelatorioONS.write(linhaAtualRelatorio,3,ponto[3])
                linhaAtualRelatorio= linhaAtualRelatorio + 1
    # retorna o arquivo formatado ----------------------------------------------
    arquivoONS.close()
    return nome_arq_saida
    
  
