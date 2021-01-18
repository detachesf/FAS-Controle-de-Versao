import gi
import winsound
import threading

gi.require_version("Gtk", "3.0")
from gi.repository import Gtk

builder = Gtk.Builder()  # inicia o construtor Gtk
builder.add_from_file('user_interface.glade')

from os import path, startfile, listdir, popen, getcwd
from sys import stdout
from traceback import print_exc
import pickle  # serve para armazenar objetos e variáveis em arquivos
import re


# Caixa de diálogo

def mensagem_erro(titulo, msg):
    mensagem_erro: Gtk.MessageDialog = builder.get_object('message_erro')
    mensagem_erro.props.text = titulo
    mensagem_erro.set_title('Erro')
    mensagem_erro.props.secondary_text = msg
    mensagem_erro.props.icon_name = 'dialog-error-symbolic'
    mensagem_erro.show_all()
    mensagem_erro.run()
    mensagem_erro.hide()


def mensagem_aviso(titulo, msg):
    message_aviso: Gtk.MessageDialog = builder.get_object('message_aviso')
    message_aviso.props.text = titulo
    message_aviso.props.secondary_text = msg
    message_aviso.props.icon_name = 'dialog-warning-symbolic'
    message_aviso.show_all()
    message_aviso.run()
    message_aviso.hide()


def pergunta_sim_nao(titulo, msg):
    perguntasimnao: Gtk.MessageDialog = builder.get_object('pergunta_sim_nao')
    perguntasimnao.props.text = titulo
    perguntasimnao.props.secondary_text = msg
    perguntasimnao.props.icon_name = 'dialog-question-symbolic'
    perguntasimnao.show_all()
    resposta = perguntasimnao.run()
    perguntasimnao.hide()
    if resposta == Gtk.ResponseType.YES:
        return True
    elif resposta == Gtk.ResponseType.NO:
        return False


try:
    from xlrd import open_workbook
except:
    mensagem_erro('Erro', 'Módulo xlrd não instalado')

try:
    from lp_lib.Gerar_LP import gerar
except:
    mensagem_erro('Erro', 'Módulo Gerar_LP não instalado')

try:
    from lp_lib.Checar_LP import checar
except:
    mensagem_erro('Erro', 'Módulo Checar_LP não instalado')

try:
    from lp_lib.func import processing
except:
    mensagem_erro('Erro', 'Módulo func não instalado')

try:
    from lp_lib.func import createToolTip
except:
    mensagem_erro('Erro',
                  'Arquivo "func.pyc" deve estar no diretório "lp_lib"')


class Manipulador(object):

    def __init__(self):
        # Vairáveis Gerais
        self.arqconf_novo = True
        self.pathchecar = getcwd
        self.versao = '2.0.12'
        self.data = '10/11/2020'

        # Arrays com os nomes padrão dos objetos de cada linha
        self.NotbkLT_Linha = ['selec_linha_LT_', 'LT_entry_codlinha_', 'LT_entry_codpainel_', 'LT_entry_ltremota_',
                              'LT_entry_camarapass_',
                              'LT_entry_conjuntosecc_', 'LT_combobox_arranjo_', 'LT_combobox_religamento_',
                              'LT_checkbtt_rdp_', 'LT_checkbtt_painelteleprot_', 'LT_checkbtt_f9_', 'LT_checkbtt_87l_']

        self.NotbkTrafo_Linha = ['selec_linha_Trafo_', 'Trafo_entry_codtrafo_', 'Trafo_entry_codpainelH_',
                                 'Trafo_entry_codpainelX_',
                                 'Trafo_entry_camarapass_',
                                 'Trafo_entry_conjuntosecc_', 'Trafo_combobox_arranjoH_', 'Trafo_combobox_arranjoX_',
                                 'Trafo_checkbtt_rdp_', 'Trafo_checkbtt_regapp_', 'Trafo_checkbtt_f9_',
                                 'Trafo_combobox_equip_', 'Trafo_combobox_relacao_', 'Trafo_combobox_prot_',
                                 ]
        self.NotbkVaoTrans_Linha = ['selec_linha_vaotrans_','vaotrans_entry_cod_','vaotrans_entry_painel_',
                                    'vaotrans_checkbtt_87B_','vaotrans_combobox_arranjo_','vaotrans_entry_campass_',
                                    'vaotrans_entry_conjsecc_']

        self.NotbkPaisage_Linha = ['selec_linha_paisage_','paisage_entry_painel_','paisage_combobox_sagebastidor_',
                                   'paisage_entry_sw_de_', 'paisage_entry_sw_ate_', 'paisage_entry_nportas_sw_',
                                   'paisage_checkbtt_fw_', 'paisage_entry_nporta_fw_','paisage_checkbtt_rb_', 'paisage_entry_rb_de_',
                                   'paisage_entry_rb_ate_', 'paisage_entry_nporta_rb_']

        self.NotbkReator_Linha = ['selec_linha_reator_', 'reator_entry_cod_', 'reator_entry_painel_',
                                  'reator_checkbtt_manob_', 'reator_combobox_equip_', 'reator_checkbtt_rdp_', 'reator_checkbtt_bunitf9_',
                                  'reator_entry_campass_', 'reator_entry_conjuntosecc_']

        self.NotbkAcesso_Linha = ['selec_linha_acesso_', 'acesso_entry_codvao_','acesso_entry_painelacess_','acesso_checkbtt_painelexist_',
                                  'acesso_entry_num_uc_chesf_','acesso_entry_num_uc_acessante_','acesso_combobox_arranjo_', 'acesso_checkbtt_ts_',
                                  'acesso_entry_ts_de_', 'acesso_entry_ts_ate_', 'acesso_checkbtt_rb_', 'acesso_entry_redbox_de_',
                                  'acesso_entry_redbox_ate_', 'acesso_checkbtt_multimedidor_', 'acesso_entry_mm_de_', 'acesso_entry_mm_ate_',
                                  'acesso_entry_lt_remota_']
        self.NotbkTterra_Linha = ['selec_linha_tterra_', 'tterra_entry_codigo_', 'tterra_entry_painel_','tterra_entry_camaraspass_', 'tterra_entry_conjuntosecc_']
        self.NotbkProtbarra_Linha = ['selec_linha_protbarra_','protbarra_entry_painel_', 'protbarra_entry_qtpan_','protbarra_combobox_arranjo_',
                                     'protbarra_checkbtt_bu_no_painel_','protbarra_entry_vaos_']

        self.NotbkBcapshunt_Linha = ['selec_linha_bcapshunt_','bcapshunt_entry_codigo_','bcapshunt_entry_painel_','bcapshunt_combobox_arranjo_',
                                     'bcapshunt_checkbtt_rdp_','bcapshunt_checkbtt_bunitf9_']

        self.NotbkBcapserie_Linha = ['selec_linha_bcapserie_','bcapserie_entry_codigo_','bcapserie_entry_painel_']

        self.NotbkEce_Linha = ['selec_linha_ece_', 'ece_entry_codigo_', 'ece_entry_painel_']

        self.NotbkSistreg_Linha = ['selec_linha_sistreg_','sistreg_combobox_nome_','sistreg_combobox_tesao_reg_',
                                   'sistreg_entry_painel_']

        self.NotbkPrepreen_Linha = ['selec_linha_prepreen_', 'prepreen_entry_sistema_']

        self.NotbkCompsinc_Linha = ['selec_linha_compsinc_', 'compsinc_entry_codigo_', 'compsinc_entry_painel_']

        # Variáveis Auxiliares na mecânica da tela de configuração

        self.NotbkLT_Linha_dic = {}  # dicionário para armazenar os objetos adicionados dinâmicamente
        self.NotbkTrafo_Linha_dic = {}
        self.NotbkVaoTrans_Linha_dic = {}
        self.NotbkPaisage_Linha_dic = {}
        self.NotbkReator_Linha_dic = {}
        self.NotbkAcesso_Linha_dic = {}
        self.NotbkTterra_Linha_dic = {}
        self.NotbkProtbarra_Linha_dic = {}
        self.NotbkBcapshunt_Linha_dic = {}
        self.NotbkBcapserie_Linha_dic = {}
        self.NotbkEce_Linha_dic = {}
        self.NotbkSistreg_Linha_dic = {}
        self.NotbkPrepreen_Linha_dic = {}
        self.NotbkCompsinc_Linha_dic = {}

        self.Arranjos = ['DISJ E MEIO', 'BS', 'BPT', 'BD3',
                         'BD4']  # Array com os arranjos possíveis para preencher os comboboxes

        self.Num_de_LT = [1]  # Variável que armazena o número das linhas ativas
        self.Num_de_Trafo = [1]
        self.Num_de_VaoTrans = [1]
        self.Num_de_Paisage = [1]
        self.Num_de_Reator = [1]
        self.Num_de_Acesso = [1]
        self.Num_de_Tterra = [1]
        self.Num_de_Protbarra = [1]
        self.Num_de_Bcapshunt = [1]
        self.Num_de_Bcapserie = [1]
        self.Num_de_Ece = [1]
        self.Num_de_Sistreg = [1]
        self.Num_de_Prepreen = [1]
        self.Num_de_Compsinc = [1]

        self.Linhas_Removidas_LT = []  # Variável que registra as linhas que foram removidas
        self.Linhas_Removidas_Trafo = []
        self.Linhas_Removidas_VaoTrans = []
        self.Linhas_Removidas_Paisage = []
        self.Linhas_Removidas_Reator = []
        self.Linhas_Removidas_Acesso = []
        self.Linhas_Removidas_Tterra = []
        self.Linhas_Removidas_Protbarra = []
        self.Linhas_Removidas_Bcapshunt = []
        self.Linhas_Removidas_Bcapserie = []
        self.Linhas_Removidas_Ece = []
        self.Linhas_Removidas_Sistreg = []
        self.Linhas_Removidas_Prepreen = []
        self.Linhas_Removidas_Compsinc = []

        # Carregando objetos

        self.janela_sobre: Gtk.AboutDialog = builder.get_object('janela_sobre')
        self.dialogo_diretorio: Gtk.FileChooserDialog = builder.get_object('diretorio_dialogo')
        self.diretorio_dialogo_pasta_entry: Gtk.Entry = builder.get_object('diretorio_dialogo_pasta_entry')

        self.tabela_LT: Gtk.Table = builder.get_object('tabela_LT')
        self.tabela_Trafo: Gtk.Table = builder.get_object('tabela_trafo')
        self.tabela_VaoTrans: Gtk.Table = builder.get_object('tabela_vaotransf')
        self.tabela_Paisage: Gtk.Table = builder.get_object('tabela_painel_sage')
        self.tabela_Reator: Gtk.Table = builder.get_object('tabela_reator')
        self.tabela_Acesso: Gtk.Table = builder.get_object('tabela_acesso')
        self.tabela_Tterra: Gtk.Table = builder.get_object('tabela_tterra')
        self.tabela_Protbarra: Gtk.Table = builder.get_object('tabela_protbarra')
        self.tabela_Bcapshunt: Gtk.Table = builder.get_object('tabela_bcapshunt')
        self.tabela_Bcapserie: Gtk.Table = builder.get_object('tabela_bcapserie')
        self.tabela_Ece: Gtk.Table = builder.get_object('tabela_ece')
        self.tabela_Sistreg: Gtk.Table = builder.get_object('tabela_sist_reg')
        self.tabela_Prepreen: Gtk.Table = builder.get_object('tabela_prepreen')
        self.tabela_Compsinc: Gtk.Table = builder.get_object('tabela_compsinc')

        self.notebook: Gtk.Notebook = builder.get_object('notebook1')

    def on_janela_principal_destroy(self, window):
        Gtk.main_quit()  # Encerra a aplicação quando fechar a janela no X vermelho

    # Sinais de navegação entre páginas

    # Janela de Sobre
    def on_arqconf_menubar_ajuda_sobre_activate(self, window):

        self.janela_sobre.set_version(self.versao)
        self.janela_sobre.show_all()
        resposta = self.janela_sobre.run()
        if resposta == -4:
            self.janela_sobre.hide()

    # Sinais de lógica na tela

    # Sinais da tela do evento LT

    # Ações executadas quando o botão adicionar for clicado
    def on_button_add_linha_clicked(self, button):

        Aba = self.notebook.get_current_page() #captura a aba ativa

        if Aba == 0: #Aba da LT
            self.adicionar_linha(self.Linhas_Removidas_LT, self.Num_de_LT, 'LT', self.NotbkLT_Linha,
                                 self.NotbkLT_Linha_dic,
                                 self.tabela_LT)
        elif Aba == 1: #Aba do Trafo
            self.adicionar_linha(self.Linhas_Removidas_Trafo, self.Num_de_Trafo, 'Trafo', self.NotbkTrafo_Linha,
                                 self.NotbkTrafo_Linha_dic, self.tabela_Trafo)
        elif Aba == 2: #Aba do Vão de Transferência
            self.adicionar_linha(self.Linhas_Removidas_VaoTrans, self.Num_de_VaoTrans, 'VaoTrans',
                                 self.NotbkVaoTrans_Linha, self.NotbkVaoTrans_Linha_dic, self.tabela_VaoTrans)
        elif Aba == 3: #Aba do Painel Sage e Bastidor de Rede
            self.adicionar_linha(self.Linhas_Removidas_Paisage, self.Num_de_Paisage, 'Paisage',
                                 self.NotbkPaisage_Linha, self.NotbkPaisage_Linha_dic, self.tabela_Paisage)
        elif Aba == 4: #Aba do Reator
            self.adicionar_linha(self.Linhas_Removidas_Reator, self.Num_de_Reator, 'Reator',
                                 self.NotbkReator_Linha, self.NotbkReator_Linha_dic, self.tabela_Reator)
        elif Aba == 5: #Aba do Acesso Segregado
            self.adicionar_linha(self.Linhas_Removidas_Acesso, self.Num_de_Acesso, 'Acesso',
                                 self.NotbkAcesso_Linha, self.NotbkAcesso_Linha_dic, self.tabela_Acesso)
        elif Aba == 6: #Aba do Trafo Terra
            self.adicionar_linha(self.Linhas_Removidas_Tterra, self.Num_de_Tterra, 'Tterra',
                                 self.NotbkTterra_Linha, self.NotbkTterra_Linha_dic, self.tabela_Tterra)
        elif Aba == 7:  # Aba de Proteção de Barra
            self.adicionar_linha(self.Linhas_Removidas_Protbarra, self.Num_de_Protbarra, 'Protbarra',
                                 self.NotbkProtbarra_Linha, self.NotbkProtbarra_Linha_dic, self.tabela_Protbarra)
        elif Aba == 8:  # Aba do Banco de Capacitores shunt
            self.adicionar_linha(self.Linhas_Removidas_Bcapshunt, self.Num_de_Bcapshunt, 'Bcapshunt',
                                 self.NotbkBcapshunt_Linha, self.NotbkBcapshunt_Linha_dic, self.tabela_Bcapshunt)
        elif Aba == 9:  # Aba do Banco de Capacitores série
            self.adicionar_linha(self.Linhas_Removidas_Bcapserie, self.Num_de_Bcapserie, 'Bcapserie',
                                 self.NotbkBcapserie_Linha, self.NotbkBcapserie_Linha_dic, self.tabela_Bcapserie)
        elif Aba == 10: # Aba do ECE
            self.adicionar_linha(self.Linhas_Removidas_Ece, self.Num_de_Ece, 'Ece',
                                 self.NotbkEce_Linha, self.NotbkEce_Linha_dic, self.tabela_Ece)
        elif Aba == 11: # Aba do Sistema de Regulação
            self.adicionar_linha(self.Linhas_Removidas_Sistreg, self.Num_de_Sistreg, 'Sistreg',
                                 self.NotbkSistreg_Linha, self.NotbkSistreg_Linha_dic, self.tabela_Sistreg)
        elif Aba == 12: #Aba de Preparação para Reenergização
            self.adicionar_linha(self.Linhas_Removidas_Prepreen, self.Num_de_Prepreen, 'Prepreen',
                                 self.NotbkPrepreen_Linha, self.NotbkPrepreen_Linha_dic, self.tabela_Prepreen)
        elif Aba == 13: #Aba do Compensador Síncrono
            self.adicionar_linha(self.Linhas_Removidas_Compsinc, self.Num_de_Compsinc, 'Compsinc',
                                 self.NotbkCompsinc_Linha, self.NotbkCompsinc_Linha_dic, self.tabela_Compsinc)
    # Ações executadas quando o botão excluir for clicado
    def on_button_Excluir_clicked(self, button):

        Aba = self.notebook.get_current_page()  # captura a aba ativa

        if Aba == 0:  # Aba da LT
            self.exclui_linha(self.Linhas_Removidas_LT, self.Num_de_LT, self.NotbkLT_Linha, self.NotbkLT_Linha_dic)
        elif Aba == 1:  # Aba do Trafo
            self.exclui_linha(self.Linhas_Removidas_Trafo, self.Num_de_Trafo, self.NotbkTrafo_Linha,self.NotbkTrafo_Linha_dic)
        elif Aba == 2:  # Aba do Vão de Transferência
            self.exclui_linha(self.Linhas_Removidas_VaoTrans, self.Num_de_VaoTrans,
                                 self.NotbkVaoTrans_Linha, self.NotbkVaoTrans_Linha_dic)
        elif Aba == 3:  # Aba do Painel Sage e Bastidor de Rede
            self.exclui_linha(self.Linhas_Removidas_Paisage, self.Num_de_Paisage,
                                 self.NotbkPaisage_Linha, self.NotbkPaisage_Linha_dic)
        elif Aba == 4:  # Aba do Reator
            self.exclui_linha(self.Linhas_Removidas_Reator, self.Num_de_Reator,
                                 self.NotbkReator_Linha, self.NotbkReator_Linha_dic)
        elif Aba == 5:  # Aba do Acesso Segregado
            self.exclui_linha(self.Linhas_Removidas_Acesso, self.Num_de_Acesso,
                                 self.NotbkAcesso_Linha, self.NotbkAcesso_Linha_dic)
        elif Aba == 6:  # Aba do Trafo Terra
            self.exclui_linha(self.Linhas_Removidas_Tterra, self.Num_de_Tterra,
                              self.NotbkTterra_Linha, self.NotbkTterra_Linha_dic)
        elif Aba == 7:  # Aba de Proteção de Barra
            self.exclui_linha(self.Linhas_Removidas_Protbarra, self.Num_de_Protbarra,
                              self.NotbkProtbarra_Linha, self.NotbkProtbarra_Linha_dic)
        elif Aba == 8:  # Aba do Banco de Capacitores shunt
            self.exclui_linha(self.Linhas_Removidas_Bcapshunt, self.Num_de_Bcapshunt,
                                 self.NotbkBcapshunt_Linha, self.NotbkBcapshunt_Linha_dic)
        elif Aba == 9:  # Aba do Banco de Capacitores série
            self.exclui_linha(self.Linhas_Removidas_Bcapserie, self.Num_de_Bcapserie,
                              self.NotbkBcapserie_Linha, self.NotbkBcapserie_Linha_dic)

        elif Aba == 10:  # Aba do ECE
            self.exclui_linha(self.Linhas_Removidas_Ece, self.Num_de_Ece,
                                 self.NotbkEce_Linha, self.NotbkEce_Linha_dic)

        elif Aba == 11: # Aba do Sistema de Regulação
            self.exclui_linha(self.Linhas_Removidas_Sistreg, self.Num_de_Sistreg,
                                 self.NotbkSistreg_Linha, self.NotbkSistreg_Linha_dic)

        elif Aba == 12: #Aba de Preparação para Reenergização
            self.exclui_linha(self.Linhas_Removidas_Prepreen, self.Num_de_Prepreen,
                                 self.NotbkPrepreen_Linha, self.NotbkPrepreen_Linha_dic)
        elif Aba == 13: #Aba do Compensador Síncrono
            self.exclui_linha(self.Linhas_Removidas_Compsinc, self.Num_de_Compsinc,
                                 self.NotbkCompsinc_Linha, self.NotbkCompsinc_Linha_dic)

    # Ações executadas quando o botão duplicar for clicado
    def on_button_duplicar_clicked(self, button):

        Aba = self.notebook.get_current_page()  # captura a aba ativa

        if Aba == 0:  # Aba da LT
            self.prepara_para_duplicar(self.Linhas_Removidas_LT, self.Num_de_LT, 'LT', self.NotbkLT_Linha,
                                       self.NotbkLT_Linha_dic, self.tabela_LT)
        elif Aba == 1:  # Aba do Trafo
            self.prepara_para_duplicar(self.Linhas_Removidas_Trafo, self.Num_de_Trafo, 'Trafo', self.NotbkTrafo_Linha,
                                       self.NotbkTrafo_Linha_dic, self.tabela_Trafo)
        elif Aba == 2:  # Aba do Vão de Transferência
            self.prepara_para_duplicar(self.Linhas_Removidas_VaoTrans, self.Num_de_VaoTrans, 'VaoTrans',
                                       self.NotbkVaoTrans_Linha, self.NotbkVaoTrans_Linha_dic, self.tabela_VaoTrans)
        elif Aba == 3:  # Aba do Painel Sage e Bastidor de Rede
            self.prepara_para_duplicar(self.Linhas_Removidas_Paisage, self.Num_de_Paisage, 'Paisage',
                                       self.NotbkPaisage_Linha, self.NotbkPaisage_Linha_dic, self.tabela_Paisage)
        elif Aba == 4:  # Aba do Reator
            self.prepara_para_duplicar(self.Linhas_Removidas_Reator, self.Num_de_Reator, 'Reator',
                                       self.NotbkReator_Linha, self.NotbkReator_Linha_dic, self.tabela_Reator)
        elif Aba == 5:  # Aba do Acesso Segregado
            self.prepara_para_duplicar(self.Linhas_Removidas_Acesso, self.Num_de_Acesso, 'Acesso',
                                 self.NotbkAcesso_Linha, self.NotbkAcesso_Linha_dic, self.tabela_Acesso)
        elif Aba == 6:  # Aba do Trafo Terra
            self.prepara_para_duplicar(self.Linhas_Removidas_Tterra, self.Num_de_Tterra, 'Tterra',
                                       self.NotbkTterra_Linha, self.NotbkTterra_Linha_dic, self.tabela_Tterra)
        elif Aba == 7:  # Aba de Proteção de Barra
            self.prepara_para_duplicar(self.Linhas_Removidas_Protbarra, self.Num_de_Protbarra, 'Protbarra',
                                 self.NotbkProtbarra_Linha, self.NotbkProtbarra_Linha_dic, self.tabela_Protbarra)
        elif Aba == 8:  # Aba do Banco de Capacitores shunt
            self.prepara_para_duplicar(self.Linhas_Removidas_Bcapshunt, self.Num_de_Bcapshunt, 'Bcapshunt',
                                 self.NotbkBcapshunt_Linha, self.NotbkBcapshunt_Linha_dic, self.tabela_Bcapshunt)
        elif Aba == 9:  # Aba do Banco de Capacitores série
            self.prepara_para_duplicar(self.Linhas_Removidas_Bcapserie, self.Num_de_Bcapserie, 'Bcapserie',
                                 self.NotbkBcapserie_Linha, self.NotbkBcapserie_Linha_dic, self.tabela_Bcapserie)
        elif Aba == 10:  # Aba do ECE
            self.prepara_para_duplicar(self.Linhas_Removidas_Ece, self.Num_de_Ece, 'Ece',
                                 self.NotbkEce_Linha, self.NotbkEce_Linha_dic, self.tabela_Ece)
        elif Aba == 11: # Aba do Sistema de Regulação
            self.prepara_para_duplicar(self.Linhas_Removidas_Sistreg, self.Num_de_Sistreg, 'Sistreg',
                                 self.NotbkSistreg_Linha, self.NotbkSistreg_Linha_dic, self.tabela_Sistreg)
        elif Aba == 12: #Aba de Preparação para Reenergização
            self.prepara_para_duplicar(self.Linhas_Removidas_Prepreen, self.Num_de_Prepreen, 'Prepreen',
                                 self.NotbkPrepreen_Linha, self.NotbkPrepreen_Linha_dic, self.tabela_Prepreen)
        elif Aba == 13: #Aba do Compensador Síncrono
            self.prepara_para_duplicar(self.Linhas_Removidas_Compsinc, self.Num_de_Compsinc, 'Compsinc',
                                 self.NotbkCompsinc_Linha, self.NotbkCompsinc_Linha_dic, self.tabela_Compsinc)

    # Ações executadas quando o botão limpar for clicado

    def on_button_limpar_clicked(self, button):

        Aba = self.notebook.get_current_page()  # captura a aba ativa

        if Aba == 0:  # Aba da LT
            self.limpar_linha(self.Num_de_LT, self.NotbkLT_Linha, self.NotbkLT_Linha_dic)

        elif Aba == 1:  # Aba do Trafo
            self.limpar_linha( self.Num_de_Trafo, self.NotbkTrafo_Linha, self.NotbkTrafo_Linha_dic)

        elif Aba == 2:  # Aba do Vão de Transferência
            self.limpar_linha(self.Num_de_VaoTrans, self.NotbkVaoTrans_Linha, self.NotbkVaoTrans_Linha_dic)

        elif Aba == 3:  # Aba do Painel Sage e Bastidor de Rede
            self.limpar_linha(self.Num_de_Paisage, self.NotbkPaisage_Linha, self.NotbkPaisage_Linha_dic)

        elif Aba == 4:  # Aba do Reator
            self.limpar_linha(self.Num_de_Reator, self.NotbkReator_Linha, self.NotbkReator_Linha_dic)

        elif Aba == 5:  # Aba do Acesso Segregado
            self.limpar_linha(self.Num_de_Acesso, self.NotbkAcesso_Linha, self.NotbkAcesso_Linha_dic)

        elif Aba == 6:  # Aba do Trafo Terra
            self.limpar_linha(self.Num_de_Tterra, self.NotbkTterra_Linha, self.NotbkTterra_Linha_dic)

        elif Aba == 7:  # Aba de Proteção de Barra
            self.limpar_linha(self.Num_de_Protbarra, self.NotbkProtbarra_Linha, self.NotbkProtbarra_Linha_dic)

        elif Aba == 8:  # Aba do Banco de Capacitores shunt
            self.limpar_linha(self.Num_de_Bcapshunt, self.NotbkBcapshunt_Linha, self.NotbkBcapshunt_Linha_dic)

        elif Aba == 9:  # Aba do Banco de Capacitores série
            self.limpar_linha(self.Num_de_Bcapserie, self.NotbkBcapserie_Linha, self.NotbkBcapserie_Linha_dic)
        elif Aba == 10:  # Aba do ECE
            self.limpar_linha(self.Num_de_Ece, self.NotbkEce_Linha, self.NotbkEce_Linha_dic)
        elif Aba == 11: # Aba do Sistema de Regulação
            self.limpar_linha(self.Num_de_Sistreg, self.NotbkSistreg_Linha, self.NotbkSistreg_Linha_dic)
        elif Aba == 12: #Aba de Preparação para Reenergização
            self.limpar_linha(self.Num_de_Prepreen, self.NotbkPrepreen_Linha, self.NotbkPrepreen_Linha_dic)
        elif Aba == 13: #Aba do Compensador Síncrono
            self.limpar_linha(self.Num_de_Compsinc, self.NotbkCompsinc_Linha, self.NotbkCompsinc_Linha_dic)


    # Funções de ação gerais

    # Função para adicionar uma linha
    def adicionar_linha(self, Linhas_Removidas, Numero_linhas_ativas, tipo_evento, array_nomes_objetos,
                        dicionario_objetos, tabela_evento):
        if tipo_evento == 'LT':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.ComboBoxText(),
                             Gtk.ComboBoxText(),
                             Gtk.CheckButton(),
                             Gtk.CheckButton(),
                             Gtk.CheckButton(),
                             Gtk.CheckButton()]  # Array que cria novos objetos do evento LT na sequência da tela
        elif tipo_evento == 'Trafo':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.ComboBoxText(),
                             Gtk.ComboBoxText(),
                             Gtk.CheckButton(),
                             Gtk.CheckButton(),
                             Gtk.CheckButton(),
                             Gtk.ComboBoxText(),
                             Gtk.ComboBoxText(),
                             Gtk.ComboBoxText()]
        elif tipo_evento == 'VaoTrans':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.CheckButton(),
                             Gtk.ComboBoxText(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             ]
        elif tipo_evento == 'Paisage':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.ComboBoxText(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry()]
        elif tipo_evento == 'Reator':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.CheckButton(),
                             Gtk.ComboBoxText(),
                             Gtk.CheckButton(),
                             Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry()]

        elif tipo_evento == 'Acesso':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.ComboBoxText(),
                             Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry()]
        elif tipo_evento == 'Tterra':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.Entry()]
        elif tipo_evento == 'Protbarra':
            array_objetos = [Gtk.CheckButton(),
                         Gtk.Entry(),
                         Gtk.Entry(),
                         Gtk.ComboBoxText(),
                         Gtk.CheckButton(),
                         Gtk.Entry()]

        elif tipo_evento == 'Bcapshunt':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry(),
                             Gtk.ComboBoxText(),
                             Gtk.CheckButton(),
                             Gtk.CheckButton()]

        elif tipo_evento == 'Bcapserie':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry()]

        elif tipo_evento == 'Ece':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry()]

        elif tipo_evento == 'Sistreg':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.ComboBoxText(),
                             Gtk.ComboBoxText(),
                             Gtk.Entry()]

        elif tipo_evento == 'Prepreen':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry()]

        elif tipo_evento == 'Compsinc':
            array_objetos = [Gtk.CheckButton(),
                             Gtk.Entry(),
                             Gtk.Entry()]

        if not Linhas_Removidas:  # identifica se já foi removida alguma linha anteriormente
            # Caso nenhuma linha tenha sido removida, a nova linha a adicionar será uma a mais da ultima
            indice_a_adicionar = max(Numero_linhas_ativas, key=int) + 1
            Numero_linhas_ativas.append(indice_a_adicionar)
        else:
            # Caso já tenha sido removida alguma linha, a linha a adicionar será a de índice menor dentre as removidas
            indice_a_adicionar = min(Linhas_Removidas, key=int)
            Numero_linhas_ativas.append(indice_a_adicionar)
            del Linhas_Removidas[Linhas_Removidas.index(indice_a_adicionar)]
        for i in range(0, len(array_nomes_objetos)):  # For para tratar cada objeto da linha
            objeto = array_objetos[i]
            objeto.set_name(array_nomes_objetos[i] + str(
                indice_a_adicionar))  # Seta o nome do objeto para obedecer o padrão, junto ao número sequencial da linha
            objeto.props.visible = True  # Faz o objeto ficar visível na tela
            dicionario_objetos[
                objeto.get_name()] = objeto  # Armazena o objeto no dicionário para que seja possível acessá-lo posteriormente
            tabela_evento.attach(objeto, i, i + 1, indice_a_adicionar + 1,
                                 indice_a_adicionar + 2)  # organiza os objetos na tela, dentro do objeto tabela

            # Configurações adicionais a objetos específicos
            if objeto.get_name().__contains__('selec_linha') or objeto.get_name().__contains__('checkbtt'):
                objeto.set_halign(Gtk.Align.CENTER)
                objeto.set_valign(Gtk.Align.CENTER)
            if objeto.get_name().__contains__('arranjo'):
                self.preenche_arranjo(objeto)

            if tipo_evento == 'LT':
                if objeto.get_name().__contains__('religamento') and tipo_evento == 'LT':
                    objeto.append_text('MONO/TRI')
                    objeto.append_text('TRIPOLAR')
                if objeto.get_name().__contains__('f9') and tipo_evento == 'LT':
                    objeto.set_property('margin-start', 10)
                if objeto.get_name().__contains__('87l') and tipo_evento == 'LT':
                    objeto.set_property('margin-start', 20)
                    objeto.set_property('margin-end', 30)
            if tipo_evento == 'Trafo':
                if objeto.get_name().__contains__('equip'):
                    objeto.append_text('Banco Monof.')
                    objeto.append_text('Trifásico')
                if objeto.get_name().__contains__('combobox_relacao'):
                    objeto.append_text('500/230')
                    objeto.append_text('500/230/13,8')
                    objeto.append_text('230/138')
                    objeto.append_text('230/138/13,8')
                    objeto.append_text('230/69')
                    objeto.append_text('230/69/13,8')
                    objeto.append_text('138/69')
                    objeto.append_text('138/69/13,8')
                    objeto.append_text('230/6,9')
                    objeto.append_text('69/13,8')
                if objeto.get_name().__contains__('prot'):
                    objeto.append_text('PP/PA')
                    objeto.append_text('PU/PG')
                if objeto.get_name().__contains__('f9'):
                    objeto.set_property('margin-start', 10)
                    objeto.set_property('margin-end', 10)
            if tipo_evento == 'Paisage':
                if objeto.get_name().__contains__('sagebastidor'):
                    objeto.append_text('SAGE')
                    objeto.append_text('BASTIDOR')
                if objeto.get_name().__contains__('nporta_fw'):
                    objeto.set_halign(Gtk.Align.CENTER)
                    objeto.set_valign(Gtk.Align.CENTER)
            if tipo_evento == 'Reator':
                if objeto.get_name().__contains__('equip'):
                    objeto.append_text('Banco Monof.')
                    objeto.append_text('Trifásico')
            if tipo_evento == 'Sistreg':
                if objeto.get_name().__contains__('nome'):
                    objeto.append_text('SAGE')
                    objeto.append_text('UTR-')
                    objeto.append_text('PCPG')
                    objeto.append_text('SART')
                if objeto.get_name().__contains__('tesao_reg'):
                    objeto.append_text('230kV')
                    objeto.append_text('138kV')
                    objeto.append_text('69kV')
                    objeto.append_text('13,8kV')
        return indice_a_adicionar

    # Função principal para realizar a cópia das linhas selecionadas
    def prepara_para_duplicar(self, Linhas_Removidas, Numero_linhas_ativas, tipo_evento, array_nomes_objetos,
                              dicionario_objetos, tabela_evento):
        for linha in Numero_linhas_ativas:  # Varre todas as linhas para achar os checkboxes selecionados
            try:  # Caso para os objetos que foram criados no botão adicionar (dinamicamente)
                check_select: Gtk.CheckButton = dicionario_objetos[
                    array_nomes_objetos[0] + str(linha)]  # Resgatando o objeto checkbutton da linha
                if check_select.get_active():
                    linhas_ocupadas, linha_a_duplicar = self.todas_linhas_preenchidas(linha, Numero_linhas_ativas,
                                                                                      array_nomes_objetos,
                                                                                      dicionario_objetos)
                    if linhas_ocupadas:
                        linha_a_duplicar = self.adicionar_linha(Linhas_Removidas, Numero_linhas_ativas, tipo_evento,
                                                                array_nomes_objetos, dicionario_objetos,
                                                                tabela_evento)
                        self.duplicar_linha(linha, linha_a_duplicar, True, array_nomes_objetos, dicionario_objetos)
                    else:
                        self.duplicar_linha(linha, linha_a_duplicar, True, array_nomes_objetos, dicionario_objetos)
            except:
                check_select: Gtk.CheckButton = builder.get_object(array_nomes_objetos[0] + str(1))
                if check_select.get_active():
                    linhas_ocupadas, linha_a_duplicar = self.todas_linhas_preenchidas(linha, Numero_linhas_ativas,
                                                                                      array_nomes_objetos,
                                                                                      dicionario_objetos)
                    if linhas_ocupadas:
                        linha_a_duplicar = self.adicionar_linha(Linhas_Removidas, Numero_linhas_ativas, tipo_evento,
                                                                array_nomes_objetos, dicionario_objetos,
                                                                tabela_evento)
                        self.duplicar_linha(linha, linha_a_duplicar, False, array_nomes_objetos, dicionario_objetos)
                    else:
                        self.duplicar_linha(linha, linha_a_duplicar, False, array_nomes_objetos, dicionario_objetos)

    # Função que transfere as propriedades da linha selecionada e para a(s) linha(s) inferior(es)
    def duplicar_linha(self, linha, linha_a_duplicar, widget_dinamico, array_nomes_objetos, dicionario_objetos):
        if widget_dinamico:
            for objeto in array_nomes_objetos:
                if objeto != array_nomes_objetos[0]:
                    if 'entry' in objeto:
                        objeto_selecionado: Gtk.Entry = dicionario_objetos[objeto + str(linha)]
                        objeto_duplicado: Gtk.Entry = dicionario_objetos[objeto + str(linha_a_duplicar)]
                        objeto_duplicado.set_text(objeto_selecionado.get_text())

                    elif 'combobox' in objeto:
                        objeto_selecionado: Gtk.ComboBoxText = dicionario_objetos[objeto + str(linha)]
                        objeto_duplicado: Gtk.ComboBoxText = dicionario_objetos[objeto + str(linha_a_duplicar)]
                        objeto_duplicado.set_active(objeto_selecionado.get_active())

                    elif 'checkbtt' in objeto:
                        objeto_selecionado: Gtk.CheckButton = dicionario_objetos[objeto + str(linha)]
                        objeto_duplicado: Gtk.CheckButton = dicionario_objetos[objeto + str(linha_a_duplicar)]
                        objeto_duplicado.set_active(objeto_selecionado.get_active())
        else:
            for objeto in array_nomes_objetos:
                if objeto != array_nomes_objetos[0]:
                    if 'entry' in objeto:
                        objeto_selecionado: Gtk.Entry = builder.get_object(objeto + str(1))
                        objeto_duplicado: Gtk.Entry = dicionario_objetos[objeto + str(linha_a_duplicar)]
                        objeto_duplicado.set_text(objeto_selecionado.get_text())

                    elif 'combobox' in objeto:
                        objeto_selecionado: Gtk.ComboBoxText = builder.get_object(objeto + str(1))
                        objeto_duplicado: Gtk.ComboBoxText = dicionario_objetos[objeto + str(linha_a_duplicar)]
                        objeto_duplicado.set_active(objeto_selecionado.get_active())

                    elif 'checkbtt' in objeto:
                        objeto_selecionado: Gtk.CheckButton = builder.get_object(objeto + str(1))
                        objeto_duplicado: Gtk.CheckButton = dicionario_objetos[objeto + str(linha_a_duplicar)]
                        objeto_duplicado.set_active(objeto_selecionado.get_active())

    # Função que limpa as configurações da(s) linha(s) selecionada(s)
    def limpar_linha(self, Numero_linhas_ativas, array_nomes_objetos, dicionario_objetos):
        for linha in Numero_linhas_ativas:  # Varre todas as linhas para achar os checkboxes selecionados
            try:  # Caso para os objetos que foram criados no botão adicionar (dinamicamente)
                check_select: Gtk.CheckButton = dicionario_objetos[
                    array_nomes_objetos[0] + str(linha)]  # Resgatando o objeto checkbutton da linha
                if check_select.get_active():
                    for objeto in array_nomes_objetos:
                        if objeto != array_nomes_objetos[0]:
                            if 'entry' in objeto:
                                objeto_selecionado: Gtk.Entry = dicionario_objetos[objeto + str(linha)]
                                objeto_selecionado.set_text('')
                            elif 'combobox' in objeto:
                                objeto_selecionado: Gtk.ComboBoxText = dicionario_objetos[objeto + str(linha)]
                                objeto_selecionado.set_active(-1)
                            elif 'checkbtt' in objeto:
                                objeto_selecionado: Gtk.CheckButton = dicionario_objetos[objeto + str(linha)]
                                objeto_selecionado.set_active(False)
            except:
                check_select: Gtk.CheckButton = builder.get_object(array_nomes_objetos[0] + str(linha))
                if check_select.get_active():
                    for objeto in array_nomes_objetos:
                        if objeto != array_nomes_objetos[0]:
                            if 'entry' in objeto:
                                objeto_selecionado: Gtk.Entry = builder.get_object(objeto + str(linha))
                                objeto_selecionado.set_text('')
                            elif 'combobox' in objeto:
                                objeto_selecionado: Gtk.ComboBoxText = builder.get_object(objeto + str(linha))
                                objeto_selecionado.set_active(-1)
                            elif 'checkbtt' in objeto:
                                objeto_selecionado: Gtk.CheckButton = builder.get_object(objeto + str(linha))
                                objeto_selecionado.set_active(False)

    # Função que exclui a(s) linha(s) selecionada(s)
    def exclui_linha(self, Linhas_Removidas, Numero_linhas_ativas, array_nomes_objetos, dicionario_objetos):
        linhas_removidas_agora=[]
        for linha in Numero_linhas_ativas:  # Varre todas as linhas para achar os checkboxes selecionados
            try:  # Caso para os objetos que foram criados no botão adicionar (dinamicamente)
                check_select: Gtk.CheckButton = dicionario_objetos[
                    array_nomes_objetos[0] + str(linha)]  # Resgatando o objeto checkbutton da linha
                if check_select.get_active():
                    Linhas_Removidas.append(linha)
                    linhas_removidas_agora.append(linha)
                    for item in array_nomes_objetos:
                        objeto = dicionario_objetos[
                            item + str(linha)]  # Resgatando os objetos armazenados no dicionário
                        objeto.destroy()  # Comando para deletar os objetos
                        del dicionario_objetos[item + str(linha)]

            except:  # Caso para os objetos que foram criados no glade (primeira Linha)
                check_select: Gtk.CheckButton = builder.get_object(array_nomes_objetos[0] + str(linha))
                if check_select.get_active():
                    Linhas_Removidas.append(linha)
                    for item in array_nomes_objetos:
                        objeto = builder.get_object(item + str(1))
                        objeto.destroy()

        for linha_remov in linhas_removidas_agora:  # Remove as linhas removidas do array de linhas ativas
            del Numero_linhas_ativas[Numero_linhas_ativas.index(linha_remov)]

    # Funções Adicionais Auxiliares

    def preenche_arranjo(self, objeto):

        for arranjo in self.Arranjos:
            objeto.append_text(arranjo)

    # Função que identifica se todas as linhas já estão preenchidas (Usada na função de prapara_para_duplicar)
    def todas_linhas_preenchidas(self, linha, Numero_linhas_ativas, array_nomes_objetos, dicionario_objetos):

        preenchidas = True
        for linha_a_duplicar in Numero_linhas_ativas:
            if linha_a_duplicar != linha:
                if (array_nomes_objetos[0] + str(linha_a_duplicar)) in dicionario_objetos:
                    codigo_evento: Gtk.Entry = dicionario_objetos[array_nomes_objetos[1] + str(linha_a_duplicar)]
                    if codigo_evento.get_name().__contains__('entry'):
                        if codigo_evento.get_text() == '':
                            preenchidas = False
                            break
                        else:
                            preenchidas = True
                    elif codigo_evento.get_name().__contains__('combobox'):
                        if codigo_evento.get_active() == -1:
                            preenchidas = False
                            break
                        else:
                            preenchidas = True
        if preenchidas:
            return [preenchidas, 0]
        else:
            return [preenchidas, linha_a_duplicar]

    def on_menubar_Base_SAGE_para_LP_Excel_activate(self, menubar):
        self.dialogo_diretorio.show()

    def on_diretorio_dialogo_pasta_button_cancelar_clicked(self, button):
        self.dialogo_diretorio.hide()

    def on_diretorio_dialogo_selection_changed(self, selection):
        self.diretorio_dialogo_pasta_entry.set_text(self.dialogo_diretorio.get_filename())

    def on_diretorio_dialogo_pasta_button_selecionar_clicked(self, button):
        try:
            from lp_lib.base2lp import base2lp
        except:
            mensagem_erro('Erro', 'Módulo base2lp não instalado')
            return 0
        diretorio = self.diretorio_dialogo_pasta_entry.get_text()
        self.dialogo_diretorio.hide()
        self.diretorio_dialogo_pasta_entry.set_text("")
        if diretorio:
            try:
                base2lp(diretorio)
            except:
                print_exc(file=stdout)
                mensagem_erro('Erro', 'Erro inesperado ao tentar checar lista de pontos.')

    def on_arqconf_menubar_ferramentas_cepel2excel_activate(self, menubar):
        mensagem_aviso('Aviso', 'você foi avisado')

    def on_arqconf_menubar_arquivo_sair_activate(self, button):
        Gtk.main_quit()


if __name__ == '__main__':
    builder.connect_signals(Manipulador())  # Conecta os sinais da interface com a classe manipuladora "manipulador"
    window = builder.get_object('janela_principal')  # Pega o Objeto da janela princial
    window.show_all()  # Mostra a janela principal
    Gtk.main()
