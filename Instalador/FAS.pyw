�
P��_D  �            '   @   s   d  d l  Z  d  d l Z d  d l m Z d  d l m Z d  d l m Z m Z d  d l m	 Z	 m
 Z
 m Z m Z m Z d  d l m Z d  d l m Z d  d l Z y d  d l m Z Wn e d	 d
 � Yn Xy d  d l m Z Wn e d	 d � Yn Xy d  d l m Z Wn e d	 d � Yn Xy d  d l m Z Wn e d	 d � Yn Xy d  d l m Z Wn e d	 d � Yn XGd d �  d � Z e d k r�e j �  Z e j  d � y e j! d d � Wn Yn Xe j" d  d  � e e � e j# �  n  d S)�    N)�	showerror)�ttk)�askopenfilename�askdirectory)�path�	startfile�listdir�popen�getcwd)�stdout)�	print_exc)�open_workbook�Errou   Módulo xlrd não instalado)�geraru   Módulo Gerar_LP não instalado)�checaru    Módulo Checar_LP não instalado)�
processingu   Módulo func não instalado)�createToolTipu4   Arquivo "func.pyc" deve estar no diretório "lp_lib"c               @   s�   e  Z d  Z d d �  Z d d �  Z d d �  Z d d �  Z d	 d
 �  Z d d �  Z d d �  Z	 d d �  Z
 d d �  Z d d �  Z d d �  Z d d �  Z d d �  Z d d �  Z d d �  Z d S) �Janelac             C   s~	  y$ d d �  t  d � D� d@ |  _ Wn d |  _ Yn Xy$ d d �  t  d � D� dA |  _ Wn d |  _ Yn Xy$ d d �  t  d � D� dB |  _ Wn d |  _ Yn Xd |  _ d |  _ t |  _ d |  _ d	 |  _	 d
 } d } t
 j | � |  _ t
 j |  j d d �|  _ |  j j d d d d d |  j � |  j j d d d d d |  j � |  j j �  |  j j d d d d d t � |  j j d d d d d |  j � t
 j |  j d d �|  _ |  j j d d d d d |  j � |  j j d d d d d |  j � |  j j d d d d d |  j � |  j j d d d d d |  j � |  j j d d d d d |  j � t
 j |  j d d �|  _ |  j j d d d d d |  j � |  j | d <t j �  |  _ |  j j d t
 j � t j �  |  _  |  j  j d t
 j! � t j" |  j d d d t# | � d  t# | � �|  _$ |  j$ j% d! d d" d d# d$ d% d$ � |  j$ j& d � t j' |  j$ d |  j �|  _( |  j( j% d! d d" d � t j) |  j$ d d& d  t# | d' � d |  j* �|  _+ |  j+ j% d! d( d" d d) t
 j, t
 j- t
 j. d% d* d# d+ � t j" |  j d d, d | d  | �|  _/ |  j/ j% d! d( d" d d# d$ d% d$ � |  j/ j& d � t j' |  j/ d |  j �|  _0 |  j0 j% d! d d" d d- d( � |  j0 j& d � t j) |  j/ d d& d  t# | d. � d |  j1 �|  _2 |  j2 j% d! d( d" d d) t
 j3 d% d( d# d/ � |  j2 j& d � t j) |  j/ d d0 d  t# | d. � d |  j4 �|  _5 |  j5 j% d! d( d" d( d% d( d# d/ � |  j5 j& d � t j" |  j d d1 d d2 | d  | �|  _6 |  j6 j% d! d$ d" d d# d$ d% d$ � t j' |  j6 d d3 �|  _7 |  j7 j% d! d d" d � t j) |  j6 d d& d  t# | d' � d |  j8 �|  _9 |  j9 j% d! d( d" d d) t
 j, t
 j3 t
 j- t
 j. d% d* d# d+ � t j: |  j6 � |  _; |  j; j% d! d$ d" d d) t
 j, t
 j3 t
 j- t
 j. d% d+ d# d+ � t j |  j d | d  | �|  _< |  j< j% d! d4 d" d d# d$ d% d$ � t j) |  j< d d5 d  t# | d6 � d |  j= �|  _> |  j> j% d! d d" d d) t
 j. d% d7 d# d* � t j) |  j< d d8 d  t# | d6 � d9 t
 j? d |  j@ �|  _A |  jA j% d! d d" d( d) t
 j3 d% d7 d# d* � t j) |  j< d d: d  t# | d6 � d |  jB �|  _C |  jC j% d! d d" d$ d% d7 d# d* � t j" |  j  d d; d d< | d  | �|  _D |  jD j% d! d d" d d- d( d# d$ d% d+ � |  jD j& d � t
 jE |  jD d  t# d( | d= � d t# d( | d> � �|  _F |  jF j% d! d d" d d) t
 j, t
 j- � |  jF j& d � t
 jG |  jD d? t
 jH d |  jF jI �|  _J |  jJ j% d! d d" d d) t
 j, t
 j- � d  S)CNc             S   s@   g  |  ]6 } | j  d  � d k r | j  d � d k r | � q S)ZPadr�   ZPlanilha�����r   )�find)�.0�arq� r   �FAS.pyw�
<listcomp>.   s   	 z#Janela.__init__.<locals>.<listcomp>�.r   � c             S   s+   g  |  ]! } | j  d  � d k r | � q S)�Configr   r   )r   )r   r   r   r   r   r   2   s   	 c             S   s+   g  |  ]! } | j  d  � d k r | � q S)r   r   r   )r   )r   r   r   r   r   r   6   s   	 z2.0.12z
16/10/2020iJ  �P   �tearoffr   �labelzAbrir pasta do programa�	underline�commandu   Limpar relatórioZSairZArquivo�menuzComparar Listas de Pontos...z Base SAGE para LP Excel...[Beta]zPlanilha Cepel para LP Excel...zGerar Planilha ONS...Z
Ferramenta�Sobre�side�textu   Arquivo LP Padrão�height�width�row�column�padx�   �padyZ
Selecionarg      @�   �sticky�   �
   u   Arquivo de Parametrização�
columnspan�   �   ZEditarzArquivo LP a ser checadog      �?zDefina o arquivo...�   z
Gerar
�   �   z
Checar
�statez
Arquivo Gerado
u    Relatório Geração  g������@�   �   �orientr   r   r   )Kr   �caminhoArquivoLP_PadraoZPlanilhaArquivoLP_Comfig�caminhoArquivoLP_Comfig�PlanilhaArquivoLPEditado�caminhoArquivoLPEditador
   �
pathchecar�versao�data�tkinter�Menu�menubarZ	mnArquivo�add_command�
fcExplorer�fcLimparRelatorio�add_separator�exit�add_cascadeZmnFerramentas�
fcComparar�	fcbase2lp�
fccepel2lp�
fcGerarONSZmnAjuda�sobreClickButtonr   �FrameZfrmE�pack�LEFTZfrmD�RIGHT�
LabelFrame�intZfrm11�grid�grid_propagate�Label�nomeArquivoDOMO�Button�btArqDOMOClickZbotaoEscolheArquivoDOMO�N�S�WZfrm21�nomeArquivoLP_Comfig� botaoEscolheArquivoLPConfigClickZbotaoEscolheLPConfig�E�btEditarArqLPConfigClickZbotaoEditarLPConfigZfrm31�nomeArquivoLPEditado�!botaoEscolheArquivoLPEditadoClickZbotaoEscolheLPEditadoZCombobox�	comboplanZfrm41�gerarClickButtonZ
botaoGerar�DISABLED�checarClickButton�botaoChecar�arquivoClickButtonZbotaoArquivoZfrm12�Listbox�Lb�	Scrollbar�VERTICAL�yviewZscrollY)�self�toplevelZ
frmlarguraZ	frmalturar   r   r   �__init__)   s�    $$$					"""""""""!%<'%.%%<C!%..%+=)*zJanela.__init__c             C   sD   t  d d d d	 g � } | r@ | |  _ t j | � |  j d <n  d  S)
N�	filetypes�Arquivo do Excel�xls�xlsx�xlsmr'   )rv   rw   )rv   rx   )rv   ry   )r   r=   r   �basenamer[   )rr   �tempr   r   r   r]   �   s
    	zJanela.btArqDOMOClickc             C   s   t  |  j � d  S)N)r   r>   )rr   r   r   r   rd   �   s    zJanela.btEditarArqLPConfigClickc             C   s_   t  d d d	 d
 g d |  j � } | r[ t j | � |  _ | |  _ t j | � |  j d <n  d  S)Nru   �Arquivo do Excelrw   rx   ry   �
initialdirr'   )r|   zxls)r|   zxlsx)r|   zxlsm)r   rA   r   �dirnamer>   rz   ra   )rr   r{   r   r   r   rb   �   s    	z'Janela.botaoEscolheArquivoLPConfigClickc          	   C   s�   t  d d d d g d |  j � } | r� | |  _ t j | � |  j d <y t | � } Wn# d | d	 } t d
 | � Yn Xg  } x6 t | j	 � D]% } | j
 | � } | j | j � q� Wt | � |  j d <|  j j d � |  j j d t j � n  d  S)Nru   �Arquivo do Excelrw   rx   ry   r}   r'   z	Arquivo "u   " não encontrador   �valuesr   r9   )r   zxls)r   zxlsx)r   zxlsm)r   rA   r@   r   rz   re   r   r   �rangeZnsheets�sheet_by_index�append�name�tuplerg   �currentrk   �configrD   �NORMAL)rr   r{   Zbook�avisoZarray_comboZ
plan_index�sheetr   r   r   rf   �   s$    	z(Janela.botaoEscolheArquivoLPEditadoClickc             C   s@  |  j  j d t j � y t |  j � } Wn& d |  j d } t d | � Yn Xy� | j d � } t j	 d t
 | j d d � � � d j d � } t t t | � � } | d d d	 g k  r� t d d
 � nT y/ t t i |  j d 6|  j  d 6|  j d 6� Wn" t d t � t d d � Yn XWn t d d � Yn Xd  S)Nr   z	Arquivo "u   " não encontrador   z\d+\.\d+\.\d+�n   r   r/   �   uH   Deve ser usado arquivo LP_Config.xls com versão igual ou maior a 2.0.12�	LP_Padrao�	relatorio�	LP_Config�filez0Erro inesperado ao tentar gerar lista de pontos.uG   Arquivo indicado não corresponde a arquivo de parametrização válido)rn   �deleterD   �ENDr   r>   r   r�   �re�findall�str�cell�split�list�maprW   r   r   r=   r   r   )rr   �arq_confr�   r�   �versr   r   r   rh   �   s(    1zJanela.gerarClickButtonc             C   sf  |  j  j �  |  _ |  j j d t j � y t |  j � } Wn& d |  j d } t	 d | � Yn Xy� | j
 d � } t j d t | j d d � � � d j d � } t t t | � � } | d d d	 g k  r� t	 d d
 � nh yC t t i |  j d 6|  j d 6|  j d 6|  j d 6|  j d 6� Wn" t d t � t	 d d � Yn XWn t	 d d � Yn Xd  S)Nr   z	Arquivo "u   " não encontrador   z\d+\.\d+\.\d+r�   r   r/   �   uH   Deve ser usado arquivo LP_Config.xls com versão igual ou maior a 2.0.11r�   Z
LP_EditadoZplanilhar�   r�   r�   z1Erro inesperado ao tentar checar lista de pontos.uG   Arquivo indicado não corresponde a arquivo de parametrização válido)rg   �getr?   rn   r�   rD   r�   r   r>   r   r�   r�   r�   r�   r�   r�   r�   r�   rW   r   r   r=   r@   r   r   )rr   r�   r�   r�   r�   r   r   r   rj   �   s.    1zJanela.checarClickButtonc          
   C   sF   y* t  j t d d � � } t | d � Wn t d d � Yn Xd  S)Nzfas.p�r�arquivor   u   Não existe arquivo definido)�pickle�load�openr   r   )rr   Zconfr   r   r   rl     s
    zJanela.arquivoClickButtonc             C   s   t  d � d  S)Nz
explorer .)r	   )rr   r   r   r   rH     s    zJanela.fcExplorerc             C   s   |  j  j d t j � d  S)Nr   )rn   r�   rD   r�   )rr   r   r   r   rI     s    zJanela.fcLimparRelatorioc             C   s�   y d d l  m } Wn t d d � d SYn Xt j �  } | j d � y | j d d � Wn Yn X| j d d � | | |  j � | j	 �  d  S)Nr   )�
JanelaCompr   u"   Módulo LP_Comparar não instaladozComparar Arquivos�defaultzlp_lib/chesf.ico)
Zlp_lib.LP_Compararr�   r   rD   �Toplevel�title�
iconbitmap�	resizablern   �mainloop)rr   r�   Zjncompr   r   r   rM   !  s    	zJanela.fcCompararc             C   s   y d d l  m } Wn t d d � d SYn Xt d d � } | r{ y | | � Wq{ t d t � t d d � Yq{ Xn  d  S)	Nr   )�base2lpr   u   Módulo base2lp não instalador�   u2   Selecione o diretório que estão os arquivos .datr�   z1Erro inesperado ao tentar checar lista de pontos.)Zlp_lib.base2lpr�   r   r   r   r   )rr   r�   Z	diretorior   r   r   rN   2  s    	zJanela.fcbase2lpc             C   s�   y d d l  m } Wn t d d � d SYn Xt d d d d g � } | r� y | | � Wq� t d
 t � t d d � Yq� Xn  d  S)Nr   )�cepel2lpr   u   Módulo cepel2lp não instaladoru   �Arquivo do Excelrw   rx   ry   r�   z1Erro inesperado ao tentar checar lista de pontos.)r�   zxls)r�   zxlsx)r�   zxlsm)Zlp_lib.cepel2lpr�   r   r   r   r   )rr   r�   Zarqcepelr   r   r   rO   @  s    	zJanela.fccepel2lpc             C   s�   y d d l  m } Wn t d d � d SYn Xt j �  } | j d � y | j d d � Wn Yn X| j d d � | | |  j � | j	 �  d  S)Nr   )�JanelaGerarONSr   u    Módulo Gerar_ONS não instaladozGerar Planilha ONSr�   zlp_lib/chesf.ico)
Zlp_lib.Gerar_ONSr�   r   rD   r�   r�   r�   r�   rn   r�   )rr   r�   Z
jngeraronsr   r   r   rP   O  s    	zJanela.fcGerarONSc             C   s�   t  j �  } | j d � d |  j d |  j } t  j | d d d d d t  j d	 d �j d d d d d t  j t  j	 t  j
 t  j � t  j | d | d t  j d	 d �j d d d d d t  j t  j	 t  j
 t  j d d d d � | j d � d  S)Nr%   u   
        Versão u�   
        Ferramenta de Automatização para Projetos de Sistemas Supervisórios
        Produzido e mantido pelo DETA
        Atualização do programa: r'   z
F A S�fg�blue�anchor�font�Verdana�14�bold italicr*   r   r+   r0   �8r/   r.   r2   r,   �   r   )r�   z14r�   )r�   r�   )rD   �Tkr�   rB   rC   rZ   �CENTERrX   r^   rc   r_   r`   r�   )rr   ZsobreZtextor   r   r   rQ   `  s    '!zJanela.sobreClickButtonN)�__name__�
__module__�__qualname__rt   r]   rd   rb   rf   rh   rj   rl   rH   rI   rM   rN   rO   rP   rQ   r   r   r   r   r   (   s   �	r   �__main__uL   FAS - Ferramenta de Automatização para Projetos de Sistemas Supervisóriosr�   zlp_lib/chesf.ico)$r�   rD   �tkinter.messageboxr   r   �tkinter.filedialogr   r   �osr   r   r   r	   r
   �sysr   �	tracebackr   r�   Zxlrdr   Zlp_lib.Gerar_LPr   Zlp_lib.Checar_LPr   Zlp_lib.funcr   r   r   r�   r�   Zappr�   r�   r�   r�   r   r   r   r   �<module>   sR   (� O
