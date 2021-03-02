# -*- coding: cp860 -*-

import os

import re

import datetime

import inspect

import copy

import gi
gi.require_version("Gtk", "3.0")
from gi.repository import Gtk,GObject
from FASgtkui import mensagem_aviso, message_aviso_dialog, mensagem_erro, mensagem_erro_dialog

import time
DAT_KEYS = {}
DAT_KEYS["pds"]=DAT_KEYS["pdf"]=DAT_KEYS["pdd"]=DAT_KEYS["pas"]=DAT_KEYS["noh"]=DAT_KEYS["cnf"]=DAT_KEYS["nv1"]= \
    DAT_KEYS["tac"]=DAT_KEYS["paf"]=DAT_KEYS["cgs"]=DAT_KEYS["cgf"]=DAT_KEYS["pts"]=DAT_KEYS["ptf"]=\
    DAT_KEYS["pad"]=DAT_KEYS["lsc"]=DAT_KEYS["nv1"]=DAT_KEYS["nv2"]=DAT_KEYS["tdd"]=\
    DAT_KEYS["ctx"]=DAT_KEYS["enm"]=DAT_KEYS["grupo"]=DAT_KEYS["gsd"]=DAT_KEYS["ptd"]=DAT_KEYS["map"]="ID"
DAT_KEYS["rca"]=DAT_KEYS["rfc"]=DAT_KEYS["rfi"]=DAT_KEYS["e2m"]=DAT_KEYS["grcmp"]=DAT_KEYS['ocr']=DAT_KEYS["cxp"]= \
    DAT_KEYS["psv"]=""

DAT_FIELDS = {
    "pas" : ["ID","NOME","TAC","LIA","LIE","LIU","LSA","LSE","LSU","ALINT","ALRIN","TCL","OCR","TIPO","BNDMO", \
             "CDINIC","VLINIC","CMT"], \
    "paf" : ["ID","DESC1","NV2","PNT","TPPNT","KCONV1","KCONV2","KCONV3","CMT"], \
    "pds" : ["ID","NOME","TIPO","TAC","TCL","TPFIL","OCR","ALINT","ALRIN","SOEIN","CDINIC","STINI","STNOR","CMT"], \
    "pdf" : ["ID","DESC1","NV2","KCONV","TPPNT","PNT","ORDEM","CMT"], \
    "cgf" : ["ID","DESC1","NV2","KCONV","CMT"], \
    "cgs" : ["ID","NOME","TAC","LMI1C","LMI2C","LMS1C","LMS2C","PAC","PINT","TIPO","TIPOE","TPCTL","TRRAC","CMT"], \
    "rca" : ["ORDEM","PARC","PNT","TIPOP","TPPARC","TPPNT"], \
    "noh" : ["ID","NOME","ENDIP","TPNOH","NTATV"], \
    "cnf" : ["ID", "CONFIG", "LSC"],
    "tac" : ["ID", "TPAQS", "NOME", "LSC"],
    "lsc" : ["ID", "TIPO", "NOME", "TCV", "TTP", "GSD", "NSRV1", "NSRV2", "MAP", "VERBD"],
    "rfc" : ["PNT"],
    "rfi" : ["ORDEM"],
    "nv1" : ["ID"],
    "nv2" : ["ID"],
    "pts" : ["ID"],
    "ptf" : ["ID"],
    "tdd" : ["ID"],
    'pdd' : ['ID'],
    'gsd' : ['ID'],
    'ocr' : ['ID'],
    'pad' : ['ID'],
    'inp': ['NOH','ORDEM','PRO'],
    'cxu': ['ID'],
    'utr': ['ID'],
    'enm': ['ID','ORDEM'],
    'pro': ['ID'],
    'enu': ['ID'],
    'ins': ['ID'],
    'mul': ['ID'],
    'tn1': ['ID'],
    'tn2': ['ID'],
    'sxp': ['PRO'],
    'tctl': ['ID'],
    'sev': ['ID'],
    'ttp': ['ID'],
    'tcv': ['ID'],
    'inm': ['CLASSE'],
    'psv': ['GRUPO'],
    'ctx': ['ID'],
    'cxp': ['CLASSE'],
    'e2m': ['IDPTO','MAP','TIPO'],
    'tcl': ['ID'],
    'map': ['ID'],
    'grupo': ['ID'],
    'grcmp': ['GRUPO','ORDEM1','ORDEM2']
    }

MSG_ERR = "erro"
MSG_WARN = "aviso"
MSG_INFO = "info"


def _make_include_str(dat_type, **kwargs):
    base = kwargs.get("base","")
    source_path = kwargs.get("source_path","")
    source_path = fix_path(source_path)
    add_to = kwargs.get("add_to","")
    #print(add_to)
    add_to = fix_path(add_to)
    #print(add_to)
    #print(source_path)
    dat_path=""
    # Define o caminho do dat em uma das opáîes na seguinte ordem de prioridade: base, caminho,
    #  ou diret¢rio corrente
    if base:
        base_path = "/export/sage/sage/config/"+base+"/bd/dados/"
    elif source_path:
        base_path = source_path
    else:
        base_path = os.path.curdir

    if add_to:
        base_path = os.path.join(base_path, add_to)
        if not base_path[-4:] == ".dat":
            dat_path = os.path.join(base_path, dat_type + ".dat")
        else:
            dat_path = base_path
    return dat_path

def make_include_str(dat_type, **kwargs):
    add_to = kwargs.get("add_to","")
    #print(add_to)
    add_to = add_to.replace("\\","/")
    #print(add_to)
    #print(source_path)
    # Define o caminho do dat em uma das opáîes na seguinte ordem de prioridade: base, caminho,
    #  ou diret¢rio corrente
    if add_to:
        if not add_to[-4:] == ".dat":
            add_to = os.path.join(add_to, dat_type + ".dat")
        if not add_to.startswith("/"):
            add_to = "/"+add_to
        add_to = "#"+add_to
        include_str = add_to
    else:
        include_str = dat_type
    return include_str

def get_dat_pattern(dat_type):

    '''get_dat_pattern(dat_type) retorna uma raw string preparada para regex, que encontra o tipo de registro passado em dat_type.
    Ex.: get_dat_pattern("pds") retorna um padrÑo de expressÑo regular que encontra todas as entradas PDS de linha £nica
    :param dat_type
    :rtype: str
    '''
    dat_pattern = r"^;*[\s]*"
    for c in dat_type:
        if c.isalpha():
            dat_pattern+=r"["+c.swapcase()+c+"]"
        else:
            dat_pattern+=r"["+c+"]"
    dat_pattern+=r"[\s]*\n"
    return dat_pattern


def is_dat_type(dat_type, dat_str):
    '''
    Retorna True caso a string dat_str tenha o formato de dados de dat_type
    :param dat_type
    :param dat_str
    :rtype: str
    '''
    if dat_type and dat_str:
        return re.search(pattern=get_dat_pattern(dat_type),string=dat_str,flags=re.MULTILINE)


def get_field_pattern():
    '''
    Retorna uma string para ser usada em regex a fim de determinar campos de entradas dat
    :return:
    '''
    # original-> r"^([;]+)?[\s]*[a-zA-Z\d]+([\s]+)?=([\s]+)?([-:=_\$\.a-zA-Z0-9\s]+)?(\n)?"
    return r"^([;]+)?[\s]*[_a-zA-Z\d]+([\s]+)?=([\s]+)?([-:=_\$\.a-zA-Z0-9\s]+)?(\n)?"


def is_dat_field(field_str):
    '''
    Retorna True caso field_str tenha o formato de um campo de um registro dat. ATENÄéO: Campos comentados
    tambÇm sÑo identificados como v†lidos e retornam True.
    :param field_str:
    :return:
    '''
    z=None
    if field_str:
        field_str=field_str.strip()
        z= re.search(pattern=get_field_pattern(),string=field_str)
    if z is None:
        return False
    else:
        return True


def get_rec_list(dat_str,dat_type):
    '''
    Retorna uma lista de entradas do tipo dat_type de dat_str
    :param dat_str:
    :param dat_type:
    :return:
    '''
    return re.split(pattern=get_dat_pattern(dat_type), string=dat_str,flags=re.MULTILINE)


def _is_comment(dat_line):
    return str(dat_line).strip().startswith(";")


def _is_comment_field(dat_line):
    return is_dat_field(dat_line) and _is_comment(dat_line)


def _is_include(dat_line):
    return str(dat_line).startswith("#include ")


def get_file(dat_path, **kwargs):
    s=""
    if not os.path.exists(dat_path):
        dat_path = os.path.join(os.curdir,dat_path)
    if not os.path.exists(dat_path):
        GObject.idle_add(mensagem_aviso,'Aviso','Caminho para o arquivo nÑo exite: {}'.format(dat_path))
        time.sleep(1)
        while message_aviso_dialog.get_visible() == True:
            time.sleep(1)
        #print_msg(__name__, "caminho para o arquivo nÑo existe: (0)".format(dat_path), **kwargs)
    if os.path.isfile(dat_path):
        # leitura do arquivo e criaáÑo da string
        try:
            dat_file = open(dat_path, "r", encoding="iso-8859-1")
            for line in dat_file.readlines():
                s+=line
        except:
            raise
        finally:
            dat_file.close()
    return s


def _load_datx(dat_type, **kwargs):

    dat_content=[]     # vari†vel de retorno com a lista de pontos/dicion†rios
    dat=""      # aramazena o texto puro do .dat
    is_nested = False
    base_path = ""  # armazena o diret¢rio de trabalho (base ou diret¢rio passado como parÉmetro)
    no_comments=kwargs.get("no_comments",False)
    no_field_comment = kwargs.get("no_field_comment",False)
    ignore_includes = kwargs.get("ignore_includes",False)
    base = kwargs.get("base","")
    source_path = kwargs.get("source_path","")
    source_str = kwargs.get("source_str","")
    concatenate = kwargs.get("concatenate",False)

    # Carrega o dat em uma das opáîes na seguinte ordem de prioridade: base, caminho, string ou diret¢rio corrente
    if base != "":
        dat_path = "/export/sage/sage/config/"+base+"/bd/dados/"+dat_type+".dat"
        try:
            dat = get_file(dat_path, **kwargs)
            base_path = os.path.dirname(dat_path)
        except:
            print("Erro: base {0} nÑo encontrada".format(base))

    elif source_path != "":
        try:

            dat = get_file(source_path, **kwargs)

            base_path = os.path.dirname(source_path)
        except:
            print("Erro na leitura de {0}".format(source_path))

    elif source_str != "":
        dat = source_str
        base_path="."
    else:
        dat_path = os.path.join(os.path.curdir, dat_type + ".dat")
        try:
            dat = get_file(dat_path, **kwargs)
            base_path = os.path.dirname(dat_path)
        except:
            print("Erro: nenhuma entrada especificada")

    # checa se o dat Ç do tipo informado
    if not is_dat_type(dat_type,dat):
        print("Erro ao carregar {0}: {1}".format(dat_type, source_path))
        return dat_content

    # constr¢i a lista de entradas para iterar
    reclist = get_rec_list(dat_str=dat,dat_type=dat_type)

    # para cada entrada/ponto
    for rec in reclist:
            # dic ser† preenchido com os campos do ponto
            dic={}
            # caso a entrada possua coment†rios, serÑo guardados em comment_rec
            comment_rec=""
            # analisa cada linha da entrada Ö procura de campos v†lidos, inclusive comentados
            for field in rec.split("\n"):
                # caso nÑo seja permitido coment†rios ou campo comentado, pule para o pr¢ximo
                if (no_comments and _is_comment(field) and not _is_comment_field(field)) \
                        or (no_field_comment and _is_comment_field(field)):
                    continue
                if is_dat_field(field):
                    # caso campo v†lido, acrescente a dic
                    fieldcontents = field.strip().split("=",1)
                    fieldname = fieldcontents[0].upper().strip().replace(" ","")
                    fieldvalue = fieldcontents[1].strip()
                    dic[fieldname]=fieldvalue
                elif _is_comment(field):
                    # caso seja um coment†rio, acrescente a comment_rec
                    comment_rec+=field.strip()+"\n"
                elif _is_include(field) and not ignore_includes:
                    '''
                    Caso seja um include e deva ser processado, este objeto dicion†rio da lista dat_content ter† como £nica
                    chave uma string do tipo '#include {caminho do include}', e como valor uma lista contendo todos os
                    pontos (como dicion†rios) do dat
                    referenciado no include. Para receber uma lista com includes concatenados, deve-se passar o
                    argumento concatenate = False
                    '''
                    #include_path = os.path.join(base_path, field.lstrip("#include").strip().lstrip("//"))
                    include_path = field.lstrip("#include").strip()
                    include_path = fix_path(include_path)
                    base_path = fix_path(base_path)
                    include_path = os.path.join(base_path, include_path )

                    inc_recs = _load_datx(dat_type=dat_type, source_path=include_path, no_comments=no_comments, \
                                          no_field_comment=no_field_comment, concatenate=concatenate, \
                                          ignore_includes=ignore_includes)
                    fieldname = "#"+include_path
                    if len(inc_recs)>0:
                        dic[fieldname]=inc_recs
                        is_nested=True
                        dat_content.append(dic)
                        dic={}


            # caso tenha lido ponto (dic existe), acrescentar Ö lista dat_content
            if len(dic.items()) !=0:
                dat_content.append(dic)
            # caso haja coment†rios, acrescente Ö lista dat_content com chave 'comment'
            if comment_rec:
                dat_content.append({"comment":comment_rec})

    # opáÑo para retornar lista concatenada, sem os includes aninhados
    if concatenate and is_nested:
        for r in dat_content:
            #assert isinstance(r, dict)
            k=list(r.keys())
            #assert isinstance(k, str)
            if k[0].startswith("#"):
                # pega a lista nova
                new_l = r[k[0]]
                # remove item da lista
                r.popitem()
                # insere coment†rio de in°cio da concatenaáÑo
                # dat_content.append({"include":k[0].strip('#include ').lstrip(base_path)})
                for n in new_l:
                    # adiciona novos itens de dicion†rio
                    dat_content.append(n)
                # limpa a lista removendo dicion†rios vazios
                new_set = list((v) for v in dat_content if v)
                # substitui a lista dat_content com a nova lista limpa
                dat_content = new_set
    return dat_content


def fix_path(path_string):
    if os.name == "posix":
        path_string = path_string.replace("\\","/")
        #path_string = path_string.lstrip("/")
    elif os.name == "nt":
        path_string = path_string.replace("/", "\\")
        #path_string = path_string.lstrip("\\")
    return path_string




def load_dat(dat_type, **kwargs):

    dat_type = dat_type.lower()
    dat_content={}     # vari†vel de retorno com a lista de pontos/dicion†rios
    dat=""      # aramazena o texto puro do .dat
    is_nested = False
    base_path = ""  # armazena o diret¢rio de trabalho (base ou diret¢rio passado como parÉmetro)
    no_comments=kwargs.get("no_comments",False)
    no_field_comment = kwargs.get("no_field_comment",False)
    ignore_includes = kwargs.get("ignore_includes",False)
    base = kwargs.get("base","")
    source_path = kwargs.get("source_path","")
    source_str = kwargs.get("source_str","")
    concatenate = kwargs.get("concatenate",False)
    label_progressbar: Gtk.Label = kwargs.get('label_progressbar',None)


    # Carrega o dat em uma das opáîes na seguinte ordem de prioridade: base, caminho, string ou diret¢rio corrente
    if base != "":
        dat_path = "/export/home/sage/sage/config/"+base+"/bd/dados/"+dat_type+".dat"
        try:
            dat = get_file(dat_path, **kwargs)
            base_path = os.path.dirname(dat_path)
        except:
            GObject.idle_add(mensagem_erro,'Erro',"base {0} nÑo encontrada".format(base))
            time.sleep(1)
            while mensagem_erro_dialog.get_visible() == True:
                time.sleep(1)
            #print("(datapi.load_dat) erro: base {0} nÑo encontrada".format(base))

    elif source_path != "":
        source_path = fix_path(source_path)
        if not ".dat" in source_path:
            dat_path = os.path.join(source_path, dat_type+".dat")
        else:
            dat_path = source_path
        #dat_path = os.path.join(os.path.curdir,source_path)
        try:
            dat = get_file(dat_path, **kwargs)
            base_path = os.path.dirname(dat_path)
        except:
            GObject.idle_add(mensagem_erro, 'Erro', "erro na leitura: {0}".format(source_path))
            time.sleep(1)
            while mensagem_erro_dialog.get_visible() == True:
                time.sleep(1)
            #print("(datapi.load_dat) erro na leitura: {0}".format(source_path))

    elif source_str != "":
        dat = source_str
        base_path="."
    else:
        dat_path = os.path.join(os.path.curdir, dat_type + ".dat")
        try:
            dat = get_file(dat_path, **kwargs)
            base_path = os.path.dirname(dat_path)
        except:
            GObject.idle_add(mensagem_erro, 'Erro', "erro: nenhuma entrada especificada")
            time.sleep(1)
            while mensagem_erro_dialog.get_visible() == True:
                time.sleep(1)
            #print("(datapi.load_dat) erro: nenhuma entrada especificada")

    # checa se o dat Ç do tipo informado
    #if not is_dat_type(dat_type,dat):
     #   print("Erro ao carregar {0}: {1}".format(dat_type, source_path))
      #  return dat_content

    # constr¢i a lista de entradas para iterar
    reclist = get_rec_list(dat_str=dat,dat_type=dat_type)

    dat_list = []
    # para cada entrada/ponto
    for rec in reclist:
            # dic ser† preenchido com os campos do ponto
            dic={}
            # caso a entrada possua coment†rios, serÑo guardados em comment_rec
            comment_rec=""
            # analisa cada linha da entrada Ö procura de campos v†lidos, inclusive comentados
            for field in rec.split("\n"):
                # caso nÑo seja permitido coment†rios ou campo comentado, pule para o pr¢ximo
                if (no_comments and _is_comment(field) and not _is_comment_field(field)) \
                        or (no_field_comment and _is_comment_field(field)):
                    continue
                if is_dat_field(field):
                    # caso campo v†lido, acrescente a dic
                    fieldcontents = field.strip().split("=", 1)
                    fieldname = fieldcontents[0].upper().strip().replace(" ","")
                    fieldvalue = fieldcontents[1].strip()
                    #print("fieldname= "+fieldname+"\n"+"fieldvalue= "+fieldvalue)
                    dic[fieldname]=fieldvalue
                elif _is_comment(field):
                    # caso seja um coment†rio, acrescente a comment_rec
                    comment_rec+=field.strip()+"\n"
                elif _is_include(field) and not ignore_includes:
                    '''
                    Caso seja um include e deva ser processado, este objeto dicion†rio da lista dat_content ter† como £nica
                    chave uma string do tipo '#include {caminho do include}', e como valor uma lista contendo todos os
                    pontos (como dicion†rios) do dat
                    referenciado no include. Para receber uma lista com includes concatenados, deve-se passar o
                    argumento concatenate = False
                    '''
                    include_str = field.lstrip("#include").strip() # ser† usado como chave do dicion†rio
                    include_path = include_str.lstrip("/")
                    include_path = fix_path(include_path)
                    #TESTE include_path = os.path.join(base_path, include_path )
                    fname = inspect.currentframe().f_back.f_code.co_name
                    label_progressbar.set_text("({0}.{1}) {2}: {3}".format(__name__, fname, MSG_INFO, include_path))

                    tmp_path = os.path.join(base_path, include_path)
                    inc_recs = load_dat(dat_type=dat_type, source_path=tmp_path, no_comments=no_comments, \
                                          no_field_comment=no_field_comment, concatenate=concatenate, \
                                          ignore_includes=ignore_includes)
                    fieldname = "#"+include_str
                    if len(inc_recs)>0:
                            dat_content[fieldname]=list(inc_recs[dat_type])
                            is_nested=True

            # caso tenha lido ponto (dic existe), acrescentar Ö lista dat_list
            if len(list(dic.items())) !=0:
                dat_list.append(dic)
            # caso haja coment†rios, acrescente Ö lista dat_content com chave 'comment'
            if comment_rec:
                dat_list.append({"comment":comment_rec})

    dat_content[dat_type] = dat_list

    if concatenate and is_nested:
        flat_content = {}
        dat_content[dat_type].insert(0,{"comment":";==== in°cio do arquivo {0} ====".format(dat_path)+"\n\n"})
        flat_content[dat_type]=dat_content[dat_type][:] # faz shalllow copy
        for k in list(dat_content.keys()):
            if k != dat_type:
                flat_content[dat_type].append({"comment":"\n\n;==== in°cio do include {0} ====".format(k)})
                flat_content[dat_type].extend(list(dat_content[k]))
        dat_content = flat_content

    return dat_content


def _ds_clear_comments(dataset):
    if len(dataset) > 0:
        dataset = list((v) for v in dataset if not "comment" in v.keys())


def _dc_clear_comments(dat_content):
    for k in list(dat_content.keys()):
        _ds_clear_comments(dat_content[k])


def clear_comments(generic_set):
    if type(generic_set) == dict:
        _dc_clear_comments(generic_set)
    elif type(generic_set) == list:
        _ds_clear_comments(generic_set)


def _ds_clear_includes(dataset):
    # remove include aninhado
    output = []
    for d in dataset:
        k = list(d.keys())
        if not k[0].startswith("#"):
            output.append(d)
        else:
            continue
    dataset = output


def _dc_clear_includes(dat_content):
    dat_content = dict({k:dat_content[k] for k in list(dat_content.keys()) if not k.startswith("#")})


def clear_includes(generic_set):
    if type(generic_set) == dict:
        _dc_clear_includes(generic_set)
    elif type(generic_set) == list:
        _ds_clear_includes(generic_set)


def make_dat_str(dat_type, dataset, **kwargs):
    '''
    Retorna uma tupla output, c onde output Ç uma string para salvar o .dat e c Ç o total de pontos criados.
    Caso seja passada uma lista de campos como field_order, a escrita dos campos no dat obedece a esta sequància,
    caso contr†rio eles sÑo escritos em ordem alfabÇtica
    :param dat_type:
    :param dataset:
    :return:
    '''
    field_order = list(kwargs.get("field_order",[]))
    output = ""
    c = 0
    for d in dataset:

        keys = list(d.keys())
        keys.sort()
        if "comment" in keys:
            output += "\n" + d["comment"] + "\n"
        else:
            c += 1
            p =''
            if is_commented_point(d):
                p=';'
            output += "\n" + p + dat_type.upper() + "\n"
            if field_order != []:
                # remove de field_order os elementos q n estÑo em keys
                field_order = list((f) for f in field_order if f in keys)
                # remove de keys os elementos que j† estÑo em field_order
                keys = list((k) for k in keys if not k in field_order)
                # cria lista completa organizada e atribui a keys
                field_order.extend(keys)
                keys = field_order[:]
            for k in keys:
                output += "\t"+str(k).upper()+" = "+d[k]+"\n"
    now = datetime.datetime.now()
    data = "{0}-{1}-{2} Ös {3}:{4}:{5}".format(now.year, now.month, now.day, now.hour, now.minute, now.second)
    total_lines = len(output.split("\n"))
    footer = \
'''
;=================================================================================
;Arquivo criado pelo datapi.py beta 0.1 em {0}
;Total de linhas = {1}
;Total de pontos = {2}
;=================================================================================
\n'''.format(data, total_lines, c)
    #output += footer
    return output, c


def save_to_disk(dat_path, text, **kwargs):
    if os.path.isfile(dat_path):
        if kwargs.get("do_backup",True):
            c = 1
            while 1:
                filename = os.path.join(os.path.splitext(dat_path)[0] + ".old_"+str(c))
                if os.path.isfile(filename):
                   c +=1
                else:
                    break
            #print_msg(__name__, "dat existente: "+dat_path, MSG_WARN, **kwargs)
            print_msg(__name__, "renomeando dat existente para: "+filename, MSG_INFO, **kwargs)
            os.rename(dat_path,filename)
        else:
            #print_msg(__name__, "removendo arquivo existente: "+dat_path, MSG_INFO, **kwargs)
            os.remove(dat_path)
    #print("(datapi.save_to_disk) tentando escrever em : "+dat_path)
    dir = os.path.dirname(dat_path)
    os.makedirs(dir, exist_ok=True)
    with open(dat_path,"w",encoding="iso-8859-1") as f:
    #with open(dat_path,"w") as f:
        f.write(text)

    print_msg(__name__, "sucesso ao escrever : "+dat_path, MSG_INFO, **kwargs)

def write_dat(dat_type, dat_content, **kwargs):
    '''
    Salva em disco o conte£do de dat_content. Caso dat_content tenha novos includes, os diret¢rios serÑo criados
    de acordo.
    :param dat_type: (str) tipo do dat que ser† escrito
    :param dat_content: (dict) dicion†rio com o conte£do do dat e includes
    :param kwargs:
        source_path = (str) caminho da base onde ser† escrito
        base = (str) nome da base onde deve ser escrito
        dests = (list) caso queira escrever apenas
    o dat principal ou alguns includes, deve ser passado na lista dests as chaves de dat_content que devem ser
    salvas em disco. Ex: dests = ["lsc", "#/siemens/lsc.dat"] escreve apenas o conte£do principal e o include
    siemens, ignorando os demais includes, se houver.
        do_backup = (bool=True) faz com que seja salvo um backup do arquivo atual caso exista, no formato "dat.old_x"
    onde x Ç um n£mero sequencial.
    :return:
    '''
    if (dat_content == []) or (dat_type == ""):
        return 1
    base_path = ""  # armazena o diret¢rio de trabalho (base ou diret¢rio passado como parÉmetro)
    no_comments=kwargs.get("no_comments",False)
    no_field_comment = kwargs.get("no_field_comment",False)
    ignore_includes = kwargs.get("ignore_includes",False)
    base = kwargs.get("base","")
    source_path = kwargs.get("source_path","")
    do_backup = kwargs.get("do_backup",True)
    verbose = kwargs.get('verbose', True)
    dests = kwargs.get("dests",[])
    dat_path=""

    # Define o caminho do dat em uma das opáîes na seguinte ordem de prioridade: base, caminho,
    # string ou diret¢rio corrente
    if base != "":
        dat_path = "/export/sage/sage/config/"+base+"/bd/dados/"+dat_type+".dat"
        base_path = os.path.dirname(dat_path)
    elif source_path != "":
        source_path = fix_path(source_path)
        if not ".dat" in source_path:
            dat_path = os.path.join(source_path,dat_type+".dat")
        else:
            dat_path = source_path
        #dat_path = os.path.join(source_path,dat_type+".dat")
        #dat_path = fix_path(dat_path)
        base_path = os.path.dirname(dat_path)
    else:
        dat_path = os.path.join(os.path.curdir, dat_type + ".dat")
        dat_path = fix_path(dat_path)
        base_path = os.path.dirname(dat_path)

    if not os.path.exists(base_path):
        print_msg(__name__,"diret¢rio nÑo existe. Criando: {0} ".format(base_path),MSG_INFO, **kwargs)
        try:
            os.makedirs(base_path)
            print_msg(__name__,'diret¢rio criado com sucesso', MSG_INFO, **kwargs)
        except:
            return False


    if no_comments:
        clear_comments(dat_content)
    if ignore_includes:
        clear_includes(dat_content)

    output_list = []
    include_str = ""

    # se h† algum include novo em dat_content, o dat principal deve ser atualizado
    update_includes = False
    current_includes = get_curr_includes(dat_path, **kwargs)
    #print("currrent includes:"+str(current_includes))
    for dest in dests:
        if not dest in current_includes:
            print_msg(__name__," dat principal ser† atualizado", MSG_INFO, **kwargs)
            update_includes = True
            break

    for k in list(dat_content.keys()):
        #print_msg(__name__, "criando dat para {0}".format(k), MSG_INFO, **kwargs)
        output, c = make_dat_str(dat_type, dat_content[k], field_order=DAT_FIELDS.get(dat_type,[]))
        if k.startswith("#"):
            include_path = str(k)
            # print("write_dat: include_path = "+include_path)
            include_path = include_path.lstrip("#").strip()
            # print("write_dat: include_path = "+include_path)
            # print("write_dat: base_path ="+base_path)
            include_str = include_str + "#include "+include_path+"\n"
            include_path = include_path.lstrip("/")
            include_path = fix_path(include_path)
            include_path = os.path.join(base_path, include_path)

            # print("write_dat: include_str = "+include_str)
            if (dests == []) or (k in dests):
                # colocar na lista output se for marcado para ser escrito
                output_list.append([include_path,output])
        else: # caso seja o dat principal, inserir no in°cio da lista de sa°da
            if (dests == []) or (k in dests) or (update_includes):
                # colocar na lista output de escrita se deve ser atualizado
                output_list.insert(0,[dat_path, output])

    # insere na string do dat principal as linhas de include, caso o dat principal esteja sendo editado
    if (dests == []) or (dat_type in dests) or (update_includes):
        output_list[0][1] = include_str + output_list[0][1]

    # cria os arquivos
    for o in output_list:
        # print("dat_path ="+o[0])
        print_msg(__name__,"escrevendo " + o[0], MSG_INFO, **kwargs)
        save_to_disk(dat_path=o[0],text=o[1], do_backup=do_backup, verbose=verbose)
    return 0


def bulk_write_dat(write_list, **kwargs):
    assert isinstance(write_list, list)
    for w in write_list:
        assert isinstance(w,tuple)
        dat_type, dat, dests = w
        try:
            write_dat(dat_type, dat, dests=dests, **kwargs)
        except IOError:
            print_msg(__name__, "erro ao salvar arquivos ".join(dat_type), **kwargs)
            return False
    return True



def _write_dat(dat_type, dat_content, **kwargs):
    if (dat_content == []) or (dat_type == ""):
        return 1
    base_path = ""  # armazena o diret¢rio de trabalho (base ou diret¢rio passado como parÉmetro)
    clear_comments=kwargs.get("clear_comments",False)
    clear_field_comment = kwargs.get("clear_field_comment",False)
    ignore_includes = kwargs.get("ignore_includes",False)
    base = kwargs.get("base","")
    source_path = kwargs.get("source_path","")
    do_backup = kwargs.get("do_backup",True)
    dat_path=""

    # Define o caminho do dat em uma das opáîes na seguinte ordem de prioridade: base, caminho,
    # string ou diret¢rio corrente
    if base != "":
        dat_path = "/export/sage/sage/config/"+base+"/bd/dados/"+dat_type+".dat"
        base_path = os.path.dirname(dat_path)
    elif source_path != "":
        dat_path = source_path
        base_path = os.path.dirname(source_path)
    else:
        dat_path = os.path.join(os.path.curdir, dat_type + ".dat")
        base_path = os.path.dirname(dat_path)

    now = datetime.datetime.now()

    if os.path.isfile(dat_path):

        if do_backup:
            c = 1
            while 1:
                filename = os.path.join(base_path, dat_type + ".old_"+str(c))
                if os.path.isfile(filename):
                   c +=1
                else:
                    break
            print("dat: "+dat_path)
            print("novo: "+filename)
            os.rename(dat_path,filename)
        else:
            print("Removendo arquivo existente")
            os.remove(dat_path)

    if clear_comments:
        clear_comments(dat_content)
    if ignore_includes:
        dat_content = clear_includes(dat_content)


    output=""
    c = 0
    for d in dat_content:
        keys = list(d.keys())
        if "#" in keys[0]:
            include_path = keys[0]
            print(include_path)
            include_path = include_path.lstrip("#").strip()
            print(include_path)
            output += "#include "+include_path.lstrip(base_path)+"\n\n"
            _write_dat(dat_type, d[keys[0]], source_path=include_path, clear_comments=clear_comments, \
                       ignore_includes=ignore_includes, do_backup=do_backup)

            continue
        if d.get("comment") != None:
            output += d["comment"]+"\n"
        else:
            c += 1
            output += "\n" + dat_type.upper() + "\n"
            for k in d.keys():
                output += "\t"+str(k).upper()+" = "+d[k]+"\n"

    data = "{0}-{1}-{2} Ös {3}:{4}:{5}".format(now.year, now.month, now.day, now.hour, now.minute, now.second)
    total_lines = len(output.split("\n"))
    footer = \
    '''
;=================================================================================
;Arquivo criado pelo datapi.py beta 0.1 em {0}
;Total de linhas = {1}
;Total de pontos = {2}
;=================================================================================
\n'''.format(data, total_lines, c)
    output += footer
    print(dat_path)
    f = open(dat_path,"w",encoding="ISO-8859-1")
    f.write(output)
    f.close()

    return 0


def _ds_exists_in(dat_type, dataset, item_id="", item={}):
    '''
    Retorna o °ndice na lista em que a entrada com id_value foi encontrada, do contr†rio retorna None
    :param item:
    :param dat_type:
    :param item_id:
    :return:
    '''
    dat_pk = DAT_KEYS.get(dat_type,"")
    found = False
    if ((dat_pk) and (item_id)):
        for d in dataset:
            if d.get(DAT_KEYS[dat_type],"") == item_id:
                found = True
                break
    elif ((not dat_pk) and (item)) or ((dat_pk) and (item.get(dat_pk,""))):
        if item in dataset:
            found = True
    return found


def print_msg(mname, msg, msg_type=MSG_ERR, **kwargs):
    if kwargs.get("verbose", True):
        fname = inspect.currentframe().f_back.f_code.co_name
        print("({0}.{1}) {2}: {3}".format(mname, fname, msg_type, msg))



def exists_in(dat_type, generic_set, item_id="", item={}):
    '''
    Retorna True caso o item_id exista em generic_set, False caso contr†rio
    :param item:
    :param item_id:
    :return:
    '''

    if type(generic_set)==list:
        return _ds_exists_in(dat_type, generic_set, item_id=item_id, item=item)
    elif type(generic_set)==dict:
        found = False
        for k in list(generic_set.keys()):
            if _ds_exists_in(dat_type, generic_set[k], item_id=item_id, item=item):
                found = True
                break
        return found


def replace_item_text(item, old, new, fields=[], **kwargs):
    '''
    Substitui a string old pela string new no item passado, que deve ser no formato dict (ponto de dados).
    ê poss°vel passar uma lista fields com os campos que devem ser considerados para a substituiáÑo. Caso nÑo seja
    passado argumento fields, todos os campos de item sÑo considerados
    :param old:
    :param new:
    :return:
    '''
    if type(item)==dict:
        for k in list(item.keys()):
            if fields:
                if not k in fields:
                    continue
            item[k] = str(item[k]).replace(old, new)
    else:
        print_msg(__name__,"item nÑo Ç um dicion†rio", **kwargs)

def replace_text(dat_type, generic_set, old, new, fields=[]):
    '''
    Substitui o texto old por new em todos os itens do dataset (list) ou dat_content (dict).
    Pode ser passado uma lista de campos fields onde a substituiáÑo deve ocorrer. Caso nÑo seja passado, o texto
    ser† substitu°do em todos os campos
    :param old:
    :param new:
    :return:
    '''
    if type(generic_set) == list:
        for item in generic_set:
            replace_item_text(item, old, new, fields=fields)
    elif type(generic_set) == dict:
        for k in list(generic_set.keys()):
            for item in generic_set[k]:
                replace_item_text(item, old, new, fields=fields)



def find_item(dat_type, generic_set, item={}, item_id=""):
    """
    Retorna uma string com o °ndice de generic_set onde o item se encontra, e "" se o item nÑo for encontrado.
    :param dat_type:
    :param generic_set:
    :param item:
    :return:
    """
    if type(generic_set) == list:
        if item:
            if item in generic_set:
                return generic_set.index(item)
            else:
                return ""
        elif item_id:
            c=0
            for d in generic_set:
                if d.get(DAT_KEYS[dat_type]) == item_id:
                    return c
                c +=1
            return ""
    elif type(generic_set) == dict:
        if item:
            for k in list(generic_set.keys()):
                if item in generic_set[k]:
                    return str(k)
        elif item_id:
            for k in list(generic_set.keys()):
                if exists_in(dat_type, generic_set[k], item_id=item_id):
                    return str(k)

    return ""

def find_items(dat_type, generic_set, items=[], item_ids=[], where={}, op="and"):
    output = []
    if type(generic_set) == list:
        if items:
            for item in items:
                if item in generic_set:
                    output.append(generic_set.index(item))
        elif item_ids:
            c=0
            for d in generic_set:
                if d.get(DAT_KEYS[dat_type]) in item_ids:
                    output.append(c)
                c +=1
        elif where:
            for d in generic_set:
                if conditions_apply(datapoint=d, where_fields=where, op=op):
                    output.append(generic_set.index(d))

    elif type(generic_set) == dict:
        if items:
            for k in list(generic_set.keys()):
                for item in items:
                    if item in generic_set[k]:
                        output.append(str(k))
        elif item_ids:
            for k in list(generic_set.keys()):
                for item_id in item_ids:
                    if exists_in(dat_type, generic_set[k], item_id=item_id):
                        output.append(str(k))
        elif where:
            for k in list(generic_set.keys()):
                for d in generic_set[k]:
                    if conditions_apply(datapoint=d, where_fields=where, op=op):
                        output.append(str(k))
                        break

    return output


def field_to_string(field, rec):
    if field in rec:
        return str(field.strip().upper()+"= "+rec[field]+"\n")
    else:
        return None


def apostrophe(s):
    return "\"" + str(s)+"\""


def parse_expr(s):
    pattern = r"^\s*(?P<op1>(([><=!]?[<>=])|(has )|(has_prefix )|(has_sufix )))\s*(?P<arg>(.+))"
    regexp = re.compile(pattern)
    result = regexp.search(s)
    if result == None:
        return ()
    else:
        op1 = result.group("op1").strip()
        arg = result.group("arg").strip()
        return op1, arg


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def is_comment(item):
    '''
    Retorna True caso o item dict passado seja um coment†rio
    :param item:
    :return:
    '''
    return 'comment' in list(item.keys())

def is_commented_point(item):
    result = True
    for k in list(item.keys()):
        result = result and str(k).startswith(';')
    return result

def clean_commented_point(item):
    new_item={}
    for k in item.keys():
        nk=str(k).lstrip(';').strip()
        new_item[nk]=item[k]
    return new_item

def comment_point(item):
    new_item = {}
    for k in item.keys():
        nk = ';'+str(k)
        new_item[nk]=item[k]
    return new_item


def _ds_filter_fields(dataset, fields, allow_commented=False):

    """
    Retorna apenas os campos especificados na lista <fields> das entradas de <dataset>.
    :param dataset:
    :param fields:
    :param allow_commented:
    :return:
    """
    output = []
    for datapoint in dataset:
        filtered_point={}
        if fields != []:
            for field in fields:
                field = field.upper()
                if (field in datapoint):
                    filtered_point[field]=datapoint[field]
                elif ((allow_commented) and (";"+field in datapoint)):
                    filtered_point[";"+field]=datapoint[";"+field]
        else:
            output.append(datapoint)

        if len(filtered_point)>0:
            output.append(filtered_point)
    return output

def filter_fields(generic_set, fields, allow_commented=False):
    if type(generic_set) == list:
        return _ds_filter_fields(generic_set, fields, allow_commented=allow_commented)
    elif type(generic_set) == dict:
        output = []
        for k in list(generic_set.keys()):
            output.extend(_ds_filter_fields(generic_set[k], fields, allow_commented=allow_commented))
        output = list((o) for o in output if o)
        return output


def conditions_apply(datapoint, where_fields, op="and"):
    '''
    Retorna True caso a entrada <datapoint> obedeáa Ös condiáîes de <where_fields>
    :param datapoint:
    :param where_fields:
    :param op:
    :return:
    '''
    condition = ""
    for where_field in where_fields.keys():
        t = parse_expr(where_fields[where_field].strip())
        if t != ():
            op1, arg = t
            #print("op1="+op1)
            #print("arg="+arg)

        where_field = where_field.upper().strip()
        if where_field in datapoint:
            value = datapoint[where_field]
            #print("value="+value)
            #print("field="+where_field)
            if op1 == "has":
                condition += " ("+apostrophe(arg)+" in "+apostrophe(value)+") "+op
            elif op1 == "has_prefix":
                x=len(arg)
                condition += " ("+apostrophe(value[:x])+" == "+apostrophe(arg)+") "+op
            elif op1 == "has_sufix":
                x=len(arg)
                condition += " ("+apostrophe(value[-x:])+" == "+apostrophe(arg)+") "+op
            elif (is_number(arg)) and (op1 in ("==", ">=", "<=", "!=", "<>")):
                condition += " ("+value+op1+arg+") "+op
            elif op1 in ("==", ">=", "<=", "!=", "<>"):
                condition += " ("+apostrophe(value)+op1+apostrophe(arg)+") "+op
            else:
                #condition = "False"
                #break
                return False
        elif op=="and":
            #condition = "False"
            #break
            # caso o campo nÑo exista e a expressÑo seja AND, retorna como False
            return False
        else:
            # caso o campo nÑo exista e a expressÑo seja OR, continua, ignorando o campo
            continue
    condition = condition.rstrip(op).strip() # tira o £ltimo operador pra n dar erro no eval
    #print(condition)
    try:
        return eval(condition)
    except:
        return False


def get_all_dataset(dat_type, fields=[], base="", source_path="", source_str="", ignore_includes=False, \
                    allow_commented=False, **kwargs):

    clear_field_comment = not allow_commented
    dataset = _load_datx(dat_type=dat_type, base=base, source_path=source_path, source_str=source_str, \
                         ignore_includes=ignore_includes, clear_comments=True, concatenate=True, \
                         clear_field_comment=clear_field_comment)

    # caso condiáîes sejam passadas, aplicar filtro de condiáîes Ö lista
    if "where_fields" in kwargs:
        where_fields = kwargs["where_fields"]
        op = kwargs.get("op","and").lower()
        dataset = list((v) for v in dataset if conditions_apply(v,where_fields,op))

    # aplica o filtro de campos selecionados apenas, com a opáÑo de considerar campos comentados
    dataset = filter_fields(dataset, fields, allow_commented)

    return dataset

'''
def _ds_add_item(dat_type, dataset, item, **kwargs):
    item_id = item.get(DAT_KEYS[dat_type],"")
    add_or_update = kwargs.get("add_or_update",False)
    ignore_id = kwargs.get("ignore_id",False)

    #print(item_id)
    if (item_id == "") and (not ignore_id):
        print_msg(__name__, "id nÑo pode ser nulo", **kwargs)
    elif (item_id!="") and (exists_in(dat_type, dataset, item_id=item_id)):
        if add_or_update:
            update_item(dat_type, dataset, item=item, item_id=item_id)
        else:
            print_msg(__name__, "item j† existe no dataset", **kwargs)
    else:
        dataset.append(item)
'''

def add_item(dat_type, generic_set, item, **kwargs):
    '''
    Adiciona o item ao generic_set.
    :param dat_type: (str) tipo do dat ("pds", "pas", ...)
    :param generic_set: uma lista (dataset) com pontos ou um dict com o conte£do de um dat (dat_content)
    :param item: (dict) item a ser adicionado a generic_set
    :param kwargs:
        allow_duplicate: caso True inclui o item mesmo se j† existir em generic_set
        to_include: include onde deve ser adicionado (no formato #/dir/nome.dat), caso generic_set seja um dat_content
    :return:
    '''
    item = copy.deepcopy(item)
    dat_pk = DAT_KEYS.get(dat_type, "")
    item_id = item.get(dat_pk,"")
    to_include = kwargs.get("to_include","")
    allow_duplicate = kwargs.get("allow_duplicate", False)


    if (dat_pk) and (not item_id):
        print_msg(__name__, "item nÑo possui {0}".format(dat_pk), **kwargs)
        return False

    if (dat_pk) and exists_in(dat_type, generic_set, item_id=item_id) and (not allow_duplicate):
        print_msg(__name__, "j† existe item com mesmo {0} no destino".format(dat_pk), **kwargs)
        return False

    if (not dat_pk) and exists_in(dat_type, generic_set, item=item) and (not allow_duplicate):
        print_msg(__name__, "j† existe item idàntico no destino", **kwargs)
        return False

    if type(generic_set) == list:
        generic_set.append(item)
    elif type(generic_set) == dict:
        if to_include:
            if to_include not in list_includes(generic_set):
                generic_set[to_include]=[]
            generic_set[to_include].append(item)
        else:
            generic_set[dat_type].append(item)
    return True


def _ds_get_item(dat_type, dataset, item_id="", item={}, **kwargs):
    where = kwargs.get("where","")
    op = kwargs.setdefault("op","and")
    output = {}
    if item_id:
        for d in dataset:
            if d.get(DAT_KEYS[dat_type]) == item_id:
                output = d.copy()
                break
    elif item:
        for d in dataset:
            if d == item:
                output = d.copy()
                break
    elif where:
        for d in dataset:
            if conditions_apply(d, where, op):
                output = d.copy()
                break
    #print(str(len(dataset)) + ' ' + dat_type)
    return output


def get_item(dat_type, generic_set, item_id="", item={}, where={},op="and"):
    '''
    Retorna o primeiro item que encontrar, procurando primeiro pelo item_id e depois pelo parÉmetro item. Caso
    generic_set seja uma lista, retorna apenas o item no formado dict. Caso generic_set seja um dicion†rio de
    listas (formato do dat completo), get_item retorna uma tuple com o local (chave) do dicion†rio cuja lista
    se encontra o item, e o item em si.
    Exemplo:
    pds = load_pds(base="demo")
    location, item = get_item("pds", pds, item_id="S1_TIE_803")
    :param dat_type:
    :param generic_set:
    :param item_id:
    :param item:
    :param where:
    :param op:
    :return dict(item) ou str(location), dict(item):
    '''
    if type(generic_set) == list:
        return _ds_get_item(dat_type,generic_set,item_id=item_id, item=item, where=where, op=op)
    elif type(generic_set) == dict:
        output = {}
        location = ""
        output = _ds_get_item(dat_type, generic_set[dat_type],item_id=item_id,
                              item=item, where=where, op=op)
        if output:
            location = str(dat_type.lower())
        includes = list_includes(generic_set)
        if (len(includes) >0) and (output=={}): # se n achou item, procurar nos includes, caso haja includes
            for i in includes:
                output = _ds_get_item(dat_type,generic_set[i],item_id=item_id,
                                      item=item, where=where, op=op)
                if output != {}:
                    location = str(i)
                    break
        return location, output




def _ds_get_dataset(dat_type, dataset, id_set=[], item_set=[], **kwargs):
    #print(id_set)
    where = kwargs.get("where","")
    op = kwargs.get("op","and")
    output = []
    if id_set:
        for item_id in id_set:
            i = get_item(dat_type, dataset, item_id=item_id)
            output.append(i)
    elif item_set:
        for item in item_set:
            i = get_item(dat_type, dataset, item=item)
            output.append(i)
    elif where:
        output = list((d for d in dataset if conditions_apply(d, where, op)))
    else:
        output = copy.deepcopy(dataset)
    return output


def get_dataset(dat_type, generic_set, id_set=[], item_set = [], **kwargs):
    output = []
    if type(generic_set) == list:
        return _ds_get_dataset(dat_type=dat_type, dataset=generic_set, id_set=id_set,
                               item_set=item_set, **kwargs)
    elif type(generic_set) == dict:
        output = list(_ds_get_dataset(dat_type, generic_set[dat_type], id_set=id_set,
                                      item_set=item_set, **kwargs))
        includes = list_includes(generic_set)
        if (len(includes) >0):
            for i in includes:
                l = list(_ds_get_dataset(dat_type,generic_set[i],id_set=id_set,
                                         item_set=item_set, **kwargs))
                if l:
                    output.extend(l)
        output = list((o) for o in output if o)
        return output


def add_dataset(dat_type, generic_set, new_dataset, **kwargs):
    for d in new_dataset:
        add_item(dat_type, generic_set, d, **kwargs)


def delete_dataset(dat_type, generic_set, items=[], item_ids=[], **kwargs):
    '''
    Remove um ou mais itens de generic_set, que pode ser um objeto do tipo dict com o conte£do de um dat ou
    um dataset simples (lista).
    :param dat_type:
    :param generic_set:
    :param items: lista com os pontos (dict) que devem ser removidos
    :param item_ids: lista com os ids dos pontos (str) que devem ser removidos
    :param kwargs:
        where: dict com opáîes de seleáÑo de remoáÑo
        op: "and" (default) ou "or", operador l¢gico para os campos de where
    :return:
    '''
    ignore_id = kwargs.get("ignore_id",False)
    where = kwargs.get("where", {})
    op = kwargs.get("op", "and")
    d_id = DAT_KEYS.get(dat_type,"")
    if items:
        for d in items:
            delete_item(dat_type, generic_set, item=d)
    elif (item_ids) and (not ignore_id) and (d_id!=""):
        for item_id in item_ids:
            delete_item(dat_type, generic_set, item_id=item_id)
    elif where:
        delete_item(dat_type, generic_set, where=where, op=op)




def update_dataset(dat_type, generic_set, field_set, id_set=[], item_set=[], **kwargs):
    if id_set:
        c = 0
        if len(id_set) != len(field_set):
            print("(datapi.update_dataset) erro: field_set com tamanho diferente de id_set")
        else:
            for i in id_set:
                if (i == "") or (len(field_set[c]) == 0):
                    print("Erro: parÉmetros inv†lidos")
                else:
                    update_item(dat_type, generic_set, field_set[c], item_id=i)
                c += 1
    if item_set:
        c = 0
        if len(field_set) != len(item_set):
            print("(datapi.update_dataset) erro: field_set com tamanho diferente de item_set")
        else:
            for i in item_set:
                update_item(dat_type=dat_type, generic_set=generic_set, fields=field_set[c], \
                            item= item_set[c])



def _ds_delete_item(dat_type, dataset, item_id="", item={}, where={}, op="and"):
    if item_id:
        for d in dataset:
            if d.get(DAT_KEYS[dat_type]) == item_id:
                dataset.remove(d)
    elif item:
        if item in dataset: dataset.remove(item)
    elif where:
        temp = copy.deepcopy(dataset)
        for d in temp:
            if conditions_apply(datapoint=d, where_fields=where, op=op):
                dataset.remove(d)


def delete_item(dat_type, generic_set, item_id="", item={}, where={}, op="and"):
    if type(generic_set)==list:
        _ds_delete_item(dat_type,generic_set,item_id=item_id, item=item, where=where, op=op)
    elif type(generic_set)==dict:
        for k in list(generic_set.keys()):
            _ds_delete_item(dat_type,generic_set[k],item_id=item_id, item=item, where=where, op=op)


def _ds_update_item(dat_type, dataset, fields, item_id="", item={}, where={}, op="and"):
    cmp = item.copy()
    for d in dataset:
        if (d.get(DAT_KEYS[dat_type]) == item_id) or (d==cmp) or \
                conditions_apply(datapoint=d, where_fields=where, op=op):
            for k in fields.keys():
                d[k] = fields[k]




def update_item(dat_type, generic_set, fields, item_id="", item={}, where={}, op="and"):
    if type(generic_set) == list:
        _ds_update_item(dat_type, generic_set, fields, item_id=item_id, item=item, where=where, op=op)
    elif type(generic_set) == dict:
        for k in list(generic_set.keys()):
            _ds_update_item(dat_type, generic_set[k], fields, item_id=item_id, item=item, where=where, op=op)


def _ds_get_values(dat_type, dataset, field):
    output = []
    for d in dataset:
        output.append(d.get(field.upper(),""))
    return output


def get_values(dat_type, generic_set, field, unique = True):
    '''
    Retorna valores de um campo passado como parÉmetro. Se unique = True, retorna sem valores duplicados
    :param dat_type:
    :param generic_set:
    :param field:
    :return:
    '''
    output = []
    if type(generic_set) == list:
        output = _ds_get_values(dat_type, generic_set, field)
    elif type(generic_set) == dict:
        output = []
        for k in list(generic_set.keys()):
            output.extend(_ds_get_values(dat_type, generic_set[k],field))
        output = list((o) for o in output if o)
    if unique:
        # elimina duplicados transfromando em conjunto
        output = set(output)
        output = list(output)
    return output


def get_keys(dat_type, generic_set):
    return get_values(dat_type, generic_set, DAT_KEYS[dat_type])


def _ds_get_fieldnames(dataset):
    output = []
    for d in dataset:
        for k in list(d.keys()):
            if (not k in output) and (not k.startswith(";")): output.append(k)
    return output


def get_fieldnames(generic_set):
    '''
    Retorna uma lista de todos os campos utilizados no .dat, desconsiderando campos comentados
    :param generic_set:
    :return:
    '''
    output = []
    if type(generic_set) == list:
        output = _ds_get_fieldnames(generic_set)
    elif type(generic_set) == dict:
        for k in list(generic_set.keys()):
            output.extend(_ds_get_fieldnames(generic_set[k]))
    output = set(output)
    if "comment" in output: output.remove("comment")
    output = list(output)
    return output


def dist_report(nv2="", base="", source_path="", source_str=""):
    where_fields = {"tppnt": "== PDD"}
    if str(nv2) != "": where_fields["nv2"] = "== "+str(nv2)
    dataset = get_all_dataset(dat_type="pdf", base=base, source_path=source_path, source_str=source_str, \
                              where_fields=where_fields)
    max_address = 0
    pnt_list = ""
    for r in dataset:
        address = int(r["ORDEM"])
        if address > max_address: max_address = address
        pnt = r["PNT"]
        pnt_list += "{0} - {1} \n".format(pnt,address)

    output = \
    """
    ========================================================================================================
    Relat¢rio de distribuiáÑo {0}
    ========================================================================================================
    Total de pontos distribu°dos:       {1}
    Maior endereáo encontrado:          {2}

    Lista de pontos:
    --------------------------------------------------------------------------------------------------------
    Sinal - Endereáo
    {3}
    """.format(nv2, len(dataset), max_address, pnt_list)
    print(output)


def get_points(dat_type, ids, dataset = [], base="", source_path="", source_str="", \
              allow_commented=False, ignore_includes=False, **kwargs):

    if len(ids) > 0:
        if len(dataset) >0:
            output = dataset
        else:
            output = get_all_dataset(dat_type=dat_type, base=base, source_str=source_str, source_path=source_path, \
                                     allow_commented=allow_commented, ignore_includes=ignore_includes, kwargs=kwargs)
        output = list((v) for v in output if v.get(DAT_KEYS[dat_type],"") in ids)
        return output
    else:
        return []

#def get_pds_info(ids):


def list_includes(dat_content):
    output = []
    for k in list(dat_content.keys()):
            if k.startswith("#"):output.append(k)
    return output

def list_keys(dat_content):
    output = []
    output = sorted(list(dat_content.keys())[:],reverse=True)
    return output


def get_curr_includes(dat_path, **kwargs):
    #print(dat_path)
    try:
        s = get_file(dat_path, **kwargs)
    except:
        raise
    output = []
    dirname = os.path.dirname(dat_path)
    for line in s.splitlines():
        if line.strip().startswith("#"):
            line = line.lstrip("#include").strip()
            #line = fix_path(line)
            # alterado para ficar de acordo com as chaves do novo load_dat
            #line = os.path.join(dirname,line)
            line = "#"+line
            output.append(line)
    return output




def print_dat(dat_content, **kwargs):
    keys = list(dat_content.keys())
    keys.sort(reverse=True)
    dat_type = keys[0]
    kwargs.setdefault("field_order",DAT_FIELDS[dat_type])
    for d in list(dat_content.keys()):
        print("Listagem do arquivo: {0}".format(d))
        o,c = make_dat_str(dat_type,dat_content[d],**kwargs)
        print(o)

        #print_dataset(dat_content[d])

    print("\n")


def print_dataset(dataset):
    for d in dataset:
        for k in list(d.keys()):
            print("{0}   =   {1}".format(k, d[k]))
        print("\n")

# dt = "pds"
#
# recs = load_dat(dat_type=dt, source_path="bd/gvm/" ,concatenate=True)
#
# print_dat(recs)
#
# print(len(recs))
#
# d = get_dataset(dat_type=dt,generic_set=recs)
# n = {"ID":"srv3-dc-ssl-mid","TPNOH":"IHM"}
#
# key_list = filter_fields(recs, [DAT_KEYS[dt]])
#
# print(key_list)
#
# print(get_values(dt, recs, "ocr"))
#
# print(get_fieldnames(recs))
# #print("Adicionando "+str(n))
# #add_item(dt,recs,n)
#
# #print_dat(recs)
# #write_dat(dt,recs)
#
# print("===============  QUERY ================")
# #print(list_includes(recs))
# print(d)
# print(len(d))
# print(len(d)==len(recs[dt]))
