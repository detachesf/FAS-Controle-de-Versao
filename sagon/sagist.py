# -*- coding: cp860 -*-

from sagon.datapi import *

import copy

I_NSEQ = 0
I_ID = 1
I_DESCR = 2
I_PROC = 3

TCV_NONE = (0, "NLCN", "Ausˆncia de Conversor de Protocolo", "")
TCV_CNO = (1, "CNVA", "Conversor SINSC Modo Mestre", "cno")
TCV_C32 = (2, "CNVB", "Conversor CONITEL C3x00", "c32")
TCV_COS = (3, "CNVC", "Conversor SINSC Modo Escravo", "cos")
TCV_CNUL = (4, "CNVD", "Conversor Nulo Protocolo SAC", "cnul")
TCV_RDAC = (5, "CNVE", "Conversor REDAC-70 Westinghouse", "rdac")
TCV_LN57 = (6, "CNVF", "Conversor L&N IEC/60870-5 (LN57)", "ln57")
TCV_I101 = (7, "CNVG", "Conversor Siemens IEC/60870-5-101 (I101)", "i101")
TCV_DNP3 = (8, "CNVH", "Conversor DNP 3.0", "dnp3")
TCV_ABB = (9, "CNVI", "Conversor ABB 1771/X3.28", "abb")
TCV_MODB = (10, "CNVJ", "Conversor ModBus", "modb")
TCV_ALTS = (11, "CNVK", "Conversor ALTUS AL - 1000", "alts")
TCV_MLAB = (12, "CNVL", "Conversor MicroLab", "mlab")
TCV_I104 = (13, "CNVM", "Conversor IEC/60870-5-104", "i104")
TCV_ICCP = (14, "CNVN", "Conversor TASE2/ICCP-MMS", "iccp")
TCV_I61850 = (15, "CNVO", "Conversor IEC/61850", "i61850")
TCV_MLHD = (16, "CNVP", "Conversor Microlab-HLDC", "mlhd")


TCV = {
    TCV_NONE[I_ID]: TCV_NONE,
    TCV_CNO[I_ID]: TCV_CNO,
    TCV_C32[I_ID]: TCV_C32,
    TCV_COS[I_ID]: TCV_COS,
    TCV_CNUL[I_ID]: TCV_CNUL,
    TCV_RDAC[I_ID]: TCV_RDAC,
    TCV_LN57[I_ID]: TCV_LN57,
    TCV_I101[I_ID]: TCV_I101,
    TCV_DNP3[I_ID]: TCV_DNP3,
    TCV_ABB[I_ID]: TCV_ABB,
    TCV_MODB[I_ID]: TCV_MODB,
    TCV_ALTS[I_ID]: TCV_ALTS,
    TCV_MLAB[I_ID]: TCV_MLAB,
    TCV_I104[I_ID]: TCV_I104,
    TCV_ICCP[I_ID]: TCV_ICCP,
    TCV_I61850[I_ID]: TCV_I61850,
    TCV_MLHD[I_ID]: TCV_MLHD
}

TPAQS_ASAC = (0,"ASAC", "ASAC - Aquisi‡„o e controle")
TPAQS_ACSC = (1, "ACSC", "ACSC - C lculo est tico compilado")
TPAQS_AFIL = (2, "AFIL", "AFIL - Filtro localizado (parcelas de mesma CNF/LSC)")
TPAQS_AFID = (3, "AFID", "AFID - Filtro distribu¡do (parcelas de CNF/LSC diferentes)")
TPAQS_ACDI = (4, "ACDI", "ACDI - C lculo interpretado ou dinƒmico")
TPAQS_ASDE = (5, "ASDE", "ASDE - Sequˆncia de eventos (Conitel/ModBus)")


TPAQS = {
    TPAQS_ASAC[I_ID]: TPAQS_ASAC,
    TPAQS_ACSC[I_ID]: TPAQS_ACSC,
    TPAQS_AFIL[I_ID]: TPAQS_AFIL,
    TPAQS_AFID[I_ID]: TPAQS_AFID,
    TPAQS_ACDI[I_ID]: TPAQS_ACDI,
    TPAQS_ASDE[I_ID]: TPAQS_ASDE,
    }


TTP_NONE = (0, "NLTP", "Ausˆncia de Transportador de Protocolo", "")
TTP_MLX25 = (1, "MLX25", "Transp. de Multiliga‡”es em Protocolo X25/X75", "mlx25")
TTP_MLTCP = (2, "MLTCP", "Transp. de Multiliga‡”es em Protocolo TCP/IP", "tcpd")
TTP_DNPF3 = (3, "DNPF3", "Transp. em Frames FT3 modo balanceado do DNP", "iec3d")
TTP_IECF3 = (4, "IECF3", "Transp. em Frames FT3 do IEC/60870-5", "iecd")
TTP_CXTCP = (5, "CXTCP", "Transp. de Conex”es Virtuais UTR em TCP/IP", "tcps")
TTP_CXX25 = (6, "CXX25", "Transp. de Conex”es Virtuais UTR em X25", "x25d")
TTP_CXA32 = (7, "CXA32", "Transp. de Conex”es Virtuais UTR em Frames Ass. 32 Bits", "a32d")
TTP_IECF1 = (8, "IECF1", "Transp. em Frames balanceado/n„o balanceado FT1.2 do IEC/870-5", "iec1d/iec2d")
TTP_CX328 = (9, "CX328", "Transp. em Frames ANSI X3.28", "x328d")
TTP_UDPF3 = (10, "UDPF3", "Transp. em Frames FT3-DNP do IEC/60870 sobre UDP", "iec3u")
TTP_TCPF1 = (11, "TCPF1", "Transp. Balanceado em Frames FT1.2 do IEC/60870 sobre TCP", "iec1t")
TTP_CXTT3 = (12, "CXTT3", "Transp. em Frames FT3-LN57 do IEC/60870 sobre TTY", "iecy")
TTP_CXTT1 = (13, "CXTT1", "Transp. N„o balanceado em Frames FT1.2 do IEC/60870 sobre TTY", "iec2y")
TTP_CXTTD = (14, "CXTTD", "Transp. em Frames FT3 DNP do IEC/60870 sobre TTY", "iec3y")
TTP_YMBUS = (15, "YMBUS", "Transp. de Frames ModBus em Enlaces TTY/TCP-IP", "ybus")
TTP_ALTUS = (16, "ALTUS", "Transp. de Frames ALTUS AI-1000 em enlaces TTY/TCP-IP", "alty")
TTP_CXTB1 = (17, "DNPF3", "Transp. Balanceado em Frames FT1.2 do IEC/60870 sobre TTY", "iec1y")
TTP_SPTDS = (18, "SPTDS", "Transp. de Multiliga‡”es em Linhas Seriais Ass¡ncronas", "sptd")
TTP_YMLAB = (19, "YMLAB", "Transp. de Frames do Protocolo MicroLab sobre TTY", "ylab")
TTP_CX104 = (20, "CX104", "Transp. TCP-IP do IEC/60870-5-104", "iec4t")
TTP_PCTR = (21, "PCTR", "Transp. de Multiliga‡”es em Datagramas UDP", "pctr")
TTP_MMST = (22, "MMST", "Transp. de Multiliga‡”es em Conex”es MMS/TCP-IP", "mmst")
TTP_CXTU1 = (23, "CXTU1", "Transp. Balanceado Bidirecional FT 1.2 do IEC/60870 sobre UDP", "iec1u")
TTP_CXTY1 = (24, "CXTY1", "Transp. Balanceado Bidirecional FT 1.2 do IEC/60870 sobre TTY", "iec1b")
TTP_CA32Y = (25, "CA32Y", "Transp. de Conex”es Virtuais UTR em Frames Ass¡nc. 32 bits sob TTY", "a32y")
TTP_TMBUS = (26, "TMBUS", "Transportador de Frames Open-MODBUS TCP/IP", "tbus")
TTP_YHDLC = (27, "YHDLC", "Transportador de Frames Ass¡ncronos HDLC sobre TTY", "hdlc")
TTP_IEC1S = (28, "IEC1S", "Transp. Balanceado em Frames FT1.2 do IEC/60870 p/ Term. Server", "iec1s")
TTP_IEC2S = (29, "IEC2S", "Transp. N„o Balanceado em Frames FT1.2 do IEC/60870 p/ Term. Server", "iec2s")
TTP_IEC3S = (30, "IEC3S", "Transportador em Frames FT3-DNP do IEC/60870 para Terminal Server", "iec3s")
TTP_SMBUS = (31, "SMBUS", "Transporte de Frames MODBUS para Terminal Server", "smbus")
TTP_A32S = (32, "A32S", "Transportador em Frames Ass¡ncronos 32 Bits para terminal server", "a32s")
TTP_SMLAB = (33, "SMLAB", "Transportador de Frames do Protocolo Microlab para terminal server", "smlab")
TTP_SHDLC = (34, "SHDLC", "Transporte de Frames Ass¡ncronos HDLC para Terminal Server", "shdlc")
TTP_TSNMP = (35, "TSNMP", "Transporte de SNMP", "tsnmp")
TTP_IEC2T = (36, "IEC2T", "Transp. Nao Balanceado em Frames FT1.2 do IEC/60870 sobre TCP", "iec2t")

TTP = {
    TTP_NONE[I_ID]: TTP_NONE,
    TTP_MLX25[I_ID]: TTP_MLX25,
    TTP_MLTCP[I_ID]: TTP_MLTCP,
    TTP_DNPF3[I_ID]: TTP_DNPF3,
    TTP_IECF3[I_ID]: TTP_IECF3,
    TTP_CXTCP[I_ID]: TTP_CXTCP,
    TTP_CXX25[I_ID]: TTP_CXX25,
    TTP_CXA32[I_ID]: TTP_CXA32,
    TTP_IECF1[I_ID]: TTP_IECF1,
    TTP_CX328[I_ID]: TTP_CX328,
    TTP_UDPF3[I_ID]: TTP_UDPF3,
    TTP_TCPF1[I_ID]: TTP_TCPF1,
    TTP_CXTT3[I_ID]: TTP_CXTT3,
    TTP_CXTT1[I_ID]: TTP_CXTT1,
    TTP_CXTTD[I_ID]: TTP_CXTTD,
    TTP_YMBUS[I_ID]: TTP_YMBUS,
    TTP_ALTUS[I_ID]: TTP_ALTUS,
    TTP_CXTB1[I_ID]: TTP_CXTB1,
    TTP_SPTDS[I_ID]: TTP_SPTDS,
    TTP_YMLAB[I_ID]: TTP_YMLAB,
    TTP_CX104[I_ID]: TTP_CX104,
    TTP_PCTR[I_ID]: TTP_PCTR,
    TTP_MMST[I_ID]: TTP_MMST,
    TTP_CXTU1[I_ID]: TTP_CXTU1,
    TTP_CXTY1[I_ID]: TTP_CXTY1,
    TTP_CA32Y[I_ID]: TTP_CA32Y,
    TTP_TMBUS[I_ID]: TTP_TMBUS,
    TTP_YHDLC[I_ID]: TTP_YHDLC,
    TTP_IEC1S[I_ID]: TTP_IEC1S,
    TTP_IEC2S[I_ID]: TTP_IEC2S,
    TTP_IEC3S[I_ID]: TTP_IEC3S,
    TTP_SMBUS[I_ID]: TTP_SMBUS,
    TTP_A32S[I_ID]: TTP_A32S,
    TTP_SMLAB[I_ID]: TTP_SMLAB,
    TTP_SHDLC[I_ID]: TTP_SHDLC,
    TTP_TSNMP[I_ID]: TTP_TSNMP,
    TTP_IEC2T[I_ID]: TTP_IEC2T
    }

LSC_TIPO_AA = (0, "AA", "AA - Liga‡„o de aquisi‡„o")
LSC_TIPO_DD = (1, "DD", "DD - Liga‡„o de distribui‡„o")
LSC_TIPO_AD = (2, "AD", "AD - Liga‡„o de aquisi‡„o e distribui‡„o")

LSC_TIPO = {
    LSC_TIPO_AA[I_ID]: LSC_TIPO_AA,
    LSC_TIPO_DD[I_ID]: LSC_TIPO_DD,
    LSC_TIPO_AD[I_ID]: LSC_TIPO_AD
}



PHY_SIGNAL = "Ponto de aquisi‡„o ou controle"
CAL_SIGNAL = "Ponto calculado"
FIL_SIGNAL = "Ponto resultado de filtro"

def load_pds(**kwargs):
    return load_dat(dat_type="pds", **kwargs)


def load_pdf(**kwargs):
    return load_dat(dat_type="pdf", **kwargs)


def load_pdd(**kwargs):
    return load_dat(dat_type="pdd", **kwargs)


def load_lsc(**kwargs):
    return load_dat(dat_type="lsc", **kwargs)


def load_nv1(**kwargs):
    return load_dat(dat_type="nv1", **kwargs)


def load_nv2(**kwargs):
    return load_dat(dat_type="nv2", **kwargs)


def load_cnf(**kwargs):
    return load_dat(dat_type="cnf", **kwargs)


def load_tac(**kwargs):
    return load_dat(dat_type="tac", **kwargs)


def load_pas(**kwargs):
    return load_dat(dat_type="pas", **kwargs)


def load_paf(**kwargs):
    return load_dat(dat_type="paf", **kwargs)


def load_pad(**kwargs):
    return load_dat(dat_type="pad", **kwargs)


def load_cgf(**kwargs):
    return load_dat(dat_type="cgf", **kwargs)


def load_cgs(**kwargs):
    return load_dat(dat_type="cgs", **kwargs)


def load_rca(**kwargs):
    return load_dat(dat_type="rca", **kwargs)


def load_pts(**kwargs):
    return load_dat(dat_type="pts", **kwargs)


def load_ptf(**kwargs):
    return load_dat(dat_type="ptf", **kwargs)


def load_ptd(**kwargs):
    return load_dat(dat_type="ptd", **kwargs)


def load_ocr(**kwargs):
    return load_dat(dat_type="ocr", **kwargs)


def load_gsd(**kwargs):
    return load_dat(dat_type="gsd", **kwargs)


def load_noh(**kwargs):
    return load_dat(dat_type="noh", **kwargs)


def load_pro(**kwargs):
    return load_dat(dat_type="pro", **kwargs)


def load_inp(**kwargs):
    return load_dat(dat_type="inp", **kwargs)


def load_rfi(**kwargs):
    return load_dat(dat_type="rfi", **kwargs)


def load_rfc(**kwargs):
    return load_dat(dat_type="rfc", **kwargs)

def load_mul(**kwargs):
    return load_dat(dat_type="mul", **kwargs)

def load_enm(**kwargs):
    return load_dat(dat_type="enm", **kwargs)

def load_cxu(**kwargs):
    return load_dat(dat_type="cxu", **kwargs)

def load_utr(**kwargs):
    return load_dat(dat_type="utr", **kwargs)

def load_map(**kwargs):
    return load_dat(dat_type="map", **kwargs)

def load_ctx(**kwargs):
    return load_dat(dat_type="ctx", **kwargs)

def load_cxp(**kwargs):
    return load_dat(dat_type="cxp", **kwargs)

def load_e2m(**kwargs):
    return load_dat(dat_type="e2m", **kwargs)

def load_grcmp(**kwargs):
    return load_dat(dat_type="grcmp", **kwargs)

def load_grupo(**kwargs):
    return load_dat(dat_type="grupo", **kwargs)

def load_inm(**kwargs):
    return load_dat(dat_type="inm", **kwargs)

def load_ins(**kwargs):
    return load_dat(dat_type="ins", **kwargs)

def load_psv(**kwargs):
    return load_dat(dat_type="psv", **kwargs)

def load_sev(**kwargs):
    return load_dat(dat_type="sev", **kwargs)

def load_sxp(**kwargs):
    return load_dat(dat_type="sxp", **kwargs)

def load_tcl(**kwargs):
    return load_dat(dat_type="tcl", **kwargs)

def load_tctl(**kwargs):
    return load_dat(dat_type="tctl", **kwargs)

def load_tcv(**kwargs):
    return load_dat(dat_type="tcv", **kwargs)

def load_tdd(**kwargs):
    return load_dat(dat_type="tdd", **kwargs)

def load_tn1(**kwargs):
    return load_dat(dat_type="tn1", **kwargs)

def load_tn2(**kwargs):
    return load_dat(dat_type="tn2", **kwargs)

def load_ttp(**kwargs):
    return load_dat(dat_type="ttp", **kwargs)





def add_pds_calc_(pds_id, function="", fields={}, parcs=[], clone_id="", **kwargs):
    """

    :param pds_id:
    :param function:
    :param fields:
    :param parcs: lista com as parcelas do c lculo, onde cada elemento ‚ um dicion rio com os atributos da
    entrada na tabela RCA. Exemplo de uso:
    parc1 = {"PARC": "PONTO_PDS_1", "TIPOP": "EDC", "ORDEM": "1", "TPPARC": "PDS"}
    parc2 = {"PARC": "PONTO_PDS_2", "TIPOP": "EDC", "ORDEM": "2", "TPPARC": "PDS"}
    parcs = [parc1, parc2]
    add_pds_calc(pds_id="PONTO_PDS_3", function="OU", parcs=parcs)
    :param clone_id:
    :param kwargs:
    :return:
    """
    pds = load_dat(dat_type="pds",**kwargs)
    rca = load_dat(dat_type="rca", **kwargs)

    if exists_in(dat_type="pds", generic_set=pds, item_id=pds_id):
        #print_msg(**kwargs, "j  existe pds com este ID", msg_type=ERR_MSG)
        #print("(sagist.add_pds_calc): j  existe pds com este ID")
        return False

    pds_item = {}

    rca_set = []

    if clone_id:
        w, clone_item = get_item(dat_type="pds",generic_set=pds, item_id=clone_id)
        if not clone_item:
            #print_msg("ponto a ser clonado n„o existe", **kwargs)
            return False
        if clone_item.get("TCL") == "NLCL":
            #print_msg("ponto a ser clonado n„o ‚ calculado", **kwargs)
            return False
        print(clone_id)
        #print(clone_item)
        rca_clone_set = get_dataset("rca", rca, where={"PNT": "== "+str(clone_id)})
        #print(rca_clone_set)
        if len(rca_clone_set) == 0:
            #print_msg(**kwargs, "n„o h  parcelas configuradas para o ponto a ser clonado", msg_type=ERR_MSG)
            return False
        rca_set = copy.deepcopy(rca_clone_set)
        pds_item = clone_item.copy()
    else:
        if not parcs:
            #print_msg("parcelas n„o foram definidas e n„o h  ponto a ser clonado", **kwargs)
            return False
        rca_set = copy.deepcopy(parcs)

    i=1
    for r in rca_set:
        r.setdefault("TPPNT","PDS")
        r.setdefault("ORDEM",str(i))
        r["PNT"] = pds_id
        i += 1

    for k in list(fields.keys()):
        pds_item[str(k.upper())] = fields[k]

    pds_item["ID"] = pds_id
    pds_item.setdefault("TPFIL","NLFL")
    pds_item.setdefault("NOME",pds_id)
    if function:
        pds_item["TCL"] = function
    elif pds_item.get("TCL") is None:
        print("(sagist.add_pds_calc) erro: n„o h  fun‡„o especificada")
        return False


    tac = load_dat("tac", **kwargs)
    if not pds_item.get("TAC"):
        # buscar primeira TAC de c lculos que encontrar
        #print_msg(__name__,"n„o foi informada TAC, buscando uma TAC de c lculos na base...",
         #           msg_type=WARN_MSG, **kwargs)
        w, tac_item = get_item("tac", tac, where={"TPAQS": "== ACSC"})
        if not tac_item:
            #print_msg(__name__, "n„o h  TAC de c lculos configurada na base", **kwargs)
            return False
        pds_item["TAC"] = tac_item["ID"]
    else:
        w, tac_item = get_item("tac", tac, item_id=pds_item["TAC"])


    add_to = kwargs.get("add_to","")

    if add_to:
        print(add_to)
        pds_location = make_include_str("pds", **kwargs)
        pds_location = "#"+pds_location
        print("PDS LOCAL="+pds_location)
        rca_location = make_include_str("rca", **kwargs)
        rca_location = "#"+rca_location
        print("RCA LOCAL="+rca_location)
    elif clone_id:
        pds_location = find_item("pds", pds, clone_item)
        rca_location = find_item("rca", rca, rca_clone_set[0])
        print(pds_location)
        print(rca_location)
    else:
        pds_location="pds"
        rca_location = "rca"

    if pds_location in list(pds.keys()):
        add_item("pds",pds[pds_location],pds_item)
    else:
        pds[pds_location]=[pds_item]

    if rca_location in list(rca.keys()):
        add_dataset("rca",rca[rca_location],rca_set, ignore_id=True)
    else:
        rca[rca_location]=[]
        add_dataset("rca", rca[rca_location],rca_set, ignore_id=True)

    print(list_includes(rca))

    print(pds_item)
    print(rca_set)

    print("\n\n")

    print_dataset(pds[pds_location])

    print("\n\n")

    print_dataset(rca[rca_location])

    try:
        write_dat("pds", pds, dests=[pds_location], **kwargs)
    except IOError:
        print_msg(__name__, "erro ao salvar arquivo: " + pds_location.lstrip("#"), **kwargs)
        return False
    try:
        write_dat("rca", rca, dests=[rca_location], **kwargs)
    except IOError:
        print_msg(__name__, "erro ao salvar arquivo: " + rca_location.lstrip("#"), **kwargs)
        return False

    print_msg(__name__, "ponto adicionado com sucesso a {0} e {1}".format(pds_location.lstrip("#"),
                                                                          rca_location.lstrip("#")), msg_type=MSG_WARN, **kwargs)
    return True


def add_item_calc(dat_type, item_id, function="", fields={}, parcs=[], clone_id="", **kwargs):
    """

    :param item_id:
    :param function:
    :param fields:
    :param parcs: lista com as parcelas do c lculo, onde cada elemento ‚ um dicion rio com os atributos da
    entrada na tabela RCA. Exemplo de uso:
    parc1 = {"PARC": "PONTO_PDS_1", "TIPOP": "EDC", "ORDEM": "1", "TPPARC": "PDS"}
    parc2 = {"PARC": "PONTO_PDS_2", "TIPOP": "EDC", "ORDEM": "2", "TPPARC": "PDS"}
    parcs = [parc1, parc2]
    add_pds_calc(item_id="PONTO_PDS_3", function="OU", parcs=parcs)
    :param clone_id:
    :param kwargs:
    :return:
    """

    dat_type = dat_type.lower()
    if not dat_type in ("pds", "pas", "pts"):
        print_msg(__name__, "o item deve ser um ponto l¢gico (pds, pas ou pts)", **kwargs)
        return False

    ds = load_dat(dat_type=dat_type,**kwargs)
    rca = load_dat(dat_type="rca", **kwargs)

    if exists_in(dat_type=dat_type, generic_set=ds, item_id=item_id):
        print_msg(__name__, "j  existe {0} com este ID".format(dat_type), **kwargs)
        #print("(sagist.add_pds_calc): j  existe pds com este ID")
        return False

    ds_item = {}

    rca_set = []

    if clone_id:
        w, clone_item = get_item(dat_type=dat_type,generic_set=ds, item_id=clone_id)
        if not clone_item:
            print_msg(__name__, "ponto a ser clonado n„o existe", **kwargs)
            return False
        if clone_item.get("TCL") == "NLCL":
            print_msg(__name__, "ponto a ser clonado n„o ‚ calculado", **kwargs)
            return False
        #print(clone_id)
        #print(clone_item)
        rca_clone_set = get_dataset("rca", rca, where={"PNT": "== "+str(clone_id)})
        #print(rca_clone_set)
        if len(rca_clone_set) == 0:
            print_msg(__name__, "n„o h  parcelas configuradas para o ponto a ser clonado", **kwargs)
            return False
        rca_set = copy.deepcopy(rca_clone_set)
        ds_item = clone_item.copy()
    else:
        if not parcs:
            print_msg(__name__, "parcelas n„o foram definidas e n„o h  ponto a ser clonado", **kwargs)
            return False
        rca_set = copy.deepcopy(parcs)

    i=1
    for r in rca_set:
        r.setdefault("TPPNT",dat_type.upper())
        r.setdefault("ORDEM",str(i))
        r["PNT"] = item_id
        i += 1

    for k in list(fields.keys()):
        ds_item[str(k.upper())] = fields[k]

    ds_item["ID"] = item_id
    ds_item.setdefault("TPFIL","NLFL")
    ds_item.setdefault("NOME", item_id)
    if function:
        ds_item["TCL"] = function
    elif ds_item.get("TCL") is None:
        print_msg(__name__, "n„o h  fun‡„o especificada", **kwargs)
        return False

    tac = load_dat("tac", **kwargs)
    if not ds_item.get("TAC"):
        # buscar primeira TAC de c lculos que encontrar
        print_msg(__name__, "n„o foi informada TAC, buscando uma TAC de c lculos na base...",
                  msg_type=MSG_WARN, **kwargs)
        w, tac_item = get_item("tac", tac, where={"TPAQS": "== ACSC"})
        if not tac_item:
            print_msg(__name__, "n„o h  TAC de c lculos configurada na base", **kwargs)
            return False
        ds_item["TAC"] = tac_item["ID"]
    else:
        w, tac_item = get_item("tac", tac, item_id=ds_item["TAC"])

    add_to = kwargs.get("add_to","")

    if add_to:
        #print(add_to)
        ds_location = make_include_str(dat_type, **kwargs)
        #ds_location = "#"+ds_location
        #print("PDS LOCAL="+ds_location)
        rca_location = make_include_str("rca", **kwargs)
        #rca_location = "#"+rca_location
        #print("RCA LOCAL="+rca_location)
    elif clone_id:
        ds_location = find_item(dat_type, ds, clone_item)
        rca_location = find_item("rca", rca, rca_clone_set[0])
        #print(ds_location)
        #print(rca_location)
    else:
        ds_location = dat_type
        rca_location = "rca"

    if ds_location in list(ds.keys()):
        add_item(dat_type, ds[ds_location], ds_item)
    else:
        ds[ds_location]=[ds_item]

    if rca_location in list(rca.keys()):
        add_dataset("rca",rca[rca_location],rca_set, ignore_id=True)
    else:
        rca[rca_location]=[]
        add_dataset("rca", rca[rca_location],rca_set, ignore_id=True)

    #print(list_includes(rca))

    #print(ds_item)
    #print(rca_set)

    #print("\n\n")

    #print_dataset(ds[ds_location])

    #print("\n\n")

    #print_dataset(rca[rca_location])

    try:
        write_dat(dat_type, ds, dests=[ds_location], **kwargs)
    except:
        print_msg(__name__, "erro ao salvar arquivo: " + ds_location.lstrip("#"), **kwargs)
        return False
    try:
        write_dat("rca", rca, dests=[rca_location], **kwargs)
    except:
        print(__name__, "erro ao salvar arquivo: "+rca_location.lstrip("#"), **kwargs)
        return False

    print_msg(__name__, "ponto adicionado com sucesso a {0} e {1}".format(ds_location.lstrip("#"),
                                                                          rca_location.lstrip("#")), msg_type=MSG_WARN)
    return True


def add_pds_calc(pds_id, function="", fields={}, parcs=[], clone_id="", **kwargs):
    add_item_calc(dat_type="pds", item_id=pds_id, function=function, fields=fields, \
                  parcs=parcs, clone_id=clone_id, **kwargs)


def add_pas_calc(pas_id, function="", fields={}, parcs=[], clone_id="", **kwargs):
    add_item_calc(dat_type="pas", item_id=pas_id, function=function, fields=fields, \
                  parcs=parcs, clone_id=clone_id, **kwargs)


def add_pts_calc(pts_id, function="", fields={}, parcs=[], clone_id="", **kwargs):
    add_item_calc(dat_type="pts", item_id=pts_id, function=function, fields=fields, \
                  parcs=parcs, clone_id=clone_id, **kwargs)


def add_item_61850(dat_type, item_id, logical_device="", address="", fields={}, clone_id ="", **kwargs):
    '''
    Insere ponto em l¢gico e faz as altera‡”es necess rias na tabela do ponto f¡sico.

    :param logical_device: string com valor do atributo CONFIG da tabela NV1. Representa a primeira parte do
    endere‡o 61850 do ponto. Caso o logical_device n„o seja definido, deve ser passado o argumento clone para
    que o novo ponto seja adicionado com o mesmo logical device do ponto clone
    :param address: string com valor do endere‡o 61850, sem o logical device. Este campo, junto com o ID do NV1 referente ao
    logical device, forma o atributo ID do ponto PDF relacionado. Caso address n„o seja definido, o argumento clone
    deve ser passado para que o endere‡o possa ser copiado do clone
    :param clone_id: string com ID do ponto PDS a ser clonado. Caso este argumento exista, o novo ponto PDS ter  seus
    atributos copiados do ponto clone, assim como o ponto PDF relacionado. Caso logical_device ou address n„o sejam
    definidos, os valores do clone s„o usados.
    :param fields: dicion rio com os valores que se quer atribuir aos atributos PDS ou PDF do novo ponto.
    Formato do campo fields: {pds.nome : valor, pds.alrin: valor, pdf.desc1: valor, etc...}
    Os atributos devem ser strings come‡ando por "pds." ou "pdf.", indicando a que tabela o atributo deve ser adicionado.
    Observe que mesmo que um ponto clone seja passado como parƒmetro, ‚ poss¡vel sobrescrever atributos individualmente
    caso os mesmos estejam no parƒmetro fields
    :param kwargs:
        do_backup (default True): faz backup dos arquivos dat e pdf antes de adicionar o ponto. Os arquivos
     existentes s„o renomeados para a forma dat.old_00x, onde x ‚ um n£mero sequencial.
        source_path (string): caminho para o diret¢rio principal da base. N„o inserir / ou \ no in¡cio. Ex.:
    source_path = "bd/demo/"
        base (string): nome da base onde o ponto deve ser adicionado. Serve apenas para utiliza‡„o no Linux, pois
    procura no diret¢rio /exports/home/sage/sage/config/<base>/bd/dados/

    :return: False - caso n„o consiga adicionar o ponto, True - caso o ponto seja adicionado com sucesso

    O ponto criado considera os seguintes valors padr”es, caso os mesmos n„o sejam definidos no parƒmetro fields:
    PDS.ALRIN = PDS.SOIN = NAO
    PDS.NONME = PDS.ID
    PDF.DESC1 = PDS.NOME
    PDF.CMT = logical_device
    Caso base ou source_path n„o existam, considera o diret¢rio atual como reposit¢rio principal dos .dats.

    Exemplo de uso:
    fields = {"pds.nome": "Sele‡„o de Religamento Tripolar"}
    fields["pdf.kconv"] = "SPS0"
    add_pds_61850(pds_id="GVM:04C7:F1:SRET", fields=fields, logical_device="UC1_04C7CONTROL", \
              address="GosGGIO1$ST$Ind23", source_path = "bd/gvm/", do_backup = False)

    ou:

    fields = {"pds.nome": "Sele‡„o de Religamento Tripolar"}
    add_pds_61850(pds_id="GVM:04C7:F1:SRET", fields=fields, clone_id="GVM:04C7:F1:SREM", base = "ssl-gvm", \
                address="GosGGIO1$ST$Ind23", do_backup = False)

    '''
    dat_type = dat_type.lower()
    if not dat_type in ("pds", "cgs", "pas", "pts"):
        print_msg(__name__, "item deve ser ponto ou controle l¢gico (pds, pas, pts, cgs)", **kwargs)
        return False
    dat_s = load_dat(dat_type, **kwargs)
    nv1 = load_dat("nv1", **kwargs)
    dat_type_f = dat_type.replace("s","f")
    dat_f = load_dat(dat_type_f, **kwargs)

    ds_item = {}
    df_item = {}

    if dat_type == "cgs":
        foreign_key_field = "CGS"
    else:
        foreign_key_field = "PNT"

    if clone_id:
        w, ds_clone_item = get_item(dat_type=dat_type,generic_set=dat_s, item_id=clone_id)
        if not ds_clone_item:
            print_msg(__name__, "ponto a ser clonado n„o existe em " + dat_type, **kwargs)
            return False
        w, df_clone_item = get_item(dat_type=dat_type_f, generic_set=dat_f, where={foreign_key_field: "== "+ds_clone_item["ID"]})
        if not df_clone_item:
            print_msg(__name__, "ponto a ser clonado n„o possui ponto f¡sico configurado", **kwargs)
            return False

        if not logical_device: # caso n„o tenha sido definido logical device, pegar do clone
            nv1_clone_id = str(df_clone_item["ID"]).split("-")[0]
            print_msg(__name__, "clone nv1 = "+ nv1_clone_id, MSG_INFO, **kwargs)
            w, nv1_clone_item = get_item(dat_type="nv1", generic_set=nv1,item_id=nv1_clone_id)
            if not nv1_clone_item:
                print_msg(__name__,"ponto clonado n„o possui NV1 definido e n„o h  logical device de " \
                      "entrada", **kwargs)
                return False
            logical_device = nv1_clone_item["CONFIG"]
            print_msg(__name__, "clone logical device = "+ logical_device, MSG_INFO, **kwargs)

        if not address: # caso n„o tenha sido passado address, pegar do clone
            address = str(df_clone_item["ID"]).split("-")[1]


        # popula pds_item e pdf_item com os valores clones
        for k in list(ds_clone_item.keys()):
            ds_item[str(k)] = ds_clone_item[k]
        for k in list(df_clone_item.keys()):
            df_item[str(k)] = df_clone_item[k]

    if (not item_id) or (not logical_device) or (not address):
        print_msg(__name__, "especifique valores de entrada ou um ponto para clonagem", **kwargs)
        return False

    if exists_in(dat_type, dat_s, item_id=item_id):
        print_msg(__name__, "ponto j  existe em "+dat_type, **kwargs)
        return False

    w, nv1_item = get_item("nv1", nv1, where={"CONFIG": "== "+logical_device})
    if not nv1_item:
        print_msg(__name__, "o logical device n„o ‚ v lido", **kwargs)
        return False

    nv2 = load_dat("nv2", **kwargs)
    w, nv2_item = get_item("nv2", nv2, where={"NV1": "== "+nv1_item["ID"], "TPPNT": "== "+dat_type_f.upper()})
    if not nv2_item:
        print_msg(__name__, "o logical device n„o possui aquisi‡„o de pontos digitais configurada", **kwargs)
        return False

    df_item_id = nv1_item["ID"]+"-"+address
    if exists_in(dat_type_f, dat_f, item_id=df_item_id):
        print_msg(__name__, "j  existe ponto configurado com o mesmo endere‡o em "+dat_type_f, **kwargs)
        return False

    cnf = load_dat("cnf", **kwargs)
    w, cnf_item = get_item("cnf", cnf, where={"ID": "== "+nv1_item["CNF"]})
    if not cnf_item:
        print_msg(__name__, "o CNF n„o parece estar corretamente configurado para este LOGICAL DEVICE (nv1.dat)", **kwargs)
        return False

    tac = load_dat("tac", **kwargs)
    w, tac_item = get_item("tac", tac, where={"LSC": "== "+cnf_item["LSC"] })
    if not tac_item:
        print_msg(__name__, "n„o h  TAC configurada para este LOGICAL DEVICE (tac.dat)", **kwargs)
        return False

    # preenche ds_item e df_item com os valores dos campos passados em fields
    for k in list(fields.keys()):
        table, field = str(k).split(".")
        if table.lower() == dat_type:
            ds_item[str(field).upper()] = fields[k]
        elif table.lower() == dat_type_f:
            df_item[str(field).upper()] = fields[k]
        print_msg(__name__, "adicionando {0} a {1} com valor {2}".format(field, table, fields[k]), MSG_INFO, **kwargs)

    # seta valores dos campos obrigat¢rios
    ds_item["ID"] = item_id
    if not clone_id: ds_item["TAC"] = tac_item["ID"]
    ds_item["TCL"] = "NLCL"
    ds_item["TPFIL"] = "NLFL"

    df_item["ID"] = df_item_id
    if not clone_id: df_item["NV2"] = nv2_item["ID"]
    df_item[foreign_key_field] = item_id
    if dat_type != "cgs": df_item["TPPNT"] = dat_type.upper()

    # checa e seta valores default
    #ds_item.setdefault("ALRIN","NAO")
    #ds_item.setdefault("SOEIN","NAO")
    ds_item.setdefault("NOME",item_id)
    # caso seja cgs, seta ponto de retorno padr„o um pds de mesmo nome
    if dat_type == "cgs":
        ds_item.setdefault("PAC",item_id)
        ds_item.setdefault("TIPO","PDS")


    df_item.setdefault("DESC1",ds_item["NOME"])
    df_item.setdefault("CMT", logical_device)

    add_to = kwargs.get("add_to","")

    if add_to:
        #print(add_to)
        ds_location = make_include_str(dat_type, **kwargs)
        print_msg(__name__, "o {0} ser  salvo em {1}".format(dat_type,ds_location), MSG_WARN, **kwargs)
        df_location = make_include_str(dat_type_f, **kwargs)
        print_msg(__name__, "o {0} ser  salvo em {1}".format(dat_type_f, df_location), MSG_WARN, **kwargs)
    else:
         # checa onde est  o logical device (include ou .dat principal, para inserir o ponto no mesmo local
        item_location = find_item("nv1",nv1,nv1_item)
        #item_path = fix_path(os.path.dirname(item_location).lstrip("#"))
        ds_location = str(item_location).replace("nv1",dat_type)
        df_location = str(item_location).replace("nv1",dat_type_f)


    #print(pds_item)
    #print(pdf_item)
    #print(pds_location)
    #print(pdf_location)
    #print(item_path)

    if ds_location in list(dat_s.keys()):
        add_item(dat_type,dat_s[ds_location],ds_item)
    else:
        dat_s[ds_location]=[ds_item]

    if df_location in list(dat_f.keys()):
        add_item(dat_type_f,dat_f[df_location],df_item)
    else:
        dat_f[df_location]=[df_item]

    #print_dataset(dat[pds_location])
    #print_dataset(pdf[pdf_location])
    try:
        write_dat(dat_type, dat_s, dests=[ds_location], **kwargs)
    except:
        print_msg(__name__, "n„o foi poss¡vel salvar arquivo: "+ds_location.lstrip("#"), **kwargs)
        return False
    try:
        write_dat(dat_type_f, dat_f, dests=[df_location], **kwargs)
    except:
        print_msg(__name__, "n„o foi poss¡vel salvar arquivo: "+df_location.lstrip("#"), **kwargs)
        return False

    print_msg(__name__, "ponto adicionado com sucesso a {0} e {1}".format(dat_type, dat_type_f), MSG_INFO, **kwargs)
    return True


def add_pds_61850(pds_id, logical_device="", address="", fields={}, clone_id="", **kwargs):
    add_item_61850("pds", item_id=pds_id, logical_device=logical_device, address=address, \
                   fields=fields, clone_id=clone_id, **kwargs)


def add_pas_61850(pas_id, logical_device="", address="", fields={}, clone_id="", **kwargs):
    add_item_61850("pas", item_id=pas_id, logical_device=logical_device, address=address, \
                   fields=fields, clone_id=clone_id, **kwargs)


def add_pts_61850(pts_id, logical_device="", address="", fields={}, clone_id="", **kwargs):
    add_item_61850("pts", item_id=pts_id, logical_device=logical_device, address=address, \
                   fields=fields, clone_id=clone_id, **kwargs)


def add_cgs_61850(cgs_id, logical_device="", address="", fields={}, clone_id="", **kwargs):
    add_item_61850("cgs", item_id=cgs_id, logical_device=logical_device, address=address, \
                   fields=fields, clone_id=clone_id, **kwargs)


def remove_item_61850(dat_type, item_id="", item={}, **kwargs):
    dts = dat_type.lower()
    s_item_id = item_id
    dtf = dts.replace("s", "f")
    force_remove = kwargs.get("force_remove", False)
    cascade = kwargs.get("cascade", False)

    xxs = load_dat(dat_type=dts, **kwargs)
    xxf = load_dat(dat_type=dtf, **kwargs)
    rca = load_rca(**kwargs)
    rfi = load_rfi(**kwargs)
    rfc = load_rfc(**kwargs)

    if dts== "cgs":
        where_df = {"CGS": "== " + s_item_id}

    else:
        where_df = {"PNT": "== " + s_item_id, "TPPNT": "== " + dts.upper()}
        dtd = dts.replace("s", "d")
        xxd = load_dat(dat_type=dtd, **kwargs)

    w, s_item = get_item(dat_type=dts, generic_set=xxs, item_id=item_id, item=item)

    s_item_id = s_item.get("ID","")

    if not s_item:
        print_msg(__name__, "o item {0} n„o existe em {1}".format(item_id, dts), **kwargs)
        return False
    elif not s_item_id:
        print_msg(__name__, "o item n„o possui id", **kwargs)
        return False

    print_msg(__name__, "item encontrado: {0} - {1}".format(s_item["ID"],s_item["NOME"]), MSG_INFO, **kwargs)

    f_items = get_dataset(dat_type=dtf, generic_set=xxf, where=where_df)
    if not f_items:
        print_msg(__name__, "ponto n„o possui ponto f¡sico associado", **kwargs)
        return False

    f_item_ids = list(v["ID"] for v in f_items if v.get("ID") is not None)
    if not f_item_ids:
        print_msg(__name__, "n„o h  id(s) associado(s) em {0}".format(xxf), **kwargs)
        return False
    else:
        print_msg(__name__, "encontrados os seguintes pontos f¡sicos em {0}: {1}".format(dtf, f_item_ids),
                  MSG_INFO, **kwargs)

    if not force_remove: # faz com que pontos que estejam sendo usados n„o possam ser removidos
        # checar se o item ‚ parcela de algum ponto calculado
        parcs_rca = get_dataset("rca", rca, where={"PARC": "== " + s_item_id, "TPPARC": "== " + dts.upper()})
        if parcs_rca:
            print(parcs_rca)
            print_msg(__name__, "o ponto est  sendo usado como parcela de {0} em rca.dat."
                  " Para for‡ar a remo‡„o use a op‡„o force_remove=True".format(parcs_rca[0]["PNT"]), **kwargs)
            return False
        # checar se os endere‡os f¡sicos fazem parte de um filtro
        found_rfc = False
        for f_item_id in f_item_ids:
            parcs_rfc = get_dataset(dat_type="rfc", generic_set=rfc,
                                    where={"PARC":"== "+f_item_id, "TPPARC": "== "+dtf.upper()})
            if parcs_rfc:
                print_msg(__name__, "o ponto est  sendo usado como parcela de {0} em rfc.dat."
                      " Para for‡ar a remo‡„o use a op‡„o force_remove=True".format(parcs_rfc[0]["PNT"]), **kwargs)
                found_rfc = True

        if found_rfc: return False

        # checar se est  sendo usado em algum comando
        cgs = load_cgs(**kwargs)
        linked_cgs = get_dataset(dat_type="cgs", generic_set=cgs,
                                 where={"PAC":"== " + s_item_id, "TIPO": "== " + dts.upper()})
        if linked_cgs:
            print(linked_cgs)
            print_msg(__name__, "o ponto est  sendo usado como retorno do comando {0} em cgs.dat."
                  " Para for‡ar a remo‡„o use a op‡„o force_remove=True".format(linked_cgs[0]["PAC"]), **kwargs)
            return False

        linked_cgs = get_dataset(dat_type="cgs", generic_set=cgs, where={"PINT":"== " + s_item_id})

        if (linked_cgs) and (dts == "pds"):
            print(linked_cgs)
            print_msg(__name__, "o ponto est  sendo usado como intertravamento do comando {0} "
                  "em cgs.dat. Para for‡ar a remo‡„o use a op‡„o force_remove=True".format(linked_cgs[0]["ID"]),
                      **kwargs)
            return False

        # checar se est  sendo usado em filtro simples
        found_rfi = False
        for f_item_id in f_item_ids:
            parcs_rfi = get_dataset(dat_type="rfi", generic_set=rfi,
                                    where={"PNT":"== "+f_item_id, "TIPOP": "== "+dtf.upper()})
            if parcs_rfi:
                print_msg(__name__, "o ponto {0} est  sendo usado como parcela em rfi.dat."
                            " Para for‡ar a remo‡„o use a op‡„o force_remove=True".format(f_item_id), **kwargs)
                found_rfi = True
        if found_rfi: return False

    s_location = find_item(dat_type=dts, generic_set=xxs, item=s_item)

    write_list = []
    print_msg(__name__, "removendo {0} de {1} ...".format(s_item_id, dts), MSG_INFO, **kwargs)
    delete_item(dat_type=dts, generic_set=xxs, item_id=s_item_id)
    write_list.append((dts, xxs, [s_location]))

    if cascade:
        print_msg(__name__, "remo‡„o em cascata selecionada, procurando itens vinculados...", MSG_INFO, **kwargs)
        f_locations = []
        if not dts == "cgs":
            # procura o pdd/pad/ptd para deletar tamb‚m
            # pode haver mais de uma distribui‡„o para o mesmo ponto, por isso get_dataset (?)
            d_items = get_dataset(dat_type=dtd, generic_set=xxd,
                                    where={str(dts.upper()): "== " + s_item_id})
            d_locations = []
            for d_item in d_items:

                d_locations.append(find_item(dat_type=dtd, generic_set=xxd, item=d_item))

                w, fd_item = get_item(dat_type=dtf, generic_set=xxf,
                                    where={"TPPNT": "== "+dtd.upper(), "PNT": "== "+d_item.get("ID","")})
                if fd_item:
                    dsfd_location = find_item(dat_type=dtf, generic_set=xxf, item=fd_item)
                    f_locations.append(dsfd_location)
                    print_msg(__name__, "removendo {0} de {1}".format(str(fd_item["ID"]),str(dtf)), MSG_INFO, **kwargs)
                    delete_item(dat_type=dtf, generic_set=xxf, item=fd_item)

                print_msg(__name__, "removendo {0} de {1}".format(d_item["ID"],dtd), MSG_INFO, **kwargs)
                delete_item(dat_type=dtd, generic_set=xxd, item=d_item)

            d_locations = set(d_locations)
            write_list.append((dtd, xxd, d_locations))

        for f_item in f_items:
            f_locations.append(find_item(dat_type=dtf, generic_set=xxf, item=f_item))
            print_msg(__name__, "removendo {0} de {1}".format(f_item["ID"],dtf), MSG_INFO, **kwargs)
            delete_item(dat_type=dtf, generic_set=xxf, item=f_item)

        f_locations = set(f_locations)

        write_list.append((dtf, xxf, f_locations))

    bulk_write_dat(write_list, **kwargs)


def remove_item_calc(dat_type, item_id="", item={}, **kwargs):
    dts = dat_type.lower()
    dtf = dts.replace("s", "f")
    dtd = dts.replace("s", "d")

    s_item_id = item_id
    s_item = item

    force_remove = kwargs.get("force_remove", False)
    cascade = kwargs.get("cascade", False)

    xxs = load_dat(dat_type=dts, **kwargs)
    xxf = load_dat(dat_type=dtf, **kwargs)
    xxd = load_dat(dat_type=dtd, **kwargs)
    rca = load_rca(**kwargs)

    w, s_item = get_item(dat_type=dts, generic_set=xxs, item_id=item_id, item=item)

    s_item_id = s_item.get("ID","")

    if not s_item:
        print_msg(__name__, "o item {0} n„o existe em {1}".format(item_id, dts), **kwargs)
        return False
    elif not s_item_id:
        print_msg(__name__, "o item n„o possui id", **kwargs)
        return False

    print_msg(__name__, "item encontrado: {0} - {1}".format(s_item["ID"],s_item["NOME"]), MSG_INFO, **kwargs)

    if not force_remove: # faz com que pontos que estejam sendo usados n„o possam ser removidos
        # checar se o item ‚ parcela de algum ponto calculado
        parcs_rca = get_dataset("rca", rca, where={"PARC": "== " + s_item_id, "TPPARC": "== " + dts.upper()})
        if parcs_rca:
            #print(parcs_rca)
            print_msg(__name__, "o ponto est  sendo usado como parcela de {0} em rca.dat."
                  " Para for‡ar a remo‡„o use a op‡„o force_remove=True".format(parcs_rca[0]["PNT"]), **kwargs)
            return False

        # checar se est  sendo usado em algum comando
        cgs = load_cgs(**kwargs)
        linked_cgs = get_dataset(dat_type="cgs", generic_set=cgs,
                                 where={"PAC":"== " + s_item_id, "TIPO": "== " + dts.upper()})
        if linked_cgs:
            #print(linked_cgs)
            print_msg(__name__, "o ponto est  sendo usado como retorno do comando {0} em cgs.dat."
                  " Para for‡ar a remo‡„o use a op‡„o force_remove=True".format(linked_cgs[0]["PAC"]), **kwargs)
            return False

        linked_cgs = get_dataset(dat_type="cgs", generic_set=cgs, where={"PINT":"== " + s_item_id})

        if (linked_cgs) and (dts == "pds"):
            #print(linked_cgs)
            print_msg(__name__, "o ponto est  sendo usado como intertravamento do comando {0} "
                  "em cgs.dat. Para for‡ar a remo‡„o use a op‡„o force_remove=True".format(linked_cgs[0]["ID"]),
                      **kwargs)
            return False

    s_location = find_item(dat_type=dts, generic_set=xxs, item=s_item)

    write_list = []
    print_msg(__name__, "removendo {0} de {1} ...".format(s_item_id, dts), MSG_INFO, **kwargs)
    delete_item(dat_type=dts, generic_set=xxs, item_id=s_item_id)
    write_list.append((dts, xxs, [s_location]))

    if cascade:
        print_msg(__name__, "remo‡„o em cascata selecionada, procurando itens vinculados...", MSG_INFO, **kwargs)
        f_locations = []

        # procura o pdd/pad/ptd para deletar tamb‚m
        # pode haver mais de uma distribui‡„o para o mesmo ponto, por isso get_dataset (?)
        d_items = get_dataset(dat_type=dtd, generic_set=xxd,
                                where={str(dts.upper()): "== " + s_item_id})
        d_locations = []
        for d_item in d_items:

            d_locations.append(find_item(dat_type=dtd, generic_set=xxd, item=d_item))

            w, fd_item = get_item(dat_type=dtf, generic_set=xxf,
                                where={"TPPNT": "== "+dtd.upper(), "PNT": "== "+d_item.get("ID","")})
            if fd_item:
                dsfd_location = find_item(dat_type=dtf, generic_set=xxf, item=fd_item)
                f_locations.append(dsfd_location)
                print_msg(__name__, "removendo {0} de {1}".format(fd_item["ID"],dtf), MSG_INFO, **kwargs)
                delete_item(dat_type=dtf, generic_set=xxf, item=fd_item)

            print_msg(__name__, "removendo {0} de {1}".format(d_item["ID"],dtd), MSG_INFO, **kwargs)
            delete_item(dat_type=dtd, generic_set=xxd, item=d_item)

        rca_locations = []
        rca_items = get_dataset(dat_type="rca", generic_set=rca,
                                where={"PNT": "== "+s_item_id, "TPPNT": "== "+dts.upper()})
        for rca_item in rca_items:
            rca_locations.append(find_item(dat_type="rca", generic_set=rca, item=rca_item))
            print_msg(__name__, "removendo parcela {0} ({1}) de {2} em rca".format(rca_item["ORDEM"],
                      rca_item["PARC"], rca_item["PNT"]), MSG_INFO, **kwargs)
            delete_item(dat_type="rca", generic_set=rca, item=rca_item)

        rca_locations = set(rca_locations)
        if rca_locations:
            write_list.append(("rca", rca, rca_locations))

        d_locations = set(d_locations)
        if d_locations:
            write_list.append((dtd, xxd, d_locations))

        f_locations = set(f_locations)
        if f_locations:
            write_list.append((dtf, xxf, f_locations))

    bulk_write_dat(write_list, **kwargs)


def update_item_61850(dat_type, item_id="", item={}, fields={}, **kwargs):
    dts = dat_type.lower()
    s_item_id = item_id
    dtf = dts.replace("s", "f")
    update_related = kwargs.get("update_related", False)

    xxs = load_dat(dat_type=dts, **kwargs)
    xxf = load_dat(dat_type=dtf, **kwargs)
    rca = load_rca(**kwargs)
    rfi = load_rfi(**kwargs)
    rfc = load_rfc(**kwargs)

    if dts== "cgs":
        where_df = {"CGS": "== " + s_item_id}

    else:
        where_df = {"PNT": "== " + s_item_id, "TPPNT": "== " + dts.upper()}
        dtd = dts.replace("s", "d")
        xxd = load_dat(dat_type=dtd, **kwargs)

    w, s_item = get_item(dat_type=dts, generic_set=xxs, item_id=item_id, item=item)

    s_item_id = s_item.get("ID","")

    if not s_item:
        print_msg(__name__, "o item {0} n„o existe em {1}".format(item_id, dts), **kwargs)
        return False
    elif not s_item_id:
        print_msg(__name__, "o item n„o possui id", **kwargs)
        return False

    print_msg(__name__, "item encontrado: {0} - {1}".format(s_item["ID"],s_item["NOME"]), MSG_INFO, **kwargs)

    f_items = get_dataset(dat_type=dtf, generic_set=xxf, where=where_df)
    if not f_items:
        print_msg(__name__, "ponto n„o possui ponto f¡sico associado", **kwargs)
        return False

    f_item_ids = list(v["ID"] for v in f_items if v.get("ID") is not None)
    if not f_item_ids:
        print_msg(__name__, "n„o h  id(s) associado(s) em {0}".format(xxf), **kwargs)
        return False
    else:
        print_msg(__name__, "encontrados os seguintes pontos f¡sicos em {0}: {1}".format(dtf, f_item_ids),
                  MSG_INFO, **kwargs)

    d_items = []
    d_item_ids = []
    if not dts == "cgs":
        d_items = get_dataset(dat_type=dtd, generic_set=xxd, where={dts.upper(): "== "+s_item_id})
        if not d_items:
            print_msg(__name__, "ponto n„o possui distribui‡„o", MSG_INFO, **kwargs)
        else:
            d_item_ids = list(v["ID"] for v in d_items if v.get("ID") is not None)
            if not d_item_ids:
                print_msg(__name__, "n„o h  id(s) associado(s) em {0}".format(xxd), **kwargs)
                return False
            else:
                print_msg(__name__, "encontrados os seguintes pontos de distribui‡„o em {0}: {1}".format(dtd, d_item_ids),
                          MSG_INFO, **kwargs)


    write_list = []

    s_fields = {}
    f_fields = {}
    d_fields = {}

    for k in list(fields.keys()):
        table, field = str(k).split(".")
        if table.lower() == dts:
            s_fields[str(field).upper()] = fields[k]
        elif table.lower() == dtf:
            f_fields[str(field).upper()] = fields[k]
        elif table.lower() == dtd:
            d_fields[str(field).upper()] = fields[k]

    s_id_change = False
    f_id_change = False
    d_id_change = False

    if (s_fields.get("ID") is not None) and (s_fields.get("ID") != s_item_id):
        new_s_item_id = s_fields.get("ID")
        if not new_s_item_id:
            print_msg(__name__, "n„o ‚ poss¡vel atualizar para um valor de id nulo em {0}".format(dts),
                      **kwargs)
            return False
        elif exists_in(dat_type=dts, generic_set=xxs, item_id=new_s_item_id):
            print_msg(__name__, "n„o ‚ poss¡vel atualizar: id {0} existe em {1}".format(new_s_item_id, dts),
                      **kwargs)
            return False
        print_msg(__name__, "h  mudan‡a de id em {0}: {1} -> {2}".format(dts, s_item_id, new_s_item_id),
                  MSG_WARN, **kwargs)
        s_id_change = True

    if (f_fields.get("ID") is not None) and (not set(f_fields.get("ID")).issubset(set(f_item_ids))):
        new_f_item_ids = []
        if type(f_fields.get("ID")) == str:
            new_f_item_ids.append(f_fields.get("ID"))
        elif type(f_fields.get("ID")) == list:
            new_f_item_ids.extend(f_fields.get("ID"))
        if not new_f_item_ids:
            print_msg(__name__, "n„o ‚ poss¡vel atualizar para um valor de id nulo em {0}".format(dtf),
                      **kwargs)
            return False
        elif len(new_f_item_ids) != len(f_item_ids):
            print_msg(__name__, "para atualizar mais de um endere‡o f¡sico do mesmo ponto, deve-se passar "
                      "uma lista com o mesmo n£mero de endere‡os f¡sicos existentes no campo pdf.id: {0}".format(
                len(f_item_ids)
            ), **kwargs)
            return False
        else:
            for i in new_f_item_ids:
                if exists_in(dat_type=dtf, generic_set=xxf, item_id=i):
                    print_msg(__name__, "n„o ‚ poss¡vel atualizar: id {0} existe em {1}".format(i, dtf),
                              **kwargs)
                    return False
        print_msg(__name__, "h  mudan‡a de id em {0}: {1} -> {2}".format(dtf, f_item_ids, new_f_item_ids),
                  MSG_WARN, **kwargs)
        f_id_change = True

    if (d_fields.get("ID") is not None) and (not set(d_fields.get("ID")).issubset(set(d_item_ids))) and (dts != "cgs"):
        new_d_item_ids = []
        if type(d_fields.get("ID"))== str:
            new_d_item_ids.append(d_fields.get("ID"))
        elif type(d_fields.get("ID")) == list:
            new_d_item_ids.extend(d_fields.get("ID"))
        #print(new_d_item_ids)
        if not new_d_item_ids:
            print_msg(__name__, "n„o ‚ poss¡vel atualizar para um valor de id nulo em {0}".format(dtd),
                      **kwargs)
            return False
        elif len(new_d_item_ids) != len(d_item_ids):
            print_msg(__name__, "para atualizar mais de um endere‡o de distribui‡„o do mesmo ponto, deve-se passar "
                      "uma lista com o mesmo n£mero de endere‡os existentes no campo id: {0}".format(
                len(d_item_ids)
            ), **kwargs)
            return False
        else:
            for i in new_d_item_ids:
                if exists_in(dat_type=dtd, generic_set=xxd, item_id=i):
                    print_msg(__name__, "n„o ‚ poss¡vel atualizar: id {0} existe em {1}".format(i, dtd),
                              **kwargs)
                    return False
        print_msg(__name__, "h  mudan‡a de id em {0}: {1} -> {2}".format(dtd, d_item_ids, new_d_item_ids),
                  MSG_WARN, **kwargs)
        d_id_change = True

    s_locations = []
    f_locations = []
    d_locations = []

    if update_related:

        rca_locations = []
        cgs_locations = []
        rfi_locations = []
        rfc_locations = []

        if s_id_change:
            print_msg(__name__,"atualizando dats relacionados … mudan‡a de id l¢gico...", MSG_INFO, **kwargs)
            # atualiza itens de rca que usam este ponto
            rca_related_set = get_dataset(dat_type="rca", generic_set=rca,
                                      where={"PARC": "== "+s_item_id, "TPPARC": "== "+dts.upper()})
            for rca_related in rca_related_set:
                rca_locations.append(find_item("rca", rca, rca_related))
                print_msg(__name__, "atualizando parcela {0} do ponto {1} em rca: {2} -> {3}".format(
                    rca_related["ORDEM"], rca_related["PNT"], rca_related["PARC"], new_s_item_id
                ), MSG_INFO, **kwargs)
                update_item(dat_type="rca", generic_set=rca, fields={"PARC": new_s_item_id},
                            item=rca_related)

            cgs = load_cgs(**kwargs)

            # atualiza PACs de cgs que usam este ponto
            cgs_related_set = get_dataset(dat_type="cgs", generic_set=cgs,
                                          where={"PAC": "== "+s_item_id, "TIPO": "== "+dts.upper()})
            for cgs_related in cgs_related_set:
                cgs_locations.append(find_item("cgs", cgs, cgs_related))
                print_msg(__name__, "atualizando PAC do ponto {0} em cgs: {1} -> {2}".format(
                    cgs_related["ID"], cgs_related["PAC"], new_s_item_id), MSG_INFO, **kwargs)
                update_item(dat_type="cgs", generic_set=cgs, fields={"PAC": new_s_item_id},
                            item=cgs_related)

            # atualiza PINT de cgs que usam este ponto, caso seja um pds
            if dts == "pds":
                cgs_related_set = get_dataset(dat_type="cgs", generic_set=cgs,
                                          where={"PINT": "== "+s_item_id})
                for cgs_related in cgs_related_set:
                    cgs_locations.append(find_item("cgs", cgs, cgs_related))
                    print_msg(__name__, "atualizando PINT do ponto {0} em cgs: {1} -> {2}".format(
                        cgs_related["ID"], cgs_related["PINT"], new_s_item_id), MSG_INFO, **kwargs)
                    update_item(dat_type="cgs", generic_set=cgs, fields={"PINT": new_s_item_id},
                                item=cgs_related)

            # atualiza id do ponto l¢gico no ponto f¡sico
            for f_item in f_items:
                f_locations.append(find_item(dtf, xxf, f_item))
                print_msg(__name__, "atualizando ponto {0} em {1} com novo id l¢gico: {2}".format(
                        f_item["ID"], dtf, new_s_item_id), MSG_INFO, **kwargs)
                if dts == "cgs":
                    up_fields = {"CGS": new_s_item_id}
                else:
                    up_fields = {"PNT": new_s_item_id}
                update_item(dat_type=dtf, generic_set=xxf, fields=up_fields, item=f_item)

            # atualiza id do ponto l¢gico no ponto de distribui‡„o
            if not dts == "cgs":
                for d_item in d_items:
                    d_locations.append(find_item(dtd, xxd, d_item))
                    print_msg(__name__, "atualizando ponto {0} em {1} com novo id l¢gico: {2}".format(
                        d_item["ID"], dtd, new_s_item_id), MSG_INFO, **kwargs)
                    update_item(dat_type=dtd, generic_set=xxd, fields={dts.upper(): new_s_item_id}, item=d_item)

        # caso haja mudan‡a de endere‡o f¡sico
        if f_id_change:
            print_msg(__name__,"atualizando dats relacionados … mudan‡a de id f¡sico...", MSG_INFO, **kwargs)

            # atualiza itens de rfc que usam este ponto
            c = 0
            for f_item_id in f_item_ids:
                rfc_related_set = get_dataset(dat_type="rfc", generic_set=rfc,
                                          where={"PARC": "== "+f_item_id, "TPPARC": "== "+dtf.upper()})
                for rfc_related in rfc_related_set:
                    rfc_locations.append(find_item("rfc", rfc, rfc_related))
                    print_msg(__name__, "atualizando parcela {0} do ponto {1} em rfc: {2} -> {3}".format(
                        rfc_related["ORDEM"], rfc_related["PNT"], rfc_related["PARC"], new_f_item_ids[c]
                    ), MSG_INFO, **kwargs)
                    update_item(dat_type="rfc", generic_set=rfc, fields={"PARC": new_f_item_ids[c]},
                                item=rfc_related)

                # atualiza itens de rfi que usam este ponto
                rfi_related_set = get_dataset(dat_type="rfi", generic_set=rfi,
                                          where={"PNT": "== "+f_item_id, "TIPOP": "== "+dtf.upper()})
                for rfi_related in rfi_related_set:
                    rfi_locations.append(find_item("rfi", rfi, rfi_related))
                    print_msg(__name__, "atualizando parcela {0} do ponto em rfi: {1} -> {2}".format(
                        rfi_related["ORDEM"], rfi_related["PNT"], new_f_item_ids[c]
                    ), MSG_INFO, **kwargs)
                    update_item(dat_type="rfi", generic_set=rfi, fields={"PNT": new_f_item_ids[c]},
                                item=rfi_related)
                c += 1

        if d_id_change:
            print_msg(__name__,"atualizando dats relacionados … mudan‡a de id de distribui‡„o...", MSG_INFO, **kwargs)

            c = 0
            for d_item_id in d_item_ids:
                f_related_set = get_dataset(dat_type=dtf, generic_set=xxf,
                                            where={"PNT": "== "+d_item_id, "TPPNT": "== "+dtd.upper()})
                for f_related in f_related_set:
                    f_locations.append(find_item(dtf, xxf, f_related))
                    print_msg(__name__, "atualizando id de distribui‡„o de {0} em {1}: {2} -> {3}".format(
                        f_related["ID"], dtf, d_item_id, new_d_item_ids[c]
                    ), MSG_INFO, **kwargs)
                    update_item(dat_type=dtf, generic_set=xxf, fields={"PNT": new_d_item_ids[c]},
                                item=f_related)
                c += 1

        rca_locations = list(set(rca_locations))
        rfc_locations = list(set(rfc_locations))
        rfi_locations = list(set(rfi_locations))
        cgs_locations = list(set(cgs_locations))

        if rca_locations: write_list.append(("rca", rca, rca_locations))
        if rfc_locations: write_list.append(("rfc", rfc, rfc_locations))
        if rfi_locations: write_list.append(("rfi", rfi, rfi_locations))
        if cgs_locations: write_list.append(("cgs", cgs, cgs_locations))

    print_msg(__name__,"atualizando {0} com os novos valores: {1}".format(dts, s_fields), MSG_INFO, **kwargs)
    s_locations.append(find_item(dts, xxs, s_item))
    update_item(dat_type=dts, generic_set=xxs, item=s_item, fields=s_fields)
    s_locations = list(set(s_locations))
    print(dts+":"+str(s_locations))
    if s_locations: write_list.append((dts, xxs, s_locations))

    print_msg(__name__,"atualizando {0} com os novos valores: {1}".format(dtf, f_fields), MSG_INFO, **kwargs)
    c = 0
    for f_item_id in f_item_ids:
        f_locations.append(find_item(dtf, xxf, item_id=f_item_id))
        if f_id_change:
            f_fields["ID"]=new_f_item_ids[c]
        else:
            f_fields["ID"]=f_item_id
        update_item(dat_type=dtf, generic_set=xxf, item_id=f_item_id, fields=f_fields)
        c += 1
    f_locations = list(set(f_locations))
    print(dtf+":"+str(f_locations))
    if f_locations: write_list.append((dtf, xxf, f_locations))

    if dts != "cgs":
        c = 0
        print_msg(__name__,"atualizando {0} com os novos valores: {1}".format(dtd, d_fields), MSG_INFO, **kwargs)
        for d_item_id in d_item_ids:
            d_locations.append(find_item(dtd, xxd, item_id=d_item_id))
            if d_id_change:
                d_fields["ID"]=new_d_item_ids[c]
            else:
                d_fields["ID"]=d_item_id
            update_item(dat_type=dtd, generic_set=xxd, item_id=d_item_id, fields=d_fields)
            c += 1
        d_locations = list(set(d_locations))
        print(dtd+":"+str(d_locations))
        if d_locations: write_list.append((dtd, xxd, d_locations))

    bulk_write_dat(write_list, **kwargs)


def find_item_in_base(dat_type, item_id="", item={}, **kwargs):
    dat = load_dat(dat_type, **kwargs)
    if dat:
        return find_item(dat_type=dat_type, generic_set=dat, item_id=item_id, item=item)
    else:
        return ""

def find_items_in_base(dat_type, item_id="", items=[], item_ids=[], where={}, op="and", **kwargs):
    dat = load_dat(dat_type, **kwargs)
    if dat:
        return find_items(dat_type=dat_type, generic_set=dat, item_ids=item_ids, items=items, where=where, op=op)
    else:
        return []



def find_pds(item_id="", item={}, **kwargs):
    dat = load_dat("pds", **kwargs)
    if dat:
        return find_item(dat_type="pds", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def find_pdf(item_id="", item={}, **kwargs):
    dat = load_dat("pdf", **kwargs)
    if dat:
        return find_item(dat_type="pdf", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def find_pas(item_id="", item={}, **kwargs):
    dat = load_dat("pas", **kwargs)
    if dat:
        return find_item(dat_type="pas", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def find_paf(item_id="", item={}, **kwargs):
    dat = load_dat("paf", **kwargs)
    if dat:
        return find_item(dat_type="paf", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def find_pts(item_id="", item={}, **kwargs):
    dat = load_dat("pts", **kwargs)
    if dat:
        return find_item(dat_type="pts", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def find_ptf(item_id="", item={}, **kwargs):
    dat = load_dat("ptf", **kwargs)
    if dat:
        return find_item(dat_type="ptf", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def find_pdd(item_id="", item={}, **kwargs):
    dat = load_dat("pdd", **kwargs)
    if dat:
        return find_item(dat_type="pdd", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def find_pad(item_id="", item={}, **kwargs):
    dat = load_dat("pad", **kwargs)
    if dat:
        return find_item(dat_type="pad", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def find_cgs(item_id="", item={}, **kwargs):
    dat = load_dat("cgs", **kwargs)
    if dat:
        return find_item(dat_type="cgs", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def find_cgf(item_id="", item={}, **kwargs):
    dat = load_dat("cgf", **kwargs)
    if dat:
        return find_item(dat_type="cgf", generic_set=dat, item_id=item_id, item=item)
    else:
        return ""


def get_item_from_base(dat_type, item_id="", item={}, **kwargs):
    '''
    Retorna um dicion rio com os campos do item da base passado como parƒmetro em item_id ou item. Caso
    o argumento base_item seja passado com uma base completa, o item ‚ procurado no objeto, do contr rio o item
    ‚ carregado do arquivo dat do tipo dat_type
    :param dat_type:
    :param item_id:
    :param item:
    :param kwargs:
    :return:
    '''
    dat_type = dat_type.lower()
    where = kwargs.get("where",{})
    op = kwargs.get("op","and")
    base_item = kwargs.get('base_item',[])
    if base_item != []:
        dat = base_item.get(dat_type,[])
    else:
        dat = load_dat(dat_type, **kwargs)
    if dat:
        return get_item(dat_type=dat_type, generic_set=dat, item_id=item_id, item=item, where=where, op=op)
    else:
        return "", {}



def get_dataset_from_base(dat_type, id_set=[], item_set = [], **kwargs):
    dat_type = dat_type.lower()
    #where = kwargs.get("where",{})
    #op = kwargs.get("op","and")
    dat = load_dat(dat_type, **kwargs)
    if dat:
        return get_dataset(dat_type=dat_type, generic_set=dat, id_set=id_set, item_set=item_set,
                           **kwargs)
    else:
        return []


def add_item_to_base(dat_type, item, **kwargs):
    '''
    Adiciona o item ao dat, caso n„o exista um com o mesmo id.
    :param dat_type:
    :param item:
    :param kwargs: add_to: str com o diret¢rio a ser adicionado. Caso n„o exista, ‚ criado. source_path: str com caminho
    da base a ser usada. base: str com o nome da base a ser usada.
    :return:
    '''
    dat_type = dat_type.lower()
    dat = load_dat(dat_type, **kwargs)
    dat_pk = DAT_KEYS.get(dat_type)
    if dat_pk == "":
        kwargs.setdefault("ignore_id", True)
        search_path = find_item(dat_type=dat_type, generic_set=dat, item=item).lstrip("#")
    elif dat_pk == "ID":
        item_pk = item.get(dat_pk)
        if item_pk is None:
            print_msg(__name__, "o item n„o possui chave prim ria", **kwargs)
            return False
        search_path = find_item(dat_type=dat_type, generic_set=dat, item_id=item_pk).lstrip("#")
    add_to = kwargs.get("add_to","")

    if (search_path != "") and (not kwargs.get("add_or_update")==True):
        print_msg(__name__, "item j  existente em {0}".format(search_path), **kwargs)
        return False
    if add_to:
        #print(add_to)
        location = make_include_str(dat_type, **kwargs)
        print_msg(__name__, "o {0} ser  salvo em {1}".format(dat_type,location), MSG_WARN, **kwargs)
        add_item(dat_type=dat_type, generic_set=dat, item=item, to_include=location, **kwargs)
    else:
        location = dat_type
        print_msg(__name__, "o item ser  salvo no {0} raiz".format(dat_type), MSG_WARN, **kwargs)
        add_item(dat_type=dat_type, generic_set=dat[dat_type], item=item, **kwargs)
    write_dat(dat_type, dat, dests=[location], **kwargs)


def add_items_to_base(dat_type, items, **kwargs):
    '''
    Adiciona o items ao dat, caso n„o existam com o mesmo id.
    :param dat_type:
    :param items: lista com objetos dict representando cada item
    :param kwargs: add_to: str com o diret¢rio a ser adicionado. Caso n„o exista, ‚ criado. source_path: str com caminho
    da base a ser usada. base: str com o nome da base a ser usada.
    :return:
    '''
    dat_type = dat_type.lower()
    dat = load_dat(dat_type, **kwargs)

    #locations = find_items(dat_type=dat_type, generic_set=dat, items=items)
    locations = []
    add_to = kwargs.get("add_to","")
    for item in items:
        print_msg(__name__, "adicionando: {0}".format(item), MSG_INFO, **kwargs)
        search = find_item(dat_type=dat_type, generic_set=dat, item=item)
        if (search):
            print_msg(__name__, "item {0} j  existente em {1}".format(item, search), **kwargs)
        elif add_to:
            #print(add_to)
            location = make_include_str(dat_type, **kwargs)
            print_msg(__name__, "o {0} ser  salvo em {1}".format(dat_type,location), MSG_WARN, **kwargs)
            add_item(dat_type=dat_type, generic_set=dat, item=item, to_include=location, **kwargs)
            if not location in locations:
                locations.append(location)
        else:
            location = dat_type
            print_msg(__name__, "o item ser  salvo no {0} raiz".format(dat_type), MSG_WARN, **kwargs)
            add_item(dat_type=dat_type, generic_set=dat[dat_type], item=item, **kwargs)
            if not location in locations:
                locations.append(location)
    if locations:
        write_dat(dat_type, dat, dests=locations, **kwargs)


def update_items_in_base(dat_type, fields, item_ids=[], items=[], **kwargs):
    where = kwargs.get("where", {})
    op = kwargs.get("op","and")
    dat_type = dat_type.lower()
    dat = load_dat(dat_type, **kwargs)
    update_or_add = kwargs.get("update_or_add", False)
    locations = []
    if item_ids:
        iter_list = copy.deepcopy(item_ids)
    elif items:
        iter_list = copy.deepcopy(items)
    else:
        print_msg(__name__, "parƒmetros insuficientes. Passe uma lista de IDs ou itens a serem atualizados", **kwargs)
        return False
    if len(fields) != len(iter_list):
        print_msg(__name__, "lista com campos e lista de itens/ids possuem tamanhos diferentes", **kwargs)
        return False
    i = 0
    for item in iter_list:
        if item_ids:
            search = find_item(dat_type=dat_type, generic_set=dat, item_id=item)
        else:
            search = find_item(dat_type=dat_type, generic_set=dat, item=item)
        if (not search) and (update_or_add):
            # item n„o existe
            if item_ids:
                new_item = copy.deepcopy(fields[i])
                new_item[DAT_KEYS.get(dat_type)]=item
            else:
                new_item = copy.deepcopy(item)
                new_item.update(fields[i])
            print_msg(__name__,
                      "item novo encontrado. Op‡„o de adicionar selecionada, adicionando {0}...".format(new_item),
                      MSG_INFO, **kwargs)
            add_item(dat_type=dat_type, generic_set=dat, item=new_item)
            locations.append(dat_type)
        elif (search):
            print_msg(__name__, "atualizando {0} em {1}".format(item, search), MSG_INFO, **kwargs)
            if item_ids:
                update_item(dat_type=dat_type,generic_set=dat,fields=fields[i],item_id=item, where=where, op=op)
            else:
                update_item(dat_type=dat_type,generic_set=dat,fields=fields[i],item=item, where=where, op=op)
            if search not in locations:
                locations.append(search)
        i +=1
    if locations:
        write_dat(dat_type=dat_type,dat_content=dat,dests=locations,**kwargs)





def remove_items_from_base(dat_type, items=[], item_ids=[], **kwargs):
    dat_type = dat_type.lower()
    dat = load_dat(dat_type, **kwargs)
    dat_pk = DAT_KEYS.get(dat_type)
    where = kwargs.get("where", {})
    op = kwargs.get("op", "and")
    locations = find_items(dat_type=dat_type, generic_set=dat, items=items, item_ids=item_ids, where=where, op=op)
    if locations:
        print_msg(__name__, "removendo de {0}...".format(locations), MSG_INFO, **kwargs)
        delete_dataset(dat_type=dat_type, generic_set=dat, item_ids=item_ids, items=items, where=where, op=op)
        write_dat(dat_type=dat_type, dat_content=dat, dests=locations, **kwargs)
        return True
    else:
        print_msg(__name__, "ponto(s) n„o encontrado(s)", MSG_INFO, **kwargs)
        return False




def get_pds(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="pds", item_id=item_id, item=item, **kwargs)


def get_pas(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="pas", item_id=item_id, item=item, **kwargs)


def get_pts(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="pts", item_id=item_id, item=item, **kwargs)


def get_cgs(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="cgs", item_id=item_id, item=item, **kwargs)


def get_pdf(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="pdf", item_id=item_id, item=item, **kwargs)


def get_paf(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="paf", item_id=item_id, item=item, **kwargs)


def get_ptf(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="ptf", item_id=item_id, item=item, **kwargs)


def get_cgf(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="cgf", item_id=item_id, item=item, **kwargs)


def get_pdd(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="pdd", item_id=item_id, item=item, **kwargs)


def get_pad(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="pad", item_id=item_id, item=item, **kwargs)


def get_ptd(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="ptd", item_id=item_id, item=item, **kwargs)


def get_tac(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="tac", item_id=item_id, item=item, **kwargs)

def get_tdd(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="tdd", item_id=item_id, item=item, **kwargs)


def get_lsc(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="lsc", item_id=item_id, item=item, **kwargs)


def get_cnf(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="cnf", item_id=item_id, item=item, **kwargs)


def get_nv2(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="nv2", item_id=item_id, item=item, **kwargs)


def get_nv1(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="nv1", item_id=item_id, item=item, **kwargs)


def get_utr(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="utr", item_id=item_id, item=item, **kwargs)


def get_cxu(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="cxu", item_id=item_id, item=item, **kwargs)


def get_enm(item_id="", item={}, **kwargs):
    return get_item_from_base(dat_type="enm", item_id=item_id, item=item, **kwargs)


def get_physical_conf(dat_type, item_id="", item={}, **kwargs):
    '''
    Retorna um dicion rio com a seguinte estrutura, referente ao ponto f¡sico passado como parƒmetro:
    {
    [pdf|paf|ptf|cgf]:{"location": <local>, "item": <item(dict)>},
     "nv2": {"location": <local>, "item": <item(dict)>},
     "nv1": {"location": <local>, "item": <item(dict)>},
     "cnf": {"location": <local>, "item": <item(dict)>},
     "lsc": {"location": <local>, "item": <item(dict)>}
     }

     Onde local ‚ uma string com o local do include onde o ponto existe (ou simplesmente o nome do .dat caso
     esteja na raiz) e item ‚ um dicion rio com os campos e valores do ponto

    :param dat_type:
    :param item_id:
    :param item:
    :param kwargs:
    :return:
    '''
    dat_type = dat_type.lower()
    if not item_id:
        item_id = item.get("ID")
    f_location, f_item = get_item_from_base(dat_type=dat_type, item_id=item_id, **kwargs)
    nv2_location, nv2_item = get_nv2(item_id=f_item.get("NV2"),
                                     where={"TPPNT": "== "+dat_type.upper()},
                                     **kwargs)
    nv1_location, nv1_item = get_nv1(item_id=nv2_item.get("NV1"), **kwargs)

    cnf_location, cnf_item = get_cnf(item_id=nv1_item.get("CNF"), **kwargs)

    lsc_location, lsc_item = get_lsc(item_id=cnf_item.get("LSC"), **kwargs)

    output = {
        dat_type: {
            "location": f_location,
            "item": f_item
        },
        "nv2": {
            "location": nv2_location,
            "item": nv2_item
        },
        "nv1": {
            "location": nv1_location,
            "item": nv1_item
        },
        "cnf": {
            "location": cnf_location,
            "item": cnf_item
        },
        "lsc": {
            "location": lsc_location,
            "item": lsc_item
        }
    }

    return output


def get_aconf_from_base(dat_type, item_id="", item={}, **kwargs):
    '''
    Retorna um dicion rio com o ponto/controle l¢gico passado como parƒmetro e as configura‡”es relacionadas de tac,
    cnf, lsc, ponto/controle f¡sico, nv1, nv2, rca e rfi. O dicion rio possui o seguinte formato:
    {
    str [pds|pas|pts|cgs]: {
            "item": {str campo1: str valor1, str campo2: str valor2, ...},
            "location": str include_path
            },
    str [pdf/paf/ptf/cgf]: {
            "items": list [{str campo1: str valor1, str campo2: str valor2,..},
                            {str campo1: str valor1, str campo2: str valor2,..}],
            "locations": list [str include_path1, str include_path2, ...]
            },
    "rca":  {
            "items": list [{str campo1: str valor1, str campo2: str valor2,..},
                            {str campo1: str valor1, str campo2: str valor2,..}],
            "locations": list [str include_path1, str include_path2, ...],
            "parcs_ids": list [str id1, str id2, ...]
            },
    "rfi":  {
            "items": list [{str campo1: str valor1, str campo2: str valor2,..},
                            {str campo1: str valor1, str campo2: str valor2,..}],
            "locations": list [str include_path1, str include_path2, ...],
            "parcs_ids": list [str id1, str id2, ...]
            },
    "tac": {
            "item": {str campo1: str valor1, str campo2: str valor2, ...},
            "location": str include_path
            },
    "cnf": {
            "item": {str campo1: str valor1, str campo2: str valor2, ...},
            "location": str include_path
            },
    "lsc": {
            "item": {str campo1: str valor1, str campo2: str valor2, ...},
            "location": str include_path
            },
    "nv1": {
            "items": list [{str campo1: str valor1, str campo2: str valor2,..},
                            {str campo1: str valor1, str campo2: str valor2,..}],
            "locations": list [str include_path1, str include_path2, ...]
            },
    "nv2": {
            "items": list [{str campo1: str valor1, str campo2: str valor2,..},
                            {str campo1: str valor1, str campo2: str valor2,..}],
            "locations": list [str include_path1, str include_path2, ...]
            },

    }
    :param dat_type:
    :param item_id:
    :param item:
    :param kwargs:
    :return:
    '''

    dat_type = dat_type.lower()
    dtf = dat_type.replace("s", "f")
    base_item = kwargs.get('base_item',{})

    output = {}
    xxf_items = []
    xxf_locations = []
    xxf_conf = {}

    rca_items = []
    rca_locations = []
    rca_parcs_ids = []

    rfc_items = []
    rfc_locations = []
    rfc_parcs_ids = []

    rfi_items = []
    rfi_locations = []

    xxf_rfi_items = []
    xxf_rfi_locations = []

    nv1_items = []
    nv1_locations = []

    nv2_items = []
    nv2_locations = []

    s_location, s_item = get_item_from_base(dat_type=dat_type, item_id=item_id, item=item, **kwargs)
    if not s_item:
        print_msg(__name__, "o item informado n„o existe na base", **kwargs)
        return {}

    #print(s_item)
    output[dat_type] = {"location": s_location, "item": s_item}

    # lˆ a tac do item
    s_tac_id = s_item.get("TAC")
    if s_tac_id is None:
        print_msg(__name__, "o item n„o possui TAC configurada", **kwargs)
        return {}
    tac_location, tac_item = get_tac(item_id=s_tac_id, **kwargs)

    #print(tac_item)
    output["tac"] = {"location": tac_location, "item": tac_item}

    # dependendo da tac, define o que ser  lido de configura‡„o

    tpaqs = TPAQS[tac_item.get("TPAQS", "ASAC")]
    #s_type = tpaqs[I_DESCR]

    # lˆ a lsc do item
    s_lsc_id = tac_item.get("LSC")
    if s_lsc_id is None:
        print_msg(__name__, "o item n„o tem LSC configurada", **kwargs)
        print(item_id)
        print(tac_item.get('ID'))
        return {}
    lsc_location, lsc_item = get_lsc(item_id=s_lsc_id, **kwargs)

    #print(lsc_item)
    output["lsc"] = {"location": lsc_location, "item": lsc_item}

    # lˆ a cnf

    cnf_location, cnf_item = get_cnf(where={"LSC": "== "+s_lsc_id}, **kwargs)

    output["cnf"] = {"location": cnf_location, "item": cnf_item}

    s_tcv = TCV.get(lsc_item.get("TCV"))
    s_ttp = TTP.get(lsc_item.get("TTP"))

    if (s_tcv is None) or (s_ttp is None):
        print_msg(__name__, "a LSC do item n„o possui TCV ou TTP corretamente configurado", **kwargs)
        return {}

    #print(s_tcv)
    #print(s_ttp)

    # caso seja um ponto de aquisi‡„o e controle, pegar pdf/paf/ptf
    if tpaqs == TPAQS_ASAC:
        if dat_type == "cgs":
            where = {"CGS": "== "+s_item.get("ID")}
        else:
            where = {"PNT": "== "+s_item.get("ID"),
                    "TPPNT": "== "+dat_type.upper()}
        #print('dtf: ' + str(dtf))
        xxf_location, xxf_item = get_item_from_base(dat_type=dtf, where=where, **kwargs)
        xxf_conf = get_physical_conf(dtf, item=xxf_item, **kwargs)
        #xxf_item = xxf_conf.get(dtf).get('item')
        #xxf_location = xxf_conf.get(dtf).get('location')
        #print('xxf item: '+ str(xxf_item) + '\n' + 'xxf_conf: '+ str(xxf_conf))
        xxf_locations.append(xxf_location)
        xxf_items.append(xxf_item)
        nv1_items.append(xxf_conf["nv1"]["item"])
        nv1_locations.append(xxf_conf["nv1"]["location"])
        nv2_items.append(xxf_conf["nv2"]["item"])
        nv2_locations.append(xxf_conf["nv2"]["location"])

        output[dtf] = {"locations": xxf_locations, "items": xxf_items}
        output["nv1"] = {"locations": nv1_locations, "items": nv1_items}
        output["nv2"] = {"locations": nv2_locations, "items": nv2_items}

    # caso seja um c lculo, pegar as parcelas de rca
    elif tpaqs == TPAQS_ACSC:
        if base_item != {}:
            rca_dat = base_item.get('rca',[])
        else:
            rca_dat = load_rca(**kwargs)
        rca_items = get_dataset(dat_type="rca", generic_set=rca_dat,
                                where={"PNT": "== "+s_item.get("ID"), "TPPNT": "== "+dat_type.upper()}, **kwargs)
        if rca_items:
            for rca_item in rca_items:
                rca_locations.append(find_item(dat_type="rca", generic_set=rca_dat, item=rca_item))
                rca_parcs_ids.append(rca_item["PARC"])
            #rca_locations = list(set(rca_locations))

        output["rca"] = {
            "locations": rca_locations,
            "items": rca_items,
            "parcs_ids": rca_parcs_ids
        }

    # caso seja um filtro composto, pegar as parcelas
    elif tpaqs in [TPAQS_AFID, TPAQS_AFIL]:
        if base_item != {}:
            rfc_dat = base_item['rfc']
        else:
            rfc_dat = load_rfc(**kwargs)
        rfc_items = get_dataset(dat_type="rfc", generic_set=rfc_dat,
                                where={"PNT": "== "+s_item.get("ID"), "TPPNT": "== "+dat_type.upper()}, **kwargs)
        if rfc_items:
            for rfc_item in rfc_items:
                rfc_locations.append(find_item("rfc", rfc_dat, item=rfc_item))
                related_dat = str(rfc_item.get("TPPARC")).lower()
                related_id = rfc_item.get("PARC")
                xxf_related_location, xxf_related_item = get_item_from_base(dat_type=related_dat,
                                                                            item_id=related_id, **kwargs)
                rfc_parcs_ids.append(xxf_related_item.get("PNT",''))

            #rfc_locations = list(set(rfc_locations))
            output["rfc"] = {
                "locations": rfc_locations,
                "items": rfc_items,
                "parcs_ids": rfc_parcs_ids
            }
        else:

            # filtros simples
            if base_item != {}:
                rfi_dat = base_item['rfi']
                xxf_dat = base_item[dtf]
            else:
                rfi_dat = load_rfi(**kwargs)
                xxf_dat = load_dat(dat_type=dtf, **kwargs)
            xxf_rfi_items = get_dataset(dat_type=dtf, generic_set=xxf_dat,
                                    where={"PNT": "== "+s_item.get("ID"), "TPPNT": "== "+dat_type.upper()}, **kwargs)
            for xxf_rfi_item in xxf_rfi_items:
                rfi_location, rfi_item = get_item(dat_type="rfi", generic_set=rfi_dat,
                                                  where={"PNT": "== "+xxf_rfi_item.get("ID"), "TIPOP": "== "+dtf.upper()})
                rfi_locations.append(rfi_location)
                rfi_items.append(rfi_item)

                xxf_rfi_locations.append(find_item(dat_type=dtf, generic_set=xxf_dat, item=xxf_rfi_item))
                xxf_conf = get_physical_conf(dtf, item=xxf_rfi_item, **kwargs)
                nv1_items.append(xxf_conf["nv1"]["item"])
                nv1_locations.append(xxf_conf["nv1"]["location"])
                nv2_items.append(xxf_conf["nv2"]["item"])
                nv2_locations.append(xxf_conf["nv2"]["location"])

            #xxf_rfi_locations = list(set(xxf_rfi_locations))
            #rfi_locations = list(set(rfi_locations))

            output["rfi"] = {"items": rfi_items, "locations": rfi_locations}
            output[dtf] = {"items": xxf_rfi_items, "locations": xxf_rfi_locations}
            output["nv1"] = {"items": nv1_items, "locations": nv1_locations}
            output["nv2"] = {"items": nv2_items, "locations": nv2_locations}

    return output


def load_base(**kwargs):
    base = {}
    base['cgf'] = load_cgf(**kwargs)
    base['cgs'] = load_cgs(**kwargs)

    base['pds'] = load_pds(**kwargs)
    base['pdf'] = load_pdf(**kwargs)
    base['pdd'] = load_pdd(**kwargs)

    base['pts'] = load_pts(**kwargs)
    base['ptf'] = load_ptf(**kwargs)
    base['ptd'] = load_ptd(**kwargs)

    base['pas'] = load_pas(**kwargs)
    base['paf'] = load_paf(**kwargs)
    base['pad'] = load_pad(**kwargs)

    base['rca'] = load_rca(**kwargs)
    base['rfc'] = load_rfc(**kwargs)
    base['rfi'] = load_rfi(**kwargs)

    source_path = kwargs.get('source_path')
    # Ajusta o caminho da base para a pasta raiz dados, caso tenha sido selecionado um inlcude
    # Esse procedimento ‚ feito apenas para carregar os arquivos de configura‡„o abaixo
    if source_path != None:
        if source_path.rsplit('\\', 3)[2] == 'dados':
            source_path = source_path.rsplit('\\', 1)[0]
            kwargs['source_path'] = source_path
        elif source_path.rsplit('\\', 3)[2] == 'bd':
            pass

    base['tac'] = load_tac(**kwargs)
    base['ocr'] = load_ocr(**kwargs)

    base['cnf'] = load_cnf(**kwargs)
    base['gsd'] = load_gsd(**kwargs)
    base['noh'] = load_noh(**kwargs)
    base['lsc'] = load_lsc(**kwargs)
    base['nv1'] = load_nv1(**kwargs)
    base['nv2'] = load_nv2(**kwargs)

    base['pro'] = load_pro(**kwargs)
    base['enu'] = load_dat('enu', **kwargs)
    base['cxu'] = load_cxu(**kwargs)
    base['utr'] = load_utr(**kwargs)
    base['mul'] = load_mul(**kwargs)
    base['enm'] = load_enm(**kwargs)
    base['inp'] = load_inp(**kwargs)
    base['ins'] = load_ins(**kwargs)
    base['tdd'] = load_tdd(**kwargs)
    base['map'] = load_map(**kwargs)
    base['grupo'] = load_grupo(**kwargs)
    base['grcmp'] = load_grcmp(**kwargs)
    base['tcl'] = load_tcl(**kwargs)
    base['tctl'] = load_tctl(**kwargs)
    base['ttp'] = load_ttp(**kwargs)
    base['tcv'] = load_tcv(**kwargs)
    base['sxp'] = load_sxp(**kwargs)
    base['sev'] = load_sev(**kwargs)
    base['inm'] = load_inm(**kwargs)
    base['psv'] = load_psv(**kwargs)
    base['tn1'] = load_tn1(**kwargs)
    base['tn2'] = load_tn2(**kwargs)
    base['e2m'] = load_e2m(**kwargs)
    base['ctx'] = load_ctx(**kwargs)
    base['cxp'] = load_cxp(**kwargs)



    return base

def save_base(base, **kwargs):
    for key in list(base.keys()):
        dat_content = base[key]
        if dat_content != {}:
            write_dat(key, dat_content, **kwargs)


def get_ttp_conf(dat_type, item_id="", item={}, **kwargs):
    '''
    Retorna uma tuple com o tipo TTP do ponto l¢gico passado.
    Ex.: TTP_IEC3S = (30, "IEC3S", "Transportador em Frames FT3-DNP do IEC/60870 para Terminal Server", "iec3s")

    :param dat_type:
    :param item_id:
    :param item:
    :param kwargs:
    :return:
    '''
    conf = get_aconf_from_base(dat_type, item_id=item_id, item=item, **kwargs)
    return TTP[conf["lsc"]["item"]["TTP"]]


def get_tcv_conf(dat_type, item_id="", item={}, **kwargs):
    '''
    Retorna uma tuple com o tipo TCV do ponto l¢gico passado.
    Ex.: (8, "CNVH", "Conversor DNP 3.0", "dnp3")
    :param dat_type:
    :param item_id:
    :param item:
    :param kwargs:
    :return:
    '''
    aconf = kwargs.get('aconf')
    if aconf:
        return TCV[aconf["lsc"]["item"]["TCV"]]
    else:
        conf = get_aconf_from_base(dat_type, item_id=item_id, item=item, **kwargs)
        return TCV[conf["lsc"]["item"]["TCV"]]

def is_61850(dat_type, item_id="", item={}, **kwargs):
    '''
    Retorna True caso o item l¢gico passado como parƒmetro seja proveniente de uma liga‡„o 61850
    :param dat_type:
    :param item_id:
    :param item:
    :param kwargs:
    :return:
    '''
    return get_tcv_conf(dat_type, item_id=item_id, item=item, **kwargs) == TCV_I61850

def is_dnp3(dat_type, item_id="", item={}, **kwargs):
    '''
    Retorna True caso o item l¢gico passado como parƒmetro seja proveniente de uma liga‡„o dnp3
    :param dat_type:
    :param item_id:
    :param item:
    :param kwargs:
    :return:
    '''
    return get_tcv_conf(dat_type, item_id=item_id, item=item, **kwargs) == TCV_DNP3



def get_logical_dist(dat_type, item_id="", item={}, **kwargs):
    '''
    Retorna um dicion rio com configura‡”es de distribui‡„o do ponto l¢gico passado como parƒmetro. Formato da sa¡da:
    {
    [pdd|pad|ptd] : {
                    "item": dict item,
                    "location": str include_location
                    },
    [pdf|paf|ptf] : {
                    "item": dict item,
                    "location": str include_location
                    },
    "tdd" :         {
                    "item": dict item,
                    "location": str include_location
                    },
    "nv1" :         {
                    "item": dict item,
                    "location": str include_location
                    },
    "nv2" :         {
                    "item": dict item,
                    "location": str include_location
                    },
    "lsc" :         {
                    "item": dict item,
                    "location": str include_location
                    },
    "cnf" :         {
                    "item": dict item,
                    "location": str include_location
                    },
    }
    :param dat_type:
    :param item_id:
    :param item:
    :param kwargs:
    :return:
    '''
    output = []
    dat_type = dat_type.lower()
    dtf = dat_type.replace("s", "f")
    dtd = dat_type.replace("s", "d")
    where = {
        dat_type.upper(): "== "+item_id
        }
    # um ponto pode ter v rias distribui‡”es configuradas

    xxd_items = get_dataset_from_base(dat_type=dtd, where=where, **kwargs)
    print('xxd_items: ' + str(xxd_items))
    for xxd_item in xxd_items:
        xxd_location = find_item_in_base(dat_type, item=xxd_item, **kwargs)
        tdd_location, tdd_item = get_tdd(item_id=item_id, **kwargs)
        xxf_location, xxf_item = get_item_from_base(dtf,
                                                    where={
                                                      "PNT": "== "+xxd_item.get("ID"),
                                                      "TPPNT": "== "+dtd.upper()
                                                  }, **kwargs)

        item = {
            dtd: {
                "item": xxd_item,
                "location": xxd_location
            },
            dtf: {
                "item": xxf_item,
                "location": xxf_location
            },
            "tdd": {
                "item": tdd_item,
                "location": tdd_location
            }
        }

        phy_conf = get_physical_conf(dat_type=dtf, item=xxf_item, **kwargs)
        item.update(phy_conf)

        output.append(item)
    return output

def get_endN3_dist(dat_type, item_id="", item={}, **kwargs):
    '''
     Retorna a ordem de distribui‡„o de N3
     :param dat_type:
     :param item_id:
     :param item:
     :param kwargs:
     :return:
     '''
    output = ''
    dat = kwargs.get('base')
    dat_type = dat_type.lower()
    dtf = dat_type.replace("s", "f")
    dtd = dat_type.replace("s", "d")
    if dat_type == 'cgs':
        dtd = 'pdd'
    where = {
        dat_type.upper(): "== " + item_id
    }
    # um ponto pode ter v rias distribui‡”es configuradas
    xxd_items = get_dataset(dat_type=dtd.lower(), generic_set=dat, id_set=[], item_set=[],
                            **kwargs)
    xxd_item = xxd_items[len(xxd_items)-1]
    xxd_location, xxd_dic = get_item_from_base(dat_type=dtd, where={"PDS": "== " + item_id}, base_item = dat, **kwargs)
    if xxd_dic != {}:
        #tdd_location, tdd_item = get_tdd(item_id = xxd_dic.get('TDD'), base_item = dat, **kwargs)
        xxf_location, xxf_item = get_item_from_base(dtf,
                                                    where={
                                                        "PNT": "== " + xxd_dic.get("ID"),
                                                        "TPPNT": "== " + dtd.upper()
                                                    },base_item = dat, **kwargs)
        output = xxf_item.get('ORDEM')
    return output

def get_control_dist(item_id="", item={}, **kwargs):
    '''
    Retorna um dicion rio com as configura‡”es de distribui‡„o do controle cgs passado como parƒmetro, da seguinte
    forma:
    {
    cgf_d :         {
                    "item": dict item,
                    "location": str include_location
                    },
    "nv1_d" :       {
                    "item": dict item,
                    "location": str include_location
                    },
    "nv2_d" :       {
                    "item": dict item,
                    "location": str include_location
                    },
    "lsc_d" :       {
                    "item": dict item,
                    "location": str include_location
                    },
    "cnf_d" :       {
                    "item": dict item,
                    "location": str include_location
                    },
    }

    Comandos 61850 ainda possuem as entradas cgf_r e cgs_r com as configura‡”es do comando roteado

    :param item_id: string com o cgs
    :param item: dict com o cgs
    :param kwargs: source_path = string com o caminho da base
    :return:
    '''
    output = []

    if not item_id:
        item_id = item.get("ID", "")

    # um ponto pode ter v rias distribui‡”es configuradas
    if not is_61850("cgs", item_id=item_id, item=item, **kwargs):
        where = {
            "KCONV": "== "+item_id
        }
        cgf_items = get_dataset_from_base(dat_type="cgf", where=where, **kwargs)
        for cgf_item in cgf_items:
            cgf_location = find_item_in_base("cgf", item=cgf_item, **kwargs)
            #cnf_location, cnf_item = get_cnf(item_id=cgf_item.get("CNF"))

            item = {
                "cgf_d": {
                    "item": cgf_item,
                    "location": cgf_location
                }
            }

            phy_conf = get_physical_conf(dat_type="cgf", item=cgf_item, **kwargs)
            #item.update(phy_conf)
            item["cnf_d"] = phy_conf.get("cnf", {}).copy()
            item["lsc_d"] = phy_conf.get("lsc", {}).copy()
            item["nv1_d"] = phy_conf.get("nv1", {}).copy()
            item["nv2_d"] = phy_conf.get("nv2", {}).copy()

            output.append(item)
    else:
        where = {
            "CGS": "== "+ item_id
        }
        cgf_location, cgf_item = get_cgf(where=where, **kwargs)
        cgf_id = cgf_item.get("ID")
        # encontra o cgf roteado 61850, que possui id com sufixo e cgs com ponto roteado
        routed_location, routed_item = get_cgf(where={
                                                        "ID": "has "+cgf_id,
                                                        "CGS": "!= "+item_id
                                                        }, **kwargs)
        cgs_location, cgs_item = get_cgs(item_id=routed_item.get("CGS",""), **kwargs)

        dist_location, dist_item = get_cgf(where={
            "KCONV": "has "+cgs_item.get("ID","")
        }, **kwargs)

        item = {
                "cgf_r": {
                    "item": routed_item,
                    "location": routed_location
                },
                "cgs_r": {
                        "item": cgs_item,
                        "location": cgs_location
                    },
                "cgf_d": {
                        "item": dist_item,
                        "location": dist_location
                    },
            }

        phy_conf = get_physical_conf(dat_type="cgf", item=dist_item, **kwargs)
        #item.update(phy_conf)
        item["cnf_d"] = phy_conf.get("cnf", {}).copy()
        item["lsc_d"] = phy_conf.get("lsc", {}).copy()
        item["nv1_d"] = phy_conf.get("nv1", {}).copy()
        item["nv2_d"] = phy_conf.get("nv2", {}).copy()
        output.append(item)
    return output


def print_logical_conf(lconf):
    def print_block(dt):
        if lconf.get(dt) is not None:
            field_order = DAT_FIELDS[dt]
            print("Configura‡„o de "+dt.upper())
            print("Locais: "+str(lconf[dt].get("location",""))+str(lconf[dt].get("locations","")))
            field_order = DAT_FIELDS[dt]
            dataset = lconf[dt].get("item")
            if dataset is None:
                dataset = lconf[dt].get("items")
            else:
                dataset = [dataset]
            text, count = make_dat_str(dat_type=dt, dataset=dataset, field_order=field_order)
            print(text)

    if "pds" in lconf.keys():
        dts = "pds"
        dtf = "pdf"
    elif "pas" in lconf.keys():
        dts = "pas"
        dtf = "paf"
    elif "pts" in lconf.keys():
        dts = "pts"
        dtf = "ptf"
    elif "cgs" in lconf.keys():
        dts = "cgs"
        dtf = "cgf"
    print_block(dts)
    print_block("tac")
    print_block("lsc")
    print_block("cnf")
    print_block("rca")
    print_block("rfi")
    print_block(dtf)
    print_block("nv2")
    print_block("nv1")


def change_tags(old, new, fields=[], **kwargs):
    '''
    Muda tags passadas na lista old para as tags passadas na lista new. As tags s„o mudadas apenas nos campos fields,
    se este for passado, caso contr rio em todos os campos. As tags s„o mudadas apenas nas seguintes tabelas: pds,
    pdf, pdd, pas, paf, pad, pts, ptf, ptd, cgs, cgf, rca, rfi, rfc
    :param old: lista com as tags antigas
    :param new: lista com as tags novas
    :param fields: campos onde tags devem ser subsitu¡das
    :param kwargs: source_path= str com o caminho pra base; base=nome da base (sage apenas)
    :return:
    '''
    pds = load_dat("pds", **kwargs)
    pdf = load_dat("pdf", **kwargs)
    pdd = load_dat("pdd", **kwargs)

    pas = load_dat("pas", **kwargs)
    paf = load_dat("paf", **kwargs)
    pad = load_dat("pad", **kwargs)

    pts = load_dat("pts", **kwargs)
    ptf = load_dat("ptf", **kwargs)
    ptd = load_dat("ptd", **kwargs)

    cgs = load_dat("cgs", **kwargs)
    cgf = load_dat("cgf", **kwargs)

    rca = load_dat("rca", **kwargs)
    rfi = load_dat("rfi", **kwargs)
    rfc = load_dat("rfc", **kwargs)

    if kwargs.get("change_devices") == True:
        nv1 = load_nv1(**kwargs)
        nv2 = load_nv2(**kwargs)
        tac = load_tac(**kwargs)
        cnf = load_cnf(**kwargs)
        lsc = load_lsc(**kwargs)
        tdd = load_dat("tdd", **kwargs)
        mul = load_dat("mul", **kwargs)
        enm = load_dat("enm", **kwargs)
        cxu = load_dat("cxu", **kwargs)
        utr = load_dat("utr", **kwargs)
        map = load_dat("map", **kwargs)

    if type(old) != list:
        old = [old]
    if type(new) != list:
        new = [new]

    if len(old) != len(new):
        print_msg(__name__, "argumentos old e new devem ter mesmo tamanho", **kwargs)
        return 1

    k = 0
    for o in old:

        replace_text("pds", pds, o, new[k], fields=fields)
        replace_text("pdf", pdf, o, new[k], fields=fields)
        replace_text("pdd", pdd, o, new[k], fields=fields)

        replace_text("pas", pas, o, new[k], fields=fields)
        replace_text("paf", paf, o, new[k], fields=fields)
        replace_text("pad", pad, o, new[k], fields=fields)

        replace_text("pts", pts, o, new[k], fields=fields)
        replace_text("ptf", ptf, o, new[k], fields=fields)
        replace_text("ptd", ptd, o, new[k], fields=fields)

        replace_text("cgs", cgs, o, new[k], fields=fields)
        replace_text("cgf", cgf, o, new[k], fields=fields)
        replace_text("rca", rca, o, new[k], fields=fields)
        replace_text("rfi", rfi, o, new[k], fields=fields)
        replace_text("rfc", rfc, o, new[k], fields=fields)

        if kwargs.get("change_devices") == True:
            replace_text("cnf", cnf, o, new[k], fields=fields)
            replace_text("nv1", nv1, o, new[k], fields=fields)
            replace_text("nv2", nv2, o, new[k], fields=fields)
            replace_text("lsc", lsc, o, new[k], fields=fields)
            replace_text("mul", mul, o, new[k], fields=fields)
            replace_text("enm", enm, o, new[k], fields=fields)
            replace_text("tac", tac, o, new[k], fields=fields)
            replace_text("cxu", cxu, o, new[k], fields=fields)
            replace_text("tdd", tdd, o, new[k], fields=fields)
            replace_text("utr", utr, o, new[k], fields=fields)
            replace_text("map", map, o, new[k], fields=fields)


        k += 1


    write_dat("pds", pds, **kwargs)
    write_dat("pdf", pdf, **kwargs)
    write_dat("pdd", pdd, **kwargs)

    write_dat("pas", pas, **kwargs)
    write_dat("paf", paf, **kwargs)
    write_dat("pad", pad, **kwargs)

    write_dat("pts", pts, **kwargs)
    write_dat("ptf", ptf, **kwargs)
    write_dat("ptd", ptd, **kwargs)

    write_dat("cgs", cgs, **kwargs)
    write_dat("cgf", cgf, **kwargs)
    write_dat("rca", rca, **kwargs)
    write_dat("rfi", rfi, **kwargs)
    write_dat("rfc", rfc, **kwargs)

    if kwargs.get("change_devices") == True:
        write_dat("cnf", cnf, **kwargs)
        write_dat("lsc", lsc, **kwargs)
        write_dat("nv1", nv1, **kwargs)
        write_dat("nv2", nv2, **kwargs)
        write_dat("tac", tac, **kwargs)
        write_dat("tdd", tdd, **kwargs)
        write_dat("mul", mul, **kwargs)
        write_dat("enm", enm, **kwargs)
        write_dat("cxu", cxu, **kwargs)
        write_dat("utr", utr, **kwargs)
        write_dat("map", map, **kwargs)


def _change_eqp_tag(old_tag, new_tag, fields=[], eqp_type="LINE", **kwargs):
    '''
    Substutui a tag de um equipamento. old_tag indica a tag_atual e new_tag a nova tag. A fun‡„o substitui automaticamente
    os pontos de disjuntor e chave. Ex.: change_eqp_tag("04C1","04F1",base="ssl-gvm"). As seguintes tabelas s„o
    alteradas: pds, pdf, pdd, pas, paf, pad, pts, ptf, ptd, cgs, cgf, rca, rfi, rfc
    Por padr„o, os seguintes campos s„o afetados: ID, PNT, NOME, DESC1, CMT, PNT, PDS, PAS, PTS, PINT,
    PAC, CGS, KCONV, PARC, OBSRV
    :param old_tag: tag atual do equipamento
    :param new_tag: nova tag do euipamento
    :param kwargs: source_path= str com o caminho pra base; base=nome da base (sage apenas)
            change_devices = True muda as tabelas de configura‡„o f¡sica (LSC, CNF, MUL, TAC, etc). default = False
    :return:
    '''
    old = []
    new = []
    if fields == []:
        fields=["ID", "PNT", "NOME", "DESC1", "CMT", "PNT", "PDS", "PAS", "PTS", "PINT",
            "PAC", "CGS", "KCONV", "PARC", "OBSRV"]
    if kwargs.get("change_devices") == True:
        print_msg(__name__,"as configura‡”es de liga‡„o f¡sica ser„o alteradas", msg_type=MSG_INFO, **kwargs)
        fields.extend(["TAC","LSC","TDD","NV1","CNF","MUL","ENM","CXU","NV2","CONFIG","UTR","MAP","NARRT"])
    if eqp_type == "LINE":
        old = [old_tag, "1"+old_tag[1:], "3"+old_tag[1:]]
        new = [new_tag, "1"+new_tag[1:], "3"+new_tag[1:]]
    elif eqp_type == "TRF":
        old = [old_tag, "1"+old_tag[1:], "3"+old_tag[1:], "12"+old_tag[2:], "32"+old_tag[2:], "11"+old_tag[2:]]
        new = [new_tag, "1"+new_tag[1:], "3"+new_tag[1:], "12"+new_tag[2:], "32"+new_tag[2:], "11"+new_tag[2:]]

    if (old != []) and (new != []):
        change_tags(old, new, fields=fields, **kwargs)

def change_trf_tag(old_tag, new_tag, fields=[], **kwargs):
    _change_eqp_tag(old_tag=old_tag, new_tag=new_tag, fields=fields, eqp_type="TRF", **kwargs)

def change_line_tag(old_tag, new_tag, fields=[], **kwargs):
    _change_eqp_tag(old_tag=old_tag, new_tag=new_tag, fields=fields, eqp_type="LINE", **kwargs)





'''
# TESTE DE ADD_ITEM_TO_BASE
item = {
    "ID": "LSC_TESTE_01",
    "CNF": "TESTE"
}
add_item_to_base("lsc", item=item, source_path="bd/tad", add_to="novo")
'''


'''
# TESTE DE GET LOGICAL CONF
w = get_aconf_from_base("cgs", item_id="i12-LLN0$CO$LEDRs", source_path="bd/curso61850")
#print(w)
print_logical_conf(w)

#w = get_tcv_conf("cgs", item_id="GVM:34C7-4:89", source_path="bd/gvm")
#print(w)

#print(str(is_dnp3("cgs", item_id="PFAIL_ENUP_GVM", source_path="bd/gvm")))
'''

'''
#TESTE DE GET LOGICAL / CONTROL DIST
#w = get_logical_dist("pds", item_id="NOVO_ID", source_path="bd/curso")

#print(w)

w = get_control_dist(item_id="TAD:34M3-2:89", source_path="bd/tad")
print(str(w))

'''



'''

# TESTE DE UPDATE ITEM
source = "bd/gvm/"
dt = "pds"
fields = {
    "pds.id":"GVM:34C8-6:89",
    "pds.nome": "Seccionadora 34C8-6",
    "pdf.id": "h11-CSWI4$ST$Pos",
    "pdf.nv2": "h11_ADAQ"
}

update_item_61850(dat_type=dt, item_id="GVM:34C7-6:89", fields=fields, update_related=True, source_path=source)
'''

'''
# TESTE DE REMOVE_ITEM_61850
source = "bd/gvm/"
dt="pds"
remove_item_61850(dat_type=dt, item_id="GVM:34C7-4:00:CSBL", source_path=source, cascade=True, force_remove=False)
'''

'''
# TESTE DE REMOVE_ITEM_CALC
source = "bd/gvm/"
dt="pds"
remove_item_calc(dat_type=dt, item_id="TESTE_PTS_CALC_1", source_path=source, cascade=True, force_remove=False)
'''

'''
# TESTE DE ADD_PDS_CALC
source = "bd/gvm/"
parc1 = {"PARC": "PONTO_PDS_1", "TIPOP": "EDC", "TPPARC": "PDS"}
parc2 = {"PARC": "PONTO_PDS_2", "TIPOP": "EDC", "TPPARC": "PDS"}
#parcs = [parc1, parc2]
parcs = []
clone_id="GVM:14C7:52:IATVD"
add_to = "TESTE"
add_pds_calc(pds_id="TESTE_PDS_CALC_2", parcs=parcs, clone_id=clone_id, source_path=source, verbose=True, add_to=add_to)
'''



'''
# TESTE DE ADD_PDS_61850
source = "bd/gvm/"
#recs = load_pds(source_path=source)
#print_dat(recs)
dt = "cgs"
fields = {"pds.nome": "Esse ‚ um teste 13"}
fields["pds.kconv"] = "NPS5"
fields["pds.desc1"] = "Descri‡„o para este ponto 13"
#add_pds_61850(pds_id="GVM:04C7:F1:SRET27", fields=fields, logical_device="F1_04C7Protection", \
#              address="teste sret27", source_path = source, add_to="novo5")
add_item_61850(dat_type=dt, item_id="GVM:34C8-10:89", clone_id="GVM:34C7-4:89", \
             source_path = source, add_to="NOVO3", address="teste-cgs-10")
'''