import streamlit as st
import xml.etree.ElementTree as ET
from datetime import datetime

# ─────────────────────────────────────────────
# VERSÃO
# ─────────────────────────────────────────────
VERSAO = "V1.1"

# ─────────────────────────────────────────────
# TEMA THOMSON REUTERS
# ─────────────────────────────────────────────
def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] {
            font-family: 'Segoe UI', 'Arial', sans-serif;
            color: #444444;
        }
        h1, h2, h3 {
            color: #FF8000;
            font-weight: 700;
        }
        section[data-testid="stSidebar"] {
            background-color: #444444;
            color: #FFFFFF;
        }
        section[data-testid="stSidebar"] * {
            color: #FFFFFF !important;
        }
        .stButton > button {
            background-color: #FF8000;
            color: #FFFFFF;
            border: none;
            border-radius: 4px;
            font-weight: bold;
        }
        .stButton > button:hover {
            background-color: #D64001;
            color: #FFFFFF;
        }
        .stDownloadButton > button {
            background-color: #FF8000;
            color: #FFFFFF;
            border: none;
            border-radius: 4px;
            font-weight: bold;
        }
        .stDownloadButton > button:hover {
            background-color: #D64001;
            color: #FFFFFF;
        }
        hr {
            border-color: #FF8000;
        }
        [data-testid="metric-container"] {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px;
            padding: 10px;
        }
        .instrucoes-box {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px;
            padding: 16px 20px;
            margin: 12px 0;
            color: #444444;
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        .instrucoes-box h4 {
            color: #FF8000;
            margin-top: 14px;
            margin-bottom: 6px;
        }
        .instrucoes-box h4:first-child {
            margin-top: 0;
        }
        .cnpj-badge {
            background-color: #444444;
            color: #FF8000;
            border: 1px solid #FF8000;
            border-radius: 4px;
            padding: 6px 12px;
            font-family: Consolas, monospace;
            font-size: 13px;
            display: inline-block;
            margin: 4px 0;
        }
        .info-origem {
            background-color: #FFF3E0;
            border-left: 3px solid #FF8000;
            border-radius: 3px;
            padding: 8px 12px;
            font-size: 12px;
            color: #444444;
            margin: 6px 0;
        }
        </style>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Domínio Sistemas | Thomson Reuters",
    page_icon="🟠",
    layout="wide",
    initial_sidebar_state="expanded",
)
apply_tr_theme()

# ─────────────────────────────────────────────
# CABEÇALHO
# ─────────────────────────────────────────────
st.markdown(
    f"""
    <div style="background:#444444; padding:24px 28px 18px 28px; border-radius:8px;
                border-top:6px solid #FF8000; margin-bottom:28px;">
        <h2 style="color:#FF8000; margin:0; font-family:'Segoe UI',Arial,sans-serif;">
            🔄 Conversor XML NF-e → Domínio Sistemas &nbsp;|&nbsp; {VERSAO}
        </h2>
        <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
            Converte arquivos XML de NF-e para o leiaute padrão de importação do
            <strong>Domínio Sistemas</strong>. Saída em <strong>ANSI (Latin-1)</strong>.
            O CNPJ da empresa é lido automaticamente do XML (<code>&lt;dest&gt;</code>).
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────
# NAMESPACE NF-e
# ─────────────────────────────────────────────
NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

# ─────────────────────────────────────────────
# TABELA DE IMPOSTOS – Impostos.xls
# ─────────────────────────────────────────────
TABELA_IMPOSTOS = {
    1:   "ICMS",
    2:   "IPI",
    3:   "ISS",
    4:   "PIS",
    5:   "COFINS",
    6:   "CONTRIBUICAO SOCIAL",
    7:   "IRPJ-LP",
    8:   "DIFALI",
    9:   "SUBST. TRIBUTARIA",
    10:  "SIMPLES",
    11:  "ICMS RETIDO",
    12:  "ICMS SUBSTITUTO (RS)",
    14:  "REFIS",
    15:  "FIA",
    16:  "IRRF",
    17:  "PIS(nao cumulativo)",
    18:  "ISS Retido",
    19:  "COFINS(nao cumulativa)",
    20:  "Fundo de Fomento e Desenvolvimento",
    21:  "Recomposicao de Aliquota",
    22:  "PIS(Retido)",
    23:  "COFINS(Retido)",
    24:  "Contribuicao Social(Retida)",
    25:  "Contribuicao Retida na Fonte",
    26:  "INSS Retido",
    27:  "ICMS Antecipado(Atacadista)",
    28:  "FUNRURAL",
    29:  "Contr p/ Prog. Inc Arrec Educ tributaria",
    30:  "IPI-M",
    31:  "ICMS ST Antecipacao Total",
    32:  "FUNDOSOCIAL",
    33:  "IRPJ postergado",
    34:  "ICMS Antecipado(Farmacia)",
    35:  "FUNCULTURAL",
    36:  "FUNTURISMO",
    37:  "FUNDESPORTE",
    38:  "Contribuicoes Retidas na Fonte - Orgaos Publicos",
    39:  "IRRF Propaganda",
    40:  "INSS COOPERATIVAS",
    41:  "Fundo para o Desenv. Tec. das Telec.",
    42:  "ICMS Fato Gerador",
    43:  "ICMS Adicional(2%)",
    44:  "SIMPLES NACIONAL",
    45:  "ICMS Importacao",
    46:  "ICMS Antecipado Combustiveis",
    47:  "Subst. Trib. Antecipada Combustiveis",
    48:  "ICMS Garantido Integral",
    49:  "FECOEP-FUNDO EST. DE COMB E ERRAD. POBRE",
    50:  "Fundo de Transportes e Habitacao- FETHAB",
    55:  "FECOP-ICMS Normal",
    56:  "FECOP-ICMS ST OP.Internas",
    57:  "Fundo de Combate a Pobreza ST - Operacoes Interestaduais",
    58:  "Fundo Pro-Emprego",
    59:  "Apoio a Manut. e Desenv. Educ. Superior",
    60:  "Fundo Protecao Social do estado de Goias",
    61:  "FUNDAF",
    62:  "Fundo Geracao de Emprego e Renda do DF",
    63:  "IRRF alugueis pessoa fisica",
    64:  "Simples MEI",
    65:  "TX ADIC. FOMENTAR/MICROPRODUZIR/PRODUZIR",
    66:  "Fundo Universal dos Servicos de Telecom.",
    67:  "PIS - Substituicao Tributaria",
    68:  "COFINS - Substituicao Tributaria",
    69:  "ICMS Subst.Tributaria Serv. Transportes",
    70:  "Comp. Financeira pela Expl. Rec.Minerais",
    82:  "FECP - ICMS Importacao",
    83:  "RET - Regime especial incorporacoes imobiliarias",
    84:  "PIS - Regime especial incorporacoes imobiliarias",
    85:  "COFINS - Regime especial incorporacoes imobiliarias",
    86:  "IRPJ - Regime especial incorporacoes imobiliarias",
    87:  "CSLL - Regime especial incorporacoes imobiliarias",
    88:  "Regime especial PMCMV",
    89:  "PIS Regime especial PMCMV",
    90:  "COFINS Regime especial PMCMV",
    91:  "IRPJ - Regime especial PMCMV",
    92:  "CSLL - Regime especial PMCMV",
    93:  "ICMS Antecipado - Farinha de Trigo",
    94:  "FECOP - ICMS Antecipado",
    95:  "FUNDESA",
    96:  "FEM - ICMS Normal",
    97:  "FEM - ICMS Substituicao Tributaria",
    98:  "ICMS Antecipado Servicos de Transportes",
    99:  "Fundo Estadual de Defesa Civil - FUNDEC",
    100: "FUNDEIC",
    101: "FUNDED",
    102: "FUPIS",
    103: "Contribuicao Previdenciaria sobre a Receita Bruta",
    104: "PIS Entidades Financeiras e Equiparadas",
    105: "COFINS Entidades Financeiras e Equiparadas",
    106: "PIS Nao cumulativo - SCP",
    107: "COFINS Nao cumulativo - SCP",
    108: "PIS - SCP",
    109: "COFINS - SCP",
    110: "Contribuicao Social - SCP",
    111: "IRPJ - SCP",
    112: "Fundo de Combate - ICMS Antecipado",
    113: "Fundo de Combate - ICMS NORMAL",
    114: "Fundo de Combate - ICMS ST",
    115: "Fundo Desenvolvimento Economico-FDE",
    116: "ICMS DIFERIDO",
    117: "CIDE",
    118: "FAPESC",
    119: "FITUR",
    120: "PROUNIV",
    121: "ICMS Antecipado Telecomunicacoes",
    122: "FAI",
    125: "ICMS Complementar",
    126: "FUNDEPEC",
    127: "FUNPRODUZIR",
    128: "ICMS Carga Media",
    129: "ICMS ST Carga Media",
    130: "FUNDAP",
    132: "INSS Retido - SCP",
    133: "PIS Importacao",
    134: "COFINS Importacao",
    135: "Fundo Estadual de Saude - FES",
    136: "Programa Bolsa Universitaria",
    138: "PIS - Codigo de recolhimento",
    139: "COFINS - Codigo de recolhimento",
    140: "PIS Nao cumulativo - Codigo de recolhimento",
    141: "COFINS Nao cumulativa - Codigo de recolhimento",
    142: "CPRB - SCP",
    145: "ICMS Diferencial de Aliquota Antecipado",
    146: "Fundo Estadual de Combate a Pobreza",
    147: "FECOEP-SP - ICMS Proprio",
    148: "FECOEP-SP - ICMS ST",
    149: "FECOEP-TO - ICMS Proprio",
    150: "FECOEP-TO - ICMS ST",
    152: "FECOP-PR - ICMS Proprio",
    153: "FECOP-PR - ICMS ST",
    154: "FEEF",
    155: "FECOMP-MS - ICMS Proprio",
    156: "FECOMP-MS - ICMS ST",
    157: "FECOEP-RO - ICMS Proprio",
    158: "FECOEP-RO - ICMS ST",
    159: "FECEP-MT - ICMS ES",
    160: "FECEP-MT - ICMS ST",
    161: "IOF",
    162: "Fundo Desenv. Sistema Rodoviario MS",
    163: "Fundo Milho e Soja MS",
    164: "ICMS Equalizacao Simples Nacional",
    165: "FIDER-RO",
    166: "PROLEITE",
    167: "FITHA-RO",
    168: "RET - SCP",
    169: "PMCMV - SCP",
    170: "FUNTUR",
    172: "FUNEF",
    177: "Fundo Estadual - Incentivos",
    178: "FDI",
    179: "Programa Mais IDH",
    180: "FUNDEINFRA",
    183: "IBS - Imposto sobre Bens e Servicos",
    184: "CBS - Contribuicao Social sobre Bens e Servicos",
    187: "IS - Imposto Seletivo",
}

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def get_text(element, path: str, default: str = "") -> str:
    if element is None:
        return default
    found = element.find(path, NS)
    if found is not None and found.text:
        return found.text.strip()
    return default


def fmt_decimal(value: str, decimals: int = 2) -> str:
    if not value:
        return ""
    try:
        v = float(str(value).replace(",", "."))
        return f"{v:.{decimals}f}".replace(".", ",")
    except (ValueError, TypeError):
        return str(value)


def fmt_date(iso_date: str) -> str:
    if not iso_date:
        return ""
    try:
        d = datetime.strptime(iso_date[:10], "%Y-%m-%d")
        return d.strftime("%d/%m/%Y")
    except ValueError:
        return iso_date


def pipe_join(fields: list) -> str:
    return "|" + "|".join(str(f) for f in fields) + "|"


def encode_ansi(conteudo: str) -> bytes:
    resultado = []
    for char in conteudo:
        try:
            resultado.append(char.encode("latin-1"))
        except UnicodeEncodeError:
            resultado.append(b"?")
    return b"".join(resultado)


# ─────────────────────────────────────────────
# EXTRAÇÃO DO CNPJ DESTINATÁRIO DO XML
# Prioridade: dest/CNPJ → dest/CPF → dest/idEstrangeiro
# Fallback: campo manual da sidebar
# ─────────────────────────────────────────────
def extrair_cnpj_dest(nfe_root) -> tuple:
    """
    Retorna (cnpj_str, origem_str).
    origem pode ser: 'XML - dest/CNPJ', 'XML - dest/CPF',
                     'XML - dest/idEstrangeiro', 'Manual'
    """
    dest = nfe_root.find("nfe:infNFe/nfe:dest", NS)
    if dest is not None:
        # Tenta CNPJ
        cnpj = get_text(dest, "nfe:CNPJ")
        if cnpj:
            return cnpj, "XML — <dest><CNPJ>"

        # Tenta CPF (pessoa física)
        cpf = get_text(dest, "nfe:CPF")
        if cpf:
            return cpf, "XML — <dest><CPF>"

        # Tenta idEstrangeiro (exportação)
        id_ext = get_text(dest, "nfe:idEstrangeiro")
        if id_ext:
            return id_ext, "XML — <dest><idEstrangeiro>"

    return "", "Não encontrado no XML"


def extrair_nome_dest(nfe_root) -> str:
    dest = nfe_root.find("nfe:infNFe/nfe:dest", NS)
    if dest is not None:
        return get_text(dest, "nfe:xNome")
    return ""


# ─────────────────────────────────────────────
# REGISTRO 0000
# ─────────────────────────────────────────────
def gerar_registro_0000(cnpj_empresa: str) -> str:
    return pipe_join(["0000", cnpj_empresa])


# ─────────────────────────────────────────────
# REGISTRO 0020 – Cadastro de fornecedor (33 campos)
# ─────────────────────────────────────────────
def gerar_registro_0020(emit) -> str:
    cnpj        = get_text(emit, "nfe:CNPJ")
    razao       = get_text(emit, "nfe:xNome")[:150]
    fantasia    = get_text(emit, "nfe:xFant")[:40]
    ender       = emit.find("nfe:enderEmit", NS)
    logradouro  = get_text(ender, "nfe:xLgr")
    numero      = get_text(ender, "nfe:nro")
    complemento = get_text(ender, "nfe:xCpl")
    bairro      = get_text(ender, "nfe:xBairro")
    cod_mun     = get_text(ender, "nfe:cMun")
    uf          = get_text(ender, "nfe:UF")
    cep         = get_text(ender, "nfe:CEP")
    ie          = get_text(emit, "nfe:IE")
    crt         = get_text(emit, "nfe:CRT")

    regime_map   = {"1": "M", "2": "E", "3": "N"}
    regime       = regime_map.get(crt, "N")
    contrib_icms = "S" if ie and ie.upper() not in ("ISENTO", "NAO CONTRIBUINTE", "") else "N"

    return pipe_join([
        "0020", cnpj, razao, fantasia, logradouro, numero, complemento,
        bairro, cod_mun, uf, "", cep, ie, "", "", "", "", "", "", "", "",
        "N", "7", regime, contrib_icms, "", "", "", "", "N", "N", "", "",
    ])


# ─────────────────────────────────────────────
# REGISTRO 0100 – Cadastro de produtos (91 campos)
# ─────────────────────────────────────────────
def gerar_registro_0100(det) -> str:
    prod      = det.find("nfe:prod", NS)
    cod_prod  = get_text(prod, "nfe:cProd")[:14]
    descricao = get_text(prod, "nfe:xProd")
    ncm       = get_text(prod, "nfe:NCM")
    unidade   = get_text(prod, "nfe:uCom")
    val_unit  = get_text(prod, "nfe:vUnCom")
    cest      = get_text(prod, "nfe:CEST")

    imposto   = det.find("nfe:imposto", NS)
    cst_icms  = ""
    aliq_icms = ""
    aliq_ipi  = ""

    if imposto is not None:
        for icms_type in ["ICMS00","ICMS10","ICMS20","ICMS30","ICMS40",
                          "ICMS51","ICMS60","ICMS70","ICMS90","ICMSSN101",
                          "ICMSSN102","ICMSSN201","ICMSSN202","ICMSSN500","ICMSSN900"]:
            node = imposto.find(f"nfe:ICMS/nfe:{icms_type}", NS)
            if node is not None:
                cst_icms  = get_text(node, "nfe:CST") or get_text(node, "nfe:CSOSN")
                aliq_icms = fmt_decimal(get_text(node, "nfe:pICMS"))
                break
        ipi_trib = imposto.find("nfe:IPI/nfe:IPITrib", NS)
        if ipi_trib is not None:
            aliq_ipi = fmt_decimal(get_text(ipi_trib, "nfe:pIPI"))

    return pipe_join([
        "0100", cod_prod, descricao, "", ncm, "", "", "", "", unidade,
        "N", "O", "", "", "", "N", "", fmt_decimal(val_unit, 3), "", "",
        cst_icms, aliq_icms, aliq_ipi, "", "", "N", "", "", "", "", "", "",
        "", "", "N", "", "", "", "N", "", "", "", "N", "", "", "", "N", "",
        "", "", "", "", "", "", "N", "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", "", "", "", "", "", "N", "", "", "N", "", "",
        "", "", "", "", "", "", "", "", "", "", cest, "", "",
    ])


# ─────────────────────────────────────────────
# REGISTRO 1000 – NF de Entrada (98 campos)
# ─────────────────────────────────────────────
def gerar_registro_1000(nfe_root, cnpj_empresa: str,
                        acumulador: str = "1157",
                        especie: str = "36") -> str:
    ide   = nfe_root.find("nfe:infNFe/nfe:ide", NS)
    emit  = nfe_root.find("nfe:infNFe/nfe:emit", NS)
    total = nfe_root.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)

    cnpj_emit  = get_text(emit, "nfe:CNPJ")
    nNF        = get_text(ide, "nfe:nNF")
    serie      = get_text(ide, "nfe:serie")
    dhEmi      = fmt_date(get_text(ide, "nfe:dhEmi"))
    ie_emit    = get_text(emit, "nfe:IE")
    c_mun_fg   = get_text(ide, "nfe:cMunFG")

    det_list   = nfe_root.findall("nfe:infNFe/nfe:det", NS)
    cfop_first = ""
    if det_list:
        cfop_first = get_text(det_list[0].find("nfe:prod", NS), "nfe:CFOP")

    v_nf     = fmt_decimal(get_text(total, "nfe:vNF"))
    v_pis    = fmt_decimal(get_text(total, "nfe:vPIS"))
    v_cofins = fmt_decimal(get_text(total, "nfe:vCOFINS"))
    v_ipi    = fmt_decimal(get_text(total, "nfe:vIPI"))
    v_st     = fmt_decimal(get_text(total, "nfe:vST"))
    v_prod   = fmt_decimal(get_text(total, "nfe:vProd"))
    v_frete  = fmt_decimal(get_text(total, "nfe:vFrete"))
    v_seg    = fmt_decimal(get_text(total, "nfe:vSeg"))
    v_outro  = fmt_decimal(get_text(total, "nfe:vOutro"))
    v_icms_d = fmt_decimal(get_text(total, "nfe:vICMSDeson"))

    inf_nfe  = nfe_root.find("nfe:infNFe", NS)
    chave    = ""
    if inf_nfe is not None:
        id_attr = inf_nfe.get("Id", "")
        chave   = id_attr.replace("NFe", "") if id_attr else ""

    transp        = nfe_root.find("nfe:infNFe/nfe:transp", NS)
    mod_frete_cod = get_text(transp, "nfe:modFrete")
    frete_map     = {"0": "C", "1": "F", "2": "S", "3": "T", "4": "R", "5": "D", "9": "S"}
    mod_frete     = frete_map.get(mod_frete_cod, "C")

    inf_adic  = nfe_root.find("nfe:infNFe/nfe:infAdic", NS)
    obs_fisco = ""
    if inf_adic is not None:
        obs_fisco = get_text(inf_adic, "nfe:infAdFisco")[:300]

    n_di = ""
    if det_list:
        di_node = det_list[0].find("nfe:prod/nfe:DI", NS)
        if di_node is not None:
            n_di = get_text(di_node, "nfe:nDI")

    return pipe_join([
        "1000", especie, cnpj_emit, "", acumulador, cfop_first, "",
        nNF, serie, "", dhEmi, dhEmi, v_nf, "", obs_fisco, mod_frete,
        "T", "", "", "", "", "", "", "", "", v_frete, v_seg, v_outro,
        v_pis, "", v_cofins, "", "", "", "", "", "", "", v_prod,
        c_mun_fg, "0", "", "", ie_emit, "", "", "", "", "", "", "",
        n_di, "N", chave, "", "", "", "", "", "1", "", "", "", "", "",
        "", "", "", "", "", "", "", "", "", "", "10", "", "", "", "",
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
        v_ipi, v_st, "", "", "", "", "", v_icms_d, "",
    ])


# ─────────────────────────────────────────────
# REGISTROS 1020 – Impostos (um por imposto)
# ─────────────────────────────────────────────
def gerar_registros_1020(nfe_root) -> list:
    total  = nfe_root.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)
    ibs_t  = nfe_root.find("nfe:infNFe/nfe:total/nfe:IBSCBSTot", NS)
    v_nf   = fmt_decimal(get_text(total, "nfe:vNF"))
    linhas = []

    def r1020(cod, perc_red="", base="", aliq="", valor="",
              isentas="", outras="", v_ipi="", v_st="", v_cont="", cod_rec=""):
        return pipe_join([
            "1020", cod, perc_red, base, aliq, valor,
            isentas, outras, v_ipi, v_st, v_cont, cod_rec,
            "", "", "", "", "", "", "",
        ])

    v_bc_icms    = get_text(total, "nfe:vBC")
    v_icms       = get_text(total, "nfe:vICMS")
    v_ipi_tot    = get_text(total, "nfe:vIPI")
    v_st_tot     = get_text(total, "nfe:vST")
    v_pis_tot    = get_text(total, "nfe:vPIS")
    v_cofins_tot = get_text(total, "nfe:vCOFINS")
    v_icms_deson = get_text(total, "nfe:vICMSDeson")

    try:
        bc_f  = float(v_bc_icms)
        icm_f = float(v_icms)
        aliq_icms_med = fmt_decimal(str(icm_f / bc_f * 100)) if bc_f > 0 else ""
    except (ValueError, ZeroDivisionError):
        aliq_icms_med = ""

    det_list     = nfe_root.findall("nfe:infNFe/nfe:det", NS)
    bc_ipi_total = 0.0
    bc_pis_total = 0.0
    bc_cof_total = 0.0

    for det in det_list:
        imp = det.find("nfe:imposto", NS)
        if imp is None:
            continue
        ipi_t = imp.find("nfe:IPI/nfe:IPITrib", NS)
        if ipi_t is not None:
            try:
                bc_ipi_total += float(get_text(ipi_t, "nfe:vBC") or "0")
            except ValueError:
                pass
        for pt in ["PISAliq", "PISQtde", "PISNT", "PISOutr"]:
            pn = imp.find(f"nfe:PIS/nfe:{pt}", NS)
            if pn is not None:
                try:
                    bc_pis_total += float(get_text(pn, "nfe:vBC") or "0")
                except ValueError:
                    pass
                break
        for ct in ["COFINSAliq", "COFINSQtde", "COFINSNT", "COFINSOutr"]:
            cn = imp.find(f"nfe:COFINS/nfe:{ct}", NS)
            if cn is not None:
                try:
                    bc_cof_total += float(get_text(cn, "nfe:vBC") or "0")
                except ValueError:
                    pass
                break

    # 1 – ICMS
    if v_icms and float(v_icms) > 0:
        linhas.append(r1020(1,
            base   = fmt_decimal(v_bc_icms),
            aliq   = aliq_icms_med,
            valor  = fmt_decimal(v_icms),
            v_ipi  = fmt_decimal(v_ipi_tot),
            v_st   = fmt_decimal(v_st_tot),
            v_cont = v_nf,
        ))

    # 2 – IPI
    if v_ipi_tot and float(v_ipi_tot) > 0:
        linhas.append(r1020(2,
            base   = fmt_decimal(str(bc_ipi_total)),
            valor  = fmt_decimal(v_ipi_tot),
            v_cont = v_nf,
        ))

    # 4 – PIS
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(r1020(4,
            base   = fmt_decimal(str(bc_pis_total)),
            valor  = fmt_decimal(v_pis_tot),
            v_cont = v_nf,
        ))

    # 5 – COFINS
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(r1020(5,
            base   = fmt_decimal(str(bc_cof_total)),
            valor  = fmt_decimal(v_cofins_tot),
            v_cont = v_nf,
        ))

    # 45 – ICMS Importação (desonerado)
    if v_icms_deson and float(v_icms_deson) > 0:
        linhas.append(r1020(45,
            base   = fmt_decimal(v_bc_icms),
            aliq   = aliq_icms_med,
            valor  = fmt_decimal(v_icms_deson),
            v_cont = v_nf,
        ))

    # 133 – PIS Importação
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(r1020(133,
            base   = fmt_decimal(str(bc_pis_total)),
            valor  = fmt_decimal(v_pis_tot),
            v_cont = v_nf,
        ))

    # 134 – COFINS Importação
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(r1020(134,
            base   = fmt_decimal(str(bc_cof_total)),
            valor  = fmt_decimal(v_cofins_tot),
            v_cont = v_nf,
        ))

    # 183 – IBS
    if ibs_t is not None:
        v_ibs_uf = get_text(ibs_t, "nfe:gIBS/nfe:gIBSUF/nfe:vIBSUF")
        bc_ibs   = get_text(ibs_t, "nfe:vBCIBSCBS")
        if v_ibs_uf and float(v_ibs_uf) > 0:
            linhas.append(r1020(183,
                base   = fmt_decimal(bc_ibs),
                valor  = fmt_decimal(v_ibs_uf),
                v_cont = v_nf,
            ))

    # 184 – CBS
    if ibs_t is not None:
        v_cbs  = get_text(ibs_t, "nfe:gCBS/nfe:vCBS")
        bc_cbs = get_text(ibs_t, "nfe:vBCIBSCBS")
        if v_cbs and float(v_cbs) > 0:
            linhas.append(r1020(184,
                base   = fmt_decimal(bc_cbs),
                valor  = fmt_decimal(v_cbs),
                v_cont = v_nf,
            ))

    return linhas


# ─────────────────────────────────────────────
# REGISTRO 1030 – Estoque / item (111 campos)
# ─────────────────────────────────────────────
def gerar_registro_1030(det, seq: int) -> str:
    prod    = det.find("nfe:prod", NS)
    imposto = det.find("nfe:imposto", NS)

    cod_prod = get_text(prod, "nfe:cProd")[:14]
    qtd      = fmt_decimal(get_text(prod, "nfe:qCom"), 2)
    v_prod   = get_text(prod, "nfe:vProd")
    v_outro  = get_text(prod, "nfe:vOutro")
    v_desc   = get_text(prod, "nfe:vDesc")
    cfop     = get_text(prod, "nfe:CFOP")
    unidade  = get_text(prod, "nfe:uCom")
    v_unit   = get_text(prod, "nfe:vUnCom")
    cest     = get_text(prod, "nfe:CEST")

    di_node = prod.find("nfe:DI", NS)
    n_di    = get_text(di_node, "nfe:nDI") if di_node is not None else ""
    d_di    = fmt_date(get_text(di_node, "nfe:dDI")) if di_node is not None else ""

    icms_node  = None
    v_bc_icms  = ""
    aliq_icms  = ""
    v_icms     = ""
    cst_icms   = ""
    v_icms_des = ""
    v_bc_st    = ""

    if imposto is not None:
        for icms_type in ["ICMS00","ICMS10","ICMS20","ICMS30","ICMS40",
                          "ICMS51","ICMS60","ICMS70","ICMS90","ICMSSN101",
                          "ICMSSN102","ICMSSN201","ICMSSN202","ICMSSN500","ICMSSN900"]:
            node = imposto.find(f"nfe:ICMS/nfe:{icms_type}", NS)
            if node is not None:
                icms_node = node
                break
        if icms_node is not None:
            v_bc_icms  = fmt_decimal(get_text(icms_node, "nfe:vBC"))
            aliq_icms  = fmt_decimal(get_text(icms_node, "nfe:pICMS"))
            v_icms     = fmt_decimal(get_text(icms_node, "nfe:vICMS"))
            cst_icms   = get_text(icms_node, "nfe:CST") or get_text(icms_node, "nfe:CSOSN")
            v_icms_des = fmt_decimal(get_text(icms_node, "nfe:vICMSDeson"))
            v_bc_st    = fmt_decimal(get_text(icms_node, "nfe:vBCST"))

        ipi_trib = imposto.find("nfe:IPI/nfe:IPITrib", NS)
        ipi_nt   = imposto.find("nfe:IPI/nfe:IPINT", NS)
        if ipi_trib is not None:
            v_ipi    = fmt_decimal(get_text(ipi_trib, "nfe:vIPI"))
            aliq_ipi = fmt_decimal(get_text(ipi_trib, "nfe:pIPI"))
            cst_ipi  = get_text(ipi_trib, "nfe:CST")
        elif ipi_nt is not None:
            v_ipi    = "0,00"
            aliq_ipi = "0,00"
            cst_ipi  = get_text(ipi_nt, "nfe:CST")
        else:
            v_ipi = aliq_ipi = cst_ipi = ""

        v_pis = aliq_pis = cst_pis = bc_pis = ""
        pis_node = imposto.find("nfe:PIS", NS)
        if pis_node is not None:
            for pt in ["PISAliq", "PISQtde", "PISNT", "PISOutr"]:
                pn = pis_node.find(f"nfe:{pt}", NS)
                if pn is not None:
                    v_pis    = fmt_decimal(get_text(pn, "nfe:vPIS"))
                    aliq_pis = fmt_decimal(get_text(pn, "nfe:pPIS"), 4)
                    cst_pis  = get_text(pn, "nfe:CST")
                    bc_pis   = fmt_decimal(get_text(pn, "nfe:vBC"))
                    break

        v_cof = aliq_cof = cst_cof = bc_cof = ""
        cof_node = imposto.find("nfe:COFINS", NS)
        if cof_node is not None:
            for ct in ["COFINSAliq", "COFINSQtde", "COFINSNT", "COFINSOutr"]:
                cn = cof_node.find(f"nfe:{ct}", NS)
                if cn is not None:
                    v_cof    = fmt_decimal(get_text(cn, "nfe:vCOFINS"))
                    aliq_cof = fmt_decimal(get_text(cn, "nfe:pCOFINS"), 4)
                    cst_cof  = get_text(cn, "nfe:CST")
                    bc_cof   = fmt_decimal(get_text(cn, "nfe:vBC"))
                    break

        ibs_node       = imposto.find("nfe:IBSCBS", NS)
        ibs_class_trib = ibs_bc = ibs_aliq = ibs_val = ""
        cbs_bc = cbs_aliq = cbs_val = ""
        if ibs_node is not None:
            ibs_class_trib = get_text(ibs_node, "nfe:cClassTrib")
            gibs = ibs_node.find("nfe:gIBSCBS", NS)
            if gibs is not None:
                ibs_bc = fmt_decimal(get_text(gibs, "nfe:vBC"))
                guf    = gibs.find("nfe:gIBSUF", NS)
                if guf is not None:
                    ibs_aliq = fmt_decimal(get_text(guf, "nfe:pIBSUF"))
                    ibs_val  = fmt_decimal(get_text(guf, "nfe:vIBSUF"))
                gcbs = gibs.find("nfe:gCBS", NS)
                if gcbs is not None:
                    cbs_bc   = ibs_bc
                    cbs_aliq = fmt_decimal(get_text(gcbs, "nfe:pCBS"))
                    cbs_val  = fmt_decimal(get_text(gcbs, "nfe:vCBS"))
    else:
        v_ipi = aliq_ipi = cst_ipi = ""
        v_pis = aliq_pis = cst_pis = bc_pis = ""
        v_cof = aliq_cof = cst_cof = bc_cof = ""
        ibs_class_trib = ibs_bc = ibs_aliq = ibs_val = ""
        cbs_bc = cbs_aliq = cbs_val = ""

    try:
        vp = float(v_prod or "0")
        vi_str = ""
        if imposto is not None:
            ipi_t2 = imposto.find("nfe:IPI/nfe:IPITrib", NS)
            if ipi_t2 is not None:
                vi_str = get_text(ipi_t2, "nfe:vIPI")
        vi = float(vi_str or "0")
        v_total = fmt_decimal(str(vp + vi))
    except (ValueError, TypeError):
        v_total = fmt_decimal(v_prod)

    return pipe_join([
        "1030", cod_prod, qtd, v_total, v_ipi, fmt_decimal(v_prod),
        "1", d_di, n_di, cst_icms, fmt_decimal(v_prod), fmt_decimal(v_desc),
        v_bc_icms, v_bc_st, aliq_icms, "", "", "", "", fmt_decimal(v_outro),
        "", v_icms, "", "", "", "", fmt_decimal(v_unit, 6), "", cst_ipi,
        aliq_ipi, "", "", "", cfop, "", aliq_pis, v_pis, aliq_cof, v_cof,
        fmt_decimal(v_prod), cst_pis, bc_pis, cst_cof, bc_cof,
        "", "", "", "", "", "", "", "", "", "", "", "S", unidade, "", "",
        fmt_decimal(v_prod), "", "", "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", cest,
        "", "", "", "", "", v_icms_des, "", "", "", "", "", "",
        ibs_class_trib, ibs_bc, ibs_aliq, ibs_val,
        ibs_class_trib, cbs_bc, cbs_aliq, cbs_val,
    ])


# ─────────────────────────────────────────────
# REGISTROS 1150 / 1151 – IBS / CBS
# ─────────────────────────────────────────────
def gerar_registro_1150(ct, bc, aliq, valor) -> str:
    return pipe_join(["1150", ct, bc, aliq, valor])


def gerar_registro_1151(ct, bc, aliq, valor) -> str:
    return pipe_join(["1151", ct, bc, aliq, valor])


# ─────────────────────────────────────────────
# CONVERSÃO PRINCIPAL
# ─────────────────────────────────────────────
def converter_xml(
    xml_content: bytes,
    cnpj_fallback: str,
    acumulador: str = "1157",
    especie: str = "36",
    incluir_0000: bool = True,
    incluir_0020: bool = True,
    incluir_0100: bool = True,
) -> tuple:

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        return "", {"erro": str(e)}

    nfe = root.find("nfe:NFe", NS)
    if nfe is None:
        nfe = root

    # ── CNPJ: lê do XML, usa fallback se vazio ────────────────────────
    cnpj_xml, origem_cnpj = extrair_cnpj_dest(nfe)
    nome_dest             = extrair_nome_dest(nfe)

    if cnpj_xml:
        cnpj_empresa = cnpj_xml
    elif cnpj_fallback:
        cnpj_empresa = cnpj_fallback
        origem_cnpj  = "Manual (fallback)"
    else:
        return "", {"erro": "CNPJ do destinatário não encontrado no XML nem informado manualmente."}

    lines    = []
    resumo   = {}
    det_list = nfe.findall("nfe:infNFe/nfe:det", NS)
    emit     = nfe.find("nfe:infNFe/nfe:emit", NS)
    ide      = nfe.find("nfe:infNFe/nfe:ide", NS)
    total    = nfe.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)

    resumo["nNF"]         = get_text(ide, "nfe:nNF")
    resumo["Emitente"]    = get_text(emit, "nfe:xNome")
    resumo["CNPJ Emit"]   = get_text(emit, "nfe:CNPJ")
    resumo["Destinatário"]= nome_dest
    resumo["CNPJ Dest"]   = cnpj_empresa
    resumo["Origem CNPJ"] = origem_cnpj
    resumo["Emissão"]     = fmt_date(get_text(ide, "nfe:dhEmi"))
    resumo["Itens"]       = len(det_list)
    resumo["vNF"]         = fmt_decimal(get_text(total, "nfe:vNF"))
    resumo["vICMS"]       = fmt_decimal(get_text(total, "nfe:vICMS"))
    resumo["vICMSDes"]    = fmt_decimal(get_text(total, "nfe:vICMSDeson"))
    resumo["vIPI"]        = fmt_decimal(get_text(total, "nfe:vIPI"))
    resumo["vPIS"]        = fmt_decimal(get_text(total, "nfe:vPIS"))
    resumo["vCOFINS"]     = fmt_decimal(get_text(total, "nfe:vCOFINS"))
    resumo["vII"]         = fmt_decimal(get_text(total, "nfe:vII"))

    if incluir_0000:
        lines.append(gerar_registro_0000(cnpj_empresa))
    if incluir_0020 and emit is not None:
        lines.append(gerar_registro_0020(emit))
    if incluir_0100:
        produtos_gerados = set()
        for det in det_list:
            cod = get_text(det.find("nfe:prod", NS), "nfe:cProd")
            if cod not in produtos_gerados:
                lines.append(gerar_registro_0100(det))
                produtos_gerados.add(cod)

    lines.append(gerar_registro_1000(nfe, cnpj_empresa, acumulador, especie))

    for r in gerar_registros_1020(nfe):
        lines.append(r)

    for seq, det in enumerate(det_list, start=1):
        lines.append(gerar_registro_1030(det, seq))

    ibs_gerados = {}
    for det in det_list:
        imp = det.find("nfe:imposto", NS)
        if imp is None:
            continue
        ibs_node = imp.find("nfe:IBSCBS", NS)
        if ibs_node is None:
            continue
        ct   = get_text(ibs_node, "nfe:cClassTrib")
        gibs = ibs_node.find("nfe:gIBSCBS", NS)
        if not ct or gibs is None:
            continue
        if ct not in ibs_gerados:
            ibs_gerados[ct] = {"bc_ibs": 0.0, "v_ibs": 0.0, "aliq_ibs": "",
                               "bc_cbs": 0.0, "v_cbs": 0.0, "aliq_cbs": ""}
        try:
            ibs_gerados[ct]["bc_ibs"] += float(get_text(gibs, "nfe:vBC") or "0")
        except ValueError:
            pass
        guf = gibs.find("nfe:gIBSUF", NS)
        if guf is not None:
            try:
                ibs_gerados[ct]["v_ibs"]   += float(get_text(guf, "nfe:vIBSUF") or "0")
                ibs_gerados[ct]["aliq_ibs"] = get_text(guf, "nfe:pIBSUF")
            except ValueError:
                pass
        gcbs = gibs.find("nfe:gCBS", NS)
        if gcbs is not None:
            try:
                ibs_gerados[ct]["bc_cbs"]  += float(get_text(gibs, "nfe:vBC") or "0")
                ibs_gerados[ct]["v_cbs"]   += float(get_text(gcbs, "nfe:vCBS") or "0")
                ibs_gerados[ct]["aliq_cbs"] = get_text(gcbs, "nfe:pCBS")
            except ValueError:
                pass

    for ct, d in ibs_gerados.items():
        lines.append(gerar_registro_1150(ct, fmt_decimal(str(d["bc_ibs"])),
                                         fmt_decimal(d["aliq_ibs"]),
                                         fmt_decimal(str(d["v_ibs"]))))
        lines.append(gerar_registro_1151(ct, fmt_decimal(str(d["bc_cbs"])),
                                         fmt_decimal(d["aliq_cbs"]),
                                         fmt_decimal(str(d["v_cbs"]))))

    return "\n".join(lines), resumo


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"### ℹ Sobre")
    st.markdown(f"**Versão:** {VERSAO}")
    st.markdown("**Thomson Reuters**")
    st.markdown("**Domínio Sistemas**")
    st.markdown("---")

    st.markdown("### ⚙ Parâmetros")

    st.markdown(
        """
        <div style='font-size:12px; color:#DDDDDD; margin-bottom:4px;'>
        🔍 <b>CNPJ da empresa</b> é lido automaticamente do XML
        (<code>&lt;dest&gt;&lt;CNPJ&gt;</code>).<br>
        Use o campo abaixo apenas como <b>fallback</b>.
        </div>
        """,
        unsafe_allow_html=True,
    )
    cnpj_fallback = st.text_input(
        "CNPJ Fallback (opcional)",
        value="",
        max_chars=14,
        help="Usado somente se o XML não contiver o CNPJ do destinatário.",
    )
    acumulador = st.text_input("Código do Acumulador", value="1157",
                               help="1157 = CFOP 3102 (Importação)")
    especie    = st.text_input("Código da Espécie", value="36",
                               help="36 = NF-e modelo 55")

    st.markdown("---")
    st.markdown("### 📋 Registros de cadastro")
    inc_0000 = st.checkbox("0000 – Identificação da empresa", value=True)
    inc_0020 = st.checkbox("0020 – Fornecedor (emitente)", value=True)
    inc_0100 = st.checkbox("0100 – Produtos", value=True)
    st.caption("1000 / 1020 / 1030 / 1150 / 1151 sempre gerados.")

    st.markdown("---")
    st.markdown("### ⚙ Encoding")
    st.markdown("**Entrada:** UTF-8 / Latin-1")
    st.markdown("**Saída:** ANSI (Latin-1)")

    st.markdown("---")
    with st.expander("📋 Tabela de Impostos Domínio"):
        for cod, nome in sorted(TABELA_IMPOSTOS.items()):
            st.caption(f"`{cod:3d}` – {nome}")


# ─────────────────────────────────────────────
# INSTRUÇÕES
# ─────────────────────────────────────────────
with st.expander("📖 Instruções de Uso — clique para expandir", expanded=False):
    st.markdown(
        """
        <div class="instrucoes-box">

        <h4>🔹 Como o CNPJ é obtido</h4>
        <p>O sistema lê automaticamente o CNPJ do <b>destinatário</b> diretamente do XML
        da NF-e, na tag <code>&lt;dest&gt;&lt;CNPJ&gt;</code>.<br>
        Caso o XML não possua CNPJ (ex.: CPF ou idEstrangeiro), o valor correspondente
        é utilizado. O campo <b>CNPJ Fallback</b> na barra lateral é usado apenas quando
        nenhuma inscrição for encontrada no XML.</p>

        <h4>🔹 Passo 1 — Parâmetros</h4>
        <p>Confirme o <b>Acumulador</b> (<code>1157</code> para CFOP 3102) e a
        <b>Espécie</b> (<code>36</code> para NF-e modelo 55) na barra lateral.</p>

        <h4>🔹 Passo 2 — Upload dos XMLs</h4>
        <p>Faça o upload de um ou mais arquivos <b>XML de NF-e</b>
        (nfeProc ou NFe direta).</p>

        <h4>🔹 Passo 3 — Verificar resumo</h4>
        <p>Confira o <b>CNPJ Dest</b> e a <b>Origem CNPJ</b> na tabela de resumo
        para garantir que o CNPJ correto foi lido do XML.</p>

        <h4>🔹 Passo 4 — Baixar e importar</h4>
        <p>Clique em <b>⬇ Baixar Arquivo Domínio (.TXT ANSI)</b> e importe no
        módulo fiscal do <b>Domínio Sistemas</b>.</p>

        <hr>

        <h4>⚠ Observações</h4>
        <ul>
            <li>Saída em <b>ANSI (Latin-1)</b>, padrão Domínio Sistemas.</li>
            <li>Registros gerados: <b>0000, 0020, 0100, 1000, 1020, 1030, 1150, 1151</b>.</li>
            <li>O registro <b>1020</b> é gerado um por imposto presente na nota
                (tabela oficial Domínio).</li>
            <li>Impostos mapeados: <b>1-ICMS, 2-IPI, 4-PIS, 5-COFINS,
                45-ICMS Importação, 133-PIS Importação, 134-COFINS Importação,
                183-IBS, 184-CBS</b>.</li>
        </ul>

        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("---")

# ─────────────────────────────────────────────
# UPLOAD
# ─────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "📂 Selecione um ou mais arquivos XML de NF-e",
    type=["xml"],
    accept_multiple_files=True,
    help="XML de NF-e (nfeProc ou NFe) — modelo 55",
)

# ─────────────────────────────────────────────
# PROCESSAMENTO
# ─────────────────────────────────────────────
if uploaded_files:

    all_lines   = []
    all_resumos = []
    erros       = []

    progress = st.progress(0, text="Processando arquivos...")

    for i, f in enumerate(uploaded_files):
        xml_bytes = f.read()
        texto, resumo = converter_xml(
            xml_bytes,
            cnpj_fallback=cnpj_fallback,
            acumulador=acumulador,
            especie=especie,
            incluir_0000=inc_0000,
            incluir_0020=inc_0020,
            incluir_0100=inc_0100,
        )
        if "erro" in resumo:
            erros.append({"Arquivo": f.name, "Erro": resumo["erro"]})
        else:
            all_lines.append(texto)
            all_resumos.append({"Arquivo": f.name, **resumo})

        progress.progress(
            (i + 1) / len(uploaded_files),
            text=f"Processando {f.name}..."
        )

    progress.empty()

    # Erros
    if erros:
        st.error("❌ Erros encontrados:")
        st.dataframe(erros, use_container_width=True)

    # Sucesso
    if all_resumos:
        st.success(f"✅ {len(all_resumos)} arquivo(s) convertido(s) com sucesso!")

        # ── Destaque do CNPJ lido do XML ─────────────────────────────
        cnpjs_unicos = list({r["CNPJ Dest"]: r for r in all_resumos}.values())
        if cnpjs_unicos:
            st.markdown("#### 🏢 Empresa Destinatária (lida do XML)")
            cols = st.columns(len(cnpjs_unicos[:4]))
            for idx, r in enumerate(cnpjs_unicos[:4]):
                with cols[idx]:
                    st.markdown(
                        f"""
                        <div class="cnpj-badge">
                            CNPJ: {r['CNPJ Dest']}
                        </div>
                        <div class="info-origem">
                            📌 {r['Origem CNPJ']}<br>
                            🏷️ {r['Destinatário'][:60] if r['Destinatário'] else '—'}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

        st.markdown("---")

        with st.expander("📊 Resumo das notas processadas", expanded=True):
            st.dataframe(all_resumos, use_container_width=True)

        saida_final = "\n".join(all_lines)

        with st.expander("👁️ Preview do arquivo gerado (primeiras 80 linhas)"):
            preview = saida_final.split("\n")[:80]
            st.code("\n".join(preview), language="text")

        st.markdown("---")

        saida_ansi = encode_ansi(saida_final)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="⬇ Baixar Arquivo Domínio (.TXT ANSI)",
                data=saida_ansi,
                file_name="importacao_dominio.txt",
                mime="text/plain",
                use_container_width=True,
                type="primary",
            )
        with col2:
            st.download_button(
                label="⬇ Baixar Arquivo (.TXT UTF-8)",
                data=saida_final.encode("utf-8"),
                file_name="importacao_dominio_utf8.txt",
                mime="text/plain",
                use_container_width=True,
            )

        # Métricas
        st.markdown("---")
        st.markdown("#### 📈 Estatísticas")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Notas", len(all_resumos))
        c2.metric("Itens", sum(r.get("Itens", 0) for r in all_resumos))
        total_linhas = len([l for l in saida_final.split("\n") if l.startswith("|")])
        c3.metric("Linhas geradas", total_linhas)
        c4.metric("Erros", len(erros))
        try:
            total_nf = sum(
                float(r.get("vNF", "0").replace(",", "."))
                for r in all_resumos if r.get("vNF")
            )
            c5.metric("Total NF (R$)",
                      f"{total_nf:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        except Exception:
            c5.metric("Total NF", "—")

else:
    st.info("👆 Faça o upload de um ou mais arquivos XML de NF-e para iniciar a conversão.")

# ─────────────────────────────────────────────
# RODAPÉ
# ─────────────────────────────────────────────
st.markdown("---")
st.caption(
    f"Conversor XML NF-e → Domínio Sistemas | Thomson Reuters | "
    f"Python + Streamlit | {VERSAO}"
)
