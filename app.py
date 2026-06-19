import streamlit as st
import xml.etree.ElementTree as ET
from datetime import datetime

# ─────────────────────────────────────────────
# VERSÃO
# ─────────────────────────────────────────────
VERSAO = "V1.3"

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
        h1, h2, h3 { color: #FF8000; font-weight: 700; }
        section[data-testid="stSidebar"] {
            background-color: #444444; color: #FFFFFF;
        }
        section[data-testid="stSidebar"] * { color: #FFFFFF !important; }
        .stButton > button {
            background-color: #FF8000; color: #FFFFFF;
            border: none; border-radius: 4px; font-weight: bold;
        }
        .stButton > button:hover { background-color: #D64001; }
        .stDownloadButton > button {
            background-color: #FF8000; color: #FFFFFF;
            border: none; border-radius: 4px; font-weight: bold;
        }
        .stDownloadButton > button:hover { background-color: #D64001; }
        hr { border-color: #FF8000; }
        [data-testid="metric-container"] {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 10px;
        }
        .instrucoes-box {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 16px 20px;
            margin: 12px 0; color: #444444;
        }
        .instrucoes-box h4 {
            color: #FF8000; margin-top: 14px; margin-bottom: 6px;
        }
        .instrucoes-box h4:first-child { margin-top: 0; }
        .cnpj-badge {
            background-color: #444444;
            border: 1px solid #FF8000; border-radius: 4px;
            padding: 6px 12px; font-family: Consolas, monospace;
            font-size: 13px; display: inline-block; margin: 4px 0;
        }
        .info-origem {
            background-color: #FFF3E0;
            border-left: 3px solid #FF8000;
            border-radius: 3px; padding: 8px 12px;
            font-size: 12px; color: #444444; margin: 6px 0;
        }
        </style>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Dominio Sistemas | Thomson Reuters",
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
    <div style="background:#444444; padding:24px 28px 18px 28px;
                border-radius:8px; border-top:6px solid #FF8000;
                margin-bottom:28px;">
        <h2 style="color:#FF8000; margin:0;
                   font-family:'Segoe UI',Arial,sans-serif;">
            🔄 Conversor XML NF-e → Dominio Sistemas
            &nbsp;|&nbsp; {VERSAO}
        </h2>
        <p style="color:#DDDDDD; margin:6px 0 0 0;
                  font-family:'Segoe UI',Arial,sans-serif;">
            Converte XML de NF-e para leiaute padrao de importacao do
            <strong>Dominio Sistemas</strong>.
            Saida em <strong>ANSI (Latin-1)</strong>.
            CNPJ lido automaticamente do XML.
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
# TABELA DE GRUPOS – Grupos.xls
# ─────────────────────────────────────────────
TABELA_GRUPOS = {
    0:   "Automatico (por CFOP/NCM)",
    1:   "GERAL",
    2:   "MERCADORIA PARA REVENDA",
    3:   "MATERIA PRIMA",
    4:   "EMBALAGENS",
    5:   "PRODUTO EM PROCESSO",
    6:   "PRODUTO ACABADO",
    7:   "SUBPRODUTO",
    8:   "PRODUTOS INTERMEDIARIOS",
    9:   "MATERIAL DE USO E CONSUMO",
    10:  "ATIVO IMOBILIZADO",
    11:  "SERVICOS",
    12:  "OUTROS INSUMOS",
    100: "PRODUTOS NOVOS - ENTRADAS",
    500: "PRODUTOS NOVOS - SAIDAS",
}

# ─────────────────────────────────────────────
# TABELA DE IMPOSTOS – Impostos.xls
# ─────────────────────────────────────────────
TABELA_IMPOSTOS = {
    1:"ICMS", 2:"IPI", 3:"ISS", 4:"PIS", 5:"COFINS",
    6:"CONTRIBUICAO SOCIAL", 7:"IRPJ-LP", 8:"DIFALI",
    9:"SUBST. TRIBUTARIA", 10:"SIMPLES", 11:"ICMS RETIDO",
    12:"ICMS SUBSTITUTO (RS)", 14:"REFIS", 15:"FIA", 16:"IRRF",
    17:"PIS(nao cumulativo)", 18:"ISS Retido",
    19:"COFINS(nao cumulativa)", 20:"Fundo de Fomento e Desenvolvimento",
    21:"Recomposicao de Aliquota", 22:"PIS(Retido)", 23:"COFINS(Retido)",
    24:"Contribuicao Social(Retida)", 25:"Contribuicao Retida na Fonte",
    26:"INSS Retido", 27:"ICMS Antecipado(Atacadista)", 28:"FUNRURAL",
    29:"Contr p/ Prog. Inc Arrec Educ tributaria", 30:"IPI-M",
    31:"ICMS ST Antecipacao Total", 32:"FUNDOSOCIAL",
    33:"IRPJ postergado", 34:"ICMS Antecipado(Farmacia)",
    35:"FUNCULTURAL", 36:"FUNTURISMO", 37:"FUNDESPORTE",
    38:"Contribuicoes Retidas na Fonte - Orgaos Publicos",
    39:"IRRF Propaganda", 40:"INSS COOPERATIVAS",
    41:"Fundo para o Desenv. Tec. das Telec.", 42:"ICMS Fato Gerador",
    43:"ICMS Adicional(2%)", 44:"SIMPLES NACIONAL",
    45:"ICMS Importacao", 46:"ICMS Antecipado Combustiveis",
    47:"Subst. Trib. Antecipada Combustiveis",
    48:"ICMS Garantido Integral",
    49:"FECOEP-FUNDO EST. DE COMB E ERRAD. POBRE",
    50:"Fundo de Transportes e Habitacao- FETHAB",
    55:"FECOP-ICMS Normal", 56:"FECOP-ICMS ST OP.Internas",
    57:"Fundo de Combate a Pobreza ST - Operacoes Interestaduais",
    58:"Fundo Pro-Emprego",
    59:"Apoio a Manut. e Desenv. Educ. Superior",
    60:"Fundo Protecao Social do estado de Goias", 61:"FUNDAF",
    62:"Fundo Geracao de Emprego e Renda do DF",
    63:"IRRF alugueis pessoa fisica", 64:"Simples MEI",
    65:"TX ADIC. FOMENTAR/MICROPRODUZIR/PRODUZIR",
    66:"Fundo Universal dos Servicos de Telecom.",
    67:"PIS - Substituicao Tributaria",
    68:"COFINS - Substituicao Tributaria",
    69:"ICMS Subst.Tributaria Serv. Transportes",
    70:"Comp. Financeira pela Expl. Rec.Minerais",
    82:"FECP - ICMS Importacao",
    83:"RET - Regime especial incorporacoes imobiliarias",
    84:"PIS - Regime especial incorporacoes imobiliarias",
    85:"COFINS - Regime especial incorporacoes imobiliarias",
    86:"IRPJ - Regime especial incorporacoes imobiliarias",
    87:"CSLL - Regime especial incorporacoes imobiliarias",
    88:"Regime especial PMCMV", 89:"PIS Regime especial PMCMV",
    90:"COFINS Regime especial PMCMV",
    91:"IRPJ - Regime especial PMCMV",
    92:"CSLL - Regime especial PMCMV",
    93:"ICMS Antecipado - Farinha de Trigo",
    94:"FECOP - ICMS Antecipado", 95:"FUNDESA",
    96:"FEM - ICMS Normal", 97:"FEM - ICMS Substituicao Tributaria",
    98:"ICMS Antecipado Servicos de Transportes",
    99:"Fundo Estadual de Defesa Civil - FUNDEC",
    100:"FUNDEIC", 101:"FUNDED", 102:"FUPIS",
    103:"Contribuicao Previdenciaria sobre a Receita Bruta",
    104:"PIS Entidades Financeiras e Equiparadas",
    105:"COFINS Entidades Financeiras e Equiparadas",
    106:"PIS Nao cumulativo - SCP",
    107:"COFINS Nao cumulativo - SCP",
    108:"PIS - SCP", 109:"COFINS - SCP",
    110:"Contribuicao Social - SCP", 111:"IRPJ - SCP",
    112:"Fundo de Combate - ICMS Antecipado",
    113:"Fundo de Combate - ICMS NORMAL",
    114:"Fundo de Combate - ICMS ST",
    115:"Fundo Desenvolvimento Economico-FDE",
    116:"ICMS DIFERIDO", 117:"CIDE", 118:"FAPESC",
    119:"FITUR", 120:"PROUNIV",
    121:"ICMS Antecipado Telecomunicacoes", 122:"FAI",
    125:"ICMS Complementar", 126:"FUNDEPEC", 127:"FUNPRODUZIR",
    128:"ICMS Carga Media", 129:"ICMS ST Carga Media",
    130:"FUNDAP", 132:"INSS Retido - SCP",
    133:"PIS Importacao", 134:"COFINS Importacao",
    135:"Fundo Estadual de Saude - FES",
    136:"Programa Bolsa Universitaria",
    138:"PIS - Codigo de recolhimento",
    139:"COFINS - Codigo de recolhimento",
    140:"PIS Nao cumulativo - Codigo de recolhimento",
    141:"COFINS Nao cumulativa - Codigo de recolhimento",
    142:"CPRB - SCP",
    145:"ICMS Diferencial de Aliquota Antecipado",
    146:"Fundo Estadual de Combate a Pobreza",
    147:"FECOEP-SP - ICMS Proprio", 148:"FECOEP-SP - ICMS ST",
    149:"FECOEP-TO - ICMS Proprio", 150:"FECOEP-TO - ICMS ST",
    152:"FECOP-PR - ICMS Proprio", 153:"FECOP-PR - ICMS ST",
    154:"FEEF", 155:"FECOMP-MS - ICMS Proprio",
    156:"FECOMP-MS - ICMS ST", 157:"FECOEP-RO - ICMS Proprio",
    158:"FECOEP-RO - ICMS ST", 159:"FECEP-MT - ICMS ES",
    160:"FECEP-MT - ICMS ST", 161:"IOF",
    162:"Fundo Desenv. Sistema Rodoviario MS",
    163:"Fundo Milho e Soja MS",
    164:"ICMS Equalizacao Simples Nacional",
    165:"FIDER-RO", 166:"PROLEITE", 167:"FITHA-RO",
    168:"RET - SCP", 169:"PMCMV - SCP", 170:"FUNTUR",
    172:"FUNEF", 177:"Fundo Estadual - Incentivos",
    178:"FDI", 179:"Programa Mais IDH", 180:"FUNDEINFRA",
    183:"IBS - Imposto sobre Bens e Servicos",
    184:"CBS - Contribuicao Social sobre Bens e Servicos",
    187:"IS - Imposto Seletivo",
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
# DETECÇÃO AUTOMÁTICA DE GRUPO
# Por CFOP (primário) e NCM (fallback)
# ─────────────────────────────────────────────
def get_grupo_por_cfop(cfop: str) -> int:
    if not cfop:
        return 1
    dois = cfop[:2] if len(cfop) >= 2 else cfop
    mapa = {
        "11": 2, "12": 2, "13": 3, "14": 3,
        "15": 9, "16": 10, "17": 11,
        "20": 2, "21": 2, "22": 3, "25": 9,
        "30": 2, "31": 3, "35": 9,
        "40": 2, "41": 2, "55": 12, "60": 2,
    }
    return mapa.get(dois, 1)


def get_grupo_por_ncm(ncm: str) -> int:
    if not ncm or len(ncm) < 2:
        return 1
    cap = ncm[:2]
    mapa = {
        "84": 10, "85": 10, "86": 10, "87": 10,
        "88": 10, "89": 10, "90": 10, "91": 10,
        "28": 3, "29": 3, "30": 3, "31": 3,
        "32": 3, "33": 3, "34": 3, "38": 3,
        "39": 3, "40": 3, "44": 3, "47": 3,
        "48": 3, "72": 3, "73": 3, "74": 3,
        "75": 3, "76": 3, "82": 3, "83": 3,
        "94": 10,
        "01": 2, "02": 2, "03": 2, "04": 2,
        "07": 2, "08": 2, "09": 2, "10": 2,
        "16": 2, "17": 2, "18": 2, "19": 2,
        "20": 2, "21": 2, "22": 2, "27": 2,
    }
    return mapa.get(cap, 1)


def detectar_grupo(cfop: str, ncm: str, grupo_padrao: int) -> int:
    if grupo_padrao > 0:
        return grupo_padrao
    g = get_grupo_por_cfop(cfop)
    if g == 1 and ncm:
        g2 = get_grupo_por_ncm(ncm)
        return g2 if g2 != 1 else g
    return g


# ─────────────────────────────────────────────
# EXTRAÇÃO DO CNPJ DESTINATÁRIO
# Prioridade: dest/CNPJ → dest/CPF →
#   dest/idEstrangeiro (com valor) →
#   emit/CNPJ (UF=EX ou idEstrangeiro vazio) →
#   fallback manual
# ─────────────────────────────────────────────
def extrair_cnpj_dest(nfe_root) -> tuple:
    dest = nfe_root.find("nfe:infNFe/nfe:dest", NS)
    if dest is not None:
        cnpj = get_text(dest, "nfe:CNPJ")
        if cnpj:
            return cnpj, "XML — <dest><CNPJ>", False
        cpf = get_text(dest, "nfe:CPF")
        if cpf:
            return cpf, "XML — <dest><CPF>", False
        id_ext = get_text(dest, "nfe:idEstrangeiro")
        if id_ext and id_ext.strip():
            return id_ext.strip(), "XML — <dest><idEstrangeiro>", True
        id_ext_node = dest.find("nfe:idEstrangeiro", NS)
        ender_dest  = dest.find("nfe:enderDest", NS)
        uf_dest     = get_text(ender_dest, "nfe:UF") if ender_dest is not None else ""
        if id_ext_node is not None or uf_dest == "EX":
            emit = nfe_root.find("nfe:infNFe/nfe:emit", NS)
            cnpj_emit = get_text(emit, "nfe:CNPJ") if emit is not None else ""
            if cnpj_emit:
                return cnpj_emit, "XML — <emit><CNPJ> (dest. exterior/UF=EX)", True
            return "", "Destinatario exterior sem CNPJ", True
    return "", "Nao encontrado no XML", False


def extrair_nome_dest(nfe_root) -> str:
    dest = nfe_root.find("nfe:infNFe/nfe:dest", NS)
    return get_text(dest, "nfe:xNome") if dest is not None else ""


def extrair_uf_dest(nfe_root) -> str:
    dest = nfe_root.find("nfe:infNFe/nfe:dest", NS)
    if dest is not None:
        ender = dest.find("nfe:enderDest", NS)
        return get_text(ender, "nfe:UF") if ender is not None else ""
    return ""


# ─────────────────────────────────────────────
# REGISTRO 0000 (2 campos)
# ─────────────────────────────────────────────
def gerar_registro_0000(cnpj_empresa: str) -> str:
    return pipe_join(["0000", cnpj_empresa])


# ─────────────────────────────────────────────
# REGISTRO 0020 – Fornecedor (33 campos)
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
    regime_map  = {"1": "M", "2": "E", "3": "N"}
    regime      = regime_map.get(crt, "N")
    contrib     = "S" if ie and ie.upper() not in ("ISENTO", "NAO CONTRIBUINTE", "") else "N"
    return pipe_join([
        "0020", cnpj, razao, fantasia, logradouro, numero, complemento,
        bairro, cod_mun, uf, "", cep, ie, "", "", "", "", "", "", "", "",
        "N", "7", regime, contrib, "", "", "", "", "N", "N", "", "",
    ])


# ─────────────────────────────────────────────
# REGISTRO 0100 – Produtos (91 campos)
# Campo 9 = Grupo de produtos (Grupos.xls)
# ─────────────────────────────────────────────
def gerar_registro_0100(det, grupo_padrao: int = 0) -> str:
    prod      = det.find("nfe:prod", NS)
    cod_prod  = get_text(prod, "nfe:cProd")[:14]
    descricao = get_text(prod, "nfe:xProd")
    ncm       = get_text(prod, "nfe:NCM")
    unidade   = get_text(prod, "nfe:uCom")
    val_unit  = get_text(prod, "nfe:vUnCom")
    cest      = get_text(prod, "nfe:CEST")
    cfop      = get_text(prod, "nfe:CFOP")
    cod_grupo = detectar_grupo(cfop, ncm, grupo_padrao)

    imposto   = det.find("nfe:imposto", NS)
    cst_icms  = ""
    aliq_icms = ""
    aliq_ipi  = ""
    if imposto is not None:
        for tp in ["ICMS00","ICMS10","ICMS20","ICMS30","ICMS40",
                   "ICMS51","ICMS60","ICMS70","ICMS90","ICMSSN101",
                   "ICMSSN102","ICMSSN201","ICMSSN202","ICMSSN500","ICMSSN900"]:
            node = imposto.find(f"nfe:ICMS/nfe:{tp}", NS)
            if node is not None:
                cst_icms  = get_text(node, "nfe:CST") or get_text(node, "nfe:CSOSN")
                aliq_icms = fmt_decimal(get_text(node, "nfe:pICMS"))
                break
        ipi_trib = imposto.find("nfe:IPI/nfe:IPITrib", NS)
        if ipi_trib is not None:
            aliq_ipi = fmt_decimal(get_text(ipi_trib, "nfe:pIPI"))

    return pipe_join([
        "0100", cod_prod, descricao, "", ncm, "", "", "",
        cod_grupo,                    # 9 – Grupo de produtos
        unidade, "N", "O", "", "", "", "N", "",
        fmt_decimal(val_unit, 3), "", "", cst_icms, aliq_icms, aliq_ipi,
        "", "", "N", "", "", "", "", "", "", "", "", "N", "", "", "",
        "N", "", "", "", "N", "", "", "", "N", "", "", "", "", "", "",
        "N", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
        "", "", "", "N", "", "", "N", "", "", "", "", "", "", "", "",
        "", "", cest, "", "",
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

    inf_nfe = nfe_root.find("nfe:infNFe", NS)
    chave   = ""
    if inf_nfe is not None:
        id_attr = inf_nfe.get("Id", "")
        chave   = id_attr.replace("NFe", "") if id_attr else ""

    transp        = nfe_root.find("nfe:infNFe/nfe:transp", NS)
    mod_frete_cod = get_text(transp, "nfe:modFrete")
    frete_map     = {"0":"C","1":"F","2":"S","3":"T","4":"R","5":"D","9":"S"}
    mod_frete     = frete_map.get(mod_frete_cod, "C")

    # Campo 15 — Observação de interesse do fisco (infAdFisco)
    inf_adic  = nfe_root.find("nfe:infNFe/nfe:infAdic", NS)
    obs_fisco = ""
    if inf_adic is not None:
        obs_fisco = get_text(inf_adic, "nfe:infAdFisco")[:300]

    # Campo 52 — Número DI do primeiro item
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
# REGISTRO 1010 – Informacoes Complementares
# Filho do 1000 | 3 campos
# Origem: infAdFisco (cod=1) e infCpl (cod=2)
# ─────────────────────────────────────────────
def gerar_registros_1010(nfe_root) -> list:
    linhas   = []
    inf_adic = nfe_root.find("nfe:infNFe/nfe:infAdic", NS)
    if inf_adic is None:
        return linhas
    # Codigo 1 = Interesse do Fisco
    obs_fisco = get_text(inf_adic, "nfe:infAdFisco")
    if obs_fisco:
        for bloco in [obs_fisco[i:i+300] for i in range(0, len(obs_fisco), 300)]:
            linhas.append(pipe_join(["1010", "1", bloco]))
    # Codigo 2 = Interesse do Contribuinte
    obs_cpl = get_text(inf_adic, "nfe:infCpl")
    if obs_cpl:
        for bloco in [obs_cpl[i:i+300] for i in range(0, len(obs_cpl), 300)]:
            linhas.append(pipe_join(["1010", "2", bloco]))
    return linhas


# ─────────────────────────────────────────────
# REGISTRO 1015 – Observacoes
# Filho do 1000 | 3 campos
# Espelha o 1010 com mesmo conteudo
# ─────────────────────────────────────────────
def gerar_registros_1015(nfe_root) -> list:
    linhas   = []
    inf_adic = nfe_root.find("nfe:infNFe/nfe:infAdic", NS)
    if inf_adic is None:
        return linhas
    obs_fisco = get_text(inf_adic, "nfe:infAdFisco")
    if obs_fisco:
        for bloco in [obs_fisco[i:i+300] for i in range(0, len(obs_fisco), 300)]:
            linhas.append(pipe_join(["1015", "1", bloco]))
    obs_cpl = get_text(inf_adic, "nfe:infCpl")
    if obs_cpl:
        for bloco in [obs_cpl[i:i+300] for i in range(0, len(obs_cpl), 300)]:
            linhas.append(pipe_join(["1015", "2", bloco]))
    return linhas


# ─────────────────────────────────────────────
# REGISTRO 1020 – Impostos (19 campos cada)
# Um registro por imposto — tabela Impostos.xls
# ─────────────────────────────────────────────
def gerar_registros_1020(nfe_root) -> list:
    total  = nfe_root.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)
    ibs_t  = nfe_root.find("nfe:infNFe/nfe:total/nfe:IBSCBSTot", NS)
    v_nf   = fmt_decimal(get_text(total, "nfe:vNF"))
    linhas = []

    def r1020(cod, perc_red="", base="", aliq="", valor="",
              isentas="", outras="", v_ipi="", v_st="",
              v_cont="", cod_rec=""):
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
            base=fmt_decimal(v_bc_icms), aliq=aliq_icms_med,
            valor=fmt_decimal(v_icms), v_ipi=fmt_decimal(v_ipi_tot),
            v_st=fmt_decimal(v_st_tot), v_cont=v_nf))
    # 2 – IPI
    if v_ipi_tot and float(v_ipi_tot) > 0:
        linhas.append(r1020(2,
            base=fmt_decimal(str(bc_ipi_total)),
            valor=fmt_decimal(v_ipi_tot), v_cont=v_nf))
    # 4 – PIS
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(r1020(4,
            base=fmt_decimal(str(bc_pis_total)),
            valor=fmt_decimal(v_pis_tot), v_cont=v_nf))
    # 5 – COFINS
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(r1020(5,
            base=fmt_decimal(str(bc_cof_total)),
            valor=fmt_decimal(v_cofins_tot), v_cont=v_nf))
    # 45 – ICMS Importacao
    if v_icms_deson and float(v_icms_deson) > 0:
        linhas.append(r1020(45,
            base=fmt_decimal(v_bc_icms), aliq=aliq_icms_med,
            valor=fmt_decimal(v_icms_deson), v_cont=v_nf))
    # 133 – PIS Importacao
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(r1020(133,
            base=fmt_decimal(str(bc_pis_total)),
            valor=fmt_decimal(v_pis_tot), v_cont=v_nf))
    # 134 – COFINS Importacao
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(r1020(134,
            base=fmt_decimal(str(bc_cof_total)),
            valor=fmt_decimal(v_cofins_tot), v_cont=v_nf))
    # 183 – IBS
    if ibs_t is not None:
        v_ibs_uf = get_text(ibs_t, "nfe:gIBS/nfe:gIBSUF/nfe:vIBSUF")
        bc_ibs   = get_text(ibs_t, "nfe:vBCIBSCBS")
        if v_ibs_uf and float(v_ibs_uf) > 0:
            linhas.append(r1020(183,
                base=fmt_decimal(bc_ibs),
                valor=fmt_decimal(v_ibs_uf), v_cont=v_nf))
    # 184 – CBS
    if ibs_t is not None:
        v_cbs  = get_text(ibs_t, "nfe:gCBS/nfe:vCBS")
        bc_cbs = get_text(ibs_t, "nfe:vBCIBSCBS")
        if v_cbs and float(v_cbs) > 0:
            linhas.append(r1020(184,
                base=fmt_decimal(bc_cbs),
                valor=fmt_decimal(v_cbs), v_cont=v_nf))
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
        for tp in ["ICMS00","ICMS10","ICMS20","ICMS30","ICMS40",
                   "ICMS51","ICMS60","ICMS70","ICMS90","ICMSSN101",
                   "ICMSSN102","ICMSSN201","ICMSSN202","ICMSSN500","ICMSSN900"]:
            node = imposto.find(f"nfe:ICMS/nfe:{tp}", NS)
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
        v_bc_icms, v_bc_st, aliq_icms, "", "", "", "",
        fmt_decimal(v_outro), "", v_icms, "", "", "", "",
        fmt_decimal(v_unit, 6), "", cst_ipi, aliq_ipi, "", "", "",
        cfop, "", aliq_pis, v_pis, aliq_cof, v_cof,
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
# REGISTRO 1097 – Dados do Frete SP (35 campos)
# Filho do 1000
# ─────────────────────────────────────────────
def gerar_registro_1097(nfe_root) -> str:
    transp = nfe_root.find("nfe:infNFe/nfe:transp", NS)
    if transp is None:
        return ""

    mod_frete_cod = get_text(transp, "nfe:modFrete")
    frete_map     = {"0":"C","1":"F","2":"S","3":"T","4":"R","5":"D","9":"S"}
    mod_frete     = frete_map.get(mod_frete_cod, "C")

    # Modalidade de transporte (tpViaTransp do primeiro DI)
    det_list = nfe_root.findall("nfe:infNFe/nfe:det", NS)
    tp_via   = ""
    if det_list:
        di_node = det_list[0].find("nfe:prod/nfe:DI", NS)
        if di_node is not None:
            tp_via = get_text(di_node, "nfe:tpViaTransp")

    # Transportadora
    transporta   = transp.find("nfe:transporta", NS)
    cnpj_transp  = get_text(transporta, "nfe:CNPJ")   if transporta is not None else ""
    razao_transp = get_text(transporta, "nfe:xNome")  if transporta is not None else ""
    ie_transp    = get_text(transporta, "nfe:IE")     if transporta is not None else ""
    end_transp   = get_text(transporta, "nfe:xEnder") if transporta is not None else ""
    mun_transp   = get_text(transporta, "nfe:xMun")   if transporta is not None else ""
    uf_transp    = get_text(transporta, "nfe:UF")     if transporta is not None else ""
    tipo_insc    = "1" if cnpj_transp else ""

    # Volumes
    vol     = transp.find("nfe:vol", NS)
    q_vol   = get_text(vol, "nfe:qVol")   if vol is not None else ""
    esp_vol = get_text(vol, "nfe:esp")    if vol is not None else ""
    marca   = get_text(vol, "nfe:marca")  if vol is not None else ""
    peso_l  = fmt_decimal(get_text(vol, "nfe:pesoL"), 3) if vol is not None else ""
    peso_b  = fmt_decimal(get_text(vol, "nfe:pesoB"), 3) if vol is not None else ""

    return pipe_join([
        "1097",         # 1  - Identificacao
        mod_frete,      # 2  - Modalidade do frete (C/F/S/T/R/D)
        tp_via,         # 3  - Modalidade do transporte (tpViaTransp)
        "",             # 4  - Frete por conta do (D/E/O)
        "",             # 5  - Placa 1
        "",             # 6  - UF Placa 1
        "",             # 7  - Placa 2
        "",             # 8  - UF Placa 2
        "",             # 9  - Placa 3
        "",             # 10 - UF Placa 3
        razao_transp[:150] if razao_transp else "",  # 11 - Razao Social
        tipo_insc,      # 12 - Tipo da Inscricao (1=CNPJ)
        cnpj_transp,    # 13 - CNPJ/CPF Transportador
        ie_transp,      # 14 - Inscricao Estadual
        end_transp,     # 15 - Endereco
        "",             # 16 - Numero
        "",             # 17 - Bairro
        "",             # 18 - Complemento
        mun_transp,     # 19 - Cidade
        uf_transp,      # 20 - UF
        "",             # 21 - CEP
        q_vol,          # 22 - Qtd volumes transportados
        esp_vol,        # 23 - Especie
        marca,          # 24 - Marca
        "",             # 25 - Numeracao
        peso_l,         # 26 - Peso Liquido (3 casas)
        peso_b,         # 27 - Peso Bruto (3 casas)
        "E",            # 28 - Tipo Local saida (E=Emitente)
        "",             # 29 - CNPJ Local saida
        "",             # 30 - UF Local saida
        "",             # 31 - IE Local saida
        "D",            # 32 - Tipo Local recebimento (D=Destinatario)
        "",             # 33 - CNPJ Local recebimento
        "",             # 34 - UF Local recebimento
        "",             # 35 - IE Local recebimento
    ])


# ─────────────────────────────────────────────
# REGISTROS 1150 / 1151 – IBS / CBS (5 campos)
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
    incluir_1010: bool = True,
    incluir_1015: bool = False,
    incluir_1097: bool = True,
    grupo_padrao: int = 0,
) -> tuple:

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        return "", {"erro": str(e)}

    nfe = root.find("nfe:NFe", NS)
    if nfe is None:
        nfe = root

    cnpj_xml, origem_cnpj, is_exterior = extrair_cnpj_dest(nfe)
    nome_dest = extrair_nome_dest(nfe)
    uf_dest   = extrair_uf_dest(nfe)

    if cnpj_xml:
        cnpj_empresa = cnpj_xml
    elif cnpj_fallback:
        cnpj_empresa = cnpj_fallback
        origem_cnpj  = "Manual (fallback)"
        is_exterior  = False
    else:
        return "", {"erro": "CNPJ nao encontrado no XML nem informado manualmente."}

    lines    = []
    resumo   = {}
    det_list = nfe.findall("nfe:infNFe/nfe:det", NS)
    emit     = nfe.find("nfe:infNFe/nfe:emit", NS)
    ide      = nfe.find("nfe:infNFe/nfe:ide", NS)
    total    = nfe.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)

    resumo["nNF"]          = get_text(ide, "nfe:nNF")
    resumo["Emitente"]     = get_text(emit, "nfe:xNome")
    resumo["CNPJ Emit"]    = get_text(emit, "nfe:CNPJ")
    resumo["Destinatario"] = nome_dest
    resumo["UF Dest"]      = uf_dest
    resumo["CNPJ Empresa"] = cnpj_empresa
    resumo["Origem CNPJ"]  = origem_cnpj
    resumo["Exterior"]     = "Sim" if is_exterior else "Nao"
    resumo["Emissao"]      = fmt_date(get_text(ide, "nfe:dhEmi"))
    resumo["Itens"]        = len(det_list)
    resumo["vNF"]          = fmt_decimal(get_text(total, "nfe:vNF"))
    resumo["vICMS"]        = fmt_decimal(get_text(total, "nfe:vICMS"))
    resumo["vICMSDes"]     = fmt_decimal(get_text(total, "nfe:vICMSDeson"))
    resumo["vIPI"]         = fmt_decimal(get_text(total, "nfe:vIPI"))
    resumo["vPIS"]         = fmt_decimal(get_text(total, "nfe:vPIS"))
    resumo["vCOFINS"]      = fmt_decimal(get_text(total, "nfe:vCOFINS"))
    resumo["vII"]          = fmt_decimal(get_text(total, "nfe:vII"))
    resumo["Grupo"]        = (
        f"{grupo_padrao} - {TABELA_GRUPOS.get(grupo_padrao,'GERAL')}"
        if grupo_padrao > 0 else "Auto (CFOP/NCM)"
    )

    # Cadastros
    if incluir_0000:
        lines.append(gerar_registro_0000(cnpj_empresa))
    if incluir_0020 and emit is not None:
        lines.append(gerar_registro_0020(emit))
    if incluir_0100:
        produtos_gerados = set()
        for det in det_list:
            cod = get_text(det.find("nfe:prod", NS), "nfe:cProd")
            if cod not in produtos_gerados:
                lines.append(gerar_registro_0100(det, grupo_padrao=grupo_padrao))
                produtos_gerados.add(cod)

    # Cabecalho da nota
    lines.append(gerar_registro_1000(nfe, cnpj_empresa, acumulador, especie))

    # Informacoes complementares
    if incluir_1010:
        for r in gerar_registros_1010(nfe):
            lines.append(r)

    # Observacoes
    if incluir_1015:
        for r in gerar_registros_1015(nfe):
            lines.append(r)

    # Impostos
    for r in gerar_registros_1020(nfe):
        lines.append(r)

    # Itens
    for seq, det in enumerate(det_list, start=1):
        lines.append(gerar_registro_1030(det, seq))

    # Frete SP
    if incluir_1097:
        r1097 = gerar_registro_1097(nfe)
        if r1097:
            lines.append(r1097)

    # IBS/CBS agrupados por cClassTrib
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
            ibs_gerados[ct] = {
                "bc_ibs": 0.0, "v_ibs": 0.0, "aliq_ibs": "",
                "bc_cbs": 0.0, "v_cbs": 0.0, "aliq_cbs": "",
            }
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
        lines.append(gerar_registro_1150(
            ct, fmt_decimal(str(d["bc_ibs"])),
            fmt_decimal(d["aliq_ibs"]), fmt_decimal(str(d["v_ibs"]))))
        lines.append(gerar_registro_1151(
            ct, fmt_decimal(str(d["bc_cbs"])),
            fmt_decimal(d["aliq_cbs"]), fmt_decimal(str(d["v_cbs"]))))

    return "\n".join(lines), resumo


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"### Info")
    st.markdown(f"**Versao:** {VERSAO}")
    st.markdown("**Thomson Reuters**")
    st.markdown("**Dominio Sistemas**")
    st.markdown("---")

    st.markdown("### Parametros")
    st.markdown(
        "<div style='font-size:12px;color:#DDDDDD;margin-bottom:4px;'>"
        "CNPJ lido automaticamente do XML.<br>"
        "Use o campo abaixo apenas como <b>fallback</b>.</div>",
        unsafe_allow_html=True,
    )
    cnpj_fallback = st.text_input(
        "CNPJ Fallback (opcional)", value="", max_chars=14,
        help="Usado somente se o XML nao contiver CNPJ do destinatario.")
    acumulador = st.text_input("Codigo do Acumulador", value="1157",
                               help="1157 = CFOP 3102 (Importacao)")
    especie    = st.text_input("Codigo da Especie", value="36",
                               help="36 = NF-e modelo 55")

    st.markdown("---")
    st.markdown("### Registros de cadastro")
    inc_0000 = st.checkbox("0000 - Identificacao da empresa", value=True)
    inc_0020 = st.checkbox("0020 - Fornecedor (emitente)", value=True)
    inc_0100 = st.checkbox("0100 - Produtos", value=True)

    st.markdown("---")
    st.markdown("### Registros filhos do 1000")
    inc_1010 = st.checkbox("1010 - Inf. Complementares (fisco/contrib.)", value=True)
    inc_1015 = st.checkbox("1015 - Observacoes", value=False)
    inc_1097 = st.checkbox("1097 - Dados do Frete (SP)", value=True)
    st.caption("1020 / 1030 / 1150 / 1151 sempre gerados.")

    st.markdown("---")
    st.markdown("### Grupo de Produtos (0100 campo 9)")
    grupo_selecionado = st.selectbox(
        "Grupo",
        options=list(TABELA_GRUPOS.keys()),
        format_func=lambda x: f"{x} - {TABELA_GRUPOS[x]}" if x > 0 else TABELA_GRUPOS[x],
        index=0,
        help="0 = Detecta automaticamente pelo CFOP e NCM do produto.",
    )
    if grupo_selecionado == 0:
        st.caption(
            "Auto: CFOP 3102 → Grupo 2 (Mercadoria para Revenda)\n"
            "NCM 94xx → Grupo 10 (Ativo Imobilizado)"
        )
    else:
        st.caption(
            f"Todos os produtos receberao: "
            f"{grupo_selecionado} - {TABELA_GRUPOS[grupo_selecionado]}"
        )

    st.markdown("---")
    st.markdown("### Encoding")
    st.markdown("**Entrada:** UTF-8 / Latin-1")
    st.markdown("**Saida:** ANSI (Latin-1)")

    st.markdown("---")
    with st.expander("Tabela de Impostos Dominio"):
        for cod, nome in sorted(TABELA_IMPOSTOS.items()):
            st.caption(f"`{cod:3d}` - {nome}")

    with st.expander("Tabela de Grupos Dominio"):
        for cod, desc in sorted(TABELA_GRUPOS.items()):
            if cod > 0:
                st.caption(f"`{cod:3d}` - {desc}")


# ─────────────────────────────────────────────
# INSTRUCOES
# ─────────────────────────────────────────────
with st.expander("Instrucoes de Uso — clique para expandir", expanded=False):
    st.markdown(
        """
        <div class="instrucoes-box">
        <h4>Como o CNPJ e obtido</h4>
        <p>Lido automaticamente do XML na seguinte prioridade:<br>
        1. <code>&lt;dest&gt;&lt;CNPJ&gt;</code><br>
        2. <code>&lt;dest&gt;&lt;CPF&gt;</code><br>
        3. <code>&lt;dest&gt;&lt;idEstrangeiro&gt;</code> com valor<br>
        4. <code>&lt;emit&gt;&lt;CNPJ&gt;</code> quando UF=EX ou idEstrangeiro vazio<br>
        5. Campo <b>CNPJ Fallback</b> na sidebar</p>
        <h4>Grupo de Produtos (campo 9 do 0100)</h4>
        <p>Detectado automaticamente pelo CFOP (primario) e NCM (fallback).
        Pode ser fixado manualmente na sidebar para todos os produtos.</p>
        <h4>Registros gerados</h4>
        <ul>
        <li><b>0000</b> Identificacao da empresa</li>
        <li><b>0020</b> Cadastro do fornecedor</li>
        <li><b>0100</b> Cadastro de produtos (com grupo)</li>
        <li><b>1000</b> Cabecalho da nota (98 campos)</li>
        <li><b>1010</b> Informacoes complementares (fisco/contribuinte)</li>
        <li><b>1015</b> Observacoes (opcional)</li>
        <li><b>1020</b> Impostos por codigo (tabela oficial)</li>
        <li><b>1030</b> Itens / estoque (111 campos)</li>
        <li><b>1097</b> Dados do frete SP</li>
        <li><b>1150</b> IBS por cClassTrib</li>
        <li><b>1151</b> CBS por cClassTrib</li>
        </ul>
        <h4>Saida</h4>
        <p>Arquivo <b>.TXT em ANSI (Latin-1)</b> — padrao Dominio Sistemas.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("---")

# ─────────────────────────────────────────────
# UPLOAD
# ─────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Selecione um ou mais arquivos XML de NF-e",
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
            cnpj_fallback  = cnpj_fallback,
            acumulador     = acumulador,
            especie        = especie,
            incluir_0000   = inc_0000,
            incluir_0020   = inc_0020,
            incluir_0100   = inc_0100,
            incluir_1010   = inc_1010,
            incluir_1015   = inc_1015,
            incluir_1097   = inc_1097,
            grupo_padrao   = grupo_selecionado,
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

    if erros:
        st.error("Erros encontrados:")
        st.dataframe(erros, use_container_width=True)

    if all_resumos:
        st.success(f"{len(all_resumos)} arquivo(s) convertido(s) com sucesso!")

        # Badge CNPJ
        cnpjs_unicos = list({r["CNPJ Empresa"]: r for r in all_resumos}.values())
        if cnpjs_unicos:
            st.markdown("#### Empresa Identificada (lida do XML)")
            cols = st.columns(min(len(cnpjs_unicos), 4))
            for idx, r in enumerate(cnpjs_unicos[:4]):
                is_ext = r.get("Exterior", "Nao") == "Sim"
                cor    = "#1565C0" if is_ext else "#FF8000"
                with cols[idx]:
                    st.markdown(
                        f"""
                        <div class="cnpj-badge"
                             style="color:{cor};border-color:{cor};">
                            CNPJ: {r['CNPJ Empresa']}
                        </div>
                        <div class="info-origem"
                             style="border-left-color:{cor};">
                            {r.get('Origem CNPJ','')}<br>
                            {(r.get('Emitente','') if is_ext
                              else r.get('Destinatario',''))[:60]}<br>
                            {'<b>Importacao/Exterior</b>' if is_ext else ''}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

        st.markdown("---")

        with st.expander("Resumo das notas processadas", expanded=True):
            st.dataframe(all_resumos, use_container_width=True)

        saida_final = "\n".join(all_lines)

        with st.expander("Preview do arquivo gerado (primeiras 80 linhas)"):
            preview = saida_final.split("\n")[:80]
            st.code("\n".join(preview), language="text")

        st.markdown("---")

        saida_ansi = encode_ansi(saida_final)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="Baixar Arquivo Dominio (.TXT ANSI)",
                data=saida_ansi,
                file_name="importacao_dominio.txt",
                mime="text/plain",
                use_container_width=True,
                type="primary",
            )
        with col2:
            st.download_button(
                label="Baixar Arquivo (.TXT UTF-8)",
                data=saida_final.encode("utf-8"),
                file_name="importacao_dominio_utf8.txt",
                mime="text/plain",
                use_container_width=True,
            )

        st.markdown("---")
        st.markdown("#### Estatisticas")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Notas", len(all_resumos))
        c2.metric("Itens", sum(r.get("Itens", 0) for r in all_resumos))
        total_linhas = len([l for l in saida_final.split("\n") if l.startswith("|")])
        c3.metric("Linhas geradas", total_linhas)
        c4.metric("Erros", len(erros))
        try:
            total_nf = sum(
                float(r.get("vNF","0").replace(",","."))
                for r in all_resumos if r.get("vNF")
            )
            c5.metric("Total NF (R$)",
                      f"{total_nf:,.2f}".replace(",","X").replace(".",",").replace("X","."))
        except Exception:
            c5.metric("Total NF", "-")

else:
    st.info("Faca o upload de um ou mais arquivos XML de NF-e para iniciar.")

st.markdown("---")
st.caption(
    f"Conversor XML NF-e → Dominio Sistemas | "
    f"Thomson Reuters | Python + Streamlit | {VERSAO}"
)
