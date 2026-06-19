import streamlit as st
import xml.etree.ElementTree as ET
from datetime import datetime

# ─────────────────────────────────────────────
# VERSÃO
# ─────────────────────────────────────────────
VERSAO = "V1.0"

# ─────────────────────────────────────────────
# TEMA THOMSON REUTERS (idêntico ao RPA V3.9)
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
# CABEÇALHO THOMSON REUTERS
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
# TABELA DE IMPOSTOS – conforme Impostos.xls
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
# REGISTRO 0000 – Identificação da empresa
# Campos: 1=Identificação(fixo 0000), 2=Inscrição
# ─────────────────────────────────────────────
def gerar_registro_0000(cnpj_empresa: str) -> str:
    return pipe_join(["0000", cnpj_empresa])


# ─────────────────────────────────────────────
# REGISTRO 0020 – Cadastro de fornecedores
# 33 campos conforme leiaute oficial
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

    # Campo 24 - Regime de apuração
    regime_map = {"1": "M", "2": "E", "3": "N"}
    regime = regime_map.get(crt, "N")

    # Campo 25 - Contribuinte ICMS
    contrib_icms = "S" if ie and ie.upper() not in ("ISENTO", "NAO CONTRIBUINTE", "") else "N"

    fields = [
        "0020",       # 1  - Identificação do registro
        cnpj,         # 2  - Inscrição
        razao,        # 3  - Razão Social
        fantasia,     # 4  - Apelido
        logradouro,   # 5  - Endereço
        numero,       # 6  - Número do endereço
        complemento,  # 7  - Complemento
        bairro,       # 8  - Bairro
        cod_mun,      # 9  - Código do município
        uf,           # 10 - UF
        "",           # 11 - Código do País (só exterior)
        cep,          # 12 - CEP
        ie,           # 13 - Inscrição Estadual
        "",           # 14 - Inscrição Municipal
        "",           # 15 - Inscrição Suframa
        "",           # 16 - DDD
        "",           # 17 - Telefone
        "",           # 18 - FAX
        "",           # 19 - Data do cadastro
        "",           # 20 - Conta contábil
        "",           # 21 - Conta contábil cliente
        "N",          # 22 - Agropecuário
        "7",          # 23 - Natureza jurídica (7=Empresa Privada)
        regime,       # 24 - Regime de apuração
        contrib_icms, # 25 - Contribuinte ICMS
        "",           # 26 - Alíquota ICMS
        "",           # 27 - Categoria do estabelecimento
        "",           # 28 - Inscrição Estadual ST
        "",           # 29 - Email
        "N",          # 30 - Interdependência com a empresa
        "N",          # 31 - Contribuinte da CPRB
        "",           # 32 - Processo administrativo/judicial
        "",           # 33 - Tipo Inscrição
    ]
    return pipe_join(fields)


# ─────────────────────────────────────────────
# REGISTRO 0100 – Cadastro de produtos
# 91 campos conforme leiaute oficial
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
        icms_node = None
        for icms_type in ["ICMS00","ICMS10","ICMS20","ICMS30","ICMS40",
                          "ICMS51","ICMS60","ICMS70","ICMS90","ICMSSN101",
                          "ICMSSN102","ICMSSN201","ICMSSN202","ICMSSN500","ICMSSN900"]:
            node = imposto.find(f"nfe:ICMS/nfe:{icms_type}", NS)
            if node is not None:
                icms_node = node
                break
        if icms_node is not None:
            cst_icms  = get_text(icms_node, "nfe:CST") or get_text(icms_node, "nfe:CSOSN")
            aliq_icms = fmt_decimal(get_text(icms_node, "nfe:pICMS"))
        ipi_trib = imposto.find("nfe:IPI/nfe:IPITrib", NS)
        if ipi_trib is not None:
            aliq_ipi = fmt_decimal(get_text(ipi_trib, "nfe:pIPI"))

    fields = [
        "0100",                   # 1  - Identificação do registro
        cod_prod,                 # 2  - Código do produto
        descricao,                # 3  - Descrição do produto
        "",                       # 4  - Código NBM
        ncm,                      # 5  - Código NCM
        "",                       # 6  - Código NCM Exterior
        "",                       # 7  - Código de barras
        "",                       # 8  - Código do imposto de importação
        "",                       # 9  - Código do grupo de produtos
        unidade,                  # 10 - Unidade de medida
        "N",                      # 11 - Unidade inventária diferente
        "O",                      # 12 - Tipo do produto (O=Outros)
        "",                       # 13 - Tipo da arma de fogo
        "",                       # 14 - Descrição da arma de fogo
        "",                       # 15 - Tipo de medicamento
        "N",                      # 16 - Serviço tributado pelo ISSQN
        "",                       # 17 - Código do chassi do veículo
        fmt_decimal(val_unit, 3), # 18 - Valor unitário (3 casas)
        "",                       # 19 - Quantidade inicial em estoque
        "",                       # 20 - Valor inicial em estoque
        cst_icms,                 # 21 - CST ICMS
        aliq_icms,                # 22 - Alíquota do ICMS
        aliq_ipi,                 # 23 - Alíquota do IPI
        "",                       # 24 - Periodicidade do IPI
        "",                       # 25 - Observação
        "N",                      # 26 - Exporta produto para DNF
        "",                       # 27 - Ex TIPI
        "",                       # 28 - DNF – Código da espécie
        "",                       # 29 - DNF – Unidade de medida padrão
        "",                       # 30 - DNF – Fator de conversão
        "",                       # 31 - DNF – Código do produto
        "",                       # 32 - DNF – Capacidade Volumétrica
        "",                       # 33 - SE/DIC – Código EAN
        "",                       # 34 - SE/DIC – Código do produto relevante
        "N",                      # 35 - SCANC – Gerar para o SCANC
        "",                       # 36 - SCANC – Código do produto
        "",                       # 37 - SCANC – Contém gasolina A
        "",                       # 38 - SCANC – Tipo de produto
        "N",                      # 39 - GRF – CTB – Gera para o GRF
        "",                       # 40 - GRF – CTB – Código do produto
        "",                       # 41 - DIEF – Unidade
        "",                       # 42 - DIEF – Tipo de produto/serviço
        "N",                      # 43 - 88ST – Informa registro 88ST
        "",                       # 44 - 88ST – Código do produto
        "",                       # 45 - GO – Informações complementares IPM
        "",                       # 46 - GO – Código do produto/serviço IPM
        "N",                      # 47 - GO – Produto relacionado
        "N",                      # 48 - AM – Cesta básica
        "",                       # 49 - AM – Código do produto na DAM
        "",                       # 50 - RS – Produto incluído ST
        "",                       # 51 - RS – Data início ST
        "",                       # 52 - RS – Produto com preço tabelado
        "",                       # 53 - RS – Valor unitário ST
        "",                       # 54 - RS – MVA ST
        "",                       # 55 - RS – Grupo ST
        "N",                      # 56 - PR – Equipamento ECF
        "",                       # 57 - MS – Possui incentivo fiscal
        "",                       # 58 - DF – Produto sujeito regime especial
        "",                       # 59 - DF – Item padrão regime especial
        "",                       # 60 - PE – Tipo do produto
        "N",                      # 61 - SP – Controla ressarcimento Cat 17/99
        "",                       # 62 - SP – Data saldo inicial Cat 17/99
        "",                       # 63 - SP – Valor unitário Cat 17/99
        "",                       # 64 - SP – Quantidade Cat 17/99
        "",                       # 65 - SP – Valor final Cat 17/99
        "",                       # 66 - SPED – Gênero
        "",                       # 67 - SPED – Código do Serviço
        "",                       # 68 - SPED – Tipo do item
        "",                       # 69 - SPED – Classificação
        "",                       # 70 - SPED – Conta Contábil em seu poder
        "",                       # 71 - SPED – Conta Contábil em poder de terceiros
        "",                       # 72 - SPED – Conta Contábil de terceiros em seu poder
        "",                       # 73 - SPED – Tipo de receita
        "",                       # 74 - SPED – Energia elétrica/Gás
        "",                       # 75 - Data do cadastro
        "N",                      # 76 - Produto escriturado no LMC
        "",                       # 77 - Código combustível DF
        "",                       # 78 - Código combustível ANP
        "N",                      # 79 - Produto relacionado MP 540
        "",                       # 80 - Permitir descrição complementar
        "",                       # 81 - Código de atividade INSS Folha
        "",                       # 82 - DACON – Tipo do Produto
        "",                       # 83 - DACON – Crédito Presumido
        "",                       # 84 - Desconsiderar
        "",                       # 85 - SPED – Conta Contábil em processo
        "",                       # 86 - SPED – Conta Contábil histórico em processo
        "",                       # 87 - SPED – Conta Contábil acabado
        "",                       # 88 - SPED – Conta Contábil histórico acabado
        cest,                     # 89 - Código CEST
        "",                       # 90 - Registro de Exportação (RE)
        "",                       # 91 - Identificador
    ]
    return pipe_join(fields)


# ─────────────────────────────────────────────
# REGISTRO 1000 – Nota Fiscal de Entrada
# 98 campos conforme leiaute oficial
# ─────────────────────────────────────────────
def gerar_registro_1000(nfe_root, acumulador: str = "1157", especie: str = "36") -> str:
    ide   = nfe_root.find("nfe:infNFe/nfe:ide", NS)
    emit  = nfe_root.find("nfe:infNFe/nfe:emit", NS)
    total = nfe_root.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)

    cnpj_emit = get_text(emit, "nfe:CNPJ")
    nNF       = get_text(ide, "nfe:nNF")
    serie     = get_text(ide, "nfe:serie")
    dhEmi     = fmt_date(get_text(ide, "nfe:dhEmi"))
    ie_emit   = get_text(emit, "nfe:IE")

    # CFOP do primeiro item
    det_list   = nfe_root.findall("nfe:infNFe/nfe:det", NS)
    cfop_first = ""
    if det_list:
        cfop_first = get_text(det_list[0].find("nfe:prod", NS), "nfe:CFOP")

    # Município de origem (cMunFG da ide)
    c_mun_fg = get_text(ide, "nfe:cMunFG")

    # Valores totais
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

    # Chave NF-e
    inf_nfe = nfe_root.find("nfe:infNFe", NS)
    chave   = ""
    if inf_nfe is not None:
        id_attr = inf_nfe.get("Id", "")
        chave   = id_attr.replace("NFe", "") if id_attr else ""

    # Modalidade frete
    transp        = nfe_root.find("nfe:infNFe/nfe:transp", NS)
    mod_frete_cod = get_text(transp, "nfe:modFrete")
    frete_map     = {"0": "C", "1": "F", "2": "S", "3": "T", "4": "R", "5": "D", "9": "S"}
    mod_frete     = frete_map.get(mod_frete_cod, "C")

    # Observação fiscal (campo 15 - Observação de interesse do fisco)
    inf_adic  = nfe_root.find("nfe:infNFe/nfe:infAdic", NS)
    obs_fisco = ""
    if inf_adic is not None:
        obs_fisco = get_text(inf_adic, "nfe:infAdFisco")[:300]

    # Número DI do primeiro item (campo 52)
    n_di = ""
    if det_list:
        di_node = det_list[0].find("nfe:prod/nfe:DI", NS)
        if di_node is not None:
            n_di = get_text(di_node, "nfe:nDI")

    fields = [
        "1000",          # 1  - Identificação do registro
        especie,         # 2  - Código da espécie (36=NF-e mod.55)
        cnpj_emit,       # 3  - Inscrição fornecedor
        "",              # 4  - Código de Exclusão da DIEF
        acumulador,      # 5  - Código do acumulador (1157)
        cfop_first,      # 6  - CFOP
        "",              # 7  - Segmento
        nNF,             # 8  - Número do documento
        serie,           # 9  - Série
        "",              # 10 - Número do documento final
        dhEmi,           # 11 - Data da entrada (dd/mm/aaaa)
        dhEmi,           # 12 - Data emissão (dd/mm/aaaa)
        v_nf,            # 13 - Valor contábil (2 casas)
        "",              # 14 - Valor da exclusão da DIEF
        obs_fisco,       # 15 - Observação (interesse do fisco)
        mod_frete,       # 16 - Modalidade do frete (C/F/S/T/R/D)
        "T",             # 17 - Emitente da nota (T=Terceiros)
        "",              # 18 - CFOP estendido (SE)
        "",              # 19 - Código transferência de crédito (RS)
        "",              # 20 - Código recolhimento ISS Retido
        "",              # 21 - Código recolhimento IRRF
        "",              # 22 - Código da observação
        "",              # 23 - Data do visto (MG) dd/mm/aaaa
        "",              # 24 - Fato gerador da CRF (E/P)
        "",              # 25 - Fato gerador do IRRF (E/P)
        v_frete,         # 26 - Valor do frete (2 casas)
        v_seg,           # 27 - Valor do seguro (2 casas)
        v_outro,         # 28 - Valor das despesas (2 casas)
        v_pis,           # 29 - Valor do PIS (2 casas)
        "",              # 30 - Código antecipação tributária
        v_cofins,        # 31 - Valor do COFINS (2 casas)
        "",              # 32 - Valor DARE (SE) (2 casas)
        "",              # 33 - Alíquota DARE (SE) (2 casas)
        "",              # 34 - Base cálculo ICMS ST
        "",              # 35 - Entradas isentas saídas (MG) (2 casas)
        "",              # 36 - Outras entradas isentas (MG) (2 casas)
        "",              # 37 - Valor transporte base (MG) (2 casas)
        "",              # 38 - Código de ressarcimento
        v_prod,          # 39 - Valor produtos (2 casas)
        c_mun_fg,        # 40 - Município Origem (cMunFG da ide)
        "0",             # 41 - Situação da Nota (0=Documento Regular)
        "",              # 42 - Código da situação tributária
        "",              # 43 - Sub série
        ie_emit,         # 44 - Inscrição estadual do fornecedor
        "",              # 45 - Inscrição municipal do fornecedor
        "",              # 46 - Código da operação e prestação
        "",              # 47 - Valor deduzido da receita tributável (2 casas)
        "",              # 48 - Competência (dd/mm/aaaa)
        "",              # 49 - Operação (PA)
        "",              # 50 - Número do parecer fiscal
        "",              # 51 - Data do parecer fiscal (dd/mm/aaaa)
        n_di,            # 52 - Número da declaração de Importação
        "N",             # 53 - Possui benefício fiscal (S/N)
        chave,           # 54 - Chave da nota fiscal eletrônica
        "",              # 55 - Código recolhimento FETHAB
        "",              # 56 - Responsável recolhimento FETHAB (E/C)
        "",              # 57 - CFOP documento fiscal
        "",              # 58 - Tipo de CT-e (0/1/2)
        "",              # 59 - CT-e referência
        "1",             # 60 - Modalidade da importação (1=Com direito a crédito)
        "",              # 61 - Código da informação complementar
        "",              # 62 - Informação complementar
        "",              # 63 - Classe de consumo
        "",              # 64 - Tipo de ligação
        "",              # 65 - Grupo de tensão
        "",              # 66 - Tipo de assinante
        "",              # 67 - KWH consumido
        "",              # 68 - Valor fornecido energia/gás (2 casas)
        "",              # 69 - Valor cobrado de terceiros (2 casas)
        "10",            # 70 - Tipo doc importação (10=DI / 1=DSI)
        "",              # 71 - Número Ato Concessório Drawback
        "",              # 72 - Natureza do frete PIS/COFINS
        "",              # 73 - CST PIS/COFINS
        "",              # 74 - Base do crédito PIS/COFINS
        "",              # 75 - Valor serviços/itens PIS/COFINS (2 casas)
        "",              # 76 - Base de cálculo PIS/COFINS (2 casas)
        "",              # 77 - Alíquota de PIS (2 casas)
        "",              # 78 - Alíquota de COFINS (2 casas)
        "",              # 79 - Chave de NFSe
        "",              # 80 - Número processo/ato concessório
        "",              # 81 - Origem do processo
        "",              # 82 - Data da escrituração (dd/mm/aaaa)
        "",              # 83 - CFPS (DF)
        "",              # 84 - Natureza da receita PIS/COFINS
        "",              # 85 - CST IPI
        "",              # 86 - Lançamentos de SCP
        "",              # 87 - Tipo de serviço (1/2)
        "",              # 88 - Município destino
        "",              # 89 - Pedágio (2 casas)
        v_ipi,           # 90 - IPI (2 casas)
        v_st,            # 91 - ICMS ST (2 casas)
        "",              # 92 - Classificação Serviços EFD-Reinf
        "",              # 93 - Indicativo Prestação EFD-Reinf
        "",              # 94 - Número doc arrecadação (RS)
        "",              # 95 - Tipo do título (0-4/99)
        "",              # 96 - Identificação
        v_icms_d,        # 97 - ICMS Desonerado (2 casas)
        "",              # 98 - IPI Devolução (2 casas)
    ]
    return pipe_join(fields)


# ─────────────────────────────────────────────
# REGISTRO 1020 – Impostos da nota
# 19 campos | Um registro por imposto presente
# Códigos conforme tabela oficial Impostos.xls
# ─────────────────────────────────────────────
def gerar_registros_1020(nfe_root) -> list:
    total  = nfe_root.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)
    ibs_t  = nfe_root.find("nfe:infNFe/nfe:total/nfe:IBSCBSTot", NS)
    v_nf   = fmt_decimal(get_text(total, "nfe:vNF"))
    linhas = []

    def linha_1020(cod_imp, perc_red="", base="", aliq="", valor="",
                   isentas="", outras="", v_ipi="", v_st="", v_cont="", cod_rec=""):
        fields = [
            "1020",   # 1  - Identificação do registro
            cod_imp,  # 2  - Código do imposto (tabela Impostos.xls)
            perc_red, # 3  - % redução base de cálculo (2 casas)
            base,     # 4  - Base de cálculo (2 casas)
            aliq,     # 5  - Alíquota (2 casas; 3 casas p/ imposto 39)
            valor,    # 6  - Valor do imposto (2 casas)
            isentas,  # 7  - Valor de isentas (2 casas)
            outras,   # 8  - Valor de outras (2 casas)
            v_ipi,    # 9  - Valor do IPI (2 casas)
            v_st,     # 10 - Valor da substituição tributária (2 casas)
            v_cont,   # 11 - Valor contábil (2 casas)
            cod_rec,  # 12 - Código do recolhimento do imposto
            "",       # 13 - Valor não tributadas (GO) (2 casas)
            "",       # 14 - Valor parcela reduzida (GO) (2 casas)
            "",       # 15 - Alíq. Interestadual
            "",       # 16 - Nat. rendimentos
            "",       # 17 - Tipo de Dedução
            "",       # 18 - Tipo de Isenção
            "",       # 19 - Descrição
        ]
        return pipe_join(fields)

    # Lê valores do total
    v_bc_icms   = get_text(total, "nfe:vBC")
    v_icms      = get_text(total, "nfe:vICMS")
    v_ipi_tot   = get_text(total, "nfe:vIPI")
    v_st_tot    = get_text(total, "nfe:vST")
    v_pis_tot   = get_text(total, "nfe:vPIS")
    v_cofins_tot= get_text(total, "nfe:vCOFINS")
    v_icms_deson= get_text(total, "nfe:vICMSDeson")

    # Alíquota média ICMS
    try:
        bc_f  = float(v_bc_icms)
        icm_f = float(v_icms)
        aliq_icms_med = fmt_decimal(str(icm_f / bc_f * 100)) if bc_f > 0 else ""
    except (ValueError, ZeroDivisionError):
        aliq_icms_med = ""

    # Soma bases por imposto nos itens
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
        for pis_type in ["PISAliq", "PISQtde", "PISNT", "PISOutr"]:
            pn = imp.find(f"nfe:PIS/nfe:{pis_type}", NS)
            if pn is not None:
                try:
                    bc_pis_total += float(get_text(pn, "nfe:vBC") or "0")
                except ValueError:
                    pass
                break
        for cof_type in ["COFINSAliq", "COFINSQtde", "COFINSNT", "COFINSOutr"]:
            cn = imp.find(f"nfe:COFINS/nfe:{cof_type}", NS)
            if cn is not None:
                try:
                    bc_cof_total += float(get_text(cn, "nfe:vBC") or "0")
                except ValueError:
                    pass
                break

    # ── 1 – ICMS ──────────────────────────────────────────────────────
    if v_icms and float(v_icms) > 0:
        linhas.append(linha_1020(
            cod_imp = 1,
            base    = fmt_decimal(v_bc_icms),
            aliq    = aliq_icms_med,
            valor   = fmt_decimal(v_icms),
            v_ipi   = fmt_decimal(v_ipi_tot),
            v_st    = fmt_decimal(v_st_tot),
            v_cont  = v_nf,
        ))

    # ── 2 – IPI ───────────────────────────────────────────────────────
    if v_ipi_tot and float(v_ipi_tot) > 0:
        linhas.append(linha_1020(
            cod_imp = 2,
            base    = fmt_decimal(str(bc_ipi_total)),
            valor   = fmt_decimal(v_ipi_tot),
            v_cont  = v_nf,
        ))

    # ── 4 – PIS ───────────────────────────────────────────────────────
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(linha_1020(
            cod_imp = 4,
            base    = fmt_decimal(str(bc_pis_total)),
            valor   = fmt_decimal(v_pis_tot),
            v_cont  = v_nf,
        ))

    # ── 5 – COFINS ────────────────────────────────────────────────────
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(linha_1020(
            cod_imp = 5,
            base    = fmt_decimal(str(bc_cof_total)),
            valor   = fmt_decimal(v_cofins_tot),
            v_cont  = v_nf,
        ))

    # ── 45 – ICMS Importação (valor desonerado) ───────────────────────
    if v_icms_deson and float(v_icms_deson) > 0:
        linhas.append(linha_1020(
            cod_imp = 45,
            base    = fmt_decimal(v_bc_icms),
            aliq    = aliq_icms_med,
            valor   = fmt_decimal(v_icms_deson),
            v_cont  = v_nf,
        ))

    # ── 133 – PIS Importação ─────────────────────────────────────────
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(linha_1020(
            cod_imp = 133,
            base    = fmt_decimal(str(bc_pis_total)),
            valor   = fmt_decimal(v_pis_tot),
            v_cont  = v_nf,
        ))

    # ── 134 – COFINS Importação ──────────────────────────────────────
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(linha_1020(
            cod_imp = 134,
            base    = fmt_decimal(str(bc_cof_total)),
            valor   = fmt_decimal(v_cofins_tot),
            v_cont  = v_nf,
        ))

    # ── 183 – IBS ────────────────────────────────────────────────────
    if ibs_t is not None:
        v_ibs_uf = get_text(ibs_t, "nfe:gIBS/nfe:gIBSUF/nfe:vIBSUF")
        bc_ibs   = get_text(ibs_t, "nfe:vBCIBSCBS")
        if v_ibs_uf and float(v_ibs_uf) > 0:
            linhas.append(linha_1020(
                cod_imp = 183,
                base    = fmt_decimal(bc_ibs),
                valor   = fmt_decimal(v_ibs_uf),
                v_cont  = v_nf,
            ))

    # ── 184 – CBS ────────────────────────────────────────────────────
    if ibs_t is not None:
        v_cbs  = get_text(ibs_t, "nfe:gCBS/nfe:vCBS")
        bc_cbs = get_text(ibs_t, "nfe:vBCIBSCBS")
        if v_cbs and float(v_cbs) > 0:
            linhas.append(linha_1020(
                cod_imp = 184,
                base    = fmt_decimal(bc_cbs),
                valor   = fmt_decimal(v_cbs),
                v_cont  = v_nf,
            ))

    return linhas


# ─────────────────────────────────────────────
# REGISTRO 1030 – Estoque (item da nota)
# 111 campos conforme leiaute oficial
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

    # DI
    di_node = prod.find("nfe:DI", NS)
    n_di    = get_text(di_node, "nfe:nDI") if di_node is not None else ""
    d_di    = fmt_date(get_text(di_node, "nfe:dDI")) if di_node is not None else ""

    # ICMS
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

        # IPI
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

        # PIS
        v_pis = aliq_pis = cst_pis = bc_pis = ""
        pis_node = imposto.find("nfe:PIS", NS)
        if pis_node is not None:
            for pis_type in ["PISAliq", "PISQtde", "PISNT", "PISOutr"]:
                pn = pis_node.find(f"nfe:{pis_type}", NS)
                if pn is not None:
                    v_pis    = fmt_decimal(get_text(pn, "nfe:vPIS"))
                    aliq_pis = fmt_decimal(get_text(pn, "nfe:pPIS"), 4)
                    cst_pis  = get_text(pn, "nfe:CST")
                    bc_pis   = fmt_decimal(get_text(pn, "nfe:vBC"))
                    break

        # COFINS
        v_cof = aliq_cof = cst_cof = bc_cof = ""
        cof_node = imposto.find("nfe:COFINS", NS)
        if cof_node is not None:
            for cof_type in ["COFINSAliq", "COFINSQtde", "COFINSNT", "COFINSOutr"]:
                cn = cof_node.find(f"nfe:{cof_type}", NS)
                if cn is not None:
                    v_cof    = fmt_decimal(get_text(cn, "nfe:vCOFINS"))
                    aliq_cof = fmt_decimal(get_text(cn, "nfe:pCOFINS"), 4)
                    cst_cof  = get_text(cn, "nfe:CST")
                    bc_cof   = fmt_decimal(get_text(cn, "nfe:vBC"))
                    break

        # IBS / CBS
        ibs_node       = imposto.find("nfe:IBSCBS", NS)
        ibs_class_trib = ""
        ibs_bc = ibs_aliq = ibs_val = ""
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

    # Campo 4 = Base Cal. + IPI (conforme leiaute)
    try:
        vp = float(v_prod or "0")
        vi = float(get_text(
            imposto.find("nfe:IPI/nfe:IPITrib", ET.Element("x")) if imposto else None,
            "nfe:vIPI") or "0")
        v_total = fmt_decimal(str(vp + vi))
    except (ValueError, AttributeError, TypeError):
        v_total = fmt_decimal(v_prod)

    fields = [
        "1030",                    # 1   - Identificação do registro
        cod_prod,                  # 2   - Código do produto
        qtd,                       # 3   - Quantidade (2 casas)
        v_total,                   # 4   - Valor total (Base Cal. + IPI) (2 casas)
        v_ipi,                     # 5   - Valor IPI (2 casas)
        fmt_decimal(v_prod),       # 6   - Base de cálculo (2 casas)
        "1",                       # 7   - Tipo de Lançamento (1=vinculado à nota)
        d_di,                      # 8   - Data (dd/mm/aaaa)
        n_di,                      # 9   - Número da DI (2 casas)
        cst_icms,                  # 10  - Código da Situação Tributária
        fmt_decimal(v_prod),       # 11  - Valor bruto do produto (2 casas)
        fmt_decimal(v_desc),       # 12  - Valor do desconto (2 casas)
        v_bc_icms,                 # 13  - Base de cálculo do ICMS (2 casas)
        v_bc_st,                   # 14  - Base de cálculo do ICMS ST (2 casas)
        aliq_icms,                 # 15  - Alíquota do ICMS (2 casas)
        "",                        # 16  - Produto Incentivado (PE) S/N
        "",                        # 17  - Código da apuração (PE)
        "",                        # 18  - Valor do frete (2 casas)
        "",                        # 19  - Valor do seguro (2 casas)
        fmt_decimal(v_outro),      # 20  - Valor das despesas acessórias (2 casas)
        "",                        # 21  - Quantidade de gasolina (3 casas)
        v_icms,                    # 22  - Valor do ICMS (2 casas)
        "",                        # 23  - Valor da SUBTRI (2 casas)
        "",                        # 24  - Valor de isentas IPI (2 casas)
        "",                        # 25  - Valor de outras IPI (2 casas)
        "",                        # 26  - ICMS NFP (2 casas)
        fmt_decimal(v_unit, 6),    # 27  - Valor Unitário (6 casas)
        "",                        # 28  - Alíquota da Substituição Tributária (2 casas)
        cst_ipi,                   # 29  - Código de Tributação do IPI
        aliq_ipi,                  # 30  - Alíquota do IPI (2 casas)
        "",                        # 31  - Base de cálculo ISSQN (2 casas)
        "",                        # 32  - Alíquota do ISSQN (2 casas)
        "",                        # 33  - Valor ISSQN (2 casas)
        cfop,                      # 34  - CFOP
        "",                        # 35  - Série de fabricação do ECF
        aliq_pis,                  # 36  - Alíquota do PIS (4 casas)
        v_pis,                     # 37  - Valor do PIS (2 casas)
        aliq_cof,                  # 38  - Alíquota da COFINS (4 casas)
        v_cof,                     # 39  - Valor da COFINS (2 casas)
        fmt_decimal(v_prod),       # 40  - Custo total do produto (2 casas)
        cst_pis,                   # 41  - CST do PIS
        bc_pis,                    # 42  - Base de cálculo do PIS (2 casas)
        cst_cof,                   # 43  - CST da COFINS
        bc_cof,                    # 44  - Base de cálculo da COFINS (2 casas)
        "",                        # 45  - Chassi do veículo
        "",                        # 46  - Tipo de operação com o veículo (00/01/02/99)
        "",                        # 47  - Lote do medicamento
        "",                        # 48  - Quantidade de item por lote
        "",                        # 49  - Data de validade (dd/mm/aaaa)
        "",                        # 50  - Data de fabricação (dd/mm/aaaa)
        "",                        # 51  - Referência base de cálculo (00-05)
        "",                        # 52  - Valor tabelado/máximo (2 casas)
        "",                        # 53  - Número de série da arma
        "",                        # 54  - Número de série do cano
        "",                        # 55  - Enquadramento do IPI
        "S",                       # 56  - Movimentação física do produto (S/N)
        unidade,                   # 57  - Unidade comercializada
        "",                        # 58  - Complemento da CFOP (CAT 17/99)
        "",                        # 59  - Tanque do combustível
        fmt_decimal(v_prod),       # 60  - Valor contábil produto (2 casas)
        "",                        # 61  - Qtd tributada PIS por unid. medida (3 casas)
        "",                        # 62  - Valor unidade PIS por unid. medida (4 casas)
        "",                        # 63  - Valor PIS por unid. medida (2 casas)
        "",                        # 64  - Qtd tributada COFINS por unid. medida (3 casas)
        "",                        # 65  - Valor unidade COFINS por unid. medida (4 casas)
        "",                        # 66  - Valor COFINS por unid. medida (2 casas)
        "",                        # 67  - Base do crédito (01-18)
        "",                        # 68  - Número da Nota/Redução Z devolvido
        "",                        # 69  - Descrição complementar
        "",                        # 70  - Nota devolvida – CST PIS
        "",                        # 71  - Nota devolvida – CST COFINS
        "",                        # 72  - Vínculo de Crédito PIS
        "",                        # 73  - Vínculo de Crédito COFINS
        "",                        # 74  - Exclusão PIS (2 casas)
        "",                        # 75  - Exclusão COFINS (2 casas)
        "",                        # 76  - Base de cálculo ICMS Carga Média (2 casas)
        "",                        # 77  - Alíquota ICMS Carga Média (2 casas)
        "",                        # 78  - Valor ICMS Carga Média (2 casas)
        "",                        # 79  - Número série ECF devolvido
        "",                        # 80  - PIS/COFINS % redução base de cálculo (2 casas)
        "",                        # 81  - Código recolhimento PIS devolvido
        "",                        # 82  - Código recolhimento COFINS devolvido
        "",                        # 83  - Código de recolhimento PIS
        "",                        # 84  - Código de recolhimento COFINS
        "",                        # 85  - Crédito Presumido PIS/COFINS – PIS (2 casas)
        "",                        # 86  - Crédito Presumido PIS/COFINS – COFINS (2 casas)
        "",                        # 87  - ICMS ST Antecipação Total – Base (2 casas)
        "",                        # 88  - ICMS ST Antecipação Total – Alíquota (2 casas)
        "",                        # 89  - ICMS ST Antecipação Total – Valor (2 casas)
        "",                        # 90  - Código de recolhimento IPI
        cest,                      # 91  - Código CEST
        "",                        # 92  - ICMS ST Retido – Base de cálculo (2 casas)
        "",                        # 93  - ICMS ST Retido – Valor (2 casas)
        "",                        # 94  - ICMS ST Retido – Possui tag no XML (S/N)
        "",                        # 95  - Identificador
        "",                        # 96  - ICMS Próprio do Substituto – Valor (2 casas)
        v_icms_des,                # 97  - Valor Desonerado (2 casas)
        "",                        # 98  - Código (1-12/90)
        "",                        # 99  - ICMS Não creditado
        "",                        # 100 - ICMS Monofásico Qtde Trib. (4 casas)
        "",                        # 101 - ICMS Monofásico Alíq. Fixa (4 casas)
        "",                        # 102 - ICMS Monofásico Valor (2 casas)
        "",                        # 103 - ICMS Monofásico FCV (4 casas)
        ibs_class_trib,            # 104 - IBS – cClass Trib
        ibs_bc,                    # 105 - IBS – Base de cálculo (2 casas)
        ibs_aliq,                  # 106 - IBS – Alíquota (2 casas)
        ibs_val,                   # 107 - IBS – Valor (2 casas)
        ibs_class_trib,            # 108 - CBS – cClass Trib
        cbs_bc,                    # 109 - CBS – Base de cálculo (2 casas)
        cbs_aliq,                  # 110 - CBS – Alíquota (2 casas)
        cbs_val,                   # 111 - CBS – Valor (2 casas)
    ]
    return pipe_join(fields)


# ─────────────────────────────────────────────
# REGISTRO 1150 – IBS (filho do 1000)
# 5 campos conforme leiaute oficial
# ─────────────────────────────────────────────
def gerar_registro_1150(class_trib, bc, aliq, valor) -> str:
    fields = [
        "1150",      # 1 - Identificação (fixo 1150)
        class_trib,  # 2 - IBS – cClassTrib
        bc,          # 3 - IBS – Base de cálculo (2 casas)
        aliq,        # 4 - IBS – Alíquota (2 casas)
        valor,       # 5 - IBS – Valor (2 casas)
    ]
    return pipe_join(fields)


# ─────────────────────────────────────────────
# REGISTRO 1151 – CBS (filho do 1000)
# 5 campos conforme leiaute oficial
# ─────────────────────────────────────────────
def gerar_registro_1151(class_trib, bc, aliq, valor) -> str:
    fields = [
        "1151",      # 1 - Identificação (fixo 1151)
        class_trib,  # 2 - CBS – cClassTrib
        bc,          # 3 - CBS – Base de cálculo (2 casas)
        aliq,        # 4 - CBS – Alíquota (2 casas)
        valor,       # 5 - CBS – Valor (2 casas)
    ]
    return pipe_join(fields)


# ─────────────────────────────────────────────
# FUNÇÃO PRINCIPAL DE CONVERSÃO
# ─────────────────────────────────────────────
def converter_xml(
    xml_content: bytes,
    cnpj_empresa: str,
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

    # Suporte a nfeProc ou NFe direto
    nfe = root.find("nfe:NFe", NS)
    if nfe is None:
        nfe = root

    lines    = []
    resumo   = {}
    det_list = nfe.findall("nfe:infNFe/nfe:det", NS)
    emit     = nfe.find("nfe:infNFe/nfe:emit", NS)
    ide      = nfe.find("nfe:infNFe/nfe:ide", NS)
    total    = nfe.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)

    resumo["nNF"]       = get_text(ide, "nfe:nNF")
    resumo["Emitente"]  = get_text(emit, "nfe:xNome")
    resumo["CNPJ Emit"] = get_text(emit, "nfe:CNPJ")
    resumo["Emissão"]   = fmt_date(get_text(ide, "nfe:dhEmi"))
    resumo["Itens"]     = len(det_list)
    resumo["vNF"]       = fmt_decimal(get_text(total, "nfe:vNF"))
    resumo["vICMS"]     = fmt_decimal(get_text(total, "nfe:vICMS"))
    resumo["vICMSDes"]  = fmt_decimal(get_text(total, "nfe:vICMSDeson"))
    resumo["vIPI"]      = fmt_decimal(get_text(total, "nfe:vIPI"))
    resumo["vPIS"]      = fmt_decimal(get_text(total, "nfe:vPIS"))
    resumo["vCOFINS"]   = fmt_decimal(get_text(total, "nfe:vCOFINS"))
    resumo["vII"]       = fmt_decimal(get_text(total, "nfe:vII"))

    # ── Registros de cadastro ─────────────────────────────────────────
    if incluir_0000:
        lines.append(gerar_registro_0000(cnpj_empresa))
    if incluir_0020 and emit is not None:
        lines.append(gerar_registro_0020(emit))
    if incluir_0100:
        produtos_gerados = set()
        for det in det_list:
            prod     = det.find("nfe:prod", NS)
            cod_prod = get_text(prod, "nfe:cProd")
            if cod_prod not in produtos_gerados:
                lines.append(gerar_registro_0100(det))
                produtos_gerados.add(cod_prod)

    # ── Registro 1000 ─────────────────────────────────────────────────
    lines.append(gerar_registro_1000(nfe, acumulador, especie))

    # ── Registros 1020 (um por imposto) ──────────────────────────────
    for r1020 in gerar_registros_1020(nfe):
        lines.append(r1020)

    # ── Registros 1030 (um por item) ─────────────────────────────────
    for seq, det in enumerate(det_list, start=1):
        lines.append(gerar_registro_1030(det, seq))

    # ── Registros 1150 / 1151 (IBS/CBS agrupados por cClassTrib) ─────
    ibs_gerados = {}
    for det in det_list:
        imposto  = det.find("nfe:imposto", NS)
        if imposto is None:
            continue
        ibs_node = imposto.find("nfe:IBSCBS", NS)
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
                ibs_gerados[ct]["v_ibs"]  += float(get_text(guf, "nfe:vIBSUF") or "0")
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

    for ct, dados in ibs_gerados.items():
        lines.append(gerar_registro_1150(
            ct,
            fmt_decimal(str(dados["bc_ibs"])),
            fmt_decimal(dados["aliq_ibs"]),
            fmt_decimal(str(dados["v_ibs"])),
        ))
        lines.append(gerar_registro_1151(
            ct,
            fmt_decimal(str(dados["bc_cbs"])),
            fmt_decimal(dados["aliq_cbs"]),
            fmt_decimal(str(dados["v_cbs"])),
        ))

    return "\n".join(lines), resumo


# ─────────────────────────────────────────────
# SIDEBAR THOMSON REUTERS
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"### ℹ Sobre")
    st.markdown(f"**Versão:** {VERSAO}")
    st.markdown("**Thomson Reuters**")
    st.markdown("**Domínio Sistemas**")
    st.markdown("---")

    st.markdown("### ⚙ Parâmetros")
    cnpj_empresa = st.text_input(
        "CNPJ da Empresa (destinatária)",
        value="",
        max_chars=14,
        help="Somente números, sem pontos ou traços."
    )
    acumulador = st.text_input(
        "Código do Acumulador",
        value="1157",
        help="1157 = CFOP 3102 (Importação)"
    )
    especie = st.text_input(
        "Código da Espécie",
        value="36",
        help="36 = NF-e modelo 55"
    )

    st.markdown("---")
    st.markdown("### 📋 Registros de cadastro")
    inc_0000 = st.checkbox("0000 – Identificação da empresa", value=True)
    inc_0020 = st.checkbox("0020 – Fornecedor (emitente)", value=True)
    inc_0100 = st.checkbox("0100 – Produtos", value=True)
    st.caption("1000 / 1020 / 1030 / 1150 / 1151 são sempre gerados.")

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

        <h4>🔹 Passo 1 — Informe o CNPJ</h4>
        <p>Informe o <b>CNPJ da empresa destinatária</b> (sua empresa) na barra lateral,
        apenas números.</p>

        <h4>🔹 Passo 2 — Confirme os parâmetros</h4>
        <p>Confirme o <b>Acumulador</b> (<code>1157</code> para CFOP 3102) e a
        <b>Espécie</b> (<code>36</code> para NF-e modelo 55).</p>

        <h4>🔹 Passo 3 — Upload dos XMLs</h4>
        <p>Faça o upload de um ou mais arquivos <b>XML de NF-e</b> (nfeProc ou NFe).</p>

        <h4>🔹 Passo 4 — Baixar e importar</h4>
        <p>Clique em <b>⬇ Baixar Arquivo Domínio</b> e importe o arquivo <code>.TXT</code>
        gerado no módulo fiscal do <b>Domínio Sistemas</b>.</p>

        <hr>

        <h4>⚠ Observações</h4>
        <ul>
            <li>O arquivo de saída é gerado em <b>ANSI (Latin-1)</b>, padrão do Domínio.</li>
            <li>Registros gerados: <b>0000, 0020, 0100, 1000, 1020, 1030, 1150, 1151</b>.</li>
            <li>O registro <b>1020</b> é gerado um por imposto presente na nota,
                conforme a tabela oficial de impostos do Domínio Sistemas.</li>
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
    help="Arquivos XML de NF-e (nfeProc ou NFe) — modelo 55",
)

# ─────────────────────────────────────────────
# PROCESSAMENTO
# ─────────────────────────────────────────────
if uploaded_files:
    if not cnpj_empresa:
        st.warning("⚠️ Informe o CNPJ da empresa destinatária na barra lateral.")
        st.stop()

    all_lines   = []
    all_resumos = []
    erros       = []

    progress = st.progress(0, text="Processando arquivos...")

    for i, f in enumerate(uploaded_files):
        xml_bytes = f.read()
        texto, resumo = converter_xml(
            xml_bytes,
            cnpj_empresa=cnpj_empresa,
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

        with st.expander("📊 Resumo das notas processadas", expanded=True):
            st.dataframe(all_resumos, use_container_width=True)

        saida_final = "\n".join(all_lines)

        with st.expander("👁️ Preview do arquivo gerado (primeiras 80 linhas)"):
            preview = saida_final.split("\n")[:80]
            st.code("\n".join(preview), language="text")

        st.markdown("---")

        # Download ANSI (padrão Domínio)
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
            c5.metric("Total NF", "-")

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
