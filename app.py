import streamlit as st
import xml.etree.ElementTree as ET
from datetime import datetime

# ─────────────────────────────────────────────
# VERSÃO
# ─────────────────────────────────────────────
VERSAO = "V1.2"

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
        .stButton > button:hover { background-color: #D64001; color: #FFFFFF; }
        .stDownloadButton > button {
            background-color: #FF8000; color: #FFFFFF;
            border: none; border-radius: 4px; font-weight: bold;
        }
        .stDownloadButton > button:hover { background-color: #D64001; color: #FFFFFF; }
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
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        .instrucoes-box h4 {
            color: #FF8000; margin-top: 14px; margin-bottom: 6px;
        }
        .instrucoes-box h4:first-child { margin-top: 0; }
        .cnpj-badge {
            background-color: #444444; color: #FF8000;
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
    <div style="background:#444444; padding:24px 28px 18px 28px;
                border-radius:8px; border-top:6px solid #FF8000;
                margin-bottom:28px;">
        <h2 style="color:#FF8000; margin:0;
                   font-family:'Segoe UI',Arial,sans-serif;">
            🔄 Conversor XML NF-e → Domínio Sistemas &nbsp;|&nbsp; {VERSAO}
        </h2>
        <p style="color:#DDDDDD; margin:6px 0 0 0;
                  font-family:'Segoe UI',Arial,sans-serif;">
            Converte arquivos XML de NF-e para o leiaute padrão de importação do
            <strong>Domínio Sistemas</strong>.
            Saída em <strong>ANSI (Latin-1)</strong>.
            CNPJ lido automaticamente do XML — inclusive para
            <strong>NF-e de importação (dest. exterior)</strong>.
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
    88:"Regime especial PMCMV",
    89:"PIS Regime especial PMCMV",
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
# EXTRAÇÃO DO CNPJ DESTINATÁRIO
# Prioridade:
#   1. <dest><CNPJ>          → NF-e doméstica
#   2. <dest><CPF>           → Pessoa física
#   3. <dest><idEstrangeiro> → Exterior com valor
#   4. <emit><CNPJ>          → Exterior sem id (idEstrangeiro vazio / UF=EX)
#   5. Fallback manual
# ─────────────────────────────────────────────
def extrair_cnpj_dest(nfe_root) -> tuple:
    """Retorna (inscricao, origem, is_exterior)."""
    dest = nfe_root.find("nfe:infNFe/nfe:dest", NS)

    if dest is not None:
        # 1. CNPJ
        cnpj = get_text(dest, "nfe:CNPJ")
        if cnpj:
            return cnpj, "XML — <dest><CNPJ>", False

        # 2. CPF
        cpf = get_text(dest, "nfe:CPF")
        if cpf:
            return cpf, "XML — <dest><CPF>", False

        # 3. idEstrangeiro com valor
        id_ext = get_text(dest, "nfe:idEstrangeiro")
        if id_ext and id_ext.strip():
            return id_ext.strip(), "XML — <dest><idEstrangeiro>", True

        # 4. idEstrangeiro vazio OU UF=EX → usa emit/CNPJ
        id_ext_node = dest.find("nfe:idEstrangeiro", NS)
        ender_dest  = dest.find("nfe:enderDest", NS)
        uf_dest     = get_text(ender_dest, "nfe:UF") if ender_dest is not None else ""

        if id_ext_node is not None or uf_dest == "EX":
            emit = nfe_root.find("nfe:infNFe/nfe:emit", NS)
            cnpj_emit = get_text(emit, "nfe:CNPJ") if emit is not None else ""
            if cnpj_emit:
                return cnpj_emit, "XML — <emit><CNPJ> (dest. exterior / UF=EX)", True
            return "", "Destinatario exterior sem CNPJ identificado", True

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
# REGISTRO 0000 – Identificação da empresa (2 campos)
# ─────────────────────────────────────────────
def gerar_registro_0000(cnpj_empresa: str) -> str:
    return pipe_join(["0000", cnpj_empresa])


# ─────────────────────────────────────────────
# REGISTRO 0020 – Cadastro de fornecedores (33 campos)
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
        "0020",       # 1  - Identificação do registro
        cnpj,         # 2  - Inscrição (CNPJ)
        razao,        # 3  - Razão Social
        fantasia,     # 4  - Apelido
        logradouro,   # 5  - Endereço
        numero,       # 6  - Número
        complemento,  # 7  - Complemento
        bairro,       # 8  - Bairro
        cod_mun,      # 9  - Código do município
        uf,           # 10 - UF
        "",           # 11 - Código do País
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
        "7",          # 23 - Natureza jurídica
        regime,       # 24 - Regime de apuração
        contrib_icms, # 25 - Contribuinte ICMS
        "",           # 26 - Alíquota ICMS
        "",           # 27 - Categoria do estabelecimento
        "",           # 28 - Inscrição Estadual ST
        "",           # 29 - Email
        "N",          # 30 - Interdependência
        "N",          # 31 - Contribuinte da CPRB
        "",           # 32 - Processo adm/judicial
        "",           # 33 - Tipo Inscrição
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
        "0100",                   # 1
        cod_prod,                 # 2
        descricao,                # 3
        "",                       # 4  - NBM
        ncm,                      # 5  - NCM
        "",                       # 6  - NCM Exterior
        "",                       # 7  - Código de barras
        "",                       # 8  - Cód. imposto importação
        "",                       # 9  - Cód. grupo de produtos
        unidade,                  # 10 - Unidade de medida
        "N",                      # 11 - Unidade inventária diferente
        "O",                      # 12 - Tipo do produto
        "",                       # 13 - Tipo da arma
        "",                       # 14 - Descrição da arma
        "",                       # 15 - Tipo de medicamento
        "N",                      # 16 - ISSQN
        "",                       # 17 - Chassi do veículo
        fmt_decimal(val_unit, 3), # 18 - Valor unitário (3 casas)
        "",                       # 19 - Qtd inicial estoque
        "",                       # 20 - Valor inicial estoque
        cst_icms,                 # 21 - CST ICMS
        aliq_icms,                # 22 - Alíquota ICMS
        aliq_ipi,                 # 23 - Alíquota IPI
        "",                       # 24 - Periodicidade IPI
        "",                       # 25 - Observação
        "N",                      # 26 - Exporta DNF
        "",                       # 27 - Ex TIPI
        "", "", "", "", "",        # 28-32 DNF
        "", "",                   # 33-34 SE/DIC
        "N",                      # 35 - SCANC
        "", "", "",               # 36-38 SCANC
        "N",                      # 39 - GRF
        "",                       # 40 - GRF cód.
        "", "",                   # 41-42 DIEF
        "N",                      # 43 - 88ST
        "",                       # 44 - 88ST cód.
        "", "",                   # 45-46 GO
        "N",                      # 47 - GO produto rel.
        "N",                      # 48 - AM cesta
        "",                       # 49 - AM cód.
        "", "", "", "", "",        # 50-54 RS
        "",                       # 55 - RS grupo
        "N",                      # 56 - PR ECF
        "", "",                   # 57-58 MS/DF
        "",                       # 59 - DF item
        "",                       # 60 - PE tipo
        "N",                      # 61 - SP Cat 17/99
        "", "", "", "",           # 62-65 SP
        "", "",                   # 66-67 SPED gênero/serviço
        "", "", "", "", "", "",   # 68-73 SPED
        "",                       # 74 - SPED energia
        "",                       # 75 - Data cadastro
        "N",                      # 76 - LMC
        "", "",                   # 77-78 Combustível
        "N",                      # 79 - MP 540
        "",                       # 80 - Desc. complementar
        "", "",                   # 81-82 INSS/DACON
        "",                       # 83 - DACON crédito
        "",                       # 84 - Desconsiderar
        "", "", "", "",           # 85-88 SPED bloco K
        cest,                     # 89 - CEST
        "",                       # 90 - RE
        "",                       # 91 - Identificador
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
        "1000",        # 1  - Identificação
        especie,       # 2  - Código da espécie (36)
        cnpj_emit,     # 3  - Inscrição fornecedor
        "",            # 4  - Cód. exclusão DIEF
        acumulador,    # 5  - Código do acumulador (1157)
        cfop_first,    # 6  - CFOP
        "",            # 7  - Segmento
        nNF,           # 8  - Número do documento
        serie,         # 9  - Série
        "",            # 10 - Número doc final
        dhEmi,         # 11 - Data da entrada
        dhEmi,         # 12 - Data emissão
        v_nf,          # 13 - Valor contábil
        "",            # 14 - Valor exclusão DIEF
        obs_fisco,     # 15 - Observação (interesse do fisco)
        mod_frete,     # 16 - Modalidade do frete
        "T",           # 17 - Emitente (T=Terceiros)
        "",            # 18 - CFOP estendido (SE)
        "",            # 19 - Cód. transf. crédito (RS)
        "",            # 20 - Cód. recolh. ISS Retido
        "",            # 21 - Cód. recolh. IRRF
        "",            # 22 - Cód. observação
        "",            # 23 - Data visto (MG)
        "",            # 24 - Fato gerador CRF
        "",            # 25 - Fato gerador IRRF
        v_frete,       # 26 - Valor frete
        v_seg,         # 27 - Valor seguro
        v_outro,       # 28 - Valor despesas
        v_pis,         # 29 - Valor PIS
        "",            # 30 - Cód. antecipação tributária
        v_cofins,      # 31 - Valor COFINS
        "",            # 32 - Valor DARE (SE)
        "",            # 33 - Alíq. DARE (SE)
        "",            # 34 - Base cálc. ICMS ST
        "",            # 35 - Entradas isentas (MG)
        "",            # 36 - Outras entradas isentas (MG)
        "",            # 37 - Valor transp. base (MG)
        "",            # 38 - Cód. ressarcimento
        v_prod,        # 39 - Valor produtos
        c_mun_fg,      # 40 - Município Origem (cMunFG)
        "0",           # 41 - Situação da Nota (0=Regular)
        "",            # 42 - Cód. situação tributária
        "",            # 43 - Sub série
        ie_emit,       # 44 - IE fornecedor
        "",            # 45 - IM fornecedor
        "",            # 46 - Cód. operação e prestação
        "",            # 47 - Valor dedução receita
        "",            # 48 - Competência
        "",            # 49 - Operação (PA)
        "",            # 50 - Nº parecer fiscal
        "",            # 51 - Data parecer fiscal
        n_di,          # 52 - Nº declaração de importação
        "N",           # 53 - Possui benefício fiscal
        chave,         # 54 - Chave NF-e
        "",            # 55 - Cód. recolh. FETHAB
        "",            # 56 - Resp. recolh. FETHAB
        "",            # 57 - CFOP doc. fiscal
        "",            # 58 - Tipo CT-e
        "",            # 59 - CT-e referência
        "1",           # 60 - Modalidade importação (1=Com direito a crédito)
        "",            # 61 - Cód. inf. complementar
        "",            # 62 - Informação complementar
        "",            # 63 - Classe de consumo
        "",            # 64 - Tipo de ligação
        "",            # 65 - Grupo de tensão
        "",            # 66 - Tipo de assinante
        "",            # 67 - KWH consumido
        "",            # 68 - Valor energia/gás
        "",            # 69 - Valor cobrado terceiros
        "10",          # 70 - Tipo doc. importação (10=DI)
        "",            # 71 - Ato Concessório Drawback
        "",            # 72 - Natureza frete PIS/COFINS
        "",            # 73 - CST PIS/COFINS
        "",            # 74 - Base crédito PIS/COFINS
        "",            # 75 - Valor serviços PIS/COFINS
        "",            # 76 - Base cálc. PIS/COFINS
        "",            # 77 - Alíq. PIS
        "",            # 78 - Alíq. COFINS
        "",            # 79 - Chave NFSe
        "",            # 80 - Nº processo/ato concessório
        "",            # 81 - Origem do processo
        "",            # 82 - Data escrituração
        "",            # 83 - CFPS (DF)
        "",            # 84 - Natureza receita PIS/COFINS
        "",            # 85 - CST IPI
        "",            # 86 - Lançamentos SCP
        "",            # 87 - Tipo de serviço
        "",            # 88 - Município destino
        "",            # 89 - Pedágio
        v_ipi,         # 90 - IPI
        v_st,          # 91 - ICMS ST
        "",            # 92 - Classif. serviços EFD-Reinf
        "",            # 93 - Indicativo prestação EFD-Reinf
        "",            # 94 - Nº doc. arrecadação (RS)
        "",            # 95 - Tipo do título
        "",            # 96 - Identificação
        v_icms_d,      # 97 - ICMS Desonerado
        "",            # 98 - IPI Devolução
    ])


# ─────────────────────────────────────────────
# REGISTROS 1020 – Impostos (19 campos cada)
# Um registro por imposto presente na NF-e
# Códigos conforme tabela oficial Impostos.xls
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
            "1020",   # 1  - Identificação
            cod,      # 2  - Código do imposto
            perc_red, # 3  - % redução base cálculo
            base,     # 4  - Base de cálculo
            aliq,     # 5  - Alíquota
            valor,    # 6  - Valor do imposto
            isentas,  # 7  - Valor de isentas
            outras,   # 8  - Valor de outras
            v_ipi,    # 9  - Valor do IPI
            v_st,     # 10 - Valor da ST
            v_cont,   # 11 - Valor contábil
            cod_rec,  # 12 - Cód. recolhimento
            "",       # 13 - Valor não tributadas (GO)
            "",       # 14 - Valor parcela reduzida (GO)
            "",       # 15 - Alíq. Interestadual
            "",       # 16 - Nat. rendimentos
            "",       # 17 - Tipo de Dedução
            "",       # 18 - Tipo de Isenção
            "",       # 19 - Descrição
        ])

    # Totais do ICMSTot
    v_bc_icms    = get_text(total, "nfe:vBC")
    v_icms       = get_text(total, "nfe:vICMS")
    v_ipi_tot    = get_text(total, "nfe:vIPI")
    v_st_tot     = get_text(total, "nfe:vST")
    v_pis_tot    = get_text(total, "nfe:vPIS")
    v_cofins_tot = get_text(total, "nfe:vCOFINS")
    v_icms_deson = get_text(total, "nfe:vICMSDeson")

    # Alíquota média ICMS
    try:
        bc_f  = float(v_bc_icms)
        icm_f = float(v_icms)
        aliq_icms_med = fmt_decimal(str(icm_f / bc_f * 100)) if bc_f > 0 else ""
    except (ValueError, ZeroDivisionError):
        aliq_icms_med = ""

    # Soma bases por imposto percorrendo os itens
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

    # ── 1 – ICMS ──────────────────────────────────────────────────
    if v_icms and float(v_icms) > 0:
        linhas.append(r1020(1,
            base   = fmt_decimal(v_bc_icms),
            aliq   = aliq_icms_med,
            valor  = fmt_decimal(v_icms),
            v_ipi  = fmt_decimal(v_ipi_tot),
            v_st   = fmt_decimal(v_st_tot),
            v_cont = v_nf,
        ))

    # ── 2 – IPI ───────────────────────────────────────────────────
    if v_ipi_tot and float(v_ipi_tot) > 0:
        linhas.append(r1020(2,
            base   = fmt_decimal(str(bc_ipi_total)),
            valor  = fmt_decimal(v_ipi_tot),
            v_cont = v_nf,
        ))

    # ── 4 – PIS ───────────────────────────────────────────────────
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(r1020(4,
            base   = fmt_decimal(str(bc_pis_total)),
            valor  = fmt_decimal(v_pis_tot),
            v_cont = v_nf,
        ))

    # ── 5 – COFINS ────────────────────────────────────────────────
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(r1020(5,
            base   = fmt_decimal(str(bc_cof_total)),
            valor  = fmt_decimal(v_cofins_tot),
            v_cont = v_nf,
        ))

    # ── 45 – ICMS Importação (desonerado) ─────────────────────────
    if v_icms_deson and float(v_icms_deson) > 0:
        linhas.append(r1020(45,
            base   = fmt_decimal(v_bc_icms),
            aliq   = aliq_icms_med,
            valor  = fmt_decimal(v_icms_deson),
            v_cont = v_nf,
        ))

    # ── 133 – PIS Importação ──────────────────────────────────────
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(r1020(133,
            base   = fmt_decimal(str(bc_pis_total)),
            valor  = fmt_decimal(v_pis_tot),
            v_cont = v_nf,
        ))

    # ── 134 – COFINS Importação ───────────────────────────────────
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(r1020(134,
            base   = fmt_decimal(str(bc_cof_total)),
            valor  = fmt_decimal(v_cofins_tot),
            v_cont = v_nf,
        ))

    # ── 183 – IBS ─────────────────────────────────────────────────
    if ibs_t is not None:
        v_ibs_uf = get_text(ibs_t, "nfe:gIBS/nfe:gIBSUF/nfe:vIBSUF")
        bc_ibs   = get_text(ibs_t, "nfe:vBCIBSCBS")
        if v_ibs_uf and float(v_ibs_uf) > 0:
            linhas.append(r1020(183,
                base   = fmt_decimal(bc_ibs),
                valor  = fmt_decimal(v_ibs_uf),
                v_cont = v_nf,
            ))

    # ── 184 – CBS ─────────────────────────────────────────────────
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

    # ICMS
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
            for pt in ["PISAliq", "PISQtde", "PISNT", "PISOutr"]:
                pn = pis_node.find(f"nfe:{pt}", NS)
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
            for ct in ["COFINSAliq", "COFINSQtde", "COFINSNT", "COFINSOutr"]:
                cn = cof_node.find(f"nfe:{ct}", NS)
                if cn is not None:
                    v_cof    = fmt_decimal(get_text(cn, "nfe:vCOFINS"))
                    aliq_cof = fmt_decimal(get_text(cn, "nfe:pCOFINS"), 4)
                    cst_cof  = get_text(cn, "nfe:CST")
                    bc_cof   = fmt_decimal(get_text(cn, "nfe:vBC"))
                    break

        # IBS / CBS
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

    # Campo 4 = vProd + vIPI
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
        "1030",                  # 1
        cod_prod,                # 2
        qtd,                     # 3   - Quantidade (2 casas)
        v_total,                 # 4   - Valor total (Base Cal. + IPI)
        v_ipi,                   # 5   - Valor IPI
        fmt_decimal(v_prod),     # 6   - Base de cálculo
        "1",                     # 7   - Tipo Lançamento (1=vinculado à nota)
        d_di,                    # 8   - Data
        n_di,                    # 9   - Número DI
        cst_icms,                # 10  - CST ICMS
        fmt_decimal(v_prod),     # 11  - Valor bruto
        fmt_decimal(v_desc),     # 12  - Valor desconto
        v_bc_icms,               # 13  - Base cálculo ICMS
        v_bc_st,                 # 14  - Base cálculo ICMS ST
        aliq_icms,               # 15  - Alíquota ICMS
        "",                      # 16  - Produto Incentivado (PE)
        "",                      # 17  - Cód. apuração (PE)
        "",                      # 18  - Valor frete
        "",                      # 19  - Valor seguro
        fmt_decimal(v_outro),    # 20  - Despesas acessórias
        "",                      # 21  - Qtd gasolina
        v_icms,                  # 22  - Valor ICMS
        "",                      # 23  - Valor SUBTRI
        "",                      # 24  - Valor isentas IPI
        "",                      # 25  - Valor outras IPI
        "",                      # 26  - ICMS NFP
        fmt_decimal(v_unit, 6),  # 27  - Valor Unitário (6 casas)
        "",                      # 28  - Alíq. ST
        cst_ipi,                 # 29  - CST IPI
        aliq_ipi,                # 30  - Alíquota IPI
        "",                      # 31  - Base ISSQN
        "",                      # 32  - Alíquota ISSQN
        "",                      # 33  - Valor ISSQN
        cfop,                    # 34  - CFOP
        "",                      # 35  - Série ECF
        aliq_pis,                # 36  - Alíquota PIS (4 casas)
        v_pis,                   # 37  - Valor PIS
        aliq_cof,                # 38  - Alíquota COFINS (4 casas)
        v_cof,                   # 39  - Valor COFINS
        fmt_decimal(v_prod),     # 40  - Custo total
        cst_pis,                 # 41  - CST PIS
        bc_pis,                  # 42  - Base cálculo PIS
        cst_cof,                 # 43  - CST COFINS
        bc_cof,                  # 44  - Base cálculo COFINS
        "",                      # 45  - Chassi
        "",                      # 46  - Tipo operação veículo
        "",                      # 47  - Lote medicamento
        "",                      # 48  - Qtd lote
        "",                      # 49  - Data validade
        "",                      # 50  - Data fabricação
        "",                      # 51  - Ref. base cálculo
        "",                      # 52  - Valor tabelado
        "",                      # 53  - Nº série arma
        "",                      # 54  - Nº série cano
        "",                      # 55  - Enquadramento IPI
        "S",                     # 56  - Movimentação física
        unidade,                 # 57  - Unidade comercializada
        "",                      # 58  - Complemento CFOP
        "",                      # 59  - Tanque combustível
        fmt_decimal(v_prod),     # 60  - Valor contábil produto
        "",                      # 61  - Qtd trib. PIS unid. medida
        "",                      # 62  - Valor unit. PIS unid. medida
        "",                      # 63  - Valor PIS unid. medida
        "",                      # 64  - Qtd trib. COFINS unid. medida
        "",                      # 65  - Valor unit. COFINS unid. medida
        "",                      # 66  - Valor COFINS unid. medida
        "",                      # 67  - Base do crédito
        "",                      # 68  - Nº nota devolvida
        "",                      # 69  - Descrição complementar
        "",                      # 70  - Nota dev. CST PIS
        "",                      # 71  - Nota dev. CST COFINS
        "",                      # 72  - Vínculo crédito PIS
        "",                      # 73  - Vínculo crédito COFINS
        "",                      # 74  - Exclusão PIS
        "",                      # 75  - Exclusão COFINS
        "",                      # 76  - Base ICMS Carga Média
        "",                      # 77  - Alíq. ICMS Carga Média
        "",                      # 78  - Valor ICMS Carga Média
        "",                      # 79  - Nº série ECF devolvido
        "",                      # 80  - PIS/COFINS % redução
        "",                      # 81  - Cód. recolh. PIS dev.
        "",                      # 82  - Cód. recolh. COFINS dev.
        "",                      # 83  - Cód. recolh. PIS
        "",                      # 84  - Cód. recolh. COFINS
        "",                      # 85  - Créd. Presumido PIS
        "",                      # 86  - Créd. Presumido COFINS
        "",                      # 87  - ICMS ST Antec. Total Base
        "",                      # 88  - ICMS ST Antec. Total Alíq.
        "",                      # 89  - ICMS ST Antec. Total Valor
        "",                      # 90  - Cód. recolh. IPI
        cest,                    # 91  - Código CEST
        "",                      # 92  - ICMS ST Retido Base
        "",                      # 93  - ICMS ST Retido Valor
        "",                      # 94  - ICMS ST Retido tag XML
        "",                      # 95  - Identificador
        "",                      # 96  - ICMS Próprio Substituto
        v_icms_des,              # 97  - Valor Desonerado
        "",                      # 98  - Código (motDesICMS)
        "",                      # 99  - ICMS Não creditado
        "",                      # 100 - ICMS Monofásico Qtde
        "",                      # 101 - ICMS Monofásico Alíq.
        "",                      # 102 - ICMS Monofásico Valor
        "",                      # 103 - ICMS Monofásico FCV
        ibs_class_trib,          # 104 - IBS cClassTrib
        ibs_bc,                  # 105 - IBS Base cálculo
        ibs_aliq,                # 106 - IBS Alíquota
        ibs_val,                 # 107 - IBS Valor
        ibs_class_trib,          # 108 - CBS cClassTrib
        cbs_bc,                  # 109 - CBS Base cálculo
        cbs_aliq,                # 110 - CBS Alíquota
        cbs_val,                 # 111 - CBS Valor
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
) -> tuple:

    try:
        root = ET.fromstring(xml_content)
    except ET.ParseError as e:
        return "", {"erro": str(e)}

    # Suporte a nfeProc ou NFe direto
    nfe = root.find("nfe:NFe", NS)
    if nfe is None:
        nfe = root

    # ── CNPJ: XML → fallback manual ───────────────────────────────
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

    # ── Registros de cadastro ─────────────────────────────────────
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

    # ── Registro 1000 ─────────────────────────────────────────────
    lines.append(gerar_registro_1000(nfe, cnpj_empresa, acumulador, especie))

    # ── Registros 1020 ────────────────────────────────────────────
    for r in gerar_registros_1020(nfe):
        lines.append(r)

    # ── Registros 1030 ────────────────────────────────────────────
    for seq, det in enumerate(det_list, start=1):
        lines.append(gerar_registro_1030(det, seq))

    # ── Registros 1150 / 1151 (IBS/CBS agrupados por cClassTrib) ──
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
            ct,
            fmt_decimal(str(d["bc_ibs"])),
            fmt_decimal(d["aliq_ibs"]),
            fmt_decimal(str(d["v_ibs"])),
        ))
        lines.append(gerar_registro_1151(
            ct,
            fmt_decimal(str(d["bc_cbs"])),
            fmt_decimal(d["aliq_cbs"]),
            fmt_decimal(str(d["v_cbs"])),
        ))

    return "\n".join(lines), resumo


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"### ℹ Sobre")
    st.markdown(f"**Versao:** {VERSAO}")
    st.markdown("**Thomson Reuters**")
    st.markdown("**Dominio Sistemas**")
    st.markdown("---")
    st.markdown("### ⚙ Parametros")
    st.markdown(
        "<div style='font-size:12px;color:#DDDDDD;margin-bottom:4px;'>"
        "O CNPJ e lido automaticamente do XML.<br>"
        "Use o campo abaixo apenas como <b>fallback</b>.</div>",
        unsafe_allow_html=True,
    )
    cnpj_fallback = st.text_input(
        "CNPJ Fallback (opcional)", value="", max_chars=14,
        help="Usado somente se o XML nao contiver CNPJ do destinatario.",
    )
    acumulador = st.text_input("Codigo do Acumulador", value="1157",
                               help="1157 = CFOP 3102 (Importacao)")
    especie    = st.text_input("Codigo da Especie", value="36",
                               help="36 = NF-e modelo 55")
    st.markdown("---")
    st.markdown("### 📋 Registros de cadastro")
    inc_0000 = st.checkbox("0000 - Identificacao da empresa", value=True)
    inc_0020 = st.checkbox("0020 - Fornecedor (emitente)", value=True)
    inc_0100 = st.checkbox("0100 - Produtos", value=True)
    st.caption("1000 / 1020 / 1030 / 1150 / 1151 sempre gerados.")
    st.markdown("---")
    st.markdown("### ⚙ Encoding")
    st.markdown("**Entrada:** UTF-8 / Latin-1")
    st.markdown("**Saida:** ANSI (Latin-1)")
    st.markdown("---")
    with st.expander("📋 Tabela de Impostos Dominio"):
        for cod, nome in sorted(TABELA_IMPOSTOS.items()):
            st.caption(f"`{cod:3d}` - {nome}")


# ─────────────────────────────────────────────
# INSTRUÇÕES
# ─────────────────────────────────────────────
with st.expander("📖 Instrucoes de Uso — clique para expandir", expanded=False):
    st.markdown(
        """
        <div class="instrucoes-box">
        <h4>🔹 Como o CNPJ e obtido</h4>
        <p>Lido automaticamente do XML na seguinte prioridade:<br>
        1. <code>&lt;dest&gt;&lt;CNPJ&gt;</code> — NF-e domestica<br>
        2. <code>&lt;dest&gt;&lt;CPF&gt;</code> — Pessoa fisica<br>
        3. <code>&lt;dest&gt;&lt;idEstrangeiro&gt;</code> — Exterior com valor<br>
        4. <code>&lt;emit&gt;&lt;CNPJ&gt;</code> — <b>NF-e de importacao</b>
           (idEstrangeiro vazio / UF=EX)<br>
        5. Campo <b>CNPJ Fallback</b> na sidebar — ultimo recurso</p>
        <h4>🔹 Passo 1 — Parametros</h4>
        <p>Confirme o <b>Acumulador</b> (1157) e a <b>Especie</b> (36) na sidebar.</p>
        <h4>🔹 Passo 2 — Upload dos XMLs</h4>
        <p>Faca o upload de um ou mais arquivos XML de NF-e.</p>
        <h4>🔹 Passo 3 — Verificar resumo</h4>
        <p>Confira o <b>CNPJ Empresa</b> e a <b>Origem CNPJ</b> na tabela de resumo.</p>
        <h4>🔹 Passo 4 — Baixar e importar</h4>
        <p>Clique em <b>Baixar Arquivo Dominio (.TXT ANSI)</b> e importe no
        modulo fiscal do Dominio Sistemas.</p>
        <hr>
        <h4>⚠ Observacoes</h4>
        <ul>
        <li>Saida em <b>ANSI (Latin-1)</b> — padrao Dominio Sistemas.</li>
        <li>Registros: <b>0000, 0020, 0100, 1000, 1020, 1030, 1150, 1151</b>.</li>
        <li>1020 gerado por imposto: <b>1-ICMS, 2-IPI, 4-PIS, 5-COFINS,
            45-ICMS Importacao, 133-PIS Imp., 134-COFINS Imp., 183-IBS, 184-CBS</b>.</li>
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

    if erros:
        st.error("❌ Erros encontrados:")
        st.dataframe(erros, use_container_width=True)

    if all_resumos:
        st.success(f"✅ {len(all_resumos)} arquivo(s) convertido(s) com sucesso!")

        # Badge CNPJ
        cnpjs_unicos = list({r["CNPJ Empresa"]: r for r in all_resumos}.values())
        if cnpjs_unicos:
            st.markdown("#### 🏢 Empresa Identificada (lida do XML)")
            cols = st.columns(min(len(cnpjs_unicos), 4))
            for idx, r in enumerate(cnpjs_unicos[:4]):
                is_ext = r.get("Exterior", "Nao") == "Sim"
                cor    = "#1565C0" if is_ext else "#FF8000"
                with cols[idx]:
                    st.markdown(
                        f"""
                        <div class="cnpj-badge"
                             style="border-color:{cor};color:{cor};">
                            CNPJ: {r['CNPJ Empresa']}
                        </div>
                        <div class="info-origem"
                             style="border-left-color:{cor};">
                            📌 {r.get('Origem CNPJ','')}<br>
                            🏷️ {(r.get('Emitente','') if is_ext
                                  else r.get('Destinatario',''))[:60]}<br>
                            {'🌍 <b>Operacao de Importacao/Exterior</b>' if is_ext else ''}
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
                label="⬇ Baixar Arquivo Dominio (.TXT ANSI)",
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

        st.markdown("---")
        st.markdown("#### 📈 Estatisticas")
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
            c5.metric("Total NF","—")

else:
    st.info("👆 Faca o upload de um ou mais arquivos XML de NF-e para iniciar.")

# ─────────────────────────────────────────────
# RODAPÉ
# ─────────────────────────────────────────────
st.markdown("---")
st.caption(
    f"Conversor XML NF-e → Dominio Sistemas | "
    f"Thomson Reuters | Python + Streamlit | {VERSAO}"
)
