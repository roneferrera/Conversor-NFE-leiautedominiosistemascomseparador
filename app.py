import streamlit as st
import xml.etree.ElementTree as ET
from datetime import datetime
import re

VERSAO = "V1.7"

# ─────────────────────────────────────────────
# TEMA
# ─────────────────────────────────────────────
def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] {
            font-family: 'Segoe UI', 'Arial', sans-serif; color: #444444;
        }
        h1, h2, h3 { color: #FF8000; font-weight: 700; }
        section[data-testid="stSidebar"] { background-color: #444444; color: #FFFFFF; }
        section[data-testid="stSidebar"] * { color: #FFFFFF !important; }
        .stButton > button { background-color: #FF8000; color: #FFFFFF; border: none; border-radius: 4px; font-weight: bold; }
        .stButton > button:hover { background-color: #D64001; }
        .stDownloadButton > button { background-color: #FF8000; color: #FFFFFF; border: none; border-radius: 4px; font-weight: bold; }
        .stDownloadButton > button:hover { background-color: #D64001; }
        hr { border-color: #FF8000; }
        [data-testid="metric-container"] { background-color: #E9E9E9; border-left: 4px solid #FF8000; border-radius: 4px; padding: 10px; }
        .instrucoes-box { background-color: #E9E9E9; border-left: 4px solid #FF8000; border-radius: 4px; padding: 16px 20px; margin: 12px 0; color: #444444; }
        .instrucoes-box h4 { color: #FF8000; margin-top: 14px; margin-bottom: 6px; }
        .instrucoes-box h4:first-child { margin-top: 0; }
        .cnpj-badge { background-color: #444444; border: 1px solid #FF8000; border-radius: 4px; padding: 6px 12px; font-family: Consolas, monospace; font-size: 13px; display: inline-block; margin: 4px 0; }
        .info-origem { background-color: #FFF3E0; border-left: 3px solid #FF8000; border-radius: 3px; padding: 8px 12px; font-size: 12px; color: #444444; margin: 6px 0; }
        </style>
    """, unsafe_allow_html=True)

st.set_page_config(page_title="Dominio Sistemas | Thomson Reuters", page_icon="🟠", layout="wide", initial_sidebar_state="expanded")
apply_tr_theme()

st.markdown(f"""
    <div style="background:#444444; padding:24px 28px 18px 28px; border-radius:8px; border-top:6px solid #FF8000; margin-bottom:28px;">
        <h2 style="color:#FF8000; margin:0; font-family:'Segoe UI',Arial,sans-serif;">
            🔄 Conversor XML NF-e → Dominio Sistemas &nbsp;|&nbsp; {VERSAO}
        </h2>
        <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
            Converte XML de NF-e para leiaute padrao de importacao do <strong>Dominio Sistemas</strong>.
            Saida em <strong>ANSI (Latin-1)</strong>.
        </p>
    </div>
""", unsafe_allow_html=True)

NS = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

TABELA_GRUPOS = {
    0:"Automatico (por CFOP/NCM)", 1:"GERAL", 2:"MERCADORIA PARA REVENDA",
    3:"MATERIA PRIMA", 4:"EMBALAGENS", 5:"PRODUTO EM PROCESSO",
    6:"PRODUTO ACABADO", 7:"SUBPRODUTO", 8:"PRODUTOS INTERMEDIARIOS",
    9:"MATERIAL DE USO E CONSUMO", 10:"ATIVO IMOBILIZADO",
    11:"SERVICOS", 12:"OUTROS INSUMOS",
    100:"PRODUTOS NOVOS - ENTRADAS", 500:"PRODUTOS NOVOS - SAIDAS",
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
        return datetime.strptime(iso_date[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
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

def somente_numeros(valor: str) -> str:
    return re.sub(r"[^0-9]", "", valor or "")

# ─────────────────────────────────────────────
# DETECÇÃO DE GRUPO
# ─────────────────────────────────────────────
def get_grupo_por_cfop(cfop: str) -> int:
    if not cfop:
        return 1
    dois = cfop[:2]
    mapa = {
        "11": 2, "12": 2, "13": 3, "14": 3, "15": 9, "16": 10, "17": 11,
        "20": 2, "21": 2, "22": 3, "25": 9, "30": 2, "31": 3, "35": 9,
        "40": 2, "41": 2, "55": 12, "60": 2,
    }
    return mapa.get(dois, 1)

def get_grupo_por_ncm(ncm: str) -> int:
    if not ncm or len(ncm) < 2:
        return 1
    cap = ncm[:2]
    mapa = {
        "84": 10, "85": 10, "86": 10, "87": 10, "88": 10, "89": 10, "90": 10, "91": 10, "94": 10,
        "28": 3, "29": 3, "30": 3, "31": 3, "32": 3, "33": 3, "34": 3, "38": 3,
        "39": 3, "40": 3, "44": 3, "47": 3, "48": 3, "72": 3, "73": 3, "74": 3,
        "75": 3, "76": 3, "82": 3, "83": 3,
        "01": 2, "02": 2, "03": 2, "04": 2, "07": 2, "08": 2, "09": 2, "10": 2,
        "16": 2, "17": 2, "18": 2, "19": 2, "20": 2, "21": 2, "22": 2, "27": 2,
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
# DETECÇÃO SE É IMPORTAÇÃO
# ─────────────────────────────────────────────
def is_nota_importacao(nfe_root) -> bool:
    dest = nfe_root.find("nfe:infNFe/nfe:dest", NS)
    if dest is not None:
        ender = dest.find("nfe:enderDest", NS)
        if ender is not None and get_text(ender, "nfe:UF") == "EX":
            return True
    det_list = nfe_root.findall("nfe:infNFe/nfe:det", NS)
    if det_list:
        cfop = get_text(det_list[0].find("nfe:prod", NS), "nfe:CFOP")
        if cfop.startswith("3"):
            return True
    return False

# ─────────────────────────────────────────────
# EXTRAÇÃO CNPJ DESTINATÁRIO
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
        ender_dest  = dest.find("nfe:enderDest", NS)
        uf_dest     = get_text(ender_dest, "nfe:UF") if ender_dest is not None else ""
        id_ext_node = dest.find("nfe:idEstrangeiro", NS)
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
# REGISTRO 0000
# ─────────────────────────────────────────────
def gerar_registro_0000(cnpj_empresa: str) -> str:
    return pipe_join(["0000", cnpj_empresa])

# ─────────────────────────────────────────────
# REGISTRO 0020 – Cadastro de Fornecedores
# V1.7: NF importação → dados do <dest> (fornecedor estrangeiro)
#        NF doméstica  → dados do <emit>
# ─────────────────────────────────────────────
def gerar_registro_0020(emit, dest=None, is_importacao: bool = False) -> str:
    if is_importacao and dest is not None:
        # ── Fornecedor ESTRANGEIRO: dados do <dest> ──────────────
        razao    = get_text(dest, "nfe:xNome")[:150]
        fantasia = ""
        ender    = dest.find("nfe:enderDest", NS)

        logradouro  = get_text(ender, "nfe:xLgr")                         if ender is not None else ""
        numero      = somente_numeros(get_text(ender, "nfe:nro"))          if ender is not None else ""
        complemento = ""
        bairro      = get_text(ender, "nfe:xBairro")                      if ender is not None else ""
        cod_mun     = somente_numeros(get_text(ender, "nfe:cMun"))         if ender is not None else ""
        cep         = get_text(ender, "nfe:CEP")                          if ender is not None else ""
        cod_pais    = somente_numeros(get_text(ender, "nfe:cPais"))        if ender is not None else ""

        # Exterior: inscrição vazia, UF = EX
        inscricao = ""
        uf_campo  = "EX"
        ie        = ""
        regime    = "N"
        contrib   = "N"

    else:
        # ── Fornecedor DOMÉSTICO: dados do <emit> ────────────────
        inscricao   = get_text(emit, "nfe:CNPJ")
        razao       = get_text(emit, "nfe:xNome")[:150]
        fantasia    = get_text(emit, "nfe:xFant")[:40]
        ender       = emit.find("nfe:enderEmit", NS)

        logradouro  = get_text(ender, "nfe:xLgr")                         if ender is not None else ""
        numero      = somente_numeros(get_text(ender, "nfe:nro"))          if ender is not None else ""
        complemento = get_text(ender, "nfe:xCpl")                         if ender is not None else ""
        bairro      = get_text(ender, "nfe:xBairro")                      if ender is not None else ""
        cod_mun     = somente_numeros(get_text(ender, "nfe:cMun"))         if ender is not None else ""
        cep         = get_text(ender, "nfe:CEP")                          if ender is not None else ""
        uf_campo    = get_text(ender, "nfe:UF")                           if ender is not None else ""
        cod_pais    = ""

        ie          = get_text(emit, "nfe:IE")
        crt         = get_text(emit, "nfe:CRT")
        regime_map  = {"1": "M", "2": "E", "3": "N"}
        regime      = regime_map.get(crt, "N")
        contrib     = "S" if ie and ie.upper() not in ("ISENTO", "NAO CONTRIBUINTE", "") else "N"

    return pipe_join([
        "0020",       # 1  - Identificacao
        inscricao,    # 2  - Inscricao (vazio quando exterior) ← V1.7
        razao,        # 3  - Razao Social (HILLROM quando importacao) ← V1.7
        fantasia,     # 4  - Apelido
        logradouro,   # 5  - Endereco (STATE ROUTE 46 EAST) ← V1.7
        numero,       # 6  - Numero (Numerico)
        complemento,  # 7  - Complemento
        bairro,       # 8  - Bairro (MAIL CODE J36) ← V1.7
        cod_mun,      # 9  - Codigo municipio (9999999) ← V1.7
        uf_campo,     # 10 - UF (EX) ← V1.7
        cod_pais,     # 11 - Codigo Pais (2496) ← V1.7
        cep,          # 12 - CEP
        ie,           # 13 - IE
        "",           # 14 - IM
        "",           # 15 - Suframa
        "",           # 16 - DDD
        "",           # 17 - Telefone
        "",           # 18 - FAX
        "",           # 19 - Data cadastro
        "",           # 20 - Conta contabil
        "",           # 21 - Conta contabil cliente
        "N",          # 22 - Agropecuario
        "7",          # 23 - Natureza juridica
        regime,       # 24 - Regime apuracao
        contrib,      # 25 - Contribuinte ICMS
        "",           # 26 - Aliquota ICMS
        "",           # 27 - Categoria estabelecimento
        "",           # 28 - IE ST
        "",           # 29 - Email
        "N",          # 30 - Interdependencia
        "N",          # 31 - Contribuinte CPRB
        "",           # 32 - Processo
        "",           # 33 - Tipo Inscricao
    ])

# ─────────────────────────────────────────────
# REGISTRO 0100 – Cadastro de produtos (91 campos)
# V1.7: campo 24 = "M" (Periodicidade IPI Mensal)
#        campo 54 = "" (RS-MVA — Decimal, NAO "N")
#        pad automático garante sempre 92 posições
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

    campos = [
        "0100",                   # 1
        cod_prod,                 # 2
        descricao,                # 3
        "",                       # 4  - NBM
        ncm,                      # 5  - NCM
        "",                       # 6  - NCM Exterior (Numerico)
        "",                       # 7  - Codigo barras
        "",                       # 8  - Cod. imposto importacao (Numerico)
        cod_grupo,                # 9  - Grupo (Numerico)
        unidade,                  # 10 - Unidade
        "N",                      # 11 - Unidade inventaria diferente
        "O",                      # 12 - Tipo produto
        "",                       # 13 - Tipo arma (Numerico)
        "",                       # 14 - Descricao arma
        "",                       # 15 - Tipo medicamento (Numerico)
        "N",                      # 16 - ISSQN
        "",                       # 17 - Chassi
        fmt_decimal(val_unit, 3), # 18 - Valor unitario (3 casas)
        "",                       # 19 - Qtd inicial estoque
        "",                       # 20 - Valor inicial estoque
        cst_icms,                 # 21 - CST ICMS (Numerico)
        aliq_icms,                # 22 - Aliq. ICMS (2 casas)
        aliq_ipi,                 # 23 - Aliq. IPI (2 casas)
        "M",                      # 24 - Periodicidade IPI ← V1.7 fixo "M"
        "",                       # 25 - Observacao
        "N",                      # 26 - Exporta DNF
        "",                       # 27 - Ex TIPI
        "",                       # 28 - DNF especie (Numerico)
        "",                       # 29 - DNF unidade (Numerico)
        "",                       # 30 - DNF fator conversao
        "",                       # 31 - DNF cod produto (Numerico)
        "",                       # 32 - DNF cap. volumetrica (Numerico)
        "",                       # 33 - SE/DIC EAN
        "",                       # 34 - SE/DIC cod produto (Numerico)
        "N",                      # 35 - SCANC
        "",                       # 36 - SCANC cod produto (Numerico)
        "",                       # 37 - SCANC gasolina A
        "",                       # 38 - SCANC tipo produto
        "N",                      # 39 - GRF
        "",                       # 40 - GRF cod produto (Numerico)
        "",                       # 41 - DIEF unidade
        "",                       # 42 - DIEF tipo produto (Numerico)
        "N",                      # 43 - 88ST
        "",                       # 44 - 88ST cod produto (Numerico)
        "",                       # 45 - GO complementar IPM
        "",                       # 46 - GO cod produto IPM (Numerico)
        "N",                      # 47 - GO produto relacionado
        "N",                      # 48 - AM cesta basica
        "",                       # 49 - AM cod produto DAM (Numerico)
        "",                       # 50 - RS produto incluido ST
        "",                       # 51 - RS data inicio ST (Data)
        "",                       # 52 - RS produto preco tabelado
        "",                       # 53 - RS valor unitario ST (Decimal 2)
        "",                       # 54 - RS MVA ST (Decimal 2) ← V1.7 "" NAO "N"
        "",                       # 55 - RS grupo ST (Numerico)
        "N",                      # 56 - PR equipamento ECF
        "",                       # 57 - MS incentivo fiscal (Numerico)
        "",                       # 58 - DF regime especial (Numerico)
        "",                       # 59 - DF item padrao regime (Numerico)
        "",                       # 60 - PE tipo produto (Numerico)
        "N",                      # 61 - SP Cat 17/99
        "",                       # 62 - SP data saldo inicial (Data)
        "",                       # 63 - SP valor unitario (3 casas)
        "",                       # 64 - SP quantidade (3 casas)
        "",                       # 65 - SP valor final (2 casas)
        "",                       # 66 - SPED genero (Numerico)
        "",                       # 67 - SPED cod servico (Numerico)
        "",                       # 68 - SPED tipo item (Numerico)
        "",                       # 69 - SPED classificacao (Numerico)
        "",                       # 70 - SPED conta em seu poder (Numerico)
        "",                       # 71 - SPED conta terceiros (Numerico)
        "",                       # 72 - SPED conta de terceiros (Numerico)
        "",                       # 73 - SPED tipo receita
        "",                       # 74 - SPED energia (Numerico)
        "",                       # 75 - Data cadastro (Data)
        "N",                      # 76 - LMC
        "",                       # 77 - Cod combustivel DF
        "",                       # 78 - Cod combustivel ANP
        "N",                      # 79 - MP 540
        "",                       # 80 - Desc complementar
        "",                       # 81 - INSS folha
        "",                       # 82 - DACON tipo produto
        "",                       # 83 - DACON credito presumido (Numerico)
        "",                       # 84 - Desconsiderar (Numerico)
        "",                       # 85 - SPED bloco K em processo (Numerico)
        "",                       # 86 - SPED bloco K hist processo (Numerico)
        "",                       # 87 - SPED bloco K acabado (Numerico)
        "",                       # 88 - SPED bloco K hist acabado (Numerico)
        cest,                     # 89 - CEST (Numerico)
        "",                       # 90 - RE (Numerico)
        "",                       # 91 - Identificador
    ]

    # Garante exatamente 92 posições sem crash em producao
    while len(campos) < 92:
        campos.append("")
    campos = campos[:92]
    return pipe_join(campos)

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

    # Campo 52 = Número DI no 1000 é Caractere — mantém original
    n_di = ""
    if det_list:
        di_node = det_list[0].find("nfe:prod/nfe:DI", NS)
        if di_node is not None:
            n_di = get_text(di_node, "nfe:nDI")

    return pipe_join([
        "1000",        # 1
        especie,       # 2
        cnpj_emit,     # 3  - Inscricao FORNECEDOR (emitente da NF)
        "",            # 4
        acumulador,    # 5
        cfop_first,    # 6
        "",            # 7
        nNF,           # 8
        serie,         # 9
        "",            # 10
        dhEmi,         # 11
        dhEmi,         # 12
        v_nf,          # 13
        "",            # 14
        obs_fisco,     # 15
        mod_frete,     # 16
        "T",           # 17
        "",            # 18
        "",            # 19
        "",            # 20
        "",            # 21
        "",            # 22
        "",            # 23
        "",            # 24
        "",            # 25
        v_frete,       # 26
        v_seg,         # 27
        v_outro,       # 28
        v_pis,         # 29
        "",            # 30
        v_cofins,      # 31
        "",            # 32
        "",            # 33
        "",            # 34
        "",            # 35
        "",            # 36
        "",            # 37
        "",            # 38
        v_prod,        # 39
        c_mun_fg,      # 40
        "0",           # 41
        "",            # 42
        "",            # 43
        ie_emit,       # 44
        "",            # 45
        "",            # 46
        "",            # 47
        "",            # 48
        "",            # 49
        "",            # 50
        "",            # 51
        n_di,          # 52 - Numero DI (Caractere no 1000)
        "N",           # 53
        chave,         # 54
        "",            # 55
        "",            # 56
        "",            # 57
        "",            # 58
        "",            # 59
        "1",           # 60
        "",            # 61
        "",            # 62
        "",            # 63
        "",            # 64
        "",            # 65
        "",            # 66
        "",            # 67
        "",            # 68
        "",            # 69
        "10",          # 70
        "",            # 71
        "",            # 72
        "",            # 73
        "",            # 74
        "",            # 75
        "",            # 76
        "",            # 77
        "",            # 78
        "",            # 79
        "",            # 80
        "",            # 81
        "",            # 82
        "",            # 83
        "",            # 84
        "",            # 85
        "",            # 86
        "",            # 87
        "",            # 88
        "",            # 89
        v_ipi,         # 90
        v_st,          # 91
        "",            # 92
        "",            # 93
        "",            # 94
        "",            # 95
        "",            # 96
        v_icms_d,      # 97
        "",            # 98
    ])

# ─────────────────────────────────────────────
# REGISTROS 1010 / 1015
# ─────────────────────────────────────────────
def gerar_registros_1010(nfe_root) -> list:
    linhas   = []
    inf_adic = nfe_root.find("nfe:infNFe/nfe:infAdic", NS)
    if inf_adic is None:
        return linhas
    for txt, cod in [
        (get_text(inf_adic, "nfe:infAdFisco"), "1"),
        (get_text(inf_adic, "nfe:infCpl"),     "2"),
    ]:
        if txt:
            for bloco in [txt[i:i+300] for i in range(0, len(txt), 300)]:
                linhas.append(pipe_join(["1010", cod, bloco]))
    return linhas

def gerar_registros_1015(nfe_root) -> list:
    linhas   = []
    inf_adic = nfe_root.find("nfe:infNFe/nfe:infAdic", NS)
    if inf_adic is None:
        return linhas
    for txt, cod in [
        (get_text(inf_adic, "nfe:infAdFisco"), "1"),
        (get_text(inf_adic, "nfe:infCpl"),     "2"),
    ]:
        if txt:
            for bloco in [txt[i:i+300] for i in range(0, len(txt), 300)]:
                linhas.append(pipe_join(["1015", cod, bloco]))
    return linhas

# ─────────────────────────────────────────────
# REGISTROS 1020 – Impostos
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
    bc_ipi_total = bc_pis_total = bc_cof_total = 0.0

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

    if v_icms and float(v_icms) > 0:
        linhas.append(r1020(1, base=fmt_decimal(v_bc_icms), aliq=aliq_icms_med,
            valor=fmt_decimal(v_icms), v_ipi=fmt_decimal(v_ipi_tot),
            v_st=fmt_decimal(v_st_tot), v_cont=v_nf))
    if v_ipi_tot and float(v_ipi_tot) > 0:
        linhas.append(r1020(2, base=fmt_decimal(str(bc_ipi_total)),
            valor=fmt_decimal(v_ipi_tot), v_cont=v_nf))
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(r1020(4, base=fmt_decimal(str(bc_pis_total)),
            valor=fmt_decimal(v_pis_tot), v_cont=v_nf))
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(r1020(5, base=fmt_decimal(str(bc_cof_total)),
            valor=fmt_decimal(v_cofins_tot), v_cont=v_nf))
    if v_icms_deson and float(v_icms_deson) > 0:
        linhas.append(r1020(45, base=fmt_decimal(v_bc_icms), aliq=aliq_icms_med,
            valor=fmt_decimal(v_icms_deson), v_cont=v_nf))
    if v_pis_tot and float(v_pis_tot) > 0:
        linhas.append(r1020(133, base=fmt_decimal(str(bc_pis_total)),
            valor=fmt_decimal(v_pis_tot), v_cont=v_nf))
    if v_cofins_tot and float(v_cofins_tot) > 0:
        linhas.append(r1020(134, base=fmt_decimal(str(bc_cof_total)),
            valor=fmt_decimal(v_cofins_tot), v_cont=v_nf))
    if ibs_t is not None:
        v_ibs_uf = get_text(ibs_t, "nfe:gIBS/nfe:gIBSUF/nfe:vIBSUF")
        bc_ibs   = get_text(ibs_t, "nfe:vBCIBSCBS")
        if v_ibs_uf and float(v_ibs_uf) > 0:
            linhas.append(r1020(183, base=fmt_decimal(bc_ibs),
                valor=fmt_decimal(v_ibs_uf), v_cont=v_nf))
        v_cbs  = get_text(ibs_t, "nfe:gCBS/nfe:vCBS")
        bc_cbs = get_text(ibs_t, "nfe:vBCIBSCBS")
        if v_cbs and float(v_cbs) > 0:
            linhas.append(r1020(184, base=fmt_decimal(bc_cbs),
                valor=fmt_decimal(v_cbs), v_cont=v_nf))
    return linhas

# ─────────────────────────────────────────────
# REGISTRO 1030 – Estoque / item (111 campos)
# V1.7: campo 9  (Numero DI)    → somente_numeros() ← corrige "26BR5021449"
#        campo 98 (motDesICMS)   → lido da tag XML <motDesICMS>
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

    di_node  = prod.find("nfe:DI", NS)
    # Campo 9 = Numero DI → tipo Decimal → somente numeros ← V1.7
    n_di     = somente_numeros(get_text(di_node, "nfe:nDI")) if di_node is not None else ""
    d_di     = fmt_date(get_text(di_node, "nfe:dDI"))        if di_node is not None else ""

    # Defaults
    icms_node  = None
    v_bc_icms  = aliq_icms = v_icms = cst_icms = v_icms_des = v_bc_st = ""
    mot_des    = ""   # ← V1.7: motDesICMS
    v_ipi = aliq_ipi = cst_ipi = ""
    v_pis = aliq_pis = cst_pis = bc_pis = ""
    v_cof = aliq_cof = cst_cof = bc_cof = ""
    ibs_class_trib = ibs_bc = ibs_aliq = ibs_val = ""
    cbs_bc = cbs_aliq = cbs_val = ""

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
            mot_des    = get_text(icms_node, "nfe:motDesICMS")  # ← V1.7

        ipi_trib = imposto.find("nfe:IPI/nfe:IPITrib", NS)
        ipi_nt   = imposto.find("nfe:IPI/nfe:IPINT", NS)
        if ipi_trib is not None:
            v_ipi    = fmt_decimal(get_text(ipi_trib, "nfe:vIPI"))
            aliq_ipi = fmt_decimal(get_text(ipi_trib, "nfe:pIPI"))
            cst_ipi  = get_text(ipi_trib, "nfe:CST")
        elif ipi_nt is not None:
            v_ipi = "0,00"; aliq_ipi = "0,00"
            cst_ipi = get_text(ipi_nt, "nfe:CST")

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

        ibs_node = imposto.find("nfe:IBSCBS", NS)
        if ibs_node is not None:
            ibs_class_trib = get_text(ibs_node, "nfe:cClassTrib")
            gibs = ibs_node.find("nfe:gIBSCBS", NS)
            if gibs is not None:
                ibs_bc = fmt_decimal(get_text(gibs, "nfe:vBC"))
                guf = gibs.find("nfe:gIBSUF", NS)
                if guf is not None:
                    ibs_aliq = fmt_decimal(get_text(guf, "nfe:pIBSUF"))
                    ibs_val  = fmt_decimal(get_text(guf, "nfe:vIBSUF"))
                gcbs = gibs.find("nfe:gCBS", NS)
                if gcbs is not None:
                    cbs_bc   = ibs_bc
                    cbs_aliq = fmt_decimal(get_text(gcbs, "nfe:pCBS"))
                    cbs_val  = fmt_decimal(get_text(gcbs, "nfe:vCBS"))

    try:
        vp = float(v_prod or "0")
        vi = float(get_text(
            imposto.find("nfe:IPI/nfe:IPITrib", NS) if imposto else None,
            "nfe:vIPI") or "0") if imposto else 0.0
        v_total = fmt_decimal(str(vp + vi))
    except (ValueError, TypeError):
        v_total = fmt_decimal(v_prod)

    return pipe_join([
        "1030",                  # 1
        cod_prod,                # 2
        qtd,                     # 3
        v_total,                 # 4
        v_ipi,                   # 5
        fmt_decimal(v_prod),     # 6
        "1",                     # 7
        d_di,                    # 8
        n_di,                    # 9  - DI somente numeros ← V1.7
        cst_icms,                # 10
        fmt_decimal(v_prod),     # 11
        fmt_decimal(v_desc),     # 12
        v_bc_icms,               # 13
        v_bc_st,                 # 14
        aliq_icms,               # 15
        "",                      # 16
        "",                      # 17
        "",                      # 18
        "",                      # 19
        fmt_decimal(v_outro),    # 20
        "",                      # 21
        v_icms,                  # 22
        "",                      # 23
        "",                      # 24
        "",                      # 25
        "",                      # 26
        fmt_decimal(v_unit, 6),  # 27
        "",                      # 28
        cst_ipi,                 # 29
        aliq_ipi,                # 30
        "",                      # 31
        "",                      # 32
        "",                      # 33
        cfop,                    # 34
        "",                      # 35
        aliq_pis,                # 36
        v_pis,                   # 37
        aliq_cof,                # 38
        v_cof,                   # 39
        fmt_decimal(v_prod),     # 40
        cst_pis,                 # 41
        bc_pis,                  # 42
        cst_cof,                 # 43
        bc_cof,                  # 44
        "",                      # 45
        "",                      # 46
        "",                      # 47
        "",                      # 48
        "",                      # 49
        "",                      # 50
        "",                      # 51
        "",                      # 52
        "",                      # 53
        "",                      # 54
        "",                      # 55
        "S",                     # 56
        unidade,                 # 57
        "",                      # 58
        "",                      # 59
        fmt_decimal(v_prod),     # 60
        "",                      # 61
        "",                      # 62
        "",                      # 63
        "",                      # 64
        "",                      # 65
        "",                      # 66
        "",                      # 67
        "",                      # 68
        "",                      # 69
        "",                      # 70
        "",                      # 71
        "",                      # 72
        "",                      # 73
        "",                      # 74
        "",                      # 75
        "",                      # 76
        "",                      # 77
        "",                      # 78
        "",                      # 79
        "",                      # 80
        "",                      # 81
        "",                      # 82
        "",                      # 83
        "",                      # 84
        "",                      # 85
        "",                      # 86
        "",                      # 87
        "",                      # 88
        "",                      # 89
        "",                      # 90
        cest,                    # 91
        "",                      # 92
        "",                      # 93
        "",                      # 94
        "",                      # 95
        "",                      # 96
        v_icms_des,              # 97
        mot_des,                 # 98 - motDesICMS ← V1.7
        "",                      # 99
        "",                      # 100
        "",                      # 101
        "",                      # 102
        "",                      # 103
        ibs_class_trib,          # 104
        ibs_bc,                  # 105
        ibs_aliq,                # 106
        ibs_val,                 # 107
        ibs_class_trib,          # 108
        cbs_bc,                  # 109
        cbs_aliq,                # 110
        cbs_val,                 # 111
    ])

# ─────────────────────────────────────────────
# REGISTRO 1097 – Dados do Frete SP
# V1.7: campo 19 (Cidade) → codigo IBGE (Numerico)
#        transportadora FIORDE não tem <cMun> → envia vazio
# ─────────────────────────────────────────────
def gerar_registro_1097(nfe_root) -> str:
    transp = nfe_root.find("nfe:infNFe/nfe:transp", NS)
    if transp is None:
        return ""

    mod_frete_cod = get_text(transp, "nfe:modFrete")
    frete_map     = {"0":"C","1":"F","2":"S","3":"T","4":"R","5":"D","9":"S"}
    mod_frete     = frete_map.get(mod_frete_cod, "C")

    det_list = nfe_root.findall("nfe:infNFe/nfe:det", NS)
    tp_via   = ""
    if det_list:
        di_node = det_list[0].find("nfe:prod/nfe:DI", NS)
        if di_node is not None:
            tp_via = get_text(di_node, "nfe:tpViaTransp")

    transporta   = transp.find("nfe:transporta", NS)
    cnpj_transp  = get_text(transporta, "nfe:CNPJ")   if transporta is not None else ""
    razao_transp = get_text(transporta, "nfe:xNome")  if transporta is not None else ""
    ie_transp    = get_text(transporta, "nfe:IE")     if transporta is not None else ""
    end_transp   = get_text(transporta, "nfe:xEnder") if transporta is not None else ""
    uf_transp    = get_text(transporta, "nfe:UF")     if transporta is not None else ""

    # Campo 19 = Cidade → Numerico (cod IBGE)
    # XML da FIORDE tem <xMun>Guarulhos</xMun> mas NAO tem <cMun>
    # → envia vazio para nao gerar erro de tipo
    cmun_transp = get_text(transporta, "nfe:cMun") if transporta is not None else ""
    cidade_cod  = somente_numeros(cmun_transp) if cmun_transp else ""

    tipo_insc = "1" if cnpj_transp else ""

    vol     = transp.find("nfe:vol", NS)
    q_vol   = get_text(vol, "nfe:qVol")  if vol is not None else ""
    esp_vol = get_text(vol, "nfe:esp")   if vol is not None else ""
    marca   = get_text(vol, "nfe:marca") if vol is not None else ""
    peso_l  = fmt_decimal(get_text(vol, "nfe:pesoL"), 3) if vol is not None else ""
    peso_b  = fmt_decimal(get_text(vol, "nfe:pesoB"), 3) if vol is not None else ""

    return pipe_join([
        "1097",             # 1
        mod_frete,          # 2  - C=CIF ou F=FOB
        tp_via,             # 3  - Modalidade transporte (Numerico)
        "",                 # 4  - Frete por conta
        "",                 # 5  - Placa 1
        "",                 # 6  - UF Placa 1
        "",                 # 7  - Placa 2
        "",                 # 8  - UF Placa 2
        "",                 # 9  - Placa 3
        "",                 # 10 - UF Placa 3
        razao_transp[:150], # 11 - Razao Social
        tipo_insc,          # 12 - Tipo inscricao (Numerico)
        cnpj_transp,        # 13 - CNPJ
        ie_transp,          # 14 - IE
        end_transp,         # 15 - Endereco
        "",                 # 16 - Numero (Numerico)
        "",                 # 17 - Bairro
        "",                 # 18 - Complemento
        cidade_cod,         # 19 - Cidade (Numerico cod IBGE) ← V1.7 vazio se sem cMun
        uf_transp,          # 20 - UF
        "",                 # 21 - CEP
        q_vol,              # 22 - Qtd volumes (Numerico)
        esp_vol,            # 23 - Especie
        marca,              # 24 - Marca
        "",                 # 25 - Numeracao
        peso_l,             # 26 - Peso Liquido (3 casas)
        peso_b,             # 27 - Peso Bruto (3 casas)
        "E",                # 28 - Tipo local saida
        "",                 # 29 - CNPJ local saida
        "",                 # 30 - UF local saida
        "",                 # 31 - IE local saida
        "D",                # 32 - Tipo local recebimento
        "",                 # 33 - CNPJ local recebimento
        "",                 # 34 - UF local recebimento
        "",                 # 35 - IE local recebimento
    ])

# ─────────────────────────────────────────────
# REGISTROS 1150 / 1151
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

    # Detecta se é importação
    importacao = is_nota_importacao(nfe)

    # Nó do destinatário (fornecedor estrangeiro na importação)
    dest_node = nfe.find("nfe:infNFe/nfe:dest", NS)

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
    resumo["Importacao"]   = "Sim" if importacao  else "Nao"
    resumo["Emissao"]      = fmt_date(get_text(ide, "nfe:dhEmi"))
    resumo["Itens"]        = len(det_list)
    resumo["vNF"]          = fmt_decimal(get_text(total, "nfe:vNF"))
    resumo["vICMS"]        = fmt_decimal(get_text(total, "nfe:vICMS"))
    resumo["vICMSDes"]     = fmt_decimal(get_text(total, "nfe:vICMSDeson"))
    resumo["vIPI"]         = fmt_decimal(get_text(total, "nfe:vIPI"))
    resumo["vPIS"]         = fmt_decimal(get_text(total, "nfe:vPIS"))
    resumo["vCOFINS"]      = fmt_decimal(get_text(total, "nfe:vCOFINS"))
    resumo["Grupo"]        = (
        f"{grupo_padrao} - {TABELA_GRUPOS.get(grupo_padrao,'GERAL')}"
        if grupo_padrao > 0 else "Auto (CFOP/NCM)"
    )

    # ── Registros ─────────────────────────────────────────────────
    if incluir_0000:
        lines.append(gerar_registro_0000(cnpj_empresa))

    if incluir_0020 and emit is not None:
        # V1.7: importação → dados do <dest> (fornecedor estrangeiro)
        lines.append(gerar_registro_0020(emit, dest=dest_node, is_importacao=importacao))

    if incluir_0100:
        produtos_gerados = set()
        for det in det_list:
            cod = get_text(det.find("nfe:prod", NS), "nfe:cProd")
            if cod not in produtos_gerados:
                lines.append(gerar_registro_0100(det, grupo_padrao=grupo_padrao))
                produtos_gerados.add(cod)

    lines.append(gerar_registro_1000(nfe, cnpj_empresa, acumulador, especie))

    if incluir_1010:
        for r in gerar_registros_1010(nfe):
            lines.append(r)
    if incluir_1015:
        for r in gerar_registros_1015(nfe):
            lines.append(r)

    for r in gerar_registros_1020(nfe):
        lines.append(r)

    for seq, det in enumerate(det_list, start=1):
        lines.append(gerar_registro_1030(det, seq))

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
    cnpj_fallback = st.text_input("CNPJ Fallback (opcional)", value="", max_chars=14,
        help="Usado somente se o XML nao contiver CNPJ do destinatario.")
    acumulador = st.text_input("Codigo do Acumulador", value="1157")
    especie    = st.text_input("Codigo da Especie", value="36")
    st.markdown("---")
    st.markdown("### Registros de cadastro")
    inc_0000 = st.checkbox("0000 - Identificacao da empresa", value=True)
    inc_0020 = st.checkbox("0020 - Fornecedor", value=True)
    inc_0100 = st.checkbox("0100 - Produtos", value=True)
    st.markdown("---")
    st.markdown("### Registros filhos do 1000")
    inc_1010 = st.checkbox("1010 - Inf. Complementares", value=True)
    inc_1015 = st.checkbox("1015 - Observacoes", value=False)
    inc_1097 = st.checkbox("1097 - Dados do Frete SP", value=True)
    st.caption("1020 / 1030 / 1150 / 1151 sempre gerados.")
    st.markdown("---")
    st.markdown("### Grupo de Produtos (0100 campo 9)")
    grupo_selecionado = st.selectbox(
        "Grupo",
        options=list(TABELA_GRUPOS.keys()),
        format_func=lambda x: f"{x} - {TABELA_GRUPOS[x]}" if x > 0 else TABELA_GRUPOS[x],
        index=0,
    )
    st.caption("Auto: CFOP 3102 → Grupo 2" if grupo_selecionado == 0
               else f"Fixo: {grupo_selecionado} - {TABELA_GRUPOS[grupo_selecionado]}")
    st.markdown("---")
    with st.expander("Tabela de Grupos"):
        for cod, desc in sorted(TABELA_GRUPOS.items()):
            if cod > 0:
                st.caption(f"`{cod:3d}` - {desc}")

# ─────────────────────────────────────────────
# INSTRUÇÕES
# ─────────────────────────────────────────────
with st.expander("Instrucoes / Historico de versoes", expanded=False):
    st.markdown("""
        <div class="instrucoes-box">
        <h4>Correcoes V1.7</h4>
        <ul>
          <li><b>0020 campo 3 (Razao Social)</b>: NF importacao → <code>HILLROM COMPANY, INC USA</code> (dados do &lt;dest&gt;)</li>
          <li><b>0020 campo 5 (Endereco)</b>: <code>STATE ROUTE 46 EAST</code></li>
          <li><b>0020 campo 8 (Bairro)</b>: <code>MAIL CODE J36</code></li>
          <li><b>0020 campo 9 (Cod. Municipio)</b>: <code>9999999</code> (exterior)</li>
          <li><b>0020 campo 10 (UF)</b>: <code>EX</code></li>
          <li><b>0020 campo 11 (Cod. Pais)</b>: <code>2496</code> (EUA)</li>
          <li><b>0100 campo 54 (RS-MVA)</b>: <code>""</code> — Decimal, nunca "N"</li>
          <li><b>0100 campo 24 (Periodicidade IPI)</b>: fixo <code>M</code></li>
          <li><b>0100 tamanho</b>: pad automatico garante 92 posicoes sem crash</li>
          <li><b>1030 campo 9 (Numero DI)</b>: <code>somente_numeros()</code> remove "26BR" → <code>265021449</code></li>
          <li><b>1030 campo 98 (motDesICMS)</b>: lido da tag XML → <code>9</code></li>
          <li><b>1097 campo 19 (Cidade)</b>: codigo IBGE numerico; vazio se sem &lt;cMun&gt; (FIORDE)</li>
        </ul>
        </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ─────────────────────────────────────────────
# UPLOAD
# ─────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Selecione um ou mais arquivos XML de NF-e",
    type=["xml"],
    accept_multiple_files=True,
)

# ─────────────────────────────────────────────
# PROCESSAMENTO
# ─────────────────────────────────────────────
if uploaded_files:
    all_lines   = []
    all_resumos = []
    erros       = []
    progress    = st.progress(0, text="Processando arquivos...")

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
        progress.progress((i + 1) / len(uploaded_files), text=f"Processando {f.name}...")

    progress.empty()

    if erros:
        st.error("Erros encontrados:")
        st.dataframe(erros, use_container_width=True)

    if all_resumos:
        st.success(f"{len(all_resumos)} arquivo(s) convertido(s) com sucesso!")

        cnpjs_unicos = list({r["CNPJ Empresa"]: r for r in all_resumos}.values())
        if cnpjs_unicos:
            st.markdown("#### Empresa Identificada")
            cols = st.columns(min(len(cnpjs_unicos), 4))
            for idx, r in enumerate(cnpjs_unicos[:4]):
                is_ext = r.get("Exterior", "Nao") == "Sim"
                cor    = "#1565C0" if is_ext else "#FF8000"
                with cols[idx]:
                    st.markdown(
                        f'<div class="cnpj-badge" style="color:{cor};border-color:{cor};">'
                        f'CNPJ: {r["CNPJ Empresa"]}</div>'
                        f'<div class="info-origem" style="border-left-color:{cor};">'
                        f'{r.get("Origem CNPJ","")}<br>'
                        f'{"Importacao/Exterior" if is_ext else r.get("Destinatario","")[:60]}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

        st.markdown("---")
        with st.expander("Resumo das notas processadas", expanded=True):
            st.dataframe(all_resumos, use_container_width=True)

        saida_final = "\n".join(all_lines)

        with st.expander("Preview do arquivo gerado (primeiras 80 linhas)"):
            st.code("\n".join(saida_final.split("\n")[:80]), language="text")

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
        c1.metric("Notas",   len(all_resumos))
        c2.metric("Itens",   sum(r.get("Itens", 0) for r in all_resumos))
        total_linhas = len([l for l in saida_final.split("\n") if l.startswith("|")])
        c3.metric("Linhas geradas", total_linhas)
        c4.metric("Erros",   len(erros))
        try:
            total_nf = sum(float(r.get("vNF","0").replace(",","."))
                           for r in all_resumos if r.get("vNF"))
            c5.metric("Total NF (R$)",
                      f"{total_nf:,.2f}".replace(",","X").replace(".",",").replace("X","."))
        except Exception:
            c5.metric("Total NF", "-")

else:
    st.info("Faca o upload de um ou mais arquivos XML de NF-e para iniciar.")

st.markdown("---")
st.caption(f"Conversor XML NF-e → Dominio Sistemas | Thomson Reuters | Python + Streamlit | {VERSAO}")
