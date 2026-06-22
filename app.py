import streamlit as st
import xml.etree.ElementTree as ET
from datetime import datetime
import re

VERSAO = "V3.1-FINAL"

def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] { font-family: 'Segoe UI', 'Arial', sans-serif; color: #444444; }
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

CST_ENTRADA_SAIDA = {
    "50": "01", "51": "02", "52": "08",
    "73": "06", "74": "08", "70": "04", "99": "49",
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
        "84": 10, "85": 10, "86": 10, "87": 10, "88": 10, "89": 10,
        "90": 10, "91": 10, "94": 10,
        "28": 3, "29": 3, "30": 3, "31": 3, "32": 3, "33": 3, "34": 3,
        "38": 3, "39": 3, "40": 3, "44": 3, "47": 3, "48": 3,
        "72": 3, "73": 3, "74": 3, "75": 3, "76": 3, "82": 3, "83": 3,
        "01": 2, "02": 2, "03": 2, "04": 2, "07": 2, "08": 2, "09": 2,
        "10": 2, "16": 2, "17": 2, "18": 2, "19": 2, "20": 2, "21": 2,
        "22": 2, "27": 2,
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

def is_nota_importacao(nfe_root) -> bool:
    """Detecta importação por UF=EX no dest ou CFOP iniciando em 3."""
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

def extrair_cnpj_empresa(nfe_root, cnpj_fallback: str) -> tuple:
    """
    CNPJ para o registro 0000 (empresa que lança a nota):
    - Nacional:   CNPJ do <dest>
    - Importação: fallback informado pelo usuário
    Retorna: (cnpj, origem)
    """
    importacao = is_nota_importacao(nfe_root)

    if not importacao:
        dest = nfe_root.find("nfe:infNFe/nfe:dest", NS)
        if dest is not None:
            cnpj = get_text(dest, "nfe:CNPJ")
            if cnpj:
                return cnpj, "XML — <dest><CNPJ>"
            cpf = get_text(dest, "nfe:CPF")
            if cpf:
                return cpf, "XML — <dest><CPF>"

    if cnpj_fallback:
        return somente_numeros(cnpj_fallback), "Manual (fallback)"

    # Último recurso: emitente
    emit = nfe_root.find("nfe:infNFe/nfe:emit", NS)
    cnpj_emit = get_text(emit, "nfe:CNPJ") if emit is not None else ""
    if cnpj_emit:
        return cnpj_emit, "XML — <emit><CNPJ> (fallback)"

    return "", "Nao encontrado"

# ─────────────────────────────────────────────
# REGISTRO 0000
# ─────────────────────────────────────────────
def gerar_registro_0000(cnpj_empresa: str) -> str:
    return pipe_join(["0000", cnpj_empresa])

# ─────────────────────────────────────────────
# REGISTRO 0020 – Fornecedor
# Layout: 33 campos
# Importação → <dest> exterior: inscrição vazia, UF=EX, cPais BACEN
# Nacional   → <emit>: CNPJ, dados normais
# ─────────────────────────────────────────────
def gerar_registro_0020(emit, dest=None, is_importacao: bool = False) -> str:
    if is_importacao and dest is not None:
        razao       = get_text(dest, "nfe:xNome")[:150]
        fantasia    = razao[:40]
        ender       = dest.find("nfe:enderDest", NS)
        logradouro  = get_text(ender, "nfe:xLgr")                    if ender is not None else ""
        numero      = somente_numeros(get_text(ender, "nfe:nro"))     if ender is not None else ""
        complemento = ""
        bairro      = get_text(ender, "nfe:xBairro")                 if ender is not None else ""
        cod_mun     = somente_numeros(get_text(ender, "nfe:cMun"))    if ender is not None else ""
        cep         = get_text(ender, "nfe:CEP")                     if ender is not None else ""
        cod_pais    = somente_numeros(get_text(ender, "nfe:cPais"))   if ender is not None else ""
        inscricao   = ""
        uf_campo    = "EX"
        ie          = ""
        regime      = "N"
        contrib     = "N"
    else:
        inscricao    = get_text(emit, "nfe:CNPJ")
        razao        = get_text(emit, "nfe:xNome")[:150]
        fantasia_raw = get_text(emit, "nfe:xFant")
        fantasia     = fantasia_raw[:40] if fantasia_raw else razao[:40]
        ender        = emit.find("nfe:enderEmit", NS)
        logradouro   = get_text(ender, "nfe:xLgr")                   if ender is not None else ""
        numero       = somente_numeros(get_text(ender, "nfe:nro"))    if ender is not None else ""
        complemento  = get_text(ender, "nfe:xCpl")                   if ender is not None else ""
        bairro       = get_text(ender, "nfe:xBairro")                if ender is not None else ""
        cod_mun      = somente_numeros(get_text(ender, "nfe:cMun"))   if ender is not None else ""
        cep          = get_text(ender, "nfe:CEP")                    if ender is not None else ""
        uf_campo     = get_text(ender, "nfe:UF")                     if ender is not None else ""
        cod_pais     = ""
        ie           = get_text(emit, "nfe:IE")
        crt          = get_text(emit, "nfe:CRT")
        regime_map   = {"1": "M", "2": "E", "3": "N"}
        regime       = regime_map.get(crt, "N")
        contrib      = "S" if ie and ie.upper() not in ("ISENTO", "NAO CONTRIBUINTE", "") else "N"

    return pipe_join([
        "0020",       # 1  - Fixo
        inscricao,    # 2  - Inscrição (vazia para exterior)
        razao,        # 3  - Razão Social
        fantasia,     # 4  - Apelido
        logradouro,   # 5  - Endereço
        numero,       # 6  - Número
        complemento,  # 7  - Complemento
        bairro,       # 8  - Bairro
        cod_mun,      # 9  - Código município
        uf_campo,     # 10 - UF (EX para exterior)
        cod_pais,     # 11 - Código País BACEN (apenas exterior)
        cep,          # 12 - CEP
        ie,           # 13 - IE
        "",           # 14 - IM
        "",           # 15 - Suframa
        "",           # 16 - DDD
        "",           # 17 - Telefone
        "",           # 18 - FAX
        "",           # 19 - Data cadastro
        "",           # 20 - Conta contábil
        "",           # 21 - Conta contábil cliente
        "N",          # 22 - Agropecuário
        "7",          # 23 - Natureza jurídica
        regime,       # 24 - Regime apuração
        contrib,      # 25 - Contribuinte ICMS
        "",           # 26 - Alíquota ICMS
        "",           # 27 - Categoria estabelecimento
        "",           # 28 - IE ST
        "",           # 29 - Email
        "N",          # 30 - Interdependência
        "N",          # 31 - Contribuinte CPRB
        "",           # 32 - Processo adm/judicial
        "",           # 33 - Tipo inscrição
    ])

# ─────────────────────────────────────────────
# EXTRAIR PIS/COFINS E cClassTrib
# ─────────────────────────────────────────────
def extrair_pis_cofins(det) -> dict:
    imposto = det.find("nfe:imposto", NS)
    resultado = {
        "cst_e": "", "aliq_pis_e": "", "aliq_cof_e": "",
        "cst_s": "", "aliq_pis_s": "", "aliq_cof_s": "",
        "class_trib": "",
    }
    if imposto is None:
        return resultado

    pis_node = imposto.find("nfe:PIS", NS)
    if pis_node is not None:
        for pt in ["PISAliq", "PISQtde", "PISNT", "PISOutr"]:
            pn = pis_node.find(f"nfe:{pt}", NS)
            if pn is not None:
                resultado["cst_e"] = get_text(pn, "nfe:CST")
                aliq = get_text(pn, "nfe:pPIS") or get_text(pn, "nfe:vAliqProd")
                resultado["aliq_pis_e"] = fmt_decimal(aliq, 4)
                break

    cof_node = imposto.find("nfe:COFINS", NS)
    if cof_node is not None:
        for ct in ["COFINSAliq", "COFINSQtde", "COFINSNT", "COFINSOutr"]:
            cn = cof_node.find(f"nfe:{ct}", NS)
            if cn is not None:
                aliq = get_text(cn, "nfe:pCOFINS") or get_text(cn, "nfe:vAliqProd")
                resultado["aliq_cof_e"] = fmt_decimal(aliq, 4)
                break

    resultado["cst_s"]      = CST_ENTRADA_SAIDA.get(resultado["cst_e"], "")
    resultado["aliq_pis_s"] = resultado["aliq_pis_e"]
    resultado["aliq_cof_s"] = resultado["aliq_cof_e"]

    ibs_node = imposto.find("nfe:IBSCBS", NS)
    if ibs_node is not None:
        resultado["class_trib"] = get_text(ibs_node, "nfe:cClassTrib")

    return resultado

# ─────────────────────────────────────────────
# REGISTRO 0100 – 92 campos
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
        "0100", cod_prod, descricao, "", ncm, "", "", "", cod_grupo,
        unidade, "N", "O", "", "", "", "N", "",
        fmt_decimal(val_unit, 3), "", "", cst_icms, aliq_icms, aliq_ipi, "M",
        "", "N", "", "", "", "", "", "", "", "", "N", "", "", "", "N", "",
        "", "", "N", "", "", "", "N", "N", "", "", "", "", "", "", "N",
        "", "", "", "", "N", "", "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", "", cest, "", "",
    ]
    while len(campos) < 92:
        campos.append("")
    campos = campos[:92]
    return pipe_join(campos)

# ─────────────────────────────────────────────
# REGISTRO 0110 – 70 campos
# ─────────────────────────────────────────────
def gerar_registro_0110(det) -> str:
    pc = extrair_pis_cofins(det)
    ct = pc["class_trib"]

    return pipe_join([
        "0110",           # 1  - Fixo
        "Inicial",        # 2  - Descrição vigência
        pc["cst_e"],      # 3  - CST Entrada
        "",               # 4  - Vínculo do Crédito
        "01",             # 5  - Base do Crédito
        "N",              # 6  - Crédito proporcional
        "N",              # 7  - Crédito alíq. diferenciada
        pc["aliq_pis_e"], # 8  - Alíquota PIS Entrada
        pc["aliq_cof_e"], # 9  - Alíquota COFINS Entrada
        "N",              # 10 - Crédito por unidade medida
        "N",              # 11 - Unidade tributada diferente
        "",               # 12 - Unidade tributável Entrada
        "",               # 13 - Fator conversão Entrada
        "",               # 14 - Valor PIS Entrada
        "",               # 15 - Valor COFINS Entrada
        pc["cst_s"],      # 16 - CST Saída
        "N",              # 17 - Tipo contribuição
        "",               # 18 - Natureza receita
        "",               # 19 - Cód. recolhimento PIS Saída
        "",               # 20 - Cód. recolhimento COFINS Saída
        "N",              # 21 - Débito alíq. diferenciada
        pc["aliq_pis_s"], # 22 - Alíquota PIS Saída
        pc["aliq_cof_s"], # 23 - Alíquota COFINS Saída
        "N",              # 24 - Débito por unidade medida
        "N",              # 25 - Unidade tributada diferente Saída
        "",               # 26 - Unidade tributável Saída
        "",               # 27 - Fator conversão Saída
        "",               # 28 - Valor PIS Saída
        "",               # 29 - Valor COFINS Saída
        "",               # 30 - Tabela SPED
        "",               # 31 - Marca/Grupo SPED
        "N",              # 32 - PIS cumulativo
        "N",              # 33 - COFINS cumulativo
        "",               # 34 - ICMS CST Entradas
        "",               # 35 - ICMS CST Saídas
        "",               # 36 - ICMS Alíquota
        "",               # 37 - IPI CST Entradas
        "",               # 38 - IPI CST Saídas
        "M",              # 39 - IPI Periodicidade
        "",               # 40 - IPI Alíquota
        "N",              # 41 - SN PIS/COFINS
        "N",              # 42 - Excluir base importação
        "N",              # 43 - FUNDEPEC GO
        "",               # 44 - Tipo produto FUNDEPEC
        "N",              # 45 - PRODEPE PE
        "",               # 46 - Cód. apuração PRODEPE
        "N",              # 47 - Redução base
        "",               # 48 - % redução base PIS/COFINS
        "",               # 49 - SN tipo tributação
        "",               # 50 - Cód. recolhimento PIS Entrada
        "",               # 51 - Cód. recolhimento COFINS Entrada
        "",               # 52 - Base cálculo ST
        "",               # 53 - % margem ST
        "",               # 54 - Valor unitário ST
        "",               # 55 - IPI cód. recolhimento
        "N",              # 56 - RS Detalhamento VA/VB
        "",               # 57 - RS Cód. VA
        "",               # 58 - RS Cód. VB
        "N",              # 59 - Bebidas frias SN
        "",               # 60 - Alíquota PIS alt.
        "",               # 61 - Alíquota COFINS alt.
        "N",              # 62 - RS ressarcimento ICMS ST
        "",               # 63 - RS % base cálculo
        "N",              # 64 - RS PMPF combustíveis
        "N",              # 65 - ES interestaduais
        "N",              # 66 - ES internas
        ct,               # 67 - IBS cClassTrib
        ct,               # 68 - CBS cClassTrib
        "N",              # 69 - IBS tabela NCM/NBS
        "N",              # 70 - CBS tabela NCM/NBS
    ])

# ─────────────────────────────────────────────
# REGISTRO 1000 – NF de Entrada (98 campos)
# V3.1: Campo 70 preenchido para notas de importação
# ─────────────────────────────────────────────
def gerar_registro_1000(nfe_root, cnpj_empresa: str,
                        acumulador: str = "1157",
                        especie: str = "36",
                        importacao: bool = False) -> str:
    ide   = nfe_root.find("nfe:infNFe/nfe:ide", NS)
    emit  = nfe_root.find("nfe:infNFe/nfe:emit", NS)
    dest  = nfe_root.find("nfe:infNFe/nfe:dest", NS)
    total = nfe_root.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)

    # ── Campo 3: Inscrição do fornecedor ─────────────────────────────
    if importacao and dest is not None:
        id_ext    = get_text(dest, "nfe:idEstrangeiro").strip()
        cnpj_forn = id_ext if id_ext else ""
    else:
        cnpj_forn = get_text(emit, "nfe:CNPJ")

    # ── Campo 17: P=Próprio (importação) / T=Terceiros (nacional) ────
    emitente_nf = "P" if importacao else "T"

    # ── Campo 44: IE do fornecedor ────────────────────────────────────
    ie_forn = "" if importacao else get_text(emit, "nfe:IE")

    nNF      = get_text(ide, "nfe:nNF")
    serie    = get_text(ide, "nfe:serie")
    dhEmi    = fmt_date(get_text(ide, "nfe:dhEmi"))
    c_mun_fg = get_text(ide, "nfe:cMunFG")

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

    # ── Número da DI (campo 52) ───────────────────────────────────────
    n_di = ""
    if det_list:
        di_node = det_list[0].find("nfe:prod/nfe:DI", NS)
        if di_node is not None:
            n_di = get_text(di_node, "nfe:nDI")

    # ── Campo 70: Tipo do documento de importação ─────────────────────
    # 10 = Declaração de Importação (DI)
    #  1 = Declaração Simplificada de Importação (DSI)
    tipo_doc_importacao = ""
    if importacao and n_di:
        if n_di.upper().startswith("DSI"):
            tipo_doc_importacao = "1"
        else:
            tipo_doc_importacao = "10"

    return pipe_join([
        "1000",               # 1  - Fixo
        especie,              # 2  - Código espécie
        cnpj_forn,            # 3  - Inscrição fornecedor
        "",                   # 4  - Código exclusão DIEF
        acumulador,           # 5  - Código acumulador
        cfop_first,           # 6  - CFOP
        "",                   # 7  - Segmento
        nNF,                  # 8  - Número documento
        serie,                # 9  - Série
        "",                   # 10 - Número documento final
        dhEmi,                # 11 - Data entrada
        dhEmi,                # 12 - Data emissão
        v_nf,                 # 13 - Valor contábil
        "",                   # 14 - Valor exclusão DIEF
        obs_fisco,            # 15 - Observação fiscal
        mod_frete,            # 16 - Modalidade frete
        emitente_nf,          # 17 - Emitente NF (P/T)
        "",                   # 18 - CFOP estendido SE
        "",                   # 19 - Cód. transferência crédito RS
        "",                   # 20 - Cód. recolhimento ISS retido
        "",                   # 21 - Cód. recolhimento IRRF
        "",                   # 22 - Código observação
        "",                   # 23 - Data visto MG
        "",                   # 24 - Fato gerador CRF
        "",                   # 25 - Fato gerador IRRF
        v_frete,              # 26 - Valor frete
        v_seg,                # 27 - Valor seguro
        v_outro,              # 28 - Valor despesas acessórias
        v_pis,                # 29 - Valor PIS
        "",                   # 30 - Tipo antecipação tributária
        v_cofins,             # 31 - Valor COFINS
        "",                   # 32 - Valor DARE SE
        "",                   # 33 - Alíquota DARE SE
        "",                   # 34 - Valor BC ICMS ST
        "",                   # 35 - Entradas isentas MG
        "",                   # 36 - Outras entradas isentas MG
        "",                   # 37 - Valor transporte incluído base MG
        "",                   # 38 - Código ressarcimento
        v_prod,               # 39 - Valor produtos
        c_mun_fg,             # 40 - Município origem
        "0",                  # 41 - Situação da nota (0=Regular)
        "",                   # 42 - Código situação tributária
        "",                   # 43 - Sub série
        ie_forn,              # 44 - IE fornecedor
        "",                   # 45 - IM fornecedor
        "",                   # 46 - Código operação e prestação
        "",                   # 47 - Valor deduzido receita tributável
        "",                   # 48 - Competência
        "",                   # 49 - Operação PA
        "",                   # 50 - Número parecer fiscal
        "",                   # 51 - Data parecer fiscal
        n_di,                 # 52 - Número declaração de importação
        "N",                  # 53 - Possui benefício fiscal
        chave,                # 54 - Chave NF-e
        "",                   # 55 - Cód. recolhimento FETHAB
        "",                   # 56 - Responsável recolhimento FETHAB
        "",                   # 57 - CFOP documento fiscal
        "",                   # 58 - Tipo CT-e
        "",                   # 59 - CT-e referência
        "1",                  # 60 - Modalidade importação (1=Com direito a crédito)
        "",                   # 61 - Cód. informação complementar
        "",                   # 62 - Informação complementar
        "",                   # 63 - Classe de consumo
        "",                   # 64 - Tipo de ligação
        "",                   # 65 - Grupo de tensão
        "",                   # 66 - Tipo de assinante
        "",                   # 67 - KWH consumido
        "",                   # 68 - Valor fornecido energia/gás
        "",                   # 69 - Valor cobrado de terceiros
        tipo_doc_importacao,  # 70 - Tipo doc. importação (10=DI / 1=DSI)
        "",                   # 71 - Número Ato Concessório Drawback
        "",                   # 72 - Natureza frete PIS/COFINS
        "",                   # 73 - CST PIS/COFINS
        "",                   # 74 - Base crédito PIS/COFINS
        "",                   # 75 - Valor serviços PIS/COFINS
        "",                   # 76 - Base cálculo PIS/COFINS
        "",                   # 77 - Alíquota PIS
        "",                   # 78 - Alíquota COFINS
        "",                   # 79 - Chave NFSe
        "",                   # 80 - Número processo/ato concessório
        "",                   # 81 - Origem processo
        "",                   # 82 - Data escrituração
        "",                   # 83 - CFPS DF
        "",                   # 84 - Natureza receita PIS/COFINS
        "",                   # 85 - CST IPI
        "",                   # 86 - Lançamentos SCP
        "",                   # 87 - Tipo serviço
        "",                   # 88 - Município destino
        "",                   # 89 - Pedágio
        v_ipi,                # 90 - Valor IPI
        v_st,                 # 91 - Valor ICMS ST
        "",                   # 92 - Classificação serviços EFD-Reinf tipo
        "",                   # 93 - Indicativo prestação serviço EFD-Reinf
        "",                   # 94 - Número documento arrecadação RS
        "",                   # 95 - Tipo do título
        "",                   # 96 - Identificação
        v_icms_d,             # 97 - ICMS desonerado
        "",                   # 98 - IPI devolução
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
# REGISTROS 1020 – Impostos (19 campos)
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
# REGISTRO 1030 – Estoque/Item (111 campos)
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
    n_di    = somente_numeros(get_text(di_node, "nfe:nDI")) if di_node is not None else ""
    d_di    = fmt_date(get_text(di_node, "nfe:dDI"))        if di_node is not None else ""

    icms_node  = None
    v_bc_icms  = aliq_icms = v_icms = cst_icms = v_icms_des = v_bc_st = ""
    mot_des    = ""
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
            mot_des    = get_text(icms_node, "nfe:motDesICMS")

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
        "1030",                   # 1  - Fixo
        cod_prod,                 # 2  - Código produto
        qtd,                      # 3  - Quantidade
        v_total,                  # 4  - Valor total (vProd + vIPI)
        v_ipi,                    # 5  - Valor IPI
        fmt_decimal(v_prod),      # 6  - Base de cálculo (vProd)
        "1",                      # 7  - Tipo lançamento
        d_di,                     # 8  - Data DI
        n_di,                     # 9  - Número DI (somente dígitos)
        cst_icms,                 # 10 - CST ICMS
        fmt_decimal(v_prod),      # 11 - Valor bruto produto
        fmt_decimal(v_desc),      # 12 - Valor desconto
        v_bc_icms,                # 13 - BC ICMS
        v_bc_st,                  # 14 - BC ICMS ST
        aliq_icms,                # 15 - Alíquota ICMS
        "",                       # 16 - Produto incentivado PE
        "",                       # 17 - Cód. apuração PE
        "",                       # 18 - Valor frete
        "",                       # 19 - Valor seguro
        fmt_decimal(v_outro),     # 20 - Despesas acessórias
        "",                       # 21 - Qtd gasolina
        v_icms,                   # 22 - Valor ICMS
        "",                       # 23 - Valor SUBTRI
        "",                       # 24 - Isentas IPI
        "",                       # 25 - Outras IPI
        "",                       # 26 - ICMS NFP
        fmt_decimal(v_unit, 6),   # 27 - Valor unitário
        "",                       # 28 - Alíquota ST
        cst_ipi,                  # 29 - CST IPI
        aliq_ipi,                 # 30 - Alíquota IPI
        "",                       # 31 - BC ISSQN
        "",                       # 32 - Alíquota ISSQN
        "",                       # 33 - Valor ISSQN
        cfop,                     # 34 - CFOP
        "",                       # 35 - Série ECF
        aliq_pis,                 # 36 - Alíquota PIS
        v_pis,                    # 37 - Valor PIS
        aliq_cof,                 # 38 - Alíquota COFINS
        v_cof,                    # 39 - Valor COFINS
        fmt_decimal(v_prod),      # 40 - Custo total produto
        cst_pis,                  # 41 - CST PIS
        bc_pis,                   # 42 - BC PIS
        cst_cof,                  # 43 - CST COFINS
        bc_cof,                   # 44 - BC COFINS
        "",                       # 45 - Chassi veículo
        "",                       # 46 - Tipo operação veículo
        "",                       # 47 - Lote medicamento
        "",                       # 48 - Qtd lote medicamento
        "",                       # 49 - Data validade
        "",                       # 50 - Data fabricação
        "",                       # 51 - Referência BC
        "",                       # 52 - Valor tabelado
        "",                       # 53 - Nº série arma
        "",                       # 54 - Nº série cano
        "",                       # 55 - Enquadramento IPI
        "S",                      # 56 - Movimentação física
        unidade,                  # 57 - Unidade comercializada
        "",                       # 58 - Complemento CFOP SP
        "",                       # 59 - Tanque combustível
        fmt_decimal(v_prod),      # 60 - Valor contábil produto
        "",                       # 61 - Qtd trib. PIS/unid.
        "",                       # 62 - Valor unit. PIS/unid.
        "",                       # 63 - Valor PIS/unid.
        "",                       # 64 - Qtd trib. COFINS/unid.
        "",                       # 65 - Valor unit. COFINS/unid.
        "",                       # 66 - Valor COFINS/unid.
        "",                       # 67 - Base do crédito
        "",                       # 68 - Nº nota devolvida
        "",                       # 69 - Descrição complementar
        "",                       # 70 - CST PIS nota devolvida
        "",                       # 71 - CST COFINS nota devolvida
        "",                       # 72 - Vínculo crédito PIS
        "",                       # 73 - Vínculo crédito COFINS
        "",                       # 74 - Exclusão PIS
        "",                       # 75 - Exclusão COFINS
        "",                       # 76 - BC ICMS carga média
        "",                       # 77 - Alíq. ICMS carga média
        "",                       # 78 - Valor ICMS carga média
        "",                       # 79 - Nº série ECF devolvido
        "",                       # 80 - % redução BC PIS/COFINS
        "",                       # 81 - Cód. recolhimento PIS dev.
        "",                       # 82 - Cód. recolhimento COFINS dev.
        "",                       # 83 - Cód. recolhimento PIS
        "",                       # 84 - Cód. recolhimento COFINS
        "",                       # 85 - Crédito presumido PIS
        "",                       # 86 - Crédito presumido COFINS
        "",                       # 87 - ICMS ST Antec. BC
        "",                       # 88 - ICMS ST Antec. Alíq.
        "",                       # 89 - ICMS ST Antec. Valor
        "",                       # 90 - Cód. recolhimento IPI
        cest,                     # 91 - Código CEST
        "",                       # 92 - ICMS ST Retido BC
        "",                       # 93 - ICMS ST Retido Valor
        "",                       # 94 - ICMS ST Retido tag XML
        "",                       # 95 - Identificador
        "",                       # 96 - ICMS próprio substituto
        v_icms_des,               # 97 - Valor desonerado
        "",                       # 98 - Código desoneração
        "",                       # 99 - ICMS não creditado
        "",                       # 100 - ICMS monofásico qtd.
        "",                       # 101 - ICMS monofásico alíq.
        "",                       # 102 - ICMS monofásico valor
        "",                       # 103 - ICMS monofásico FCV
        ibs_class_trib,           # 104 - IBS cClassTrib
        ibs_bc,                   # 105 - IBS BC
        ibs_aliq,                 # 106 - IBS Alíquota
        ibs_val,                  # 107 - IBS Valor
        ibs_class_trib,           # 108 - CBS cClassTrib
        cbs_bc,                   # 109 - CBS BC
        cbs_aliq,                 # 110 - CBS Alíquota
        cbs_val,                  # 111 - CBS Valor
    ])

# ─────────────────────────────────────────────
# REGISTRO 1097 – Frete SP (35 campos)
# ─────────────────────────────────────────────
def gerar_registro_1097(nfe_root) -> str:
    transp = nfe_root.find("nfe:infNFe/nfe:transp", NS)
    if transp is None:
        return ""

    mod_frete_cod = get_text(transp, "nfe:modFrete")
    frete_map     = {"0":"C","1":"F","2":"S","3":"T","4":"R","5":"D","9":"S"}
    mod_frete     = frete_map.get(mod_frete_cod, "C")
    frete_conta   = "E" if mod_frete == "C" else "D"

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
    cmun_transp  = get_text(transporta, "nfe:cMun")   if transporta is not None else ""
    cidade_cod   = somente_numeros(cmun_transp) if cmun_transp else ""
    tipo_insc    = "1" if cnpj_transp else ""

    vol     = transp.find("nfe:vol", NS)
    q_vol   = get_text(vol, "nfe:qVol")  if vol is not None else ""
    esp_vol = get_text(vol, "nfe:esp")   if vol is not None else ""
    marca   = get_text(vol, "nfe:marca") if vol is not None else ""
    peso_l  = fmt_decimal(get_text(vol, "nfe:pesoL"), 3) if vol is not None else ""
    peso_b  = fmt_decimal(get_text(vol, "nfe:pesoB"), 3) if vol is not None else ""

    return pipe_join([
        "1097",              # 1  - Fixo
        mod_frete,           # 2  - Modalidade frete
        tp_via,              # 3  - Modalidade transporte
        frete_conta,         # 4  - Frete por conta
        "",                  # 5  - Placa 1
        "",                  # 6  - UF Placa 1
        "",                  # 7  - Placa 2
        "",                  # 8  - UF Placa 2
        "",                  # 9  - Placa 3
        "",                  # 10 - UF Placa 3
        razao_transp[:150],  # 11 - Razão Social transportador
        tipo_insc,           # 12 - Tipo inscrição
        cnpj_transp,         # 13 - CNPJ transportador
        ie_transp,           # 14 - IE transportador
        end_transp,          # 15 - Endereço
        "",                  # 16 - Número
        "",                  # 17 - Bairro
        "",                  # 18 - Complemento
        cidade_cod,          # 19 - Cidade (código IBGE)
        uf_transp,           # 20 - UF
        "",                  # 21 - CEP
        q_vol,               # 22 - Qtd volumes
        esp_vol,             # 23 - Espécie
        marca,               # 24 - Marca
        "",                  # 25 - Numeração
        peso_l,              # 26 - Peso líquido
        peso_b,              # 27 - Peso bruto
        "E",                 # 28 - Tipo local saída
        "",                  # 29 - CNPJ local saída
        "",                  # 30 - UF local saída
        "",                  # 31 - IE local saída
        "D",                 # 32 - Tipo local recebimento
        "",                  # 33 - CNPJ local recebimento
        "",                  # 34 - UF local recebimento
        "",                  # 35 - IE local recebimento
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
    incluir_0110: bool = True,
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

    importacao = is_nota_importacao(nfe)
    cnpj_empresa, origem_cnpj = extrair_cnpj_empresa(nfe, cnpj_fallback)

    if not cnpj_empresa:
        return "", {"erro": "CNPJ nao encontrado. Para importacao, informe o CNPJ Fallback."}

    lines     = []
    resumo    = {}
    det_list  = nfe.findall("nfe:infNFe/nfe:det", NS)
    emit      = nfe.find("nfe:infNFe/nfe:emit", NS)
    dest_node = nfe.find("nfe:infNFe/nfe:dest", NS)
    ide       = nfe.find("nfe:infNFe/nfe:ide", NS)
    total     = nfe.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)

    if importacao and dest_node is not None:
        nome_forn = get_text(dest_node, "nfe:xNome")
        uf_forn   = "EX"
    else:
        nome_forn = get_text(emit, "nfe:xNome") if emit is not None else ""
        uf_forn   = ""
        if emit is not None:
            ender_e = emit.find("nfe:enderEmit", NS)
            uf_forn = get_text(ender_e, "nfe:UF") if ender_e is not None else ""

    resumo["nNF"]          = get_text(ide, "nfe:nNF")
    resumo["Emitente"]     = get_text(emit, "nfe:xNome") if emit is not None else ""
    resumo["CNPJ Emit"]    = get_text(emit, "nfe:CNPJ")  if emit is not None else ""
    resumo["Fornecedor"]   = nome_forn
    resumo["UF Forn"]      = uf_forn
    resumo["CNPJ Empresa"] = cnpj_empresa
    resumo["Origem CNPJ"]  = origem_cnpj
    resumo["Importacao"]   = "Sim" if importacao else "Nao"
    resumo["Emitente NF"]  = "P (Proprio)" if importacao else "T (Terceiros)"
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

    if incluir_0000:
        lines.append(gerar_registro_0000(cnpj_empresa))

    if incluir_0020 and emit is not None:
        lines.append(gerar_registro_0020(emit, dest=dest_node, is_importacao=importacao))

    if incluir_0100:
        produtos_gerados = set()
        for det in det_list:
            cod = get_text(det.find("nfe:prod", NS), "nfe:cProd")
            if cod not in produtos_gerados:
                lines.append(gerar_registro_0100(det, grupo_padrao=grupo_padrao))
                if incluir_0110:
                    lines.append(gerar_registro_0110(det))
                produtos_gerados.add(cod)

    lines.append(gerar_registro_1000(nfe, cnpj_empresa, acumulador, especie, importacao))

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
    cnpj_fallback = st.text_input(
        "CNPJ da Empresa (obrigatorio para importacao)",
        value="", max_chars=14,
        help="Para notas de importacao (CFOP 3xxx / dest. exterior), informe o CNPJ da empresa importadora."
    )
    acumulador = st.text_input("Codigo do Acumulador", value="1157")
    especie    = st.text_input("Codigo da Especie", value="36")
    st.markdown("---")
    st.markdown("### Registros de cadastro")
    inc_0000 = st.checkbox("0000 - Identificacao da empresa", value=True)
    inc_0020 = st.checkbox("0020 - Fornecedor", value=True)
    inc_0100 = st.checkbox("0100 - Produtos", value=True)
    inc_0110 = st.checkbox("0110 - Vigencia PIS/COFINS", value=True)
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
    with st.expander("Mapeamento CST Entrada → Saida"):
        for ce, cs in CST_ENTRADA_SAIDA.items():
            st.caption(f"CST {ce} → {cs}")

# ─────────────────────────────────────────────
# INSTRUÇÕES
# ─────────────────────────────────────────────
with st.expander("Instrucoes / Historico de versoes", expanded=False):
    st.markdown("""
        <div class="instrucoes-box">
        <h4>V3.1-FINAL — Campo 70 do Registro 1000 corrigido</h4>
        <ul>
          <li><b>1000 campo 70</b>: Preenchido automaticamente em notas de importacao</li>
          <li><b>10</b> = Declaracao de Importacao (DI) — numero nao inicia com DSI</li>
          <li><b>1</b> = Declaracao Simplificada de Importacao (DSI) — numero inicia com DSI</li>
          <li><b>Vazio</b> = Nota nacional (nao importacao)</li>
        </ul>
        <h4>V3.0-FINAL — Baseado nos layouts oficiais do Dominio</h4>
        <ul>
          <li><b>0000 campo 2</b>: CNPJ da empresa (fallback para importacao)</li>
          <li><b>0020</b>: Importacao → fornecedor EXTERIOR (<code>&lt;dest&gt;</code>), CNPJ vazio, cPais BACEN</li>
          <li><b>1000 campo 3</b>: CNPJ do <code>&lt;emit&gt;</code> — valida chave NF-e</li>
          <li><b>1000 campo 17</b>: <b>P=Proprio</b> (importacao) / <b>T=Terceiros</b> (nacional)</li>
          <li><b>1000</b>: 98 campos conforme layout oficial</li>
          <li><b>1030</b>: 111 campos conforme layout oficial</li>
          <li><b>1097</b>: 35 campos conforme layout oficial</li>
          <li><b>0110 campos 8/9</b>: Aliquota PIS/COFINS lida do XML</li>
          <li><b>1030 campo 9</b>: Numero DI = somente digitos</li>
          <li><b>1097 campo 19</b>: Cidade = codigo IBGE numerico</li>
        </ul>
        </div>
    """, unsafe_allow_html=True)

st.markdown("---")

# ─────────────────────────────────────────────
# UPLOAD E PROCESSAMENTO
# ─────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Selecione um ou mais arquivos XML de NF-e",
    type=["xml"],
    accept_multiple_files=True,
)

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
            incluir_0110   = inc_0110,
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
            st.markdown("#### Empresa / Fornecedor")
            cols = st.columns(min(len(cnpjs_unicos), 4))
            for idx, r in enumerate(cnpjs_unicos[:4]):
                is_imp = r.get("Importacao", "Nao") == "Sim"
                cor    = "#1565C0" if is_imp else "#FF8000"
                with cols[idx]:
                    st.markdown(
                        f'<div class="cnpj-badge" style="color:{cor};border-color:{cor};">'
                        f'Empresa: {r["CNPJ Empresa"]}</div>'
                        f'<div class="info-origem" style="border-left-color:{cor};">'
                        f'{r.get("Origem CNPJ","")}<br>'
                        f'{"Importacao | Forn: " + r.get("Fornecedor","")[:45] if is_imp else "Forn: " + r.get("Fornecedor","")[:50]}'
                        f'<br><small>Emitente NF: {r.get("Emitente NF","")}</small>'
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
