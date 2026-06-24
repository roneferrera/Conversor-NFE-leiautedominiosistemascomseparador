import streamlit as st
import xml.etree.ElementTree as ET
from datetime import datetime
import re

VERSAO = "V3.9-FINAL"

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

PAISES_BACEN_PARA_DOMINIO = {
    "0132": 1,   "7560": 2,   "0175": 3,   "0230": 4,   "0370": 5,
    "0400": 6,   "0418": 7,   "0434": 8,   "0477": 9,   "0531": 10,
    "0590": 11,  "0639": 12,  "0647": 13,  "0655": 14,  "0698": 15,
    "0728": 16,  "0736": 17,  "0779": 18,  "0809": 19,  "0817": 20,
    "0833": 21,  "0850": 22,  "0876": 23,  "0884": 24,  "0906": 25,
    "0930": 26,  "0973": 27,  "0981": 28,  "1015": 29,  "1058": 30,
    "1082": 31,  "1112": 32,  "1155": 33,  "1198": 34,  "1279": 35,
    "1376": 36,  "1414": 37,  "1457": 38,  "1490": 39,  "1504": 40,
    "1511": 41,  "1546": 42,  "1554": 43,  "1589": 44,  "1600": 45,
    "1635": 46,  "1651": 47,  "1694": 48,  "1716": 49,  "1750": 50,
    "1767": 51,  "1792": 52,  "1830": 53,  "1872": 54,  "1880": 55,
    "1902": 56,  "1937": 57,  "1953": 58,  "1961": 59,  "1988": 60,
    "2003": 61,  "2070": 62,  "2100": 63,  "2127": 64,  "2151": 65,
    "2186": 66,  "2291": 67,  "2321": 68,  "2356": 69,  "2399": 70,
    "2402": 71,  "2445": 72,  "2453": 73,  "2461": 74,
    "2484": 75,  "2496": 76,  "2518": 77,  "2534": 78,  "2550": 79,
    "2593": 80,  "2674": 81,  "2712": 82,  "2755": 83,  "2810": 84,
    "2836": 85,  "2895": 86,  "2917": 87,  "2933": 88,  "2976": 89,
    "3018": 90,  "3050": 91,  "3093": 92,  "3107": 93,  "3131": 94,
    "3174": 95,  "3204": 96,  "3212": 97,  "3255": 98,  "3298": 99,
    "3310": 100, "3344": 101, "3352": 102, "3360": 103, "3417": 104,
    "3450": 105, "3484": 106, "3492": 107, "3557": 108, "3573": 109,
    "3611": 110, "3654": 111, "3697": 112, "3727": 113, "3751": 114,
    "3794": 115, "3808": 116, "3832": 117, "3867": 118, "3883": 119,
    "3913": 120, "3964": 121, "3999": 122, "4030": 123, "4111": 124,
    "4200": 125, "4235": 126, "4260": 127, "4278": 128, "4316": 129,
    "4340": 130, "4383": 131, "4405": 132, "4421": 133, "4456": 134,
    "4472": 135, "4499": 136, "4502": 137, "4553": 138, "4588": 139,
    "4618": 140, "4642": 141, "4677": 142, "4723": 143, "4740": 144,
    "4766": 145, "4774": 146, "4855": 147, "4880": 148, "4936": 149,
    "4944": 150, "4952": 151, "5010": 152, "5053": 153, "5070": 154,
    "5088": 155, "5096": 156, "5177": 157, "5215": 158, "5258": 159,
    "5282": 160, "5380": 161, "5428": 162, "5487": 163, "5568": 164,
    "5665": 165, "5738": 166, "5754": 167, "5762": 168, "5800": 169,
    "5835": 170, "5851": 171, "5894": 172, "5932": 173, "5991": 174,
    "6033": 175, "6076": 176, "6114": 177, "6238": 178, "6254": 179,
    "6289": 180, "6300": 181, "6327": 182, "6408": 183, "6432": 184,
    "6459": 185, "6505": 186, "6513": 187, "6548": 188, "6580": 189,
    "6599": 190, "6645": 191, "6700": 192, "6750": 193, "6769": 194,
    "6777": 195, "6781": 196, "6793": 197, "6807": 198, "6815": 199,
    "6858": 200, "6866": 201, "6904": 202, "6912": 203, "7005": 204,
    "7030": 205, "7056": 206, "7102": 207, "7153": 208, "7200": 209,
    "7285": 210, "7315": 211, "7358": 212, "7370": 213, "7412": 214,
    "7447": 215, "7455": 216, "7501": 217, "7544": 218, "7552": 219,
    "7590": 220, "7595": 221, "7641": 222, "7676": 223, "7706": 224,
    "7722": 225, "7757": 226, "7765": 227, "7773": 228, "7781": 229,
    "7820": 230, "7838": 231, "7889": 232, "7919": 233, "7951": 234,
    "7994": 235, "8001": 236, "8052": 237, "8079": 238, "8087": 239,
    "8109": 240, "8150": 241, "8168": 242, "8176": 243, "8230": 244,
    "8273": 245, "8281": 248, "8311": 249, "8338": 250, "8346": 251,
    "8451": 252, "8478": 253, "8486": 254, "8494": 255, "8508": 256,
    "8516": 257, "8524": 258, "8532": 259, "8540": 260, "8559": 261,
    "8567": 262, "8575": 263, "8583": 264, "8591": 265, "8605": 266,
    "8613": 267, "8630": 268,
}

PAISES_NOME_PARA_DOMINIO = {
    "AFEGANISTAO": 1, "AFRICA DO SUL": 2, "ALBANIA": 3, "ALEMANHA": 4,
    "ANDORRA": 5, "ANGOLA": 6, "ANGUILLA": 7, "ANTIGUA E BARBUDA": 8,
    "ANTILHAS HOLANDESAS": 9, "ARABIA SAUDITA": 10, "ARGELIA": 11,
    "ARGENTINA": 12, "ARMENIA": 13, "ARUBA": 14, "AUSTRALIA": 15,
    "AUSTRIA": 16, "AZERBAIJAO": 17, "BAHAMAS": 18, "BAHREIN": 19,
    "BANGLADESH": 20, "BARBADOS": 21, "BELARUS": 22, "BELGICA": 23,
    "BELIZE": 24, "BENIN": 25, "BERMUDAS": 26, "BOLIVIA": 27,
    "BOSNIA": 28, "BOTSUANA": 29, "BRASIL": 30, "BRUNEI": 31,
    "BULGARIA": 32, "BURKINA FASO": 33, "BURUNDI": 34, "BUTAO": 35,
    "CABO VERDE": 36, "CAMAROES": 37, "CAMBOJA": 38, "CANADA": 39,
    "CANARIAS": 41, "CATAR": 42, "CAYMAN": 43, "CAZAQUISTAO": 44,
    "CHADE": 45, "CHILE": 46, "CHINA": 47, "CHIPRE": 48,
    "CHRISTMAS": 49, "CINGAPURA": 50, "SINGAPURA": 50,
    "COCOS": 51, "COLOMBIA": 52, "COMORES": 53, "CONGO": 54,
    "COOK": 56, "COREIA DO NORTE": 57, "COREIA DO SUL": 58,
    "COSTA DO MARFIM": 59, "COSTA RICA": 60, "KUWAIT": 61,
    "CROACIA": 62, "CUBA": 63, "DINAMARCA": 64, "DJIBUTI": 65,
    "DOMINICA": 66, "EGITO": 67, "EL SALVADOR": 68,
    "EMIRADOS ARABES UNIDOS": 69, "EQUADOR": 70, "ERITREIA": 71,
    "ESCOCIA": 72, "ESLOVACA": 73, "ESLOVENIA": 74, "ESPANHA": 75,
    "ESTADOS UNIDOS": 76, "ESTONIA": 77, "ETIOPIA": 78,
    "FALKLAND": 79, "FEROE": 80, "FIJI": 81, "FILIPINAS": 82,
    "FINLANDIA": 83, "FORMOSA": 84, "TAIWAN": 84, "FRANCA": 85,
    "GABAO": 86, "GALES": 87, "GAMBIA": 88, "GANA": 89,
    "GEORGIA": 90, "GIBRALTAR": 91, "GRA-BRETANHA": 92,
    "GRANADA": 93, "GRECIA": 94, "GROENLANDIA": 95, "GUADALUPE": 96,
    "GUAM": 97, "GUATEMALA": 98, "GUIANA": 99, "GUIANA FRANCESA": 100,
    "GUINE": 101, "GUINE-BISSAU": 102, "GUINE-EQUATORIAL": 103,
    "HAITI": 104, "HOLANDA": 105, "PAISES BAIXOS": 105,
    "HONDURAS": 106, "HONG KONG": 107, "HUNGRIA": 108, "IEMEN": 109,
    "INDIA": 110, "INDONESIA": 111, "INGLATERRA": 112, "IRA": 113,
    "IRAQUE": 114, "IRLANDA": 115, "IRLANDA DO NORTE": 116,
    "ISLANDIA": 117, "ISRAEL": 118, "ITALIA": 119, "SERVIA": 120,
    "JAMAICA": 121, "JAPAO": 122, "JOHNSTON": 123, "JORDANIA": 124,
    "KIRIBATI": 125, "LAOS": 126, "LEBUAN": 127, "LESOTO": 128,
    "LETONIA": 129, "LIBANO": 130, "LIBERIA": 131, "LIBIA": 132,
    "LIECHTENSTEIN": 133, "LITUANIA": 134, "LUXEMBURGO": 135,
    "MACAU": 136, "MACEDONIA DO NORTE": 137, "MADAGASCAR": 138,
    "MADEIRA": 139, "MALASIA": 140, "MALAVI": 141, "MALDIVAS": 142,
    "MALI": 143, "MALTA": 144, "MAN": 145, "MARIANAS DO NORTE": 146,
    "MARROCOS": 147, "MARSHALL": 148, "MARTINICA": 149,
    "MAURICIO": 150, "MAURITANIA": 151, "MEXICO": 152,
    "MIANMAR": 153, "BIRMANIA": 153, "MICRONESIA": 154,
    "MIDWAY": 155, "MOCAMBIQUE": 156, "MOLDAVIA": 157, "MONACO": 158,
    "MONGOLIA": 159, "MONTSERRAT": 160, "NAMIBIA": 161, "NAURU": 162,
    "NEPAL": 163, "NICARAGUA": 164, "NIGER": 165, "NIGERIA": 166,
    "NIUE": 167, "NORFOLK": 168, "NORUEGA": 169,
    "NOVA CALEDONIA": 170, "NOVA ZELANDIA": 171, "OMA": 172,
    "PALAU": 173, "PANAMA": 174, "PAPUA NOVA GUINE": 175,
    "PAQUISTAO": 176, "PARAGUAI": 177, "PERU": 178, "PITCAIRN": 179,
    "POLINESIA FRANCESA": 180, "POLONIA": 181, "PORTO RICO": 182,
    "PORTUGAL": 183, "QUENIA": 184, "QUIRGUIZ": 185,
    "REINO UNIDO": 186, "REPUBLICA CENTRO-AFRICANA": 187,
    "REPUBLICA DOMINICANA": 188, "REUNIAO": 189, "ROMENIA": 190,
    "RUANDA": 191, "RUSSIA": 192, "SAARA OCIDENTAL": 193,
    "SALOMAO": 194, "SAMOA": 195, "SAMOA AMERICANA": 196,
    "SAN MARINO": 197, "SANTA HELENA": 198, "SANTA LUCIA": 199,
    "SAO CRISTOVAO E NEVES": 200, "SAO PEDRO E MIQUELON": 201,
    "SAO TOME E PRINCIPE": 202, "SAO VICENTE E GRANADINA": 203,
    "SENEGAL": 204, "SERRA LEOA": 205, "SEYCHELLE": 206,
    "SIRIA": 207, "SOMALIA": 208, "SRI LANKA": 209,
    "ESWATINI": 210, "SUAZILANDIA": 210, "SUDAO": 211,
    "SUECIA": 212, "SUICA": 213, "SURINAME": 214,
    "TADJIQUISTAO": 215, "TAILANDIA": 216, "TANZANIA": 217,
    "TCHECA": 218, "TERRITORIO BRITANICO": 219, "TIMOR LESTE": 220,
    "TOGO": 221, "TONGA": 222, "TOQUELAU": 223,
    "TRINIDAD E TOBAGO": 224, "TUNISIA": 225,
    "TURCAS E CAICOS": 226, "TURCOMENISTAO": 227, "TURQUIA": 228,
    "TUVALU": 229, "UCRANIA": 230, "UGANDA": 231, "URUGUAI": 232,
    "UZBEQUISTAO": 233, "VANUATU": 234, "VATICANO": 235,
    "VENEZUELA": 236, "VIETNA": 237, "VIRGENS BRITANICAS": 238,
    "VIRGENS EUA": 239, "WAKE": 240, "WALLIS E FUTUNA": 241,
    "ZAMBIA": 242, "ZIMBABUE": 243, "ZONA DO CANAL DO PANAMA": 244,
    "MONTENEGRO": 245, "QATAR": 249, "SAINT KITTS E NEVIS": 250,
    "CURACAO": 256, "MAYOTTE": 261, "PALESTINA": 266,
    "SUDAO DO SUL": 267,
}

def resolver_codigo_pais_dominio(c_pais_xml: str, x_pais_xml: str) -> str:
    c_pais_norm = (c_pais_xml or "").strip().zfill(4)
    if c_pais_norm in PAISES_BACEN_PARA_DOMINIO:
        return str(PAISES_BACEN_PARA_DOMINIO[c_pais_norm])
    if x_pais_xml:
        nome_upper = x_pais_xml.upper().strip()
        if nome_upper in PAISES_NOME_PARA_DOMINIO:
            return str(PAISES_NOME_PARA_DOMINIO[nome_upper])
        for chave, cod in PAISES_NOME_PARA_DOMINIO.items():
            if chave in nome_upper or nome_upper in chave:
                return str(cod)
    return c_pais_xml or ""

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

def extrair_chave_nfe(nfe_root) -> str:
    inf_nfe = nfe_root.find("nfe:infNFe", NS)
    if inf_nfe is None:
        return ""
    id_attr = inf_nfe.get("Id", "")
    chave = re.sub(r"^NFe", "", id_attr)
    chave_num = re.sub(r"[^0-9]", "", chave)
    return chave_num if len(chave_num) == 44 else chave

def safe_float(v: str) -> float:
    try:
        return float(str(v).replace(",", ".") or "0")
    except (ValueError, TypeError):
        return 0.0

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
# REGISTRO 0020
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
        c_pais_xml  = get_text(ender, "nfe:cPais")                   if ender is not None else ""
        x_pais_xml  = get_text(ender, "nfe:xPais")                   if ender is not None else ""
        cod_pais    = resolver_codigo_pais_dominio(c_pais_xml, x_pais_xml)
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
        "0020", inscricao, razao, fantasia, logradouro, numero, complemento,
        bairro, cod_mun, uf_campo, cod_pais, cep, ie, "", "", "", "", "",
        "", "", "", "N", "7", regime, contrib, "", "", "", "", "N", "N", "", "",
    ])

# ─────────────────────────────────────────────
# EXTRAIR PIS/COFINS
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
        "0110", "Inicial", pc["cst_e"], "", "01", "N", "N",
        pc["aliq_pis_e"], pc["aliq_cof_e"], "N", "N", "", "", "", "",
        pc["cst_s"], "N", "", "", "", "N", pc["aliq_pis_s"], pc["aliq_cof_s"],
        "N", "N", "", "", "", "", "", "", "N", "N", "", "", "", "", "", "M",
        "", "N", "N", "N", "", "N", "", "N", "", "", "", "", "", "", "N",
        "", "", "N", "", "", "N", "", "N", "N", "N", ct, ct, "N", "N",
    ])

# ─────────────────────────────────────────────
# REGISTRO 1000 – 98 campos
# ─────────────────────────────────────────────
def gerar_registro_1000(nfe_root, cnpj_empresa: str,
                        acumulador: str = "1157",
                        especie: str = "36",
                        importacao: bool = False) -> str:
    ide   = nfe_root.find("nfe:infNFe/nfe:ide", NS)
    emit  = nfe_root.find("nfe:infNFe/nfe:emit", NS)
    dest  = nfe_root.find("nfe:infNFe/nfe:dest", NS)
    total = nfe_root.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)

    if importacao and dest is not None:
        id_ext    = get_text(dest, "nfe:idEstrangeiro").strip()
        cnpj_forn = id_ext if id_ext else ""
    else:
        cnpj_forn = get_text(emit, "nfe:CNPJ")

    emitente_nf = "P" if importacao else "T"
    ie_forn     = "" if importacao else get_text(emit, "nfe:IE")

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

    chave = extrair_chave_nfe(nfe_root)

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

    tipo_doc_importacao = "1" if importacao else ""

    campos = [
        "1000",               # 1
        especie,              # 2
        cnpj_forn,            # 3
        "",                   # 4
        acumulador,           # 5
        cfop_first,           # 6
        "",                   # 7
        nNF,                  # 8
        serie,                # 9
        "",                   # 10
        dhEmi,                # 11
        dhEmi,                # 12
        v_nf,                 # 13
        "",                   # 14
        obs_fisco,            # 15
        mod_frete,            # 16
        emitente_nf,          # 17
        "",                   # 18
        "",                   # 19
        "",                   # 20
        "",                   # 21
        "",                   # 22
        "",                   # 23
        "",                   # 24
        "",                   # 25
        v_frete,              # 26
        v_seg,                # 27
        v_outro,              # 28
        v_pis,                # 29
        "",                   # 30
        v_cofins,             # 31
        "",                   # 32
        "",                   # 33
        "",                   # 34
        "",                   # 35
        "",                   # 36
        "",                   # 37
        "",                   # 38
        v_prod,               # 39
        c_mun_fg,             # 40
        "0",                  # 41
        "",                   # 42
        "",                   # 43
        ie_forn,              # 44
        "",                   # 45
        "",                   # 46
        "",                   # 47
        "",                   # 48
        "",                   # 49
        "",                   # 50
        "",                   # 51
        n_di,                 # 52
        "N",                  # 53
        chave,                # 54
        "",                   # 55
        "",                   # 56
        "",                   # 57
        "",                   # 58
        "",                   # 59
        "1",                  # 60
        "",                   # 61
        "",                   # 62
        "",                   # 63
        "",                   # 64
        "",                   # 65
        "",                   # 66
        "",                   # 67
        "",                   # 68
        "",                   # 69
        tipo_doc_importacao,  # 70
        "",                   # 71
        "",                   # 72
        "",                   # 73
        "",                   # 74
        "",                   # 75
        "",                   # 76
        "",                   # 77
        "",                   # 78
        "",                   # 79
        "",                   # 80
        "",                   # 81
        "",                   # 82
        "",                   # 83
        "",                   # 84
        "",                   # 85
        "",                   # 86
        "",                   # 87
        "",                   # 88
        "",                   # 89
        v_ipi,                # 90
        v_st,                 # 91
        "",                   # 92
        "",                   # 93
        "",                   # 94
        "",                   # 95
        "",                   # 96
        v_icms_d,             # 97
        "",                   # 98
    ]
    assert len(campos) == 98, f"1000: esperado 98, gerado {len(campos)}"
    return pipe_join(campos)

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
# REGISTROS 1020 – por alíquota
# ─────────────────────────────────────────────
def gerar_registros_1020(nfe_root) -> list:
    total    = nfe_root.find("nfe:infNFe/nfe:total/nfe:ICMSTot", NS)
    det_list = nfe_root.findall("nfe:infNFe/nfe:det", NS)
    v_nf     = fmt_decimal(get_text(total, "nfe:vNF"))
    linhas   = []

    def r1020(cod, perc_red="", base="", aliq="", valor="",
              isentas="", outras="", v_ipi="", v_st="", v_cont="", cod_rec=""):
        return pipe_join([
            "1020", cod, perc_red, base, aliq, valor,
            isentas, outras, v_ipi, v_st, v_cont, cod_rec,
            "", "", "", "", "", "", "",
        ])

    # ── ICMS: agrupa por alíquota ─────────────────────────────────────
    icms_por_aliq = {}
    for det in det_list:
        imp = det.find("nfe:imposto", NS)
        if imp is None:
            continue
        for tp in ["ICMS00","ICMS10","ICMS20","ICMS30","ICMS40",
                   "ICMS51","ICMS60","ICMS70","ICMS90","ICMSSN101",
                   "ICMSSN102","ICMSSN201","ICMSSN202","ICMSSN500","ICMSSN900"]:
            node = imp.find(f"nfe:ICMS/nfe:{tp}", NS)
            if node is not None:
                aliq_str = get_text(node, "nfe:pICMS") or "0"
                bc       = safe_float(get_text(node, "nfe:vBC"))
                valor    = safe_float(get_text(node, "nfe:vICMS"))
                deson    = safe_float(get_text(node, "nfe:vICMSDeson"))
                key = aliq_str
                if key not in icms_por_aliq:
                    icms_por_aliq[key] = {"bc": 0.0, "valor": 0.0, "deson": 0.0}
                icms_por_aliq[key]["bc"]    += bc
                icms_por_aliq[key]["valor"] += valor
                icms_por_aliq[key]["deson"] += deson
                break

    v_ipi_tot = fmt_decimal(get_text(total, "nfe:vIPI"))
    v_st_tot  = fmt_decimal(get_text(total, "nfe:vST"))
    for aliq_str, dados in sorted(icms_por_aliq.items(),
                                   key=lambda x: safe_float(x[0])):
        if dados["valor"] > 0 or dados["bc"] > 0:
            linhas.append(r1020(
                1,
                base  = fmt_decimal(str(dados["bc"])),
                aliq  = fmt_decimal(aliq_str),
                valor = fmt_decimal(str(dados["valor"])),
                v_ipi = v_ipi_tot,
                v_st  = v_st_tot,
                v_cont= v_nf,
            ))

    # ── IPI: agrupa por alíquota ──────────────────────────────────────
    # Estrutura: aliq_str → {"bc": float, "valor": float, "isentas": float}
    ipi_por_aliq = {}
    for det in det_list:
        imp = det.find("nfe:imposto", NS)
        if imp is None:
            continue
        ipi_trib = imp.find("nfe:IPI/nfe:IPITrib", NS)
        ipi_nt   = imp.find("nfe:IPI/nfe:IPINT", NS)
        if ipi_trib is not None:
            aliq_str = get_text(ipi_trib, "nfe:pIPI") or "0"
            bc       = safe_float(get_text(ipi_trib, "nfe:vBC"))
            valor    = safe_float(get_text(ipi_trib, "nfe:vIPI"))
            key = aliq_str
            if key not in ipi_por_aliq:
                ipi_por_aliq[key] = {"bc": 0.0, "valor": 0.0, "isentas": 0.0}
            ipi_por_aliq[key]["bc"]    += bc
            ipi_por_aliq[key]["valor"] += valor
        elif ipi_nt is not None:
            # NT: base isenta = vProd do item
            v_prod_item = safe_float(get_text(det.find("nfe:prod", NS), "nfe:vProd"))
            key = "0"
            if key not in ipi_por_aliq:
                ipi_por_aliq[key] = {"bc": 0.0, "valor": 0.0, "isentas": 0.0}
            ipi_por_aliq[key]["bc"]      += v_prod_item
            ipi_por_aliq[key]["isentas"] += v_prod_item

    # Gera linhas 1020 IPI — NT primeiro (aliq=0), depois tributadas
    for aliq_str, dados in sorted(ipi_por_aliq.items(),
                                   key=lambda x: safe_float(x[0])):
        aliq_f = safe_float(aliq_str)
        if aliq_f == 0 and dados["isentas"] > 0:
            # Linha NT/isentas: base = isentas, aliq e valor vazios
            linhas.append(r1020(
                2,
                base    = fmt_decimal(str(dados["bc"])),
                aliq    = "",
                valor   = "",
                isentas = fmt_decimal(str(dados["isentas"])),
                v_cont  = v_nf,
            ))
        elif dados["valor"] > 0 or (dados["bc"] > 0 and aliq_f > 0):
            linhas.append(r1020(
                2,
                base  = fmt_decimal(str(dados["bc"])),
                aliq  = fmt_decimal(aliq_str),
                valor = fmt_decimal(str(dados["valor"])),
                v_cont= v_nf,
            ))

    # ── PIS: agrupa por alíquota ──────────────────────────────────────
    pis_por_aliq = {}
    for det in det_list:
        imp = det.find("nfe:imposto", NS)
        if imp is None:
            continue
        pis_node = imp.find("nfe:PIS", NS)
        if pis_node is not None:
            for pt in ["PISAliq", "PISQtde", "PISNT", "PISOutr"]:
                pn = pis_node.find(f"nfe:{pt}", NS)
                if pn is not None:
                    aliq_str = get_text(pn, "nfe:pPIS") or "0"
                    bc       = safe_float(get_text(pn, "nfe:vBC"))
                    valor    = safe_float(get_text(pn, "nfe:vPIS"))
                    key = aliq_str
                    if key not in pis_por_aliq:
                        pis_por_aliq[key] = {"bc": 0.0, "valor": 0.0}
                    pis_por_aliq[key]["bc"]    += bc
                    pis_por_aliq[key]["valor"] += valor
                    break

    for aliq_str, dados in sorted(pis_por_aliq.items(),
                                   key=lambda x: safe_float(x[0])):
        if dados["valor"] > 0 or dados["bc"] > 0:
            linhas.append(r1020(
                4,
                base  = fmt_decimal(str(dados["bc"])),
                aliq  = fmt_decimal(aliq_str),
                valor = fmt_decimal(str(dados["valor"])),
                v_cont= v_nf,
            ))

    # ── COFINS: agrupa por alíquota ───────────────────────────────────
    cof_por_aliq = {}
    for det in det_list:
        imp = det.find("nfe:imposto", NS)
        if imp is None:
            continue
        cof_node = imp.find("nfe:COFINS", NS)
        if cof_node is not None:
            for ct in ["COFINSAliq", "COFINSQtde", "COFINSNT", "COFINSOutr"]:
                cn = cof_node.find(f"nfe:{ct}", NS)
                if cn is not None:
                    aliq_str = get_text(cn, "nfe:pCOFINS") or "0"
                    bc       = safe_float(get_text(cn, "nfe:vBC"))
                    valor    = safe_float(get_text(cn, "nfe:vCOFINS"))
                    key = aliq_str
                    if key not in cof_por_aliq:
                        cof_por_aliq[key] = {"bc": 0.0, "valor": 0.0}
                    cof_por_aliq[key]["bc"]    += bc
                    cof_por_aliq[key]["valor"] += valor
                    break

    for aliq_str, dados in sorted(cof_por_aliq.items(),
                                   key=lambda x: safe_float(x[0])):
        if dados["valor"] > 0 or dados["bc"] > 0:
            linhas.append(r1020(
                5,
                base  = fmt_decimal(str(dados["bc"])),
                aliq  = fmt_decimal(aliq_str),
                valor = fmt_decimal(str(dados["valor"])),
                v_cont= v_nf,
            ))

    # ── ICMS Desonerado: agrupa por alíquota (cód 45) ─────────────────
    for aliq_str, dados in sorted(icms_por_aliq.items(),
                                   key=lambda x: safe_float(x[0])):
        if dados["deson"] > 0:
            linhas.append(r1020(
                45,
                base  = fmt_decimal(str(dados["bc"])),
                aliq  = fmt_decimal(aliq_str),
                valor = fmt_decimal(str(dados["deson"])),
                v_cont= v_nf,
            ))

    # ── PIS SPED (133) e COFINS SPED (134): totais únicos ────────────
    v_pis_tot    = get_text(total, "nfe:vPIS")
    v_cofins_tot = get_text(total, "nfe:vCOFINS")
    bc_pis_total = bc_cof_total = 0.0
    for det in det_list:
        imp = det.find("nfe:imposto", NS)
        if imp is None:
            continue
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

    if v_pis_tot and safe_float(v_pis_tot) > 0:
        linhas.append(r1020(133,
            base  = fmt_decimal(str(bc_pis_total)),
            valor = fmt_decimal(v_pis_tot),
            v_cont= v_nf))
    if v_cofins_tot and safe_float(v_cofins_tot) > 0:
        linhas.append(r1020(134,
            base  = fmt_decimal(str(bc_cof_total)),
            valor = fmt_decimal(v_cofins_tot),
            v_cont= v_nf))

    return linhas

# ─────────────────────────────────────────────
# REGISTRO 1030 – 111 campos
# V3.9: assert removido, campos contados explicitamente 1 a 111
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
    v_ipi = aliq_ipi = cst_ipi = ""
    v_pis = aliq_pis = cst_pis = bc_pis = ""
    v_cof = aliq_cof = cst_cof = bc_cof = ""
    ibs_class_trib = ibs_bc = ibs_aliq = ibs_val = ""
    cbs_class_trib = cbs_bc = cbs_aliq = cbs_val = ""

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
            cbs_class_trib = ibs_class_trib
            gibs = ibs_node.find("nfe:gIBSCBS", NS)
            if gibs is not None:
                ibs_bc = fmt_decimal(get_text(gibs, "nfe:vBC"))
                cbs_bc = ibs_bc
                guf = gibs.find("nfe:gIBSUF", NS)
                if guf is not None:
                    ibs_aliq = fmt_decimal(get_text(guf, "nfe:pIBSUF"))
                    ibs_val  = fmt_decimal(get_text(guf, "nfe:vIBSUF"))
                gcbs = gibs.find("nfe:gCBS", NS)
                if gcbs is not None:
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

    # ── 111 campos exatos conforme layout 46 - Registro 1030 ─────────
    campos = [
        "1030",                   # 1
        cod_prod,                 # 2
        qtd,                      # 3
        v_total,                  # 4
        v_ipi,                    # 5
        fmt_decimal(v_prod),      # 6
        "1",                      # 7
        d_di,                     # 8
        n_di,                     # 9
        cst_icms,                 # 10
        fmt_decimal(v_prod),      # 11
        fmt_decimal(v_desc),      # 12
        v_bc_icms,                # 13
        v_bc_st,                  # 14
        aliq_icms,                # 15
        "",                       # 16
        "",                       # 17
        "",                       # 18
        "",                       # 19
        fmt_decimal(v_outro),     # 20
        "",                       # 21
        v_icms,                   # 22
        "",                       # 23
        "",                       # 24
        "",                       # 25
        "",                       # 26
        fmt_decimal(v_unit, 6),   # 27
        "",                       # 28
        cst_ipi,                  # 29
        aliq_ipi,                 # 30
        "",                       # 31
        "",                       # 32
        "",                       # 33
        cfop,                     # 34
        "",                       # 35
        aliq_pis,                 # 36
        v_pis,                    # 37
        aliq_cof,                 # 38
        v_cof,                    # 39
        fmt_decimal(v_prod),      # 40
        cst_pis,                  # 41
        bc_pis,                   # 42
        cst_cof,                  # 43
        bc_cof,                   # 44
        "",                       # 45
        "",                       # 46
        "",                       # 47
        "",                       # 48
        "",                       # 49
        "",                       # 50
        "",                       # 51
        "",                       # 52
        "",                       # 53
        "",                       # 54
        "",                       # 55
        "S",                      # 56
        unidade,                  # 57
        "",                       # 58
        "",                       # 59
        fmt_decimal(v_prod),      # 60
        "",                       # 61
        "",                       # 62
        "",                       # 63
        "",                       # 64
        "",                       # 65
        "",                       # 66
        "",                       # 67
        "",                       # 68
        "",                       # 69
        "",                       # 70
        "",                       # 71
        "",                       # 72
        "",                       # 73
        "",                       # 74
        "",                       # 75
        "",                       # 76
        "",                       # 77
        "",                       # 78
        "",                       # 79
        "",                       # 80
        "",                       # 81
        "",                       # 82
        "",                       # 83
        "",                       # 84
        "",                       # 85
        "",                       # 86
        "",                       # 87
        "",                       # 88
        "",                       # 89
        "",                       # 90
        cest,                     # 91
        "",                       # 92
        "",                       # 93
        "",                       # 94
        "",                       # 95
        "",                       # 96
        v_icms_des,               # 97
        "",                       # 98
        "",                       # 99
        "",                       # 100
        "",                       # 101
        "",                       # 102
        "",                       # 103
        ibs_class_trib,           # 104
        ibs_bc,                   # 105
        ibs_aliq,                 # 106
        ibs_val,                  # 107
        cbs_class_trib,           # 108
        cbs_bc,                   # 109
        cbs_aliq,                 # 110
        cbs_val,                  # 111
    ]
    # Validação de segurança — não lança assert, apenas corrige silenciosamente
    if len(campos) != 111:
        if len(campos) < 111:
            campos.extend([""] * (111 - len(campos)))
        else:
            campos = campos[:111]
    return pipe_join(campos)

# ─────────────────────────────────────────────
# REGISTRO 1097
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
        "1097", mod_frete, tp_via, frete_conta, "", "", "", "", "", "",
        razao_transp[:150], tipo_insc, cnpj_transp, ie_transp, end_transp,
        "", "", "", cidade_cod, uf_transp, "",
        q_vol, esp_vol, marca, "", peso_l, peso_b,
        "E", "", "", "", "D", "", "", "",
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
        ender_dest = dest_node.find("nfe:enderDest", NS)
        c_pais_xml = get_text(ender_dest, "nfe:cPais") if ender_dest is not None else ""
        x_pais_xml = get_text(ender_dest, "nfe:xPais") if ender_dest is not None else ""
        cod_pais_dominio = resolver_codigo_pais_dominio(c_pais_xml, x_pais_xml)
    else:
        nome_forn        = get_text(emit, "nfe:xNome") if emit is not None else ""
        uf_forn          = ""
        cod_pais_dominio = ""
        if emit is not None:
            ender_e = emit.find("nfe:enderEmit", NS)
            uf_forn = get_text(ender_e, "nfe:UF") if ender_e is not None else ""

    chave_resumo = extrair_chave_nfe(nfe)

    resumo["nNF"]            = get_text(ide, "nfe:nNF")
    resumo["Emitente"]       = get_text(emit, "nfe:xNome") if emit is not None else ""
    resumo["CNPJ Emit"]      = get_text(emit, "nfe:CNPJ")  if emit is not None else ""
    resumo["Fornecedor"]     = nome_forn
    resumo["UF Forn"]        = uf_forn
    resumo["CNPJ Empresa"]   = cnpj_empresa
    resumo["Origem CNPJ"]    = origem_cnpj
    resumo["Importacao"]     = "Sim" if importacao else "Nao"
    resumo["Emitente NF"]    = "P (Proprio)" if importacao else "T (Terceiros)"
    resumo["Emissao"]        = fmt_date(get_text(ide, "nfe:dhEmi"))
    resumo["Itens"]          = len(det_list)
    resumo["vNF"]            = fmt_decimal(get_text(total, "nfe:vNF"))
    resumo["vICMS"]          = fmt_decimal(get_text(total, "nfe:vICMS"))
    resumo["vICMSDes"]       = fmt_decimal(get_text(total, "nfe:vICMSDeson"))
    resumo["vIPI"]           = fmt_decimal(get_text(total, "nfe:vIPI"))
    resumo["vPIS"]           = fmt_decimal(get_text(total, "nfe:vPIS"))
    resumo["vCOFINS"]        = fmt_decimal(get_text(total, "nfe:vCOFINS"))
    resumo["Chave NF-e"]     = chave_resumo
    resumo["Cod Pais (Dom)"] = cod_pais_dominio
    resumo["Grupo"]          = (
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
        <h4>V3.9-FINAL — Corrigido AssertionError no Registro 1030</h4>
        <ul>
          <li><b>Causa raiz</b>: o <code>assert len(campos) == 111</code> lançava erro porque versões anteriores do código tinham campos faltando entre as posições 56-103</li>
          <li><b>Correção</b>: lista de 111 campos reescrita explicitamente (campos 1 a 111) com autocorreção silenciosa ao invés de assert</li>
          <li><b>1030 campos 56-103</b>: todos os 48 campos intermediários agora estão presentes e numerados</li>
          <li>Campos 104-111 (IBS/CBS) mantidos corretamente</li>
        </ul>
        <h4>V3.8-FINAL — 1020: uma linha por alíquota (ICMS e IPI)</h4>
        <h4>V3.7-FINAL — Contagem exata 98 campos no 1000</h4>
        <h4>V3.6-FINAL — Removidos 183/184 do 1020</h4>
        <h4>V3.5-FINAL — Campos numéricos/decimais no 0100 e 0110</h4>
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
                pais_info = f" | Pais Dom.: {r.get('Cod Pais (Dom)', '')}" if is_imp else ""
                with cols[idx]:
                    st.markdown(
                        f'<div class="cnpj-badge" style="color:{cor};border-color:{cor};">'
                        f'Empresa: {r["CNPJ Empresa"]}</div>'
                        f'<div class="info-origem" style="border-left-color:{cor};">'
                        f'{r.get("Origem CNPJ","")}<br>'
                        f'{"Importacao | Forn: " + r.get("Fornecedor","")[:40] + pais_info if is_imp else "Forn: " + r.get("Fornecedor","")[:50]}'
                        f'<br><small>Emitente NF: {r.get("Emitente NF","")}</small>'
                        f'<br><small>Chave: {r.get("Chave NF-e","")[:22]}...</small>'
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
