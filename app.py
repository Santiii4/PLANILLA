import io
import hashlib
import os
import re
from datetime import datetime
import pandas as pd
import pypdf
import streamlit as st

# Módulo para Google Sheets
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# ==============================================================================
# CONFIGURACIÓN DE RUTAS Y AUTENTICACIÓN
# ==============================================================================
DIRECTORIO_APP = os.path.dirname(os.path.abspath(__file__))

def obtener_ruta_credenciales():
    """Busca el archivo credentials.json localmente si existe."""
    posibles_nombres = [
        "credentials.json",
        "credentials.json.json",
        "credentials",
    ]
    for nombre in posibles_nombres:
        ruta_script = os.path.join(DIRECTORIO_APP, nombre)
        if os.path.exists(ruta_script):
            return ruta_script
        if os.path.exists(nombre):
            return os.path.abspath(nombre)
    return None

def verificar_credenciales_disponibles():
    """Verifica si hay credenciales en Streamlit Secrets o en archivo local."""
    try:
        if "gcp_service_account" in st.secrets:
            return True
    except Exception:
        pass
    return obtener_ruta_credenciales() is not None

# ==============================================================================
# 1. CONFIGURACIÓN DE LA PÁGINA
# ==============================================================================
st.set_page_config(
    page_title="Control de Cargas & Google Sheets",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==============================================================================
# 2. MOTOR DE EXTRACCIÓN AVANZADO (MIC/DTA & CRT)
# ==============================================================================
def extraer_texto_pdf(archivo_pdf) -> str:
    """Lee texto seleccionable; los errores los muestra la interfaz por archivo."""
    archivo_pdf.seek(0)
    lector = pypdf.PdfReader(archivo_pdf)
    return "\n".join(pagina.extract_text() or "" for pagina in lector.pages).strip()

"""Extracción conservadora de manifiestos con texto, sin servicios externos."""

import re
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation


_COLUMNAS_MANIFIESTO = (
    "ORIGEN", "ADUANA DESTINO", "ADUANA DE SALIDA", "EXPORTADOR", "IMPORTADOR",
    "fecha", "MIC ELEC.", "CRT", "FACTURA", "VALOR", "FLETE EN REALES", "FRETE",
    "TRACTOR", "CARRETA", "CHOFER", "DNI", "SEGURO",
)


def _normalizar_busqueda(texto):
    # Un carácter por carácter mantiene los índices del texto original.
    return "".join(
        next((c for c in unicodedata.normalize("NFD", letra) if not unicodedata.combining(c)), letra)
        for letra in texto
    ).upper()


def _rotulo(patron, numeros=""):
    prefijo = rf"(?:(?:{numeros})[.)]?\s+)?" if numeros else ""
    return rf"(?<!\w){prefijo}(?:{patron})(?!\w)"


# Los encabezados también delimitan los campos anteriores. No se busca un
# importe, una patente o una aduana en todo el resto del documento.
_ROTULOS_MANIFIESTO = [
    ("ignorar", _rotulo(r"(?:FECHA|DATA)\s*(?:DE\s*)?(?:NACIMIENTO|NASCIMENTO|VENCIMIENTO|VENCIMENTO)")),
    ("origen", _rotulo(r"ADUANA(?:\s*,?\s*CIUDAD\s*Y\s*PAIS)?\s*(?:DE\s*)?PARTIDA(?:\s*/\s*ALFANDEGA(?:\s*,?\s*CIDADE\s*E\s*PAIS)?\s*DE\s*PARTIDA)?|PAIS\s*DE\s*ORIGEN|ORIGEN", "7|26")),
    ("destino", _rotulo(r"ADUANA\s*(?:DE\s*)?DESTINO(?:\s*/\s*ALFANDEGA\s*(?:DE\s*)?DESTINO)?|CIUDAD\s*Y\s*PAIS\s*DE\s*DESTINO(?:\s*FINAL)?", "24|8")),
    ("salida", _rotulo(r"ADUANA\s*(?:DE\s*)?(?:SALIDA|FRONTERA)|PASO\s*FRONTERIZO")),
    ("ruta", _rotulo(r"RUTA(?:\s*(?:Y\s*PLAZO\s*DE\s*TRANSPORTE|PREVISTA|DE\s*TRANSPORTE))?|ITINERARIO", "40")),
    ("exportador", _rotulo(r"(?:REMITENTE|REMETENTE|EXPORTADOR)(?:\s*/\s*(?:REMITENTE|REMETENTE|EXPORTADOR))?", "33|6|1")),
    ("importador", _rotulo(r"(?:DESTINATARIO|IMPORTADOR)(?:\s*/\s*(?:DESTINATARIO|IMPORTADOR))?", "34|7|4")),
    ("fecha", _rotulo(r"F\.?\s*OFIC\.?|(?:FECHA|DATA)(?:\s*(?:DE\s*)?(?:EMISION|EMISSAO|OFICIALIZACION))?")),
    ("crt", _rotulo(r"(?:N[°º?O.]?\s*(?:DE\s*)?|NUMEROS?\s*(?:DE\s*)?)CARTAS?\s*DE\s*PORTE|CRTS?", "23")),
    ("crt", r"(?<!\w)2[.)]?\s+NUMERO(?!\w)"),
    ("factura", _rotulo(r"FACTURA(?:\s*COMERCIAL)?|FATURA(?:\s*COMERCIAL)?|INVOICE")),
    ("valor", _rotulo(r"(?:MONEDA\s*Y\s*)?VALOR\s*FO[BT]", "27|15")),
    ("valor", r"(?<!\w)14[.)]?\s+VALOR(?!\w)"),
    ("reales", _rotulo(r"(?:FLETE|FRETE)\s*(?:(?:EN|EM)\s*)?(?:REALES|REAIS|BRL|R\$)", "28")),
    ("flete", _rotulo(r"(?:FLETE|FRETE)(?:\s*/\s*(?:FLETE|FRETE))?(?:\s*(?:EN|EM)\s*(?:US[S$]|USD|DOLARES))?", "28")),
    ("seguro", _rotulo(r"SEGURO(?:\s*(?:/|X)\s*SEGURO)?(?:\s*(?:(?:EN|EM)\s*)?(?:US[S$]|USD))?", "29")),
    ("tractor", _rotulo(r"PLACA\s*(?:DE[L]?\s*)?(?:CAMION|CAMINHAO|TRACTOR)|PATENTE\s*(?:DE[L]?\s*)?(?:CAMION|TRACTOR)|TRACTOR", "11|18")),
    ("carreta", _rotulo(r"SEMIR?REMOLQUE|SEMIREMOLQUE|SEMI[-\s]?REBOQUE|CARRETA|ACOPLADO", "15|20")),
    ("chofer", _rotulo(r"CONDUCTOR(?:\s*1)?|CHOFER|MOTORISTA")),
    ("dni", _rotulo(r"DOC(?:UMENTO)?(?:\s*(?:DE\s*IDENTIDAD|DEL\s*CHOFER))?|DNI|CEDULA(?:\s*DE\s*IDENTIDAD)?")),
    ("ignorar", _rotulo(r"MIC(?:\s*ELEC(?:TRONICO)?\.?)?(?:\s*/\s*DTA)?|CONSIGNATARIO|PESO(?:\s*BRUTO|\s*NETO)?|BULTOS|PRECINTOS?|MARCAS\s*Y\s*NUMEROS|DESCRIPCION\s*DE\s*(?:LA\s*)?MERCADERIA|GASTOS\s*A\s*PAGAR|OTROS\s*GASTOS|TRANSPORTISTA|TRANSPORTADOR|PERMISO|CUIT|CNPJ|RUT", "35|30|31|32|36|37|38|39|15|1|2|3|4|5")),
]
_ENCABEZADOS_MANIFIESTO = re.compile("|".join(
    rf"(?P<c{i}>{patron})" for i, (_, patron) in enumerate(_ROTULOS_MANIFIESTO)
))
_NUMERO_IMPORTE = r"[+-]?(?:\d{1,3}(?:[ \u00a0]\d{3})+(?:[.,]\d{1,2})?|\d+(?:[.,]\d+)*)"
_MONEDAS = r"(?:USD|US\$|U\$S|USS|BRL|R\$|EUR|ARS|REALES|REAIS|PESOS)"


def _secciones_manifiesto(texto):
    normalizado = _normalizar_busqueda(texto)
    matches = list(_ENCABEZADOS_MANIFIESTO.finditer(normalizado))
    secciones = {}
    for indice, match in enumerate(matches):
        tipo = _ROTULOS_MANIFIESTO[int(match.lastgroup[1:])][0]
        fin = matches[indice + 1].start() if indice + 1 < len(matches) else len(texto)
        contenido = texto[match.end():fin]
        # Un campo numerado desconocido en una línea nueva también es límite.
        corte = re.search(r"(?m)^\s*\d{1,2}[.)]?\s+[A-Za-zÁÉÍÓÚÑÜáéíóúñü]", contenido)
        if corte:
            contenido = contenido[:corte.start()]
        secciones.setdefault(tipo, []).append((texto[match.start():match.end()], contenido))
    return secciones


def _primera_linea_campo(contenido):
    contenido = re.sub(r"^\s*[:/\-–.]*\s*", "", contenido)
    contenido = re.sub(r"^\([^\n)]*\)\s*[:/\-–.]*\s*", "", contenido)
    for linea in contenido.splitlines():
        linea = linea.strip(" \t:/–")
        if not linea:
            continue
        if _normalizar_busqueda(linea).strip(" .:-") in {
            "NOMBRE Y DOMICILIO", "NOME E ENDERECO", "NOMBRE Y DIRECCION",
            "NOMBRE", "NOME", "NOMBRE Y DOMICILIO / NOME E ENDERECO",
        }:
            continue
        return re.sub(r"\s+", " ", linea)
    return ""


def _localidad_campo(contenido, quitar_pais=False):
    for linea in contenido.splitlines():
        linea = linea.strip(" \t:/–")
        if not linea or not re.search(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]", linea):
            continue
        linea = re.sub(r"^\d+\s*[-.]?\s*", "", linea)
        if quitar_pais:
            linea = re.sub(r"\s*[-/]\s*(?:ARGENTINA|BRASIL|CHILE|URUGUAY|PARAGUAY|BOLIVIA)\s*$", "", linea, flags=re.I)
        return re.sub(r"\s+", " ", linea).strip()
    return ""


def _decimal_manifiesto(token):
    """Admite miles locales/internacionales; rechaza agrupaciones ambiguas."""
    token = re.sub(r"[ \u00a0]", "", token)
    signo = ""
    if token.startswith(("+", "-")):
        signo, token = token[0], token[1:]
    if not token or not re.fullmatch(r"\d+(?:[.,]\d+)*", token):
        return None
    if "." in token and "," in token:
        decimal = "." if token.rfind(".") > token.rfind(",") else ","
        miles = "," if decimal == "." else "."
        entero, fraccion = token.rsplit(decimal, 1)
        if decimal in entero or len(fraccion) not in (1, 2):
            return None
        grupos = entero.split(miles)
        if not (1 <= len(grupos[0]) <= 3 and all(len(g) == 3 for g in grupos[1:])):
            return None
        canonico = "".join(grupos) + "." + fraccion
    elif "." in token or "," in token:
        separador = "." if "." in token else ","
        grupos = token.split(separador)
        if len(grupos) == 2 and len(grupos[-1]) in (1, 2):
            canonico = grupos[0] + "." + grupos[1]
        elif 1 <= len(grupos[0]) <= 3 and all(len(g) == 3 for g in grupos[1:]):
            canonico = "".join(grupos)
        else:
            return None
    else:
        canonico = token
    try:
        return Decimal(signo + canonico).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None


def _importe_campo(rotulo, contenido, moneda):
    # El valor debe estar inmediatamente debajo/junto al encabezado. Una nota
    # posterior con otro importe no puede completar un campo que estaba vacío.
    lineas = [linea.strip(" \t:") for linea in contenido.splitlines() if linea.strip(" \t:")]
    if not lineas:
        return None
    texto = _normalizar_busqueda(lineas[0])
    if re.fullmatch(_MONEDAS, texto) and len(lineas) > 1:
        texto += " " + _normalizar_busqueda(lineas[1])
    todas = list(re.finditer(_MONEDAS, texto))
    deseada = r"(?:USD|US\$|U\$S|USS)" if moneda == "USD" else r"(?:BRL|R\$|REALES|REAIS)"
    # Si se declara una moneda en el valor, solamente se acepta la pedida.
    if todas:
        patrones = (
            rf"{deseada}\s*[:$]?\s*({_NUMERO_IMPORTE})",
            rf"({_NUMERO_IMPORTE})\s*{deseada}",
        )
        for patron in patrones:
            match = re.fullmatch(patron, texto)
            if match:
                valor = _decimal_manifiesto(match.group(1))
                if valor is not None:
                    return valor
        return None
    rotulo_normalizado = _normalizar_busqueda(rotulo)
    moneda_rotulo = re.search(_MONEDAS, rotulo_normalizado)
    if moneda_rotulo and not re.fullmatch(deseada, moneda_rotulo.group()):
        return None
    # Primera línea con contenido, evitando buscar números en notas posteriores.
    match = re.fullmatch(rf"\s*[$:]?\s*({_NUMERO_IMPORTE})\s*", texto)
    return _decimal_manifiesto(match.group(1)) if match else None


def _formatear_importe(valor, moneda):
    return moneda + " " + format(valor, ",.2f").replace(",", "@").replace(".", ",").replace("@", ".")


def _identificador_campo(contenido):
    linea = _primera_linea_campo(contenido)
    linea = re.sub(r"^(?:NRO\.?|NR\.?|NUMERO|N[°º?O.]?)\s*[:.-]?\s*", "", linea, flags=re.I)
    match = re.match(r"([A-Za-z0-9]+(?:[./-][A-Za-z0-9]+)*)", linea)
    if match and re.search(r"\d", match.group(1)):
        return match.group(1)
    return ""


def _mic_manifiesto(texto):
    normalizado = _normalizar_busqueda(texto)
    match = re.search(
        r"(?<![A-Z0-9])(?:\d{2}[ \t]*AR[ \t]*\d{6}[ \t]*[A-Z]|\d{2}[ \t]*\d{3}[ \t]*[A-Z]{3,5}[ \t]*\d{4,8}[A-Z0-9]?)(?![A-Z0-9])",
        normalizado,
    )
    return re.sub(r"\s+", "", match.group()) if match else ""


def _crts_campo(contenido):
    """Recoge los identificadores de un campo CRT, incluso en varias líneas."""
    encontrados = []
    for linea in contenido.splitlines():
        linea = linea.strip(" \t:/–")
        linea = re.sub(r"^(?:NROS?\.?|NRS?\.?|NUMEROS?|N[°º?O.]?)\s*[:.-]?\s*", "", linea, flags=re.I)
        if not linea:
            continue
        # Una barra entre dos códigos AR completos separa CRT, mientras que
        # una barra interna de otro identificador conserva su significado.
        linea = re.sub(r"(?<=\d)/(?=(?:038[.-]?)?AR\d)", ";", linea, flags=re.I)
        tokens = re.split(r"\s*[,;|]\s*|\s+/\s+|\s+(?:Y|E|AND)\s+|\s+", linea, flags=re.I)
        numeros = []
        for token in tokens:
            token = token.strip().upper()
            if (not re.fullmatch(r"[A-Z0-9]+(?:[./-][A-Z0-9]+)*", token)
                    or not re.search(r"\d", token)
                    or len(re.sub(r"[^A-Z0-9]", "", token)) < 4):
                # Detenerse ante descripciones o notas: sus números no son CRT.
                return encontrados
            token = re.sub(r"^038[.-]?(?=AR\d)", "", token)
            numeros.append(token)
        for numero in numeros:
            if numero not in encontrados:
                encontrados.append(numero)
    return encontrados


def procesar_manifiesto(texto: str, nombre_archivo: str = "") -> dict:
    """Devuelve las 17 columnas; un dato no encontrado permanece vacío."""
    datos = dict.fromkeys(_COLUMNAS_MANIFIESTO, "")
    texto = texto or ""
    datos["MIC ELEC."] = _mic_manifiesto(texto) or _mic_manifiesto(nombre_archivo or "")
    if not texto.strip():
        return datos
    secciones = _secciones_manifiesto(texto)

    for tipo, columna in (("exportador", "EXPORTADOR"), ("importador", "IMPORTADOR"), ("chofer", "CHOFER")):
        for _, contenido in secciones.get(tipo, []):
            valor = _primera_linea_campo(contenido)
            if valor:
                datos[columna] = valor
                break
    for tipo, columna in (("origen", "ORIGEN"), ("destino", "ADUANA DESTINO"), ("salida", "ADUANA DE SALIDA")):
        for _, contenido in secciones.get(tipo, []):
            valor = _localidad_campo(contenido, quitar_pais=(tipo == "origen"))
            if valor:
                datos[columna] = valor
                break
    if not datos["ADUANA DE SALIDA"]:
        for _, contenido in secciones.get("ruta", []):
            match = re.search(r"\b(PASO DE LOS LIBRES|IGUAZU|CRISTO REDENTOR|SAN JAVIER|SANTO TOME|GUALEGUAYCHU|CLORINDA|POCITOS|LA QUIACA|PTM)\b", _normalizar_busqueda(contenido))
            if match:
                datos["ADUANA DE SALIDA"] = match.group(1)
                break

    for _, contenido in secciones.get("fecha", []):
        match = re.search(r"(?<!\d)(\d{1,2}[/.\-]\d{1,2}[/.\-](?:\d{4}|\d{2}))(?!\d)", _primera_linea_campo(contenido))
        if match:
            partes = re.split(r"[/.\-]", match.group(1))
            try:
                fecha = datetime.strptime("/".join(partes), "%d/%m/%Y" if len(partes[2]) == 4 else "%d/%m/%y")
            except ValueError:
                continue
            datos["fecha"] = fecha.strftime("%d/%m/%Y")
            break

    crts = []
    for _, contenido in secciones.get("crt", []):
        for numero in _crts_campo(contenido):
            if numero not in crts:
                crts.append(numero)
    datos["CRT"] = "; ".join(crts)
    for _, contenido in secciones.get("factura", []):
        valor = _identificador_campo(contenido)
        if valor:
            datos["FACTURA"] = valor
            break
    for tipo, columna, moneda in (("valor", "VALOR", "USD"), ("flete", "FRETE", "USD"), ("reales", "FLETE EN REALES", "BRL"), ("seguro", "SEGURO", "USD")):
        for rotulo, contenido in secciones.get(tipo, []):
            valor = _importe_campo(rotulo, contenido, moneda)
            if valor is not None:
                datos[columna] = format(valor, ".2f") if columna == "SEGURO" else _formatear_importe(valor, moneda)
                break
    if not datos["FLETE EN REALES"]:
        for rotulo, contenido in secciones.get("flete", []):
            if re.search(r"BRL|R\$|REALES|REAIS", _normalizar_busqueda(contenido)):
                valor = _importe_campo(rotulo, contenido, "BRL")
                if valor is not None:
                    datos["FLETE EN REALES"] = _formatear_importe(valor, "BRL")
                    break

    patente = r"(?<![A-Z0-9])(?:[A-Z]{2}[ \t]*\d{3}[ \t]*[A-Z]{2}|[A-Z]{3}[ \t]*\d[A-Z0-9]\d{2}|[A-Z]{3}[ \t]*\d{3,4})(?![A-Z0-9])"
    for tipo, columna in (("tractor", "TRACTOR"), ("carreta", "CARRETA")):
        for _, contenido in secciones.get(tipo, []):
            match = re.search(patente, _normalizar_busqueda(contenido))
            if match:
                datos[columna] = re.sub(r"\s+", "", match.group())
                break
    for _, contenido in secciones.get("dni", []):
        match = re.match(r"\s*[:.-]?\s*(?:CI\s*)?([0-9][0-9.\-]{4,18}[0-9])(?!\d)", contenido, re.I)
        if match:
            datos["DNI"] = match.group(1)
            break
    return datos

# ==============================================================================
# 3. CONECTOR DE GOOGLE SHEETS (SECRETS & LOCAL)
# ==============================================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def conectar_google_sheets(sheet_url: str):
    """Inicializa la conexión con Google Sheets tanto en Streamlit Cloud como en Local."""
    creds = None
    
    # 1. Intentar cargar desde Secrets de Streamlit Cloud
    try:
        if "gcp_service_account" in st.secrets:
            creds_info = dict(st.secrets["gcp_service_account"])
            if "private_key" in creds_info:
                creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
            creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    except Exception:
        pass
    
    # 2. Si no está en secrets, intentar cargar credentials.json local
    if creds is None:
        ruta = obtener_ruta_credenciales()
        if ruta:
            creds = Credentials.from_service_account_file(ruta, scopes=SCOPES)
    
    if creds is None:
        raise Exception("No se encontraron credenciales válidas en Streamlit Secrets ni en credentials.json.")
        
    client = gspread.authorize(creds)
    if sheet_url.startswith("https://"):
        return client.open_by_url(sheet_url)
    return client.open(sheet_url)

def guardar_en_google_sheets(df: pd.DataFrame, sheet_target: str, worksheet_name: str = "Septiembre"):
    """Agrega cargas y completa los CRT de un MIC existente sin pisar otros datos."""
    spreadsheet = conectar_google_sheets(sheet_target)
    try:
        ws = spreadsheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        raise ValueError(f"No existe la pestaña '{worksheet_name}'. Revisá su nombre antes de guardar.") from None

    valores_existentes = ws.get_all_values()
    columnas = list(df.columns)
    if not any(str(celda).strip() for fila in valores_existentes for celda in fila):
        ws.update(values=[columnas], range_name="A1")
        valores_existentes = [columnas]
    encabezados = [str(celda).strip() for celda in valores_existentes[0]]
    while encabezados and not encabezados[-1]:
        encabezados.pop()
    if len(encabezados) != len(columnas) or set(encabezados) != set(columnas):
        raise ValueError("Los encabezados de la planilla no coinciden con las columnas de la app. Revisalos antes de guardar.")

    idx_mic, idx_crt = encabezados.index("MIC ELEC."), encabezados.index("CRT")
    existentes = {}
    for numero_fila, valores in enumerate(valores_existentes[1:], start=2):
        valores = valores + [""] * max(0, len(encabezados) - len(valores))
        mic = normalizar_mic(valores[idx_mic])
        if mic:
            existentes.setdefault(mic, []).append((numero_fila, valores))

    nuevas_filas, nuevas_por_mic, cambios_crt = [], {}, {}
    for _, row in df.iterrows():
        valores = ["" if pd.isna(row.get(col, "")) else str(row.get(col, "")) for col in encabezados]
        if not any(valor.strip() for valor in valores):
            continue
        mic = normalizar_mic(valores[idx_mic])
        valores[idx_mic] = mic
        valores[idx_crt] = combinar_crts(valores[idx_crt])
        if mic and mic in existentes:
            if len(existentes[mic]) != 1:
                raise ValueError(f"El MIC {mic} aparece en varias filas de Google Sheets. Revisá ese duplicado antes de completar sus CRT.")
            numero_fila, anterior = existentes[mic][0]
            combinado = combinar_crts(cambios_crt.get(numero_fila, anterior[idx_crt]), valores[idx_crt])
            if combinado != combinar_crts(anterior[idx_crt]):
                cambios_crt[numero_fila] = combinado
            continue
        if mic and mic in nuevas_por_mic:
            anterior = nuevas_filas[nuevas_por_mic[mic]]
            anterior[idx_crt] = combinar_crts(anterior[idx_crt], valores[idx_crt])
            continue
        if mic:
            nuevas_por_mic[mic] = len(nuevas_filas)
        nuevas_filas.append(valores)

    if cambios_crt:
        ws.batch_update([
            {"range": gspread.utils.rowcol_to_a1(numero_fila, idx_crt + 1), "values": [[crt]]}
            for numero_fila, crt in cambios_crt.items()
        ], value_input_option="RAW")
    if nuevas_filas:
        ws.append_rows(nuevas_filas, value_input_option="RAW")
    return len(nuevas_filas) + len(cambios_crt)

# ==============================================================================
# 4. INTERFAZ WEB STREAMLIT
# ==============================================================================
def normalizar_mic(valor):
    if valor is None or pd.isna(valor):
        return ""
    return re.sub(r"\s+", "", str(valor)).upper()


def combinar_crts(*valores):
    numeros = []
    for valor in valores:
        if valor is None or pd.isna(valor):
            continue
        for numero in re.split(r"[;,\n]+|\s+/\s+", str(valor)):
            numero = re.sub(r"^038[.-]?(?=AR\d)", "", numero.strip().upper())
            if numero and numero not in numeros:
                numeros.append(numero)
    return "; ".join(numeros)


def cargar_archivos(archivos, registros, procesados):
    """Procesa cada contenido una sola vez y evita MIC repetidos en el lote."""
    mics = {normalizar_mic(r.get("MIC ELEC.")) for r in registros}
    nuevos, avisos = 0, []
    for archivo in archivos:
        contenido = archivo.getvalue()
        identidad = hashlib.sha256(contenido).hexdigest()
        if identidad in procesados:
            continue
        procesados.add(identidad)
        try:
            texto = extraer_texto_pdf(io.BytesIO(contenido))
        except Exception:
            avisos.append(f"{archivo.name}: no se pudo leer el PDF. Comprobá que no esté dañado o protegido con contraseña.")
            continue
        if not texto:
            avisos.append(f"{archivo.name}: el PDF no tiene texto legible. Si es un escaneo, necesita reconocimiento de texto (OCR). No se agregó una fila vacía.")
            continue
        datos = procesar_manifiesto(texto, archivo.name)
        if not any(datos.values()):
            avisos.append(f"{archivo.name}: no se reconocieron los campos del manifiesto. No se agregó una fila vacía.")
            continue
        mic = normalizar_mic(datos.get("MIC ELEC."))
        if mic and mic in mics:
            existente = next(r for r in registros if normalizar_mic(r.get("MIC ELEC.")) == mic)
            anteriores = combinar_crts(existente.get("CRT"))
            combinados = combinar_crts(anteriores, datos.get("CRT"))
            if combinados != anteriores:
                existente["CRT"] = combinados
                avisos.append(f"{archivo.name}: se incorporaron los CRT adicionales al MIC {mic} en la misma fila.")
                continue
            avisos.append(f"{archivo.name}: el MIC {mic} ya está cargado. No se agregó otra fila.")
            continue
        registros.append(datos)
        if mic:
            mics.add(mic)
        nuevos += 1
        campos_revision = ("MIC ELEC.", "fecha", "EXPORTADOR", "IMPORTADOR", "VALOR", "FRETE", "TRACTOR", "CARRETA", "CHOFER", "DNI")
        faltantes = [campo for campo in campos_revision if not datos.get(campo)]
        if faltantes:
            avisos.append(f"{archivo.name}: revisá los campos sin detectar: {', '.join(faltantes)}.")
    return nuevos, avisos


def aplicar_cambios_editor(registros, cambios):
    """Aplica una edición a una copia para no repetir altas/bajas al recargar."""
    columnas = tuple(procesar_manifiesto(""))
    filas = [dict(fila) for fila in registros]
    for indice, valores in cambios.get("edited_rows", {}).items():
        indice = int(indice)
        if 0 <= indice < len(filas):
            filas[indice].update({col: "" if valor is None else valor for col, valor in valores.items() if col in columnas})
    borradas = {int(indice) for indice in cambios.get("deleted_rows", [])}
    filas = [fila for indice, fila in enumerate(filas) if indice not in borradas]
    for fila in cambios.get("added_rows", []):
        filas.append({col: "" if fila.get(col) is None else fila[col] for col in columnas})
    return filas


def confirmar_edicion(clave_editor):
    cambios = st.session_state.get(clave_editor, {})
    st.session_state.registros = aplicar_cambios_editor(st.session_state.registros, cambios)
    st.session_state.revision_cargas += 1


st.title("🚚 Registro de Cargas y Control de Camiones")
st.markdown("""
Sube tus **Manifiestos de Carga (MIC/DTA, CRT)** en PDF. La app identifica sus datos y los prepara con las columnas de tu planilla. Revisalos antes de guardarlos en tu **Google Sheet**.
""")

if "registros" not in st.session_state:
    st.session_state.registros = []
for clave, valor in (("archivos_procesados", set()), ("avisos_carga", []), ("revision_cargas", 0), ("cargador_version", 0)):
    if clave not in st.session_state:
        st.session_state[clave] = valor

creds_disponibles = verificar_credenciales_disponibles()

# Barra lateral: Configuración de Google Sheets
with st.sidebar:
    st.header("🌐 Configuración Google Sheets")
    sheet_url = st.text_input(
        "URL de tu Google Sheet:",
        value="https://docs.google.com/spreadsheets/d/1-9AkVFnZkx1miHjsh5USFifcfFp-o6-1mMtJ94KfyQ8/edit?usp=sharing",
        help="Enlace configurado a tu planilla de Google Sheets."
    )
    nombre_pestana = st.text_input("Nombre de la Pestaña:", value="Septiembre")
    
    if creds_disponibles:
        st.success("✅ Credenciales de Google activas.")
    else:
        st.warning("⚠️ No se encontraron credenciales de Google.")
        st.caption("Configura Secrets en Streamlit Cloud o coloca credentials.json en local.")

    st.divider()
    if st.button("🗑️ Limpiar registros", use_container_width=True):
        st.session_state.registros = []
        st.session_state.archivos_procesados = set()
        st.session_state.avisos_carga = []
        st.session_state.revision_cargas += 1
        st.session_state.cargador_version += 1
        st.rerun()

# 1. ZONA DRAG & DROP
st.subheader("📄 1. Soltar Manifiestos de Carga (PDF)")
archivos = st.file_uploader(
    "Arrastra tus archivos PDF aquí (puedes subir varios a la vez):",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"manifiestos_{st.session_state.cargador_version}",
    help="Arrastra tus PDFs de MIC/DTA o Manifiestos de carga."
)

if archivos:
    registros_antes = [dict(fila) for fila in st.session_state.registros]
    nuevos, avisos = cargar_archivos(archivos, st.session_state.registros, st.session_state.archivos_procesados)
    st.session_state.avisos_carga.extend(avisos)
    if st.session_state.registros != registros_antes:
        st.session_state.revision_cargas += 1
    if nuevos > 0:
        st.success(f"✅ Se procesaron {nuevos} nuevo(s) manifiesto(s).")

if st.session_state.avisos_carga:
    with st.expander("Revisar los archivos cargados", expanded=True):
        for aviso in st.session_state.avisos_carga:
            st.warning(aviso)

# 2. PLANILLA INTERACTIVA
st.subheader("📊 2. Planilla de Cargas del Mes")

if st.session_state.registros:
    df_actual = pd.DataFrame(st.session_state.registros)
    
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Total Camiones", len(df_actual))
    
    st.caption("✏️ Puedes hacer doble clic en cualquier celda para corregir o agregar información antes de guardar.")
    st.caption("Los datos que no se pudieron identificar quedan en blanco; completalos antes de guardar.")
    st.caption("Si un manifiesto tiene varios CRT, aparecen juntos en la columna CRT, separados por punto y coma.")
    clave_editor = f"editor_cargas_{st.session_state.revision_cargas}"
    st.data_editor(
        df_actual,
        use_container_width=True,
        num_rows="dynamic",
        key=clave_editor,
        on_change=confirmar_edicion,
        args=(clave_editor,),
    )

    # 3. GUARDADO EN GOOGLE SHEETS
    st.divider()
    st.subheader("💾 3. Sincronizar en Google Sheets & Copia Local")
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.markdown("#### ☁️ Google Sheets")
        st.caption("Agrega cargas nuevas y completa los CRT de un MIC ya guardado. Conserva los demás datos de las filas existentes.")
        if st.button("📤 Guardar / Sincronizar en Google Sheets", type="primary", use_container_width=True):
            if not sheet_url:
                st.error("Por favor, ingresa el enlace de tu Google Sheet en la barra lateral.")
            elif not creds_disponibles:
                st.error("Faltan las credenciales para autenticar con Google Sheets.")
            else:
                try:
                    with st.spinner("Sincronizando con tu Google Sheet..."):
                        filas_guardadas = guardar_en_google_sheets(
                            pd.DataFrame(st.session_state.registros),
                            sheet_url,
                            nombre_pestana
                        )
                        st.success(f"🎉 ¡Éxito! Se sincronizaron {filas_guardadas} fila(s) en tu Google Sheet.")
                        st.markdown(f"👉 [Abrir Google Sheet en el navegador]({sheet_url})")
                except Exception as err:
                    st.error(f"Error al conectar con Google Sheets: {err}")
    
    with col_g2:
        st.markdown("#### 📥 Copia de Respaldo (Excel)")
        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
            pd.DataFrame(st.session_state.registros).to_excel(writer, index=False, sheet_name=nombre_pestana)
        
        st.download_button(
            label="Descargar Planilla Excel (.xlsx)",
            data=out_excel.getvalue(),
            file_name=f"Registro_Cargas_{datetime.now().strftime('%Y_%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
else:
    st.info("ℹ️ Arrastra o selecciona tus archivos PDF de manifiestos arriba para comenzar.")
