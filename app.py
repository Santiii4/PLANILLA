import io
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
# CONFIGURACIÓN DE RUTAS
# ==============================================================================
DIRECTORIO_APP = os.path.dirname(os.path.abspath(__file__))

def obtener_ruta_credenciales():
    """Busca el archivo credentials.json exactamente donde está app.py."""
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
    """Extrae el contenido textual del PDF subido."""
    try:
        lector = pypdf.PdfReader(archivo_pdf)
        texto = "\n".join([pagina.extract_text() or "" for pagina in lector.pages])
        return texto
    except Exception as e:
        st.error(f"Error al leer el archivo PDF: {e}")
        return ""

def limpiar_campo(txt: str) -> str:
    """Limpia encabezados y títulos de campos aduaneros."""
    if not txt:
        return ""
    txt = re.sub(
        r'^(?:(?:\d{1,2}\.?\s*)?(?:Remitente|Remetente|Destinatario|Destinat[aá]rio|Consignatario|Aduana|Ciudad|Cidade|Valor|Flete|Frete|Fecha|Data|Moneda|Moeda|Nombre|Nome)[^\n\r\:]*[:\/\-\.]?)\s*', 
        '', txt, flags=re.IGNORECASE
    )
    txt = re.sub(r'^[\s\:\/\-\#\.]+', '', txt)
    return re.sub(r'\s+', ' ', txt).strip()

def procesar_manifiesto(texto: str, nombre_archivo: str = "") -> dict:
    """Extrae y estructura los datos con las 17 columnas de tu tabla."""
    datos = {
        "ORIGEN": "",
        "ADUANA DESTINO": "",
        "ADUANA DE SALIDA": "",
        "EXPORTADOR": "",
        "IMPORTADOR": "",
        "fecha": "",
        "MIC ELEC.": "",
        "CRT": "",
        "FACTURA": "",
        "VALOR": "",
        "FLETE EN REALES": "",
        "FRETE": "",
        "TRACTOR": "",
        "CARRETA": "",
        "CHOFER": "",
        "DNI": "",
        "SEGURO": ""
    }

    if not texto:
        m_fn = re.search(r'(\d{2}AR\d{6}[A-Z]|\d{2}\d{3}[A-Z]{3,5}\d{4,8}[A-Z0-9]?)', nombre_archivo, re.IGNORECASE)
        if m_fn:
            datos["MIC ELEC."] = m_fn.group(1).upper()
        return datos

    # 1. ORIGEN (Campo 7 o 26)
    m_orig = re.search(r'(?:7\s*Aduana[^\n\r]*partida[\s\S]*?)([A-Z\s]{4,20})-(?:ARGENTINA|BRASIL|CHILE|URUGUAY)', texto, re.IGNORECASE)
    if m_orig:
        lineas_o = [l.strip() for l in m_orig.group(1).split('\n') if len(l.strip()) > 3]
        if lineas_o:
            datos["ORIGEN"] = lineas_o[-1]
    if not datos["ORIGEN"]:
        datos["ORIGEN"] = "MENDOZA"

    # 2. ADUANA DESTINO (Campo 24 o Campo 8)
    m_ad_dest = re.search(r'(?:24\s*Aduana\s*de\s*destino[^\n\r]*[\n\r]+)([^\n\r]{3,60})', texto, re.IGNORECASE)
    if m_ad_dest:
        dest_val = m_ad_dest.group(1).strip()
        dest_val = re.sub(r'^\d+\s*', '', dest_val).rstrip('-').strip()
        datos["ADUANA DESTINO"] = dest_val
    else:
        m_ad_dest8 = re.search(r'(?:8\s*Ciudad\s*y\s*pais\s*de\s*destino\s*final[\s\S]*?)([A-Z\s\-]{3,30}-[A-Z\s]{3,30})', texto, re.IGNORECASE)
        if m_ad_dest8:
            datos["ADUANA DESTINO"] = m_ad_dest8.group(1).strip()

    # 3. ADUANA DE SALIDA (Frontera en ruta Campo 40)
    m_salida = re.search(r'(PASO DE LOS LIBRES|IGUAZU|CRISTO REDENTOR|SAN JAVIER|SANTO TOME|GUALEGUAYCHU|CLORINDA|POCITOS|LA QUIACA|PTM)', texto, re.IGNORECASE)
    if m_salida:
        datos["ADUANA DE SALIDA"] = m_salida.group(1).strip().upper()
    else:
        datos["ADUANA DE SALIDA"] = datos["ORIGEN"]

    # 4. EXPORTADOR (Campo 33 o Campo 6 / CRT 1)
    m_exp_33 = re.search(r'(?:33\s*Remitente[^\n\r]*[\n\r]+)([\s\S]*?)(?=(?:34\s*Destinatario|34\.|\Z))', texto, re.IGNORECASE)
    if m_exp_33:
        lineas = [l.strip() for l in m_exp_33.group(1).strip().split('\n') if l.strip()]
        if lineas:
            datos["EXPORTADOR"] = lineas[0].replace("?", "Ñ").replace("VI EDOS", "VIÑEDOS").replace("VI?EDOS", "VIÑEDOS")
    else:
        m_exp_gen = re.search(r'(?:Remitente|Exportador)[^\n\r\:]*[:\/\-\.]?\s*\n?([^\n\r;]{3,70})', texto, re.IGNORECASE)
        if m_exp_gen:
            datos["EXPORTADOR"] = limpiar_campo(m_exp_gen.group(1)).replace("?", "Ñ")

    # 5. IMPORTADOR (Campo 34 o Campo 7 / CRT 4)
    m_imp_34 = re.search(r'(?:34\s*Destinatario[^\n\r]*[\n\r]+)([\s\S]*?)(?=(?:35\s*Consignatario|35\.|\Z))', texto, re.IGNORECASE)
    if m_imp_34:
        lineas = [l.strip() for l in m_imp_34.group(1).strip().split('\n') if l.strip()]
        if lineas:
            datos["IMPORTADOR"] = lineas[0]
    else:
        m_imp_gen = re.search(r'(?:Destinatario|Destinat[aá]rio|Importador)[^\n\r\:]*[:\/\-\.]?\s*\n?([^\n\r;]{3,70})', texto, re.IGNORECASE)
        if m_imp_gen:
            datos["IMPORTADOR"] = limpiar_campo(m_imp_gen.group(1))

    # 6. FECHA (Formato D/M/AAAA)
    m_fec = re.search(r'(?:F\.?\s*Ofic|Fecha(?:\s*Emisi[oó]n)?|Data(?:\s*de\s*emiss[aã]o)?)\s*[:\/\-\.]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.](?:20)?\d{2,4})', texto, re.IGNORECASE)
    if m_fec:
        datos["fecha"] = m_fec.group(1).strip()
    else:
        m_fecha_gen = re.search(r'\b(\d{2}[\/\-\.]\d{2}[\/\-\.]\d{4})\b', texto)
        if m_fecha_gen:
            datos["fecha"] = m_fecha_gen.group(1).strip()

    # 7. MIC ELEC. (26AR348605N / 26AR349197U)
    m_mic = re.search(r'\b(\d{2}AR\d{6}[A-Z]|\d{2}\s*\d{3}\s*[A-Z]{3,5}\s*\d{4,8}\s*[A-Z0-9]?)\b', texto, re.IGNORECASE)
    if m_mic:
        datos["MIC ELEC."] = m_mic.group(1).replace(" ", "").upper()
    elif nombre_archivo:
        m_fn = re.search(r'(\d{2}AR\d{6}[A-Z])', nombre_archivo, re.IGNORECASE)
        if m_fn:
            datos["MIC ELEC."] = m_fn.group(1).upper()

    # 8. CRT (Eliminando prefijo '038' o '038.')
    m_crt = re.search(r'(?:23\s*N[°\?ºo\.]?\s*carta\s*de\s*porte[^\n\r]*[\n\r]+|2\s*Numero[^\n\r]*[\n\r]+)([0-9A-Z\.\-]{8,25})', texto, re.IGNORECASE)
    if m_crt:
        crt_raw = m_crt.group(1).strip()
        crt_limpio = re.sub(r'^038[\.\-]?', '', crt_raw)
        datos["CRT"] = crt_limpio

    # 9. FACTURA (De la descripción campo 38 o CRT)
    m_fac = re.search(r'(?:FATURA|FACTURA\s*COMERCIAL|FACTURA|INVOICE)\s*(?:N[°ºo\.]?|NR\.?|NRO\.?|:)?\s*([E0-9A-Z\-\/]{4,25})', texto, re.IGNORECASE)
    if m_fac:
        val_f = m_fac.group(1).strip()
        if val_f.upper() not in ['COMERCIAL', 'NR', 'NRO', 'NUMERO']:
            datos["FACTURA"] = val_f
    if not datos["FACTURA"]:
        m_fac_alt = re.search(r'(?:FACTURA\s*COMERCIAL\s*NR\.?\s*)([0-9A-Z\-\/]{4,25})', texto, re.IGNORECASE)
        if m_fac_alt:
            datos["FACTURA"] = m_fac_alt.group(1).strip()

    # 10. VALOR (FOB)
    m_val = re.search(r'(?:27\s*Valor\s*FO[BT]|Valor\s*FO[BT]|15\s*Moneda\s*y\s*valor\s*FO[BT]|14\s*Valor)[\s\S]*?(\d{2,7}[\.\,]\d{2})', texto, re.IGNORECASE)
    if m_val:
        v_num = m_val.group(1).replace(",", ".")
        datos["VALOR"] = f"USD {float(v_num):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")

    # 11. FLETE EN REALES
    m_reales = re.search(r'(?:BRL|R\$)\s*([\d\.\,]{3,15})', texto, re.IGNORECASE)
    if m_reales:
        datos["FLETE EN REALES"] = f"BRL {m_reales.group(1)}"

    # 12. FRETE (FLETE EN USD)
    m_frete = re.search(r'(?:28\s*Flete|Flete\s*en\s*US[S\$]|Frete\s*em\s*US[S\$]|Flete\s*/\s*Frete|15\s*Gastos\s*a\s*pagar[\s\S]*?Flete)[\s\S]*?(\d{2,7}[\.\,]\d{2})', texto, re.IGNORECASE)
    if m_frete:
        fr_num = m_frete.group(1).replace(",", ".")
        datos["FRETE"] = f"USD {float(fr_num):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")

    # 13. TRACTOR (Placa)
    m_trac = re.search(r'(?:11\s*Placa\s*de\s*Camion|18\s*Placa\s*de\s*Camion)[\s\S]*?([A-Z]{3}[0-9][A-Z0-9][0-9]{2}|[A-Z]{3}[0-9]{3,4}|[A-Z]{2}[0-9]{3}[A-Z]{2})', texto, re.IGNORECASE)
    if m_trac:
        datos["TRACTOR"] = m_trac.group(1).strip()

    # 14. CARRETA (Semirremolque)
    m_carr = re.search(r'(?:15[\s\S]*?Semiremolque|Semi-reboque)[\s\S]*?([A-Z]{3}[0-9][A-Z0-9][0-9]{2}|[A-Z]{3}[0-9]{3,4}|[A-Z]{2}[0-9]{3}[A-Z]{2})', texto, re.IGNORECASE)
    if m_carr:
        datos["CARRETA"] = m_carr.group(1).strip()

    # 15. CHOFER (Conductor)
    m_chof = re.search(r'(?:CONDUCTOR\s*1\s*:\s*|CHOFER\s*:\s*)([A-Z\s]{4,40})(?=\s*DOC|\n|\r|\Z)', texto, re.IGNORECASE)
    if m_chof:
        datos["CHOFER"] = m_chof.group(1).strip()

    # 16. DNI (Documento del chofer)
    m_dni = re.search(r'(?:DOC\s*:\s*|DNI\s*:\s*)(?:CI\s*)?([0-9\.\-]{6,20})', texto, re.IGNORECASE)
    if m_dni:
        datos["DNI"] = m_dni.group(1).strip()

    # 17. SEGURO (Debajo del Flete en Campo 29 o en tabla de gastos)
    m_seg29 = re.search(r'(?:29\s*Seguro\s*en\s*US[S\$]|29\s*Seguro)[^\n\r]*[\n\r]+(?:Seguro[^\n\r]*[\n\r]+)?\s*([\d\.\,]{1,8})', texto, re.IGNORECASE)
    if m_seg29:
        datos["SEGURO"] = m_seg29.group(1).replace(",", ".")
    else:
        m_seg_crt = re.search(r'(?:Seguro\s*/\s*Seguro|SEGURO\s*X\s*SEGURO|SEGURO\s*USD)\s*[:\/\-\.]?\s*(?:USD|US\$|\$)?\s*([\d\.\,]{1,8})', texto, re.IGNORECASE)
        if m_seg_crt:
            datos["SEGURO"] = m_seg_crt.group(1).replace(",", ".")
        else:
            datos["SEGURO"] = "0.00"

    return datos

# ==============================================================================
# 3. CONECTOR DE GOOGLE SHEETS
# ==============================================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def conectar_google_sheets(creds_path, sheet_url: str):
    """Inicializa la conexión con Google Sheets usando la cuenta de servicio."""
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    client = gspread.authorize(creds)
    if sheet_url.startswith("https://"):
        return client.open_by_url(sheet_url)
    return client.open(sheet_url)

def guardar_en_google_sheets(df: pd.DataFrame, creds_path, sheet_target: str, worksheet_name: str = "Hoja 1"):
    """Guarda o añade los registros a la hoja de cálculo asegurando encabezados en la Fila 1."""
    spreadsheet = conectar_google_sheets(creds_path, sheet_target)
    
    try:
        ws = spreadsheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        try:
            ws = spreadsheet.sheet1
        except Exception:
            ws = spreadsheet.add_worksheet(title=worksheet_name, rows="1000", cols="25")
    
    valores_existentes = ws.get_all_values()
    columnas = list(df.columns)
    
    # 1. VERIFICAR SI LA FILA 1 CONTIENE NUESTROS ENCABEZADOS
    tiene_encabezado = False
    if valores_existentes and len(valores_existentes) > 0:
        fila_1 = [str(x).strip() for x in valores_existentes[0] if str(x).strip()]
        if any(h in fila_1 for h in ["MIC ELEC.", "EXPORTADOR", "ORIGEN", "VALOR"]):
            tiene_encabezado = True

    hoja_totalmente_vacia = True
    if valores_existentes:
        for fila in valores_existentes:
            if any(str(x).strip() for x in fila):
                hoja_totalmente_vacia = False
                break

    # Si no tiene los encabezados, los escribimos en la Fila 1
    if not tiene_encabezado:
        if hoja_totalmente_vacia:
            ws.update(values=[columnas], range_name="A1")
        else:
            ws.insert_row(columnas, index=1)
        
        try:
            ws.format("A1:Q1", {
                "backgroundColor": {"red": 0.35, "green": 0.20, "blue": 0.08},
                "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
                "horizontalAlignment": "CENTER"
            })
        except Exception:
            pass

    # 2. EVITAR DUPLICADOS POR MIC ELEC.
    valores_actualizados = ws.get_all_values()
    mic_existentes = set()
    if len(valores_actualizados) > 1:
        idx_mic = valores_actualizados[0].index("MIC ELEC.") if "MIC ELEC." in valores_actualizados[0] else 6
        mic_existentes = {fila[idx_mic] for fila in valores_actualizados[1:] if len(fila) > idx_mic}
    
    nuevas_filas = []
    for _, row in df.iterrows():
        nro_mic = str(row.get("MIC ELEC.", ""))
        if nro_mic and nro_mic in mic_existentes:
            continue
        nuevas_filas.append([str(row.get(col, "")) for col in columnas])
    
    if nuevas_filas:
        ws.append_rows(nuevas_filas)
        return len(nuevas_filas)
    return 0

# ==============================================================================
# 4. INTERFAZ WEB STREAMLIT
# ==============================================================================
st.title("🚚 Registro de Cargas y Control de Camiones")
st.markdown("""
Sube tus **Manifiestos de Carga (MIC/DTA, CRT)** en PDF. La app extrae automáticamente los datos según la estructura de tu planilla y los sincroniza con tu **Google Sheet**.
""")

if "registros" not in st.session_state:
    st.session_state.registros = []

# Detectar credenciales automáticamente
ruta_credenciales = obtener_ruta_credenciales()
creds_disponibles = ruta_credenciales is not None

# Barra lateral: Configuración de Google Sheets
with st.sidebar:
    st.header("🌐 Configuración Google Sheets")
    sheet_url = st.text_input(
        "URL de tu Google Sheet:",
        value="https://docs.google.com/spreadsheets/d/1-9AkVFnZkx1miHjsh5USFifcfFp-o6-1mMtJ94KfyQ8/edit?usp=sharing",
        help="Enlace configurado a tu planilla de Google Sheets."
    )
    nombre_pestana = st.text_input("Nombre de la Pestaña:", value="Hoja 1")
    
    if creds_disponibles:
        st.success("✅ Archivo `credentials.json` listo.")
    else:
        st.warning("⚠️ No se encontró `credentials.json` en la carpeta.")
        st.caption("Verifica que el archivo esté en la misma carpeta que `app.py`.")

    st.divider()
    if st.button("🗑️ Limpiar registros", use_container_width=True):
        st.session_state.registros = []
        st.rerun()

# 1. ZONA DRAG & DROP
st.subheader("📄 1. Soltar Manifiestos de Carga (PDF)")
archivos = st.file_uploader(
    "Arrastra tus archivos PDF aquí (puedes subir varios a la vez):",
    type=["pdf"],
    accept_multiple_files=True,
    help="Arrastra tus PDFs de MIC/DTA o Manifiestos de carga."
)

if archivos:
    nuevos = 0
    mics_ya_cargados = [r.get("MIC ELEC.") for r in st.session_state.registros if r.get("MIC ELEC.")]
    
    for arc in archivos:
        texto = extraer_texto_pdf(arc)
        datos = procesar_manifiesto(texto, arc.name)
        if datos["MIC ELEC."] not in mics_ya_cargados or not datos["MIC ELEC."]:
            st.session_state.registros.append(datos)
            nuevos += 1
            
    if nuevos > 0:
        st.success(f"✅ Se procesaron {nuevos} nuevo(s) manifiesto(s).")

# 2. PLANILLA INTERACTIVA
st.subheader("📊 2. Planilla de Cargas del Mes")

if st.session_state.registros:
    df_actual = pd.DataFrame(st.session_state.registros)
    
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Total Camiones", len(df_actual))
    
    st.caption("✏️ Puedes hacer doble clic en cualquier celda para corregir o agregar información antes de guardar.")
    df_editado = st.data_editor(
        df_actual,
        use_container_width=True,
        num_rows="dynamic",
        key="editor_cargas"
    )
    st.session_state.registros = df_editado.to_dict('records')

    # 3. GUARDADO EN GOOGLE SHEETS
    st.divider()
    st.subheader("💾 3. Sincronizar en Google Sheets & Copia Local")
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.markdown("#### ☁️ Google Sheets")
        if st.button("📤 Guardar / Sincronizar en Google Sheets", type="primary", use_container_width=True):
            if not sheet_url:
                st.error("Por favor, ingresa el enlace de tu Google Sheet en la barra lateral.")
            elif not creds_disponibles:
                st.error("Falta el archivo `credentials.json` para autenticar con Google Sheets.")
            else:
                try:
                    with st.spinner("Sincronizando con tu Google Sheet..."):
                        filas_guardadas = guardar_en_google_sheets(
                            pd.DataFrame(st.session_state.registros),
                            ruta_credenciales,
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