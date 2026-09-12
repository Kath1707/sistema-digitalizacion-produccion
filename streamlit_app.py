"""
App de Registro de Parámetros de Calidad - Producto Terminado B2B Starbucks (v2)
==================================================================================
Streamlit + Google Sheets (histórico) + GitHub deploy.

Estructura esperada del repo:
  streamlit_app.py
  requirements.txt
  data/MA_BASE_DATOS_PRODUCTOS_TERMINADO.xlsx

Secrets necesarios en Streamlit Cloud (Manage app -> Settings -> Secrets):

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "...@...iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."

ROOT_FOLDER_ID = "1h46IR2nUS8TZUSIgl1lD-M4nCtp-p_wW"
"""

import io
import re
from datetime import date, timedelta

import pandas as pd
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build as build_drive_service
    GSHEETS_DISPONIBLE = True
except ImportError:
    GSHEETS_DISPONIBLE = False

# ----------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Registro de Calidad - Producto Terminado B2B Starbucks",
    page_icon="✅",
    layout="wide",
)

EXCEL_PATH = "MA_BASE_DATOS_PRODUCTOS_TERMINADO.xlsx"
SHEET_NAME = "ALMENARA-PT_2026"
CLIENTE_FIJO = "STARBUCKS"
AREA_FIJA = "EMPAQUE"

EQUIPO_CALIDAD = [
    "Verónica Iriarte",
    "Cristina Merino",
    "Lisseth Aspíllaga",
    "Sandra Chavez",
    "Alejandro Herrera",
    "Katherin Hidalgo",
]

SUBCARPETA_DRIVE = "Productos Terminados"
PLANTILLA_NOMBRE = "Liberación PT - Plantilla Base"
SUFIJO_MES = " - PT"
MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


def nombre_mes_es(fecha: date) -> str:
    return f"{MESES_ES[fecha.month]} {fecha.year}{SUFIJO_MES}"


TEMP_LIBERACION_OPCIONES = ["0 a 4 °C", "<23 °C", "menor a -16°C"]

# Encabezados finales del historial / exportable (mismo orden que la plantilla FR_Liberacion
# editada — SIN "Grados Brix", que se eliminó de la plantilla porque ningún producto lo usa)
HEADERS_EXPORT = [
    "FECHA", "AREA", "CLIENTE", "N° de Muestra", "Producto", "Línea HACCP",
    "Lote (Juliano)", "Fecha Producción", "Fecha Vencimiento",
    "T° Liberación",
    "Peso (g)", "Diámetro (cm)", "Largo (cm)", "Ancho (cm)", "Altura (cm)",
    "Sabor/Olor/Color", "Textura", "Apariencia",
    "Integridad del Empaque (C/NC)", "Rotulado (C/NC)",
    "Conclusión (C/NC)", "Iniciales",
]


def iniciales(nombre_completo: str) -> str:
    partes = nombre_completo.strip().split()
    if len(partes) < 2:
        return nombre_completo[:2].upper()
    return (partes[0][0] + partes[-1][0]).upper()


# ----------------------------------------------------------------------------
# TABLA MIL-STD-105E - Nivel de Inspección Especial S-2, Inspección Rigurosa
# (Tightened, Tabla II-B), AQL 4.0%
# (lote_min, lote_max, letra código, n, Ac, Re)
# ----------------------------------------------------------------------------
SAMPLING_TABLE_S2 = [
    (2, 8, "A", 2, 0, 1),
    (9, 15, "A", 2, 0, 1),
    (16, 25, "B", 3, 0, 1),
    (26, 50, "B", 3, 0, 1),
    (51, 90, "B", 3, 0, 1),
    (91, 150, "C", 5, 0, 1),
    (151, 280, "C", 5, 0, 1),
    (281, 500, "C", 5, 0, 1),
    (501, 1200, "D", 8, 1, 2),
    (1201, 3200, "D", 8, 1, 2),
    (3201, 10000, "D", 8, 1, 2),
    (10001, 35000, "E", 13, 1, 2),
    (35001, 150000, "E", 13, 1, 2),
    (150001, 10_000_000, "E", 13, 1, 2),
]


def get_sample_size(lot_size: int):
    """Devuelve (letra_codigo, n, Ac, Re) según MIL-STD-105E Nivel S-2, Tightened, AQL 4.0%."""
    for low, high, letter, n, ac, re_ in SAMPLING_TABLE_S2:
        if low <= lot_size <= high:
            return letter, n, ac, re_
    if lot_size < 2:
        return "A", 2, 0, 1
    return "E", 13, 1, 2


# ----------------------------------------------------------------------------
# CARGA Y LIMPIEZA DE DATOS DEL EXCEL
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_specs(excel_bytes: bytes) -> pd.DataFrame:
    df_raw = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=SHEET_NAME,
        header=None,
        skiprows=7,
    )

    cols = [
        "col_extra", "producto", "linea_produccion", "linea_haccp", "tipo",
        "peso", "diametro", "largo_ancho", "altura", "organolepticas",
        "temp_almacenamiento", "vida_util",
        "envase_tipo", "presentacion", "unidades_empaque", "material_empaque",
        "rotulado",
    ]
    df_raw = df_raw.iloc[:, : len(cols)]
    df_raw.columns = cols
    df_raw = df_raw.dropna(subset=["producto"]).copy()

    def clean_txt(x):
        if pd.isna(x):
            return x
        return re.sub(r"\s+", " ", str(x)).strip()

    text_cols = [
        "producto", "linea_produccion", "linea_haccp", "peso", "diametro",
        "largo_ancho", "altura", "organolepticas", "temp_almacenamiento",
        "envase_tipo", "presentacion", "unidades_empaque", "material_empaque",
        "rotulado",
    ]
    for c in text_cols:
        df_raw[c] = df_raw[c].apply(clean_txt)

    df_raw["producto"] = df_raw["producto"].str.upper()
    df_raw["linea_haccp"] = df_raw["linea_haccp"].str.upper()
    df_raw["vida_util"] = pd.to_numeric(df_raw["vida_util"], errors="coerce")
    df_raw = df_raw[df_raw["tipo"].astype(str).str.upper().str.strip() == "PT"]

    return df_raw.reset_index(drop=True)


def campo_aplica(valor) -> bool:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return False
    v = str(valor).strip()
    return v not in ("", "-", "—", "nan", "None")


def construir_texto_envase(filas_producto: pd.DataFrame) -> str:
    """Combina las líneas de envase (primario/secundario) de todas las filas de un producto."""
    lineas = []
    vistos = set()
    contador_generico = 1
    for _, fila in filas_producto.iterrows():
        tipo = fila["envase_tipo"]
        pres = fila["presentacion"]
        unid = fila["unidades_empaque"]
        mat = fila["material_empaque"]
        clave = (tipo, pres, unid, mat)
        if clave in vistos:
            continue
        if not any(campo_aplica(x) for x in [tipo, pres, unid, mat]):
            continue
        vistos.add(clave)
        etiqueta = tipo if campo_aplica(tipo) else f"Envase adicional {contador_generico}:"
        if not campo_aplica(tipo):
            contador_generico += 1
        detalle_partes = []
        if campo_aplica(pres):
            detalle_partes.append(f"presentación {pres}")
        if campo_aplica(unid):
            detalle_partes.append(f"{unid}")
        if campo_aplica(mat):
            detalle_partes.append(f"material {mat}")
        lineas.append(f"- {etiqueta} " + ", ".join(detalle_partes))
    return "\n".join(lineas) if lineas else "Sin información de envase registrada."


def calcular_juliano(fecha: date) -> str:
    return str(fecha.timetuple().tm_yday).zfill(3)


# ----------------------------------------------------------------------------
# CONEXIÓN A GOOGLE DRIVE / SHEETS
# ----------------------------------------------------------------------------
def get_gsheet_client_and_drive():
    if not GSHEETS_DISPONIBLE:
        return None, None
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)
        drive = build_drive_service("drive", "v3", credentials=creds)
        return gc, drive
    except Exception:
        return None, None


TEMPLATE_SHEET_NAME = "Hoja 1"  # pestaña con el logo, encabezado y "CONSIDERACIONES"
FOOTER_MARCA = "CONSIDERACIONES"
FILA_ENCABEZADO_PLANTILLA = 4  # fila donde están los títulos de columna (FECHA, AREA, ...)


def _encontrar_fila_footer(ws) -> int | None:
    """Busca la fila donde empieza el bloque 'CONSIDERACIONES:' dentro de la hoja."""
    valores = ws.get_all_values()
    for idx, fila in enumerate(valores, start=1):
        if fila and fila[0].strip().upper().startswith(FOOTER_MARCA):
            return idx
    return None


def _buscar_archivo_en_carpeta(drive, nombre: str, carpeta_id: str, es_carpeta: bool = False):
    """Busca por nombre exacto dentro de una carpeta. Devuelve el ID o None."""
    tipo_mime = "application/vnd.google-apps.folder" if es_carpeta else "application/vnd.google-apps.spreadsheet"
    query = (
        f"name = '{nombre}' and '{carpeta_id}' in parents "
        f"and mimeType = '{tipo_mime}' and trashed = false"
    )
    resultado = drive.files().list(q=query, fields="files(id, name)").execute()
    archivos = resultado.get("files", [])
    return archivos[0]["id"] if archivos else None


def get_or_create_month_spreadsheet_id(drive, root_folder_id: str, fecha: date) -> str:
    """Devuelve el ID del Google Sheets del mes correspondiente.

    IMPORTANTE: las cuentas de servicio no tienen cuota propia de almacenamiento
    en Drive, así que NO pueden crear archivos nuevos (falla con
    'storageQuotaExceeded'). Por eso, si el archivo del mes no existe todavía,
    le pedimos a la persona que lo duplique manualmente una vez (queda con su
    propia cuota de Drive; la cuenta de servicio ya tiene acceso heredado de la carpeta).
    """
    subcarpeta_id = _buscar_archivo_en_carpeta(drive, SUBCARPETA_DRIVE, root_folder_id, es_carpeta=True)
    if subcarpeta_id is None:
        raise RuntimeError(f"No encontré la subcarpeta '{SUBCARPETA_DRIVE}' dentro de la carpeta raíz.")

    nombre_mes = nombre_mes_es(fecha)
    archivo_mes_id = _buscar_archivo_en_carpeta(drive, nombre_mes, subcarpeta_id)
    if archivo_mes_id:
        return archivo_mes_id

    raise RuntimeError(
        f"Todavía no existe el archivo del mes '{nombre_mes}' dentro de "
        f"'{SUBCARPETA_DRIVE}'. Duplica manualmente '{PLANTILLA_NOMBRE}' en esa "
        f"carpeta de Drive y renómbralo exactamente '{nombre_mes}' (Google no "
        f"permite que la cuenta de servicio cree archivos nuevos automáticamente)."
    )


def get_or_create_daily_worksheet(spreadsheet, fecha_produccion: date):
    """Abre (o crea, duplicando la plantilla) la pestaña del día."""
    titulo_hoja = fecha_produccion.strftime("%Y-%m-%d")

    try:
        return spreadsheet.worksheet(titulo_hoja)
    except gspread.exceptions.WorksheetNotFound:
        pass

    # No existe todavía la hoja del día: duplicamos la plantilla oficial
    plantilla = spreadsheet.worksheet(TEMPLATE_SHEET_NAME)
    ws = spreadsheet.duplicate_sheet(
        source_sheet_id=plantilla.id,
        new_sheet_name=titulo_hoja,
    )

    # Quitamos las filas vacías de ejemplo (entre el encabezado y "CONSIDERACIONES")
    fila_footer = _encontrar_fila_footer(ws)
    primera_fila_datos = FILA_ENCABEZADO_PLANTILLA + 1
    if fila_footer and fila_footer > primera_fila_datos:
        ws.delete_rows(primera_fila_datos, fila_footer - 1)

    return ws


def guardar_en_google_sheets(filas: list) -> tuple:
    """Guarda las filas en la hoja del día correspondiente, dentro del Sheets del mes."""
    gc, drive = get_gsheet_client_and_drive()
    if gc is None or drive is None:
        return False, "No se pudo conectar a Google Drive/Sheets (revisa los Secrets configurados)."
    try:
        root_folder_id = st.secrets.get("ROOT_FOLDER_ID")
        if not root_folder_id:
            return False, "No se encontró ROOT_FOLDER_ID en los Secrets."

        fecha_prod = st.session_state.fecha_produccion

        spreadsheet_id = get_or_create_month_spreadsheet_id(drive, root_folder_id, fecha_prod)
        spreadsheet = gc.open_by_key(spreadsheet_id)
        ws = get_or_create_daily_worksheet(spreadsheet, fecha_prod)

        fila_footer = _encontrar_fila_footer(ws)
        if fila_footer is None:
            # No se encontró el bloque de consideraciones: agregamos al final como respaldo
            ws.append_rows(filas)
        else:
            ws.insert_rows(filas, row=fila_footer)

        return True, (
            f"Guardado en '{nombre_mes_es(fecha_prod)}' → hoja "
            f"'{fecha_prod.strftime('%Y-%m-%d')}'."
        )
    except Exception as e:
        return False, f"Error al guardar en Google Sheets: {e}"


# ----------------------------------------------------------------------------
# CARGA DEL ARCHIVO EXCEL DE ESPECIFICACIONES
# ----------------------------------------------------------------------------
def get_excel_bytes():
    try:
        with open(EXCEL_PATH, "rb") as f:
            return f.read()
    except FileNotFoundError:
        st.warning(
            "No encontré el archivo en `data/MA_BASE_DATOS_PRODUCTOS_TERMINADO.xlsx` "
            "dentro del repositorio."
        )
        up = st.file_uploader("Sube el Excel de especificaciones (.xlsx)", type=["xlsx"])
        if up is not None:
            return up.read()
        st.stop()


# ----------------------------------------------------------------------------
# ESTADO / WIZARD
# ----------------------------------------------------------------------------
if "step" not in st.session_state:
    st.session_state.step = 1


def go_next():
    st.session_state.step += 1


def go_back():
    st.session_state.step -= 1


excel_bytes = get_excel_bytes()
specs_df = load_specs(excel_bytes)

# ============================================================================
# PASO 1 - PORTADA
# ============================================================================
if st.session_state.step == 1:
    st.title("✅ Registro de Parámetros de Calidad")
    st.subheader("Producto Terminado - Línea de Producción B2B Starbucks")
    st.markdown(
        """
        Esta aplicación permite al **equipo de calidad** registrar la inspección
        de producto terminado de la línea B2B Starbucks, siguiendo el plan de
        muestreo **MIL-STD-105E (Nivel Especial S-2, Inspección Rigurosa, AQL 4.0%)**.

        El registro queda guardado automáticamente en un historial en Google Sheets.
        """
    )
    st.button("Comenzar registro ➜", on_click=go_next, type="primary")

# ============================================================================
# PASO 2 - EQUIPO DE CALIDAD Y FECHA DE PRODUCCIÓN
# ============================================================================
elif st.session_state.step == 2:
    st.header("1️⃣ Datos del registro")

    responsable = st.selectbox(
        "Equipo de calidad (responsable del registro)",
        EQUIPO_CALIDAD,
        index=EQUIPO_CALIDAD.index(st.session_state.get("responsable", EQUIPO_CALIDAD[0]))
        if st.session_state.get("responsable") in EQUIPO_CALIDAD else 0,
    )

    fecha_produccion = st.date_input(
        "Fecha de producción (= fecha de registro)",
        value=st.session_state.get("fecha_produccion", None),
        format="DD/MM/YYYY",
    )
    if fecha_produccion is None:
        st.caption("⚠️ Selecciona la fecha de producción para continuar (no se precarga sola).")

    st.info(f"**Cliente:** {CLIENTE_FIJO}  |  **Área:** {AREA_FIJA}")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary", disabled=fecha_produccion is None):
            st.session_state.responsable = responsable
            st.session_state.fecha_produccion = fecha_produccion
            go_next()
            st.rerun()

# ============================================================================
# PASO 3 - LINEA HACCP Y PRODUCTO
# ============================================================================
elif st.session_state.step == 3:
    st.header("2️⃣ Línea HACCP y producto")

    lineas = sorted(specs_df["linea_haccp"].dropna().unique().tolist())
    linea_sel = st.selectbox("Línea de producción HACCP", lineas)

    productos_linea = sorted(
        specs_df.loc[specs_df["linea_haccp"] == linea_sel, "producto"].unique().tolist()
    )
    producto_sel = st.selectbox("Producto", productos_linea)

    filas_prod = specs_df[specs_df["producto"] == producto_sel]
    fila_ref = filas_prod.iloc[0]
    tiene_diametro = campo_aplica(fila_ref["diametro"])
    tiene_largo_ancho = campo_aplica(fila_ref["largo_ancho"])

    with st.expander("📋 Ficha técnica de referencia (solo informativa)", expanded=True):
        c1, c2, c3 = st.columns(3)
        c1.metric("Peso especificado", fila_ref["peso"])
        if tiene_diametro:
            c2.metric("Diámetro especificado", fila_ref["diametro"])
        elif tiene_largo_ancho:
            c2.metric("Largo x Ancho especificado", fila_ref["largo_ancho"])
        else:
            c2.metric("Diámetro / Largo-Ancho", "No aplica")
        c3.metric("Altura/Espesor especificado", fila_ref["altura"])
        st.markdown(f"**Organolépticas esperadas:**\n\n{fila_ref['organolepticas']}")
        st.markdown("**Empaque de referencia:**")
        st.text(construir_texto_envase(filas_prod))
        st.markdown(f"**Rotulado:** {fila_ref['rotulado']}")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.linea_haccp = linea_sel
            st.session_state.producto = producto_sel
            go_next()
            st.rerun()

# ============================================================================
# PASO 4 - TEMPERATURAS (ALMACENAMIENTO Y LIBERACIÓN) Y VIDA ÚTIL
# ============================================================================
elif st.session_state.step == 4:
    st.header("3️⃣ Temperaturas")

    producto_sel = st.session_state.producto
    filas_prod = specs_df[specs_df["producto"] == producto_sel]

    opciones_temp = filas_prod["temp_almacenamiento"].dropna().unique().tolist()
    temp_almac_sel = st.selectbox("Temperatura de almacenamiento", opciones_temp)

    temp_liberacion_sel = st.selectbox("Temperatura de liberación", TEMP_LIBERACION_OPCIONES)

    fila_temp = filas_prod[filas_prod["temp_almacenamiento"] == temp_almac_sel].iloc[0]
    vida_util_dias = int(fila_temp["vida_util"])
    fecha_produccion = st.session_state.fecha_produccion
    fecha_venc = fecha_produccion + timedelta(days=vida_util_dias)

    col1, col2 = st.columns(2)
    col1.metric("Vida útil", f"{vida_util_dias} días")
    col2.metric("Fecha de vencimiento", fecha_venc.strftime("%d/%m/%Y"))

    colA, colB = st.columns(2)
    with colA:
        st.button("⬅ Atrás", on_click=go_back)
    with colB:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.temp_almacenamiento = temp_almac_sel
            st.session_state.temp_liberacion = temp_liberacion_sel
            st.session_state.vida_util = vida_util_dias
            st.session_state.fecha_vencimiento = fecha_venc
            st.session_state.fila_spec = fila_temp.to_dict()
            st.session_state.filas_producto_completo = filas_prod
            go_next()
            st.rerun()

# ============================================================================
# PASO 5 - LOTE (CÓDIGO JULIANO)
# ============================================================================
elif st.session_state.step == 5:
    st.header("4️⃣ Lote (código Juliano)")

    modo_juliano = st.radio(
        "¿Cómo deseas obtener el código Juliano?",
        ["Calcular automáticamente (según fecha de producción)", "Ingresar manualmente"],
        horizontal=False,
    )

    if modo_juliano.startswith("Calcular"):
        juliano = calcular_juliano(st.session_state.fecha_produccion)
        st.success(f"Código Juliano calculado: **{juliano}** "
                   f"(día N° {int(juliano)} del año {st.session_state.fecha_produccion.year})")
    else:
        juliano = st.text_input("Ingresa el código Juliano manualmente", value="")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary", disabled=not str(juliano).strip()):
            st.session_state.lote_juliano = str(juliano).strip()
            go_next()
            st.rerun()

# ============================================================================
# PASO 6 - TAMAÑO DE BATCH Y MUESTREO
# ============================================================================
elif st.session_state.step == 6:
    st.header("5️⃣ Tamaño del batch y muestreo (MIL-STD-105E, S-2, Rigurosa, AQL 4.0%)")

    batch_size = st.number_input(
        "¿Cuántas unidades tiene el batch?", min_value=1, step=1,
        value=st.session_state.get("batch_size", 300),
    )

    letra, n_muestras, ac, re_ = get_sample_size(int(batch_size))
    st.success(
        f"Para un batch de **{int(batch_size)}** unidades: letra código **{letra}** → "
        f"**n = {n_muestras}** muestras. Criterio: Aceptar con **{ac}** o menos no conformes, "
        f"Rechazar con **{re_}** o más."
    )

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.batch_size = int(batch_size)
            st.session_state.n_muestras = n_muestras
            st.session_state.letra_codigo = letra
            st.session_state.ac = ac
            st.session_state.re_ = re_
            go_next()
            st.rerun()

# ============================================================================
# PASO 7 - REGISTRO DE PARÁMETROS POR MUESTRA
# ============================================================================
elif st.session_state.step == 7:
    st.header("6️⃣ Registro de parámetros por muestra")

    fila_spec = st.session_state.fila_spec
    n = st.session_state.n_muestras
    tiene_diametro = campo_aplica(fila_spec["diametro"])
    tiene_largo_ancho = campo_aplica(fila_spec["largo_ancho"])

    st.markdown(f"**Producto:** {st.session_state.producto} &nbsp;|&nbsp; "
                f"**N° de muestras a evaluar:** {n}")

    # (clave, etiqueta_pestaña, pregunta)
    parametros = [("peso", "Peso", f"¿El peso oscila en el rango especificado? ({fila_spec['peso']} g)")]
    if tiene_diametro:
        parametros.append(
            ("diametro", "Diámetro", f"¿El diámetro oscila en el rango especificado? ({fila_spec['diametro']} cm)")
        )
    elif tiene_largo_ancho:
        parametros.append(
            ("largo_ancho", "Largo/Ancho",
             f"¿Las medidas de largo y ancho cumplen con lo especificado? ({fila_spec['largo_ancho']})")
        )
    parametros.append(
        ("altura", "Altura/Espesor", f"¿La altura/espesor oscila en el rango especificado? ({fila_spec['altura']} cm)")
    )
    parametros.append(("sabor_olor_color", "Sabor/Olor/Color",
                        "¿El sabor, olor y color cumplen con lo especificado?"))
    parametros.append(("textura", "Textura", "¿La textura cumple con lo especificado?"))
    parametros.append(("apariencia", "Apariencia", "¿La apariencia (incluye decorado, si aplica) cumple con lo especificado?"))
    parametros.append(("empaque", "Empaque", "¿El empaque cumple con lo especificado (ver referencia abajo)?"))
    parametros.append(
        ("rotulado", "Rotulado", f"¿El rotulado cumple con lo especificado? ({fila_spec['rotulado']})")
    )

    with st.expander("Ver detalle de características organolépticas esperadas"):
        st.text(fila_spec["organolepticas"])

    st.caption("Completa cada muestra tocando 'Conforme' o 'No conforme'. Si algún parámetro "
               "tiene 1 o más 'No conforme', se habilitará un cuadro de comentario al final de esa pestaña.")

    if "respuestas" not in st.session_state:
        st.session_state.respuestas = {}
    if "comentarios_parametro" not in st.session_state:
        st.session_state.comentarios_parametro = {}

    tabs = st.tabs([label for _, label, _ in parametros])
    for (clave, label, pregunta), tab in zip(parametros, tabs):
        with tab:
            st.markdown(f"**{pregunta}**")
            if clave == "empaque":
                with st.expander("Ver referencia completa de empaque especificado", expanded=True):
                    st.text(construir_texto_envase(st.session_state.filas_producto_completo))
            hay_no_conforme = False
            for i in range(n):
                key = f"resp_{clave}_{i}"
                if key not in st.session_state.respuestas:
                    st.session_state.respuestas[key] = "Conforme"
                valor_actual = st.session_state.respuestas[key]
                idx_default = 0 if valor_actual == "Conforme" else 1
                st.session_state.respuestas[key] = st.radio(
                    f"Muestra {i + 1}",
                    ["Conforme", "No conforme"],
                    index=idx_default,
                    horizontal=True,
                    key=f"widget_{key}",
                )
                if st.session_state.respuestas[key] == "No conforme":
                    hay_no_conforme = True

            if hay_no_conforme:
                st.session_state.comentarios_parametro[clave] = st.text_area(
                    f"Comentario / corrección para '{label}' (hay al menos 1 muestra No conforme)",
                    value=st.session_state.comentarios_parametro.get(clave, ""),
                    key=f"comentario_{clave}",
                )
            else:
                st.session_state.comentarios_parametro[clave] = ""

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.parametros_muestra = parametros
            go_next()
            st.rerun()

# ============================================================================
# PASO 8 - CONCLUSIÓN DEL REGISTRO (liberación del batch)
# ============================================================================
elif st.session_state.step == 8:
    st.header("7️⃣ Conclusión del registro")

    st.markdown("Con base en todo lo evaluado, indica si el batch se libera o no.")
    conclusion = st.radio(
        "Conclusión (¿se libera el batch?)",
        ["Conforme", "No conforme"],
        index=0 if st.session_state.get("conclusion", "Conforme") == "Conforme" else 1,
        horizontal=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Generar resumen ➜", type="primary"):
            st.session_state.conclusion = conclusion
            go_next()
            st.rerun()

# ============================================================================
# PASO 9 - RESUMEN FINAL, GUARDADO EN GOOGLE SHEETS Y EXPORTACIÓN
# ============================================================================
elif st.session_state.step == 9:
    st.header("8️⃣ Resumen final")

    fila_spec = st.session_state.fila_spec
    n = st.session_state.n_muestras
    parametros = st.session_state.parametros_muestra

    st.subheader("Datos del registro")
    resumen_info = pd.DataFrame(
        {
            "Campo": [
                "Equipo de calidad", "Fecha de producción", "Cliente", "Área",
                "Línea HACCP", "Producto", "Temperatura de almacenamiento",
                "Temperatura de liberación", "Vida útil (días)", "Fecha de vencimiento",
                "Lote (Juliano)", "Tamaño de batch", "Letra código muestreo",
                "N° de muestras", "Conclusión",
            ],
            "Valor": [
                st.session_state.responsable,
                st.session_state.fecha_produccion.strftime("%d/%m/%Y"),
                CLIENTE_FIJO, AREA_FIJA,
                st.session_state.linea_haccp, st.session_state.producto,
                st.session_state.temp_almacenamiento, st.session_state.temp_liberacion,
                st.session_state.vida_util,
                st.session_state.fecha_vencimiento.strftime("%d/%m/%Y"),
                st.session_state.lote_juliano, st.session_state.batch_size,
                st.session_state.letra_codigo, n,
                st.session_state.conclusion,
            ],
        }
    )
    st.dataframe(resumen_info, use_container_width=True, hide_index=True)

    st.subheader("Resultados por parámetro (muestras)")
    conteo = {}
    for clave, label, _ in parametros:
        valores = [st.session_state.respuestas[f"resp_{clave}_{i}"] for i in range(n)]
        conteo[label] = {
            "Conforme": valores.count("Conforme"),
            "No conforme": valores.count("No conforme"),
        }
    st.dataframe(pd.DataFrame(conteo).T, use_container_width=True)

    # ------------------------------------------------------------------
    # Armado de las filas exportables (una fila por muestra)
    # ------------------------------------------------------------------
    valores_por_parametro = {
        clave: [st.session_state.respuestas[f"resp_{clave}_{i}"] for i in range(n)]
        for clave, _, _ in parametros
    }

    def valor_muestra(clave, i):
        valor = valores_por_parametro.get(clave, [None] * n)[i]
        if valor in (None, ""):
            return "No aplica"
        if valor == "No conforme":
            comentario = st.session_state.comentarios_parametro.get(clave, "").strip()
            if comentario:
                return f"No conforme: {comentario}"
        return valor

    filas_export = []
    for i in range(n):
        fila = [
            st.session_state.fecha_produccion.strftime("%d/%m/%Y"),  # FECHA
            AREA_FIJA,
            CLIENTE_FIJO,
            i + 1,  # N° de Muestra
            st.session_state.producto,
            st.session_state.linea_haccp,
            st.session_state.lote_juliano,
            st.session_state.fecha_produccion.strftime("%d/%m/%Y"),
            st.session_state.fecha_vencimiento.strftime("%d/%m/%Y"),
            st.session_state.temp_liberacion,
            valor_muestra("peso", i),
            valor_muestra("diametro", i) if "diametro" in valores_por_parametro else "No aplica",
            "No aplica",  # Largo (se completa abajo si el producto usa largo x ancho)
            "No aplica",  # Ancho (no se usa: la pregunta es combinada, va toda en "Largo")
            valor_muestra("altura", i),
            valor_muestra("sabor_olor_color", i),
            valor_muestra("textura", i),
            valor_muestra("apariencia", i),
            valor_muestra("empaque", i),
            valor_muestra("rotulado", i),
            st.session_state.conclusion,
            iniciales(st.session_state.responsable),
        ]
        # Si el producto usa largo x ancho en vez de diámetro, lo ponemos en la columna Largo
        # (antes se escribía por error en la columna Ancho, índice 13 en vez de 12)
        if "largo_ancho" in valores_por_parametro:
            fila[12] = valor_muestra("largo_ancho", i)
        filas_export.append(fila)

    export_df = pd.DataFrame(filas_export, columns=HEADERS_EXPORT)

    with st.expander("Ver todas las filas que se guardarán/exportarán"):
        st.dataframe(export_df, use_container_width=True, hide_index=True)

    # ------------------------------------------------------------------
    # Guardar en Google Sheets
    # ------------------------------------------------------------------
    st.divider()
    st.subheader("Guardar historial")

    if st.button("💾 Guardar en Google Sheets (historial)", type="primary"):
        exito, mensaje = guardar_en_google_sheets(filas_export)
        if exito:
            st.success(mensaje)
        else:
            st.error(mensaje)

    # ------------------------------------------------------------------
    # Exportación local
    # ------------------------------------------------------------------
    csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Registro")
    excel_buffer.seek(0)

    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        st.download_button(
            "⬇ Descargar CSV",
            data=csv_bytes,
            file_name=f"registro_{st.session_state.producto}_{st.session_state.fecha_produccion}.csv",
            mime="text/csv",
        )
    with col3:
        st.download_button(
            "⬇ Descargar Excel",
            data=excel_buffer,
            file_name=f"registro_{st.session_state.producto}_{st.session_state.fecha_produccion}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    if st.button("🔄 Nuevo registro"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
