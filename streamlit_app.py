"""
App de Registro de Parámetros de Calidad - Producto Terminado B2B Starbucks
=============================================================================
Streamlit + GitHub deploy ready.

Estructura esperada del repo:
  app.py
  requirements.txt
  data/MA_BASE_DATOS_PRODUCTOS_TERMINADO.xlsx   <-- este Excel debe viajar en el repo
"""

import io
import re
from datetime import date, timedelta

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Registro de Calidad - Producto Terminado B2B Starbucks",
    page_icon="✅",
    layout="wide",
)

EXCEL_PATH = "data/MA_BASE_DATOS_PRODUCTOS_TERMINADO.xlsx"
SHEET_NAME = "ALMENARA-PT_2026"

# ----------------------------------------------------------------------------
# TABLA MIL-STD-105E / ANSI-ASQ Z1.4 - Nivel de Inspección General II
# Muestreo simple normal, AQL 2.5%
# (rango de tamaño de lote, letra código, tamaño de muestra n)
# ----------------------------------------------------------------------------
SAMPLING_TABLE = [
    (2, 8, "A", 2),
    (9, 15, "B", 3),
    (16, 25, "C", 5),
    (26, 50, "D", 8),
    (51, 90, "E", 13),
    (91, 150, "F", 20),
    (151, 280, "G", 32),
    (281, 500, "H", 50),
    (501, 1200, "J", 80),
    (1201, 3200, "K", 125),
    (3201, 10000, "L", 200),
    (10001, 35000, "M", 315),
    (35001, 150000, "N", 500),
    (150001, 500000, "P", 800),
    (500001, 10_000_000, "Q", 1250),
]


def get_sample_size(lot_size: int):
    """Devuelve (letra_codigo, n) según MIL-STD-105E Nivel II, AQL 2.5%."""
    for low, high, letter, n in SAMPLING_TABLE:
        if low <= lot_size <= high:
            return letter, n
    if lot_size < 2:
        return "A", 2
    return "Q", 1250


# ----------------------------------------------------------------------------
# CARGA Y LIMPIEZA DE DATOS DEL EXCEL
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_specs(excel_bytes: bytes) -> pd.DataFrame:
    """Lee la hoja de producto terminado y arma un dataframe limpio."""
    df_raw = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=SHEET_NAME,
        header=None,
        skiprows=7,  # los datos empiezan en la fila 8 (1-indexed)
    )

    cols = [
        "col_extra", "producto", "linea_produccion", "linea_haccp", "tipo",
        "peso", "diametro", "altura", "organolepticas", "envase",
        "temp_almacenamiento", "vida_util", "inspeccion",
    ]
    df_raw = df_raw.iloc[:, : len(cols)]
    df_raw.columns = cols

    df_raw = df_raw.dropna(subset=["producto"]).copy()

    # Normalización de texto: quitar espacios extra, pasar a formato consistente
    def clean_txt(x):
        if pd.isna(x):
            return x
        return re.sub(r"\s+", " ", str(x)).strip()

    for c in ["producto", "linea_produccion", "linea_haccp", "peso",
              "diametro", "altura", "organolepticas", "envase",
              "temp_almacenamiento"]:
        df_raw[c] = df_raw[c].apply(clean_txt)

    df_raw["producto"] = df_raw["producto"].str.upper()
    df_raw["linea_haccp"] = df_raw["linea_haccp"].str.upper()

    df_raw["vida_util"] = pd.to_numeric(df_raw["vida_util"], errors="coerce")

    df_raw = df_raw[df_raw["tipo"].astype(str).str.upper().str.strip() == "PT"]

    return df_raw.reset_index(drop=True)


def diametro_aplica(valor: str) -> bool:
    if valor is None:
        return False
    v = str(valor).strip()
    return v not in ("", "-", "—", "nan")


# ----------------------------------------------------------------------------
# CARGA DEL ARCHIVO (repo por defecto, o subida manual como respaldo)
# ----------------------------------------------------------------------------
def get_excel_bytes():
    try:
        with open(EXCEL_PATH, "rb") as f:
            return f.read()
    except FileNotFoundError:
        st.warning(
            "No encontré el archivo en `data/MA_BASE_DATOS_PRODUCTOS_TERMINADO.xlsx` "
            "dentro del repositorio. Puedes subirlo manualmente mientras lo agregas al repo."
        )
        up = st.file_uploader("Sube el Excel de especificaciones (.xlsx)", type=["xlsx"])
        if up is not None:
            return up.read()
        st.stop()


# ----------------------------------------------------------------------------
# ESTADO DE LA APP (wizard de pasos)
# ----------------------------------------------------------------------------
if "step" not in st.session_state:
    st.session_state.step = 1

def go_next():
    st.session_state.step += 1

def go_back():
    st.session_state.step -= 1


# ----------------------------------------------------------------------------
# CARGA DE DATOS
# ----------------------------------------------------------------------------
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
        Esta aplicación permite al **supervisor de calidad** registrar de forma
        estandarizada la inspección de producto terminado de la línea B2B
        Starbucks, siguiendo el plan de muestreo **Military Standard 105E
        (Nivel de Inspección General II, AQL 2.5%)**.

        El flujo de la aplicación es el siguiente:
        1. Datos del supervisor y fecha de registro.
        2. Selección de la línea HACCP y el producto a inspeccionar.
        3. Temperatura de almacenamiento y cálculo automático de la fecha de vencimiento.
        4. Tamaño del batch y cálculo automático del número de muestras a evaluar.
        5. Registro de parámetros (peso, diámetro, altura, organolépticas) por muestra.
        6. Verificación general de envase, correcciones y exportación del registro.
        """
    )
    st.button("Comenzar registro ➜", on_click=go_next, type="primary")

# ============================================================================
# PASO 2 - SUPERVISOR Y FECHA
# ============================================================================
elif st.session_state.step == 2:
    st.header("1️⃣ Datos del registro")

    supervisor = st.text_input(
        "Nombre del supervisor de calidad", value=st.session_state.get("supervisor", "")
    )
    hoy = date.today()
    st.info(f"📅 Fecha de registro: **{hoy.strftime('%d/%m/%Y')}** (automática)")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary", disabled=not supervisor.strip()):
            st.session_state.supervisor = supervisor.strip()
            st.session_state.fecha_registro = hoy
            go_next()
            st.rerun()

    if not supervisor.strip():
        st.caption("Ingresa el nombre del supervisor para continuar.")

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

    # Ficha de referencia (solo informativa)
    filas_prod = specs_df[specs_df["producto"] == producto_sel]
    fila_ref = filas_prod.iloc[0]

    with st.expander("📋 Ficha técnica de referencia (solo informativa)", expanded=True):
        c1, c2, c3 = st.columns(3)
        c1.metric("Peso especificado", fila_ref["peso"])
        c2.metric(
            "Diámetro especificado",
            fila_ref["diametro"] if diametro_aplica(fila_ref["diametro"]) else "No aplica",
        )
        c3.metric("Altura/Espesor especificado", fila_ref["altura"])
        st.markdown(f"**Características organolépticas esperadas:**\n\n{fila_ref['organolepticas']}")
        st.markdown(f"**Envase:** {fila_ref['envase']}")

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
# PASO 4 - TEMPERATURA Y VIDA ÚTIL
# ============================================================================
elif st.session_state.step == 4:
    st.header("3️⃣ Temperatura de almacenamiento")

    producto_sel = st.session_state.producto
    filas_prod = specs_df[specs_df["producto"] == producto_sel]

    opciones_temp = filas_prod["temp_almacenamiento"].dropna().unique().tolist()
    temp_sel = st.selectbox("Temperatura de almacenamiento", opciones_temp)

    fila_temp = filas_prod[filas_prod["temp_almacenamiento"] == temp_sel].iloc[0]
    vida_util_dias = int(fila_temp["vida_util"])
    fecha_registro = st.session_state.fecha_registro
    fecha_venc = fecha_registro + timedelta(days=vida_util_dias)

    col1, col2 = st.columns(2)
    col1.metric("Vida útil", f"{vida_util_dias} días")
    col2.metric("Fecha de vencimiento", fecha_venc.strftime("%d/%m/%Y"))

    colA, colB = st.columns(2)
    with colA:
        st.button("⬅ Atrás", on_click=go_back)
    with colB:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.temp_almacenamiento = temp_sel
            st.session_state.vida_util = vida_util_dias
            st.session_state.fecha_vencimiento = fecha_venc
            st.session_state.fila_spec = fila_temp.to_dict()
            go_next()
            st.rerun()

# ============================================================================
# PASO 5 - TAMAÑO DE BATCH Y MUESTREO
# ============================================================================
elif st.session_state.step == 5:
    st.header("4️⃣ Tamaño del batch y muestreo (MIL-STD-105E, Nivel II, AQL 2.5%)")

    batch_size = st.number_input(
        "¿Cuántas unidades tiene el batch?", min_value=1, step=1,
        value=st.session_state.get("batch_size", 100),
    )

    letra, n_muestras = get_sample_size(int(batch_size))
    st.success(
        f"Para un batch de **{int(batch_size)}** unidades, corresponde la letra código "
        f"**{letra}** → tamaño de muestra **n = {n_muestras}**."
    )

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.batch_size = int(batch_size)
            st.session_state.n_muestras = n_muestras
            st.session_state.letra_codigo = letra
            go_next()
            st.rerun()

# ============================================================================
# PASO 6 - REGISTRO DE PARÁMETROS POR MUESTRA (GRID)
# ============================================================================
elif st.session_state.step == 6:
    st.header("5️⃣ Registro de parámetros por muestra")

    fila_spec = st.session_state.fila_spec
    n = st.session_state.n_muestras
    tiene_diametro = diametro_aplica(fila_spec["diametro"])

    st.markdown(f"**Producto:** {st.session_state.producto} &nbsp;|&nbsp; "
                f"**N° de muestras a evaluar:** {n}")

    st.markdown(f"- ¿El **peso** oscila en el rango especificado? ({fila_spec['peso']} g)")
    if tiene_diametro:
        st.markdown(f"- ¿El **diámetro** oscila en el rango especificado? ({fila_spec['diametro']} cm)")
    st.markdown(f"- ¿La **altura/espesor** oscila en el rango especificado? ({fila_spec['altura']} cm)")
    st.markdown("- ¿Las **características organolépticas** (apariencia, sabor, olor, color, "
                "textura) cumplen con lo especificado?")
    with st.expander("Ver detalle de características organolépticas esperadas"):
        st.text(fila_spec["organolepticas"])

    # Construcción de la tabla editable
    columnas = {
        "N° Muestra": [f"Muestra {i+1}" for i in range(n)],
        f"Peso ({fila_spec['peso']} g) - Conforme?": ["Conforme"] * n,
    }
    if tiene_diametro:
        columnas[f"Diámetro ({fila_spec['diametro']} cm) - Conforme?"] = ["Conforme"] * n
    columnas[f"Altura/Espesor ({fila_spec['altura']} cm) - Conforme?"] = ["Conforme"] * n
    columnas["Organolépticas - Conforme?"] = ["Conforme"] * n

    df_grid = pd.DataFrame(columnas)

    column_config = {
        col: st.column_config.SelectboxColumn(
            col, options=["Conforme", "No conforme"], required=True
        )
        for col in df_grid.columns
        if col != "N° Muestra"
    }
    column_config["N° Muestra"] = st.column_config.TextColumn("N° Muestra", disabled=True)

    df_editado = st.data_editor(
        df_grid,
        column_config=column_config,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="grid_muestras",
    )

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.df_muestras = df_editado
            go_next()
            st.rerun()

# ============================================================================
# PASO 7 - ENVASE, CORRECCIONES, RESUMEN Y EXPORTACIÓN
# ============================================================================
elif st.session_state.step == 7:
    st.header("6️⃣ Envase, correcciones y resumen final")

    fila_spec = st.session_state.fila_spec

    envase_conforme = st.radio(
        f"¿El envase del batch cumple con: **{fila_spec['envase']}**?",
        ["Conforme", "No conforme"],
        horizontal=True,
    )

    correcciones = st.text_area(
        "Correcciones / observaciones (campo libre)",
        placeholder="Escribe aquí cualquier corrección, acción tomada u observación adicional...",
    )

    st.divider()
    st.subheader("Resumen del registro")

    df_muestras = st.session_state.df_muestras
    resumen_cols = [c for c in df_muestras.columns if c != "N° Muestra"]
    conteo = {
        c: {
            "Conforme": int((df_muestras[c] == "Conforme").sum()),
            "No conforme": int((df_muestras[c] == "No conforme").sum()),
        }
        for c in resumen_cols
    }
    df_conteo = pd.DataFrame(conteo).T
    st.dataframe(df_conteo, use_container_width=True)

    resumen_info = pd.DataFrame(
        {
            "Campo": [
                "Supervisor de calidad", "Fecha de registro", "Línea HACCP", "Producto",
                "Temperatura de almacenamiento", "Vida útil (días)", "Fecha de vencimiento",
                "Tamaño de batch", "Letra código muestreo", "N° de muestras",
                "Envase (batch)", "Correcciones/observaciones",
            ],
            "Valor": [
                st.session_state.supervisor,
                st.session_state.fecha_registro.strftime("%d/%m/%Y"),
                st.session_state.linea_haccp,
                st.session_state.producto,
                st.session_state.temp_almacenamiento,
                st.session_state.vida_util,
                st.session_state.fecha_vencimiento.strftime("%d/%m/%Y"),
                st.session_state.batch_size,
                st.session_state.letra_codigo,
                st.session_state.n_muestras,
                envase_conforme,
                correcciones if correcciones.strip() else "-",
            ],
        }
    )
    st.dataframe(resumen_info, use_container_width=True, hide_index=True)

    # ------------------------------------------------------------------
    # Armado del registro final exportable
    # ------------------------------------------------------------------
    export_df = df_muestras.copy()
    export_df.insert(0, "Supervisor", st.session_state.supervisor)
    export_df.insert(1, "Fecha de registro", st.session_state.fecha_registro.strftime("%d/%m/%Y"))
    export_df.insert(2, "Línea HACCP", st.session_state.linea_haccp)
    export_df.insert(3, "Producto", st.session_state.producto)
    export_df.insert(4, "Temperatura de almacenamiento", st.session_state.temp_almacenamiento)
    export_df.insert(5, "Fecha de vencimiento", st.session_state.fecha_vencimiento.strftime("%d/%m/%Y"))
    export_df.insert(6, "Tamaño de batch", st.session_state.batch_size)
    export_df.insert(7, "N° de muestras", st.session_state.n_muestras)
    export_df["Envase (batch) - Conforme?"] = envase_conforme
    export_df["Correcciones / Observaciones"] = correcciones if correcciones.strip() else "-"

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
            file_name=f"registro_calidad_{st.session_state.producto}_{st.session_state.fecha_registro}.csv",
            mime="text/csv",
        )
    with col3:
        st.download_button(
            "⬇ Descargar Excel",
            data=excel_buffer,
            file_name=f"registro_calidad_{st.session_state.producto}_{st.session_state.fecha_registro}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    if st.button("🔄 Nuevo registro"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
