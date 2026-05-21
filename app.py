import streamlit as st
import pandas as pd
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import json
import base64
import os

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
FICHERO_TRABAJADORES = "Informe_trabajadores_domicilio_imputacion.XLS"
FICHERO_CENTROS      = "DIRECCIONES_CENTROS.xlsx"
LOGO_PATH            = "LOGO_ARGIA_2026.png"
ORS_API_KEY          = st.secrets.get("ORS_API_KEY", "")
GOOGLE_SHEETS_NAME   = st.secrets.get("GOOGLE_SHEETS_NAME", "HuellaCarbonoTransporte2025")

VEHICULOS_CON_COMBUSTIBLE = ["Coche", "Furgoneta", "Moto"]
TIPOS_COMBUSTIBLE         = ["Gasolina", "Diésel", "Eléctrico", "Híbrido"]
MODOS_SIN_COMBUSTIBLE     = ["Transporte público (bus / metro / tren)", "A pie o en bicicleta"]
TODOS_MODOS               = VEHICULOS_CON_COMBUSTIBLE + MODOS_SIN_COMBUSTIBLE

# Colores corporativos Argia
COLOR_ROSA      = "#F4A7B4"
COLOR_VERDE     = "#A8D5B5"
COLOR_TURQUESA  = "#7BC4C4"
COLOR_GRIS      = "#6B6B6B"
COLOR_FONDO     = "#F9F9F9"
COLOR_BLANCO    = "#FFFFFF"

# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────
def get_logo_base64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return None

@st.cache_data
def cargar_datos():
    df_trab = pd.read_excel(FICHERO_TRABAJADORES, engine="xlrd", sheet_name="TRABAJADORES")
    df_imp  = pd.read_excel(FICHERO_TRABAJADORES, engine="xlrd", sheet_name="IMPUTACIONES")
    df_cent = pd.read_excel(FICHERO_CENTROS)

    df_trab.columns = df_trab.columns.str.strip()
    df_imp.columns  = df_imp.columns.str.strip()
    df_cent.columns = df_cent.columns.str.strip()

    # Código como string con ceros a la izquierda (4 dígitos)
    df_trab["CODIGO"] = df_trab["CODIGO"].astype(int).apply(lambda x: str(x).zfill(4))
    df_imp["CODIGO"]  = df_imp["CODIGO"].astype(int).apply(lambda x: str(x).zfill(4))
    df_trab["DNI"]    = df_trab["DNI"].astype(str).str.strip().str.upper()
    df_imp["CENTRO"]  = df_imp["CENTRO"].astype(str).str.strip()
    df_imp["IMPUTACION"] = pd.to_numeric(df_imp["IMPUTACION"], errors="coerce")

    return df_trab, df_imp, df_cent

def geocodificar(direccion, api_key):
    url    = "https://api.openrouteservice.org/geocode/search"
    params = {"api_key": api_key, "text": direccion, "boundary.country": "ES", "size": 1}
    try:
        r    = requests.get(url, params=params, timeout=10)
        data = r.json()
        coords = data["features"][0]["geometry"]["coordinates"]
        return coords[0], coords[1]
    except Exception:
        return None, None

def calcular_km(origen, destino, api_key):
    url     = "https://api.openrouteservice.org/v2/directions/driving-car"
    headers = {"Authorization": api_key, "Content-Type": "application/json"}
    body    = {"coordinates": [list(origen), list(destino)]}
    try:
        r      = requests.post(url, json=body, headers=headers, timeout=15)
        data   = r.json()
        metros = data["routes"][0]["summary"]["distance"]
        return round(metros / 1000, 2)
    except Exception:
        return None

def construir_dir_centro(row):
    return f"{row['CALLE']} {row['NUMERO']}, {row['MUNICIPIO']}, {row['COD_POSTAL']}, España"

def guardar_en_sheets(filas):
    try:
        scope = ["https://spreadsheets.google.com/feeds",
                 "https://www.googleapis.com/auth/drive"]
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        creds      = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client     = gspread.authorize(creds)
        sheet      = client.open(GOOGLE_SHEETS_NAME).sheet1

        if not sheet.get_all_values():
            sheet.append_row([
                "FECHA", "CODIGO", "DNI", "NOMBRE",
                "DOMICILIO_USADO", "MUNICIPIO", "CP",
                "CENTRO", "IMPUTACION_%", "KM_IDA",
                "MODO_TRANSPORTE", "COMBUSTIBLE", "DOMICILIO_CORREGIDO"
            ])
        for fila in filas:
            sheet.append_row(fila)
        return True
    except Exception as e:
        st.error(f"Error al guardar en Google Sheets: {e}")
        return False

# ─────────────────────────────────────────────
# CSS CORPORATIVO
# ─────────────────────────────────────────────
def inyectar_css():
    st.markdown(f"""
    <style>
    /* Fondo general */
    .stApp {{ background-color: {COLOR_FONDO}; }}

    /* Ocultar menú hamburguesa y footer */
    #MainMenu, footer {{ visibility: hidden; }}

    /* Header corporativo */
    .header-box {{
        background: {COLOR_BLANCO};
        border-bottom: 4px solid {COLOR_TURQUESA};
        padding: 1rem 1.5rem 0.8rem 1.5rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 1.5rem;
        border-radius: 0 0 8px 8px;
    }}
    .header-titulo {{
        font-size: 1.4rem;
        font-weight: bold;
        color: {COLOR_GRIS};
        margin: 0;
    }}
    .header-subtitulo {{
        font-size: 0.9rem;
        color: #999;
        margin: 0;
    }}

    /* Cajas de información */
    .info-box {{
        background: {COLOR_BLANCO};
        border-left: 5px solid {COLOR_TURQUESA};
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin-bottom: 0.8rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}
    .aviso {{
        background: #FFF8E1;
        border-left: 5px solid #FFD54F;
        padding: 0.8rem 1rem;
        border-radius: 6px;
        font-size: 0.88rem;
        color: #6D5000;
        margin-bottom: 1rem;
    }}
    .exito {{
        background: #E8F5E9;
        border-left: 5px solid {COLOR_VERDE};
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin-bottom: 0.8rem;
    }}

    /* Tags de centros */
    .centro-tag {{
        background: {COLOR_TURQUESA};
        color: white;
        padding: 4px 12px;
        border-radius: 14px;
        font-size: 0.83rem;
        display: inline-block;
        margin: 3px;
        font-weight: 500;
    }}

    /* Secciones */
    .seccion-titulo {{
        font-size: 1.05rem;
        font-weight: bold;
        color: {COLOR_GRIS};
        border-bottom: 2px solid {COLOR_ROSA};
        padding-bottom: 4px;
        margin-top: 1.5rem;
        margin-bottom: 0.8rem;
    }}

    /* Resultado km */
    .km-box {{
        background: {COLOR_BLANCO};
        border: 1px solid {COLOR_VERDE};
        border-left: 5px solid {COLOR_VERDE};
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin-bottom: 0.6rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }}

    /* Botón principal */
    .stButton > button {{
        background-color: {COLOR_TURQUESA} !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        padding: 0.6rem 1.2rem !important;
    }}
    .stButton > button:hover {{
        background-color: {COLOR_VERDE} !important;
    }}

    /* Inputs */
    .stTextInput > div > div > input {{
        border-radius: 6px !important;
        border: 1px solid #DDD !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# APP PRINCIPAL
# ─────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="Huella de Carbono 2025 — Argia Fundazioa",
        page_icon="🌱",
        layout="centered",
    )
    inyectar_css()

    # ── HEADER CON LOGO ───────────────────────
    logo_b64 = get_logo_base64(LOGO_PATH)
    col_txt, col_logo = st.columns([3, 1])
    with col_txt:
        st.markdown(f"""
        <div>
            <p class="header-titulo">🌱 Huella de Carbono 2025</p>
            <p class="header-subtitulo">Registro de desplazamientos al trabajo</p>
        </div>
        """, unsafe_allow_html=True)
    with col_logo:
        if logo_b64:
            st.markdown(f"""
            <div style="text-align:right; padding-top:4px;">
                <img src="data:image/png;base64,{logo_b64}" style="max-height:70px; max-width:180px;">
            </div>
            """, unsafe_allow_html=True)

    st.markdown("""
    <div class="aviso">
    ⚖️ <strong>Protección de datos:</strong> Los datos introducidos se utilizarán exclusivamente
    para el cálculo de la huella de carbono de Argia Fundazioa 2025, en cumplimiento del RGPD.
    Solo el personal autorizado tiene acceso a esta información.
    </div>
    """, unsafe_allow_html=True)

    df_trab, df_imp, df_cent = cargar_datos()

    # ── PASO 1: IDENTIFICACIÓN ────────────────
    st.markdown('<div class="seccion-titulo">🔐 Identificación</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        codigo_input = st.text_input("Código de trabajador/a (4 dígitos)", placeholder="Ej: 0129", max_chars=4).strip()
    with col2:
        dni = st.text_input("DNI", placeholder="Ej: 72263087P").strip().upper()

    if not codigo_input or not dni:
        st.stop()

    # Normalizar código a 4 dígitos con ceros
    try:
        codigo = str(int(codigo_input)).zfill(4)
    except ValueError:
        st.error("❌ El código debe ser numérico.")
        st.stop()

    trabajador = df_trab[(df_trab["CODIGO"] == codigo) & (df_trab["DNI"] == dni)]

    if trabajador.empty:
        st.error("❌ Código o DNI incorrecto. Por favor comprueba los datos.")
        st.stop()

    row    = trabajador.iloc[0]
    nombre = row["NOMBRE"]

    st.markdown(f"""
    <div class="exito">
    ✅ Bienvenida/o, <strong>{nombre}</strong>
    </div>
    """, unsafe_allow_html=True)

    # ── PASO 2: DATOS DEL TRABAJADOR ─────────
    st.markdown('<div class="seccion-titulo">👤 Tus datos</div>', unsafe_allow_html=True)

    domicilio_original = row["DIREC.TRABAJ"]
    municipio_original = row["POBLACION"]
    cp_original        = str(int(row["COD. POSTAL TRAB."])) if pd.notna(row["COD. POSTAL TRAB."]) else ""

    st.markdown(f"""
    <div class="info-box">
    📍 <strong>Domicilio registrado:</strong> {domicilio_original}, {municipio_original} ({cp_original})
    </div>
    """, unsafe_allow_html=True)

    centros_trab = df_imp[df_imp["CODIGO"] == codigo][["CENTRO", "IMPUTACION"]].copy()
    centros_trab["IMPUTACION_%"] = (centros_trab["IMPUTACION"] * 100).round(1)

    st.markdown("**Centros de trabajo en 2025:**")
    tags = "".join([
        f'<span class="centro-tag">{r["CENTRO"]} ({r["IMPUTACION_%"]}%)</span>'
        for _, r in centros_trab.iterrows()
    ])
    st.markdown(tags + "<br><br>", unsafe_allow_html=True)

    # ── PASO 3: VERIFICAR DOMICILIO ───────────
    st.markdown('<div class="seccion-titulo">🏠 Domicilio habitual</div>', unsafe_allow_html=True)

    domicilio_correcto = st.radio(
        "¿Tu domicilio habitual es correcto?",
        options=["Sí, es correcto", "No, quiero corregirlo"],
        index=0,
    )

    domicilio_final  = domicilio_original
    municipio_final  = municipio_original
    cp_final         = cp_original
    domicilio_corregido = False

    if domicilio_correcto == "No, quiero corregirlo":
        st.markdown("**Introduce tu domicilio habitual actual:**")
        col1, col2, col3 = st.columns([3, 1, 1])
        with col1:
            domicilio_final = st.text_input("Calle y número", placeholder="Ej: Calle Mayor 5 2A")
        with col2:
            municipio_final = st.text_input("Municipio", placeholder="Ej: Bilbao")
        with col3:
            cp_final = st.text_input("Código postal", placeholder="Ej: 48001")

        if not domicilio_final or not municipio_final or not cp_final:
            st.warning("Por favor completa todos los campos del domicilio.")
            st.stop()
        domicilio_corregido = True

    # ── PASO 4: MODO DE TRANSPORTE POR CENTRO ─
    st.markdown('<div class="seccion-titulo">🚗 Modo de transporte</div>', unsafe_allow_html=True)
    st.caption("Indica cómo te desplazas habitualmente a cada centro de trabajo.")

    modos_por_centro = {}
    combustible_por_centro = {}
    todo_completado = True

    for _, centro_row in centros_trab.iterrows():
        centro = centro_row["CENTRO"]
        pct    = centro_row["IMPUTACION_%"]

        st.markdown(f"**🏢 {centro}** ({pct}% del tiempo)")

        modo = st.selectbox(
            f"Modo de transporte — {centro}",
            options=["— Selecciona —"] + TODOS_MODOS,
            key=f"modo_{centro}"
        )

        if modo == "— Selecciona —":
            todo_completado = False
            combustible_por_centro[centro] = ""
            modos_por_centro[centro] = ""
            continue

        modos_por_centro[centro] = modo

        if modo in VEHICULOS_CON_COMBUSTIBLE:
            combustible = st.selectbox(
                f"Tipo de combustible — {centro}",
                options=["— Selecciona —"] + TIPOS_COMBUSTIBLE,
                key=f"comb_{centro}"
            )
            if combustible == "— Selecciona —":
                todo_completado = False
                combustible_por_centro[centro] = ""
            else:
                combustible_por_centro[centro] = combustible
        else:
            combustible_por_centro[centro] = "—"

        st.markdown("---")

    if not todo_completado:
        st.stop()

    # ── PASO 5: CALCULAR KM ───────────────────
    st.markdown('<div class="seccion-titulo">📏 Cálculo de distancias</div>', unsafe_allow_html=True)

    if st.button("✅ Calcular distancias y confirmar", type="primary", use_container_width=True):

        if not ORS_API_KEY:
            st.error("⚠️ No se ha configurado la clave de OpenRouteService. Contacta con el administrador.")
            st.stop()

        direccion_origen = f"{domicilio_final}, {municipio_final}, {cp_final}, España"

        with st.spinner("Calculando distancias..."):
            lon_orig, lat_orig = geocodificar(direccion_origen, ORS_API_KEY)

            if lon_orig is None:
                st.error("❌ No se ha podido geolocalizar tu domicilio. Comprueba la dirección e inténtalo de nuevo.")
                st.stop()

            resultados   = []
            filas_sheets = []
            errores      = []

            for _, centro_row in centros_trab.iterrows():
                centro       = centro_row["CENTRO"]
                imputacion   = centro_row["IMPUTACION_%"]
                modo         = modos_por_centro.get(centro, "")
                combustible  = combustible_por_centro.get(centro, "")

                centro_data = df_cent[df_cent.iloc[:, 0].str.strip() == centro]
                if centro_data.empty:
                    errores.append(centro)
                    continue

                dir_centro           = construir_dir_centro(centro_data.iloc[0])
                lon_dest, lat_dest   = geocodificar(dir_centro, ORS_API_KEY)

                if lon_dest is None:
                    errores.append(centro)
                    continue

                km = calcular_km((lon_orig, lat_orig), (lon_dest, lat_dest), ORS_API_KEY)

                if km is None:
                    errores.append(centro)
                    continue

                resultados.append({
                    "centro": centro, "imputacion": imputacion,
                    "km": km, "modo": modo, "combustible": combustible
                })

                filas_sheets.append([
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    codigo, dni, nombre,
                    f"{domicilio_final}, {municipio_final}",
                    municipio_final, cp_final,
                    centro, imputacion, km,
                    modo, combustible,
                    "Sí" if domicilio_corregido else "No",
                ])

        if resultados:
            st.markdown("**Distancias calculadas:**")
            for r in resultados:
                comb_txt = f" — {r['combustible']}" if r['combustible'] and r['combustible'] != "—" else ""
                st.markdown(f"""
                <div class="km-box">
                🏢 <strong>{r['centro']}</strong> ({r['imputacion']}% del tiempo)<br>
                🚗 {r['modo']}{comb_txt}<br>
                📏 Distancia de ida: <strong>{r['km']} km</strong>
                </div>
                """, unsafe_allow_html=True)

            if errores:
                st.warning(f"⚠️ No se pudo calcular la distancia para: {', '.join(errores)}.")

            guardado = guardar_en_sheets(filas_sheets)
            if guardado:
                st.markdown("""
                <div class="exito">
                ✅ <strong>¡Gracias!</strong> Tus datos han sido registrados correctamente.
                </div>
                """, unsafe_allow_html=True)
                st.balloons()
            else:
                st.warning("⚠️ Los datos se han calculado pero no se han podido guardar. Contacta con el administrador.")
        else:
            st.error("❌ No se han podido calcular las distancias. Inténtalo de nuevo.")


if __name__ == "__main__":
    main()
