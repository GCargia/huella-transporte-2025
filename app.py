import streamlit as st
import pandas as pd
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import json
import os

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
FICHERO_TRABAJADORES = "Informe_trabajadores_domicilio_imputacion.XLS"
FICHERO_CENTROS = "DIRECCIONES_CENTROS.xlsx"
ORS_API_KEY = st.secrets.get("ORS_API_KEY", "")
GOOGLE_SHEETS_NAME = st.secrets.get("GOOGLE_SHEETS_NAME", "HuellaCarbonoTransporte2025")

MODOS_TRANSPORTE = [
    "Coche gasolina",
    "Coche diésel",
    "Coche híbrido",
    "Coche eléctrico",
    "Transporte público (bus / metro / tren)",
    "A pie o en bicicleta",
]

# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────
@st.cache_data
def cargar_datos():
    df_trab = pd.read_excel(FICHERO_TRABAJADORES, engine="xlrd", sheet_name="TRABAJADORES")
    df_imp = pd.read_excel(FICHERO_TRABAJADORES, engine="xlrd", sheet_name="IMPUTACIONES")
    df_centros = pd.read_excel(FICHERO_CENTROS)

    # Limpiar espacios en columnas clave
    df_trab.columns = df_trab.columns.str.strip()
    df_imp.columns = df_imp.columns.str.strip()
    df_centros.columns = df_centros.columns.str.strip()

    df_trab["CODIGO"] = df_trab["CODIGO"].astype(str).str.strip()
    df_trab["DNI"] = df_trab["DNI"].astype(str).str.strip().str.upper()
    df_imp["CODIGO"] = df_imp["CODIGO"].astype(str).str.strip()
    df_imp["CENTRO"] = df_imp["CENTRO"].astype(str).str.strip()
    df_imp["IMPUTACION"] = pd.to_numeric(df_imp["IMPUTACION"], errors="coerce")

    return df_trab, df_imp, df_centros


# ─────────────────────────────────────────────
# GEOCODIFICACIÓN Y KM
# ─────────────────────────────────────────────
def geocodificar(direccion: str, api_key: str) -> tuple:
    """Devuelve (lon, lat) para una dirección dada usando OpenRouteService."""
    url = "https://api.openrouteservice.org/geocode/search"
    params = {
        "api_key": api_key,
        "text": direccion,
        "boundary.country": "ES",
        "size": 1,
    }
    try:
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        coords = data["features"][0]["geometry"]["coordinates"]
        return coords[0], coords[1]  # lon, lat
    except Exception:
        return None, None


def calcular_km(origen: tuple, destino: tuple, api_key: str) -> float:
    """Calcula la distancia en km por carretera entre dos puntos (lon, lat)."""
    url = "https://api.openrouteservice.org/v2/directions/driving-car"
    headers = {"Authorization": api_key, "Content-Type": "application/json"}
    body = {"coordinates": [list(origen), list(destino)]}
    try:
        r = requests.post(url, json=body, headers=headers, timeout=15)
        data = r.json()
        metros = data["routes"][0]["summary"]["distance"]
        return round(metros / 1000, 2)
    except Exception:
        return None


def construir_direccion_trabajador(row) -> str:
    return f"{row['DIREC.TRABAJ']}, {row['POBLACION']}, {row['COD. POSTAL TRAB.']}, España"


def construir_direccion_centro(row) -> str:
    return f"{row['CALLE']} {row['NUMERO']}, {row['MUNICIPIO']}, {row['COD_POSTAL']}, España"


# ─────────────────────────────────────────────
# GOOGLE SHEETS
# ─────────────────────────────────────────────
def guardar_en_sheets(datos: dict):
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet = client.open(GOOGLE_SHEETS_NAME).sheet1

        # Cabecera si está vacío
        if sheet.row_count == 0 or sheet.cell(1, 1).value is None:
            sheet.append_row([
                "FECHA", "CODIGO", "DNI", "NOMBRE",
                "DOMICILIO_USADO", "MUNICIPIO", "CP",
                "CENTRO", "IMPUTACION_%", "KM_IDA",
                "MODO_TRANSPORTE", "DOMICILIO_CORREGIDO"
            ])

        for fila in datos["filas"]:
            sheet.append_row(fila)
        return True
    except Exception as e:
        st.error(f"Error al guardar en Google Sheets: {e}")
        return False


# ─────────────────────────────────────────────
# APP PRINCIPAL
# ─────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="Huella de Carbono 2025 — Argia Fundazioa",
        page_icon="🌱",
        layout="centered",
    )

    # CSS personalizado
    st.markdown("""
        <style>
        .titulo { font-size: 1.6rem; font-weight: bold; color: #1F4E79; margin-bottom: 0.2rem; }
        .subtitulo { font-size: 1rem; color: #555; margin-bottom: 1.5rem; }
        .info-box { background: #D6E4F0; padding: 1rem 1.2rem; border-radius: 8px; margin-bottom: 1rem; }
        .centro-tag { background: #1F4E79; color: white; padding: 3px 10px; border-radius: 12px;
                      font-size: 0.85rem; display: inline-block; margin: 3px; }
        .aviso { background: #FFF2CC; padding: 0.8rem 1rem; border-radius: 6px;
                 font-size: 0.9rem; color: #7B5C00; margin-bottom: 1rem; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="titulo">🌱 Huella de Carbono 2025</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitulo">Argia Fundazioa — Registro de desplazamientos al trabajo</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="aviso">
    ℹ️ <strong>Protección de datos:</strong> Los datos que introduzcas se utilizarán exclusivamente
    para el cálculo de la huella de carbono de Argia Fundazioa 2025, en cumplimiento del RGPD.
    Solo el personal autorizado tiene acceso a esta información.
    </div>
    """, unsafe_allow_html=True)

    df_trab, df_imp, df_centros = cargar_datos()

    # ── PASO 1: IDENTIFICACIÓN ────────────────
    st.markdown("### 🔐 Identificación")
    col1, col2 = st.columns(2)
    with col1:
        codigo = st.text_input("Código de trabajador/a", placeholder="Ej: 156").strip()
    with col2:
        dni = st.text_input("DNI", placeholder="Ej: 72263087P").strip().upper()

    if not codigo or not dni:
        st.stop()

    # Verificar credenciales
    trabajador = df_trab[
        (df_trab["CODIGO"] == codigo) & (df_trab["DNI"] == dni)
    ]

    if trabajador.empty:
        st.error("❌ Código o DNI incorrecto. Por favor comprueba los datos.")
        st.stop()

    row = trabajador.iloc[0]
    nombre = row["NOMBRE"]

    st.success(f"✅ Bienvenida/o, **{nombre}**")

    # ── PASO 2: DATOS DEL TRABAJADOR ─────────
    st.markdown("---")
    st.markdown("### 👤 Tus datos")

    domicilio_original = row["DIREC.TRABAJ"]
    municipio_original = row["POBLACION"]
    cp_original = str(int(row["COD. POSTAL TRAB."])) if pd.notna(row["COD. POSTAL TRAB."]) else ""

    st.markdown(f"""
    <div class="info-box">
    📍 <strong>Domicilio registrado:</strong> {domicilio_original}, {municipio_original} ({cp_original})
    </div>
    """, unsafe_allow_html=True)

    # Centros del trabajador
    centros_trab = df_imp[df_imp["CODIGO"] == codigo][["CENTRO", "IMPUTACION"]].copy()
    centros_trab["IMPUTACION_%"] = (centros_trab["IMPUTACION"] * 100).round(1)

    st.markdown("**Centros de trabajo en 2025:**")
    tags = "".join([
        f'<span class="centro-tag">{r["CENTRO"]} ({r["IMPUTACION_%"]}%)</span>'
        for _, r in centros_trab.iterrows()
    ])
    st.markdown(tags, unsafe_allow_html=True)

    # ── PASO 3: VERIFICAR DOMICILIO ───────────
    st.markdown("---")
    st.markdown("### 🏠 ¿Tu domicilio habitual es correcto?")

    domicilio_correcto = st.radio(
        f"{domicilio_original}, {municipio_original} ({cp_original})",
        options=["Sí, es correcto", "No, quiero corregirlo"],
        index=0,
    )

    domicilio_final = domicilio_original
    municipio_final = municipio_original
    cp_final = cp_original
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

    # ── PASO 4: MODO DE TRANSPORTE ────────────
    st.markdown("---")
    st.markdown("### 🚗 Modo de transporte habitual")
    st.caption("Indica el medio que usas habitualmente para desplazarte al trabajo.")

    modo = st.selectbox("¿Cómo te desplazas al trabajo?", options=["— Selecciona —"] + MODOS_TRANSPORTE)

    if modo == "— Selecciona —":
        st.stop()

    # ── PASO 5: CÁLCULO DE KM ─────────────────
    st.markdown("---")
    st.markdown("### 📏 Cálculo de distancias")

    if st.button("✅ Calcular distancias y confirmar", type="primary", use_container_width=True):

        direccion_origen = f"{domicilio_final}, {municipio_final}, {cp_final}, España"

        if not ORS_API_KEY:
            st.error("⚠️ No se ha configurado la clave de la API de OpenRouteService. Contacta con el administrador.")
            st.stop()

        with st.spinner("Calculando distancias..."):
            lon_origen, lat_origen = geocodificar(direccion_origen, ORS_API_KEY)

            if lon_origen is None:
                st.error("❌ No se ha podido geolocalizar tu domicilio. Comprueba la dirección e inténtalo de nuevo.")
                st.stop()

            resultados = []
            filas_sheets = []
            errores = []

            for _, centro_row in centros_trab.iterrows():
                centro_nombre = centro_row["CENTRO"]
                imputacion_pct = centro_row["IMPUTACION_%"]

                # Buscar dirección del centro
                centro_data = df_centros[df_centros.iloc[:, 0].str.strip() == centro_nombre]
                if centro_data.empty:
                    errores.append(centro_nombre)
                    continue

                direccion_centro = construir_direccion_centro(centro_data.iloc[0])
                lon_dest, lat_dest = geocodificar(direccion_centro, ORS_API_KEY)

                if lon_dest is None:
                    errores.append(centro_nombre)
                    continue

                km = calcular_km(
                    (lon_origen, lat_origen),
                    (lon_dest, lat_dest),
                    ORS_API_KEY
                )

                if km is None:
                    errores.append(centro_nombre)
                    continue

                resultados.append({
                    "centro": centro_nombre,
                    "imputacion": imputacion_pct,
                    "km": km,
                })

                filas_sheets.append([
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    codigo, dni, nombre,
                    f"{domicilio_final}, {municipio_final}",
                    municipio_final, cp_final,
                    centro_nombre, imputacion_pct, km,
                    modo,
                    "Sí" if domicilio_corregido else "No",
                ])

        # Mostrar resultados
        if resultados:
            st.markdown("#### Distancias calculadas:")
            for r in resultados:
                st.markdown(f"""
                <div class="info-box">
                🏢 <strong>{r['centro']}</strong> ({r['imputacion']}% del tiempo)<br>
                📏 Distancia de ida: <strong>{r['km']} km</strong>
                </div>
                """, unsafe_allow_html=True)

            if errores:
                st.warning(f"⚠️ No se pudo calcular la distancia para: {', '.join(errores)}. Contacta con el administrador.")

            # Guardar en Google Sheets
            guardado = guardar_en_sheets({"filas": filas_sheets})

            if guardado:
                st.success("✅ Tus datos han sido registrados correctamente. ¡Gracias por tu colaboración!")
                st.balloons()
            else:
                st.warning("⚠️ Los datos se han calculado pero no se han podido guardar. Contacta con el administrador.")
        else:
            st.error("❌ No se han podido calcular las distancias. Comprueba tu conexión e inténtalo de nuevo.")


if __name__ == "__main__":
    main()
