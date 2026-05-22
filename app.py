import streamlit as st
import pandas as pd
import requests
import gspread
import re
import io
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import json
import base64
import os
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
DRIVE_FILE_ID_TRAB   = "1vW-2w-KTfAkOd-QrDOg-Tg40MSjMDdpr"
DRIVE_FILE_ID_DIAS   = "13rMzp0j4I_OpZAqt5gc_gJ5Do8bjEY49"
FICHERO_CENTROS      = "DIRECCIONES_CENTROS.xlsx"
LOGO_PATH            = "LOGO_ARGIA_2026.png"
ORS_API_KEY          = st.secrets.get("ORS_API_KEY", "")
GOOGLE_SHEETS_NAME   = st.secrets.get("GOOGLE_SHEETS_NAME", "HuellaCarbonoTransporte2025")
DIAS_BASE            = 360
DIAS_LABORABLES_2025 = 226

VEHICULOS_CON_COMBUSTIBLE = ["Coche", "Furgoneta", "Motozikleta"]
TIPOS_COMBUSTIBLE         = ["Gasolina", "Diesela", "Elektrikoa", "Hibridoa"]
MODOS_SIN_COMBUSTIBLE     = ["Garraio publikoa (autobusa / metroa / trena)", "Oinez edo bizikletaz"]
TODOS_MODOS               = VEHICULOS_CON_COMBUSTIBLE + MODOS_SIN_COMBUSTIBLE

COLOR_TURQUESA = "#7BC4C4"
COLOR_ROSA     = "#F4A7B4"
COLOR_VERDE    = "#A8D5B5"
COLOR_GRIS     = "#6B6B6B"
COLOR_FONDO    = "#F9F9F9"
COLOR_BLANCO   = "#FFFFFF"

# ─────────────────────────────────────────────
# TEXTOS BILINGÜES
# ─────────────────────────────────────────────
TEXTOS = {
    "eu": {
        "titulo": "Karbono Aztarna 2025",
        "subtitulo": "Lanerako joan-etorrien erregistroa",
        "aviso_titulo": "Datuen babesa",
        "aviso_texto": "Sartutako datuak Argia Fundazioa 2025en karbono-aztarna kalkulatzeko soilik erabiliko dira, DBEO betez. Baimendutako langileek bakarrik dute informazio hori eskuratzeko aukera.",
        "identificacion": "🔐 Identifikazioa",
        "codigo": "Langilearen kodea (4 digitu)",
        "dni": "NAN",
        "error_credenciales": "❌ Kodea edo NANa ez da zuzena. Mesedez, egiaztatu datuak.",
        "error_codigo": "❌ Kodea numerikoa izan behar da.",
        "bienvenido": "✅ Ongi etorri,",
        "tus_datos": "👤 Zure datuak",
        "domicilio_registrado": "📍 Helbidea erregistratu da:",
        "centros_trabajo": "2025eko lan zentroak:",
        "domicilio_habitual": "🏠 Ohiko bizilekua",
        "domicilio_correcto": "Zure ohiko bizilekua zuzena da?",
        "si_correcto": "Bai, zuzena da",
        "no_correcto": "Ez, zuzendu nahi dut",
        "introduce_domicilio": "Sartu zure egungo ohiko helbidea:",
        "calle": "Kalea eta zenbakia",
        "municipio": "Udalerria",
        "cp": "Posta kodea",
        "aviso_domicilio": "Bete helbidearen eremu guztiak, mesedez.",
        "modo_transporte": "🚗 Garraio modua",
        "modo_caption": "Adierazi nola joaten zaren ohikoki lan zentro bakoitzera.",
        "selecciona": "— Hautatu —",
        "modo_label": "Garraio modua —",
        "combustible_label": "Erregai mota —",
        "calculo": "📏 Distantzien kalkulua",
        "boton_calcular": "✅ Distantziak kalkulatu eta baieztatu",
        "error_ors": "⚠️ OpenRouteService gakoa ez dago konfiguratuta.",
        "error_geolocalizacion": "❌ Ezin izan da zure helbidea geolokalizatu. Hautatu 'Ez, zuzendu nahi dut' eta sartu helbidea formatu honetan: Kalea Zenbakia, Udalerria, Posta kodea.",
        "distancias_calculadas": "Kalkulatutako distantziak:",
        "distancia_ida": "Joaneko distantzia:",
        "km_anuales": "Urteko KM (joan-etorri):",
        "error_centro": "⚠️ Ezin izan da distantzia kalkulatu:",
        "gracias": "✅ Eskerrik asko! Zure datuak behar bezala erregistratu dira.",
        "error_sheets": "⚠️ Datuak kalkulatu dira baina ezin izan dira gorde. Jarri harremanetan administratzailearekin.",
        "error_distancias": "❌ Ezin izan dira distantziak kalkulatu. Saiatu berriro.",
        "denboraren": "denboraren",
        "idioma": "🌐 Hizkuntza",
        "spinner": "Distantziak kalkulatzen...",
        "modos_display": {
            "Coche": "Autoa",
            "Furgoneta": "Furgoneta",
            "Motozikleta": "Motozikleta",
            "Garraio publikoa (autobusa / metroa / trena)": "Garraio publikoa",
            "Oinez edo bizikletaz": "Oinez edo bizikletaz",
        }
    },
    "es": {
        "titulo": "Huella de Carbono 2025",
        "subtitulo": "Registro de desplazamientos al trabajo",
        "aviso_titulo": "Protección de datos",
        "aviso_texto": "Los datos introducidos se utilizarán exclusivamente para el cálculo de la huella de carbono de Argia Fundazioa 2025, en cumplimiento del RGPD. Solo el personal autorizado tiene acceso a esta información.",
        "identificacion": "🔐 Identificación",
        "codigo": "Código de trabajador/a (4 dígitos)",
        "dni": "DNI",
        "error_credenciales": "❌ Código o DNI incorrecto. Por favor comprueba los datos.",
        "error_codigo": "❌ El código debe ser numérico.",
        "bienvenido": "✅ Bienvenida/o,",
        "tus_datos": "👤 Tus datos",
        "domicilio_registrado": "📍 Domicilio registrado:",
        "centros_trabajo": "Centros de trabajo en 2025:",
        "domicilio_habitual": "🏠 Domicilio habitual",
        "domicilio_correcto": "¿Tu domicilio habitual es correcto?",
        "si_correcto": "Sí, es correcto",
        "no_correcto": "No, quiero corregirlo",
        "introduce_domicilio": "Introduce tu domicilio habitual actual:",
        "calle": "Calle y número",
        "municipio": "Municipio",
        "cp": "Código postal",
        "aviso_domicilio": "Por favor completa todos los campos del domicilio.",
        "modo_transporte": "🚗 Modo de transporte",
        "modo_caption": "Indica cómo te desplazas habitualmente a cada centro de trabajo.",
        "selecciona": "— Selecciona —",
        "modo_label": "Modo de transporte —",
        "combustible_label": "Tipo de combustible —",
        "calculo": "📏 Cálculo de distancias",
        "boton_calcular": "✅ Calcular distancias y confirmar",
        "error_ors": "⚠️ No se ha configurado la clave de OpenRouteService.",
        "error_geolocalizacion": "❌ No se ha podido geolocalizar tu domicilio. Selecciona 'No, quiero corregirlo' e introduce la dirección en formato: Calle Mayor 5, Bilbao, 48001.",
        "distancias_calculadas": "Distancias calculadas:",
        "distancia_ida": "Distancia de ida:",
        "km_anuales": "KM anuales (ida y vuelta):",
        "error_centro": "⚠️ No se pudo calcular la distancia para:",
        "gracias": "✅ ¡Gracias! Tus datos han sido registrados correctamente.",
        "error_sheets": "⚠️ Los datos se han calculado pero no se han podido guardar. Contacta con el administrador.",
        "error_distancias": "❌ No se han podido calcular las distancias. Inténtalo de nuevo.",
        "denboraren": "del tiempo",
        "idioma": "🌐 Idioma",
        "spinner": "Calculando distancias...",
        "modos_display": {
            "Coche": "Coche",
            "Furgoneta": "Furgoneta",
            "Motozikleta": "Moto",
            "Garraio publikoa (autobusa / metroa / trena)": "Transporte público",
            "Oinez edo bizikletaz": "A pie o en bicicleta",
        }
    }
}

COMBUSTIBLE_DISPLAY = {
    "eu": {"Gasolina": "Gasolina", "Diesela": "Diesela", "Elektrikoa": "Elektrikoa", "Hibridoa": "Hibridoa"},
    "es": {"Gasolina": "Gasolina", "Diesela": "Diésel",  "Elektrikoa": "Eléctrico",  "Hibridoa": "Híbrido"}
}

# ─────────────────────────────────────────────
# LOGO
# ─────────────────────────────────────────────
def get_logo_base64():
    try:
        if os.path.exists(LOGO_PATH):
            with open(LOGO_PATH, "rb") as f:
                return base64.b64encode(f.read()).decode()
    except Exception:
        pass
    return None

# Cargar logo una vez al arrancar el módulo
LOGO_B64 = get_logo_base64()

# ─────────────────────────────────────────────
# GOOGLE
# ─────────────────────────────────────────────
def get_google_creds():
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]
    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    return ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

def descargar_desde_drive(file_id):
    creds   = get_google_creds()
    service = build("drive", "v3", credentials=creds)
    request = service.files().get_media(fileId=file_id)
    buffer  = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buffer.seek(0)
    return buffer

@st.cache_data(ttl=0)
def cargar_datos():
    buf_trab = descargar_desde_drive(DRIVE_FILE_ID_TRAB)
    df_trab  = pd.read_excel(buf_trab, sheet_name="TRABAJADORES")
    buf_trab.seek(0)
    df_imp   = pd.read_excel(buf_trab, sheet_name="IMPUTACIONES")

    buf_dias = descargar_desde_drive(DRIVE_FILE_ID_DIAS)
    df_dias  = pd.read_excel(buf_dias)

    df_cent  = pd.read_excel(FICHERO_CENTROS)

    for df in [df_trab, df_imp, df_dias, df_cent]:
        df.columns = df.columns.str.strip()

    df_trab["CODIGO"] = df_trab["CODIGO"].astype(int).apply(lambda x: str(x).zfill(4))
    df_imp["CODIGO"]  = df_imp["CODIGO"].astype(int).apply(lambda x: str(x).zfill(4))
    df_dias["CODIGO"] = df_dias["CODIGO"].astype(int).apply(lambda x: str(x).zfill(4))
    df_trab["DNI"]    = df_trab["DNI"].astype(str).str.strip().str.upper()
    df_imp["CENTRO"]  = df_imp["CENTRO"].astype(str).str.strip()
    df_imp["IMPUTACION"]  = pd.to_numeric(df_imp["IMPUTACION"], errors="coerce")
    df_dias["Nº DÍAS"]    = pd.to_numeric(df_dias["Nº DÍAS"], errors="coerce")
    df_dias["% JORNADA"]  = pd.to_numeric(df_dias["% JORNADA"], errors="coerce")

    def parsear_coords(val):
        try:
            partes = str(val).split(",")
            return float(partes[0].strip()), float(partes[1].strip())
        except Exception:
            return None, None

    df_cent["LAT"], df_cent["LON"] = zip(*df_cent["COORDENADAS"].apply(parsear_coords))
    return df_trab, df_imp, df_cent, df_dias

# ─────────────────────────────────────────────
# GEOCODIFICACIÓN Y RUTAS
# ─────────────────────────────────────────────
def limpiar_direccion(direccion):
    reemplazos = [
        (r"^CL\b","Calle"),(r"^C/\b","Calle"),(r"^C\b","Calle"),
        (r"^AV\b","Avenida"),(r"^AVD\b","Avenida"),(r"^AVDA\b","Avenida"),
        (r"^PZ\b","Plaza"),(r"^PL\b","Plaza"),
        (r"^PS\b","Paseo"),(r"^PSO\b","Paseo"),
        (r"^BO\b","Barrio"),(r"^Bº\b","Barrio"),
        (r"^URB\b","Urbanización"),(r"^CTRA\b","Carretera"),(r"^CR\b","Carretera"),
    ]
    d = direccion.strip()
    for patron, reemplazo in reemplazos:
        d = re.sub(patron, reemplazo, d, flags=re.IGNORECASE)
    d = re.sub(r"(\d+)\s+\d+\s*[A-Za-z]?\s*$", r"\1", d)
    return d.strip()

def geocodificar_nominatim(domicilio, municipio, cp):
    domicilio_limpio = limpiar_direccion(domicilio)
    intentos = [
        f"{domicilio_limpio}, {municipio}, {cp}, España",
        f"{domicilio_limpio}, {municipio}, España",
        f"{domicilio_limpio}, {cp}, España",
        f"{municipio}, {cp}, España",
        f"{cp}, España",
    ]
    headers = {"User-Agent": "ArgiaCarbonApp/1.0"}
    for intento in intentos:
        try:
            r    = requests.get("https://nominatim.openstreetmap.org/search",
                                params={"q": intento, "format": "json",
                                        "limit": 1, "countrycodes": "es"},
                                headers=headers, timeout=10)
            data = r.json()
            if data:
                lat = float(data[0]["lat"])
                lon = float(data[0]["lon"])
                if 41.0 <= lat <= 44.5 and -5.0 <= lon <= 2.0:
                    return lon, lat
        except Exception:
            continue
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

# ─────────────────────────────────────────────
# SHEETS
# ─────────────────────────────────────────────
def guardar_en_sheets(filas):
    try:
        creds  = get_google_creds()
        client = gspread.authorize(creds)
        sheet  = client.open(GOOGLE_SHEETS_NAME).sheet1
        if not sheet.get_all_values():
            sheet.append_row([
                "FECHA","CODIGO","DNI","NOMBRE","DOMICILIO_USADO","MUNICIPIO","CP",
                "CENTRO","IMPUTACION_%","KM_IDA","KM_IDA_VUELTA_ANUALES",
                "MODO_TRANSPORTE","COMBUSTIBLE","DOMICILIO_CORREGIDO"
            ])
        for fila in filas:
            sheet.append_row(fila)
        return client
    except Exception as e:
        st.error(f"Error al guardar en Google Sheets: {e}")
        return None

def actualizar_sheet_calculos(client, resultados):
    try:
        try:
            sheet_calc = client.open(GOOGLE_SHEETS_NAME).worksheet("CALCULOS")
        except Exception:
            sheet_calc = client.open(GOOGLE_SHEETS_NAME).add_worksheet(
                title="CALCULOS", rows=100, cols=20)

        modos_cols = [
            "Coche - Gasolina","Coche - Diésel","Coche - Eléctrico","Coche - Híbrido",
            "Furgoneta - Gasolina","Furgoneta - Diésel","Furgoneta - Eléctrico","Furgoneta - Híbrido",
            "Moto - Gasolina","Moto - Diésel","Moto - Eléctrico","Moto - Híbrido",
            "Transporte público","A pie o en bicicleta","TOTAL"
        ]
        cabecera = ["CENTRO"] + modos_cols

        data = sheet_calc.get_all_values()
        if not data or data[0] != cabecera:
            sheet_calc.clear()
            sheet_calc.append_row(cabecera)
            data = [cabecera]

        centros_idx = {row[0]: i + 2 for i, row in enumerate(data[1:])}

        for r in resultados:
            centro     = r["centro"]
            km_anuales = r["km_ida_vuelta_anuales"]
            modo       = r["modo"]
            combustible = r["combustible"]

            if combustible and combustible not in ["—", ""]:
                col_key = f"{modo} - {combustible}"
            else:
                col_key = modo

            if col_key not in modos_cols:
                col_key = "TOTAL"

            col_idx   = modos_cols.index(col_key) + 2
            col_total = len(modos_cols) + 1

            if centro not in centros_idx:
                sheet_calc.append_row([centro] + [0] * len(modos_cols))
                data = sheet_calc.get_all_values()
                centros_idx = {row[0]: i + 2 for i, row in enumerate(data[1:])}

            fila_idx = centros_idx[centro]
            val_actual   = float(sheet_calc.cell(fila_idx, col_idx).value or 0)
            val_total    = float(sheet_calc.cell(fila_idx, col_total).value or 0)
            sheet_calc.update_cell(fila_idx, col_idx,   round(val_actual + km_anuales, 2))
            sheet_calc.update_cell(fila_idx, col_total, round(val_total  + km_anuales, 2))

    except Exception as e:
        st.warning(f"No se pudo actualizar la hoja de cálculos: {e}")

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
def inyectar_css():
    st.markdown(f"""
    <style>
    .stApp {{ background-color: {COLOR_FONDO}; }}
    #MainMenu, footer {{ visibility: hidden; }}
    .info-box {{ background:{COLOR_BLANCO}; border-left:5px solid {COLOR_TURQUESA};
        padding:0.8rem 1rem; border-radius:6px; margin-bottom:0.8rem;
        box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
    .aviso {{ background:#FFF8E1; border-left:5px solid #FFD54F;
        padding:0.8rem 1rem; border-radius:6px; font-size:0.88rem;
        color:#6D5000; margin-bottom:1rem; }}
    .exito {{ background:#E8F5E9; border-left:5px solid {COLOR_VERDE};
        padding:0.8rem 1rem; border-radius:6px; margin-bottom:0.8rem; }}
    .centro-tag {{ background:{COLOR_TURQUESA}; color:white; padding:4px 12px;
        border-radius:14px; font-size:0.83rem; display:inline-block;
        margin:3px; font-weight:500; }}
    .seccion-titulo {{ font-size:1.05rem; font-weight:bold; color:{COLOR_GRIS};
        border-bottom:2px solid {COLOR_ROSA}; padding-bottom:4px;
        margin-top:1.5rem; margin-bottom:0.8rem; }}
    .km-box {{ background:{COLOR_BLANCO}; border:1px solid {COLOR_VERDE};
        border-left:5px solid {COLOR_VERDE}; padding:0.8rem 1rem;
        border-radius:6px; margin-bottom:0.6rem;
        box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
    .stButton > button {{ background-color:{COLOR_TURQUESA} !important;
        color:white !important; border:none !important; border-radius:8px !important;
        font-weight:bold !important; padding:0.6rem 1.2rem !important; }}
    .stButton > button:hover {{ background-color:{COLOR_VERDE} !important; }}
    </style>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# APP PRINCIPAL
# ─────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="Karbono Aztarna 2025 — Argia Fundazioa",
        page_icon="🌱", layout="centered",
    )
    inyectar_css()

    # ── SELECTOR DE IDIOMA INICIAL ────────────
    if "idioma" not in st.session_state:
        st.session_state["idioma"] = None

    if st.session_state["idioma"] is None:
        if LOGO_B64:
            st.markdown(f"""
            <div style="text-align:center; padding:2rem 0 1rem 0;">
                <img src="data:image/png;base64,{LOGO_B64}"
                     style="max-height:100px; max-width:280px;">
            </div>
            """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("""
            <div style="text-align:center; margin-bottom:2rem;">
                <p style="font-size:1.1rem; color:#6B6B6B; font-weight:500;">
                    Hautatu hizkuntza / Selecciona idioma</p>
            </div>
            """, unsafe_allow_html=True)
            col_eu, col_es = st.columns(2)
            with col_eu:
                if st.button("🇪🇺  Euskera", use_container_width=True, type="primary"):
                    st.session_state["idioma"] = "eu"
                    st.rerun()
            with col_es:
                if st.button("🇪🇸  Castellano", use_container_width=True):
                    st.session_state["idioma"] = "es"
                    st.rerun()
        st.stop()

    idioma = st.session_state["idioma"]
    T      = TEXTOS[idioma]

    # ── SELECTOR IDIOMA ARRIBA DERECHA ────────
    col_h1, col_h2 = st.columns([4, 1])
    with col_h2:
        opciones = {"eu": "🇪🇺 Euskera", "es": "🇪🇸 Castellano"}
        nuevo = st.selectbox(T["idioma"], options=list(opciones.keys()),
                             format_func=lambda x: opciones[x],
                             index=0 if idioma == "eu" else 1,
                             label_visibility="collapsed", key="sel_idioma")
        if nuevo != idioma:
            st.session_state["idioma"] = nuevo
            st.rerun()

    # ── HEADER ────────────────────────────────
    col_txt, col_logo = st.columns([3, 1])
    with col_txt:
        st.markdown(f"""
        <div>
            <p style="font-size:1.4rem;font-weight:bold;color:{COLOR_GRIS};margin:0;">
                🌱 {T['titulo']}</p>
            <p style="font-size:0.9rem;color:#999;margin:0;">{T['subtitulo']}</p>
        </div>
        """, unsafe_allow_html=True)
    with col_logo:
        if LOGO_B64:
            st.markdown(f"""
            <div style="text-align:right;padding-top:4px;">
                <img src="data:image/png;base64,{LOGO_B64}"
                     style="max-height:70px;max-width:180px;">
            </div>
            """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="aviso">
    ⚖️ <strong>{T['aviso_titulo']}:</strong> {T['aviso_texto']}
    </div>
    """, unsafe_allow_html=True)

    df_trab, df_imp, df_cent, df_dias = cargar_datos()

    # ── PASO 1: IDENTIFICACIÓN ────────────────
    st.markdown(f'<div class="seccion-titulo">{T["identificacion"]}</div>',
                unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        codigo_input = st.text_input(T["codigo"], placeholder="Ej: 0129",
                                     max_chars=4).strip()
    with col2:
        dni = st.text_input(T["dni"], placeholder="Ej: 72263087P").strip().upper()

    if not codigo_input or not dni:
        st.stop()

    try:
        codigo = str(int(codigo_input)).zfill(4)
    except ValueError:
        st.error(T["error_codigo"])
        st.stop()

    trabajador = df_trab[(df_trab["CODIGO"] == codigo) & (df_trab["DNI"] == dni)]
    if trabajador.empty:
        st.error(T["error_credenciales"])
        st.stop()

    row   = trabajador.iloc[0]
    nombre = row["NOMBRE"]

    dias_row    = df_dias[df_dias["CODIGO"] == codigo]
    dias_trab   = float(dias_row.iloc[0]["Nº DÍAS"])   if not dias_row.empty else DIAS_BASE
    pct_jornada = float(dias_row.iloc[0]["% JORNADA"]) if not dias_row.empty else 1.0

    st.markdown(f'<div class="exito">{T["bienvenido"]} <strong>{nombre}</strong></div>',
                unsafe_allow_html=True)

    # ── PASO 2: DATOS ─────────────────────────
    st.markdown(f'<div class="seccion-titulo">{T["tus_datos"]}</div>',
                unsafe_allow_html=True)

    domicilio_original = row["DIREC.TRABAJ"]
    municipio_original = row["POBLACION"]
    cp_original = str(int(row["COD. POSTAL TRAB."])) if pd.notna(
        row["COD. POSTAL TRAB."]) else ""

    st.markdown(f"""
    <div class="info-box">
    {T['domicilio_registrado']} {domicilio_original}, {municipio_original} ({cp_original})
    </div>
    """, unsafe_allow_html=True)

    centros_trab = df_imp[df_imp["CODIGO"] == codigo][["CENTRO","IMPUTACION"]].copy()
    centros_trab["IMPUTACION_%"] = (centros_trab["IMPUTACION"] * 100).round(1)

    st.markdown(f"**{T['centros_trabajo']}**")
    tags = "".join([
        f'<span class="centro-tag">{r["CENTRO"]} ({r["IMPUTACION_%"]}%)</span>'
        for _, r in centros_trab.iterrows()
    ])
    st.markdown(tags + "<br><br>", unsafe_allow_html=True)

    # ── PASO 3: DOMICILIO ─────────────────────
    st.markdown(f'<div class="seccion-titulo">{T["domicilio_habitual"]}</div>',
                unsafe_allow_html=True)

    dom_correcto = st.radio(T["domicilio_correcto"],
                            options=[T["si_correcto"], T["no_correcto"]], index=0)

    domicilio_final = domicilio_original
    municipio_final = municipio_original
    cp_final        = cp_original
    dom_corregido   = False

    if dom_correcto == T["no_correcto"]:
        st.markdown(f"**{T['introduce_domicilio']}**")
        c1, c2, c3 = st.columns([3, 1, 1])
        with c1: domicilio_final = st.text_input(T["calle"], placeholder="Ej: Calle Mayor 5 2A")
        with c2: municipio_final = st.text_input(T["municipio"], placeholder="Ej: Bilbao")
        with c3: cp_final        = st.text_input(T["cp"], placeholder="Ej: 48001")
        if not domicilio_final or not municipio_final or not cp_final:
            st.warning(T["aviso_domicilio"])
            st.stop()
        dom_corregido = True

    # ── PASO 4: TRANSPORTE ────────────────────
    st.markdown(f'<div class="seccion-titulo">{T["modo_transporte"]}</div>',
                unsafe_allow_html=True)
    st.caption(T["modo_caption"])

    modos_por_centro       = {}
    combustible_por_centro = {}
    todo_completado        = True

    for _, cr in centros_trab.iterrows():
        centro = cr["CENTRO"]
        pct    = cr["IMPUTACION_%"]
        st.markdown(f"**🏢 {centro}** ({pct}% {T['denboraren']})")

        modo = st.selectbox(
            f"{T['modo_label']} {centro}",
            options=[T["selecciona"]] + TODOS_MODOS,
            format_func=lambda x: T["modos_display"].get(x, x) if x != T["selecciona"] else x,
            key=f"modo_{centro}"
        )

        if modo == T["selecciona"]:
            todo_completado = False
            modos_por_centro[centro] = ""
            combustible_por_centro[centro] = ""
            continue

        modos_por_centro[centro] = modo

        if modo in VEHICULOS_CON_COMBUSTIBLE:
            comb = st.selectbox(
                f"{T['combustible_label']} {centro}",
                options=[T["selecciona"]] + TIPOS_COMBUSTIBLE,
                format_func=lambda x: COMBUSTIBLE_DISPLAY[idioma].get(x, x) if x != T["selecciona"] else x,
                key=f"comb_{centro}"
            )
            if comb == T["selecciona"]:
                todo_completado = False
                combustible_por_centro[centro] = ""
            else:
                combustible_por_centro[centro] = comb
        else:
            combustible_por_centro[centro] = "—"

        st.markdown("---")

    if not todo_completado:
        st.stop()

    # ── PASO 5: CALCULAR ──────────────────────
    st.markdown(f'<div class="seccion-titulo">{T["calculo"]}</div>',
                unsafe_allow_html=True)

    if st.button(T["boton_calcular"], type="primary", use_container_width=True):
        if not ORS_API_KEY:
            st.error(T["error_ors"])
            st.stop()

        with st.spinner(T["spinner"]):
            lon_orig, lat_orig = geocodificar_nominatim(
                domicilio_final, municipio_final, cp_final)

            if lon_orig is None:
                st.error(T["error_geolocalizacion"])
                st.stop()

            resultados   = []
            filas_sheets = []
            errores      = []

            for _, cr in centros_trab.iterrows():
                centro      = cr["CENTRO"]
                imputacion  = cr["IMPUTACION_%"]
                modo        = modos_por_centro.get(centro, "")
                combustible = combustible_por_centro.get(centro, "")

                cd = df_cent[df_cent["CENTRO"].str.strip() == centro]
                if cd.empty or pd.isna(cd.iloc[0]["LAT"]):
                    errores.append(centro)
                    continue

                km = calcular_km((lon_orig, lat_orig),
                                 (cd.iloc[0]["LON"], cd.iloc[0]["LAT"]),
                                 ORS_API_KEY)
                if km is None:
                    errores.append(centro)
                    continue

                km_anuales   = round(km * 2 * (dias_trab / DIAS_BASE * DIAS_LABORABLES_2025) * pct_jornada, 2)
                modo_display = T["modos_display"].get(modo, modo)
                comb_display = COMBUSTIBLE_DISPLAY[idioma].get(combustible, combustible)
                # Versión en español para la hoja CALCULOS
                modo_es = TEXTOS["es"]["modos_display"].get(modo, modo)
                comb_es = COMBUSTIBLE_DISPLAY["es"].get(combustible, combustible)

                resultados.append({
                    "centro": centro, "imputacion": imputacion,
                    "km": km, "km_ida_vuelta_anuales": km_anuales,
                    "modo": modo_es, "combustible": comb_es,
                    "modo_display": modo_display, "comb_display": comb_display,
                })

                filas_sheets.append([
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    codigo, dni, nombre,
                    f"{domicilio_final}, {municipio_final}",
                    municipio_final, cp_final,
                    centro, imputacion, km, km_anuales,
                    modo_display, comb_display,
                    ("Bai" if dom_corregido else "Ez") if idioma == "eu"
                    else ("Sí" if dom_corregido else "No"),
                ])

        if resultados:
            st.markdown(f"**{T['distancias_calculadas']}**")
            for r in resultados:
                comb_txt = f" — {r['comb_display']}" if r['combustible'] not in ["—",""] else ""
                st.markdown(f"""
                <div class="km-box">
                🏢 <strong>{r['centro']}</strong> ({r['imputacion']}% {T['denboraren']})<br>
                🚗 {r['modo_display']}{comb_txt}<br>
                📏 {T['distancia_ida']} <strong>{r['km']} km</strong><br>
                📅 {T['km_anuales']} <strong>{r['km_ida_vuelta_anuales']} km</strong>
                </div>
                """, unsafe_allow_html=True)

            if errores:
                st.warning(f"{T['error_centro']} {', '.join(errores)}")

            client = guardar_en_sheets(filas_sheets)
            if client:
                actualizar_sheet_calculos(client, resultados)
                st.markdown(f'<div class="exito">{T["gracias"]}</div>',
                            unsafe_allow_html=True)
                st.balloons()
            else:
                st.warning(T["error_sheets"])
        else:
            st.error(T["error_distancias"])


if __name__ == "__main__":
    main()
