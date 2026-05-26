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
        "aviso_2025": "⚠️ Tresna hau 2025. urtean Argia Fundazioan lan egin duten pertsonentzat soilik da. Gogoratu sartu beharreko datuak 2025ean gertatu direnak direla.",
        "aviso_titulo": "Datuen babesa",
        "aviso_texto": "Sartutako datuak Argia Fundazioa 2025en karbono-aztarna kalkulatzeko soilik erabiliko dira, DBEO betez. Baimendutako langileek bakarrik dute informazio hori eskuratzeko aukera.",
        "identificacion": "🔐 Identifikazioa",
        "correo": "Posta elektronikoa korporatiboa",
        "dni": "NAN",
        "error_credenciales": "❌ Posta elektronikoa edo NANa ez da zuzena. Mesedez, egiaztatu datuak.",
        "error_codigo": "❌ Sartu zure posta elektroniko korporatiboa.",
        "bienvenido": "✅ Ongi etorri,",
        "aviso_reenvio": "⚠️ Datuak aurretik bidali dituzu ({}. bidalketa). Zuzendu eta berriro bidal ditzakezu, baina azken bidalketako datuak bakarrik hartuko dira kontuan.",
        "tus_datos": "👤 Zure datuak",
        "domicilio_registrado": "📍 Helbidea erregistratu da:",
        "centros_trabajo": "2025eko lan zentroak:",
        "domicilio_habitual": "🏠 Ohiko bizilekua",
        "anadir_domicilio": "➕ Bigarren helbidea gehitu (2025ean bizilekuz aldatu bazara)",
        "eliminar_domicilio": "✕ Bigarren helbidea kendu",
        "domicilio_1": "1. helbidea",
        "domicilio_2": "2. helbidea",
        "pct_domicilio": "Denbora %",
        "pct_domicilio_tooltip": "Estimazioa: 2025ean zenbat denbora eman zenuen helbide honetan. Adibidez, 6 hilabete bakoitzean egon bazinen, %50/%50 jar dezakezu.",
        "error_pct_domicilio": "⚠️ Helbideen portzentajeen batura 100% izan behar da.",
        "domicilio_correcto": "Zure ohiko bizilekua zuzena da?",
        "si_correcto": "Bai, zuzena da",
        "no_correcto": "Ez, zuzendu nahi dut",
        "introduce_domicilio": "Sartu zure egungo ohiko helbidea:",
        "calle": "Kalea eta zenbakia",
        "municipio": "Udalerria",
        "cp": "Posta kodea",
        "aviso_domicilio": "Bete helbidearen eremu guztiak, mesedez.",
        "modo_transporte": "🚗 Garraio modua",
        "modo_caption": "Adierazi nola joaten zaren ohikoki lan zentro bakoitzera. Bi garraio mota erabil baditzazke, gehitu bigarren bat.",
        "selecciona": "— Hautatu —",
        "modo_label": "Garraio modua",
        "combustible_label": "Erregai mota",
        "pct_uso": "Erabilera %",
        "pct_tooltip": "Erabileraren estimazioa: normalean autoz joaten bazara baina noizean behin garraio publikoa erabiltzen baduzu, %80 autoa eta %20 garraio publikoa jar dezakezu. Ez da beharrezkoa zehatza izatea; hurbilketa bat nahikoa da.",
        "anadir_modo": "➕ Beste garraio mota bat gehitu",
        "eliminar_modo": "✕ Kendu",
        "error_pct": "⚠️ Garraio moduen portzentajeen batura 100% izan behar da.",
        "calculo": "📏 Distantzien kalkulua",
        "boton_calcular": "✅ Distantziak kalkulatu eta baieztatu",
        "error_ors": "⚠️ OpenRouteService gakoa ez dago konfiguratuta.",
        "error_geolocalizacion": "❌ Ezin izan da zure helbidea geolokalizatu. Hautatu 'Ez, zuzendu nahi dut' eta sartu helbidea formatu honetan: Kalea Zenbakia, Udalerria, Posta kodea.",
        "distancias_calculadas": "Kalkulatutako distantziak:",
        "distancia_ida": "Joaneko distantzia:",
        "km_anuales": "Urteko KM (joan-etorri):",
        "error_centro": "⚠️ Ezin izan da distantzia kalkulatu:",
        "gracias": "✅ Eskerrik asko! Zure datuak behar bezala erregistratu dira.",
        "piloto_titulo": "📋 Pilotoaren balorazioa",
        "piloto_subtitulo": "Zure iritzia lagungarria izango da tresna hobetzeko. Minutu bat baino gutxiago irauten du.",
        "piloto_p1": "1. Nola baloratuko zenuke tresna erabiltzeko erreztasuna?",
        "piloto_p1_ops": ["— Hautatu —", "⭐ Oso zaila", "⭐⭐ Zaila", "⭐⭐⭐ Normala", "⭐⭐⭐⭐ Erraza", "⭐⭐⭐⭐⭐ Oso erraza"],
        "piloto_p2": "2. Zentroen inputazioaren datuak ondo agertzen ziren sartzean? (Ez da helbideari buruzkoa, hori aplikazioan zuzendu daiteke)",
        "piloto_p2_ops": ["— Hautatu —", "✅ Bai, dena zuzen", "⚠️ Akatsen bat zegoen", "❌ Datuak okerrak ziren"],
        "piloto_p3": "3. Ondo ulertu al zenuen tresnak zer eskatzen zizun?",
        "piloto_p3_ops": ["— Hautatu —", "Bai", "Gutxi gorabehera", "Ez"],
        "piloto_p4": "4. Iruzkinak edo iradokizunak (aukerakoa)",
        "piloto_boton": "Balorazioa bidali",
        "piloto_gracias": "✅ Eskerrik asko zure balorazioagatik!",
        "piloto_error": "⚠️ Balorazioa ezin izan da gorde. Saiatu berriro.",
        "piloto_incompleto": "⚠️ Mesedez erantzun 1, 2 eta 3 galderak.",
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
        "aviso_2025": "⚠️ Esta herramienta es únicamente para las personas que trabajaron en Argia Fundazioa durante el año 2025. Recuerda que los datos que debes introducir son de lo ocurrido en 2025.",
        "aviso_titulo": "Protección de datos",
        "aviso_texto": "Los datos introducidos se utilizarán exclusivamente para el cálculo de la huella de carbono de Argia Fundazioa 2025, en cumplimiento del RGPD. Solo el personal autorizado tiene acceso a esta información.",
        "identificacion": "🔐 Identificación",
        "correo": "Correo corporativo",
        "dni": "DNI",
        "error_credenciales": "❌ Correo o DNI incorrecto. Por favor comprueba los datos.",
        "error_codigo": "❌ Introduce tu correo corporativo.",
        "bienvenido": "✅ Bienvenida/o,",
        "aviso_reenvio": "⚠️ Ya enviaste tus datos anteriormente (envío nº {}). Puedes corregirlos y reenviar, pero solo se tendrá en cuenta el último envío.",
        "tus_datos": "👤 Tus datos",
        "domicilio_registrado": "📍 Domicilio registrado:",
        "centros_trabajo": "Centros de trabajo en 2025:",
        "domicilio_habitual": "🏠 Domicilio habitual",
        "anadir_domicilio": "➕ Añadir segundo domicilio (si te mudaste en 2025)",
        "eliminar_domicilio": "✕ Eliminar segundo domicilio",
        "domicilio_1": "1er domicilio",
        "domicilio_2": "2º domicilio",
        "pct_domicilio": "% de tiempo",
        "pct_domicilio_tooltip": "Estimación: qué porcentaje del tiempo de 2025 viviste en este domicilio. Por ejemplo, si estuviste 6 meses en cada uno, pon 50%/50%.",
        "error_pct_domicilio": "⚠️ Los porcentajes de los domicilios deben sumar 100%.",
        "domicilio_correcto": "¿Tu domicilio habitual es correcto?",
        "si_correcto": "Sí, es correcto",
        "no_correcto": "No, quiero corregirlo",
        "introduce_domicilio": "Introduce tu domicilio habitual actual:",
        "calle": "Calle y número",
        "municipio": "Municipio",
        "cp": "Código postal",
        "aviso_domicilio": "Por favor completa todos los campos del domicilio.",
        "modo_transporte": "🚗 Modo de transporte",
        "modo_caption": "Indica cómo te desplazas habitualmente a cada centro de trabajo. Si usas más de un medio de transporte, puedes añadir un segundo.",
        "selecciona": "— Selecciona —",
        "modo_label": "Modo de transporte",
        "combustible_label": "Tipo de combustible",
        "pct_uso": "% de uso",
        "pct_tooltip": "Estimación del uso: si habitualmente vas en coche pero ocasionalmente usas el transporte público, puedes poner 80% coche y 20% transporte público. No hace falta que sea exacto, una aproximación es suficiente.",
        "anadir_modo": "➕ Añadir otro modo de transporte",
        "eliminar_modo": "✕ Eliminar",
        "error_pct": "⚠️ Los porcentajes de los modos de transporte deben sumar 100%.",
        "calculo": "📏 Cálculo de distancias",
        "boton_calcular": "✅ Calcular distancias y confirmar",
        "error_ors": "⚠️ No se ha configurado la clave de OpenRouteService.",
        "error_geolocalizacion": "❌ No se ha podido geolocalizar tu domicilio. Selecciona 'No, quiero corregirlo' e introduce la dirección en formato: Calle Mayor 5, Bilbao, 48001.",
        "distancias_calculadas": "Distancias calculadas:",
        "distancia_ida": "Distancia de ida:",
        "km_anuales": "KM anuales (ida y vuelta):",
        "error_centro": "⚠️ No se pudo calcular la distancia para:",
        "gracias": "✅ ¡Gracias! Tus datos han sido registrados correctamente.",
        "piloto_titulo": "📋 Valoración del piloto",
        "piloto_subtitulo": "Tu opinión nos ayudará a mejorar la herramienta. Menos de un minuto.",
        "piloto_p1": "1. ¿Cómo valorarías la facilidad de uso de la herramienta?",
        "piloto_p1_ops": ["— Selecciona —", "⭐ Muy difícil", "⭐⭐ Difícil", "⭐⭐⭐ Normal", "⭐⭐⭐⭐ Fácil", "⭐⭐⭐⭐⭐ Muy fácil"],
        "piloto_p2": "2. ¿Los datos de imputación de centros aparecían correctamente al entrar? (No hace referencia al domicilio, que puede corregirse en la app)",
        "piloto_p2_ops": ["— Selecciona —", "✅ Sí, todo correcto", "⚠️ Había algún error", "❌ Los datos eran incorrectos"],
        "piloto_p3": "3. ¿Entendiste bien qué te pedía la herramienta?",
        "piloto_p3_ops": ["— Selecciona —", "Sí", "Más o menos", "No"],
        "piloto_p4": "4. Comentarios o sugerencias (opcional)",
        "piloto_boton": "Enviar valoración",
        "piloto_gracias": "✅ ¡Gracias por tu valoración!",
        "piloto_error": "⚠️ No se ha podido guardar la valoración. Inténtalo de nuevo.",
        "piloto_incompleto": "⚠️ Por favor responde las preguntas 1, 2 y 3.",
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

    df_trab["CODIGO"]  = df_trab["CODIGO"].astype(int).apply(lambda x: str(x).zfill(4))
    df_imp["CODIGO"]   = df_imp["CODIGO"].astype(int).apply(lambda x: str(x).zfill(4))
    df_dias["CODIGO"]  = df_dias["CODIGO"].astype(int).apply(lambda x: str(x).zfill(4))
    df_trab["DNI"]     = df_trab["DNI"].astype(str).str.strip().str.upper()
    df_trab["CORREO"]  = df_trab["CORREO"].astype(str).str.strip().str.lower()
    df_imp["CENTRO"]   = df_imp["CENTRO"].astype(str).str.strip()
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
# COMPROBACIÓN DE ENVÍOS PREVIOS
# ─────────────────────────────────────────────
def obtener_envio_anterior(codigo):
    """
    Busca en el Sheet si el empleado (por CODIGO) ya ha enviado datos.
    Devuelve el número de envío más alto encontrado, o 0 si no hay envíos previos.
    """
    try:
        creds  = get_google_creds()
        client = gspread.authorize(creds)
        sheet  = client.open(GOOGLE_SHEETS_NAME).sheet1
        datos  = sheet.get_all_values()
        if len(datos) <= 1:
            return 0
        cabecera = datos[0]
        try:
            idx_codigo  = cabecera.index("CODIGO")
            idx_envio   = cabecera.index("ENVIO_NUM")
        except ValueError:
            # La columna ENVIO_NUM aún no existe (datos históricos sin ella)
            idx_codigo = cabecera.index("CODIGO")
            filas_empleado = [f for f in datos[1:] if f[idx_codigo] == codigo]
            return 1 if filas_empleado else 0

        filas_empleado = [f for f in datos[1:] if f[idx_codigo] == codigo]
        if not filas_empleado:
            return 0
        envios = []
        for f in filas_empleado:
            try:
                envios.append(int(f[idx_envio]))
            except (ValueError, IndexError):
                envios.append(1)
        return max(envios)
    except Exception:
        return 0

# ─────────────────────────────────────────────
# GEOCODIFICACIÓN Y RUTAS
# ─────────────────────────────────────────────
def limpiar_direccion(direccion):
    """
    Limpia y normaliza la dirección para mejorar la geocodificación con Nominatim.
    - Expande abreviaturas de tipo de vía
    - Elimina información de piso/puerta (todo lo que va tras el número de portal)
    - Colapsa espacios múltiples
    """
    reemplazos = [
        (r"^CL\b",   "Calle"),
        (r"^C/\b",   "Calle"),
        (r"^AV\b",   "Avenida"),
        (r"^AVD\b",  "Avenida"),
        (r"^AVDA\b", "Avenida"),
        (r"^PZ\b",   "Plaza"),
        (r"^PL\b",   "Plaza"),
        (r"^PS\b",   "Paseo"),
        (r"^PSO\b",  "Paseo"),
        (r"^BO\b",   "Barrio"),
        (r"^Bº\b",   "Barrio"),
        (r"^GR\b",   "Grupo"),
        (r"^PA\b",   "Pasaje"),
        (r"^CM\b",   "Camino"),
        (r"^CT\b",   "Carretera"),
        (r"^CTRA\b", "Carretera"),
        (r"^CR\b",   "Carretera"),
        (r"^URB\.",  "Urbanización"),
        (r"^URB\b",  "Urbanización"),
    ]

    d = direccion.strip()

    # Expandir abreviaturas de tipo de vía
    for patron, reemplazo in reemplazos:
        d = re.sub(patron, reemplazo, d, flags=re.IGNORECASE)

    # Colapsar espacios múltiples
    d = re.sub(r"\s{2,}", " ", d)

    # Eliminar piso/puerta: todo lo que viene después del número de portal.
    # El número de portal puede ir seguido opcionalmente de una letra pegada (ej: "21B", "3B").
    # Patrón: número + letra_opcional + (espacio + más_texto_que_no_queremos)
    # Ejemplos que debe limpiar:
    #   "Calle Languileria 64 2 B"     → "Calle Languileria 64"
    #   "Calle Porton Urarte 19 2 DCHA B" → "Calle Porton Urarte 19"
    #   "Calle Iturribide 59 4 IZ-IZ"  → "Calle Iturribide 59"
    #   "Calle Peña Santa Marina 2B"   → "Calle Peña Santa Marina 2B"  (letra pegada: ok)
    #   "Calle La Dinamita 21B BAJO B" → "Calle La Dinamita 21B"
    d = re.sub(r"(\d+[A-Za-z]?)\s+\d.*$", r"\1", d)
    d = re.sub(
        r"(\d+[A-Za-z]?)\s+(BAJO|ÁTICO|ATICO|PRINCIPAL|PPAL|DCHA\.?|IZDA\.?|IZQ\.?|IZD|DC|IZ|CTRO|CENTRO|BLOQUE|BLQ|BAJO)\b.*$",
        r"\1", d, flags=re.IGNORECASE
    )

    return d.strip()


def geocodificar_nominatim(domicilio, municipio, cp):
    """
    Geocodifica una dirección usando Nominatim con múltiples intentos progresivamente
    más permisivos. Valida que el resultado esté en el País Vasco / norte de España.

    En Bizkaia/Gipuzkoa, OpenStreetMap indexa muchas calles en euskera (ej: "Los Tilos Kalea"
    en vez de "Calle Los Tilos"), por lo que se añaden intentos sin el prefijo de tipo de vía
    y con el nombre de calle solo + municipio/CP.
    """
    domicilio_limpio = limpiar_direccion(domicilio)

    # Nombre de la vía sin tipo (ej: "Calle LOS TILOS 16" → "LOS TILOS 16")
    nombre_sin_tipo = re.sub(
        r"^(Calle|Avenida|Plaza|Paseo|Barrio|Grupo|Pasaje|Camino|Carretera|Urbanización)\s+",
        "", domicilio_limpio, flags=re.IGNORECASE
    ).strip()

    # Nombre de la vía sin tipo y sin número (ej: "LOS TILOS")
    nombre_sin_numero = re.sub(r"\s+\d+.*$", "", nombre_sin_tipo).strip()

    intentos = [
        # 1. Dirección completa expandida
        f"{domicilio_limpio}, {municipio}, {cp}, España",
        f"{domicilio_limpio}, {municipio}, España",
        # 2. Sin prefijo de tipo de vía (funciona mejor con OSM en euskera)
        f"{nombre_sin_tipo}, {municipio}, {cp}, España",
        f"{nombre_sin_tipo}, {municipio}, España",
        # 3. Solo nombre de calle sin número (para calles con nombre en euskera)
        f"{nombre_sin_numero}, {municipio}, {cp}, España",
        f"{nombre_sin_numero}, {municipio}, España",
        # 4. Fallback por municipio y CP
        f"{municipio}, {cp}, España",
        f"{cp}, España",
    ]

    headers = {"User-Agent": "ArgiaCarbonApp/1.0"}
    for intento in intentos:
        if not intento.strip().replace(",", "").replace("España", "").strip():
            continue
        try:
            r = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={"q": intento, "format": "json", "limit": 1, "countrycodes": "es"},
                headers=headers,
                timeout=10,
            )
            data = r.json()
            if data:
                lat = float(data[0]["lat"])
                lon = float(data[0]["lon"])
                # Validar que está en el norte de España (Euskadi + alrededores)
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
                "FECHA","CODIGO","CORREO","NOMBRE","DOMICILIO_USADO","MUNICIPIO","CP",
                "CENTRO","IMPUTACION_%","KM_IDA","KM_IDA_VUELTA_ANUALES",
                "MODO_TRANSPORTE","COMBUSTIBLE","PCT_MODO","DOMICILIO_CORREGIDO","ENVIO_NUM"
            ])
        for fila in filas:
            sheet.append_row(fila)
        return client
    except Exception as e:
        st.error(f"Error al guardar en Google Sheets: {e}")
        return None


def recalcular_sheet_calculos(client):
    """
    Borra y recalcula la pestaña CALCULOS desde cero,
    usando solo el último envío (ENVIO_NUM más alto) de cada empleado.
    """
    try:
        sheet_main = client.open(GOOGLE_SHEETS_NAME).sheet1
        datos      = sheet_main.get_all_values()

        if len(datos) <= 1:
            return

        cabecera = datos[0]
        filas    = datos[1:]

        # Índices de columnas necesarias (ENVIO_NUM puede no existir en datos históricos)
        try:
            idx_codigo = cabecera.index("CODIGO")
            idx_centro = cabecera.index("CENTRO")
            idx_km     = cabecera.index("KM_IDA_VUELTA_ANUALES")
            idx_modo   = cabecera.index("MODO_TRANSPORTE")
            idx_comb   = cabecera.index("COMBUSTIBLE")
        except ValueError as e:
            st.warning(f"No se pudo recalcular CALCULOS: columna no encontrada ({e})")
            return

        # ENVIO_NUM puede no existir en datos históricos: si falta, asumimos 1
        idx_envio = cabecera.index("ENVIO_NUM") if "ENVIO_NUM" in cabecera else None

        def get_envio_num(fila):
            if idx_envio is None:
                return 1
            try:
                return int(fila[idx_envio])
            except (ValueError, IndexError):
                return 1

        # Para cada empleado, quedarse solo con el último envío
        envios_por_codigo = {}
        for f in filas:
            cod = f[idx_codigo]
            num = get_envio_num(f)
            if cod not in envios_por_codigo or num > envios_por_codigo[cod]:
                envios_por_codigo[cod] = num

        # Filtrar filas que corresponden al último envío de cada empleado
        filas_validas = []
        for f in filas:
            cod = f[idx_codigo]
            num = get_envio_num(f)
            if num == envios_por_codigo.get(cod, num):
                filas_validas.append(f)

        # Agregar km por centro y modo
        modos_cols = [
            "Coche - Gasolina","Coche - Diésel","Coche - Eléctrico","Coche - Híbrido",
            "Furgoneta - Gasolina","Furgoneta - Diésel","Furgoneta - Eléctrico","Furgoneta - Híbrido",
            "Moto - Gasolina","Moto - Diésel","Moto - Eléctrico","Moto - Híbrido",
            "Transporte público","A pie o en bicicleta","TOTAL"
        ]

        # acumulador: {centro: {col: km}}
        acum = {}
        for f in filas_validas:
            centro = f[idx_centro]
            try:
                km = float(str(f[idx_km]).replace(",", "."))
            except ValueError:
                km = 0.0
            modo        = f[idx_modo].strip()
            combustible = f[idx_comb].strip()

            if combustible and combustible not in ["—", "", "None"]:
                col_key = f"{modo} - {combustible}"
            else:
                col_key = modo

            col_key_norm = col_key.strip().lower()
            col_found    = next(
                (mc for mc in modos_cols if mc.strip().lower() == col_key_norm),
                "TOTAL"
            )

            if centro not in acum:
                acum[centro] = {mc: 0.0 for mc in modos_cols}
            acum[centro][col_found] = round(acum[centro].get(col_found, 0.0) + km, 2)
            acum[centro]["TOTAL"]   = round(acum[centro].get("TOTAL", 0.0) + km, 2)

        # Escribir hoja CALCULOS desde cero
        try:
            sheet_calc = client.open(GOOGLE_SHEETS_NAME).worksheet("CALCULOS")
        except Exception:
            sheet_calc = client.open(GOOGLE_SHEETS_NAME).add_worksheet(
                title="CALCULOS", rows=100, cols=20)

        sheet_calc.clear()
        cabecera_calc = ["CENTRO"] + modos_cols
        sheet_calc.append_row(cabecera_calc)

        for centro, vals in acum.items():
            fila_calc = [centro] + [vals.get(mc, 0.0) for mc in modos_cols]
            sheet_calc.append_row(fila_calc)

    except Exception as e:
        st.warning(f"No se pudo recalcular la hoja CALCULOS: {e}")


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
    .aviso-2025 {{ background:#FFE8E8; border-left:5px solid #C00000;
        padding:0.8rem 1rem; border-radius:6px; font-size:0.92rem;
        color:#C00000; margin-bottom:1rem; font-weight:500; }}
    .aviso-reenvio {{ background:#FFF3E0; border-left:5px solid #FF9800;
        padding:0.8rem 1rem; border-radius:6px; font-size:0.92rem;
        color:#E65100; margin-bottom:1rem; font-weight:500; }}
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
    .modo-box {{ background:{COLOR_FONDO}; border:1px solid #DDD;
        padding:0.8rem 1rem; border-radius:8px; margin-bottom:0.6rem; }}
    .domicilio-header {{ background:{COLOR_TURQUESA}20; border-left:4px solid {COLOR_TURQUESA};
        padding:0.5rem 0.8rem; border-radius:4px; margin-bottom:0.6rem;
        font-weight:600; color:{COLOR_GRIS}; }}
    .stButton > button {{ background-color:{COLOR_TURQUESA} !important;
        color:white !important; border:none !important; border-radius:8px !important;
        font-weight:bold !important; padding:0.6rem 1.2rem !important; }}
    .stButton > button:hover {{ background-color:{COLOR_VERDE} !important; }}
    </style>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# BLOQUE DE MODOS DE TRANSPORTE POR DOMICILIO
# ─────────────────────────────────────────────
def render_modos_domicilio(centro, d_idx, dom_label, T, idioma):
    """
    Renderiza el bloque de selección de modos de transporte
    para un centro y un domicilio concretos.
    Devuelve lista de {modo, combustible, pct_modo} y un booleano todo_ok.
    """
    key_n    = f"n_modos_{centro}_{d_idx}"
    if key_n not in st.session_state:
        st.session_state[key_n] = 1

    n_modos     = st.session_state[key_n]
    modos_lista = []
    pct_total   = 0
    todo_ok     = True

    st.markdown(
        f'<div class="domicilio-header">🏠 {dom_label}</div>',
        unsafe_allow_html=True
    )

    for i in range(n_modos):
        st.markdown('<div class="modo-box">', unsafe_allow_html=True)

        modo = st.selectbox(
            f"{T['modo_label']} {i+1}",
            options=[T["selecciona"]] + TODOS_MODOS,
            format_func=lambda x: T["modos_display"].get(x, x) if x != T["selecciona"] else x,
            key=f"modo_{centro}_{d_idx}_{i}"
        )

        combustible = "—"
        if n_modos > 1:
            if modo != T["selecciona"] and modo in VEHICULOS_CON_COMBUSTIBLE:
                col_comb, col_pct, col_del = st.columns([3, 2, 1])
                with col_comb:
                    combustible = st.selectbox(
                        T["combustible_label"],
                        options=[T["selecciona"]] + TIPOS_COMBUSTIBLE,
                        format_func=lambda x: COMBUSTIBLE_DISPLAY[idioma].get(x, x) if x != T["selecciona"] else x,
                        key=f"comb_{centro}_{d_idx}_{i}"
                    )
                    if combustible == T["selecciona"]:
                        combustible = "—"
            else:
                col_pct, col_del = st.columns([4, 1])

            with col_pct:
                with st.expander("ℹ️"):
                    st.caption(T["pct_tooltip"])
                pct_modo = st.number_input(
                    T["pct_uso"],
                    min_value=1, max_value=99,
                    value=50,
                    key=f"pct_{centro}_{d_idx}_{i}",
                )
            with col_del:
                st.markdown("<br>", unsafe_allow_html=True)
                if i > 0 and st.button(T["eliminar_modo"], key=f"del_{centro}_{d_idx}_{i}"):
                    st.session_state[key_n] = n_modos - 1
                    st.rerun()
        else:
            if modo != T["selecciona"] and modo in VEHICULOS_CON_COMBUSTIBLE:
                combustible = st.selectbox(
                    T["combustible_label"],
                    options=[T["selecciona"]] + TIPOS_COMBUSTIBLE,
                    format_func=lambda x: COMBUSTIBLE_DISPLAY[idioma].get(x, x) if x != T["selecciona"] else x,
                    key=f"comb_{centro}_{d_idx}_{i}"
                )
                if combustible == T["selecciona"]:
                    combustible = "—"
            pct_modo = 100

        st.markdown('</div>', unsafe_allow_html=True)

        if modo == T["selecciona"]:
            todo_ok = False
        else:
            pct_total += pct_modo
            modos_lista.append({
                "modo": modo,
                "combustible": combustible,
                "pct_modo": pct_modo
            })

    # Botón añadir modo
    if n_modos < 5:
        if st.button(T["anadir_modo"], key=f"add_{centro}_{d_idx}"):
            st.session_state[key_n] = n_modos + 1
            st.rerun()

    # Validar porcentajes
    if n_modos > 1 and modos_lista and round(pct_total) != 100:
        st.warning(T["error_pct"])
        todo_ok = False

    return modos_lista, todo_ok


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

    # ── AVISO 2025 ────────────────────────────
    st.markdown(f'<div class="aviso-2025">{T["aviso_2025"]}</div>', unsafe_allow_html=True)
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
        correo_input = st.text_input(T["correo"],
                                     placeholder="nombre@argiafundazioa.org").strip().lower()
    with col2:
        dni = st.text_input(T["dni"], placeholder="Ej: 72263087P").strip().upper()

    if not correo_input or not dni:
        st.stop()

    if "@" not in correo_input:
        st.error(T["error_codigo"])
        st.stop()

    trabajador = df_trab[(df_trab["CORREO"] == correo_input) & (df_trab["DNI"] == dni)]
    if trabajador.empty:
        st.error(T["error_credenciales"])
        st.stop()

    row    = trabajador.iloc[0]
    nombre = row["NOMBRE"]
    codigo = str(row["CODIGO"]).zfill(4)

    dias_row    = df_dias[df_dias["CODIGO"] == codigo]
    dias_trab   = float(dias_row.iloc[0]["Nº DÍAS"])   if not dias_row.empty else DIAS_BASE
    pct_jornada = float(dias_row.iloc[0]["% JORNADA"]) if not dias_row.empty else 1.0

    st.markdown(f'<div class="exito">{T["bienvenido"]} <strong>{nombre}</strong></div>',
                unsafe_allow_html=True)

    # ── COMPROBACIÓN DE ENVÍO PREVIO ──────────
    envio_anterior = obtener_envio_anterior(codigo)
    if envio_anterior > 0:
        st.markdown(
            f'<div class="aviso-reenvio">{T["aviso_reenvio"].format(envio_anterior)}</div>',
            unsafe_allow_html=True
        )
    nuevo_envio_num = envio_anterior + 1

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

    if "n_domicilios" not in st.session_state:
        st.session_state["n_domicilios"] = 1

    domicilios  = []
    dom_pct_total = 0
    dom_valido  = True

    for d_idx in range(st.session_state["n_domicilios"]):
        label = f"**{T['domicilio_1'] if d_idx == 0 else T['domicilio_2']}**"
        st.markdown(label)

        if d_idx == 0:
            dom_correcto = st.radio(
                T["domicilio_correcto"],
                options=[T["si_correcto"], T["no_correcto"]],
                index=0,
                key=f"dom_correcto_{d_idx}"
            )
            if dom_correcto == T["no_correcto"]:
                st.markdown(f"**{T['introduce_domicilio']}**")
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1: dom_calle = st.text_input(T["calle"], placeholder="Ej: Calle Mayor 5 2A", key=f"dom_calle_{d_idx}")
                with c2: dom_muni  = st.text_input(T["municipio"], placeholder="Ej: Bilbao", key=f"dom_muni_{d_idx}")
                with c3: dom_cp    = st.text_input(T["cp"], placeholder="Ej: 48001", key=f"dom_cp_{d_idx}")
                if not dom_calle or not dom_muni or not dom_cp:
                    st.warning(T["aviso_domicilio"])
                    dom_valido = False
                    dom_calle  = domicilio_original
                    dom_muni   = municipio_original
                    dom_cp     = cp_original
                corregido = True
            else:
                dom_calle = domicilio_original
                dom_muni  = municipio_original
                dom_cp    = cp_original
                corregido = False
        else:
            st.markdown(f"**{T['introduce_domicilio']}**")
            c1, c2, c3 = st.columns([3, 1, 1])
            with c1: dom_calle = st.text_input(T["calle"], placeholder="Ej: Calle Mayor 5 2A", key=f"dom_calle_{d_idx}")
            with c2: dom_muni  = st.text_input(T["municipio"], placeholder="Ej: Bilbao", key=f"dom_muni_{d_idx}")
            with c3: dom_cp    = st.text_input(T["cp"], placeholder="Ej: 48001", key=f"dom_cp_{d_idx}")
            if not dom_calle or not dom_muni or not dom_cp:
                st.warning(T["aviso_domicilio"])
                dom_valido = False
                dom_calle  = ""
                dom_muni   = ""
                dom_cp     = ""
            corregido = True

        pct_dom = 100
        if st.session_state["n_domicilios"] > 1:
            col_pct, _ = st.columns([2, 3])
            with col_pct:
                with st.expander("ℹ️"):
                    st.caption(T["pct_domicilio_tooltip"])
                pct_dom = st.number_input(
                    T["pct_domicilio"],
                    min_value=1, max_value=99,
                    value=50,
                    key=f"dom_pct_{d_idx}"
                )
            dom_pct_total += pct_dom

        domicilios.append({
            "calle": dom_calle, "municipio": dom_muni, "cp": dom_cp,
            "pct": pct_dom, "corregido": corregido,
            "label": T['domicilio_1'] if d_idx == 0 else T['domicilio_2']
        })

    if st.session_state["n_domicilios"] == 1:
        if st.button(T["anadir_domicilio"], key="add_domicilio"):
            st.session_state["n_domicilios"] = 2
            st.rerun()
    else:
        if st.button(T["eliminar_domicilio"], key="del_domicilio"):
            st.session_state["n_domicilios"] = 1
            st.rerun()
        if round(dom_pct_total) != 100:
            st.warning(T["error_pct_domicilio"])
            dom_valido = False

    if not dom_valido:
        st.stop()

    # ── PASO 4: TRANSPORTE ────────────────────
    # Ahora los modos se configuran POR DOMICILIO dentro de cada centro
    st.markdown(f'<div class="seccion-titulo">{T["modo_transporte"]}</div>',
                unsafe_allow_html=True)
    st.caption(T["modo_caption"])

    # Estructura: {centro: [{domicilio_idx, modos:[{modo,combustible,pct_modo}]}]}
    modos_por_centro = {}
    todo_completado  = True

    for _, cr in centros_trab.iterrows():
        centro = cr["CENTRO"]
        pct    = cr["IMPUTACION_%"]

        if idioma == "eu":
            pct_txt_centro = f"Denboraren %{pct}ean"
        else:
            pct_txt_centro = f"{pct}% {T['denboraren']}"

        st.markdown(f"**🏢 {centro}** ({pct_txt_centro})")

        modos_centro = []

        for d_idx, dom in enumerate(domicilios):
            dom_label = f"{dom['label']} — {dom['municipio']} ({dom['pct']}%)"
            modos_lista, todo_ok = render_modos_domicilio(
                centro, d_idx, dom_label, T, idioma
            )
            if not todo_ok:
                todo_completado = False
            modos_centro.append({
                "d_idx":  d_idx,
                "modos":  modos_lista
            })

        modos_por_centro[centro] = modos_centro
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
            # Geocodificar todos los domicilios
            domicilios_geo = []
            geo_error = False
            for dom in domicilios:
                lon, lat = geocodificar_nominatim(dom["calle"], dom["municipio"], dom["cp"])
                if lon is None:
                    st.error(T["error_geolocalizacion"])
                    geo_error = True
                    break
                domicilios_geo.append({**dom, "lon": lon, "lat": lat})

            if geo_error:
                st.stop()

            resultados   = []
            filas_sheets = []
            errores      = []

            for _, cr in centros_trab.iterrows():
                centro     = cr["CENTRO"]
                imputacion = cr["IMPUTACION_%"]

                cd = df_cent[df_cent["CENTRO"].str.strip() == centro]
                if cd.empty or pd.isna(cd.iloc[0]["LAT"]):
                    errores.append(centro)
                    continue

                # Para cada domicilio, calcular km con sus propios modos
                for d_idx, dom_geo in enumerate(domicilios_geo):
                    km = calcular_km(
                        (dom_geo["lon"], dom_geo["lat"]),
                        (cd.iloc[0]["LON"], cd.iloc[0]["LAT"]),
                        ORS_API_KEY
                    )
                    if km is None:
                        errores.append(f"{centro} ({dom_geo['municipio']})")
                        continue

                    dom_corregido_txt = (
                        ("Bai" if dom_geo["corregido"] else "Ez")
                        if idioma == "eu"
                        else ("Sí" if dom_geo["corregido"] else "No")
                    )

                    # Obtener los modos específicos de este domicilio para este centro
                    modos_dom = next(
                        (mc["modos"] for mc in modos_por_centro.get(centro, [])
                         if mc["d_idx"] == d_idx),
                        []
                    )

                    for m in modos_dom:
                        modo        = m["modo"]
                        combustible = m["combustible"]
                        pct_modo    = m["pct_modo"] / 100

                        km_anuales = round(
                            km * 2
                            * (dias_trab / DIAS_BASE * DIAS_LABORABLES_2025)
                            * pct_jornada
                            * (imputacion / 100)
                            * pct_modo
                            * (dom_geo["pct"] / 100),
                            2
                        )

                        modo_es      = TEXTOS["es"]["modos_display"].get(modo, modo)
                        comb_es      = COMBUSTIBLE_DISPLAY["es"].get(combustible, combustible)
                        modo_display = T["modos_display"].get(modo, modo)
                        comb_display = COMBUSTIBLE_DISPLAY[idioma].get(combustible, combustible)

                        resultados.append({
                            "centro":              centro,
                            "imputacion":          imputacion,
                            "km":                  km,
                            "km_ida_vuelta_anuales": km_anuales,
                            "modo":                modo_es,
                            "combustible":         comb_es,
                            "modo_display":        modo_display,
                            "comb_display":        comb_display,
                            "pct_modo":            int(m["pct_modo"]),
                            "domicilio":           f"{dom_geo['calle']}, {dom_geo['municipio']}",
                            "pct_dom":             dom_geo["pct"],
                            "dom_label":           dom_geo["label"],
                        })

                        filas_sheets.append([
                            datetime.now().strftime("%Y-%m-%d %H:%M"),
                            codigo, correo_input, nombre,
                            f"{dom_geo['calle']}, {dom_geo['municipio']} ({dom_geo['pct']}%)",
                            dom_geo["municipio"], dom_geo["cp"],
                            centro, imputacion, km, km_anuales,
                            modo_display, comb_display, int(m["pct_modo"]),
                            dom_corregido_txt,
                            nuevo_envio_num,          # ← ENVIO_NUM
                        ])

        if resultados:
            st.markdown(f"**{T['distancias_calculadas']}**")
            centro_actual = None
            for r in resultados:
                if r["centro"] != centro_actual:
                    centro_actual = r["centro"]
                    if idioma == "eu":
                        pct_label = f"Denboraren %{r['imputacion']}ean"
                    else:
                        pct_label = f"{r['imputacion']}% {T['denboraren']}"
                    st.markdown(f"**🏢 {r['centro']}** ({pct_label})")

                comb_txt = f" — {r['comb_display']}" if r['combustible'] not in ["—",""] else ""
                pct_txt  = f" ({r['pct_modo']}%)" if r["pct_modo"] < 100 else ""
                dom_txt  = (
                    f"🏠 {r.get('dom_label','')} — {r.get('domicilio', '')} ({r.get('pct_dom', 100)}%)<br>"
                    if len(domicilios_geo) > 1 else ""
                )
                st.markdown(f"""
                <div class="km-box">
                {dom_txt}🚗 {r['modo_display']}{comb_txt}{pct_txt}<br>
                📏 {T['distancia_ida']} <strong>{r['km']} km</strong><br>
                📅 {T['km_anuales']} <strong>{r['km_ida_vuelta_anuales']} km</strong>
                </div>
                """, unsafe_allow_html=True)

            if errores:
                st.warning(f"{T['error_centro']} {', '.join(errores)}")

            client = guardar_en_sheets(filas_sheets)
            if client:
                # Recalcular CALCULOS desde cero con solo el último envío de cada empleado
                recalcular_sheet_calculos(client)
                st.session_state["datos_enviados"] = True
                st.session_state["piloto_codigo"]  = codigo
                st.session_state["piloto_correo"]  = correo_input
                st.session_state["piloto_nombre"]  = nombre
                st.session_state["piloto_client"]  = client
                st.session_state["piloto_idioma"]  = idioma
                st.markdown(f'<div class="exito">{T["gracias"]}</div>',
                            unsafe_allow_html=True)
                st.balloons()
            else:
                st.warning(T["error_sheets"])
        else:
            st.error(T["error_distancias"])

    # ── FORMULARIO PILOTO ─────────────────────
    if st.session_state.get("datos_enviados"):
        idioma_p = st.session_state.get("piloto_idioma", idioma)
        TP       = TEXTOS[idioma_p]
        codigo_p = st.session_state["piloto_codigo"]
        correo_p = st.session_state["piloto_correo"]
        nombre_p = st.session_state["piloto_nombre"]
        client_p = st.session_state["piloto_client"]

        st.markdown("---")
        st.markdown(f'<div class="seccion-titulo">{TP["piloto_titulo"]}</div>',
                    unsafe_allow_html=True)
        st.caption(TP["piloto_subtitulo"])

        p1 = st.selectbox(TP["piloto_p1"], options=TP["piloto_p1_ops"], key="piloto_p1")
        p2 = st.selectbox(TP["piloto_p2"], options=TP["piloto_p2_ops"], key="piloto_p2")
        p3 = st.selectbox(TP["piloto_p3"], options=TP["piloto_p3_ops"], key="piloto_p3")
        p4 = st.text_area(TP["piloto_p4"], placeholder="...", key="piloto_p4", height=100)

        if not st.session_state.get("piloto_enviado"):
            if st.button(TP["piloto_boton"], key="piloto_enviar"):
                selecciona = TP["piloto_p1_ops"][0]
                if p1 == selecciona or p2 == selecciona or p3 == selecciona:
                    st.warning(TP["piloto_incompleto"])
                else:
                    try:
                        try:
                            sheet_piloto = client_p.open(GOOGLE_SHEETS_NAME).worksheet("PILOTO")
                        except Exception:
                            sheet_piloto = client_p.open(GOOGLE_SHEETS_NAME).add_worksheet(
                                title="PILOTO", rows=200, cols=10)
                            sheet_piloto.append_row([
                                "FECHA", "CODIGO", "CORREO", "NOMBRE",
                                "FACILIDAD_USO", "DATOS_CORRECTOS",
                                "COMPRENSION", "COMENTARIOS"
                            ])
                        sheet_piloto.append_row([
                            datetime.now().strftime("%Y-%m-%d %H:%M"),
                            codigo_p, correo_p, nombre_p,
                            p1, p2, p3, p4
                        ])
                        st.session_state["piloto_enviado"] = True
                        st.rerun()
                    except Exception:
                        st.warning(TP["piloto_error"])
        else:
            st.markdown(f'<div class="exito">{TP["piloto_gracias"]}</div>',
                        unsafe_allow_html=True)


if __name__ == "__main__":
    main()
