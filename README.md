# App Huella de Carbono 2025 — Argia Fundazioa
## Desplazamientos de empleados

---

## Archivos necesarios

```
app.py                        ← código principal de la app
requirements.txt              ← dependencias Python
Informe_trabajadores_domicilio_imputacion.XLS  ← datos de trabajadores
DIRECCIONES_CENTROS.xlsx      ← direcciones de los centros
.streamlit/secrets.toml       ← claves y credenciales (NO subir a GitHub)
```

---

## Pasos para poner en marcha

### 1. Obtener clave API de OpenRouteService (gratuita)
1. Ve a https://openrouteservice.org/dev/#/signup
2. Crea una cuenta gratuita
3. En el dashboard copia tu API Key
4. Pégala en secrets.toml como ORS_API_KEY

### 2. Configurar Google Sheets
1. Ve a https://console.cloud.google.com
2. Crea un proyecto nuevo
3. Activa la API de Google Sheets y Google Drive
4. Crea una cuenta de servicio → descarga el JSON de credenciales
5. Crea un Google Sheet llamado "HuellaCarbonoTransporte2025"
6. Comparte ese Sheet con el email de la cuenta de servicio (editor)
7. Pega el contenido del JSON en secrets.toml como GOOGLE_CREDENTIALS

### 3. Configurar secrets.toml
Crea la carpeta .streamlit en tu proyecto y dentro el fichero secrets.toml:
```
.streamlit/
    secrets.toml    ← rellena con tus claves reales
```

### 4. Desplegar en Streamlit Community Cloud
1. Sube todos los archivos a un repositorio GitHub (privado)
2. Ve a https://share.streamlit.io
3. Conecta tu cuenta de GitHub
4. Selecciona el repositorio y el fichero app.py
5. En "Advanced settings" → "Secrets" pega el contenido de tu secrets.toml
6. Haz clic en Deploy

---

## Estructura del Google Sheet de resultados

La app vuelca automáticamente una fila por cada trabajador con estas columnas:

| FECHA | CODIGO | DNI | NOMBRE | DOMICILIO_USADO | MUNICIPIO | CP | CENTRO | IMPUTACION_% | KM_IDA | MODO_TRANSPORTE | DOMICILIO_CORREGIDO |
|---|---|---|---|---|---|---|---|---|---|---|---|

---

## Notas importantes
- Los km calculados son de **ida** solamente
- Si un trabajador tiene varios centros, se genera una fila por centro
- Si un trabajador corrige su domicilio, queda registrado en la columna DOMICILIO_CORREGIDO
- La app no permite ver datos de otros trabajadores: cada uno solo ve los suyos al introducir su código y DNI
