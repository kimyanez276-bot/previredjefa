# 🚀 Guía de Despliegue en Streamlit Community Cloud vinculado a Google Sheets

Esta guía detalla paso a paso cómo publicar el **Dashboard Contable de Previred y DJ 1887** de forma **100% gratuita** en la nube de Streamlit ([Streamlit Community Cloud](https://share.streamlit.io)), permitiendo que Marlene (desde su casa) y las 3 analistas (en la oficina) trabajen en tiempo real con sincronización a Google Drive / Google Sheets.

---

## 📋 Arquitectura de la Solución Gratuita

```
[Marlene en Casa]       [3 Analistas en Oficina]
        │                           │
        ▼                           ▼
 ┌──────────────────────────────────────────────┐
 │   Streamlit Community Cloud (Servidor Web)   │
 │   - Autenticación por roles                  │
 │   - Parseo PDF con pdfplumber                │
 │   - Edición en vivo con st.data_editor       │
 └──────────────────────┬───────────────────────┘
                        │ (vía gspread y st.secrets)
                        ▼
 ┌──────────────────────────────────────────────┐
 │   Google Sheets en Google Drive              │
 │   - Hoja compartida en tiempo real           │
 │   - Sincronización instantánea               │
 │   - Respaldo permanente en la nube           │
 └──────────────────────────────────────────────┘
```

---

## 🛠️ PASO 1: Subir el Proyecto a GitHub

1. Entra a [github.com](https://github.com) con tu cuenta (si no tienes, crea una gratis).
2. Crea un **Nuevo Repositorio** (puedes llamarlo `conta-marlene-previred`).
3. Sube los archivos del proyecto que hemos creado:
   - `app.py`
   - `requirements.txt`
   - Carpeta `modules/` (con `__init__.py`, `parser_previred.py`, `tax_engine.py`, `storage.py`, `report_generator.py`)
   - `.gitignore` (ignora la carpeta `data/` y archivos `.db` locales).

---

## 🔑 PASO 2: Crear Credenciales de Google Cloud (Service Account Gratuita)

Para que Streamlit pueda escribir y leer de tu Google Drive sin pedir contraseñas interactivas, Google utiliza una **Cuenta de Servicio (Service Account)** gratuita:

1. Ve a [Google Cloud Console](https://console.cloud.google.com/).
2. Crea un proyecto nuevo llamado **"Conta Marlene"**.
3. En el menú lateral izquierdo, ve a **APIs y Servicios > Biblioteca**:
   - Busca **"Google Sheets API"** y presiona **Habilitar**.
   - Busca **"Google Drive API"** y presiona **Habilitar**.
4. Ve a **APIs y Servicios > Credenciales**:
   - Presiona **"Crear Credenciales"** > **"Cuenta de Servicio" (Service Account)**.
   - Dale un nombre (ej. `streamlit-sync`) y presiona **Crear y Continuar**.
   - En el rol, selecciona **Proyecto > Editor** y finaliza.
5. Haz clic sobre la cuenta de servicio recién creada, entra a la pestaña **Claves (Keys)**:
   - Presiona **"Agregar Clave" > "Crear clave nueva"**.
   - Elige formato **JSON** y presiona **Crear**.
   - Se descargará a tu computador un archivo `.json` con tus credenciales.

---

## 📊 PASO 3: Crear la Planilla en Google Drive

1. Entra a tu [Google Drive](https://drive.google.com) y crea una nueva **Hoja de cálculo de Google (Google Sheets)**.
2. Nómbrala: `Conta_Marlene_76519206-4` (o el nombre de tu empresa cliente).
3. Presiona el botón verde **Compartir** (arriba a la derecha).
4. Copia el correo electrónico de tu cuenta de servicio (aparece en el JSON en `"client_email"`, ej. `streamlit-sync@conta-marlene.iam.gserviceaccount.com`).
5. Pégalo en el cuadro de compartir con permiso de **"Editor"** y desmarca notificar. ¡Listo! Ahora tu aplicación tiene permiso para escribir en esa hoja.

---

## ☁️ PASO 4: Desplegar en Streamlit Community Cloud

1. Entra a [share.streamlit.io](https://share.streamlit.io/) e inicia sesión con tu cuenta de GitHub.
2. Presiona el botón **"New app"** (Nueva aplicación).
3. Configura:
   - **Repository:** `tu-usuario/conta-marlene-previred`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. **IMPORTANTE - Configurar los Secretos:**
   - Antes de presionar Deploy, haz clic en **"Advanced Settings"** (Configuración avanzada).
   - En la sección **Secrets**, pega el contenido de tu archivo JSON descargado en formato TOML y el título de tu planilla:

```toml
google_sheet_title = "Conta_Marlene_76519206-4"

[gcp_service_account]
type = "service_account"
project_id = "conta-marlene-XXXXXX"
private_key_id = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
private_key = "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQD...==\n-----END PRIVATE KEY-----\n"
client_email = "streamlit-sync@conta-marlene.iam.gserviceaccount.com"
client_id = "123456789012345678901"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/streamlit-sync%40conta-marlene.iam.gserviceaccount.com"
```

5. Presiona **Save** y luego **Deploy!**.
6. En 2 a 3 minutos, tu aplicación estará funcionando en una URL pública y segura (ej. `https://conta-marlene.streamlit.app`).

---

## 👩‍💻 PASO 5: Cómo Trabajar las 4 Usuarias

1. **Marlene (Jefa Remota):**
   - Ingresa desde su casa con el usuario `marlene`.
   - Puede revisar en vivo los semáforos de cuadratura, verificar las inconsistencias, autorizar los cierres anuales, consultar quién hizo cada cambio en la pestaña de Auditoría y descargar el archivo plano DJ 1887 final para el SII.
2. **Las 3 Analistas en la Oficina:**
   - Ingresan con sus usuarios `analista1`, `analista2` o `analista3`.
   - Suben los comprobantes mensuales de Previred (drag-and-drop).
   - Realizan ajustes en la tabla interactiva (`st.data_editor`).
   - Al presionar **"Guardar Cambios"**, el sistema audita automáticamente el cambio con su nombre de usuaria y fecha.
   - Al presionar **"Sincronizar a Google Sheets"**, la planilla en la nube se actualiza de inmediato para que Marlene la vea en tiempo real.

---

## 🖥️ Cómo ejecutarlo en el computador de la oficina (Modo Local)

Si en la oficina prefieren usarlo en sus computadores sin internet o en red local:
1. Haz doble clic en el archivo **`Iniciar_Previred_DJ1887.bat`**.
2. Se abrirá automáticamente tu navegador en `http://localhost:8501`.
3. Todos los datos se guardan en la base de datos SQLite local dentro de la carpeta `data/`.
