# Desplegar el convertidor en Render

Como tu cotizador ya vive en Render, aquí ponemos el convertidor igual. Render
levanta LibreOffice solo, usando el `Dockerfile` que viene incluido. Tú no
instalas nada en tu compu.

## Qué tipo de servicio

Un **Background Worker** (trabajador en segundo plano). No es una página web:
solo se queda encendido escuchando el grupo y convirtiendo. En Render un
Background Worker requiere un plan de pago (el más chico, "Starter", basta).

## Paso a paso

### 1. Sube esta carpeta a un repositorio de GitHub

Puede ser un repo nuevo solo para el convertidor (lo más limpio). Sube TODOS
los archivos de esta carpeta, incluyendo `Dockerfile` y `render.yaml`.
No subas `.env` (tiene tus credenciales); ya está en `.gitignore`.

### 2. Crea el servicio en Render

**Opción rápida (Blueprint):**
1. En Render: **New +** → **Blueprint**.
2. Conecta el repo que subiste. Render leerá `render.yaml` y creará el worker
   `pdf-convertidor` automáticamente.

**Opción manual:**
1. **New +** → **Background Worker**.
2. Conecta el repo. En *Runtime* elige **Docker**.
3. Plan: **Starter**.

### 3. Pon las variables de entorno (Environment)

En el servicio → pestaña **Environment**, agrega:

| Variable | Valor |
|---|---|
| `GREENAPI_ID_INSTANCE` | `7107629985` |
| `GREENAPI_API_TOKEN` | *(tu apiTokenInstance de Green API)* |
| `GREENAPI_API_URL` | `https://7107.api.greenapi.com` |
| `GREENAPI_MEDIA_URL` | `https://7107.api.greenapi.com` |
| `ALLOWED_CHAT_ID` | *(déjala vacía por ahora)* |
| `PYTHON_BIN` | `python3` |

(Si usaste el Blueprint, las URLs ya vienen puestas; solo llena el id, el token
y luego el `ALLOWED_CHAT_ID`.)

### 4. Deploy y descubre el chatId del grupo

1. Render construye la imagen (tarda unos minutos la primera vez por LibreOffice).
2. Abre la pestaña **Logs**. Debe decir `Bot Excel/Word → PDF iniciado...`.
3. Sube cualquier Excel al grupo "PDF convertidor".
4. En los Logs verás: `Archivo recibido en 120363XXXXXXXXXX@g.us`.
5. Copia ese `...@g.us`, ponlo en la variable `ALLOWED_CHAT_ID` y guarda.
   Render reinicia solo.

¡Listo! A partir de ahí, cada Excel o Word que suban al grupo regresa como PDF.

## Requisito en Green API

En la consola de Green API, en los ajustes de la instancia, deja activadas las
notificaciones de **mensajes entrantes** (`incomingWebhook` en "on"). Sin eso,
el convertidor no "ve" los archivos.

## Notas

- El worker no expone URL pública (no es una web). Si prefieres correrlo como
  **Web Service**, también funciona: el código levanta un mini endpoint de salud
  cuando Render define `PORT`. Pero para un servicio que debe estar SIEMPRE
  encendido, el Background Worker es lo correcto (los Web Service gratuitos se
  duermen).
- Todo lo que necesita LibreOffice para escribir se maneja en carpetas
  temporales, así que funciona bien en el disco efímero de Render.
