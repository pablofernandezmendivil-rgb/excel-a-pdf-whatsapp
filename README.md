# Bot Excel/Word → PDF por WhatsApp (Green API)

Bot para un **grupo de WhatsApp** (por ejemplo tu grupo "pdf"): cuando un
proveedor sube un **Excel** o un **Word**, el bot responde con el **PDF** en el
mismo grupo. Nada de texto, solo el archivo.

## Reglas de conversión (así opera)

- **Solo responde en el grupo configurado** (`ALLOWED_CHAT_ID`).
- **Excel**: se convierte **solo la primera pestaña** (sin importar su nombre).
  Las demás hojas y las vacías se descartan.
- **Ancho de una hoja**: todas las columnas se ajustan al **ancho de una sola
  página**; hacia abajo usa las páginas que hagan falta. Se **respeta la
  orientación** original del archivo.
- **Conserva** formato, colores e **imágenes/logos** (usa el motor interno de
  LibreOffice, no una conversión "plana").
- **Word**: se convierte tal cual a PDF.
- **Nombre del PDF**: el **caption** con el que subieron el archivo; si no hay
  caption, el **nombre original** del archivo.
- **Con contraseña**: responde el texto `Error archivo con contraseña`.
- **Cualquier otra cosa** (foto, PDF ya hecho, texto, otro formato): **silencio**.

---

## 1. Requisitos

- **Node.js 18+** → https://nodejs.org
- **LibreOffice** con enlace UNO para Python:
  - Ubuntu/Debian: `sudo apt install -y libreoffice-calc libreoffice-writer python3-uno`
  - Windows: instala [LibreOffice](https://www.libreoffice.org/download/); trae
    el módulo `uno` para el Python que incluye. En Windows conviene apuntar
    `PYTHON_BIN` al Python de LibreOffice
    (`C:\Program Files\LibreOffice\program\python.exe`).
- **Python 3** con la librería de detección de contraseña:
  `pip install -r requirements.txt`  (instala `msoffcrypto-tool`)
- Una **instancia de Green API** vinculada a tu número de WhatsApp Business.

---

## 2. Configurar Green API

1. Entra a https://console.green-api.com y crea/abre tu instancia.
2. Vincula tu WhatsApp Business: escanea el **QR** desde el teléfono
   (WhatsApp → Ajustes → Dispositivos vinculados → Vincular un dispositivo).
3. Copia de la instancia: `idInstance` y `apiTokenInstance`.
4. En los ajustes de la instancia activa las notificaciones de **mensajes
   entrantes** (`incomingWebhook` en "on"). Este bot lee por *polling*: no
   necesitas abrir puertos ni tener URL pública.
5. Agrega el número del bot al grupo "pdf".

---

## 3. Instalar

```bash
cd excel-a-pdf-whatsapp
npm install
pip install -r requirements.txt        # msoffcrypto-tool
cp .env.example .env                    # Windows: copy .env.example .env
```

Edita `.env` con tu `GREENAPI_ID_INSTANCE` y `GREENAPI_API_TOKEN`.

### Fijar el grupo "pdf"

1. Deja `ALLOWED_CHAT_ID` vacío y arranca el bot (`npm start`).
2. Sube cualquier archivo al grupo "pdf".
3. En la consola verás una línea `Archivo recibido en 120363...@g.us`.
   Copia ese `@g.us` y pégalo en `ALLOWED_CHAT_ID`.
4. Reinicia el bot. A partir de ahí solo responderá en ese grupo.

---

## 4. Probar la conversión sin WhatsApp

```bash
python3 convert.py algun_archivo.xlsx salida.pdf ; echo "codigo: $?"
```

Códigos: `0` ok · `10` con contraseña · `2` tipo no soportado · `1` otro error.

---

## 5. Arrancar

```bash
npm start
```

Sube un Excel o Word al grupo y el bot responderá con el PDF. 🎉

---

## 6. Dejarlo corriendo 24/7

Debe estar encendido para escuchar el grupo.

- **Tu PC**: sirve mientras la dejes prendida (bueno para pruebas).
- **Servidor / VPS** (producción). En Linux, con **pm2**:

  ```bash
  npm install -g pm2
  pm2 start bot.js --name pdf-wa
  pm2 save
  pm2 startup     # sigue lo que imprime para que arranque solo al reiniciar
  ```

---

## Cómo funciona (resumen técnico)

1. `bot.js` consulta `receiveNotification` de Green API (long polling) y solo
   atiende el `ALLOWED_CHAT_ID`.
2. Si es un `documentMessage` con extensión de Excel o Word, descarga el archivo.
3. Llama a `convert.py` (LibreOffice vía UNO):
   - detecta contraseña con `msoffcrypto` (→ código 10);
   - en Excel deja la 1ª hoja y aplica *fit-to-width = 1 página*;
   - en Word convierte directo; exporta el PDF.
4. Según el código de salida: envía el PDF (nombre = caption o nombre original),
   o el texto de contraseña, o guarda silencio.
5. Borra la notificación para no reprocesarla. Los temporales se borran solos.

## Solución de problemas

| Síntoma | Solución |
|---|---|
| `Faltan variables en el archivo .env` | Crea `.env` con tu id y token. |
| Error de `import uno` / no convierte | Instala `python3-uno` (o apunta `PYTHON_BIN` al Python de LibreOffice en Windows). |
| No detecta contraseña | Instala `msoffcrypto-tool` (`pip install -r requirements.txt`). |
| Responde en chats que no quieres | Configura `ALLOWED_CHAT_ID` con el grupo. |
| El PDF encoge mucho el texto | Es efecto de "todas las columnas en un ancho de página" en hojas muy anchas; es el comportamiento pedido. |

## Notas

- Green API tiene límites según tu plan (archivos/mes). Revísalo si el volumen es alto.
- El bot conserva la orientación (vertical/horizontal) que traiga cada archivo.
- Los archivos temporales se guardan en `tmp/` y se borran tras cada envío
  (no se guarda respaldo).
