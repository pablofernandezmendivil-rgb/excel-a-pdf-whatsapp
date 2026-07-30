# Guía de despliegue (servidor en la nube) — recomendado

Esta es la forma más simple y confiable de dejar el bot corriendo 24/7 para tu
grupo "pdf convertidor".

## 1. Crea un servidor Ubuntu

Contrata un VPS chico (alcanza y sobra el más barato, ~4–6 USD/mes) con
**Ubuntu 22.04 o 24.04**. Opciones comunes: Hetzner, DigitalOcean, Vultr.
Al crearlo te dan una **IP** y una forma de entrar por **SSH**.

Tamaño sugerido: 1 vCPU y 1–2 GB de RAM es suficiente.

## 2. Entra al servidor

Desde tu compu:

```bash
ssh root@LA_IP_DEL_SERVIDOR
```

## 3. Sube el proyecto

Opción fácil: instala git y clona/copía la carpeta. O usa `scp` desde tu compu:

```bash
scp -r excel-a-pdf-whatsapp root@LA_IP_DEL_SERVIDOR:/root/
```

## 4. Instala todo con un comando

```bash
cd excel-a-pdf-whatsapp
bash setup.sh
```

Ese script instala LibreOffice, el motor UNO, Python, Node, las dependencias
y pm2.

## 5. Pon tus credenciales

```bash
nano .env
```

Llena `GREENAPI_ID_INSTANCE` y `GREENAPI_API_TOKEN` (guardar: Ctrl+O, Enter; salir: Ctrl+X).

## 6. Descubre el chatId del grupo

```bash
npm start
```

Sube cualquier archivo al grupo "pdf convertidor". En pantalla verás:

```
Archivo recibido en 120363XXXXXXXXXX@g.us: "..."
```

Copia ese `...@g.us`. Detén con Ctrl+C, edita `.env` otra vez y pégalo en
`ALLOWED_CHAT_ID`. Guarda.

## 7. Déjalo encendido 24/7

```bash
pm2 start bot.js --name pdf-wa
pm2 save
pm2 startup     # ejecuta el comando que te imprima
```

Comandos útiles:

```bash
pm2 logs pdf-wa      # ver actividad en vivo
pm2 restart pdf-wa   # reiniciar (después de editar .env)
pm2 stop pdf-wa      # detener
```

¡Listo! Ya puede tu equipo/proveedores subir Excel o Word al grupo y recibir el PDF.

---

## Alternativa gratis para SOLO probar: tu propia PC

Sirve para una prueba rápida, pero **no** para el uso real (si la PC se apaga o
se duerme, el bot deja de responder).

- **Linux/Mac**: mismos pasos, corre `bash setup.sh` (en Mac instala LibreOffice
  y Node aparte) y luego `npm start`.
- **Windows**: es más enredado porque el motor UNO vive dentro del Python que
  trae LibreOffice. Tendrías que instalar Node y LibreOffice, apuntar
  `PYTHON_BIN` en `.env` a
  `C:\Program Files\LibreOffice\program\python.exe`, e instalar ahí
  `msoffcrypto-tool`. Si te decides por Windows, dime y te paso los pasos
  exactos.
