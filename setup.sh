#!/usr/bin/env bash
# Instalador para servidor Ubuntu/Debian.
# Deja todo listo para correr el bot Excel/Word -> PDF por WhatsApp.
# Uso:  bash setup.sh
set -e

echo "==> Actualizando paquetes..."
sudo apt-get update -y

echo "==> Instalando LibreOffice + motor UNO + Python + Node..."
sudo apt-get install -y \
  libreoffice-calc libreoffice-writer \
  python3 python3-pip python3-uno \
  nodejs npm

echo "==> Instalando dependencias de Python (detección de contraseña)..."
pip3 install -r requirements.txt --break-system-packages 2>/dev/null || pip3 install -r requirements.txt

echo "==> Instalando dependencias de Node..."
npm install

echo "==> Instalando pm2 (para mantener el bot 24/7)..."
sudo npm install -g pm2

if [ ! -f .env ]; then
  cp .env.example .env
  echo "==> Se creó .env — EDÍTALO con tus credenciales de Green API antes de arrancar."
fi

echo ""
echo "======================================================================"
echo " Listo. Pasos finales:"
echo "   1) Edita el archivo .env con GREENAPI_ID_INSTANCE y GREENAPI_API_TOKEN"
echo "   2) Arranca en modo prueba:   npm start"
echo "      (sube un archivo al grupo y copia el chatId ...@g.us que salga)"
echo "   3) Pega ese chatId en ALLOWED_CHAT_ID dentro de .env"
echo "   4) Deja el bot 24/7:"
echo "        pm2 start bot.js --name pdf-wa"
echo "        pm2 save && pm2 startup   (sigue lo que imprima)"
echo "======================================================================"
