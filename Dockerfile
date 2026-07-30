# Imagen para Render (u otro Docker): Node + LibreOffice + puente UNO.
FROM node:20-bookworm-slim

# LibreOffice (Calc + Writer), Python con el puente UNO, y fuentes para que
# los PDF se vean como el Excel original (Carlito ~ Calibri, Liberation ~ Arial).
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-calc \
      libreoffice-writer \
      python3 \
      python3-uno \
      python3-pip \
      fonts-crosextra-carlito \
      fonts-liberation \
      fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Dependencias de Node (se cachean si no cambia package.json)
COPY package.json package-lock.json* ./
RUN npm install --omit=dev

# Detección de contraseña (opcional; convert.py tiene respaldo sin esta librería)
COPY requirements.txt ./
RUN pip3 install --no-cache-dir --break-system-packages -r requirements.txt || true

# Código de la app
COPY . .

ENV PYTHON_BIN=python3

# El convertidor escucha el grupo por polling; no necesita exponer puertos.
CMD ["node", "bot.js"]
