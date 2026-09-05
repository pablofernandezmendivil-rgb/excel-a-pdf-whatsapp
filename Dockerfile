# Imagen para Render: Node + LibreOffice + UNO + poppler (pdfunite) + fuentes.
FROM node:20-bookworm-slim

# LibreOffice (Calc+Writer), Python/UNO, poppler-utils (pdfunite para unir PDFs),
# y fuentes: Carlito~Calibri, Liberation~Arial, y Noto CJK para chino/japonés/coreano
# (necesaria para que la cotización original en chino se vea correcta en el PDF).
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-calc \
      libreoffice-writer \
      python3 \
      python3-uno \
      python3-pip \
      poppler-utils \
      fonts-crosextra-carlito \
      fonts-liberation \
      fonts-dejavu \
      fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm install --omit=dev

COPY requirements.txt ./
RUN pip3 install --no-cache-dir --break-system-packages -r requirements.txt || true

COPY . .

ENV PYTHON_BIN=python3

CMD ["node", "server.js"]
