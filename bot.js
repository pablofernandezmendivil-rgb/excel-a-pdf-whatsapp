'use strict';

require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const axios = require('axios');
const FormData = require('form-data');

// ---------------------------------------------------------------------------
// Configuración
// ---------------------------------------------------------------------------
const ID_INSTANCE = process.env.GREENAPI_ID_INSTANCE;
const API_TOKEN = process.env.GREENAPI_API_TOKEN;
const API_URL = (process.env.GREENAPI_API_URL || 'https://api.green-api.com').replace(/\/$/, '');
const MEDIA_URL = (process.env.GREENAPI_MEDIA_URL || 'https://media.green-api.com').replace(/\/$/, '');
const ALLOWED_CHAT_ID = (process.env.ALLOWED_CHAT_ID || '').trim();
const RECEIVE_TIMEOUT = parseInt(process.env.RECEIVE_TIMEOUT || '15', 10);
const PYTHON_BIN = process.env.PYTHON_BIN || 'python3';

// Mensaje que se envía cuando el Excel/Word viene con contraseña.
const PASSWORD_MSG = process.env.PASSWORD_MSG || 'Error archivo con contraseña';

const EXCEL_EXT = ['.xlsx', '.xls', '.xlsm', '.xlsb', '.ods', '.csv', '.xltx'];
const WORD_EXT = ['.doc', '.docx', '.docm', '.odt', '.rtf', '.dotx'];
const CONVERTIBLE_EXT = new Set([...EXCEL_EXT, ...WORD_EXT]);

const TMP_DIR = path.join(__dirname, 'tmp');
fs.mkdirSync(TMP_DIR, { recursive: true });
const CONVERTER = path.join(__dirname, 'convert.py');

// Códigos de salida de convert.py
const EXIT_OK = 0;
const EXIT_PASSWORD = 10;
const EXIT_UNSUPPORTED = 2;

// ---------------------------------------------------------------------------
// Utilidades
// ---------------------------------------------------------------------------
function log(...args) {
  console.log(`[${new Date().toISOString()}]`, ...args);
}

function checkConfig() {
  const missing = [];
  if (!ID_INSTANCE) missing.push('GREENAPI_ID_INSTANCE');
  if (!API_TOKEN) missing.push('GREENAPI_API_TOKEN');
  if (missing.length) {
    console.error(
      `Faltan variables en el archivo .env: ${missing.join(', ')}.\n` +
        'Copia .env.example a .env y complétalo con tus credenciales de Green API.'
    );
    process.exit(1);
  }
  if (!fs.existsSync(CONVERTER)) {
    console.error(`No se encontró el conversor: ${CONVERTER}`);
    process.exit(1);
  }
}

function isConvertible(fileName) {
  const ext = path.extname(fileName || '').toLowerCase();
  return CONVERTIBLE_EXT.has(ext);
}

// Nombre del PDF: usa el caption si existe; si no, el nombre original.
function pdfNameFrom(caption, originalName) {
  let base = (caption || '').trim();
  if (!base) {
    base = path.basename(originalName || 'archivo', path.extname(originalName || ''));
  }
  // Quitar una extensión que el usuario haya escrito en el caption.
  base = base.replace(/\.(pdf|xlsx?|xlsm|xlsb|ods|csv|docx?|docm|odt|rtf)$/i, '');
  // Sanitizar: quitar separadores de ruta y caracteres problemáticos, conservar acentos/espacios.
  base = base.replace(/[\\/:*?"<>|\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!base) base = 'archivo';
  return base.slice(0, 100) + '.pdf';
}

function runConverter(inputPath, outputPath) {
  return new Promise((resolve) => {
    const proc = spawn(PYTHON_BIN, [CONVERTER, inputPath, outputPath], {
      stdio: ['ignore', 'ignore', 'pipe'],
    });
    let stderr = '';
    proc.stderr.on('data', (d) => (stderr += d.toString()));
    proc.on('error', (err) => resolve({ code: -1, stderr: err.message }));
    proc.on('close', (code) => resolve({ code, stderr }));
  });
}

// ---------------------------------------------------------------------------
// Green API
// ---------------------------------------------------------------------------
async function receiveNotification() {
  const url = `${API_URL}/waInstance${ID_INSTANCE}/receiveNotification/${API_TOKEN}?receiveTimeout=${RECEIVE_TIMEOUT}`;
  const { data } = await axios.get(url, { timeout: (RECEIVE_TIMEOUT + 10) * 1000 });
  return data;
}

async function deleteNotification(receiptId) {
  const url = `${API_URL}/waInstance${ID_INSTANCE}/deleteNotification/${API_TOKEN}/${receiptId}`;
  await axios.delete(url, { timeout: 20000 });
}

async function downloadFile(downloadUrl, destPath) {
  const resp = await axios.get(downloadUrl, { responseType: 'stream', timeout: 60000 });
  await new Promise((resolve, reject) => {
    const w = fs.createWriteStream(destPath);
    resp.data.pipe(w);
    w.on('finish', resolve);
    w.on('error', reject);
  });
}

async function sendPdf(chatId, pdfPath, fileName) {
  const form = new FormData();
  form.append('chatId', chatId);
  form.append('caption', ''); // sin texto, solo el PDF
  form.append('fileName', fileName);
  form.append('file', fs.createReadStream(pdfPath), { filename: fileName });

  const url = `${MEDIA_URL}/waInstance${ID_INSTANCE}/sendFileByUpload/${API_TOKEN}`;
  await axios.post(url, form, {
    headers: form.getHeaders(),
    maxContentLength: Infinity,
    maxBodyLength: Infinity,
    timeout: 120000,
  });
}

async function sendMessage(chatId, message) {
  const url = `${API_URL}/waInstance${ID_INSTANCE}/sendMessage/${API_TOKEN}`;
  await axios.post(url, { chatId, message }, { timeout: 20000 });
}

// ---------------------------------------------------------------------------
// Procesamiento de un mensaje entrante
// ---------------------------------------------------------------------------
async function handleNotification(body) {
  if (!body || body.typeWebhook !== 'incomingMessageReceived') return;

  const chatId = body.senderData && body.senderData.chatId;
  const md = body.messageData || {};

  // Solo el grupo permitido (Excel/Word de ESE grupo).
  if (ALLOWED_CHAT_ID && chatId !== ALLOWED_CHAT_ID) return;

  // Solo documentos.
  const fileData = md.fileMessageData || null;
  if (md.typeMessage !== 'documentMessage' || !fileData || !fileData.downloadUrl) return;

  const originalName = fileData.fileName || 'archivo';
  const caption = fileData.caption || '';

  // Cualquier otra cosa que no sea Excel/Word -> silencio.
  if (!isConvertible(originalName)) {
    log(`Ignorado (no es Excel/Word): "${originalName}"`);
    return;
  }

  log(`Archivo recibido en ${chatId}: "${originalName}" (caption: "${caption}"). Procesando...`);

  const stamp = `${Date.now()}_${process.hrtime()[1]}`;
  const ext = path.extname(originalName) || '.xlsx';
  const inputPath = path.join(TMP_DIR, `${stamp}${ext}`);
  const outputPath = path.join(TMP_DIR, `${stamp}.pdf`);
  const pdfName = pdfNameFrom(caption, originalName);

  try {
    await downloadFile(fileData.downloadUrl, inputPath);

    const { code, stderr } = await runConverter(inputPath, outputPath);

    if (code === EXIT_PASSWORD) {
      log(`Archivo con contraseña: "${originalName}". Avisando al grupo.`);
      await sendMessage(chatId, PASSWORD_MSG);
      return;
    }
    if (code === EXIT_UNSUPPORTED) {
      log(`Tipo no soportado por el conversor: "${originalName}". Silencio.`);
      return;
    }
    if (code !== EXIT_OK || !fs.existsSync(outputPath)) {
      log(`❌ Falló la conversión de "${originalName}" (code ${code}). ${stderr || ''}`.trim());
      return; // silencio ante errores no relacionados con contraseña
    }

    await sendPdf(chatId, outputPath, pdfName);
    log(`✅ PDF enviado a ${chatId} como "${pdfName}".`);
  } catch (err) {
    log('❌ Error procesando el archivo:', err.message);
  } finally {
    for (const p of [inputPath, outputPath]) {
      if (p) fs.rm(p, { force: true }, () => {});
    }
  }
}

// ---------------------------------------------------------------------------
// Loop principal
// ---------------------------------------------------------------------------
async function mainLoop() {
  log('Bot Excel/Word → PDF iniciado. Escuchando Green API...');
  if (ALLOWED_CHAT_ID) log(`Solo responderé en el grupo: ${ALLOWED_CHAT_ID}`);
  else log('⚠️  ALLOWED_CHAT_ID vacío: responderé en CUALQUIER chat (úsalo solo para descubrir el chatId del grupo).');

  // eslint-disable-next-line no-constant-condition
  while (true) {
    let notification;
    try {
      notification = await receiveNotification();
    } catch (err) {
      const status = err.response && err.response.status;
      log(`Error al recibir${status ? ` (HTTP ${status})` : ''}: ${err.message}. Reintento en 5s.`);
      await new Promise((r) => setTimeout(r, 5000));
      continue;
    }

    if (!notification || !notification.receiptId) continue;

    const { receiptId, body } = notification;
    try {
      await handleNotification(body);
    } catch (err) {
      log('Error inesperado:', err.message);
    } finally {
      try {
        await deleteNotification(receiptId);
      } catch (err) {
        log(`No se pudo borrar la notificación ${receiptId}: ${err.message}`);
      }
    }
  }
}

// Endpoint de salud opcional: si Render (u otro) define PORT, levantamos un
// mini servidor para poder correr como "Web Service". Como "Background Worker"
// no se define PORT y el bot solo escucha el grupo.
if (process.env.PORT) {
  require('http')
    .createServer((req, res) => {
      res.writeHead(200, { 'Content-Type': 'text/plain' });
      res.end('pdf-convertidor ok');
    })
    .listen(process.env.PORT, () => log(`Health server en puerto ${process.env.PORT}`));
}

checkConfig();
mainLoop().catch((err) => {
  console.error('Error fatal:', err);
  process.exit(1);
});
