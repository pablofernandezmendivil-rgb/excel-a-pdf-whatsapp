'use strict';

// Servicio HTTP del convertidor, para ser llamado desde Make.
// Recibe un Excel/Word (multipart 'file' o JSON {fileUrl}) y devuelve el PDF.
// El envío a WhatsApp y el guardado en Drive los hace Make.

require('dotenv').config();
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');
const axios = require('axios');

const PORT = process.env.PORT || 3000;
const PYTHON_BIN = process.env.PYTHON_BIN || 'python3';
const CONVERTER = path.join(__dirname, 'convert.py');
const PASSWORD_MSG = process.env.PASSWORD_MSG || 'Error archivo con contraseña';

// Token opcional para proteger el endpoint (Make lo manda en header x-auth-token).
const AUTH_TOKEN = (process.env.CONVERT_AUTH_TOKEN || '').trim();

// --- Green API (para que el convertidor pueda ENVIAR el PDF al grupo) ---
// Enviar NO choca con el webhook de Make (eso solo aplica a recibir).
const FormData = require('form-data');
const GA_ID = process.env.GREENAPI_ID_INSTANCE;
const GA_TOKEN = process.env.GREENAPI_API_TOKEN;
const GA_API = (process.env.GREENAPI_API_URL || 'https://api.green-api.com').replace(/\/$/, '');
const GA_MEDIA = (process.env.GREENAPI_MEDIA_URL || 'https://media.green-api.com').replace(/\/$/, '');
const DEFAULT_CHAT_ID = (process.env.DEFAULT_CHAT_ID || '').trim();

const TMP = path.join(os.tmpdir(), 'convertidor');
fs.mkdirSync(TMP, { recursive: true });

const EXCEL_EXT = ['.xlsx', '.xls', '.xlsm', '.xlsb', '.ods', '.csv', '.xltx'];
const WORD_EXT = ['.doc', '.docx', '.docm', '.odt', '.rtf', '.dotx'];
const OK_EXT = new Set([...EXCEL_EXT, ...WORD_EXT]);

const EXIT_OK = 0;
const EXIT_PASSWORD = 10;
const EXIT_UNSUPPORTED = 2;

const upload = multer({ dest: TMP, limits: { fileSize: 60 * 1024 * 1024 } });

function log(...a) {
  console.log(`[${new Date().toISOString()}]`, ...a);
}

// Nombre del PDF: caption si existe; si no, el nombre original. Conserva el nombre.
function pdfNameFrom(caption, originalName) {
  let base = (caption || '').trim();
  if (!base) base = path.basename(originalName || 'archivo', path.extname(originalName || ''));
  base = base.replace(/\.(pdf|xlsx?|xlsm|xlsb|ods|csv|docx?|docm|odt|rtf)$/i, '');
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

async function downloadToFile(url, destPath) {
  const r = await axios.get(url, { responseType: 'arraybuffer', timeout: 60000 });
  fs.writeFileSync(destPath, Buffer.from(r.data));
}

async function waSendPdf(chatId, pdfPath, fileName) {
  const form = new FormData();
  form.append('chatId', chatId);
  form.append('caption', '');
  form.append('fileName', fileName);
  form.append('file', fs.createReadStream(pdfPath), { filename: fileName });
  const url = `${GA_MEDIA}/waInstance${GA_ID}/sendFileByUpload/${GA_TOKEN}`;
  await axios.post(url, form, {
    headers: form.getHeaders(),
    maxContentLength: Infinity,
    maxBodyLength: Infinity,
    timeout: 120000,
  });
}

async function waSendText(chatId, message) {
  const url = `${GA_API}/waInstance${GA_ID}/sendMessage/${GA_TOKEN}`;
  await axios.post(url, { chatId, message }, { timeout: 20000 });
}

const app = express();
app.use(express.json({ limit: '2mb' }));

app.get('/', (req, res) => res.status(200).send('convertidor ok'));
app.get('/health', (req, res) => res.status(200).json({ ok: true }));

app.post('/convert', upload.single('file'), async (req, res) => {
  // Auth opcional
  if (AUTH_TOKEN && (req.headers['x-auth-token'] || '') !== AUTH_TOKEN) {
    return res.status(401).json({ ok: false, reason: 'unauthorized' });
  }

  const caption = req.body.caption || req.query.caption || '';
  const originalName =
    req.body.fileName || req.query.fileName || (req.file && req.file.originalname) || 'archivo.xlsx';
  const ext = path.extname(originalName).toLowerCase() || '.xlsx';

  const cleanup = [];
  const doCleanup = () => cleanup.forEach((p) => fs.rm(p, { force: true }, () => {}));

  try {
    let inputPath;
    if (req.file) {
      inputPath = req.file.path + ext;
      fs.renameSync(req.file.path, inputPath);
    } else if (req.body.fileUrl) {
      inputPath = path.join(TMP, `${Date.now()}${ext}`);
      const r = await axios.get(req.body.fileUrl, { responseType: 'arraybuffer', timeout: 60000 });
      fs.writeFileSync(inputPath, Buffer.from(r.data));
    } else {
      return res.status(400).json({ ok: false, reason: 'no_file' });
    }
    cleanup.push(inputPath);

    if (!OK_EXT.has(ext)) {
      doCleanup();
      return res.status(200).json({ ok: false, reason: 'unsupported' });
    }

    const outputPath = inputPath + '.pdf';
    cleanup.push(outputPath);

    const { code, stderr } = await runConverter(inputPath, outputPath);

    if (code === EXIT_PASSWORD) {
      doCleanup();
      log(`Archivo con contraseña: "${originalName}"`);
      return res.status(200).json({ ok: false, reason: 'password', message: PASSWORD_MSG });
    }
    if (code === EXIT_UNSUPPORTED) {
      doCleanup();
      return res.status(200).json({ ok: false, reason: 'unsupported' });
    }
    if (code !== EXIT_OK || !fs.existsSync(outputPath)) {
      doCleanup();
      log(`Error de conversión "${originalName}" (code ${code}): ${stderr || ''}`.trim());
      return res.status(500).json({ ok: false, reason: 'error' });
    }

    const pdfName = pdfNameFrom(caption, originalName);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('X-Pdf-Filename', encodeURIComponent(pdfName));
    res.setHeader('Content-Disposition', `attachment; filename="${pdfName.replace(/"/g, '')}"`);
    log(`OK "${originalName}" -> "${pdfName}"`);

    const stream = fs.createReadStream(outputPath);
    stream.on('close', doCleanup);
    stream.on('error', () => {
      doCleanup();
      if (!res.headersSent) res.status(500).json({ ok: false, reason: 'stream' });
    });
    stream.pipe(res);
  } catch (err) {
    doCleanup();
    log('Error inesperado:', err.message);
    if (!res.headersSent) res.status(500).json({ ok: false, reason: 'error', message: err.message });
  }
});

// Todo-en-uno para Make: recibe la info del archivo, convierte y ENVÍA el PDF
// al grupo por Green API. Make solo necesita UN módulo HTTP que llame aquí.
app.post('/convert-and-send', async (req, res) => {
  if (AUTH_TOKEN && (req.headers['x-auth-token'] || '') !== AUTH_TOKEN) {
    return res.status(401).json({ ok: false, reason: 'unauthorized' });
  }
  if (!GA_ID || !GA_TOKEN) {
    return res.status(500).json({ ok: false, reason: 'greenapi_no_configurado' });
  }

  const url = req.body.fileUrl || req.body.downloadUrl;
  const chatId = (req.body.chatId || DEFAULT_CHAT_ID || '').trim();
  const caption = req.body.caption || '';
  const originalName = req.body.fileName || 'archivo.xlsx';
  const ext = path.extname(originalName).toLowerCase() || '.xlsx';

  if (!url || !chatId) return res.status(400).json({ ok: false, reason: 'faltan_datos' });

  // Cualquier cosa que no sea Excel/Word -> silencio (no se envía nada).
  if (!OK_EXT.has(ext)) return res.status(200).json({ ok: false, reason: 'unsupported' });

  const inputPath = path.join(TMP, `${Date.now()}${ext}`);
  const outputPath = inputPath + '.pdf';
  const cleanup = () => [inputPath, outputPath].forEach((p) => fs.rm(p, { force: true }, () => {}));

  try {
    await downloadToFile(url, inputPath);
    const { code, stderr } = await runConverter(inputPath, outputPath);

    if (code === EXIT_PASSWORD) {
      await waSendText(chatId, PASSWORD_MSG);
      cleanup();
      return res.status(200).json({ ok: false, reason: 'password' });
    }
    if (code !== EXIT_OK || !fs.existsSync(outputPath)) {
      cleanup();
      log(`Error conversión "${originalName}" (code ${code}): ${stderr || ''}`.trim());
      return res.status(200).json({ ok: false, reason: 'error' });
    }

    const pdfName = pdfNameFrom(caption, originalName);
    await waSendPdf(chatId, outputPath, pdfName);
    cleanup();
    log(`✅ Enviado a ${chatId}: "${pdfName}"`);
    return res.status(200).json({ ok: true, pdfName });
  } catch (err) {
    cleanup();
    log('Error convert-and-send:', err.message);
    return res.status(500).json({ ok: false, reason: 'error', message: err.message });
  }
});

app.listen(PORT, () => log(`Convertidor HTTP escuchando en el puerto ${PORT}`));
