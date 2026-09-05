'use strict';

// Servicio HTTP del convertidor (para Make).
//  POST /convert           -> devuelve el PDF
//  POST /convert-and-send  -> convierte y ENVÍA el PDF al grupo por Green API.
// Extra: si un Excel viene en chino y hay ANTHROPIC_API_KEY, traduce (chino->
// idioma destino) conservando formato e imágenes, y arma un PDF con la
// versión traducida ARRIBA y el original ABAJO.

require('dotenv').config();
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawn } = require('child_process');
const axios = require('axios');
const FormData = require('form-data');

const PORT = process.env.PORT || 3000;
const PYTHON_BIN = process.env.PYTHON_BIN || 'python3';
const CONVERTER = path.join(__dirname, 'convert.py');
const TRANSLATOR = path.join(__dirname, 'translate_xlsx.py');
const PASSWORD_MSG = process.env.PASSWORD_MSG || 'Error archivo con contraseña';
const AUTH_TOKEN = (process.env.CONVERT_AUTH_TOKEN || '').trim();

const GA_ID = process.env.GREENAPI_ID_INSTANCE;
const GA_TOKEN = process.env.GREENAPI_API_TOKEN;
const GA_API = (process.env.GREENAPI_API_URL || 'https://api.green-api.com').replace(/\/$/, '');
const GA_MEDIA = (process.env.GREENAPI_MEDIA_URL || 'https://media.green-api.com').replace(/\/$/, '');
const DEFAULT_CHAT_ID = (process.env.DEFAULT_CHAT_ID || '').trim();

// Traducción activa solo si hay llave de IA.
const TRANSLATE_ENABLED = !!(process.env.ANTHROPIC_API_KEY || '').trim();

const TMP = path.join(os.tmpdir(), 'convertidor');
fs.mkdirSync(TMP, { recursive: true });

const EXCEL_EXT = ['.xlsx', '.xls', '.xlsm', '.xlsb', '.ods', '.csv', '.xltx'];
const WORD_EXT = ['.doc', '.docx', '.docm', '.odt', '.rtf', '.dotx'];
const OK_EXT = new Set([...EXCEL_EXT, ...WORD_EXT]);
// Solo estos formatos (OOXML) se pueden traducir conservando todo.
const TRANSLATABLE_EXT = new Set(['.xlsx', '.xlsm', '.xltx']);

const EXIT_OK = 0;
const EXIT_PASSWORD = 10;
const EXIT_UNSUPPORTED = 2;

const upload = multer({ dest: TMP, limits: { fileSize: 60 * 1024 * 1024 } });

function log(...a) {
  console.log(`[${new Date().toISOString()}]`, ...a);
}

function pdfNameFrom(caption, originalName) {
  let base = (caption || '').trim();
  if (!base) base = path.basename(originalName || 'archivo', path.extname(originalName || ''));
  base = base.replace(/\.(pdf|xlsx?|xlsm|xlsb|ods|csv|docx?|docm|odt|rtf)$/i, '');
  base = base.replace(/[\\/:*?"<>|\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!base) base = 'archivo';
  return base.slice(0, 100) + '.pdf';
}

function run(cmd, args) {
  return new Promise((resolve) => {
    const proc = spawn(cmd, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let stderr = '';
    proc.stderr.on('data', (d) => (stderr += d.toString()));
    proc.on('error', (err) => resolve({ code: -1, stderr: err.message }));
    proc.on('close', (code) => resolve({ code, stderr }));
  });
}

const runConverter = (inp, outp) => run(PYTHON_BIN, [CONVERTER, inp, outp]);
const runTranslator = (inp, outp) => run(PYTHON_BIN, [TRANSLATOR, inp, outp]);
const runPdfUnite = (a, b, outp) => run('pdfunite', [a, b, outp]);

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
    maxContentLength: Infinity, maxBodyLength: Infinity, timeout: 120000,
  });
}

async function waSendText(chatId, message) {
  const url = `${GA_API}/waInstance${GA_ID}/sendMessage/${GA_TOKEN}`;
  await axios.post(url, { chatId, message }, { timeout: 20000 });
}

/**
 * Construye el PDF final a partir del archivo de entrada.
 * Devuelve { status: 'ok'|'password'|'unsupported'|'error', pdfPath, extras: [] }
 * Si es Excel en chino (y hay IA), el PDF = traducido (arriba) + original (abajo).
 */
async function buildPdf(inputPath, ext, stamp) {
  const origPdf = path.join(TMP, `${stamp}_orig.pdf`);
  const extras = [origPdf];

  const conv = await runConverter(inputPath, origPdf);
  if (conv.code === EXIT_PASSWORD) return { status: 'password', extras };
  if (conv.code === EXIT_UNSUPPORTED) return { status: 'unsupported', extras };
  if (conv.code !== EXIT_OK || !fs.existsSync(origPdf)) {
    log(`Conversión falló (code ${conv.code}): ${conv.stderr || ''}`.trim());
    return { status: 'error', extras };
  }

  // ¿Traducir? Solo Excel OOXML + IA disponible.
  if (TRANSLATE_ENABLED && TRANSLATABLE_EXT.has(ext)) {
    const transXlsx = path.join(TMP, `${stamp}_en${ext}`);
    const transPdf = path.join(TMP, `${stamp}_en.pdf`);
    const finalPdf = path.join(TMP, `${stamp}_final.pdf`);
    extras.push(transXlsx, transPdf, finalPdf);
    try {
      const tr = await runTranslator(inputPath, transXlsx);
      if (tr.code === 0 && fs.existsSync(transXlsx)) {
        const c2 = await runConverter(transXlsx, transPdf);
        if (c2.code === EXIT_OK && fs.existsSync(transPdf)) {
          const u = await runPdfUnite(transPdf, origPdf, finalPdf);
          if (u.code === 0 && fs.existsSync(finalPdf)) {
            return { status: 'ok', pdfPath: finalPdf, extras };
          }
        }
      } else if (tr.code === 2) {
        // no hay chino -> PDF normal
      } else {
        log(`Traducción no aplicada (code ${tr.code}): ${tr.stderr || ''}`.trim());
      }
    } catch (e) {
      log('Error en traducción (se envía original):', e.message);
    }
  }

  return { status: 'ok', pdfPath: origPdf, extras };
}

const app = express();
app.use(express.json({ limit: '2mb' }));

app.get('/', (req, res) => res.status(200).send('convertidor ok'));
app.get('/health', (req, res) => res.status(200).json({ ok: true, translate: TRANSLATE_ENABLED }));

// Devuelve el PDF (Make maneja el envío). Incluye traducción si aplica.
app.post('/convert', upload.single('file'), async (req, res) => {
  if (AUTH_TOKEN && (req.headers['x-auth-token'] || '') !== AUTH_TOKEN) {
    return res.status(401).json({ ok: false, reason: 'unauthorized' });
  }
  const caption = req.body.caption || req.query.caption || '';
  const originalName = req.body.fileName || req.query.fileName || (req.file && req.file.originalname) || 'archivo.xlsx';
  const ext = path.extname(originalName).toLowerCase() || '.xlsx';
  const stamp = `${Date.now()}_${process.hrtime()[1]}`;
  let inputPath;
  const cleanup = [];
  try {
    if (req.file) { inputPath = req.file.path + ext; fs.renameSync(req.file.path, inputPath); }
    else if (req.body.fileUrl) { inputPath = path.join(TMP, `${stamp}${ext}`); await downloadToFile(req.body.fileUrl, inputPath); }
    else return res.status(400).json({ ok: false, reason: 'no_file' });
    cleanup.push(inputPath);
    if (!OK_EXT.has(ext)) { cleanup.forEach((p) => fs.rm(p, { force: true }, () => {})); return res.status(200).json({ ok: false, reason: 'unsupported' }); }

    const r = await buildPdf(inputPath, ext, stamp);
    cleanup.push(...r.extras);
    const done = () => cleanup.forEach((p) => fs.rm(p, { force: true }, () => {}));
    if (r.status === 'password') { done(); return res.status(200).json({ ok: false, reason: 'password', message: PASSWORD_MSG }); }
    if (r.status !== 'ok' || !r.pdfPath) { done(); return res.status(r.status === 'unsupported' ? 200 : 500).json({ ok: false, reason: r.status }); }
    const pdfName = pdfNameFrom(caption, originalName);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('X-Pdf-Filename', encodeURIComponent(pdfName));
    const asciiName = (pdfName.replace(/[^\x20-\x7E]/g, '_').replace(/"/g, '') || 'archivo.pdf');
    res.setHeader('Content-Disposition',
      `attachment; filename="${asciiName}"; filename*=UTF-8''${encodeURIComponent(pdfName)}`);
    const stream = fs.createReadStream(r.pdfPath);
    stream.on('close', done);
    stream.on('error', () => { done(); if (!res.headersSent) res.status(500).end(); });
    stream.pipe(res);
  } catch (err) {
    cleanup.forEach((p) => fs.rm(p, { force: true }, () => {}));
    log('Error /convert:', err.message);
    if (!res.headersSent) res.status(500).json({ ok: false, reason: 'error', message: err.message });
  }
});

// Todo-en-uno para Make: convierte (y traduce si aplica) y ENVÍA el PDF al grupo.
app.post('/convert-and-send', async (req, res) => {
  if (AUTH_TOKEN && (req.headers['x-auth-token'] || '') !== AUTH_TOKEN) {
    return res.status(401).json({ ok: false, reason: 'unauthorized' });
  }
  if (!GA_ID || !GA_TOKEN) return res.status(500).json({ ok: false, reason: 'greenapi_no_configurado' });

  const url = req.body.fileUrl || req.body.downloadUrl;
  const chatId = (req.body.chatId || DEFAULT_CHAT_ID || '').trim();
  const caption = req.body.caption || '';
  const originalName = req.body.fileName || 'archivo.xlsx';
  const ext = path.extname(originalName).toLowerCase() || '.xlsx';
  if (!url || !chatId) return res.status(400).json({ ok: false, reason: 'faltan_datos' });
  if (!OK_EXT.has(ext)) return res.status(200).json({ ok: false, reason: 'unsupported' });

  const stamp = `${Date.now()}_${process.hrtime()[1]}`;
  const inputPath = path.join(TMP, `${stamp}${ext}`);
  const cleanup = [inputPath];
  const done = () => cleanup.forEach((p) => fs.rm(p, { force: true }, () => {}));

  try {
    await downloadToFile(url, inputPath);
    const r = await buildPdf(inputPath, ext, stamp);
    cleanup.push(...r.extras);

    if (r.status === 'password') { await waSendText(chatId, PASSWORD_MSG); done(); return res.status(200).json({ ok: false, reason: 'password' }); }
    if (r.status !== 'ok' || !r.pdfPath) { done(); return res.status(200).json({ ok: false, reason: r.status }); }

    const pdfName = pdfNameFrom(caption, originalName);
    await waSendPdf(chatId, r.pdfPath, pdfName);
    done();
    log(`✅ Enviado a ${chatId}: "${pdfName}"`);
    return res.status(200).json({ ok: true, pdfName });
  } catch (err) {
    done();
    log('Error convert-and-send:', err.message);
    return res.status(500).json({ ok: false, reason: 'error', message: err.message });
  }
});

app.listen(PORT, () => log(`Convertidor HTTP escuchando en el puerto ${PORT} (traducción: ${TRANSLATE_ENABLED ? 'ON' : 'OFF'})`));
