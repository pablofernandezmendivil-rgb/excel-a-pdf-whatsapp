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

app.listen(PORT, () => log(`Convertidor HTTP escuchando en el puerto ${PORT}`));
