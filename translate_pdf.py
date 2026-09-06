#!/usr/bin/env python3
"""
Traduce un PDF en chino a ingles (o TRANSLATE_TARGET_LANG) usando IA de vision.

A diferencia del Excel, un PDF no tiene texto "editable", asi que NO se puede
conservar el diseno exacto. En su lugar:
  1) Se decide si el PDF esta en chino.
       - Si tiene capa de texto y NO hay chino -> exit 2 (no se hace nada).
       - Si tiene texto con chino -> se traduce.
       - Si NO tiene texto (escaneado) -> se le pregunta a la IA por la 1a pagina;
         si no es chino -> exit 2.
  2) Cada pagina se convierte a imagen y la IA la traduce a un fragmento HTML
     (conservando tablas, filas, columnas y numeros).
  3) Se arma un HTML y se convierte a PDF con LibreOffice.

Uso:
  python3 translate_pdf.py <entrada.pdf> <salida.pdf> [--mock]

Codigos de salida:
  0 -> se tradujo (salida.pdf creado)
  2 -> no es chino (no se crea salida; el llamador no hace nada)
  1 -> error

Requiere ANTHROPIC_API_KEY (salvo --mock). Config opcional:
  TRANSLATE_TARGET_LANG (default "English"), CLAUDE_MODEL (default sonnet).
Herramientas: pdftotext, pdftoppm (poppler) y libreoffice (ya instaladas).
"""
import os
import re
import sys
import json
import base64
import shutil
import tempfile
import subprocess
import urllib.request

CJK = re.compile(r"[㐀-䶿一-鿿豈-﫿]")
TARGET_LANG = os.environ.get("TRANSLATE_TARGET_LANG", "English")
MODEL = os.environ.get("CLAUDE_MODEL", "claude-sonnet-5")
MAX_PAGES = int(os.environ.get("PDF_MAX_PAGES", "15"))
DPI = int(os.environ.get("PDF_DPI", "150"))


def has_cjk(text):
    return bool(text) and bool(CJK.search(text))


def run(cmd):
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return p.returncode, p.stdout, p.stderr


def pdf_text(pdf):
    code, out, _ = run(["pdftotext", "-q", pdf, "-"])
    if code != 0:
        return ""
    try:
        return out.decode("utf-8", "ignore")
    except Exception:
        return ""


def render_pages(pdf, workdir):
    """Convierte el PDF a PNG (una por pagina). Devuelve lista de rutas."""
    prefix = os.path.join(workdir, "page")
    code, _, err = run(["pdftoppm", "-png", "-r", str(DPI), pdf, prefix])
    if code != 0:
        raise RuntimeError("pdftoppm fallo: " + err.decode("utf-8", "ignore"))
    files = sorted(
        [os.path.join(workdir, f) for f in os.listdir(workdir)
         if f.startswith("page") and f.endswith(".png")]
    )
    return files[:MAX_PAGES]


def anthropic_vision(png_path, prompt, mock=False, max_tokens=4000):
    if mock:
        return "<table><tr><td>EN (mock)</td></tr></table>"
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("Falta ANTHROPIC_API_KEY")
    with open(png_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    body = json.dumps({
        "model": MODEL,
        "max_tokens": max_tokens,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": "image/png", "data": b64}},
                {"type": "text", "text": prompt},
            ],
        }],
    }).encode("utf-8")
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages", data=body,
        headers={"x-api-key": api_key, "anthropic-version": "2023-06-01",
                 "content-type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=180) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return "".join(p.get("text", "") for p in data.get("content", [])).strip()


def is_chinese_image(png_path, mock=False):
    if mock:
        return True
    ans = anthropic_vision(
        png_path,
        "Does this document contain any Chinese characters? "
        "Answer with only one word: YES or NO.",
        mock=mock, max_tokens=10,
    ).upper()
    return "YES" in ans


def translate_image(png_path, mock=False):
    prompt = (
        "This is one page of a Chinese commercial document (often a price "
        "quotation). Translate ALL of its text into " + TARGET_LANG + ". "
        "Reproduce the content as a clean, self-contained HTML fragment: use a "
        "<table> with borders for tabular data, keeping the same rows, columns "
        "and order. Keep all numbers, codes, units and prices EXACTLY as shown. "
        "Do not add commentary. Output ONLY the HTML fragment (no markdown "
        "fences, no <html> or <body> tags)."
    )
    html = anthropic_vision(png_path, prompt, mock=mock, max_tokens=8000)
    # limpiar posibles ``` fences
    html = html.strip()
    if html.startswith("```"):
        html = html.strip("`")
        if html.lstrip().lower().startswith("html"):
            html = html.lstrip()[4:]
    return html.strip()


HTML_HEAD = (
    "<!DOCTYPE html><html><head><meta charset=\"utf-8\">"
    "<style>"
    "body{font-family:'Liberation Sans',Arial,sans-serif;font-size:10pt;color:#000;}"
    "table{border-collapse:collapse;width:100%;margin:0 0 8px 0;}"
    "td,th{border:1px solid #444;padding:3px 5px;font-size:9pt;vertical-align:top;}"
    "th{background:#eee;}"
    ".page{page-break-after:always;}"
    "h1,h2,h3{margin:6px 0;}"
    "</style></head><body>"
)
HTML_TAIL = "</body></html>"


def html_to_pdf(html_path, out_pdf):
    workdir = os.path.dirname(html_path)
    profile = tempfile.mkdtemp(prefix="lo-html-")
    profile_url = "file://" + profile
    soffice = None
    for c in ("libreoffice", "soffice", "/usr/bin/libreoffice", "/usr/bin/soffice"):
        if shutil.which(c) or os.path.exists(c):
            soffice = c
            break
    if not soffice:
        raise RuntimeError("No se encontro LibreOffice")
    code, _, err = run([
        soffice, "--headless", "--norestore", "--nologo", "--nolockcheck",
        "-env:UserInstallation=" + profile_url,
        "--convert-to", "pdf:writer_pdf_Export", "--outdir", workdir, html_path,
    ])
    produced = os.path.join(workdir, os.path.splitext(os.path.basename(html_path))[0] + ".pdf")
    if code != 0 or not os.path.exists(produced):
        raise RuntimeError("LibreOffice HTML->PDF fallo: " + err.decode("utf-8", "ignore"))
    if os.path.abspath(produced) != os.path.abspath(out_pdf):
        os.replace(produced, out_pdf)


def main(inp, outp, mock=False):
    # 1) decidir si es chino
    text = pdf_text(inp)
    if text.strip():
        if not has_cjk(text):
            return 2  # PDF digital sin chino -> no hacer nada
    workdir = tempfile.mkdtemp(prefix="pdftr-")
    pages = render_pages(inp, workdir)
    if not pages:
        return 1
    if not text.strip():
        # escaneado: preguntar a la IA por la 1a pagina
        if not is_chinese_image(pages[0], mock=mock):
            return 2

    # 2) traducir pagina por pagina
    parts = []
    for i, png in enumerate(pages):
        frag = translate_image(png, mock=mock)
        if not frag:
            continue
        # salto de pagina antes de cada pagina (menos la primera): LibreOffice
        # respeta page-break-before en linea de forma confiable.
        if parts:
            parts.append("<p style=\"page-break-before:always;margin:0;\"></p>")
        parts.append(frag)
    if not parts:
        return 1

    # 3) HTML -> PDF
    html_path = os.path.join(workdir, "translated.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(HTML_HEAD + "".join(parts) + HTML_TAIL)
    html_to_pdf(html_path, outp)
    return 0 if os.path.exists(outp) else 1


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    mock = "--mock" in sys.argv or bool(os.environ.get("TRANSLATE_MOCK"))
    if len(args) != 2:
        sys.stderr.write("Uso: python3 translate_pdf.py <entrada.pdf> <salida.pdf> [--mock]\n")
        sys.exit(1)
    try:
        sys.exit(main(args[0], args[1], mock=mock))
    except Exception as e:
        sys.stderr.write("Error traduccion PDF: %s\n" % e)
        sys.exit(1)
