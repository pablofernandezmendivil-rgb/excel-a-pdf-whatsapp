#!/usr/bin/env python3
"""
Traduce el texto de un .xlsx de chino a otro idioma (por defecto ingles),
CONSERVANDO 100% del formato, colores e imagenes: reemplaza solo el texto
dentro de los nodos <t> (en sharedStrings.xml y en las hojas) y deja el
resto del archivo intacto byte por byte.

Uso:
  python3 translate_xlsx.py <entrada.xlsx> <salida.xlsx> [--mock]

Codigos de salida:
  0  -> se detecto chino y se tradujo (salida.xlsx creado)
  2  -> no hay chino (no se creo salida; el llamador usa el original)
  1  -> error

Requiere ANTHROPIC_API_KEY (salvo --mock). Config opcional:
  TRANSLATE_TARGET_LANG (default "English"), CLAUDE_MODEL (default haiku)
"""
import os
import re
import sys
import json
import zipfile
import html
import urllib.request
from xml.sax.saxutils import escape

CJK = re.compile(r"[㐀-䶿一-鿿豈-﫿]")
T_RE = re.compile(r"(<t\b[^>]*>)(.*?)(</t>)", re.DOTALL)
TARGET_LANG = os.environ.get("TRANSLATE_TARGET_LANG", "English")
MODEL = os.environ.get("CLAUDE_MODEL", "claude-haiku-4-5-20251001")
BATCH = 50


def has_chinese(text):
    return bool(text) and bool(CJK.search(text))


def anthropic_translate(strings, mock=False):
    if mock:
        return ["EN:" + s for s in strings]
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("Falta ANTHROPIC_API_KEY")
    out = []
    for i in range(0, len(strings), BATCH):
        chunk = strings[i:i + BATCH]
        numbered = {str(j): chunk[j] for j in range(len(chunk))}
        prompt = (
            "You are a professional translator for commercial price quotations. "
            "Translate each value from Chinese to " + TARGET_LANG + ". "
            "Keep numbers, product codes, units and any non-Chinese text unchanged. "
            "Return ONLY a JSON object mapping the same keys to the translated "
            "strings, no markdown, no explanation.\n\nInput JSON:\n"
            + json.dumps(numbered, ensure_ascii=False)
        )
        body = json.dumps({
            "model": MODEL,
            "max_tokens": 4000,
            "messages": [{"role": "user", "content": prompt}],
        }).encode("utf-8")
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages", data=body,
            headers={"x-api-key": api_key, "anthropic-version": "2023-06-01",
                     "content-type": "application/json"}, method="POST")
        with urllib.request.urlopen(req, timeout=120) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        text = "".join(p.get("text", "") for p in data.get("content", [])).strip()
        if text.startswith("```"):
            text = text.strip("`")
            if text.lstrip().lower().startswith("json"):
                text = text.lstrip()[4:]
        mapping = json.loads(text)
        for j in range(len(chunk)):
            out.append(mapping.get(str(j), chunk[j]))
    return out


def main(inp, outp, mock=False):
    with zipfile.ZipFile(inp) as z:
        names = z.namelist()
        parts = [n for n in names
                 if n == "xl/sharedStrings.xml" or re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
        contents = {n: z.read(n).decode("utf-8") for n in parts}

    # 1) recolectar textos en chino (unicos, preservando orden)
    found = []
    for content in contents.values():
        for m in T_RE.finditer(content):
            inner = html.unescape(m.group(2))
            if has_chinese(inner):
                found.append(inner)
    uniq = list(dict.fromkeys(found))
    if not uniq:
        return 2

    # 2) traducir
    translated = anthropic_translate(uniq, mock=mock)
    tmap = dict(zip(uniq, translated))

    # 3) sustituir en cada parte
    def repl(m):
        inner = html.unescape(m.group(2))
        if has_chinese(inner) and inner in tmap:
            return m.group(1) + escape(tmap[inner]) + m.group(3)
        return m.group(0)

    new_contents = {n: T_RE.sub(repl, c) for n, c in contents.items()}

    # 4) reescribir el zip (solo cambian las partes de texto)
    with zipfile.ZipFile(inp) as zin, zipfile.ZipFile(outp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in new_contents:
                data = new_contents[item.filename].encode("utf-8")
            zout.writestr(item, data)
    return 0


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    mock = "--mock" in sys.argv or bool(os.environ.get("TRANSLATE_MOCK"))
    if len(args) != 2:
        sys.stderr.write("Uso: python3 translate_xlsx.py <entrada.xlsx> <salida.xlsx> [--mock]\n")
        sys.exit(1)
    try:
        sys.exit(main(args[0], args[1], mock=mock))
    except Exception as e:
        sys.stderr.write("Error traduccion: %s\n" % e)
        sys.exit(1)
