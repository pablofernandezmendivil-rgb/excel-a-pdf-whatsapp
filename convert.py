#!/usr/bin/env python3
"""
Convierte un Excel o Word a PDF conservando formato e imagenes (LibreOffice UNO).
Excel, dos modos:
  - "primera" (el de siempre): la primera pestana VISIBLE, ajustada a una
    sola pagina (ancho y alto) si cabe sin encoger de mas.
  - "cliente": TODAS las pestanas visibles, en su orden, y las filas con texto
    ajustado crecen lo que haga falta para que no se corte (nunca se encogen).
    Ninguna pestana se quita: las ocultas se quedan ocultas (no se imprimen) y
    en el libro, para que las formulas que leen de ellas calculen.
  En los dos, una pestana oculta NUNCA sale en el PDF: ahi suelen ir costos,
  margenes y el proveedor. Si no se puede asegurar eso, en modo "cliente" no
  sale PDF (exit 1).
Word: se convierte tal cual.
Contrasena -> exit 10.

Uso: python3 convert.py <entrada> <salida.pdf> [--modo=cliente]
"""
import math
import os
import re
import sys
import time
import subprocess
import tempfile
import shutil

EXIT_OK = 0
EXIT_PASSWORD = 10
EXIT_UNSUPPORTED = 2
EXIT_ERROR = 1

SPREADSHEET_EXT = {".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv", ".xltx"}
WORD_EXT = {".doc", ".docx", ".docm", ".odt", ".rtf", ".dotx"}
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
OOXML_EXT = {".xlsx", ".xlsm", ".xltx", ".docx", ".docm", ".dotx", ".xlsb"}


def is_encrypted(path):
    try:
        import msoffcrypto
        with open(path, "rb") as f:
            try:
                return bool(msoffcrypto.OfficeFile(f).is_encrypted())
            except Exception:
                return False
    except Exception:
        try:
            if os.path.splitext(path)[1].lower() in OOXML_EXT:
                with open(path, "rb") as f:
                    return f.read(8) == OLE_MAGIC
        except Exception:
            pass
        return False


def _find_soffice():
    for c in ("libreoffice", "soffice", "/usr/bin/libreoffice", "/usr/bin/soffice"):
        if shutil.which(c) or os.path.exists(c):
            return c
    raise RuntimeError("No se encontro LibreOffice.")


MODOS = ("primera", "cliente")


def convert(input_path, output_path, modo="primera"):
    if modo not in MODOS:
        raise ValueError("modo desconocido: %r" % (modo,))
    import uno
    import unohelper
    from com.sun.star.beans import PropertyValue

    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SPREADSHEET_EXT and ext not in WORD_EXT:
        return EXIT_UNSUPPORTED
    if is_encrypted(input_path):
        return EXIT_PASSWORD

    pipe_name = "lo_conv_%d_%d" % (os.getpid(), int(time.time() * 1000) % 100000)
    profile_dir = tempfile.mkdtemp(prefix="lo-profile-")
    profile_url = unohelper.systemPathToFileUrl(profile_dir)
    soffice = _find_soffice()
    proc = subprocess.Popen(
        [soffice, "--headless", "--invisible", "--nologo", "--nofirststartwizard",
         "--norestore", "--nolockcheck", "-env:UserInstallation=" + profile_url,
         "--accept=pipe,name=%s;urp;StarOffice.ComponentContext" % pipe_name],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    doc = None
    desktop = None
    try:
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext)
        conn = "uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipe_name
        ctx = None
        last = None
        for _ in range(60):
            try:
                ctx = resolver.resolve(conn)
                break
            except Exception as e:
                last = e
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("No se pudo conectar con LibreOffice: %s" % last)
        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

        def prop(n, v):
            p = PropertyValue(); p.Name = n; p.Value = v; return p

        in_url = unohelper.systemPathToFileUrl(input_path)
        doc = desktop.loadComponentFromURL(in_url, "_blank", 0, (prop("Hidden", True), prop("ReadOnly", True)))
        if doc is None:
            raise RuntimeError("No se pudo abrir el documento.")
        is_calc = doc.supportsService("com.sun.star.sheet.SpreadsheetDocument")
        if is_calc:
            _prepare_spreadsheet(doc, modo)
            filter_name = "calc_pdf_Export"
        else:
            filter_name = "writer_pdf_Export"
        out_url = unohelper.systemPathToFileUrl(output_path)
        doc.storeToURL(out_url, (prop("FilterName", filter_name),))
        return EXIT_OK
    finally:
        try:
            if doc is not None:
                doc.close(False)
        except Exception:
            pass
        try:
            if desktop is not None:
                desktop.terminate()
        except Exception:
            pass
        try:
            proc.terminate(); proc.wait(timeout=10)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
        shutil.rmtree(profile_dir, ignore_errors=True)


# Si para meter todo en UNA hoja habria que encoger a menos de este % del
# tamano que ya cabe a lo ancho, mejor NO encoger tanto: usar varias hojas
# legibles. (1.0 = igual; 0.72 = se permite hasta ~28% de encogido extra.)
ONE_PAGE_FLOOR = 0.72
# Flags de contenido real (valor, fecha, texto, formula) para ignorar celdas
# con solo formato/color ("basura") al medir el area de datos.
CONTENT_FLAGS = 1 | 2 | 4 | 16


def _fit_width_only(ps):
    ps.setPropertyValue("ScaleToPagesX", 1)
    ps.setPropertyValue("ScaleToPagesY", 0)


def _fit_one_page(ps):
    try:
        ps.setPropertyValue("ScaleToPagesX", 1)
        ps.setPropertyValue("ScaleToPagesY", 1)
    except Exception:
        ps.setPropertyValue("ScaleToPages", 1)


def _data_bounds(sheet):
    """Ultima columna/fila con datos REALES (ignora celdas solo con formato)."""
    try:
        ranges = sheet.queryContentCells(CONTENT_FLAGS)
        addrs = ranges.RangeAddresses
        if not addrs:
            return None
        max_col = max(a.EndColumn for a in addrs)
        max_row = max(a.EndRow for a in addrs)
        return max_col, max_row
    except Exception:
        return None


# Banderas de queryContentCells para texto: cadenas y formulas (una formula
# puede dar texto). Sirven para hallar las celdas con texto ajustado.
TEXT_FLAGS = 4 | 16


def _es_visible(sheet, estricto):
    """IsVisible de la pestana.

    Si no se puede leer: en modo estricto truena (mejor no sacar PDF que sacar
    una pestana oculta); si no, se asume visible, que es lo que pasaba antes.
    """
    try:
        return bool(sheet.IsVisible)
    except Exception:
        if estricto:
            raise
        return True


# Letras que ocupan el doble de ancho (chino, japones, coreano, ancho completo).
ANCHAS = re.compile("[\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]")
# 1 punto tipografico en centesimas de milimetro (las unidades de UNO).
PUNTO = 35.28


def _alto_necesario(sheet, celda):
    """Lo que pide (en 1/100 mm) el texto ajustado de una celda, contado a mano.

    Caracteres por renglon segun el ancho de la columna, o del area combinada
    (LibreOffice no mide las combinadas), y el tamano de la letra. Se cuenta de
    mas a proposito: un renglon de sobra deja aire; uno de menos corta la
    descripcion. Si la celda combina varias filas, no se calcula (0).
    """
    texto = celda.getString()
    if not texto:
        return 0
    c0 = c1 = celda.CellAddress.Column
    try:
        cur = sheet.createCursorByRange(celda)
        cur.collapseToMergedArea()
        a = cur.getRangeAddress()
        if a.EndRow > a.StartRow:
            return 0
        c0, c1 = a.StartColumn, a.EndColumn
    except Exception:
        pass
    ancho = sum(sheet.Columns.getByIndex(c).Width for c in range(c0, c1 + 1)) - 100
    pts = float(getattr(celda, "CharHeight", 0) or 10)
    por_renglon = max(1, int(ancho / (pts * 0.5 * PUNTO)))
    renglones = 0
    for parrafo in texto.split("\n"):
        largo = len(parrafo) + len(ANCHAS.findall(parrafo))
        renglones += max(1, int(math.ceil(largo * 1.1 / por_renglon)))
    return int(renglones * pts * 1.25 * PUNTO + 80)


def _crece_filas(sheet):
    """Las filas con texto ajustado crecen hasta que el texto quepa.

    Dos cuentas y gana la mayor: la altura optima de LibreOffice y la de
    _alto_necesario. Nunca se encogen: una fila alta puede traer una foto, y
    la altura optima no ve las fotos. Las filas ocultas no se tocan.
    Devuelve cuantas filas crecieron.
    """
    necesita = {}
    try:
        celdas = sheet.queryContentCells(TEXT_FLAGS).getCells().createEnumeration()
        while celdas.hasMoreElements():
            c = celdas.nextElement()
            try:
                if c.IsTextWrapped:
                    r = c.CellAddress.Row
                    necesita[r] = max(necesita.get(r, 0), _alto_necesario(sheet, c))
            except Exception:
                pass
    except Exception:
        return 0
    rows = sheet.Rows
    crecieron = 0
    for r in sorted(necesita):
        try:
            row = rows.getByIndex(r)
            if not row.IsVisible:
                continue
            antes = row.Height
            # Un xlsx sin altos (openpyxl no los escribe) llega con las filas
            # marcadas como optimas pero sin calcular, y volver a poner True no
            # hace nada (CQ#303 de la prueba ciega). Se apaga y se prende.
            optima = antes
            try:
                row.OptimalHeight = False
                row.OptimalHeight = True
                optima = row.Height
            except Exception:
                pass
            final = max(antes, optima, necesita[r])
            if row.Height != final:
                row.Height = final
            if final > antes:
                crecieron += 1
        except Exception:
            pass
    return crecieron


def _ajusta_pagina(doc, sheet, idx, sola=True):
    """Area de impresion y escala de una pestana.

    `sola` = la pestana tiene su propio estilo de pagina. Si lo comparte con
    otra, la escala de una seria la de todas: ahi se usa lo seguro, todas las
    columnas a lo ancho de la hoja y hacia abajo las paginas que hagan falta.
    """
    ps = None
    try:
        import uno
        page_styles = doc.StyleFamilies.getByName("PageStyles")
        ps = page_styles.getByName(sheet.PageStyle)
        for pname, pval in (("PageScale", 0), ("ScaleToPages", 0)):
            try:
                ps.setPropertyValue(pname, pval)
            except Exception:
                pass

        bounds = _data_bounds(sheet)
        if not bounds:
            # sin datos: comportamiento simple (una hoja)
            if sola:
                _fit_one_page(ps)
            else:
                _fit_width_only(ps)
            return
        max_col, max_row = bounds

        # Recortar el area de impresion a los datos reales (quita "basura"
        # que inflaria el tamano y encogeria de mas).
        try:
            addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
            addr.Sheet = idx
            addr.StartColumn = 0
            addr.StartRow = 0
            addr.EndColumn = max_col
            addr.EndRow = max_row
            sheet.setPrintAreas((addr,))
        except Exception:
            pass

        if not sola:
            _fit_width_only(ps)
            return

        # Medir contenido vs area imprimible de la hoja (1/100 mm).
        cols = sheet.Columns
        rows = sheet.Rows
        content_w = sum(cols.getByIndex(c).Width for c in range(max_col + 1))
        content_h = sum(rows.getByIndex(r).Height for r in range(max_row + 1))
        printable_w = ps.Width - ps.LeftMargin - ps.RightMargin
        printable_h = ps.Height - ps.TopMargin - ps.BottomMargin

        if content_w <= 0 or content_h <= 0 or printable_w <= 0 or printable_h <= 0:
            _fit_one_page(ps)
            return

        # Escala para que las columnas quepan a lo ancho (nunca agrandar).
        s_width = min(1.0, printable_w / content_w)
        # Escala para meter TODO en una sola hoja.
        s_one = min(1.0, printable_w / content_w, printable_h / content_h)
        ratio = s_one / s_width if s_width > 0 else 0.0

        if ratio >= ONE_PAGE_FLOOR:
            # Cabe en una hoja con un ajuste leve -> una sola hoja limpia.
            _fit_one_page(ps)
        else:
            # Encogeria demasiado -> ancho a una hoja, alto en varias (legible).
            _fit_width_only(ps)
    except Exception:
        # Ante cualquier problema, comportamiento seguro.
        try:
            if sola:
                _fit_one_page(ps)
            else:
                _fit_width_only(ps)
        except Exception:
            pass


def _prepare_spreadsheet(doc, modo="primera"):
    sheets = doc.Sheets
    names = list(sheets.ElementNames)
    if not names:
        return
    cliente = (modo == "cliente")
    visibles = [n for n in names if _es_visible(sheets.getByName(n), cliente)]

    if not cliente:
        # La primera pestana VISIBLE. Antes era la primera a secas, y si estaba
        # oculta (costos, margen, proveedor) salia ESA en el PDF.
        first = visibles[0] if visibles else names[0]
        for n in names:
            if n != first:
                try:
                    sheets.removeByName(n)
                except Exception:
                    pass
        _ajusta_pagina(doc, sheets.getByName(first), 0, True)
        return

    # Modo cliente: todas las visibles, en su orden. Ninguna pestana se QUITA:
    # la visible puede sacar sus precios de una oculta (costos en una, precios
    # en otra), y quitarla los vuelve #REF! (CQ#305 de la prueba ciega). Las
    # ocultas se quedan ocultas: LibreOffice no imprime ni exporta las ocultas.
    if not visibles:
        raise RuntimeError("el libro no tiene ninguna pestana visible")
    estilos = {}
    for n in visibles:
        estilos.setdefault(sheets.getByName(n).PageStyle, []).append(n)
    for n in visibles:
        sheet = sheets.getByName(n)
        _crece_filas(sheet)
        _ajusta_pagina(doc, sheet, names.index(n), len(estilos[sheet.PageStyle]) == 1)


def _lee_modo(args):
    """`--modo=cliente` (o `--modo=primera`) despues de entrada y salida."""
    modo = "primera"
    for a in args:
        if a.startswith("--modo="):
            modo = a.split("=", 1)[1].strip().lower()
        else:
            raise ValueError("argumento desconocido: %r" % (a,))
    if modo not in MODOS:
        raise ValueError("modo desconocido: %r" % (modo,))
    return modo


def main():
    if len(sys.argv) not in (3, 4):
        sys.stderr.write("Uso: python3 convert.py <entrada> <salida.pdf> [--modo=cliente]\n")
        return EXIT_ERROR
    inp, outp = sys.argv[1], sys.argv[2]
    try:
        modo = _lee_modo(sys.argv[3:])
    except ValueError as e:
        sys.stderr.write("Error: %s\n" % e)
        return EXIT_ERROR
    if not os.path.exists(inp):
        sys.stderr.write("No existe: %s\n" % inp)
        return EXIT_ERROR
    try:
        code = convert(inp, outp, modo)
        if code == EXIT_OK and not os.path.exists(outp):
            return EXIT_ERROR
        return code
    except Exception as e:
        sys.stderr.write("Error: %s\n" % e)
        return EXIT_ERROR


if __name__ == "__main__":
    sys.exit(main())
