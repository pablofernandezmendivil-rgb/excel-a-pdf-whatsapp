#!/usr/bin/env python3
"""
Convierte un Excel o Word a PDF conservando formato e imagenes (LibreOffice UNO).
Excel: solo la primera pestana, ajustada a UNA sola pagina (ancho y alto).
Word: se convierte tal cual.
Contrasena -> exit 10.

Uso: python3 convert.py <entrada> <salida.pdf>
"""
import os
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


def convert(input_path, output_path):
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
            _prepare_spreadsheet(doc)
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


def _prepare_spreadsheet(doc):
    sheets = doc.Sheets
    names = list(sheets.ElementNames)
    if not names:
        return
    first = names[0]
    for n in names[1:]:
        try:
            sheets.removeByName(n)
        except Exception:
            pass
    try:
        import uno
        sheet = sheets.getByName(first)
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
            _fit_one_page(ps)
            return
        max_col, max_row = bounds

        # Recortar el area de impresion a los datos reales (quita "basura"
        # que inflaria el tamano y encogeria de mas).
        try:
            addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
            addr.Sheet = 0
            addr.StartColumn = 0
            addr.StartRow = 0
            addr.EndColumn = max_col
            addr.EndRow = max_row
            sheet.setPrintAreas((addr,))
        except Exception:
            pass

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
        # Ante cualquier problema, comportamiento seguro: una sola hoja.
        try:
            _fit_one_page(ps)
        except Exception:
            pass


def main():
    if len(sys.argv) != 3:
        sys.stderr.write("Uso: python3 convert.py <entrada> <salida.pdf>\n")
        return EXIT_ERROR
    inp, outp = sys.argv[1], sys.argv[2]
    if not os.path.exists(inp):
        sys.stderr.write("No existe: %s\n" % inp)
        return EXIT_ERROR
    try:
        code = convert(inp, outp)
        if code == EXIT_OK and not os.path.exists(outp):
            return EXIT_ERROR
        return code
    except Exception as e:
        sys.stderr.write("Error: %s\n" % e)
        return EXIT_ERROR


if __name__ == "__main__":
    sys.exit(main())
