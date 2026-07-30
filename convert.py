#!/usr/bin/env python3
"""
Convierte un Excel o Word a PDF conservando formato e imagenes, usando el
motor interno de LibreOffice (UNO).

Reglas:
  - Excel: se conserva SOLO la primera pestana (sin importar su nombre) y se
    ajustan todas las columnas al ancho de UNA pagina (fit-to-width = 1),
    respetando la orientacion original del archivo.
  - Word: se convierte tal cual.
  - Si el archivo esta protegido con contrasena -> exit code 10.

Uso:
  python3 convert.py <archivo_entrada> <archivo_salida.pdf>

Codigos de salida:
  0  -> PDF generado correctamente
  10 -> archivo protegido con contrasena
  2  -> tipo de archivo no soportado
  1  -> cualquier otro error
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
    """Detecta si un archivo OOXML/OLE esta protegido con contrasena.

    Usa msoffcrypto si esta disponible; si no, cae a una heuristica por firma
    de bytes (un .xlsx/.docx cifrado es un contenedor OLE, no un ZIP)."""
    try:
        import msoffcrypto
        with open(path, "rb") as f:
            try:
                office = msoffcrypto.OfficeFile(f)
                return bool(office.is_encrypted())
            except Exception:
                return False
    except Exception:
        # Respaldo sin dependencias: un OOXML normal empieza con "PK" (ZIP);
        # si empieza con la firma OLE, esta cifrado.
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in OOXML_EXT:
                with open(path, "rb") as f:
                    return f.read(8) == OLE_MAGIC
        except Exception:
            pass
        return False


def _find_soffice():
    for c in ("libreoffice", "soffice",
              "/usr/bin/libreoffice", "/usr/bin/soffice"):
        if shutil.which(c) or os.path.exists(c):
            return c
    raise RuntimeError("No se encontro LibreOffice (libreoffice/soffice).")


def convert(input_path, output_path):
    import uno  # noqa
    import unohelper  # noqa
    from com.sun.star.beans import PropertyValue

    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    ext = os.path.splitext(input_path)[1].lower()

    if ext not in SPREADSHEET_EXT and ext not in WORD_EXT:
        return EXIT_UNSUPPORTED

    if is_encrypted(input_path):
        return EXIT_PASSWORD

    # ---- Lanzar una instancia headless de LibreOffice con un pipe unico ----
    pipe_name = "lo_conv_%d_%d" % (os.getpid(), int(time.time() * 1000) % 100000)
    profile_dir = tempfile.mkdtemp(prefix="lo-profile-")
    profile_url = unohelper.systemPathToFileUrl(profile_dir)
    soffice = _find_soffice()

    proc = subprocess.Popen(
        [
            soffice, "--headless", "--invisible", "--nologo", "--nofirststartwizard",
            "--norestore", "--nolockcheck",
            "-env:UserInstallation=" + profile_url,
            "--accept=pipe,name=%s;urp;StarOffice.ComponentContext" % pipe_name,
        ],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )

    doc = None
    desktop = None
    try:
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext
        )
        conn = ("uno:pipe,name=%s;urp;StarOffice.ComponentContext" % pipe_name)

        ctx = None
        last_err = None
        for _ in range(60):  # ~30s de espera a que arranque soffice
            try:
                ctx = resolver.resolve(conn)
                break
            except Exception as e:
                last_err = e
                time.sleep(0.5)
        if ctx is None:
            raise RuntimeError("No se pudo conectar con LibreOffice: %s" % last_err)

        smgr = ctx.ServiceManager
        desktop = smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx
        )

        def prop(name, value):
            p = PropertyValue()
            p.Name = name
            p.Value = value
            return p

        in_url = unohelper.systemPathToFileUrl(input_path)
        load_props = (prop("Hidden", True), prop("ReadOnly", True))
        doc = desktop.loadComponentFromURL(in_url, "_blank", 0, load_props)
        if doc is None:
            # Suele pasar con archivos corruptos o (raro) protegidos
            raise RuntimeError("LibreOffice no pudo abrir el documento.")

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
            proc.terminate()
            proc.wait(timeout=10)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
        shutil.rmtree(profile_dir, ignore_errors=True)


def _prepare_spreadsheet(doc):
    """Deja solo la primera hoja y ajusta las columnas a 1 pagina de ancho."""
    sheets = doc.Sheets
    names = list(sheets.ElementNames)
    if not names:
        return
    first = names[0]

    # Eliminar todas las hojas menos la primera.
    for n in names[1:]:
        try:
            sheets.removeByName(n)
        except Exception:
            pass

    # Ajustar la hoja restante: todas las columnas en el ancho de 1 pagina,
    # alto ilimitado (0 = las paginas hacia abajo que hagan falta).
    try:
        sheet = sheets.getByName(first)
        style_name = sheet.PageStyle
        page_styles = doc.StyleFamilies.getByName("PageStyles")
        ps = page_styles.getByName(style_name)

        # Desactivar otras escalas para que no compitan.
        for pname, pval in (("PageScale", 0), ("ScaleToPages", 0)):
            try:
                ps.setPropertyValue(pname, pval)
            except Exception:
                pass

        applied = False
        try:
            ps.setPropertyValue("ScaleToPagesX", 1)  # 1 pagina de ancho
            ps.setPropertyValue("ScaleToPagesY", 0)  # alto ilimitado
            applied = True
        except Exception:
            applied = False

        if not applied:
            # Version de LibreOffice sin ScaleToPagesX: como respaldo,
            # forzamos todo el contenido a 1 pagina total.
            try:
                ps.setPropertyValue("ScaleToPages", 1)
            except Exception:
                pass
    except Exception:
        # Si algo falla en el ajuste, igual exportamos el PDF de la 1a hoja.
        pass


def main():
    if len(sys.argv) != 3:
        sys.stderr.write("Uso: python3 convert.py <entrada> <salida.pdf>\n")
        return EXIT_ERROR
    inp, outp = sys.argv[1], sys.argv[2]
    if not os.path.exists(inp):
        sys.stderr.write("No existe el archivo de entrada: %s\n" % inp)
        return EXIT_ERROR
    try:
        code = convert(inp, outp)
        if code == EXIT_OK and not os.path.exists(outp):
            sys.stderr.write("No se genero el PDF esperado.\n")
            return EXIT_ERROR
        return code
    except Exception as e:
        sys.stderr.write("Error en la conversion: %s\n" % e)
        return EXIT_ERROR


if __name__ == "__main__":
    sys.exit(main())
