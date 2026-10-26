# -*- coding: utf-8 -*-
"""Pruebas de convert.py sin LibreOffice: un libro de mentiras con la misma
forma que el de UNO (Sheets, IsVisible, PageStyle, Rows, queryContentCells...).

Lo que se prueba es la decision —que pestanas salen, que filas crecen, que
escala lleva cada pagina— no el PDF. El PDF de verdad se prueba contra el
servicio desplegado (los casos ciegos del 25-oct-26).

El modo normal (sin `modo`: el chat PDF convertidor y el generador de POs)
tiene que hacer EXACTAMENTE lo que hace 21da4e5, que es lo que esta vivo y a
Pablo le funciona («el convertidor no le muevas»). `MismoQue21da4e5` lo amarra:
corre los dos convert.py sobre el mismo libro y compara cada llamada UNO.

    python pruebas_convert.py
"""
import importlib.util
import os
import subprocess
import sys
import tempfile
import types
import unittest

# convert.py importa `uno` adentro de las funciones; aqui no hay LibreOffice.
uno = types.ModuleType("uno")
uno.createUnoStruct = lambda nombre: types.SimpleNamespace()
sys.modules.setdefault("uno", uno)

import convert  # noqa: E402

AQUI = os.path.dirname(os.path.abspath(__file__))
VIVO = "21da4e5"
DESCRIPCION = ("Cabecera tapizada en lino color arena para cama king size, estructura de madera de "
               "pino tratada, relleno de espuma de alta densidad. Tiempo de fabricación: 25 días "
               "hábiles a partir del anticipo.")


# ══════════════════════════════════════════════════════════════════════════════
# El libro de mentiras
# ══════════════════════════════════════════════════════════════════════════════

class Celda:
    def __init__(self, fila, ajustada, texto="", col=1, puntos=11, combinada=None):
        self.CellAddress = types.SimpleNamespace(Row=fila, Column=col)
        self.IsTextWrapped = ajustada
        self.texto, self.CharHeight = texto, puntos
        # (col0, col1, fila0, fila1) si esta combinada
        self.combinada = combinada or (col, col, fila, fila)

    def getString(self):
        return self.texto


class Enum:
    def __init__(self, cosas):
        self.cosas = list(cosas)

    def hasMoreElements(self):
        return bool(self.cosas)

    def nextElement(self):
        return self.cosas.pop(0)


class Acceso:
    def __init__(self, cosas):
        self.cosas = cosas

    def createEnumeration(self):
        return Enum(self.cosas)


class Rangos:
    def __init__(self, hoja):
        self.hoja = hoja
        fin = hoja.fin
        self.RangeAddresses = [types.SimpleNamespace(EndColumn=fin[0], EndRow=fin[1])] if fin else []

    def getCells(self):
        return Acceso(self.hoja.celdas)


class Cursor:
    def __init__(self, celda):
        self.celda = celda

    def collapseToMergedArea(self):
        pass

    def getRangeAddress(self):
        c0, c1, f0, f1 = self.celda.combinada
        return types.SimpleNamespace(StartColumn=c0, EndColumn=c1, StartRow=f0, EndRow=f1)


class Fila:
    """`optima` es lo que LibreOffice calcularia; None = no calcula nada (CQ#303)."""

    def __init__(self, alto, optima=None, visible=True):
        self._alto, self.optima, self.IsVisible = alto, optima, visible
        self.toques = 0

    @property
    def Height(self):
        return self._alto

    @Height.setter
    def Height(self, v):
        self.toques += 1
        self._alto = v

    @property
    def OptimalHeight(self):
        return False

    @OptimalHeight.setter
    def OptimalHeight(self, v):
        self.toques += 1
        if v and self.optima is not None:
            self._alto = self.optima


class Filas:
    def __init__(self, filas):
        self.filas = filas

    def getByIndex(self, i):
        return self.filas.setdefault(i, Fila(452))


class Columnas:
    def __init__(self, anchos):
        self.anchos = anchos

    def getByIndex(self, i):
        return types.SimpleNamespace(Width=self.anchos.get(i, 2000))


class Hoja:
    def __init__(self, nombre, visible=True, estilo="Default", celdas=(), filas=None,
                 fin=(5, 10), visible_truena=False, anchos=None):
        self.nombre, self._visible, self.PageStyle = nombre, visible, estilo
        self.celdas, self.fin, self.visible_truena = list(celdas), fin, visible_truena
        self.Rows = Filas(filas or {})
        self.Columns = Columnas(anchos or {})
        self.areas = None

    @property
    def IsVisible(self):
        if self.visible_truena:
            raise RuntimeError("UNO no contesto")
        return self._visible

    @IsVisible.setter
    def IsVisible(self, v):
        self._visible = v

    def queryContentCells(self, banderas):
        return Rangos(self)

    def createCursorByRange(self, celda):
        return Cursor(celda)

    def setPrintAreas(self, areas):
        self.areas = areas


class Hojas:
    def __init__(self, hojas, quitar_truena=()):
        self.hojas = list(hojas)
        self.quitar_truena = set(quitar_truena)
        self.quitadas = []

    @property
    def ElementNames(self):
        return tuple(h.nombre for h in self.hojas)

    def getByName(self, n):
        return next(h for h in self.hojas if h.nombre == n)

    def removeByName(self, n):
        if n in self.quitar_truena:
            raise RuntimeError("no se deja quitar")
        self.quitadas.append(n)
        self.hojas = [h for h in self.hojas if h.nombre != n]


class Estilo:
    def __init__(self):
        self.Width, self.Height = 21000, 29700
        self.LeftMargin = self.RightMargin = self.TopMargin = self.BottomMargin = 1000
        self.puesto = {}

    def setPropertyValue(self, k, v):
        self.puesto[k] = v


class Familia:
    def __init__(self, libro):
        self.libro = libro

    def getByName(self, n):
        return self.libro.estilos.setdefault(n, Estilo())


class Familias:
    def __init__(self, libro):
        self.libro = libro

    def getByName(self, n):
        assert n == "PageStyles"
        return Familia(self.libro)


class Libro:
    def __init__(self, hojas, quitar_truena=()):
        self.Sheets = Hojas(hojas, quitar_truena)
        self.estilos = {}
        self.StyleFamilies = Familias(self)

    def visibles(self):
        return tuple(h.nombre for h in self.Sheets.hojas if h._visible)


def ancho_solo(est):
    return est.puesto.get("ScaleToPagesX") == 1 and est.puesto.get("ScaleToPagesY") == 0


def una_hoja(est):
    return est.puesto.get("ScaleToPagesX") == 1 and est.puesto.get("ScaleToPagesY") == 1


# ══════════════════════════════════════════════════════════════════════════════
# Grabadora de llamadas UNO
# ══════════════════════════════════════════════════════════════════════════════

PRIMITIVOS = (str, int, float, bool, type(None))


def _texto(v):
    """Una representacion estable (sin direcciones de memoria) de lo que paso por UNO."""
    if isinstance(v, PRIMITIVOS):
        return repr(v)
    if isinstance(v, (list, tuple)):
        return "[" + ", ".join(_texto(x) for x in v) + "]"
    if isinstance(v, types.SimpleNamespace):
        return "{" + ", ".join("%s=%s" % (k, _texto(x)) for k, x in sorted(vars(v).items())) + "}"
    if isinstance(v, Graba):
        return "<%s>" % object.__getattribute__(v, "_r")
    return "<%s>" % type(v).__name__


def _crudo(v):
    if isinstance(v, Graba):
        return object.__getattribute__(v, "_o")
    if isinstance(v, tuple):
        return tuple(_crudo(x) for x in v)
    return v


def _envuelve(v, ruta, log):
    if isinstance(v, PRIMITIVOS) or (isinstance(v, tuple) and all(isinstance(x, PRIMITIVOS) for x in v)):
        return v
    if isinstance(v, list):
        return [_envuelve(x, "%s[%d]" % (ruta, i), log) for i, x in enumerate(v)]
    return Graba(v, ruta, log)


class Graba:
    """Envuelve el libro y apunta cada lectura, escritura y llamada, en orden."""

    def __init__(self, obj, ruta, log):
        object.__setattr__(self, "_o", obj)
        object.__setattr__(self, "_r", ruta)
        object.__setattr__(self, "_log", log)

    def __getattr__(self, n):
        o, r, log = (object.__getattribute__(self, k) for k in ("_o", "_r", "_log"))
        ruta = r + "." + n
        try:
            v = getattr(o, n)
        except Exception as e:
            log.append(("lee", ruta, "TRUENA " + type(e).__name__))
            raise
        if callable(v) and not isinstance(v, type):
            def llamada(*a):
                args = tuple(_crudo(x) for x in a)
                try:
                    res = v(*args)
                except Exception as e:
                    log.append(("llama", ruta, _texto(args), "TRUENA " + type(e).__name__))
                    raise
                log.append(("llama", ruta, _texto(args), _texto(res)))
                return _envuelve(res, ruta + "()", log)
            return llamada
        log.append(("lee", ruta, _texto(v)))
        return _envuelve(v, ruta, log)

    def __setattr__(self, n, v):
        o, r, log = (object.__getattribute__(self, k) for k in ("_o", "_r", "_log"))
        log.append(("pone", r + "." + n, _texto(v)))
        setattr(o, n, _crudo(v))


def convert_vivo():
    """El convert.py de 21da4e5, tal cual esta en git."""
    try:
        src = subprocess.run(["git", "show", VIVO + ":convert.py"], cwd=AQUI, capture_output=True,
                             check=True).stdout
    except Exception:
        return None
    d = tempfile.mkdtemp(prefix="convert-vivo-")
    ruta = os.path.join(d, "convert_vivo.py")
    with open(ruta, "wb") as f:
        f.write(src)
    spec = importlib.util.spec_from_file_location("convert_vivo", ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# ══════════════════════════════════════════════════════════════════════════════
# Las pruebas
# ══════════════════════════════════════════════════════════════════════════════

class MismoQue21da4e5(unittest.TestCase):
    """El modo normal hace exactamente las mismas llamadas UNO que lo que esta vivo."""

    @classmethod
    def setUpClass(cls):
        cls.vivo = convert_vivo()

    def setUp(self):
        if self.vivo is None:
            self.skipTest("no hay git para leer %s" % VIVO)

    LIBROS = {
        "primera oculta, dos visibles, filas ajustadas": lambda: Libro([
            Hoja("Cálculo", visible=False),
            Hoja("Cotización", celdas=[Celda(9, True, DESCRIPCION), Celda(4, False, "x")],
                 filas={9: Fila(452, optima=None)}, anchos={1: 8300}),
            Hoja("Notas")]),
        "todas visibles": lambda: Libro([Hoja("Hoja1"), Hoja("Hoja2", estilo="P2"), Hoja("Hoja3")]),
        "una sola": lambda: Libro([Hoja("Cotización", celdas=[Celda(3, True, DESCRIPCION)])]),
        "no se deja quitar una": lambda: Libro([Hoja("A"), Hoja("B"), Hoja("C")], quitar_truena={"B"}),
        "no se lee la visibilidad": lambda: Libro([Hoja("A", visible_truena=True), Hoja("B", visible_truena=True)]),
        "sin datos": lambda: Libro([Hoja("Vacia", fin=None), Hoja("Otra")]),
        "contenido mas alto que la hoja": lambda: Libro([Hoja("Larga", fin=(4, 400))]),
        "libro sin pestanas": lambda: Libro([]),
    }

    def corre(self, mod, hacer, modo_arg):
        log = []
        libro = hacer()
        doc = Graba(libro, "doc", log)
        try:
            if modo_arg is None:
                mod._prepare_spreadsheet(doc)
            else:
                mod._prepare_spreadsheet(doc, modo_arg)
            fin = "termina"
        except Exception as e:
            fin = "truena " + type(e).__name__
        estado = (libro.Sheets.ElementNames, libro.visibles(),
                  {n: dict(sorted(e.puesto.items())) for n, e in libro.estilos.items()},
                  [(h.nombre, _texto(h.areas)) for h in libro.Sheets.hojas],
                  [(h.nombre, sorted((r, f.Height) for r, f in h.Rows.filas.items())) for h in libro.Sheets.hojas])
        return log, fin, estado

    def test_mismas_llamadas_sin_modo_y_con_primera(self):
        for nombre, hacer in self.LIBROS.items():
            for modo_arg in (None, "primera"):
                with self.subTest(libro=nombre, modo=modo_arg):
                    viejo = self.corre(self.vivo, hacer, modo_arg)
                    nuevo = self.corre(convert, hacer, modo_arg)
                    self.assertTrue(viejo[0] or nombre == "libro sin pestanas")
                    self.assertEqual(nuevo[0], viejo[0])     # cada llamada, en orden
                    self.assertEqual(nuevo[1:], viejo[1:])   # y lo que quedo

    def test_las_funciones_del_camino_normal_son_las_mismas(self):
        import inspect
        for f in ("_es_visible", "_ajusta_pagina", "_data_bounds", "_fit_one_page", "_fit_width_only",
                  "is_encrypted", "convert", "main", "_lee_modo"):
            with self.subTest(funcion=f):
                self.assertEqual(inspect.getsource(getattr(convert, f)), inspect.getsource(getattr(self.vivo, f)))
        self.assertEqual(convert.ONE_PAGE_FLOOR, self.vivo.ONE_PAGE_FLOOR)
        self.assertEqual(convert.CONTENT_FLAGS, self.vivo.CONTENT_FLAGS)


class ModoPrimera(unittest.TestCase):
    """Lo de siempre (chat PDF convertidor, generador de POs): una pestana, como en 21da4e5."""

    def test_primera_oculta_sale_la_visible(self):
        libro = Libro([Hoja("Cálculo", visible=False), Hoja("Cotización"), Hoja("Notas")])
        convert._prepare_spreadsheet(libro, "primera")
        self.assertEqual(libro.Sheets.ElementNames, ("Cotización",))
        self.assertEqual(libro.Sheets.getByName("Cotización").areas[0].Sheet, 0)

    def test_todas_visibles_queda_la_primera_como_antes(self):
        libro = Libro([Hoja("Hoja1"), Hoja("Hoja2")])
        convert._prepare_spreadsheet(libro)
        self.assertEqual(libro.Sheets.ElementNames, ("Hoja1",))
        self.assertTrue(una_hoja(libro.estilos["Default"]))

    def test_si_no_se_lee_la_visibilidad_se_porta_como_antes(self):
        libro = Libro([Hoja("A", visible_truena=True), Hoja("B", visible_truena=True)])
        convert._prepare_spreadsheet(libro, "primera")
        self.assertEqual(libro.Sheets.ElementNames, ("A",))

    def test_no_toca_alturas(self):
        fila = Fila(452, optima=1800)
        libro = Libro([Hoja("A", celdas=[Celda(8, True, DESCRIPCION)], filas={8: fila})])
        convert._prepare_spreadsheet(libro, "primera")
        self.assertEqual(fila.Height, 452)
        self.assertEqual(fila.toques, 0)


class ModoCliente(unittest.TestCase):
    """La cotizacion al cliente: todas las visibles, ninguna oculta sale, ninguna se quita."""

    def test_salen_las_visibles_en_orden_y_las_ocultas_se_quedan_en_el_libro(self):
        # CQ#305: los precios de «Cotización» leen de «Cálculo», oculta.
        libro = Libro([Hoja("Cálculo", visible=False), Hoja("Cotización", estilo="P1"),
                       Hoja("Oculta2", visible=False), Hoja("Referencias", estilo="P2")])
        convert._prepare_spreadsheet(libro, "cliente")
        self.assertEqual(libro.Sheets.quitadas, [])
        self.assertEqual(libro.Sheets.ElementNames, ("Cálculo", "Cotización", "Oculta2", "Referencias"))
        self.assertEqual(libro.visibles(), ("Cotización", "Referencias"))
        self.assertEqual(libro.Sheets.getByName("Cotización").areas[0].Sheet, 1)
        self.assertEqual(libro.Sheets.getByName("Referencias").areas[0].Sheet, 3)
        self.assertIsNone(libro.Sheets.getByName("Cálculo").areas)

    def test_estilo_compartido_entre_visibles_va_a_lo_ancho(self):
        libro = Libro([Hoja("A"), Hoja("B")])          # las dos con «Default»
        convert._prepare_spreadsheet(libro, "cliente")
        self.assertTrue(ancho_solo(libro.estilos["Default"]))

    def test_una_oculta_con_el_mismo_estilo_no_cuenta(self):
        libro = Libro([Hoja("Cálculo", visible=False), Hoja("Cotización")])
        convert._prepare_spreadsheet(libro, "cliente")
        self.assertTrue(una_hoja(libro.estilos["Default"]))

    def test_estilo_propio_sigue_la_regla_de_una_hoja(self):
        libro = Libro([Hoja("A", estilo="P1"), Hoja("B", estilo="P2")])
        convert._prepare_spreadsheet(libro, "cliente")
        self.assertTrue(una_hoja(libro.estilos["P1"]))
        self.assertTrue(una_hoja(libro.estilos["P2"]))

    def test_si_no_se_lee_la_visibilidad_no_hay_pdf(self):
        libro = Libro([Hoja("A", visible_truena=True), Hoja("B")])
        with self.assertRaises(RuntimeError):
            convert._prepare_spreadsheet(libro, "cliente")

    def test_sin_visibles_no_hay_pdf(self):
        libro = Libro([Hoja("A", visible=False)])
        with self.assertRaises(RuntimeError):
            convert._prepare_spreadsheet(libro, "cliente")


class FilasQueCrecen(unittest.TestCase):
    """CQ#303: descripciones largas con ajuste de texto en filas de un renglon (modo cliente)."""

    ANCHO_B = 8300          # 45 caracteres de Excel, en 1/100 mm

    def hoja(self, celdas, filas):
        return Hoja("Cotización", celdas=celdas, filas=filas, anchos={1: self.ANCHO_B})

    def test_en_modo_cliente_crecen(self):
        fila = Fila(452, optima=None)
        libro = Libro([Hoja("Cotización", celdas=[Celda(9, True, DESCRIPCION)], filas={9: fila},
                            anchos={1: self.ANCHO_B})])
        convert._prepare_spreadsheet(libro, "cliente")
        self.assertGreater(fila.Height, 452)

    def test_libreoffice_no_calcula_y_crece_por_la_cuenta_propia(self):
        fila = Fila(452, optima=None)
        h = self.hoja([Celda(9, True, DESCRIPCION)], {9: fila})
        self.assertEqual(convert._crece_filas(h), 1)
        # ~206 caracteres en ~42 por renglon -> 5 o 6 renglones de 11 pt.
        self.assertGreaterEqual(fila.Height, 5 * 11 * 1.25 * convert.PUNTO)

    def test_libreoffice_calcula_mas_y_gana_libreoffice(self):
        fila = Fila(452, optima=9000)
        h = self.hoja([Celda(9, True, DESCRIPCION)], {9: fila})
        convert._crece_filas(h)
        self.assertEqual(fila.Height, 9000)

    def test_la_fila_alta_de_una_foto_no_se_encoge(self):
        foto = Fila(6000, optima=600)
        h = self.hoja([Celda(4, True, "Silla")], {4: foto})
        self.assertEqual(convert._crece_filas(h), 0)
        self.assertEqual(foto.Height, 6000)

    def test_ocultas_y_sin_ajuste_no_se_tocan(self):
        oculta = Fila(0, optima=900, visible=False)
        sin_ajuste = Fila(452, optima=1800)
        h = self.hoja([Celda(12, True, DESCRIPCION), Celda(7, False, DESCRIPCION)],
                      {12: oculta, 7: sin_ajuste})
        convert._crece_filas(h)
        self.assertEqual((oculta.Height, oculta.toques), (0, 0))
        self.assertEqual((sin_ajuste.Height, sin_ajuste.toques), (452, 0))

    def test_combinada_cuenta_el_ancho_de_toda_el_area(self):
        sola = convert._alto_necesario(self.hoja([], {}), Celda(9, True, DESCRIPCION, col=1))
        h = Hoja("C", anchos={1: self.ANCHO_B, 2: self.ANCHO_B, 3: self.ANCHO_B})
        combinada = convert._alto_necesario(h, Celda(9, True, DESCRIPCION, col=1, combinada=(1, 3, 9, 9)))
        self.assertLess(combinada, sola)
        self.assertEqual(convert._alto_necesario(h, Celda(9, True, DESCRIPCION, combinada=(1, 1, 9, 10))), 0)

    def test_varios_parrafos_y_letras_anchas_piden_mas(self):
        h = self.hoja([], {})
        uno = convert._alto_necesario(h, Celda(9, True, "a" * 30))
        dos = convert._alto_necesario(h, Celda(9, True, "a" * 15 + "\n" + "a" * 15))
        chino = convert._alto_necesario(h, Celda(9, True, "门" * 30))
        self.assertGreater(dos, uno)
        self.assertGreater(chino, uno)


class Argumentos(unittest.TestCase):
    def test_lee_modo(self):
        self.assertEqual(convert._lee_modo([]), "primera")
        self.assertEqual(convert._lee_modo(["--modo=cliente"]), "cliente")
        self.assertEqual(convert._lee_modo(["--modo=CLIENTE"]), "cliente")
        with self.assertRaises(ValueError):
            convert._lee_modo(["--modo=todo"])
        with self.assertRaises(ValueError):
            convert._lee_modo(["cliente"])

    def test_convert_rechaza_modo_raro(self):
        with self.assertRaises(ValueError):
            convert.convert("x.xlsx", "x.pdf", "todas")

    def test_main_con_modo_raro_sale_con_error(self):
        viejo = sys.argv
        try:
            sys.argv = ["convert.py", "no-existe.xlsx", "x.pdf", "--modo=raro"]
            self.assertEqual(convert.main(), convert.EXIT_ERROR)
        finally:
            sys.argv = viejo


class Servidor(unittest.TestCase):
    """server.js no se puede correr aqui (no hay Node): se revisa que el modo
    llegue al conversor y que en modo cliente no se traduzca."""

    def setUp(self):
        with open(os.path.join(AQUI, "server.js"), encoding="utf-8") as f:
            self.js = f.read()

    def test_pasa_el_modo(self):
        self.assertIn("['--modo=cliente']", self.js)
        self.assertEqual(self.js.count("buildPdf(inputPath, ext, stamp, modo)"), 2)
        self.assertIn("runConverter(inputPath, origPdf, modo)", self.js)

    def test_cliente_no_traduce(self):
        self.assertIn("TRANSLATE_ENABLED && !cliente && TRANSLATABLE_EXT.has(ext)", self.js)
        self.assertIn("if (!TRANSLATE_ENABLED || cliente) return { status: 'skip'", self.js)

    def test_sintaxis(self):
        try:
            import esprima
        except ImportError:
            self.skipTest("sin esprima")
        esprima.parseScript(self.js)


if __name__ == "__main__":
    unittest.main(verbosity=2)
