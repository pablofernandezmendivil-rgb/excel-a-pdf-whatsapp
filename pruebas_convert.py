# -*- coding: utf-8 -*-
"""Pruebas de convert.py sin LibreOffice: un libro de mentiras con la misma
forma que el de UNO (Sheets, IsVisible, PageStyle, Rows, queryContentCells...).

Lo que se prueba es la decision —que pestanas quedan, que filas crecen, que
escala lleva cada pagina— no el PDF. El PDF de verdad se prueba contra el
servicio desplegado (los casos ciegos del 25-oct-26).

    python pruebas_convert.py
"""
import sys
import types
import unittest

# convert.py importa `uno` adentro de las funciones; aqui no hay LibreOffice.
uno = types.ModuleType("uno")
uno.createUnoStruct = lambda nombre: types.SimpleNamespace()
sys.modules.setdefault("uno", uno)

import convert  # noqa: E402


class Celda:
    def __init__(self, fila, ajustada):
        self.CellAddress = types.SimpleNamespace(Row=fila)
        self.IsTextWrapped = ajustada


class Enum:
    def __init__(self, cosas):
        self.cosas = list(cosas)

    def hasMoreElements(self):
        return bool(self.cosas)

    def nextElement(self):
        return self.cosas.pop(0)


class Rangos:
    def __init__(self, hoja, banderas):
        self.hoja, self.banderas = hoja, banderas
        fin = hoja.fin
        self.RangeAddresses = [types.SimpleNamespace(EndColumn=fin[0], EndRow=fin[1])] if fin else []

    def getCells(self):
        hoja = self

        class Acceso:
            def createEnumeration(self_):
                return Enum(hoja.hoja.celdas)
        return Acceso()


class Fila:
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
        return self.filas.setdefault(i, Fila(500))


class Columnas:
    def getByIndex(self, i):
        return types.SimpleNamespace(Width=2000)


class Hoja:
    def __init__(self, nombre, visible=True, estilo="Default", celdas=(), filas=None,
                 fin=(5, 10), visible_truena=False):
        self.nombre, self._visible, self.PageStyle = nombre, visible, estilo
        self.celdas, self.fin, self.visible_truena = list(celdas), fin, visible_truena
        self.Rows = Filas(filas or {})
        self.Columns = Columnas()
        self.areas = None

    @property
    def IsVisible(self):
        if self.visible_truena:
            raise RuntimeError("UNO no contesto")
        return self._visible

    def queryContentCells(self, banderas):
        return Rangos(self, banderas)

    def setPrintAreas(self, areas):
        self.areas = areas


class Hojas:
    def __init__(self, hojas, quitar_truena=()):
        self.hojas = list(hojas)
        self.quitar_truena = set(quitar_truena)

    @property
    def ElementNames(self):
        return tuple(h.nombre for h in self.hojas)

    def getByName(self, n):
        return next(h for h in self.hojas if h.nombre == n)

    def removeByName(self, n):
        if n in self.quitar_truena:
            raise RuntimeError("no se deja quitar")
        self.hojas = [h for h in self.hojas if h.nombre != n]


class Estilo:
    def __init__(self):
        self.Width, self.Height = 21000, 29700
        self.LeftMargin = self.RightMargin = self.TopMargin = self.BottomMargin = 1000
        self.puesto = {}

    def setPropertyValue(self, k, v):
        self.puesto[k] = v


class Libro:
    def __init__(self, hojas, quitar_truena=()):
        self.Sheets = Hojas(hojas, quitar_truena)
        self.estilos = {}
        libro = self

        class Familia:
            def getByName(self_, n):
                return libro.estilos.setdefault(n, Estilo())

        class Familias:
            def getByName(self_, n):
                assert n == "PageStyles"
                return Familia()
        self.StyleFamilies = Familias()


def ancho_solo(est):
    return est.puesto.get("ScaleToPagesX") == 1 and est.puesto.get("ScaleToPagesY") == 0


def una_hoja(est):
    return est.puesto.get("ScaleToPagesX") == 1 and est.puesto.get("ScaleToPagesY") == 1


class ModoPrimera(unittest.TestCase):
    """Lo de siempre (chat PDF convertidor, generador de POs): una pestana."""

    def test_primera_oculta_sale_la_visible(self):
        # CQ#305 de la prueba ciega: «Calculo» oculta con costos y margen.
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
        fila = Fila(450, optima=1800)
        libro = Libro([Hoja("A", celdas=[Celda(8, True)], filas={8: fila})])
        convert._prepare_spreadsheet(libro, "primera")
        self.assertEqual(fila.Height, 450)
        self.assertEqual(fila.toques, 0)


class ModoCliente(unittest.TestCase):
    """La cotizacion al cliente: todas las visibles, ninguna oculta."""

    def test_quita_ocultas_y_deja_visibles_en_orden(self):
        libro = Libro([Hoja("Proveedor", visible=False), Hoja("Cotización", estilo="P1"),
                       Hoja("Oculta2", visible=False), Hoja("Referencias", estilo="P2")])
        convert._prepare_spreadsheet(libro, "cliente")
        self.assertEqual(libro.Sheets.ElementNames, ("Cotización", "Referencias"))
        self.assertEqual(libro.Sheets.getByName("Cotización").areas[0].Sheet, 0)
        self.assertEqual(libro.Sheets.getByName("Referencias").areas[0].Sheet, 1)

    def test_estilo_compartido_va_a_lo_ancho(self):
        libro = Libro([Hoja("A"), Hoja("B")])          # las dos con «Default»
        convert._prepare_spreadsheet(libro, "cliente")
        self.assertTrue(ancho_solo(libro.estilos["Default"]))

    def test_estilo_propio_sigue_la_regla_de_una_hoja(self):
        libro = Libro([Hoja("A", estilo="P1"), Hoja("B", estilo="P2")])
        convert._prepare_spreadsheet(libro, "cliente")
        self.assertTrue(una_hoja(libro.estilos["P1"]))
        self.assertTrue(una_hoja(libro.estilos["P2"]))

    def test_una_sola_visible_es_como_antes_mas_las_filas(self):
        libro = Libro([Hoja("Cotización")])
        convert._prepare_spreadsheet(libro, "cliente")
        self.assertEqual(libro.Sheets.ElementNames, ("Cotización",))
        self.assertTrue(una_hoja(libro.estilos["Default"]))

    def test_filas_ajustadas_crecen_y_nunca_se_encogen(self):
        larga = Fila(450, optima=1800)        # descripcion larga cortada (CQ#303)
        foto = Fila(3000, optima=600)         # fila alta por una foto
        oculta = Fila(0, optima=900, visible=False)
        sin_ajuste = Fila(450, optima=1800)   # su texto no esta ajustado
        hoja = Hoja("Cotización", celdas=[Celda(9, True), Celda(9, True), Celda(4, True),
                                          Celda(12, True), Celda(7, False)],
                    filas={9: larga, 4: foto, 12: oculta, 7: sin_ajuste})
        self.assertEqual(convert._crece_filas(hoja), 1)
        self.assertEqual(larga.Height, 1800)
        self.assertEqual(foto.Height, 3000)
        self.assertEqual(oculta.Height, 0)
        self.assertEqual(oculta.toques, 0)
        self.assertEqual(sin_ajuste.Height, 450)
        self.assertEqual(sin_ajuste.toques, 0)

    def test_si_una_oculta_no_se_deja_quitar_no_hay_pdf(self):
        libro = Libro([Hoja("Cálculo", visible=False), Hoja("Cotización")],
                      quitar_truena={"Cálculo"})
        with self.assertRaises(RuntimeError):
            convert._prepare_spreadsheet(libro, "cliente")

    def test_si_no_se_lee_la_visibilidad_no_hay_pdf(self):
        libro = Libro([Hoja("A", visible_truena=True), Hoja("B")])
        with self.assertRaises(RuntimeError):
            convert._prepare_spreadsheet(libro, "cliente")

    def test_sin_visibles_no_hay_pdf(self):
        libro = Libro([Hoja("A", visible=False)])
        with self.assertRaises(RuntimeError):
            convert._prepare_spreadsheet(libro, "cliente")


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
        with open("server.js", encoding="utf-8") as f:
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
