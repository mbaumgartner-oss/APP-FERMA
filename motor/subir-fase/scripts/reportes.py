# -*- coding: utf-8 -*-
"""
Lectura y comparación de reportes de Tekla (.xlsx) - Skill "subir-fase" (reentregas).
Regla: Reglas de carga y revisiones.md, sección 8 (pedido del usuario 2026-11-05).

Se compara SOLO la tabla de datos. Se ignoran: título, modelo, fecha, hora, página, filas de totales,
y diferencias de formato en los encabezados ("P. UNIT (kg)" = "P.UNIT (kg)"). Cada modelador usa su
propia macro: por eso se busca la fila de encabezados en vez de suponer una posición fija.

Tipos conocidos (vistos en 674L FASE 56):
  SUBMATERIALES: Marca | Material | Grado | Longitud | Cant. | Miembro | Peso Unit. | Peso Total
  PESO:          Cant. | Código | Material | Categoría | Grado | Long. | Alto | Ancho | P.UNIT | P.TOTAL | Sec. | Área de pintura
Otros (bulones, tuercas): comparación genérica fila por fila.

Uso suelto (prueba):  python reportes.py <viejo.xlsx> <nuevo.xlsx>
Lee .xlsx (openpyxl) y .xls (xlrd): python -m pip install openpyxl xlrd
"""
import os
import re
import sys
import unicodedata

CLAVES = {"MARCA": "SUBMATERIALES", "CODIGO": "PESO"}   # columna clave -> tipo de tabla
# columnas que describen la pieza (las demás se derivan o son de presentación)
CAMPOS = {
    "SUBMATERIALES": ["MATERIAL", "GRADO", "LONGITUD"],
    "PESO": ["MATERIAL", "GRADO", "LONG", "ALTO", "ANCHO", "CATEGORIA"],
}


def _n(t):
    t = unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode()
    t = re.sub(r"\(.*?\)", "", t)                 # saca unidades: (kg), (mm), (m2)
    t = re.sub(r"[^A-Za-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip().upper()


def _celda(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:.2f}".rstrip("0").rstrip(".") if v != int(v) else str(int(v))
    s = str(v).strip()
    if re.fullmatch(r"-?\d+[.,]\d+", s):          # "0.48" / "0,48"
        f = float(s.replace(",", "."))
        return f"{f:.2f}".rstrip("0").rstrip(".") if f != int(f) else str(int(f))
    return re.sub(r"\s+", " ", s)


class Tabla:
    def __init__(self, ruta):
        self.ruta = ruta
        self.tipo = None          # SUBMATERIALES / PESO / GENERICO
        self.cab = []             # encabezados normalizados
        self.filas = []           # listas de celdas normalizadas
        self.error = None
        self._leer()

    def _leer(self):
        ext = os.path.splitext(self.ruta)[1].lower()
        if ext not in (".xlsx", ".xls"):
            self.error = "formato no soportado (solo .xls y .xlsx)"
            return
        if os.path.getsize(self.ruta) == 0:
            self.error = "archivo vacío"
            return
        try:
            if ext == ".xlsx":
                import openpyxl
                wb = openpyxl.load_workbook(self.ruta, read_only=True, data_only=True)
                datos = [[_celda(v) for v in r] for r in wb.worksheets[0].iter_rows(values_only=True)]
                wb.close()
            else:                                   # .xls viejo (formato de los reportes de Tekla de hoy)
                import xlrd
                s = xlrd.open_workbook(self.ruta).sheet_by_index(0)
                datos = [[_celda(v) for v in s.row_values(i)] for i in range(s.nrows)]
        except ImportError as e:
            self.error = f"falta la librería ({e.name}): python -m pip install openpyxl xlrd"
            return
        except Exception as e:
            self.error = f"no se pudo abrir: {type(e).__name__}"
            return
        # líneas separadoras "-----" de los reportes de Tekla
        datos = [r if any(c and not re.fullmatch(r"[-=_ ]+", c) for c in r) else [] for r in datos]
        # encabezado: primera fila con al menos 4 celdas no vacías y sin números puros
        idx = None
        if not datos:
            self.error = "hoja vacía"
            return
        for i, r in enumerate(datos):
            llenas = [c for c in r if c]
            if len(llenas) >= 4 and not any(re.fullmatch(r"-?[\d.]+", c) for c in llenas):
                idx = i
                break
        if idx is None:
            self.error = "no se encontró la fila de encabezados"
            return
        self.cab = [_n(c) for c in datos[idx]]
        while self.cab and not self.cab[-1]:     # columnas vacías al final (Excel hechos con otra macro)
            self.cab.pop()
        self.tipo = next((t for k, t in CLAVES.items() if k in self.cab), "GENERICO")
        clave = self.col_clave()
        for r in datos[idx + 1:]:
            r = (r + [""] * len(self.cab))[:len(self.cab)]
            if not any(r):
                continue
            if self.tipo == "GENERICO":
                if any(_n(c) in ("TOTAL", "CONJUNTOS") for c in r):
                    continue
            else:
                k = r[clave]
                if not k or _n(k) in ("TOTAL", "CONJUNTOS") or any(_n(c) == "TOTAL" for c in r):
                    continue           # fila de totales
            self.filas.append(r)

    def col_clave(self):
        for k in CLAVES:
            if k in self.cab:
                return self.cab.index(k)
        return 0

    def col(self, nombre):
        nombre = _n(nombre)
        for i, c in enumerate(self.cab):
            if c == nombre or c.startswith(nombre + " ") or c.replace(" ", "") == nombre.replace(" ", ""):
                return i
        return None

    def por_pieza(self):
        """{clave: {'campos': {...}, 'cant': total, 'miembros': {...}}} para SUBMATERIALES y PESO."""
        out = {}
        ck = self.col_clave()
        cc = self.col("CANT")
        cm = self.col("MIEMBRO")
        for r in self.filas:
            k = r[ck]
            d = out.setdefault(k, {"campos": {}, "cant": 0.0, "miembros": {}})
            for campo in CAMPOS.get(self.tipo, []):
                i = self.col(campo)
                if i is not None and campo not in d["campos"]:
                    d["campos"][campo] = r[i]
            try:
                q = float(r[cc]) if cc is not None and r[cc] else 0.0
            except ValueError:
                q = 0.0
            d["cant"] += q
            if cm is not None:
                d["miembros"][r[cm]] = d["miembros"].get(r[cm], 0) + q
        return out

    def firma(self):
        """Contenido comparable: filas ordenadas (sin metadatos ni totales)."""
        return sorted(tuple(r) for r in self.filas)


def comparar(viejo, nuevo):
    """Devuelve dict: igual (bool), tipo, cambiadas {clave: [textos]}, nuevas [..], eliminadas [..], error."""
    a, b = Tabla(viejo), Tabla(nuevo)
    res = {"igual": False, "tipo": b.tipo or a.tipo, "cambiadas": {}, "nuevas": [], "eliminadas": [], "error": None}
    if a.error or b.error:
        res["error"] = f"viejo: {a.error or 'ok'} · nuevo: {b.error or 'ok'}"
        return res
    if a.firma() == b.firma():
        res["igual"] = True
        return res
    if a.tipo != b.tipo or a.tipo == "GENERICO":
        res["cambiadas"]["(tabla)"] = [f"{len(set(a.firma()) ^ set(b.firma()))} fila(s) distintas"]
        return res
    pa, pb = a.por_pieza(), b.por_pieza()
    res["nuevas"] = sorted(k for k in pb if k not in pa)
    res["eliminadas"] = sorted(k for k in pa if k not in pb)
    for k in sorted(set(pa) & set(pb)):
        x, y = pa[k], pb[k]
        txt = [f"{c.title()} {x['campos'].get(c, '')} → {y['campos'].get(c, '')}"
               for c in CAMPOS[a.tipo] if x["campos"].get(c, "") != y["campos"].get(c, "")]
        if abs(x["cant"] - y["cant"]) > 1e-6:
            txt.append(f"Cant. {x['cant']:g} → {y['cant']:g}")
        if x["miembros"] != y["miembros"]:
            txt.append("Conjuntos donde va: " + ", ".join(sorted(x["miembros"])) + " → " + ", ".join(sorted(y["miembros"])))
        if not txt:
            # cambió algo menor (peso, pintura): se informa igual
            fa = sorted(tuple(r) for r in a.filas if r[a.col_clave()] == k)
            fb = sorted(tuple(r) for r in b.filas if r[b.col_clave()] == k)
            if fa != fb:
                txt.append("cambian peso u otros datos")
        if txt:
            res["cambiadas"][k] = txt
    return res


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    r = comparar(sys.argv[1], sys.argv[2])
    print("Tipo:", r["tipo"], "· Igual:", r["igual"], "· Error:", r["error"])
    print("Nuevas:", r["nuevas"])
    print("Eliminadas:", r["eliminadas"])
    for k, v in r["cambiadas"].items():
        print(f"  {k}: " + " · ".join(v))
