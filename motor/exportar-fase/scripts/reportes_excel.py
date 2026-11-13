# -*- coding: utf-8 -*-
"""
Genera el Excel (.xlsx) de un reporte de Tekla (.csv) con el FORMATO ESTÁNDAR de Ferma
(tomado de los reportes de la 696A FASE 102 formateados por el usuario, 2026-11-06).
Reemplaza la macro de Excel de cada modelador: todos los reportes quedan iguales.

Formato:
  fila 1        vacía con borde fino arriba
  fila 2        título (ej. "LISTADO DE PERFILES DE TEKLA STRUCTURES   Página: 1"), combinada, borde medio
  fila 3        "Modelo: ...   Proyecto: ...", combinada, borde medio
  filas 4-7     4 filas vacías con grilla fina (lugar de Fecha / Hora, que no se copian)
  encabezado    relleno gris, borde medio; si hay fila de unidades ((mm), (kg)...) va debajo
  fila vacía, datos con grilla fina, fila vacía, totales con grilla fina
  Calibri 11, centrado. Los números quedan como números.

Uso:  python -P reportes_excel.py <reporte.csv> [<otro.csv> ...]   -> crea <reporte>.xlsx al lado
      Nunca pisa un .xlsx existente (lo informa y sigue).
"""
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FINO, MEDIO = Side(style="thin"), Side(style="medium")
GRIS = PatternFill("solid", fgColor="D9D9D9")
FUENTE = Font(name="Calibri", size=11)
CENTRO = Alignment(horizontal="center", vertical="center")
# anchos de columna copiados de los Excel del usuario (por cantidad de columnas del reporte)
ANCHOS = {
    "CONJUNTOS": [7.7, 12.6, 22.6, 12.3, 12.7, 7.1, 6.7, 7.4, 8.7, 9.1, 9.9, 15.1],
    "PARTES": [14.9, 19.9, 8.4, 9.6, 9.7, 14.4, 11.3, 11.9],
}
NUM = re.compile(r"^-?\d+(?:\.\d+)?$")


def limpiar(s):
    # Tekla a veces mete un byte de control antes de "²" (ej. "Total m\x9d² Pintura"): se saca
    return re.sub(r"[\x80-\x9f�]", "", s)


def valor(s):
    s = limpiar(s).strip()
    if NUM.match(s):
        return float(s) if "." in s else int(s)
    return s


def leer_csv(ruta):
    with open(ruta, encoding="cp1252", errors="replace") as f:
        lineas = [l.rstrip("\r\n") for l in f]
    titulo, modelo, filas = "", "", []
    for l in lineas:
        t = l.strip()
        if not t or re.fullmatch(r"-+", t):
            filas.append(None)                      # separador / vacía
        elif ";" in l:
            filas.append([limpiar(c).strip() for c in l.split(";")])
        elif t.upper().startswith("MODELO"):
            modelo = re.sub(r"\s{2,}", "   ", t)
        elif t.upper().startswith(("FECHA", "HORA")):
            continue
        elif not titulo:
            titulo = re.sub(r"\s{2,}", "   ", t)
    # sacar vacías repetidas, del principio y del final
    limpias = []
    for f in filas:
        if f is None and (not limpias or limpias[-1] is None):
            continue
        limpias.append(f)
    while limpias and limpias[0] is None:
        limpias.pop(0)
    while limpias and limpias[-1] is None:
        limpias.pop()
    return titulo, modelo, limpias


def caja(ws, fila, c1, c2, lado):
    for c in range(c1, c2 + 1):
        ws.cell(fila, c).border = Border(top=lado, bottom=lado,
                                         left=lado if c == c1 else None, right=lado if c == c2 else None)


def grilla(ws, fila, ncol, lado=FINO):
    for c in range(1, ncol + 1):
        cel = ws.cell(fila, c)
        cel.border = Border(left=lado, right=lado, top=lado, bottom=lado)
        cel.alignment, cel.font = CENTRO, FUENTE


# A4 con márgenes de ~1 cm: entran unas 100 unidades de ancho de Excel en vertical y 145 en horizontal
# (1 unidad ≈ 1,9 mm). Si en vertical el texto queda por debajo del 80 % (~9 pt) se pasa a horizontal.
ANCHO_A4_VERTICAL, ANCHO_A4_HORIZONTAL, ESCALA_MINIMA = 100, 145, 0.80


def orientacion(ancho_total):
    return "portrait" if ANCHO_A4_VERTICAL / max(ancho_total, 1) >= ESCALA_MINIMA else "landscape"


def generar(ruta_csv, ruta_xlsx=None):
    ruta_xlsx = ruta_xlsx or os.path.splitext(ruta_csv)[0] + ".xlsx"
    if os.path.exists(ruta_xlsx):
        return None, f"ya existe, no se pisa: {os.path.basename(ruta_xlsx)}"
    titulo, modelo, filas = leer_csv(ruta_csv)
    datos = [f for f in filas if f]
    if not datos:
        return None, f"sin datos: {os.path.basename(ruta_csv)}"
    ncol = max(len(f) for f in datos)
    wb = Workbook()
    ws = wb.active
    ws.title =re.sub(r"[^A-Za-z0-9_ ]", "_", os.path.splitext(os.path.basename(ruta_xlsx))[0])[:31].strip("_ ") or "Reporte"
    # encabezado de página
    caja(ws, 1, 1, ncol, FINO)
    for fila, texto in ((2, titulo), (3, modelo)):
        ws.cell(fila, 1, texto).alignment = CENTRO
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=ncol)
        caja(ws, fila, 1, ncol, MEDIO)
    for fila in range(4, 8):
        grilla(ws, fila, ncol)
    r = 8
    encabezado_hecho = False
    for f in filas:
        if f is None:
            grilla(ws, r, ncol)
            r += 1
            continue
        es_unidades = encabezado_hecho and r == fila_enc + 1 and all(not x or x.startswith("(") for x in f)
        for c, x in enumerate(f, start=1):
            v = valor(x)
            cel = ws.cell(r, c, v)
            if isinstance(v, float):
                cel.number_format = "0.00"          # 516.20 y no 516.2
        grilla(ws, r, ncol)
        if not encabezado_hecho:                       # primera fila con ';' = encabezado de columnas
            for c in range(1, ncol + 1):
                ws.cell(r, c).fill = GRIS
                ws.cell(r, c).border = Border(left=MEDIO if c == 1 else FINO, right=MEDIO if c == ncol else FINO,
                                              top=MEDIO, bottom=MEDIO)
            encabezado_hecho, fila_enc = True, r
        elif es_unidades:
            for c in range(1, ncol + 1):
                ws.cell(r, c).fill = GRIS
        r += 1
    # anchos
    tipo = "CONJUNTOS" if "PERFILES" in titulo.upper() and ncol == 12 else ("PARTES" if "PARTES" in titulo.upper() and ncol == 8 else None)
    ancho_total = 0
    for c in range(1, ncol + 1):
        if tipo:
            w = ANCHOS[tipo][c - 1]
        else:
            w = max([len(str(ws.cell(rr, c).value or "")) for rr in range(8, r)] + [6]) + 3
        ws.column_dimensions[get_column_letter(c)].width = w
        ancho_total += w
    for fila in (1, 2, 3, 7, 8):
        ws.row_dimensions[fila].height = 15.75
    ws.sheet_view.showGridLines = False
    # impresión (pedido del usuario 2026-11-07): A4, 1 hoja de ancho como máximo y las hojas
    # de alto que hagan falta. Vertical si entra a un tamaño legible; si no, horizontal.
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = orientacion(ancho_total)
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4          # pulgadas (~1 cm)
    ws.page_margins.top = ws.page_margins.bottom = 0.6
    # Nota: repetir el encabezado en cada hoja (print_title_rows) no lo toma este Excel: probado 2026-11-07.
    wb.save(ruta_xlsx)
    return ruta_xlsx, f"{os.path.basename(ruta_xlsx)}: {len(datos)} filas, {ncol} columnas"


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    for p in sys.argv[1:]:
        _, msg = generar(p)
        print(msg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
