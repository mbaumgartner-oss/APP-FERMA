# -*- coding: utf-8 -*-
"""
Control de fase (solo lectura) - Skill "Subir fase al servidor" de Ferma.

Revisa una fase contra las reglas y lista los errores. NO mueve, NO renombra,
NO borra nada. Reglas en el vault: Cerebro IA Ferma/WIKI/Relevamiento/
(Mapeo local a servidor, Reglas de nomenclatura, Tipos de fase y qué se sube,
Reglas de carga y revisiones, Alcance Skill v1 - Subir fase).

Uso:
  # Fase en la zona compartida (antes de subir) - uso normal de la Skill
  python control_fase.py zona  --fase-dir "<EN CONTROL>\\674L\\FASE 05" --obra 674L --planilla 674L.csv

  # Obra ya subida al servidor (auditoría) - para probar o revisar obras
  python control_fase.py servidor --constructivos "<...>\\CONSTRUCTIVOS" --obra 674L --planilla 674L.csv [--fases 5,9,22]

Opciones: --json (salida para la Skill), --salida <archivo.md>
Código de salida: 0 sin errores, 1 hay errores, 2 error de uso.
"""
import argparse
import csv
import json
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field

# ---------------------------------------------------------------- reglas
PREFIJOS_SUB = [  # Prefijos de submateriales (Z = letra de obra)
    "wpw", "wpb", "cfc", "cfz", "hss", "tws", "blt",
    "wt", "tb", "st", "cl", "bp", "sb", "rb", "dk", "gr", "gt", "ss",
    "a", "c", "s", "m", "w", "p",
]
TIPOS_REPORTE = {
    "SUBMATERIALES": "REPORTE DE SUBMATERIALES",
    "PESO": "REPORTE DE PESO",
    "BULONES OBRA": "REPORTE DE BULONES OBRA",
    "BULONES TALLER": "REPORTE DE BULONES TALLER",
    "TUERCAS Y ARANDELAS": "REPORTE DE TUERCAS Y ARANDELAS",
}
# Se modelan pero se COMPRAN fabricados: no se detallan ni se suben (confirmado 2026-11-04)
PREFIJOS_COMPRADOS = {"nut": "tuerca", "washer": "arandela"}
# Se compran hechos pero SÍ llevan PDF (y se suben): no tienen CNC. Todo "gr" (confirmado por el usuario 2026-11-10)
PREFIJOS_SIN_CNC = {"gr": "reja / escalón de reja"}
EXT_NO_SE_SUBEN = {".bak", ".log", ".dwl", ".dwl2", ".tmp", ".zip", ".db", ".ini"}
ESTADO_PARA_SUBIR = "APROBADO"
ESTADO_SUBIDO = "SUBIDO AL SERVIDOR"
# Encabezado de la primera columna en la planilla: son equivalentes (confirmado 2026-11-05, ej. 702Q usa "FASES")
ENCABEZADOS_SECUENCIA = {"SECUENCIA", "SECUENCIAS", "FASE", "FASES"}

ERROR, AVISO, INFO = "ERROR", "AVISO", "INFO"


# ---------------------------------------------------------------- utilidades
def norm(texto):
    """Mayúsculas, sin acentos, espacios simples."""
    t = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", t).strip().upper()


# Archivos que crean Windows, Google Drive o Excel solos: no son de la fase, nunca se suben ni se cuentan
# (desktop.ini está en todas las carpetas de PASADO A CONTROL; ~$x.xlsx aparece al abrir un reporte). 2026-11-11
SISTEMA = {"desktop.ini", "thumbs.db", ".ds_store"}


def es_de_sistema(nombre):
    n = os.path.basename(nombre).lower()
    return n in SISTEMA or n.startswith("~$")


def archivos(carpeta, ext=None):
    if not carpeta or not os.path.isdir(carpeta):
        return []
    res = []
    for n in sorted(os.listdir(carpeta)):
        p = os.path.join(carpeta, n)
        if es_de_sistema(n):
            continue
        if os.path.isfile(p) and (ext is None or n.lower().endswith(ext)):
            res.append(n)
    return res


def subcarpeta_fase(base, n):
    """Encuentra 'FASE 05', 'FASE 5', 'FASE 05_CATEGORIA', 'SEC 05'... dentro de base."""
    if not base or not os.path.isdir(base):
        return None
    pat = re.compile(r"^(FASE|SEC)\s*0*%d(?:\b|_|\s|$)" % n, re.I)
    for d in sorted(os.listdir(base)):
        if os.path.isdir(os.path.join(base, d)) and pat.match(d):
            return os.path.join(base, d)
    return None


# ---------------------------------------------------------------- planilla
@dataclass
class FilaPlanilla:
    secuencia: int
    categoria: str
    estado: str
    fila: int


def leer_planilla(ruta):
    """Lee el CSV de una pestaña de obra. Tolera varios bloques, datos sucios y totales."""
    with open(ruta, encoding="utf-8-sig", newline="") as f:
        return leer_planilla_filas(list(csv.reader(f)))


def leer_planilla_filas(datos):
    """Igual que leer_planilla, pero desde una lista de filas (CSV o planilla en vivo)."""
    # La columna de secuencia NO siempre es la primera (ej. 696A tiene SECTOR antes, pedido de ING):
    # se busca el encabezado en toda la fila y se usa esa columna. La planilla no se modifica.
    filas, estados_validos, repetidas = {}, set(), []
    idx_seq = 0
    idx_cat = idx_est = None
    for i, r in enumerate(datos, start=1):
        celdas = [str(c).strip() for c in r]
        cab = [norm(c) for c in celdas]
        pos = next((j for j, c in enumerate(cab) if c in ENCABEZADOS_SECUENCIA), None)
        if pos is not None:
            idx_seq = pos
            idx_cat = cab.index("CATEGORIA") if "CATEGORIA" in cab else None
            idx_est = cab.index("ESTADO") if "ESTADO" in cab else None
            continue
        # lista de estados válidos: columna con "ESTADO DE LAS SECUENCIAS" y lo que sigue abajo
        for j, c in enumerate(celdas):
            if norm(c) == "ESTADO DE LAS SECUENCIAS":
                for r2 in datos[i:]:
                    if j < len(r2) and r2[j].strip():
                        estados_validos.add(norm(r2[j]))
        if idx_cat is None or idx_seq >= len(celdas) or not re.fullmatch(r"\d+", celdas[idx_seq]):
            continue
        n = int(celdas[idx_seq])
        cat = celdas[idx_cat] if idx_cat is not None and idx_cat < len(celdas) else ""
        est = celdas[idx_est] if idx_est is not None and idx_est < len(celdas) else ""
        if n in filas:
            repetidas.append(n)
        filas.setdefault(n, []).append(FilaPlanilla(n, cat.strip(), norm(est), i))
    return filas, estados_validos, repetidas


# ---------------------------------------------------------------- fase
@dataclass
class Fase:
    numero: int
    obra: str                       # código corto, ej. 674L
    sub_pdf: list = field(default_factory=list)
    cnc: list = field(default_factory=list)
    tipicos: list = field(default_factory=list)
    reportes: list = field(default_factory=list)   # nombres de archivo
    strumis: list = field(default_factory=list)
    otros: list = field(default_factory=list)      # archivos que no son de ningún tipo
    carpeta_strumis: str = ""
    origen: str = ""
    donde: dict = field(default_factory=dict)      # reporte -> carpetas donde está (modo servidor)

    @property
    def letra(self):
        m = re.match(r"^\d+([A-Z]+)$", self.obra.upper())
        return m.group(1) if m else ""


# Subcarpetas de entrada equivalentes: SUB = SUBMATERIALES, TIP = TIPICOS (696A, confirmado 2026-11-06)
ALIAS_ENTRADA = {"SUBMATERIALES": ("SUBMATERIALES", "SUB"), "TIPICOS": ("TIPICOS", "TIP"), "CNC": ("CNC",)}


def dir_entrada(fase_dir, clave):
    """Carpeta real de entrada para SUBMATERIALES / TIPICOS / CNC (acepta los nombres cortos)."""
    for n_ in ALIAS_ENTRADA.get(clave, (clave,)):
        p = os.path.join(fase_dir, n_)
        if os.path.isdir(p):
            return p
    return os.path.join(fase_dir, clave)


def fase_desde_zona(fase_dir, obra, n):
    f = Fase(n, obra, origen=fase_dir)
    f.cnc = archivos(dir_entrada(fase_dir, "CNC"))
    f.sub_pdf = archivos(dir_entrada(fase_dir, "SUBMATERIALES"))
    f.tipicos = archivos(dir_entrada(fase_dir, "TIPICOS"))
    for a in archivos(fase_dir):
        al = a.lower()
        if al.endswith(".bswx"):
            f.strumis.append(a)
        elif "REPORTE" in a.upper() or "BULONES" in a.upper():
            f.reportes.append(a)
        elif al.endswith("_log.txt") or al == "_fase.json":
            continue  # log de Strumis y manifiesto: no se suben y no son error
        else:
            f.otros.append(a)
    return f


def fase_desde_servidor(constructivos, obra, n, multi_letra=False):
    """Lee una fase ya subida. Con multi_letra, la letra de obra va como subcarpeta (TIPICOS\\678Z\\PDF)."""
    L = obra if multi_letra else ""

    def base(*partes):
        partes = [partes[0], L, *partes[1:]] if L else list(partes)
        return os.path.join(constructivos, *partes)

    f = Fase(n, obra, origen=constructivos)
    f.sub_pdf = archivos(subcarpeta_fase(base("SUBMATERIALES", "PDF"), n))
    f.cnc = archivos(subcarpeta_fase(base("SUBMATERIALES", "CNC"), n))
    f.tipicos = archivos(subcarpeta_fase(base("TIPICOS", "PDF"), n))
    pat_rep = re.compile(r"(FASE|SECUENCIA)\s*0*%d(?:[_ .]|$)" % n, re.I)
    for b in (base("SUBMATERIALES", "REPORTES"), base("TIPICOS", "REPORTES"), base("REPORTE DE BULONES")):
        sub = subcarpeta_fase(b, n)            # forma nueva: REPORTES\FASE 05\
        f.reportes += archivos(sub)
        f.reportes += [a for a in archivos(b) if pat_rep.search(a)]   # forma vieja: sueltos
        for a in archivos(sub) + [a for a in archivos(b) if pat_rep.search(a)]:
            f.donde[a] = f.donde.get(a, []) + [os.path.relpath(sub if a in archivos(sub) else b, constructivos)]
    cs = subcarpeta_fase(base("STRUMIS"), n)
    if cs:
        f.carpeta_strumis = os.path.basename(cs)
        for a in archivos(cs):
            (f.strumis if a.lower().endswith(".bswx") else f.otros).append(a)
    for lista in (f.sub_pdf, f.cnc, f.tipicos):
        for a in list(lista):
            if os.path.splitext(a)[1].lower() in EXT_NO_SE_SUBEN:
                lista.remove(a)
                f.otros.append(a)
    return f


# ---------------------------------------------------------------- reglas de control
@dataclass
class Hallazgo:
    nivel: str
    codigo: str
    texto: str
    archivos: list = field(default_factory=list)


def controlar(f, filas_planilla, estados_validos, modo):
    h = []
    add = lambda *a: h.append(Hallazgo(*a))
    NN = "%02d" % f.numero
    obra = f.obra.upper()

    # --- planilla
    fp = filas_planilla.get(f.numero)
    categoria = ""
    if not fp:
        add(ERROR, "PLAN-01", f"La fase {f.numero} no está en la planilla. Si es no constructiva o un prearmado, no se sube.")
    else:
        if len(fp) > 1:
            add(AVISO, "PLAN-03", f"La secuencia {f.numero} está repetida en la planilla (filas {', '.join(str(x.fila) for x in fp)}).")
        p = fp[0]
        categoria = p.categoria
        if not p.categoria:
            add(ERROR, "PLAN-02", f"La fila {p.fila} de la planilla no tiene CATEGORIA.")
        if estados_validos and p.estado and p.estado not in estados_validos | {ESTADO_PARA_SUBIR}:
            add(AVISO, "PLAN-04", f"Estado '{p.estado}' no está en la lista de estados de la obra.")
        if modo == "zona" and p.estado != ESTADO_PARA_SUBIR:
            add(ERROR, "PLAN-05", f"Estado '{p.estado or '(vacío)'}': solo se sube con '{ESTADO_PARA_SUBIR}' (OK del checker).")
        if modo == "servidor" and p.estado != ESTADO_SUBIDO:
            add(AVISO, "PLAN-06", f"Hay archivos en el servidor pero la planilla dice '{p.estado or '(vacío)'}'.")

    # --- tipo de fase
    if not f.sub_pdf and not f.cnc:
        if f.tipicos:
            add(AVISO, "FASE-01", "No hay submateriales ni CNC. ¿Es una fase solo de típicos (correas sin accesorios, plegados)? Confirmar.")
        else:
            add(ERROR, "FASE-02", "No hay submateriales, CNC ni típicos.")

    # --- submateriales <-> CNC
    pat_sub = re.compile(r"^(\d+)_([a-z]+)%s(\d+)$" % re.escape(f.letra))
    pat_comprado = re.compile(r"^\d+_(%s)" % "|".join(PREFIJOS_COMPRADOS))
    comprados = [a for a in f.sub_pdf + f.cnc if pat_comprado.match(a)]
    if comprados:
        if modo == "zona":
            add(INFO, "SUB-06", "Tuercas y arandelas compradas: se modelan pero NO se suben. Se excluyen al subir.", comprados)
        else:
            add(AVISO, "SUB-06", "Tuercas y arandelas compradas que están en el servidor y no deberían.", comprados)
    base_pdf = {os.path.splitext(a)[0]: a for a in f.sub_pdf if a not in comprados}
    base_cnc = {}
    for a in f.cnc:
        if a in comprados:
            continue
        b = re.sub(r"_\d+$", "", os.path.splitext(a)[0])   # 100_wL27_1 -> 100_wL27
        base_cnc.setdefault(b, a)
    malos, prefijo_raro = [], []
    for b, a in list(base_pdf.items()) + list(base_cnc.items()):
        m = pat_sub.match(b)
        if not m:
            malos.append(a)
        elif int(m.group(1)) != f.numero:
            malos.append(a + f"  (dice fase {m.group(1)})")
        elif m.group(2) not in PREFIJOS_SUB:
            prefijo_raro.append(a)
    if malos:
        add(ERROR, "SUB-01", f"Nombre de submaterial fuera de regla (<fase>_<prefijo>{f.letra}<n>, prefijo en minúscula).", sorted(set(malos)))
    if prefijo_raro:
        add(AVISO, "SUB-05", "Prefijo que no está en la lista de prefijos de submateriales. Confirmar si es válido.", sorted(set(prefijo_raro)))
    sin_cnc = sorted(base_pdf[b] for b in base_pdf if b not in base_cnc)
    sin_pdf = sorted(base_cnc[b] for b in base_cnc if b not in base_pdf)
    # Rejas / escalones de reja: se compran hechos, llevan PDF pero no CNC (usuario, 2026-11-10, 696A 108_grA83)
    pat_sin_cnc = re.compile(r"^\d+_(%s)[A-Z]" % "|".join(PREFIJOS_SIN_CNC))
    comprados_pdf = [a for a in sin_cnc if pat_sin_cnc.match(a)]
    if comprados_pdf:
        sin_cnc = [a for a in sin_cnc if a not in comprados_pdf]
        add(INFO, "SUB-07", "Rejas / escalones de reja: se compran hechos, van con PDF y sin CNC.", comprados_pdf)
    if sin_cnc:
        add(ERROR, "SUB-02", f"{len(sin_cnc)} PDF de submaterial sin su CNC.", sin_cnc)
    if sin_pdf:
        add(ERROR, "SUB-03", f"{len(sin_pdf)} CNC sin su PDF de submaterial.", sin_pdf)
    ext_mal = [a for a in f.sub_pdf if not a.lower().endswith(".pdf")] + [a for a in f.cnc if not a.lower().endswith(".nc1")]
    if ext_mal:
        add(ERROR, "SUB-04", "Archivo con extensión equivocada en submateriales o CNC.", ext_mal)

    # --- típicos
    pat_tip = re.compile(r"^[A-Z]+%s\d+$" % re.escape(f.letra))
    tip_mal = [a for a in f.tipicos if not a.lower().endswith(".pdf") or not pat_tip.match(os.path.splitext(a)[0])]
    if tip_mal:
        add(AVISO, "TIP-01", f"Típico con nombre fuera de regla (<tipo>{f.letra}<n>, ej. C{f.letra}73).", tip_mal)
    if (f.sub_pdf or f.cnc) and not f.tipicos:
        add(AVISO, "TIP-02", "Hay submateriales pero ningún típico.")

    # --- reportes
    pat_rep = re.compile(r"^%s_(REPORTE DE (?:SUBMATERIALES|PESO|BULONES OBRA|BULONES TALLER|TUERCAS Y ARANDELAS))_FASE %s_(.+)\.xlsx$" % (re.escape(obra), NN))
    tipos_ok, rep_mal, cat_mal = set(), [], []
    for a in f.reportes:
        m = pat_rep.match(a)
        if m:
            tipos_ok.add(m.group(1))
            if categoria and norm(m.group(2)) != norm(categoria):
                cat_mal.append(f"{a}  (planilla: {categoria.strip()})")
        else:
            rep_mal.append(a)
            u = norm(a)
            for k, v in TIPOS_REPORTE.items():   # reconoce el tipo aunque el nombre esté mal
                if k.split()[0] in u and (k != "BULONES OBRA" or "TALLER" not in u):
                    tipos_ok.add(v + "*")
                    break
    if rep_mal:
        add(ERROR, "REP-01", f"Reporte con nombre o formato fuera de regla ({obra}_REPORTE DE <TIPO>_FASE {NN}_<CATEGORIA>.xlsx).", sorted(set(rep_mal)))
    if cat_mal:
        add(AVISO, "REP-02", "La categoría del nombre del reporte no coincide con la planilla.", sorted(set(cat_mal)))
    dobles = [f"{a}  (en {' y '.join(c)})" for a, c in f.donde.items() if len(c) > 1]
    if dobles:
        add(AVISO, "REP-06", "El mismo reporte está en más de una carpeta.", dobles)
    tiene = lambda t: t in tipos_ok or (t + "*") in tipos_ok
    n_sub = len(set(base_pdf) | set(base_cnc))
    if n_sub > 1 and not tiene("REPORTE DE SUBMATERIALES"):
        add(ERROR, "REP-03", "Falta el reporte de submateriales.")
    elif n_sub == 1 and not tiene("REPORTE DE SUBMATERIALES"):
        # Regla 2026-11-05: con un solo submaterial no se sube (su información ya está en los típicos)
        add(INFO, "REP-07", "Un solo submaterial: no lleva reporte de submateriales.")
    if f.tipicos and not tiene("REPORTE DE PESO"):
        add(ERROR, "REP-04", "Falta el reporte de peso (conjuntos).")
    if not any(tiene(TIPOS_REPORTE[k]) for k in ("BULONES OBRA", "BULONES TALLER")):
        add(AVISO, "REP-05", "No hay reporte de bulones. ¿La fase no lleva bulones? Confirmar.")

    # --- Strumis
    # Se aceptan ENT 01 y ENT01; la forma unificada es "ENT 01" (decidido 2026-11-04)
    pat_str = re.compile(r"^%s_FASE %s_ENT ?(\d{2})\.bswx$" % (re.escape(obra), NN))
    if not f.strumis:
        add(ERROR, "STR-01", "Falta el Strumis (.bswx). Se exporta a mano, nunca se renumera.")
    else:
        mal = [a for a in f.strumis if not pat_str.match(a)]
        if mal:
            add(ERROR, "STR-02", f"Strumis con nombre fuera de regla ({obra}_FASE {NN}_ENT 01.bswx).", mal)
        sin_espacio = [a for a in f.strumis if pat_str.match(a) and "_ENT " not in a]
        if sin_espacio and modo == "zona":
            add(INFO, "STR-04", "Strumis con 'ENT01': se acepta y al subir se unifica a 'ENT 01'.", sin_espacio)
    if modo == "servidor" and f.carpeta_strumis and categoria:
        esperado = f"FASE {NN}_{categoria.strip().upper()}"
        if norm(f.carpeta_strumis) != norm(esperado):
            add(INFO, "STR-03", f"Carpeta de Strumis '{f.carpeta_strumis}'. Con la regla nueva sería '{esperado}'.")

    # --- archivos que no se suben
    if f.otros:
        add(ERROR if modo == "zona" else AVISO, "BAS-01", "Archivos que no corresponden a la fase (no se suben).", f.otros)
    return h


# ---------------------------------------------------------------- salida
def informe_md(resultados, titulo):
    out = [f"# {titulo}", ""]
    tot = {ERROR: 0, AVISO: 0, INFO: 0}
    for f, hs in resultados:
        for x in hs:
            tot[x.nivel] += 1
    out.append(f"**{tot[ERROR]} errores · {tot[AVISO]} avisos · {tot[INFO]} info** en {len(resultados)} fase(s).")
    out.append("")
    out.append("| Fase | Sub PDF | CNC | Típicos | Reportes | Strumis | Errores | Avisos |")
    out.append("|---|---|---|---|---|---|---|---|")
    for f, hs in resultados:
        e = sum(1 for x in hs if x.nivel == ERROR)
        a = sum(1 for x in hs if x.nivel == AVISO)
        marca = "❌" if e else ("⚠️" if a else "✅")
        out.append(f"| {marca} {f.numero} | {len(f.sub_pdf)} | {len(f.cnc)} | {len(f.tipicos)} | {len(f.reportes)} | {len(f.strumis)} | {e} | {a} |")
    for f, hs in resultados:
        if not hs:
            continue
        out += ["", f"## Fase {f.numero}"]
        for x in sorted(hs, key=lambda x: [ERROR, AVISO, INFO].index(x.nivel)):
            out.append(f"- **{x.nivel}** `{x.codigo}` {x.texto}")
            for a in x.archivos[:15]:
                out.append(f"  - `{a}`")
            if len(x.archivos) > 15:
                out.append(f"  - … y {len(x.archivos) - 15} más")
    return "\n".join(out) + "\n"


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Control de fase (solo lectura)")
    ap.add_argument("modo", choices=["zona", "servidor"])
    ap.add_argument("--obra", required=True, help="código corto, ej. 674L")
    ap.add_argument("--planilla", required=True, help="CSV de la pestaña de la obra")
    ap.add_argument("--fase-dir", help="modo zona: carpeta FASE xx en la zona compartida")
    ap.add_argument("--constructivos", help="modo servidor: carpeta CONSTRUCTIVOS de la obra")
    ap.add_argument("--fases", help="modo servidor: lista 5,9,22 (por defecto, todas las de la planilla)")
    ap.add_argument("--multi-letra", action="store_true", help="la obra tiene varias letras (subcarpeta por letra)")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--salida", help="guardar el informe .md en este archivo")
    a = ap.parse_args()

    filas, estados, repetidas = leer_planilla(a.planilla)
    resultados = []
    if a.modo == "zona":
        if not a.fase_dir or not os.path.isdir(a.fase_dir):
            print("Falta --fase-dir o no existe.", file=sys.stderr)
            return 2
        m = re.search(r"(\d+)\s*$", os.path.basename(os.path.normpath(a.fase_dir)))
        if not m:
            print("La carpeta de fase debe terminar en el número (FASE 05).", file=sys.stderr)
            return 2
        f = fase_desde_zona(a.fase_dir, a.obra, int(m.group(1)))
        resultados.append((f, controlar(f, filas, estados, "zona")))
        titulo = f"Control de fase {a.obra} FASE {f.numero:02d} (zona compartida)"
    else:
        if not a.constructivos or not os.path.isdir(a.constructivos):
            print("Falta --constructivos o no existe.", file=sys.stderr)
            return 2
        nums = [int(x) for x in a.fases.split(",")] if a.fases else sorted(filas)
        for n in nums:
            f = fase_desde_servidor(a.constructivos, a.obra, n, a.multi_letra)
            if not (f.sub_pdf or f.cnc or f.tipicos or f.reportes or f.strumis) and n in filas and not a.fases:
                continue  # en la planilla pero sin nada subido: no se audita
            resultados.append((f, controlar(f, filas, estados, "servidor")))
        titulo = f"Auditoría de {a.obra} en el servidor"

    if a.json:
        print(json.dumps([{"fase": f.numero, "hallazgos": [x.__dict__ for x in hs]} for f, hs in resultados], ensure_ascii=False, indent=1))
    else:
        md = informe_md(resultados, titulo)
        print(md)
        if a.salida:
            with open(a.salida, "w", encoding="utf-8") as fo:
                fo.write(md)
    return 1 if any(x.nivel == ERROR for _, hs in resultados for x in hs) else 0


if __name__ == "__main__":
    sys.exit(main())
