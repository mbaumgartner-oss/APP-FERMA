# -*- coding: utf-8 -*-
"""
Arma el PLAN de subida de una fase copiando la convención que la obra YA TIENE en el servidor
(decidido 2026-11-05: "seguí su propia estructura"). Solo lee: no copia nada. El plan lo ejecuta subir_plan.py.

Detecta en CONSTRUCTIVOS (mirando las fases ya subidas):
  - carpeta de planos: PDF o DXF (en SUBMATERIALES y en TIPICOS)
  - nombre de carpeta de fase: "FASE 5 - <CATEGORIA>" o "FASE 05"
  - reportes sueltos en REPORTES\\ o en REPORTES\\FASE xx\\
  - carpeta de bulones: BULONES o REPORTE DE BULONES
  - nombres de reportes y de Strumis (copia el patrón de uno existente)
Si la obra no tiene nada subido todavía, usa el estándar nuevo (674).

Origen: carpeta de la fase exportada (ej. OBRAS\\PASADO A CONTROL\\702Q\\FASE 50 - TENSORES AUXILIARES\\) con
CNC\\, SUBMATERIALES\\, TIPICOS\\, STRUMIS\\*.bswx y reportes (con el nombre de la plantilla de Tekla o ya renombrados).

Uso:
  python plan_obra.py --obra 702Q --fase 50 --fase-dir "<...>\\FASE 50 - TENSORES AUXILIARES" \
      --constructivos "<...>\\CONSTRUCTIVOS" [--categoria "TENSORES AUXILIARES" | --pestana 702Q] --plan "<salida.txt>"
Código de salida: 0 plan listo · 1 errores (no se arma el plan) · 2 uso
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from control_fase import PREFIJOS_COMPRADOS, PREFIJOS_SIN_CNC, archivos, dir_entrada, norm, subcarpeta_fase  # noqa: E402

# Rejas / escalones de reja (<fase>_gr...): se compran hechos, llevan PDF y NO llevan CNC (2026-11-10)
SIN_CNC = re.compile(r"^\d+_(%s)[A-Z]" % "|".join(PREFIJOS_SIN_CNC))

# tipo -> (palabras que lo identifican en el nombre del export, texto estándar, carpeta)
TIPOS = [
    ("SUB", ["SUBMATERIALES"], "REPORTE DE SUBMATERIALES", ("SUBMATERIALES", "REPORTES")),
    ("PESO", ["CONJUNTOS"], "REPORTE DE PESO", ("TIPICOS", "REPORTES")),
    ("PESO", ["PESO"], "REPORTE DE PESO", ("TIPICOS", "REPORTES")),
    ("TUERCAS", ["TUERCAS"], "REPORTE DE TUERCAS Y ARANDELAS", ("BULONES",)),
    ("BUL_TALLER", ["BULONES", "TALLER"], "REPORTE DE BULONES TALLER", ("BULONES",)),
    ("BUL_OBRA", ["BULONES"], "REPORTE DE BULONES OBRA", ("BULONES",)),
]


def tipo_de(nombre):
    u = norm(nombre)
    for clave, palabras, _, _ in TIPOS:
        if all(p in u for p in palabras):
            return clave
    return None


def limpiar_nombre(texto):
    """Categoría apta para nombre de carpeta: sin \\ / : * ? " < > | y sin espacios de más."""
    return re.sub(r"\s+", " ", re.sub(r'[\\/:*?"<>|]', "-", texto or "")).strip(" .")


def subdirs(p):
    return sorted(d for d in os.listdir(p) if os.path.isdir(os.path.join(p, d))) if os.path.isdir(p) else []


# Carpeta de letra de obra (obras con varias letras, ej. 670: "670U", "670K - TEKLA", "670K TEKLA", "678Z").
LETRA_RE = re.compile(r"^(\d+[A-Z]+)(?=$|[\s_\-])", re.I)


def carpetas_letra(p):
    """Carpetas de letra de obra. (2026-11-14) No cuentan las que dicen FASE / REPORTE / OBS (en la 695 había una
    carpeta mal nombrada "695L_REPORTE DE SUBMATERIALES_FASE 40..." que hacía creer que la obra tenía varias letras)."""
    return [d for d in subdirs(p) if LETRA_RE.match(d) and not re.search(r"FASE|REPORTE|\bOBS\b", d, re.I)]


# Carpeta de fase: "FASE 05", "FASE 5 - CAT", "FASE 05_CAT", "FASE 41-COLUMNAS"
FASE_DIR_RE = re.compile(r"^FASE ?(\d+)(?:\s*[-_]\s*(\S.*))?$", re.I)
SEP_FASE = " - "        # (2026-11-14) las carpetas nuevas llevan siempre la categoría: "FASE 46 - ESCALERAS"


class Convencion:
    def __init__(self, c, obra):
        self.c, self.obra = c, obra
        self.avisos, self.errores = [], []
        self.nuevas_letra = []          # carpetas de letra que hay que crear (no existen todavía)
        self.pl_sub = "DXF" if os.path.isdir(os.path.join(c, "SUBMATERIALES", "DXF")) and not \
            os.path.isdir(os.path.join(c, "SUBMATERIALES", "PDF")) else "PDF"
        self.pl_tip = "DXF" if os.path.isdir(os.path.join(c, "TIPICOS", "DXF")) and not \
            os.path.isdir(os.path.join(c, "TIPICOS", "PDF")) else "PDF"
        self.bulones = "BULONES" if os.path.isdir(os.path.join(c, "BULONES")) else "REPORTE DE BULONES"
        # ¿Obra con varias letras? Las carpetas de letra pueden ir ANTES (TIPICOS\678Z\PDF) o DESPUÉS (TIPICOS\PDF\670U)
        antes = carpetas_letra(os.path.join(c, "SUBMATERIALES")) + carpetas_letra(os.path.join(c, "TIPICOS"))
        despues = carpetas_letra(os.path.join(c, "SUBMATERIALES", self.pl_sub)) + \
            carpetas_letra(os.path.join(c, "TIPICOS", self.pl_tip))
        self.modo_letra = "antes" if antes else ("despues" if despues else None)
        self.multi = self.modo_letra is not None
        if self.multi:
            # carpetas de fase: se miran dentro de las carpetas de letra existentes
            fases = []
            for base in (("SUBMATERIALES", self.pl_sub), ("TIPICOS", self.pl_tip)):
                for L in self._letras_de(base):
                    fases += subdirs(os.path.join(c, self._unir(base, L)))
        else:
            fases = subdirs(os.path.join(c, "SUBMATERIALES", self.pl_sub)) + subdirs(os.path.join(c, "TIPICOS", self.pl_tip))
        # (2026-11-14, pedido del usuario) Las carpetas de fase nuevas llevan SIEMPRE la categoría ("FASE 46 - CAT").
        # De la obra solo se copia si el número va con cero adelante (FASE 05) o no (FASE 5).
        nums = [m.group(1) for m in (FASE_DIR_RE.match(d) for d in fases) if m]
        self.estilo_fase = "CAT"
        con_cero = sum(1 for x in nums if len(x) > 1 and x.startswith("0"))
        sin_cero = sum(1 for x in nums if len(x) == 1)
        self.pad = 2 if con_cero >= sin_cero else 0
        # Reportes sueltos (695, 702) o en subcarpeta por fase (674): manda lo que más hay en la obra.
        rep = os.path.join(c, self.ruta(("SUBMATERIALES", "REPORTES"), _nueva=False))
        sueltos = [x for x in archivos(rep) if re.search(r"FASE ?\d+", x, re.I) and x.lower().endswith((".xls", ".xlsx", ".xlsm", ".csv"))]
        por_fase = [d for d in subdirs(rep) if FASE_DIR_RE.match(d)]
        self.rep_por_fase = len(por_fase) > len(sueltos) or not (sueltos or por_fase)
        # patrones de nombre de reportes: se copia el "tipo" tal como lo escribe la obra
        self.texto_tipo, self.ext_rep = {}, None
        carpetas = [rep, os.path.join(c, self.ruta(("TIPICOS", "REPORTES"), _nueva=False)),
                    os.path.join(c, self.ruta((self.bulones,), _nueva=False))]
        for carpeta in list(carpetas):
            carpetas += [os.path.join(carpeta, d) for d in subdirs(carpeta)]     # reportes dentro de FASE xx
        for carpeta in carpetas:
            for x in archivos(carpeta):
                m = re.match(r"^%s_(.+?)_FASE \d+_.+\.(xlsx?)$" % re.escape(obra), x, re.I)
                t = tipo_de(x)
                if m and t and t not in self.texto_tipo:
                    self.texto_tipo[t] = m.group(1)
                    self.ext_rep = self.ext_rep or m.group(2).lower()
        # patrón de Strumis: se toma el primer .bswx de cualquier fase (de esta letra si la obra tiene varias)
        self.str_patron = None
        base_str = os.path.join(c, self.ruta(("STRUMIS",), _nueva=False))
        for d in subdirs(base_str):
            for x in archivos(os.path.join(base_str, d), ".bswx"):
                p = re.sub(r"(FASE ?)\d+", r"\1{F}", x, flags=re.I)
                p = re.sub(r"(ENT ?)\d+", r"\1{E}", p, flags=re.I)
                if "{F}" in p and "{E}" in p:
                    self.str_patron = p
                    self.str_pad = len(re.search(r"FASE ?(\d+)", x, re.I).group(1))
                    break
            if self.str_patron:
                break

    # ---- carpetas de letra
    def _unir(self, base, L):
        base = tuple(base)
        if self.modo_letra == "antes" and len(base) > 1:
            return os.path.join(base[0], L, *base[1:])
        return os.path.join(*base, L)

    def _padre_letra(self, base):
        base = tuple(base)
        return os.path.join(self.c, base[0]) if self.modo_letra == "antes" and len(base) > 1 else os.path.join(self.c, *base)

    def _letras_de(self, base):
        return carpetas_letra(self._padre_letra(base))

    def letra_en(self, base, _nueva=True):
        """Nombre de la carpeta de letra de ESTA obra dentro de base. Si hay varias (ej. '670K - SDS' y
        '670K - TEKLA') se prefiere la de TEKLA. Si no existe, se usa el código de obra (ej. '670T') y se anota
        para crearla."""
        propias = [d for d in self._letras_de(base) if LETRA_RE.match(d).group(1).upper() == self.obra.upper()]
        if len(propias) > 1:
            tekla = [d for d in propias if "TEKLA" in d.upper()]
            if len(tekla) == 1:
                msg = f"{os.path.join(*base)}: hay {', '.join(propias)}; se usa '{tekla[0]}'"
                if msg not in self.avisos:
                    self.avisos.append(msg)
                return tekla[0]
            msg = f"{os.path.join(*base)}: hay varias carpetas para {self.obra} ({', '.join(propias)}); no sé cuál usar"
            if msg not in self.errores:
                self.errores.append(msg)
            return propias[0]
        if propias:
            return propias[0]
        if _nueva:
            p = os.path.join(self._padre_letra(base), self.obra)
            if p not in self.nuevas_letra:
                self.nuevas_letra.append(p)
        return self.obra

    def ruta(self, base, *resto, _nueva=True):
        """Ruta relativa a CONSTRUCTIVOS: base (ej. ('TIPICOS','PDF')) + carpeta de letra si la obra tiene varias + resto."""
        if not self.multi:
            return os.path.join(*base, *resto)
        return os.path.join(self._unir(base, self.letra_en(base, _nueva)), *resto)

    def fase_dir(self, n, cat):
        num = f"{n:0{self.pad}d}" if self.pad else str(n)
        cat = limpiar_nombre(cat)
        return f"FASE {num}{SEP_FASE}{cat}" if cat else f"FASE {num}"

    def nombre_rep(self, tipo, n, cat, ext_origen):
        estandar = next(t[2] for t in TIPOS if t[0] == tipo)
        texto = self.texto_tipo.get(tipo, estandar)
        num = f"{n:0{self.pad}d}" if self.pad else str(n)
        return f"{self.obra}_{texto}_FASE {num}_{cat}.{self.ext_rep or ext_origen}"

    def carpeta_rep(self, tipo, n, cat):
        base = next(t[3] for t in TIPOS if t[0] == tipo)
        if base == ("BULONES",):
            base = (self.bulones,)
        return self.ruta(base, self.fase_dir(n, cat)) if self.rep_por_fase else self.ruta(base)

    def nombre_strumis(self, n, ent):
        if not self.str_patron:
            return f"{self.obra}_FASE {n:02d}_ENT {ent:02d}.bswx"
        return self.str_patron.replace("{F}", f"{n:0{self.str_pad}d}").replace("{E}", f"{ent:02d}")


def armar(obra, n, fd, c, categoria=None, pestana=None, libro="real"):
    """Arma el plan de subida con la convención de la obra. Solo lee.
    Devuelve dict: obra, n, cat, conv, fdir, plan [(origen, destino_relativo)], errores, avisos, excluidos.
    Error de uso (carpetas que no existen, sin categoría) -> ValueError."""
    obra = obra.upper()
    if not os.path.isdir(fd) or not os.path.isdir(c):
        raise ValueError("No existe --fase-dir o --constructivos.")
    cat = (categoria or "").strip()
    if not cat and pestana:
        os.environ["FERMA_PLANILLA"] = libro
        import planilla  # noqa: E402
        filas, _, _ = planilla.leer_pestana(pestana)
        cat = filas[n][0].categoria.strip() if filas.get(n) else ""
    if not cat:
        m = re.match(r"^FASE\s*\d+\s*-\s*(.+)$", os.path.basename(os.path.normpath(fd)))
        cat = m.group(1).strip() if m else ""
    if not cat:
        raise ValueError("No se pudo saber la CATEGORIA (pasar --categoria o --pestana).")
    cat = re.sub(r"\s+", " ", cat).upper()

    conv = Convencion(c, obra)
    errores, avisos, excluidos, plan = [], [], [], []
    comprado = re.compile(r"^\d+_(%s)" % "|".join(PREFIJOS_COMPRADOS))
    de_esta_fase = re.compile(r"^0*%d_" % n)
    fdir = conv.fase_dir(n, cat)

    def fd_en(base):
        """Carpeta de la fase dentro de base: la que ya exista (con cualquier nombre) o la nueva con categoría."""
        ya = subcarpeta_fase(os.path.join(c, conv.ruta(base, _nueva=False)), n)
        return os.path.basename(ya) if ya else fdir

    def base(x):
        return re.sub(r"_\d+$", "", os.path.splitext(x)[0])

    pdfs, cncs = {}, {}
    d_sub, d_cnc, d_tip = (dir_entrada(fd, k) for k in ("SUBMATERIALES", "CNC", "TIPICOS"))
    for x in archivos(d_sub):
        if comprado.match(x):
            excluidos.append(f"SUBMATERIALES\\{x} (comprado)")
        elif not de_esta_fase.match(x):
            excluidos.append(f"SUBMATERIALES\\{x} (es de otra fase)")
        else:
            pdfs[base(x)] = x
    for x in archivos(d_cnc):
        if comprado.match(x):
            excluidos.append(f"CNC\\{x} (comprado)")
        elif not de_esta_fase.match(x):
            excluidos.append(f"CNC\\{x} (es de otra fase)")
        else:
            cncs[base(x)] = x
    for b in sorted(set(pdfs) ^ set(cncs)):
        if b in pdfs and SIN_CNC.match(b):
            continue                      # reja: PDF sí, CNC no
        errores.append(f"{b}: {'falta el CNC' if b in pdfs else 'falta el PDF'}")
    rejas = sorted(b for b in pdfs if SIN_CNC.match(b) and b not in cncs)
    if rejas:
        avisos.append("Rejas / escalones de reja (se compran hechos): van con PDF y sin CNC: " + ", ".join(rejas))
    for b, x in sorted(pdfs.items()):
        plan.append((os.path.join(d_sub, x), conv.ruta(("SUBMATERIALES", conv.pl_sub), fd_en(("SUBMATERIALES", conv.pl_sub)), x)))
    for b, x in sorted(cncs.items()):
        plan.append((os.path.join(d_cnc, x), conv.ruta(("SUBMATERIALES", "CNC"), fd_en(("SUBMATERIALES", "CNC")), x)))
    tips = archivos(d_tip)
    for x in tips:
        plan.append((os.path.join(d_tip, x), conv.ruta(("TIPICOS", conv.pl_tip), fd_en(("TIPICOS", conv.pl_tip)), x)))
    # reportes
    vistos = {}
    for x in archivos(fd):
        if not x.lower().endswith((".xls", ".xlsx")):
            if not x.lower().endswith(("_log.txt", ".bswx")) and x.lower() != "_fase.json":
                excluidos.append(f"{x} (no es de ningún tipo)")
            continue
        t = tipo_de(x)
        if not t:
            avisos.append(f"Reporte sin tipo reconocido, no se sube: {x}")
            continue
        if t in vistos:
            errores.append(f"Dos reportes del mismo tipo: {vistos[t]} y {x}")
            continue
        vistos[t] = x
        ext = os.path.splitext(x)[1].lstrip(".").lower()
        destino = conv.nombre_rep(t, n, cat, ext)
        if conv.ext_rep and ext != conv.ext_rep:
            avisos.append(f"{x}: la obra usa .{conv.ext_rep}; se sube como .{ext}")
            destino = os.path.splitext(destino)[0] + "." + ext
        plan.append((os.path.join(fd, x), os.path.join(conv.carpeta_rep(t, n, cat), destino)))
    n_sub = len(set(pdfs) | set(cncs))
    if n_sub > 1 and "SUB" not in vistos:
        errores.append("Falta el reporte de submateriales.")
    if n_sub == 1 and "SUB" not in vistos:
        avisos.append("Un solo submaterial: no lleva reporte de submateriales (regla 2026-11-05).")
    if tips and "PESO" not in vistos:
        errores.append("Falta el reporte de peso / conjuntos.")
    if "BUL_OBRA" not in vistos and "BUL_TALLER" not in vistos:
        avisos.append("No hay reporte de bulones: ¿la fase no lleva bulones? Confirmar.")
    # Strumis (entrada siguiente a las que ya existan en esa carpeta)
    bswx = [x for x in archivos(fd, ".bswx")] + [os.path.join("STRUMIS", x) for x in archivos(os.path.join(fd, "STRUMIS"), ".bswx")]
    if not bswx:
        errores.append("Falta el Strumis (.bswx). Se exporta a mano y nunca se renumera.")
    elif len(bswx) > 1:
        errores.append(f"Hay más de un Strumis: {', '.join(bswx)}")
    else:
        car_str = conv.ruta(("STRUMIS",), fd_en(("STRUMIS",)))
        ents = [int(m.group(1)) for m in (re.search(r"ENT ?(\d+)", x, re.I) for x in archivos(os.path.join(c, car_str), ".bswx")) if m]
        plan.append((os.path.join(fd, bswx[0]), os.path.join(car_str, conv.nombre_strumis(n, max(ents, default=0) + 1))))

    # Obras con varias letras: la carpeta de la letra se crea SOLO si la obra está en la planilla de secuencias
    errores += conv.errores
    avisos += [x for x in conv.avisos if x not in avisos]
    if conv.nuevas_letra:
        try:
            os.environ.setdefault("FERMA_PLANILLA", libro)
            import planilla  # noqa: E402
            planilla.CONFIG = planilla.ruta_config(libro)
            planilla.leer_pestana(obra)          # sale con error si no existe la pestaña
            avisos.append(f"{obra} tiene pestaña en la planilla: se crean sus carpetas de letra:")
            avisos += ["   " + p for p in conv.nuevas_letra]
        except SystemExit as e:
            errores.append(f"La carpeta de letra {obra} no existe en el servidor y no se pudo confirmar la pestaña "
                           f"{obra} en la planilla ({e}). No se crea nada.")

    return {"obra": obra, "n": n, "cat": cat, "conv": conv, "fdir": fdir, "plan": plan,
            "errores": errores, "avisos": avisos, "excluidos": excluidos}


def texto_plan(r):
    """Líneas del resumen del plan (convención detectada, excluidos, avisos, errores)."""
    obra, n, cat, conv, fdir = r["obra"], r["n"], r["cat"], r["conv"], r["fdir"]
    excluidos, avisos, errores = r["excluidos"], r["avisos"], r["errores"]
    out = []
    out.append(f"# Plan {obra} FASE {n} - {cat}")
    out.append("## Convención detectada en el servidor")
    if conv.multi:
        out.append(f"- Obra con varias letras: carpeta de letra {'antes' if conv.modo_letra == 'antes' else 'después'} "
                   f"de PDF/CNC/REPORTES (ej. {conv.ruta(('TIPICOS', conv.pl_tip), _nueva=False)})")
    out.append(f"- Planos: SUBMATERIALES\\{conv.pl_sub} · TIPICOS\\{conv.pl_tip}")
    out.append(f"- Carpeta de fase: {fdir}")
    out.append(f"- Reportes: {'en subcarpeta por fase' if conv.rep_por_fase else 'sueltos en REPORTES'} · bulones en {conv.bulones}\\")
    out.append(f"- Nombres de reportes copiados de la obra: {conv.texto_tipo or '(ninguno: estándar nuevo)'}")
    out.append(f"- Strumis: {conv.str_patron or '(estándar nuevo)'}")
    if excluidos:
        out.append("## No se suben")
        for x in excluidos:
            out.append(f"- {x}")
    if avisos:
        out.append("## Avisos")
        for x in avisos:
            out.append(f"- {x}")
    if errores:
        out.append("## ERRORES (no se arma el plan)")
        for x in errores:
            out.append(f"- {x}")

    return out


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--obra", required=True)
    ap.add_argument("--fase", type=int, required=True)
    ap.add_argument("--fase-dir", required=True)
    ap.add_argument("--constructivos", required=True)
    ap.add_argument("--categoria")
    ap.add_argument("--pestana", help="lee la CATEGORIA de la planilla en vivo")
    ap.add_argument("--libro", choices=["real", "prueba"], default="real")
    ap.add_argument("--plan", required=True)
    a = ap.parse_args()
    try:
        r = armar(a.obra, a.fase, a.fase_dir, a.constructivos, a.categoria, a.pestana, a.libro)
    except ValueError as e:
        print(e, file=sys.stderr)
        return 2
    print("\n".join(texto_plan(r)))
    if r["errores"]:
        return 1
    obra, n, cat, plan, excluidos = r["obra"], r["n"], r["cat"], r["plan"], r["excluidos"]
    with open(a.plan, "w", encoding="utf-8") as f:
        f.write(f"# Plan {obra} FASE {n} - {cat} (armado por plan_obra.py con la convención de la obra)\n")
        for x in excluidos:
            f.write(f"# No se sube: {x}\n")
        for o, d in plan:
            f.write(f"{o} | {d}\n")
    print(f"## Plan: {len(plan)} archivos → {a.plan}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
