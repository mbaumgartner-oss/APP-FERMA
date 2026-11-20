# -*- coding: utf-8 -*-
"""
Reentrega de una fase ya subida - Skill "subir-fase" de Ferma (modo reentrega).
Reglas: vault Cerebro IA Ferma/WIKI/Relevamiento/Reglas de carga y revisiones.md (secciones 2, 5, 6 y 7).

Origen:  EN CONTROL\\<obra>\\FASE xx\\REENTREGA NN\\
           CNC\\ SUBMATERIALES\\ TIPICOS\\ <reportes rehechos> <.bswx nuevo>   <- solo lo que cambió
           LISTA CNC COMPLETA\\                                              <- export CNC de toda la fase (solo nombres)

Qué hace:
  - Archivo entregado que ya está en el servidor -> el viejo va a OBSOLETOS\\<nombre>_OBS AAAA-MM-DD y se sube el nuevo.
  - Archivo entregado que no está -> pieza nueva o renombrada: se agrega.
  - Pieza en el servidor que no está en LISTA CNC COMPLETA -> CANDIDATA a eliminar. Nunca se mueve sola:
    el modelador decide cada una con --retirar o --mantener.
  - Reporte rehecho con otro nombre (otra categoría) -> el del mismo tipo de esa fase va a OBSOLETOS.
  - Strumis -> entrada siguiente (ENT 02, 03...). Nada va a OBSOLETOS.
  - Planilla: la fase tiene que estar en SUBIDO AL SERVIDOR. No cambia de estado; con --pestana se agrega una nota.
Nunca borra. Sin --ejecutar solo simula. Para ejecutar hace falta --avise-checker (el modelador confirmó
en el chat que avisó al checker: única excepción a la regla de APROBADO, aprobada 2026-11-05).

Uso:
  python reentrega.py --obra 674L (--pestana 674L | --planilla 674L.csv) --reentrega-dir "<...>\\FASE 05\\REENTREGA 01" \
      --constructivos "<...>\\CONSTRUCTIVOS" [--retirar 5_pL20,CL73] [--mantener 5_pL21] [--sin-lista] \
      [--avise-checker --ejecutar] [--informe x.md]
Código de salida: 0 ok · 1 bloqueado por el control · 2 uso · 3 falta decidir (candidatos o lista)
                  5 subido pero planilla sin nota · 6 falta la confirmación del checker
"""
import argparse
import datetime
import getpass
import json
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from control_fase import (ERROR, AVISO, INFO, ESTADO_SUBIDO, PREFIJOS_COMPRADOS, PREFIJOS_SIN_CNC, PREFIJOS_SUB,  # noqa: E402
                          Hallazgo, archivos, dir_entrada, fase_desde_zona, leer_planilla, norm, subcarpeta_fase)
from plan_obra import Convencion, TIPOS, tipo_de  # noqa: E402
import reportes  # noqa: E402

# Tipos de reporte por palabras clave (sirve con nombres finales y con nombres de plantilla de Tekla,
# ej. FMA_LISTADO_CONJUNTOS_TOTALES_MFD.xls = peso). Ver plan_obra.TIPOS.
SUB_REP, PESO_REP = "SUB", "PESO"
NOMBRE_TIPO = {t[0]: t[2] for t in TIPOS}

LISTA = "LISTA CNC COMPLETA"
MARCA_SUBIDA = "_reentrega_subida.json"   # se escribe en REENTREGA NN al terminar: esa reentrega no se sube de nuevo


def base_pieza(nombre):
    """5_pL13.pdf -> 5_pL13 · 100_wL27_1.nc1 -> 100_wL27"""
    return re.sub(r"_\d+$", "", os.path.splitext(nombre)[0])


def tipo_reporte(nombre):
    return tipo_de(nombre) if nombre.lower().endswith((".xls", ".xlsx")) else None


def nombre_obs(carpeta_obs, nombre, hoy):
    """5_pL13.pdf -> 5_pL13_OBS 2026-11-05.pdf (y _2, _3... si ya existe ese día)."""
    stem, ext = os.path.splitext(nombre)
    cand = f"{stem}_OBS {hoy}{ext}"
    k = 2
    while os.path.exists(os.path.join(carpeta_obs, cand)):
        cand = f"{stem}_OBS {hoy}_{k}{ext}"
        k += 1
    return cand


class Servidor:
    """Carpetas de la fase en CONSTRUCTIVOS, con la convención de la obra (plan_obra.Convencion):
    PDF o DXF, 'FASE 05' o 'FASE 5 - <CAT>', reportes sueltos o por fase, BULONES o REPORTE DE BULONES.
    Si la carpeta de la fase ya existe (con cualquier nombre), se usa esa."""

    def __init__(self, constructivos, obra, n, categoria, multi_letra):
        self.c, self.n, self.NN = constructivos, n, "%02d" % n
        self.categoria = categoria
        # La carpeta de letra (obras con varias letras: 670U, 670K - TEKLA…) la resuelve la Convencion sola.
        self.conv = Convencion(constructivos, obra)

    def carpeta(self, primera, *resto, nombre_nuevo=None):
        base = os.path.join(self.c, self.conv.ruta((primera,) + resto))
        return subcarpeta_fase(base, self.n) or os.path.join(base, nombre_nuevo or self.conv.fase_dir(self.n, norm(self.categoria)))

    def pdf(self):
        return self.carpeta("SUBMATERIALES", self.conv.pl_sub)

    def cnc(self):
        return self.carpeta("SUBMATERIALES", "CNC")

    def tip(self):
        return self.carpeta("TIPICOS", self.conv.pl_tip)

    def rep(self, tipo):
        base = next(t[3] for t in TIPOS if t[0] == tipo)
        if base == ("BULONES",):
            base = (self.conv.bulones,)
        if self.conv.rep_por_fase:
            return self.carpeta(*base)
        return os.path.join(self.c, self.conv.ruta(base))      # reportes sueltos (ej. 702)

    def nombre_rep(self, tipo, ext):
        return self.conv.nombre_rep(tipo, self.n, norm(self.categoria), ext)

    def strumis(self):
        nuevo = self.conv.fase_dir(self.n, norm(self.categoria)) if self.conv.estilo_fase == "CAT" \
            else f"FASE {self.NN}_{norm(self.categoria)}"
        return self.carpeta("STRUMIS", nombre_nuevo=nuevo)

    def nombre_strumis(self, ent):
        return self.conv.nombre_strumis(self.n, ent)


def controlar_entrega(f, obra, NN):
    """Control liviano: en una reentrega llega solo lo que cambió, así que no se exige la fase completa."""
    h = []
    add = lambda *a: h.append(Hallazgo(*a))
    if not (f.sub_pdf or f.cnc or f.tipicos or f.reportes or f.strumis):
        add(ERROR, "RE-01", "La carpeta de reentrega no tiene nada para subir.")
    pat_sub = re.compile(r"^(\d+)_([a-z]+)%s(\d+)$" % re.escape(f.letra))
    comprado = re.compile(r"^\d+_(%s)" % "|".join(PREFIJOS_COMPRADOS))
    comprados = [a for a in f.sub_pdf + f.cnc if comprado.match(a)]
    if comprados:
        add(INFO, "SUB-06", "Tuercas y arandelas compradas: no se suben.", comprados)
    pdf = {base_pieza(a): a for a in f.sub_pdf if a not in comprados}
    cnc = {base_pieza(a): a for a in f.cnc if a not in comprados}
    malos = [a for b, a in list(pdf.items()) + list(cnc.items())
             if not pat_sub.match(b) or int(pat_sub.match(b).group(1)) != f.numero]
    if malos:
        add(ERROR, "SUB-01", f"Nombre de submaterial fuera de regla (<fase>_<prefijo>{f.letra}<n>).", sorted(set(malos)))
    raros = [a for b, a in pdf.items() if pat_sub.match(b) and pat_sub.match(b).group(2) not in PREFIJOS_SUB]
    if raros:
        add(AVISO, "SUB-05", "Prefijo que no está en la lista de prefijos. Confirmar.", raros)
    # Regla: el PDF y el CNC de un submaterial van siempre juntos (también a OBSOLETOS)
    # Rejas (<fase>_gr...): se compran hechas, PDF sí y CNC no (2026-11-10)
    sin_cnc_ok = re.compile(r"^\d+_(%s)[A-Z]" % "|".join(PREFIJOS_SIN_CNC))
    sin_cnc = sorted(pdf[b] for b in pdf if b not in cnc and not sin_cnc_ok.match(b))
    sin_pdf = sorted(cnc[b] for b in cnc if b not in pdf)
    if sin_cnc:
        add(ERROR, "RE-02", "Submaterial entregado sin su CNC (en una reentrega van siempre juntos).", sin_cnc)
    if sin_pdf:
        add(ERROR, "RE-03", "CNC entregado sin su PDF (en una reentrega van siempre juntos).", sin_pdf)
    ext_mal = [a for a in f.sub_pdf if not a.lower().endswith(".pdf")] + [a for a in f.cnc if not a.lower().endswith(".nc1")]
    if ext_mal:
        add(ERROR, "SUB-04", "Extensión equivocada en submateriales o CNC.", ext_mal)
    pat_tip = re.compile(r"^[A-Z]+%s\d+$" % re.escape(f.letra))
    tip_mal = [a for a in f.tipicos if not a.lower().endswith(".pdf") or not pat_tip.match(os.path.splitext(a)[0])]
    if tip_mal:
        add(AVISO, "TIP-01", f"Típico con nombre fuera de regla (<tipo>{f.letra}<n>).", tip_mal)
    # Reportes: se aceptan con nombre final o con el de la plantilla de Tekla; al subir se renombran
    rep_mal = [a for a in f.reportes if not tipo_reporte(a)]
    if rep_mal:
        add(ERROR, "REP-01", "Reporte de tipo no reconocido (submateriales, peso/conjuntos, bulones, tuercas).", rep_mal)
    # Strumis: cualquier nombre (ej. STRUMIS.bswx); al subir toma el nombre de la entrada siguiente
    tipos = [tipo_reporte(a) for a in f.reportes if tipo_reporte(a)]
    dobles = sorted({t for t in tipos if tipos.count(t) > 1})
    if dobles:
        add(ERROR, "RE-10", "Hay más de un reporte del mismo tipo en la reentrega.", dobles)
    if len(f.strumis) > 1:
        add(ERROR, "RE-04", "Hay más de un Strumis en la reentrega.", f.strumis)
    if f.otros:
        add(ERROR, "BAS-01", "Archivos que no corresponden (no se suben).", f.otros)
    return h, comprados


def main():
    """(2026-11-20) Todo lo que se muestra queda también en ~/.ferma/registros/AAAA-MM-DD.log."""
    aqui = os.path.dirname(os.path.abspath(__file__))
    if aqui not in sys.path:
        sys.path.insert(0, aqui)
    import registro
    return registro.con_registro("reentrega " + " ".join(sys.argv[1:3]), _main)


def _main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Reentrega de una fase ya subida")
    ap.add_argument("--obra", required=True)
    fuente = ap.add_mutually_exclusive_group(required=True)
    fuente.add_argument("--planilla", help="CSV de la pestaña (sin nota en la planilla)")
    fuente.add_argument("--pestana", help="planilla EN VIVO: lee y, al ejecutar, agrega la nota")
    ap.add_argument("--libro", choices=["real", "prueba"], default="real", help="con --pestana: planilla real (por defecto) o PRUEBA OBRAS")
    ap.add_argument("--reentrega-dir", required=True, help="carpeta ...\\FASE 05\\REENTREGA 01")
    ap.add_argument("--constructivos", required=True)
    ap.add_argument("--multi-letra", action="store_true")
    ap.add_argument("--retirar", default="", help="piezas eliminadas confirmadas: van a OBSOLETOS (ej. 5_pL20,CL73)")
    ap.add_argument("--mantener", default="", help="candidatas que el modelador decidió dejar")
    ap.add_argument("--sin-lista", action="store_true", help="no hay LISTA CNC COMPLETA: el modelador confirmó qué retirar a mano")
    ap.add_argument("--confirmo", default="", help="avisos confirmados por el modelador (ej. RE-09)")
    ap.add_argument("--avise-checker", action="store_true", help="el modelador confirmó en el chat que avisó al checker")
    ap.add_argument("--ejecutar", action="store_true")
    ap.add_argument("--informe")
    a = ap.parse_args()

    rd = os.path.normpath(a.reentrega_dir)
    m_re = re.search(r"REENTREGA\s*(\d+)$", os.path.basename(rd), re.I)
    m_fa = re.match(r"^\s*(?:FASE|SEC|SECUENCIA)\s*0*(\d+)", os.path.basename(os.path.dirname(rd)), re.I) \
        or re.search(r"(\d+)\s*$", os.path.basename(os.path.dirname(rd)))
    if not os.path.isdir(rd) or not m_re or not m_fa:
        print("--reentrega-dir tiene que existir y ser <...>\\FASE NN\\REENTREGA NN.", file=sys.stderr)
        return 2
    if not os.path.isdir(a.constructivos):
        print("No existe --constructivos.", file=sys.stderr)
        return 2
    if os.path.isfile(os.path.join(rd, MARCA_SUBIDA)):
        with open(os.path.join(rd, MARCA_SUBIDA), encoding="utf-8") as fo:
            ya = json.load(fo)
        print(f"BLOQUEADO: {os.path.basename(rd)} ya se subió el {ya.get('fecha')} ({ya.get('usuario')}). "
              "Los cambios nuevos van en la REENTREGA siguiente.")
        return 1
    n, nre = int(m_fa.group(1)), int(m_re.group(1))
    NN, obra = "%02d" % n, a.obra.upper()
    hoy = datetime.date.today().isoformat()

    if a.pestana:
        os.environ["FERMA_PLANILLA"] = a.libro
        import planilla  # noqa: E402
        print(f"[planilla en vivo: {a.libro}]")
        filas, _, _ = planilla.leer_pestana(a.pestana)
    else:
        filas, _, _ = leer_planilla(a.planilla)

    # ---- 1. control
    hall = []
    fp = filas.get(n)
    categoria = fp[0].categoria if fp else ""
    if not fp:
        hall.append(Hallazgo(ERROR, "PLAN-01", f"La fase {n} no está en la planilla."))
    elif fp[0].estado != ESTADO_SUBIDO:
        hall.append(Hallazgo(ERROR, "RE-05", f"Estado '{fp[0].estado or '(vacío)'}': una reentrega solo se hace sobre una fase "
                                             f"en '{ESTADO_SUBIDO}'. Si es la primera subida, usar subir.py."))
    f = fase_desde_zona(rd, obra, n)
    # export tal como sale de Tekla: reportes con nombre de plantilla (ej. FMA_LISTADO_CONJUNTOS...) y STRUMIS\*.bswx
    for x in list(f.otros):
        if tipo_reporte(x):
            f.otros.remove(x)
            f.reportes.append(x)
    f.strumis += [os.path.join("STRUMIS", x) for x in archivos(os.path.join(rd, "STRUMIS"), ".bswx")]
    h2, comprados = controlar_entrega(f, obra, NN)
    hall += h2
    errores = [x for x in hall if x.nivel == ERROR]
    if errores:
        print(f"BLOQUEADO: {len(errores)} error(es) en la reentrega:")
        for x in hall:
            print(f"- {x.nivel} {x.codigo} {x.texto}" + "".join(f"\n    {y}" for y in x.archivos[:10]))
        return 1

    srv = Servidor(a.constructivos, obra, n, categoria, a.multi_letra)
    # plan: (accion, origen, carpeta_destino, nombre_destino)   accion: REEMPLAZA / NUEVO / RETIRA
    plan = []

    def entregar(origen, carpeta, nombre, por_base=False):
        plan.append(("REEMPLAZA" if os.path.exists(os.path.join(carpeta, nombre)) else "NUEVO", origen, carpeta, nombre))
        if por_base:   # misma pieza con otro nombre de archivo (ej. 100_wL27.nc1 vs 100_wL27_1.nc1)
            for viejo in archivos(carpeta):
                if viejo != nombre and base_pieza(viejo) == base_pieza(nombre) and \
                        os.path.splitext(viejo)[1].lower() == os.path.splitext(nombre)[1].lower():
                    plan.append(("RETIRA", None, carpeta, viejo))

    for x in f.sub_pdf:
        if x not in comprados:
            entregar(os.path.join(dir_entrada(rd, "SUBMATERIALES"), x), srv.pdf(), x, por_base=True)
    for x in f.cnc:
        if x not in comprados:
            entregar(os.path.join(dir_entrada(rd, "CNC"), x), srv.cnc(), x, por_base=True)
    for x in f.tipicos:
        entregar(os.path.join(dir_entrada(rd, "TIPICOS"), x), srv.tip(), x, por_base=True)
    # ---- reportes: en toda reentrega van TODOS; se comparan con los del servidor (sección 8)
    pat_fase = re.compile(r"_FASE 0*%d[_ .]" % n)
    srv_reps = {}
    for t in dict.fromkeys(x[0] for x in TIPOS):
        car = srv.rep(t)
        for viejo in archivos(car):
            if tipo_reporte(viejo) == t and pat_fase.search(viejo):
                srv_reps.setdefault(t, []).append((car, viejo))
    entregados_t = {tipo_reporte(x): x for x in f.reportes}
    faltan_rep = sorted(t for t in srv_reps if t not in entregados_t)
    if faltan_rep:
        print("BLOQUEADO: en toda reentrega van TODOS los reportes de la fase. Faltan:")
        for t in faltan_rep:
            print(f"    {NOMBRE_TIPO.get(t, t)}   (en el servidor: {', '.join(x for _, x in srv_reps[t])})")
        return 1
    rep_res, difs = {}, {}            # tipo -> texto · tipo -> resultado de comparar
    for t, x in sorted(entregados_t.items()):
        car = srv.rep(t)
        viejos = srv_reps.get(t, [])
        final = srv.nombre_rep(t, os.path.splitext(x)[1].lstrip(".").lower())   # nombre según la obra
        if not viejos:
            entregar(os.path.join(rd, x), car, final)
            rep_res[t] = "nuevo (no había en el servidor)"
            continue
        cmp = reportes.comparar(os.path.join(*viejos[0]), os.path.join(rd, x))
        if cmp["igual"]:
            rep_res[t] = "sin cambios → no se sube"
            continue
        rep_res[t] = ("no se pudo comparar (" + cmp["error"] + ") → se reemplaza") if cmp["error"] else "cambió → se reemplaza"
        difs[t] = cmp
        entregar(os.path.join(rd, x), car, final)
        for vcar, vnom in viejos:          # mismo tipo con otro nombre (cambió la categoría o la extensión)
            if vnom != final:
                plan.append(("RETIRA", None, vcar, vnom))

    # ---- doble chequeo: lo que dice el reporte vs lo que se entregó
    sub_ent = {base_pieza(x) for x in f.sub_pdf + f.cnc if x not in comprados}
    tip_ent = {base_pieza(x) for x in f.tipicos}
    errores_dc, solo_plano, elim_rep, avisos_dc = [], [], set(), []
    for t, entregadas_t, nombre in ((SUB_REP, sub_ent, "submaterial"), (PESO_REP, tip_ent, "típico")):
        if t not in entregados_t:
            if entregadas_t:
                avisos_dc.append(f"No hay {t}: no se puede hacer el doble chequeo de los {nombre}s.")
            continue
        cmp = difs.get(t)
        if cmp is None:          # igual o nuevo
            if rep_res.get(t, "").startswith("sin cambios"):
                solo_plano += sorted(entregadas_t)
            continue
        if cmp["error"] or "(tabla)" in cmp["cambiadas"]:
            avisos_dc.append(f"{t}: no se pudo comparar pieza por pieza; doble chequeo incompleto.")
            continue
        es_comprado = re.compile(r"^\d+_(%s)" % "|".join(PREFIJOS_COMPRADOS))
        deben = {k for k in set(cmp["cambiadas"]) | set(cmp["nuevas"]) if not es_comprado.match(k)}
        cmp["eliminadas"] = [k for k in cmp["eliminadas"] if not es_comprado.match(k)]
        for k in sorted(deben - entregadas_t):
            motivo = "nuevo en el reporte" if k in cmp["nuevas"] else " · ".join(cmp["cambiadas"][k])
            errores_dc.append(f"{nombre} {k}: {motivo} → falta reexportarlo")
        solo_plano += sorted(entregadas_t - deben)
        elim_rep |= set(cmp["eliminadas"])
    if errores_dc:
        print("BLOQUEADO (doble chequeo con los reportes): el reporte muestra cambios en piezas que NO se entregaron:")
        for x in errores_dc:
            print("   ", x)
        return 1
    confirmados = {c.strip().upper() for c in a.confirmo.split(",") if c.strip()}
    if solo_plano and "RE-09" not in confirmados:
        print("FALTA CONFIRMAR (RE-09): estas piezas se entregaron pero el reporte no muestra cambios. "
              "¿Es un cambio solo de plano (cota, nota, soldadura)? Si sí, pasar --confirmo RE-09:")
        for x in solo_plano:
            print("   ", x)
        return 3
    str_nombre = None
    if f.strumis:
        car = srv.strumis()
        ents = [int(m.group(1)) for m in (re.search(r"ENT ?(\d+)\.bswx$", x, re.I) for x in archivos(car)) if m]
        str_nombre = srv.nombre_strumis(max(ents, default=0) + 1)
        plan.append(("NUEVO", os.path.join(rd, f.strumis[0]), car, str_nombre))

    # ---- 2. piezas eliminadas (LISTA CNC COMPLETA)
    entregadas = {base_pieza(x) for x in f.sub_pdf + f.cnc + f.tipicos}
    srv_sub = {}
    for car in (srv.pdf(), srv.cnc()):
        for x in archivos(car):
            srv_sub.setdefault(base_pieza(x), []).append((car, x))
    srv_tip = {}
    for x in archivos(srv.tip()):
        srv_tip.setdefault(base_pieza(x), []).append((srv.tip(), x))
    retirar = {s.strip() for s in a.retirar.split(",") if s.strip()}
    mantener = {s.strip() for s in a.mantener.split(",") if s.strip()}
    lista_dir = os.path.join(rd, LISTA)
    candidatas, pendientes, avisos = [], [], []
    if os.path.isdir(lista_dir):
        lista = {base_pieza(x) for x in archivos(lista_dir) if x.lower().endswith(".nc1")}
        if not lista:
            print(f"BLOQUEADO: '{LISTA}' está vacía. Exportar el CNC de toda la fase ahí.")
            return 1
        no_en_lista = sorted(b for b in entregadas if b in {base_pieza(x) for x in f.cnc} and b not in lista)
        if no_en_lista:
            print(f"BLOQUEADO: hay CNC entregados que no están en '{LISTA}' (la lista no es de esta versión):")
            for b in no_en_lista:
                print("   ", b)
            return 1
        # las rejas no tienen CNC: que no estén en la lista CNC no quiere decir que se eliminaron
        candidatas = sorted(b for b in srv_sub if b not in lista and b not in entregadas
                            and not re.match(r"^\d+_(%s)" % "|".join(PREFIJOS_COMPRADOS), b)
                            and not re.match(r"^\d+_(%s)[A-Z]" % "|".join(PREFIJOS_SIN_CNC), b))
        pendientes = [b for b in candidatas if b not in retirar and b not in mantener]
    elif not a.sin_lista:
        print(f"FALTA DECIDIR: no hay '{LISTA}' en la reentrega. Exportar el CNC de toda la fase ahí, "
              f"o pasar --sin-lista con las piezas a retirar (--retirar) que diga el modelador.")
        return 3
    # piezas que el reporte nuevo ya no tiene (submateriales y típicos) también son candidatas
    elim_srv = {b for b in elim_rep if (b in srv_sub or b in srv_tip) and b not in entregadas}
    origen_cand = {b: "no está en la lista CNC" for b in candidatas}
    for b in elim_srv:
        origen_cand[b] = (origen_cand[b] + " y " if b in origen_cand else "") + "ya no está en el reporte"
    candidatas = sorted(origen_cand)
    pendientes = [b for b in candidatas if b not in retirar and b not in mantener]
    if pendientes:
        print("FALTA DECIDIR con el modelador: estas piezas están en el servidor pero parecen eliminadas "
              "(¿se eliminaron o cambiaron de nombre?). Pasar cada una en --retirar o en --mantener:")
        for b in pendientes:
            print(f"    {b}   ({origen_cand[b]}; {', '.join(x for _, x in srv_sub.get(b, []) + srv_tip.get(b, []))})")
        return 3
    for b in sorted(retirar):
        encontrados = srv_sub.get(b, []) + srv_tip.get(b, [])
        if not encontrados:
            avisos.append(f"`{b}` (en --retirar) no está en el servidor: no se hace nada.")
        for car, x in encontrados:
            plan.append(("RETIRA", None, car, x))
    avisos += avisos_dc
    if PESO_REP not in difs and not rep_res.get(PESO_REP, "").startswith("sin cambios"):
        avisos.append("Sin comparación del reporte de peso: los típicos eliminados no se detectan; indicarlos con --retirar.")
    if solo_plano:
        avisos.append(f"Cambio solo de plano confirmado por el modelador (RE-09): {', '.join(solo_plano)}")

    # ---- 3. ejecutar
    movidos, copiados = [], []
    if a.ejecutar:
        if not a.avise_checker:
            print("FALTA CONFIRMAR: el modelador tiene que confirmar en el chat que le avisó al checker (--avise-checker).")
            return 6
        # Orden seguro (2026-11-11): 1) lo nuevo se copia con nombre temporal, 2) lo viejo va a OBSOLETOS,
        # 3) lo nuevo toma su nombre. Si algo falla, el servidor nunca queda sin la pieza.
        for accion, origen, car, nombre in plan:
            destino = os.path.join(car, nombre)
            try:
                tmp = None
                if origen:
                    os.makedirs(car, exist_ok=True)
                    tmp = destino + ".subiendo"
                    shutil.copy2(origen, tmp)
                    if os.path.getsize(tmp) != os.path.getsize(origen):
                        raise IOError("tamaño distinto después de copiar")
                if accion in ("REEMPLAZA", "RETIRA") and os.path.exists(destino):
                    obs = os.path.join(car, "OBSOLETOS")
                    os.makedirs(obs, exist_ok=True)
                    nuevo_obs = nombre_obs(obs, nombre, hoy)
                    shutil.move(destino, os.path.join(obs, nuevo_obs))      # mover, nunca borrar
                    movidos.append((os.path.relpath(car, a.constructivos), nombre, nuevo_obs))
                if tmp:
                    os.rename(tmp, destino)
                    copiados.append(destino)
            except Exception as e:
                print(f"FRENA: la reentrega se cortó en {os.path.relpath(destino, a.constructivos)} ({type(e).__name__}: {e}).")
                print(f"       Hechos: {len(copiados)} copiados, {len(movidos)} a OBSOLETOS. La planilla NO se anotó.")
                for c_, x_, o_ in movidos:
                    print(f"         a OBSOLETOS: {c_}\\{x_} → {o_}")
                print("       No repetir sin revisar: avisale al responsable (MB) con esta lista.")
                return 1
        # Marca de reentrega subida: evita subir dos veces la misma REENTREGA NN (2026-11-11)
        with open(os.path.join(rd, MARCA_SUBIDA), "w", encoding="utf-8") as fo:
            json.dump({"fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "usuario": getpass.getuser(),
                       "copiados": len(copiados), "obsoletos": len(movidos)}, fo, ensure_ascii=False)

    marca = None
    reemplazos = sum(1 for p in plan if p[0] in ("REEMPLAZA", "RETIRA"))
    if a.ejecutar and a.pestana:
        texto = (f"REENTREGA {nre:02d}: {len(copiados)} archivo(s) subidos, {reemplazos} a OBSOLETOS")
        try:
            marca = planilla.nota(a.pestana, n, texto)
        except SystemExit as e:
            marca = {"ok": False, "error": str(e)}
        except Exception as e:
            marca = {"ok": False, "error": f"{type(e).__name__}: {e}"}

    # ---- 4. informe
    rel = lambda p: os.path.relpath(p, a.constructivos)
    out = [f"# Reentrega {nre:02d} — {obra} FASE {NN} — {'EJECUTADA' if a.ejecutar else 'SIMULACIÓN (no se tocó nada)'}", "",
           f"- Fecha: {datetime.datetime.now():%Y-%m-%d %H:%M}", f"- Origen: `{rd}`", f"- Destino: `{a.constructivos}`",
           f"- Aviso al checker: {'confirmado por el modelador' if a.avise_checker else 'NO confirmado todavía (necesario para ejecutar)'}"]
    out += ["", "## Plan", "| Acción | Carpeta | Archivo |", "|---|---|---|"]
    etiqueta = {"REEMPLAZA": "Reemplaza (viejo → OBSOLETOS)", "NUEVO": "Nuevo", "RETIRA": "Retira → OBSOLETOS (sin reemplazo)"}
    for accion, origen, car, nombre in sorted(plan, key=lambda p: (p[0], rel(p[2]), p[3])):
        extra = f" (entregado como `{os.path.basename(origen)}`)" if origen and os.path.basename(origen) != nombre else ""
        out.append(f"| {etiqueta[accion]} | `{rel(car)}` | `{nombre}`{extra} |")
    tot = {k: sum(1 for p in plan if p[0] == k) for k in etiqueta}
    out += ["", f"**{tot['REEMPLAZA']} reemplazos · {tot['NUEVO']} nuevos · {tot['RETIRA']} retirados**"]
    out += ["", "## Reportes (comparados con el servidor)", "| Reporte | Resultado |", "|---|---|"]
    out += [f"| {NOMBRE_TIPO.get(t, t)} | {r} |" for t, r in sorted(rep_res.items())]
    for t, cmp in sorted(difs.items()):
        if cmp["error"]:
            continue
        lineas = [f"- `{k}`: " + " · ".join(v) for k, v in cmp["cambiadas"].items()]
        lineas += [f"- `{k}`: nuevo" for k in cmp["nuevas"]] + [f"- `{k}`: ya no está" for k in cmp["eliminadas"]]
        out += ["", f"### Cambios según {NOMBRE_TIPO.get(t, t)} ({len(lineas)})"] + lineas[:40]
        if len(lineas) > 40:
            out.append(f"- … y {len(lineas) - 40} más")
    if candidatas:
        out += ["", "## Piezas que parecen eliminadas"]
        out += [f"- `{b}` ({origen_cand[b]}) → {'retirar' if b in retirar else 'mantener'} (decidido por el modelador)"
                for b in candidatas]
    if movidos:
        out += ["", "## Movidos a OBSOLETOS"] + [f"- `{c}\\{x}` → `OBSOLETOS\\{o}`" for c, x, o in movidos]
    if comprados:
        out += ["", "## No se suben (tuercas y arandelas)"] + [f"- `{x}`" for x in comprados]
    out += ["", "## Avisos"] + [f"- {x}" for x in avisos]
    out += ["", "## Planilla"]
    if not a.pestana:
        out.append("- Se usó un CSV: el estado no cambia; si se quiere, anotar la reentrega a mano en la celda.")
    elif not a.ejecutar:
        out.append(f"- Leída en vivo: fase {n} en `{ESTADO_SUBIDO}`. Al ejecutar, el estado **no cambia**; se agrega una nota con fecha.")
    elif marca and marca.get("ok"):
        out.append(f"- Nota agregada en la fila {marca['fila']} ({marca['fecha']}). Estado sin cambios.")
    else:
        out.append(f"- **No se pudo agregar la nota** ({(marca or {}).get('error', 'sin respuesta')}). Los archivos ya están subidos.")
    md = "\n".join(out) + "\n"
    print(md)
    if a.informe:
        os.makedirs(os.path.dirname(os.path.abspath(a.informe)), exist_ok=True)   # (2026-11-19)
        with open(a.informe, "w", encoding="utf-8") as fo:
            fo.write(md)
    if marca is not None and not marca.get("ok"):
        return 5
    return 0


if __name__ == "__main__":
    sys.exit(main())
