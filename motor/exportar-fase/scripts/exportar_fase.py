# -*- coding: utf-8 -*-
"""
Exportar fase (Skill "exportar-fase", Ferma) - arma el destino, controla y cierra lo que exportan las macros de Tekla.

El destino sale de la OBRA (código del modelo abierto, ej. 696A) y del NÚMERO DE FASE:
    <base>\\<obra>\\FASE <n>\\  con  CNC\\  SUB\\  TIP\\  + reportes      (formato confirmado 2026-11-06)
<base> = OBRAS\\PASADO A CONTROL del Drive compartido: se busca solo en esta PC (o --base / ~\\.ferma\\exportar.json).

Comandos, en el orden de uso:
  preparar        Lee el modelo abierto (log RECIENTE de MCP_ProbeModelo.cs), confirma la fila en la planilla,
                  muestra las rutas. Con --crear: crea carpetas, escribe export_params.txt (con fecha) y el bloqueo.
                  Siempre borra antes los parámetros viejos: si frena, las macros no pueden exportar la fase anterior.
  reportes        Lista las plantillas a sacar (plantillas.json) y la carpeta temporal donde dejarlas.
  traer-reportes  Pasa a la fase SOLO los reportes con datos (los vacíos no van). Avisa si falta alguno.
  revisar         Con el JSON de get_drawings_properties: cada plano de la fase tiene revisión y está actualizado.
  finalizar       Lee los logs de CNC y PDF (frena si hay planos sin actualizar o errores), arma los Excel,
                  cruza SUB/CNC/TIP contra los reportes, avisa del Strumis y levanta el bloqueo.
  planilla        Carga PESO y M2 solo si están vacíos (si difieren, avisa; --reemplazar para pisar) y, con
                  --pasado-a-control, controla que el Drive tenga todo y marca el estado + manifiesto.
  macros          Controla que las macros instaladas en Tekla sean las de la Skill.

Códigos de salida: 0 OK · 1 frena (ver mensaje) · 2 error de uso.
"""
import argparse
import getpass
import glob
import hashlib
import json
import os
import re
import shutil
import sys
import time
from datetime import datetime

DIR_LOG = os.path.join(os.path.expanduser("~"), "TeklaMCP_logs")
PROBE_LOG = os.path.join(DIR_LOG, "probe_modelo_log.txt")
PARAMS = os.path.join(DIR_LOG, "export_params.txt")
DIR_REPORTES = os.path.join(DIR_LOG, "reportes")
CONFIG_PC = os.path.join(os.path.expanduser("~"), ".ferma", "exportar.json")
SKILL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLANTILLAS = os.path.join(SKILL, "plantillas.json")
BASE_CONOCIDA = r"G:\.shortcut-targets-by-id\1p8yAck6fXKpSit08dYenjyFLlO_z5VE9\OBRAS\PASADO A CONTROL"
CONFIG_IMPRESION_BASE = "MFD"          # confirmado por el usuario 2026-11-06
SUBCARPETAS = ("CNC", "SUB", "TIP")
COMPRADOS = re.compile(r"^\d+_(nut|washer|tuerca|arandela)", re.I)   # se compran fabricadas: nunca van
SIN_CNC = re.compile(r"^\d+_gr[A-Z]")   # rejas / escalones de reja: se compran hechos, PDF sí, CNC no (2026-11-10)
BLOQUEO = "_EXPORTANDO.txt"
PROBE_MAX_MIN = 5          # el log del modelo tiene que ser de recién
BLOQUEO_MAX_H = 4          # un bloqueo de otra persona más viejo que esto se considera abandonado
LARGO_MAX = 240            # Windows se complica arriba de 260 caracteres
MACROS = ("MCP_ProbeModelo.cs", "MCP_SeleccionarFase.cs", "MCP_ExportCNCFase.cs", "MCP_ImprimirFasePDF.cs")


# ---------------------------------------------------------------- utilidades
def leer_kv(ruta):
    d = {}
    if os.path.isfile(ruta):
        with open(ruta, encoding="utf-8-sig", errors="replace") as f:
            for linea in f:
                if "=" in linea:
                    k, v = linea.rstrip("\n").split("=", 1)
                    d[k.strip()] = v.strip()
    return d


def minutos_desde(ruta):
    return (time.time() - os.path.getmtime(ruta)) / 60


def buscar_base():
    """OBRAS\\PASADO A CONTROL en esta PC: config propia, después la ruta conocida, después se busca en las unidades."""
    if os.path.isfile(CONFIG_PC):
        with open(CONFIG_PC, encoding="utf-8") as f:
            b = json.load(f).get("base", "")
        if b and os.path.isdir(b):
            return b
    if os.path.isdir(BASE_CONOCIDA):
        return BASE_CONOCIDA
    patrones = [r"{u}:\.shortcut-targets-by-id\*\OBRAS\PASADO A CONTROL", r"{u}:\Unidades compartidas\*\OBRAS\PASADO A CONTROL",
                r"{u}:\Shared drives\*\OBRAS\PASADO A CONTROL", r"{u}:\Mi unidad\OBRAS\PASADO A CONTROL",
                r"{u}:\My Drive\OBRAS\PASADO A CONTROL", r"{u}:\Mi unidad\*\OBRAS\PASADO A CONTROL"]
    hallados = []
    for u in "GHIJKLMNOPQRSTUVWXYZDEF":
        for p in patrones:
            hallados += [h for h in glob.glob(p.format(u=u)) if os.path.isdir(h)]
    if len(hallados) == 1:
        return hallados[0]
    if hallados:
        sys.exit("FRENA: encontré varias carpetas PASADO A CONTROL; elegí una con --base o en " + CONFIG_PC + ":\n  " + "\n  ".join(hallados))
    sys.exit("FRENA: no encontré OBRAS\\PASADO A CONTROL en esta PC. Pasala con --base o guardala en " + CONFIG_PC + ' como {"base": "..."}')


def leer_probe(max_min=120):
    """preparar exige un log de recién (PROBE_MAX_MIN); los pasos siguientes aceptan hasta 2 h, como los parámetros."""
    if not os.path.isfile(PROBE_LOG):
        sys.exit("FRENA: no está el log del modelo. Correr antes MCP_ProbeModelo.cs.")
    edad = minutos_desde(PROBE_LOG)
    if edad > max_min:
        sys.exit(f"FRENA: el log del modelo tiene {edad:.0f} min. Correr MCP_ProbeModelo.cs de nuevo (si falla, reconectar tekla-mcp).")
    d = leer_kv(PROBE_LOG)
    if any(k.startswith("ERROR") or k.startswith("EXCEPCION") for k in d) or "MODEL_NAME" not in d or "MODEL_PATH" not in d:
        sys.exit("FRENA: el log del modelo no tiene MODEL_NAME / MODEL_PATH (¿Tekla sin modelo abierto?).")
    d["MODELO"] = re.sub(r"\.db1$", "", d["MODEL_NAME"], flags=re.I)
    m = re.match(r"^\s*(\d+[A-Za-z]+)\b", d["MODELO"])
    if not m:
        sys.exit("FRENA: no se pudo sacar el código de obra del nombre del modelo: " + d["MODELO"])
    d["OBRA"] = m.group(1).upper()
    return d


def carpeta_fase(base_obra, n):
    """Usa una carpeta existente de esa fase si la hay; si no, propone 'FASE <n>' (2 dígitos mínimo)."""
    pat = re.compile(r"^FASE\s*0*%d(?:\b|_|\s|$)" % n, re.I)
    if os.path.isdir(base_obra):
        for d in sorted(os.listdir(base_obra)):
            if os.path.isdir(os.path.join(base_obra, d)) and pat.match(d):
                return os.path.join(base_obra, d), True
    return os.path.join(base_obra, "FASE %02d" % n), False


def rutas(args, max_min=120):
    pr = leer_probe(max_min)
    if args.obra and args.obra.upper() != pr["OBRA"]:
        sys.exit(f"FRENA: pediste la obra {args.obra.upper()} pero el modelo abierto es {pr['OBRA']} ({pr['MODELO']}).")
    base = args.base or buscar_base()
    fdir, existe = carpeta_fase(os.path.join(base, pr["OBRA"]), args.fase)
    return pr, fdir, existe


def modulo_planilla(libro=None):
    sys.path.insert(0, os.path.join(SKILL, "..", "subir-fase", "scripts"))
    import planilla  # noqa: E402  (Skill subir-fase)
    if libro:
        planilla.CONFIG = planilla.ruta_config(libro)
    return planilla


def num(v):
    try:
        return float(str(v).replace(",", ".").strip())
    except ValueError:
        return None


def marcas_reporte(fdir, clave):
    """Marcas de la fase según los reportes: clave 'CONJUNTOS' (típicos) o 'SUBMATERIALES' (piezas)."""
    marcas = set()
    for x in os.listdir(fdir):
        if clave in x.upper() and x.lower().endswith(".csv"):
            with open(os.path.join(fdir, x), encoding="cp1252", errors="replace") as f:
                for linea in f:
                    c = [y.strip() for y in linea.split(";")]
                    if clave == "CONJUNTOS" and len(c) > 2 and re.fullmatch(r"\d+", c[0]) and re.fullmatch(r"[A-Za-z]+\d+[\w-]*", c[1]):
                        marcas.add(c[1])
                    if clave == "SUBMATERIALES" and len(c) > 2 and re.fullmatch(r"\d+_[A-Za-z]+\d+[\w-]*", c[0]):
                        marcas.add(c[0])
    return marcas


def reporte_vacio(ruta):
    """True si el reporte de Tekla no tiene ninguna fila de datos (solo encabezados).
    Fila de datos = línea con ';' cuyo primer campo tiene algún dígito (cantidad o marca)."""
    with open(ruta, encoding="cp1252", errors="replace") as f:
        for linea in f:
            if ";" in linea and re.search(r"\d", linea.split(";", 1)[0]):
                return False
    return True


def lista_plantillas(obra):
    with open(PLANTILLAS, encoding="utf-8") as f:
        p = json.load(f)
    return p.get("obras", {}).get(obra, p["defecto"])


def staging(obra, fase):
    return os.path.join(DIR_REPORTES, f"{obra}_FASE_{fase}")


def leer_bloqueo(fdir):
    ruta = os.path.join(fdir, BLOQUEO)
    return (leer_kv(ruta), minutos_desde(ruta) / 60) if os.path.isfile(ruta) else (None, 0)


# ---------------------------------------------------------------- comandos
def config_impresion(attr_dir, nombre, destino):
    """(Viejo, solo --solo-pdf) Copia MFD.PdfPrintOptions.xml cambiando solo la carpeta de salida."""
    base = os.path.join(attr_dir, CONFIG_IMPRESION_BASE + ".PdfPrintOptions.xml")
    if not os.path.isfile(base):
        sys.exit(f"FRENA: no está la config de impresión {CONFIG_IMPRESION_BASE} en {attr_dir}")
    with open(base, encoding="utf-8-sig") as f:
        xml = f.read()
    nuevo = "<PDFAndPlotFileLocation>%s</PDFAndPlotFileLocation>" % destino.replace("&", "&amp;")
    xml = re.sub(r"<PDFAndPlotFileLocation>.*?</PDFAndPlotFileLocation>", lambda m: nuevo, xml, flags=re.S)
    with open(os.path.join(attr_dir, nombre + ".PdfPrintOptions.xml"), "w", encoding="utf-8") as f:
        f.write(xml)


ESTADOS_CERRADOS = ("APROBADO", "SUBIDO AL SERVIDOR")


def frena_por_estado(fila, fase):
    """(2026-11-11) Una fase APROBADA o ya SUBIDA no se vuelve a exportar: "Retomar" sobre ella devolvía el estado a
    PASADO A CONTROLAR y reescribía el manifiesto (la foto de lo que vio el checker). Los cambios van por reentrega."""
    est = re.sub(r"\s+", " ", str(fila.get("ESTADO", ""))).strip().upper()
    if est in ESTADOS_CERRADOS:
        print(f"FRENA: la fase {fase} está en '{est}' en la planilla: no se vuelve a exportar ni se cambia el estado.")
        print("       Si hay cambios después del OK del checker, van como reentrega (carpeta REENTREGA NN).")
        return True
    return False


def preparar(args):
    # (2) Nunca quedan parámetros de la vez anterior: si este comando frena, las macros no tienen qué exportar.
    if os.path.isfile(PARAMS):
        os.remove(PARAMS)
    pr, fdir, existe = rutas(args, PROBE_MAX_MIN)
    print(f"Modelo abierto : {pr['MODELO']}")
    print(f"Obra           : {pr['OBRA']}")
    print(f"Fase           : {args.fase}")

    # (11) La obra y la fase tienen que existir en la planilla: confirma que el código de obra se leyó bien.
    if not args.sin_planilla:
        try:
            fila = modulo_planilla(args.libro).fila_secuencia(args.pestana or pr["OBRA"], args.fase)
        except SystemExit as e:
            print(f"FRENA: no encontré la fase {args.fase} en la pestaña {args.pestana or pr['OBRA']} de la planilla ({e}).")
            print("       Si la pestaña tiene otro nombre, usá --pestana.")
            return 1
        print(f"Planilla       : {fila.get('CATEGORIA', '?')} · detalló {fila.get('DETALLO', '?')} · estado {fila.get('ESTADO', '?')}")
        if frena_por_estado(fila, args.fase):
            return 1

    print(f"Carpeta de fase: {fdir}  ({'ya existe' if existe else 'SE CREA'})")
    ocupadas = []
    for s in SUBCARPETAS:
        p = os.path.join(fdir, s)
        n = len(os.listdir(p)) if os.path.isdir(p) else 0
        print(f"   {s:4}: {p}  ({'SE CREA' if not os.path.isdir(p) else str(n) + ' archivo(s)'})")
        if n:
            ocupadas.append(s)

    # (14) Rutas largas
    largo = len(os.path.join(fdir, "SUB", "")) + 45
    if largo > LARGO_MAX:
        print(f"AVISO: las rutas van a tener unos {largo} caracteres; arriba de 260 Windows falla. Conviene acortar el nombre de la carpeta.")

    # (12) Otra persona exportando la misma fase
    bl, horas = leer_bloqueo(fdir)
    if bl and bl.get("USUARIO") != getpass.getuser() and horas < BLOQUEO_MAX_H:
        print(f"FRENA: {bl.get('USUARIO')} está exportando esta fase desde {bl.get('DESDE')} (en {bl.get('PC')}). Hablalo antes de seguir.")
        return 1

    # (15) Fase a medias: con --retomar se sigue sin pisar nada (las macros saltean lo que ya está).
    if ocupadas and not args.retomar:
        print("FRENA: ya hay archivos en " + ", ".join(ocupadas) + ". No se exporta encima (reentrega = otro flujo).")
        print("       Si es la misma exportación que se cortó a la mitad, repetir con --retomar.")
        return 1
    if not args.crear:
        print("\nSimulación: no se creó nada. Repetir con --crear después del OK del usuario.")
        return 0
    for s in SUBCARPETAS:
        os.makedirs(os.path.join(fdir, s), exist_ok=True)
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(os.path.join(fdir, BLOQUEO), "w", encoding="utf-8") as f:
        f.write(f"USUARIO={getpass.getuser()}\nPC={os.environ.get('COMPUTERNAME', '?')}\nDESDE={ahora}\nMODELO={pr['MODELO']}\n")
    os.makedirs(DIR_LOG, exist_ok=True)
    with open(PARAMS, "w", encoding="utf-8") as f:
        f.write(f"CREADO={ahora}\nMODELO={pr['MODELO']}\nFASE={args.fase}\nDESTINO={os.path.join(fdir, 'CNC')}\nCONFIG_CNC=Todo\n"
                f"DESTINO_SUB={os.path.join(fdir, 'SUB')}\nDESTINO_TIP={os.path.join(fdir, 'TIP')}\n")
    st = staging(pr["OBRA"], args.fase)
    if os.path.isdir(st):
        shutil.rmtree(st)          # carpeta temporal propia de la Skill, en TeklaMCP_logs
    os.makedirs(st)
    print("\nListo: carpetas creadas, bloqueo puesto y export_params.txt escrito (vale 2 horas).")
    print("Siguiente: MCP_SeleccionarFase.cs -> reportes -> traer-reportes -> revisar -> MCP_ExportCNCFase.cs -> MCP_ImprimirFasePDF.cs -> finalizar.")
    return 0


def cmd_reportes(args):
    pr = leer_probe()
    st = staging(pr["OBRA"], args.fase)
    os.makedirs(st, exist_ok=True)
    print(json.dumps({"carpeta": st, "plantillas": lista_plantillas(pr["OBRA"])}, ensure_ascii=False, indent=1))
    return 0


def traer_reportes(args):
    pr, fdir, _ = rutas(args)
    st = staging(pr["OBRA"], args.fase)
    esperadas = lista_plantillas(pr["OBRA"])
    faltan, vacios, traidos = [], [], []
    for t in esperadas:
        base = os.path.splitext(t)[0]
        cand = [os.path.join(st, base + e) for e in (".csv", ".xsr")]
        src = next((c for c in cand if os.path.isfile(c)), None)
        if not src:
            faltan.append(t)
            continue
        if reporte_vacio(src):
            vacios.append(t)       # regla del usuario (2026-11-10): un reporte vacío no va
            continue
        dst = os.path.join(fdir, base + ".csv")
        if os.path.exists(dst):
            print(f"AVISO: ya existe {os.path.basename(dst)} en la fase; no se pisa.")
            continue
        shutil.copy2(src, dst)
        traidos.append(os.path.basename(dst))
    print("Reportes a la fase: " + (", ".join(traidos) or "ninguno"))
    if vacios:
        print("Vacíos, no van: " + ", ".join(vacios))
    if faltan:
        print("FRENA: no se generaron: " + ", ".join(faltan) + ". Si la plantilla tiene otro nombre en esta obra, preguntar cuál es y anotarla en plantillas.json.")
        return 1
    en_fase = [x for x in os.listdir(fdir) if x.lower().endswith(".csv")]
    if not any("CONJUNTOS" in x.upper() for x in en_fase) or not any("SUBMATERIALES" in x.upper() for x in en_fase):
        print("FRENA: falta el reporte de conjuntos o el de submateriales con datos: sin ellos no se puede controlar la fase.")
        return 1
    return 0


def revisar(args):
    """(5) Revisión y estado de cada plano, con el JSON que devuelve get_drawings_properties (solo lectura)."""
    pr, fdir, _ = rutas(args)
    with open(args.json, encoding="utf-8-sig") as f:
        datos = json.load(f)
    planos = {d["mark"].strip("[] "): d for d in datos.get("drawings", [])}
    esperados = sorted(marcas_reporte(fdir, "SUBMATERIALES")) + sorted(marcas_reporte(fdir, "CONJUNTOS"))
    problemas = []
    for m in esperados:
        if COMPRADOS.match(m):
            continue
        d = planos.get(m)
        if not d:
            problemas.append(f"{m}: no tiene plano en Tekla")
            continue
        if not str(d.get("revision_mark", "")).strip():
            problemas.append(f"{m}: sin revisión")
        if d.get("up_to_date_status") != "DrawingIsUpToDate":
            problemas.append(f"{m}: plano no actualizado ({d.get('up_to_date_status')})")
    revs = sorted({f"rev {d.get('revision_mark')} · {d.get('user_properties', {}).get('text4_1', '?')}" for m, d in planos.items() if m in esperados})
    print(f"Planos controlados: {len(esperados)} ({', '.join(revs) or '-'})")
    if problemas:
        print("FRENA:")
        for p in problemas:
            print("  - " + p)
        return 1
    print("OK: todos con revisión y actualizados.")
    return 0


def finalizar(args):
    pr, fdir, _ = rutas(args)
    if not os.path.isdir(fdir):
        print("FRENA: no existe " + fdir)
        return 1
    errores = []

    # (4) Lo que dicen las macros: planos sin actualizar, errores de PDF, CNC que no salieron.
    pdf, cnc = leer_kv(os.path.join(DIR_LOG, "imprimir_pdf_log.txt")), os.path.join(DIR_LOG, "export_cnc_log.txt")
    if pdf.get("FASE") != str(args.fase):
        errores.append("el log de PDF no es de esta fase: correr MCP_ImprimirFasePDF.cs")
    else:
        if int(pdf.get("NO_ACTUALIZADOS", "0") or 0):
            errores.append(f"{pdf['NO_ACTUALIZADOS']} plano(s) no actualizados al imprimir el PDF: actualizarlos, borrar esos PDF y repetir")
        if int(pdf.get("ERRORES", "0") or 0):
            errores.append(f"{pdf['ERRORES']} PDF con error (ver imprimir_pdf_log.txt)")
    if os.path.isfile(cnc):
        with open(cnc, encoding="utf-8-sig", errors="replace") as f:
            sin = [l.split()[-1] for l in f if l.startswith("  SIN") and not SIN_CNC.match(l.split()[-1])]
        if sin:
            errores.append("CNC que no se generaron: " + ", ".join(sin))

    for x in os.listdir(fdir):
        if x.lower().endswith(".xsr"):
            nuevo = os.path.join(fdir, x[:-4] + ".csv")
            if not os.path.exists(nuevo):
                os.rename(os.path.join(fdir, x), nuevo)
                print(f"renombrado: {x} -> {os.path.basename(nuevo)}")

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # con python -P no está en el path
    import reportes_excel  # noqa: E402
    for x in sorted(os.listdir(fdir)):
        if x.lower().endswith(".csv"):
            if reporte_vacio(os.path.join(fdir, x)):
                print(f"vacío: {x} no tiene datos; no va (no se arma su Excel).")
                continue
            _, msg = reportes_excel.generar(os.path.join(fdir, x))
            print("excel: " + msg)

    def lista(s, ext):
        p = os.path.join(fdir, s)
        return sorted(os.path.splitext(x)[0] for x in os.listdir(p) if x.lower().endswith(ext)) if os.path.isdir(p) else []

    sub, cnc_l, tip = lista("SUB", ".pdf"), lista("CNC", ".nc1"), lista("TIP", ".pdf")
    sin_cnc = [b for b in sub if SIN_CNC.match(b) and b not in cnc_l]
    if sin_cnc:
        print("  Rejas / escalones de reja (se compran hechos: PDF sí, CNC no): " + ", ".join(sin_cnc))
    for b in sorted(set(sub) ^ set(cnc_l)):
        if not COMPRADOS.match(b) and b not in sin_cnc:
            errores.append(f"{b}: {'falta el CNC' if b in sub else 'falta el PDF'}")
    otra = [x for x in sub + cnc_l if not re.match(r"^0*%d_" % args.fase, x)]
    if otra:
        errores.append("archivos de otra fase: " + ", ".join(otra))
    # (3) piezas del reporte de submateriales contra lo exportado
    pzs = {m for m in marcas_reporte(fdir, "SUBMATERIALES") if not COMPRADOS.match(m)}
    if pzs:
        for m in sorted(pzs - set(sub)):
            errores.append(f"{m}: está en el reporte de submateriales pero no tiene PDF")
        for m in sorted(set(sub) - pzs):
            errores.append(f"{m}: tiene PDF pero no está en el reporte de submateriales")
    esperados = marcas_reporte(fdir, "CONJUNTOS")
    for t in sorted(esperados - set(tip)):
        errores.append(f"típico {t}: está en el reporte de conjuntos pero falta su PDF")
    for t in sorted(set(tip) - esperados) if esperados else []:
        errores.append(f"típico {t}: tiene PDF pero no está en el reporte de conjuntos")
    comprados = [x for x in cnc_l + sub if COMPRADOS.match(x)]
    # (14) rutas largas y (13) archivos vacíos
    for raiz, _, archivos in os.walk(fdir):
        for a in archivos:
            r = os.path.join(raiz, a)
            if len(r) > 255:
                errores.append(f"ruta de {len(r)} caracteres: {r}")
            if os.path.getsize(r) == 0 and a != BLOQUEO:
                errores.append(f"archivo vacío (0 bytes): {a}")

    print(f"\nFase {args.fase} en {fdir}")
    print(f"  SUB {len(sub)} PDF · CNC {len(cnc_l)} · TIP {len(tip)} PDF (reporte de conjuntos: {len(esperados)}, de submateriales: {len(pzs)})")
    reps = sorted(x for x in os.listdir(fdir) if x.lower().endswith(".xlsx"))
    print("  Reportes: " + (", ".join(reps) or "ninguno"))
    if comprados:
        print("  Tuercas/arandelas (no se suben nunca): " + ", ".join(comprados))
    # (16) Strumis: es manual; solo se recuerda.
    if not glob.glob(os.path.join(fdir, "*.bswx")):
        print("  RECORDATORIO: después del APROBADO del checker, el modelador exporta el Strumis (.bswx) a mano con BIMReview Export y NUNCA renumera.")
    if errores:
        print("FRENA:")
        for e in errores:
            print("  - " + e)
        return 1
    bl = os.path.join(fdir, BLOQUEO)
    if os.path.isfile(bl):
        os.remove(bl)              # el bloqueo es de la Skill: se levanta al terminar bien
    print("OK: todo cruzado. Bloqueo levantado.")
    return 0


def totales_conjuntos(fdir):
    """(peso_kg, m2_pintura) de la línea de total del reporte de conjuntos (FMA_LISTADO_CONJUNTOS_TOTALES_*)."""
    for x in os.listdir(fdir):
        if "CONJUNTOS" in x.upper() and x.lower().endswith(".csv"):
            with open(os.path.join(fdir, x), encoding="cp1252", errors="replace") as f:
                for linea in f:
                    c = [y.strip() for y in linea.split(";")]
                    if len(c) > 3 and c[1].upper() == "CONJUNTOS":
                        nums = [y for y in c[2:] if re.fullmatch(r"\d+(?:[.,]\d+)?", y)]
                        if len(nums) >= 2:
                            return nums[0], nums[-1]
    return None, None


def drive_estable(fdir, espera=4):
    """(13) El Drive terminó de escribir: ningún archivo de 0 bytes y los tamaños no cambian en unos segundos."""
    def foto():
        return {os.path.join(r, a): os.path.getsize(os.path.join(r, a)) for r, _, aa in os.walk(fdir) for a in aa}
    a = foto()
    time.sleep(espera)
    b = foto()
    cambian = [os.path.basename(k) for k in b if a.get(k) != b[k]]
    vacios = [os.path.basename(k) for k, v in b.items() if v == 0 and not k.endswith(BLOQUEO)]
    return cambian, vacios


def cmd_planilla(args):
    planilla = modulo_planilla(args.libro)
    pr, fdir, _ = rutas(args)
    if not os.path.isdir(fdir):
        print("FRENA: no existe " + fdir)
        return 1
    pest = args.pestana or pr["OBRA"]
    fila = planilla.fila_secuencia(pest, args.fase)
    print(f"Planilla ({os.path.basename(planilla.CONFIG)}) · pestaña {pest} · fase {args.fase}")
    for k in ("CATEGORIA", "MODELO", "DETALLO", "ESTADO", "PESO (KG)", "M2 - PINTURA"):
        print(f"  {k:<14} {fila.get(k, '(no hay columna)')}")
    if args.pasado_a_control and frena_por_estado(fila, args.fase):
        return 1
    peso, m2 = totales_conjuntos(fdir)
    print(f"Reporte de conjuntos: peso {peso} kg · pintura {m2} m2")
    if args.completar:
        if not peso:
            print("FRENA: no se encontró el total en el reporte de conjuntos.")
            return 1
        # (10) Solo se carga lo vacío. Si la planilla tiene otro valor, se avisa y se pisa solo con --reemplazar.
        enviar, distintos = {}, []
        for col, clave, val in (("PESO (Kg)", "PESO (KG)", peso), ("M2 - PINTURA", "M2 - PINTURA", m2)):
            actual = fila.get(clave, "")
            if not str(actual).strip():
                enviar[col] = float(val)
            elif num(actual) is not None and abs(num(actual) - float(val)) < 0.01:
                pass
            else:
                distintos.append(f"{col}: planilla {actual} · reporte {val}")
                if args.reemplazar:
                    enviar[col] = float(val)
        if distintos:
            print("AVISO, la planilla tiene otro valor: " + " | ".join(distintos) + ("  → se reemplaza (--reemplazar)" if args.reemplazar else "  → no se toca (usar --reemplazar si corresponde)"))
        if enviar:
            r = planilla.completar(pest, args.fase, enviar)
            print(f"  cargado: {r.get('escritos')}  reemplazado: {r.get('reemplazados')}")
            if r.get("sinColumna"):
                print(f"  AVISO sin columna / no permitido: {r.get('sinColumna')}")
        else:
            print("  nada para cargar.")
    if args.pasado_a_control:
        if os.path.isfile(os.path.join(fdir, BLOQUEO)):
            print("FRENA: la exportación no terminó bien (sigue el bloqueo). Correr finalizar primero.")
            return 1
        cambian, vacios = drive_estable(fdir)
        if cambian or vacios:
            print("FRENA: el Drive todavía está escribiendo (" + ", ".join((cambian + vacios)[:8]) + "). Esperar un minuto y repetir.")
            return 1
        r = planilla.marcar(pest, args.fase, "PASADO A CONTROLAR", "Claude (Skill exportar-fase)")
        print(f"  estado: '{r['anterior']}' → '{r['nuevo']}' ({r['fecha']})")
        # Foto de lo que se manda a controlar: al subir se compara contra esto (bloquea y pregunta, 2026-11-07)
        import manifiesto  # noqa: E402  (Skill subir-fase)
        ruta = manifiesto.escribir(fdir, {"obra": pr["OBRA"], "fase": args.fase,
                                          "categoria": fila.get("CATEGORIA", ""), "detallo": fila.get("DETALLO", "")})
        print(f"  manifiesto: {ruta} ({len(manifiesto.leer(fdir)['archivos'])} archivos)")
    return 0


def cmd_manifiesto(args):
    """Escribe el manifiesto sin tocar la planilla (cuando PASADO A CONTROLAR se marcó a mano)."""
    modulo_planilla()
    import manifiesto  # noqa: E402
    fdir = args.fase_dir or rutas(args)[1]
    if not os.path.isdir(fdir):
        print("FRENA: no existe " + fdir)
        return 1
    if manifiesto.leer(fdir) and not args.reescribir:
        print(f"FRENA: ya hay manifiesto en {fdir}. Solo se reescribe si la fase vuelve a control (--reescribir).")
        return 1
    ruta = manifiesto.escribir(fdir, {"fase": args.fase})
    print(f"manifiesto: {ruta} ({len(manifiesto.leer(fdir)['archivos'])} archivos)")
    return 0


def cmd_macros(args):
    """(7) Las macros instaladas en cada Tekla tienen que ser iguales a las de la Skill."""
    def h(p):
        with open(p, "rb") as f:
            return hashlib.sha256(f.read()).hexdigest()
    # --dirs: las carpetas donde la Tekla abierta busca las macros (las pasa la App, que se lo pregunta a Tekla).
    dirs = [d for d in (getattr(args, "dirs", None) or "").split(";") if d.strip()]
    if not dirs:
        dirs = glob.glob(r"C:\TeklaStructures\*\Environments\common\macros\modeling")
    if not dirs:
        print("FRENA: no encontré la carpeta de macros de Tekla. Instalar con scripts\\instalar_macros.ps1.")
        return 1
    mal = []
    for d in dirs:
        for m in MACROS:
            inst, orig = os.path.join(d, m), os.path.join(SKILL, "macros", m)
            estado = "falta" if not os.path.isfile(inst) else ("ok" if h(inst) == h(orig) else "VIEJA")
            if estado != "ok":
                mal.append(f"{d}\\{m}: {estado}")
    if mal:
        print("AVISO, macros para instalar o actualizar (correr scripts\\instalar_macros.ps1):")
        for x in mal:
            print("  - " + x)
        return 1
    print(f"OK: las {len(MACROS)} macros están al día en " + ", ".join(dirs))
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    for nombre in ("preparar", "reportes", "traer-reportes", "revisar", "finalizar", "planilla", "manifiesto", "macros"):
        s = sp.add_parser(nombre)
        if nombre == "macros":
            s.add_argument("--dirs", help="carpetas de macros de modelado separadas por ';' (por defecto C:\\TeklaStructures\\*)")
            continue
        s.add_argument("--fase", type=int, required=True)
        s.add_argument("--base", help="carpeta madre (por defecto se busca OBRAS\\PASADO A CONTROL en esta PC)")
        s.add_argument("--obra", help="código de obra esperado; si no coincide con el modelo abierto, frena")
        if nombre in ("preparar", "planilla"):
            s.add_argument("--pestana", help="pestaña de la planilla (por defecto, el código de obra)")
            s.add_argument("--libro", choices=("real", "prueba"))
        if nombre == "preparar":
            s.add_argument("--crear", action="store_true")
            s.add_argument("--retomar", action="store_true", help="la misma exportación se cortó: sigue sin pisar nada")
            s.add_argument("--sin-planilla", action="store_true", help="pruebas: no controla la fila en la planilla")
        if nombre == "revisar":
            s.add_argument("--json", required=True, help="respuesta de get_drawings_properties guardada en un archivo")
        if nombre == "planilla":
            s.add_argument("--completar", action="store_true", help="carga PESO y M2 del reporte si están vacíos")
            s.add_argument("--reemplazar", action="store_true", help="pisa PESO/M2 aunque la planilla tenga otro valor")
            s.add_argument("--pasado-a-control", action="store_true",
                           help="marca PASADO A CONTROLAR y escribe el manifiesto (exige finalizar OK)")
        if nombre == "manifiesto":
            s.add_argument("--fase-dir", help="carpeta de la fase (si se pasa, no hace falta el log del modelo)")
            s.add_argument("--reescribir", action="store_true", help="la fase volvió a control: reemplaza el manifiesto")
    a = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    return {"preparar": preparar, "reportes": cmd_reportes, "traer-reportes": traer_reportes, "revisar": revisar,
            "finalizar": finalizar, "planilla": cmd_planilla, "manifiesto": cmd_manifiesto, "macros": cmd_macros}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
