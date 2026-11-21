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
# (2026-11-21) Correas Z (CFZ / cfz): no llevan submateriales ni CNC, solo típicos. Si la fase es toda de correas Z,
# no se pregunta: se sigue sola. Cualquier otra fase sin planos de submateriales pregunta.
CORREA_Z_TIP = re.compile(r"^CFZ", re.I)
CORREA_Z_SUB = re.compile(r"^\d+_cfz", re.I)
SIN_SUB_MARCA = "_sin_submateriales.txt"   # va en la carpeta temporal de la Skill (staging), nunca al Drive
BLOQUEO = "_EXPORTANDO.txt"
PROBE_MAX_MIN = 5          # el log del modelo tiene que ser de recién
BLOQUEO_MAX_H = 4          # un bloqueo de otra persona más viejo que esto se considera abandonado
LARGO_MAX = 240            # Windows se complica arriba de 260 caracteres
MACROS = ("MCP_ProbeModelo.cs", "MCP_SeleccionarFase.cs", "MCP_ExportCNCFase.cs", "MCP_ImprimirFasePDF.cs")
# (2026-11-14) Se arma todo en una carpeta LOCAL (rápido) y al final se copia al Drive de una vez.
DIR_TRABAJO = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Ferma", "exportando")
# (2026-11-14) Reportes elegidos a mano para una obra (la App los recuerda por obra).
REPORTES_OBRA = os.path.join(os.path.expanduser("~"), ".ferma", "reportes_por_obra.json")
SELECCION = os.path.join(DIR_LOG, "seleccionar_fase.txt")
PARAMS_MAX_MIN = 240       # igual que las macros
GALV = re.compile(r"\bGALV|\bGAL\b|\bGAL\.", re.I)   # galvanizado (en caliente o en frío): sin pintura (2026-11-12)


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


def limpiar_nombre(texto):
    """Categoría apta para nombre de carpeta: sin \\ / : * ? " < > | y sin espacios de más."""
    return re.sub(r"\s+", " ", re.sub(r'[\\/:*?"<>|]', "-", texto or "")).strip(" .")


def carpeta_fase(base_obra, n, categoria=""):
    """Usa una carpeta existente de esa fase si la hay (con o sin categoría; no se renombra).
    Si no, propone 'FASE <nn> - <CATEGORIA>' (pedido del usuario 2026-11-11)."""
    pat = re.compile(r"^FASE\s*0*%d(?:\b|_|\s|$)" % n, re.I)
    if os.path.isdir(base_obra):
        for d in sorted(os.listdir(base_obra)):
            if os.path.isdir(os.path.join(base_obra, d)) and pat.match(d):
                return os.path.join(base_obra, d), True
    cat = limpiar_nombre(categoria)
    return os.path.join(base_obra, "FASE %02d" % n + (" - " + cat if cat else "")), False


def carpeta_trabajo(obra, fdir):
    """Carpeta LOCAL donde se arma la fase antes de copiarla al Drive (mismo nombre que en el Drive)."""
    return os.path.join(DIR_TRABAJO, obra, os.path.basename(fdir))


def rutas(args, max_min=PARAMS_MAX_MIN):
    pr = leer_probe(max_min)
    if args.obra and args.obra.upper() != pr["OBRA"]:
        sys.exit(f"FRENA: pediste la obra {args.obra.upper()} pero el modelo abierto es {pr['OBRA']} ({pr['MODELO']}).")
    base = args.base or buscar_base()
    fdir, existe = carpeta_fase(os.path.join(base, pr["OBRA"]), args.fase, getattr(args, "categoria", "") or "")
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


def tipo_csv(ruta):
    """(2026-11-14) Tipo de un reporte de Tekla por su CONTENIDO, no por el nombre (sirve para reportes elegidos a
    mano, ej. "MB_REPORTE DE PESO Y CLASIFICION"): 'CONJUNTOS' (Código + P.TOTAL), 'SUBMATERIALES' (Marca + Miembro)
    o None. Si el encabezado no lo dice, se mira el nombre del archivo."""
    try:
        with open(ruta, encoding="cp1252", errors="replace") as f:
            for i, linea in enumerate(f):
                if i > 40:
                    break
                if ";" not in linea:
                    continue
                cab = re.sub(r"[^A-Z;]", "", linea.upper().replace("Ó", "O"))
                celdas = cab.split(";")
                if "CODIGO" in celdas and "PTOTAL" in celdas:
                    return "CONJUNTOS"
                if "MARCA" in celdas and "MIEMBRO" in celdas:
                    return "SUBMATERIALES"
    except OSError:
        return None
    n = os.path.basename(ruta).upper()
    return "CONJUNTOS" if "CONJUNTOS" in n else ("SUBMATERIALES" if "SUBMATERIALES" in n else None)


def marcas_reporte(fdir, clave):
    """Marcas de la fase según los reportes: clave 'CONJUNTOS' (típicos) o 'SUBMATERIALES' (piezas)."""
    marcas = set()
    for x in os.listdir(fdir):
        if x.lower().endswith(".csv") and tipo_csv(os.path.join(fdir, x)) == clave:
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


def lista_plantillas(obra, por_defecto=False):
    """Plantillas de reportes de una obra: las elegidas a mano para esa obra (~\\.ferma\\reportes_por_obra.json),
    si no las de plantillas.json para la obra, si no las de 'defecto'. por_defecto=True ignora lo elegido a mano."""
    if not por_defecto and os.path.isfile(REPORTES_OBRA):
        try:
            with open(REPORTES_OBRA, encoding="utf-8") as f:
                elegidas = json.load(f).get(obra)
            if elegidas:
                return list(elegidas)
        except (OSError, ValueError):
            pass
    with open(PLANTILLAS, encoding="utf-8") as f:
        p = json.load(f)
    return p.get("obras", {}).get(obra, p["defecto"])


def elegidas_a_mano(obra):
    """True si la obra tiene reportes elegidos a mano (entonces un reporte que falta no frena)."""
    try:
        with open(REPORTES_OBRA, encoding="utf-8") as f:
            return bool(json.load(f).get(obra))
    except (OSError, ValueError):
        return False


def guardar_elegidas(obra, plantillas):
    """Recuerda los reportes elegidos para la obra. Lista vacía = volver a los de por defecto."""
    d = {}
    if os.path.isfile(REPORTES_OBRA):
        try:
            with open(REPORTES_OBRA, encoding="utf-8") as f:
                d = json.load(f)
        except (OSError, ValueError):
            d = {}
    if plantillas:
        d[obra] = list(plantillas)
    else:
        d.pop(obra, None)
    os.makedirs(os.path.dirname(REPORTES_OBRA), exist_ok=True)
    with open(REPORTES_OBRA, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)


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


def leer_params():
    return leer_kv(PARAMS)


def params_vigentes(modelo):
    """export_params.txt de esta misma exportación (mismo modelo y de menos de PARAMS_MAX_MIN), o None."""
    p = leer_params()
    if not p or p.get("MODELO") != modelo or "CREADO" not in p:
        return None
    try:
        edad = (datetime.now() - datetime.strptime(p["CREADO"], "%Y-%m-%d %H:%M:%S")).total_seconds() / 60
    except ValueError:
        return None
    return p if edad <= PARAMS_MAX_MIN else None


def escribir_params(p):
    """Escribe export_params.txt. Con una sola fase deja también las claves viejas (FASE, DESTINO...)."""
    fases = [x for x in p["FASES"].split(",") if x]
    lineas = [f"CREADO={p['CREADO']}", f"MODELO={p['MODELO']}", f"FASES={','.join(fases)}", "CONFIG_CNC=Todo"]
    for f in fases:
        for k in ("CNC", "SUB", "TIP"):
            lineas.append(f"DESTINO_{k}_{f}={p[f'DESTINO_{k}_{f}']}")
    if len(fases) == 1:
        f = fases[0]
        lineas += [f"FASE={f}", f"DESTINO={p[f'DESTINO_CNC_{f}']}", f"DESTINO_SUB={p[f'DESTINO_SUB_{f}']}",
                   f"DESTINO_TIP={p[f'DESTINO_TIP_{f}']}"]
    os.makedirs(DIR_LOG, exist_ok=True)
    with open(PARAMS, "w", encoding="utf-8") as fo:
        fo.write("\n".join(lineas) + "\n")


def con_archivos(p):
    """Archivos reales en una carpeta (sin desktop.ini y similares)."""
    if not os.path.isdir(p):
        return []
    return [x for x in os.listdir(p) if x.lower() not in ("desktop.ini", "thumbs.db") and not x.startswith("~$")]


def preparar(args):
    # (2) Nunca quedan parámetros de la vez anterior: si este comando frena, las macros no tienen qué exportar.
    #     Con --agregar (varias fases en la misma corrida) se suma esta fase a los parámetros vigentes.
    if os.path.isfile(PARAMS) and not args.agregar:
        os.remove(PARAMS)
    pr = leer_probe(PARAMS_MAX_MIN if args.agregar else PROBE_MAX_MIN)
    if args.obra and args.obra.upper() != pr["OBRA"]:
        sys.exit(f"FRENA: pediste la obra {args.obra.upper()} pero el modelo abierto es {pr['OBRA']} ({pr['MODELO']}).")
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
        args.categoria = args.categoria or fila.get("CATEGORIA", "")
    _, fdir, existe = rutas(args, PARAMS_MAX_MIN)
    wdir = carpeta_trabajo(pr["OBRA"], fdir)

    print(f"Carpeta de fase: {fdir}  ({'ya existe' if existe else 'SE CREA'})")
    ocupadas = []
    for s in SUBCARPETAS:
        n = len(con_archivos(os.path.join(fdir, s)))
        if n:
            ocupadas.append(f"{s} ({n})")
    print(f"   Se arma en esta PC: {wdir}  (CNC, SUB, TIP y reportes; al terminar se copia todo junto a la carpeta de fase)")

    # (14) Rutas largas
    largo = len(os.path.join(fdir, "SUB", "")) + 45
    if largo > LARGO_MAX:
        print(f"AVISO: las rutas van a tener unos {largo} caracteres; arriba de 260 Windows falla. Conviene acortar el nombre de la carpeta.")

    # (12) Otra persona exportando la misma fase
    bl, horas = leer_bloqueo(fdir)
    if bl and bl.get("USUARIO") != getpass.getuser() and horas < BLOQUEO_MAX_H:
        print(f"FRENA: {bl.get('USUARIO')} está exportando esta fase desde {bl.get('DESDE')} (en {bl.get('PC')}). Hablalo antes de seguir.")
        return 1

    # (15) Fase con archivos en el Drive: no se exporta encima (reentrega = otro flujo).
    if ocupadas and not args.retomar:
        print("FRENA: la carpeta de fase ya tiene archivos en " + ", ".join(ocupadas) + ". No se exporta encima (reentrega = otro flujo).")
        print("       Si es la misma exportación que se cortó a la mitad, repetir con --retomar.")
        return 1
    a_medias = [s for s in SUBCARPETAS if con_archivos(os.path.join(wdir, s))]
    if a_medias:
        print(f"AVISO: en esta PC hay una exportación a medias de esta fase ({', '.join(a_medias)})"
              + (": se sigue desde ahí (--retomar)." if args.retomar else ": se aparta y se empieza de cero."))
    if not args.crear:
        print("\nSimulación: no se creó nada. Repetir con --crear después del OK del usuario.")
        return 0

    if a_medias and not args.retomar:
        aparte = wdir + " (anterior " + datetime.now().strftime("%Y-%m-%d %H%M%S") + ")"
        os.rename(wdir, aparte)          # se aparta, no se borra
        print(f"La exportación a medias quedó en: {aparte}")
    os.makedirs(fdir, exist_ok=True)
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(os.path.join(fdir, BLOQUEO), "w", encoding="utf-8") as f:
        f.write(f"USUARIO={getpass.getuser()}\nPC={os.environ.get('COMPUTERNAME', '?')}\nDESDE={ahora}\nMODELO={pr['MODELO']}\n")
    for s in SUBCARPETAS:
        os.makedirs(os.path.join(wdir, s), exist_ok=True)
    p = params_vigentes(pr["MODELO"]) if args.agregar else None
    if not p:
        p = {"CREADO": ahora, "MODELO": pr["MODELO"], "FASES": ""}
    fases = [x for x in p["FASES"].split(",") if x and x != str(args.fase)] + [str(args.fase)]
    p["FASES"] = ",".join(fases)
    for k in ("CNC", "SUB", "TIP"):
        p[f"DESTINO_{k}_{args.fase}"] = os.path.join(wdir, k)
    escribir_params(p)
    st = staging(pr["OBRA"], args.fase)
    if os.path.isdir(st):
        shutil.rmtree(st)          # carpeta temporal propia de la Skill, en TeklaMCP_logs
    os.makedirs(st)
    print(f"\nListo: carpetas creadas, bloqueo puesto y export_params.txt escrito (fases {p['FASES']}, vale {PARAMS_MAX_MIN // 60} horas).")
    print("Siguiente: seleccionar -> MCP_SeleccionarFase.cs -> reportes -> traer-reportes -> revisar -> "
          "MCP_ExportCNCFase.cs y MCP_ImprimirFasePDF.cs (una vez para todas las fases) -> finalizar.")
    return 0


def cmd_seleccionar(args):
    """Indica a MCP_SeleccionarFase.cs qué fase seleccionar (cuando export_params.txt tiene varias)."""
    p = leer_params()
    fases = [x for x in p.get("FASES", p.get("FASE", "")).split(",") if x]
    if str(args.fase) not in fases:
        print(f"FRENA: la fase {args.fase} no está en export_params.txt (fases {','.join(fases) or '-'}). Correr preparar --crear.")
        return 1
    with open(SELECCION, "w", encoding="utf-8") as f:
        f.write(f"FASE={args.fase}\n")
    print(f"Fase a seleccionar: {args.fase}. Siguiente: MCP_SeleccionarFase.cs")
    return 0


def plantillas_de(args, obra):
    if getattr(args, "plantillas", None):
        return [x.strip() for x in args.plantillas.split(";") if x.strip()]
    return lista_plantillas(obra, por_defecto=getattr(args, "por_defecto", False))


def sin_frenar(args, obra):
    """Reportes elegidos a mano: si falta uno, avisa y sigue (pedido del usuario 2026-11-14)."""
    return bool(getattr(args, "sin_frenar", False) or getattr(args, "plantillas", None)
                or (not getattr(args, "por_defecto", False) and elegidas_a_mano(obra)))


def cmd_reportes(args):
    pr = leer_probe(PARAMS_MAX_MIN)
    st = staging(pr["OBRA"], args.fase)
    os.makedirs(st, exist_ok=True)
    print(json.dumps({"carpeta": st, "plantillas": plantillas_de(args, pr["OBRA"])}, ensure_ascii=False, indent=1))
    return 0


def traer_reportes(args):
    pr, fdir, _ = rutas(args)
    wdir = carpeta_trabajo(pr["OBRA"], fdir)
    st = staging(pr["OBRA"], args.fase)
    esperadas = plantillas_de(args, pr["OBRA"])
    blando = sin_frenar(args, pr["OBRA"])
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
        dst = os.path.join(wdir, base + ".csv")
        if os.path.exists(dst):
            print(f"AVISO: ya existe {os.path.basename(dst)} en la fase; no se pisa.")
            continue
        shutil.copy2(src, dst)
        traidos.append(os.path.basename(dst))
    print("Reportes a la fase: " + (", ".join(traidos) or "ninguno"))
    if vacios:
        print("Vacíos, no van: " + ", ".join(vacios))
    if faltan:
        if blando:
            print("AVISO (reportes elegidos a mano, se sigue): no se generaron: " + ", ".join(faltan))
        else:
            print("FRENA: no se generaron: " + ", ".join(faltan) + ". Si la plantilla tiene otro nombre en esta obra, "
                  "elegila con 'Elegir reportes' (App) o anotala en plantillas.json.")
            return 1
    tipos = {tipo_csv(os.path.join(wdir, x)) for x in con_archivos(wdir) if x.lower().endswith(".csv")}
    falta_tipo = [t for t in ("CONJUNTOS", "SUBMATERIALES") if t not in tipos]
    sub_vacio = any(tipo_csv(os.path.join(st, os.path.splitext(t)[0] + e)) == "SUBMATERIALES"
                    for t in vacios for e in (".csv", ".xsr") if os.path.isfile(os.path.join(st, os.path.splitext(t)[0] + e)))
    if falta_tipo == ["SUBMATERIALES"] and (sub_vacio or not blando):
        # (2026-11-21) El reporte de submateriales salió vacío: puede ser una fase solo de típicos (correas Z,
        # plegados). No frena acá: lo decide "revisar" (sigue solo con correas CFZ; si no, pregunta).
        print("AVISO: el reporte de submateriales no tiene datos: se revisa si la fase lleva submateriales.")
        return 0
    if falta_tipo:
        if blando:
            print("AVISO: no hay reporte de " + " ni de ".join(t.lower() for t in falta_tipo)
                  + " con datos: esos controles se saltean.")
        else:
            print("FRENA: falta el reporte de conjuntos o el de submateriales con datos: sin ellos no se puede controlar la fase.")
            return 1
    return 0


def revisar(args):
    """(5) Revisión y estado de cada plano, con el JSON que devuelve get_drawings_properties (solo lectura)."""
    pr, fdir, _ = rutas(args)
    wdir = carpeta_trabajo(pr["OBRA"], fdir)
    with open(args.json, encoding="utf-8-sig") as f:
        datos = json.load(f)
    # (2026-11-14) el JSON puede traer los planos de varias fases juntas (se pide una sola vez): solo cuentan los de esta
    planos = {d["mark"].strip("[] "): d for d in datos.get("drawings", [])}
    subs = sorted(m for m in marcas_reporte(wdir, "SUBMATERIALES") if not COMPRADOS.match(m))
    tips = sorted(marcas_reporte(wdir, "CONJUNTOS"))
    marca_sin = os.path.join(staging(pr["OBRA"], args.fase), SIN_SUB_MARCA)
    if os.path.isfile(marca_sin):
        os.remove(marca_sin)       # archivo propio de la Skill: se decide de nuevo en cada revisión
    # (2026-11-21) Fase sin submateriales: ninguna pieza de los conjuntos tiene plano de submaterial creado (o el
    # reporte de submateriales salió vacío) y sí hay típicos. Correas Z (CFZ): se sigue sola. Otra: se pregunta.
    if tips and all(m not in planos for m in subs):
        cfz = all(CORREA_Z_TIP.match(t) for t in tips) or (subs and all(CORREA_Z_SUB.match(m) for m in subs))
        if cfz:
            print(f"Correas Z (CFZ): la fase {args.fase} no lleva submateriales. Se sigue solo con típicos (sin submateriales ni CNC).")
        elif getattr(args, "sin_submateriales", False):
            print(f"Fase {args.fase} sin submateriales (confirmado por el modelador): se sigue solo con típicos (sin CNC).")
        else:
            print(f"PREGUNTA: la fase {args.fase} no tiene ningún plano de submaterial creado"
                  + (f" ({len(subs)} pieza(s) en el reporte sin plano)" if subs else " (el reporte de submateriales está vacío)")
                  + ". ¿Esta fase no lleva submateriales (solo típicos)? Con el sí, repetir con --sin-submateriales.")
            return 3
        os.makedirs(os.path.dirname(marca_sin), exist_ok=True)
        with open(marca_sin, "w", encoding="utf-8") as fo:
            fo.write(f"FASE={args.fase}\nMOTIVO={'correas Z (CFZ)' if cfz else 'confirmado por el modelador'}\n"
                     f"PIEZAS_SIN_PLANO={','.join(subs)}\n")
        subs = []
    esperados = subs + tips
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


def log_fase(nombre, fase):
    """Log de una macro para una fase: el propio de la fase (nombre_<fase>.txt) o, si no hay, el general."""
    propio = os.path.join(DIR_LOG, f"{nombre}_{fase}.txt")
    return leer_kv(propio if os.path.isfile(propio) else os.path.join(DIR_LOG, f"{nombre}.txt")), \
        (propio if os.path.isfile(propio) else os.path.join(DIR_LOG, f"{nombre}.txt"))


def piezas_por_tipicos(fase):
    """(otra_fase, fuera) de partes_fase_<n>.txt: piezas de típicos de la fase con otra fase puesta, y piezas con la
    fase puesta que son de típicos de otra fase (no se exportan)."""
    ruta = os.path.join(DIR_LOG, f"partes_fase_{fase}.txt")
    otra, fuera = [], []
    if os.path.isfile(ruta):
        with open(ruta, encoding="utf-8-sig", errors="replace") as f:
            for l in f:
                if l.startswith("OTRA_FASE="):
                    otra.append(l.split("=", 1)[1].strip())
                elif l.startswith("FUERA="):
                    fuera.append(l.split("=", 1)[1].strip())
                elif l.startswith("BULON_FUERA="):
                    fuera.append("bulón de " + l.split("=", 1)[1].strip())
    return otra, fuera


def publicar(wdir, fdir):
    """(2026-11-14) Copia lo armado en esta PC a la carpeta de fase del Drive, de una vez.
    Nunca pisa: lo que ya está igual se saltea; lo que está distinto se informa. Devuelve la lista de problemas."""
    problemas, copiados = [], 0
    for raiz, _, nombres in os.walk(wdir):
        for n in nombres:
            if n.lower() in ("desktop.ini", "thumbs.db") or n.startswith("~$"):
                continue
            o = os.path.join(raiz, n)
            d = os.path.join(fdir, os.path.relpath(o, wdir))
            if os.path.exists(d):
                if os.path.getsize(d) != os.path.getsize(o):
                    problemas.append(f"{os.path.relpath(d, fdir)}: ya está en el Drive y es distinto (no se pisa)")
                continue
            os.makedirs(os.path.dirname(d), exist_ok=True)
            tmp = d + ".copiando"
            shutil.copy2(o, tmp)
            if os.path.getsize(tmp) != os.path.getsize(o):
                problemas.append(f"{os.path.relpath(d, fdir)}: la copia quedó incompleta")
                continue
            os.replace(tmp, d)
            copiados += 1
    print(f"  Copiados al Drive: {copiados} archivo(s) → {fdir}")
    return problemas


def finalizar(args):
    pr, fdir, _ = rutas(args)
    wdir = carpeta_trabajo(pr["OBRA"], fdir)
    if not os.path.isdir(wdir):
        if os.path.isdir(os.path.join(fdir, "SUB")):
            wdir = fdir            # exportación vieja, armada directo en el Drive
        else:
            print("FRENA: no existe " + wdir)
            return 1
    blando = sin_frenar(args, pr["OBRA"])
    errores, avisos = [], []
    sin_sub = leer_kv(os.path.join(staging(pr["OBRA"], args.fase), SIN_SUB_MARCA))
    sin_sub = sin_sub if sin_sub.get("FASE") == str(args.fase) else None
    if sin_sub:
        # (2026-11-21) Fase solo de típicos: el CNC que haya sacado la macro (de las piezas principales) no va.
        # Se aparta fuera de la carpeta de trabajo (no se borra) y no se copia al Drive.
        print(f"Fase sin submateriales ({sin_sub.get('MOTIVO', '?')}): van solo los típicos y los reportes; sin CNC.")
        d_cnc = os.path.join(wdir, "CNC")
        if con_archivos(d_cnc):
            aparte = wdir + " (CNC que no va, sin submateriales " + datetime.now().strftime("%Y-%m-%d %H%M%S") + ")"
            os.makedirs(aparte, exist_ok=True)
            shutil.move(d_cnc, os.path.join(aparte, "CNC"))
            os.makedirs(d_cnc, exist_ok=True)
            print(f"  El CNC que salió de la macro quedó aparte (no va): {aparte}")

    # (4) Lo que dicen las macros: planos sin actualizar, errores de PDF, CNC que no salieron.
    pdf, ruta_pdf = log_fase("imprimir_pdf_log", args.fase)
    if pdf.get("FASE") != str(args.fase) or pdf.get("FIN") != "OK":
        errores.append(f"el log de PDF no es de esta fase o no terminó ({os.path.basename(ruta_pdf)}): correr MCP_ImprimirFasePDF.cs")
    else:
        if int(pdf.get("NO_ACTUALIZADOS", "0") or 0):
            errores.append(f"{pdf['NO_ACTUALIZADOS']} plano(s) no actualizados al imprimir el PDF: actualizarlos, borrar esos PDF y repetir")
        if int(pdf.get("ERRORES", "0") or 0):
            errores.append(f"{pdf['ERRORES']} PDF con error (ver {os.path.basename(ruta_pdf)})")
    cnc, ruta_cnc = log_fase("export_cnc_log", args.fase)
    if sin_sub:
        pass                       # sin submateriales: el CNC no va, no se controla
    elif cnc.get("FASE") != str(args.fase) or cnc.get("FIN") != "OK":
        motivo = next((k + v for k, v in cnc.items() if k.startswith(("ERROR", "EXCEPCION"))), "")
        if args.retomar and "ya tiene archivos .nc1" in motivo:
            avisos.append("CNC: ya estaban de la vez anterior (retomar); se controlan contra los PDF.")
        else:
            errores.append(f"el CNC no terminó bien ({os.path.basename(ruta_cnc)}){': ' + motivo[:160] if motivo else ''}")
    else:
        with open(ruta_cnc, encoding="utf-8-sig", errors="replace") as f:
            sin = [l.split()[-1] for l in f if l.startswith("  SIN") and not SIN_CNC.match(l.split()[-1])]
        if sin:
            errores.append("CNC que no se generaron: " + ", ".join(sin))

    for x in os.listdir(wdir):
        if x.lower().endswith(".xsr"):
            nuevo = os.path.join(wdir, x[:-4] + ".csv")
            if not os.path.exists(nuevo):
                os.rename(os.path.join(wdir, x), nuevo)
                print(f"renombrado: {x} -> {os.path.basename(nuevo)}")

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # con python -P no está en el path
    import reportes_excel  # noqa: E402
    for x in sorted(os.listdir(wdir)):
        if x.lower().endswith(".csv"):
            if reporte_vacio(os.path.join(wdir, x)):
                print(f"vacío: {x} no tiene datos; no va (no se arma su Excel).")
                continue
            if sin_sub and tipo_csv(os.path.join(wdir, x)) == "SUBMATERIALES":
                print(f"sin submateriales: {x} no va (no se arma su Excel).")
                continue
            _, msg = reportes_excel.generar(os.path.join(wdir, x))
            print("excel: " + msg)

    def lista(s, ext):
        p = os.path.join(wdir, s)
        return sorted(os.path.splitext(x)[0] for x in os.listdir(p) if x.lower().endswith(ext)) if os.path.isdir(p) else []

    sub, cnc_l, tip = lista("SUB", ".pdf"), lista("CNC", ".nc1"), lista("TIP", ".pdf")
    sin_cnc = [b for b in sub if SIN_CNC.match(b) and b not in cnc_l]
    if sin_cnc:
        print("  Rejas / escalones de reja (se compran hechos: PDF sí, CNC no): " + ", ".join(sin_cnc))
    for b in sorted(set(sub) ^ set(cnc_l)):
        if not COMPRADOS.match(b) and b not in sin_cnc:
            errores.append(f"{b}: {'falta el CNC' if b in sub else 'falta el PDF'}")
    # (2026-11-14) Se exporta por TÍPICOS: una pieza de un típico de la fase puede tener otro número (mal puesto).
    otra_fase, fuera = piezas_por_tipicos(args.fase)
    otra = [x for x in sub + cnc_l if not re.match(r"^0*%d_" % args.fase, x)]
    if otra:
        avisos.append("piezas con otro número de fase dentro de típicos de esta fase (van igual, por típico): "
                      + ", ".join(sorted(set(otra))))
    if otra_fase:
        avisos.append(f"{len(otra_fase)} pieza(s) de típicos de esta fase tienen otra fase en el modelo (van igual): "
                      + ", ".join(otra_fase[:15]) + (" …" if len(otra_fase) > 15 else ""))
    if fuera:
        avisos.append(f"{len(fuera)} pieza(s) o bulón(es) tienen la fase {args.fase} pero son de típicos de otra fase (NO van): "
                      + ", ".join(fuera[:15]) + (" …" if len(fuera) > 15 else ""))
    # (3) piezas del reporte de submateriales contra lo exportado
    pzs = {m for m in marcas_reporte(wdir, "SUBMATERIALES") if not COMPRADOS.match(m)}
    if sin_sub:
        if sub:
            errores.append("la fase se marcó sin submateriales pero salieron PDF de submateriales: " + ", ".join(sub[:15]))
        pzs = set()
    elif pzs:
        for m in sorted(pzs - set(sub)):
            errores.append(f"{m}: está en el reporte de submateriales pero no tiene PDF")
        for m in sorted(set(sub) - pzs):
            errores.append(f"{m}: tiene PDF pero no está en el reporte de submateriales")
    elif blando:
        avisos.append("sin reporte de submateriales: no se cruzan las piezas con un reporte.")
    esperados = marcas_reporte(wdir, "CONJUNTOS")
    for t in sorted(esperados - set(tip)):
        errores.append(f"típico {t}: está en el reporte de conjuntos pero falta su PDF")
    for t in sorted(set(tip) - esperados) if esperados else []:
        errores.append(f"típico {t}: tiene PDF pero no está en el reporte de conjuntos")
    if not esperados and blando:
        avisos.append("sin reporte de conjuntos: no se cruzan los típicos con un reporte.")
    comprados = [x for x in cnc_l + sub if COMPRADOS.match(x)]
    # (14) rutas largas (se miden como van a quedar en el Drive) y (13) archivos vacíos
    for raiz, _, archivos in os.walk(wdir):
        for a in archivos:
            r = os.path.join(raiz, a)
            en_drive = os.path.join(fdir, os.path.relpath(r, wdir))
            if len(en_drive) > 255:
                errores.append(f"ruta de {len(en_drive)} caracteres: {en_drive}")
            if os.path.getsize(r) == 0 and a != BLOQUEO:
                errores.append(f"archivo vacío (0 bytes): {a}")

    print(f"\nFase {args.fase} armada en {wdir}")
    print(f"  SUB {len(sub)} PDF · CNC {len(cnc_l)} · TIP {len(tip)} PDF (reporte de conjuntos: {len(esperados)}, de submateriales: {len(pzs)})")
    reps = sorted(x for x in os.listdir(wdir) if x.lower().endswith(".xlsx"))
    print("  Reportes: " + (", ".join(reps) or "ninguno"))
    if comprados:
        print("  Tuercas/arandelas (no se suben nunca): " + ", ".join(comprados))
    for a in avisos:
        print("AVISO: " + a)
    # (16) Strumis: es manual; solo se recuerda.
    if not glob.glob(os.path.join(fdir, "*.bswx")):
        print("  RECORDATORIO: después del APROBADO del checker, el modelador exporta el Strumis (.bswx) a mano con BIMReview Export y NUNCA renumera.")
    if errores:
        print("FRENA:")
        for e in errores:
            print("  - " + e)
        return 1
    if wdir != fdir:
        problemas = publicar(wdir, fdir)
        if problemas:
            print("FRENA al copiar al Drive:")
            for e in problemas:
                print("  - " + e)
            return 1
    bl = os.path.join(fdir, BLOQUEO)
    if os.path.isfile(bl):
        os.remove(bl)              # el bloqueo es de la Skill: se levanta al terminar bien
    print("OK: todo cruzado y copiado al Drive. Bloqueo levantado.")
    return 0


def totales_conjuntos(fdir):
    """(peso_kg, m2_pintura) de la línea de total del reporte de conjuntos (se reconoce por el contenido)."""
    for x in os.listdir(fdir):
        if x.lower().endswith(".csv") and tipo_csv(os.path.join(fdir, x)) == "CONJUNTOS":
            with open(os.path.join(fdir, x), encoding="cp1252", errors="replace") as f:
                for linea in f:
                    c = [y.strip() for y in linea.split(";")]
                    if len(c) > 3 and c[1].upper() == "CONJUNTOS":
                        nums = [y for y in c[2:] if re.fullmatch(r"\d+(?:[.,]\d+)?", y)]
                        if len(nums) >= 2:
                            return nums[0].replace(",", "."), nums[-1].replace(",", ".")
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
        # (2026-11-12) Galvanizado (en caliente o en frío) o un "-" en M2 - PINTURA: no lleva pintura, no se carga el M2.
        acabado = " ".join(str(fila.get(k, "")) for k in ("ACABADO SUPERFICIAL", "CATEGORIA"))
        guion = str(fila.get("M2 - PINTURA", "")).strip() in ("-", "–", "—")
        sin_pintura = bool(GALV.search(acabado)) or guion
        if sin_pintura:
            print("  Sin pintura (" + ("galvanizado" if GALV.search(acabado) else "'-' en M2 - PINTURA")
                  + "): el M2 no se carga; el peso sí.")
        for col, clave, val in (("PESO (Kg)", "PESO (KG)", peso), ("M2 - PINTURA", "M2 - PINTURA", m2)):
            if clave == "M2 - PINTURA" and (sin_pintura or val is None):
                continue
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
    """(2026-11-20) Todo lo que se muestra queda también en ~/.ferma/registros/AAAA-MM-DD.log."""
    aqui = os.path.dirname(os.path.abspath(__file__))
    if aqui not in sys.path:
        sys.path.insert(0, aqui)
    import registro
    return registro.con_registro("exportar " + " ".join(sys.argv[1:3]), _main)


def _main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    for nombre in ("preparar", "seleccionar", "reportes", "traer-reportes", "revisar", "finalizar", "planilla",
                   "manifiesto", "macros"):
        s = sp.add_parser(nombre)
        if nombre == "macros":
            s.add_argument("--dirs", help="carpetas de macros de modelado separadas por ';' (por defecto C:\\TeklaStructures\\*)")
            continue
        s.add_argument("--fase", type=int, required=True)
        if nombre == "seleccionar":
            continue
        s.add_argument("--base", help="carpeta madre (por defecto se busca OBRAS\\PASADO A CONTROL en esta PC)")
        s.add_argument("--obra", help="código de obra esperado; si no coincide con el modelo abierto, frena")
        if nombre in ("preparar", "planilla"):
            s.add_argument("--pestana", help="pestaña de la planilla (por defecto, el código de obra)")
            s.add_argument("--libro", choices=("real", "prueba"))
        if nombre == "preparar":
            s.add_argument("--crear", action="store_true")
            s.add_argument("--retomar", action="store_true", help="la misma exportación se cortó: sigue sin pisar nada")
            s.add_argument("--agregar", action="store_true",
                           help="varias fases en la misma corrida: suma esta fase a export_params.txt")
            s.add_argument("--categoria", default="", help="categoría para el nombre de la carpeta (por defecto, la de la planilla)")
            s.add_argument("--sin-planilla", action="store_true", help="pruebas: no controla la fila en la planilla")
        if nombre in ("reportes", "traer-reportes", "finalizar"):
            s.add_argument("--plantillas", help="reportes elegidos a mano, separados por ';' (si falta uno, no frena)")
            s.add_argument("--sin-frenar", action="store_true", help="si falta un reporte, avisa y sigue")
            s.add_argument("--por-defecto", action="store_true", help="usa los reportes por defecto aunque la obra tenga elegidos")
        if nombre == "finalizar":
            s.add_argument("--retomar", action="store_true", help="la exportación se retomó: CNC ya hechos no son error")
        if nombre == "revisar":
            s.add_argument("--json", required=True, help="respuesta de get_drawings_properties guardada en un archivo")
            s.add_argument("--sin-submateriales", action="store_true",
                           help="el modelador confirmó que la fase no lleva submateriales (solo típicos, sin CNC)")
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
    return {"preparar": preparar, "seleccionar": cmd_seleccionar, "reportes": cmd_reportes, "traer-reportes": traer_reportes,
            "revisar": revisar, "finalizar": finalizar, "planilla": cmd_planilla, "manifiesto": cmd_manifiesto,
            "macros": cmd_macros}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
