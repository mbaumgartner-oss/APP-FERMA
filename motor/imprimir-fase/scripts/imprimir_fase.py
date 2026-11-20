# -*- coding: utf-8 -*-
"""
Imprimir fase (Skill "imprimir-fase", Ferma) - imprime una fase de PASADO A CONTROL para el checker.

Reglas (relevamiento 2026-11-09, [[Proceso - Imprimir fase para el checker]]):
- Se imprime todo menos los CNC (y la LISTA CNC COMPLETA). El orden no importa.
- Cada hoja con su tamaño y orientación: A4/A3 -> RICOH MP 4002sp · A2/A1 -> HP DesignJet T250.
- 1 copia, simple faz, blanco y negro.
- Reportes: se imprimen los .xlsx (los .csv no).
- Reentrega: solo la carpeta REENTREGA NN (lo que cambió).
- Registro de impresiones: PASADO A CONTROL\\_IMPRESIONES\\<usuario>.csv (uno por usuario, sin conflictos en el Drive).
  Sirve para no imprimir 2 veces, avisar de planos que quedaron viejos y medir el volumen del mes.

Uso:
  python -P imprimir_fase.py --obra 696A --fase 104                    -> PLAN (no imprime nada)
  python -P imprimir_fase.py --obra 696A --fase 104 --si               -> imprime (solo con el OK del modelador)
  Opciones: --reentrega 1 · --solo-cambios · --forzar (reimprimir lo ya impreso igual)
            --prueba <carpeta>  (imprime a PDF con "Microsoft Print to PDF", sin papel ni registro)
            --base <PASADO A CONTROL>
Códigos de salida: 0 OK · 1 frena (ver mensaje) · 2 error de uso.
"""
import argparse
import csv
import datetime
import getpass
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import time

AQUI = os.path.dirname(os.path.abspath(__file__))
BASE_DEFECTO = r"G:\.shortcut-targets-by-id\1p8yAck6fXKpSit08dYenjyFLlO_z5VE9\OBRAS\PASADO A CONTROL"
CARPETA_REGISTRO = "_IMPRESIONES"
IMPRESORAS = {"A4": "RICOH MP 4002sp", "A3": "RICOH MP 4002sp",
              "A2": "HP5E87AB (HP DesignJet T250 24-in)", "A1": "HP5E87AB (HP DesignJet T250 24-in)"}
IMPRESORA_PRUEBA = "Microsoft Print to PDF"
# (2026-11-20) En cada PC Windows le pone otro nombre a la misma impresora (ej. "RICOH Aficio MP 4002 [0026734B5DF9]").
# Si no está el nombre exacto, se busca por modelo. Para un caso raro: %USERPROFILE%\.ferma\impresoras.json
# con {"RICOH MP 4002sp": "<nombre en esta PC>", ...}.
MODELO_IMPRESORA = {"RICOH MP 4002sp": r"4002", "HP5E87AB (HP DesignJet T250 24-in)": r"T250"}
CONFIG_IMPRESORAS = os.path.join(os.path.expanduser("~"), ".ferma", "impresoras.json")
HOJAS = {"A4": (210, 297), "A3": (297, 420), "A2": (420, 594), "A1": (594, 841), "A0": (841, 1189)}
NO_SE_IMPRIMEN = {"CNC", "LISTA CNC COMPLETA", "STRUMIS", "OBSOLETOS"}
CAMPOS = ["fecha", "usuario", "obra", "fase", "reentrega", "archivo", "hoja", "orientacion", "paginas",
          "impresora", "sha256"]


def hoja_de(ancho, alto, tol=0.03):
    """Nombre ISO (A4..A0) de una página en mm, o None si no es estándar."""
    corto, largo = sorted((ancho, alto))
    for n, (c, l) in HOJAS.items():
        if abs(corto - c) <= c * tol and abs(largo - l) <= l * tol:
            return n
    return None


def carpeta_fase(base, obra, n):
    bo = os.path.join(base, obra)
    if not os.path.isdir(bo):
        sys.exit(f"FRENA: no existe la obra en PASADO A CONTROL: {bo}")
    pat = re.compile(r"^FASE\s*0*%d(?:\b|_|\s|$)" % n, re.I)
    for d in sorted(os.listdir(bo)):
        if os.path.isdir(os.path.join(bo, d)) and pat.match(d):
            return os.path.join(bo, d)
    sys.exit(f"FRENA: no hay carpeta de la FASE {n} en {bo}")


def archivos_a_imprimir(fdir, reentrega):
    """PDF y Excel de la fase (o de la reentrega), sin CNC ni LISTA CNC COMPLETA."""
    if reentrega is not None:
        raiz = None
        for d in os.listdir(fdir):
            m = re.match(r"^REENTREGA\s*0*(\d+)$", d, re.I)
            if m and int(m.group(1)) == reentrega and os.path.isdir(os.path.join(fdir, d)):
                raiz = os.path.join(fdir, d)
        if not raiz:
            sys.exit(f"FRENA: no hay carpeta REENTREGA {reentrega:02d} en {fdir}")
    else:
        raiz = fdir
    res = []
    for actual, dirs, nombres in os.walk(raiz):
        dirs[:] = [d for d in dirs if d.upper() not in NO_SE_IMPRIMEN and not re.match(r"^REENTREGA", d, re.I)]
        for n in nombres:
            if n.lower().endswith((".pdf", ".xlsx", ".xls")) and not n.startswith("~$"):
                res.append(os.path.join(actual, n))
    return raiz, sorted(res)


def huella(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def impresoras_instaladas():
    """Nombres de las impresoras de esta PC (None si no se pudieron leer)."""
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Command",
                            "[Console]::OutputEncoding=[Text.Encoding]::UTF8; Add-Type -AssemblyName System.Drawing; "
                            "[System.Drawing.Printing.PrinterSettings]::InstalledPrinters | ForEach-Object { $_ }"],
                           capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=60,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:
        return None
    nombres = [x.strip() for x in r.stdout.splitlines() if x.strip()]
    return nombres or None


def resolver_impresora(nombre, instaladas):
    """Nombre real en esta PC de la impresora `nombre`. Devuelve (real, None) o (None, motivo)."""
    try:
        with open(CONFIG_IMPRESORAS, encoding="utf-8") as f:
            propia = json.load(f).get(nombre)
    except (OSError, ValueError):
        propia = None
    if instaladas is None:
        return (propia or nombre), None          # no se pudo listar: se intenta igual y frena al imprimir si no está
    if propia:
        return (propia, None) if propia in instaladas else (None, f"'{propia}' (de impresoras.json) no está en esta PC")
    if nombre in instaladas:
        return nombre, None
    patron = MODELO_IMPRESORA.get(nombre)
    cand = [x for x in instaladas if patron and re.search(patron, x, re.I)]
    if len(cand) == 1:
        return cand[0], None
    if not cand:
        return None, f"no hay ninguna impresora '{nombre}' (ni del mismo modelo) en esta PC"
    return None, f"hay {len(cand)} impresoras del modelo de '{nombre}': {', '.join(cand)}"


def trabajador(modo, trabajos):
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
        json.dump(trabajos, f, ensure_ascii=False)
        entrada = f.name
    # (2026-11-14) Tiempo máximo (un cartel del driver no deja la App esperando para siempre) y sin consola negra.
    paginas = sum(len(t.get("paginas", [])) or 1 for t in trabajos)
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                            os.path.join(AQUI, "imprimir.ps1"), "-modo", modo, "-entrada", entrada],
                           capture_output=True, text=True, encoding="utf-8", errors="replace",
                           timeout=300 + 60 * paginas, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired:
        sys.exit(f"FRENA: imprimir.ps1 ({modo}) no terminó a tiempo. Fijate si la impresora o Excel tienen un cartel abierto.")
    finally:
        os.remove(entrada)
    salida = r.stdout.strip().splitlines()
    if r.returncode != 0 or not salida:
        sys.exit(f"FRENA: falló imprimir.ps1 ({modo}): {r.stderr.strip()[:800]}")
    datos = json.loads(salida[-1])
    return datos if isinstance(datos, list) else [datos]


def leer_registro(base, obra, fase):
    """Última impresión de cada archivo de esa obra/fase (todos los usuarios): {rel: fila}."""
    d = os.path.join(base, CARPETA_REGISTRO)
    ult = {}
    if not os.path.isdir(d):
        return ult
    for n in os.listdir(d):
        if not n.lower().endswith(".csv"):
            continue
        with open(os.path.join(d, n), encoding="utf-8-sig", newline="") as f:
            for fila in csv.DictReader(f, delimiter=";"):
                if fila.get("obra") == obra and fila.get("fase") == str(fase):
                    if fila["archivo"] not in ult or fila["fecha"] > ult[fila["archivo"]]["fecha"]:
                        ult[fila["archivo"]] = fila
    return ult


def escribir_registro(base, filas):
    d = os.path.join(base, CARPETA_REGISTRO)
    os.makedirs(d, exist_ok=True)
    ruta = os.path.join(d, re.sub(r"[^\w.-]", "_", getpass.getuser()) + ".csv")
    nuevo = not os.path.isfile(ruta)
    with open(ruta, "a", encoding="utf-8-sig" if nuevo else "utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS, delimiter=";")
        if nuevo:
            w.writeheader()
        w.writerows(filas)
    return ruta


def avisar_manifiesto(fdir):
    sys.path.insert(0, os.path.join(AQUI, "..", "..", "subir-fase", "scripts"))
    try:
        import manifiesto  # noqa: E402  (Skill subir-fase)
    except ImportError:
        return
    r = manifiesto.verificar(fdir)
    if r is None:
        print("Manifiesto: la fase no tiene _fase.json (todavía no se marcó PASADO A CONTROLAR).")
    elif manifiesto.hay_diferencias(r):
        print("AVISO: la carpeta no coincide con lo que se pasó a control:")
        for linea in manifiesto.texto(r):
            print("  " + linea)
    else:
        print(f"Manifiesto: coincide con lo pasado a control ({r['fecha']}).")


def main():
    ap = argparse.ArgumentParser(description="Imprimir una fase de PASADO A CONTROL para el checker")
    ap.add_argument("--obra", required=True)
    ap.add_argument("--fase", type=int, required=True)
    ap.add_argument("--reentrega", type=int)
    ap.add_argument("--base", help="OBRAS\\PASADO A CONTROL (por defecto se busca en esta PC)")
    ap.add_argument("--si", action="store_true", help="imprimir (solo con el OK del modelador)")
    ap.add_argument("--solo-cambios", action="store_true", help="solo lo nuevo o cambiado desde la última impresión")
    ap.add_argument("--forzar", action="store_true", help="reimprimir también lo ya impreso igual")
    ap.add_argument("--prueba", metavar="CARPETA", help="imprimir a PDF en esa carpeta (sin papel ni registro)")
    a = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")   # (2026-11-14) sin esto se cae al escribir "→" fuera de la consola
    obra = a.obra.upper()
    if not a.base:
        # La misma búsqueda que exportar-fase (otra letra de unidad, otro nombre del Drive, ~\.ferma\exportar.json)
        try:
            sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "exportar-fase", "scripts"))
            import exportar_fase  # noqa: E402
            a.base = exportar_fase.buscar_base()
        except ImportError:
            a.base = BASE_DEFECTO

    fdir = carpeta_fase(a.base, obra, a.fase)
    raiz, arch = archivos_a_imprimir(fdir, a.reentrega)
    print(f"Obra {obra} · FASE {a.fase}" + (f" · REENTREGA {a.reentrega:02d}" if a.reentrega is not None else ""))
    print(f"Carpeta: {raiz}")
    if not arch:
        print("FRENA: no hay PDF ni Excel para imprimir.")
        return 1
    avisar_manifiesto(fdir)

    # Excel -> PDF temporal (no se toca el .xlsx), después todo es PDF
    tmp = tempfile.mkdtemp(prefix="imprimir_fase_")
    xls = [x for x in arch if not x.lower().endswith(".pdf")]
    pdf_de = {x: x for x in arch if x.lower().endswith(".pdf")}
    hoja_excel = {}   # la hoja de un Excel sale de su configuración de página, no del PDF temporal
    if xls:
        r = trabajador("xlsx2pdf", [{"archivo": x, "pdf": os.path.join(tmp, "%02d.pdf" % i)} for i, x in enumerate(xls)])
        for t in r:
            if not t.get("ok"):
                print(f"FRENA: no se pudo leer {t['archivo']}: {t.get('error')}")
                return 1
            if not t.get("hoja"):
                if t.get("papel") in (0, None):
                    # (2026-11-14) Excel sin tamaño de hoja definido (pasa si se volvió a guardar a mano): va en A4
                    print(f"AVISO: {os.path.basename(t['archivo'])} no tiene tamaño de hoja: se imprime en A4.")
                    t["hoja"] = "A4"
                else:
                    print(f"FRENA: {t['archivo']}: hoja de Excel no reconocida (código {t.get('papel')})")
                    return 1
            pdf_de[t["archivo"]] = t["pdf"]
            hoja_excel[t["archivo"]] = (t["hoja"], bool(t["horizontal"]))

    info = {t["archivo"]: t for t in trabajador("info", [{"archivo": pdf_de[x]} for x in arch])}
    previo = leer_registro(a.base, obra, a.fase)
    filas, errores = [], []
    for x in arch:
        i = info[pdf_de[x]]
        rel = os.path.relpath(x, fdir).replace("\\", "/")
        if not i.get("ok"):
            errores.append(f"{rel}: no se pudo abrir ({i.get('error')})")
            continue
        hojas = []
        for p in i["paginas"]:
            w, h = (p["alto_mm"], p["ancho_mm"]) if p["giro"] in (1, 3) else (p["ancho_mm"], p["alto_mm"])
            if x in hoja_excel:
                hojas.append({"hoja": hoja_excel[x][0], "horizontal": hoja_excel[x][1], "mm": f"{w:.0f}x{h:.0f}"})
            else:
                hojas.append({"hoja": hoja_de(w, h), "horizontal": w > h, "mm": f"{w:.0f}x{h:.0f}"})
        malas = [hj["mm"] for hj in hojas if not hj["hoja"]]
        if malas:
            errores.append(f"{rel}: tamaño de hoja no estándar ({', '.join(malas)} mm)")
            continue
        tamanos = sorted({hj["hoja"] for hj in hojas})
        impresoras = sorted({IMPRESORAS.get(t, "?") for t in tamanos})
        if "?" in impresoras or len(impresoras) > 1:
            errores.append(f"{rel}: hojas {tamanos} sin una única impresora asignada")
            continue
        sha = huella(x)
        ant = previo.get(rel)
        estado = "nuevo" if not ant else ("ya impreso igual" if ant["sha256"] == sha else "CAMBIÓ desde la última impresión")
        filas.append({"rel": rel, "archivo": x, "pdf": pdf_de[x], "hojas": hojas, "tamanos": tamanos,
                      "impresora": impresoras[0], "sha": sha, "estado": estado, "anterior": ant})
    if errores:
        print("FRENA:")
        for e in errores:
            print("  - " + e)
        return 1

    # (2026-11-20) cada impresora con el nombre que tiene en ESTA PC
    instaladas = impresoras_instaladas()
    reales, faltan = {}, []
    for nombre in sorted({f["impresora"] for f in filas}):
        real, motivo = resolver_impresora(nombre, instaladas)
        if real:
            reales[nombre] = real
            if real != nombre:
                print(f"[impresora: '{nombre}' en esta PC se llama '{real}']")
        elif a.prueba:
            reales[nombre] = nombre
            print(f"AVISO: {motivo} (en la prueba no importa: va a PDF).")
        else:
            faltan.append(motivo)
    if faltan:
        print("FRENA: falta la impresora:")
        for m in faltan:
            print("  - " + m)
        print(f"  Instaladas en esta PC: {', '.join(instaladas or ['(no se pudieron leer)'])}")
        print("  Agregala en Windows (Configuración > Impresoras y escáneres) o pedile a MB que la configure.")
        return 1
    for f in filas:
        f["impresora"] = reales[f["impresora"]]

    # qué se manda
    for f in filas:
        f["va"] = (f["estado"] != "ya impreso igual" or a.forzar) and (not a.solo_cambios or f["estado"] != "ya impreso igual")
    ancho = max(46, max(len(f["rel"]) for f in filas))
    print(f"\n{'Archivo':<{ancho}} {'Hoja':<16} {'Pág':>3}  {'Impresora':<22} Estado")
    for f in filas:
        hj = ", ".join(sorted({f"{h['hoja']} {'horiz.' if h['horizontal'] else 'vert.'}" for h in f["hojas"]}))
        imp = f["impresora"].split(" (")[0]
        marca = "" if f["va"] else "  (no se imprime)"
        extra = f" · última {f['anterior']['fecha']}" if f["anterior"] else ""
        print(f"{f['rel']:<{ancho}} {hj:<16} {len(f['hojas']):>3}  {imp:<22} {f['estado']}{extra}{marca}")
    van = [f for f in filas if f["va"]]
    tot = {}
    for f in van:
        for h in f["hojas"]:
            tot[(f["impresora"], h["hoja"])] = tot.get((f["impresora"], h["hoja"]), 0) + 1
    print("\nTotal a imprimir: " + (" · ".join(f"{n} hoja(s) {hj} en {imp}" for (imp, hj), n in sorted(tot.items())) or "nada"))
    viejos = [f["rel"] for f in filas if f["estado"].startswith("CAMBIÓ")]
    if viejos:
        print(f"AVISO: {len(viejos)} archivo(s) cambiaron desde la última impresión: lo impreso antes quedó viejo.")
    if not van:
        print("Nada nuevo para imprimir (todo ya se imprimió igual). Para reimprimir: --forzar.")
        return 0
    if not a.si and not a.prueba:
        print("\nPLAN: no se imprimió nada. Repetir con --si después del OK del modelador"
              " (o --prueba <carpeta> para probar sin papel).")
        return 0

    # imprimir: UN envío por impresora y tamaño de hoja (como lo hace el usuario: mandar de a un archivo
    # es muchísimo más lento en la impresora, 2026-11-09). La orientación va página por página.
    grupos = {}
    for f in van:
        for i, h in enumerate(f["hojas"]):
            grupos.setdefault((f["impresora"], h["hoja"]), []).append(
                {"archivo": f["pdf"], "pagina": i, "hoja": h["hoja"], "horizontal": h["horizontal"], "rel": f["rel"],
                 "excel": f["archivo"] in hoja_excel})   # informes: grises tramados (2026-11-17)
    trabajos = []
    for (imp, hj), pags in sorted(grupos.items(), key=lambda x: list(HOJAS).index(x[0][1])):
        nombre = f"{obra} FASE {a.fase}" + (f" REENTREGA {a.reentrega:02d}" if a.reentrega is not None else "") + f" - {hj} ({len(pags)} hojas)"
        t = {"nombre": nombre, "impresora": IMPRESORA_PRUEBA if a.prueba else imp, "imp_real": imp, "hoja": hj,
             "paginas": [{k: v for k, v in p.items() if k != "rel"} for p in pags], "rels": [p["rel"] for p in pags]}
        if a.prueba:
            os.makedirs(a.prueba, exist_ok=True)
            t["salida"] = os.path.join(a.prueba, re.sub(r"[^\w .()-]", "_", nombre) + ".pdf")
        trabajos.append(t)
    print("\nEnvíos: " + " · ".join(f"{t['hoja']} {len(t['paginas'])} hoja(s) → {t['imp_real'].split(' (')[0]}" for t in trabajos))
    t0 = time.time()
    res = {r["nombre"]: r for r in trabajador("imprimir", [{k: v for k, v in t.items() if k not in ("rels", "imp_real", "hoja")}
                                                         for t in trabajos])}
    seg = time.time() - t0
    ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    fallados, mal = set(), []
    for t in trabajos:
        r = res.get(t["nombre"], {})
        if not (r.get("ok") and r.get("paginas") == len(t["paginas"])):
            fallados.update(t["rels"])
            mal.append(f"{t['nombre']}: {r.get('error') or 'hojas enviadas %s de %d' % (r.get('paginas'), len(t['paginas']))}")
    ok, reg = 0, []
    for f in van:
        if f["rel"] in fallados:
            continue
        ok += 1
        for hj in sorted({h["hoja"] for h in f["hojas"]}):
            pags = [h for h in f["hojas"] if h["hoja"] == hj]
            reg.append({"fecha": ahora, "usuario": getpass.getuser(), "obra": obra, "fase": a.fase,
                        "reentrega": a.reentrega if a.reentrega is not None else "", "archivo": f["rel"],
                        "hoja": hj, "orientacion": "horizontal" if pags[0]["horizontal"] else "vertical",
                        "paginas": len(pags), "impresora": f["impresora"], "sha256": f["sha"]})
    print(f"\nEnviados {len(trabajos) - len(mal)} de {len(trabajos)} envío(s) · {ok} de {len(van)} archivo(s) en {seg:.0f} s"
          + (f" (PRUEBA, a {a.prueba})" if a.prueba else ""))
    if mal:
        print("ERRORES (no se enviaron completos):")
        for m in mal:
            print("  - " + m)
    if reg and not a.prueba:
        print("Registro: " + escribir_registro(a.base, reg))
    return 1 if mal else 0


if __name__ == "__main__":
    sys.exit(main())
