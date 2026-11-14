# -*- coding: utf-8 -*-
"""
Manifiesto de fase (_fase.json) - "lo que se sube es lo que vio el checker" (decidido 2026-11-07).

- Se ESCRIBE cuando el modelador confirma que la fase pasó a control (exportar_fase.py planilla
  --pasado-a-control, o exportar_fase.py manifiesto). Es la foto de lo que se mandó a controlar.
  Si la fase vuelve a control (cambio de modelo), se escribe de nuevo.
- Se VERIFICA al subir (subir.py / subir_plan.py). Si un archivo cambió, falta o es nuevo,
  la subida FRENA y pregunta. Solo sigue con --acepto-cambios (confirmación del modelador),
  y la lista queda en el informe.
- Fases exportadas a mano (sin manifiesto): no frena, avisa.

No cuentan: el propio manifiesto, el Strumis (.bswx, se exporta después del OK), su _Log.txt
y los .xsr (formato intermedio de Tekla).
"""
import datetime
import getpass
import hashlib
import json
import os

NOMBRE = "_fase.json"
NO_CUENTAN = (".bswx", "_log.txt", ".xsr")
# Archivos que crean Windows, Google Drive o Excel solos (2026-11-11): no cuentan, ni en manifiestos viejos.
SISTEMA = {"desktop.ini", "thumbs.db", ".ds_store"}


def _cuenta(nombre):
    n = nombre.lower()
    base = n.replace("\\", "/").rsplit("/", 1)[-1]
    if base in SISTEMA or base.startswith("~$"):
        return False
    return n != NOMBRE and not n.endswith(NO_CUENTAN)


def _huella(ruta):
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _archivos(fase_dir):
    res = {}
    for raiz, dirs, nombres in os.walk(fase_dir):
        dirs[:] = [d for d in dirs if d.upper() != "OBSOLETOS"]   # lo reemplazado no es parte de la fase (2026-11-14)
        for n in nombres:
            if _cuenta(n):
                p = os.path.join(raiz, n)
                res[os.path.relpath(p, fase_dir).replace("\\", "/")] = p
    return res


def escribir(fase_dir, datos=None):
    """Escribe (o reescribe) el manifiesto con todos los archivos de la fase. Devuelve la ruta."""
    arch = {}
    for rel, p in sorted(_archivos(fase_dir).items()):
        arch[rel] = {"bytes": os.path.getsize(p), "sha256": _huella(p)}
    m = {
        "version": 1,
        "fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "usuario": getpass.getuser(),
        "motivo": "pasado a control",
        **(datos or {}),
        "archivos": arch,
    }
    ruta = os.path.join(fase_dir, NOMBRE)
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(m, f, ensure_ascii=False, indent=1)
    return ruta


def leer(fase_dir):
    ruta = os.path.join(fase_dir, NOMBRE)
    if not os.path.isfile(ruta):
        return None
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


def verificar(fase_dir):
    """Compara la carpeta con su manifiesto.
    Devuelve None si no hay manifiesto; si no, {"fecha", "usuario", "cambiados", "faltan", "nuevos"}."""
    m = leer(fase_dir)
    if m is None:
        return None
    antes = {k: v for k, v in m.get("archivos", {}).items() if _cuenta(k)}   # manifiestos viejos tenían desktop.ini
    ahora = _archivos(fase_dir)
    cambiados, faltan, nuevos = [], [], []
    for rel, info in sorted(antes.items()):
        p = ahora.get(rel)
        if p is None:
            faltan.append(rel)
        elif os.path.getsize(p) != info.get("bytes") or _huella(p) != info.get("sha256"):
            cambiados.append(rel)
    nuevos = sorted(set(ahora) - set(antes))
    return {"fecha": m.get("fecha"), "usuario": m.get("usuario"),
            "cambiados": cambiados, "faltan": faltan, "nuevos": nuevos}


def hay_diferencias(r):
    return bool(r and (r["cambiados"] or r["faltan"] or r["nuevos"]))


def buscar_dir(origenes):
    """Para un plan (lista de rutas de origen): la carpeta de fase que tiene el manifiesto, o None.
    Mira la carpeta de cada archivo y hasta 2 niveles arriba (FASE xx\\CNC\\archivo.nc1)."""
    vistos = set()
    for o in origenes:
        d = os.path.dirname(os.path.abspath(o))
        for _ in range(3):
            if d in vistos:
                break
            vistos.add(d)
            if os.path.isfile(os.path.join(d, NOMBRE)):
                return d
            d = os.path.dirname(d)
    return None


def texto(r, limite=15):
    """Líneas legibles para mostrar al usuario."""
    out = [f"Manifiesto del {r['fecha']} ({r['usuario']}), cuando la fase pasó a control:"]
    for k, t in (("cambiados", "CAMBIÓ después de pasar a control"), ("faltan", "YA NO ESTÁ"),
                 ("nuevos", "NUEVO, no estaba cuando pasó a control")):
        if r[k]:
            out.append(f"  {t} ({len(r[k])}):")
            out += [f"    - {x}" for x in r[k][:limite]]
            if len(r[k]) > limite:
                out.append(f"    - … y {len(r[k]) - limite} más")
    return out
