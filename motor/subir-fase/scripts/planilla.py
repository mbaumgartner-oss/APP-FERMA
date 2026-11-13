# -*- coding: utf-8 -*-
"""
Cliente de la planilla en vivo, a través del Apps Script `apps-script/PlanillaSkill.gs`.
Por ahora solo para la copia de prueba "PRUEBA OBRAS".

La URL y el token se guardan FUERA del Drive (no se sincronizan):
  %USERPROFILE%\\.ferma\\planilla_prueba.json   -> {"url": "...", "token": "..."}

Uso:
  python planilla.py configurar --url "<URL /exec de la aplicación web>"   (genera el token y lo guarda)
  python planilla.py leer   --pestana 674L [--csv salida.csv]
  python planilla.py estado --pestana 674L --secuencia 5 --estado "SUBIDO AL SERVIDOR"
"""
import argparse
import csv
import json
import os
import secrets
import sys
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from control_fase import leer_planilla_filas  # noqa: E402

LIBROS = ("real", "prueba")   # real = "SECUENCIAS PARA OBRAS" · prueba = copia "PRUEBA OBRAS"


def ruta_config(libro=None):
    """Config de cada libro: %USERPROFILE%\\.ferma\\planilla_<libro>.json. Se elige con FERMA_PLANILLA (por defecto: real)."""
    libro = (libro or os.environ.get("FERMA_PLANILLA") or "real").lower()
    if libro not in LIBROS:
        sys.exit(f"Libro desconocido '{libro}'. Opciones: {', '.join(LIBROS)}")
    return os.path.join(os.path.expanduser("~"), ".ferma", f"planilla_{libro}.json")


CONFIG = ruta_config()


def cargar_config():
    if not os.path.exists(CONFIG):
        sys.exit(f"Falta la configuración {CONFIG}. Corré primero: planilla.py configurar --url ...")
    with open(CONFIG, encoding="utf-8") as f:
        return json.load(f)


def llamar(pedido):
    c = cargar_config()
    pedido = dict(pedido, token=c["token"])
    req = urllib.request.Request(c["url"], data=json.dumps(pedido).encode("utf-8"),
                                 headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=60) as r:   # Apps Script redirige; urllib lo sigue
        cuerpo = r.read().decode("utf-8")
    try:
        res = json.loads(cuerpo)
    except json.JSONDecodeError:
        sys.exit("La respuesta no es JSON. ¿La aplicación web pide inicio de sesión? Revisar 'Quién tiene acceso'.\n" + cuerpo[:300])
    if not res.get("ok"):
        sys.exit("Error de la planilla: " + str(res.get("error")))
    return res


def leer_pestana(pestana):
    """Devuelve (filas_por_secuencia, estados_validos, repetidas) igual que leer_planilla."""
    return leer_planilla_filas(llamar({"accion": "leer", "pestana": pestana})["valores"])


def nota(pestana, secuencia, texto, quien="Claude (Skill subir-fase)"):
    """Agrega una nota con fecha en la celda ESTADO, sin cambiar el estado (reentregas)."""
    return llamar({"accion": "nota", "pestana": pestana, "secuencia": secuencia, "texto": texto, "quien": quien})


def marcar(pestana, secuencia, estado, quien="Claude (Skill subir-fase)"):
    """Cambia el ESTADO de una secuencia. Devuelve la respuesta del Apps Script (fila, anterior, nuevo, fecha)."""
    return llamar({"accion": "estado", "pestana": pestana, "secuencia": secuencia, "estado": estado, "quien": quien})


def fila_secuencia(pestana, secuencia):
    """Devuelve {ENCABEZADO: valor} de la fila de UNA secuencia (busca el encabezado en cualquier columna).
    Sirve para leer DETALLO, MODELO, PESO (KG), M2 - PINTURA, ESTADO... sin asumir posiciones (2026-11-06)."""
    from control_fase import ENCABEZADOS_SECUENCIA, norm
    valores = llamar({"accion": "leer", "pestana": pestana})["valores"]
    cab, idx, encontradas = None, 0, []
    for r in valores:
        celdas = [str(c).strip() for c in r]
        n = [norm(c) for c in celdas]
        pos = next((j for j, c in enumerate(n) if c in ENCABEZADOS_SECUENCIA), None)
        if pos is not None:
            cab, idx = n, pos
            continue
        if cab and idx < len(celdas) and celdas[idx] == str(secuencia):
            encontradas.append({cab[j]: celdas[j] for j in range(min(len(cab), len(celdas))) if cab[j]})
    if len(encontradas) != 1:
        sys.exit(f"La secuencia {secuencia} aparece {len(encontradas)} veces en la pestaña {pestana}: no se sigue.")
    return encontradas[0]


def completar(pestana, secuencia, valores, quien="Claude (Skill exportar-fase)"):
    """Carga PESO (Kg) / M2 - PINTURA solo si la celda está vacía. Devuelve escritos / yaTenia / sinColumna."""
    return llamar({"accion": "completar", "pestana": pestana, "secuencia": secuencia, "valores": valores, "quien": quien})


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--libro", choices=LIBROS, help="planilla real (por defecto) o la copia de prueba")
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("configurar")
    c.add_argument("--url", required=True)
    al = sub.add_parser("alta", help="genera la clave de un compañero (piloto); no muestra la clave en pantalla")
    al.add_argument("--nombre", required=True, help="nombre con el que van a quedar firmados sus cambios")
    l = sub.add_parser("leer")
    l.add_argument("--pestana", required=True)
    l.add_argument("--csv")
    e = sub.add_parser("estado")
    e.add_argument("--pestana", required=True)
    e.add_argument("--secuencia", type=int, required=True)
    e.add_argument("--estado", required=True)
    e.add_argument("--quien", default="Claude (Skill subir-fase)")
    fi = sub.add_parser("fila", help="muestra todos los datos de una secuencia")
    fi.add_argument("--pestana", required=True)
    fi.add_argument("--secuencia", type=int, required=True)
    a = ap.parse_args()
    global CONFIG
    if a.libro:
        CONFIG = ruta_config(a.libro)
    print(f"[planilla: {os.path.basename(CONFIG)}]")

    if a.cmd == "configurar":
        if not a.url.startswith("https://script.google.com/"):
            sys.exit("La URL tiene que ser la de la aplicación web de Apps Script (https://script.google.com/...).")
        os.makedirs(os.path.dirname(CONFIG), exist_ok=True)
        conf = {"url": a.url, "token": secrets.token_urlsafe(32)}
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump(conf, f)
        print(f"Configuración guardada en {CONFIG}.")
        print("Copiá el valor de 'token' de ese archivo en Apps Script > Configuración del proyecto > Propiedades > TOKEN.")
        return 0
    if a.cmd == "alta":
        # Una clave por persona: el Apps Script firma cada cambio con este nombre (propiedad TOKENS).
        base = cargar_config()
        libro = os.path.basename(CONFIG)[len("planilla_"):-len(".json")]
        dir_alta = os.path.join(os.path.dirname(CONFIG), "altas")
        os.makedirs(dir_alta, exist_ok=True)
        mapa_ruta = os.path.join(os.path.dirname(CONFIG), f"tokens_{libro}.json")
        mapa = {}
        if os.path.exists(mapa_ruta):
            with open(mapa_ruta, encoding="utf-8") as f:
                mapa = json.load(f)
        if a.nombre in mapa.values():
            sys.exit(f"'{a.nombre}' ya tiene clave en {mapa_ruta}.")
        tok = secrets.token_urlsafe(32)
        mapa[tok] = a.nombre
        with open(mapa_ruta, "w", encoding="utf-8") as f:
            json.dump(mapa, f, ensure_ascii=False)
        seguro = "".join(ch if ch.isalnum() else "_" for ch in a.nombre)
        para_el = os.path.join(dir_alta, f"planilla_{libro} - {seguro}.json")
        with open(para_el, "w", encoding="utf-8") as f:
            json.dump({"url": base["url"], "token": tok, "usuario": a.nombre}, f, ensure_ascii=False)
        import subprocess
        subprocess.run("clip", input=("﻿" + json.dumps(mapa, ensure_ascii=False)).encode("utf-16-le"), check=True)
        print(f"Alta de '{a.nombre}' en el libro '{libro}'.")
        print(f"1. En el portapapeles quedó la lista completa de claves: pegarla en Apps Script > Configuración del proyecto > "
              f"Propiedades > TOKENS (reemplaza el valor anterior) y guardar.")
        print(f"2. Pasarle a {a.nombre} el archivo '{para_el}' por un medio privado (no el Drive).")
        print(f"   En su PC va en %USERPROFILE%\\.ferma\\planilla_{libro}.json")
        return 0
    if a.cmd == "leer":
        res = llamar({"accion": "leer", "pestana": a.pestana})
        if a.csv:
            with open(a.csv, "w", encoding="utf-8", newline="") as f:
                csv.writer(f).writerows(res["valores"])
        filas, estados, rep = leer_planilla_filas(res["valores"])
        print(f"Pestaña {res['pestana']}: {len(filas)} secuencias. Estados válidos: {', '.join(sorted(estados)) or '(no encontrados)'}")
        for n in sorted(filas):
            p = filas[n][0]
            print(f"  {n:>3}  {p.categoria:<30} {p.estado}")
        return 0
    if a.cmd == "fila":
        for k, v in fila_secuencia(a.pestana, a.secuencia).items():
            print(f"  {k:<22} {v}")
        return 0
    if a.cmd == "estado":
        res = marcar(a.pestana, a.secuencia, a.estado, a.quien)
        print(f"Fila {res['fila']}: '{res['anterior']}' → '{res['nuevo']}' ({res['fecha']})"
              + (" · se agregó el color del estado a la pestaña" if res.get("colorAgregado") else ""))
        return 0


if __name__ == "__main__":
    sys.exit(main())
