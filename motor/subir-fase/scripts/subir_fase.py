# -*- coding: utf-8 -*-
"""
ATAJO (2026-11-07): la subida se unificó en subir.py. Este script queda para no romper órdenes viejas
y pasa los argumentos a subir.py:
  --confirmo <códigos>  ->  --avisos-vistos
  --multi-letra         ->  (se ignora: la convención de la obra se detecta sola)
Usar directamente:  python subir.py --obra ... --fase-dir ... --constructivos ... (--pestana | --planilla) [--ejecutar]
La versión anterior (mapeo fijo + control estricto) quedó guardada fuera del Drive; el control sigue en control_fase.py.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import subir  # noqa: E402


def traducir(argv):
    out, i = [], 0
    while i < len(argv):
        x = argv[i]
        if x == "--multi-letra":
            pass
        elif x == "--confirmo":
            i += 1
            if i < len(argv) and argv[i].strip():
                out.append("--avisos-vistos")
        elif x.startswith("--confirmo="):
            if x.split("=", 1)[1].strip():
                out.append("--avisos-vistos")
        else:
            out.append(x)
        i += 1
    return out


if __name__ == "__main__":
    print("[subir_fase.py es un atajo: se usa subir.py]")
    sys.argv = ["subir.py"] + traducir(sys.argv[1:])
    sys.exit(subir.main())
