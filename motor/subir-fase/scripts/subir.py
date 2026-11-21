# -*- coding: utf-8 -*-
"""
SUBIR UNA FASE - punto de entrada único de la Skill "subir-fase" (unificado 2026-11-07).

Una sola orden para el modelador: "subí la fase N". El script decide el modo:
  - PRIMERA SUBIDA: la fase está en APROBADO. Arma el plan con la convención que la obra ya tiene en el
    servidor (plan_obra.armar) y lo ejecuta (subir_plan.ejecutar_plan): manifiesto, nada se pisa y marca
    SUBIDO AL SERVIDOR.
  - REENTREGA: la fase está en SUBIDO AL SERVIDOR (o --fase-dir es una carpeta REENTREGA NN). Usa la última
    REENTREGA NN de la carpeta de la fase y la pasa a reentrega.py (OBSOLETOS, ENT siguiente, nota en la planilla).
Sin --ejecutar siempre simula. Nunca borra.

Uso:
  python subir.py --obra 702Q --fase-dir "<PASADO A CONTROL>\\702Q\\FASE 50" --constructivos "<...>\\CONSTRUCTIVOS" \
      --pestana 702Q [--libro real|prueba] [--ejecutar] [--avisos-vistos] [--acepto-cambios] [--informe x.md]
  Sin planilla en vivo (pruebas): --planilla <csv> o --categoria "<CAT>".
  Reentrega: además --avise-checker, --retirar, --mantener, --sin-lista, --confirmo (ver reentrega.py).
Códigos: 0 ok · 1 bloqueado · 2 uso · 3 falta confirmar (avisos o decisiones) · 4 ya existe
         5 subido pero planilla sin marcar · 6 cambió después de pasar a control / falta aviso al checker
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import plan_obra  # noqa: E402
import subir_plan  # noqa: E402
from control_fase import ESTADO_PARA_SUBIR, ESTADO_SUBIDO, leer_planilla  # noqa: E402

RE_REENTREGA = re.compile(r"^REENTREGA\s*(\d+)$", re.I)
RE_FASE = re.compile(r"^\s*(?:FASE|SEC|SECUENCIA)\s*0*(\d+)", re.I)


def ultima_reentrega(fase_dir):
    subs = [(int(m.group(1)), d) for d in os.listdir(fase_dir)
            if os.path.isdir(os.path.join(fase_dir, d)) and (m := RE_REENTREGA.match(d))]
    return os.path.join(fase_dir, max(subs)[1]) if subs else None


def main():
    """(2026-11-20) Todo lo que se muestra queda también en ~/.ferma/registros/AAAA-MM-DD.log."""
    aqui = os.path.dirname(os.path.abspath(__file__))
    if aqui not in sys.path:
        sys.path.insert(0, aqui)
    import registro
    return registro.con_registro("subir " + " ".join(sys.argv[1:3]), _main)


def _main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Subir una fase (primera subida o reentrega)")
    ap.add_argument("--obra", required=True)
    ap.add_argument("--fase-dir", required=True, help="carpeta de la fase (o su REENTREGA NN)")
    ap.add_argument("--constructivos", required=True)
    ap.add_argument("--fase", type=int, help="número de fase (por defecto, se lee del nombre de la carpeta)")
    fuente = ap.add_mutually_exclusive_group(required=True)
    fuente.add_argument("--pestana", help="planilla EN VIVO (lee el estado y, al ejecutar, marca o anota)")
    fuente.add_argument("--planilla", help="CSV de la pestaña (pruebas; la planilla queda como paso manual)")
    fuente.add_argument("--categoria", help="sin planilla (pruebas): categoría de la fase")
    ap.add_argument("--libro", choices=["real", "prueba"], default="real")
    ap.add_argument("--ejecutar", action="store_true")
    ap.add_argument("--avisos-vistos", action="store_true", help="el modelador vio y aceptó los avisos del plan")
    ap.add_argument("--acepto-cambios", action="store_true", help="sube aunque la fase cambió después de pasar a control")
    ap.add_argument("--informe")
    # reentrega (se pasan tal cual a reentrega.py)
    ap.add_argument("--avise-checker", action="store_true")
    ap.add_argument("--retirar", default="")
    ap.add_argument("--mantener", default="")
    ap.add_argument("--sin-lista", action="store_true")
    ap.add_argument("--confirmo", default="")
    a = ap.parse_args()

    fd = os.path.normpath(a.fase_dir)
    if not os.path.isdir(fd) or not os.path.isdir(a.constructivos):
        print("No existe --fase-dir o --constructivos (no se crea la raíz).", file=sys.stderr)
        return 2
    es_re_dir = bool(RE_REENTREGA.match(os.path.basename(fd)))
    carpeta_fase = os.path.dirname(fd) if es_re_dir else fd
    m = RE_FASE.match(os.path.basename(carpeta_fase))
    n = a.fase if a.fase is not None else (int(m.group(1)) if m else None)
    if n is None:
        print("No se pudo leer el número de fase del nombre de la carpeta: pasar --fase.", file=sys.stderr)
        return 2
    obra = a.obra.upper()

    # ---- estado y categoría en la planilla
    estado, categoria = None, a.categoria
    if a.pestana or a.planilla:
        if a.pestana:
            os.environ["FERMA_PLANILLA"] = a.libro
            import planilla  # noqa: E402
            print(f"[planilla en vivo: {a.libro}]")
            # (2026-11-19) Con la App abierta, lo leído hace un rato puede estar viejo (el checker acaba de poner
            # APROBADO). Al ejecutar se lee siempre de nuevo; en la simulación, si la fase no está en APROBADO se
            # confirma con una lectura nueva antes de frenar. Varias fases seguidas siguen usando una sola lectura.
            filas, _, _ = planilla.leer_pestana(a.pestana, fresco=a.ejecutar)
            fp = filas.get(n)
            if not a.ejecutar and not (fp and len(fp) == 1 and fp[0].estado == ESTADO_PARA_SUBIR):
                filas, _, _ = planilla.leer_pestana(a.pestana, fresco=True)
        else:
            filas, _, _ = leer_planilla(a.planilla)
        fp = filas.get(n)
        if not fp:
            print(f"BLOQUEADO: la fase {n} no está en la planilla. Si es no constructiva o un prearmado, no se sube.")
            return 1
        if len(fp) > 1:
            print(f"BLOQUEADO: la fase {n} aparece {len(fp)} veces en la planilla. Corregir la planilla.")
            return 1
        estado, categoria = fp[0].estado, fp[0].categoria

    # ---- ¿primera subida o reentrega?
    reentrega_dir = fd if es_re_dir else None
    if not reentrega_dir and estado == ESTADO_SUBIDO:
        reentrega_dir = ultima_reentrega(fd)
        if not reentrega_dir:
            print(f"BLOQUEADO: la fase {n} ya está en '{ESTADO_SUBIDO}'. Para una reentrega, poné lo que cambió en "
                  f"{os.path.join(fd, 'REENTREGA 01')} (o la siguiente) con su LISTA CNC COMPLETA.")
            return 1
    if reentrega_dir:
        print(f"[modo REENTREGA: {reentrega_dir}]")
        import reentrega  # noqa: E402
        argv = ["reentrega.py", "--obra", obra, "--reentrega-dir", reentrega_dir, "--constructivos", a.constructivos]
        if a.pestana:
            argv += ["--pestana", a.pestana, "--libro", a.libro]
        elif a.planilla:
            argv += ["--planilla", a.planilla]
        else:
            print("Una reentrega necesita la planilla (--pestana o --planilla).", file=sys.stderr)
            return 2
        for flag, val in (("--retirar", a.retirar), ("--mantener", a.mantener), ("--confirmo", a.confirmo)):
            if val:
                argv += [flag, val]
        for flag, val in (("--sin-lista", a.sin_lista), ("--avise-checker", a.avise_checker), ("--ejecutar", a.ejecutar)):
            if val:
                argv.append(flag)
        if a.informe:
            argv += ["--informe", a.informe]
        sys.argv = argv
        return reentrega.main()

    # ---- primera subida
    print("[modo PRIMERA SUBIDA]")
    if estado is not None and estado != ESTADO_PARA_SUBIR:
        print(f"BLOQUEADO: la fase {n} está en '{estado or '(vacío)'}'. Solo se sube con '{ESTADO_PARA_SUBIR}' (OK del checker).")
        return 1
    try:
        r = plan_obra.armar(obra, n, fd, a.constructivos, categoria=categoria, libro=a.libro)
    except ValueError as e:
        print(e, file=sys.stderr)
        return 2
    tips = [os.path.basename(d) for _, d in r["plan"] if d.startswith(os.path.join("TIPICOS", ""))
            and d.lower().endswith((".pdf", ".dxf"))]
    if not any(d.startswith(os.path.join("SUBMATERIALES", "")) for _, d in r["plan"]) and tips:
        if all(re.match(r"^CFZ", t, re.I) for t in tips):
            # (2026-11-21) Correas Z (CFZ): no llevan submateriales ni CNC. No se pregunta.
            print("Correas Z (CFZ): fase solo de típicos, sin submateriales ni CNC.")
        else:
            r["avisos"].append("No hay submateriales ni CNC: ¿es una fase solo de típicos (correas sin accesorios, plegados)?")
    resumen = plan_obra.texto_plan(r)
    print("\n".join(resumen))
    if r["errores"]:
        return 1
    if r["avisos"] and a.ejecutar and not a.avisos_vistos:
        print("FALTA CONFIRMAR: mostrale los avisos al modelador. Con su sí, repetir con --avisos-vistos.")
        return 3
    return subir_plan.ejecutar_plan(r["plan"], a.constructivos, pestana=a.pestana, secuencia=n, libro=a.libro,
                                    ejecutar=a.ejecutar, acepto_cambios=a.acepto_cambios, informe=a.informe,
                                    extra_md=resumen)


if __name__ == "__main__":
    sys.exit(main())
