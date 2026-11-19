# -*- coding: utf-8 -*-
"""
Subida según un PLAN explícito (obras con estructura propia, como la 702, decidido 2026-11-05:
"seguí su propia estructura"). Claude arma el plan copiando la convención que la obra ya tiene en
el servidor; este script solo lo valida y lo ejecuta.

Plan: archivo de texto UTF-8, una línea por archivo:   <origen> | <destino relativo a CONSTRUCTIVOS>
Líneas vacías o que empiezan con # se ignoran.

Reglas:
  - Copia (no mueve). Nunca pisa: si un destino existe, FRENA (es una reentrega).
  - Sin --ejecutar solo muestra el plan.
  - Con --pestana/--secuencia verifica APROBADO antes y marca SUBIDO AL SERVIDOR después (planilla en vivo).
Uso:
  python subir_plan.py --plan plan.txt --constructivos "<...>\\CONSTRUCTIVOS" --pestana 702Q --secuencia 50 [--ejecutar] [--informe x.md]
Manifiesto (_fase.json, 2026-11-07): si la carpeta de la fase lo tiene y algo cambió después de pasar a control,
FRENA (código 6) hasta que el modelador confirme (--acepto-cambios). Se busca al lado de los archivos del plan.
Código de salida: 0 ok · 1 bloqueado · 2 uso · 4 ya existe · 5 subido pero planilla sin marcar · 6 cambió después de pasar a control
"""
import argparse
import datetime
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import manifiesto  # noqa: E402


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", required=True)
    ap.add_argument("--constructivos", required=True)
    ap.add_argument("--pestana")
    ap.add_argument("--secuencia", type=int)
    ap.add_argument("--libro", choices=["real", "prueba"], default="real")
    ap.add_argument("--ejecutar", action="store_true")
    ap.add_argument("--acepto-cambios", action="store_true",
                    help="el modelador confirmó subir aunque la fase cambió después de pasar a control")
    ap.add_argument("--informe")
    a = ap.parse_args()
    c = a.constructivos
    if not os.path.isdir(c):
        print("No existe --constructivos (no se crea la raíz).", file=sys.stderr)
        return 2
    plan = []
    with open(a.plan, encoding="utf-8") as f:
        for i, linea in enumerate(f, 1):
            linea = linea.strip()
            if not linea or linea.startswith("#"):
                continue
            if "|" not in linea:
                print(f"Línea {i} sin '|': {linea}", file=sys.stderr)
                return 2
            o, d = (x.strip() for x in linea.split("|", 1))
            if os.path.isabs(d) or ".." in d.replace("\\", "/").split("/"):
                print(f"Línea {i}: el destino tiene que ser relativo a CONSTRUCTIVOS: {d}", file=sys.stderr)
                return 2
            plan.append((o, d))
    return ejecutar_plan(plan, c, a.pestana, a.secuencia, a.libro, a.ejecutar, a.acepto_cambios, a.informe)


def mismo_archivo(a, b):
    """Mismo tamaño y mismo contenido (sha256)."""
    if os.path.getsize(a) != os.path.getsize(b):
        return False
    return manifiesto._huella(a) == manifiesto._huella(b)


def copiar_seguro(origen, destino):
    """Copia a un nombre temporal y recién al final le pone el nombre real: un corte nunca deja un archivo a medias
    con el nombre bueno. Nunca pisa: si el destino apareció mientras tanto, falla."""
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    tmp = destino + ".subiendo"        # si quedó de un corte anterior, se vuelve a escribir encima (es propio)
    shutil.copy2(origen, tmp)
    if os.path.getsize(tmp) != os.path.getsize(origen):
        raise IOError("tamaño distinto después de copiar")
    os.rename(tmp, destino)            # os.rename falla si el destino ya existe: no pisa


def ejecutar_plan(plan, c, pestana=None, secuencia=None, libro="real", ejecutar=False, acepto_cambios=False,
                  informe=None, extra_md=None):
    """Valida y (con ejecutar=True) copia un plan [(origen, destino relativo a CONSTRUCTIVOS)].
    Controla APROBADO en la planilla, el manifiesto y que nada se pise; al terminar marca SUBIDO AL SERVIDOR.
    extra_md: líneas que se agregan al informe (ej. el resumen de plan_obra). Devuelve el código de salida."""
    if not os.path.isdir(c):
        print("No existe --constructivos (no se crea la raíz).", file=sys.stderr)
        return 2
    for o, d in plan:
        if os.path.isabs(d) or ".." in d.replace("\\", "/").split("/"):
            print(f"El destino tiene que ser relativo a CONSTRUCTIVOS: {d}", file=sys.stderr)
            return 2
    faltan = [o for o, _ in plan if not os.path.isfile(o)]
    if faltan:
        print("BLOQUEADO: no existen estos archivos de origen:")
        for o in faltan:
            print("   ", o)
        return 1
    dests = [d.lower() for _, d in plan]
    if len(set(dests)) != len(dests):
        print("BLOQUEADO: hay dos archivos con el mismo destino.")
        return 1
    # Subida cortada a la mitad (2026-11-11): lo que ya está en el servidor IGUAL al origen se saltea, así repetir la
    # misma orden termina la subida. Lo que está DISTINTO frena como siempre (es una reentrega: no se pisa).
    ya, iguales = [], set()
    for o, d in plan:
        destino = os.path.join(c, d)
        if os.path.exists(destino):
            if os.path.isfile(destino) and mismo_archivo(o, destino):
                iguales.add(d)
            else:
                ya.append(d)
    if iguales:
        print(f"AVISO: {len(iguales)} archivo(s) ya están en el servidor iguales al origen (subida anterior cortada): "
              "no se vuelven a copiar.")
    if ya:
        print(f"FRENADO: {len(ya)} destino(s) ya existen (reentrega, no se pisa):")
        for d in ya:
            print("   ", d)
        return 4

    # manifiesto: lo que se sube tiene que ser lo que vio el checker (bloquea y pregunta)
    dir_man = manifiesto.buscar_dir([o for o, _ in plan])
    man = manifiesto.verificar(dir_man) if dir_man else None
    if manifiesto.hay_diferencias(man) and not acepto_cambios:
        print("FRENADO: la fase cambió después de pasar a control. Preguntar al modelador si sigue")
        print("(ej. corrección menor que el checker ya sabe). Con su sí: --acepto-cambios.")
        print("\n".join(manifiesto.texto(man)))
        return 6

    planilla = None
    if pestana:
        if secuencia is None:
            print("Con --pestana hace falta --secuencia.", file=sys.stderr)
            return 2
        os.environ["FERMA_PLANILLA"] = libro
        import planilla  # noqa: E402
        filas, _, _ = planilla.leer_pestana(pestana)
        fp = filas.get(secuencia)
        if not fp or len(fp) != 1:
            print(f"BLOQUEADO: la secuencia {secuencia} aparece {len(fp or [])} veces en {pestana}.")
            return 1
        if fp[0].estado != "APROBADO":
            print(f"BLOQUEADO: la secuencia {secuencia} está en '{fp[0].estado or '(vacío)'}'. Solo se sube con APROBADO.")
            return 1

    carpetas_nuevas = sorted({os.path.dirname(d) for _, d in plan if not os.path.isdir(os.path.join(c, os.path.dirname(d)))})
    marca = None
    if ejecutar:
        copiados = 0
        for o, d in plan:
            if d in iguales:
                continue
            destino = os.path.join(c, d)
            try:
                copiar_seguro(o, destino)
                copiados += 1
            except Exception as e:
                print(f"FRENA: la subida se cortó en {d} ({type(e).__name__}: {e}).")
                print(f"       Se copiaron {copiados} de {len(plan) - len(iguales)} archivo(s); la planilla NO se marcó.")
                print("       Revisá la conexión o si el archivo está abierto y repetí la MISMA orden: sigue donde quedó.")
                return 1
        if planilla:
            try:
                marca = planilla.marcar(pestana, secuencia, "SUBIDO AL SERVIDOR")
            except SystemExit as e:
                marca = {"ok": False, "error": str(e)}
            except Exception as e:
                marca = {"ok": False, "error": f"{type(e).__name__}: {e}"}

    out = [f"# Subida {pestana or ''} FASE {secuencia if secuencia is not None else ''} — "
           f"{'EJECUTADA' if ejecutar else 'SIMULACIÓN (no se copió nada)'}", "",
           f"- Fecha: {datetime.datetime.now():%Y-%m-%d %H:%M}", f"- Destino: `{c}`", f"- Archivos: {len(plan)}", "",
           "## Carpetas " + ("creadas" if ejecutar else "que se crean")]
    out += [f"- `{x}`" for x in carpetas_nuevas] or ["- (ninguna)"]
    out += ["", "## Archivos", "| Origen | Destino |", "|---|---|"]
    out += [f"| `{os.path.basename(o)}` | `{d}` |" for o, d in plan]
    out += ["", "## Planilla"]
    if not pestana:
        out.append("- Sin planilla en vivo: marcar a mano.")
    elif not ejecutar:
        out.append(f"- `{pestana}` secuencia {secuencia} en `APROBADO` (leído en vivo). Al ejecutar se marca `SUBIDO AL SERVIDOR`.")
    elif marca and marca.get("ok"):
        out.append(f"- Fila {marca['fila']}: `{marca['anterior']}` → `{marca['nuevo']}` ({marca['fecha']}).")
    else:
        out.append(f"- **NO se pudo marcar** ({(marca or {}).get('error')}). Marcar a mano `SUBIDO AL SERVIDOR`.")
    out += ["", "## Manifiesto"]
    if man is None:
        out.append("- Sin manifiesto (fase exportada a mano): no se pudo comparar con lo que vio el checker.")
    elif not manifiesto.hay_diferencias(man):
        out.append(f"- Igual a lo que pasó a control el {man['fecha']} ({man['usuario']}).")
    else:
        out.append("- **Cambios aceptados por el modelador** (`--acepto-cambios`):")
        out += ["  " + x for x in manifiesto.texto(man, limite=1000)]
    if extra_md:
        out += [""] + list(extra_md)
    md = "\n".join(out) + "\n"
    print(md)
    if informe:
        os.makedirs(os.path.dirname(os.path.abspath(informe)), exist_ok=True)   # (2026-11-19) se cortaba sin la carpeta
        with open(informe, "w", encoding="utf-8") as fo:
            fo.write(md)
    return 5 if (marca is not None and not marca.get("ok")) else 0



if __name__ == "__main__":
    sys.exit(main())
