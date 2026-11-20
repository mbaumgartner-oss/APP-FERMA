# -*- coding: utf-8 -*-
"""
Registro de lo que hacen exportar, imprimir y subir (2026-11-20).

Todo lo que el motor muestra en la ventana de la App (o en la consola) se copia también a
    %USERPROFILE%\\.ferma\\registros\\AAAA-MM-DD.log
un archivo por día, en la PC de quien lo usa. Si algo falla y no hay captura, alcanza con mandar ese archivo.
No guarda claves: solo la salida del motor (las claves viven en planilla_real.json y nunca se imprimen).
Si no se puede escribir el registro, se sigue igual: nunca frena el trabajo.

Este archivo está repetido igual en las 3 Skills (exportar-fase, imprimir-fase, subir-fase).
"""
import contextlib
import datetime
import getpass
import os
import sys

DIR_REGISTROS = os.path.join(os.path.expanduser("~"), ".ferma", "registros")
_activo = [0]          # subir.py llama a reentrega.main(): se registra una sola vez (la de afuera)


class _Copia:
    """Escribe en la salida de siempre y además en el archivo del registro."""

    def __init__(self, original, archivo):
        self.original, self.archivo = original, archivo

    def write(self, s):
        n = self.original.write(s)
        try:
            self.archivo.write(s)
        except Exception:
            pass
        return n

    def flush(self):
        self.original.flush()
        try:
            self.archivo.flush()
        except Exception:
            pass

    def reconfigure(self, **kw):
        if hasattr(self.original, "reconfigure"):
            self.original.reconfigure(**kw)

    def __getattr__(self, nombre):
        return getattr(self.original, nombre)


@contextlib.contextmanager
def registrar(que):
    """with registrar("exportar"): ...  — copia stdout y stderr al registro del día."""
    r = {}
    if _activo[0]:
        yield r
        return
    try:
        os.makedirs(DIR_REGISTROS, exist_ok=True)
        ahora = datetime.datetime.now()
        f = open(os.path.join(DIR_REGISTROS, f"{ahora:%Y-%m-%d}.log"), "a", encoding="utf-8")
        f.write(f"\n===== {ahora:%Y-%m-%d %H:%M:%S} · {que} · {getpass.getuser()} en {os.environ.get('COMPUTERNAME', '?')}\n"
                f"      {' '.join(sys.argv[1:])}\n")
        f.flush()
    except Exception:
        yield r
        return
    viejo_out, viejo_err = sys.stdout, sys.stderr
    sys.stdout, sys.stderr = _Copia(viejo_out, f), _Copia(viejo_err, f)
    _activo[0] += 1
    fin = "terminó"
    try:
        yield r
        if "codigo" in r:
            fin = f"terminó (código {r['codigo'] or 0})"
    except SystemExit as e:
        fin = f"salió: {e.code}"
        raise
    except BaseException as e:
        fin = f"ERROR inesperado: {type(e).__name__}: {e}"
        raise
    finally:
        _activo[0] -= 1
        sys.stdout, sys.stderr = viejo_out, viejo_err
        try:
            f.write(f"----- {datetime.datetime.now():%H:%M:%S} {fin}\n")
            f.close()
        except Exception:
            pass


def con_registro(que, funcion):
    """Corre funcion() dentro del registro y anota el código de salida."""
    with registrar(que) as r:
        r["codigo"] = funcion()
        return r["codigo"]
