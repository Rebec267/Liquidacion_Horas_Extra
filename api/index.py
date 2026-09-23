"""Entry point WSGI para Vercel.

Vercel solo reconoce como funciones Python los archivos bajo `api/` y exige
exponer la aplicacion como variable `app` a nivel de modulo en
`api/index.py`. Este archivo no contiene logica: reutiliza la Flask app de
`web/app.py` (que a su vez usa el modulo compartido
`app/logica_liquidacion.py`, el mismo de la version de escritorio).
"""
import os
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)  # raiz del repo (contiene web/ y app/)
sys.path.insert(0, os.path.join(RAIZ, "web"))

from app import app  # noqa: E402  (la Flask app de web/app.py)
