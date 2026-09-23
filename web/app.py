"""Banco Horizonte - Version web para Vercel (Flask).

Subir los 2 Excel, liquidar con R1-R8 y descargar Excel+PDF.
Sin estado en servidor y sin carpeta reportes/: los archivos se generan
en memoria y se entregan como enlaces data-URI (compatible serverless).
La logica vive en app/logica_liquidacion.py (modulo compartido).
"""
import base64
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "app"))

from flask import Flask, render_template, request  # noqa: E402
import pandas as pd  # noqa: E402

import logica_liquidacion as L  # noqa: E402

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB por subida


def _b64(datos: bytes, mime: str) -> str:
    return f"data:{mime};base64," + base64.b64encode(datos).decode("ascii")


@app.get("/")
def index():
    return render_template("index.html", error=None)


@app.post("/liquidar")
def liquidar():
    tope_txt = (request.form.get("tope") or "40").replace(",", ".")
    try:
        tope = float(tope_txt)
        if tope <= 0:
            raise ValueError
    except ValueError:
        return render_template("index.html",
                               error="El tope debe ser un número mayor que 0."), 400
    f_maestro = request.files.get("maestro")
    f_reg = request.files.get("registros")
    if not f_maestro or not f_maestro.filename or not f_reg or not f_reg.filename:
        return render_template("index.html",
                               error="Debe subir ambos archivos Excel."), 400
    try:
        df_m = L.cargar_maestro(f_maestro)
        df_r = L.cargar_registros(f_reg)
    except Exception as exc:  # hoja o columnas invalidas
        return render_template("index.html",
                               error=f"No se pudieron leer los Excel: {exc}"), 400
    res = L.aplicar_reglas(df_m, df_r, tope)
    resumen = res["resumen"]
    estados = [(e, resumen["por_estado"].get(e, 0)) for e in L.ESTADOS]
    xlsx_uri = _b64(L.exportar_excel_bytes(res),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    pdf_uri = _b64(L.exportar_pdf_bytes(res), "application/pdf")
    exceden = res["por_empleado"][res["por_empleado"]["Excede_Tope"] == "Sí"] \
        if not res["por_empleado"].empty else []
    deptos = sorted(df_m["Departamento"].unique().tolist())
    filas_emp = res["por_empleado"].to_dict("records") \
        if not res["por_empleado"].empty else []
    filas_dep = res["por_departamento"].to_dict("records") \
        if not res["por_departamento"].empty else []
    filas_ale = res["no_validos"].to_dict("records")
    for f in filas_ale:  # NaN -> texto legible
        if pd.isna(f.get("Horas")):
            f["Horas"] = "—"
    return render_template(
        "resultado.html", tope=tope_txt, mes=resumen["mes"],
        estados=estados, horas_total=L.fmt_horas(resumen["horas_total"]),
        horas_d=L.fmt_horas(resumen["horas_diurnas"]),
        horas_n=L.fmt_horas(resumen["horas_nocturnas"]),
        monto_total=L.fmt_crc(resumen["monto_total"]),
        filas_dep=filas_dep, filas_emp=filas_emp, filas_ale=filas_ale,
        top5=res["top5"].to_dict("records"), exceden=exceden, deptos=deptos,
        xlsx_uri=xlsx_uri, pdf_uri=pdf_uri, fmt_crc=L.fmt_crc,
        fmt_horas=L.fmt_horas)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
