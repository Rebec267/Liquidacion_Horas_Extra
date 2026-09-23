"""Banco Horizonte - Logica de liquidacion de horas extra (R1-R8).

Modulo COMPARTIDO sin dependencias graficas: lo importan tanto la app de
escritorio (Tkinter) como la version web (Flask). Solo pandas/openpyxl/reportlab.
"""
from io import BytesIO
from pathlib import Path
from datetime import date, datetime, timedelta
from collections import Counter

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# --------------------------------------------------------------------------
# Constantes
# --------------------------------------------------------------------------
AZUL = "#1F3A5F"            # encabezado y barra de estado
DORADO = "#C9A227"          # linea de acento institucional
BLANCO = "#FFFFFF"
ROJO_SUAVE = "#FADBD8"      # fila que excede el tope
FONDO = "#F1F4F8"           # fondo general gris muy claro
TARJETA = "#FFFFFF"         # tarjetas / paneles
TEXTO = "#1F2937"           # texto gris oscuro
TEXTO_SUAVE = "#5B6472"     # texto secundario
PRIMARIO = "#2456A6"        # botones principales (azul)
PRIMARIO_OSCURO = "#1B447F" # hover del boton principal
SECUNDARIO = "#E2E6EB"      # botones secundarios (gris)
CABECERA_TABLA = "#1F3A5F"  # fondo de cabeceras de tabla
FILA_PAR = "#F2F5F9"        # cebreado de tablas

HOJA_MAESTRO = "Empleados"
HOJA_REGISTROS = "Registros"
COLS_MAESTRO = ["Cedula", "Nombre", "Departamento", "Salario_Mensual_CRC",
                "Tope_Horas_Extra_Mes"]
COLS_REGISTROS = ["ID_Registro", "Cedula", "Fecha", "Horas", "Tipo",
                  "Aprobado_Por"]
FACTOR = {"Diurna": 1.5, "Nocturna": 2.0}
ESTADOS = ("VALIDO", "SIN EMPLEADO", "HORAS INVALIDAS", "DUPLICADO")

MESES_ES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
            "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


# --------------------------------------------------------------------------
# Utilidades de formato y normalizacion
# --------------------------------------------------------------------------
def limpiar_cedula(valor):
    """R1: Cedula sin guiones y sin espacios al inicio/fin. No toca el original."""
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.endswith(".0"):  # por si Excel entrego la cedula como numero
        texto = texto[:-2]
    return texto.replace("-", "")


def normalizar_depto(valor):
    """Repara el mojibake conocido de 'Tecnologia' (tildes corruptas en el xlsx)."""
    texto = str(valor).strip()
    if texto.lower().startswith("tecnolog"):
        return "Tecnología"
    return texto


def normalizar_fecha(valor):
    """Devuelve la fecha como 'YYYY-MM-DD' venga como serial, datetime o texto."""
    if pd.isna(valor) or str(valor).strip() == "":
        return ""
    if isinstance(valor, pd.Timestamp):
        return valor.date().isoformat()
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    texto = str(valor).strip()
    try:
        serial = float(texto)
        if 20000 < serial < 80000:  # serial Excel: dias desde 1899-12-30
            return (date(1899, 12, 30) + timedelta(days=int(serial))).isoformat()
    except ValueError:
        pass
    try:
        return pd.to_datetime(texto).date().isoformat()
    except Exception:
        return texto


def fmt_crc(valor):
    """Monto con separador de miles punto y simbolo de colon, ej. 1.980.448."""
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        numero = 0.0
    return "₡" + f"{numero:,.0f}".replace(",", ".")


def fmt_horas(valor):
    """Horas con un decimal, ej. 14.5."""
    try:
        return f"{float(valor):.1f}"
    except (TypeError, ValueError):
        return "0.0"


def mes_liquidado(fechas_iso):
    """Mes mas frecuente de los registros, ej. 'Agosto 2026'."""
    ym = [f[:7] for f in fechas_iso if len(str(f)) >= 7]
    if not ym:
        return "Mes no determinado"
    top = Counter(ym).most_common(1)[0][0]
    anio, mes = top.split("-")
    return f"{MESES_ES[int(mes) - 1]} {anio}"


# --------------------------------------------------------------------------
# Carga de archivos
# --------------------------------------------------------------------------
def cargar_maestro(path):
    df = pd.read_excel(path, sheet_name=HOJA_MAESTRO,
                       dtype={"Cedula": str})
    faltan = [c for c in COLS_MAESTRO if c not in df.columns]
    if faltan:
        raise ValueError(f"Maestro sin columnas {faltan} (hoja {HOJA_MAESTRO}).")
    df = df[COLS_MAESTRO].copy()
    df["Cedula_Limpia"] = df["Cedula"].apply(limpiar_cedula)
    df["Departamento"] = df["Departamento"].apply(normalizar_depto)
    df["Salario_Mensual_CRC"] = pd.to_numeric(df["Salario_Mensual_CRC"],
                                              errors="coerce")
    return df


def cargar_registros(path):
    df = pd.read_excel(path, sheet_name=HOJA_REGISTROS,
                       dtype={"Cedula": str, "ID_Registro": str})
    faltan = [c for c in COLS_REGISTROS if c not in df.columns]
    if faltan:
        raise ValueError(f"Registros sin columnas {faltan} (hoja {HOJA_REGISTROS}).")
    df = df[COLS_REGISTROS].copy()
    df["Cedula_Limpia"] = df["Cedula"].apply(limpiar_cedula)
    df["Fecha_Norm"] = df["Fecha"].apply(normalizar_fecha)
    df["Horas"] = pd.to_numeric(df["Horas"], errors="coerce")
    df["Tipo"] = df["Tipo"].astype(str).str.strip()
    return df


# --------------------------------------------------------------------------
# Motor de reglas R1-R8
# --------------------------------------------------------------------------
def aplicar_reglas(df_maestro, df_registros, tope):
    """Aplica R1-R8 al pie de la letra. Devuelve dict con DataFrames y resumen."""
    maestro = {str(c): r for c, r in
               df_maestro.set_index("Cedula_Limpia").iterrows()}

    val = df_registros.sort_values("ID_Registro").copy()  # R4: anterior = menor ID
    estados, tarifas, factores, montos = [], [], [], []
    vistos = {}  # (limpia, fecha, tipo, horas) -> ID_Registro del primero

    for _, fila in val.iterrows():
        limpia = str(fila["Cedula_Limpia"])
        horas = fila["Horas"]
        if limpia not in maestro:                       # R2
            estado = "SIN EMPLEADO"
        elif pd.isna(horas) or horas <= 0 or horas > 12:  # R3
            estado = "HORAS INVALIDAS"
        else:                                           # R4 (tras R2 y R3)
            clave = (limpia, str(fila["Fecha_Norm"]), str(fila["Tipo"]),
                     round(float(horas), 4))
            if clave in vistos:
                estado = "DUPLICADO"
            else:
                vistos[clave] = str(fila["ID_Registro"])
                estado = "VALIDO"
        estados.append(estado)
        if estado == "VALIDO":                          # R5 + R6
            salario = float(maestro[limpia]["Salario_Mensual_CRC"])
            tarifa = salario / 240.0
            factor = FACTOR.get(str(fila["Tipo"]), 1.5)
            monto = float(horas) * tarifa * factor
        else:
            tarifa, factor, monto = 0.0, 0.0, 0.0
        tarifas.append(tarifa)
        factores.append(factor if estado == "VALIDO" else 0.0)
        montos.append(round(monto, 2))

    val["Estado"] = estados
    val["Tarifa"] = [round(t, 2) for t in tarifas]
    val["Factor"] = factores
    val["Monto"] = montos

    validos = val[val["Estado"] == "VALIDO"].copy()
    validos["Nombre"] = validos["Cedula_Limpia"].apply(
        lambda c: maestro[str(c)]["Nombre"])
    validos["Departamento"] = validos["Cedula_Limpia"].apply(
        lambda c: maestro[str(c)]["Departamento"])
    validos["Salario"] = validos["Cedula_Limpia"].apply(
        lambda c: float(maestro[str(c)]["Salario_Mensual_CRC"]))

    # Por empleado (R7: excede tope pero se paga todo)
    filas_emp = []
    for ced, g in validos.groupby("Cedula_Limpia"):
        h_d = float(g[g["Tipo"] == "Diurna"]["Horas"].sum())
        h_n = float(g[g["Tipo"] == "Nocturna"]["Horas"].sum())
        filas_emp.append({
            "Cedula": ced,
            "Nombre": maestro[ced]["Nombre"],
            "Departamento": maestro[ced]["Departamento"],
            "Horas_Diurnas": round(h_d, 2),
            "Horas_Nocturnas": round(h_n, 2),
            "Horas_Totales": round(h_d + h_n, 2),
            "Monto": round(float(g["Monto"].sum()), 2),
            "Excede_Tope": "Sí" if (h_d + h_n) > tope else "No",
        })
    por_emp = pd.DataFrame(filas_emp)
    if not por_emp.empty:
        por_emp = por_emp.sort_values("Horas_Totales",
                                      ascending=False).reset_index(drop=True)

    # Por departamento (R8)
    if not validos.empty:
        por_dep = (validos.groupby("Departamento")
                   .agg(Registros=("ID_Registro", "count"),
                        Horas=("Horas", "sum"),
                        Monto=("Monto", "sum"))
                   .reset_index())
        por_dep["Horas"] = por_dep["Horas"].round(2)
        por_dep["Monto"] = por_dep["Monto"].round(2)
    else:
        por_dep = pd.DataFrame(columns=["Departamento", "Registros",
                                        "Horas", "Monto"])

    # Alertas: no validos + excede tope
    alertas = []
    for _, f in val[val["Estado"] != "VALIDO"].iterrows():
        alertas.append({"Tipo_Alerta": f["Estado"],
                        "ID_Registro": f["ID_Registro"],
                        "Cedula_Limpia": f["Cedula_Limpia"],
                        "Nombre": maestro.get(str(f["Cedula_Limpia"]),
                                              {}).get("Nombre", "—"),
                        "Departamento": maestro.get(str(f["Cedula_Limpia"]),
                                                    {}).get("Departamento", "—"),
                        "Fecha": f["Fecha_Norm"], "Horas": f["Horas"],
                        "Tipo": f["Tipo"],
                        "Detalle": f"Registro {f['Estado'].lower()}."})
    if not por_emp.empty:
        for _, e in por_emp[por_emp["Excede_Tope"] == "Sí"].iterrows():
            alertas.append({"Tipo_Alerta": "EXCEDE TOPE", "ID_Registro": "",
                            "Cedula_Limpia": e["Cedula"], "Nombre": e["Nombre"],
                            "Departamento": e["Departamento"], "Fecha": "",
                            "Horas": e["Horas_Totales"], "Tipo": "",
                            "Detalle": (f"{e['Nombre']} ({e['Departamento']}) "
                                        f"suma {e['Horas_Totales']}h > tope {tope}h.")})
    df_alertas = pd.DataFrame(alertas)

    resumen = {
        "por_estado": val["Estado"].value_counts().to_dict(),
        "horas_diurnas": round(float(validos[validos["Tipo"] == "Diurna"]
                                     ["Horas"].sum()), 2) if not validos.empty else 0.0,
        "horas_nocturnas": round(float(validos[validos["Tipo"] == "Nocturna"]
                                       ["Horas"].sum()), 2) if not validos.empty else 0.0,
        "monto_total": round(float(validos["Monto"].sum()), 2) if not validos.empty else 0.0,
        "mes": mes_liquidado(val["Fecha_Norm"].tolist()),
    }
    resumen["horas_total"] = round(resumen["horas_diurnas"]
                                   + resumen["horas_nocturnas"], 2)
    top5 = por_emp.head(5).copy() if not por_emp.empty else por_emp
    return {"validados": val, "por_empleado": por_emp, "por_departamento": por_dep,
            "alertas": df_alertas, "no_validos": val[val["Estado"] != "VALIDO"].copy(),
            "resumen": resumen, "top5": top5}


# --------------------------------------------------------------------------
# Exportacion (en memoria para la web, a disco para el escritorio)
# --------------------------------------------------------------------------
def exportar_excel_bytes(resultado):
    """Devuelve el Excel de liquidacion como bytes (sin usar reportes/)."""
    val = resultado["validados"].copy()
    cols_val = (["ID_Registro", "Cedula", "Cedula_Limpia", "Fecha", "Horas",
                 "Tipo", "Aprobado_Por", "Estado", "Tarifa", "Factor", "Monto"])
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        val[cols_val].to_excel(writer, sheet_name="Registros_Validados",
                               index=False)
        resultado["por_empleado"].to_excel(
            writer, sheet_name="Liquidacion_Por_Empleado", index=False)
        resultado["por_departamento"].to_excel(
            writer, sheet_name="Por_Departamento", index=False)
        resultado["alertas"].to_excel(writer, sheet_name="Alertas", index=False)
    return buffer.getvalue()


def exportar_pdf_bytes(resultado):
    """Devuelve el PDF de liquidacion como bytes (sin usar reportes/)."""
    res = resultado["resumen"]
    estilos = getSampleStyleSheet()
    titulo = estilos["Title"]
    titulo.textColor = colors.HexColor(AZUL)
    normal = estilos["Normal"]
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            title="Banco Horizonte - Liquidación de Horas Extra")
    partes = [Paragraph("Banco Horizonte - Liquidación de Horas Extra", titulo),
              Spacer(1, 6),
              Paragraph(f"Mes liquidado: {res['mes']} &nbsp;&nbsp;|&nbsp;&nbsp; "
                        f"Generado: {date.today().isoformat()}", normal),
              Spacer(1, 12)]

    def tabla(datos, anchos=None):
        t = Table(datos, colWidths=anchos, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(AZUL)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#F2F4F7")]),
        ]))
        return t

    filas = [["Estado", "Registros"]]
    for e in list(ESTADOS) + ["EXCEDE TOPE"]:
        if e in ESTADOS:
            n = res["por_estado"].get(e, 0)
        elif not resultado["por_empleado"].empty:
            n = int((resultado["por_empleado"]["Excede_Tope"] == "Sí").sum())
        else:
            n = 0
        filas.append([e, n])
    partes += [Paragraph("Resumen por estado", estilos["Heading2"]),
               tabla(filas), Spacer(1, 12)]

    dep = [["Departamento", "Registros", "Horas", "Monto"]]
    for _, f in resultado["por_departamento"].iterrows():
        dep.append([f["Departamento"], f["Registros"], fmt_horas(f["Horas"]),
                    fmt_crc(f["Monto"])])
    dep.append(["TOTAL", sum(r[1] for r in dep[1:]) if len(dep) > 1 else 0,
                fmt_horas(res["horas_total"]), fmt_crc(res["monto_total"])])
    partes += [Paragraph("Monto y horas por departamento", estilos["Heading2"]),
               tabla(dep), Spacer(1, 12)]

    top = [["Cédula", "Nombre", "Horas", "Monto"]]
    for _, f in resultado["top5"].iterrows():
        top.append([f["Cedula"], f["Nombre"], fmt_horas(f["Horas_Totales"]),
                    fmt_crc(f["Monto"])])
    partes += [Paragraph("Top 5 empleados con más horas válidas",
                         estilos["Heading2"]),
               tabla(top), Spacer(1, 12)]

    ale = [["Alerta", "ID", "Cédula", "Detalle"]]
    for _, f in resultado["alertas"].iterrows():
        ale.append([f["Tipo_Alerta"], f["ID_Registro"], f["Cedula_Limpia"],
                    Paragraph(str(f["Detalle"]), normal)])
    if len(ale) == 1:
        ale.append(["—", "—", "—", "Sin alertas."])
    partes += [Paragraph("Alertas", estilos["Heading2"]), tabla(ale)]
    doc.build(partes)
    return buffer.getvalue()


def ruta_reportes():
    """Carpeta reportes/ junto a la raiz del proyecto (solo escritorio)."""
    aqui = Path(__file__).resolve().parent
    raiz = aqui.parent if aqui.name == "app" else Path.cwd()
    destino = raiz / "reportes"
    destino.mkdir(parents=True, exist_ok=True)
    return destino


def exportar_excel(resultado, destino):
    """Guarda el Excel en disco (app de escritorio)."""
    with open(destino, "wb") as fh:
        fh.write(exportar_excel_bytes(resultado))


def exportar_pdf(resultado, destino):
    """Guarda el PDF en disco (app de escritorio)."""
    with open(destino, "wb") as fh:
        fh.write(exportar_pdf_bytes(resultado))
