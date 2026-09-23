"""Banco Horizonte (Costa Rica) - Liquidacion de Horas Extra.

App de escritorio (Tkinter) en UN solo archivo. Reglas R1-R8 segun AGENTS.md.
Uso:  python app/horas_extra_app.py
Dependencias: pandas, openpyxl, reportlab, tkinter (stdlib).
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
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
# Exportacion a reportes/
# --------------------------------------------------------------------------
def ruta_reportes():
    aqui = Path(__file__).resolve().parent
    raiz = aqui.parent if aqui.name == "app" else Path.cwd()
    destino = raiz / "reportes"
    destino.mkdir(parents=True, exist_ok=True)
    return destino


def exportar_excel(resultado, destino):
    val = resultado["validados"].copy()
    cols_val = (["ID_Registro", "Cedula", "Cedula_Limpia", "Fecha", "Horas",
                 "Tipo", "Aprobado_Por", "Estado", "Tarifa", "Factor", "Monto"])
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        val[cols_val].to_excel(writer, sheet_name="Registros_Validados",
                               index=False)
        resultado["por_empleado"].to_excel(
            writer, sheet_name="Liquidacion_Por_Empleado", index=False)
        resultado["por_departamento"].to_excel(
            writer, sheet_name="Por_Departamento", index=False)
        resultado["alertas"].to_excel(writer, sheet_name="Alertas", index=False)


def exportar_pdf(resultado, destino):
    res = resultado["resumen"]
    estilos = getSampleStyleSheet()
    titulo = estilos["Title"]
    titulo.textColor = colors.HexColor(AZUL)
    normal = estilos["Normal"]
    doc = SimpleDocTemplate(str(destino), pagesize=A4,
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
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(AZUL)),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#F2F4F7")]),
        ]))
        return t

    filas = [["Estado", "Registros"]]
    for e in list(ESTADOS) + ["EXCEDE TOPE"]:
        n = res["por_estado"].get(e, 0) if e in ESTADOS else \
            int((resultado["por_empleado"]["Excede_Tope"] == "Sí").sum()) \
            if not resultado["por_empleado"].empty else 0
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


# --------------------------------------------------------------------------
# Interfaz Tkinter
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Banco Horizonte - Liquidación de Horas Extra")
        self.geometry("1100x700")
        self.configure(bg=FONDO)
        self.df_maestro = None
        self.df_registros = None
        self.resultado = None
        self._aplicar_estilo()
        self._armar_ui()

    # -- estilo visual (solo apariencia, no toca logica) --------------------
    def _aplicar_estilo(self):
        self.option_add("*Font", ("Segoe UI", 10))
        estilo = ttk.Style(self)
        estilo.theme_use("clam")
        estilo.configure("Primario.TButton", background=PRIMARIO,
                         foreground="white", font=("Segoe UI", 10, "bold"),
                         padding=6, borderwidth=0)
        estilo.map("Primario.TButton", background=[("active", PRIMARIO_OSCURO)])
        estilo.configure("Liquidar.TButton", background=PRIMARIO,
                         foreground="white", font=("Segoe UI", 13, "bold"),
                         padding=10, borderwidth=0)
        estilo.map("Liquidar.TButton", background=[("active", PRIMARIO_OSCURO)])
        estilo.configure("Secundario.TButton", background=SECUNDARIO,
                         foreground=TEXTO, font=("Segoe UI", 10),
                         padding=6, borderwidth=0)
        estilo.map("Secundario.TButton", background=[("active", "#D3D9E0")])
        estilo.configure("Treeview", font=("Segoe UI", 9), rowheight=24,
                         background="white", fieldbackground="white",
                         foreground=TEXTO)
        estilo.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                         background=CABECERA_TABLA, foreground="white")
        estilo.map("Treeview", background=[("selected", PRIMARIO)],
                   foreground=[("selected", "white")])
        estilo.configure("TLabelFrame", background=TARJETA)
        estilo.configure("TLabelFrame.Label", background=TARJETA,
                         foreground=AZUL, font=("Segoe UI", 10, "bold"))

    # -- construcción ----------------------------------------------------
    def _armar_ui(self):
        cab = tk.Frame(self, bg=AZUL, height=64)
        cab.pack(fill="x")
        tk.Label(cab, text="Banco Horizonte · Liquidación de Horas Extra",
                 bg=AZUL, fg=BLANCO,
                 font=("Segoe UI", 14, "bold")).pack(pady=(10, 0))
        tk.Label(cab, text="Recursos Humanos · Costa Rica",
                 bg=AZUL, fg="#C9D4E3", font=("Segoe UI", 9)).pack(pady=(0, 8))
        tk.Frame(self, bg=DORADO, height=3).pack(fill="x")

        carga = tk.Frame(self, bg=TARJETA, padx=8, pady=6)
        carga.pack(fill="x", padx=10, pady=(8, 0))
        ttk.Button(carga, text="Cargar Maestro de Empleados",
                   style="Secundario.TButton",
                   command=self.cargar_maestro).grid(row=0, column=0, padx=4)
        self.lbl_maestro = tk.Label(carga, text="Maestro: sin cargar",
                                    bg=TARJETA, fg=TEXTO_SUAVE, anchor="w")
        self.lbl_maestro.grid(row=0, column=1, sticky="w", padx=4)
        ttk.Button(carga, text="Cargar Registros de Horas Extra",
                   style="Secundario.TButton",
                   command=self.cargar_registros).grid(row=1, column=0, padx=4,
                                                       pady=4)
        self.lbl_reg = tk.Label(carga, text="Registros: sin cargar",
                                bg=TARJETA, fg=TEXTO_SUAVE, anchor="w")
        self.lbl_reg.grid(row=1, column=1, sticky="w", padx=4)
        tk.Label(carga, text="Tope de horas extra al mes:", bg=TARJETA,
                 fg=TEXTO).grid(row=0, column=2, padx=(16, 4), sticky="e")
        self.var_tope = tk.StringVar(value="40")
        tk.Entry(carga, textvariable=self.var_tope, width=8).grid(
            row=0, column=3, sticky="w")
        tk.Label(carga, text="Departamento:", bg=TARJETA,
                 fg=TEXTO).grid(row=1, column=2, padx=(16, 4), sticky="e")
        self.var_depto = tk.StringVar(value="Todos")
        self.cmb_depto = ttk.Combobox(carga, textvariable=self.var_depto,
                                      values=["Todos"], state="readonly",
                                      width=22)
        self.cmb_depto.grid(row=1, column=3, sticky="w")
        self.cmb_depto.bind("<<ComboboxSelected>>",
                            lambda _e: self.refrescar_vista())

        btns = tk.Frame(self, bg=FONDO)
        btns.pack(fill="x", padx=10, pady=6)
        ttk.Button(btns, text="LIQUIDAR", style="Liquidar.TButton",
                   command=self.liquidar).pack(side="left", padx=4)
        ttk.Button(btns, text="Exportar (Excel + PDF)",
                   style="Primario.TButton",
                   command=self.exportar).pack(side="left", padx=4)

        self.frm_res = tk.LabelFrame(self, text="Resumen", bg=TARJETA,
                                     fg=AZUL, padx=6, pady=4)
        self.frm_res.pack(fill="x", padx=10)
        self.lbl_estados = tk.Label(self.frm_res, bg=TARJETA, fg=TEXTO,
                                    justify="left", anchor="w")
        self.lbl_estados.pack(side="left", padx=10, pady=4)
        self.lbl_totales = tk.Label(self.frm_res, bg=TARJETA, fg=TEXTO,
                                    justify="left", anchor="w",
                                    font=("Segoe UI", 10, "bold"))
        self.lbl_totales.pack(side="left", padx=10, pady=4)
        self.lbl_deptos = tk.Label(self.frm_res, bg=TARJETA, fg=TEXTO,
                                   justify="left", anchor="w")
        self.lbl_deptos.pack(side="left", padx=10, pady=4)

        mid = tk.Frame(self, bg=FONDO)
        mid.pack(fill="both", expand=True, padx=10, pady=6)
        izq = tk.LabelFrame(mid, text="Liquidación por empleado", bg=TARJETA,
                            fg=AZUL)
        izq.pack(side="left", fill="both", expand=True, padx=(0, 4))
        cols_e = ("cedula", "nombre", "depto", "hd", "hn", "ht", "monto",
                  "excede")
        self.tab_emp = ttk.Treeview(izq, columns=cols_e, show="headings",
                                    height=10)
        for c, t, w in [("cedula", "Cédula", 80), ("nombre", "Nombre", 150),
                        ("depto", "Departamento", 110),
                        ("hd", "H. diurnas", 70), ("hn", "H. nocturnas", 70),
                        ("ht", "H. totales", 70), ("monto", "Monto", 100),
                        ("excede", "Excede tope", 80)]:
            self.tab_emp.heading(c, text=t)
            self.tab_emp.column(c, width=w, anchor="center" if c not in
                                ("nombre", "depto") else "w")
        self.tab_emp.tag_configure("excede", background=ROJO_SUAVE)
        self.tab_emp.tag_configure("par", background=FILA_PAR)
        se = ttk.Scrollbar(izq, orient="vertical", command=self.tab_emp.yview)
        self.tab_emp.configure(yscrollcommand=se.set)
        self.tab_emp.pack(side="left", fill="both", expand=True)
        se.pack(side="right", fill="y")

        der = tk.LabelFrame(mid, text="Alertas (registros no válidos)",
                            bg=TARJETA, fg=AZUL)
        der.pack(side="right", fill="both", expand=True, padx=(4, 0))
        cols_a = ("id", "cedula", "fecha", "horas", "tipo", "estado")
        self.tab_ale = ttk.Treeview(der, columns=cols_a, show="headings",
                                    height=10)
        for c, t, w in [("id", "ID_Registro", 80), ("cedula", "Cédula", 80),
                        ("fecha", "Fecha", 85), ("horas", "Horas", 55),
                        ("tipo", "Tipo", 70), ("estado", "Estado", 110)]:
            self.tab_ale.heading(c, text=t)
            self.tab_ale.column(c, width=w, anchor="center")
        sa = ttk.Scrollbar(der, orient="vertical", command=self.tab_ale.yview)
        self.tab_ale.configure(yscrollcommand=sa.set)
        self.tab_ale.pack(side="left", fill="both", expand=True)
        sa.pack(side="right", fill="y")

        self.var_status = tk.StringVar(value="Cargue ambos Excel y pulse LIQUIDAR.")
        tk.Label(self, textvariable=self.var_status, bg=AZUL, fg=BLANCO,
                 anchor="w").pack(fill="x", side="bottom")

    # -- carga -----------------------------------------------------------
    def _limpieza_info(self, df):
        return int((df["Cedula"].astype(str).str.strip()
                    != df["Cedula_Limpia"]).sum())

    def cargar_maestro(self):
        path = filedialog.askopenfilename(
            title="Maestro de empleados",
            filetypes=[("Excel", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.df_maestro = cargar_maestro(path)
            n = len(self.df_maestro)
            self.lbl_maestro.config(
                text=f"Maestro: {n} filas, "
                     f"{self._limpieza_info(self.df_maestro)} con limpieza.")
            deptos = sorted(self.df_maestro["Departamento"].unique().tolist())
            self.cmb_depto.config(values=["Todos"] + deptos)
            self.var_status.set("Maestro cargado correctamente.")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo cargar el maestro:\n{exc}")
            self.var_status.set("Error al cargar el maestro.")

    def cargar_registros(self):
        path = filedialog.askopenfilename(
            title="Registros de horas extra",
            filetypes=[("Excel", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.df_registros = cargar_registros(path)
            n = len(self.df_registros)
            self.lbl_reg.config(
                text=f"Registros: {n} filas, "
                     f"{self._limpieza_info(self.df_registros)} con limpieza.")
            self.var_status.set("Registros cargados correctamente.")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo cargar registros:\n{exc}")
            self.var_status.set("Error al cargar los registros.")

    # -- liquidacion -----------------------------------------------------
    def liquidar(self):
        if self.df_maestro is None or self.df_registros is None:
            messagebox.showwarning("Faltan datos",
                                   "Cargue el maestro y los registros primero.")
            return
        try:
            tope = float(str(self.var_tope.get()).replace(",", "."))
            if tope <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Tope inválido",
                                   "El tope debe ser un número mayor que 0.")
            return
        try:
            self.resultado = aplicar_reglas(self.df_maestro,
                                            self.df_registros, tope)
            self.refrescar_vista()
            self.var_status.set(
                f"Liquidación lista: {len(self.resultado['validados'])} registros, "
                f"{fmt_crc(self.resultado['resumen']['monto_total'])} total.")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo liquidar:\n{exc}")
            self.var_status.set("Error en la liquidación.")

    def refrescar_vista(self):
        if self.resultado is None:
            return
        res = self.resultado["resumen"]
        por_emp = self.resultado["por_empleado"]
        depto = self.var_depto.get()
        if depto != "Todos" and not por_emp.empty:
            por_emp = por_emp[por_emp["Departamento"] == depto]
        conteo = {e: res["por_estado"].get(e, 0) for e in ESTADOS}
        self.lbl_estados.config(
            text="Registros por estado:\n" + "\n".join(
                f"{e}: {conteo[e]}" for e in ESTADOS))
        self.lbl_totales.config(
            text=f"Horas válidas: {fmt_horas(res['horas_total'])} "
                 f"(diurnas {fmt_horas(res['horas_diurnas'])}, "
                 f"nocturnas {fmt_horas(res['horas_nocturnas'])})\n"
                 f"Monto total: {fmt_crc(res['monto_total'])}")
        lineas = ["Monto y horas por departamento:"]
        for _, f in self.resultado["por_departamento"].iterrows():
            marca = " ◀" if f["Departamento"] == depto else ""
            lineas.append(f"{f['Departamento']}: {fmt_horas(f['Horas'])}h, "
                          f"{fmt_crc(f['Monto'])}{marca}")
        self.lbl_deptos.config(text="\n".join(lineas))

        for item in self.tab_emp.get_children():
            self.tab_emp.delete(item)
        for i, (_, f) in enumerate(por_emp.iterrows()):
            if f["Excede_Tope"] == "Sí":
                tag = ("excede",)
            else:
                tag = ("par",) if i % 2 else ()
            self.tab_emp.insert("", "end", tags=tag, values=(
                f["Cedula"], f["Nombre"], f["Departamento"],
                fmt_horas(f["Horas_Diurnas"]), fmt_horas(f["Horas_Nocturnas"]),
                fmt_horas(f["Horas_Totales"]), fmt_crc(f["Monto"]),
                f["Excede_Tope"]))
        for item in self.tab_ale.get_children():
            self.tab_ale.delete(item)
        no_val = self.resultado["no_validos"]
        if depto != "Todos":
            ced_ok = set(por_emp["Cedula"].tolist()) | set(
                self.df_maestro[self.df_maestro["Departamento"] == depto]
                ["Cedula_Limpia"].tolist())
            no_val = no_val[no_val["Cedula_Limpia"].isin(ced_ok) |
                            ~no_val["Cedula_Limpia"].isin(
                                set(self.df_maestro["Cedula_Limpia"].tolist()))]
        for _, f in no_val.iterrows():
            self.tab_ale.insert("", "end", values=(
                f["ID_Registro"], f["Cedula"], f["Fecha_Norm"],
                fmt_horas(f["Horas"]) if pd.notna(f["Horas"]) else "—",
                f["Tipo"], f["Estado"]))

    # -- exportar --------------------------------------------------------
    def exportar(self):
        if self.resultado is None:
            messagebox.showwarning("Nada que exportar",
                                   "Primero pulse LIQUIDAR.")
            return
        try:
            carpeta = ruta_reportes()
            ruta_x = carpeta / "Liquidacion_Horas_Extra.xlsx"
            ruta_p = carpeta / "Reporte_Horas_Extra.pdf"
            exportar_excel(self.resultado, ruta_x)
            exportar_pdf(self.resultado, ruta_p)
            self.var_status.set(f"Exportado en {carpeta}.")
            messagebox.showinfo("Éxito", f"Archivos generados en:\n{carpeta}")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo exportar:\n{exc}")
            self.var_status.set("Error al exportar.")


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
