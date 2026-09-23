"""Banco Horizonte (Costa Rica) - Liquidacion de Horas Extra.

App de escritorio (Tkinter) en UN solo archivo. Reglas R1-R8 segun AGENTS.md.
Uso:  python app/horas_extra_app.py
Dependencias: pandas, openpyxl, reportlab, tkinter (stdlib).
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from logica_liquidacion import (
    AZUL, DORADO, BLANCO, ROJO_SUAVE, FONDO, TARJETA, TEXTO, TEXTO_SUAVE,
    PRIMARIO, PRIMARIO_OSCURO, SECUNDARIO, CABECERA_TABLA, FILA_PAR, ESTADOS,
    fmt_crc, fmt_horas, cargar_maestro, cargar_registros, aplicar_reglas,
    exportar_excel, exportar_pdf, ruta_reportes,
)

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
