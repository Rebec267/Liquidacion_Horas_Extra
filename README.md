# Liquidación de Horas Extra — Banco Horizonte (RR. HH., Costa Rica)

Aplicación de escritorio (Tkinter) para liquidar las horas extra mensuales:
valida los registros de las jefaturas, calcula el monto a pagar por persona
y genera el informe para planilla. Las reglas de negocio (R1–R8) están en
`AGENTS.md` y se aplican al pie de la letra.

## Requisitos

- Python 3.11+
- `pip install -r requirements.txt` (pandas, openpyxl, reportlab)

## Uso

```bash
python app/horas_extra_app.py
```

1. Cargar el maestro de empleados y los registros de horas extra (los
   Excel de ejemplo están en esta carpeta).
2. Ajustar el tope de horas y el departamento si hace falta.
3. Pulsar **LIQUIDAR** y luego **Exportar (Excel + PDF)**. Las salidas se
   generan en `reportes/`.

## Estructura

- `app/horas_extra_app.py` — toda la aplicación en un solo archivo.
- `Maestro_Empleados.xlsx`, `Horas_Extra_Agosto.xlsx` — insumos de ejemplo.
- `AGENTS.md` — esquema de datos y reglas del proyecto.
