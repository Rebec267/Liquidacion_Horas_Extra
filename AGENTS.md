# AGENTS.md – Liquidación horas extra (Banco Horizonte, Costa Rica, RR.HH.)

Liquidar horas extra mensuales: validar registros de jefaturas, calcular monto por persona e informe para planilla.

## Archivos (verificado por XML interno del xlsx)

- `Maestro_Empleados.xlsx` → hoja `Empleados`, rango `A1:E41` (header + 40 filas).
  Columnas: `Cedula | Nombre | Departamento | Salario_Mensual_CRC | Tope_Horas_Extra_Mes`.
  `Tope_Horas_Extra_Mes` = 40 en las 40 filas. Deptos: Contabilidad, Operaciones, Seguridad, Servicio al Cliente, Tecnología.
- `Horas_Extra_Agosto.xlsx` → hoja `Registros`, rango `A1:F112` (header + 111 filas, `ID_Registro` únicos HE-0001…HE-0111, sin duplicados).
  Columnas: `ID_Registro | Cedula | Fecha | Horas | Tipo | Aprobado_Por`.
  `Fecha` = serial Excel 46235–46265 = 2026-08-01 a 2026-08-31. `Tipo` solo `Diurna|Nocturna`. `Aprobado_Por` solo `jefatura.{operaciones,servicio,contabilidad,tecnologia,seguridad}`.

## Clave `Cedula`: normalizar antes de cruzar

- Hacer `TRIM + quitar "-"` en ambos archivos. Sin esto el join falla.
- Casos verificados: maestro `" 4174679 "` (con espacios) ↔ horas `"4174679"`; horas `"7-6245-87"` ↔ maestro `"7624587"`; maestro `"7-8107-49"` → normalizado `7810749`.
- Tras normalizar: única cédula huérfana en horas = `999111222` (no existe en maestro). 7 cédulas del maestro sin horas: `2857744, 2987495, 4718180, 5975701, 6821501, 7724436, 7810749`.

## Validaciones necesarias (datos sucios intencionales)

- `Horas` tiene fracciones (`1.5`) y valores a rechazar/marcar: fila 67 `HE-0107` = 14h, fila 75 `HE-0108` = 0h, fila 88 `HE-0109` = -2h. Rango normal 1–6.
- Nombres y `Tecnología` vienen con mojibake de tildes en `sharedStrings.xml` (ej. `Tecnolog��a`). Normalizar texto antes de agrupar por departamento/empleado.

## Reglas de negocio (obligatorias, aplicar al pie de la letra)

- R1 Limpieza: `Cedula_Limpia` = `Cedula` sin guiones y sin espacios inicio/fin. Nunca modificar la columna original.
- R2 SIN EMPLEADO: `Cedula_Limpia` no existe en maestro. No se paga; se reporta.
- R3 HORAS INVALIDAS: `Horas <= 0` o `Horas > 12`. No se paga; se reporta.
- R4 DUPLICADO: misma `Cedula_Limpia` + misma `Fecha` + mismo `Tipo` + mismas `Horas` que un registro anterior (menor `ID_Registro`). Solo se paga el primero. Evaluar R2 y R3 antes que R4.
- R5 Tarifa hora ordinaria = `Salario_Mensual_CRC / 240`.
- R6 Registro VALIDO: `Monto = Horas * Tarifa * Factor`; `Factor = 1.5` si `Diurna`, `2.0` si `Nocturna`.
- R7 EXCEDE TOPE: suma de `Horas` válidas por empleado > `Tope_Horas_Extra_Mes` (40). Se pagan todas; alerta con nombre, departamento y horas.
- R8 Informe: monto total a pagar, monto y horas por departamento, y top-5 empleados con más horas válidas.

## Salidas (`reportes/`)

- `Liquidacion_Horas_Extra.xlsx` con hojas: `Registros_Validados` (todos los registros + `Cedula_Limpia, Estado, Tarifa, Factor, Monto`), `Liquidacion_Por_Empleado` (cédula, nombre, departamento, horas diurnas/nocturnas/totales, monto, excede tope), `Por_Departamento`, `Alertas`.
- `Reporte_Horas_Extra.pdf` con resumen ejecutivo y alertas.

## Tecnología objetivo

- Python 3.11+, Tkinter, `pandas + openpyxl`, `reportlab`. Toda la app en UN solo archivo: `app/horas_extra_app.py`. No usar otras librerías.

## Entorno

- En esta máquina no hay `python`/`node` utilizables (stubs a Microsoft Store). Para inspeccionar xlsx usar PowerShell + `System.IO.Compression.ZipFile` sobre `xl/sharedStrings.xml` y `xl/worksheets/sheet1.xml`; fechas = días desde `1899-12-30`.
