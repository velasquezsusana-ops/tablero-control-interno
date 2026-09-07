# -*- coding: utf-8 -*-
"""
Extrae los datos de la hoja "Indicadores" de Control interno.xlsx y genera
control_interno_data.json, que luego se inyecta en Tablero_Control_Interno.html
(bloque `const CI = {...}` dentro del <script>).

Volver a correr este script (con el mismo interprete de Python usado en las
otras sesiones: C:\\Users\\Montolivo\\AppData\\Local\\Programs\\Python\\Python313\\python.exe)
cada vez que se actualice Control interno.xlsx, y luego pedirle a Claude que
vuelva a "splicear" control_interno_data.json dentro del HTML.
"""
import json
import openpyxl
import os

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(BASE, "Control interno.xlsx")

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb["Indicadores"]


def col(cell_range):
    """cell_range: e.g. 'X4:X16' -> list of values in that column range."""
    vals = []
    for row in ws[cell_range]:
        for c in row:
            vals.append(c.value)
    return vals


def clean_labels(labels):
    out = []
    for l in labels:
        if l is None or str(l).strip() in ("", " "):
            break
        out.append(str(l).strip())
    return out


MESES_ABR = {
    "Enero": "Ene", "Febrero": "Feb", "Marzo": "Mar", "Abril": "Abr",
    "Mayo": "May", "Junio": "Jun", "Julio": "Jul", "Agosto": "Ago",
    "Septiembre": "Sep", "Octubre": "Oct", "Noviembre": "Nov", "Diciembre": "Dic",
}


def short_label(full):
    # "Diciembre 2024" -> "Dic 24"
    parts = full.split(" ")
    mes = MESES_ABR.get(parts[0], parts[0][:3])
    anio = parts[1][-2:] if len(parts) > 1 else ""
    return f"{mes} {anio}"


def series(row_start, row_end, cols):
    """
    row_start/row_end: 1-based inclusive Excel rows for the data (X = mes label col)
    cols: dict name -> column letter
    Returns dict with 'meses' (full), 'mesesCorto', plus each series as list of numbers
    (None stays None; '' becomes None too).
    """
    meses_full = col(f"X{row_start}:X{row_end}")
    meses_full = clean_labels(meses_full)
    n = len(meses_full)
    out = {"meses": meses_full, "mesesCorto": [short_label(m) for m in meses_full]}
    for name, letter in cols.items():
        vals = col(f"{letter}{row_start}:{letter}{row_end}")[:n]
        clean = []
        for v in vals:
            if v is None or v == "":
                clean.append(None)
            else:
                clean.append(v)
        out[name] = clean
    return out


data = {}

# 1. Cumplimiento del plan de auditoria (filas 4-16)
data["auditoria"] = series(4, 16, {
    "registros": "AA", "errores": "AB", "exactitud": "AC", "objetivo": "AD",
})

# 2. N de hallazgos (filas 24-36)
data["hallazgos"] = series(24, 36, {
    "criticos": "AA", "significativos": "AB", "total": "AC",
})

# 3. Cierre de hallazgos (filas 43-55)
data["cierre"] = series(43, 55, {
    "conPlan": "AA", "cerrados": "AB", "pctCierre": "AC", "objetivo": "AD",
})

# 4. Tiempo de respuesta (filas 62-74)
data["tiempoResp"] = series(62, 74, {
    "cantidad": "AA", "promDias": "AB", "objetivo": "AC",
})

# 5. Controles implementados (filas 81-93)
data["controles"] = series(81, 93, {
    "planificados": "AA", "implementados": "AB", "pct": "AC", "objetivo": "AD",
})

# 6. Matriz de riesgos - actualizacion (filas 99-111)
data["matrizRiesgos"] = series(99, 111, {
    "programados": "AA", "actualizados": "AB", "pct": "AC", "objetivo": "AD",
})

# 7. Mitigacion de riesgos (filas 115-127)
data["mitigacion"] = series(115, 127, {
    "programadas": "AA", "ejecutadas": "AB", "pct": "AC", "objetivo": "AD",
})

# 8. Confiabilidad de inventario (fila de meses 196, cols Y:AK; datos filas 197-202)
meses_conf_full = col("Y196:AK196")
meses_conf_full = [str(m).strip() for m in meses_conf_full if m]
def rowvals(r, ncols):
    vals = col(f"Y{r}:AK{r}")[:ncols]
    return [None if (v is None or v == "") else v for v in vals]

n_conf = len(meses_conf_full)
data["confiabilidad"] = {
    "meses": meses_conf_full,
    "mesesCorto": [short_label(m) for m in meses_conf_full],
    "periodoBase": rowvals(197, n_conf),
    "objetivo1": rowvals(198, n_conf),
    "objetivo2": rowvals(199, n_conf),
    "objetivo3": rowvals(200, n_conf),
    "pct": rowvals(201, n_conf),
}

# 9. Costos de inventario (filas 214-226)
data["costos"] = series(214, 226, {
    "impacto": "AA", "objetivo": "AB",
})


# ─────────────────────────────────────────────────────────────────
# DETALLE: datos "tal cual" de cada una de las 9 hojas fuente (no el
# resumen recortado de 13 meses que usa Indicadores, sino el historico
# completo, en general Enero 2024 - Diciembre 2025 = 24 meses).
# ─────────────────────────────────────────────────────────────────

def find_sheet(prefix):
    for n in wb.sheetnames:
        if n.startswith(prefix):
            return wb[n]
    raise KeyError(prefix)


def jval(v):
    if hasattr(v, "isoformat"):
        return v.isoformat()[:10]
    if v == "":
        return None
    return v


def rows_until_empty(sheet, start_row, key_col, cols):
    """cols: list of 1-based column indices to pull. Stops when key_col is empty."""
    out = []
    r = start_row
    while True:
        key = sheet.cell(row=r, column=key_col).value
        if key is None or key == "":
            break
        out.append([jval(sheet.cell(row=r, column=c).value) for c in cols])
        r += 1
    return out


detalle = {}

ws_a = find_sheet("Cumplimiento del plan de audi")
detalle["auditoria"] = {
    "headers": ["Año", "Mes", "Auditorías programadas", "Auditorías realizadas", "Exactitud (%)", "Objetivo"],
    "rows": rows_until_empty(ws_a, 2, 1, [1, 2, 3, 4, 5, 6]),
}

ws_h = find_sheet("N de hallazgos")
detalle["hallazgos"] = {
    "headers": ["Año", "Mes", "Hallazgos críticos", "Hallazgos significativos", "Total"],
    "rows": rows_until_empty(ws_h, 2, 1, [1, 2, 3, 4, 5]),
}

ws_c = find_sheet("Cierre de hallazgos")
detalle["cierre"] = {
    "headers": ["Año", "Mes", "Hallazgos con plan de acción", "Hallazgos cerrados", "% cierre", "Objetivo"],
    "rows": rows_until_empty(ws_c, 2, 1, [1, 2, 3, 4, 5, 6]),
}

ws_t = find_sheet("Tiempo de respuesta")
detalle["tiempoRespDetalle"] = {
    "headers": ["Año", "Mes", "Hallazgo", "Fecha detección", "Fecha cierre", "Días de cierre"],
    "rows": rows_until_empty(ws_t, 3, 1, [1, 2, 3, 4, 5, 6]),
}
detalle["tiempoRespResumen"] = {
    "headers": ["Año", "Mes", "Cantidad de hallazgos", "Promedio de días", "Objetivo"],
    "rows": rows_until_empty(ws_t, 3, 9, [9, 10, 11, 12, 13]),
}

ws_k = find_sheet("Controles implementados")
detalle["controles"] = {
    "headers": ["Año", "Mes", "Controles planificados", "Controles implementados", "% Cumplimiento", "Objetivo"],
    "rows": rows_until_empty(ws_k, 2, 1, [1, 2, 3, 4, 5, 6]),
}

ws_m = find_sheet("Matriz de riesgos")
detalle["matrizRiesgos"] = {
    "headers": ["Año", "Mes", "Riesgos programados", "Riesgos actualizados", "% Cumplimiento", "Objetivo"],
    "rows": rows_until_empty(ws_m, 2, 1, [1, 2, 3, 4, 5, 6]),
}

ws_g = find_sheet("Cumplimiento mitigaci")
detalle["mitigacion"] = {
    "headers": ["Año", "Mes", "Acciones programadas", "Acciones ejecutadas", "% Cumplimiento", "Objetivo"],
    "rows": rows_until_empty(ws_g, 2, 1, [1, 2, 3, 4, 5, 6]),
}

ws_i = find_sheet("Confiabilidad de inventario")
detalle["confiabilidad"] = {
    "headers": ["Año", "Mes", "Total Ref", "Sin Diferencia", "Con Diferencia", "% de cumplimiento"],
    "rows": rows_until_empty(ws_i, 2, 1, [1, 2, 4, 5, 6, 7]),
}

ws_o = find_sheet("Costos de inventario")
detalle["costosDetalle"] = {
    "headers": ["Año", "Mes", "Referencia", "Sistema", "Costo Unitario", "Físico", "Diferencia", "Tipo", "Impacto"],
    "rows": rows_until_empty(ws_o, 2, 1, [1, 2, 3, 4, 5, 6, 7, 8, 9]),
}
detalle["costosResumen"] = {
    "headers": ["Año", "Mes", "Impacto", "Objetivo"],
    "rows": rows_until_empty(ws_o, 2, 13, [13, 14, 15, 16]),
}

data["detalle"] = detalle

out_path = os.path.join(BASE, "control_interno_data.json")
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("OK ->", out_path)
for k, v in data.items():
    if k == "detalle":
        continue
    n = len(v.get("meses", []))
    print(f"  {k}: {n} meses")
for k, v in detalle.items():
    print(f"  detalle.{k}: {len(v['rows'])} filas")
