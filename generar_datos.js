/* ════════════════════════════════════════════════════════════════════════
   generar_datos.js — regenera el snapshot de datos incrustado en
   Tablero_Control_Interno.html a partir de los 8 Excel de esta carpeta.

   Uso:  node generar_datos.js

   Qué hace:
   1. Lee los 8 Excel de origen con la misma lógica de extracción que usaba
      el tablero al leer en vivo (para mantener resultados idénticos).
   2. Los ajustes de inventario y las salidas se agregan (por año+mes+referencia
      y por año+mes+motivo respectivamente) porque son datasets enormes
      (~48.000 y ~101.000 filas) y ninguna vista del tablero necesita el
      detalle fila por fila — solo sumas y conteos por mes.
   3. Reemplaza el bloque "const DATA_RAW = ...;" dentro del HTML (entre los
      marcadores DATA_RAW_START/END) con el snapshot nuevo.
   4. Guarda una copia de respaldo del HTML anterior con fecha/hora.

   Cuándo correrlo: cada vez que se actualicen los Excel de esta carpeta y
   se quiera que el tablero refleje los datos nuevos (el tablero ya NO lee
   los Excel en vivo — solo muestra el snapshot generado aquí).
   ════════════════════════════════════════════════════════════════════════ */
const fs = require("fs");
const path = require("path");

const DIR = __dirname;
const HTML_PATH = path.join(DIR, "Tablero_Control_Interno.html");
const XLSX = require(path.join(DIR, "vendor", "xlsx.full.min.js"));

const FILES = {
  auditoria: "Plan Auditoria Anual 2026.xlsx",
  hallazgos: "Matriz de seguimiento de hallazgos.xlsx",
  controles: "Seguimiento controles de hallazgos.xlsx",
  riesgos:   "Matriz de Riesgos Montolivo.xlsx",
  confiab:   "REFERENCIAS MONTOLIVO.xlsx",
  ajustes:   "BASE AJUSTES.xlsx",
  salidas:   "BASE SALIDAS.xlsx",
  ventas:    "Consolidado de Ventas 2026.xlsx"
};
const RIESGOS_SHEETS_IGNORAR = ["formato", "hoja1", "hoja2"];
const MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
const MESES_UP = MESES.map(m => m.toUpperCase());

/* ───────── helpers (idénticos a los del tablero) ───────── */
function readWorkbook(filePath) {
  return XLSX.read(fs.readFileSync(filePath), { type: "buffer", cellDates: true });
}
function findSheet(wb, wantedName) {
  if (!wb) return null;
  const name = wb.SheetNames.find(n => n.trim().toLowerCase() === wantedName.trim().toLowerCase());
  return name ? wb.Sheets[name] : null;
}
function sheetRows(ws) {
  return ws ? XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) : [];
}
function headerIdx(header, name) {
  if (!header) return -1;
  const norm = s => String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase();
  const target = norm(name);
  return header.findIndex(h => norm(h) === target);
}
function parseFecha(v) {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) {
    if (isNaN(v)) return null;
    return new Date(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate());
  }
  if (typeof v === "number") {
    try {
      const d = XLSX.SSF.parse_date_code(v);
      if (!d) return null;
      return new Date(d.y, d.m - 1, d.d);
    } catch (e) { return null; }
  }
  if (typeof v === "string") {
    const s = v.trim();
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }
  return null;
}
function d2s(d) {
  return d ? d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0") : null;
}
function guessAnioFromFilename(name) {
  const m = String(name).match(/(20\d{2})/);
  return m ? +m[1] : new Date().getFullYear();
}

/* ───────── extracción: Auditoría (Cronograma Anual) ───────── */
function extractAuditoria(wb) {
  const rows = sheetRows(findSheet(wb, "Cronograma Anual"));
  const out = [];
  for (let i = 6; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r[1] == null || r[2] == null) continue;
    const fechaProg = parseFecha(r[3]);
    out.push({
      area: String(r[1]).trim(),
      mesProgramado: String(r[2]).trim(),
      fechaProgramada: d2s(fechaProg),
      tipo: r[4], responsable: r[5], alcance: r[6],
      estado: r[7] == null ? "(sin estado)" : String(r[7]).trim(),
      fechaCierre: d2s(parseFecha(r[8])),
      resultadoGeneral: r[9],
      anioSiguiente: r[10],
      observaciones: r[11],
      anio: fechaProg ? fechaProg.getFullYear() : null,
      mes: MESES.indexOf(String(r[2]).trim())
    });
  }
  return out;
}

/* ───────── extracción: Hallazgos (Matriz de Hallazgos) ───────── */
function extractHallazgos(wb) {
  const rows = sheetRows(findSheet(wb, "Matriz de Hallazgos"));
  const out = [];
  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r[0] == null) continue;
    const fechaId = parseFecha(r[1]);
    if (!fechaId) continue;
    out.push({
      id: r[0],
      fechaId: d2s(fechaId),
      area: r[2] == null ? "(sin área)" : String(r[2]).trim(),
      proceso: r[3],
      tipo: r[4] == null ? "(sin tipo)" : String(r[4]).trim(),
      descripcion: r[5], accion: r[6], responsable: r[7],
      fechaCompromiso: d2s(parseFecha(r[8])),
      fechaReal: d2s(parseFecha(r[9])),
      estado: r[10] == null ? "Abierto" : String(r[10]).trim(),
      diasAtraso: r[11],
      evidencia: r[12], observaciones: r[13],
      anio: fechaId.getFullYear(), mes: fechaId.getMonth()
    });
  }
  return out;
}
function extractHallazgosListas(wb) {
  const rows = sheetRows(findSheet(wb, "Listas"));
  const areas = [], tipos = [], estados = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    if (r[0] != null) areas.push(String(r[0]).trim());
    if (r[1] != null) tipos.push(String(r[1]).trim());
    if (r[2] != null) estados.push(String(r[2]).trim());
  }
  return { areas, tipos, estados };
}

/* ───────── extracción: Controles implementados ───────── */
function extractControles(wb) {
  const rows = sheetRows(findSheet(wb, "Seguimiento_Controles"));
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r[2] == null) continue;
    const fc = parseFecha(r[4]);
    const fi = parseFecha(r[6]);
    const fechaFiltro = fc || fi;
    out.push({
      area: r[0] == null ? "(sin área)" : String(r[0]).trim(),
      riesgo: r[1], controlPropuesto: r[2], responsable: r[3],
      fechaCompromiso: d2s(fc),
      estado: r[5] == null ? "Pendiente" : String(r[5]).trim(),
      fechaImplementacion: d2s(fi),
      anio: fechaFiltro ? fechaFiltro.getFullYear() : null, mes: fechaFiltro ? fechaFiltro.getMonth() : null
    });
  }
  return out;
}

/* ───────── extracción: Riesgos (una hoja por área) ───────── */
function extractRiesgos(wb) {
  const out = [];
  wb.SheetNames.forEach(name => {
    if (RIESGOS_SHEETS_IGNORAR.includes(name.trim().toLowerCase())) return;
    const rows = sheetRows(wb.Sheets[name]);
    for (let i = 11; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r[2] == null || typeof r[7] !== "number" || typeof r[8] !== "number") continue;
      const fo = parseFecha(r[19]);
      out.push({
        areaHoja: name.trim(),
        proceso: r[1], riesgo: r[2], evento: r[3], causa: r[4], impactoTxt: r[5],
        tipoRiesgo: r[6], probabilidad: r[7], impactoNum: r[8],
        nivelInherentePxI: r[9], nivelInherente: r[10] == null ? null : String(r[10]).trim(),
        controlesExistentes: r[11], tipoControl: r[12], frecuenciaControl: r[13],
        responsableControl: r[14], efectividadControl: r[15] == null ? null : String(r[15]).trim(),
        nivelResidual: r[16] == null ? null : String(r[16]).trim(),
        planAccion: r[17], responsablePlan: r[18],
        fechaObjetivo: d2s(fo),
        estado: r[20] == null ? "Abierto" : String(r[20]).trim(),
        anio: fo ? fo.getFullYear() : null, mes: fo ? fo.getMonth() : null
      });
    }
  });
  return out;
}

/* ───────── extracción: catálogo de referencias (REFERENCIAS MONTOLIVO) ───────── */
function extractCatalogo(wb) {
  const rows = sheetRows(findSheet(wb, "REFERENCIAS"));
  if (!rows.length) return { referencias: [], porTipo: {}, total: 0, nombres: {} };
  const header = rows[0];
  const iRef = headerIdx(header, "Referencia");
  const iTipo = headerIdx(header, "Tipo item");
  const iDesc = headerIdx(header, "Desc. item");
  const refTipo = new Map();
  const nombres = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || iRef < 0 || r[iRef] == null) continue;
    const ref = String(r[iRef]).trim().toUpperCase();
    if (!ref) continue;
    if (!refTipo.has(ref)) refTipo.set(ref, (iTipo >= 0 && r[iTipo] != null) ? String(r[iTipo]).trim() : "(sin tipo)");
    if (!nombres[ref] && iDesc >= 0 && r[iDesc] != null) nombres[ref] = String(r[iDesc]).trim();
  }
  const porTipo = {};
  refTipo.forEach(tipo => { porTipo[tipo] = (porTipo[tipo] || 0) + 1; });
  return { referencias: [...refTipo.keys()], porTipo, total: refTipo.size, nombres };
}

/* ───────── extracción agregada: ajustes de inventario (BASE AJUSTES) ─────────
   Se agrega por (año, mes, referencia) — cada fila resultante trae "count" (n° de
   movimientos reales que representa) y "costoNeto" (suma), para no perder precisión
   frente a la extracción fila-por-fila del tablero en vivo. Los desgloses por tipo de
   inventario y por bodega se calculan aparte en extractAjustesDimension — agregar esas
   dos dimensiones AQUÍ (junto con referencia) casi no reduce nada porque cada referencia
   ya vive prácticamente en una sola bodega/tipo, así que el archivo se disparó de 0.5MB
   a más de 8MB sin necesidad; separarlo lo mantiene liviano.
   La hoja se busca por nombre "GENERAL" o, si no existe, la primera que empiece con
   "AJUSTES" (el archivo la ha tenido con ambos nombres) — y las columnas se ubican
   por encabezado, no por posición fija, porque ya han cambiado de orden una vez. */
function extractAjustesAgregado(wb) {
  const rows = ajustesRows(wb);
  if (!rows.length) return [];
  const header = rows[0];
  const iRef = headerIdx(header, "Referencia");
  const iCostoNeto = headerIdx(header, "Costo neto (prom.)");
  const iFecha = headerIdx(header, "Fecha");
  const iMes = headerIdx(header, "MES");
  const map = new Map();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || iRef < 0 || r[iRef] == null) continue;
    const fecha = iFecha >= 0 ? parseFecha(r[iFecha]) : null;
    const mesNum = iMes >= 0 ? r[iMes] : null;
    const mes = mesNum != null && !isNaN(+mesNum) ? (+mesNum - 1) : (fecha ? fecha.getMonth() : null);
    const anio = fecha ? fecha.getFullYear() : null;
    const referencia = String(r[iRef]).trim().toUpperCase();
    const k = anio + "|" + mes + "|" + referencia;
    if (!map.has(k)) map.set(k, { anio, mes, referencia, count: 0, costoNeto: 0, costoPositivo: 0, costoNegativo: 0 });
    const o = map.get(k);
    o.count++;
    const c = (iCostoNeto >= 0 && typeof r[iCostoNeto] === "number") ? r[iCostoNeto] : 0;
    o.costoNeto += c;
    // Positivo y negativo se acumulan por movimiento individual (no por el signo del agregado):
    // dos movimientos que se compensan entre sí dentro del mismo mes/referencia (ej. +5M y -3M)
    // deben seguir contando como +5M de sobrante y -3M de faltante, no netearse a +2M antes de clasificar.
    if (c > 0) o.costoPositivo += c; else if (c < 0) o.costoNegativo += c;
  }
  return [...map.values()];
}
/* Ubica y devuelve las filas de la hoja de ajustes (misma búsqueda de hoja que extractAjustesAgregado). */
function ajustesRows(wb) {
  let ws = findSheet(wb, "GENERAL");
  if (!ws) {
    const name = wb.SheetNames.find(n => /^AJUSTES/i.test(n.trim()));
    if (name) ws = wb.Sheets[name];
  }
  return sheetRows(ws);
}
/* Desglose de ajustes por una dimensión adicional (tipo de inventario, bodega) — agregado
   solo por (año, mes, valor de la dimensión), sin referencia, para que quede compacto
   (decenas de filas, no miles) sin importar cuántas referencias distintas haya. */
function extractAjustesDimension(wb, columnaHeader) {
  const rows = ajustesRows(wb);
  if (!rows.length) return [];
  const header = rows[0];
  const iDim = headerIdx(header, columnaHeader);
  const iCostoNeto = headerIdx(header, "Costo neto (prom.)");
  const iFecha = headerIdx(header, "Fecha");
  const iMes = headerIdx(header, "MES");
  const map = new Map();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const fecha = iFecha >= 0 ? parseFecha(r[iFecha]) : null;
    const mesNum = iMes >= 0 ? r[iMes] : null;
    const mes = mesNum != null && !isNaN(+mesNum) ? (+mesNum - 1) : (fecha ? fecha.getMonth() : null);
    const anio = fecha ? fecha.getFullYear() : null;
    const valor = (iDim >= 0 && r[iDim] != null) ? String(r[iDim]).trim() : "(sin dato)";
    const k = anio + "|" + mes + "|" + valor;
    if (!map.has(k)) map.set(k, { anio, mes, valor, count: 0, costoNeto: 0, costoPositivo: 0, costoNegativo: 0 });
    const o = map.get(k);
    o.count++;
    const c = (iCostoNeto >= 0 && typeof r[iCostoNeto] === "number") ? r[iCostoNeto] : 0;
    o.costoNeto += c;
    if (c > 0) o.costoPositivo += c; else if (c < 0) o.costoNegativo += c;
  }
  return [...map.values()];
}

/* ───────── extracción agregada: salidas de inventario (BASE SALIDAS › SALIDAS ####) ─────────
   Se agrega por (año, mes, motivo) — la referencia y el detalle por transacción no
   se usan en ninguna vista del tablero, solo el total de $ salidas por mes/motivo. */
function extractSalidasAgregado(wb) {
  const ws = wb.Sheets[wb.SheetNames.find(n => /^SALIDAS/i.test(n.trim()))] || null;
  const rows = sheetRows(ws);
  if (!rows.length) return [];
  const header = rows[0];
  const iMes = headerIdx(header, "Mes"), iAnio = headerIdx(header, "Año"),
        iMotivo = headerIdx(header, "Desc. motivo"), iRef = headerIdx(header, "Referencia"),
        iCosto = headerIdx(header, "Costo salidas (prom.)");
  const map = new Map();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || iRef < 0 || r[iRef] == null) continue;
    const mesNum = iMes >= 0 ? r[iMes] : null, anioNum = iAnio >= 0 ? r[iAnio] : null;
    const anio = anioNum != null && !isNaN(+anioNum) ? +anioNum : null;
    const mes = mesNum != null && !isNaN(+mesNum) ? (+mesNum - 1) : null;
    const motivo = (iMotivo >= 0 && r[iMotivo] != null) ? String(r[iMotivo]).trim() : "(sin motivo)";
    const k = anio + "|" + mes + "|" + motivo;
    if (!map.has(k)) map.set(k, { anio, mes, motivo, costoSalida: 0 });
    map.get(k).costoSalida += (iCosto >= 0 && typeof r[iCosto] === "number") ? r[iCosto] : 0;
  }
  return [...map.values()];
}

/* ───────── extracción: ventas (Consolidado de Ventas, hojas por mes) ───────── */
function extractVentas(wb, anioDefault) {
  const out = [];
  wb.SheetNames.forEach(name => {
    const idx = MESES_UP.indexOf(name.trim().toUpperCase());
    if (idx < 0) return;
    const rows = sheetRows(wb.Sheets[name]);
    let total = 0;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r[1] == null) continue;
      if (typeof r[4] === "number") total += r[4];
    }
    out.push({ mes: idx, anio: anioDefault, total });
  });
  return out;
}

/* ───────── construcción del snapshot ───────── */
function build() {
  const missing = [];
  const wbs = {};
  for (const key of Object.keys(FILES)) {
    const fp = path.join(DIR, FILES[key]);
    if (!fs.existsSync(fp)) { missing.push(FILES[key]); continue; }
    wbs[key] = readWorkbook(fp);
  }
  if (missing.length) {
    console.warn("ADVERTENCIA — no se encontraron estos archivos (se omiten del snapshot):");
    missing.forEach(m => console.warn("  -", m));
  }

  const ajustesAgregado = wbs.ajustes ? extractAjustesAgregado(wbs.ajustes) : [];
  const ajustesPorTipo = wbs.ajustes ? extractAjustesDimension(wbs.ajustes, "Desc. tipo inventario") : [];
  const ajustesPorBodega = wbs.ajustes ? extractAjustesDimension(wbs.ajustes, "Desc. Bodega") : [];

  const out = {
    generadoEn: new Date().toLocaleString("es-CO", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }),
    auditoria: { records: wbs.auditoria ? extractAuditoria(wbs.auditoria) : [] },
    hallazgos: {
      records: wbs.hallazgos ? extractHallazgos(wbs.hallazgos) : [],
      listas: wbs.hallazgos ? extractHallazgosListas(wbs.hallazgos) : { areas: [], tipos: [], estados: [] }
    },
    controles: { records: wbs.controles ? extractControles(wbs.controles) : [] },
    riesgos: { records: wbs.riesgos ? extractRiesgos(wbs.riesgos) : [] },
    confiabilidad: {
      catalogo: wbs.confiab ? extractCatalogo(wbs.confiab) : { referencias: [], porTipo: {}, total: 0 },
      ajustes: ajustesAgregado,
      ajustesPorTipo, ajustesPorBodega
    },
    costos: {
      salidas: wbs.salidas ? extractSalidasAgregado(wbs.salidas) : [],
      ventas: wbs.ventas ? extractVentas(wbs.ventas, guessAnioFromFilename(FILES.ventas)) : []
    }
  };
  return out;
}

function main() {
  const data = build();
  const json = JSON.stringify(data);
  console.log("Snapshot generado:");
  console.log("  auditoria:", data.auditoria.records.length, "registros");
  console.log("  hallazgos:", data.hallazgos.records.length, "registros");
  console.log("  controles:", data.controles.records.length, "registros");
  console.log("  riesgos:", data.riesgos.records.length, "registros");
  console.log("  catálogo confiabilidad:", data.confiabilidad.catalogo.total, "referencias");
  console.log("  ajustes (agregados):", data.confiabilidad.ajustes.length, "filas");
  console.log("  salidas (agregadas):", data.costos.salidas.length, "filas");
  console.log("  ventas:", data.costos.ventas.length, "meses");
  console.log("  tamaño JSON:", (json.length / 1e6).toFixed(2), "MB");

  let html = fs.readFileSync(HTML_PATH, "utf8");
  const startMarker = "/* ==DATA_RAW_START== NO EDITAR A MANO — generado por generar_datos.js */";
  const endMarker = "/* ==DATA_RAW_END== */";
  const startIdx = html.indexOf(startMarker);
  const endIdx = html.indexOf(endMarker);
  if (startIdx === -1 || endIdx === -1 || endIdx < startIdx) {
    throw new Error("No se encontraron los marcadores DATA_RAW_START/END en el HTML — revisa que no se hayan borrado.");
  }

  // Respaldo con fecha/hora antes de sobrescribir, igual que el flujo de Gestión Humana.
  const stamp = new Date().toISOString().replace(/[:T]/g, "-").slice(0, 16);
  const backupPath = path.join(DIR, `Tablero_Control_Interno (backup ${stamp}).html`);
  fs.copyFileSync(HTML_PATH, backupPath);
  console.log("Respaldo guardado en:", path.basename(backupPath));

  const before = html.slice(0, startIdx);
  const after = html.slice(endIdx + endMarker.length);
  const newBlock = startMarker + "\nconst DATA_RAW = " + json + ";\n" + endMarker;
  html = before + newBlock + after;

  fs.writeFileSync(HTML_PATH, html, "utf8");
  console.log("Tablero_Control_Interno.html actualizado con el snapshot nuevo.");
}

main();
