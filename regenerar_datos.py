# Regenera los 3 bloques de datos de Tablero_Control_Interno.html desde los Excel:
#   1) DATA_RAW            (port de generar_datos.js)
#   2) data-blob-confiabilidad (port de ReliabilityEngine.parseAjustesSheet/parseCatalogoSheet/buildPayload)
#   3) window.__CS_DATA__   (port de tryParseSalidas/Ventas/AjustesWorkbook, extendiendo diccionarios existentes)
# Uso: python regenerar_datos.py "<carpeta>" [--validar]   (--validar: solo meses 1-8, compara con el HTML actual, no escribe)
import sys, os, re, json, math, shutil, pickle, datetime as dt, warnings
import openpyxl
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")

DIR = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else os.path.dirname(os.path.abspath(__file__))
VALIDAR = "--validar" in sys.argv
MAXMES = 8 if VALIDAR else 12
HTML = os.path.join(DIR, "Tablero_Control_Interno.html")
import tempfile
CACHE = os.path.join(tempfile.gettempdir(), "montolivo_ci_cache_excel.pkl")
F = {
    "auditoria": "Plan Auditoria Anual 2026.xlsx",
    "hallazgos": "Matriz de seguimiento de hallazgos 2026.xlsx",
    "controles": "Tablero de Controles 2026.xlsx",
    "riesgos": "Matriz de Riesgos Montolivo - Auditoria 2026.xlsx",
    "confiab": "REFERENCIAS MONTOLIVO.xlsx",
    "ajustes": "BASE AJUSTES.xlsx",
    "salidas": "BASE SALIDAS.xlsx",
    "ventas": "Consolidado de Ventas 2026.xlsx",
}
MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
MESES_UP = [m.upper() for m in MESES]

# ───────── lectura (cacheada por mtime) ─────────
def read_all():
    key = {k: os.path.getmtime(os.path.join(DIR, f)) for k, f in F.items()}
    if os.path.exists(CACHE):
        c = pickle.load(open(CACHE, "rb"))
        if c["key"] == key: return c["wbs"]
    wbs = {}
    for k, f in F.items():
        wb = openpyxl.load_workbook(os.path.join(DIR, f), read_only=True, data_only=True)
        sheets = {}
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows = []
            for r in ws.iter_rows(min_row=1, min_col=1, values_only=True):
                rows.append(list(r))
            # SheetJS: rango empieza en la primera celda usada; openpyxl desde A1 (todas las hojas usadas aquí empiezan en A1)
            while rows and all(v is None for v in rows[-1]): rows.pop()
            sheets[sn] = rows
        wbs[k] = (wb.sheetnames, sheets)
        print("leído", f, flush=True)
    pickle.dump({"key": key, "wbs": wbs}, open(CACHE, "wb"))
    return wbs

def jsnum(v):  # Number(v) || 0
    if v is None or isinstance(v, bool): return float(v or 0) if isinstance(v, bool) else 0
    if isinstance(v, (int, float)): return 0 if (isinstance(v, float) and math.isnan(v)) else v
    if isinstance(v, str):
        s = v.strip()
        if s == "": return 0
        try: return float(s)
        except ValueError: return 0
    return 0
def jsstr(v):
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v)
def isnum(v): return isinstance(v, (int, float)) and not isinstance(v, bool)
def truthy(v): return not (v is None or v == "" or v == 0 or v is False)
def round2(v): return math.floor(v * 100 + 0.5) / 100
def g(r, i): return r[i] if 0 <= i < len(r) else None
def norm(h): return "" if h is None else jsstr(h).strip()

def month_ok(m): return m is not None and m <= MAXMES

# ═════════════ 1) DATA_RAW (generar_datos.js) ═════════════
def parse_fecha(v):
    if v is None or v == "": return None
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isnum(v):
        try: return (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(v))).date()
        except Exception: return None
    if isinstance(v, str):
        s = v.strip()
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
        if m: return dt.date(int(m[3]), int(m[2]), int(m[1]))
        m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
        if m: return dt.date(int(m[1]), int(m[2]), int(m[3]))
    return None
def d2s(d): return d.strftime("%Y-%m-%d") if d else None
def cv(v):
    if isinstance(v, float) and v.is_integer(): return int(v)
    if isinstance(v, (dt.datetime, dt.date)): return v.strftime("%Y-%m-%dT05:00:16.000Z")
    if isinstance(v, dt.time): return str(v)
    return v
def S(v): return jsstr(v).strip()
def find_sheet(wbt, name):
    names, sheets = wbt
    for n in names:
        if n.strip().lower() == name.strip().lower(): return sheets[n]
    return []
def pad(r): return list(r) + [None] * 30
def hidx(header, name):
    nz = lambda s: re.sub(r"\s+", " ", "" if s is None else jsstr(s)).strip().lower()
    t = nz(name)
    for i, h in enumerate(header or []):
        if nz(h) == t: return i
    return -1

def ext_auditoria(w):
    out = []
    for r in map(pad, find_sheet(w, "Cronograma Anual")[6:]):
        if r[1] is None or r[2] is None: continue
        fp = parse_fecha(r[3])
        out.append({"area": S(r[1]), "mesProgramado": S(r[2]), "fechaProgramada": d2s(fp), "tipo": cv(r[4]),
            "responsable": cv(r[5]), "alcance": cv(r[6]), "estado": "(sin estado)" if r[7] is None else S(r[7]),
            "fechaCierre": d2s(parse_fecha(r[8])), "resultadoGeneral": cv(r[9]), "anioSiguiente": cv(r[10]),
            "observaciones": cv(r[11]), "anio": fp.year if fp else None,
            "mes": MESES.index(S(r[2])) if S(r[2]) in MESES else -1})
    return out
def ext_hallazgos(w):
    out = []
    for r in map(pad, find_sheet(w, "Matriz de Hallazgos")[2:]):
        if r[0] is None: continue
        f = parse_fecha(r[1])
        if not f: continue
        out.append({"id": cv(r[0]), "fechaId": d2s(f), "area": "(sin área)" if r[2] is None else S(r[2]),
            "proceso": cv(r[3]), "tipo": "(sin tipo)" if r[4] is None else S(r[4]), "descripcion": cv(r[5]),
            "accion": cv(r[6]), "responsable": cv(r[7]), "fechaCompromiso": d2s(parse_fecha(r[8])),
            "fechaReal": d2s(parse_fecha(r[9])), "estado": "Abierto" if r[10] is None else S(r[10]),
            "diasAtraso": cv(r[11]), "evidencia": cv(r[12]), "observaciones": cv(r[13]),
            "anio": f.year, "mes": f.month - 1})
    return out
def ext_listas(w):
    a, t, e = [], [], []
    for r in map(pad, find_sheet(w, "Listas")[1:]):
        if r[0] is not None: a.append(S(r[0]))
        if r[1] is not None: t.append(S(r[1]))
        if r[2] is not None: e.append(S(r[2]))
    return {"areas": a, "tipos": t, "estados": e}
def ext_controles(w):
    # "Tablero de Controles 2026.xlsx" — hoja Seguimiento_Controles: una fila por control con el
    # resultado mensual (C / P / NC / NA) en 12 columnas de mes; parámetros en la hoja Listas.
    rows = find_sheet(w, "Seguimiento_Controles")
    nz = lambda v: re.sub(r"\s+", " ", "" if v is None else jsstr(v)).strip().lower()
    hi = next((i for i, r in enumerate(rows) if r and nz(r[0]) == "id"), None)
    out, meses = [], []
    if hi is not None:
        H = pad(rows[hi])
        col = lambda name: next((i for i, h in enumerate(H) if nz(h) == name), -1)
        mcols = [(i, parse_fecha(h)) for i, h in enumerate(H) if isinstance(h, (dt.datetime, dt.date)) or (i >= 13 and parse_fecha(h))]
        mcols = [(i, d) for i, d in mcols if d][:12]
        meses = [{"anio": d.year, "mes": d.month - 1} for _, d in mcols]
        iObs = col("observaciones de la última revisión")
        for r in map(pad, rows[hi + 1:]):
            if r[0] is None or r[1] is None: continue
            res = []
            for (ci, d) in mcols:
                v = r[ci]
                if v is None or S(v) == "": continue
                res.append({"anio": d.year, "mes": d.month - 1, "valor": S(v).upper()})
            out.append({"id": S(r[0]), "area": S(r[1]), "ref": cv(r[2]), "control": cv(r[3]), "objetivo": cv(r[4]),
                "prueba": cv(r[5]), "criterio": cv(r[6]), "evidencia": cv(r[7]), "tipoPrueba": None if r[8] is None else S(r[8]),
                "frecuencia": None if r[9] is None else S(r[9]), "prioridad": None if r[10] is None else S(r[10]),
                "responsable": cv(r[11]), "responsableAuditoria": cv(r[12]),
                "observaciones": cv(r[iObs]) if iObs >= 0 else None, "resultados": res})
    params = {"inicio": d2s(dt.date(meses[0]["anio"], meses[0]["mes"] + 1, 1)) if meses else None,
              "parcial": 0.5, "meta": 0.9, "umbral": 0.7, "frecuencias": {}}
    for r in map(pad, find_sheet(w, "Listas")[1:]):
        k = nz(r[9]); v = r[10]
        if r[4] is not None and isnum(r[5]): params["frecuencias"][S(r[4])] = r[5]
        if not k or v is None: continue
        if "inicio del periodo" in k and parse_fecha(v): params["inicio"] = d2s(parse_fecha(v))
        elif "parcial pondera" in k and isnum(v): params["parcial"] = v
        elif "meta de cumplimiento" in k and isnum(v): params["meta"] = v
        elif "umbral de alerta" in k and isnum(v): params["umbral"] = v
    return {"records": out, "meses": meses, "params": params}
def ext_riesgos(w):
    names, sheets = w; out = []
    for name in names:
        if name.strip().lower() in ("formato", "hoja1", "hoja2"): continue
        for r in map(pad, sheets[name][11:]):
            if r[2] is None or not isnum(r[7]) or not isnum(r[8]): continue
            fo = parse_fecha(r[19])
            out.append({"areaHoja": name.strip(), "proceso": cv(r[1]), "riesgo": cv(r[2]), "evento": cv(r[3]),
                "causa": cv(r[4]), "impactoTxt": cv(r[5]), "tipoRiesgo": cv(r[6]), "probabilidad": cv(r[7]),
                "impactoNum": cv(r[8]), "nivelInherentePxI": cv(r[9]), "nivelInherente": None if r[10] is None else S(r[10]),
                "controlesExistentes": cv(r[11]), "tipoControl": cv(r[12]), "frecuenciaControl": cv(r[13]),
                "responsableControl": cv(r[14]), "efectividadControl": None if r[15] is None else S(r[15]),
                "nivelResidual": None if r[16] is None else S(r[16]), "planAccion": cv(r[17]), "responsablePlan": cv(r[18]),
                "fechaObjetivo": d2s(fo), "estado": "Abierto" if r[20] is None else S(r[20]),
                "anio": fo.year if fo else None, "mes": fo.month - 1 if fo else None})
    return out
def ext_catalogo(w):
    rows = find_sheet(w, "REFERENCIAS")
    h = rows[0]; iR, iT, iD = hidx(h, "Referencia"), hidx(h, "Tipo item"), hidx(h, "Desc. item")
    ref_tipo, nombres = {}, {}
    for r in rows[1:]:
        if iR < 0 or g(r, iR) is None: continue
        ref = S(r[iR]).upper()
        if not ref: continue
        if ref not in ref_tipo: ref_tipo[ref] = S(r[iT]) if (iT >= 0 and g(r, iT) is not None) else "(sin tipo)"
        if ref not in nombres and iD >= 0 and g(r, iD) is not None: nombres[ref] = S(r[iD])
    por = {}
    for t in ref_tipo.values(): por[t] = por.get(t, 0) + 1
    return {"referencias": list(ref_tipo), "porTipo": por, "total": len(ref_tipo), "nombres": nombres}

def ajustes_sheet(w):
    names, sheets = w
    for n in names:
        if n.strip().lower() == "general": return sheets[n]
    for n in names:
        if re.match(r"^AJUSTES", n.strip(), re.I): return sheets[n]
    return []
def ajustes_mes_anio(r, iF, iM):
    fecha = parse_fecha(g(r, iF)) if iF >= 0 else None
    mn = g(r, iM) if iM >= 0 else None
    mes = (int(jsnum(mn)) - 1) if (mn is not None and isnum(jsnum(mn)) and (isnum(mn) or (isinstance(mn, str) and mn.strip()))) else (fecha.month - 1 if fecha else None)
    return (fecha.year if fecha else None), mes
def ext_ajustes_agregado(w):
    rows = ajustes_sheet(w); h = rows[0]
    iR, iC, iF, iM = hidx(h, "Referencia"), hidx(h, "Costo neto (prom.)"), hidx(h, "Fecha"), hidx(h, "MES")
    m = {}
    for r in rows[1:]:
        if iR < 0 or g(r, iR) is None: continue
        anio, mes = ajustes_mes_anio(r, iF, iM)
        if not month_ok(None if mes is None else mes + 1): continue
        ref = S(r[iR]).upper(); k = (anio, mes, ref)
        o = m.get(k)
        if o is None: o = m[k] = {"anio": anio, "mes": mes, "referencia": ref, "count": 0, "costoNeto": 0, "costoPositivo": 0, "costoNegativo": 0}
        o["count"] += 1
        c = r[iC] if (iC >= 0 and isnum(g(r, iC))) else 0
        o["costoNeto"] += c
        if c > 0: o["costoPositivo"] += c
        elif c < 0: o["costoNegativo"] += c
    return list(m.values())
def ext_ajustes_dim(w, col, fallback_map=None):
    rows = ajustes_sheet(w); h = rows[0]
    iD, iC, iF, iM = hidx(h, col), hidx(h, "Costo neto (prom.)"), hidx(h, "Fecha"), hidx(h, "MES")
    iB = hidx(h, "Bodega")
    m = {}
    for r in rows[1:]:
        anio, mes = ajustes_mes_anio(r, iF, iM)
        if not month_ok(None if mes is None else mes + 1): continue
        if iD >= 0 and g(r, iD) is not None: valor = S(r[iD])
        elif fallback_map is not None and iB >= 0 and g(r, iB) is not None and S(r[iB]) in fallback_map: valor = fallback_map[S(r[iB])]
        else: valor = "(sin dato)"
        k = (anio, mes, valor)
        o = m.get(k)
        if o is None: o = m[k] = {"anio": anio, "mes": mes, "valor": valor, "count": 0, "costoNeto": 0, "costoPositivo": 0, "costoNegativo": 0}
        o["count"] += 1
        c = r[iC] if (iC >= 0 and isnum(g(r, iC))) else 0
        o["costoNeto"] += c
        if c > 0: o["costoPositivo"] += c
        elif c < 0: o["costoNegativo"] += c
    return list(m.values())
def salidas_sheet(w):
    names, sheets = w
    for n in names:
        if re.match(r"^SALIDAS", n.strip(), re.I): return sheets[n]
    return []
def ext_salidas_agregado(w):
    rows = salidas_sheet(w); h = rows[0]
    iM, iA, iMo, iR, iC = hidx(h, "Mes"), hidx(h, "Año"), hidx(h, "Desc. motivo"), hidx(h, "Referencia"), hidx(h, "Costo salidas (prom.)")
    m = {}
    for r in rows[1:]:
        if iR < 0 or g(r, iR) is None: continue
        mn, an = g(r, iM), g(r, iA)
        anio = int(an) if isnum(an) else None
        mes = int(mn) - 1 if isnum(mn) else None
        if not month_ok(None if mes is None else mes + 1): continue
        motivo = S(r[iMo]) if (iMo >= 0 and g(r, iMo) is not None) else "(sin motivo)"
        k = (anio, mes, motivo)
        o = m.get(k)
        if o is None: o = m[k] = {"anio": anio, "mes": mes, "motivo": motivo, "costoSalida": 0}
        o["costoSalida"] += r[iC] if (iC >= 0 and isnum(g(r, iC))) else 0
    return list(m.values())
def ext_ventas(w):
    names, sheets = w; out = []
    for n in names:
        if n.strip().upper() not in MESES_UP: continue
        idx = MESES_UP.index(n.strip().upper())
        if not month_ok(idx + 1): continue
        tot = 0
        for r in sheets[n][1:]:
            if g(r, 1) is None: continue
            if isnum(g(r, 4)): tot += r[4]
        out.append({"mes": idx, "anio": 2026, "total": tot})
    return out

# ═════════════ 2) Confiabilidad (ReliabilityEngine) ═════════════
def find_col(header, pats):
    for p in pats:
        for i, h in enumerate(header):
            if re.search(p, norm(h).lower()): return i
    return -1
def parse_ajustes_engine(aoa):
    h = aoa[0]
    c = dict(tipoInv=find_col(h, [r"desc\.?\s*tipo\s*inventario"]), bodega=find_col(h, [r"^bodega$"]),
             descItem=find_col(h, [r"desc\.?\s*item"]), referencia=find_col(h, [r"^referencia$"]),
             descMotivo=find_col(h, [r"desc\.?\s*motivo"]), umInv=find_col(h, [r"^u\.?m\.?\s*inv\.?$"]),
             umSigla=find_col(h, [r"u\.?m\s*inv\s*sigla"]), entQty=find_col(h, [r"entradas.*sobra"]),
             costoEnt=find_col(h, [r"costo\s*entradas"]), salQty=find_col(h, [r"salidas.*falta"]),
             costoSal=find_col(h, [r"costo\s*salidas"]), costoNeto=find_col(h, [r"costo\s*neto"]),
             fecha=find_col(h, [r"^fecha$"]))
    groups = {}
    for row in aoa[1:]:
        bodega, ref = g(row, c["bodega"]), g(row, c["referencia"])
        fr = g(row, c["fecha"]) if c["fecha"] >= 0 else None
        if bodega is None or ref is None or fr is None: continue
        ref = ref.strip() if isinstance(ref, str) else ref
        if isinstance(fr, (dt.datetime, dt.date)): ym = fr.year * 100 + fr.month
        elif isnum(fr):
            d = (dt.datetime(1899, 12, 30) + dt.timedelta(days=int(fr))); ym = d.year * 100 + d.month
        else: continue
        if not month_ok(ym % 100): continue
        key = jsstr(bodega) + "|" + str(ym) + "|" + jsstr(ref)
        gr = groups.get(key)
        if gr is None:
            um = g(row, c["umSigla"]) if (c["umSigla"] >= 0 and truthy(g(row, c["umSigla"]))) else (g(row, c["umInv"]) if c["umInv"] >= 0 else None)
            gr = groups[key] = dict(bodega=bodega, ym=ym, referencia=ref, desc=g(row, c["descItem"]) if c["descItem"] >= 0 else None,
                tipoInv=g(row, c["tipoInv"]) if c["tipoInv"] >= 0 else None, um=um,
                entQty=0, salQty=0, ce=0, cs=0, cn=0, n=0, motivos={})
        gr["entQty"] += jsnum(g(row, c["entQty"])); gr["salQty"] += jsnum(g(row, c["salQty"]))
        gr["ce"] += abs(jsnum(g(row, c["costoEnt"]))); gr["cs"] += abs(jsnum(g(row, c["costoSal"])))
        gr["cn"] += jsnum(g(row, c["costoNeto"])); gr["n"] += 1
        mot = g(row, c["descMotivo"]) if c["descMotivo"] >= 0 else None
        if truthy(mot): gr["motivos"][mot] = gr["motivos"].get(mot, 0) + 1
    return groups
def parse_catalogo_engine(aoa):
    h = aoa[0]
    c = dict(bodega=find_col(h, [r"^bodega$"]), estado=find_col(h, [r"^estado$"]),
             nombreBodega=find_col(h, [r"nombre\s*bodega", r"desc\.?\s*bodega"]),
             zona=find_col(h, [r"^zona$", r"^ciudad$", r"^regi[oó]n$"]), referencia=find_col(h, [r"^referencia$"]),
             descItem=find_col(h, [r"desc\.?\s*item"]), tipoItem=find_col(h, [r"tipo\s*item"]), manufactura=find_col(h, [r"manufactura"]))
    uni, bn, bz = {}, {}, {}
    for row in aoa[1:]:
        bodega = g(row, c["bodega"]); ref = g(row, c["referencia"])
        if bodega is None: continue
        if c["nombreBodega"] >= 0 and bodega not in bn:
            n = norm(g(row, c["nombreBodega"]))
            if n and n != "#N/A": bn[bodega] = n
        if c["zona"] >= 0 and bodega not in bz:
            z = norm(g(row, c["zona"]))
            if z and z != "#N/A": bz[bodega] = z
        if ref is None: continue
        ref = ref.strip() if isinstance(ref, str) else ref
        key = jsstr(bodega) + "|" + jsstr(ref)
        if key in uni: continue
        uni[key] = dict(bodega=bodega, referencia=ref, desc=g(row, c["descItem"]) if c["descItem"] >= 0 else None,
            estado=g(row, c["estado"]) if c["estado"] >= 0 else None, tipoItem=g(row, c["tipoItem"]) if c["tipoItem"] >= 0 else None,
            manufactura=g(row, c["manufactura"]) if c["manufactura"] >= 0 else None)
    return list(uni.values()), bn, bz
class Enc:
    def __init__(s): s.idx, s.list = {}, []
    def enc(s, v):
        if v is None: v = ""
        k = (type(v) is str, jsstr(v))
        if k not in s.idx: s.idx[k] = len(s.list); s.list.append(v)
        return s.idx[k]
def or_nd(v): return v if truthy(v) else "N/D"
def build_payload(groups, universe, bn, bz):
    B, R, Dd, TI, UM, MOT, EST, TIT, MF = (Enc() for _ in range(9))
    rows = []
    for gr in groups.values():
        top, mx = None, -1
        for k, v in gr["motivos"].items():
            if v > mx: mx, top = v, k
        rows.append([B.enc(gr["bodega"]), gr["ym"], R.enc(gr["referencia"]), Dd.enc(gr["desc"]), TI.enc(gr["tipoInv"]), UM.enc(gr["um"]),
            round2(gr["entQty"]), round2(gr["salQty"]), round2(gr["ce"]), round2(gr["cs"]), round2(gr["cn"]), round2(gr["ce"] + gr["cs"]),
            gr["n"], MOT.enc(top), EST.enc("N/D"), TIT.enc("N/D"), MF.enc("N/D")])
    cat = {jsstr(c["bodega"]) + "|" + jsstr(c["referencia"]): c for c in universe}
    for row in rows:
        c = cat.get(jsstr(B.list[row[0]]) + "|" + jsstr(R.list[row[2]]))
        if c:
            row[14] = EST.enc(or_nd(c["estado"])); row[15] = TIT.enc(or_nd(c["tipoItem"])); row[16] = MF.enc(or_nd(c["manufactura"]))
    uni = []
    for c in universe:
        if truthy(c["desc"]): Dd.enc(c["desc"])
        uni.append([B.enc(c["bodega"]), R.enc(c["referencia"]), EST.enc(or_nd(c["estado"])), TIT.enc(or_nd(c["tipoItem"])), MF.enc(or_nd(c["manufactura"]))])
    bnl = [bn.get(code, code) for code in B.list]
    zl = [bz.get(code, "Sin zona") for code in B.list]
    return {"dicts": {"b": B.list, "r": R.list, "d": Dd.list, "ti": TI.list, "um": UM.list, "mot": MOT.list, "est": EST.list,
                      "tit": TIT.list, "mf": MF.list, "bn": bnl, "zona": zl},
            "cols": ['b','ym','r','d','ti','um','eq','sq','ce','cs','cn','cm','n','mot','est','tit','mf'],
            "universeCols": ['b','r','est','tit','mf'], "rows": rows, "universe": uni}
def detect_kind(h):
    low = [norm(x).lower() for x in h]
    has = lambda p: any(re.search(p, x) for x in low)
    if has(r"costo neto") and has(r"bodega") and has(r"referencia"): return "ajustes"
    if has(r"manufactura") and has(r"estado") and has(r"referencia") and has(r"tipo\s*inventario") and not has(r"costo neto"): return "catalogo"
    return None

# ═════════════ 3) Comportamiento de salidas (__CS_DATA__) ═════════════
CONSUMO = {'CONSUMO BODEGA','CONSUMO INTERNO POR INGRESOS','CONSUMO MENAJE Y DOTACION','REPOSICION DE SALIDA','SALIDA PRESTAMO A TERCEROS'}
def build_cs(old, w_sal, w_ven, w_aju, catalog_bn, catalog_bz):
    BOD = [dict(b) for b in old["bodegas"]]; ITEMS = [dict(i) for i in old["items"]]
    TIPOINV = list(old["tipoinv"]); UM = list(old["um"]); MOT = [dict(m) for m in old["motivos"]]
    SIG = [dict(s) for s in old["siglas"]]; CPT = list(old["cptos"]); AMOT = list(old["ajusteMotivos"])
    bidx = {b["code"]: i for i, b in enumerate(BOD)}; iidx = {it["name"]: i for i, it in enumerate(ITEMS)}
    midx = {m["name"]: i for i, m in enumerate(MOT)}
    nuevas_bod = []
    def bod(code, ciudad=None, desc=None, zona=None):
        code = norm(code)
        if not code: return -1
        if code not in bidx:
            nm = norm(desc) if desc else ""
            nm = re.sub(r"^PDV\s+", "", nm) or catalog_bn.get(code) or code
            b = {"code": code, "name": nm, "tipo": "PV" if code.startswith("PV") else "CD", "ciudad": norm(ciudad)}
            if zona and norm(zona) != "#N/A": b["zona"] = norm(zona)
            BOD.append(b); bidx[code] = len(BOD) - 1; nuevas_bod.append(code)
        else:
            b = BOD[bidx[code]]
            if ciudad and not b.get("ciudad"): b["ciudad"] = norm(ciudad)
            if zona and norm(zona) != "#N/A" and not b.get("zona"): b["zona"] = norm(zona)
        return bidx[code]
    def lst(L, name):
        name = norm(name)
        if not name: return -1
        if name not in L: L.append(name)
        return L.index(name)
    def item(name, ti, um):
        name = norm(name)
        if name not in iidx: ITEMS.append({"name": name, "tipoinv": ti, "um": um}); iidx[name] = len(ITEMS) - 1
        i = iidx[name]
        if um is not None and um >= 0 and (ITEMS[i].get("um") is None or ITEMS[i]["um"] < 0): ITEMS[i]["um"] = um
        return i
    def mot(name):
        name = norm(name)
        if name not in midx: MOT.append({"name": name, "grupo": "C" if name in CONSUMO else "P"}); midx[name] = len(MOT) - 1
        return midx[name]
    def sigla(code):
        code = norm(code)
        if not code: return None
        if not any(s["code"] == code for s in SIG): SIG.append({"code": code, "hint": []})
        return code
    def hdr(row): return {jsstr(h).strip(): i for i, h in enumerate(row or []) if h is not None}
    # salidas
    rows = salidas_sheet(w_sal); ix = hdr(rows[0]); agg = {}
    for row in rows[1:]:
        mes, bc = g(row, ix['Mes']), g(row, ix['Bodega'])
        if mes is None or bc is None: continue
        mes = int(mes) if isnum(mes) else mes
        if not month_ok(mes): continue
        costo = g(row, ix['Costo salidas (prom.)']) or 0; cant = (g(row, ix['Salidas (inv.)']) or 0) if 'Salidas (inv.)' in ix else 0
        b = bod(bc, g(row, ix.get('Desc. ciudad bodega', -1)), g(row, ix.get('Desc. bodega', -1)), g(row, ix.get('ZONA', -1)))
        m = mot(g(row, ix['Desc. motivo'])); ti = lst(TIPOINV, g(row, ix['Desc. tipo inventario']))
        u = lst(UM, g(row, ix['U.M. inv.']) if 'U.M. inv.' in ix else ""); it = item(g(row, ix['Desc. item']), ti, u)
        k = (mes, b, m, it); a = agg.get(k)
        if a is None: a = agg[k] = [0, 0, 0]
        a[0] += costo; a[1] += cant; a[2] += 1
    fact = [[k[0], k[1], k[2], k[3], v[0], v[1], v[2]] for k, v in agg.items()]
    # ventas
    names, sheets = w_ven; ventas = []
    MN = {m: i + 1 for i, m in enumerate(MESES_UP)}
    for sn in names:
        mes = MN.get(sn.strip().upper())
        if not mes or not month_ok(mes): continue
        rr = sheets[sn]
        if len(rr) < 2: continue
        ix = hdr(rr[0]); sc = ix.get('Valor subtotal', 2); bc_ = ix.get('Bodega', 1); cc = ix.get('C.O.', 0)
        for row in rr[1:]:
            co = norm(g(row, cc)) if g(row, cc) is not None else ""
            bc = g(row, bc_)
            if co == "TOTAL" or not truthy(bc): continue
            ventas.append([mes, bod(bc), g(row, sc) or 0])
    # ajustes (+ item + detalle)
    rows = ajustes_sheet(w_aju); ix = hdr(rows[0]); hasS = 'Documento SIGLA' in ix
    agg, aggI, aggD = {}, {}, {}
    for row in rows[1:]:
        mes, bc = g(row, ix['MES']), g(row, ix['Bodega'])
        if mes is None or bc is None: continue
        mes = int(mes) if isnum(mes) else mes
        if not month_ok(mes): continue
        con = norm(g(row, ix['Concepto']) or ""); sg = sigla(g(row, ix['Documento SIGLA'])) if hasS else 'AJC'
        cn = g(row, ix['Costo neto (prom.)']) or 0; b = bod(bc)
        k = (mes, b, con, sg); agg[k] = agg.get(k, 0) + cn
        ti = lst(TIPOINV, g(row, ix['Desc. tipo inventario'])); u = lst(UM, g(row, ix['U.M. inv.']))
        it = item(g(row, ix['Desc. item']), ti, u)
        k = (mes, b, con, sg, it); aggI[k] = aggI.get(k, 0) + cn
        cp = lst(CPT, g(row, ix['Desc. cpto.'])); am = lst(AMOT, g(row, ix['Desc. motivo']))
        k = (mes, b, cp, am); aggD[k] = aggD.get(k, 0) + cn
    ajustes = [[k[0], k[1], k[2], k[3], v] for k, v in agg.items()]
    ajI = [[k[0], k[1], k[2], k[3], k[4], round2(v)] for k, v in aggI.items()]
    ajD = [[k[0], k[1], k[2], k[3], round2(v)] for k, v in aggD.items()]
    meses = sorted({r[0] for r in fact} | {r[0] for r in ventas} | {r[0] for r in ajustes})
    out = {"meses": {str(m): MESES[m - 1] for m in meses}, "bodegas": BOD, "items": ITEMS, "tipoinv": TIPOINV, "um": UM,
           "motivos": MOT, "siglas": SIG, "fact": fact, "ventas": ventas, "ajustes": ajustes, "ajustesItem": ajI,
           "ajustesDetalle": ajD, "cptos": CPT, "ajusteMotivos": AMOT}
    return out, nuevas_bod

# ═════════════ HTML ═════════════
def clean(o):
    if isinstance(o, float):
        if o.is_integer() and abs(o) < 1e15: return int(o)
        return o
    if isinstance(o, list): return [clean(x) for x in o]
    if isinstance(o, dict): return {k: clean(v) for k, v in o.items()}
    return o
def dumps(o): return json.dumps(clean(o), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

html = open(HTML, encoding="utf-8").read()
START = "/* ==DATA_RAW_START== NO EDITAR A MANO — generado por generar_datos.js */"; END = "/* ==DATA_RAW_END== */"
i0, i1 = html.index(START), html.index(END)
old_raw = json.loads(html[i0 + len(START):i1].strip()[len("const DATA_RAW = "):-1])
CB_OPEN = '<script type="application/json" id="data-blob-confiabilidad">'
c0 = html.index(CB_OPEN) + len(CB_OPEN); c1 = html.index("</script>", c0)
old_cb = json.loads(html[c0:c1])
CS_OPEN = "window.__CS_DATA__ = "
s0 = html.index(CS_OPEN) + len(CS_OPEN); s1 = html.index("\n", s0)
cs_txt = html[s0:s1].rstrip(); assert cs_txt.endswith(";")
old_cs = json.loads(cs_txt[:-1])

W = read_all()
# catálogo del motor de confiabilidad
names, sheets = W["confiab"]; universe, bn, bz = [], {}, {}
for sn in names:
    aoa = sheets[sn]
    if aoa and detect_kind(aoa[0]) == "catalogo": universe, bn, bz = parse_catalogo_engine(aoa); break
names, sheets = W["ajustes"]; groups = {}
for sn in names:
    aoa = sheets[sn]
    if aoa and detect_kind(aoa[0]) == "ajustes": groups = parse_ajustes_engine(aoa); break
new_cb = build_payload(groups, universe, bn, bz)

# mapa código bodega -> "Desc. Bodega" (el archivo de ajustes ya no trae esa columna; se toma del catálogo)
cat_rows = find_sheet(W["confiab"], "REFERENCIAS"); ch = cat_rows[0]
iB, iDB = hidx(ch, "Bodega"), hidx(ch, "Desc. Bodega"); desc_bod = {}
for r in cat_rows[1:]:
    if g(r, iB) is not None and g(r, iDB) is not None and S(r[iB]) not in desc_bod: desc_bod[S(r[iB])] = S(r[iDB])
# preferir el nombre de BASE SALIDAS ("PDV DE MODA"), que es el formato que traía la antigua columna "Desc. Bodega"
sal_rows = salidas_sheet(W["salidas"]); sh = sal_rows[0]; jB, jD = hidx(sh, "Bodega"), hidx(sh, "Desc. bodega")
for r in sal_rows[1:]:
    if g(r, jB) is not None and g(r, jD) is not None: desc_bod.setdefault("~" + S(r[jB]), S(r[jD]))
desc_bod = {**desc_bod, **{k[1:]: v for k, v in desc_bod.items() if k.startswith("~")}}

new_raw = dict(old_raw)
new_raw["auditoria"] = {"records": ext_auditoria(W["auditoria"])}
new_raw["hallazgos"] = {"records": ext_hallazgos(W["hallazgos"]), "listas": ext_listas(W["hallazgos"])}
new_raw["controles"] = ext_controles(W["controles"])
new_raw["riesgos"] = {"records": ext_riesgos(W["riesgos"])}
new_raw["confiabilidad"] = {"catalogo": ext_catalogo(W["confiab"]), "ajustes": ext_ajustes_agregado(W["ajustes"]),
    "ajustesPorTipo": ext_ajustes_dim(W["ajustes"], "Desc. tipo inventario"),
    "ajustesPorBodega": ext_ajustes_dim(W["ajustes"], "Desc. Bodega", desc_bod)}
new_raw["costos"] = {"salidas": ext_salidas_agregado(W["salidas"]), "ventas": ext_ventas(W["ventas"])}
now = dt.datetime.now(); ma = ["ene","feb","mar","abr","may","jun","jul","ago","sept","oct","nov","dic"]
new_raw["generadoEn"] = f"{now.day:02d} de {ma[now.month-1]} de {now.year}, {(now.hour%12 or 12):02d}:{now.minute:02d} {'a. m.' if now.hour<12 else 'p. m.'}"
new_cs, nuevas_bod = build_cs(old_cs, W["salidas"], W["ventas"], W["ajustes"], bn, bz)

# ───────── comparación ─────────
def close(a, b): return a == b or (isnum(a) and isnum(b) and abs(a - b) <= 1e-6 * max(1, abs(a), abs(b)))
def cmp_rows(name, old, new, keyf, valf):
    O = {keyf(r): valf(r) for r in old}; N = {keyf(r): valf(r) for r in new}
    same = sum(1 for k in O if k in N and all(close(x, y) for x, y in zip(O[k], N[k])))
    print(f"  {name}: antes {len(O)} | ahora {len(N)} | iguales {same} | solo antes {len(set(O)-set(N))} | solo ahora {len(set(N)-set(O))}")
    return O, N
print("── DATA_RAW")
for k in ("auditoria", "hallazgos", "controles", "riesgos"): print(f"  {k}: {len(old_raw[k]['records'])} -> {len(new_raw[k]['records'])}")
oc, nc = old_raw["confiabilidad"], new_raw["confiabilidad"]
print("  catálogo:", oc["catalogo"]["total"], "->", nc["catalogo"]["total"], "| igual:", oc["catalogo"] == nc["catalogo"])
cmp_rows("ajustes", oc["ajustes"], nc["ajustes"], lambda r: (r["anio"], r["mes"], r["referencia"]), lambda r: (r["count"], r["costoNeto"], r["costoPositivo"], r["costoNegativo"]))
cmp_rows("ajustesPorTipo", oc["ajustesPorTipo"], nc["ajustesPorTipo"], lambda r: (r["anio"], r["mes"], r["valor"]), lambda r: (r["count"], r["costoNeto"]))
cmp_rows("ajustesPorBodega", oc["ajustesPorBodega"], nc["ajustesPorBodega"], lambda r: (r["anio"], r["mes"], r["valor"]), lambda r: (r["count"], r["costoNeto"]))
cmp_rows("salidas", old_raw["costos"]["salidas"], new_raw["costos"]["salidas"], lambda r: (r["anio"], r["mes"], r["motivo"]), lambda r: (r["costoSalida"],))
cmp_rows("ventas", old_raw["costos"]["ventas"], new_raw["costos"]["ventas"], lambda r: (r["anio"], r["mes"]), lambda r: (r["total"],))
print("── Confiabilidad blob")
def dec_cb(p):
    d = p["dicts"]
    return ([(d["b"][r[0]], r[1], jsstr(d["r"][r[2]])) + tuple(r[6:13]) + (d["mot"][r[13]], d["est"][r[14]]) for r in p["rows"]],
            [(d["b"][u[0]], jsstr(d["r"][u[1]]), d["est"][u[2]]) for u in p["universe"]])
orows, ouni = dec_cb(old_cb); nrows, nuni = dec_cb(new_cb)
cmp_rows("rows", orows, nrows, lambda r: r[:3], lambda r: r[3:])
print(f"  universe: antes {len(ouni)} ahora {len(nuni)} iguales {len(set(ouni)&set(nuni))}")
print("── Salidas (__CS_DATA__)")
def dec_cs(p):
    B = [b["code"] for b in p["bodegas"]]; I = [i["name"] for i in p["items"]]; M = [m["name"] for m in p["motivos"]]
    return {"fact": [(r[0], B[r[1]], M[r[2]], I[r[3]], r[4], r[5], r[6]) for r in p["fact"]],
            "ventas": [(r[0], B[r[1]], r[2]) for r in p["ventas"]],
            "ajustes": [(r[0], B[r[1]], r[2], r[3], r[4]) for r in p["ajustes"]],
            "ajustesItem": [(r[0], B[r[1]], r[2], r[3], I[r[4]], r[5]) for r in p["ajustesItem"]],
            "ajustesDetalle": [(r[0], B[r[1]], p["cptos"][r[2]], p["ajusteMotivos"][r[3]], r[4]) for r in p["ajustesDetalle"]]}
do, dn = dec_cs(old_cs), dec_cs(new_cs)
for k, nk in (("fact", 4), ("ventas", 2), ("ajustes", 4), ("ajustesItem", 5), ("ajustesDetalle", 4)):
    cmp_rows(k, do[k], dn[k], lambda r, nk=nk: r[:nk], lambda r, nk=nk: r[nk:])
print("  meses:", list(new_cs["meses"].values()), "| bodegas nuevas:", nuevas_bod, "| items", len(old_cs["items"]), "->", len(new_cs["items"]))

if VALIDAR: sys.exit()
stamp = now.strftime("%Y-%m-%d-%H-%M")
bk = os.path.join(DIR, "extract", f"Tablero_Control_Interno (backup {stamp}).html"); shutil.copyfile(HTML, bk)
out = html[:i0] + START + "\nconst DATA_RAW = " + dumps(new_raw) + ";\n" + END + html[i1 + len(END):]
c0 = out.index(CB_OPEN) + len(CB_OPEN); c1 = out.index("</script>", c0)
out = out[:c0] + dumps(new_cb) + out[c1:]
s0 = out.index(CS_OPEN) + len(CS_OPEN); s1 = out.index("\n", s0)
out = out[:s0] + dumps(new_cs) + ";" + out[s1:]
open(HTML, "w", encoding="utf-8", newline="").write(out)
print("\nEscrito. Respaldo:", bk, "| generadoEn:", new_raw["generadoEn"])
