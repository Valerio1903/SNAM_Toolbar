# -*- coding: utf-8 -*-
"""
Compilazione parametri da CI_xxx.xlsx, Regole mappatura e CSV di linea (AP)
Pipe Accessories + Pipe Fittings (famiglie AP)
"""

__title__ = 'AP_Accessories\nmapping'
__author__ = 'Valerio Mascia'

import clr, os, re, xlrd, csv
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import TransactionStatus
from System.Collections.Generic import List
from Autodesk.Revit.DB import StorageType, UnitTypeId

# ---------- Helper Excel/CSV ----------
def col_letter_to_index(letter):
    idx = 0
    for c in letter:
        if c.isalpha():
            idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

def format_cell_value(cell):
    v = cell.value
    if cell.ctype == 2:
        if int(v) == v:
            return str(int(v))
        return str(v)
    s = str(v).strip()
    if re.match(r'^-?\d+\.0$', s):
        return s[:-2]
    return s

def get_param_as_string(prm):
    if prm is None:
        return None
    st = prm.StorageType
    if st == StorageType.String:
        return prm.AsString().strip() if prm.AsString() else None
    if st == StorageType.Integer:
        return str(prm.AsInteger())
    if st == StorageType.Double:
        return str(prm.AsDouble())
    if st == StorageType.ElementId:
        return str(prm.AsElementId().IntegerValue)
    return None

def read_csv_rows(path):
    # autodetect delimitatore ; o ,
    with open(path, 'rb') as fb:
        raw = fb.read().decode('utf-8-sig', 'ignore')
    first_line = raw.splitlines()[0] if raw.splitlines() else ""
    delim = ';' if first_line.count(';') >= first_line.count(',') else ','
    rows = []
    for row in csv.reader(raw.splitlines(), delimiter=delim):
        rows.append([c.strip() for c in row])
    return rows

# ---------- Helper T (come NO033) ----------
def _level_elev_ft(lv):
    if not lv: return None
    try:
        p = lv.get_Parameter(BuiltInParameter.LEVEL_ELEV)
        if p and p.StorageType == StorageType.Double:
            return p.AsDouble()
    except:
        pass
    try:
        return lv.Elevation
    except:
        return None

def _ft_to_mm(x_ft):
    try:
        return UnitUtils.ConvertFromInternalUnits(x_ft, UnitTypeId.Millimeters)
    except:
        return x_ft * 304.8

def _fmt_mm(v_mm):
    s = "{:.3f}".format(v_mm)
    return s.rstrip("0").rstrip(".")

def _elem_world_z_ft(doc, el):
    # livello assegnato
    lvl = None
    try:
        if hasattr(el, "LevelId") and el.LevelId and el.LevelId.IntegerValue > 0:
            lvl = doc.GetElement(el.LevelId)
        if lvl is None:
            plev = el.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
            if plev and plev.StorageType == StorageType.ElementId:
                lvl = doc.GetElement(plev.AsElementId())
    except:
        lvl = None
    z_lvl = _level_elev_ft(lvl)

    # offset istanza
    off = None
    for bip in (BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
                BuiltInParameter.RBS_OFFSET_PARAM,
                BuiltInParameter.INSTANCE_ELEVATION_PARAM):
        try:
            p = el.get_Parameter(bip)
            if p and p.StorageType == StorageType.Double:
                off = p.AsDouble()
                break
        except:
            pass

    if z_lvl is not None and off is not None:
        return z_lvl + off

    # fallback: posizione geometrica
    try:
        loc = el.Location
        if hasattr(loc, "Point") and loc.Point:
            return loc.Point.Z
    except:
        pass
    return None

_num = re.compile(r"[-+]?\d*\.?\d+")

def _first_number(txt):
    m = _num.search(str(txt))
    if m:
        try:
            v = float(m.group().replace(",", "."))
            return int(v) if v.is_integer() else v
        except:
            pass
    return None

def _format_number_keep_decimals(n):
    try:
        s = ("{:.6f}".format(float(n))).rstrip("0").rstrip(".")
        return "0" if s in ("", "-0") else s
    except:
        return str(n)

# --- Date helper per regola P (come Placeholders v3) ---
IT_MONTHS = {
    'gen':'01','feb':'02','mar':'03','apr':'04','mag':'05','giu':'06',
    'lug':'07','ago':'08','set':'09','ott':'10','nov':'11','dic':'12'
}
_DATE_RE = re.compile(r"^(\d{1,2})[-/\.](\w{3})[-/\.](\d{2,4})$", re.I)

def to_date_ddmmyyyy(s):
    m = _DATE_RE.match((s or "").strip())
    if not m:
        return s
    d = int(m.group(1))
    mon = (m.group(2) or "").strip().lower()[:3]
    y = m.group(3)
    mm = IT_MONTHS.get(mon)
    if not mm:
        return s
    if len(y) == 2:
        yy = int(y)
        y4 = 2000 + yy if yy < 50 else 1900 + yy
    else:
        y4 = int(y)
    return ("%02d%s%04d" % (d, mm, y4))

# ---------- PATHS e CONFIG ----------
MAP_RULES_EXCEL = r"C:\Users\2Dto6D\OneDrive\Desktop\Techfem_Parametri\Regole mappatura per Revit_2Dto6D.xlsx"
CI_FOLDER        = r"C:\Users\2Dto6D\OneDrive\Desktop\Techfem_Parametri"
SCHEDA_MAP       = "AP (Accessori per tubazioni)"
SCHEDA_COMMON    = "PARAMETRI COMUNI"

# Allegato 3 (per J)
ALLEGATO3_PATH  = r"C:\Users\2Dto6D\OneDrive\Desktop\Techfem_Parametri\Allegato 3 - Classi e mappatura IFC.xlsx"
ALLEGATO3_SHEET = "Elenco AP"

# CSV per P
P_CSV_PATH      = r"C:\Users\2Dto6D\OneDrive\Desktop\Techfem_Parametri\DQ_MGRC2_P_LIN_FP_LINEA.csv"
P_MATCH_COL_G   = col_letter_to_index('G')  # 0-based

# tf_ da ignorare nei WARNING perché non fanno parte dei parametri da compilare
TF_EXCLUDE_FROM_WARN = {"Long Name", "IfcObjectType", "IFC Name"}


# ---------- Setup documento ----------
doc = __revit__.ActiveUIDocument.Document

# Collector PA + PF
filter_cat = List[ElementId]([
    ElementId(int(BuiltInCategory.OST_PipeAccessory)),
    ElementId(int(BuiltInCategory.OST_PipeFitting))
])
elements   = FilteredElementCollector(doc)\
    .WherePasses(ElementMulticategoryFilter(filter_cat))\
    .WhereElementIsNotElementType()\
    .ToElements()

# Estrai codice modello dal titolo (per CI)
title = doc.Title or ""
m_model = re.search(r"-(\d+)_", title)
if not m_model:
    print("Errore: impossibile estrarre il codice modello dal nome file.")
    raise SystemExit
model_code = m_model.group(1)

# Segmenti titolo per G e P
parts_title = title.split('-')
segG = parts_title[4] if len(parts_title) > 4 else ""  # come prima
segP_raw = parts_title[3] if len(parts_title) > 3 else ""  # tra 3° e 4° trattino

# Trasformazione speciale per P: "_" -> "/" (primo), poi "_" -> "." (secondo)
def _transform_ap_key(s):
    s = s or ""
    i1 = s.find("_")
    if i1 >= 0:
        s = s[:i1] + "/" + s[i1+1:]
        i2 = s.find("_")
        if i2 >= 0:
            s = s[:i2] + "." + s[i2+1:]
    return s
segP = _transform_ap_key(segP_raw)

# Percorsi e workbook
ci_path = os.path.join(CI_FOLDER, "CI_{0}.xlsx".format(model_code))
if not os.path.exists(ci_path):
    print("Errore: file {0} non trovato.".format(ci_path))
    raise SystemExit

wb_ci  = xlrd.open_workbook(ci_path)
wb_map = xlrd.open_workbook(MAP_RULES_EXCEL)

# Regole mappatura
ws_map = wb_map.sheet_by_name(SCHEDA_MAP)

# Parametri comuni
ws_common = wb_map.sheet_by_name(SCHEDA_COMMON)
common_params = set()
for i in range(1, ws_common.nrows):
    pname = str(ws_common.cell(i, 1).value).strip()
    if pname:
        common_params.add(pname)

# Costruisci regole
param_rules = []
rule_names  = set()
for i in range(1, ws_map.nrows):
    pname = str(ws_map.cell(i, 1).value).strip()
    if not pname:
        continue
    rule_names.add(pname)
    mode = str(ws_map.cell(i, 2).value or "").strip().upper()
    desc = str(ws_map.cell(i, 3).value or "").strip()

    if mode == "W":
        m_col   = re.search(r"colonna\s+([A-Z]+)", desc, re.I)
        m_sheet = re.search(r'foglio\s+"([^"]+)"', desc, re.I)
        if m_col and m_sheet:
            col    = m_col.group(1)
            sheet  = m_sheet.group(1)
            idx    = col_letter_to_index(col)
            prefix = pname.split("_")[0].upper()
            param_rules.append(("W", pname, prefix, sheet, idx))

    elif mode == "C":
        param_rules.append(("C", pname, format_cell_value(ws_map.cell(i, 3))))

    elif mode == "D":
        names = re.findall(r'"([^"]+)"', desc)
        mval  = re.search(r"\(([^\)]+)\)", desc)
        param_rules.append(("D", pname, names, mval.group(1).strip() if mval else ""))

    elif mode == "E":
        param_rules.append(("E", pname, re.findall(r'"([^"]+)"', desc)))

    elif mode == "F":
        names = re.findall(r'"([^"]+)"', desc)
        vals  = re.findall(r"\(([^\)]+)\)", desc)
        param_rules.append(("F", pname, names,
                            vals[0].strip() if len(vals)>0 else "",
                            vals[1].strip() if len(vals)>1 else ""))

    elif mode == "G":
        names = re.findall(r'"([^"]+)"', desc)
        vals  = re.findall(r"\(([^\)]+)\)", desc)
        param_rules.append(("G", pname, names,
                            vals[0].strip() if len(vals)>0 else "",
                            vals[1].strip() if len(vals)>1 else ""))

    # --- NUOVE ---
    elif mode == "J":
        m_col = re.search(r"colonna\s+([A-Z]+)", desc, re.I)
        if m_col:
            idx = col_letter_to_index(m_col.group(1))
            param_rules.append(("J", pname, idx))

    elif mode == "K":
        # due opzioni separate da ';' (left;right) in base a elevazione/offset < 0
        param_rules.append(("K", pname, desc or ""))

    elif mode == "T":
        # nessun extra; calcolo geometrico
        param_rules.append(("T", pname))

    elif mode == "P":
        # colonna da prendere dal CSV "colonna X"
        m_col = re.search(r"colonna\s+([A-Z]+)", desc, re.I)
        if m_col:
            idx = col_letter_to_index(m_col.group(1))
            param_rules.append(("P", pname, idx))

    elif mode == "N":
        # "famiglie" tra virgolette; coppie [TYPE6](VAL)
        fam_keys = re.findall(r'"([^"]+)"', desc or "")
        pairs    = re.findall(r"\[([^\]]+)\]\s*\(([^)]*)\)", desc or "")
        # pairs = [("TYPE6", "VAL"), ...]
        param_rules.append(("N", pname, fam_keys, pairs))

    elif mode == "R":
        # Leggi direttamente dal parametro tf_<pname>
        param_rules.append(("R", pname))

    elif mode == "Y":
        # In descrizione c'è il NOME del parametro (senza virgolette)
        pname_src = str(ws_map.cell(i, 3).value or "").strip()
        if pname_src:
            param_rules.append(("Y", pname, pname_src))



# Prepara dati per W
data_by_sheet = {}
sap_col_map   = {"Report":1, "Consistenza Impiantistica":col_letter_to_index('N')}
for rule in param_rules:
    if rule[0] == "W":
        sheet = rule[3]
        if sheet not in data_by_sheet:
            ws     = wb_ci.sheet_by_name(sheet)
            sap_col= sap_col_map.get(sheet, 1)
            rows   = []
            for r in range(1, ws.nrows):
                code = format_cell_value(ws.cell(r, sap_col))
                if not code:
                    continue
                rows.append((code, [format_cell_value(ws.cell(r, c)) for c in range(ws.ncols)]))
            data_by_sheet[sheet] = rows

# Carica Allegato3 per J (lazy in uso)
allegato_data = None

# Carica CSV per P
csv_rows = []
csv_g_index = P_MATCH_COL_G
try:
    csv_rows = read_csv_rows(P_CSV_PATH)
except:
    csv_rows = []

# Indicizzazione CSV per colonna G
csv_by_key = {}
if csv_rows:
    for i in range(1, len(csv_rows)):  # salta header
        row = csv_rows[i]
        if csv_g_index < len(row):
            k = row[csv_g_index].strip()
            if k:
                csv_by_key[k] = row

# --- Precalcolo per T: livello matchato al PBP ---
matched_level = None
pbp_elev = None
try:
    pbps = (FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
            .WhereElementIsNotElementType()
            .ToElements())
    if pbps:
        ep = pbps[0].get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
        if ep and ep.StorageType == StorageType.Double:
            pbp_elev = ep.AsDouble()
except:
    pbp_elev = None

if pbp_elev is not None:
    levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
    tol = 1e-4
    closest = None; dmin = None
    for lv in levels:
        lev = _level_elev_ft(lv)
        if lev is None: continue
        d = abs(lev - pbp_elev)
        if dmin is None or d < dmin:
            dmin = d; closest = lv
        if d < tol:
            matched_level = lv
            break
    if matched_level is None:
        matched_level = closest

# ------------------- TRANSAZIONE -------------------
trans = Transaction(doc, "AP mapping (PA + PF)")
trans.Start()
try:
    debug_log   = []
    param_count = 0

    for el in elements:
        sym        = doc.GetElement(el.GetTypeId())
        tprm       = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        type_name  = tprm.AsString().strip() if tprm and tprm.AsString() else ""
        type_label = "[{0}][ID:{1}]".format(type_name, el.Id.IntegerValue)
        fam_prm    = sym.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        fam        = fam_prm.AsString().strip() if fam_prm and fam_prm.AsString() else ""
        prefix_str = "{0} {1}".format(fam, type_label)

        if not fam.startswith("AP"):
            # opzionale: puoi non loggare nulla così non “inquina” l’output
            # debug_log.append("{0} Skip: family name '{1}' non AP".format(prefix_str, fam))
            continue

        # Controllo tf_ vs regole (come prima)
        for p in el.Parameters:
            pname_tf = p.Definition.Name
            if pname_tf.startswith("tf_"):
                base = pname_tf[3:]
                # ESCLUSIONE: questi tf_ non vanno riportati nell'output finale
                if base in TF_EXCLUDE_FROM_WARN:
                    continue
                if base in common_params:
                    continue
                if base not in rule_names:
                    debug_log.append("{0} WARNING: '{1}' Parametro non presente in Regole Mappatura Parametri".format(prefix_str, pname_tf))

        # Filtra famiglie AP
        if not fam.startswith("AP"):
            debug_log.append("{0} Skip: family name '{1}' non AP".format(prefix_str, fam))
            continue

        # Verifica codice SAP
        sap_val = get_param_as_string(el.LookupParameter("NP259_codice_sap"))
        if not sap_val:
            debug_log.append("{0} WARNING: NP259_codice_sap mancante".format(prefix_str))
            continue

        # Applicazione regole
        for rule in param_rules:
            typ, pname = rule[0], rule[1]
            prm = el.LookupParameter(pname)
            if prm is None or prm.IsReadOnly:
                debug_log.append("{0} WARNING {1}: Parametro non presente in Revit".format(prefix_str, pname))
                continue
            has_tf = el.LookupParameter("tf_{0}".format(pname)) or sym.LookupParameter("tf_{0}".format(pname))
            if not has_tf:
                debug_log.append("{0} Skip {1}: tf_ parameter not defined".format(prefix_str, pname))
                continue

            val = None


            
            if typ == "W":
                _, _, prefix, sheet, idx = rule
                rows = data_by_sheet.get(sheet, [])
                val = None

                # --- Caso speciale: foglio "Report" + colonna EE ---
                # Qui lo stesso SAP ha più righe, una per ogni parametro CAxxx.
                # Serve trovare la riga con SAP uguale E DZ che identifica il parametro (prefix = "CA002", "CA008", ...).
                # Se DZ è tutta vuota -> match su EA e valore da EG
                if sheet == "Report" and idx == col_letter_to_index('EE'):
                    dz = col_letter_to_index('DZ')
                    ea = col_letter_to_index('EA')
                    eg = col_letter_to_index('EG')

                    # verifica se in DZ esiste almeno un valore non vuoto (per le righe del SAP corrente)
                    dz_has_any = False
                    for code, rv in rows:
                        if code != sap_val:
                            continue
                        dz_val_chk = (rv[dz] if dz < len(rv) else '')
                        if str(dz_val_chk).strip() != '':
                            dz_has_any = True
                            break

                    if dz_has_any:
                        # Comportamento attuale: match su DZ e valore da EE
                        for code, rv in rows:
                            if code != sap_val:
                                continue
                            dz_val = str(rv[dz] if dz < len(rv) else '').upper().strip()
                            if dz_val.startswith(prefix):
                                if idx < len(rv):
                                    candidate = rv[idx]
                                    if candidate is not None and str(candidate).strip() != "":
                                        val = candidate
                                break
                    else:
                        # Eccezione: DZ tutta vuota -> match su EA e valore da EG
                        for code, rv in rows:
                            if code != sap_val:
                                continue
                            ea_val = str(rv[ea] if ea < len(rv) else '').upper().strip()
                            if ea_val.startswith(prefix):
                                if eg < len(rv):
                                    candidate = rv[eg]
                                    if candidate is not None and str(candidate).strip() != "":
                                        val = candidate
                                break

                # --- Caso generale: match per SAP (1 riga per SAP) ---
                else:
                    for code, rv in rows:
                        if code == sap_val:
                            if idx < len(rv):
                                candidate = rv[idx]
                                if candidate is not None and str(candidate).strip() != "":
                                    val = candidate
                            break


            elif typ == "C":
                val = rule[2]

            elif typ == "D":
                names, dval = rule[2], rule[3]
                if any(fam.startswith(n) for n in names):
                    val = dval

            elif typ == "E":
                val = "SI" if fam in rule[2] else "NO"

            elif typ == "F":
                names, v1, v2 = rule[2], rule[3], rule[4]
                val = v1 if any(fam.startswith(n) for n in names) else v2

            elif typ == "G":
                names, tv, fv = rule[2], rule[3], rule[4]
                val = tv if segG in names else fv

            # ----- NUOVE REGOLE -----
            elif typ == "J":
                # Allegato 3: prefisso 5 del Family Name
                if allegato_data is None:
                    try:
                        wb_all = xlrd.open_workbook(ALLEGATO3_PATH)
                        ws_all = wb_all.sheet_by_name(ALLEGATO3_SHEET)
                        allegato_data = []
                        for c in range(ws_all.ncols):
                            col = []
                            for r in range(ws_all.nrows):
                                col.append(ws_all.cell(r, c).value)
                            allegato_data.append(col)
                    except:
                        allegato_data = None
                if allegato_data is not None:
                    idx_out = rule[2]
                    if 0 <= idx_out < len(allegato_data):
                        prefix5 = (fam[:5] if fam else "").upper()
                        colA = allegato_data[0] if len(allegato_data) > 0 else []
                        row = None
                        for irow in range(1, len(colA)):
                            if str(colA[irow]).strip().upper() == prefix5:
                                row = irow
                                break
                        if row is not None and row < len(allegato_data[idx_out]):
                            val = allegato_data[idx_out][row]

            elif typ == "K":
                # due opzioni separate da ';' in base a elevazione/offset < 0
                desc_k = rule[2] or ""
                parts_k = desc_k.split(";", 1)
                left  = parts_k[0].strip() if len(parts_k) >= 1 else ""
                right = parts_k[1].strip() if len(parts_k) == 2 else ""
                elev = None
                # prova INSTANCE_ELEVATION_PARAM, poi RBS_OFFSET_PARAM
                for bip in (BuiltInParameter.INSTANCE_ELEVATION_PARAM,
                            BuiltInParameter.RBS_OFFSET_PARAM):
                    try:
                        p = el.get_Parameter(bip)
                        if p and p.StorageType == StorageType.Double:
                            elev = p.AsDouble()
                            break
                    except:
                        pass
                choice = (left or right) if (elev is not None and elev < 0.0) else (right or left)
                val = choice if choice != "" else None

            elif typ == "T":
                # distanza firmata rispetto al livello matchato PBP
                if matched_level is not None:
                    z_lvl = _level_elev_ft(matched_level)
                    z_el  = _elem_world_z_ft(doc, el)
                    if z_el is not None and z_lvl is not None:
                        dz_ft = z_el - z_lvl
                        # se parametro Double, set in feet, altrimenti stringa mm
                        if prm.StorageType == StorageType.Double:
                            try:
                                prm.Set(dz_ft)
                                param_count += 1
                                debug_log.append("{0} T Set {1}: val={2}".format(prefix_str, pname, dz_ft))
                                continue
                            except:
                                pass
                        dz_mm = _ft_to_mm(dz_ft)
                        val = _fmt_mm(dz_mm)

            elif typ == "P":
                # match key (trasformata) con colonna G del CSV
                idx_out = rule[2]
                row = csv_by_key.get(segP)
                if row and 0 <= idx_out < len(row):
                    cand = row[idx_out]
                    # vuoto -> "N/C"
                    if cand is None or str(cand).strip() == '':
                        val = "N/C"
                    else:
                        s = str(cand).strip()
                        # 'Esercizio' -> 'Operativo'
                        if s.lower() == 'esercizio':
                            s = 'Operativo'
                        # converti "27-feb-17" -> "27022017"
                        s = to_date_ddmmyyyy(s)
                        val = s
                else:
                    val = "N/C"

            elif typ == "N":
                # rule = ("N", pname, fam_keys, pairs)
                fam_keys, pairs = rule[2], rule[3]
                fam_prefix5  = (fam[:5] if fam else "").upper()
                type_prefix6 = (type_name[:6] if type_name else "").upper()
                # default N/C
                val = "N/C"
                # match family prefix5 con una delle virgolette
                if fam_keys and any(fam_prefix5 == k.strip().upper() for k in fam_keys):
                    # match type prefix6 con una delle quadre
                    # supporta più alternative nelle quadre separate da | , ; / o spazi
                    for k, outv in pairs:
                        alts = [a.strip().upper() for a in re.split(r'[|,;/]+', k) if a.strip()]
                        if (alts and type_prefix6 in alts) or (not alts and type_prefix6 == k.strip().upper()):
                            val = outv.strip()
                            break
                # else: resta "N/C"


            elif typ == "R":
                # Copia il valore da tf_<pname> (istanza -> tipo come fallback)
                src_name = "tf_{0}".format(pname)
                src_prm = el.LookupParameter(src_name) or sym.LookupParameter(src_name)
                if src_prm:
                    val = get_param_as_string(src_prm)
                else:
                    val = None
            
            
            elif typ == "Y":
                # rule = ("Y", target_param_name, source_param_name)
                src_name = rule[2]
                src = el.LookupParameter(src_name) or sym.LookupParameter(src_name)
                if not src:
                    debug_log.append("{0} Y WARNING: parametro sorgente '{1}' non trovato".format(prefix_str, src_name))
                    continue

                if prm.StorageType != StorageType.String:
                    debug_log.append("{0} Y WARNING: il parametro destinazione '{1}' non è di tipo Testo".format(prefix_str, pname))
                    continue

                s_out = None

                if src.StorageType == StorageType.Double:
                    # Come Z: converti da internal units a millimetri e formatta senza zeri inutili
                    try:
                        try:
                            mm = UnitUtils.ConvertFromInternalUnits(src.AsDouble(), UnitTypeId.Millimeters)
                        except:
                            from Autodesk.Revit.DB import DisplayUnitType
                            mm = UnitUtils.ConvertFromInternalUnits(src.AsDouble(), DisplayUnitType.DUT_MILLIMETERS)
                        s_out = _format_number_keep_decimals(round(mm, 6))
                    except:
                        s_out = None
                else:
                    # fallback: usa AsString e prova a prendere il primo numero
                    s = (src.AsString() or "").strip()
                    if s:
                        n = _first_number(s)
                        if n is not None:
                            s_out = _format_number_keep_decimals(n)

                if not s_out:
                    debug_log.append("{0} Y WARNING: valore non interpretabile da '{1}'".format(prefix_str, src_name))
                    continue

                try:
                    prm.Set(s_out)
                    param_count += 1
                    debug_log.append("{0} Y Set {1}: val={2}".format(prefix_str, pname, s_out))
                except:
                    try:
                        prm.SetValueString(str(s_out))
                        param_count += 1
                        debug_log.append("{0} Y SetValueString {1}: val={2}".format(prefix_str, pname, s_out))
                    except:
                        debug_log.append("{0} Y WARNING: set fallita per {1} (val={2})".format(prefix_str, pname, s_out))


            # ----- scrittura -----
            if val is None:
                debug_log.append("{0} {1} Skipped {2}: no rule value".format(prefix_str, typ, pname))
                continue

            try:
                prm.Set(val)
                param_count += 1
                debug_log.append("{0} {1} Set {2}: val={3}".format(prefix_str, typ, pname, val))
            except:
                # tentativo come ValueString (alcuni string param)
                try:
                    prm.SetValueString(str(val))
                    param_count += 1
                    debug_log.append("{0} {1} SetValueString {2}: val={3}".format(prefix_str, typ, pname, val))
                except:
                    debug_log.append("{0} {1} WARNING: set fallita per {2} (val={3})".format(prefix_str, typ, pname, val))

    trans.Commit()

except Exception:
    if trans.GetStatus() == TransactionStatus.Started:
        trans.RollBack()
    raise

# -------- Output HTML (come il tuo, invariato nella struttura) --------
from pyrevit import script
output = script.get_output()
output.print_html("<b>Completato: {0} parametri compilati</b><br>".format(param_count))

# Raggruppa e stampa solo i warning “WARNING:” per family name 
logs_by_family = {}
for line in debug_log:
    family = line.split(" ", 1)[0]
    first_close  = line.find(']')
    second_close = line.find(']', first_close + 1)
    if second_close != -1:
        text = line[second_close + 1 :].strip()
    else:
        text = line[len(family) :].strip()
    if not text.startswith("WARNING:"):
        continue
    logs_by_family.setdefault(family, []).append(text)

for family, warnings in logs_by_family.items():
    output.print_html("<h3 style='font-size:1.5em;'>{0}</h3>".format(family))
    for w in warnings:
        output.print_html(
            "<span style='color:purple;font-weight:bold;font-size:1.5em;'>{0}</span><br>".format(w)
        )