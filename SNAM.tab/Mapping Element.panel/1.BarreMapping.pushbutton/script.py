# -*- coding: utf-8 -*-
"""
Compila i parametri delle tubazioni (Pipe Curves) basandosi sulle regole
nel foglio "BARRE (CATEGORIA TUBAZIONI)" di Excel.
"""
__title__ = 'Pipe\nmapping'
__author__ = 'Valerio Mascia'

import clr
import os
import re
import xlrd
import math
import csv
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List as NetList

def col_letter_to_index(letter):
    idx = 0
    for c in letter:
        if c.isalpha():
            idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1
"""Converte lettere di colonna Excel (es. 'A', 'AA') nell'indice zero-based (A=0, B=1, ...)."""

_num = re.compile("[-+]?\d*\.?\d+")
# Regex per catturare il *primo numero* in una stringa (segno opzionale, parte intera e decimale con punto).
# N.B.: le virgole decimali vengono gestite dopo sostituendo "," -> "." in _first_number.

def _first_number(txt):
    m = _num.search(str(txt))
    if m:
        try:
            v = float(m.group().replace(",", "."))
            if v.is_integer():
                return int(v)
            return v
        except:
            pass
    return None
"""Estrae il primo numero presente nel testo (supporta virgola decimale), restituendo int se intero, altrimenti float."""

DN_PARAMS = [
    BuiltInParameter.RBS_PIPE_DIAMETER_PARAM,
    BuiltInParameter.RBS_PIPE_OUTER_DIAMETER
]

def _get_dn(el):
    for bip in DN_PARAMS:
        prm = el.get_Parameter(bip)
        if prm:
            num = _first_number(prm.AsValueString() or "")
            if num is not None:
                return int(round(num))
    return None
"""Ricava il DN del tubo leggendo i parametri diametro (visualizzati), restituisce un intero (arrotondato) o None."""


def _param_to_str(prm):
    if prm.StorageType == StorageType.Double:
        s = prm.AsValueString() or ""
        if s.strip():
            return s.strip()
        return ("{:.6f}".format(prm.AsDouble())).rstrip("0").rstrip(".")
    if prm.StorageType == StorageType.Integer:
        return str(prm.AsInteger())
    return (prm.AsString() or "").strip()
"""Converte un parametro Revit in stringa: Double via AsValueString (o fallback), Integer come stringa, altrimenti AsString."""


def _val_to_str(val):
    try:
        n = float(val)
        if n.is_integer():
            return str(int(n))
        return str(n)
    except:
        return str(val)
"""Converte un valore generico in stringa; se numero, mantiene i decimali quando presenti (es. 5.0 -> '5')."""    


def _read_cols(path, sheet):
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_name(sheet)
    cols = []
    for c in range(ws.ncols):
        col = []
        for r in range(ws.nrows):
            col.append(ws.cell(r, c).value)
        cols.append(col)
    return cols
"""Legge un foglio Excel e restituisce la matrice per colonne: cols[c][r] = cella (riga r, colonna c)."""


def _get_type_name_pipe(el, doc):
    try:
        typ = doc.GetElement(el.GetTypeId())
        if typ:
            p = typ.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            return (p.AsString() or "") if p else ""
    except:
        pass
    return ""
"""Restituisce il Type Name dell’elemento pipe (stringa vuota se non disponibile)."""


def _format_number_keep_decimals(n):
    """88.0 -> '88' ; 88.12 -> '88.12' ; niente arrotondamenti ad intero."""
    try:
        s = ("{:.6f}".format(float(n))).rstrip("0").rstrip(".")
        return "0" if s in ("", "-0") else s
    except:
        return str(n)
"""Formatta un numero in stringa senza zeri inutili: 88.0 -> '88'; 88.12 -> '88.12' (massimo 6 decimali)."""

def _read_csv_rows(path):
    with open(path, 'rb') as fb:
        raw = fb.read().decode('utf-8-sig', 'ignore')
    lines = raw.splitlines()
    if not lines:
        return []
    delim = ';' if lines[0].count(';') >= lines[0].count(',') else ','
    return [[c.strip() for c in row] for row in csv.reader(lines, delimiter=delim)]

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
    y   = m.group(3)
    mm = IT_MONTHS.get(mon)
    if not mm:
        return s
    if len(y) == 2:
        yy = int(y)
        y4 = 2000 + yy if yy < 50 else 1900 + yy
    else:
        y4 = int(y)
    return ("%02d%s%04d" % (d, mm, y4))


def process_document(doc):
    excel_path = "C:\\Users\\2Dto6D\\OneDrive\\Desktop\\Techfem_Parametri\\Regole mappatura per Revit_2Dto6D.xlsx"
    sheet = "BARRE (CATEGORIA TUBAZIONI)"
    dn_lookup = {"BARRE_GASD": "BARRE_GASD"}
    # --- Config P (CSV di linea) ---
    P_CSV_PATH    = r"C:\Users\2Dto6D\OneDrive\Desktop\Techfem_Parametri\DQ_MGRC2_P_LIN_FP_LINEA.csv"
    P_MATCH_COL_G = col_letter_to_index('G')  # chiave di match


    # Allegato 3 per la regola J
    allegato3_path = "C:\\Users\\2Dto6D\\OneDrive\\Desktop\\Techfem_Parametri\\Allegato 3 - Classi e mappatura IFC.xlsx"
    allegato3_sheet = "Elenco NO"
    allegato_data = None  # caricato lazy al primo uso di J

    cols = _read_cols(excel_path, sheet)
    names = [str(c).strip() for c in cols[1]]
    codes = [str(c).strip().upper() for c in cols[2]]
    descs = [str(c).strip() for c in cols[3]]
    rules = []
    for i in range(1, len(names)):
        if names[i]:
            rules.append((names[i], codes[i], descs[i]))

    pipes = FilteredElementCollector(doc) \
        .WherePasses(ElementMulticategoryFilter(NetList[ElementId]([
            ElementId(int(BuiltInCategory.OST_PipeCurves))
        ]))) \
        .WhereElementIsNotElementType() \
        .ToElements()


    # --- derivazione chiave P dal titolo (stessa trasformazione degli Accessori)
    def _transform_ap_key(s):
        s = s or ""
        i1 = s.find("_")
        if i1 >= 0:
            s = s[:i1] + "/" + s[i1+1:]
            i2 = s.find("_")
            if i2 >= 0:
                s = s[:i2] + "." + s[i2+1:]
        return s

    title = doc.Title or ""
    parts_title = title.split('-')
    # segP è il pezzo tra il 3° e 4° trattino, trasformato
    segP_raw = parts_title[3] if len(parts_title) > 3 else ""
    segP = _transform_ap_key(segP_raw)

    # --- leggi CSV e indicizza per colonna G (0-based)
    csv_by_key = {}
    try:
        _rows = _read_csv_rows(P_CSV_PATH)
        if _rows:
            gidx = P_MATCH_COL_G
            for i in range(1, len(_rows)):  # salta header
                row = _rows[i]
                if gidx < len(row):
                    k = row[gidx].strip()
                    if k:
                        csv_by_key[k] = row
    except:
        csv_by_key = {}

    cache = {}
    parts = (doc.Title or "").split("-")
    segG = parts[4] if len(parts) > 4 else ""
    res_params = {}
    warnings = []

    # STEP 1 (tutte le regole eccetto L)
    t1 = Transaction(doc, "Mappa Barre Step1")
    t1.Start()
    for el in pipes:
        dn_cache = None
        for tgt, code, desc in rules:
            if code == "L":
                continue
            prm = el.LookupParameter(tgt)
            if prm is None or prm.IsReadOnly:
                continue

            # --- (RIMOSSA) N/C: ora ignorata ---
            if code == "N/C":
                # Ignora: non scrive, non conta, nessun warning
                continue

            # C (costante)
            if code == "C":
                v = _val_to_str(desc)
                prm.Set(v)
                res_params[tgt] = v
                continue

            # X (lookup su foglio GASD dichiarato in descrizione: DN match)
            if code == "X":
                mcol = re.search(r"colonna\s+([A-Za-z]+)", desc, re.I)
                msht = re.search(r'foglio\s+"([^"]+)"', desc)
                if mcol and msht:
                    sht = dn_lookup.get(msht.group(1), msht.group(1))
                    if sht not in cache:
                        cache[sht] = _read_cols(excel_path, sht)
                    data = cache[sht]
                    if dn_cache is None:
                        dn_cache = _get_dn(el)
                    if dn_cache is None:
                        warnings.append((tgt, "DN non trovato"))
                        continue
                    # nel tuo file le DN stanno in data[1] (2ª colonna)
                    colDN = data[1]
                    row = None
                    for irow in range(1, len(colDN)):
                        if _first_number(colDN[irow]) == dn_cache:
                            row = irow
                            break
                    if row is None:
                        warnings.append((tgt, "DN " + str(dn_cache) + " non in " + sht))
                        continue
                    idx = col_letter_to_index(mcol.group(1))
                    if idx < len(data) and row < len(data[idx]):
                        val = data[idx][row]
                        vstr = _val_to_str(val)
                        prm.Set(vstr)
                        res_params[tgt] = vstr
                continue

            # ---- PROCESSO Z (sostituisci SOLO questo blocco) ----
            if code == "Z":
                # In Excel il BuiltInParameter è senza virgolette
                src = None
                try:
                    bip = getattr(BuiltInParameter, (desc or "").strip())
                    src = el.get_Parameter(bip)
                except:
                    src = None

                if not src:
                    warnings.append((tgt, 'Z: BuiltInParameter "{}" non trovato'.format(desc)))
                    continue

                # I target, sono di testo
                if prm.StorageType != StorageType.String:
                    warnings.append((tgt, "Z: il parametro destinazione non è di tipo Testo"))
                    continue

                # ---- dentro if code == "Z":, sostituisci SOLO la gestione dei Double ----

                s_out = None

                if src.StorageType == StorageType.Double:
                    # Converti SEMPRE dall'unità interna (feet) ai millimetri,
                    # ignorando le unità/arrotondamenti della UI
                    try:
                        # API nuove
                        try:
                            from Autodesk.Revit.DB import UnitTypeId
                            mm = UnitUtils.ConvertFromInternalUnits(src.AsDouble(), UnitTypeId.Millimeters)
                        except:
                            # API vecchie
                            from Autodesk.Revit.DB import DisplayUnitType
                            mm = UnitUtils.ConvertFromInternalUnits(src.AsDouble(), DisplayUnitType.DUT_MILLIMETERS)

                        s_out = _format_number_keep_decimals(round(mm, 6))  # 88.0 -> "88", 88.9 -> "88.9"
                    except:
                        s_out = None

                # se ancora None, prova AsString() (non AsValueString, per evitare arrotondamenti UI)
                if s_out is None:
                    s = (src.AsString() or "").strip()
                    if s:
                        n = _first_number(s)
                        if n is not None:
                            s_out = _format_number_keep_decimals(n)


                if not s_out:
                    warnings.append((tgt, "Z: nessun valore interpretabile dal parametro sorgente"))
                    continue

                try:
                    prm.Set(s_out)
                    res_params[tgt] = s_out
                except:
                    warnings.append((tgt, "Z: errore in Set('{}')".format(s_out)))
                continue

            # G (segmento titolo progetto)
            if code == "G":
                keys = re.findall(r'"([^"]+)"', desc)
                vals = re.findall(r"\(([^)]+)\)", desc)
                tv = vals[0].strip() if vals else ""
                fv = vals[1].strip() if len(vals) > 1 else ""
                chosen = tv if segG in keys else fv
                vstr = _val_to_str(chosen)
                prm.Set(vstr)
                res_params[tgt] = vstr
                continue

            # K (NUOVA) -> due opzioni separate da ';' in base a offset < 0
            if code == "K":
                left, right = "", ""
                parts_k = desc.split(";", 1)
                if len(parts_k) >= 1:
                    left = parts_k[0].strip()
                if len(parts_k) == 2:
                    right = parts_k[1].strip()
                elev = None
                try:
                    off = el.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)
                    if off and off.StorageType == StorageType.Double:
                        elev = off.AsDouble()  # feet; confronto solo segno
                except:
                    elev = None
                choice = (left or right) if (elev is not None and elev < 0.0) else (right or left)
                if choice != "":
                    vstr = _val_to_str(choice)
                    prm.Set(vstr)
                    res_params[tgt] = vstr
                continue

            # J (MODIFICATA) -> Allegato 3, 'colonna X'
            # Condizione: compila SOLO se TYPE NAME inizia con "BARRE" o "Tubaz" (case-sensitive)
            # Lookup: la chiave in colonna A è SEMPRE "BARRE"
            if code == "J":
                # carica Allegato 3 la prima volta
                if allegato_data is None:
                    try:
                        allegato_data = _read_cols(allegato3_path, allegato3_sheet)
                    except:
                        allegato_data = None
                if allegato_data is None:
                    warnings.append((tgt, "J: impossibile aprire Allegato 3"))
                    continue

                mcol = re.search(r"colonna\s+([A-Za-z]+)", desc or "", re.I)
                if not mcol:
                    warnings.append((tgt, "J: 'colonna X' non specificata"))
                    continue
                idx_out = col_letter_to_index(mcol.group(1).upper())
                if idx_out < 0 or idx_out >= len(allegato_data):
                    warnings.append((tgt, "J: colonna '{}' fuori range".format(mcol.group(1).upper())))
                    continue

                # >>> Trigger solo se Type Name inizia con "BARRE" o "Tubaz" (case-sensitive)
                type_name_raw = _get_type_name_pipe(el, doc) or ""
                if type_name_raw.startswith("BARRE") or type_name_raw.startswith("Tubaz"):
                    key = "BARRE"  # su Allegato 3 la chiave è sempre BARRE
                else:
                    warnings.append((tgt, "J: Type Name non inizia con 'BARRE' o 'Tubaz'"))
                    continue

                colA = allegato_data[0] if len(allegato_data) > 0 else []
                row = None
                for irow in range(1, len(colA)):
                    if str(colA[irow]).strip() == key:  # match esatto e case-sensitive
                        row = irow
                        break
                if row is None:
                    warnings.append((tgt, "J: chiave '{}' non trovata in col. A".format(key)))
                    continue
                if row >= len(allegato_data[idx_out]):
                    warnings.append((tgt, "J: riga {} oltre dati col. {}".format(row, mcol.group(1).upper())))
                    continue

                val = allegato_data[idx_out][row]
                vstr = _val_to_str(val)
                prm.Set(vstr)
                res_params[tgt] = vstr
                continue



            # M (NUOVA) -> mapping da parametro sorgente istanza "SRC" -> (SRC_VAL, OUT_VAL)
            if code == "M":
                msrc = re.match(r'\s*"([^"]+)"', desc or "")
                if not msrc:
                    warnings.append((tgt, "M: parametro sorgente non specificato"))
                    continue
                src_name = msrc.group(1).strip()
                srcp = el.LookupParameter(src_name)
                if srcp is None:
                    warnings.append((tgt, "M: parametro sorgente '{}' non trovato".format(src_name)))
                    continue
                src_val = _param_to_str(srcp).strip()
                chosen = None
                for pair in re.findall(r"\(([^()]*)\)", desc or ""):
                    bits = pair.split(",", 1)
                    if len(bits) >= 2 and src_val == bits[0].strip():
                        chosen = bits[1].strip()
                        break
                if chosen is not None and chosen != "":
                    vstr = _val_to_str(chosen)
                    prm.Set(vstr)
                    res_params[tgt] = vstr
                continue
            

            # P (CSV di linea come negli Accessori)
            if code == "P":
                # In descrizione di Excel è scritto "colonna X"
                m_col = re.search(r"colonna\s+([A-Za-z]+)", desc, re.I)
                if not m_col:
                    warnings.append((tgt, "P: 'colonna X' non specificata"))
                    continue

                idx_out = col_letter_to_index(m_col.group(1).upper())
                row = csv_by_key.get(segP)

                # default N/C se chiave o colonna non disponibili
                v = "N/C"
                if row and 0 <= idx_out < len(row):
                    cand = row[idx_out]
                    if cand is not None and str(cand).strip() != '':
                        s = str(cand).strip()
                        # normalizzazioni come negli Accessori
                        if s.lower() == 'esercizio':
                            s = 'Operativo'
                        s = to_date_ddmmyyyy(s)
                        v = s

                try:
                    prm.Set(_val_to_str(v))
                    res_params[tgt] = _val_to_str(v)
                except:
                    try:
                        prm.SetValueString(_val_to_str(v))
                        res_params[tgt] = _val_to_str(v)
                    except:
                        warnings.append((tgt, "P: errore in Set/SetValueString"))
                continue

    # end for rules
    t1.Commit()

    # STEP 2: regola L (immutata)
    l_rule = None
    for r in rules:
        if r[1] == "L":
            l_rule = r
            break

    if l_rule is not None:
        target_param, _, rule_desc = l_rule

        # tolgo quadre esterne
        if rule_desc.startswith("[") and rule_desc.endswith("]"):
            rule_desc = rule_desc[1:-1].strip()

        # estraggo il nome del parametro sorgente (tra le prime virgolette)
        src_match = re.match(r'"([^"]+)"', rule_desc)
        if not src_match:
            TaskDialog.Show("Errore", "Parametro sorgente non trovato in regola L.")
            return
        source_name = src_match.group(1)

        # divido per '-' ogni condizione
        partsL = rule_desc.split("-")
        cond_map = {}
        for part in partsL:
            part = part.strip()
            start = part.find("[")
            end   = part.find("]")
            if start >= 0 and end > start:
                key = part[:start].strip().lower()
                val = part[start+1:end].strip()
                cond_map[key] = val

        default_val = cond_map.get("default", "")

        t2 = Transaction(doc, "Mappa Barre Step2")
        t2.Start()
        for el in pipes:
            prm_l = el.LookupParameter(target_param)
            if prm_l is None or prm_l.IsReadOnly:
                continue
            srcp = el.LookupParameter(source_name)
            val_key = ""
            if srcp and srcp.AsString() is not None:
                val_key = srcp.AsString().strip().lower()
            chosen = cond_map.get(val_key, default_val)
            if chosen != "":
                prm_l.Set(chosen)
                res_params[target_param] = chosen
        t2.Commit()

    msg = "Parametri aggiornati: {}".format(len(res_params))
    if warnings:
        msg += "\nWarning:"
        for p, w in warnings:
            msg += "\n- {}: {}".format(p, w)

    TaskDialog.Show("Risultato", msg)
    

# punto di ingresso pyRevit
doc = __revit__.ActiveUIDocument.Document
process_document(doc)
