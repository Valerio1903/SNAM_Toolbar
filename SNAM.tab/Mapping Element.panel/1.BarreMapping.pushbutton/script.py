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
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List as NetList

def col_letter_to_index(letter):
    idx = 0
    for c in letter:
        if c.isalpha():
            idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

_num = re.compile("[-+]?\d*\.?\d+")

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

def _param_to_str(prm):
    if prm.StorageType == StorageType.Double:
        s = prm.AsValueString() or ""
        if s.strip():
            return s.strip()
        return ("{:.6f}".format(prm.AsDouble())).rstrip("0").rstrip(".")
    if prm.StorageType == StorageType.Integer:
        return str(prm.AsInteger())
    return (prm.AsString() or "").strip()

def _val_to_str(val):
    try:
        n = float(val)
        if n.is_integer():
            return str(int(n))
        return str(n)
    except:
        return str(val)

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

def _get_type_name_pipe(el, doc):
    try:
        typ = doc.GetElement(el.GetTypeId())
        if typ:
            p = typ.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            return (p.AsString() or "") if p else ""
    except:
        pass
    return ""


def process_document(doc):
    excel_path = "C:\\Users\\2Dto6D\\OneDrive\\Desktop\\Techfem_Parametri\\Regole mappatura per Revit_2Dto6D.xlsx"
    sheet = "BARRE (CATEGORIA TUBAZIONI)"
    dn_lookup = {"BARRE_GASD": "BARRE_GASD"}

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

            # Z (copia BuiltInParameter indicato in chiaro nella descrizione)
            if code == "Z":
                try:
                    bip = getattr(BuiltInParameter, desc)
                    src = el.get_Parameter(bip)
                except:
                    src = None
                if src:
                    vstr = _val_to_str(_param_to_str(src))
                    prm.Set(vstr)
                    res_params[tgt] = vstr
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

            # J (MODIFICATA) -> Allegato 3, 'colonna X', chiave = prefisso 5 del **TYPE NAME**
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

                # >>> prefisso 5 dal TYPE NAME (non dal Family Name)
                type_name_up = (_get_type_name_pipe(el, doc) or "").upper()
                prefix5 = type_name_up[:5]
                if not prefix5:
                    warnings.append((tgt, "J: prefisso Type Name vuoto"))
                    continue

                colA = allegato_data[0] if len(allegato_data) > 0 else []
                row = None
                for irow in range(1, len(colA)):
                    if str(colA[irow]).strip().upper() == prefix5:
                        row = irow
                        break
                if row is None:
                    warnings.append((tgt, "J: chiave '{}' non trovata in col. A".format(prefix5)))
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
