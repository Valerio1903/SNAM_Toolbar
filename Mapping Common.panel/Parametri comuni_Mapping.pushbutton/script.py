# -*- coding: utf-8 -*-
"""
Compilazione automatica parametri comuni
"""
__title__ = 'Parametri comuni\n mapping'
__author__ = 'Valerio Mascia'
import clr, os, re
# Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List as NetList
# Excel
import xlrd

# ------------------- CONFIGURAZIONE -------------------
MAPPE_PATH    = r"C:\Users\2Dto6D\OneDrive\Desktop\Techfem_Parametri\Regole mappatura per Revit_2Dto6D.xlsx"
ALLEGATO_PATH = r"C:\Users\2Dto6D\OneDrive\Desktop\Techfem_Parametri\Allegato 2 - Lista Asset Affidamento.xlsx"
SHEET_MAPPE   = "PARAMETRI COMUNI"
SHEET_ALLEGATO= "Lista Asset Affidamento"

# -------------------------------------------------------

def col_letter_to_index(letter):
    # A->0, B->1, ..., Z->25, AA->26, etc.
    idx = 0
    for c in letter:
        if c.isalpha(): idx = idx*26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1


def _norm(val):
    try:
        if val is None: return ""
        # xlrd numbers are float
        if isinstance(val, float):
            return str(int(val)) if val.is_integer() else str(val)
        s = str(val).strip()
        return s if not re.match(r'^-?\d+\.0$', s) else s[:-2]
    except:
        return ""


def leggi_colonne(path, sheet_name):
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_name(sheet_name)
    cols = []
    for c in range(ws.ncols):
        col = [ws.cell(r, c).value for r in range(ws.nrows)]
        cols.append(col)
    return cols


def build_param_values(map_cols, all_cols, codice_edificio, file_name):
    names = map_cols[1]   # colonna B
    rules = map_cols[2]   # colonna C
    # trova righe Allegato per codice edificio
    colF = all_cols[5]
    rows = [i for i in range(1, len(colF)) if str(colF[i]).strip() == codice_edificio]
    if not rows:
        raise Exception("Codice edificio '{}' non trovato in Allegato".format(codice_edificio))
    sel = rows[0]
    # prefer match su Impianto Tipo (colonna J=9)
    if len(all_cols) > 9:
        colJ = all_cols[9]
        for i in rows:
            if i < len(colJ) and colJ[i] and colJ[i] in file_name:
                sel = i
                break

    values = {}
    for r in range(1, len(names)):
        pname = str(names[r]).strip()
        if not pname.startswith('NP'): continue
        rule = str(rules[r]).strip()
        raw = None
        m = re.match(r'colonna\s+([A-Za-z]+)', rule, re.I)
        if m:
            # estrai lettere colonna excel
            col_letters = m.group(1)
            idx = col_letter_to_index(col_letters)
            if idx < len(all_cols) and sel < len(all_cols[idx]):
                raw = all_cols[idx][sel]
        else:
            # valore fisso
            raw = rule
        values[pname] = _norm(raw)
    return values

# Inizio esecuzione
uiapp = __revit__
doc = uiapp.ActiveUIDocument.Document
# estrai codice edificio dal nome file
fname = os.path.basename(doc.PathName)
base = os.path.splitext(fname)[0]
parts = base.split('-')
if len(parts) < 4:
    TaskDialog.Show('Errore', 'Nome file non valido, serve almeno 4 segmenti separati da -')
    raise SystemExit
codice = parts[3].strip()

# legge Excel
try:
    map_cols = leggi_colonne(MAPPE_PATH, SHEET_MAPPE)
    all_cols = leggi_colonne(ALLEGATO_PATH, SHEET_ALLEGATO)
except Exception as e:
    TaskDialog.Show('Errore', 'Impossibile leggere Excel: {}'.format(e))
    raise SystemExit

# costruisce valori
try:
    param_values = build_param_values(map_cols, all_cols, codice, fname)
except Exception as e:
    TaskDialog.Show('Errore', str(e))
    raise SystemExit

# filtra elementi target
cats = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_PipeAccessory]
cat_ids = NetList[ElementId]([ElementId(int(c)) for c in cats])
elems = FilteredElementCollector(doc)\
    .WherePasses(ElementMulticategoryFilter(cat_ids))\
    .WhereElementIsNotElementType()\
    .ToElements()

# Scrittura parametri
skipped_params = set()
updated = 0
trans = Transaction(doc, 'Set Parametri Comuni')
trans.Start()
for el in elems:
    for pname, pval in param_values.items():
        prm = el.LookupParameter(pname)
        if prm and not prm.IsReadOnly:
            prm.Set(pval)
            updated += 1
        else:
            skipped_params.add(pname)
trans.Commit()

# Output
if skipped_params:
    msg = 'Parametri con nome errato su Excel:' + '\n' + '\n'.join(sorted(skipped_params))
else:
    msg = 'Tutti i parametri sono stati aggiornati.'
TaskDialog.Show('Completato', msg)
