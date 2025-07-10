# -*- coding: utf-8 -*-
"""
Compila i parametri delle tubazioni (Pipe Curves) basandosi sulle regole
nel foglio "BARRE (CATEGORIA TUBAZIONI)" di Excel.
"""
__title__ = 'Pipe\n mapping'
__author__ = 'Valerio Mascia'

import clr, os, re
# Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List as NetList
# WinForms dialog per selezione file
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult
# Excel reader
import xlrd

# ---------------- Funzione per selezionare un file Excel ----------------
def scegli_file_excel(titolo):
    dialog = OpenFileDialog()
    dialog.Title = titolo
    dialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx"
    dialog.Multiselect = False
    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    TaskDialog.Show("Errore", "Operazione annullata: file non selezionato.")
    raise SystemExit

# ------------------- CONFIGURAZIONE DINAMICA -------------------
EXCEL_PATH   = scegli_file_excel("Seleziona il file Regole mappatura per Tubazioni")
SHEET_REGOLE = "BARRE (CATEGORIA TUBAZIONI)"
# Se hai piu fogli di lookup mettili in un dict cosi, ma il file lo scegli tu:
DN_LOOKUP_SHEETS = {'BARRE_GASD':'BARRE_GASD'}

DN_PARAMS = [
    BuiltInParameter.RBS_PIPE_DIAMETER_PARAM,
    BuiltInParameter.RBS_PIPE_OUTER_DIAMETER
]

# -----------------------------------------------------------------------

# helper: col letter to index
def col_letter_to_index(letter):
    idx = 0
    for c in letter:
        if c.isalpha():
            idx = idx*26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

def _read_cols(path, sheet):
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_name(sheet)
    return [[ws.cell(r, c).value for r in range(ws.nrows)]
            for c in range(ws.ncols)]

#  qui continua il resto del tuo script esattamente come prima, usando EXCEL_PATH e SHEET_REGOLE 

# utility functions
_num = re.compile(r"[-+]?\d*\.?\d+")
def _first_number(txt):
    m = _num.search(str(txt))
    if m:
        try:
            v = float(m.group().replace(',', '.'))
            return int(v) if v.is_integer() else v
        except:
            pass
    return None

def _get_dn(el):
    for bip in DN_PARAMS:
        prm = el.get_Parameter(bip)
        if prm:
            num = _first_number(prm.AsValueString() or '')
            if num is not None:
                return int(round(num))
    return None

def _param_to_str(prm):
    if prm.StorageType == StorageType.Double:
        s = prm.AsValueString() or ''
        if s.strip(): return s.strip()
        return ('{:.6f}'.format(prm.AsDouble())).rstrip('0').rstrip('.')
    if prm.StorageType == StorageType.Integer:
        return str(prm.AsInteger())
    return (prm.AsString() or '').strip()

# read Excel columns
def _read_cols(path, sheet):
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_name(sheet)
    cols = []
    for c in range(ws.ncols):
        col = [ws.cell(r, c).value for r in range(ws.nrows)]
        cols.append(col)
    return cols
 # evita i decimali
def _val_to_str(val):
    try:
        num = float(val)
        if num.is_integer():
            return str(int(num))
        else:
            return str(num)
    except:
        return str(val)

# MAIN
uiapp = __revit__
doc = uiapp.ActiveUIDocument.Document

# load rules
try:
    cols = _read_cols(EXCEL_PATH, SHEET_REGOLE)
except Exception as e:
    TaskDialog.Show('Errore', 'Non posso leggere Excel:\n' + str(e))
    raise SystemExit

param_names = [str(c).strip() for c in cols[1]]
codes       = [str(c).strip().upper() for c in cols[2]]
descripts   = [str(c).strip() for c in cols[3]]
rules = [(param_names[i], codes[i], descripts[i])
         for i in range(1, len(param_names)) if param_names[i]]

# collect elements
pipes = (FilteredElementCollector(doc)
    .WherePasses(ElementMulticategoryFilter(NetList[ElementId]([ElementId(int(BuiltInCategory.OST_PipeCurves))])))
    .WhereElementIsNotElementType()
    .ToElements())

# prepare
x_cache = {}
parts = (doc.Title or '').split('-')
segG = parts[4] if len(parts) > 4 else ''
res_params = {}
warnings = []

# STEP 1: apply all except L
t1 = Transaction(doc, 'Mappa Barre Step1')
t1.Start()
for el in pipes:
    dn_cache = None
    for tgt, code, desc in rules:
        if code == 'L': continue
        prm = el.LookupParameter(tgt)
        if not prm or prm.IsReadOnly: continue
        if code == 'N/C':
            prm.Set('N/C'); res_params[tgt] = 'N/C'; continue
        if code == 'C':
            val_str = _val_to_str(desc)
            prm.Set(val_str); res_params[tgt] = val_str
            continue
        if code == 'X':
            mcol = re.search(r'colonna\s+([A-Za-z]+)', desc, re.I)
            msht = re.search(r'foglio\s+"([^\"]+)"', desc)
            if mcol and msht:
                sheet = DN_LOOKUP_SHEETS.get(msht.group(1), msht.group(1))
                if sheet not in x_cache: x_cache[sheet] = _read_cols(EXCEL_PATH, sheet)
                lk = x_cache[sheet]
                if dn_cache is None: dn_cache = _get_dn(el)
                if dn_cache is None:
                    warnings.append((tgt, 'DN non trovato'))
                    continue
                colF = lk[1]
                row = next((i for i in range(1, len(colF)) if _first_number(colF[i]) == dn_cache), None)
                if row is None:
                    warnings.append((tgt, 'DN {} non in {}'.format(dn_cache, sheet)))
                    continue
                idx = col_letter_to_index(mcol.group(1))
                if idx < len(lk) and row < len(lk[idx]):
                    val = lk[idx][row]
                    val_str = _val_to_str(val)
                    prm.Set(val_str); res_params[tgt] = val_str
            continue
        if code == 'Z':
            try:
                bip = getattr(BuiltInParameter, desc)
                src = el.get_Parameter(bip)
            except:
                src = None
            if src:
                val = _param_to_str(src)
                val_str = _val_to_str(val)
                prm.Set(val_str); res_params[tgt] = val_str
            continue
        if code == 'G':
            names = re.findall(r'"([^\"]+)"', desc)
            vals = re.findall(r"\(([^\)]+)\)", desc)
            tv = vals[0].strip() if vals else ''
            fv = vals[1].strip() if len(vals) > 1 else ''
            chosen = tv if segG in names else fv
            val_str = _val_to_str(chosen)
            prm.Set(val_str); res_params[tgt] = val_str
            continue
        if code == 'K':
            opts = [v.strip() for v in desc.split(';')]
            if len(opts) >= 3:
                elev = None
                try:
                    offset_param = el.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM)
                    if offset_param:
                        elev = offset_param.AsDouble()  
                        elev_m = elev * 0.3048  
                    else:
                        elev_m = None
                except:
                    elev_m = None
                if elev_m is None:
                    chosen = opts[2]
                elif elev_m > 0:
                    chosen = opts[0]
                elif elev_m == 0:
                    chosen = opts[1]
                else:
                    chosen = opts[2]
                val_str = _val_to_str(chosen)
                prm.Set(val_str); res_params[tgt] = val_str
            continue
t1.Commit()

# STEP 2: regola L dinamica da Excel (con supporto per parentesi multiple)
l_rule = next((r for r in rules if r[1] == 'L'), None)
if l_rule:
    tgt_l, code_l, desc_l = l_rule

    # Estrai il parametro da leggere e le condizioni dal testo Excel
    m_param = re.match(r'"([^"]+)"', desc_l)
    if not m_param:
        TaskDialog.Show('Errore regola L', 'Parametro sorgente non trovato nella regola L.')
        raise SystemExit

    param_da_leggere = m_param.group(1)

    # Regex corretta per parentesi multiple nel valore
    pattern_condizioni = re.compile(r'([^\(\)-]+)\s*\((.+?)\)(?:\s*-|$)')
    condizioni_map = {}
    for match in pattern_condizioni.finditer(desc_l):
        condizione = match.group(1).strip().lower()
        valore = match.group(2).strip()
        condizioni_map[condizione] = valore

    if not condizioni_map:
        TaskDialog.Show('Errore regola L', 'Nessuna condizione trovata nella regola L.')
        raise SystemExit

    # Esegui la transazione per impostare il valore
    t2 = Transaction(doc, 'Mappa Barre Step2')
    t2.Start()
    for el in pipes:
        prm_l = el.LookupParameter(tgt_l)
        if not prm_l or prm_l.IsReadOnly:
            continue

        ref = el.LookupParameter(param_da_leggere)
        ref_val = (ref.AsString() or '').strip().lower() if ref else ''

        chosen = condizioni_map.get(ref_val.strip().lower(), condizioni_map.get('default', None))

        if chosen:
            prm_l.Set(chosen)
            res_params[tgt_l] = chosen
        else:
            warnings.append((tgt_l, 'Condizione "{}" non riconosciuta'.format(ref_val)))

    t2.Commit()




# report finale
msg = 'Parametri aggiornati: {0}\n'.format(len(res_params))
if warnings:
    msg += 'Warning:\n'
    for p, w in warnings:
        msg += '- {0}: {1}\n'.format(p, w)
TaskDialog.Show('Risultato', msg)