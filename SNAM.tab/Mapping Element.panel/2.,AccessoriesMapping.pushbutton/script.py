# -*- coding: utf-8 -*-
"""
Compilazione parametri da CI_xxx.xlsx e da Regole mappatura
Solo Pipe Accessories
"""
__title__ = 'Accessories\n mapping'
__author__ = 'Valerio Mascia'
import clr, os, re, xlrd
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List
# WinForms dialog per selezione file
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult

# converte colonna Excel (A Z, AA) in indice zero-based
def col_letter_to_index(letter):
    idx = 0
    for c in letter:
        if c.isalpha(): idx = idx * 26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

# helper per convertire valori Excel in stringhe senza ".0" per interi
def format_cell_value(cell):
    v = cell.value
    if cell.ctype == 2:  # numeric cell
        if int(v) == v: return str(int(v))
        return str(v)
    s = str(v).strip()
    if re.match(r'^-?\d+\.0$', s): return s[:-2]
    return s

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
MAP_RULES_EXCEL = scegli_file_excel("Seleziona il file Regole mappatura per Revit")
CI_FILE         = scegli_file_excel("Seleziona il file CI_xxx.xlsx")
SCHEDA_MAP      = "AP (Accessori per tubazioni)"  # rimane fisso

doc = __revit__.ActiveUIDocument.Document
filter_cat = List[ElementId]([ElementId(int(BuiltInCategory.OST_PipeAccessory))])
elements = FilteredElementCollector(doc)\
    .WherePasses(ElementMulticategoryFilter(filter_cat))\
    .WhereElementIsNotElementType()\
    .ToElements()

# Estrai codice modello
title = doc.Title or ""
m = re.search(r"-(\d+)_", title)
if not m:
    TaskDialog.Show("Errore", "Impossibile estrarre il codice modello dal nome file.")
    raise SystemExit
model_code = m.group(1)

ci_path = os.path.join(CI_FOLDER, "CI_" + model_code + ".xlsx")
if not os.path.exists(ci_path):
    TaskDialog.Show("Errore", "File CI_" + model_code + ".xlsx non trovato.")
    raise SystemExit

# Carica workbook
wb_ci = xlrd.open_workbook(ci_path)
wb_map = xlrd.open_workbook(MAP_RULES_EXCEL)
ws_map = wb_map.sheet_by_name(SCHEDA_MAP)

# Leggi regole di mappatura
param_rules = []
for i in range(1, ws_map.nrows):
    pname = str(ws_map.cell(i,1).value).strip()
    if not pname: continue
    mode = str(ws_map.cell(i,2).value or "").strip().upper()
    desc = str(ws_map.cell(i,3).value or "").strip()
    if mode == "W":
        m_col = re.search(r"colonna\s+([A-Z]+)", desc, re.I)
        m_sheet = re.search(r'foglio\s+"([^\"]+)"', desc, re.I)
        if m_col and m_sheet:
            col = m_col.group(1); sheet = m_sheet.group(1)
            idx = col_letter_to_index(col)
            prefix = pname.split("_")[0].upper()
            param_rules.append(("W", pname, prefix, sheet, idx))
    elif mode == "C":
        fixed = format_cell_value(ws_map.cell(i,3))
        param_rules.append(("C", pname, fixed))
    elif mode == "D":
        names = re.findall(r'"([^\"]+)"', desc)
        mval = re.search(r"\(([^\)]+)\)", desc)
        val = mval.group(1).strip() if mval else ""
        param_rules.append(("D", pname, names, val))
    elif mode == "E":
        names = re.findall(r'"([^\"]+)"', desc)
        param_rules.append(("E", pname, names))
    elif mode == "F":
        names = re.findall(r'"([^\"]+)"', desc)
        vals = re.findall(r"\(([^\)]+)\)", desc)
        v1 = vals[0].strip() if vals else ""; v2 = vals[1].strip() if len(vals)>1 else ""
        param_rules.append(("F", pname, names, v1, v2))
    elif mode == "G":
        names = re.findall(r'"([^\"]+)"', desc)
        vals = re.findall(r"\(([^\)]+)\)", desc)
        tv = vals[0].strip() if vals else ""; fv = vals[1].strip() if len(vals)>1 else ""
        param_rules.append(("G", pname, names, tv, fv))

# Costruisci dati per W
data_by_sheet = {}
sap_col_map = {"Report":1, "Consistenza Impiantistica": col_letter_to_index('N')}
for rule in param_rules:
    if rule[0] == "W":
        sheet = rule[3]
        if sheet not in data_by_sheet:
            ws = wb_ci.sheet_by_name(sheet)
            sap_col = sap_col_map.get(sheet,1)
            rows = []
            for r in range(1, ws.nrows):
                code = format_cell_value(ws.cell(r, sap_col))
                if not code: continue
                line = [format_cell_value(ws.cell(r,c)) for c in range(ws.ncols)]
                rows.append((code, line))
            data_by_sheet[sheet] = rows

# Avvia transazione
trans = Transaction(doc, "Compila parametri CI")
trans.Start()
param_count = 0

# Segmento G da Title
parts = title.split('-')
segG = parts[4] if len(parts)>4 else ""

for el in elements:
    sap_p = el.LookupParameter("NP259_codice_sap")
    if sap_p is None: continue
    sap_val = sap_p.AsString().strip() if sap_p.AsString() else ""
    if not sap_val: continue
    sym = doc.GetElement(el.GetTypeId())
    fam = sym.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString().strip()
    
    for rule in param_rules:
        typ = rule[0]
        # W
        if typ == "W":
            _, pname, prefix, sheet, idx = rule
            prm = el.LookupParameter(pname)
            if prm and not prm.IsReadOnly:
                val = None
                rows = data_by_sheet.get(sheet, [])
                if sheet == "Report" and idx == col_letter_to_index('EE'):
                    dz = col_letter_to_index('DZ')
                    for code, rv in rows:
                        if dz < len(rv) and rv[dz].upper().startswith(prefix):
                            if idx < len(rv): val = rv[idx]
                            break
                else:
                    for code, rv in rows:
                        if code == sap_val and idx < len(rv):
                            val = rv[idx]
                            break
                if val is not None:
                    prm.Set(val)
                    param_count += 1
        # C
        elif typ == "C":
            _, pname, fixed = rule
            prm = el.LookupParameter(pname)
            if prm and not prm.IsReadOnly:
                prm.Set(fixed)
                param_count += 1
        # D
        elif typ == "D":
            _, pname, names, dval = rule
            prm = el.LookupParameter(pname)
            if prm and not prm.IsReadOnly:
                for nm in names:
                    if fam.startswith(nm):
                        prm.Set(dval)
                        param_count += 1
                        break
        # E
        elif typ == "E":
            _, pname, names = rule
            prm = el.LookupParameter(pname)
            if prm and not prm.IsReadOnly:
                prm.Set("SI" if fam in names else "NO")
                param_count += 1
        # F
        elif typ == "F":
            _, pname, names, v1, v2 = rule
            prm = el.LookupParameter(pname)
            if prm and not prm.IsReadOnly:
                prm.Set(v1 if any(fam.startswith(n) for n in names) else v2)
                param_count += 1
        # G
        elif typ == "G":
            _, pname, names, tv, fv = rule
            prm = el.LookupParameter(pname)
            if prm and not prm.IsReadOnly:
                prm.Set(tv if segG in names else fv)
                param_count += 1

# Commit e output
trans.Commit()
TaskDialog.Show("Completato", str(param_count) + " parametri compilati da CI_" + model_code + ".xlsx")
