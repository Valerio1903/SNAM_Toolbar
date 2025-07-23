# -*- coding: utf-8 -*-
"""
Compilazione Project Information 
Valorizzazione automatica da doc.Title e Excel Allegato 2
"""
__title__ = 'Project information\nmapping'
__author__ = 'Valerio Mascia'

import clr, os, re

# Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    ProjectInfo, FilteredElementCollector,
    Transaction, BuiltInParameter
)
from Autodesk.Revit.UI import TaskDialog

# WinForms dialog
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult

# Excel
import xlrd

def scegli_file_excel(titolo):
    dialog = OpenFileDialog()
    dialog.Title = titolo
    dialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx"
    dialog.Multiselect = False
    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    else:
        TaskDialog.Show("Errore", "Operazione annullata: file non selezionato.")
        raise SystemExit

def col_letter_to_index(letter):
    idx = 0
    for c in letter:
        if c.isalpha():
            idx = idx*26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

def leggi_colonne(path, sheet_name):
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_name(sheet_name)
    cols = []
    for c in range(ws.ncols):
        cols.append([ws.cell(r, c).value for r in range(ws.nrows)])
    return cols

# CONFIGURAZIONE DINAMICA
ALLEGATO_PATH = scegli_file_excel("Seleziona il file Allegato 2 - Lista Asset Affidamento")
SHEET_ALLEGATO = "Lista Asset Affidamento"

# Legge colonne Allegato2
try:
    cols_allegato = leggi_colonne(ALLEGATO_PATH, SHEET_ALLEGATO)
except Exception as e:
    TaskDialog.Show('Errore', 'Non posso leggere Excel:\n' + str(e))
    raise SystemExit

# Inizio esecuzione
uiapp = __revit__
doc = uiapp.ActiveUIDocument.Document

# Estrai codice edificio dal nome file (4 segmento)
fname = os.path.basename(doc.PathName)
base = os.path.splitext(fname)[0]
parts = base.split('-')
if len(parts) < 4:
    TaskDialog.Show('Errore', 'Nome file non valido, servono almeno 4 segmenti separati da "-"')
    raise SystemExit
codice = parts[3].strip()

# Trova riga Allegato per codice edificio (colonna F = idx 5)
colF = cols_allegato[5]
rows = [i for i in range(1, len(colF)) if str(colF[i]).strip() == codice]
if not rows:
    TaskDialog.Show('Errore', "Codice edificio '" + codice + "' non trovato in Allegato")
    raise SystemExit
sel = rows[0]

# Estrai e normalizza valori
title = doc.Title
segs = title.split('-')

# 1) Building Name = segmento tra 3 e 4 trattino
if len(segs) >= 4:
    building_name = segs[3].strip()
else:
    building_name = title

# 2) Project Status = title senza la parte finale fino al 2 trattino da destra
last = title.rfind('-')
second_last = title.rfind('-', 0, last)
if second_last != -1:
    project_status = title[:second_last]
else:
    project_status = title

# 3) Project Name = title senza la parte finale fino al 1 trattino da destra
if last != -1:
    project_name = title[:last]
else:
    project_name = title

# 4) Project Number = valore fisso
project_number = "SNAM-DAM"

# 5) BuildingDescription = colonna G
idxG = col_letter_to_index('G')
building_desc = str(cols_allegato[idxG][sel]).strip()

# 6) IfcDescription = valore fisso
ifc_desc = "SNAM - Digitalizzazione Patrimonio"

# 7) SiteDescription = parte di doc.Title tra il 3 trattino e il successivo '_'
def extract_site_desc(txt):
    count = 0
    start = 0
    for i, ch in enumerate(txt):
        if ch == '-':
            count += 1
            if count == 3:
                start = i + 1
                break
    idx_us = txt.find('_', start)
    if idx_us != -1:
        return txt[start:idx_us].strip()
    return txt[start:].strip()
site_desc = extract_site_desc(title)

# 8) SiteLandTitleNumber = valore fisso
site_land_title = "[EPSG: 6875]"

# 9) SiteLongName = colonna E
idxE = col_letter_to_index('E')
site_long = str(cols_allegato[idxE][sel]).strip()

# 10) SiteName = tutto doc.Title
site_name = title

# 11) BuildingLongName = colonna J + " - " + colonna K
idxJ = col_letter_to_index('J')
idxK = col_letter_to_index('K')
building_long_name = str(cols_allegato[idxJ][sel]).strip() + " - " + str(cols_allegato[idxK][sel]).strip()

# Mappa BuiltInParameter valore (solo i 4 built-in richiesti)
param_map = {
    BuiltInParameter.PROJECT_BUILDING_NAME: building_name,
    BuiltInParameter.PROJECT_STATUS:        project_status,
    BuiltInParameter.PROJECT_NAME:          project_name,
    BuiltInParameter.PROJECT_NUMBER:        project_number,
}

# Recupera elemento ProjectInformation
proj_info = FilteredElementCollector(doc) \
    .OfClass(ProjectInfo) \
    .FirstElement()

# Transazione per impostare i parametri
trans = Transaction(doc, 'Set Parametri Progetto')
trans.Start()

# Imposta solo i built-in via get_Parameter
for bip, val in param_map.items():
    prm = proj_info.get_Parameter(bip)
    if prm and not prm.IsReadOnly:
        prm.Set(val)

# Imposta tutti gli altri parametri via LookupParameter
def set_lookup(name, value):
    p = proj_info.LookupParameter(name)
    if p and not p.IsReadOnly:
        p.Set(value)

set_lookup('BuildingDescription',    building_desc)
set_lookup('IfcDescription',         ifc_desc)
set_lookup('SiteDescription',        site_desc)
set_lookup('SiteLandTitleNumber',    site_land_title)
set_lookup('SiteLongName',           site_long)
set_lookup('SiteName',               site_name)
set_lookup('BuildingLongName',       building_long_name)

trans.Commit()

TaskDialog.Show('Completato', 'Parametri di progetto aggiornati.')