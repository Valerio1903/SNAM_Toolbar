# -*- coding: utf-8 -*-
"""Compilazione parametri ExportAsIfc, ExportTypeIfc, IfcName, IfcObjectType"""
__title__ = 'Ifc Class\nName Object'
__author__ = 'Valerio Mascia'
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
import xlrd
from System.Collections.Generic import List
from System.Windows.Forms import OpenFileDialog
import System

# Funzione per selezionare un file Excel
def scegli_file_excel(titolo):
    dialog = OpenFileDialog()
    dialog.Title = titolo
    dialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx"
    dialog.Multiselect = False
    if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
        return dialog.FileName
    else:
        TaskDialog.Show("Errore", "Operazione annullata: file non selezionato.")
        raise SystemExit

# Documento corrente
doc = __revit__.ActiveUIDocument.Document

# === Selezione file Excel ===
IFCNAME_EXCEL = scegli_file_excel("Seleziona il file IfcName (Allegato 1)")
PLACE_EXCEL = scegli_file_excel("Seleziona il file PLACEHOLDER")
MAPPER_EXCEL = scegli_file_excel("Seleziona il file Allegato 3 - Mappatura")

# === Configurazione ===
IFCNAME_SHEET = "IfcName"
PLACE_SHEETS = ["FU", "SE"]
MAPPER_SHEETS = ["Elenco IM", "Elenco SE", "Elenco FU", "Elenco AP", "Elenco NO"]

PARAM_IFCNAME = "IfcName"
PARAM_OBJTYPE = "IfcObjectType"
PARAM_EXPORT = BuiltInParameter.IFC_EXPORT_ELEMENT_AS
PARAM_PREDEF = BuiltInParameter.IFC_EXPORT_PREDEFINEDTYPE

cat_builtin = List[BuiltInCategory]([
    BuiltInCategory.OST_DuctTerminal, BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_DuctInsulations, BuiltInCategory.OST_DuctLinings,
    BuiltInCategory.OST_ElectricalEquipment,BuiltInCategory.OST_Furniture,
    BuiltInCategory.OST_FlexDuctCurves, BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_PipeInsulations,
    BuiltInCategory.OST_FlexPipeCurves, BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_ConduitFitting,BuiltInCategory.OST_MechanicalEquipment
])

# Funzioni helper (xlrd)
def load_ifcname_rules(path, sheet):
    workbook = xlrd.open_workbook(path)
    worksheet = workbook.sheet_by_name(sheet)
    rules = []
    for row_idx in range(1, worksheet.nrows):
        val_a = str(worksheet.cell(row_idx, 0).value).strip()
        val_b = str(worksheet.cell(row_idx, 1).value).strip().upper()
        if val_a:
            rules.append((val_a, val_b))
    return rules

def load_placeholder(path, sheets):
    workbook = xlrd.open_workbook(path)
    lookup = {}
    for sh in sheets:
        worksheet = workbook.sheet_by_name(sh)
        for row_idx in range(1, worksheet.nrows):
            val = worksheet.cell(row_idx, 0).value
            if val:
                lookup[str(val).strip()[:5].upper()] = str(val).strip()
    return lookup

def load_im_rules(path, sheet):
    # IM sheet: col0=codice, col1=IfcName, col3=IfcObjectType
    workbook = xlrd.open_workbook(path)
    worksheet = workbook.sheet_by_name(sheet)
    lookup = {}
    for row_idx in range(1, worksheet.nrows):
        code = str(worksheet.cell(row_idx, 0).value or "").strip()
        ifcname = str(worksheet.cell(row_idx, 1).value or "").strip()
        if code and ifcname:
            lookup[code.upper()] = ifcname
    return lookup

def load_mapper_rules(path, sheets):
    workbook = xlrd.open_workbook(path)
    rules = {}
    for sh in sheets:
        worksheet = workbook.sheet_by_name(sh)
        for row_idx in range(1, worksheet.nrows):
            pref = str(worksheet.cell(row_idx, 0).value or "").strip()[:5].upper()
            obj = str(worksheet.cell(row_idx, 4).value or "").strip()
            exp = str(worksheet.cell(row_idx, 5).value or "").strip() or obj
            rules[pref] = (obj, exp)
    return rules

# Collector elementi Revit
filter = ElementMulticategoryFilter(cat_builtin)
collector = FilteredElementCollector(doc)\
    .WherePasses(filter)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Carica regole Excel (xls con xlrd)
rules_ifcname = load_ifcname_rules(IFCNAME_EXCEL, IFCNAME_SHEET)
im_lookup = load_im_rules(IFCNAME_EXCEL, "IM")
ph_lookup = load_placeholder(PLACE_EXCEL, PLACE_SHEETS)
map_rules = load_mapper_rules(MAPPER_EXCEL, MAPPER_SHEETS)

# Transazione
t = Transaction(doc, "Compila parametri IFC")
t.Start()

for e in collector:
    sym = doc.GetElement(e.GetTypeId())
    if not sym:
        continue

    fam_param = sym.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
    fam_name = fam_param.AsString() if fam_param else ""
    type_param = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
    type_name = type_param.AsString() if type_param else ""

    fam_name = fam_name or ""
    type_name = type_name or ""

    head5_fam = fam_name[:5].upper()
    head5_type = type_name[:5].upper()

    target_ifcname = None
    is_placeholder = False

    # Regola SNAM_
    if type_name[:5].startswith("SNAM_") or type_name[:5].startswith("Tubaz") or type_name[:5].startswith("BARRE") :
        for a, b in rules_ifcname:
            if b == "BARRE":
                target_ifcname = a
                break
        if target_ifcname:
            e.LookupParameter(PARAM_IFCNAME).Set(target_ifcname)
            e.LookupParameter(PARAM_OBJTYPE).Set("IfcFlowSegment")
            e.get_Parameter(PARAM_EXPORT).Set("IfcPipeSegmentType")
        continue

    # Regola NO025
    if head5_fam == "NO025":
        p_pre = e.get_Parameter(PARAM_PREDEF)
        if p_pre and not p_pre.IsReadOnly:
            p_pre.Set("BEND")
        target_ifcname = next((a for a, b in rules_ifcname if b == "NO025"), None)

    # Regole AP330/AP450 (aggiornato per intercettare nuove diciture)
    elif fam_name.startswith(("AP330", "AP450")) and (
        "con_comando_manuale_con_riduttore" in fam_name or
        "con_comando_manuale_a_leva" in fam_name or
        "con_comando_manuale_riduttore" in fam_name or
        "con_comando_manuale_con_chiave_a_T" in fam_name
    ):
        # Trova tutte le IFCName presenti nel nome famiglia e seleziona la piu lunga
        matches = [a for a, _ in rules_ifcname if a.strip() in fam_name]
        if matches:
            target_ifcname = max(matches, key=len)

    # Regole FU/SE (Generic Model Placeholder - da PLACEHOLDER.xlsx)
    elif head5_type in ph_lookup:
        is_placeholder = True
        target_ifcname = ph_lookup[head5_type]

    # Regole IM (Generic Model Placeholder - da IfcName Allegato 1, sheet IM)
    elif head5_type in im_lookup:
        is_placeholder = True
        target_ifcname = im_lookup[head5_type]

    # Regola standard Allegato1: prova prima type_name, poi family_name
    else:
        target_ifcname = next((a for a, b in rules_ifcname if b == head5_type), None)
        if not target_ifcname:
            target_ifcname = next((a for a, b in rules_ifcname if b == head5_fam), None)

    if target_ifcname:
        p_ifc = e.LookupParameter(PARAM_IFCNAME)
        if p_ifc and not p_ifc.IsReadOnly:
            p_ifc.Set(target_ifcname)

    # Mapping Allegato3: per IM/SE/FU usa type_name, altrimenti family_name
    mapper_key = head5_type if is_placeholder else head5_fam
    obj_exp = map_rules.get(mapper_key)
    if obj_exp:
        obj, exp = obj_exp
        p_obj = e.LookupParameter(PARAM_OBJTYPE)
        p_exp = e.get_Parameter(PARAM_EXPORT)
        if p_obj and not p_obj.IsReadOnly:
            p_obj.Set(obj)
        if p_exp and not p_exp.IsReadOnly:
            p_exp.Set(exp)

t.Commit()

TaskDialog.Show("IFC Mapping", "Tutti i parametri IFC compilati correttamente.")
