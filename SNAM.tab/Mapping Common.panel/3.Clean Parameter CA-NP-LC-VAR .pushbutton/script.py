# -*- coding: utf-8 -*-
"""
Svuota i parametri con prefissi specifici, CA-NP-LC-VAR
"""
__title__ = 'Clean Parameter\nCA-NC-LC-VAR'
__author__ = 'Valerio Mascia'

from pyrevit import revit, DB, script, forms

# Documento Revit attivo
doc = revit.doc

official_prefixes = ("CA", "NP", "LC", "VAR")  # prefissi CASE-SENSITIVE


def has_content(param):
    """Verifica se un parametro di tipo stringa è valorizzato"""
    try:
        val = param.AsString()  # restituisce None se non valorizzato
        return bool(val and val.strip())
    except:
        return False  # se il parametro non è stringa o altro errore


def clear_prefixed_params(el, log_list):
    """
    Per ogni parametro dell'elemento, se il nome inizia con uno dei prefissi,
    è modificabile e contiene testo, lo svuota.
    """
    for p in el.Parameters:
        try:
            name = (p.Definition.Name or "").strip()
            # Salta se non inizia con prefisso, è di sola lettura o vuoto
            if (not any(name.startswith(pref) for pref in official_prefixes)
                    or p.IsReadOnly
                    or not has_content(p)):
                continue
            # Reset del valore stringa
            p.Set("")
            log_list.append((el.Id.IntegerValue, name))
        except:
            # Ignora errori su parametri non compatibili
            continue


# --- RACCOLTA ELEMENTI E TIPO ----------------------------------------------
all_instances = DB.FilteredElementCollector(doc) \
    .WhereElementIsNotElementType() \
    .ToElements()
all_types = DB.FilteredElementCollector(doc) \
    .WhereElementIsElementType() \
    .ToElements()

# Lista per tracciare i parametri svuotati
cleared_log = []  # tuple: (elementId, paramName)

# Esecuzione della transazione pyRevit
with revit.Transaction("Clear Prefixed Parameters"):
    # Svuota sui singoli elementi
    for inst in all_instances:
        clear_prefixed_params(inst, cleared_log)
    # Svuota sui tipi di elemento
    for tp in all_types:
        clear_prefixed_params(tp, cleared_log)

# Mostra risultato all'utente
message = u"Parametri svuotati: {}".format(len(cleared_log))
forms.alert(message, title="Clear Prefixed Params", ok=True)

# Log nel pannello di output di pyRevit (facoltativo)
script.get_logger().info(message)