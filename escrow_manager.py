# -*- coding: utf-8 -*-
from datetime import datetime
from pathlib import Path
import pandas as pd

EXCEL_FILE = "Clients BL.xlsx"
SHEET_DOSSIERS = "Dossiers"
SHEET_ESCROW = "Escrow"

def _init_workbook_if_needed():
    """Crée le fichier Excel s'il n'existe pas encore."""
    path = Path(EXCEL_FILE)
    if not path.exists():
        df_dossiers = pd.DataFrame(columns=[
            "Dossier N", "Nom", "Date", "Montant total", "Acompte 1",
            "Date Acompte 1", "Dossier envoyé", "Date envoi", "Escrow"
        ])
        df_escrow = pd.DataFrame(columns=[
            "Dossier N", "Nom", "Montant", "Date envoi", "État", "Date réclamation"
        ])
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df_dossiers.to_excel(w, index=False, sheet_name=SHEET_DOSSIERS)
            df_escrow.to_excel(w, index=False, sheet_name=SHEET_ESCROW)

def load_data():
    """Charge les deux feuilles du fichier Excel."""
    _init_workbook_if_needed()
    xls = pd.ExcelFile(EXCEL_FILE)
    df_dossiers = pd.read_excel(xls, SHEET_DOSSIERS)
    df_escrow = pd.read_excel(xls, SHEET_ESCROW)
    return df_dossiers, df_escrow

def save_data(df_dossiers, df_escrow):
    """Sauvegarde les deux DataFrames dans le fichier Excel."""
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as w:
        df_dossiers.to_excel(w, index=False, sheet_name=SHEET_DOSSIERS)
        df_escrow.to_excel(w, index=False, sheet_name=SHEET_ESCROW)

def add_dossier(df_dossiers, df_escrow, dossier):
    """Ajoute un dossier et, si Escrow, ajoute une ligne dans l’onglet Escrow."""
    df_dossiers = pd.concat([df_dossiers, pd.DataFrame([dossier])], ignore_index=True)
    if int(dossier.get("Escrow", 0)) == 1:
        new_esc = {
            "Dossier N": dossier.get("Dossier N"),
            "Nom": dossier.get("Nom"),
            "Montant": dossier.get("Acompte 1"),
            "Date envoi": dossier.get("Date envoi", ""),
            "État": "En attente",
            "Date réclamation": ""
        }
        df_escrow = pd.concat([df_escrow, pd.DataFrame([new]()_]()_
