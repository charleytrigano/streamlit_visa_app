# -*- coding: utf-8 -*-
from datetime import datetime
from pathlib import Path
from io import BytesIO
import pandas as pd
import requests

EXCEL_URL = "https://www.dropbox.com/scl/fi/2j7czthz1u8kvwcj4a411/Clients-BL.xlsx?rlkey=ziivmkj4jler3m49hl21hbj5n&st=x7wtd6gh&dl=1"
EXCEL_PATH = Path("Clients_BL_local.xlsx")

def _init_local_if_needed():
    if not EXCEL_PATH.exists():
        df_dossiers = pd.DataFrame(columns=[
            "Dossier N","Nom","Date","Montant total","Acompte 1","Date Acompte 1",
            "Dossier envoyé","Date envoi","Escrow"
        ])
        df_escrow = pd.DataFrame(columns=[
            "Dossier N","Nom","Montant","Date envoi","État","Date réclamation"
        ])
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            df_dossiers.to_excel(writer, index=False, sheet_name="Dossiers")
            df_escrow.to_excel(writer, index=False, sheet_name="Escrow")

def load_data():
    """Lit le fichier Excel (Dropbox ou local) et détecte automatiquement les feuilles."""
    try:
        r = requests.get(EXCEL_URL, timeout=30)
        r.raise_for_status()
        xls = pd.ExcelFile(BytesIO(r.content))
        print("✅ Lecture Dropbox OK")
    except Exception as e:
        print("⚠️ Lecture Dropbox échouée :", e)
        _init_local_if_needed()
        xls = pd.ExcelFile(EXCEL_PATH)

    print("📄 Feuilles trouvées :", xls.sheet_names)

    # Détection automatique (ignore majuscules et espaces)
    def find_sheet(name_hint):
        for sheet in xls.sheet_names:
            if sheet.strip().lower() == name_hint.lower():
                return sheet
        raise ValueError(f"Feuille '{name_hint}' introuvable. Feuilles disponibles : {xls.sheet_names}")

    sheet_dossiers = find_sheet("Dossiers")
    sheet_escrow = find_sheet("Escrow")

    df_dossiers = pd.read_excel(xls, sheet_dossiers)
    df_escrow = pd.read_excel(xls, sheet_escrow)
    return df_dossiers, df_escrow

def save_data(df_dossiers, df_escrow):
    out = Path("Clients_BL_local_save.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_dossiers.to_excel(writer, index=False, sheet_name="Dossiers")
        df_escrow.to_excel(writer, index=False, sheet_name="Escrow")
    print("💾 Sauvegarde locale :", out)
    return str(out)

def add_dossier(df_dossiers, df_escrow, dossier):
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
        df_escrow = pd.concat([df_escrow, pd.DataFrame([new_esc])], ignore_index=True)
    save_data(df_dossiers, df_escrow)
    return df_dossiers, df_escrow

def update_dossier(df_dossiers, df_escrow, dossier_num, updates):
    idx = df_dossiers.index[df_dossiers["Dossier N"].astype(str) == str(dossier_num)]
    if len(idx) == 0:
        return df_dossiers, df_escrow, False
    i = idx[0]
    for k, v in updates.items():
        df_dossiers.at[i, k] = v
    row = df_dossiers.loc[i]
    escrow_flag = int(row.get("Escrow", 0)) if pd.notna(row.get("Escrow", 0)) else 0
    sent_flag = int(row.get("Dossier envoyé", 0)) if pd.notna(row.get("Dossier envoyé", 0)) else 0
    if escrow_flag == 1 and sent_flag == 1:
        ex = df_escrow.index[df_escrow["Dossier N"].astype(str) == str(dossier_num)]
        if len(ex) == 0:
            df_escrow = pd.concat([df_escrow, pd.DataFrame([{
                "Dossier N": row.get("Dossier N"),
                "Nom": row.get("Nom"),
                "Montant": row.get("Acompte 1"),
                "Date envoi": row.get("Date envoi", ""),
                "État": "À réclamer",
                "Date réclamation": ""
            }])], ignore_index=True)
        else:
            j = ex[0]
            if df_escrow.at[j, "État"] == "En attente":
                df_escrow.at[j, "État"] = "À réclamer"
            if pd.isna(df_escrow.at[j, "Date envoi"]) or df_escrow.at[j, "Date envoi"] == "":
                df_escrow.at[j, "Date envoi"] = row.get("Date envoi", "")
    save_data(df_dossiers, df_escrow)
    return df_dossiers, df_escrow, True

def mark_reclaimed(df_escrow, dossier_num):
    idx = df_escrow.index[df_escrow["Dossier N"].astype(str) == str(dossier_num)]
    if len(idx):
        j = idx[0]
        df_escrow.at[j, "État"] = "Réclamé"
        df_escrow.at[j, "Date réclamation"] = datetime.now().strftime("%Y-%m-%d")
    return df_escrow

def a_reclamer(df_escrow):
    if "État" not in df_escrow.columns:
        return df_escrow.iloc[0:0]
    return df_escrow[df_escrow["État"].fillna("") == "À réclamer"]

def reclames(df_escrow):
    if "État" not in df_escrow.columns:
        return df_escrow.iloc[0:0]
    return df_escrow[df_escrow["État"].fillna("") == "Réclamé"]
