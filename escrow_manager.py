# -*- coding: utf-8 -*-
"""
Module Escrow Manager — Gestion autonome des dossiers Escrow
Fichier de données : "Clients BL.xlsx"
Testable seul avec Streamlit :
    streamlit run escrow_manager.py
"""

import pandas as pd
import streamlit as st
from datetime import datetime
from pathlib import Path

EXCEL_FILE = "Clients BL.xlsx"
SHEET_DOSSIERS = "Dossiers"
SHEET_ESCROW = "Escrow"

# ============================================================
# 🔹 Chargement et sauvegarde automatique du fichier Excel
# ============================================================

def load_data():
    """Charge ou initialise le fichier Excel."""
    path = Path(EXCEL_FILE)
    if not path.exists():
        # Initialisation si le fichier n'existe pas encore
        df_dossiers = pd.DataFrame(columns=[
            "Dossier N", "Nom", "Date", "Acompte 1", "Date Acompte 1",
            "Dossier envoyé", "Date envoi", "Escrow"
        ])
        df_escrow = pd.DataFrame(columns=[
            "Dossier N", "Nom", "Montant", "Date envoi", "État", "Date réclamation"
        ])
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df_dossiers.to_excel(writer, index=False, sheet_name=SHEET_DOSSIERS)
            df_escrow.to_excel(writer, index=False, sheet_name=SHEET_ESCROW)
    xls = pd.ExcelFile(path)
    df_dossiers = pd.read_excel(xls, SHEET_DOSSIERS)
    df_escrow = pd.read_excel(xls, SHEET_ESCROW)
    return df_dossiers, df_escrow


def save_data(df_dossiers, df_escrow):
    """Sauvegarde automatique dans le fichier Excel."""
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_dossiers.to_excel(writer, index=False, sheet_name=SHEET_DOSSIERS)
        df_escrow.to_excel(writer, index=False, sheet_name=SHEET_ESCROW)


# ============================================================
# 🔹 Gestion Escrow
# ============================================================

def add_dossier(df_dossiers, df_escrow, dossier):
    """Ajoute un dossier avec gestion Escrow."""
    df_dossiers = pd.concat([df_dossiers, pd.DataFrame([dossier])], ignore_index=True)
    if dossier.get("Escrow", 0) == 1:
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


def update_dossier(df_dossiers, df_escrow, dossier_id, new_data):
    """Met à jour un dossier et crée une alerte à réclamer si envoi."""
    if dossier_id not in df_dossiers.index:
        st.warning("Dossier non trouvé.")
        return df_dossiers, df_escrow
    for k, v in new_data.items():
        df_dossiers.at[dossier_id, k] = v

    # Si le dossier est envoyé et lié à Escrow → crée une alerte à réclamer
    row = df_dossiers.loc[dossier_id]
    if row.get("Escrow", 0) == 1 and row.get("Dossier envoyé") == 1:
        dossier_num = row.get("Dossier N")
        if not (df_escrow["Dossier N"] == dossier_num).any():
            df_escrow = pd.concat([df_escrow, pd.DataFrame([{
                "Dossier N": dossier_num,
                "Nom": row.get("Nom"),
                "Montant": row.get("Acompte 1"),
                "Date envoi": row.get("Date envoi", ""),
                "État": "À réclamer",
                "Date réclamation": ""
            }])], ignore_index=True)
    save_data(df_dossiers, df_escrow)
    return df_dossiers, df_escrow


def mark_reclaimed(df_escrow, dossier_num):
    """Marque un Escrow comme réclamé."""
    idx = df_escrow.index[df_escrow["Dossier N"] == dossier_num]
    if len(idx) > 0:
        df_escrow.loc[idx, "État"] = "Réclamé"
        df_escrow.loc[idx, "Date réclamation"] = datetime.now().strftime("%Y-%m-%d")
    return df_escrow


def get_escrow_status(df_escrow):
    """Retourne les Escrows à réclamer et réclamés."""
    a_reclamer = df_escrow[df_escrow["État"] == "À réclamer"]
    reclames = df_escrow[df_escrow["État"] == "Réclamé"]
    return a_reclamer, reclames


def show_alert(df_escrow):
    """Affiche une alerte visuelle si Escrows à réclamer."""
    nb = len(df_escrow[df_escrow["État"] == "À réclamer"])
    if nb > 0:
        st.error(f"⚠️ {nb} dossier(s) Escrow à réclamer ! Vérifiez l’onglet 🛡️ Escrow.")
    else:
        st.success("✅ Aucun Escrow en attente de réclamation.")


# ============================================================
# 🔹 Interface Streamlit de test
# ============================================================

def main():
    st.set_page_config(page_title="Escrow Manager", page_icon="🛡️", layout="wide")
    st.title("🛡️ Gestion des Escrows")

    df_dossiers, df_escrow = load_data()
    show_alert(df_escrow)

    st.subheader("📂 Dossiers Escrow")
    st.dataframe(df_escrow, use_container_width=True)

    st.divider()
    st.subheader("🧾 Marquer un Escrow comme réclamé")
    dossier_num = st.text_input("Numéro de dossier à marquer comme réclamé")
    if st.button("✅ Marquer comme réclamé"):
        df_escrow = mark_reclaimed(df_escrow, dossier_num)
        save_data(df_dossiers, df_escrow)
        st.success(f"Dossier {dossier_num} marqué comme réclamé.")
        st.experimental_rerun()


if __name__ == "__main__":
    main()
