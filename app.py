# =========================
# Berenbaum Law — Version Complète et Corrigée
# =========================

import os
import json
import re
import io
import zipfile
import uuid
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional
from pathlib import Path

import pandas as pd
import streamlit as st

from datetime import date, datetime
import pandas as pd


# =========================
# Fonctions utilitaires
# =========================

def _date_for_widget(val):
    """Convertit proprement une valeur Excel/pandas en date utilisable dans Streamlit."""
    if val is None or pd.isna(val):
        return date.today()
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return date.today()
        return d.date()
    except Exception:
        return date.today()


# =========================
# Constantes et configuration
# =========================

APP_TITLE = "🛂 Visa Manager"

COLS_CLIENTS = [
    "ID_Client", "Date de création", "Nom", "Prénom", "Nationalité",
    "Type de visa", "Statut dossier", "Date RDV", "Date dépôt",
    "Date résultat", "Résultat", "Date retour passeport",
    "Montant total", "Montant payé", "Solde restant",
    "Email", "Téléphone", "Observations"
]

COLS_PAIEMENTS = [
    "ID_Paiement", "ID_Client", "Date paiement", "Montant", "Mode de paiement", "Commentaire"
]


# =========================
# Chargement et fusion fichiers Excel
# =========================

def charger_fichiers_excel(uploaded_files):
    """Charge un ou deux fichiers Excel et fusionne si nécessaire."""
    dfs_clients, dfs_paiements = [], []
    last_name = None

    for up in uploaded_files:
        last_name = up.name
        data = up.read()
        with pd.ExcelFile(BytesIO(data)) as xls:
            if "Clients" in xls.sheet_names:
                dfc = pd.read_excel(xls, "Clients")
            else:
                dfc = pd.DataFrame(columns=COLS_CLIENTS)
            if "Paiements" in xls.sheet_names:
                dfp = pd.read_excel(xls, "Paiements")
            else:
                dfp = pd.DataFrame(columns=COLS_PAIEMENTS)
        dfs_clients.append(dfc)
        dfs_paiements.append(dfp)

    df_clients = pd.concat(dfs_clients, ignore_index=True) if dfs_clients else pd.DataFrame(columns=COLS_CLIENTS)
    df_paiements = pd.concat(dfs_paiements, ignore_index=True) if dfs_paiements else pd.DataFrame(columns=COLS_PAIEMENTS)

    return df_clients, df_paiements, last_name or "visa_manager.xlsx"


# =========================
# Fonctions de calcul et sauvegarde
# =========================

def recalculer_soldes(df_clients, df_paiements):
    """Met à jour les montants payés et soldes à partir des paiements."""
    if "ID_Client" not in df_clients.columns or "ID_Client" not in df_paiements.columns:
        return df_clients
    paiements_sum = df_paiements.groupby("ID_Client")["Montant"].sum().reset_index(name="Total payé")
    df = df_clients.merge(paiements_sum, on="ID_Client", how="left")
    df["Total payé"] = df["Total payé"].fillna(0)
    df["Montant total"] = pd.to_numeric(df["Montant total"], errors="coerce").fillna(0)
    df["Solde restant"] = df["Montant total"] - df["Total payé"]
    return df


def sauvegarder_excel(df_clients, df_paiements):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_clients.to_excel(writer, index=False, sheet_name="Clients")
        df_paiements.to_excel(writer, index=False, sheet_name="Paiements")
    return output.getvalue()


# =========================
# Interface utilisateur
# =========================

def page_clients(df_clients, df_paiements):
    st.subheader("👥 Gestion des clients")

    # Création client
    with st.expander("➕ Ajouter un client"):
        c1, c2 = st.columns(2)
        with c1:
            nom = st.text_input("Nom")
            prenom = st.text_input("Prénom")
            nationalite = st.text_input("Nationalité")
            type_visa = st.text_input("Type de visa")
            statut = st.text_input("Statut dossier")
        with c2:
            date_creation = st.date_input("Date de création", value=_date_for_widget(date.today()))
            date_rdv = st.date_input("Date RDV", value=_date_for_widget(date.today()))
            date_depot = st.date_input("Date dépôt", value=_date_for_widget(date.today()))
            date_resultat = st.date_input("Date résultat", value=_date_for_widget(date.today()))
            resultat = st.text_input("Résultat")
        montant_total = st.number_input("Montant total", min_value=0.0, step=100.0)
        email = st.text_input("Email")
        telephone = st.text_input("Téléphone")
        observations = st.text_area("Observations")

        if st.button("Ajouter le client"):
            new_id = str(uuid.uuid4())
            new_row = {
                "ID_Client": new_id,
                "Date de création": date_creation,
                "Nom": nom,
                "Prénom": prenom,
                "Nationalité": nationalite,
                "Type de visa": type_visa,
                "Statut dossier": statut,
                "Date RDV": date_rdv,
                "Date dépôt": date_depot,
                "Date résultat": date_resultat,
                "Résultat": resultat,
                "Montant total": montant_total,
                "Montant payé": 0.0,
                "Solde restant": montant_total,
                "Email": email,
                "Téléphone": telephone,
                "Observations": observations,
            }
            df_clients = pd.concat([df_clients, pd.DataFrame([new_row])], ignore_index=True)
            st.success("Client ajouté avec succès.")

    # Liste clients
    if len(df_clients) > 0:
        st.dataframe(df_clients)
    else:
        st.info("Aucun client enregistré.")

    return df_clients


def page_paiements(df_clients, df_paiements):
    st.subheader("💳 Gestion des paiements")

    if len(df_clients) == 0:
        st.warning("Aucun client disponible.")
        return df_clients, df_paiements

    # Sélection client
    noms_clients = df_clients["Nom"].astype(str) + " " + df_clients["Prénom"].astype(str)
    selected = st.selectbox("Choisir un client", options=noms_clients)
    idxs = df_clients.index[noms_clients == selected]
    if len(idxs) == 0:
        st.warning("Client introuvable.")
        return df_clients, df_paiements
    idx = idxs[0]
    id_client = df_clients.loc[idx, "ID_Client"]

    st.write(f"**Montant total :** {df_clients.loc[idx, 'Montant total']} USD")
    st.write(f"**Solde restant :** {df_clients.loc[idx, 'Solde restant']} USD")

    # Nouveau paiement
    with st.form("form_paiement"):
        date_p = st.date_input("Date paiement", value=_date_for_widget(date.today()))
        montant = st.number_input("Montant", min_value=0.0, step=10.0)
        mode = st.selectbox("Mode de paiement", ["Espèces", "Carte", "Virement"])
        commentaire = st.text_input("Commentaire")
        submit = st.form_submit_button("Ajouter paiement")

    if submit:
        new_payment = {
            "ID_Paiement": str(uuid.uuid4()),
            "ID_Client": id_client,
            "Date paiement": date_p,
            "Montant": montant,
            "Mode de paiement": mode,
            "Commentaire": commentaire
        }
        df_paiements = pd.concat([df_paiements, pd.DataFrame([new_payment])], ignore_index=True)
        df_clients = recalculer_soldes(df_clients, df_paiements)
        st.success("Paiement ajouté avec succès.")

    # Historique paiements
    hist = df_paiements[df_paiements["ID_Client"] == id_client]
    st.dataframe(hist)

    return df_clients, df_paiements


def page_analyses(df_clients):
    st.subheader("📊 Analyses")
    if len(df_clients) == 0:
        st.info("Aucune donnée à analyser.")
        return
    # Nettoyage des types pour éviter int(year) sur NaN
    df_clients["Année"] = pd.to_datetime(df_clients["Date de création"], errors="coerce").dt.year.fillna(date.today().year).astype(int)
    grouped = df_clients.groupby("Année")["ID_Client"].count().reset_index(name="Nombre de dossiers")
    st.bar_chart(grouped, x="Année", y="Nombre de dossiers")


# =========================
# Application principale
# =========================

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    uploaded_files = st.file_uploader("Charger un ou deux fichiers Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    if not uploaded_files:
        st.info("Veuillez charger au moins un fichier Excel.")
        return

    df_clients, df_paiements, fichier_nom = charger_fichiers_excel(uploaded_files)
    df_clients = recalculer_soldes(df_clients, df_paiements)

    onglet = st.tabs(["📋 Clients", "💳 Paiements", "📊 Analyses"])

    with onglet[0]:
        df_clients = page_clients(df_clients, df_paiements)
    with onglet[1]:
        df_clients, df_paiements = page_paiements(df_clients, df_paiements)
    with onglet[2]:
        page_analyses(df_clients)

    # Sauvegarde
    data_xlsx = sauvegarder_excel(df_clients, df_paiements)
    st.download_button("📥 Télécharger le fichier mis à jour", data=data_xlsx, file_name=fichier_nom, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    main()
