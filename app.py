# app.py — version améliorée (modularité + harmonisation des colonnes)
import json
from datetime import datetime, date
import pandas as pd
import streamlit as st

# Importer les utilitaires depuis le nouveau fichier
from utils import (
    load_all_sheets,
    to_excel_bytes_multi,
    compute_finances,
    validate_rfe_row,
    harmonize_clients_df, # NOUVELLE FONCTION IMPORTÉE
    _norm_cols 
)

st.set_page_config(page_title="Visa App", page_icon="🛂", layout="wide")

# ... (Nettoyage du cache inchangé) ...

# ... (Sidebar inchangée) ...

src = data_path if data_path.strip() else up
if not src:
    st.info("Chargez un fichier ou renseignez un chemin local pour commencer.")
    st.stop()

# load sheets
try:
    with st.spinner("Chargement et nettoyage des données..."):
        all_sheets, sheet_names = load_all_sheets(src)
except Exception as e:
    st.error(f"Erreur lecture fichier: {e}")
    st.stop()

st.success(f"Onglets trouvés: {', '.join(sheet_names)}")

visa_df = all_sheets.get("Visa")
clients_df_loaded = all_sheets.get("Clients")

# Normalize and ensure base columns
base_cols = [
    "DossierID", "DateCreation", "Nom", "TypeVisa", "Telephone", "Email",
    "DateFacture", "Honoraires", "Solde", "DateEnvoi", "Dossier envoyé",
    "DateRetour", "Dossier refusé", "Dossier approuvé", "RFE",
    "DateAnnulation", "DossierAnnule", "Notes", "Paiements" 
]

if clients_df_loaded is None:
    clients_df_loaded = pd.DataFrame(columns=base_cols)
else:
    # --- ÉTAPE CRUCIALE : HARMONISATION DES DONNÉES ---
    clients_df_loaded = harmonize_clients_df(clients_df_loaded) 
    
    # S'assurer que les colonnes de base existent après l'harmonisation
    for c in base_cols:
        if c not in clients_df_loaded.columns:
            clients_df_loaded[c] = "" if c not in ["Honoraires", "DateCreation", "DateFacture", "DateEnvoi", "DateRetour", "DateAnnulation"] else (0.0 if c == "Honoraires" else pd.NaT)

# Initialisation de la session
if "clients_df" not in st.session_state:
    date_cols = ["DateCreation", "DateFacture", "DateEnvoi", "DateRetour", "DateAnnulation"]
    for col in date_cols:
        if col in clients_df_loaded.columns:
            clients_df_loaded[col] = pd.to_datetime(clients_df_loaded[col], errors='coerce')
            
    st.session_state.clients_df = clients_df_loaded.copy()

# ... (Le reste de app.py reste inchangé) ...





