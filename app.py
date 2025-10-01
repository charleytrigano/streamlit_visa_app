# app.py — Version finale avec CRUD complet et indexation ultra-sécurisée
import json
from datetime import datetime, date
import pandas as pd
import streamlit as st
import numpy as np 

# Importer les utilitaires (assurez-vous que utils.py est à jour)
from utils import (
    load_all_sheets,
    to_excel_bytes_multi,
    compute_finances,
    validate_rfe_row,
    harmonize_clients_df,
    _norm_cols 
)

st.set_page_config(page_title="Visa App", page_icon="🛂", layout="wide")

# Clear cache via URL param ?clear=1
try:
    params = st.query_params
    clear_val = params.get("clear", "0")
    if isinstance(clear_val, list):
        clear_val = clear_val[0]
    if clear_val == "1":
        st.cache_data.clear()
        try:
            st.cache_resource.clear()
        except Exception:
            pass
        st.query_params.clear()
        st.rerun()
except Exception:
    pass

# --- 1. FONCTIONS DE GESTION DES DONNÉES (INITIALISATION ET UTILITAIRES) ---

def initialize_session_state(all_sheets):
    """Charge les DataFrames initiaux et les stocke dans st.session_state."""
    
    # ------------------- DONNÉES CLIENTS -------------------
    clients_df_loaded = all_sheets.get("Clients")
    
    if clients_df_loaded is None:
        # Recherche robuste pour "Clients"
        clients_key = next((k for k in all_sheets.keys() if "client" in str(k).lower()), None)
        if clients_key:
            clients_df_loaded = all_sheets.get(clients_key)
            if clients_df_loaded is not None:
                st.info(f"Onglet 'Clients' non trouvé. Utilisation de l'onglet '{clients_key}'.")

    # Colonnes de base pour les clients
    base_cols_clients = [
        "DossierID", "DateCreation", "Nom", "TypeVisa", "Telephone", "Email",
        "DateFacture", "Honoraires", "Solde", "DateEnvoi", "Dossier envoyé",
        "DateRetour", "Dossier refusé", "Dossier approuvé", "RFE",
        "DateAnnulation", "DossierAnnule", "Notes", "Paiements" 
    ]
    
    if clients_df_loaded is None or clients_df_loaded.empty:
        if clients_df_loaded is None:
            st.warning("Onglet Clients introuvable ou illisible. Création d'un DataFrame Clients vide.")
        clients_df_loaded = pd.DataFrame(columns=base_cols_clients)
    else:
        clients_df_loaded = harmonize_clients_df(clients_df_loaded) 
        for c in base_cols_clients:
            if c not in clients_df_loaded.columns:
                 # Initialisation par défaut basée sur le type attendu
                default_val = pd.NaT if c.startswith("Date") else (0.0 if c == "Honoraires" else ([] if c == "Paiements" else ""))
                clients_df_loaded[c] = default_val

    # Finalisation des types de colonnes (Dates)
    date_cols = [c for c in base_cols_clients if c.startswith("Date")]
    for col in date_cols:
        clients_df_loaded[col] = pd.to_datetime(clients_df_loaded.get(col), errors='coerce')
            
    st.session_state.clients_df = compute_finances(clients_df_loaded.copy())
    st.session_state.clients_df = st.session_state.clients_df.astype({"Paiements": object})

    # ------------------- DONNÉES VISA -------------------
    visa_df_loaded = all_sheets.get("Visa")

    base_cols_visa = ["Categories", "Visa", "Definition"]

    if visa_df_loaded is None or visa_df_loaded.empty:
        if visa_df_loaded is None:
            st.warning("Onglet Visa introuvable ou illisible. Création d'un DataFrame Visa vide.")
        st.session_state.visa_df = pd.DataFrame(columns=base_cols_visa)
    else:
        # Nettoyage de colonnes pour Visa (norm_cols est déjà dans load_all_sheets)
        visa_df_loaded = visa_df_loaded.rename(columns={c:c for c in visa_df_loaded.columns if c in base_cols_visa})
        # S'assurer que les colonnes existent
        for col in base_cols_visa:
            if col not in visa_df_loaded.columns:
                visa_df_loaded[col] = ""
        st.session_state.visa_df = visa_df_loaded[base_cols_visa].copy()

    # Initialisation de l'index pour le formulaire (utiliser un index propre)
    if not st.session_state.visa_df.index.name:
        st.session_state.visa_df = st.session_state.visa_df.reset_index(drop=True)


def get_date_for_input(col_name, row):
    """Fonction utilitaire pour formatter les dates pour les date_input de Streamlit."""
    dt = row.get(col_name)
    if pd.notna(dt) and isinstance(dt, (datetime, date, pd.Timestamp)):
        return dt.date()
    return date.today()

# --- 2. LOGIQUE PRINCIPALE DE L'APPLICATION ---

# Sidebar / source / save options
with st.sidebar:
    st.header("Fichier source & sauvegarde")
    up = st.file_uploader("Fichier .xlsx", type=["xlsx"], help="Classeur contenant 'Visa' et 'Clients'.")
    data_path = st.text_input("Ou chemin local vers le .xlsx (optionnel)") 
    st.markdown("---")
    st.subheader("Sauvegarde")
    save_mode = st.selectbox("Mode de sauvegarde", ["Download (toujours disponible)", "Save to local path (serveur/PC)", "Google Drive (secrets req.)", "OneDrive (secrets req.)"])
    save_path = st.text_input("Chemin local pour sauvegarde (si Save to local path)", value="data_sauvegardee.xlsx")
    st.markdown("---")
    st.info("Navigation : utilisez le menu en bas pour basculer entre Clients et Visa")

src = data_path if data_path.strip() else up
if not src:
    st.info("Chargez un fichier ou renseignez un chemin local pour commencer.")
    st.stop()

# load sheets
if "clients_df" not in st.session_state:
    try:
        with st.spinner("Chargement et nettoyage des données..."):
            all_sheets, _ = load_all_sheets(src)
            initialize_session_state(all_sheets)
    except Exception as e:
        st.error(f"Erreur lecture fichier: {e}")
        st.stop()
    st.rerun()

# Navigation principale
page = st.selectbox("Page", ["Clients", "Visa"], index=0)

# --- 3. RENDU DES PAGES ---

if page == "Clients":
    
    st.header("👥 Clients — gestion & suivi")
    df = st.session_state.clients_df
    
    # Navigation CRUD dans la page Clients
    crud_mode = st.radio("Action Clients", ["Lister/Modifier/Supprimer", "Ajouter un nouveau dossier"], index=0, horizontal=True)

    if crud_mode == "Ajouter un nouveau dossier":
        st.subheader("Ajouter un nouveau dossier client")
        
        # Créer une ligne vide pour l'ajout
        empty_row = pd.Series("", index=df.columns)
        empty_row["Paiements"] = [] 
        
        render_client_form(df, empty_row, action="add")

    elif crud_mode == "Lister/Modifier/Supprimer":
        
        # KPIs
        total_dossiers = len(df)
        total_encaissé = df["TotalAcomptes"].sum()
        total_honoraires = df["Honoraires"].sum()
        total_solde = df["SoldeCalc"].sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total dossiers", f"{total_dossiers:,}")
        c2.metric("Total encaissé", f"{total_encaissé:,.2f} €")
        c3.metric("Total honoraires", f"{total_honoraires:,.2f} €")
        c4.metric("Solde total", f"{total_solde:,.2f} €")

        # Filtrage
        with st.expander("Filtrer / Rechercher"):
            q = st.text_input("Recherche (nom / dossier / email)")
            status_filter = st.selectbox("Filtrer par statut", ["Tous", "Envoyé", "Approuvé", "Refusé", "Annulé", "RFE"])

        filtered = df.copy()
        if q:
            mask = pd.Series(False, index=filtered.index)
            for c in ["DossierID", "Nom", "Email", "TypeVisa"]:
                if c in filtered.columns:
                    mask = mask | filtered[c].astype(str).str.contains(q, case=False, na=False)
            filtered = filtered[mask]
        
        if status_filter != "Tous":
            col_map = {"Envoyé": "Dossier envoyé", "Approuvé": "Dossier approuvé", "Refusé": "Dossier refusé", "Annulé": "DossierAnnule", "RFE": "RFE"}
            col_name = col_map.get(status_filter)
            if col_name:
                 filtered = filtered[filtered.get(col_name, False) == True]


        st.dataframe(filtered.reset_index(drop=True), use_container_width=True)

        # Sélection et modification
        if len(filtered) > 0:
            max_idx = len(filtered) - 1
            
            # --- CORRECTION FINALE DE L'INDEXATION ROBUSTE ---
            # 1. Gérer la valeur de l'index dans la session pour survivre aux filtres
            if 'client_sel_idx' not in st.session_state:
                st.session_state.client_sel_idx = 0
            
            current_value = st.session_state.client_sel_idx
            # Assurer que l'index n'est pas hors bornes après un filtre
            if current_value > max_idx:
                 current_value = 0 
            
            # L'utilisateur choisit l'index affiché dans le DF filtré (0 à max_idx)
            sel_idx = st.number_input("Ouvrir dossier (index affiché)", min_value=0, max_value=max_idx, value=current_value)
            
            # Mise à jour de la valeur dans la session
            st.session_state.client_sel_idx = int(sel_idx)

            # 2. RAPPEL: Utiliser .iloc[position] pour obtenir la ligne, puis .name pour l'index original
            # Cela garantit la fiabilité en cas de changement de filtre
            sel_row_filtered = filtered.iloc[int(sel_idx)]
            original_session_index = sel_row_filtered.name # C'est l'index dans df = st.session_state.clients_df

            st.subheader(f"Modifier Dossier: {sel_row_filtered.get('DossierID','(sans id)')} — {sel_row_filtered.get('Nom','')}")
            
            # Ligne 229 de votre rapport précédent : l'appel à la fonction
            render_client_form(df, sel_row_filtered, action="update", original_index=original_session_index)
        else:
            st.info("Aucun dossier client ne correspond aux filtres.")


elif page == "Visa":
    st.header("🛂 Visa — Gestion des types")
    df = st.session_state.visa_df
    
    # Navigation CRUD dans la page Visa
    crud_mode = st.radio("Action Visa", ["Lister/Modifier/Supprimer", "Ajouter un nouveau type"], index=0, horizontal=True)

    if crud_mode == "Ajouter un nouveau type":
        st.subheader("Ajouter un nouveau type de visa")
        empty_row = pd.Series("", index=df.columns)
        render_visa_form(df, empty_row, action="add")
        
    elif crud_mode == "Lister/Modifier/Supprimer":
        
        st.dataframe(df, use_container_width=True)
        
        if len(df) > 0:
            max_idx = len(df) - 1
            
            # --- Sécurisation de l'indexation pour Visa ---
            if 'visa_sel_idx' not in st.session_state:
                st.session_state.visa_sel_idx = 0
            
            current_value = st.session_state.visa_sel_idx
            if current_value > max_idx:
                 current_value = 0
                 
            sel_idx = st.number_input("Ouvrir visa (index affiché)", min_value=0, max_value=max_idx, value=current_value)
            st.session_state.visa_sel_idx = int(sel_idx)
            
            sel_row = df.iloc[int(sel_idx)]
            
            st.subheader(f"Modifier Visa: {sel_row.get('Visa', 'N/A')}")
            
            # L'index du dataframe est l'index de la série sel_row (pour le DF non filtré)
            render_visa_form(df, sel_row, action="update", original_index=int(sel_idx)) 
        else:
            st.info("Aucun type de visa à gérer.")
        
# --- 4. DEFINITION DES FORMULAIRES (CRUD) ---

def render_client_form(df, sel_row, action, original_index=None):
    """Rendu du formulaire d'ajout/modification/suppression pour un client."""
    
    is_add = (action == "add")
    button_label = "Ajouter le dossier" if is_add else "Enregistrer les modifications"

    with st.form(f"client_form_{action}"):
        
        # Corps du formulaire CLIENTS
        cols1, cols2 = st.columns(2)
        with cols1:
            # Pour l'ajout, l'ID doit être modifiable ; pour la modification, il est figé (car il sert d'identifiant)
            dossier_id = st.text_input("DossierID", value=sel_row.get("DossierID", ""), disabled=not is_add)
            nom = st.text_input("Nom", value=sel_row.get("Nom", ""))
            typevisa = st.text_input("TypeVisa", value=sel_row.get("TypeVisa", ""))
            email = st.text_input("Email", value=sel_row.get("Email", ""))
        with cols2:
            telephone = st.text_input("Telephone", value=sel_row.get("Telephone", ""))
            honoraires = st.number_input("Honoraires", value=float(sel_row.get("Honoraires", 0.0)), format="%.2f")
            notes = st.text_area("Notes", value=sel_row.get("Notes", ""))
        
        st.markdown("---")
        st.write("Statuts / dates")
        st_col1, st_col2, st_col3 = st.columns(3)
        
        # Récupération des valeurs booléennes
        envoye = bool(sel_row.get("Dossier envoyé", False))
        refuse = bool(sel_row.get("Dossier refusé", False))
        approuve = bool(sel_row.get("Dossier approuvé", False))
        annule = bool(sel_row.get("DossierAnnule", False))
        rfe_val = bool(sel_row.get("RFE", False))
        
        with st_col1:
            dossier_envoye = st.checkbox("Dossier envoyé", value=envoye)
            dossier_refuse = st.checkbox("Dossier refusé", value=refuse)
        with st_col2:
            dossier_approuve = st.checkbox("Dossier approuvé", value=approuve)
            dossier_annule = st.checkbox("DossierAnnule (annulé)", value=annule)
        with st_col3:
            rfe = st.checkbox("RFE (doit être combiné)", value=rfe_val)
            date_envoi = st.date_input("DateEnvoi", value=get_date_for_input("DateEnvoi", sel_row))

        st.markdown("---")
        
        payments_list = sel_row.get("Paiements", [])
        # S'assurer que payments_list est une liste propre
