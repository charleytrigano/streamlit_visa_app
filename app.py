import io
import re
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# =========================
# Constantes et Configuration
# =========================
APP_TITLE = "🛂 Visa Manager - Gestion Complète"
SID = "vmgr_v5"

# Dictionnaire du modèle de classification (pour la saisie de nouveaux dossiers)
VISA_STRUCTURE = {
    "Affaires / Tourisme": {
        "B-1": ["COS", "EOS"],
        "B-2": ["COS", "EOS"],
    },
    "Etudiants": {
        "F-1": ["COS", "EOS"],
        "F-2": ["COS", "EOS"],
    },
    "Treaty": {
        "E-2": {
            "E-2 Inv.": ["CP", "USCIS"],
            "E-2 Inv. Ren.": ["CP", "USCIS"],
            "E-2 ESE": ["CP", "USCIS"],
            "E-2 ESE Ren.": ["CP", "USCIS"],
        }
    },
    "Trader": {
        "E-1": {
            "E-1 Trad.": ["CP", "USCIS"],
            "E-1 Trad. Ren.": ["CP", "USCIS"],
            "E-1 ESE": ["CP", "USCIS"],
            "E-1 ESE Ren.": ["CP", "USCIS"],
        },
        "H-1B": ["Initial", "Extension", "Transfer", "CP"],
        "L-1": ["Initial", "Extension", "Transfer", "CP"],
        "R-1": ["Initial", "Extension", "Transfer", "CP"],
        "TN": ["Initial", "Extension", "Transfer", "CP"],
        "K-1": ["Initial", "CP"],
    },
    "Residence Permanente": {
        "Employment": {
            "Executive/Manager": ["I-140", "AOS", "I-140 & AOS", "CP"],
            "EB-2/EB-3": ["Perm", "I-140", "AOS", "I-140 & AOS", "CP"],
            "EB-5": ["I-526", "AOS", "I-527 & AOS", "CP", "I-829"],
        },
        "Family": {
            "Marriage": {
                "USC": ["I-130", "AOS", "I-130 & AOS", "CP"],
                "LPR": ["I-130", "AOS", "I-130 & AOS", "CP"],
            },
            "Family": {
                "USC": ["I-130", "AOS", "I-130 & AOS", "CP"],
                "LPR": ["I-130", "AOS", "I-130 & AOS", "CP"],
            },
        },
        "DV lottery": ["CP", "AOS"],
    }
}

# =========================
# Fonctions utilitaires de DataFrames
# (Non modifiées - Omisses pour la concision)
# =========================

def skey(*args) -> str:
    """Génère une clé unique pour st.session_state."""
    return f"{SID}_{'_'.join(map(str, args))}"

@st.cache_data(show_spinner="Lecture du fichier...")
def _read_data_file(file_content: BytesIO, file_name: str, header_row: int = 0) -> pd.DataFrame:
    """Lit les données d'un fichier téléchargé (CSV ou Excel)."""
    
    if file_name.endswith(('.xls', '.xlsx')):
        try:
            df = pd.read_excel(file_content, header=header_row, engine='openpyxl', dtype=str)
        except Exception as e:
            st.error(f"Erreur de lecture Excel : {e}")
            return pd.DataFrame()
    else: # Supposer CSV
        try:
            df = pd.read_csv(file_content, header=header_row, sep=None, engine='python', encoding='utf-8', on_bad_lines='skip', dtype=str)
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(file_content, header=header_row, sep=None, engine='python', encoding='latin1', on_bad_lines='skip', dtype=str)
            except Exception as e:
                st.error(f"Erreur de lecture CSV (Latin1) : {e}")
                return pd.DataFrame()
        except Exception as e:
            st.error(f"Erreur de lecture CSV : {e}")
            return pd.DataFrame()
    
    df = df.dropna(axis=1, how='all')
    df.columns = df.columns.str.strip().fillna('')
    df = df.dropna(axis=0, how='all')
    
    return df

def _clean_clients_data(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et standardise les types de données du DataFrame Clients."""
    
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    
    # 1. Standardiser et convertir les nombres financiers
    money_cols = ['honoraires', 'payé', 'solde', 'acompte_1', 'acompte_2', 'montant', 'autres_frais_us_']
    
    for col in money_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
            df[col] = df[col].str.replace(r'[^\d.]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float) 
    
    # 2. Rétablir le solde avec la formule (Montant Facturé - Total Payé) si les deux colonnes existent
    if 'montant' in df.columns and 'payé' in df.columns:
        df['solde'] = df['montant'] - df['payé']
    elif 'honoraires' in df.columns and 'payé' in df.columns:
        df['solde'] = df['honoraires'] - df['payé']


    # 3. Conversion des Dates
    date_cols = ['date', 'dossier_envoyé', 'dossier_approuvé', 'dossier_refusé', 'dossier_annulé']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # 4. Colonne dérivée
    if 'date' in df.columns:
         df['jours_ecoules'] = (pd.to_datetime('today') - df['date']).dt.days
         
    # 5. Assurer la présence des colonnes clés pour le CRUD, même si vides
    required_cols = ['dossier_n', 'nom', 'categorie', 'sous_categorie', 'montant', 'payé', 'solde', 'date', 'commentaires']
    for col in required_cols:
        if col not in df.columns:
            df[col] = pd.NA

    st.success("Nettoyage et conversion des données Clients terminés (Robuste).")
    return df

def _clean_visa_data(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et standardise les types de données du DataFrame Visa."""
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    return df

@st.cache_data
def _summarize_data(df: pd.DataFrame) -> Dict[str, Any]:
    """Calcule des indicateurs clés à partir du DataFrame Clients."""
    
    if df.empty:
        return {"total_clients": 0, "total_honoraires": 0.0, "total_payé": 0.0, "solde_du": 0.0, "clients_actifs": 0, "clients_payés": 0}

    total_honoraires = df['montant'].sum() if 'montant' in df.columns else 0.0
    total_payé = df['payé'].sum() if 'payé' in df.columns else 0.0
    solde_du = df['solde'].sum() if 'solde' in df.columns else 0.0
    clients_payés = (df['solde'] <= 0).sum() if 'solde' in df.columns else 0
    
    end_cols = ['dossier_approuvé', 'dossier_annulé', 'dossier_refusé']
    active_mask = pd.Series([True] * len(df), index=df.index)
    
    for col in end_cols:
        if col in df.columns:
            active_mask &= df[col].isna()

    clients_actifs = active_mask.sum()
    
    summary = {
        "total_clients": len(df),
        "total_honoraires": total_honoraires,
        "total_payé": total_payé,
        "solde_du": solde_du,
        "clients_actifs": clients_actifs,
        "clients_payés": clients_payés,
    }
    return summary


def _update_client_data(df: pd.DataFrame, new_data: Dict[str, Any], action: str) -> pd.DataFrame:
    """Ajoute, Modifie ou Supprime un client. Centralisation des actions CRUD."""
    
    dossier_n = str(new_data.get('dossier_n')).strip()
    
    if not dossier_n:
        st.error("Le Numéro de Dossier ne peut pas être vide.")
        return df

    # 1. Action DELETE
    if action == "DELETE":
        if 'dossier_n' not in df.columns: return df
            
        idx_to_delete = df[df['dossier_n'].astype(str) == dossier_n].index
        
        if not idx_to_delete.empty:
            df = df.drop(idx_to_delete).reset_index(drop=True)
            st.cache_data.clear() 
            st.success(f"Dossier N° {dossier_n} supprimé avec succès.")
            return df
        else:
            st.warning(f"Dossier N° {dossier_n} introuvable pour suppression.")
            return df

    # --- Pré-traitement pour ADD/MODIFY ---
    new_df_row = pd.DataFrame([new_data])
    new_df_row.columns = new_df_row.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    
    money_cols = ['payé', 'montant'] 
    for col in money_cols:
        if col in new_df_row.columns:
             new_df_row[col] = pd.to_numeric(new_df_row[col], errors='coerce').fillna(0.0).astype(float)
    
    montant = new_df_row['montant'].iloc[0] if 'montant' in new_df_row.columns else 0.0
    paye = new_df_row['payé'].iloc[0] if 'payé' in new_df_row.columns else 0.0
    new_df_row['solde'] = montant - paye
    
    # 2. Action MODIFY
    if action == "MODIFY":
        if 'dossier_n' not in df.columns: return df
            
        matching_rows = df[df['dossier_n'].astype(str) == dossier_n]
        if not matching_rows.empty:
            idx_to_modify = matching_rows.index[0]
            
            # Mise à jour des colonnes dans le DataFrame original
            for col in new_df_row.columns:
                if col in df.columns:
                    df.loc[idx_to_modify, col] = new_df_row[col].iloc[0]
                else:
                    # Ajout de la nouvelle colonne au DF existant si elle n'existe pas
                    df[col] = pd.NA
                    df.loc[idx_to_modify, col] = new_df_row[col].iloc[0]
            
            st.cache_data.clear() 
            st.success(f"Dossier N° {dossier_n} modifié avec succès.")
            return df
        else:
            st.warning(f"Dossier N° {dossier_n} introuvable pour modification.")
            return df

    # 3. Action ADD
    if action == "ADD":
        if 'dossier_n' in df.columns and (df['dossier_n'].astype(str) == dossier_n).any():
             st.error(f"Le Dossier N° {dossier_n} existe déjà. Utilisez l'onglet 'Modifier'.")
             return df
        
        # S'assurer que toutes les colonnes de la nouvelle ligne existent dans le DF cible
        for col in new_df_row.columns:
            if col not in df.columns:
                df[col] = pd.NA
        
        updated_df = pd.concat([df, new_df_row], ignore_index=True)
        st.cache_data.clear() 
        st.success(f"Dossier Client '{new_data.get('nom')}' (N° {dossier_n}) ajouté avec succès ! Rafraîchissement des statistiques en cours...")
        return updated_df
        
    return df


# =========================
# Fonctions de l'Interface Utilisateur (UI)
# =========================

def upload_section():
    """Section de chargement des fichiers (Barre latérale)."""
    st.sidebar.header("📁 Chargement des Fichiers")
    
    # ------------------- Fichier Clients -------------------
    content_clients_loaded = st.session_state.get(skey("raw_clients_content"))
    
    uploaded_file_clients = st.sidebar.file_uploader(
        "Clients/Dossiers (.csv, .xlsx)",
        type=['csv', 'xlsx'],
        key=skey("upload", "clients"),
    )
    
    if uploaded_file_clients is not None:
        st.session_state[skey("raw_clients_content")] = uploaded_file_clients.read()
        st.session_state[skey("clients_name")] = uploaded_file_clients.name
        # On force le recalcul du DF en cas de nouvel upload
        st.session_state[skey("df_clients")] = pd.DataFrame() 
        st.sidebar.success(f"Clients : **{uploaded_file_clients.name}** chargé.")
    elif content_clients_loaded:
        st.sidebar.success(f"Clients : **{st.session_state.get(skey('clients_name'), 'Précédent')}** (Persistant)")


    # ------------------- Fichier Visa -------------------
    content_visa_loaded = st.session_state.get(skey("raw_visa_content"))
    
    uploaded_file_visa = st.sidebar.file_uploader(
        "Table de Référence Visa (.csv, .xlsx)",
        type=['csv', 'xlsx'],
        key=skey("upload", "visa"),
    )

    if uploaded_file_visa is not None:
        st.session_state[skey("raw_visa_content")] = uploaded_file_visa.read()
        st.session_state[skey("visa_name")] = uploaded_file_visa.name
        # On force le recalcul du DF en cas de nouvel upload
        st.session_state[skey("df_visa")] = pd.DataFrame() 
        st.sidebar.success(f"Visa : **{uploaded_file_visa.name}** chargé.")
    elif content_visa_loaded:
        st.sidebar.success(f"Visa : **{st.session_state.get(skey('visa_name'), 'Précédent')}** (Persistant)")


def data_processing_flow():
    """Gère le chargement, le nettoyage et le stockage des DataFrames. 
    Les DataFrames sont recalculés seulement si le contenu brut change ou s'ils sont vides.
    """
    
    st.session_state.setdefault(skey("df_clients"), pd.DataFrame())
    st.session_state.setdefault(skey("df_visa"), pd.DataFrame())

    # --- 1. Clients ---
    content_clients = st.session_state.get(skey("raw_clients_content"))
    file_name_clients = st.session_state.get(skey("clients_name"), "")
    header_clients = st.session_state.get(skey("header_clients_row"), 0)

    # Recharger seulement si le DF est vide (premier run) OU si un nouvel upload a vidé le DF
    if content_clients and file_name_clients and st.session_state.get(skey("df_clients")).empty:
        df_raw_clients = _read_data_file(BytesIO(content_clients), file_name_clients, header_row=header_clients)
        if not df_raw_clients.empty:
            df_cleaned_clients = _clean_clients_data(df_raw_clients)
            st.session_state[skey("df_clients")] = df_cleaned_clients
    
    # --- 2. Visa ---
    content_visa = st.session_state.get(skey("raw_visa_content"))
    file_name_visa = st.session_state.get(skey("visa_name"), "")
    header_visa = st.session_state.get(skey("header_visa_row"), 0)

    if content_visa and file_name_visa and st.session_state.get(skey("df_visa")).empty:
        df_raw_visa = _read_data_file(BytesIO(content_visa), file_name_visa, header_row=header_visa)
        if not df_raw_visa.empty:
            df_cleaned_visa = _clean_visa_data(df_raw_visa)
            st.session_state[skey("df_visa")] = df_cleaned_visa


# --- Onglet Accueil ---
def home_tab(df_clients: pd.DataFrame):
    """Contenu de l'onglet Accueil/Statistiques."""
    st.header("📊 Statistiques Clés")
    
    if df_clients.empty:
        st.info("Veuillez charger ou ajouter des dossiers clients pour afficher les statistiques.")
        return
        
    summary = _summarize_data(df_clients)

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Clients Totaux", f"{summary['total_clients']:,}".replace(",", " "))
    col2.metric("Total Reçu (Payé)", f"${summary['total_payé']:,.2f}".replace(",", " "))
    col3.metric("Solde Total Dû", f"${summary['solde_du']:,.2f}".replace(",", " "))
    col4.metric("Dossiers Actifs", f"{summary['clients_actifs']:,}".replace(",", " "))
    
    st.divider()
    
    st.subheader("Analyse de la Catégorie Visa")
    if 'categorie' in df_clients.columns:
        counts = df_clients['categorie'].value_counts().head(10)
        st.bar_chart(counts, use_container_width=True)
    else:
        st.warning("Colonne 'categorie' introuvable pour l'analyse. Vérifiez l'index d'en-tête.")

# --- NOUVEL ONGLET COMPTABILITÉ ---
def accounting_tab(df_clients: pd.DataFrame):
    """Contenu de l'onglet Comptabilité (Suivi financier)."""
    st.header("📈 Suivi Financier (Comptabilité Client)")
    
    if df_clients.empty:
        st.info("Veuillez charger ou ajouter des dossiers clients pour afficher les données comptables.")
        return
        
    summary = _summarize_data(df_clients)

    # 1. KPIs
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Facturé (Montant)", f"${summary['total_honoraires']:,.2f}".replace(",", " "))
    col2.metric("Total Reçu (Payé)", f"${summary['total_payé']:,.2f}".replace(",", " "))
    col3.metric("Solde Total Dû", f"${summary['solde_du']:,.2f}".replace(",", " "))
    col4.metric("Dossiers Payés (Solde <= 0)", f"{summary['clients_payés']:,}".replace(",", " "))
    
    st.divider()

    # --- Filtre Client ---
    st.subheader("Détail du Compte Client")
    
    df_clients_for_select = df_clients[['dossier_n', 'nom']].dropna(subset=['dossier_n'])
    client_options = {
        f"{r['dossier_n']} - {r['nom']}": r['dossier_n'] 
        for _, r in df_clients_for_select.iterrows()
    }
    
    selected_key_acct = st.selectbox(
        "Filtrer par Client (Optionnel)",
        ["Tous les clients"] + list(client_options.keys()),
        key=skey("acct", "select_client")
    )

    df_filtered = df_clients.copy()
    
    if selected_key_acct != "Tous les clients":
        selected_dossier_n = client_options.get(selected_key_acct)
        df_filtered = df_clients[df_clients['dossier_n'].astype(str) == selected_dossier_n].copy()

    # 2. Tableau de ventilation
    
    accounting_cols = ['dossier_n', 'nom', 'categorie', 'montant', 'payé', 'solde', 'date']
    valid_cols = [col for col in accounting_cols if col in df_filtered.columns]
    
    df_accounting = df_filtered[valid_cols].copy()
    
    # Formatage des colonnes monétaires pour l'affichage
    for col in ['montant', 'payé', 'solde']:
        if col in df_accounting.columns:
            df_accounting[col] = df_accounting[col].apply(lambda x: f"${x:,.2f}".replace(",", " "))
            
    df_accounting.rename(columns={
        'dossier_n': 'N° Dossier',
        'nom': 'Nom Client',
        'categorie': 'Catégorie Visa',
        'montant': 'Montant Facturé',
        'payé': 'Total Payé',
        'solde': 'Solde Dû',
        'date': 'Date Ouverture',
    }, inplace=True)
    
    st.dataframe(
        df_accounting.sort_values(by='Solde Dû', key=lambda x: x.str.replace(r'[^\d.]', '', regex=True).astype(float), ascending=False), 
        use_container_width=True,
    )
    st.caption("Le solde dû est calculé par `Montant Facturé - Total Payé`.")

# --- FONCTION POUR LA CLASSIFICATION EN CASCADE (MISE À JOUR) ---
def _render_visa_classification_form(key_suffix: str, initial_category: Optional[str] = None, initial_type: Optional[str] = None, initial_subtype: Optional[str] = None) -> Tuple[str, str]:
    """
    Affiche les selectbox en cascade pour la classification des visas.
    Utilise des widgets différents selon les niveaux.
    Retourne la catégorie (niveau 1) et la sous-catégorie (niveau le plus profond)
    
    initial_subtype peut être soit la clé de Niveau 3 (Ex: 'EB-2/EB-3') soit l'option de Niveau 4 (Ex: 'CP').
    """
    
    main_keys = list(VISA_STRUCTURE.keys())
    default_cat_index = main_keys.index(initial_category) + 1 if initial_category in main_keys else 0
    
    col_cat, col_type = st.columns(2)
    
    visa_category = initial_category if initial_category in main_keys else "Sélectionnez un groupe"
    final_visa_type = ""
    selected_type = ""
    
    # 1. Sélection de la Catégorie (Niveau 1)
    with col_cat:
        visa_category = st.selectbox(
            "1. Catégorie de Visa (Grand Groupe)",
            ["Sélectionnez un groupe"] + main_keys,
            index=default_cat_index,
            key=skey("form", key_suffix, "cat_main"),
        )
        
    if visa_category != "Sélectionnez un groupe":
        
        selected_options = VISA_STRUCTURE.get(visa_category, {})
        visa_types_list = list(selected_options.keys())
        
        # Logique de pré-remplissage pour Niveau 2 (Type de Visa)
        default_type_index = visa_types_list.index(initial_type) + 1 if initial_type in visa_types_list else 0
        
        # Recherche si initial_type est un niveau 3 (ex: 'E-2 Inv.') et non un niveau 2 (ex: 'E-2')
        if default_type_index == 0 and initial_type:
            for key, value in selected_options.items():
                if isinstance(value, dict):
                    if initial_type in value: # Cas où initial_type est le niveau 3 (ex: 'E-2 Inv.')
                         default_type_index = visa_types_list.index(key) + 1 # key est le niveau 2 (ex: 'E-2')
                         break
                    elif initial_type == key: # Cas où initial_type est un niveau 2 lui-même (non structuré en dictionnaire)
                         default_type_index = visa_types_list.index(key) + 1
                         break
                elif initial_type == key:
                    default_type_index = visa_types_list.index(key) + 1
                    break

        # 2. Sélection du Type de Visa (Niveau 2) - Dropdown classique
        with col_type:
            selected_type = st.selectbox(
                f"2. Type de Visa ({visa_category})",
                ["Sélectionnez un type"] + visa_types_list,
                index=default_type_index,
                key=skey("form", key_suffix, "cat_type"),
            )

        if selected_type and selected_type != "Sélectionnez un type":
            current_options = selected_options.get(selected_type)

            if isinstance(current_options, list):
                # Cas 1 : Niveau 3 (Liste simple)
                st.subheader(f"3. Option pour **{selected_type}**")
                
                # --- Utilisation de st.radio (Boutons Bascule) ---
                options_list = current_options
                # initial_subtype est l'option finale (Ex: CP ou EOS)
                default_sub_index = options_list.index(initial_subtype) if initial_subtype in options_list else 0
                
                final_selection = st.radio(
                    "Choisissez l'option finale",
                    options_list,
                    index=default_sub_index,
                    key=skey("form", key_suffix, "sub1"),
                    horizontal=True
                )
                final_visa_type = f"{selected_type} ({final_selection})"
                
            elif isinstance(current_options, dict):
                # Cas 2 : Niveau 3 (Dictionnaire/Sous-catégories)
                st.subheader(f"3. Sous-catégorie pour **{selected_type}**")
                
                nested_keys = list(current_options.keys())
                nested_key_to_select = ""

                # 1. Vérifie si initial_subtype est l'une des clés de Niveau 3 (Ex: EB-2/EB-3 ou E-2 Inv.)
                if initial_subtype in nested_keys:
                    nested_key_to_select = initial_subtype
                
                # 2. Sinon, si initial_subtype est l'option de Niveau 4 (Ex: CP), on cherche la clé Niveau 3 parente
                elif initial_subtype:
                    for k, v in current_options.items():
                         if isinstance(v, list) and initial_subtype in v:
                             nested_key_to_select = k # k est la clé Niveau 3 (Ex: 'E-2 Inv.')
                             break
                
                default_nested_index = nested_keys.index(nested_key_to_select) + 1 if nested_key_to_select in nested_keys else 0

                # --- Utilisation de st.selectbox pour le niveau 3 (Sous-catégorie) ---
                nested_key = st.selectbox(
                    f"Sous-catégorie de {selected_type}",
                    ["Sélectionnez la sous-catégorie"] + nested_keys,
                    index=default_nested_index,
                    key=skey("form", key_suffix, "nested_key"),
                )
                
                if nested_key and nested_key != "Sélectionnez la sous-catégorie":
                    # Niveau 4 : Options finales
                    nested_options = current_options.get(nested_key)
                    if nested_options and isinstance(nested_options, list):
                        st.subheader(f"4. Option finale pour **{nested_key}**")
                        
                        # --- Utilisation de st.radio (Boutons Bascule) ---
                        options_list_nested = nested_options
                        # initial_subtype est l'option finale (CP, USCIS, etc.)
                        default_sub_index = options_list_nested.index(initial_subtype) if initial_subtype in options_list_nested else 0
                        
                        final_selection = st.radio(
                            "Choisissez l'option finale",
                            options_list_nested,
                            index=default_sub_index,
                            key=skey("form", key_suffix, "sub2"),
                            horizontal=True
                        )
                        final_visa_type = f"{nested_key} ({final_selection})"
                    
                    else:
                        final_visa_type = nested_key
                else:
                    final_visa_type = selected_type

    # Retourne la Catégorie (Niveau 1) et la Sous-Catégorie (Niveau final détaillé)
    return visa_category, final_visa_type

# --- GESTION DES DOSSIERS (AJOUT/MODIF/SUPPRESSION) ---
def dossier_management_tab(df_clients: pd.DataFrame):
    """Contenu de l'onglet Saisie/Modification/Suppression de Dossiers."""
    st.header("📝 Gestion des Dossiers Clients (CRUD)")
    
    tab_add, tab_modify, tab_delete = st.tabs(["➕ Ajouter un Dossier", "✍️ Modifier un Dossier", "🗑️ Supprimer un Dossier"])

    # =========================================================================
    # LOGIQUE D'AJOUT (ADD)
    # (Non modifiée - Omisses pour la concision)
    # =========================================================================
    with tab_add:
        st.subheader("Ajouter un Nouveau Dossier")
        
        # Déterminer le prochain ID/Numéro
        next_dossier_n = 13000
        if not df_clients.empty and 'dossier_n' in df_clients.columns:
            try:
                max_n = df_clients['dossier_n'].astype(str).str.extract(r'(\d+)').astype(float).max()
                next_dossier_n = int(max_n + 1) if not pd.isna(max_n) else 13000
            except:
                 next_dossier_n = 13000

        
        with st.form("add_client_form"):
            st.markdown("---")
            
            col_id, col_name, col_date = st.columns(3)
            client_name = col_name.text_input("Nom du Client", key=skey("form_add", "nom"))
            dossier_n = col_id.text_input("Numéro de Dossier", value=str(next_dossier_n), key=skey("form_add", "dossier_n"))
            date_dossier = col_date.date_input("Date d'Ouverture du Dossier", value=pd.to_datetime('today'), key=skey("form_add", "date"))
            
            st.markdown("---")
            
            col_montant, col_paye = st.columns(2)
            montant_facture = col_montant.number_input("Total Facturé (Montant)", min_value=0.0, step=100.0, key=skey("form_add", "montant"))
            paye_initial = col_paye.number_input("Paiement Initial Reçu (Payé)", min_value=0.0, step=100.0, key=skey("form_add", "payé"))
            
            # Correction de sécurité contre les valeurs None lors de l'initialisation Streamlit
            if montant_facture is None or paye_initial is None:
                solde_calcule = 0.0
            else:
                solde_calcule = montant_facture - paye_initial
            
            st.metric("Solde Initial Dû (Calculé)", f"${solde_calcule:,.2f}".replace(",", " "))
            
            st.markdown("---")
            st.subheader("Classification de Visa Hiérarchique")
            
            # --- APPEL DE LA CLASSIFICATION EN CASCADE ---
            visa_category, visa_type = _render_visa_classification_form(key_suffix="add")
            
            st.markdown("---")
            
            commentaires = st.text_area("Notes / Commentaires sur le Dossier", key=skey("form_add", "commentaires"))
            
            submitted = st.form_submit_button("✅ Ajouter le Nouveau Dossier")
            
            if submitted:
                if not client_name or montant_facture < 0 or dossier_n.strip() == "":
                    st.error("Veuillez renseigner le Nom du Client, le Numéro de Dossier, et le Montant Facturé.")
                else:
                    new_entry = {
                        "dossier_n": dossier_n,
                        "nom": client_name,
                        "date": date_dossier.strftime('%Y-%m-%d'),
                        "categorie": visa_category if visa_category != "Sélectionnez un groupe" else "",
                        "sous_categorie": visa_type,
                        "montant": montant_facture, 
                        "payé": paye_initial,
                        "commentaires": commentaires,
                    }
                    
                    updated_df_clients = _update_client_data(df_clients, new_entry, "ADD")
                    st.session_state[skey("df_clients")] = updated_df_clients
                    st.rerun()

    # =========================================================================
    # LOGIQUE DE MODIFICATION (MODIFY)
    # =========================================================================
    with tab_modify:
        st.subheader("Modifier un Dossier Existant")
        
        if df_clients.empty or 'dossier_n' not in df_clients.columns:
            st.info("Aucun dossier client chargé ou créé.")
            return

        client_list = df_clients['dossier_n'].dropna().astype(str).unique()
        if 'nom' in df_clients.columns:
            client_options = {f"{r['dossier_n']} - {r['nom']}": r['dossier_n'] for _, r in df_clients[['dossier_n', 'nom']].iterrows() if pd.notna(r['dossier_n'])}
        else:
             client_options = {f"{n}": n for n in client_list}

        selected_key = st.selectbox(
            "Sélectionner le Dossier à Modifier",
            [""] + list(client_options.keys()),
            key=skey("modify", "select_client")
        )

        selected_dossier_n = client_options.get(selected_key)
        
        if selected_dossier_n:
            current_data = df_clients[df_clients['dossier_n'].astype(str) == selected_dossier_n].iloc[0].to_dict()
            
            st.markdown(f"---")
            st.info(f"Modification du Dossier N°: **{selected_dossier_n}**")

            with st.form("modify_client_form"):
                
                # --- Remplissage des champs ---
                col_name, col_date = st.columns(2)
                
                client_name_mod = col_name.text_input("Nom du Client", value=current_data.get('nom', ''), key=skey("form_mod", "nom"))
                
                # S'assurer que la date est un objet date pour st.date_input
                date_val = current_data.get('date')
                if pd.isna(date_val):
                    date_val = pd.to_datetime('today').date()
                elif isinstance(date_val, pd.Timestamp):
                    date_val = date_val.date()
                
                date_dossier_mod = col_date.date_input(
                    "Date d'Ouverture du Dossier", 
                    value=date_val, 
                    key=skey("form_mod", "date")
                )
                
                st.markdown("---")
                
                col_montant, col_paye = st.columns(2)
                montant_facture_mod = col_montant.number_input(
                    "Total Facturé (Montant)", 
                    min_value=0.0, 
                    step=100.0, 
                    value=current_data.get('montant', 0.0), 
                    key=skey("form_mod", "montant")
                )
                paye_mod = col_paye.number_input(
                    "Total Paiements Reçus (Payé)", 
                    min_value=0.0, 
                    step=100.0, 
                    value=current_data.get('payé', 0.0), 
                    key=skey("form_mod", "payé")
                )
                
                # --- Correction du TypeError ---
                if montant_facture_mod is None or paye_mod is None:
                    solde_mod = 0.0
                else:
                    solde_mod = montant_facture_mod - paye_mod 
                    
                st.metric("Solde Actuel Dû (Calculé)", f"${solde_mod:,.2f}".replace(",", " "))
                
                st.markdown("---")
                st.subheader("Classification de Visa Hiérarchique")
                
                # Préparation des valeurs initiales pour la cascade
                current_cat = str(current_data.get('categorie', ''))
                full_sub_cat = str(current_data.get('sous_categorie', ''))
                
                # Logique pour décortiquer la sous-catégorie complexe
                current_sub_type = full_sub_cat 
                current_final_subtype = None 

                # Extraction du sous-type final s'il est entre parenthèses
                match_paren = re.search(r'\((.+)\)', full_sub_cat)
                if match_paren:
                    current_final_subtype = match_paren.group(1).strip() # Level 4 Option (Ex: CP)
                    current_sub_type = full_sub_cat[:match_paren.start()].strip() # Level 3 Key (Ex: E-2 Inv.)
                
                # Le niveau 2 est le parent direct dans la structure VISA_STRUCTURE
                level2_type = current_sub_type # Par défaut, le Niveau 3 key ou le Niveau 2 key (si non imbriqué)
                if current_cat in VISA_STRUCTURE:
                    level2_options = VISA_STRUCTURE[current_cat]
                    for key_level2, val_level2 in level2_options.items():
                        if isinstance(val_level2, dict) and current_sub_type in val_level2:
                            level2_type = key_level2 # Ex: 'E-2'
                            break
                        elif key_level2 == current_sub_type:
                            level2_type = key_level2
                            break
                
                # NOUVEAU: Détermination de l'argument à passer pour initial_subtype (3ème argument)
                initial_level_3_or_4 = current_final_subtype # Par défaut: Level 4 option (Ex: CP)
                if initial_level_3_or_4 is None:
                    # Si Level 4 est absent, le Level 3 key est le type final (Ex: E-2 Inv. ou EB-2/EB-3)
                    initial_level_3_or_4 = current_sub_type 

                # --- APPEL DE LA CLASSIFICATION EN CASCADE AVEC VALEURS INITIALES ---
                visa_category_mod, visa_type_mod = _render_visa_classification_form(
                    key_suffix="mod",
                    initial_category=current_cat, # Niveau 1
                    initial_type=level2_type, # Niveau 2
                    initial_subtype=initial_level_3_or_4, # Niveau 3 OU Niveau 4
                )
                
                commentaires_mod = st.text_area(
                    "Notes / Commentaires sur le Dossier", 
                    value=current_data.get('commentaires', ''),
                    key=skey("form_mod", "commentaires")
                )
                
                # Bouton de soumission
                submitted_mod = st.form_submit_button("💾 Enregistrer les Modifications")
                
                if submitted_mod:
                    updated_entry = {
                        "dossier_n": selected_dossier_n,
                        "nom": client_name_mod,
                        "date": date_dossier_mod.strftime('%Y-%m-%d'),
                        "categorie": visa_category_mod if visa_category_mod != "Sélectionnez un groupe" else "",
                        "sous_categorie": visa_type_mod,
                        "montant": montant_facture_mod, 
                        "payé": paye_mod,
                        "commentaires": commentaires_mod,
                    }
                    
                    updated_df_clients = _update_client_data(df_clients, updated_entry, "MODIFY")
                    st.session_state[skey("df_clients")] = updated_df_clients
                    st.rerun() 
    
    # =========================================================================
    # LOGIQUE DE SUPPRESSION (DELETE)
    # (Non modifiée - Omisses pour la concision)
    # =========================================================================
    with tab_delete:
        st.subheader("Supprimer un Dossier Définitivement")
        st.warning("ATTENTION : Cette action est irréversible.")
        
        if df_clients.empty or 'dossier_n' not in df_clients.columns:
            st.info("Aucun dossier client chargé ou créé.")
            return

        client_list = df_clients['dossier_n'].dropna().astype(str).unique()
        if 'nom' in df_clients.columns:
            client_options = {f"{r['dossier_n']} - {r['nom']}": r['dossier_n'] for _, r in df_clients[['dossier_n', 'nom']].iterrows() if pd.notna(r['dossier_n'])}
        else:
             client_options = {f"{n}": n for n in client_list}
             
        with st.form("delete_client_form"):
            selected_key_del = st.selectbox(
                "Sélectionner le Dossier à Supprimer",
                [""] + list(client_options.keys()),
                key=skey("delete", "select_client")
            )

            selected_dossier_n_del = client_options.get(selected_key_del)
            
            st.markdown("---")
            
            delete_confirmed = False
            if selected_dossier_n_del:
                delete_confirmed = st.checkbox(f"Je confirme la suppression définitive du dossier N° **{selected_dossier_n_del}**", key=skey("delete", "confirm"))
            
            submitted_del = st.form_submit_button("💣 SUPPRIMER le Dossier", disabled=not selected_dossier_n_del or not delete_confirmed)
            
            if submitted_del and delete_confirmed:
                delete_entry = {"dossier_n": selected_dossier_n_del}
                
                updated_df_clients = _update_client_data(df_clients, delete_entry, "DELETE")
                st.session_state[skey("df_clients")] = updated_df_clients
                st.rerun()


def settings_tab():
    """Contenu de l'onglet Configuration."""
    st.header("⚙️ Configuration du Chargement")
    
    st.markdown("""
        Veuillez spécifier l'index de la ligne contenant les noms de colonnes réels.
        * **0** (par défaut) : première ligne.
        * **1** : deuxième ligne, etc.
    """)
    
    # --- Configuration Clients ---
    st.subheader("Fichier Clients")
    current_header_clients = st.session_state.get(skey("header_clients_row"), 0)
    new_header_clients = st.number_input(
        "Index de la ligne d'en-tête (Clients)",
        min_value=0,
        value=current_header_clients,
        step=1,
        key=skey("input", "header_clients"),
    )
    if new_header_clients != current_header_clients:
         st.session_state[skey("header_clients_row")] = new_header_clients
         # Forcer le rechargement du DF Clients si la configuration change
         st.session_state[skey("df_clients")] = pd.DataFrame() 
         st.rerun() 

    # --- Configuration Visa ---
    st.subheader("Fichier Visa")
    current_header_visa = st.session_state.get(skey("header_visa_row"), 0)
    new_header_visa = st.number_input(
        "Index de la ligne d'en-tête (Visa)",
        min_value=0,
        value=current_header_visa,
        step=1,
        key=skey("input", "header_visa"),
    )
    if new_header_visa != current_header_visa:
         st.session_state[skey("header_visa_row")] = new_header_visa
         # Forcer le rechargement du DF Visa si la configuration change
         st.session_state[skey("df_visa")] = pd.DataFrame() 
         st.rerun() 


def export_tab(df_clients: pd.DataFrame, df_visa: pd.DataFrame):
    """Contenu de l'onglet Export."""
    st.header("💾 Export des Données Nettoyées")
    
    colx, coly = st.columns(2)

    with colx:
        if df_clients.empty:
            st.info("Pas de données Clients nettoyées à exporter.")
        else:
            buf = BytesIO()
            # On exporte l'index=False car l'index Pandas n'est pas pertinent
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                df_clients.to_excel(w, index=False, sheet_name="Clients_Nettoyes")
            st.download_button(
                "⬇️ Exporter Clients_Nettoyes.xlsx",
                data=buf.getvalue(),
                file_name="Clients_export_nettoye.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with coly:
        if df_visa.empty:
            st.info("Pas de données Visa nettoyées à exporter.")
        else:
            bufv = BytesIO()
            with pd.ExcelWriter(bufv, engine="openpyxl") as w:
                df_visa.to_excel(w, index=False, sheet_name="Visa_Nettoyes")
            st.download_button(
                "⬇️ Exporter Visa_Nettoyes.xlsx",
                data=bufv.getvalue(),
                file_name="Visa_export_nettoye.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# =========================
# Application principale
# =========================

def main():
    """Fonction principale de l'application Streamlit."""
    st.set_page_config(
        page_title=APP_TITLE,
        layout="wide",
        initial_sidebar_state="expanded"
    )
    st.title(APP_TITLE)
    
    # 1. Section de chargement des fichiers
    upload_section()
    
    # 2. Flux de traitement des données
    data_processing_flow()
    
    # Récupérer les DataFrames nettoyés
    df_clients = st.session_state.get(skey("df_clients"), pd.DataFrame())
    df_visa = st.session_state.get(skey("df_visa"), pd.DataFrame())

    # 3. Affichage des onglets
    tab_home, tab_accounting, tab_management, tab_config, tab_clients_view, tab_visa_view, tab_export = st.tabs([
        "🏠 Accueil & Stats", 
        "📈 Comptabilité",
        "📝 Gestion Dossiers", 
        "⚙️ Configuration",
        "📄 Clients - Aperçu", 
        "📄 Visa - Aperçu", 
        "💾 Export",
    ])

    with tab_home:
        home_tab(df_clients)
        
    with tab_accounting:
        accounting_tab(df_clients) 

    with tab_management:
        dossier_management_tab(df_clients) 

    with tab_config:
        settings_tab()

    with tab_clients_view:
        st.header("📄 Clients — Aperçu des Données Nettoyées")
        if df_clients.empty:
            st.info("Aucun fichier Clients chargé ou données non valides.")
        else:
            st.dataframe(df_clients, use_container_width=True)

    with tab_visa_view:
        st.header("📄 Visa — Aperçu des Données Nettoyées")
        if df_visa.empty:
            st.info("Aucun fichier Visa chargé ou données non valides.")
        else:
            st.dataframe(df_visa, use_container_width=True)

    with tab_export:
        export_tab(df_clients, df_visa)


if __name__ == "__main__":
    main()
