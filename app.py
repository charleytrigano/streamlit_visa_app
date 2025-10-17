import io
import re
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import numpy as np # Ajouté pour les opérations numériques sécurisées

# =========================
# Constantes et Configuration
# =========================================================================
APP_TITLE = "🛂 Visa Manager - Gestion Complète"
SID = "vmgr_v6"

# Le dictionnaire codé en dur est vide, il sera rempli par la fonction _build_visa_structure
VISA_STRUCTURE = {} 


# =========================
# Fonctions utilitaires de DataFrames
# =========================================================================

def skey(*args) -> str:
    """Génère une clé unique pour st.session_state."""
    return f"{SID}_{'_'.join(map(str, args))}"

@st.cache_data(show_spinner="Lecture du fichier...")
def _read_data_file(file_content: BytesIO, file_name: str, header_row: int = 0) -> pd.DataFrame:
    """
    Lit les données d'un fichier téléchargé (CSV ou Excel).
    
    CORRECTION MAJEURE: Détection du format plus tolérante, y compris 
    les fichiers .xlsx renommés en .csv.
    """
    
    # 1. Détection du format (plus tolérante)
    # On vérifie l'extension standard OU la présence de 'xlsx'/'xls' dans le nom du fichier
    is_excel = file_name.endswith(('.xls', '.xlsx')) or 'xlsx' in file_name.lower() or 'xls' in file_name.lower()
    
    # Assurez-vous que le pointeur est au début du fichier
    file_content.seek(0) 

    if is_excel:
        try:
            # Tenter la lecture Excel pour les formats xls/xlsx
            df = pd.read_excel(file_content, header=header_row, engine='openpyxl', dtype=str)
        except Exception as e:
            st.error(f"Erreur de lecture Excel : {e}")
            return pd.DataFrame()
    else: 
        # Tenter la lecture CSV
        try:
            # Tente de détecter automatiquement le séparateur (sep=None)
            df = pd.read_csv(file_content, header=header_row, sep=None, engine='python', encoding='utf-8', on_bad_lines='skip', dtype=str)
        except UnicodeDecodeError:
            try:
                # Tente l'encodage Latin1
                file_content.seek(0) # Reset du pointeur après échec UTF-8
                df = pd.read_csv(file_content, header=header_row, sep=None, engine='python', encoding='latin1', on_bad_lines='skip', dtype=str)
            except Exception as e:
                st.error(f"Erreur de lecture CSV (Latin1) : {e}")
                return pd.DataFrame()
        except Exception as e:
            st.error(f"Erreur de lecture CSV : {e}")
            return pd.DataFrame()
    
    # Nettoyage standard
    df = df.dropna(axis=1, how='all')
    df.columns = df.columns.str.strip().fillna('')
    df = df.dropna(axis=0, how='all')
    
    return df

def _clean_clients_data(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et standardise les types de données du DataFrame Clients."""
    
    # Nettoyage des noms de colonnes : minuscule et remplacement des non-alphanumériques par '_'
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    
    # 1. Standardiser et convertir les nombres financiers 
    money_cols = ['honoraires', 'payé', 'solde', 'acompte_1', 'acompte_2', 'montant', 'autres_frais_us_']
    for col in money_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
            df[col] = df[col].str.replace(r'[^\d.]', '', regex=True)
            # Conversion en float sécurisée
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float) 
    
    # 2. Rétablir le solde avec la formule (basé sur la présence des colonnes)
    if 'montant' in df.columns and 'payé' in df.columns:
        df['solde'] = df['montant'] - df['payé']
    elif 'honoraires' in df.columns and 'payé' in df.columns:
        df['solde'] = df['honoraires'] - df['payé']

    # 3. Conversion des Dates
    date_cols = ['date', 'dossier_envoyé', 'dossier_approuvé', 'dossier_refusé', 'dossier_annulé']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # 4. Assurer la présence des colonnes clés pour le CRUD
    required_cols = ['dossier_n', 'nom', 'categorie', 'sous_categorie', 'montant', 'payé', 'solde', 'date', 'commentaires']
    for col in required_cols:
        if col not in df.columns:
            df[col] = pd.NA

    # Convertir dossier_n en string pour les opérations de matching
    if 'dossier_n' in df.columns:
        df['dossier_n'] = df['dossier_n'].astype(str).str.strip()
        
    st.success("Nettoyage et conversion des données Clients terminés (Robuste).")
    return df

def _clean_visa_data(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et standardise les types de données du DataFrame Visa."""
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    return df

# --- FONCTION CLÉ : CONSTRUIRE LA STRUCTURE DYNAMIQUE ---
@st.cache_data(show_spinner="Construction de la structure Visa...")
def _build_visa_structure(df_visa: pd.DataFrame) -> Dict[str, Any]:
    """
    Construit la structure de classification VISA à partir du DataFrame Visa.xlsx.
    Suppose que les colonnes sont dans l'ordre hiérarchique : 
    [N1_Categorie, N2_Type, N3_SousCategorie, N4_Option, ...]
    et qu'une colonne contient '1' pour indiquer une classification valide.
    """
    if df_visa.empty:
        return {}

    cols = df_visa.columns.tolist()
    
    # Renommer les 4 premières colonnes pour la hiérarchie (si elles existent)
    df_temp = df_visa.copy()
    col_map = {}
    if len(cols) > 0: col_map[cols[0]] = 'N1_Categorie'
    if len(cols) > 1: col_map[cols[1]] = 'N2_Type'
    if len(cols) > 2: col_map[cols[2]] = 'N3_SousCategorie'
    # CORRECTION SYNTAXE: Utilisation de len(cols) au lieu de len(cols]
    if len(cols) > 3: col_map[cols[3]] = 'N4_Option'

    # S'assurer que les colonnes sont présentes et renommées
    df_temp.rename(columns=col_map, inplace=True)
    
    if len(cols) < 4:
         # st.warning("Le fichier Visa ne contient pas assez de colonnes pour une classification à 4 niveaux.")
         # On continue même s'il y a moins de 4, la logique du dictionnaire gérera
         pass
         
    # Trouver la colonne d'indicateur ('1') - à partir de la 5e colonne
    indicator_col = next((col for col in df_temp.columns[4:] if df_temp[col].astype(str).str.contains('1', na=False).any()), None)
    
    if not indicator_col:
        # Fallback: S'il n'y a pas d'indicateur '1', on suppose que toutes les lignes sont valides
        # st.warning("Impossible de trouver la colonne indicatrice de type ('1'). Utilisant toutes les lignes.")
        df_valid = df_temp.copy()
    else:
        # Filtrer uniquement les lignes valides (où l'indicateur est '1')
        df_valid = df_temp[df_temp[indicator_col].astype(str).str.strip() == '1'].copy()
    
    # Convertir en dictionnaire hiérarchique
    structure = {}
    
    for _, row in df_valid.iterrows():
        # Utiliser .get pour les colonnes renommées, en cas de structure incomplète
        n1_cat = row.get('N1_Categorie', '').strip()
        n2_type = row.get('N2_Type', '').strip()
        n3_subcat = row.get('N3_SousCategorie', '').strip()
        n4_option = row.get('N4_Option', '').strip()
        
        if not n1_cat or not n2_type: continue

        if n1_cat not in structure:
            structure[n1_cat] = {}
            
        if not n3_subcat: # Cas N3 manquant (Structure 3-Niveaux ou moins)
             if n2_type not in structure[n1_cat]:
                 # Si N4 est la seule sous-option, on crée une liste
                 structure[n1_cat][n2_type] = []
                 
             if n4_option and n4_option not in structure[n1_cat][n2_type]: 
                  structure[n1_cat][n2_type].append(n4_option)

        else: # Cas N3 présent (Structure 4-Niveaux ou plus)
            if n2_type not in structure[n1_cat]:
                # On initialise N2 comme un dictionnaire pour le niveau N3
                structure[n1_cat][n2_type] = {} 
            
            if n3_subcat not in structure[n1_cat][n2_type]:
                structure[n1_cat][n2_type][n3_subcat] = []
                
            if n4_option and n4_option not in structure[n1_cat][n2_type][n3_subcat]:
                structure[n1_cat][n2_type][n3_subcat].append(n4_option)

    st.success("Structure de classification Visa construite dynamiquement.")
    return structure

# --- FONCTION DE RÉSOLUTION DES NIVEAUX HIERARCHIQUES ---
def _resolve_visa_levels(category: str, full_sub_cat: str, visa_structure: Dict) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Résout les niveaux de classification à partir des données stockées (categorie, sous_categorie) 
    en utilisant la structure VISA dynamique.
    
    Retourne (Niveau 2 Type, Niveau 3 Key, Niveau 4 Option).
    """
    level2_type = None 
    level3_key = full_sub_cat.strip()
    level4_option = None 

    if not category or category not in visa_structure:
        return None, None, None

    # 1. Extraction de l'Option Niveau 4 (entre parenthèses) et du Niveau 3 Key
    match_paren = re.search(r'\((.+)\)', level3_key)
    if match_paren:
        level4_option = match_paren.group(1).strip()
        level3_key = level3_key[:match_paren.start()].strip() # Ceci est la partie avant la parenthèse

    # 2. Détermination du Niveau 2 parent (Type)
    level2_options = visa_structure.get(category, {})
    
    # Chercher le Niveau 2 parent
    for key_level2, val_level2 in level2_options.items():
        # Cas 1: N3_Key (avant la parenthèse) correspond directement au N2 Key (fusion N2/N3)
        if key_level2 == level3_key: 
            level2_type = key_level2
            return level2_type, None, level4_option # N3 est null car N2 et N3 sont fusionnés
        
        # Cas 2: N3_Key correspond à une clé dans le dictionnaire N2 (Structure 4-niveaux)
        elif isinstance(val_level2, dict) and level3_key in val_level2: 
            level2_type = key_level2
            return level2_type, level3_key, level4_option
            
        # Cas 3: N4_Option est en réalité l'option finale dans une liste simple N2
        elif isinstance(val_level2, list) and level4_option in val_level2:
             level2_type = key_level2
             return level2_type, None, level4_option
             
        # Cas 4: Le full_sub_cat (sans parenthèse) est un élément N4 dans une liste N3 (cas où N4 = N3)
        elif isinstance(val_level2, list) and level3_key in val_level2:
             level2_type = key_level2
             return level2_type, None, level3_key # Ici, N3_key est l'option N4

    # Fallback: Si rien n'a été trouvé, on retourne l'état non résolu
    return None, None, None


# --- FONCTION POUR LA CLASSIFICATION EN CASCADE (MISE À JOUR) ---
def _render_visa_classification_form(
    key_suffix: str, 
    visa_structure: Dict, # Prend la structure en argument
    initial_category: Optional[str] = None, 
    initial_type: Optional[str] = None, 
    initial_level3_key: Optional[str] = None, 
    initial_level4_option: Optional[str] = None
) -> Tuple[str, str]:
    """
    Affiche les selectbox en cascade pour la classification des visas, en utilisant 
    la structure dynamique.
    """
    
    col_cat, col_type = st.columns(2) 
    
    main_keys = list(visa_structure.keys())
    default_cat_index = main_keys.index(initial_category) + 1 if initial_category in main_keys else 0
    
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
        
        selected_options = visa_structure.get(visa_category, {})
        visa_types_list = list(selected_options.keys())
        
        default_type_index = visa_types_list.index(initial_type) + 1 if initial_type in visa_types_list else 0
        
        # 2. Sélection du Type de Visa (Niveau 2)
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
                # Cas 1 : Niveau 3 (Liste simple) - Structure 3 Niveaux
                
                options_list = [opt for opt in current_options if opt] # Filtrer les chaînes vides
                
                if not options_list:
                    # Si la liste est vide, c'est que le N2 est le type final
                    final_visa_type = selected_type
                else:
                    st.subheader(f"3. Option pour **{selected_type}**")
                    
                    default_sub_index = options_list.index(initial_level4_option) if initial_level4_option in options_list else 0
                    
                    final_selection = st.radio(
                        "Choisissez l'option finale",
                        options_list,
                        index=default_sub_index,
                        key=skey("form", key_suffix, "sub1"),
                        horizontal=True
                    )
                    final_visa_type = f"{selected_type} ({final_selection})"
                
            elif isinstance(current_options, dict):
                # Cas 2 : Niveau 3 (Dictionnaire/Sous-catégories) - Structure 4 Niveaux
                st.subheader(f"3. Sous-catégorie pour **{selected_type}**")
                
                nested_keys = list(current_options.keys())
                nested_key_to_select = initial_level3_key if initial_level3_key in nested_keys else ""
                
                default_nested_index = nested_keys.index(nested_key_to_select) + 1 if nested_key_to_select in nested_keys else 0

                # --- Niveau 3 (Sous-catégorie) ---
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
                        options_list_nested = [opt for opt in nested_options if opt] # Filtrer les chaînes vides
                        
                        if options_list_nested:
                            st.subheader(f"4. Option finale pour **{nested_key}**")
                            
                            default_sub_index = options_list_nested.index(initial_level4_option) if initial_level4_option in options_list_nested else 0
                            
                            # --- Niveau 4 (Option finale) ---
                            final_selection = st.radio(
                                "Choisissez l'option finale",
                                options_list_nested,
                                index=default_sub_index,
                                key=skey("form", key_suffix, "sub2"),
                                horizontal=True
                            )
                            final_visa_type = f"{nested_key} ({final_selection})"
                        
                        else:
                            # Cas où le Niveau 3 est la valeur finale
                            final_visa_type = nested_key
                    else:
                         final_visa_type = nested_key
                else:
                    final_visa_type = selected_type
            
            else:
                 final_visa_type = selected_type

    # Retourne la Catégorie (Niveau 1) et la Sous-Catégorie (Niveau final détaillé)
    return visa_category, final_visa_type

# --- Fonctions de Résumé et CRUD ---

def _summarize_data(df: pd.DataFrame) -> Dict[str, Any]:
    """Calcule les métriques clés pour l'affichage."""
    if df.empty:
        return {
            "total_clients": 0, "clients_actifs": 0, "clients_payés": 0,
            "total_honoraires": 0.0, "total_payé": 0.0, "solde_du": 0.0
        }
    
    # Assurer les types float pour le calcul et utiliser np.nansum
    df['montant'] = pd.to_numeric(df['montant'], errors='coerce').fillna(0.0)
    df['payé'] = pd.to_numeric(df['payé'], errors='coerce').fillna(0.0)
    df['solde'] = pd.to_numeric(df['solde'], errors='coerce').fillna(0.0)
    
    total_honoraires = np.nansum(df['montant'])
    total_payé = np.nansum(df['payé'])
    solde_du = np.nansum(df['solde'])
    
    clients_actifs = len(df) # Simplification: tous les clients sont actifs
    clients_payés = (df['solde'] <= 0).sum()
    
    return {
        "total_clients": len(df),
        "clients_actifs": clients_actifs,
        "clients_payés": clients_payés,
        "total_honoraires": float(total_honoraires),
        "total_payé": float(total_payé),
        "solde_du": float(solde_du)
    }

def _update_client_data(df: pd.DataFrame, new_data: Dict[str, Any], action: str) -> pd.DataFrame:
    """Ajoute, Modifie ou Supprime un client. Centralisation des actions CRUD."""
    
    dossier_n = str(new_data.get('dossier_n')).strip()
    
    if not dossier_n or dossier_n.lower() in ('nan', 'none', 'na', ''):
        st.error("Le Numéro de Dossier ne peut pas être vide ou non défini.")
        return df

    # --- Actions DELETE ---
    if action == "DELETE":
        if 'dossier_n' not in df.columns: return df
            
        # Assurer que la colonne dossier_n est traitée comme une chaîne pour la comparaison
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
            
            for col in new_df_row.columns:
                if col in df.columns:
                    df.loc[idx_to_modify, col] = new_df_row[col].iloc[0]
                else:
                    # Ajouter la colonne manquante si elle n'existe pas dans df
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
        
        # S'assurer que les colonnes de la nouvelle ligne existent dans df
        for col in new_df_row.columns:
            if col not in df.columns:
                df[col] = pd.NA
        
        updated_df = pd.concat([df, new_df_row], ignore_index=True)
        st.cache_data.clear() 
        st.success(f"Dossier Client '{new_data.get('nom')}' (N° {dossier_n}) ajouté avec succès ! Rafraîchissement des statistiques en cours...")
        return updated_df
        
    return df

# --- Fonctions de l'Application ---

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
        # Stockage des données binaires
        st.session_state[skey("raw_clients_content")] = uploaded_file_clients.read()
        st.session_state[skey("clients_name")] = uploaded_file_clients.name
        # On vide le DF pour forcer le rechargement
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
        # Stockage des données binaires
        st.session_state[skey("raw_visa_content")] = uploaded_file_visa.read()
        st.session_state[skey("visa_name")] = uploaded_file_visa.name
        # On vide le DF pour forcer le rechargement
        st.session_state[skey("df_visa")] = pd.DataFrame() 
        st.sidebar.success(f"Visa : **{uploaded_file_visa.name}** chargé.")
    elif content_visa_loaded:
        st.sidebar.success(f"Visa : **{st.session_state.get(skey('visa_name'), 'Précédent')}** (Persistant)")


def data_processing_flow():
    """Gère le chargement, le nettoyage et le stockage des DataFrames."""
    
    header_clients = st.session_state.get(skey("header_clients_row"), 0)
    header_visa = st.session_state.get(skey("header_visa_row"), 0)
    
    # --- Clients ---
    raw_clients_content = st.session_state.get(skey("raw_clients_content"))
    df_clients_current = st.session_state.get(skey("df_clients"))

    if raw_clients_content is not None and df_clients_current.empty:
        with st.spinner("Traitement des données Clients..."):
            try:
                # Lire le contenu brut
                df_raw = _read_data_file(BytesIO(raw_clients_content), st.session_state[skey("clients_name")], header_clients)
                # Nettoyer les données
                df_cleaned = _clean_clients_data(df_raw)
                # Stocker le DataFrame nettoyé (seulement si le nettoyage a réussi)
                if not df_cleaned.empty:
                    st.session_state[skey("df_clients")] = df_cleaned
                else:
                    st.error("Échec du traitement des données Clients. Vérifiez le format/l'en-tête.")
            except Exception as e:
                st.error(f"Erreur fatale lors du traitement des données Clients: {e}")
                st.session_state[skey("raw_clients_content")] = None 

    # --- Visa ---
    raw_visa_content = st.session_state.get(skey("raw_visa_content"))
    df_visa_current = st.session_state.get(skey("df_visa"))

    if raw_visa_content is not None and df_visa_current.empty:
        with st.spinner("Traitement des données Visa..."):
            try:
                # Lire le contenu brut
                df_raw_visa = _read_data_file(BytesIO(raw_visa_content), st.session_state[skey("visa_name")], header_visa)
                # Nettoyer les données (surtout conversion en string)
                df_cleaned_visa = _clean_visa_data(df_raw_visa)
                # Stocker le DataFrame nettoyé
                if not df_cleaned_visa.empty:
                    st.session_state[skey("df_visa")] = df_cleaned_visa
                else:
                    st.error("Échec du traitement des données Visa. Vérifiez le format/l'en-tête.")
            except Exception as e:
                st.error(f"Erreur fatale lors du traitement des données Visa: {e}")
                st.session_state[skey("raw_visa_content")] = None


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
            # S'assurer que les valeurs sont numériques avant le formatage
            df_accounting[col] = pd.to_numeric(df_accounting[col], errors='coerce').fillna(0.0)
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


# --- GESTION DES DOSSIERS (AJOUT/MODIF/SUPPRESSION) ---
def dossier_management_tab(df_clients: pd.DataFrame, visa_structure: Dict): # Prend la structure en argument
    """Contenu de l'onglet Saisie/Modification/Suppression de Dossiers."""
    st.header("📝 Gestion des Dossiers Clients (CRUD)")
    
    if not visa_structure:
        st.warning("Veuillez charger votre fichier Visa (Table de Référence) pour activer la classification de visa.")
        return

    tab_add, tab_modify, tab_delete = st.tabs(["➕ Ajouter un Dossier", "✍️ Modifier un Dossier", "🗑️ Supprimer un Dossier"])

    # =========================================================================
    # LOGIQUE D'AJOUT (ADD)
    # =========================================================================
    with tab_add:
        next_dossier_n = 13000
        if not df_clients.empty and 'dossier_n' in df_clients.columns:
            try:
                # Extraire le numéro de dossier maximum pour trouver le suivant
                # Utiliser des méthodes plus robustes pour extraire les nombres
                numeric_dossiers = df_clients['dossier_n'].astype(str).str.extract(r'(\d+)').astype(float)
                max_n = numeric_dossiers[pd.notna(numeric_dossiers)].max()
                next_dossier_n = int(max_n + 1) if not pd.isna(max_n) and max_n >= 12000 else 13000
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
            
            solde_calcule = (montant_facture if montant_facture is not None else 0.0) - (paye_initial if paye_initial is not None else 0.0)
            st.metric("Solde Initial Dû (Calculé)", f"${solde_calcule:,.2f}".replace(",", " "))
            
            st.markdown("---")
            st.subheader("Classification de Visa Hiérarchique")
            
            # --- APPEL DE LA CLASSIFICATION EN CASCADE (DYNAMIQUE) ---
            visa_category, visa_type = _render_visa_classification_form(key_suffix="add", visa_structure=visa_structure)
            
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

        client_options = {f"{r['dossier_n']} - {r['nom']}": r['dossier_n'] for _, r in df_clients[['dossier_n', 'nom']].iterrows() if pd.notna(r['dossier_n'])}
        selected_key = st.selectbox(
            "Sélectionner le Dossier à Modifier",
            [""] + list(client_options.keys()),
            key=skey("modify", "select_client")
        )

        selected_dossier_n = client_options.get(selected_key)
        
        if selected_dossier_n:
            # S'assurer que le dossier N° existe dans le DataFrame après filtrage
            matching_rows = df_clients[df_clients['dossier_n'].astype(str) == selected_dossier_n]

            # CORRECTION LOGIQUE: Vérifier si la ligne existe avant d'y accéder (previent IndexError)
            if not matching_rows.empty:
                current_data = matching_rows.iloc[0].to_dict()
                
                st.markdown(f"---")
                st.info(f"Modification du Dossier N°: **{selected_dossier_n}**")

                with st.form("modify_client_form"):
                    
                    # --- Remplissage des champs (nom, date, financier) ---
                    # Ligne corrigée pour l'indentation
                    col_name, col_date = st.columns(2) 
                    
                    # Utilisation des clés en minuscules
                    client_name_mod = col_name.text_input("Nom du Client", 
                                                         value=current_data.get('nom', ''), 
                                                         key=skey("form_mod", "nom"))
                    
                    # Gestion de la date
                    date_val = current_data.get('date')
                    if pd.isna(date_val): date_val = pd.to_datetime('today').date()
                    elif isinstance(date_val, pd.Timestamp): date_val = date_val.date()
                    date_dossier_mod = col_date.date_input("Date d'Ouverture du Dossier", 
                                                           value=date_val, 
                                                           key=skey("form_mod", "date"))
                    
                    st.markdown("---")
                    col_montant, col_paye = st.columns(2)
                    
                    # CORRECTION INITALISATION: Sécuriser les valeurs numériques pour le number_input
                    montant_initial = float(current_data.get('montant', 0.0) or 0.0)
                    paye_initial = float(current_data.get('payé', 0.0) or 0.0)
                    
                    montant_facture_mod = col_montant.number_input("Total Facturé (Montant)", 
                                                                   min_value=0.0, step=100.0, 
                                                                   value=montant_initial, 
                                                                   key=skey("form_mod", "montant"))
                    paye_mod = col_paye.number_input("Total Paiements Reçus (Payé)", 
                                                     min_value=0.0, step=100.0, 
                                                     value=paye_initial, 
                                                     key=skey("form_mod", "payé"))
                    
                    solde_mod = (montant_facture_mod if montant_facture_mod is not None else 0.0) - (paye_mod if paye_mod is not None else 0.0)
                    st.metric("Solde Actuel Dû (Calculé)", f"${solde_mod:,.2f}".replace(",", " "))
                    
                    st.markdown("---")
                    st.subheader("Classification de Visa Hiérarchique")
                    
                    # Préparation des valeurs initiales pour la cascade
                    current_cat = str(current_data.get('categorie', ''))
                    full_sub_cat = str(current_data.get('sous_categorie', ''))
                    
                    # --- APPEL DE LA FONCTION DE RÉSOLUTION DYNAMIQUE ---
                    level2_type, level3_key, level4_option = _resolve_visa_levels(current_cat, full_sub_cat, visa_structure)

                    # --- APPEL DE LA CLASSIFICATION EN CASCADE AVEC VALEURS INITIALES ET STRUCTURE DYNAMIQUE ---
                    visa_category_mod, visa_type_mod = _render_visa_classification_form(
                        key_suffix="mod",
                        visa_structure=visa_structure, 
                        initial_category=current_cat, 
                        initial_type=level2_type, 
                        initial_level3_key=level3_key, 
                        initial_level4_option=level4_option, 
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
            else:
                 st.error(f"Erreur : Dossier N° **{selected_dossier_n}** introuvable dans la base de données actuelle.")


    # =========================================================================
    # LOGIQUE DE SUPPRESSION (DELETE)
    # =========================================================================
    with tab_delete:
        st.subheader("Supprimer un Dossier Définitivement")
        st.warning("ATTENTION : Cette action est irréversible.")
        
        if df_clients.empty or 'dossier_n' not in df_clients.columns:
            st.info("Aucun dossier client chargé ou créé.")
            return

        client_options = {f"{r['dossier_n']} - {r['nom']}": r['dossier_n'] for _, r in df_clients[['dossier_n', 'nom']].iterrows() if pd.notna(r['dossier_n'])}
             
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
# =========================================================================

def main():
    """Fonction principale de l'application Streamlit."""
    st.set_page_config(
        page_title=APP_TITLE,
        layout="wide",
        initial_sidebar_state="expanded"
    )
    st.title(APP_TITLE)
    
    # --- Initialisation garantie de la Session d'État ---
    st.session_state.setdefault(skey("raw_clients_content"), None)
    st.session_state.setdefault(skey("clients_name"), "")
    st.session_state.setdefault(skey("df_clients"), pd.DataFrame())
    
    st.session_state.setdefault(skey("raw_visa_content"), None)
    st.session_state.setdefault(skey("visa_name"), "")
    st.session_state.setdefault(skey("df_visa"), pd.DataFrame())
    
    st.session_state.setdefault(skey("header_clients_row"), 0)
    st.session_state.setdefault(skey("header_visa_row"), 0)
    # --------------------------------------------------------------------------
    
    # 1. Section de chargement des fichiers
    upload_section()
    
    # 2. Flux de traitement des données
    data_processing_flow()
    
    # Récupérer les DataFrames nettoyés
    df_clients = st.session_state.get(skey("df_clients"), pd.DataFrame())
    df_visa = st.session_state.get(skey("df_visa"), pd.DataFrame())

    # --- Étape CRUCIALE : Construire la structure à partir du fichier Visa ---
    # Si le fichier visa est chargé, on génère le dictionnaire dynamique
    visa_structure = VISA_STRUCTURE 
    if not df_visa.empty:
        # Tenter de générer la structure, sinon utiliser un dict vide
        try:
             # Utilisation de la version corrigée de la fonction
             visa_structure = _build_visa_structure(df_visa)
        except Exception as e:
            # st.error(f"Erreur de construction de la structure Visa: {e}")
            visa_structure = {}
    
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
        # Passage du dictionnaire dynamique à la fonction de gestion
        dossier_management_tab(df_clients, visa_structure) 

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
