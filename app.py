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

SIMPLE_SERVICE_OPTIONS = {
    "Derivatives": None, "Travel Permit": None, "Work Permit": None, "I-751": None, 
    "Re-entry Permit": None, "I-90": None, "Consultation": None, 
    "Analysis": None, "Referral": None, "I-407": None,
    "Naturalization": ["Traditional", "Marriage"],
    "Other": ["Détail à écrire dans une case"],
}

# =========================
# Fonctions utilitaires de DataFrames
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
    # NOTE: On utilise 'montant' comme Montant Facturé et 'payé' comme Total Reçu
    if 'montant' in df.columns and 'payé' in df.columns:
        df['solde'] = df['montant'] - df['payé']
    elif 'honoraires' in df.columns and 'payé' in df.columns: # Fallback pour les anciens fichiers
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
    required_cols = ['dossier_n', 'nom', 'categorie', 'sous_categorie', 'montant', 'payé', 'solde', 'date']
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
            
        # Trouver la ligne correspondante
        matching_rows = df[df['dossier_n'].astype(str) == dossier_n]
        if not matching_rows.empty:
            idx_to_modify = matching_rows.index[0]
            
            # Mise à jour des colonnes dans le DataFrame original
            for col in new_df_row.columns:
                if col in df.columns:
                    df.loc[idx_to_modify, col] = new_df_row[col].iloc[0]
                else:
                    df[col] = pd.NA # Ajout de la colonne si elle n'existe pas
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
                df[col] = pd.NA # Ajoute la colonne manquante
        
        # Concaténation
        updated_df = pd.concat([df, new_df_row], ignore_index=True)
        st.cache_data.clear() 
        st.success(f"Dossier Client '{new_data.get('nom')}' (N° {dossier_n}) ajouté avec succès ! Rafraîchissement des statistiques en cours...")
        return updated_df
        
    return df


# =========================
# Fonctions de l'Interface Utilisateur (UI)
# =========================

# --- NOUVELLE FONCTION POUR LA CLASSIFICATION EN CASCADE ---
def _render_visa_classification_form(key_suffix: str, initial_category: Optional[str] = None, initial_type: Optional[str] = None, initial_subtype: Optional[str] = None) -> Tuple[str, str]:
    """
    Affiche les selectbox en cascade pour la classification des visas.
    Retourne la catégorie et le type de visa sélectionnés.
    """
    
    # Déterminer les index initiaux pour les valeurs par défaut
    main_keys = list(VISA_STRUCTURE.keys())
    default_cat_index = main_keys.index(initial_category) + 1 if initial_category in main_keys else 0
    
    col_cat, col_type = st.columns(2)
    
    with col_cat:
        # 1. Sélection de la Catégorie (Affaires/Tourisme, Etudiants, etc.)
        visa_category = st.selectbox(
            "1. Catégorie de Visa (Grand Groupe)",
            ["Sélectionnez un groupe"] + main_keys,
            index=default_cat_index,
            key=skey("form", key_suffix, "cat_main"),
        )
        
    final_visa_type = ""
    
    if visa_category != "Sélectionnez un groupe":
        
        selected_options = VISA_STRUCTURE.get(visa_category, {})
        visa_types_list = list(selected_options.keys())
        
        default_type_index = visa_types_list.index(initial_type) + 1 if initial_type in visa_types_list else 0
        
        with col_type:
            # 2. Sélection du Type de Visa (B-1, F-1, E-2, etc. - les points ●)
            selected_type = st.selectbox(
                f"2. Type de Visa ({visa_category})",
                ["Sélectionnez un type"] + visa_types_list,
                index=default_type_index,
                key=skey("form", key_suffix, "cat_type"),
            )

        if selected_type and selected_type != "Sélectionnez un type":
            current_options = selected_options.get(selected_type)

            if isinstance(current_options, list):
                # Cas simple : liste d'options (ex: B-1 -> COS/EOS)
                st.subheader(f"3. Option pour {selected_type}")
                
                default_sub_index = current_options.index(initial_subtype) if initial_subtype in current_options else 0
                
                final_selection = st.radio(
                    "Choisissez l'option finale",
                    current_options,
                    index=default_sub_index,
                    key=skey("form", key_suffix, "sub1"),
                    horizontal=True
                )
                final_visa_type = f"{selected_type} ({final_selection})"
                
            elif isinstance(current_options, dict):
                # Cas complexe/imbriqué : Dictionnaire (ex: E-2)
                st.subheader(f"3. Sous-catégorie pour {selected_type}")
                
                nested_keys = list(current_options.keys())
                default_nested_index = nested_keys.index(initial_type) + 1 if initial_type in nested_keys else 0

                nested_key = st.selectbox(
                    f"Sous-catégorie de {selected_type}",
                    ["Sélectionnez la sous-catégorie"] + nested_keys,
                    index=default_nested_index,
                    key=skey("form", key_suffix, "nested_key"),
                )
                
                if nested_key and nested_key != "Sélectionnez la sous-catégorie":
                    # Niveau 4 : Radio Buttons pour les options finales
                    nested_options = current_options.get(nested_key)
                    if nested_options and isinstance(nested_options, list):
                        st.subheader(f"4. Option finale pour {nested_key}")
                        
                        default_sub_index = nested_options.index(initial_subtype) if initial_subtype in nested_options else 0
                        
                        final_selection = st.radio(
                            "Choisissez l'option finale",
                            nested_options,
                            index=default_sub_index,
                            key=skey("form", key_suffix, "sub2"),
                            horizontal=True
                        )
                        final_visa_type = f"{nested_key} ({final_selection})"
            
            # Si le type sélectionné est une clé de dictionnaire profonde (ex: E-2), on utilise ce nom
            elif isinstance(selected_options.get(selected_type), dict):
                 final_visa_type = selected_type

    # Retourne les deux champs à stocker dans le DataFrame
    return visa_category, final_visa_type if final_visa_type else selected_type if selected_type and selected_type != "Sélectionnez un type" else ""

# ... (les fonctions upload_section, data_processing_flow, home_tab, accounting_tab sont inchangées)

# --- GESTION DES DOSSIERS (AJOUT/MODIF/SUPPRESSION) ---
def dossier_management_tab(df_clients: pd.DataFrame):
    """Contenu de l'onglet Saisie/Modification/Suppression de Dossiers."""
    st.header("📝 Gestion des Dossiers Clients (CRUD)")
    
    tab_add, tab_modify, tab_delete = st.tabs(["➕ Ajouter un Dossier", "✍️ Modifier un Dossier", "🗑️ Supprimer un Dossier"])

    # =========================================================================
    # LOGIQUE D'AJOUT (ADD)
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
                        "sous_categorie": visa_type, # Contient le résultat complet de la cascade
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
                date_dossier_mod = col_date.date_input(
                    "Date d'Ouverture du Dossier", 
                    value=pd.to_datetime(current_data.get('date', pd.to_datetime('today'))).date(), 
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
                
                solde_mod = montant_facture_mod - paye_mod
                st.metric("Solde Actuel Dû (Calculé)", f"${solde_mod:,.2f}".replace(",", " "))
                
                st.markdown("---")
                st.subheader("Classification de Visa Hiérarchique")
                
                # Récupérer les valeurs initiales pour pré-remplir la cascade
                current_cat = current_data.get('categorie')
                current_sub_cat = current_data.get('sous_categorie')

                # --- APPEL DE LA CLASSIFICATION EN CASCADE AVEC VALEURS INITIALES ---
                visa_category_mod, visa_type_mod = _render_visa_classification_form(
                    key_suffix="mod",
                    initial_category=current_cat,
                    initial_type=current_sub_cat.split(" (")[0] if isinstance(current_sub_cat, str) and "(" in current_sub_cat else current_sub_cat,
                    initial_subtype=current_sub_cat.split(" (")[1].replace(")", "") if isinstance(current_sub_cat, str) and "(" in current_sub_cat else None
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
            
            # Utilisation de la condition pour afficher la confirmation
            delete_confirmed = False
            if selected_dossier_n_del:
                delete_confirmed = st.checkbox(f"Je confirme la suppression définitive du dossier N° **{selected_dossier_n_del}**", key=skey("delete", "confirm"))
            
            submitted_del = st.form_submit_button("💣 SUPPRIMER le Dossier", disabled=not selected_dossier_n_del or not delete_confirmed)
            
            if submitted_del and delete_confirmed:
                delete_entry = {"dossier_n": selected_dossier_n_del}
                
                updated_df_clients = _update_client_data(df_clients, delete_entry, "DELETE")
                st.session_state[skey("df_clients")] = updated_df_clients
                st.rerun()

# ... (Le reste des fonctions est inchangé)

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
