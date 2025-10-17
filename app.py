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
SID = "vmgr_v4"

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
    
    # 1. Standardiser et convertir les nombres financiers (Montant Facturé et Payé)
    # NOTE: L'ancien fichier peut avoir 'montant' comme total facturé et 'honoraires' comme payé 
    # ou inversement. Nous standardisons sur 'montant' (Total Facturé) et 'payé' (Total Reçu).
    money_cols = ['honoraires', 'payé', 'solde', 'acompte_1', 'acompte_2', 'montant', 'autres_frais_us_']
    
    for col in money_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
            df[col] = df[col].str.replace(r'[^\d.]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float) 
    
    # 2. Rétablir le solde avec la formule (Montant Facturé - Total Payé) si les deux colonnes existent
    if 'montant' in df.columns and 'payé' in df.columns:
        df['solde'] = df['montant'] - df['payé']

    # 3. Conversion des Dates
    date_cols = ['date', 'dossier_envoyé', 'dossier_approuvé', 'dossier_refusé', 'dossier_annulé']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # 4. Colonne dérivée
    if 'date' in df.columns:
         df['jours_ecoules'] = (pd.to_datetime('today') - df['date']).dt.days

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
    """Calcule des indicateurs clés à partir du DataFrame Clients (robuste aux colonnes manquantes)."""
    
    if df.empty:
        return {"total_clients": 0, "total_honoraires": 0.0, "total_payé": 0.0, "solde_du": 0.0, "clients_actifs": 0, "clients_payés": 0}

    # Calculs financiers robustes
    total_honoraires = df['montant'].sum() if 'montant' in df.columns else 0.0
    total_payé = df['payé'].sum() if 'payé' in df.columns else 0.0
    solde_du = df['solde'].sum() if 'solde' in df.columns else 0.0
    clients_payés = (df['solde'] <= 0).sum() if 'solde' in df.columns else 0
    
    # Logique robuste pour les clients actifs
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
    
    dossier_n = str(new_data.get('dossier_n'))
    
    # 1. Action DELETE
    if action == "DELETE":
        if 'dossier_n' not in df.columns:
            st.error("Colonne 'dossier_n' introuvable dans le DataFrame pour la suppression.")
            return df
            
        idx_to_delete = df[df['dossier_n'] == dossier_n].index
        
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
    
    # S'assurer que les valeurs numériques sont des floats pour le calcul du solde
    money_cols = ['payé', 'montant'] 
    for col in money_cols:
        if col in new_df_row.columns:
             new_df_row[col] = pd.to_numeric(new_df_row[col], errors='coerce').fillna(0.0).astype(float)
    
    # Calcul du Solde
    montant = new_df_row['montant'].iloc[0] if 'montant' in new_df_row.columns else 0.0
    paye = new_df_row['payé'].iloc[0] if 'payé' in new_df_row.columns else 0.0
    new_df_row['solde'] = montant - paye
    
    # 2. Action MODIFY
    if action == "MODIFY":
        if 'dossier_n' not in df.columns:
            st.error("Colonne 'dossier_n' introuvable dans le DataFrame pour la modification.")
            return df
            
        matching_rows = df[df['dossier_n'] == dossier_n]
        if not matching_rows.empty:
            idx_to_modify = matching_rows.index[0]
            # Mettre à jour l'original DataFrame par index
            for col in new_df_row.columns:
                 # S'assurer que la colonne existe dans l'original avant d'assigner
                if col in df.columns:
                    df.loc[idx_to_modify, col] = new_df_row[col].iloc[0]
            
            st.cache_data.clear() 
            st.success(f"Dossier N° {dossier_n} modifié avec succès.")
            return df
        else:
            st.warning(f"Dossier N° {dossier_n} introuvable pour modification.")
            return df

    # 3. Action ADD
    if action == "ADD":
        # S'assurer que le nouveau client n'existe pas déjà
        if 'dossier_n' in df.columns and (df['dossier_n'] == dossier_n).any():
             st.error(f"Le Dossier N° {dossier_n} existe déjà. Utilisez l'onglet 'Modifier'.")
             return df
        
        # S'assurer que toutes les colonnes de la nouvelle ligne existent dans le DF cible
        for col in new_df_row.columns:
            if col not in df.columns:
                # Ajoute la colonne manquante au DataFrame existant, remplie de NA
                df[col] = pd.NA 
        
        # Concaténation
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
    
    uploaded_file_clients = st.sidebar.file_uploader(
        "Clients/Dossiers (.csv, .xlsx)",
        type=['csv', 'xlsx'],
        key=skey("upload", "clients"),
    )
    
    uploaded_file_visa = st.sidebar.file_uploader(
        "Table de Référence Visa (.csv, .xlsx)",
        type=['csv', 'xlsx'],
        key=skey("upload", "visa"),
    )

    if uploaded_file_clients:
        st.session_state[skey("raw_clients_content")] = uploaded_file_clients.read()
        st.session_state[skey("clients_name")] = uploaded_file_clients.name
        
    if uploaded_file_visa:
        st.session_state[skey("raw_visa_content")] = uploaded_file_visa.read()
        st.session_state[skey("visa_name")] = uploaded_file_visa.name

def data_processing_flow():
    """Gère le chargement, le nettoyage et le stockage des DataFrames."""
    
    st.session_state.setdefault(skey("df_clients"), pd.DataFrame())
    st.session_state.setdefault(skey("df_visa"), pd.DataFrame())

    # --- 1. Clients ---
    content_clients = st.session_state.get(skey("raw_clients_content"))
    file_name_clients = st.session_state.get(skey("clients_name"), "")
    header_clients = st.session_state.get(skey("header_clients_row"), 0)

    if content_clients and file_name_clients:
        df_raw_clients = _read_data_file(BytesIO(content_clients), file_name_clients, header_row=header_clients)
        if not df_raw_clients.empty:
            df_cleaned_clients = _clean_clients_data(df_raw_clients)
            st.session_state[skey("df_clients")] = df_cleaned_clients
        else:
             st.session_state[skey("df_clients")] = pd.DataFrame()
    else:
        st.session_state[skey("df_clients")] = pd.DataFrame()

    # --- 2. Visa ---
    content_visa = st.session_state.get(skey("raw_visa_content"))
    file_name_visa = st.session_state.get(skey("visa_name"), "")
    header_visa = st.session_state.get(skey("header_visa_row"), 0)

    if content_visa and file_name_visa:
        df_raw_visa = _read_data_file(BytesIO(content_visa), file_name_visa, header_row=header_visa)
        if not df_raw_visa.empty:
            df_cleaned_visa = _clean_visa_data(df_raw_visa)
            st.session_state[skey("df_visa")] = df_cleaned_visa
        else:
             st.session_state[skey("df_visa")] = pd.DataFrame()
    else:
        st.session_state[skey("df_visa")] = pd.DataFrame()

# --- Onglet Accueil ---
def home_tab(df_clients: pd.DataFrame):
    """Contenu de l'onglet Accueil/Statistiques."""
    st.header("📊 Statistiques Clés")
    
    if df_clients.empty:
        st.info("Veuillez charger ou ajouter des dossiers clients pour afficher les statistiques.")
        return
        
    summary = _summarize_data(df_clients)

    col1, col2, col3, col4 = st.columns(4)

    # Affichage des métriques
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

    # 2. Tableau de ventilation
    st.subheader("Détail du Compte Client")
    
    accounting_cols = ['dossier_n', 'nom', 'categorie', 'montant', 'payé', 'solde', 'date']
    valid_cols = [col for col in accounting_cols if col in df_clients.columns]
    
    df_accounting = df_clients[valid_cols].copy()
    
    # Formatage des colonnes monétaires pour l'affichage (optionnel, Streamlit fait déjà bien)
    if 'montant' in df_accounting.columns:
        df_accounting['Montant Facturé'] = df_accounting['montant'].apply(lambda x: f"${x:,.2f}".replace(",", " "))
    if 'payé' in df_accounting.columns:
        df_accounting['Total Payé'] = df_accounting['payé'].apply(lambda x: f"${x:,.2f}".replace(",", " "))
    if 'solde' in df_accounting.columns:
        df_accounting['Solde Dû'] = df_accounting['solde'].apply(lambda x: f"${x:,.2f}".replace(",", " "))

    # Colonnes à afficher dans le tableau
    display_cols = ['dossier_n', 'nom', 'categorie', 'Montant Facturé', 'Total Payé', 'Solde Dû', 'date']
    display_cols = [col for col in display_cols if col in df_accounting.columns]
    
    st.dataframe(
        df_accounting[display_cols].sort_values(by='Solde Dû', ascending=False), 
        use_container_width=True,
    )
    st.caption("Le solde dû est calculé par `Montant Facturé (colonne montant) - Total Payé (colonne payé)`.")

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
                # Tente de trouver le max numérique pour le numéro de dossier
                # On utilise .str.extract pour être robuste si le N° contient des lettres/tirets
                max_n = df_clients['dossier_n'].astype(str).str.extract(r'(\d+)').astype(float).max()
                next_dossier_n = int(max_n + 1) if not pd.isna(max_n) else 13000
            except:
                 next_dossier_n = 13000 # Fallback

        
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
            st.subheader("Classification de Visa")
            
            col_cat, col_type = st.columns(2)
            visa_category = col_cat.selectbox(
                "Catégorie de Visa",
                ["Sélectionnez un groupe"] + list(VISA_STRUCTURE.keys()),
                key=skey("form_add", "categorie"),
            )
            visa_type = col_type.text_input("Sous-catégorie / Type de Visa (Entrée Manuelle)", key=skey("form_add", "sous_categorie"), placeholder="Ex: B-1, E-2 Inv., I-130 USC...")
            
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
                    st.rerun() # Rafraîchissement pour effacer le form et mettre à jour les vues
    
    # =========================================================================
    # LOGIQUE DE MODIFICATION (MODIFY)
    # =========================================================================
    with tab_modify:
        st.subheader("Modifier un Dossier Existant")
        
        if df_clients.empty:
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
            # Récupérer les données du client sélectionné
            current_data = df_clients[df_clients['dossier_n'] == selected_dossier_n].iloc[0].to_dict()
            
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
                
                # Récupérer la catégorie existante ou utiliser un défaut
                current_cat = current_data.get('categorie', 'Sélectionnez un groupe')
                if current_cat not in VISA_STRUCTURE.keys(): current_cat = 'Sélectionnez un groupe'

                col_cat, col_type = st.columns(2)
                visa_category_mod = col_cat.selectbox(
                    "Catégorie de Visa",
                    ["Sélectionnez un groupe"] + list(VISA_STRUCTURE.keys()),
                    index=list(["Sélectionnez un groupe"] + list(VISA_STRUCTURE.keys())).index(current_cat),
                    key=skey("form_mod", "categorie"),
                )
                visa_type_mod = col_type.text_input(
                    "Sous-catégorie / Type de Visa (Entrée Manuelle)", 
                    value=current_data.get('sous_categorie', current_data.get('visa', '')), 
                    key=skey("form_mod", "sous_categorie"), 
                    placeholder="Ex: B-1, E-2 Inv., I-130 USC..."
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
        
        if df_clients.empty:
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
            delete_confirmed = st.checkbox(f"Je confirme la suppression définitive du dossier N° **{selected_dossier_n_del}**", key=skey("delete", "confirm"))
            
            submitted_del = st.form_submit_button("💣 SUPPRIMER le Dossier", disabled=not selected_dossier_n_del or not delete_confirmed)
            
            if submitted_del and delete_confirmed:
                delete_entry = {"dossier_n": selected_dossier_n_del}
                
                updated_df_clients = _update_client_data(df_clients, delete_entry, "DELETE")
                st.session_state[skey("df_clients")] = updated_df_clients
                st.rerun()

# --- Le reste des onglets est inchangé ---
def settings_tab():
    """Contenu de l'onglet Configuration."""
    st.header("⚙️ Configuration du Chargement")
    
    st.markdown("""
        Veuillez spécifier l'index de la ligne contenant les noms de colonnes réels.
        * **0** (par défaut) : première ligne.
        * **1** : deuxième ligne, etc.
    """)
    
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
         st.rerun() 

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
        "📈 Comptabilité", # Nouvel onglet
        "📝 Gestion Dossiers", # Gestion (Ajout/Modif/Suppr)
        "⚙️ Configuration",
        "📄 Clients - Aperçu", 
        "📄 Visa - Aperçu", 
        "💾 Export",
    ])

    with tab_home:
        home_tab(df_clients)
        
    with tab_accounting:
        accounting_tab(df_clients) # Nouvel appel de fonction

    with tab_management:
        dossier_management_tab(df_clients) # Appel de la fonction de gestion (CRUD)

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
