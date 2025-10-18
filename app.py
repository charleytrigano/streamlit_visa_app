import io
import re
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple
from datetime import date, datetime

import pandas as pd
import streamlit as st
import numpy as np

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
    """
    is_excel = file_name.endswith(('.xls', '.xlsx')) or 'xlsx' in file_name.lower() or 'xls' in file_name.lower()

    # Assurez-vous que le pointeur est au début du fichier
    file_content.seek(0)

    if is_excel:
        try:
            df = pd.read_excel(file_content, header=header_row, engine='openpyxl', dtype=str)
        except Exception as e:
            st.error(f"Erreur de lecture Excel : {e}")
            return pd.DataFrame()
    else:
        try:
            df = pd.read_csv(
                file_content,
                header=header_row,
                sep=None,
                engine='python',
                encoding='utf-8',
                on_bad_lines='skip',
                dtype=str
            )
        except UnicodeDecodeError:
            try:
                file_content.seek(0)
                df = pd.read_csv(
                    file_content,
                    header=header_row,
                    sep=None,
                    engine='python',
                    encoding='latin1',
                    on_bad_lines='skip',
                    dtype=str
                )
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
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()

    money_cols = ['honoraires', 'payé', 'solde', 'acompte_1', 'acompte_2', 'montant', 'autres_frais_us_']
    for col in money_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
            df[col] = df[col].str.replace(r'[^\d.]', '', regex=True)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)

    if 'montant' in df.columns and 'payé' in df.columns:
        df['solde'] = df['montant'] - df['payé']
    elif 'honoraires' in df.columns and 'payé' in df.columns:
        df['montant'] = df['honoraires']
        df['solde'] = df['honoraires'] - df['payé']

    date_cols = ['date', 'dossier_envoyé', 'dossier_approuvé', 'dossier_refusé', 'dossier_annulé']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    required_cols = ['dossier_n', 'nom', 'categorie', 'sous_categorie', 'montant', 'payé', 'solde', 'date', 'commentaires']
    for col in required_cols:
        if col not in df.columns:
            df[col] = pd.NA

    if 'dossier_n' in df.columns:
        df['dossier_n'] = df['dossier_n'].astype(str).str.strip()

    st.success("Nettoyage et conversion des données Clients terminés (Robuste).")
    return df


def _clean_visa_data(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoyage simple du DataFrame Visa."""
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    return df


@st.cache_data(show_spinner="Construction de la structure Visa...")
def _build_visa_structure(df_visa: pd.DataFrame) -> Dict[str, Any]:
    """
    Construit la structure de classification VISA à partir du DataFrame Visa.xlsx.
    Structure attendue : les deux premières colonnes sont N1 (Catégorie) et N2 (Type),
    les colonnes suivantes sont des options (marquées par '1' sur chaque ligne).
    """
    if df_visa.empty:
        return {}

    df_temp = df_visa.copy()
    cols = df_temp.columns.tolist()

    if len(cols) < 2:
        st.error("Le fichier Visa doit contenir au moins les colonnes Catégorie et Sous_categories.")
        return {}

    # Renommage des deux premières colonnes
    col_map = {cols[0]: 'N1_Categorie', cols[1]: 'N2_Type'}
    df_temp.rename(columns=col_map, inplace=True)

    option_columns = cols[2:]

    # Filtrage : garder les lignes valides selon la présence d'indicateurs '1' dans les colonnes d'options,
    # sinon garder les lignes ayant N1 et N2 remplis.
    if not option_columns:
        df_valid = df_temp.dropna(subset=['N1_Categorie', 'N2_Type']).copy()
    else:
        df_options = df_temp[option_columns].astype(str).fillna('').replace('nan', '')
        has_indicator = df_options.apply(lambda row: '1' in row.values, axis=1)
        # On garde toutes les lignes qui ont soit un indicateur, soit Catégorie/Type remplis
        df_valid = df_temp[has_indicator | df_temp[['N1_Categorie', 'N2_Type']].notna().all(axis=1)].copy()


    if df_valid.empty:
        df_valid = df_temp.dropna(subset=['N1_Categorie', 'N2_Type']).copy()

    structure: Dict[str, Dict[str, List[str]]] = {}

    for _, row in df_valid.iterrows():
        n1_cat = str(row.get('N1_Categorie', '')).strip()
        n2_type = str(row.get('N2_Type', '')).strip()

        if not n1_cat or not n2_type:
            continue

        if n1_cat not in structure:
            structure[n1_cat] = {}
        if n2_type not in structure[n1_cat]:
            structure[n1_cat][n2_type] = []

        for col_name in option_columns:
            # Vérifie si la colonne d'option est activée ('1')
            if str(row.get(col_name)).strip() == '1':
                option = str(col_name).strip()
                if option and option not in structure[n1_cat][n2_type]:
                    structure[n1_cat][n2_type].append(option)

    st.success("Structure de classification Visa construite dynamiquement (Matrice corrigée).")
    return structure


def _resolve_visa_levels(category: str, full_sub_cat: str, visa_structure: Dict) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Résout les niveaux de classification à partir des valeurs stockées dans les dossiers.
    Retourne (niveau2_type, niveau3_key (toujours None ici), niveau4_option)
    """
    if not category or category not in visa_structure:
        return None, None, None

    sub_cat_stripped = str(full_sub_cat).strip()

    # Si la sous-catégorie contient une option entre parenthèses (format "Type (Option)"), on la sépare
    match_paren = re.search(r'\((.+)\)', sub_cat_stripped)
    if match_paren:
        level4_option = match_paren.group(1).strip()
        level2_type_search = sub_cat_stripped[:match_paren.start()].strip()
    else:
        level4_option = None
        level2_type_search = sub_cat_stripped

    level2_options = visa_structure.get(category, {})

    if level2_type_search in level2_options:
        # cas où la valeur sauvegardée est le type N2
        if level4_option and level4_option in level2_options[level2_type_search]:
            return level2_type_search, None, level4_option
        if not level4_option or not level2_options[level2_type_search]:
            return level2_type_search, None, None

    # fallback : la valeur sauvegardée peut être seulement une option N4 (sans N2 extrait)
    for key_level2, val_level2 in level2_options.items():
        if isinstance(val_level2, list) and level2_type_search in val_level2:
            return key_level2, None, level2_type_search

    return None, None, None


def _render_visa_classification_form(
    key_suffix: str,
    visa_structure: Dict,
    initial_category: Optional[str] = None,
    initial_type: Optional[str] = None,
    initial_level3_key: Optional[str] = None,
    initial_level4_option: Optional[str] = None
) -> Tuple[str, str]:
    """
    Affiche les selectbox en cascade pour la classification des visas et renvoie
    (categorie_n1, sous_categorie_finale) où sous_categorie_finale a le format:
    - "Type" si pas d'option N4
    - "Type (Option)" si option finale choisie (ex: 'E-2 (1-COS)')
    """
    col_cat, col_type = st.columns(2)

    main_keys = list(visa_structure.keys())
    default_cat_index = main_keys.index(initial_category) + 1 if initial_category in main_keys else 0

    visa_category = initial_category if initial_category in main_keys else "Sélectionnez un groupe"
    final_visa_type = ""
    selected_type = ""

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
                # Les options N4 sont les "boutons bascules" (COS, EOS, etc.)
                options_list = [opt for opt in current_options if opt]

                if not options_list:
                    # N2 final (pas d'options N4 définies dans la matrice)
                    final_visa_type = selected_type
                else:
                    st.subheader(f"3. Option de Dépôt / Classification pour **{selected_type}**")
                    st.caption("Ces options correspondent aux colonnes `1-COS`, `2-EOS`, etc. dans votre matrice Visa.")
                    
                    # Définir l'index initial pour les boutons radio
                    default_sub_index = 0
                    if initial_level4_option in options_list:
                        default_sub_index = options_list.index(initial_level4_option)

                    final_selection = st.radio(
                        "Choisissez l'option finale",
                        options_list,
                        index=default_sub_index,
                        key=skey("form", key_suffix, "sub1"),
                        horizontal=True # Affichage en ligne des boutons bascules
                    )
                    # Format de sauvegarde : Type (Option) -> Ex: "E-2 (1-COS)"
                    final_visa_type = f"{selected_type} ({final_selection})"
            else:
                final_visa_type = selected_type

    return visa_category, final_visa_type


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

    clients_actifs = len(df)  # Simplification: tous les clients sont actifs
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

    # DELETE
    if action == "DELETE":
        if 'dossier_n' not in df.columns:
            return df

        idx_to_delete = df[df['dossier_n'].astype(str) == dossier_n].index

        if not idx_to_delete.empty:
            df = df.drop(idx_to_delete).reset_index(drop=True)
            st.cache_data.clear()
            st.success(f"Dossier N° {dossier_n} supprimé avec succès.")
            return df
        else:
            st.warning(f"Dossier N° {dossier_n} introuvable pour suppression.")
            return df

    # Pré-traitement pour ADD/MODIFY
    new_df_row = pd.DataFrame([new_data])
    new_df_row.columns = new_df_row.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()

    # Remplir les colonnes manquantes du DF principal si elles existent dans la nouvelle ligne
    for col in new_df_row.columns:
        if col not in df.columns:
            df[col] = pd.NA
            
    money_cols = ['payé', 'montant']
    for col in money_cols:
        if col in new_df_row.columns:
            new_df_row[col] = pd.to_numeric(new_df_row[col], errors='coerce').fillna(0.0).astype(float)

    montant = new_df_row['montant'].iloc[0] if 'montant' in new_df_row.columns else 0.0
    paye = new_df_row['payé'].iloc[0] if 'payé' in new_df_row.columns else 0.0
    new_df_row['solde'] = montant - paye

    # Traitement des dates (pour l'ajout)
    date_cols = ['date', 'dossier_envoyé', 'dossier_approuvé', 'dossier_refusé', 'dossier_annulé']
    for col in date_cols:
        if col in new_df_row.columns and new_df_row[col].iloc[0] is not None:
             try:
                 # Assure que la valeur est convertie en datetime pour correspondre au type du DF
                 new_df_row[col] = pd.to_datetime(new_df_row[col])
             except:
                 new_df_row[col] = pd.NaT # Not a Time

    # MODIFY
    if action == "MODIFY":
        if 'dossier_n' not in df.columns:
            return df

        matching_rows = df[df['dossier_n'].astype(str) == dossier_n]
        if not matching_rows.empty:
            idx_to_modify = matching_rows.index[0]

            for col in new_df_row.columns:
                df.loc[idx_to_modify, col] = new_df_row[col].iloc[0]

            st.cache_data.clear()
            st.success(f"Dossier N° {dossier_n} modifié avec succès.")
            return df
        else:
            st.warning(f"Dossier N° {dossier_n} introuvable pour modification.")
            return df

    # ADD
    if action == "ADD":
        if 'dossier_n' in df.columns and (df['dossier_n'].astype(str) == dossier_n).any():
            st.error(f"Le Dossier N° {dossier_n} existe déjà. Utilisez l'onglet 'Modifier'.")
            return df

        updated_df = pd.concat([df, new_df_row], ignore_index=True)
        st.cache_data.clear()
        st.success(f"Dossier Client '{new_data.get('nom')}' (N° {dossier_n}) ajouté avec succès ! Rafraîchissement des statistiques en cours...")
        return updated_df

    return df


def upload_section():
    """Section de chargement des fichiers (Barre latérale)."""
    st.sidebar.header("📁 Chargement des Fichiers")

    header_clients = st.sidebar.number_input(
        "Ligne d'en-tête Clients (Index 0 = 1ère ligne)",
        min_value=0, value=0, key=skey("header_clients_row")
    )
    header_visa = st.sidebar.number_input(
        "Ligne d'en-tête Visa (Index 0 = 1ère ligne)",
        min_value=0, value=0, key=skey("header_visa_row")
    )

    # Clients
    content_clients_loaded = st.session_state.get(skey("raw_clients_content"))

    uploaded_file_clients = st.sidebar.file_uploader(
        "Clients/Dossiers (.csv, .xlsx)",
        type=['csv', 'xlsx'],
        key=skey("upload", "clients"),
        accept_multiple_files=False
    )

    if uploaded_file_clients is not None:
        st.session_state[skey("raw_clients_content")] = uploaded_file_clients.read()
        st.session_state[skey("clients_name")] = uploaded_file_clients.name
        st.session_state[skey("df_clients")] = pd.DataFrame()
        st.sidebar.success(f"Clients : **{uploaded_file_clients.name}** chargé.")
    elif content_clients_loaded:
        st.sidebar.success(f"Clients : **{st.session_state.get(skey('clients_name'), 'Précédent')}** (Persistant)")

    # Visa
    content_visa_loaded = st.session_state.get(skey("raw_visa_content"))

    uploaded_file_visa = st.sidebar.file_uploader(
        "Table de Référence Visa (.csv, .xlsx)",
        type=['csv', 'xlsx'],
        key=skey("upload", "visa"),
        accept_multiple_files=False
    )

    if uploaded_file_visa is not None:
        st.session_state[skey("raw_visa_content")] = uploaded_file_visa.read()
        st.session_state[skey("visa_name")] = uploaded_file_visa.name
        st.session_state[skey("df_visa")] = pd.DataFrame()
        global VISA_STRUCTURE
        VISA_STRUCTURE = {}
        st.sidebar.success(f"Visa : **{uploaded_file_visa.name}** chargé.")
    elif content_visa_loaded and not VISA_STRUCTURE:
        st.sidebar.success(f"Visa : **{st.session_state.get(skey('visa_name'), 'Précédent')}** (Persistant)")


def data_processing_flow():
    """Gère le chargement, le nettoyage et le stockage des DataFrames."""
    header_clients = st.session_state.get(skey("header_clients_row"), 0)
    header_visa = st.session_state.get(skey("header_visa_row"), 0)

    # Clients
    raw_clients_content = st.session_state.get(skey("raw_clients_content"))
    df_clients_current = st.session_state.get(skey("df_clients"), pd.DataFrame())

    if raw_clients_content is not None and df_clients_current.empty:
        with st.spinner("Traitement des données Clients..."):
            try:
                df_raw = _read_data_file(BytesIO(raw_clients_content), st.session_state[skey("clients_name")], int(header_clients))
                df_cleaned = _clean_clients_data(df_raw)
                if not df_cleaned.empty:
                    st.session_state[skey("df_clients")] = df_cleaned
                else:
                    st.error("Échec du traitement des données Clients. Vérifiez le format/l'en-tête.")
            except Exception as e:
                st.error(f"Erreur fatale lors du traitement des données Clients: {e}")
                st.session_state[skey("raw_clients_content")] = None

    # Visa
    raw_visa_content = st.session_state.get(skey("raw_visa_content"))
    df_visa_current = st.session_state.get(skey("df_visa"), pd.DataFrame())

    if raw_visa_content is not None and df_visa_current.empty:
        with st.spinner("Traitement des données Visa..."):
            try:
                df_raw_visa = _read_data_file(BytesIO(raw_visa_content), st.session_state[skey("visa_name")], int(header_visa))
                df_cleaned_visa = _clean_visa_data(df_raw_visa)
                if not df_cleaned_visa.empty:
                    st.session_state[skey("df_visa")] = df_cleaned_visa
                    global VISA_STRUCTURE
                    VISA_STRUCTURE = _build_visa_structure(df_cleaned_visa)
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

    # KPIs
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Facturé (Montant)", f"${summary['total_honoraires']:,.2f}".replace(",", " "))
    col2.metric("Total Reçu (Payé)", f"${summary['total_payé']:,.2f}".replace(",", " "))
    col3.metric("Solde Total Dû", f"${summary['solde_du']:,.2f}".replace(",", " "))
    col4.metric("Dossiers Payés (Solde <= 0)", f"{summary['clients_payés']:,}".replace(",", " "))

    st.divider()

    # Filtre Client
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

    accounting_cols = ['dossier_n', 'nom', 'categorie', 'montant', 'payé', 'solde', 'date']
    valid_cols = [col for col in accounting_cols if col in df_filtered.columns]

    df_accounting = df_filtered[valid_cols].copy()

    for col in ['montant', 'payé', 'solde']:
        if col in df_accounting.columns:
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


def dossier_management_tab(df_clients: pd.DataFrame, visa_structure: Dict):
    """Contenu de l'onglet Saisie/Modification/Suppression de Dossiers."""
    st.header("📝 Gestion des Dossiers Clients (CRUD)")

    if not visa_structure:
        st.error("Veuillez charger votre fichier Visa (Table de Référence) pour activer la classification de visa.")
        return

    tab_add, tab_modify, tab_delete = st.tabs(["➕ Ajouter un Dossier", "✍️ Modifier un Dossier", "🗑️ Supprimer un Dossier"])

    # Préparer les options clients pour les selectbox
    client_options = {}
    if not df_clients.empty and 'dossier_n' in df_clients.columns:
        df_clients_for_select = df_clients[['dossier_n', 'nom']].dropna(subset=['dossier_n'])
        client_options = {
            f"{r['dossier_n']} - {r['nom']}": r['dossier_n']
            for _, r in df_clients_for_select.iterrows()
        }

    # ADD
    with tab_add:
        st.subheader("Ajouter un nouveau dossier client")
        
        # 1. Calcul du prochain N° de Dossier Principal (Base Number)
        next_dossier_n = 13000
        if not df_clients.empty and 'dossier_n' in df_clients.columns:
            try:
                # Extraire les numéros de base (avant le premier '-') et trouver le max
                base_dossiers = df_clients['dossier_n'].astype(str).str.split('-').str[0]
                numeric_dossiers = base_dossiers.str.extract(r'(\d+)').astype(float)
                max_n = numeric_dossiers[pd.notna(numeric_dossiers)].max()
                next_dossier_n = int(max_n + 1) if not pd.isna(max_n) and max_n >= 12000 else 13000
            except:
                next_dossier_n = 13000

        with st.form("add_client_form"):
            st.markdown("---")
            
            # --- 1. Identification du Dossier (Inclus Sous-Dossier) ---
            st.markdown("#### 1. Identification du Dossier")
            col_parent, col_date = st.columns(2)
            
            # LOGIQUE SOUS-DOSSIER
            parent_dossier_options = ["Nouveau Dossier Principal"] + list(client_options.keys())
            selected_parent_key = col_parent.selectbox(
                "Type de Dossier à Ajouter",
                parent_dossier_options,
                key=skey("form_add", "parent_select"),
                help="Sélectionnez un dossier existant pour créer un sous-dossier (ex: XXX-01)."
            )
            
            is_sub_dossier = selected_parent_key != "Nouveau Dossier Principal"
            dossier_n_value = ""
            client_name_value = ""

            if is_sub_dossier:
                base_dossier_n = client_options.get(selected_parent_key, "")
                
                # Calcule le prochain suffixe (-01, -02, etc.)
                pattern = re.compile(f"^{re.escape(base_dossier_n)}(-\d+)?$")
                sub_dossiers = df_clients['dossier_n'].astype(str).loc[df_clients['dossier_n'].astype(str).apply(lambda x: pattern.match(x) is not None)]
                
                max_suffix = 0
                if not sub_dossiers.empty:
                    suffixes = sub_dossiers.str.extract(r'-(\d+)$', expand=False).dropna().astype(int)
                    if not suffixes.empty:
                        max_suffix = suffixes.max()

                next_suffix = max_suffix + 1
                dossier_n_value = f"{base_dossier_n}-{next_suffix:02d}"

                parent_name = selected_parent_key.split(' - ', 1)[1] if ' - ' in selected_parent_key else ""
                client_name_value = parent_name # Le nom du client est pré-rempli avec le nom du dossier parent
                
                col_parent.text_input(
                    "Numéro de Dossier (Sous-Dossier Auto)", 
                    value=dossier_n_value, 
                    key=skey("form_add", "dossier_n_sub"), 
                    disabled=True 
                )
                
            else:
                dossier_n_value = str(next_dossier_n)
                col_parent.text_input(
                    "Numéro de Dossier Principal (Suggéré)", 
                    value=dossier_n_value, 
                    key=skey("form_add", "dossier_n_main")
                )
            
            # Utilise la valeur calculée/suggérée/saisie pour le dictionnaire final
            dossier_n = st.session_state.get(skey("form_add", "dossier_n_sub"), st.session_state.get(skey("form_add", "dossier_n_main"), dossier_n_value))

            client_name = st.text_input("Nom du Client", value=client_name_value, key=skey("form_add", "nom"))
            
            date_dossier = col_date.date_input("Date d'Ouverture du Dossier", value=date.today(), key=skey("form_add", "date"), format="DD/MM/YYYY")
            # FIN LOGIQUE SOUS-DOSSIER

            st.markdown("### 2. Classification Visa")
            cat_n1, sub_cat_final = _render_visa_classification_form(
                key_suffix="add",
                visa_structure=visa_structure,
            )

            st.markdown("### 3. Finance")
            col_montant, col_paye = st.columns(2)
            montant_facture = col_montant.number_input("Total Facturé (Montant)", min_value=0.0, step=100.0, key=skey("form_add", "montant"))
            paye_initial = col_paye.number_input("Paiement Initial Reçu (Payé)", min_value=0.0, step=100.0, key=skey("form_add", "payé"))

            st.markdown("### 4. État du Dossier (Dates Clés)")
            # Nouveaux champs pour le suivi de l'état du dossier
            col_sent, col_approved, col_refused, col_cancelled = st.columns(4)
            date_envoye = col_sent.date_input("Dossier Envoyé", value=None, key=skey("form_add", "dossier_envoyé"), format="DD/MM/YYYY")
            date_approuve = col_approved.date_input("Dossier Approuvé", value=None, key=skey("form_add", "dossier_approuvé"), format="DD/MM/YYYY")
            date_refuse = col_refused.date_input("Dossier Refusé", value=None, key=skey("form_add", "dossier_refusé"), format="DD/MM/YYYY")
            date_annule = col_cancelled.date_input("Dossier Annulé", value=None, key=skey("form_add", "dossier_annulé"), format="DD/MM/YYYY")


            st.markdown("### 5. Notes")
            commentaires = st.text_area("Commentaires", key=skey("form_add", "commentaires"))

            submitted = st.form_submit_button("➕ Ajouter le Dossier")

            if submitted:
                new_data = {
                    'dossier_n': dossier_n,
                    'nom': client_name,
                    'date': date_dossier,
                    'categorie': cat_n1,
                    'sous_categorie': sub_cat_final,
                    'montant': montant_facture,
                    'payé': paye_initial,
                    # Ajout des nouvelles dates d'état
                    'dossier_envoyé': date_envoye,
                    'dossier_approuvé': date_approuve,
                    'dossier_refusé': date_refuse,
                    'dossier_annulé': date_annule,
                    'commentaires': commentaires,
                }
                st.session_state[skey("df_clients")] = _update_client_data(df_clients, new_data, "ADD")
                st.rerun()

    # MODIFY
    with tab_modify:
        st.subheader("Modifier un dossier client existant")
        if df_clients.empty:
            st.info("Veuillez charger ou ajouter des dossiers clients pour pouvoir les modifier.")
        else:
            selected_key_mod = st.selectbox(
                "Sélectionner le Dossier à Modifier",
                ["Sélectionnez un dossier"] + list(client_options.keys()),
                key=skey("modify", "select_client")
            )

            if selected_key_mod != "Sélectionnez un dossier":
                selected_dossier_n = client_options.get(selected_key_mod)
                current_data = df_clients[df_clients['dossier_n'].astype(str) == selected_dossier_n].iloc[0].to_dict()

                initial_cat = str(current_data.get('categorie', '')).strip()
                initial_sub_cat = str(current_data.get('sous_categorie', '')).strip()

                n2_type, n3_key, n4_option = _resolve_visa_levels(initial_cat, initial_sub_cat, visa_structure)

                with st.form("modify_client_form"):
                    st.markdown("---")
                    col_id, col_name, col_date = st.columns(3)

                    col_id.text_input("Numéro de Dossier", value=selected_dossier_n, disabled=True)
                    client_name = col_name.text_input("Nom du Client", value=current_data.get('nom', ''), key=skey("form_mod", "nom"))

                    # Fonction utilitaire locale pour gérer la conversion des dates
                    def _get_current_date(data, col_name):
                        val = data.get(col_name)
                        # Vérifie si la valeur est NaN, NaT, ou vide
                        if pd.isna(val) or val is None:
                            return None
                        try:
                            # Tente de convertir en date.date()
                            return pd.to_datetime(val).date()
                        except:
                            return None

                    date_dossier = col_date.date_input("Date d'Ouverture du Dossier", value=_get_current_date(current_data, 'date'), key=skey("form_mod", "date"), format="DD/MM/YYYY")

                    st.markdown("### Classification Visa")
                    cat_n1, sub_cat_final = _render_visa_classification_form(
                        key_suffix="mod",
                        visa_structure=visa_structure,
                        initial_category=initial_cat,
                        initial_type=n2_type,
                        initial_level4_option=n4_option
                    )

                    st.markdown("### Finance")
                    col_montant, col_paye = st.columns(2)
                    montant_facture = col_montant.number_input(
                        "Total Facturé (Montant)",
                        min_value=0.0,
                        step=100.0,
                        value=current_data.get('montant', 0.0),
                        key=skey("form_mod", "montant")
                    )
                    paye_initial = col_paye.number_input(
                        "Total Payé (Payé)",
                        min_value=0.0,
                        step=100.0,
                        value=current_data.get('payé', 0.0),
                        key=skey("form_mod", "payé")
                    )

                    st.markdown("### État du Dossier (Dates Clés)")
                    # Affichage des dates d'état existantes pour modification
                    col_sent, col_approved, col_refused, col_cancelled = st.columns(4)
                    date_envoye = col_sent.date_input("Dossier Envoyé", value=_get_current_date(current_data, 'dossier_envoyé'), key=skey("form_mod", "dossier_envoyé"), format="DD/MM/YYYY")
                    date_approuve = col_approved.date_input("Dossier Approuvé", value=_get_current_date(current_data, 'dossier_approuvé'), key=skey("form_mod", "dossier_approuvé"), format="DD/MM/YYYY")
                    date_refuse = col_refused.date_input("Dossier Refusé", value=_get_current_date(current_data, 'dossier_refusé'), key=skey("form_mod", "dossier_refusé"), format="DD/MM/YYYY")
                    date_annule = col_cancelled.date_input("Dossier Annulé", value=_get_current_date(current_data, 'dossier_annulé'), key=skey("form_mod", "dossier_annulé"), format="DD/MM/YYYY")


                    st.markdown("### Notes")
                    commentaires = st.text_area("Commentaires", value=current_data.get('commentaires', ''), key=skey("form_mod", "commentaires"))

                    submitted_mod = st.form_submit_button("✍️ Modifier le Dossier")

                    if submitted_mod:
                        new_data = {
                            'dossier_n': selected_dossier_n,
                            'nom': client_name,
                            'date': date_dossier,
                            'categorie': cat_n1,
                            'sous_categorie': sub_cat_final,
                            'montant': montant_facture,
                            'payé': paye_initial,
                            'dossier_envoyé': date_envoye,
                            'dossier_approuvé': date_approuve,
                            'dossier_refusé': date_refuse,
                            'dossier_annulé': date_annule,
                            'commentaires': commentaires,
                        }
                        st.session_state[skey("df_clients")] = _update_client_data(df_clients, new_data, "MODIFY")
                        st.rerun()

    # DELETE
    with tab_delete:
        st.subheader("Supprimer un dossier client")
        if df_clients.empty:
            st.info("Aucun dossier à supprimer.")
        else:
            selected_key_del = st.selectbox(
                "Sélectionner le Dossier à Supprimer",
                ["Sélectionnez un dossier"] + list(client_options.keys()),
                key=skey("delete", "select_client")
            )

            if selected_key_del != "Sélectionnez un dossier":
                selected_dossier_n_del = client_options.get(selected_key_del)
                st.error(f"⚠️ Êtes-vous sûr de vouloir supprimer définitivement le Dossier N° {selected_dossier_n_del} ?")

                if st.button(f"🗑️ Confirmer la Suppression de {selected_dossier_n_del}"):
                    new_data = {'dossier_n': selected_dossier_n_del}
                    st.session_state[skey("df_clients")] = _update_client_data(df_clients, new_data, "DELETE")
                    st.rerun()


def main():
    st.set_page_config(layout="wide", page_title=APP_TITLE)
    st.title(APP_TITLE)

    upload_section()
    data_processing_flow()

    df_clients = st.session_state.get(skey("df_clients"), pd.DataFrame())
    df_visa = st.session_state.get(skey("df_visa"), pd.DataFrame())

    global VISA_STRUCTURE
    # Reconstruire la structure Visa si le DF est là mais la structure est vide (ex: après un st.rerun)
    if not df_visa.empty and not VISA_STRUCTURE:
        VISA_STRUCTURE = _build_visa_structure(df_visa)
        
    if df_clients.empty and df_visa.empty:
        st.info("Bienvenue ! Veuillez commencer par charger vos fichiers Clients et Visa dans la barre latérale.")
        return

    # Tabs principaux
    tab_home, tab_acct, tab_dossier, tab_visa = st.tabs([
        "🏠 Accueil & Stats", "📈 Comptabilité", "📝 Gestion des Dossiers (CRUD)", "📑 Structure Visa"
    ])

    with tab_home:
        home_tab(df_clients)

    with tab_acct:
        accounting_tab(df_clients)

    with tab_dossier:
        dossier_management_tab(df_clients, VISA_STRUCTURE)

    with tab_visa:
        st.header("📑 Aperçu de la Structure Visa (N1 > N2 > N4)")
        if VISA_STRUCTURE:
            st.json(VISA_STRUCTURE)
        else:
            st.warning("Structure Visa non chargée ou vide. Veuillez vérifier votre fichier Visa.")


if __name__ == '__main__':
    main()
