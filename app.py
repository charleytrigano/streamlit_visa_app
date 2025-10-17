import io
import re
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# =========================
# Constantes et Configuration
# =========================
APP_TITLE = "🛂 Visa Manager - Projet Stable"
SID = "vmgr_v3"

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
    
    # ... (Le code de lecture de fichier est inchangé et reste robuste) ...
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
    
    # 1. Nettoyer les noms de colonnes : supprime les caractères spéciaux, minuscules.
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    
    # --- 2. Conversion des Nombres (Vectorisée et Renforcée) ---
    # Basé sur les colonnes trouvées dans le fichier "Clients.csv"
    money_cols = ['honoraires', 'payé', 'solde', 'acompte_1', 'acompte_2', 'montant', 'autres_frais_us_']
    
    for col in money_cols:
        if col in df.columns:
            # Remplacement ',' par '.' et suppression des non-numériques pour robustesse
            df[col] = df[col].astype(str).str.strip().str.replace(',', '.', regex=False)
            df[col] = df[col].str.replace(r'[^\d.]', '', regex=True)
            # Conversion en float, les erreurs sont à NaN, puis NaN à 0.0
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float) 

    # --- 3. Conversion des Dates (Vectorisée) ---
    date_cols = ['date', 'dossier_envoyé', 'dossier_approuvé', 'dossier_refusé', 'dossier_annulé']
    
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # --- 4. Colonne dérivée pour les statistiques actives ---
    if 'date' in df.columns:
         df['jours_ecoules'] = (pd.to_datetime('today') - df['date']).dt.days

    st.success("Nettoyage et conversion des données Clients terminés (Vectorisé et Robuste).")
    return df

def _clean_visa_data(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et standardise les types de données du DataFrame Visa."""
    # Le nettoyage pour le fichier Visa est minimal, car il est principalement une table de référence
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    return df

@st.cache_data
def _summarize_data(df: pd.DataFrame) -> Dict[str, Any]:
    """Calcule des indicateurs clés à partir du DataFrame Clients (robuste aux colonnes manquantes)."""
    
    if df.empty:
        return {"total_clients": 0, "total_honoraires": 0.0, "solde_du": 0.0, "clients_actifs": 0}

    # Calculs financiers robustes (les colonnes sont garanties float/0.0 après nettoyage)
    total_honoraires = df['honoraires'].sum() if 'honoraires' in df.columns else 0.0
    total_payé = df['payé'].sum() if 'payé' in df.columns else 0.0
    solde_du = df['solde'].sum() if 'solde' in df.columns else 0.0
    
    # Logique robuste pour les clients actifs (si aucune colonne d'état n'est présente, tous sont actifs)
    end_cols = ['dossier_approuvé', 'dossier_annulé', 'dossier_refusé']
    active_mask = pd.Series([True] * len(df), index=df.index)
    
    for col in end_cols:
        if col in df.columns:
            # Un dossier n'est PLUS actif s'il a une date dans l'une de ces colonnes
            active_mask &= df[col].isna()

    clients_actifs = active_mask.sum()
    
    summary = {
        "total_clients": len(df),
        "total_honoraires": total_honoraires,
        "total_payé": total_payé,
        "solde_du": solde_du,
        "clients_actifs": clients_actifs,
    }
    return summary


# =========================
# Fonctions de l'Interface Utilisateur (UI)
# =========================

def upload_section():
    """Section de chargement des fichiers."""
    st.sidebar.header("📁 Chargement des Fichiers")
    
    # ... (Logique de chargement inchangée)
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


def home_tab(df_clients: pd.DataFrame):
    """Contenu de l'onglet Accueil/Statistiques."""
    st.header("📊 Statistiques Clés")
    
    if df_clients.empty:
        st.info("Veuillez charger un fichier de Clients dans la barre latérale pour afficher les statistiques.")
        return
        
    summary = _summarize_data(df_clients)

    col1, col2, col3, col4 = st.columns(4)

    # Affichage des métriques
    col1.metric("Clients Totaux", f"{summary['total_clients']:,}".replace(",", " "))
    col2.metric("Honoraires Facturés", f"${summary['total_honoraires']:,.2f}".replace(",", " "))
    col3.metric("Solde Total Dû", f"${summary['solde_du']:,.2f}".replace(",", " "))
    col4.metric("Dossiers Actifs (Non Clôturés)", f"{summary['clients_actifs']:,}".replace(",", " "))
    
    st.divider()
    
    st.subheader("Analyse Rapide")
    if 'categorie' in df_clients.columns:
        counts = df_clients['categorie'].value_counts().head(10)
        st.bar_chart(counts, use_container_width=True)
    else:
        st.warning("Colonne 'categorie' introuvable pour l'analyse. Vérifiez l'index d'en-tête.")

# --- NOUVEAU: Logique de Classification de Visa ---
def visa_classification_logic():
    st.header("🛂 Saisie et Classification de Visa")
    st.markdown("---")

    # 1. Sélection de la Grande Catégorie (Affaires/Tourisme, Etudiants, etc.)
    col_main, col_type = st.columns(2)
    
    with col_main:
        main_category = st.selectbox(
            "1. Catégorie de Visa (Grand Groupe)",
            ["Sélectionnez un groupe"] + list(VISA_STRUCTURE.keys()),
            key=skey("cat", "main"),
            help="Les noms ne sont pas enregistrés, juste pour le regroupement.",
        )

    # 2. Sélection du Type de Visa (B-1, F-1, E-2, etc. - les points ●)
    selected_options = VISA_STRUCTURE.get(main_category, {})
    selected_type = None

    if selected_options:
        # Si la structure est profonde (Treaty, Residence Permanente), il y a des sous-groupes
        # Nous prenons le premier niveau de clés pour la première sélection (Selectbox)
        visa_types = list(selected_options.keys())
        
        with col_type:
            selected_type = st.selectbox(
                f"2. Type de Visa ({main_category})",
                ["Sélectionnez un type"] + visa_types,
                key=skey("cat", "type"),
            )
        
        # 3. Affichage des Sous-Catégories (Radio Buttons)
        if selected_type and selected_type != "Sélectionnez un type":
            current_options = selected_options.get(selected_type)

            if isinstance(current_options, list):
                # Cas simple : liste d'options (ex: B-1 -> COS/EOS)
                st.subheader(f"3. Option pour {selected_type} (Rond à sélectionner)")
                
                # Le rond de sélection (Radio) pour les sous-catégories
                final_selection = st.radio(
                    "Choisissez l'option finale",
                    current_options,
                    key=skey("cat", "sub1"),
                    horizontal=True
                )
                st.success(f"Dossier sélectionné : {main_category} > {selected_type} > {final_selection}")
                
            elif isinstance(current_options, dict):
                # Cas complexe/imbriqué : Dictionnaire (ex: E-2 -> E-2 Inv.)
                st.subheader(f"3. Sous-catégorie pour {selected_type}")
                
                # Niveau 3 : Selectbox pour les clés du dictionnaire imbriqué
                nested_key = st.selectbox(
                    f"Sous-catégorie de {selected_type}",
                    list(current_options.keys()),
                    key=skey("cat", "nested_key"),
                )
                
                # Niveau 4 : Radio Buttons pour les options finales
                nested_options = current_options.get(nested_key)
                if nested_options and isinstance(nested_options, list):
                    st.subheader(f"4. Option finale pour {nested_key} (Rond à sélectionner)")
                    final_selection = st.radio(
                        "Choisissez l'option finale",
                        nested_options,
                        key=skey("cat", "sub2"),
                        horizontal=True
                    )
                    st.success(f"Dossier sélectionné : {main_category} > {selected_type} > {nested_key} > {final_selection}")
    
    # --- Affichage des options simples (Derivatives, etc.) ---
    st.markdown("---")
    st.subheader("Services Simples (Affichage sur une ligne)")
    
    simple_cols = st.columns(6)
    
    for i, (key, sub_options) in enumerate(SIMPLE_SERVICE_OPTIONS.items()):
        # Utiliser st.expander pour les options avec des sous-choix comme Naturalization
        if sub_options:
             with simple_cols[i % 6]:
                 if key == "Other":
                     st.text_input("Autre service (détail)", key=skey("simple", key))
                 elif key == "Naturalization":
                     st.radio(
                         key,
                         sub_options,
                         key=skey("simple", key),
                         horizontal=False,
                         help="Sélectionnez le type de Naturalisation"
                     )
                 else:
                    st.expander(key).radio(
                        f"Option pour {key}",
                        sub_options,
                        key=skey("simple", key),
                        horizontal=True
                    )
        else:
             with simple_cols[i % 6]:
                # Utiliser un simple checkbox pour les options sans sous-choix
                st.checkbox(key, key=skey("simple", key), help="Cocher pour sélectionner")


# --- Le reste des onglets est inchangé ---
def settings_tab():
    """Contenu de l'onglet Configuration."""
    # ... (Le code de configuration d'en-tête est inchangé)
    st.header("⚙️ Configuration du Chargement")
    
    st.markdown("""
        Veuillez spécifier l'index de la ligne contenant les noms de colonnes réels.
        * **0** (par défaut) : première ligne.
        * **1** : deuxième ligne, etc.
    """)
    
    # Paramètre d'en-tête pour Clients
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

    # Paramètre d'en-tête pour Visa
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

    # Export Clients
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

    # Export Visa
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
    
    # 1. Section de chargement des fichiers (Barre latérale)
    upload_section()
    
    # 2. Flux de traitement des données (Chargement et nettoyage)
    data_processing_flow()
    
    # Récupérer les DataFrames nettoyés
    df_clients = st.session_state.get(skey("df_clients"), pd.DataFrame())
    df_visa = st.session_state.get(skey("df_visa"), pd.DataFrame())

    # 3. Affichage des onglets
    tab_home, tab_config, tab_visa_entry, tab_clients_view, tab_visa_view, tab_export = st.tabs([
        "🏠 Accueil & Stats", 
        "⚙️ Configuration",
        "📝 Saisie Dossier", # Nouvel onglet pour tester la classification
        "📄 Clients - Aperçu", 
        "📄 Visa - Aperçu", 
        "💾 Export",
    ])

    with tab_home:
        home_tab(df_clients)

    with tab_config:
        settings_tab()
        
    with tab_visa_entry:
        visa_classification_logic()

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
