import io
import re
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# =========================
# Constantes et Configuration
# =========================
APP_TITLE = "🛂 Visa Manager - Amélioré"
SID = "vmgr_v2"

# =========================
# Fonctions utilitaires
# =========================

def skey(*args) -> str:
    """Génère une clé unique pour st.session_state."""
    return f"{SID}_{'_'.join(map(str, args))}"

# Utilitaire de chargement de fichier CSV/Excel avec gestion de l'en-tête.
@st.cache_data(show_spinner="Lecture du fichier...")
def _read_data_file(file_content: BytesIO, file_name: str, header_row: int = 0) -> pd.DataFrame:
    """Lit les données d'un fichier téléchargé (CSV ou Excel)."""
    
    # 1. Déterminer le type de fichier
    if file_name.endswith(('.xls', '.xlsx')):
        try:
            # Pour Excel, utiliser la première feuille par défaut
            df = pd.read_excel(
                file_content, 
                header=header_row, 
                engine='openpyxl',
                # Tenter de lire tous les types comme des chaînes pour éviter des erreurs initiales
                dtype=str, 
            )
        except Exception as e:
            st.error(f"Erreur de lecture Excel : {e}")
            return pd.DataFrame()
    else: # Supposer CSV si ce n'est pas Excel
        try:
            # Tenter plusieurs encodages courants
            df = pd.read_csv(
                file_content, 
                header=header_row, 
                sep=None, # Détection automatique du séparateur
                engine='python', # Nécessaire pour sep=None
                encoding='utf-8',
                on_bad_lines='skip',
                # Tenter de lire tous les types comme des chaînes
                dtype=str, 
            )
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(
                    file_content, 
                    header=header_row, 
                    sep=None, 
                    engine='python', 
                    encoding='latin1',
                    on_bad_lines='skip',
                    dtype=str,
                )
            except Exception as e:
                st.error(f"Erreur de lecture CSV : {e}")
                return pd.DataFrame()
        except Exception as e:
            st.error(f"Erreur de lecture CSV : {e}")
            return pd.DataFrame()
    
    # Nettoyage des colonnes : supprimer les colonnes entièrement vides
    df = df.dropna(axis=1, how='all')
    
    # Nettoyage des noms de colonnes : supprimer les espaces de début/fin
    df.columns = df.columns.str.strip().fillna('')
    
    # Supprimer les lignes entièrement vides
    df = df.dropna(axis=0, how='all')
    
    return df

def _clean_clients_data(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et standardise les types de données du DataFrame Clients."""
    
    # Nettoyer les noms de colonnes pour une manipulation plus facile
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    
    # Colonnes attendues après nettoyage pour vérification
    COLS_CLIENTS_EXPECTED = ['id_client', 'dossier_n', 'nom', 'date', 'categorie', 'sous_categorie', 'visa']
    
    # Vérification des colonnes critiques
    if not all(col in df.columns for col in COLS_CLIENTS_EXPECTED):
        st.warning(
            "Le DataFrame Clients ne contient pas toutes les colonnes attendues après le nettoyage : "
            f"{', '.join(COLS_CLIENTS_EXPECTED)}."
        )
        return df
        
    def _clean_clients_data(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et standardise les types de données du DataFrame Clients."""
    
    # Nettoyer les noms de colonnes pour une manipulation plus facile
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    
    # Colonnes attendues après nettoyage pour vérification
    COLS_CLIENTS_EXPECTED = ['id_client', 'dossier_n', 'nom', 'date', 'categorie', 'sous_categorie', 'visa']
    
    # Vérification des colonnes critiques
    if not all(col in df.columns for col in COLS_CLIENTS_EXPECTED):
        st.warning(
            "Le DataFrame Clients ne contient pas toutes les colonnes attendues après le nettoyage : "
            f"{', '.join(COLS_CLIENTS_EXPECTED)}."
        )
        # On continue quand même avec le nettoyage des types pour les colonnes trouvées
        
    # --- 1. Conversion des Nombres (Vectorisée et Renforcée) ---
    money_cols = ['honoraires', 'payé', 'solde', 'acompte_1', 'acompte_2', 'montant', 'autres_frais_us_']
    
    for col in money_cols:
        # S'assurer que la colonne existe avant de la traiter
        if col in df.columns:
            # Étape 1: Conversion en chaîne et nettoyage des espaces
            df[col] = df[col].astype(str).str.strip()
            
            # Étape 2: Remplacement des virgules par des points (standard décimal)
            df[col] = df[col].str.replace(',', '.', regex=False)
            
            # Étape 3: Suppression des symboles monétaires/caractères non numériques pour sécurisation
            # Conserve seulement les chiffres et le point décimal.
            df[col] = df[col].str.replace(r'[^\d.]', '', regex=True)

            # Étape 4: Conversion en numérique. Les erreurs sont mises à NaN.
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
            # Étape 5: Remplacer les NaN par 0.0 et forcer le type float pour éviter les erreurs sum()
            df[col] = df[col].fillna(0.0).astype(float) # <<< FIX APPLIQUÉ ICI

    # --- 2. Conversion des Dates (Vectorisée) ---
    date_cols = ['date', 'dossier_envoyé', 'dossier_approuvé', 'dossier_refusé', 'dossier_annulé']
# ... (le reste de la fonction _clean_clients_data est inchangé)
    return df

def _clean_visa_data(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et standardise les types de données du DataFrame Visa."""
    
    # Nettoyer les noms de colonnes
    df.columns = df.columns.str.replace(r'[^a-zA-Z0-9_]', '_', regex=True).str.strip('_').str.lower()
    
    # Assurer que les valeurs sont des chaînes, puis nettoyer les espaces
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
        
    st.success("Nettoyage des données Visa terminé.")
    return df

@st.cache_data
def _summarize_data(df: pd.DataFrame) -> Dict[str, Any]:
    """Calcule des indicateurs clés à partir du DataFrame Clients."""
    
    if df.empty:
        return {"total_clients": 0, "total_honoraires": 0.0, "solde_du": 0.0}

    # Utiliser les noms de colonnes nettoyés (minuscules, underscores)
    total_honoraires = df['montant'].sum() if 'montant' in df.columns else 0.0
    total_payé = df['payé'].sum() if 'payé' in df.columns else 0.0
    solde_du = df['solde'].sum() if 'solde' in df.columns else 0.0
    
    summary = {
        "total_clients": len(df),
        "total_honoraires": total_honoraires,
        "total_payé": total_payé,
        "solde_du": solde_du,
        "clients_actifs": len(df[(df['dossier_approuvé'].isna()) & (df['dossier_annulé'].isna()) & (df['dossier_refusé'].isna())]),
        "clients_payés": len(df[df['solde'] <= 0])
    }
    return summary

# =========================
# Fonctions de l'Interface Utilisateur (UI)
# =========================

def upload_section():
    """Section de chargement des fichiers."""
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
        # Stocker le contenu du fichier dans la session pour l'utiliser avec @st.cache_data
        st.session_state[skey("raw_clients_content")] = uploaded_file_clients.read()
        st.session_state[skey("clients_name")] = uploaded_file_clients.name
        
    if uploaded_file_visa:
        st.session_state[skey("raw_visa_content")] = uploaded_file_visa.read()
        st.session_state[skey("visa_name")] = uploaded_file_visa.name

def data_processing_flow():
    """Gère le chargement, le nettoyage et le stockage des DataFrames."""
    
    # Utiliser st.session_state pour les DataFrames (état principal)
    st.session_state.setdefault(skey("df_clients"), pd.DataFrame())
    st.session_state.setdefault(skey("df_visa"), pd.DataFrame())

    # --- 1. Chargement et Nettoyage Clients ---
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

    # --- 2. Chargement et Nettoyage Visa ---
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

    col1.metric("Clients Totaux", f"{summary['total_clients']:,}".replace(",", " "))
    col2.metric("Honoraires Totaux", f"${summary['total_honoraires']:,.2f}".replace(",", " "))
    col3.metric("Solde Total Dû", f"${summary['solde_du']:,.2f}".replace(",", " "))
    col4.metric("Dossiers Actifs", f"{summary['clients_actifs']:,}".replace(",", " "))
    
    st.divider()
    
    st.subheader("Distribution des Catégories")
    # Exemple d'analyse graphique simple
    if 'categorie' in df_clients.columns:
        counts = df_clients['categorie'].value_counts().head(10)
        st.bar_chart(counts, use_container_width=True)
    else:
        st.warning("Colonne 'categorie' introuvable pour l'analyse.")


def settings_tab():
    """Contenu de l'onglet Configuration."""
    st.header("⚙️ Configuration du Chargement")
    
    st.markdown("""
        Étant donné que votre fichier d'origine semble avoir des en-têtes sur plusieurs lignes, 
        vous pouvez spécifier l'index de la ligne contenant les noms de colonnes réels.
        
        * `0` (par défaut) : première ligne (index 0).
        * `1` : deuxième ligne (index 1), etc.
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
        help="L'index de la ligne qui contient les noms de colonnes réels (commence à 0)."
    )
    if new_header_clients != current_header_clients:
         st.session_state[skey("header_clients_row")] = new_header_clients
         st.rerun() # Rechargement pour appliquer le changement

    # Paramètre d'en-tête pour Visa
    st.subheader("Fichier Visa")
    current_header_visa = st.session_state.get(skey("header_visa_row"), 0)
    new_header_visa = st.number_input(
        "Index de la ligne d'en-tête (Visa)",
        min_value=0,
        value=current_header_visa,
        step=1,
        key=skey("input", "header_visa"),
        help="L'index de la ligne qui contient les noms de colonnes réels (commence à 0)."
    )
    if new_header_visa != current_header_visa:
         st.session_state[skey("header_visa_row")] = new_header_visa
         st.rerun() # Rechargement pour appliquer le changement


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
                # Exporter le DataFrame nettoyé, pas le 'df_all' non défini dans le code original
                df_clients.to_excel(w, index=False, sheet_name="Clients_Nettoyes")
            st.download_button(
                "⬇️ Exporter Clients_Nettoyes.xlsx",
                data=buf.getvalue(),
                file_name="Clients_export_nettoye.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=skey("exp", "clients"),
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
                key=skey("exp", "visa"),
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
    tab_home, tab_config, tab_clients_view, tab_visa_view, tab_export = st.tabs([
        "🏠 Accueil & Stats", 
        "⚙️ Configuration",
        "📄 Clients - Aperçu", 
        "📄 Visa - Aperçu", 
        "💾 Export",
    ])

    with tab_home:
        home_tab(df_clients)

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
    # st.session_state sera initialisé ici
    main()
