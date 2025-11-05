import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os
import datetime as dt
from datetime import datetime
import base64
import plotly.express as px
from dateutil.relativedelta import relativedelta

APP_TITLE = "Visa Manager"

st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

# --- helper to hide Escrow in selected UI tables (we keep the column in the data model) ---
def hide_escrow(df: pd.DataFrame) -> pd.DataFrame:
    try:
        if isinstance(df, pd.DataFrame):
            return df.drop(columns=["Escrow"], errors="ignore")
    except Exception:
        pass
    return df

def st_display_df(df, **kwargs):
    # Use this helper to show dataframes with Escrow column removed (keeps internal data intact)
    try:
        df2 = hide_escrow(df)
    except Exception:
        df2 = df
    return st.dataframe(df2, **kwargs)
# --- end helpers ---

# ===============================
# Initialisation et chargement
# ===============================

if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame()

def _get_df_live_safe():
    try:
        return st.session_state.df.copy()
    except Exception:
        return pd.DataFrame()

def _fmt_money(x):
    try:
        return f"{float(x):,.0f}".replace(",", " ")
    except Exception:
        return x

def _to_num(x):
    try:
        return float(str(x).replace(",", ".").replace(" ", ""))
    except:
        return 0

# -------------------------------
# Fonctions utilitaires
# -------------------------------

def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    return df

def save_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def recalc_payments_and_solde(df):
    if df is None or df.empty:
        return df
    if "Montant total" in df.columns:
        df["Montant total"] = df["Montant total"].apply(_to_num)
    acompte_cols = [c for c in df.columns if "Acompte" in c and not "Date" in c]
    df["Total acomptes"] = df[acompte_cols].applymap(_to_num).sum(axis=1)
    if "Montant total" in df.columns:
        df["Solde"] = df["Montant total"] - df["Total acomptes"]
    return df

# -------------------------------
# Layout principal avec onglets
# -------------------------------

tabs = st.tabs([
    "📂 Importation", 
    "📊 Dashboard", 
    "📈 Analyses", 
    "➕ Ajouter", 
    "🗂️ Gestion", 
    "⚙️ Paramètres",
    "🧾 Escrow"
])
escrow_tab = tabs[-1]


# ======================================================
# 📂 Onglet Importation
# ======================================================

with tabs[0]:
    st.header("📂 Importer un fichier Excel")
    uploaded_file = st.file_uploader("Choisir un fichier Excel", type=["xls", "xlsx"])
    
    if uploaded_file is not None:
        try:
            df = load_data(uploaded_file)
            st.session_state.df = df
            st.success(f"Fichier importé avec succès ({len(df)} lignes).")
            st_display_df(df.head())
        except Exception as e:
            st.error(f"Erreur lors de l'importation : {e}")
    else:
        st.info("Veuillez importer un fichier Excel pour commencer.")
    

# ======================================================
# 📊 Onglet Dashboard
# ======================================================

with tabs[1]:
    st.header("📊 Tableau de bord")
    df = _get_df_live_safe()
    if df is None or df.empty:
        st.info("Aucune donnée disponible. Importez un fichier d'abord.")
    else:
        df = recalc_payments_and_solde(df)

        col1, col2, col3, col4 = st.columns(4)
        total_dossiers = len(df)
        total_encaissé = df["Total acomptes"].sum()
        total_montant = df["Montant total"].sum()
        total_solde = df["Solde"].sum()

        col1.metric("Nombre de dossiers", total_dossiers)
        col2.metric("Montant total", f"{_fmt_money(total_montant)} €")
        col3.metric("Total encaissé", f"{_fmt_money(total_encaissé)} €")
        col4.metric("Solde restant", f"{_fmt_money(total_solde)} €")

        # Graphique de répartition
        if "Nationalité" in df.columns:
            fig = px.pie(df, names="Nationalité", title="Répartition par nationalité")
            st.plotly_chart(fig, use_container_width=True)

        if "Date" in df.columns:
            df_date = df.copy()
            df_date["Mois"] = df_date["Date"].dt.to_period("M").astype(str)
            monthly = df_date.groupby("Mois")["Montant total"].sum().reset_index()
            fig2 = px.bar(monthly, x="Mois", y="Montant total", title="Montant total par mois")
            st.plotly_chart(fig2, use_container_width=True)

        # Tableau
        st.subheader("Aperçu des données")
        st_display_df(df, use_container_width=True, height=400)


# ======================================================
# 📈 Onglet Analyses
# ======================================================

with tabs[2]:
    st.header("📈 Analyses statistiques")
    df = _get_df_live_safe()
    if df is None or df.empty:
        st.info("Aucune donnée disponible.")
    else:
        df = recalc_payments_and_solde(df)
        numeric_cols = ["Montant total", "Total acomptes", "Solde"]
        stats = df[numeric_cols].describe().T
        st.subheader("Statistiques globales")
        st_display_df(stats)

        # Graphique de distribution des soldes
        if "Solde" in df.columns:
            fig = px.histogram(df, x="Solde", nbins=20, title="Distribution des soldes")
            st.plotly_chart(fig, use_container_width=True)

        # Moyennes par pays
        if "Nationalité" in df.columns:
            pays = df.groupby("Nationalité")[["Montant total", "Total acomptes", "Solde"]].mean().reset_index()
            fig3 = px.bar(
                pays,
                x="Nationalité",
                y="Montant total",
                title="Montant total moyen par nationalité"
            )
            st.plotly_chart(fig3, use_container_width=True)

        # Tableau de détails
        st.subheader("Données détaillées")
        st_display_df(df, use_container_width=True)


# ======================================================
# ➕ Onglet Ajouter
# ======================================================

with tabs[3]:
    st.header("➕ Ajouter un dossier manuellement")

    df = _get_df_live_safe()
    if df is None:
        df = pd.DataFrame()

    with st.form("ajout_dossier"):
        col1, col2, col3 = st.columns(3)
        dossier_num = col1.text_input("Numéro de dossier")
        nom_client = col2.text_input("Nom du client")
        nationalite = col3.text_input("Nationalité")

        col4, col5, col6 = st.columns(3)
        date_dossier = col4.date_input("Date du dossier", dt.date.today())
        montant_total = col5.text_input("Montant total (€)")
        acompte1 = col6.text_input("Acompte 1 (€)")

        col7, col8 = st.columns(2)
        date_acompte1 = col7.date_input("Date acompte 1", dt.date.today())
        dossiers_envoye = col8.checkbox("Dossier envoyé ?")

        submit = st.form_submit_button("Ajouter")

    if submit:
        try:
            new_row = {
                "Dossier N": dossier_num,
                "Nom": nom_client,
                "Nationalité": nationalite,
                "Date": pd.to_datetime(date_dossier),
                "Montant total": montant_total,
                "Acompte 1": acompte1,
                "Date Acompte 1": pd.to_datetime(date_acompte1),
                "Dossiers envoyé": 1 if dossiers_envoye else 0
            }

            st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame([new_row])], ignore_index=True)
            st.success("✅ Dossier ajouté avec succès.")
        except Exception as e:
            st.error(f"Erreur lors de l'ajout : {e}")

    if not st.session_state.df.empty:
        st.subheader("Aperçu du tableau mis à jour")
        st_display_df(st.session_state.df.tail(), use_container_width=True)


# ======================================================
# 🗂️ Onglet Gestion
# ======================================================

with tabs[4]:
    st.header("🗂️ Gestion des dossiers")

    df = _get_df_live_safe()
    if df is None or df.empty:
        st.info("Aucune donnée à gérer.")
    else:
        df = recalc_payments_and_solde(df)

        st.subheader("Recherche / Filtrage")
        search_nom = st.text_input("Rechercher par nom ou dossier N°")

        df_filtered = df.copy()
        if search_nom:
            search_nom_lower = search_nom.lower()
            df_filtered = df[df["Nom"].str.lower().str.contains(search_nom_lower, na=False) |
                             df["Dossier N"].astype(str).str.contains(search_nom_lower, na=False)]

        st_display_df(df_filtered, use_container_width=True, height=420)

        st.markdown("---")
        st.subheader("📤 Exportation Excel")

        export_df = st.button("Exporter les données actuelles")
        if export_df:
            try:
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    df_filtered.to_excel(writer, index=False, sheet_name="Export")
                st.download_button(
                    "Télécharger le fichier Excel",
                    data=buf.getvalue(),
                    file_name="export_dossiers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erreur export Excel : {e}")


# ======================================================
# ⚙️ Onglet Paramètres
# ======================================================

with tabs[5]:
    st.header("⚙️ Paramètres de l’application")

    df = _get_df_live_safe()
    if df is None or df.empty:
        st.info("Aucune donnée chargée.")
    else:
        st.subheader("Structure du tableau actuel")
        st.write(f"{len(df.columns)} colonnes, {len(df)} lignes")
        st_display_df(pd.DataFrame({"Colonnes": df.columns}), use_container_width=True)

        st.markdown("---")
        st.subheader("Télécharger la base complète")

        if st.button("Exporter la base complète"):
            try:
                excel_bytes = save_to_excel(df)
                st.download_button(
                    "Télécharger le fichier Excel complet",
                    data=excel_bytes,
                    file_name="base_complete.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erreur lors de l'export : {e}")

        st.markdown("---")
        st.subheader("Réinitialisation de la session")
        if st.button("Réinitialiser toutes les données"):
            st.session_state.df = pd.DataFrame()
            st.warning("Toutes les données ont été effacées de la session en cours.")


# ======================================================
# 🧾 Onglet Escrow (nouveau)
# ======================================================

with escrow_tab:
    st.header("🧾 Dossiers en Escrow")

    df_live_esc = recalc_payments_and_solde(_get_df_live_safe())

    if df_live_esc is None or df_live_esc.empty:
        st.info("Aucune donnée en mémoire.")
    else:
        try:
            # Filtrer uniquement les dossiers Escrow
            df_esc = df_live_esc[df_live_esc.get("Escrow", 0) == 1].copy()

            display_cols = ["Dossier N", "Nom", "Date", "Acompte 1", "Date Acompte 1", "Dossiers envoyé"]
            existing = [c for c in display_cols if c in df_esc.columns]

            if not existing:
                st.info("Aucun champ Escrow pertinent trouvé dans les données.")
            else:
                df_show = df_esc[existing].reset_index(drop=True)

                # Formatage
                for c in ["Date", "Date Acompte 1"]:
                    if c in df_show.columns:
                        try:
                            df_show[c] = pd.to_datetime(df_show[c], errors="coerce").dt.date
                        except Exception:
                            pass

                for m in ["Acompte 1"]:
                    if m in df_show.columns:
                        try:
                            df_show[m] = df_show[m].apply(lambda x: _fmt_money(_to_num(x)))
                        except Exception:
                            pass

                st_display_df(df_show, use_container_width=True, height=480)

                # Export Excel
                buf = BytesIO()
                try:
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        df_show.to_excel(writer, index=False, sheet_name="Escrow")
                    st.download_button(
                        "Exporter Escrow en Excel",
                        data=buf.getvalue(),
                        file_name="escrow_export.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Erreur export Excel: {e}")

        except Exception as e:
            st.error(f"Erreur affichage Escrow: {e}")


# ======================================================
# 🔧 Fonctions auxiliaires diverses (calculs et formats)
# ======================================================

def calc_solde(row):
    """Calcule le solde pour une ligne donnée."""
    try:
        total = _to_num(row.get("Montant total", 0))
        acomptes = 0
        for col in row.index:
            if "Acompte" in col and "Date" not in col:
                acomptes += _to_num(row[col])
        return total - acomptes
    except Exception:
        return None


def calc_total_acomptes(row):
    """Calcule la somme de tous les acomptes pour une ligne donnée."""
    try:
        acomptes = 0
        for col in row.index:
            if "Acompte" in col and "Date" not in col:
                acomptes += _to_num(row[col])
        return acomptes
    except Exception:
        return 0


def format_date(val):
    """Convertit une date ou un texte en objet datetime, ou None si invalide."""
    if pd.isna(val):
        return None
    try:
        return pd.to_datetime(val, errors="coerce")
    except Exception:
        return None


def refresh_df_stats(df):
    """Recalcule les champs dérivés et statistiques principales."""
    if df is None or df.empty:
        return df

    df = recalc_payments_and_solde(df)
    df["Total acomptes recalculé"] = df.apply(calc_total_acomptes, axis=1)
    df["Solde recalculé"] = df.apply(calc_solde, axis=1)
    return df


def summarize_nationality(df):
    """Renvoie un résumé agrégé par nationalité."""
    if df is None or df.empty:
        return pd.DataFrame()
    if "Nationalité" not in df.columns:
        return pd.DataFrame()
    grp = (
        df.groupby("Nationalité")[["Montant total", "Total acomptes", "Solde"]]
        .sum()
        .reset_index()
        .sort_values("Montant total", ascending=False)
    )
    return grp


# ======================================================
# 💾 Fonctions d'export et de sauvegarde
# ======================================================

def export_dataframe_to_excel(df, filename="export.xlsx"):
    """Crée un fichier Excel à partir d’un DataFrame."""
    try:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"Erreur export Excel : {e}")
        return None


def export_filtered_data(df, filter_condition, filename="filtre_export.xlsx"):
    """Exporte les données filtrées selon une condition donnée."""
    try:
        df_filtered = df.query(filter_condition)
        data = export_dataframe_to_excel(df_filtered, filename)
        if data:
            st.download_button(
                label=f"Télécharger le fichier {filename}",
                data=data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"Erreur lors de l'export filtré : {e}")


# ======================================================
# 🧮 Fonctions analytiques
# ======================================================

def compute_monthly_totals(df):
    """Calcule les totaux mensuels des montants."""
    if df is None or df.empty or "Date" not in df.columns:
        return pd.DataFrame()
    df_temp = df.copy()
    df_temp["Mois"] = df_temp["Date"].dt.to_period("M").astype(str)
    monthly = (
        df_temp.groupby("Mois")[["Montant total", "Total acomptes", "Solde"]]
        .sum()
        .reset_index()
    )
    return monthly


def compute_top_clients(df, n=10):
    """Renvoie les top N clients par montant total."""
    if df is None or df.empty or "Nom" not in df.columns:
        return pd.DataFrame()
    ranked = df.groupby("Nom")[["Montant total", "Total acomptes", "Solde"]].sum().reset_index()
    ranked = ranked.sort_values("Montant total", ascending=False).head(n)
    return ranked


# ======================================================
# 📊 Fonctions graphiques
# ======================================================

def plot_monthly_revenue(df):
    """Affiche un graphique des montants mensuels."""
    if df is None or df.empty:
        st.warning("Aucune donnée pour tracer le graphique mensuel.")
        return
    monthly = compute_monthly_totals(df)
    if monthly.empty:
        st.warning("Les données ne contiennent pas de dates valides.")
        return
    fig = px.bar(
        monthly,
        x="Mois",
        y="Montant total",
        title="Montant total par mois",
        labels={"Montant total": "Montant (€)"},
    )
    st.plotly_chart(fig, use_container_width=True)


def plot_top_clients(df, n=10):
    """Affiche le graphique des top clients."""
    ranked = compute_top_clients(df, n)
    if ranked.empty:
        st.warning("Aucun client trouvé pour le graphique Top clients.")
        return
    fig = px.bar(
        ranked,
        x="Nom",
        y="Montant total",
        title=f"Top {n} clients par montant total",
        labels={"Montant total": "Montant (€)"},
    )
    st.plotly_chart(fig, use_container_width=True)


# ======================================================
# 🧾 Résumé global des données
# ======================================================

def summary_section(df):
    """Affiche un résumé synthétique des données."""
    if df is None or df.empty:
        st.info("Aucune donnée disponible pour le résumé.")
        return
    st.markdown("### Résumé global")
    total_dossiers = len(df)
    total_montant = df["Montant total"].sum() if "Montant total" in df.columns else 0
    total_acomptes = df["Total acomptes"].sum() if "Total acomptes" in df.columns else 0
    total_solde = df["Solde"].sum() if "Solde" in df.columns else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total dossiers", total_dossiers)
    col2.metric("Montant total (€)", f"{_fmt_money(total_montant)} €")
    col3.metric("Total encaissé (€)", f"{_fmt_money(total_acomptes)} €")
    col4.metric("Solde restant (€)", f"{_fmt_money(total_solde)} €")

    st.markdown("---")
    plot_monthly_revenue(df)
    st.markdown("---")
    plot_top_clients(df)


# ======================================================
# 🧰 Fonctions diverses et outils internes
# ======================================================

def clean_dataframe(df):
    """Nettoie le DataFrame en supprimant les colonnes vides et doublons."""
    if df is None or df.empty:
        return df
    df = df.drop_duplicates().copy()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    return df


def add_missing_columns(df, required_columns):
    """Ajoute les colonnes manquantes avec des valeurs par défaut."""
    if df is None:
        df = pd.DataFrame()
    for col in required_columns:
        if col not in df.columns:
            df[col] = None
    return df


def ensure_column_types(df):
    """Convertit les colonnes au bon type selon leur contenu."""
    if df is None or df.empty:
        return df
    for col in df.columns:
        if "Date" in col:
            df[col] = pd.to_datetime(df[col], errors="coerce")
        elif any(x in col.lower() for x in ["montant", "acompte", "solde"]):
            df[col] = df[col].apply(_to_num)
    return df


def prepare_dataframe(df):
    """Pipeline de préparation des données avant utilisation."""
    if df is None or df.empty:
        return df
    df = clean_dataframe(df)
    df = ensure_column_types(df)
    df = recalc_payments_and_solde(df)
    return df


# ======================================================
# 🚀 Lancement principal
# ======================================================

def main():
    """Point d’entrée de l’application."""
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    df = _get_df_live_safe()
    if df is None or df.empty:
        st.info("Aucune donnée chargée. Importez un fichier pour commencer.")
    else:
        summary_section(df)

    st.markdown("---")
    st.caption("Visa Manager © 2025 – Application de suivi et gestion des dossiers de visa.")


# ======================================================
# 🏁 Point d'entrée du script
# ======================================================

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"Erreur critique : {e}")
