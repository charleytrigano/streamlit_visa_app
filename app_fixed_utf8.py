# -*- coding: utf-8 -*-
"""
Application Visa Manager - Version UTF-8
Tous les onglets actifs + intégration complète Escrow (badge, alerte, sauvegarde)
"""

import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO
from datetime import date
import escrow_manager as esc

APP_TITLE = "Visa Manager"
st.set_page_config(page_title=APP_TITLE, page_icon="🛂", layout="wide")
st.title(APP_TITLE)

# === Chargement Escrow & badge ===
if "df_dossiers" not in st.session_state or "df_escrow" not in st.session_state:
    df_dossiers, df_escrow = esc.load_data()
    st.session_state.df_dossiers = df_dossiers
    st.session_state.df_escrow = df_escrow
    st.session_state.df = df_dossiers.copy()
else:
    df_dossiers = st.session_state.df_dossiers
    df_escrow = st.session_state.df_escrow

escrow_pending = esc.a_reclamer(st.session_state.df_escrow)
n_escrow = len(escrow_pending)
escrow_label = "🛡️ Escrow 🔴" if n_escrow > 0 else "🛡️ Escrow"

# --- Définition des onglets ---
tabs = st.tabs([
    "📊 Dashboard",
    "📈 Analyses",
    "➕ Ajouter",
    "✏️ Gestion",
    "💳 Compta Client",
    "💾 Export",
    escrow_label
])

# === Fonctions utilitaires ===
def _to_num(x):
    try:
        if pd.isna(x):
            return 0.0
        s = str(x).strip().replace('\u202f', '').replace('\xa0', '')
        s = s.replace("€", "").replace(" ", "").replace(",", ".")
        return float(s)
    except Exception:
        try:
            return float(x)
        except Exception:
            return 0.0


def recalc(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()
    for col in df.columns:
        if "Date" in col and df[col].dtype == object:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "Montant total" in df.columns:
        df["Montant total"] = df["Montant total"].apply(_to_num)
    if "Acompte 1" in df.columns:
        df["Acompte 1"] = df["Acompte 1"].apply(_to_num)
    if "Montant total" in df.columns and "Acompte 1" in df.columns:
        df["Solde"] = (df["Montant total"] - df["Acompte 1"]).fillna(0)
    return df


def download_excel_button(df: pd.DataFrame, label: str, filename: str):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    st.download_button(label, data=buf.getvalue(), file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# === Onglet Dashboard ===
with tabs[0]:
    st.header("📊 Tableau de bord")
    df = recalc(st.session_state.df_dossiers)
    if df.empty:
        st.info("Aucune donnée. Ajoutez un dossier ou chargez votre fichier Excel (Clients BL.xlsx).")
    else:
        total_dossiers = len(df)
        total_montant = df["Montant total"].sum()
        total_acompte = df["Acompte 1"].sum()
        total_solde = df["Solde"].sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Nombre de dossiers", total_dossiers)
        c2.metric("Montant total", f"{total_montant:,.2f} €".replace(",", " ").replace(".", ","))
        c3.metric("Total encaissé (Acompte 1)", f"{total_acompte:,.2f} €".replace(",", " ").replace(".", ","))
        c4.metric("Solde restant", f"{total_solde:,.2f} €".replace(",", " ").replace(".", ","))

        if "Date" in df.columns:
            dft = df.copy()
            dft["Mois"] = dft["Date"].dt.to_period("M").astype(str)
            monthly = dft.groupby("Mois")[["Montant total"]].sum().reset_index()
            fig = px.bar(monthly, x="Mois", y="Montant total", title="Montant total par mois")
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Derniers dossiers")
        st.dataframe(df.sort_values("Date", ascending=False).head(20), use_container_width=True, height=380)


# === Onglet Analyses ===
with tabs[1]:
    st.header("📈 Analyses")
    df = recalc(st.session_state.df_dossiers)
    if df.empty:
        st.info("Aucune donnée.")
    else:
        if "Solde" in df.columns:
            st.subheader("Distribution des soldes")
            st.plotly_chart(px.histogram(df, x="Solde", nbins=20), use_container_width=True)
        st.subheader("Table complète (lecture seule)")
        st.dataframe(df, use_container_width=True, height=420)


# === Onglet Ajouter ===
with tabs[2]:
    st.header("➕ Ajouter un dossier")
    with st.form("form_add"):
        col1, col2, col3 = st.columns(3)
        dossier_num = col1.text_input("Dossier N")
        nom_client = col2.text_input("Nom")
        date_dossier = col3.date_input("Date", date.today())

        col4, col5, col6 = st.columns(3)
        montant_total = col4.text_input("Montant total (€)", value="0")
        acompte1 = col5.text_input("Acompte 1 (€)", value="0")
        date_acompte1 = col6.date_input("Date Acompte 1", date.today())

        col7, col8 = st.columns(2)
        dossier_envoye = col7.checkbox("Dossier envoyé ?")
        date_envoi = col8.date_input("Date envoi", date.today())

        escrow_flag = st.checkbox("Escrow ?")
        ok = st.form_submit_button("Ajouter")

    if ok:
        new_row = {
            "Dossier N": dossier_num,
            "Nom": nom_client,
            "Date": pd.to_datetime(date_dossier),
            "Montant total": montant_total,
            "Acompte 1": acompte1,
            "Date Acompte 1": pd.to_datetime(date_acompte1),
            "Dossier envoyé": 1 if dossier_envoye else 0,
            "Date envoi": pd.to_datetime(date_envoi) if dossier_envoye else "",
            "Escrow": 1 if escrow_flag else 0
        }
        st.session_state.df_dossiers, st.session_state.df_escrow = esc.add_dossier(
            st.session_state.df_dossiers, st.session_state.df_escrow, new_row
        )
        st.session_state.df = st.session_state.df_dossiers.copy()
        st.success("✅ Dossier ajouté et sauvegardé dans 'Clients BL.xlsx'.")


# === Onglet Gestion ===
with tabs[3]:
    st.header("✏️ Gestion des dossiers")
    df = st.session_state.df_dossiers.copy()
    st.dataframe(recalc(df), use_container_width=True, height=360)

    st.markdown("---")
    st.subheader("Modifier un dossier (envoi / escrow)")

    colA, colB = st.columns([1, 2])
    dossier_to_edit = colA.text_input("Dossier N à modifier")
    new_envoye = colB.checkbox("Dossier envoyé ?")
    date_envoi_new = colB.date_input("Date d'envoi", date.today())

    if st.button("Enregistrer la modification"):
        updates = {
            "Dossier envoyé": 1 if new_envoye else 0,
            "Date envoi": pd.to_datetime(date_envoi_new) if new_envoye else ""
        }
        st.session_state.df_dossiers, st.session_state.df_escrow, ok = esc.update_dossier(
            st.session_state.df_dossiers, st.session_state.df_escrow, dossier_to_edit, updates
        )
        st.session_state.df = st.session_state.df_dossiers.copy()
        if ok:
            st.success("✅ Dossier modifié et sauvegardé.")
        else:
            st.warning("Dossier non trouvé.")


# === Onglet Compta Client ===
with tabs[4]:
    st.header("💳 Compta Client")
    df = recalc(st.session_state.df_dossiers)
    if df.empty:
        st.info("Aucune donnée.")
    else:
        st.subheader("Totaux par client (Montant / Acompte / Solde)")
        grp = df.groupby("Nom")[["Montant total", "Acompte 1", "Solde"]].sum().reset_index()
        st.dataframe(grp.sort_values("Montant total", ascending=False), use_container_width=True, height=420)


# === Onglet Export ===
with tabs[5]:
    st.header("💾 Export")
    st.write("Téléchargez vos données au format Excel :")
    c1, c2 = st.columns(2)
    with c1:
        download_excel_button(st.session_state.df_dossiers, "📥 Exporter Dossiers", "dossiers_export.xlsx")
    with c2:
        download_excel_button(st.session_state.df_escrow, "📥 Exporter Escrow", "escrow_export.xlsx")


# === Onglet Escrow ===
with tabs[6]:
    st.header("🛡️ Escrow")
    if n_escrow > 0:
        st.error(f"⚠️ {n_escrow} dossier(s) Escrow à réclamer !")
    else:
        st.success("✅ Aucun Escrow à réclamer.")

    col1, col2 = st.columns(2)
    col1.subheader("À réclamer")
    col1.dataframe(escrow_pending, use_container_width=True, height=300)

    col2.subheader("Réclamés")
    col2.dataframe(esc.reclames(st.session_state.df_escrow), use_container_width=True, height=300)

    st.markdown("---")
    st.subheader("✅ Marquer un Escrow comme réclamé")
    num_rec = st.text_input("Numéro de dossier")
    if st.button("Marquer comme réclamé"):
        st.session_state.df_escrow = esc.mark_reclaimed(st.session_state.df_escrow, num_rec)
        esc.save_data(st.session_state.df_dossiers, st.session_state.df_escrow)
        st.success(f"Dossier {num_rec} marqué comme réclamé et sauvegardé.")
