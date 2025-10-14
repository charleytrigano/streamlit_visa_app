from __future__ import annotations

# =======================================================
# 🛂 VISA MANAGER — Application principale Streamlit
# Version stable avec mémoire des fichiers et structure 6 onglets
# =======================================================

import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, datetime
import plotly.express as px
import os, io, json, zipfile, re

# -------------------------------------------------------
# 🧰 CONFIGURATION GÉNÉRALE
# -------------------------------------------------------
st.set_page_config(
    page_title="Visa Manager",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------------
# 🔧 FONCTIONS UTILES
# -------------------------------------------------------
def _safe_str(x):
    """Convertit n'importe quelle valeur en texte sûr."""
    try:
        if pd.isna(x):
            return ""
        return str(x)
    except Exception:
        return ""

def _safe_num(x):
    """Convertit en float sûr."""
    try:
        return float(x)
    except Exception:
        return 0.0

def _fmt_money(v):
    """Affichage formaté en dollars."""
    try:
        return f"${float(v):,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(v):
    """Retourne une date sûre pour Streamlit."""
    if isinstance(v, (date, datetime)):
        return v
    try:
        d = pd.to_datetime(v, errors="coerce")
        if pd.notna(d):
            return d.date()
    except Exception:
        pass
    return date.today()

def _ensure_dir(path):
    """Crée un dossier si nécessaire."""
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass

# -------------------------------------------------------
# 💾 CHEMINS DE BASE (mémoire locale)
# -------------------------------------------------------
DATA_DIR = "./"
CLIENTS_FILE = os.path.join(DATA_DIR, "donnees_visa_clients1.xlsx")
VISA_FILE = os.path.join(DATA_DIR, "donnees_visa_clients1.xlsx")
MEMORY_FILE = os.path.join(DATA_DIR, "last_used_files.json")

_ensure_dir(DATA_DIR)

def save_last_used_files(clients_path: str, visa_path: str):
    """Sauvegarde les derniers chemins utilisés."""
    data = {"clients": clients_path, "visa": visa_path}
    with open(MEMORY_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f)

def load_last_used_files() -> tuple[str, str]:
    """Charge les derniers chemins utilisés."""
    if os.path.exists(MEMORY_FILE):
        try:
            with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("clients", ""), data.get("visa", "")
        except Exception:
            pass
    return "", ""

# -------------------------------------------------------
# 📦 CHARGEMENT DES FICHIERS
# -------------------------------------------------------
@st.cache_data(show_spinner=False)
def read_clients_file(path):
    """Lecture du fichier Clients (xlsx ou csv)."""
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        if path.endswith(".csv"):
            return pd.read_csv(path)
        else:
            return pd.read_excel(path)
    except Exception:
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
def read_visa_file(path):
    """Lecture du fichier Visa (xlsx ou csv)."""
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        if path.endswith(".csv"):
            return pd.read_csv(path)
        else:
            return pd.read_excel(path)
    except Exception:
        return pd.DataFrame()


# =======================================================
# 📂 CHARGEMENT & MÉMOIRE DES FICHIERS
# =======================================================

st.sidebar.header("📂 Fichiers")

# --- Lecture des derniers fichiers utilisés ---
last_clients, last_visa = load_last_used_files()

mode = st.sidebar.radio(
    "Mode de chargement",
    ["Un fichier (Clients)", "Deux fichiers (Clients + Visa)"],
    horizontal=False
)

uploaded_clients = None
uploaded_visa = None

if mode == "Un fichier (Clients)":
    uploaded_clients = st.sidebar.file_uploader(
        "Clients (xlsx/csv)", type=["xlsx", "csv"], key="file_clients"
    )
else:
    uploaded_clients = st.sidebar.file_uploader(
        "Clients (xlsx/csv)", type=["xlsx", "csv"], key="file_clients_sep"
    )
    uploaded_visa = st.sidebar.file_uploader(
        "Visa (xlsx/csv)", type=["xlsx", "csv"], key="file_visa_sep"
    )

# --- Gestion du stockage local ---
clients_path_curr = last_clients
visa_path_curr = last_visa

if uploaded_clients is not None:
    clients_path_curr = os.path.join(DATA_DIR, f"upload_{uploaded_clients.name}")
    with open(clients_path_curr, "wb") as f:
        f.write(uploaded_clients.getbuffer())

if uploaded_visa is not None:
    visa_path_curr = os.path.join(DATA_DIR, f"upload_{uploaded_visa.name}")
    with open(visa_path_curr, "wb") as f:
        f.write(uploaded_visa.getbuffer())

# --- Sauvegarde des chemins en mémoire ---
if clients_path_curr or visa_path_curr:
    save_last_used_files(clients_path_curr, visa_path_curr)

# --- Chargement des fichiers ---
df_clients_raw = read_clients_file(clients_path_curr)
df_visa_raw = read_visa_file(visa_path_curr)

# --- Vérification ---
if df_clients_raw.empty:
    st.warning("⚠️ Aucun fichier Clients valide trouvé.")
else:
    st.sidebar.success(f"✅ Clients chargés : `{os.path.basename(clients_path_curr)}`")

if mode == "Deux fichiers (Clients + Visa)":
    if df_visa_raw.empty:
        st.sidebar.warning("⚠️ Aucun fichier Visa trouvé.")
    else:
        st.sidebar.success(f"✅ Visa chargé : `{os.path.basename(visa_path_curr)}`")

# -------------------------------------------------------
# 🗂️ Méta-informations
# -------------------------------------------------------
st.markdown("### 📄 Fichiers chargés")
st.write(f"**Clients** : `{clients_path_curr or 'Non chargé'}`")
if mode == "Deux fichiers (Clients + Visa)":
    st.write(f"**Visa** : `{visa_path_curr or 'Non chargé'}`")

st.divider()

# -------------------------------------------------------
# 📁 OPTIONS DE SAUVEGARDE
# -------------------------------------------------------
st.markdown("#### 💾 Chemin de sauvegarde")
save_dir = st.text_input(
    "**Chemin de sauvegarde** (sur ton PC / Drive / OneDrive) :",
    value=DATA_DIR,
    key="save_dir"
)

colA, colB = st.columns(2)
with colA:
    if st.button("Sauvegarder Clients vers…"):
        if not df_clients_raw.empty:
            save_path = os.path.join(save_dir, "Clients_sauvegarde.xlsx")
            df_clients_raw.to_excel(save_path, index=False)
            st.success(f"Clients sauvegardés dans : {save_path}")
        else:
            st.warning("Aucune donnée Clients à sauvegarder.")
with colB:
    if st.button("Sauvegarder Visa vers…"):
        if not df_visa_raw.empty:
            save_path = os.path.join(save_dir, "Visa_sauvegarde.xlsx")
            df_visa_raw.to_excel(save_path, index=False)
            st.success(f"Visa sauvegardé dans : {save_path}")
        else:
            st.warning("Aucune donnée Visa à sauvegarder.")



# =======================================================
# 🧽 HARMONISATION / NORMALISATION DES DONNÉES
# =======================================================

# Cartographie des colonnes possibles -> noms internes standardisés
COLMAP = {
    # identifiants / info client
    "id_client": "ID_Client",
    "id client": "ID_Client",
    "id": "ID_Client",
    "dossier n": "Dossier N",
    "dossier n°": "Dossier N",
    "nom": "Nom",
    "date": "Date",
    "mois": "Mois",

    # visa
    "categories": "Categorie",
    "categorie": "Categorie",
    "category": "Categorie",
    "sous-categorie": "Sous-categorie",
    "sous categorie": "Sous-categorie",
    "sous-categories": "Sous-categorie",
    "visa": "Visa",

    # montants
    "montant honoraires (us $)": "Montant honoraires (US $)",
    "montant honoraires": "Montant honoraires (US $)",
    "honoraires (us $)": "Montant honoraires (US $)",
    "autres frais (us $)": "Autres frais (US $)",
    "autres frais": "Autres frais (US $)",
    "total (us $)": "Total (US $)",
    "payé": "Payé",
    "paye": "Payé",
    "reste": "Reste",
    "solde": "Reste",
    "acomptes": "Paiements",
    "paiements": "Paiements",
    "acompte 1": "Acompte 1",
    "acompte 2": "Acompte 2",

    # statuts
    "rfe": "RFE",
    "dossiers envoyé": "Dossier envoyé",
    "dossier envoyé": "Dossier envoyé",
    "dossier envoye": "Dossier envoyé",
    "dossier approuvé": "Dossier approuvé",
    "dossier approuve": "Dossier approuvé",
    "dossier refusé": "Dossier refusé",
    "dossier refuse": "Dossier refusé",
    "dossier annulé": "Dossier annulé",
    "dossier annule": "Dossier annulé",

    # meta
    "commentaires": "Commentaires",
}

def _norm_colname(c: str) -> str:
    if not isinstance(c, str):
        return _safe_str(c)
    s = c.strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("’", "'").replace("é","e").replace("è","e").replace("ê","e").replace("à","a").replace("ç","c")
    return s

def harmonize_clients_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    # renommer colonnes selon COLMAP
    ren = {}
    for c in df.columns:
        key = _norm_colname(c)
        ren[c] = COLMAP.get(key, c)  # si pas dans la map, garder tel quel
    df = df.rename(columns=ren)

    # forcer présence des colonnes clés si absentes
    for must in ["ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
                 "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
                 "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"]:
        if must not in df.columns:
            df[must] = pd.Series([None]*len(df))

    # normaliser numériques
    for numc in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste","Acompte 1","Acompte 2"]:
        if numc in df.columns:
            df[numc] = pd.to_numeric(df[numc], errors="coerce").fillna(0.0)

    # Total et Reste cohérents
    if "Total (US $)" in df.columns:
        # si total manquant, recalculer
        mask_total_missing = df["Total (US $)"].isna() | (df["Total (US $)"] == 0)
        if "Montant honoraires (US $)" in df.columns and "Autres frais (US $)" in df.columns:
            df.loc[mask_total_missing, "Total (US $)"] = (
                pd.to_numeric(df["Montant honoraires (US $)"], errors="coerce").fillna(0)
                + pd.to_numeric(df["Autres frais (US $)"], errors="coerce").fillna(0)
            )

    if "Reste" in df.columns:
        mask_reste_missing = df["Reste"].isna()
        df.loc[mask_reste_missing, "Reste"] = (
            pd.to_numeric(df.get("Total (US $)", 0), errors="coerce").fillna(0)
            - pd.to_numeric(df.get("Payé", 0), errors="coerce").fillna(0)
        )
        df["Reste"] = df["Reste"].clip(lower=0)

    # statuts -> 0/1
    for sc in ["Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE"]:
        if sc in df.columns:
            df[sc] = pd.to_numeric(df[sc], errors="coerce").fillna(0).astype(int)

    # Date -> date + colonnes techniques
    if "Date" in df.columns:
        d = pd.to_datetime(df["Date"], errors="coerce")
        df["_Année_"] = d.dt.year.fillna(0).astype(int)
        df["_MoisNum_"] = d.dt.month.fillna(0).astype(int)
        # Mois (MM)
        if "Mois" in df.columns:
            # si Mois absent ou invalide, recalcule
            bad = df["Mois"].isna() | (df["Mois"].astype(str).str.strip() == "") | (df["Mois"].astype(str)=="NaT")
            df.loc[bad, "Mois"] = d.dt.month.fillna(0).astype(int).map(lambda m: f"{int(m):02d}" if m>0 else "")
        else:
            df["Mois"] = d.dt.month.fillna(0).astype(int).map(lambda m: f"{int(m):02d}" if m>0 else "")

    # Categorie / Sous-categorie / Visa -> string propre
    for sc in ["Categorie","Sous-categorie","Visa","Nom"]:
        if sc in df.columns:
            df[sc] = df[sc].astype(str).fillna("").replace("nan","").str.strip()

    return df

# Appliquer l’harmonisation sur les données chargées
df_all = harmonize_clients_columns(df_clients_raw.copy())

# =======================================================
# 📊 DASHBOARD & 📈 ANALYSES — TABS
# =======================================================
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Compte client", "🧾 Gestion", "📄 Visa (aperçu)"])

# -------------------------------------------------------
# 📊 DASHBOARD
# -------------------------------------------------------
with tabs[0]:
    st.markdown("### 📊 Dashboard")

    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # KPI compacts
        c1, c2, c3, c4, c5 = st.columns([1,1,1,1,1])
        total_dossiers = len(df_all)
        sum_total = float(pd.to_numeric(df_all["Total (US $)"], errors="coerce").fillna(0).sum())
        sum_paye  = float(pd.to_numeric(df_all["Payé"], errors="coerce").fillna(0).sum())
        sum_reste = float(pd.to_numeric(df_all["Reste"], errors="coerce").fillna(0).sum())
        pct_env   = (df_all["Dossier envoyé"].fillna(0).astype(int).sum() / total_dossiers * 100) if total_dossiers else 0.0

        c1.metric("Dossiers", f"{total_dossiers}")
        c2.metric("Honoraires+Frais", _fmt_money(sum_total))
        c3.metric("Payé", _fmt_money(sum_paye))
        c4.metric("Solde", _fmt_money(sum_reste))
        c5.metric("Envoyés (%)", f"{pct_env:.0f}%")

        st.markdown("#### 🎛️ Filtres")
        cats  = sorted([c for c in df_all["Categorie"].dropna().astype(str).unique().tolist() if c])
        subs  = sorted([c for c in df_all["Sous-categorie"].dropna().astype(str).unique().tolist() if c])
        visas = sorted([c for c in df_all["Visa"].dropna().astype(str).unique().tolist() if c])

        a1,a2,a3 = st.columns(3)
        fc = a1.multiselect("Catégories", cats, default=[], key="dash_cats")
        fs = a2.multiselect("Sous-catégories", subs, default=[], key="dash_subs")
        fv = a3.multiselect("Visa", visas, default=[], key="dash_visas")

        view = df_all.copy()
        if fc:
            view = view[view["Categorie"].astype(str).isin(fc)]
        if fs:
            view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv:
            view = view[view["Visa"].astype(str).isin(fv)]

        # Graph 1 : nombre de dossiers par catégorie
        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if not view.empty and "Categorie" in view.columns:
            vc = view["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie","Nombre"]
            st.bar_chart(vc.set_index("Categorie"))

        # Graph 2 : flux par mois (Honoraires / Autres frais / Payé / Reste)
        st.markdown("#### 💵 Flux par mois")
        tmp = view.copy()
        if not tmp.empty:
            tmp["_MoisLbl_"] = tmp.apply(
                lambda r: (f"{int(r['_Année_']):04d}-{int(r['_MoisNum_']):02d}" 
                           if (int(r.get('_Année_',0))>0 and int(r.get('_MoisNum_',0))>0) else "NaT"), axis=1)
            g = tmp.groupby("_MoisLbl_", as_index=False)[["Montant honoraires (US $)","Autres frais (US $)","Payé","Reste"]].sum()
            g = g.sort_values("_MoisLbl_")
            # Plotly pour courbes superposées lisibles
            try:
                import plotly.express as px
                fg = g.melt(id_vars=["_MoisLbl_"], var_name="Type", value_name="Montant")
                fig = px.line(fg, x="_MoisLbl_", y="Montant", color="Type")
                st.plotly_chart(fig, use_container_width=True, key="dash_flow_plot")
            except Exception:
                st.line_chart(g.set_index("_MoisLbl_"))

        # Table détails
        st.markdown("#### 📋 Détails (après filtres)")
        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in view.columns]

        # mise en forme monétaire
        v2 = view.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in v2.columns:
                v2[c] = pd.to_numeric(v2[c], errors="coerce").fillna(0.0).map(_fmt_money)

        # dates lisibles
        if "Date" in v2.columns:
            try:
                v2["Date"] = pd.to_datetime(v2["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                v2["Date"] = v2["Date"].astype(str)

        sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
        v2 = v2.sort_values(by=sort_keys) if sort_keys else v2
        st.dataframe(v2[show_cols].reset_index(drop=True), use_container_width=True, key="dash_detail_table")


# -------------------------------------------------------
# 📈 ANALYSES
# -------------------------------------------------------
with tabs[1]:
    st.markdown("### 📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist() if int(y)>0])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted([c for c in df_all["Categorie"].dropna().astype(str).unique().tolist() if c])
        subsA   = sorted([c for c in df_all["Sous-categorie"].dropna().astype(str).unique().tolist() if c])
        visasA  = sorted([c for c in df_all["Visa"].dropna().astype(str).unique().tolist() if c])

        a1,a2,a3,a4,a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key="ana_years")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key="ana_months")
        fc = a3.multiselect("Catégorie", catsA, default=[], key="ana_cats")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key="ana_subs")
        fv = a5.multiselect("Visa", visasA, default=[], key="ana_visas")

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(pd.to_numeric(dfA["Montant honoraires (US $)"], errors="coerce").fillna(0).sum()))
        k3.metric("Payé", _fmt_money(pd.to_numeric(dfA["Payé"], errors="coerce").fillna(0).sum()))
        k4.metric("Reste", _fmt_money(pd.to_numeric(dfA["Reste"], errors="coerce").fillna(0).sum()))

        # Graphes d'analyse
        st.markdown("#### 📊 Répartition par catégorie")
        if not dfA.empty and "Categorie" in dfA.columns:
            vc = dfA["Categorie"].value_counts(normalize=False).reset_index()
            vc.columns = ["Categorie","Nombre"]
            vc["%"] = (vc["Nombre"] / max(1,len(dfA))) * 100
            st.dataframe(vc, use_container_width=True, key="ana_cat_table")

        st.markdown("#### 📈 Honoraires par mois")
        if not dfA.empty:
            tmp = dfA.copy()
            tmp["_MoisLbl_"] = tmp.apply(
                lambda r: (f"{int(r['_Année_']):04d}-{int(r['_MoisNum_']):02d}" 
                           if (int(r.get('_Année_',0))>0 and int(r.get('_MoisNum_',0))>0) else "NaT"), axis=1)
            gm = tmp.groupby("_MoisLbl_", as_index=False)["Montant honoraires (US $)"].sum().sort_values("_MoisLbl_")
            try:
                import plotly.express as px
                fig2 = px.bar(gm, x="_MoisLbl_", y="Montant honoraires (US $)")
                st.plotly_chart(fig2, use_container_width=True, key="ana_hono_plot")
            except Exception:
                st.bar_chart(gm.set_index("_MoisLbl_"))

        # Détails
        st.markdown("#### 🧾 Détails des dossiers filtrés")
        det = dfA.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in det.columns:
                det[c] = pd.to_numeric(det[c], errors="coerce").fillna(0.0).map(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det
        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key="ana_detail_table")



# =======================================================
# 🏦 ESCROW — SYNTHÈSE
# Onglet tabs[2] (voir création des tabs en Partie 3/6)
# =======================================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")

    if df_all.empty:
        st.info("Aucun client chargé.")
    else:
        # Normalisations sûres
        dfE = df_all.copy()
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Total (US $)", "Payé", "Reste"]:
            if c in dfE.columns:
                dfE[c] = pd.to_numeric(dfE[c], errors="coerce").fillna(0.0)

        # Rappel logique simple : l'escrow correspond aux honoraires payés mais non encore "envoyés".
        # (Sans historique paiement par paiement, on approxime : si Dossier envoyé==0,
        #   alors tout le Payé (limité aux honoraires) reste "en escrow".)
        dfE["Escrow estimé"] = 0.0
        if "Dossier envoyé" in dfE.columns:
            mask_non_envoye = (pd.to_numeric(dfE["Dossier envoyé"], errors="coerce").fillna(0).astype(int) == 0)
        else:
            mask_non_envoye = pd.Series([True] * len(dfE), index=dfE.index)

        hono = dfE["Montant honoraires (US $)"]
        paye = dfE["Payé"]
        dfE.loc[mask_non_envoye, "Escrow estimé"] = np.minimum(paye, hono)

        # KPI compacts (taille réduite)
        k1, k2, k3, k4 = st.columns([1,1,1,1])
        k1.metric("Total (US $)", _fmt_money(float(dfE["Total (US $)"].sum())))
        k2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        k3.metric("Solde", _fmt_money(float(dfE["Reste"].sum())))
        k4.metric("Escrow (estimé)", _fmt_money(float(dfE["Escrow estimé"].sum())))

        # Tableau par catégorie
        st.markdown("#### 📦 Synthèse par catégorie")
        if "Categorie" in dfE.columns:
            agg = dfE.groupby("Categorie", as_index=False)[["Montant honoraires (US $)", "Payé", "Reste", "Escrow estimé"]].sum()
            # Pourcentage payé
            agg["% Payé"] = np.where(agg["Montant honoraires (US $)"] > 0,
                                     100 * agg["Payé"] / agg["Montant honoraires (US $)"], 0.0)
            st.dataframe(agg, use_container_width=True, key="escrow_cat_table")

        # Alerte dossiers "envoyés" sans avoir vidé l'escrow (signal recouvrement/affectation)
        st.markdown("#### ⚠️ Dossiers envoyés avec encaissements à affecter")
        if "Dossier envoyé" in dfE.columns:
            mask_envoye = (pd.to_numeric(dfE["Dossier envoyé"], errors="coerce").fillna(0).astype(int) == 1)
            df_env = dfE[mask_envoye].copy()
            # si la logique escrow doit être vidée au moment "envoyé", on signale les payés > honoraires déjà affectés
            df_env["A vérifier"] = np.maximum(0.0, df_env["Payé"] - df_env["Montant honoraires (US $)"])
            df_check = df_env[df_env["A vérifier"] > 0][
                ["Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Payé","Montant honoraires (US $)","A vérifier"]
            ]
            if df_check.empty:
                st.caption("✅ RAS.")
            else:
                st.dataframe(df_check.reset_index(drop=True), use_container_width=True, key="escrow_alerts")


# =======================================================
# 👤 COMPTE CLIENT — Fiche & ajout de règlement
# Onglet tabs[3]
# =======================================================
with tabs[3]:
    st.subheader("👤 Compte client")

    if df_all.empty:
        st.info("Aucun client chargé.")
    else:
        # Sélecteurs
        names = sorted([x for x in df_all["Nom"].dropna().astype(str).unique().tolist() if x])
        ids   = sorted([x for x in df_all["ID_Client"].dropna().astype(str).unique().tolist() if x])

        csel1, csel2 = st.columns(2)
        sel_name = csel1.selectbox("Nom", [""] + names, index=0, key="acct_sel_name")
        sel_id   = csel2.selectbox("ID_Client", [""] + ids,   index=0, key="acct_sel_id")

        # Filtrage
        mask = None
        if sel_id:
            mask = (df_all["ID_Client"].astype(str) == sel_id)
        elif sel_name:
            mask = (df_all["Nom"].astype(str) == sel_name)

        if mask is None or not mask.any():
            st.stop()

        # Fiche
        row = df_all[mask].iloc[0].copy()

        # KPI fiche
        c1, c2, c3, c4 = st.columns(4)
        total_ = float(pd.to_numeric(row.get("Total (US $)"), errors="coerce") or 0.0)
        paye_  = float(pd.to_numeric(row.get("Payé"), errors="coerce") or 0.0)
        reste_ = float(pd.to_numeric(row.get("Reste"), errors="coerce") or (total_ - paye_))
        hono_  = float(pd.to_numeric(row.get("Montant honoraires (US $)"), errors="coerce") or 0.0)
        c1.metric("Total", _fmt_money(total_))
        c2.metric("Payé", _fmt_money(paye_))
        c3.metric("Solde", _fmt_money(reste_))
        c4.metric("Honoraires", _fmt_money(hono_))

        st.markdown("#### 🗂️ Dossier")
        s1, s2 = st.columns([1,2])
        s1.write(f"**Dossier N** : {_safe_str(row.get('Dossier N',''))}")
        s1.write(f"**ID_Client** : {_safe_str(row.get('ID_Client',''))}")
        s1.write(f"**Nom** : {_safe_str(row.get('Nom',''))}")
        s2.write(f"**Catégorie** : {_safe_str(row.get('Categorie',''))}")
        s2.write(f"**Sous-catégorie** : {_safe_str(row.get('Sous-categorie',''))}")
        s2.write(f"**Visa** : {_safe_str(row.get('Visa',''))}")

        # Statuts + dates (affichage)
        st.markdown("#### 📌 Statuts")
        s1, s2 = st.columns(2)
        try_int = lambda v: int(pd.to_numeric(v, errors="coerce") or 0)
        s1.write(
            "- Dossier envoyé : {} | Date : {}".format(
                try_int(row.get("Dossier envoyé", 0)),
                _safe_str(row.get("Date d'envoi", ""))
            )
        )
        s1.write(
            "- Dossier accepté : {} | Date : {}".format(
                try_int(row.get("Dossier approuvé", 0)),
                _safe_str(row.get("Date d'acceptation", ""))
            )
        )
        s2.write(
            "- Dossier refusé : {} | Date : {}".format(
                try_int(row.get("Dossier refusé", 0)),
                _safe_str(row.get("Date de refus", ""))
            )
        )
        s2.write(
            "- Dossier annulé : {} | Date : {}".format(
                try_int(row.get("Dossier annulé", 0)),
                _safe_str(row.get("Date d'annulation", ""))
            )
        )
        s2.write(f"- RFE : {try_int(row.get('RFE', 0))}")

        # Commentaires
        st.markdown("#### 📝 Commentaires")
        st.write(_safe_str(row.get("Commentaires", "")))

        # ---------------------------------------------------
        # 💸 Ajout d’un règlement (si solde > 0)
        # ---------------------------------------------------
        st.markdown("#### 💸 Ajouter un règlement")
        if reste_ <= 0:
            st.success("Ce dossier est soldé.")
        else:
            pay_col1, pay_col2, pay_col3, pay_col4 = st.columns([1,1,1,2])
            pay_date = pay_col1.date_input("Date", value=date.today(), key="acct_pay_date")
            pay_mode = pay_col2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], index=0, key="acct_pay_mode")
            pay_amt  = pay_col3.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key="acct_pay_amt")
            note     = pay_col4.text_input("Note (facultatif)", "", key="acct_pay_note")

            if st.button("💾 Enregistrer le règlement sur ce client", key="acct_pay_save"):
                if pay_amt <= 0:
                    st.warning("Le montant doit être > 0.")
                    st.stop()

                # Charger le fichier "live" (pour éviter les écarts) et mettre à jour la ligne
                live = read_clients_file(clients_path_curr).copy()
                if live.empty:
                    st.error("Fichier clients introuvable en écriture.")
                    st.stop()

                # On cherche la ligne par ID_Client (prioritaire) sinon par Nom
                if sel_id:
                    m2 = (live["ID_Client"].astype(str) == sel_id)
                else:
                    m2 = (live["Nom"].astype(str) == sel_name)

                if not m2.any():
                    st.error("Ligne introuvable dans le fichier.")
                    st.stop()

                idx = live[m2].index[0]
                # Recalcule Payé/Reste
                old_paye  = float(pd.to_numeric(live.at[idx, "Payé"], errors="coerce") or 0.0)
                new_paye  = old_paye + float(pay_amt)
                total_    = float(pd.to_numeric(live.at[idx, "Total (US $)"], errors="coerce") or 0.0)
                new_reste = max(0.0, total_ - new_paye)

                live.at[idx, "Payé"]  = new_paye
                live.at[idx, "Reste"] = new_reste

                # Historiser dans "Commentaires"
                cm = _safe_str(live.at[idx, "Commentaires"])
                line = f"[{pay_date}] Règlement {pay_mode}: ${pay_amt:,.2f}"
                if note:
                    line += f" — {note}"
                live.at[idx, "Commentaires"] = (cm + "\n" + line).strip() if cm else line

                # Ecriture
                write_clients_file(live, clients_path_curr)
                st.success("Règlement enregistré.")
                # Rafraîchir la page pour relire df_all via le cache
                st.cache_data.clear()
                st.rerun()


# =======================================================
# 🧾 GESTION (CRUD) — Onglet tabs[4]
# =======================================================
with tabs[4]:
    st.subheader("🧾 Gestion des clients")

    if df_all.empty:
        st.info("Aucun client chargé. Charge un fichier Clients dans la barre latérale.")
    else:
        # Relire le fichier "live" pour les opérations d'écriture
        df_live = read_clients_file(clients_path_curr).copy()

        # Petit util local pour sécuriser les dates dans les widgets
        def _date_for_widget(v, fallback=None):
            if isinstance(v, (date, datetime)):
                return v.date() if isinstance(v, datetime) else v
            try:
                d = pd.to_datetime(v, errors="coerce")
                if pd.notna(d):
                    return d.date()
            except Exception:
                pass
            return fallback if fallback is not None else date.today()

        op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key="crud_op")

        # ------------------ AJOUTER ------------------
        if op == "Ajouter":
            st.markdown("### ➕ Ajouter un client")

            c1, c2, c3 = st.columns(3)
            nom  = c1.text_input("Nom", "", key="add_nom")
            dcre = c2.date_input("Date de création", value=date.today(), key="add_date")
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1, 13)],
                                index=date.today().month - 1, key="add_mois")

            # Cascade Visa
            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key="add_cat")
            sel_sub = ""
            visa_final = ""
            opts_dict = {"exclusive": None, "options": []}
            info_msg = ""
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key="add_sub")
                if sel_sub:
                    visa_final, opts_dict, info_msg = build_visa_option_selector(
                        visa_map, sel_cat, sel_sub, keyprefix="add_opts", preselected={}
                    )
            if info_msg:
                st.info(info_msg)

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, step=50.0, format="%.2f", key="add_h")
            autre = f2.number_input("Autres frais (US $)", min_value=0.0, step=20.0, format="%.2f", key="add_a")

            com = st.text_area("Commentaires (notes, détails d’autres frais…)", "", key="add_comm", height=80)

            st.markdown("#### 📌 Statuts initiaux")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent = s1.checkbox("Dossier envoyé", key="add_sent")
            sent_d = s1.date_input("Date d'envoi", value=None, key="add_sent_d")
            acc = s2.checkbox("Dossier approuvé", key="add_acc")
            acc_d = s2.date_input("Date d'acceptation", value=None, key="add_acc_d")
            ref = s3.checkbox("Dossier refusé", key="add_ref")
            ref_d = s3.date_input("Date de refus", value=None, key="add_ref_d")
            ann = s4.checkbox("Dossier annulé", key="add_ann")
            ann_d = s4.date_input("Date d'annulation", value=None, key="add_ann_d")
            rfe = s5.checkbox("RFE", key="add_rfe")

            if rfe and not any([sent, acc, ref, ann]):
                st.warning("⚠️ La case RFE ne peut être cochée qu’avec un autre statut (envoyé/approuvé/refusé/annulé).")

            if st.button("💾 Enregistrer le client", key="btn_add"):
                if not nom:
                    st.warning("Le nom est requis.")
                    st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisis la catégorie et la sous-catégorie.")
                    st.stop()

                total = float(honor) + float(autre)
                paye = 0.0
                reste = max(0.0, total - paye)

                # ID client et N° dossier
                did = _make_client_id(nom, dcre)
                dossier_n = _next_dossier(df_live, start=13057)

                new_row = {
                    "Dossier N": dossier_n,
                    "ID_Client": did,
                    "Nom": nom,
                    "Date": dcre,
                    "Mois": f"{int(mois):02d}" if isinstance(mois, (int, str)) else _safe_str(mois),
                    "Categorie": sel_cat,
                    "Sous-categorie": sel_sub,
                    "Visa": visa_final if visa_final else sel_sub,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(autre),
                    "Total (US $)": total,
                    "Payé": 0.0,
                    "Reste": reste,
                    "Commentaires": com,
                    "Options": opts_dict,
                    "Dossier envoyé": 1 if sent else 0,
                    "Date d'envoi": (dcre if sent and not sent_d else sent_d),
                    "Dossier approuvé": 1 if acc else 0,
                    "Date d'acceptation": acc_d,
                    "Dossier refusé": 1 if ref else 0,
                    "Date de refus": ref_d,
                    "Dossier annulé": 1 if ann else 0,
                    "Date d'annulation": ann_d,
                    "RFE": 1 if rfe else 0,
                }

                df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                write_clients_file(df_new, clients_path_curr)
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()

        # ------------------ MODIFIER ------------------
        elif op == "Modifier":
            st.markdown("### ✏️ Modifier un client")
            if df_live.empty:
                st.info("Aucun client dans le fichier.")
            else:
                ids = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
                names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
                csel1, csel2 = st.columns(2)
                sel_id = csel1.selectbox("ID_Client", [""] + ids, index=0, key="mod_id_sel")
                sel_nm = csel2.selectbox("Nom", [""] + names, index=0, key="mod_nm_sel")

                if sel_id:
                    m = (df_live["ID_Client"].astype(str) == sel_id)
                elif sel_nm:
                    m = (df_live["Nom"].astype(str) == sel_nm)
                else:
                    m = None

                if not m is None and m.any():
                    idx = df_live[m].index[0]
                    row = df_live.loc[idx].copy()

                    d1, d2, d3 = st.columns(3)
                    nom = d1.text_input("Nom", _safe_str(row.get("Nom", "")), key="mod_nom")
                    dval = _date_for_widget(row.get("Date"), fallback=date.today())
                    dcre = d2.date_input("Date de création", value=dval, key="mod_date")
                    mois_curr = _safe_str(row.get("Mois", "01"))
                    try:
                        mois_idx = max(0, min(11, int(mois_curr) - 1))
                    except Exception:
                        mois_idx = 0
                    mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1, 13)],
                                        index=mois_idx, key="mod_mois")

                    st.markdown("#### 🎯 Choix Visa")
                    cats = sorted(list(visa_map.keys()))
                    cat0 = _safe_str(row.get("Categorie", ""))
                    sel_cat = st.selectbox("Catégorie", [""] + cats,
                                           index=(cats.index(cat0) + 1 if cat0 in cats else 0), key="mod_cat")

                    sub0 = _safe_str(row.get("Sous-categorie", ""))
                    sel_sub = ""
                    if sel_cat:
                        subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                        sel_sub = st.selectbox("Sous-catégorie", [""] + subs,
                                               index=(subs.index(sub0) + 1 if sub0 in subs else 0), key="mod_sub")

                    # Pré-sélection d’options
                    preset = row.get("Options", {})
                    if not isinstance(preset, dict):
                        try:
                            preset = json.loads(_safe_str(preset) or "{}")
                            if not isinstance(preset, dict):
                                preset = {}
                        except Exception:
                            preset = {}

                    visa_final, opts_dict, info_msg = "", {"exclusive": None, "options": []}, ""
                    if sel_cat and sel_sub:
                        visa_final, opts_dict, info_msg = build_visa_option_selector(
                            visa_map, sel_cat, sel_sub, keyprefix="mod_opts", preselected=preset
                        )
                    if info_msg:
                        st.info(info_msg)

                    f1, f2 = st.columns(2)
                    honor = f1.number_input(
                        "Montant honoraires (US $)", min_value=0.0,
                        value=float(pd.to_numeric(row.get("Montant honoraires (US $)"), errors="coerce") or 0.0),
                        step=50.0, format="%.2f", key="mod_h"
                    )
                    autre = f2.number_input(
                        "Autres frais (US $)", min_value=0.0,
                        value=float(pd.to_numeric(row.get("Autres frais (US $)"), errors="coerce") or 0.0),
                        step=20.0, format="%.2f", key="mod_a"
                    )

                    com = st.text_area(
                        "Commentaires (notes, détails d’autres frais…)",
                        _safe_str(row.get("Commentaires", "")), key="mod_comm", height=80
                    )

                    st.markdown("#### 📌 Statuts")
                    s1, s2, s3, s4, s5 = st.columns(5)
                    to_int = lambda v: int(pd.to_numeric(v, errors="coerce") or 0)

                    sent = s1.checkbox("Dossier envoyé", value=(to_int(row.get("Dossier envoyé", 0)) == 1), key="mod_sent")
                    sent_d = s1.date_input(
                        "Date d'envoi",
                        value=_date_for_widget(row.get("Date d'envoi"), fallback=None),
                        key="mod_sent_d"
                    )
                    acc = s2.checkbox("Dossier approuvé", value=(to_int(row.get("Dossier approuvé", 0)) == 1), key="mod_acc")
                    acc_d = s2.date_input(
                        "Date d'acceptation",
                        value=_date_for_widget(row.get("Date d'acceptation"), fallback=None),
                        key="mod_acc_d"
                    )
                    ref = s3.checkbox("Dossier refusé", value=(to_int(row.get("Dossier refusé", 0)) == 1), key="mod_ref")
                    ref_d = s3.date_input(
                        "Date de refus",
                        value=_date_for_widget(row.get("Date de refus"), fallback=None),
                        key="mod_ref_d"
                    )
                    ann = s4.checkbox("Dossier annulé", value=(to_int(row.get("Dossier annulé", 0)) == 1), key="mod_ann")
                    ann_d = s4.date_input(
                        "Date d'annulation",
                        value=_date_for_widget(row.get("Date d'annulation"), fallback=None),
                        key="mod_ann_d"
                    )
                    rfe = s5.checkbox("RFE", value=(to_int(row.get("RFE", 0)) == 1), key="mod_rfe")

                    if rfe and not any([sent, acc, ref, ann]):
                        st.warning("⚠️ La case RFE ne peut être cochée qu’avec un autre statut.")

                    if st.button("💾 Enregistrer les modifications", key="btn_mod"):
                        if not nom:
                            st.warning("Le nom est requis.")
                            st.stop()
                        if not sel_cat or not sel_sub:
                            st.warning("Choisis la catégorie et la sous-catégorie.")
                            st.stop()

                        total = float(honor) + float(autre)
                        paye = float(pd.to_numeric(row.get("Payé"), errors="coerce") or 0.0)
                        reste = max(0.0, total - paye)

                        df_live.at[idx, "Nom"] = nom
                        df_live.at[idx, "Date"] = dcre
                        df_live.at[idx, "Mois"] = f"{int(mois):02d}" if isinstance(mois, (int, str)) else _safe_str(mois)
                        df_live.at[idx, "Categorie"] = sel_cat
                        df_live.at[idx, "Sous-categorie"] = sel_sub
                        df_live.at[idx, "Visa"] = (visa_final if visa_final else sel_sub)
                        df_live.at[idx, "Montant honoraires (US $)"] = float(honor)
                        df_live.at[idx, "Autres frais (US $)"] = float(autre)
                        df_live.at[idx, "Total (US $)"] = total
                        df_live.at[idx, "Reste"] = reste
                        df_live.at[idx, "Commentaires"] = com
                        df_live.at[idx, "Options"] = opts_dict

                        df_live.at[idx, "Dossier envoyé"] = 1 if sent else 0
                        df_live.at[idx, "Date d'envoi"] = (dcre if sent and not sent_d else sent_d)
                        df_live.at[idx, "Dossier approuvé"] = 1 if acc else 0
                        df_live.at[idx, "Date d'acceptation"] = acc_d
                        df_live.at[idx, "Dossier refusé"] = 1 if ref else 0
                        df_live.at[idx, "Date de refus"] = ref_d
                        df_live.at[idx, "Dossier annulé"] = 1 if ann else 0
                        df_live.at[idx, "Date d'annulation"] = ann_d
                        df_live.at[idx, "RFE"] = 1 if rfe else 0

                        write_clients_file(df_live, clients_path_curr)
                        st.success("Modifications enregistrées.")
                        st.cache_data.clear()
                        st.rerun()
                else:
                    st.info("Sélectionne un client à modifier.")

        # ------------------ SUPPRIMER ------------------
        elif op == "Supprimer":
            st.markdown("### 🗑️ Supprimer un client")
            if df_live.empty:
                st.info("Aucun client.")
            else:
                ids = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
                sel_id = st.selectbox("ID_Client", [""] + ids, index=0, key="del_id_sel")
                if sel_id:
                    m = (df_live["ID_Client"].astype(str) == sel_id)
                    if m.any():
                        r = df_live[m].iloc[0]
                        st.write({
                            "Dossier N": r.get("Dossier N", ""),
                            "Nom": r.get("Nom", ""),
                            "Visa": r.get("Visa", "")
                        })
                        if st.button("❗ Confirmer la suppression", key="btn_del"):
                            df_new = df_live[~m].copy()
                            write_clients_file(df_new, clients_path_curr)
                            st.success("Client supprimé.")
                            st.cache_data.clear()
                            st.rerun()
                else:
                    st.info("Sélectionne un ID client à supprimer.")


# =======================================================
# 📄 VISA (APERÇU) — Onglet tabs[5]
# =======================================================
with tabs[5]:
    st.subheader("📄 Visa — aperçu")

    if df_visa_raw.empty:
        st.info("Aucune donnée Visa chargée.")
    else:
        # Filtres simples sur Catégorie / Sous-catégorie
        v1, v2 = st.columns(2)
        cats_v = sorted(df_visa_raw["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_visa_raw.columns else []
        subs_v = sorted(df_visa_raw["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_visa_raw.columns else []
        fc = v1.multiselect("Catégorie", cats_v, default=[])
        fs = v2.multiselect("Sous-catégorie", subs_v, default=[])

        vdf = df_visa_raw.copy()
        if fc:
            vdf = vdf[vdf["Categorie"].astype(str).isin(fc)]
        if fs:
            vdf = vdf[vdf["Sous-categorie"].astype(str).isin(fs)]

        st.dataframe(vdf.reset_index(drop=True), use_container_width=True, height=420, key="visa_preview")



# =======================================================
# 🔧 Petits utilitaires locaux (autonomes pour cet onglet)
# =======================================================
SID6 = st.session_state.get("_sid", "p6")

def _ensure_time_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Ajoute _Année_, _MoisNum_ et Mois si absents (à partir de Date/Mois)."""
    out = df.copy()
    if "Mois" not in out.columns:
        # tente de déduire depuis Date
        if "Date" in out.columns:
            try:
                m = pd.to_datetime(out["Date"], errors="coerce").dt.month
                out["Mois"] = m.fillna(1).astype(int).apply(lambda x: f"{int(x):02d}")
            except Exception:
                out["Mois"] = "01"
        else:
            out["Mois"] = "01"
    # _MoisNum_
    try:
        out["_MoisNum_"] = pd.to_numeric(out["Mois"], errors="coerce").fillna(1).astype(int)
    except Exception:
        out["_MoisNum_"] = 1
    # _Année_
    if "_Année_" not in out.columns:
        if "Date" in out.columns:
            try:
                out["_Année_"] = pd.to_datetime(out["Date"], errors="coerce").dt.year
                out["_Année_"] = out["_Année_"].fillna(out["_Année_"].mode().iloc[0] if not out["_Année_"].mode().empty else date.today().year).astype(int)
            except Exception:
                out["_Année_"] = date.today().year
        else:
            out["_Année_"] = date.today().year
    return out

def _pct(a, b):
    a = float(a or 0); b = float(b or 0)
    return (a / b * 100.0) if b > 0 else 0.0


# =======================================================
# 📈 ONGLET : Analyses (séries + comparaisons + détails)
# =======================================================
with tabs[7]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucun client chargé. (Onglet « Fichiers » → charge un Clients.)")
    else:
        dfA0 = _ensure_time_cols(df_all)
        # Listes de valeurs
        yearsA  = sorted(pd.to_numeric(dfA0["_Année_"], errors="coerce").dropna().astype(int).unique().tolist())
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(dfA0.get("Categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        subsA   = sorted(dfA0.get("Sous-categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        visasA  = sorted(dfA0.get("Visa", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

        st.markdown("#### 🎛️ Filtres (ensemble global)")
        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_years_{SID6}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"a_months_{SID6}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"a_cats_{SID6}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"a_subs_{SID6}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_visas_{SID6}")

        dfA = dfA0.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # Montants normalisés
        hono = _safe_num_series(dfA, "Montant honoraires (US $)")
        autre = _safe_num_series(dfA, "Autres frais (US $)")
        total = (_safe_num_series(dfA, "Total (US $)") if "Total (US $)" in dfA.columns else (hono + autre))
        paye  = _safe_num_series(dfA, "Payé")
        reste = _safe_num_series(dfA, "Solde") if "Solde" in dfA.columns else (total - paye)

        # KPI compacts
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires+Frais", _fmt_money(float((hono+autre).sum())))
        k3.metric("Payé", _fmt_money(float(paye.sum())))
        k4.metric("Solde", _fmt_money(float(reste.sum())))
        pct_env = _pct(dfA.get("Dossier envoyé", 0).sum(), len(dfA))
        k5.metric("Envoyés (%)", f"{pct_env:.0f}%")

        # % par catégories / sous-catégories (sur nombre de dossiers)
        st.markdown("#### 📊 Répartition (nombre de dossiers)")
        c11, c12 = st.columns(2)
        if not dfA.empty:
            df_cnt_cat = (dfA.groupby("Categorie", as_index=False)
                            .size().rename(columns={"size":"Dossiers"})).sort_values("Dossiers", ascending=False)
            df_cnt_cat["%"] = (df_cnt_cat["Dossiers"] / max(1, df_cnt_cat["Dossiers"].sum()) * 100).round(1)
            c11.dataframe(df_cnt_cat, use_container_width=True, height=240, key=f"a_cnt_cat_{SID6}")

            if "Sous-categorie" in dfA.columns:
                df_cnt_sub = (dfA.groupby("Sous-categorie", as_index=False)
                                .size().rename(columns={"size":"Dossiers"})).sort_values("Dossiers", ascending=False)
                df_cnt_sub["%"] = (df_cnt_sub["Dossiers"] / max(1, df_cnt_sub["Dossiers"].sum()) * 100).round(1)
                c12.dataframe(df_cnt_sub, use_container_width=True, height=240, key=f"a_cnt_sub_{SID6}")
            else:
                c12.info("Aucune sous-catégorie dans les données.")

        # Flux par mois (honoraires, frais, payé, solde)
        st.markdown("#### 💵 Flux par mois")
        tmp = dfA.copy()
        tmp["Mois"] = tmp["Mois"].astype(str)
        flux = (tmp.groupby("Mois", as_index=False)
                    .agg({
                        "Montant honoraires (US $)": "sum",
                        "Autres frais (US $)": "sum",
                        "Payé": "sum"
                    }))
        flux = flux.sort_values("Mois")
        flux["Solde"] = (flux["Montant honoraires (US $)"] + flux["Autres frais (US $)"] - flux["Payé"]).clip(lower=0)
        st.bar_chart(flux.set_index("Mois")[["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]])

        # Comparaison A vs B (périodes / catégories)
        st.markdown("#### ⚖️ Comparaison A vs B (périodes / filtres)")
        ca1, ca2, ca3 = st.columns(3)
        ya = ca1.multiselect("Année (A)", yearsA, default=[], key=f"cmp_ya_{SID6}")
        ma = ca2.multiselect("Mois (A)", monthsA, default=[], key=f"cmp_ma_{SID6}")
        ca = ca3.multiselect("Catégories (A)", catsA, default=[], key=f"cmp_ca_{SID6}")

        cb1, cb2, cb3 = st.columns(3)
        yb = cb1.multiselect("Année (B)", yearsA, default=[], key=f"cmp_yb_{SID6}")
        mb = cb2.multiselect("Mois (B)", monthsA, default=[], key=f"cmp_mb_{SID6}")
        cb = cb3.multiselect("Catégories (B)", catsA, default=[], key=f"cmp_cb_{SID6}")

        def _apply_filters(df, yy, mm, cc):
            d = df.copy()
            if yy: d = d[d["_Année_"].isin(yy)]
            if mm: d = d[d["Mois"].astype(str).isin(mm)]
            if cc: d = d[d["Categorie"].astype(str).isin(cc)]
            return d

        A = _apply_filters(dfA0, ya, ma, ca)
        B = _apply_filters(dfA0, yb, mb, cb)

        def _kpis(df):
            h = _safe_num_series(df, "Montant honoraires (US $)")
            a = _safe_num_series(df, "Autres frais (US $)")
            t = (h + a)
            p = _safe_num_series(df, "Payé")
            r = (t - p).clip(lower=0)
            return {
                "Dossiers": len(df),
                "Honoraires+Frais": float(t.sum()),
                "Payé": float(p.sum()),
                "Solde": float(r.sum())
            }

        kA = _kpis(A); kB = _kpis(B)

        cA, cB = st.columns(2)
        with cA:
            st.markdown("**Période A**")
            st.metric("Dossiers", f"{kA['Dossiers']}")
            st.metric("Honoraires+Frais", _fmt_money(kA["Honoraires+Frais"]))
            st.metric("Payé", _fmt_money(kA["Payé"]))
            st.metric("Solde", _fmt_money(kA["Solde"]))
        with cB:
            st.markdown("**Période B**")
            st.metric("Dossiers", f"{kB['Dossiers']}")
            st.metric("Honoraires+Frais", _fmt_money(kB["Honoraires+Frais"]))
            st.metric("Payé", _fmt_money(kB["Payé"]))
            st.metric("Solde", _fmt_money(kB["Solde"]))

        # Détails des dossiers filtrés
        st.markdown("#### 📋 Détails (après filtres globaux)")
        det = dfA.copy()
        # formats lisibles
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Total (US $)", "Payé", "Solde"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).apply(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det

        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"a_detail_{SID6}")


# =======================================================
# 💾 ONGLET : Export (Clients + Visa)
# =======================================================
with tabs[6]:
    st.subheader("💾 Export")

    colz1, colz2 = st.columns([1,3])
    with colz1:
        if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID6}"):
            try:
                buf = BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    # Clients (fichier courant nettoyé)
                    try:
                        df_export = read_clients_file(clients_path_curr)
                        with BytesIO() as xbuf:
                            with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                                df_export.to_excel(wr, sheet_name="Clients", index=False)
                            zf.writestr("Clients.xlsx", xbuf.getvalue())
                    except Exception as e:
                        st.warning(f"Clients : export partiel ({_safe_str(e)})")

                    # Visa (reprendre tel quel si possible)
                    try:
                        zf.write(visa_path_curr, "Visa.xlsx")
                    except Exception:
                        try:
                            dfv0 = pd.read_excel(visa_path_curr)
                            with BytesIO() as vb:
                                with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                    dfv0.to_excel(wr, sheet_name="Visa", index=False)
                                zf.writestr("Visa.xlsx", vb.getvalue())
                        except Exception as e2:
                            st.warning(f"Visa : export partiel ({_safe_str(e2)})")

                st.session_state[f"zip_export_{SID6}"] = buf.getvalue()
                st.success("Archive prête.")
            except Exception as e:
                st.error("Erreur de préparation : " + _safe_str(e))

    with colz2:
        if st.session_state.get(f"zip_export_{SID6}"):
            st.download_button(
                label="⬇️ Télécharger l’export (ZIP)",
                data=st.session_state[f"zip_export_{SID6}"],
                file_name="Export_Visa_Manager.zip",
                mime="application/zip",
                key=f"zip_dl_{SID6}",
            )
        else:
            st.caption("Clique sur « Préparer l’archive ZIP » pour générer un export complet.")