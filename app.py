 # ================================
# PARTIE 1/6 — Imports & Utils
# ================================
from __future__ import annotations

import json
import zipfile
from io import BytesIO
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import numpy as np
import streamlit as st

# -----------------------------
# Constantes
# -----------------------------
APP_TITLE = "🛂 Visa Manager"
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"
WORK_DIR = Path("./upload")
WORK_DIR.mkdir(parents=True, exist_ok=True)

# Persistance des derniers chemins
LAST_PATHS_FILE = Path(".visa_manager_last.json")

# Préfixe isolant les clés Streamlit (évite collisions)
if "_sid_prefix" not in st.session_state:
    st.session_state["_sid_prefix"] = "sid"
SID = st.session_state["_sid_prefix"]


# -----------------------------
# Helpers génériques
# -----------------------------
def _fmt_money(x: float | int) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _safe_str(x) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _to_num(x) -> float:
    try:
        return float(x)
    except Exception:
        return 0.0

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df))
    s = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    return s

def _date_for_widget(val):
    """
    Donne à Streamlit.date_input() une date sûre :
    - None / NaT / vide -> date.today()
    - datetime -> .date()
    - str -> tentative parse
    """
    try:
        if isinstance(val, date) and not isinstance(val, datetime):
            return val
        if isinstance(val, datetime):
            return val.date()
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return date.today()
        return d.date()
    except Exception:
        return date.today()


# -----------------------------
# Persistance "dernier fichier"
# -----------------------------
def _load_last_paths() -> dict:
    try:
        if LAST_PATHS_FILE.exists():
            with open(LAST_PATHS_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
                return {
                    "clients": d.get("clients", ""),
                    "visa": d.get("visa", ""),
                }
    except Exception:
        pass
    return {"clients": "", "visa": ""}

def _save_last_paths(clients_path: str | None = None, visa_path: str | None = None) -> None:
    try:
        d = _load_last_paths()
        if clients_path:
            d["clients"] = str(clients_path)
        if visa_path:
            d["visa"] = str(visa_path)
        with open(LAST_PATHS_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# -----------------------------
# Normalisation colonnes Clients
# -----------------------------
CLIENT_COLS_ORDER = [
    "ID_Client", "Dossier N", "Nom", "Date", "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)", "Total (US $)",
    "Payé", "Solde", "Acompte 1", "Acompte 2",
    "Dossier envoyé", "Dossier approuvé", "Dossier refusé", "Dossier annulé", "RFE",
    "Commentaires", "Mois", "_Année_", "_MoisNum_"
]

ALT_STATUS_NAMES = {
    "Dossiers envoyé": "Dossier envoyé",
    "Dossier envoye": "Dossier envoyé",
    "Dossier approuve": "Dossier approuvé",
    "Dossier refuse": "Dossier refusé",
    "Dossier Annulé": "Dossier annulé",
    "Dossier Annule": "Dossier annulé",
}

def _normalize_client_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Harmoniser noms de colonnes statut
    for alt, std in ALT_STATUS_NAMES.items():
        if alt in df.columns and std not in df.columns:
            df.rename(columns={alt: std}, inplace=True)

    # Créer colonnes manquantes
    for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Acompte 1", "Acompte 2",
              "Dossier envoyé", "Dossier approuvé", "Dossier refusé", "Dossier annulé", "RFE", "Commentaires",
              "Categories", "Sous-categorie", "Visa"]:
        if c not in df.columns:
            df[c] = 0 if c not in ["Commentaires","Categories","Sous-categorie","Visa"] else ""

    # Date + Annee + Mois
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    else:
        df["Date"] = pd.NaT

    df["_Année_"]   = df["Date"].dt.year
    df["_MoisNum_"] = df["Date"].dt.month

    if "Mois" not in df.columns:
        df["Mois"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) else "")

    # Numériques
    for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Acompte 1", "Acompte 2",
              "Dossier envoyé", "Dossier approuvé", "Dossier refusé", "Dossier annulé", "RFE"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    # Total/Payé/Solde robustes
    if "Total (US $)" not in df.columns:
        df["Total (US $)"] = 0.0
    df["Total (US $)"] = (
        _safe_num_series(df, "Montant honoraires (US $)") +
        _safe_num_series(df, "Autres frais (US $)")
    )
    # Si Payé vide, utiliser acomptes
    paid = _safe_num_series(df, "Payé")
    use_acompte = _safe_num_series(df, "Acompte 1") + _safe_num_series(df, "Acompte 2")
    df["Payé"] = np.where(paid > 0, paid, use_acompte)

    # Solde recalculé
    df["Solde"] = (df["Total (US $)"] - df["Payé"]).clip(lower=0)

    # Ordonner colonnes
    cols = [c for c in CLIENT_COLS_ORDER if c in df.columns] + \
           [c for c in df.columns if c not in CLIENT_COLS_ORDER]
    df = df[cols]
    return df


# ===============================================
# PARTIE 2/6 — Sidebar chargement & lecture
# ===============================================
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.sidebar.header("📂 Fichiers")

# Choix du mode
mode = st.sidebar.radio(
    "Mode de chargement",
    ["Un fichier (Clients+Visa, 2 onglets)", "Deux fichiers (Clients & Visa)"],
    horizontal=False,
    key=f"load_mode_{SID}",
)

clients_path_curr: str | None = None
visa_path_curr: str | None = None

_last = _load_last_paths()

if mode == "Un fichier (Clients+Visa, 2 onglets)":
    one_file = st.sidebar.file_uploader("Fichier (xlsx/csv) avec 2 onglets", type=["xlsx", "csv"], key=f"up_one_{SID}")
    if one_file is not None:
        p = WORK_DIR / f"upload_{one_file.name}"
        with open(p, "wb") as f:
            f.write(one_file.getbuffer())
        clients_path_curr = str(p)
        visa_path_curr    = str(p)
        _save_last_paths(clients_path=clients_path_curr, visa_path=visa_path_curr)
else:
    up_clients = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx","csv"], key=f"up_clients_{SID}")
    if up_clients is not None:
        p = WORK_DIR / f"upload_{up_clients.name}"
        with open(p, "wb") as f:
            f.write(up_clients.getbuffer())
        clients_path_curr = str(p)
        _save_last_paths(clients_path=clients_path_curr)

    up_visa = st.sidebar.file_uploader("Visa (xlsx/csv)", type=["xlsx","csv"], key=f"up_visa_{SID}")
    if up_visa is not None:
        p = WORK_DIR / f"upload_{up_visa.name}"
        with open(p, "wb") as f:
            f.write(up_visa.getbuffer())
        visa_path_curr = str(p)
        _save_last_paths(visa_path=visa_path_curr)

# Réutiliser derniers chemins si rien d’uploadé
if not clients_path_curr and _last.get("clients") and Path(_last["clients"]).exists():
    clients_path_curr = _last["clients"]
if not visa_path_curr and _last.get("visa") and Path(_last["visa"]).exists():
    visa_path_curr = _last["visa"]

@st.cache_data(show_spinner=False)
def read_clients_file(path: str) -> pd.DataFrame:
    if not path:
        return pd.DataFrame()
    p = Path(path)
    if not p.exists():
        return pd.DataFrame()
    if p.suffix.lower() == ".csv":
        df = pd.read_csv(p)
    else:
        try:
            xl = pd.ExcelFile(p)
            sheet = SHEET_CLIENTS if SHEET_CLIENTS in xl.sheet_names else xl.sheet_names[0]
            df = pd.read_excel(p, sheet_name=sheet)
        except Exception:
            df = pd.read_excel(p)
    return _normalize_client_df(df)

@st.cache_data(show_spinner=False)
def read_visa_file(path: str) -> pd.DataFrame:
    if not path:
        return pd.DataFrame()
    p = Path(path)
    if not p.exists():
        return pd.DataFrame()
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p)
    try:
        xl = pd.ExcelFile(p)
        sheet = SHEET_VISA if SHEET_VISA in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(p, sheet_name=sheet)
    except Exception:
        df = pd.read_excel(p)
    return df

# Lecture effective
df_all = read_clients_file(clients_path_curr) if clients_path_curr else pd.DataFrame()
df_visa_raw = read_visa_file(visa_path_curr) if visa_path_curr else pd.DataFrame()

with st.expander("📄 Fichiers chargés", expanded=True):
    st.write("**Clients** :", f"`{clients_path_curr}`" if clients_path_curr else "_non chargé_")
    st.write("**Visa** :", f"`{visa_path_curr}`" if visa_path_curr else "_non chargé_")



# ===============================================
# PARTIE 3/6 — 📊 Dashboard
# ===============================================
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Compte client", "🧾 Gestion", "📄 Visa (aperçu)", "💾 Export"])

with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # KPI (réduits)
        c1, c2, c3, c4, c5 = st.columns([1,1,1,1,1])
        c1.metric("Dossiers", f"{len(df_all)}")
        c2.metric("Honoraires+Frais", _fmt_money(float((_safe_num_series(df_all,"Montant honoraires (US $)") + _safe_num_series(df_all,"Autres frais (US $)")).sum())))
        c3.metric("Payé", _fmt_money(float(_safe_num_series(df_all,"Payé").sum())))
        c4.metric("Solde", _fmt_money(float(_safe_num_series(df_all,"Solde").sum())))
        pct_env = 0.0
        if "Dossier envoyé" in df_all.columns and len(df_all)>0:
            pct_env = 100.0 * (df_all["Dossier envoyé"].astype(float) > 0).sum() / len(df_all)
        c5.metric("Envoyés (%)", f"{pct_env:.0f}%")

        # Filtres
        st.markdown("#### 🎛️ Filtres")
        cats  = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subs  = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3 = st.columns(3)
        fc = a1.multiselect("Catégories", cats, default=[], key=f"dash_c_{SID}")
        fs = a2.multiselect("Sous-catégories", subs, default=[], key=f"dash_s_{SID}")
        fv = a3.multiselect("Visa", visas, default=[], key=f"dash_v_{SID}")

        view = df_all.copy()
        if fc: view = view[view["Categories"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        # Graph simple : Dossiers par catégorie
        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if not view.empty and "Categories" in view.columns:
            vc = view["Categories"].value_counts().reset_index()
            vc.columns = ["Categories","Nombre"]
            st.bar_chart(vc.set_index("Categories"))
        else:
            st.write("—")

        # Flux par mois (honoraires/frais/payé/solde)
        st.markdown("#### 💵 Flux par mois")
        if not view.empty and "Mois" in view.columns:
            tmp = view.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gp = tmp.groupby("Mois", as_index=False)[["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]].sum().sort_values("Mois")
            st.line_chart(gp.set_index("Mois"))
        else:
            st.write("—")

        # Détails
        st.markdown("#### 📋 Détails (après filtres)")
        detail = view.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde"]:
            if c in detail.columns:
                detail[c] = _safe_num_series(detail, c).map(_fmt_money)
        if "Date" in detail.columns:
            try:
                detail["Date"] = pd.to_datetime(detail["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                detail["Date"] = detail["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categories","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in detail.columns]

        sort_keys = [c for c in ["_Année_","_MoisNum_","Categories","Nom"] if c in detail.columns]
        detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail

        st.dataframe(detail_sorted[show_cols].reset_index(drop=True), use_container_width=True, height=400)



# ===============================================
# PARTIE 4/6 — 📈 Analyses & 🏦 Escrow
# ===============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_years_{SID}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"a_months_{SID}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"a_cats_{SID}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"a_subs_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_visas_{SID}")

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categories"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPI
        k1, k2, k3, k4 = st.columns([1,1,1,1])
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(dfA,"Montant honoraires (US $)").sum()))
        k3.metric("Payé", _fmt_money(_safe_num_series(dfA,"Payé").sum()))
        k4.metric("Solde", _fmt_money(_safe_num_series(dfA,"Solde").sum()))

        # % par catégorie / sous-catégorie
        st.markdown("#### 📊 Répartition (%)")
        cA, cB = st.columns(2)
        if not dfA.empty and "Categories" in dfA.columns:
            aggC = dfA.groupby("Categories", as_index=False)["Total (US $)"].sum()
            total = float(aggC["Total (US $)"].sum()) or 1.0
            aggC["%"] = (aggC["Total (US $)"] / total * 100).round(1)
            cA.dataframe(aggC.sort_values("%", ascending=False), use_container_width=True, height=240)
        if not dfA.empty and "Sous-categorie" in dfA.columns:
            aggS = dfA.groupby("Sous-categorie", as_index=False)["Total (US $)"].sum()
            total = float(aggS["Total (US $)"].sum()) or 1.0
            aggS["%"] = (aggS["Total (US $)"] / total * 100).round(1)
            cB.dataframe(aggS.sort_values("%", ascending=False), use_container_width=True, height=240)

        # Comparaison période A vs B
        st.markdown("#### 🔁 Comparaison période A vs B (Année/Mois)")
        ca1, ca2, cb1, cb2 = st.columns(4)
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=f"cmp_ya_{SID}")
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=f"cmp_ma_{SID}")
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=f"cmp_yb_{SID}")
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=f"cmp_mb_{SID}")

        def _subset(y, m):
            d = dfA.copy()
            if y: d = d[d["_Année_"].isin(y)]
            if m: d = d[d["Mois"].astype(str).isin(m)]
            return d

        A = _subset(pa_years, pa_month)
        B = _subset(pb_years, pb_month)

        ca, cb, cc = st.columns(3)
        ca.metric("A - Total (US $)", _fmt_money(float(_safe_num_series(A,"Total (US $)").sum())))
        cb.metric("B - Total (US $)", _fmt_money(float(_safe_num_series(B,"Total (US $)").sum())))
        cc.metric("Δ (B-A)", _fmt_money(float(_safe_num_series(B,"Total (US $)").sum() - _safe_num_series(A,"Total (US $)").sum())))

        st.markdown("#### 🧾 Détails filtrés")
        det = dfA.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).map(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        cols_show = [c for c in [
            "Dossier N","ID_Client","Nom","Categories","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]
        sort_keys = [c for c in ["_Année_","_MoisNum_","Categories","Nom"] if c in det.columns]
        det = det.sort_values(by=sort_keys) if sort_keys else det
        st.dataframe(det[cols_show].reset_index(drop=True), use_container_width=True, height=360)

# ------------- Escrow -------------
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Total (US $)"] = _safe_num_series(dfE,"Total (US $)")
        dfE["Payé"] = _safe_num_series(dfE,"Payé")
        dfE["Solde"] = _safe_num_series(dfE,"Solde")

        t1, t2, t3 = st.columns([1,1,1])
        t1.metric("Total (US $)", _fmt_money(float(dfE["Total (US $)"].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Reste", _fmt_money(float(dfE["Solde"].sum())))

        agg = dfE.groupby("Categories", as_index=False)[["Total (US $)","Payé","Solde"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg["Total (US $)"]).replace([pd.NA, pd.NaT, np.inf, -np.inf], 0).fillna(0.0) * 100
        st.dataframe(agg.sort_values("Total (US $)", ascending=False), use_container_width=True, height=380)
        st.caption("NB : on peut spécialiser l’Escrow pour suivre les honoraires encaissés avant envoi et signaler les transferts à faire quand « Dossier envoyé » est coché.")

## ================================
# PARTIE 1/6 — Imports & Utils
# ================================
from __future__ import annotations

import json
import zipfile
from io import BytesIO
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import numpy as np
import streamlit as st

# -----------------------------
# Constantes
# -----------------------------
APP_TITLE = "🛂 Visa Manager"
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"
WORK_DIR = Path("./upload")
WORK_DIR.mkdir(parents=True, exist_ok=True)

# Persistance des derniers chemins
LAST_PATHS_FILE = Path(".visa_manager_last.json")

# Préfixe isolant les clés Streamlit (évite collisions)
if "_sid_prefix" not in st.session_state:
    st.session_state["_sid_prefix"] = "sid"
SID = st.session_state["_sid_prefix"]


# -----------------------------
# Helpers génériques
# -----------------------------
def _fmt_money(x: float | int) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _safe_str(x) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _to_num(x) -> float:
    try:
        return float(x)
    except Exception:
        return 0.0

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df))
    s = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    return s

def _date_for_widget(val):
    """
    Donne à Streamlit.date_input() une date sûre :
    - None / NaT / vide -> date.today()
    - datetime -> .date()
    - str -> tentative parse
    """
    try:
        if isinstance(val, date) and not isinstance(val, datetime):
            return val
        if isinstance(val, datetime):
            return val.date()
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return date.today()
        return d.date()
    except Exception:
        return date.today()


# -----------------------------
# Persistance "dernier fichier"
# -----------------------------
def _load_last_paths() -> dict:
    try:
        if LAST_PATHS_FILE.exists():
            with open(LAST_PATHS_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
                return {
                    "clients": d.get("clients", ""),
                    "visa": d.get("visa", ""),
                }
    except Exception:
        pass
    return {"clients": "", "visa": ""}

def _save_last_paths(clients_path: str | None = None, visa_path: str | None = None) -> None:
    try:
        d = _load_last_paths()
        if clients_path:
            d["clients"] = str(clients_path)
        if visa_path:
            d["visa"] = str(visa_path)
        with open(LAST_PATHS_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# -----------------------------
# Normalisation colonnes Clients
# -----------------------------
CLIENT_COLS_ORDER = [
    "ID_Client", "Dossier N", "Nom", "Date", "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)", "Total (US $)",
    "Payé", "Solde", "Acompte 1", "Acompte 2",
    "Dossier envoyé", "Dossier approuvé", "Dossier refusé", "Dossier annulé", "RFE",
    "Commentaires", "Mois", "_Année_", "_MoisNum_"
]

ALT_STATUS_NAMES = {
    "Dossiers envoyé": "Dossier envoyé",
    "Dossier envoye": "Dossier envoyé",
    "Dossier approuve": "Dossier approuvé",
    "Dossier refuse": "Dossier refusé",
    "Dossier Annulé": "Dossier annulé",
    "Dossier Annule": "Dossier annulé",
}

def _normalize_client_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Harmoniser noms de colonnes statut
    for alt, std in ALT_STATUS_NAMES.items():
        if alt in df.columns and std not in df.columns:
            df.rename(columns={alt: std}, inplace=True)

    # Créer colonnes manquantes
    for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Acompte 1", "Acompte 2",
              "Dossier envoyé", "Dossier approuvé", "Dossier refusé", "Dossier annulé", "RFE", "Commentaires",
              "Categories", "Sous-categorie", "Visa"]:
        if c not in df.columns:
            df[c] = 0 if c not in ["Commentaires","Categories","Sous-categorie","Visa"] else ""

    # Date + Annee + Mois
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    else:
        df["Date"] = pd.NaT

    df["_Année_"]   = df["Date"].dt.year
    df["_MoisNum_"] = df["Date"].dt.month

    if "Mois" not in df.columns:
        df["Mois"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) else "")

    # Numériques
    for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Acompte 1", "Acompte 2",
              "Dossier envoyé", "Dossier approuvé", "Dossier refusé", "Dossier annulé", "RFE"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    # Total/Payé/Solde robustes
    if "Total (US $)" not in df.columns:
        df["Total (US $)"] = 0.0
    df["Total (US $)"] = (
        _safe_num_series(df, "Montant honoraires (US $)") +
        _safe_num_series(df, "Autres frais (US $)")
    )
    # Si Payé vide, utiliser acomptes
    paid = _safe_num_series(df, "Payé")
    use_acompte = _safe_num_series(df, "Acompte 1") + _safe_num_series(df, "Acompte 2")
    df["Payé"] = np.where(paid > 0, paid, use_acompte)

    # Solde recalculé
    df["Solde"] = (df["Total (US $)"] - df["Payé"]).clip(lower=0)

    # Ordonner colonnes
    cols = [c for c in CLIENT_COLS_ORDER if c in df.columns] + \
           [c for c in df.columns if c not in CLIENT_COLS_ORDER]
    df = df[cols]
    return df


# ===============================================
# PARTIE 2/6 — Sidebar chargement & lecture
# ===============================================
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.sidebar.header("📂 Fichiers")

# Choix du mode
mode = st.sidebar.radio(
    "Mode de chargement",
    ["Un fichier (Clients+Visa, 2 onglets)", "Deux fichiers (Clients & Visa)"],
    horizontal=False,
    key=f"load_mode_{SID}",
)

clients_path_curr: str | None = None
visa_path_curr: str | None = None

_last = _load_last_paths()

if mode == "Un fichier (Clients+Visa, 2 onglets)":
    one_file = st.sidebar.file_uploader("Fichier (xlsx/csv) avec 2 onglets", type=["xlsx", "csv"], key=f"up_one_{SID}")
    if one_file is not None:
        p = WORK_DIR / f"upload_{one_file.name}"
        with open(p, "wb") as f:
            f.write(one_file.getbuffer())
        clients_path_curr = str(p)
        visa_path_curr    = str(p)
        _save_last_paths(clients_path=clients_path_curr, visa_path=visa_path_curr)
else:
    up_clients = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx","csv"], key=f"up_clients_{SID}")
    if up_clients is not None:
        p = WORK_DIR / f"upload_{up_clients.name}"
        with open(p, "wb") as f:
            f.write(up_clients.getbuffer())
        clients_path_curr = str(p)
        _save_last_paths(clients_path=clients_path_curr)

    up_visa = st.sidebar.file_uploader("Visa (xlsx/csv)", type=["xlsx","csv"], key=f"up_visa_{SID}")
    if up_visa is not None:
        p = WORK_DIR / f"upload_{up_visa.name}"
        with open(p, "wb") as f:
            f.write(up_visa.getbuffer())
        visa_path_curr = str(p)
        _save_last_paths(visa_path=visa_path_curr)

# Réutiliser derniers chemins si rien d’uploadé
if not clients_path_curr and _last.get("clients") and Path(_last["clients"]).exists():
    clients_path_curr = _last["clients"]
if not visa_path_curr and _last.get("visa") and Path(_last["visa"]).exists():
    visa_path_curr = _last["visa"]

@st.cache_data(show_spinner=False)
def read_clients_file(path: str) -> pd.DataFrame:
    if not path:
        return pd.DataFrame()
    p = Path(path)
    if not p.exists():
        return pd.DataFrame()
    if p.suffix.lower() == ".csv":
        df = pd.read_csv(p)
    else:
        try:
            xl = pd.ExcelFile(p)
            sheet = SHEET_CLIENTS if SHEET_CLIENTS in xl.sheet_names else xl.sheet_names[0]
            df = pd.read_excel(p, sheet_name=sheet)
        except Exception:
            df = pd.read_excel(p)
    return _normalize_client_df(df)

@st.cache_data(show_spinner=False)
def read_visa_file(path: str) -> pd.DataFrame:
    if not path:
        return pd.DataFrame()
    p = Path(path)
    if not p.exists():
        return pd.DataFrame()
    if p.suffix.lower() == ".csv":
        return pd.read_csv(p)
    try:
        xl = pd.ExcelFile(p)
        sheet = SHEET_VISA if SHEET_VISA in xl.sheet_names else xl.sheet_names[0]
        df = pd.read_excel(p, sheet_name=sheet)
    except Exception:
        df = pd.read_excel(p)
    return df

# Lecture effective
df_all = read_clients_file(clients_path_curr) if clients_path_curr else pd.DataFrame()
df_visa_raw = read_visa_file(visa_path_curr) if visa_path_curr else pd.DataFrame()

with st.expander("📄 Fichiers chargés", expanded=True):
    st.write("**Clients** :", f"`{clients_path_curr}`" if clients_path_curr else "_non chargé_")
    st.write("**Visa** :", f"`{visa_path_curr}`" if visa_path_curr else "_non chargé_")



# ===============================================
# PARTIE 3/6 — 📊 Dashboard
# ===============================================
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Compte client", "🧾 Gestion", "📄 Visa (aperçu)", "💾 Export"])

with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # KPI (réduits)
        c1, c2, c3, c4, c5 = st.columns([1,1,1,1,1])
        c1.metric("Dossiers", f"{len(df_all)}")
        c2.metric("Honoraires+Frais", _fmt_money(float((_safe_num_series(df_all,"Montant honoraires (US $)") + _safe_num_series(df_all,"Autres frais (US $)")).sum())))
        c3.metric("Payé", _fmt_money(float(_safe_num_series(df_all,"Payé").sum())))
        c4.metric("Solde", _fmt_money(float(_safe_num_series(df_all,"Solde").sum())))
        pct_env = 0.0
        if "Dossier envoyé" in df_all.columns and len(df_all)>0:
            pct_env = 100.0 * (df_all["Dossier envoyé"].astype(float) > 0).sum() / len(df_all)
        c5.metric("Envoyés (%)", f"{pct_env:.0f}%")

        # Filtres
        st.markdown("#### 🎛️ Filtres")
        cats  = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subs  = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3 = st.columns(3)
        fc = a1.multiselect("Catégories", cats, default=[], key=f"dash_c_{SID}")
        fs = a2.multiselect("Sous-catégories", subs, default=[], key=f"dash_s_{SID}")
        fv = a3.multiselect("Visa", visas, default=[], key=f"dash_v_{SID}")

        view = df_all.copy()
        if fc: view = view[view["Categories"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        # Graph simple : Dossiers par catégorie
        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if not view.empty and "Categories" in view.columns:
            vc = view["Categories"].value_counts().reset_index()
            vc.columns = ["Categories","Nombre"]
            st.bar_chart(vc.set_index("Categories"))
        else:
            st.write("—")

        # Flux par mois (honoraires/frais/payé/solde)
        st.markdown("#### 💵 Flux par mois")
        if not view.empty and "Mois" in view.columns:
            tmp = view.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gp = tmp.groupby("Mois", as_index=False)[["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]].sum().sort_values("Mois")
            st.line_chart(gp.set_index("Mois"))
        else:
            st.write("—")

        # Détails
        st.markdown("#### 📋 Détails (après filtres)")
        detail = view.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde"]:
            if c in detail.columns:
                detail[c] = _safe_num_series(detail, c).map(_fmt_money)
        if "Date" in detail.columns:
            try:
                detail["Date"] = pd.to_datetime(detail["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                detail["Date"] = detail["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categories","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in detail.columns]

        sort_keys = [c for c in ["_Année_","_MoisNum_","Categories","Nom"] if c in detail.columns]
        detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail

        st.dataframe(detail_sorted[show_cols].reset_index(drop=True), use_container_width=True, height=400)



# ===============================================
# PARTIE 4/6 — 📈 Analyses & 🏦 Escrow
# ===============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_years_{SID}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"a_months_{SID}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"a_cats_{SID}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"a_subs_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_visas_{SID}")

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categories"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPI
        k1, k2, k3, k4 = st.columns([1,1,1,1])
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(dfA,"Montant honoraires (US $)").sum()))
        k3.metric("Payé", _fmt_money(_safe_num_series(dfA,"Payé").sum()))
        k4.metric("Solde", _fmt_money(_safe_num_series(dfA,"Solde").sum()))

        # % par catégorie / sous-catégorie
        st.markdown("#### 📊 Répartition (%)")
        cA, cB = st.columns(2)
        if not dfA.empty and "Categories" in dfA.columns:
            aggC = dfA.groupby("Categories", as_index=False)["Total (US $)"].sum()
            total = float(aggC["Total (US $)"].sum()) or 1.0
            aggC["%"] = (aggC["Total (US $)"] / total * 100).round(1)
            cA.dataframe(aggC.sort_values("%", ascending=False), use_container_width=True, height=240)
        if not dfA.empty and "Sous-categorie" in dfA.columns:
            aggS = dfA.groupby("Sous-categorie", as_index=False)["Total (US $)"].sum()
            total = float(aggS["Total (US $)"].sum()) or 1.0
            aggS["%"] = (aggS["Total (US $)"] / total * 100).round(1)
            cB.dataframe(aggS.sort_values("%", ascending=False), use_container_width=True, height=240)

        # Comparaison période A vs B
        st.markdown("#### 🔁 Comparaison période A vs B (Année/Mois)")
        ca1, ca2, cb1, cb2 = st.columns(4)
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=f"cmp_ya_{SID}")
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=f"cmp_ma_{SID}")
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=f"cmp_yb_{SID}")
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=f"cmp_mb_{SID}")

        def _subset(y, m):
            d = dfA.copy()
            if y: d = d[d["_Année_"].isin(y)]
            if m: d = d[d["Mois"].astype(str).isin(m)]
            return d

        A = _subset(pa_years, pa_month)
        B = _subset(pb_years, pb_month)

        ca, cb, cc = st.columns(3)
        ca.metric("A - Total (US $)", _fmt_money(float(_safe_num_series(A,"Total (US $)").sum())))
        cb.metric("B - Total (US $)", _fmt_money(float(_safe_num_series(B,"Total (US $)").sum())))
        cc.metric("Δ (B-A)", _fmt_money(float(_safe_num_series(B,"Total (US $)").sum() - _safe_num_series(A,"Total (US $)").sum())))

        st.markdown("#### 🧾 Détails filtrés")
        det = dfA.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).map(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        cols_show = [c for c in [
            "Dossier N","ID_Client","Nom","Categories","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]
        sort_keys = [c for c in ["_Année_","_MoisNum_","Categories","Nom"] if c in det.columns]
        det = det.sort_values(by=sort_keys) if sort_keys else det
        st.dataframe(det[cols_show].reset_index(drop=True), use_container_width=True, height=360)

# ------------- Escrow -------------
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Total (US $)"] = _safe_num_series(dfE,"Total (US $)")
        dfE["Payé"] = _safe_num_series(dfE,"Payé")
        dfE["Solde"] = _safe_num_series(dfE,"Solde")

        t1, t2, t3 = st.columns([1,1,1])
        t1.metric("Total (US $)", _fmt_money(float(dfE["Total (US $)"].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Reste", _fmt_money(float(dfE["Solde"].sum())))

        agg = dfE.groupby("Categories", as_index=False)[["Total (US $)","Payé","Solde"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg["Total (US $)"]).replace([pd.NA, pd.NaT, np.inf, -np.inf], 0).fillna(0.0) * 100
        st.dataframe(agg.sort_values("Total (US $)", ascending=False), use_container_width=True, height=380)
        st.caption("NB : on peut spécialiser l’Escrow pour suivre les honoraires encaissés avant envoi et signaler les transferts à faire quand « Dossier envoyé » est coché.")



# =======================================================
# PARTIE 5/6 — 👤 Compte client & 🧾 Gestion (CRUD)
# =======================================================

# ---------- COMPTE CLIENT ----------
with tabs[3]:
    st.subheader("👤 Compte client — suivi du dossier")
    if df_all.empty:
        st.info("Aucun client chargé.")
    else:
        ids = df_all["ID_Client"].dropna().astype(str).unique().tolist() if "ID_Client" in df_all.columns else []
        sel = st.selectbox("Choisir un client", [""] + sorted(ids), index=0, key=f"acct_sel_{SID}")
        if sel:
            row = df_all[df_all["ID_Client"].astype(str) == sel].iloc[0].copy()
            st.markdown(f"**Nom :** {_safe_str(row.get('Nom',''))}  |  **Visa :** {_safe_str(row.get('Visa',''))}")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Honoraires+Frais", _fmt_money(float(_to_num(row.get("Total (US $)", 0.0)))))
            c2.metric("Payé", _fmt_money(float(_to_num(row.get("Payé", 0.0)))))
            c3.metric("Solde", _fmt_money(float(_to_num(row.get("Solde", 0.0)))))
            sent = int(_to_num(row.get("Dossier envoyé", 0)) or 0)
            c4.metric("Envoyé", "Oui" if sent == 1 else "Non")

            # ---- Chronologie ----
            st.markdown("#### Chronologie")
            s1, s2 = st.columns(2)
            s1.write(f"- Date création : {_safe_str(row.get('Date',''))}")
            s1.write(
                f"- Dossier envoyé : {int(_to_num(row.get('Dossier envoyé',0)) or 0)}  "
                f"| Date : {_safe_str(row.get(\"Date d'envoi\",\"\"))}"
            )
            s1.write(
                f"- Dossier approuvé : {int(_to_num(row.get('Dossier approuvé',0)) or 0)}  "
                f"| Date : {_safe_str(row.get(\"Date d'acceptation\",\"\"))}"
            )
            s2.write(
                f"- Dossier refusé : {int(_to_num(row.get('Dossier refusé',0)) or 0)}  "
                f"| Date : {_safe_str(row.get(\"Date de refus\",\"\"))}"
            )
            s2.write(
                f"- Dossier annulé : {int(_to_num(row.get('Dossier annulé',0)) or 0)}  "
                f"| Date : {_safe_str(row.get(\"Date d'annulation\",\"\"))}"
            )
            st.write(f"- RFE : {int(_to_num(row.get('RFE',0)) or 0)}")
            st.write(f"- Commentaires : {_safe_str(row.get('Commentaires',''))}")

# ---------- GESTION (CRUD) ----------
with tabs[4]:
    st.subheader("🧾 Gestion des clients (Ajouter / Modifier / Supprimer)")
    if df_all.empty:
        st.info("Aucun client chargé.")
    else:
        op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_op_{SID}")
        df_live = df_all.copy()

        # --- AJOUTER ---
        if op == "Ajouter":
            st.markdown("### ➕ Ajouter un client")
            d1, d2, d3 = st.columns(3)
            nom  = d1.text_input("Nom", "", key=f"add_nom_{SID}")
            dte  = d2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=date.today().month-1, key=f"add_mois_{SID}")

            f1, f2 = st.columns(2)
            cat = f1.text_input("Catégorie", "", key=f"add_cat_{SID}")
            sub = f2.text_input("Sous-catégorie", "", key=f"add_sub_{SID}")
            vis = st.text_input("Visa", "", key=f"add_visa_{SID}")

            g1, g2 = st.columns(2)
            honor = g1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
            other = g2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")
            comm  = st.text_area("Commentaires (autres frais, remarques…)", "", key=f"add_comm_{SID}")

            st.markdown("#### Statuts")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent   = s1.checkbox("Dossier envoyé", key=f"add_sent_{SID}")
            sent_d = s1.date_input("Date d'envoi", value=None, key=f"add_sentd_{SID}")
            acc    = s2.checkbox("Dossier approuvé", key=f"add_acc_{SID}")
            acc_d  = s2.date_input("Date d'acceptation", value=None, key=f"add_accd_{SID}")
            ref    = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
            ref_d  = s3.date_input("Date de refus", value=None, key=f"add_refd_{SID}")
            ann    = s4.checkbox("Dossier annulé", key=f"add_ann_{SID}")
            ann_d  = s4.date_input("Date d'annulation", value=None, key=f"add_annd_{SID}")
            rfe    = s5.checkbox("RFE", key=f"add_rfe_{SID}")

            if rfe and not any([sent, acc, ref, ann]):
                st.warning("⚠️ RFE doit être lié à un autre statut (envoyé / approuvé / refusé / annulé).")

            if st.button("💾 Enregistrer le client", key=f"btn_add_{SID}"):
                if not nom:
                    st.warning("Le nom du client est requis.")
                    st.stop()
                total = float(honor) + float(other)
                paye  = 0.0
                solde = max(0.0, total - paye)
                base = _safe_str(nom).strip().replace(" ", "_")
                did = f"{base}-{dte.strftime('%Y%m%d')}"
                next_dossier = 13057
                if "Dossier N" in df_live.columns and pd.to_numeric(df_live["Dossier N"], errors="coerce").notna().any():
                    next_dossier = int(pd.to_numeric(df_live["Dossier N"], errors="coerce").max() or 13056) + 1

                new_row = {
                    "ID_Client": did,
                    "Dossier N": next_dossier,
                    "Nom": nom,
                    "Date": dte,
                    "Mois": f"{int(mois):02d}",
                    "Categories": cat,
                    "Sous-categorie": sub,
                    "Visa": vis,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Total (US $)": total,
                    "Payé": paye,
                    "Solde": solde,
                    "Acompte 1": 0.0, "Acompte 2": 0.0,
                    "Dossier envoyé": 1 if sent else 0,
                    "Date d'envoi": sent_d if sent_d else (dte if sent else None),
                    "Dossier approuvé": 1 if acc else 0,
                    "Date d'acceptation": acc_d if acc_d else None,
                    "Dossier refusé": 1 if ref else 0,
                    "Date de refus": ref_d if ref_d else None,
                    "Dossier annulé": 1 if ann else 0,
                    "Date d'annulation": ann_d if ann_d else None,
                    "RFE": 1 if rfe else 0,
                    "Commentaires": comm,
                }
                df_all = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                st.success("✅ Client ajouté (mémoire runtime). Utilisez Export pour sauvegarder sur disque.")

        # --- MODIFIER ---
        elif op == "Modifier":
            st.markdown("### ✏️ Modifier un client")
            names = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            target = st.selectbox("ID_Client", [""] + names, index=0, key=f"mod_id_{SID}")
            if target:
                idx = df_live[df_live["ID_Client"].astype(str) == target].index[0]
                row = df_live.loc[idx].copy()

                d1, d2, d3 = st.columns(3)
                nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nom_{SID}")
                dval = _date_for_widget(row.get("Date"))
                dte  = d2.date_input("Date de création", value=dval, key=f"mod_date_{SID}")
                mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                    index=(int(_safe_str(row.get("Mois","01"))) - 1 if str(row.get("Mois","01")).isdigit() else 0),
                    key=f"mod_mois_{SID}"
                )

                f1, f2, f3 = st.columns(3)
                cat = f1.text_input("Catégorie", _safe_str(row.get("Categories","")), key=f"mod_cat_{SID}")
                sub = f2.text_input("Sous-catégorie", _safe_str(row.get("Sous-categorie","")), key=f"mod_sub_{SID}")
                vis = f3.text_input("Visa", _safe_str(row.get("Visa","")), key=f"mod_visa_{SID}")

                g1, g2, g3 = st.columns(3)
                honor = g1.number_input("Montant honoraires (US $)", min_value=0.0,
                    value=float(_to_num(row.get("Montant honoraires (US $)",0.0))), step=50.0, format="%.2f", key=f"mod_h_{SID}")
                other = g2.number_input("Autres frais (US $)", min_value=0.0,
                    value=float(_to_num(row.get("Autres frais (US $)",0.0))), step=20.0, format="%.2f", key=f"mod_o_{SID}")
                comm  = g3.text_input("Commentaires", _safe_str(row.get("Commentaires","")), key=f"mod_comm_{SID}")

                st.markdown("#### Statuts")
                s1, s2, s3, s4, s5 = st.columns(5)
                sent   = s1.checkbox("Dossier envoyé", value=bool(int(_to_num(row.get("Dossier envoyé",0)) or 0)), key=f"mod_sent_{SID}")
                sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
                acc    = s2.checkbox("Dossier approuvé", value=bool(int(_to_num(row.get("Dossier approuvé",0)) or 0)), key=f"mod_acc_{SID}")
                acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
                ref    = s3.checkbox("Dossier refusé", value=bool(int(_to_num(row.get("Dossier refusé",0)) or 0)), key=f"mod_ref_{SID}")
                ref_d  = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=f"mod_refd_{SID}")
                ann    = s4.checkbox("Dossier annulé", value=bool(int(_to_num(row.get("Dossier annulé",0)) or 0)), key=f"mod_ann_{SID}")
                ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
                rfe    = s5.checkbox("RFE", value=bool(int(_to_num(row.get("RFE",0)) or 0)), key=f"mod_rfe_{SID}")

                if st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}"):
                    total = float(honor) + float(other)
                    paye  = float(_to_num(row.get("Payé",0.0)))
                    solde = max(0.0, total - paye)

                    df_live.at[idx, "Nom"]  = nom
                    df_live.at[idx, "Date"] = dte
                    df_live.at[idx, "Mois"] = f"{int(mois):02d}"
                    df_live.at[idx, "Categories"] = cat
                    df_live.at[idx, "Sous-categorie"] = sub
                    df_live.at[idx, "Visa"] = vis
                    df_live.at[idx, "Montant honoraires (US $)"] = float(honor)
                    df_live.at[idx, "Autres frais (US $)"]        = float(other)
                    df_live.at[idx, "Total (US $)"]               = total
                    df_live.at[idx, "Solde"]                      = solde
                    df_live.at[idx, "Commentaires"]               = comm
                    df_live.at[idx, "Dossier envoyé"]             = 1 if sent else 0
                    df_live.at[idx, "Date d'envoi"]               = sent_d
                    df_live.at[idx, "Dossier approuvé"]           = 1 if acc else 0
                    df_live.at[idx, "Date d'acceptation"]         = acc_d
                    df_live.at[idx, "Dossier refusé"]             = 1 if ref else 0
                    df_live.at[idx, "Date de refus"]              = ref_d
                    df_live.at[idx, "Dossier annulé"]             = 1 if ann else 0
                    df_live.at[idx, "Date d'annulation"]          = ann_d
                    df_live.at[idx, "RFE"]                        = 1 if rfe else 0

                    df_all = df_live.copy()
                    st.success("✅ Modifications appliquées (mémoire runtime). Exporter pour sauvegarder.")

        # --- SUPPRIMER ---
        elif op == "Supprimer":
            st.markdown("### 🗑️ Supprimer un client")
            names = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            target = st.selectbox("ID_Client", [""] + names, index=0, key=f"del_id_{SID}")
            if target:
                mask = df_live["ID_Client"].astype(str) == target
                row = df_live[mask].iloc[0]
                st.write({"Dossier N": row.get("Dossier N",""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                    df_all = df_live[~mask].copy()
                    st.success("🗑️ Client supprimé (mémoire runtime). Exporter pour sauvegarder.")

 # ===============================================
# PARTIE 6/6 — 📄 Visa (aperçu) & 💾 Export
# ===============================================

with tabs[5]:
    st.subheader("📄 Visa (aperçu brut)")
    if df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True, height=420)

with tabs[6]:
    st.subheader("💾 Export (sauvegarde sur ton disque)")
    st.caption("Télécharge un ZIP contenant le fichier Clients normalisé et le fichier Visa tel que chargé.")

    if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID}"):
        try:
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients normalisés
                if not df_all.empty:
                    with BytesIO() as xbuf:
                        with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                            df_all.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                        zf.writestr("Clients.xlsx", xbuf.getvalue())
                # Visa tel quel si dispo
                if isinstance(visa_path_curr, str) and Path(visa_path_curr).exists():
                    try:
                        zf.write(visa_path_curr, "Visa.xlsx")
                    except Exception:
                        try:
                            dfv0 = read_visa_file(visa_path_curr)
                            with BytesIO() as vb:
                                with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                    dfv0.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                                zf.writestr("Visa.xlsx", vb.getvalue())
                        except Exception:
                            pass
            st.session_state[f"zip_export_{SID}"] = buf.getvalue()
            st.success("Archive prête.")
        except Exception as e:
            st.error("Erreur export : " + _safe_str(e))

    if st.session_state.get(f"zip_export_{SID}"):
        st.download_button(
            "⬇️ Télécharger l’export (ZIP)",
            data=st.session_state[f"zip_export_{SID}"],
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
            key=f"zip_dl_{SID}",
        )