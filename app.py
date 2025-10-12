# ==============================================
# 🛂 Visa Manager – App complète (Partie 1/2)
# ==============================================
from __future__ import annotations

import json, os, zipfile, uuid, re
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

# --------- Constantes colonnes ----------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"

STATUS_COLS = [
    "Dossier envoyé", "Date d'envoi",
    "Dossier accepté", "Date d'acceptation",
    "Dossier refusé",  "Date de refus",
    "Dossier annulé",  "Date d'annulation",
    "RFE",
]

# --------- SID pour éviter collisions de clés ---------
SID = st.session_state.get("_sid") or uuid.uuid4().hex[:8]
st.session_state["_sid"] = SID

st.set_page_config(page_title="Visa Manager", page_icon="🛂", layout="wide")

# ==============================================
# Utils de conversion/formatage
# ==============================================
def _norm(s: str) -> str:
    s = str(s or "")
    s = s.strip().lower()
    s = re.sub(r"[’'`´]", "", s)
    s = re.sub(r"[^a-z0-9+/_\- ]+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s

def _safe_str(v) -> str:
    try:
        return "" if v is None else str(v)
    except Exception:
        return ""

def _to_float(v) -> float:
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = _safe_str(v)
    s = s.replace(" ", "").replace("\u00A0","")
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") >= 1:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df), dtype=float)
    return df[col].apply(_to_float)

def _fmt_money(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(v):
    """Retourne None ou date (pas Timestamp NaT) pour les date_input."""
    if v is None or (isinstance(v, float) and pd.isna(v)): 
        return None
    if isinstance(v, date): return v
    if isinstance(v, datetime): return v.date()
    try:
        d = pd.to_datetime(v, errors="coerce")
        return None if pd.isna(d) else d.date()
    except Exception:
        return None

def _month_index(mois_val) -> int:
    s = _safe_str(mois_val)
    if s.isdigit():
        i = int(s)
        if 1 <= i <= 12: return i-1
    return 0

def _best_index(options: List[str], current: str) -> int:
    cur = _norm(current)
    for i, opt in enumerate(options, start=1):
        if _norm(opt) == cur:
            return i
    return 0

# ==============================================
# Lecture / Écriture fichiers
# ==============================================
WORK_DIR = os.getcwd()
CLIENTS_DEFAULT = os.path.join(WORK_DIR, "donnees_visa_clients1_adapte.xlsx")
VISA_DEFAULT    = os.path.join(WORK_DIR, "donnees_visa_clients1.xlsx")

def _save_bytes_to(path: str, data: bytes) -> None:
    with open(path, "wb") as f:
        f.write(data)

@st.cache_data(show_spinner=False)
def read_clients_file(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        df = pd.read_excel(path)  # premier onglet
    return df

@st.cache_data(show_spinner=False)
def read_visa_file(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=SHEET_VISA)
    except Exception:
        df = pd.read_excel(path)
    return df

def write_clients_file(df: pd.DataFrame, path: str) -> Tuple[bool,str]:
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as wr:
            df.to_excel(wr, index=False, sheet_name=SHEET_CLIENTS)
        return True, ""
    except Exception as e:
        return False, str(e)

# mémorisation dernier chemin
LAST_KEY = "last_paths_v2"
def save_last_paths(clients_path: str, visa_path: str):
    st.session_state[LAST_KEY] = {"clients": clients_path, "visa": visa_path}

def load_last_paths() -> Tuple[str|None, str|None]:
    obj = st.session_state.get(LAST_KEY) or {}
    return obj.get("clients"), obj.get("visa")

# ==============================================
# Normalisation du tableau Clients
# ==============================================
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Dossier N","ID_Client","Nom","Date","Mois","Categorie","Sous-categorie","Visa",
            HONO, AUTRE, TOTAL, "Payé","Reste","Paiements","Options","Notes",*STATUS_COLS
        ])
    for c in ["Nom","Categorie","Sous-categorie","Visa","Notes"]:
        if c in df.columns:
            df[c] = df[c].astype(str)

    # Montants
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c not in df.columns:
            df[c] = 0.0
        df[c] = df[c].apply(_to_float)

    # Total si absent
    if TOTAL in df.columns:
        df[TOTAL] = df[HONO] + df[AUTRE]

    # Paiements → liste d’objets
    if "Paiements" not in df.columns:
        df["Paiements"] = [[] for _ in range(len(df))]
    else:
        def _to_list(x):
            if isinstance(x, list):
                return x
            s = _safe_str(x)
            if not s:
                return []
            try:
                v = json.loads(s)
                return v if isinstance(v, list) else []
            except Exception:
                return []
        df["Paiements"] = df["Paiements"].apply(_to_list)

    # Recalcule Payé / Reste
    pay_calc = []
    for pays in df["Paiements"]:
        s = 0.0
        for p in pays:
            try:
                s += _to_float(p.get("montant", 0.0))
            except Exception:
                pass
        pay_calc.append(s)
    df["Payé"] = pay_calc
    df["Reste"] = (df[HONO] + df[AUTRE]) - df["Payé"]
    df["Reste"] = df["Reste"].apply(lambda x: max(0.0, float(x)))

    # Dates / Mois / Année
    if "Date" in df.columns:
        dd = pd.to_datetime(df["Date"], errors="coerce")
    else:
        dd = pd.to_datetime(pd.NaT)
        df["Date"] = dd
    df["_Année_"] = dd.dt.year.fillna(0).astype(int)
    df["_MoisNum_"] = dd.dt.month.fillna(0).astype(int)
    if "Mois" not in df.columns:
        df["Mois"] = dd.dt.month.fillna(0).astype(int).apply(lambda m: f"{int(m):02d}" if m else "")
    else:
        df["Mois"] = df["Mois"].astype(str).str.zfill(2)

    # Statuts si manque
    for c in STATUS_COLS:
        if c not in df.columns:
            df[c] = 0 if not c.startswith("Date") else None

    # Dossier N / ID_Client
    if "Dossier N" not in df.columns:
        df["Dossier N"] = None
    if "ID_Client" not in df.columns:
        df["ID_Client"] = df.apply(lambda r: make_client_id(r.get("Nom",""), r.get("Date")), axis=1)

    return df

def make_client_id(nom: str, d) -> str:
    base = _norm(nom).replace(" ", "-")
    if not base: base = "client"
    if isinstance(d, (date, datetime)):
        return f"{base}-{d:%Y%m%d}"
    try:
        dd = pd.to_datetime(d, errors="coerce")
        if pd.isna(dd):
            return f"{base}-{date.today():%Y%m%d}"
        return f"{base}-{dd:%Y%m%d}"
    except Exception:
        return f"{base}-{date.today():%Y%m%d}"

def next_dossier_number(df: pd.DataFrame, start=13057) -> int:
    if "Dossier N" in df.columns:
        try:
            mx = pd.to_numeric(df["Dossier N"], errors="coerce").fillna(0).max()
            return int(max(start, mx+1))
        except Exception:
            return int(start)
    return int(start)

# ==============================================
# Visa.xlsx → dictionnaire de sélection
# - Colonnes obligatoires : "Categorie", "Sous-categorie"
# - Les autres colonnes d’en-tête sont les "options"
# - Chaque ligne indique par 1/0 quelles options sont actives
# ==============================================
@st.cache_data(show_spinner=False)
def build_visa_map(visa_df: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, List[str]]]]:
    if visa_df is None or visa_df.empty:
        return {}
    # Normalise noms de colonnes
    cols = [c for c in visa_df.columns]
    # Repère bases
    cat_col = None
    sub_col = None
    for c in cols:
        if _norm(c) in ("categorie", "category"): cat_col = c
        if _norm(c) in ("sous-categorie", "sous-categories", "subcategory", "sous-categorie 1"): sub_col = c
    if not cat_col or not sub_col:
        # essaye tout de même
        cat_col = cat_col or "Categorie"
        sub_col = sub_col or "Sous-categorie"

    opt_cols = [c for c in cols if c not in (cat_col, sub_col)]

    vmap: Dict[str, Dict[str, Dict[str, List[str]]]] = {}
    for _, row in visa_df.iterrows():
        cat = _safe_str(row.get(cat_col, ""))
        sub = _safe_str(row.get(sub_col, ""))
        if not cat or not sub:
            continue
        options = []
        for oc in opt_cols:
            val = row.get(oc, 0)
            try:
                flag = int(_to_float(val))
            except Exception:
                flag = 0
            if flag == 1:
                options.append(str(oc))
        vmap.setdefault(cat, {})
        vmap[cat][sub] = {"options": options}
    return vmap

def render_option_checkboxes(options: List[str], keyprefix: str, preselected: List[str]|None=None) -> List[str]:
    pre = set(preselected or [])
    cols = st.columns(max(1, min(4, len(options)))) if options else [st]
    selected = []
    for i, opt in enumerate(options):
        col = cols[i % len(cols)]
        checked = opt in pre
        if col.checkbox(opt, value=checked, key=f"{keyprefix}_{_norm(opt)}_{SID}"):
            selected.append(opt)
    return selected

def compute_visa_string(sub: str, options: List[str]) -> str:
    sub = _safe_str(sub)
    if not options:
        return sub
    if len(options) == 1:
        return f"{sub} {options[0]}"
    return f"{sub} {'+'.join(options)}"

# ==============================================
# Zone de chargement / remémoration des fichiers
# ==============================================
st.title("🛂 Visa Manager")

# Tente de recharger les derniers fichiers
last_clients, last_visa = load_last_paths()
clients_path = last_clients if last_clients and os.path.exists(last_clients) else CLIENTS_DEFAULT if os.path.exists(CLIENTS_DEFAULT) else ""
visa_path    = last_visa    if last_visa and os.path.exists(last_visa)     else VISA_DEFAULT    if os.path.exists(VISA_DEFAULT)    else ""

st.sidebar.header("📂 Fichiers")
up_mode = st.sidebar.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"], key=f"mode_{SID}")

if up_mode == "Deux fichiers (Clients & Visa)":
    upC = st.sidebar.file_uploader("Clients (xlsx)", type=["xlsx"], key=f"upC_{SID}")
    upV = st.sidebar.file_uploader("Visa (xlsx)", type=["xlsx"], key=f"upV_{SID}")
    if upC:
        clients_path = os.path.join(WORK_DIR, "clients_current.xlsx")
        _save_bytes_to(clients_path, upC.read())
    if upV:
        visa_path = os.path.join(WORK_DIR, "visa_current.xlsx")
        _save_bytes_to(visa_path, upV.read())
else:
    upBoth = st.sidebar.file_uploader("Fichier unique (xlsx) avec onglets Clients & Visa", type=["xlsx"], key=f"upBoth_{SID}")
    if upBoth:
        both_path = os.path.join(WORK_DIR, "both_current.xlsx")
        _save_bytes_to(both_path, upBoth.read())
        # duplique sur 2 chemins pour le reste de l’app
        clients_path = both_path
        visa_path    = both_path

# Si rien encore, essaie défaut
if not clients_path or not os.path.exists(clients_path):
    st.sidebar.warning("Veuillez fournir le fichier Clients.")
if not visa_path or not os.path.exists(visa_path):
    st.sidebar.warning("Veuillez fournir le fichier Visa.")

if clients_path and visa_path and os.path.exists(clients_path) and os.path.exists(visa_path):
    save_last_paths(clients_path, visa_path)

# Téléchargements rapides des fichiers en mémoire
st.sidebar.markdown("---")
if clients_path and os.path.exists(clients_path):
    with open(clients_path, "rb") as f:
        st.sidebar.download_button("⬇️ Télécharger Clients", f.read(), file_name=os.path.basename(clients_path), key=f"dlC_{SID}")
if visa_path and os.path.exists(visa_path):
    with open(visa_path, "rb") as f:
        st.sidebar.download_button("⬇️ Télécharger Visa", f.read(), file_name=os.path.basename(visa_path), key=f"dlV_{SID}")

# Charge dataframes
df_clients_raw = pd.DataFrame()
df_visa_raw    = pd.DataFrame()
if clients_path and os.path.exists(clients_path):
    try:
        df_clients_raw = read_clients_file(clients_path)
    except Exception as e:
        st.error(f"Lecture Clients impossible : {e}")

if visa_path and os.path.exists(visa_path):
    try:
        df_visa_raw = read_visa_file(visa_path)
    except Exception as e:
        st.error(f"Lecture Visa impossible : {e}")

# Normalise / construit visa_map
df_all = normalize_clients(df_clients_raw.copy()) if not df_clients_raw.empty else normalize_clients(pd.DataFrame())
visa_map = build_visa_map(df_visa_raw.copy()) if not df_visa_raw.empty else {}

# Tabs
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "📄 Visa (aperçu)"])


# ==============================================
# 📊 Dashboard
# ==============================================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        years = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1, 13)]
        cats = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subs = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        c1, c2, c3, c4, c5 = st.columns(5)
        fy = c1.multiselect("Année", years, default=[], key=f"dash_y_{SID}")
        fm = c2.multiselect("Mois (MM)", months, default=[], key=f"dash_m_{SID}")
        fc = c3.multiselect("Catégorie", cats, default=[], key=f"dash_c_{SID}")
        fs = c4.multiselect("Sous-catégorie", subs, default=[], key=f"dash_s_{SID}")
        fv = c5.multiselect("Visa", visas, default=[], key=f"dash_v_{SID}")

        view = df_all.copy()
        if fy: view = view[view["_Année_"].isin(fy)]
        if fm: view = view[view["Mois"].astype(str).isin(fm)]
        if fc: view = view[view["Categorie"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(view)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(view, HONO).sum()))
        k3.metric("Payé", _fmt_money(_safe_num_series(view, "Payé").sum()))
        k4.metric("Reste", _fmt_money(_safe_num_series(view, "Reste").sum()))

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in view.columns]
        show_cols = list(dict.fromkeys(show_cols))  # unique

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

        st.dataframe(view_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_tbl_{SID}")

# ==============================================
# 📈 Analyses
# ==============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
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
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(dfA, HONO).sum()))
        k3.metric("Payé", _fmt_money(_safe_num_series(dfA, "Payé").sum()))
        k4.metric("Reste", _fmt_money(_safe_num_series(dfA, "Reste").sum()))

        # Graphiques simples
        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("### 📊 Dossiers par catégorie")
            vc = dfA["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie", "Nombre"]
            st.bar_chart(vc.set_index("Categorie"))

        if not dfA.empty and "Mois" in dfA.columns:
            st.markdown("### 📈 Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois")
            st.line_chart(gm.set_index("Mois"))

        # Détails
        st.markdown("### 🧾 Détails des dossiers filtrés")
        det = dfA.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).apply(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in det.columns]
        show_cols = list(dict.fromkeys(show_cols))

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det

        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"a_detail_{SID}")

# ==============================================
# 🏦 Escrow
# ==============================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Payé"]  = _safe_num_series(dfE, "Payé")
        dfE["Reste"] = _safe_num_series(dfE, "Reste")
        dfE[TOTAL]   = _safe_num_series(dfE, TOTAL)

        agg = dfE.groupby("Categorie", as_index=False)[[TOTAL, "Payé", "Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg[TOTAL]).replace([pd.NA, pd.NaT], 0).fillna(0.0) * 100
        st.dataframe(agg, use_container_width=True, key=f"esc_agg_{SID}")

        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money(float(dfE[TOTAL].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Reste", _fmt_money(float(dfE["Reste"].sum())))

        st.caption("NB : on peut isoler les honoraires perçus avant l’envoi, puis signaler les transferts à effectuer lorsque « Dossier envoyé » est coché.")

# ==============================================
# 👤 Clients — CRUD & paiements
# ==============================================
with tabs[3]:
    st.subheader("👤 Clients — Gestion & Suivi")

    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_{SID}")

    live = normalize_clients(df_clients_raw.copy())

    # ------- Ajouter -------
    if op == "Ajouter":
        c1, c2, c3 = st.columns(3)
        nom  = c1.text_input("Nom", key=f"add_nom_{SID}")
        dt   = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=date.today().month-1, key=f"add_mois_{SID}")

        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=f"add_cat_{SID}")
        subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
        sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=f"add_sub_{SID}")

        options_available = []
        if sel_cat and sel_sub and sel_cat in visa_map and sel_sub in visa_map[sel_cat]:
            options_available = visa_map[sel_cat][sel_sub]["options"]

        opts_sel = render_option_checkboxes(options_available, keyprefix=f"add_opts_{SID}")
        visa_final = compute_visa_string(sel_sub, opts_sel) if sel_sub else ""

        f1, f2 = st.columns(2)
        honor = f1.number_input(HONO, min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
        other = f2.number_input(AUTRE, min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")

        st.markdown("#### 📌 Statuts initiaux")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent  = s1.checkbox("Dossier envoyé", key=f"add_sent_{SID}")
        sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(None), key=f"add_sentd_{SID}")
        acc   = s2.checkbox("Dossier accepté", key=f"add_acc_{SID}")
        acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(None), key=f"add_accd_{SID}")
        ref   = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
        ref_d  = s3.date_input("Date de refus", value=_date_for_widget(None), key=f"add_refd_{SID}")
        ann   = s4.checkbox("Dossier annulé", key=f"add_ann_{SID}")
        ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(None), key=f"add_annd_{SID}")
        rfe   = s5.checkbox("RFE", key=f"add_rfe_{SID}")
        if rfe and not any([sent, acc, ref, ann]):
            st.warning("⚠️ RFE ne peut être coché qu’avec un autre statut.")

        note = st.text_area("Notes", key=f"add_note_{SID}")

        if st.button("💾 Enregistrer le client", key=f"btn_add_{SID}"):
            if not nom or not sel_cat or not sel_sub:
                st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                st.stop()

            total = float(honor) + float(other)
            paye  = 0.0
            reste = total

            new_row = {
                "Dossier N": next_dossier_number(live, start=13057),
                "ID_Client": make_client_id(nom, dt),
                "Nom": nom,
                "Date": dt,
                "Mois": mois,
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": visa_final or sel_sub,
                HONO: float(honor),
                AUTRE: float(other),
                TOTAL: total,
                "Payé": paye,
                "Reste": reste,
                "Paiements": [],
                "Options": {"options": opts_sel, "exclusive": None},
                "Notes": note,
                "Dossier envoyé": 1 if sent else 0,
                "Date d'envoi": (dt if sent else None) if not sent_d else sent_d,
                "Dossier accepté": 1 if acc else 0,
                "Date d'acceptation": acc_d if acc else None,
                "Dossier refusé": 1 if ref else 0,
                "Date de refus": ref_d if ref else None,
                "Dossier annulé": 1 if ann else 0,
                "Date d'annulation": ann_d if ann else None,
                "RFE": 1 if rfe else 0,
            }
            out = pd.concat([live, pd.DataFrame([new_row])], ignore_index=True)
            ok, err = write_clients_file(out, clients_path)
            if ok:
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()
            else:
                st.error(f"Erreur d’écriture : {err}")

    # ------- Modifier -------
    elif op == "Modifier":
        if live.empty:
            st.info("Aucun client.")
        else:
            names = sorted(live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(live["ID_Client"].dropna().astype(str).unique().tolist())
            c1, c2 = st.columns(2)
            sel_name = c1.selectbox("Nom", [""]+names, index=0, key=f"mod_n_{SID}")
            sel_id   = c2.selectbox("ID_Client", [""]+ids, index=0, key=f"mod_i_{SID}")

            mask = None
            if sel_id:
                mask = (live["ID_Client"].astype(str) == sel_id)
            elif sel_name:
                mask = (live["Nom"].astype(str) == sel_name)

            if mask is None or not mask.any():
                st.stop()

            idx = live[mask].index[0]
            row = live.loc[idx].copy()

            # KPI financiers
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Honoraires", _fmt_money(float(_to_float(row.get(HONO, 0.0)))))
            k2.metric("Autres frais", _fmt_money(float(_to_float(row.get(AUTRE, 0.0)))))
            k3.metric("Payé", _fmt_money(float(_to_float(row.get("Payé", 0.0)))))
            reste_calc = max(0.0, float(_to_float(row.get(TOTAL, 0.0))) - float(_to_float(row.get("Payé", 0.0))))
            k4.metric("Reste", _fmt_money(reste_calc))

            d1, d2, d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nom_{SID}")
            dt   = d2.date_input("Date de création", value=_date_for_widget(row.get("Date")), key=f"mod_date_{SID}")
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=_month_index(row.get("Mois")), key=f"mod_mois_{SID}")

            st.markdown("#### 🎯 Choix Visa (modification)")
            cats = sorted(list(visa_map.keys()))
            preset_cat_raw = _safe_str(row.get("Categorie",""))
            cat_index = _best_index(cats, preset_cat_raw)
            sel_cat = st.selectbox("Catégorie", [""] + cats, index=cat_index, key=f"mod_cat_{SID}")

            subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
            preset_sub_raw = _safe_str(row.get("Sous-categorie",""))
            sub_index = _best_index(subs, preset_sub_raw)
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=sub_index, key=f"mod_sub_{SID}")

            options_available = []
            if sel_cat and sel_sub and sel_cat in visa_map and sel_sub in visa_map[sel_cat]:
                options_available = visa_map[sel_cat][sel_sub]["options"]

            preset_opts = row.get("Options", {})
            if not isinstance(preset_opts, dict):
                try:
                    preset_opts = json.loads(_safe_str(preset_opts) or "{}")
                except Exception:
                    preset_opts = {}
            preset_list = preset_opts.get("options", []) if isinstance(preset_opts, dict) else []
            # normalise pour précocher ce qui existe encore
            norm_available = { _norm(o): o for o in options_available }
            pre_coch = [ norm_available[_norm(p)] for p in preset_list if _norm(p) in norm_available ]

            opts_sel = render_option_checkboxes(options_available, keyprefix=f"mod_opts_{SID}", preselected=pre_coch)
            visa_final = compute_visa_string(sel_sub, opts_sel) if sel_sub else _safe_str(row.get("Visa",""))

            st.markdown("#### 💵 Montants")
            f1, f2, f3 = st.columns(3)
            honor = f1.number_input(HONO, min_value=0.0,
                                    value=float(_to_float(row.get(HONO, 0.0))), step=50.0, format="%.2f", key=f"mod_h_{SID}")
            other = f2.number_input(AUTRE, min_value=0.0,
                                    value=float(_to_float(row.get(AUTRE, 0.0))), step=20.0, format="%.2f", key=f"mod_o_{SID}")
            total = float(honor) + float(other)
            f3.metric("Total (US $)", _fmt_money(total))

            st.markdown("#### 📌 Statuts & dates")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent  = s1.checkbox("Dossier envoyé", value=bool(row.get("Dossier envoyé")), key=f"mod_sent_{SID}")
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
            acc   = s2.checkbox("Dossier accepté", value=bool(row.get("Dossier accepté")), key=f"mod_acc_{SID}")
            acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
            ref   = s3.checkbox("Dossier refusé", value=bool(row.get("Dossier refusé")), key=f"mod_ref_{SID}")
            ref_d  = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=f"mod_refd_{SID}")
            ann   = s4.checkbox("Dossier annulé", value=bool(row.get("Dossier annulé")), key=f"mod_ann_{SID}")
            ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
            rfe   = s5.checkbox("RFE", value=bool(row.get("RFE")), key=f"mod_rfe_{SID}")
            if rfe and not any([sent, acc, ref, ann]):
                st.warning("⚠️ RFE ne peut être coché qu’avec un autre statut.")

            note = st.text_area("Notes", value=_safe_str(row.get("Notes","")), key=f"mod_note_{SID}")

            if st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}"):
                if not nom or not sel_cat or not sel_sub:
                    st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                    st.stop()

                paye  = float(_to_float(row.get("Payé", 0.0)))
                reste = max(0.0, total - paye)

                live.at[idx, "Nom"] = nom
                live.at[idx, "Date"] = dt
                live.at[idx, "Mois"] = _safe_str(mois)
                live.at[idx, "Categorie"] = sel_cat
                live.at[idx, "Sous-categorie"] = sel_sub
                live.at[idx, "Visa"] = visa_final or sel_sub
                live.at[idx, HONO] = float(honor)
                live.at[idx, AUTRE] = float(other)
                live.at[idx, TOTAL] = float(total)
                live.at[idx, "Reste"] = float(reste)
                live.at[idx, "Options"] = {"options": opts_sel, "exclusive": None}
                live.at[idx, "Notes"] = note
                live.at[idx, "Dossier envoyé"] = 1 if sent else 0
                live.at[idx, "Date d'envoi"] = (dt if sent else None) if not sent_d else sent_d
                live.at[idx, "Dossier accepté"] = 1 if acc else 0
                live.at[idx, "Date d'acceptation"] = acc_d if acc else None
                live.at[idx, "Dossier refusé"] = 1 if ref else 0
                live.at[idx, "Date de refus"] = ref_d if ref else None
                live.at[idx, "Dossier annulé"] = 1 if ann else 0
                live.at[idx, "Date d'annulation"] = ann_d if ann else None
                live.at[idx, "RFE"] = 1 if rfe else 0

                ok, err = write_clients_file(live, clients_path)
                if ok:
                    st.success("Modifications enregistrées.")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(f"Erreur d’écriture : {err}")

            # ---- Paiements ----
            st.markdown("#### 💵 Paiements")
            reste_actu = float(_to_float(live.loc[idx, "Reste"]))
            st.info(f"Reste actuel : {_fmt_money(reste_actu)}")

            paycol1, paycol2, paycol3 = st.columns(3)
            if reste_actu > 0:
                pay_amt  = paycol1.number_input("Montant à encaisser", min_value=0.0, step=10.0, format="%.2f", key=f"p_add_{SID}")
                pay_date = paycol2.date_input("Date paiement", value=date.today(), key=f"p_date_{SID}")
                mode     = paycol3.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"p_mode_{SID}")
                if st.button("Ajouter le paiement", key=f"p_btn_{SID}"):
                    if pay_amt <= 0:
                        st.warning("Montant > 0 requis.")
                        st.stop()
                    pays = row.get("Paiements", [])
                    if not isinstance(pays, list):
                        try:
                            pays = json.loads(_safe_str(pays) or "[]")
                        except Exception:
                            pays = []
                    pays.append({"date": str(pay_date), "montant": float(pay_amt), "mode": mode})
                    paye_new  = float(_to_float(live.loc[idx, "Payé"])) + float(pay_amt)
                    reste_new = max(0.0, float(_to_float(live.loc[idx, TOTAL])) - paye_new)
                    live.at[idx, "Paiements"] = pays
                    live.at[idx, "Payé"] = paye_new
                    live.at[idx, "Reste"] = reste_new
                    ok, err = write_clients_file(live, clients_path)
                    if ok:
                        st.success("Paiement ajouté.")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(f"Erreur écriture : {err}")

            hist = row.get("Paiements", [])
            if not isinstance(hist, list):
                try:
                    hist = json.loads(_safe_str(hist) or "[]")
                except Exception:
                    hist = []
            if hist:
                st.write("Historique des paiements :")
                st.table(pd.DataFrame(hist))

    # ------- Supprimer -------
    elif op == "Supprimer":
        if live.empty:
            st.info("Aucun client.")
        else:
            names = sorted(live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(live["ID_Client"].dropna().astype(str).unique().tolist())
            c1, c2 = st.columns(2)
            sel_name = c1.selectbox("Nom", [""]+names, index=0, key=f"del_n_{SID}")
            sel_id   = c2.selectbox("ID_Client", [""]+ids, index=0, key=f"del_i_{SID}")

            mask = None
            if sel_id:
                mask = (live["ID_Client"].astype(str) == sel_id)
            elif sel_name:
                mask = (live["Nom"].astype(str) == sel_name)

            if mask is not None and mask.any():
                row = live[mask].iloc[0]
                st.write({"Dossier N": row.get("Dossier N",""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                    out = live[~mask].copy()
                    ok, err = write_clients_file(out, clients_path)
                    if ok:
                        st.success("Client supprimé.")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(f"Erreur écriture : {err}")

# ==============================================
# 📄 Visa (aperçu) + Export ZIP
# ==============================================
with tabs[4]:
    st.subheader("📄 Visa — aperçu & export")

    st.markdown("#### Aperçu du fichier Visa")
    if df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True, key=f"v_tbl_{SID}")

    st.markdown("#### Export global (Clients + Visa)")
    colz1, colz2 = st.columns([1,3])
    with colz1:
        if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID}"):
            try:
                buf = BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    # Clients propre
                    df_export = normalize_clients(read_clients_file(clients_path))
                    with BytesIO() as xbuf:
                        with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                            df_export.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                        zf.writestr("Clients.xlsx", xbuf.getvalue())
                    # Visa (tel quel si possible)
                    try:
                        zf.write(visa_path, "Visa.xlsx")
                    except Exception:
                        with BytesIO() as vb:
                            with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                df_visa_raw.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                            zf.writestr("Visa.xlsx", vb.getvalue())
                st.session_state[f"zip_export_{SID}"] = buf.getvalue()
                st.success("Archive prête.")
            except Exception as e:
                st.error("Erreur de préparation : " + _safe_str(e))

    with colz2:
        if st.session_state.get(f"zip_export_{SID}"):
            st.download_button(
                label="⬇️ Télécharger l’export (ZIP)",
                data=st.session_state[f"zip_export_{SID}"],
                file_name="Export_Visa_Manager.zip",
                mime="application/zip",
                key=f"zip_dl_{SID}",
            )