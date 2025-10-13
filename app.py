# ===============================
# 🛂 Visa Manager — PARTIE 1/2
# ===============================
from __future__ import annotations

import os, re, json, zipfile, uuid
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st

# ---------- Config & clé de session ----------
SID = st.session_state.get("_sid") or uuid.uuid4().hex[:8]
st.session_state["_sid"] = SID

st.set_page_config(page_title="Visa Manager", page_icon="🛂", layout="wide")
st.title("🛂 Visa Manager")

# ---------- Styles compacts KPI & menus auto-larges ----------
st.markdown("""
<style>
/* KPI (st.metric) compacts */
.small-metrics [data-testid="stMetricValue"] { font-size: 1.05rem; line-height: 1.05rem; }
.small-metrics [data-testid="stMetricLabel"] { font-size: 0.80rem; opacity: 0.85; }
.small-metrics [data-testid="stMetricDelta"] { font-size: 0.75rem; }
.small-metrics [data-testid="stVerticalBlock"] { gap: 0.35rem; }

/* Menus déroulants : largeur selon contenu (sans dépasser le conteneur) */
.stSelectbox [data-baseweb="select"],
.stMultiSelect [data-baseweb="select"] {
  width: max-content !important;
  min-width: 12ch;
  max-width: 100% !important;
}
.stSelectbox div[role="combobox"],
.stMultiSelect div[role="combobox"] {
  min-width: 0 !important;
}
/* Un peu plus compacts */
.stSelectbox [data-baseweb="select"] div,
.stMultiSelect [data-baseweb="select"] div {
  padding-top: 2px; padding-bottom: 2px;
}
</style>
""", unsafe_allow_html=True)

# ---------- Constantes colonnes ----------
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

# ===============================
# Helpers format / parse
# ===============================
def _norm(s: str) -> str:
    s = str(s or "")
    s = s.strip().lower()
    s = re.sub(r"[’'`´]", "", s)
    s = re.sub(r"[^a-z0-9+/_\\- ]+", " ", s)
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

def widget_date(val, fallback=None):
    """Retourne un date (ou fallback) compatible st.date_input (évite NaT/NaN)."""
    if val is None:
        return fallback
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return fallback
        return d.date()
    except Exception:
        return fallback

def _month_index(mois_val) -> int:
    s = _safe_str(mois_val)
    if s.isdigit():
        i = int(s)
        if 1 <= i <= 12:
            return i - 1
    return 0

def _best_index(options: List[str], current: str) -> int:
    cur = _norm(current)
    for i, opt in enumerate(options, start=1):
        if _norm(opt) == cur:
            return i
    return 0

# ===============================
# I/O — chemins, dernier fichier, lecture/écriture
# ===============================
WORK_DIR = os.getcwd()
CLIENTS_DEFAULT = os.path.join(WORK_DIR, "donnees_visa_clients1_adapte.xlsx")
VISA_DEFAULT    = os.path.join(WORK_DIR, "donnees_visa_clients1.xlsx")

def _save_bytes_to(path: str, data: bytes) -> None:
    with open(path, "wb") as f:
        f.write(data)

@st.cache_data(show_spinner=False)
def read_clients_file(path: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        return pd.read_excel(path)

@st.cache_data(show_spinner=False)
def read_visa_file(path: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=SHEET_VISA)
    except Exception:
        return pd.read_excel(path)

def write_workbook(clients_df: pd.DataFrame, clients_path: str,
                   visa_df: pd.DataFrame|None, visa_path: str) -> Tuple[bool, str]:
    """Écrit la feuille Clients et, si fichier unique, la feuille Visa aussi."""
    try:
        if clients_path == visa_path and visa_df is not None:
            with pd.ExcelWriter(clients_path, engine="openpyxl") as wr:
                clients_df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                visa_df.to_excel(wr, sheet_name=SHEET_VISA, index=False)
        else:
            with pd.ExcelWriter(clients_path, engine="openpyxl") as wr:
                clients_df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        return True, ""
    except Exception as e:
        return False, str(e)

LAST_KEY = "last_paths_v4"
def save_last_paths(clients_path: str, visa_path: str):
    st.session_state[LAST_KEY] = {"clients": clients_path, "visa": visa_path}

def load_last_paths() -> Tuple[str|None, str|None]:
    obj = st.session_state.get(LAST_KEY) or {}
    return obj.get("clients"), obj.get("visa")

# ===============================
# Normalisation Clients
# ===============================
def make_client_id(nom: str, d) -> str:
    base = _norm(nom).replace(" ", "-") or "client"
    try:
        dd = pd.to_datetime(d, errors="coerce")
        if pd.isna(dd):
            return f"{base}-{date.today():%Y%m%d}"
        return f"{base}-{dd:%Y%m%d}"
    except Exception:
        return f"{base}-{date.today():%Y%m%d}"

def next_dossier_number(df: pd.DataFrame, start=13057) -> int:
    if "Dossier N" in df.columns:
        mx = pd.to_numeric(df["Dossier N"], errors="coerce").fillna(0).max()
        try:
            return int(max(start, mx + 1))
        except Exception:
            return int(start)
    return int(start)

def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Dossier N","ID_Client","Nom","Date","Mois","Categorie","Sous-categorie","Visa",
            HONO, AUTRE, TOTAL, "Payé","Reste","Paiements","Options","Notes",*STATUS_COLS
        ])

    for c in ["Nom","Categorie","Sous-categorie","Visa","Notes"]:
        if c in df.columns:
            df[c] = df[c].astype(str)

    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c not in df.columns:
            df[c] = 0.0
        df[c] = df[c].apply(_to_float)

    if TOTAL in df.columns:
        df[TOTAL] = df[HONO] + df[AUTRE]

    # Paiements -> liste structurée
    if "Paiements" not in df.columns:
        df["Paiements"] = [[] for _ in range(len(df))]
    else:
        def _to_list(x):
            if isinstance(x, list): return x
            s = _safe_str(x)
            if not s: return []
            try:
                v = json.loads(s)
                return v if isinstance(v, list) else []
            except Exception:
                return []
        df["Paiements"] = df["Paiements"].apply(_to_list)

    # Payé / Reste
    pay_calc = []
    for pays in df["Paiements"]:
        s = 0.0
        for p in pays:
            try: s += _to_float(p.get("montant", 0.0))
            except Exception: pass
        pay_calc.append(s)
    df["Payé"] = pay_calc
    df["Reste"] = (df[HONO] + df[AUTRE]) - df["Payé"]
    df["Reste"] = df["Reste"].apply(lambda x: max(0.0, float(x)))

    # Dates principales
    dd = pd.to_datetime(df.get("Date", pd.NaT), errors="coerce")
    df["Date"] = dd
    df["_Année_"] = dd.dt.year.fillna(0).astype(int)
    df["_MoisNum_"] = dd.dt.month.fillna(0).astype(int)
    if "Mois" not in df.columns:
        df["Mois"] = dd.dt.month.fillna(0).astype(int).apply(lambda m: f"{int(m):02d}" if m else "")
    else:
        df["Mois"] = df["Mois"].astype(str).str.zfill(2)

    # Statuts par défaut
    for c in STATUS_COLS:
        if c not in df.columns:
            df[c] = 0 if not c.startswith("Date") else None

    if "Dossier N" not in df.columns:
        df["Dossier N"] = None
    if "ID_Client" not in df.columns:
        df["ID_Client"] = df.apply(lambda r: make_client_id(r.get("Nom",""), r.get("Date")), axis=1)

    return df

# ===============================
# Visa → carte des options
# - colonnes "Categorie", "Sous-categorie"
# - toutes les autres colonnes = options ; valeur 1 => active
# ===============================
@st.cache_data(show_spinner=False)
def build_visa_map(visa_df: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, List[str]]]]:
    if visa_df is None or visa_df.empty:
        return {}
    cols = list(visa_df.columns)
    # noms robustes
    cat_col = None
    sub_col = None
    for c in cols:
        nm = _norm(c)
        if nm in ("categorie", "category"): cat_col = c
        if nm in ("sous-categorie", "sous-categories", "subcategory", "sous-categorie 1"): sub_col = c
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
        chk = opt in pre
        if col.checkbox(opt, value=chk, key=f"{keyprefix}_{_norm(opt)}_{SID}"):
            selected.append(opt)
    return selected

def compute_visa_string(sub: str, options: List[str]) -> str:
    sub = _safe_str(sub)
    if not options:
        return sub
    if len(options) == 1:
        return f"{sub} {options[0]}"
    return f"{sub} {'+'.join(options)}"

# ===============================
# Sidebar — chargement fichiers + mémoire dernier chemin
# ===============================
WORK_DIR = os.getcwd()
last_clients, last_visa = load_last_paths()
clients_path = last_clients if last_clients and os.path.exists(last_clients) else (os.path.join(WORK_DIR, "donnees_visa_clients1_adapte.xlsx") if os.path.exists(os.path.join(WORK_DIR, "donnees_visa_clients1_adapte.xlsx")) else "")
visa_path    = last_visa    if last_visa and os.path.exists(last_visa)       else (os.path.join(WORK_DIR, "donnees_visa_clients1.xlsx")          if os.path.exists(os.path.join(WORK_DIR, "donnees_visa_clients1.xlsx"))          else "")

st.sidebar.header("📂 Fichiers")
mode = st.sidebar.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"], key=f"mode_{SID}")

if mode == "Deux fichiers (Clients & Visa)":
    upC = st.sidebar.file_uploader("Clients (xlsx)", type=["xlsx"], key=f"upC_{SID}")
    upV = st.sidebar.file_uploader("Visa (xlsx)", type=["xlsx"], key=f"upV_{SID}")
    if upC:
        clients_path = os.path.join(WORK_DIR, "clients_current.xlsx")
        _save_bytes_to(clients_path, upC.read())
    if upV:
        visa_path = os.path.join(WORK_DIR, "visa_current.xlsx")
        _save_bytes_to(visa_path, upV.read())
else:
    upBoth = st.sidebar.file_uploader("Fichier unique (xlsx) avec onglets Clients & Visa", type=["xlsx"], key=f"upB_{SID}")
    if upBoth:
        both_path = os.path.join(WORK_DIR, "both_current.xlsx")
        _save_bytes_to(both_path, upBoth.read())
        clients_path = both_path
        visa_path    = both_path

if clients_path and visa_path and os.path.exists(clients_path) and os.path.exists(visa_path):
    save_last_paths(clients_path, visa_path)

st.sidebar.markdown("---")
if clients_path and os.path.exists(clients_path):
    with open(clients_path, "rb") as f:
        st.sidebar.download_button("⬇️ Télécharger Clients", f.read(), file_name=os.path.basename(clients_path), key=f"dlC_{SID}")
if visa_path and os.path.exists(visa_path):
    with open(visa_path, "rb") as f:
        st.sidebar.download_button("⬇️ Télécharger Visa", f.read(), file_name=os.path.basename(visa_path), key=f"dlV_{SID}")

# ===============================
# Lecture effective
# ===============================
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

df_all  = normalize_clients(df_clients_raw.copy()) if not df_clients_raw.empty else normalize_clients(pd.DataFrame())
visa_map = build_visa_map(df_visa_raw.copy()) if not df_visa_raw.empty else {}

# ===============================
# Tabs
# ===============================
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "📄 Visa (aperçu)"])


# ===============================
# 📊 Dashboard
# ===============================
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

        # KPIs compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(view)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(view, HONO).sum()))
        k3.metric("Payé", _fmt_money(_safe_num_series(view, "Payé").sum()))
        k4.metric("Reste", _fmt_money(_safe_num_series(view, "Reste").sum()))
        st.markdown('</div>', unsafe_allow_html=True)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in view.columns]
        show_cols = list(dict.fromkeys(show_cols))
        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

        st.dataframe(view_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_tbl_{SID}")

# ===============================
# 📈 Analyses
# ===============================
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

        # KPIs compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(dfA, HONO).sum()))
        k3.metric("Payé", _fmt_money(_safe_num_series(dfA, "Payé").sum()))
        k4.metric("Reste", _fmt_money(_safe_num_series(dfA, "Reste").sum()))
        st.markdown('</div>', unsafe_allow_html=True)

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

# ===============================
# 🏦 Escrow
# ===============================
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

        # KPIs compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money(float(dfE[TOTAL].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Reste", _fmt_money(float(dfE["Reste"].sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        st.caption("NB : pour un escrow « strict », isolez les honoraires perçus avant l’envoi, puis transférez à l’envoi.")

# ===============================
# 👤 Clients — CRUD & paiements
# ===============================
with tabs[3]:
    st.subheader("👤 Clients — Gestion & Suivi")
    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_{SID}")

    live = normalize_clients(df_clients_raw.copy())

    # ---- Ajouter ----
    if op == "Ajouter":
        c1, c2, c3 = st.columns(3)
        nom  = c1.text_input("Nom", key=f"add_nom_{SID}")
        dt   = c2.date_input("Date de création", value=widget_date(date.today(), fallback=date.today()), key=f"add_date_{SID}")
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=date.today().month-1, key=f"add_m_{SID}")

        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""]+cats, index=0, key=f"add_cat_{SID}")
        subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
        sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"add_sub_{SID}")

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
        sent_d = s1.date_input("Date d'envoi", value=widget_date(None, fallback=None), key=f"add_sentd_{SID}")
        acc   = s2.checkbox("Dossier accepté", key=f"add_acc_{SID}")
        acc_d  = s2.date_input("Date d'acceptation", value=widget_date(None, fallback=None), key=f"add_accd_{SID}")
        ref   = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
        ref_d  = s3.date_input("Date de refus", value=widget_date(None, fallback=None), key=f"add_refd_{SID}")
        ann   = s4.checkbox("Dossier annulé", key=f"add_ann_{SID}")
        ann_d  = s4.date_input("Date d'annulation", value=widget_date(None, fallback=None), key=f"add_annd_{SID}")
        rfe   = s5.checkbox("RFE", key=f"add_rfe_{SID}")
        if rfe and not any([sent, acc, ref, ann]):
            st.warning("⚠️ RFE ne peut être coché qu’avec un autre statut.")

        note = st.text_area("Notes", key=f"add_note_{SID}")

        if st.button("💾 Enregistrer le client", key=f"btn_add_{SID}"):
            if not nom or not sel_cat or not sel_sub:
                st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                st.stop()
            total = float(honor) + float(other)
            new_row = {
                "Dossier N": next_dossier_number(live, start=13057),
                "ID_Client": make_client_id(nom, dt),
                "Nom": nom,
                "Date": dt,
                "Mois": _safe_str(mois),
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": visa_final or sel_sub,
                HONO: float(honor),
                AUTRE: float(other),
                TOTAL: float(total),
                "Payé": 0.0,
                "Reste": float(total),
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
            ok, err = write_workbook(out, clients_path, df_visa_raw if clients_path==visa_path else None, visa_path)
            if ok:
                st.success("Client ajouté.")
                st.cache_data.clear(); st.rerun()
            else:
                st.error(f"Erreur d’écriture : {err}")

    # ---- Modifier ----
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

            # KPIs compacts info client
            st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Honoraires", _fmt_money(float(_to_float(row.get(HONO, 0.0)))))
            k2.metric("Autres frais", _fmt_money(float(_to_float(row.get(AUTRE, 0.0)))))
            k3.metric("Payé", _fmt_money(float(_to_float(row.get("Payé", 0.0)))))
            k4.metric("Reste", _fmt_money(max(0.0, float(_to_float(row.get(TOTAL, 0.0))) - float(_to_float(row.get("Payé", 0.0))))))
            st.markdown('</div>', unsafe_allow_html=True)

            d1, d2, d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nom_{SID}")
            dt   = d2.date_input("Date de création", value=widget_date(row.get("Date"), fallback=date.today()), key=f"mod_date_{SID}")
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=_month_index(row.get("Mois")), key=f"mod_mois_{SID}")

            st.markdown("#### 🎯 Choix Visa (modification)")
            cats = sorted(list(visa_map.keys()))
            preset_cat_raw = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""]+cats, index=_best_index(cats, preset_cat_raw), key=f"mod_cat_{SID}")

            subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
            preset_sub_raw = _safe_str(row.get("Sous-categorie",""))
            sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=_best_index(subs, preset_sub_raw), key=f"mod_sub_{SID}")

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
            norm_available = { _norm(o): o for o in options_available }
            pre_sel = [norm_available[_norm(p)] for p in preset_list if _norm(p) in norm_available]

            opts_sel = render_option_checkboxes(options_available, keyprefix=f"mod_opts_{SID}", preselected=pre_sel)
            visa_final = compute_visa_string(sel_sub, opts_sel) if sel_sub else _safe_str(row.get("Visa",""))

            st.markdown("#### 💵 Montants")
            f1, f2, f3 = st.columns(3)
            honor = f1.number_input(HONO, min_value=0.0, value=float(_to_float(row.get(HONO, 0.0))), step=50.0, format="%.2f", key=f"mod_h_{SID}")
            other = f2.number_input(AUTRE, min_value=0.0, value=float(_to_float(row.get(AUTRE, 0.0))), step=20.0, format="%.2f", key=f"mod_o_{SID}")
            total = float(honor) + float(other)
            f3.metric("Total (US $)", _fmt_money(total))

            st.markdown("#### 📌 Statuts & dates")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent  = s1.checkbox("Dossier envoyé", value=bool(row.get("Dossier envoyé")), key=f"mod_sent_{SID}")
            sent_d = s1.date_input("Date d'envoi", value=widget_date(row.get("Date d'envoi"), fallback=None), key=f"mod_sentd_{SID}")
            acc   = s2.checkbox("Dossier accepté", value=bool(row.get("Dossier accepté")), key=f"mod_acc_{SID}")
            acc_d  = s2.date_input("Date d'acceptation", value=widget_date(row.get("Date d'acceptation"), fallback=None), key=f"mod_accd_{SID}")
            ref   = s3.checkbox("Dossier refusé", value=bool(row.get("Dossier refusé")), key=f"mod_ref_{SID}")
            ref_d  = s3.date_input("Date de refus", value=widget_date(row.get("Date de refus"), fallback=None), key=f"mod_refd_{SID}")
            ann   = s4.checkbox("Dossier annulé", value=bool(row.get("Dossier annulé")), key=f"mod_ann_{SID}")
            ann_d  = s4.date_input("Date d'annulation", value=widget_date(row.get("Date d'annulation"), fallback=None), key=f"mod_annd_{SID}")
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

                ok, err = write_workbook(live, clients_path, df_visa_raw if clients_path==visa_path else None, visa_path)
                if ok:
                    st.success("Modifications enregistrées.")
                    st.cache_data.clear(); st.rerun()
                else:
                    st.error(f"Erreur d’écriture : {err}")

            # Paiements
            st.markdown("#### 💵 Paiements")
            reste_actu = float(_to_float(live.loc[idx, "Reste"]))
            st.info(f"Reste actuel : {_fmt_money(reste_actu)}")

            payc1, payc2, payc3 = st.columns(3)
            if reste_actu > 0:
                pay_amt  = payc1.number_input("Montant à encaisser", min_value=0.0, step=10.0, format="%.2f", key=f"p_add_{SID}")
                pay_date = payc2.date_input("Date paiement", value=date.today(), key=f"p_date_{SID}")
                mode     = payc3.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"p_mode_{SID}")
                if st.button("Ajouter le paiement", key=f"p_btn_{SID}"):
                    if pay_amt <= 0:
                        st.warning("Montant > 0 requis."); st.stop()
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
                    ok, err = write_workbook(live, clients_path, df_visa_raw if clients_path==visa_path else None, visa_path)
                    if ok:
                        st.success("Paiement ajouté."); st.cache_data.clear(); st.rerun()
                    else:
                        st.error(f"Erreur écriture : {err}")

            hist = row.get("Paiements", [])
            if not isinstance(hist, list):
                try: hist = json.loads(_safe_str(hist) or "[]")
                except Exception: hist = []
            if hist:
                st.write("Historique des paiements :")
                st.table(pd.DataFrame(hist))

    # ---- Supprimer ----
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
                    ok, err = write_workbook(out, clients_path, df_visa_raw if clients_path==visa_path else None, visa_path)
                    if ok:
                        st.success("Client supprimé."); st.cache_data.clear(); st.rerun()
                    else:
                        st.error(f"Erreur écriture : {err}")

# ===============================
# 📄 Visa (aperçu) & Export ZIP
# ===============================
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
                    # Clients normalisés
                    df_export = normalize_clients(read_clients_file(clients_path))
                    with BytesIO() as xbuf:
                        with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                            df_export.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                        zf.writestr("Clients.xlsx", xbuf.getvalue())
                    # Visa tel quel si possible
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



# ===============================
# ======   PARTIE 2 / 2    ======
# ===============================

# ---------- petits helpers locaux (écriture/lecture clients normalisés) ----------
def _read_clients(path: str) -> pd.DataFrame:
    try:
        df = read_clients_file(path)
    except Exception:
        df = pd.DataFrame()
    return normalize_clients(df)

def _write_clients(df_clients: pd.DataFrame, c_path: str) -> None:
    # si fichier unique => écrire aussi la feuille Visa telle qu'actuelle
    visa_df = None
    if c_path == visa_path:
        try:
            visa_df = read_visa_file(visa_path)
        except Exception:
            visa_df = None
    ok, err = write_workbook(df_clients, c_path, visa_df, visa_path)
    if not ok:
        st.error("Erreur d'écriture : " + _safe_str(err))

def _fmt_money_us(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(val, fallback=None):
    """Valeur sûre pour st.date_input (évite NaT)."""
    if isinstance(val, (date, datetime)):
        return val if isinstance(val, date) else val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d): return fallback
        return d.date()
    except Exception:
        return fallback

# Recharger en mémoire normalisée (au cas où la partie 1 n'a pas pu lire)
df_all   = df_all if 'df_all' in globals() else _read_clients(clients_path)
visa_map = visa_map if 'visa_map' in globals() else build_visa_map(read_visa_file(visa_path)) if (visa_path and os.path.exists(visa_path)) else {}

# =========================================================
# 🧭 Barre latérale — navigation / actions rapides
# =========================================================
with st.sidebar:
    st.markdown("### Navigation")
    st.caption("Utilise les onglets en haut : Dashboard / Analyses / Escrow / Clients / Visa.")
    st.markdown("---")
    st.markdown("### Actions rapides")
    st.write("• Ajouter / Modifier / Supprimer dans **👤 Clients**")
    st.write("• Éditer le paramétrage dans **📄 Visa (aperçu)**")

# =========================================================
# 📊 ONGLET : Dashboard (liste + KPI + filtres)
# =========================================================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        # Listes filtres
        years  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1, 13)]
        cats   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        f1, f2, f3, f4, f5 = st.columns([1,1,1,1,1])
        fy = f1.multiselect("Année", years, default=[], key=f"dash_years_{SID}")
        fm = f2.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{SID}")
        fc = f3.multiselect("Catégorie", cats, default=[], key=f"dash_cats_{SID}")
        fs = f4.multiselect("Sous-catégorie", subs, default=[], key=f"dash_subs_{SID}")
        fv = f5.multiselect("Visa", visas, default=[], key=f"dash_visas_{SID}")

        # Filtrage
        ff = df_all.copy()
        if fy: ff = ff[ff["_Année_"].isin(fy)]
        if fm: ff = ff[ff["Mois"].astype(str).isin(fm)]
        if fc: ff = ff[ff["Categorie"].astype(str).isin(fc)]
        if fs: ff = ff[ff["Sous-categorie"].astype(str).isin(fs)]
        if fv: ff = ff[ff["Visa"].astype(str).isin(fv)]

        # KPI compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(ff)}")
        k2.metric("Honoraires", _fmt_money_us(float(_safe_num_series(ff, HONO).sum())))
        k3.metric("Autres frais", _fmt_money_us(float(_safe_num_series(ff, AUTRE).sum())))
        k4.metric("Payé", _fmt_money_us(float(_safe_num_series(ff, "Payé").sum())))
        k5.metric("Reste", _fmt_money_us(float(_safe_num_series(ff, "Reste").sum())))
        st.markdown("</div>", unsafe_allow_html=True)

        # Liste
        view = ff.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in view.columns:
                view[c] = _safe_num_series(view, c).map(_fmt_money_us)
        if "Date" in view.columns:
            try:
                view["Date"] = pd.to_datetime(view["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                view["Date"] = view["Date"].astype(str)

        # colonnes à montrer
        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in view.columns]

        # Trier si colonnes présentes
        sort_cols = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        v = view.copy()
        if sort_cols:
            v = v.sort_values(by=sort_cols)
        # éviter doublons de noms de colonnes pour Arrow
        v = v.loc[:, ~v.columns.duplicated()].copy()

        st.dataframe(v[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_tbl_{SID}")

# =========================================================
# 📈 ONGLET : Analyses (séries & détails)
# =========================================================
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
        fy = a1.multiselect("Année", yearsA, default=[], key=f"anal_years_{SID}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"anal_months_{SID}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"anal_cats_{SID}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"anal_subs_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"anal_visas_{SID}")

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPI
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money_us(float(_safe_num_series(dfA, HONO).sum())))
        k3.metric("Payé", _fmt_money_us(float(_safe_num_series(dfA, "Payé").sum())))
        k4.metric("Reste", _fmt_money_us(float(_safe_num_series(dfA, "Reste").sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        # Graphiques rapides
        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("#### Dossiers par catégorie")
            vc = dfA["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie", "Nombre"]
            st.bar_chart(vc.set_index("Categorie"))

        if not dfA.empty and "Mois" in dfA.columns and HONO in dfA.columns:
            st.markdown("#### Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois")
            st.line_chart(gm.set_index("Mois"))

        # Détails
        st.markdown("#### Détails des dossiers filtrés")
        det = dfA.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).map(_fmt_money_us)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det
        det_sorted = det_sorted.loc[:, ~det_sorted.columns.duplicated()].copy()

        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"a_detail_{SID}")

# =========================================================
# 🏦 ONGLET : Escrow — synthèse & transferts
# =========================================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse & transferts")

    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Payé"]  = _safe_num_series(dfE, "Payé")
        dfE["Reste"] = _safe_num_series(dfE, "Reste")
        dfE[HONO]    = _safe_num_series(dfE, HONO)

        # Colonne historique transferts escrow
        if "EscrowTransfers" not in dfE.columns:
            dfE["EscrowTransfers"] = [[] for _ in range(len(dfE))]

        # Montant transféré cumulé
        def _esc_sum(lst):
            if isinstance(lst, list):
                s = 0.0
                for it in lst:
                    try:
                        s += _to_float(it.get("montant", 0.0))
                    except Exception:
                        pass
                return s
            try:
                v = json.loads(_safe_str(lst) or "[]")
                if isinstance(v, list):
                    return sum([_to_float(x.get("montant", 0.0)) for x in v])
            except Exception:
                pass
            return 0.0

        dfE["Escrow transféré"] = dfE["EscrowTransfers"].apply(_esc_sum)
        # Escrow dispo = min(Payé, Honoraires) - déjà transféré
        dfE["ESCROW dispo"] = (dfE[[ "Payé", HONO]].min(axis=1) - dfE["Escrow transféré"]).clip(lower=0)

        # vue agrégée par catégorie
        agg = dfE.groupby("Categorie", as_index=False)[[HONO, "Payé", "Reste", "ESCROW dispo"]].sum()
        st.dataframe(agg, use_container_width=True, key=f"esc_agg_{SID}")

        # KPIs
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        t1, t2, t3, t4 = st.columns(4)
        t1.metric("Honoraires", _fmt_money_us(float(dfE[HONO].sum())))
        t2.metric("Payé", _fmt_money_us(float(dfE["Payé"].sum())))
        t3.metric("Reste", _fmt_money_us(float(dfE["Reste"].sum())))
        t4.metric("Escrow dispo total", _fmt_money_us(float(dfE["ESCROW dispo"].sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("#### Dossiers « envoyés » avec escrow disponible")
        dfS = dfE[(dfE["Dossier envoyé"] == 1) & (dfE["ESCROW dispo"] > 0)].copy()
        if dfS.empty:
            st.info("Aucun transfert à réclamer.")
        else:
            # Écran de marquage de transfert
            live = _read_clients(clients_path)  # dataframe *éditable*
            for i, r in dfS.iterrows():
                b1, b2, b3, b4 = st.columns([2,2,2,2])
                b1.write(f"**{_safe_str(r.get('Nom',''))}** — {_safe_str(r.get('Visa',''))}")
                b2.write(f"Escrow dispo : **{_fmt_money_us(float(r['ESCROW dispo']))}**")
                amt = b3.number_input("Montant à marquer comme transféré (US $)",
                                      min_value=0.0,
                                      value=float(r["ESCROW dispo"]),
                                      step=10.0, format="%.2f",
                                      key=f"esc_amt_{SID}_{i}")
                if b4.button("Marquer transféré", key=f"esc_btn_{SID}_{i}"):
                    # append dans EscrowTransfers
                    mask = (live["ID_Client"].astype(str) == str(r["ID_Client"]))
                    if mask.any():
                        idx = live[mask].index[0]
                        lst = live.at[idx, "EscrowTransfers"] if "EscrowTransfers" in live.columns else None
                        if not isinstance(lst, list):
                            try:
                                lst = json.loads(_safe_str(lst) or "[]")
                                if not isinstance(lst, list): lst = []
                            except Exception:
                                lst = []
                        lst.append({"date": date.today().isoformat(), "montant": float(amt)})
                        live.at[idx, "EscrowTransfers"] = lst
                        _write_clients(live, clients_path)
                        st.success("Transfert marqué.")
                        st.cache_data.clear()
                        st.rerun()

# =========================================================
# 👤 ONGLET : Clients (CRUD + paiements + statuts)
# =========================================================
with tabs[3]:
    st.subheader("👤 Clients — Ajouter / Modifier / Supprimer")

    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_op_{SID}")
    df_live = _read_clients(clients_path)

    # ---------- AJOUT ----------
    if op == "Ajouter":
        c1, c2, c3 = st.columns(3)
        nom  = c1.text_input("Nom", "", key=f"add_nom_{SID}")
        dt   = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=date.today().month-1, key=f"add_mois_{SID}")

        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""]+cats, index=0, key=f"add_cat_{SID}")
        sel_sub = ""
        opts_selected: List[str] = []
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"add_sub_{SID}")
            if sel_sub:
                opt_list = visa_map[sel_cat][sel_sub].get("options", [])
                st.caption("Options disponibles (cases à cocher)")
                opts_selected = render_option_checkboxes(opt_list, keyprefix=f"add_opts_{SID}", preselected=[])

        visa_final = compute_visa_string(sel_sub, opts_selected)

        f1, f2 = st.columns(2)
        honor = f1.number_input(HONO, min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
        other = f2.number_input(AUTRE, min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")

        st.markdown("#### 📌 Statuts")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent   = s1.checkbox("Dossier envoyé", key=f"add_sent_{SID}")
        sent_d = s1.date_input("Date d'envoi", value=None, key=f"add_sentd_{SID}")
        acc    = s2.checkbox("Dossier accepté", key=f"add_acc_{SID}")
        acc_d  = s2.date_input("Date d'acceptation", value=None, key=f"add_accd_{SID}")
        ref    = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
        ref_d  = s3.date_input("Date de refus", value=None, key=f"add_refd_{SID}")
        ann    = s4.checkbox("Dossier annulé", key=f"add_ann_{SID}")
        ann_d  = s4.date_input("Date d'annulation", value=None, key=f"add_annd_{SID}")
        rfe    = s5.checkbox("RFE", key=f"add_rfe_{SID}")
        if rfe and not any([sent, acc, ref, ann]):
            st.warning("RFE doit être associé à un autre statut (envoyé/accepté/refusé/annulé).")

        if st.button("💾 Enregistrer le client", key=f"btn_add_{SID}"):
            if not nom:
                st.warning("Veuillez saisir le nom.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Choisir Catégorie et Sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            paye  = 0.0
            reste = max(0.0, total - paye)
            did   = make_client_id(nom, dt)
            dossier_n = next_dossier_number(df_live, start=13057)

            new_row = {
                "Dossier N": dossier_n,
                "ID_Client": did,
                "Nom": nom,
                "Date": dt,
                "Mois": f"{int(mois):02d}",
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": visa_final if visa_final else sel_sub,
                HONO: float(honor),
                AUTRE: float(other),
                TOTAL: total,
                "Payé": paye,
                "Reste": reste,
                "Paiements": [],
                "Options": opts_selected,
                "Notes": "",
                "Dossier envoyé": 1 if sent else 0,
                "Date d'envoi": sent_d if sent_d else (dt if sent else None),
                "Dossier accepté": 1 if acc else 0,
                "Date d'acceptation": acc_d,
                "Dossier refusé": 1 if ref else 0,
                "Date de refus": ref_d,
                "Dossier annulé": 1 if ann else 0,
                "Date d'annulation": ann_d,
                "RFE": 1 if rfe else 0,
                "EscrowTransfers": [],
            }
            df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            _write_clients(df_new, clients_path)
            st.success("Client ajouté.")
            st.cache_data.clear()
            st.rerun()

    # ---------- MODIFIER ----------
    elif op == "Modifier":
        if df_live.empty:
            st.info("Aucun client.")
        else:
            left, right = st.columns(2)
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
            sel_name = left.selectbox("Nom", [""]+names, index=0, key=f"mod_nom_{SID}")
            sel_id   = right.selectbox("ID_Client", [""]+ids, index=0, key=f"mod_id_{SID}")

            mask = None
            if sel_id:
                mask = (df_live["ID_Client"].astype(str) == sel_id)
            elif sel_name:
                mask = (df_live["Nom"].astype(str) == sel_name)

            if mask is None or not mask.any():
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()

            d1, d2, d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nomv_{SID}")
            dt   = d2.date_input("Date de création", value=_date_for_widget(row.get("Date"), date.today()), key=f"mod_date_{SID}")
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=_month_index(row.get("Mois")), key=f"mod_mois_{SID}")

            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""]+cats,
                                   index=_best_index(cats, preset_cat),
                                   key=f"mod_cat_{SID}")

            preset_sub = _safe_str(row.get("Sous-categorie",""))
            sel_sub = ""
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""]+subs,
                                       index=_best_index(subs, preset_sub),
                                       key=f"mod_sub_{SID}")

            # options existantes
            preset_opts = row.get("Options", [])
            if not isinstance(preset_opts, list):
                try:
                    preset_opts = json.loads(_safe_str(preset_opts) or "[]")
                    if not isinstance(preset_opts, list): preset_opts = []
                except Exception:
                    preset_opts = []

            opts_selected: List[str] = []
            if sel_cat and sel_sub:
                opt_list = visa_map[sel_cat][sel_sub].get("options", [])
                st.caption("Options disponibles (cases à cocher)")
                opts_selected = render_option_checkboxes(opt_list, keyprefix=f"mod_opts_{SID}", preselected=preset_opts)

            visa_final = compute_visa_string(sel_sub, opts_selected)

            f1, f2 = st.columns(2)
            honor = f1.number_input(HONO, min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0]),
                                    step=50.0, format="%.2f", key=f"mod_h_{SID}")
            other = f2.number_input(AUTRE, min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0]),
                                    step=20.0, format="%.2f", key=f"mod_o_{SID}")

            st.markdown("#### 📌 Statuts")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent   = s1.checkbox("Dossier envoyé", value=bool(row.get("Dossier envoyé")), key=f"mod_sent_{SID}")
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
            acc    = s2.checkbox("Dossier accepté", value=bool(row.get("Dossier accepté")), key=f"mod_acc_{SID}")
            acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
            ref    = s3.checkbox("Dossier refusé", value=bool(row.get("Dossier refusé")), key=f"mod_ref_{SID}")
            ref_d  = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=f"mod_refd_{SID}")
            ann    = s4.checkbox("Dossier annulé", value=bool(row.get("Dossier annulé")), key=f"mod_ann_{SID}")
            ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
            rfe    = s5.checkbox("RFE", value=bool(row.get("RFE")), key=f"mod_rfe_{SID}")
            if rfe and not any([sent, acc, ref, ann]):
                st.warning("RFE doit être associé à un autre statut (envoyé/accepté/refusé/annulé).")

            # Paiements (ajout)
            st.markdown("#### 💵 Paiements")
            pay_col1, pay_col2, pay_col3, pay_col4 = st.columns([1.2,1,1,1.2])
            pay_date = pay_col1.date_input("Date paiement", value=date.today(), key=f"mod_paydate_{SID}")
            pay_mode = pay_col2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"mod_paymode_{SID}")
            pay_amt  = pay_col3.number_input("Montant (US $)", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=f"mod_payamt_{SID}")
            add_pay  = pay_col4.button("➕ Ajouter paiement", key=f"mod_payadd_{SID}")

            if add_pay:
                lst = row.get("Paiements", [])
                if not isinstance(lst, list):
                    try:
                        lst = json.loads(_safe_str(lst) or "[]")
                        if not isinstance(lst, list): lst = []
                    except Exception:
                        lst = []
                if float(pay_amt) <= 0:
                    st.warning("Montant invalide.")
                    st.stop()
                lst.append({"date": _safe_str(pay_date), "mode": pay_mode, "montant": float(pay_amt)})
                df_live.at[idx, "Paiements"] = lst
                # recalcul Payé/Reste
                new_paid = sum([_to_float(p.get("montant",0.0)) for p in lst])
                df_live.at[idx, "Payé"] = new_paid
                df_live.at[idx, "Reste"] = max(0.0, (float(honor)+float(other)) - new_paid)

            # Enregistrer modifications
            if st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}"):
                if not nom:
                    st.warning("Le nom est requis.")
                    st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisir Catégorie et Sous-catégorie.")
                    st.stop()

                total = float(honor) + float(other)
                paye  = float(_safe_num_series(pd.DataFrame([df_live.loc[idx]]), "Payé").iloc[0])
                reste = max(0.0, total - paye)

                df_live.at[idx, "Nom"] = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(mois):02d}"
                df_live.at[idx, "Categorie"] = sel_cat
                df_live.at[idx, "Sous-categorie"] = sel_sub
                df_live.at[idx, "Visa"] = visa_final if visa_final else sel_sub
                df_live.at[idx, HONO] = float(honor)
                df_live.at[idx, AUTRE] = float(other)
                df_live.at[idx, TOTAL] = total
                df_live.at[idx, "Reste"] = reste
                df_live.at[idx, "Options"] = opts_selected

                df_live.at[idx, "Dossier envoyé"] = 1 if sent else 0
                df_live.at[idx, "Date d'envoi"] = sent_d
                df_live.at[idx, "Dossier accepté"] = 1 if acc else 0
                df_live.at[idx, "Date d'acceptation"] = acc_d
                df_live.at[idx, "Dossier refusé"] = 1 if ref else 0
                df_live.at[idx, "Date de refus"] = ref_d
                df_live.at[idx, "Dossier annulé"] = 1 if ann else 0
                df_live.at[idx, "Date d'annulation"] = ann_d
                df_live.at[idx, "RFE"] = 1 if rfe else 0

                _write_clients(df_live, clients_path)
                st.success("Modifications enregistrées.")
                st.cache_data.clear()
                st.rerun()

            # Historique paiements (lecture)
            st.markdown("#### 🧾 Historique des paiements")
            pay_list = row.get("Paiements", [])
            if not isinstance(pay_list, list):
                try:
                    pay_list = json.loads(_safe_str(pay_list) or "[]")
                    if not isinstance(pay_list, list): pay_list = []
                except Exception:
                    pay_list = []
            if pay_list:
                jdf = pd.DataFrame(pay_list)
                if "montant" in jdf.columns:
                    jdf["montant"] = jdf["montant"].apply(_to_float).apply(_fmt_money_us)
                if "date" in jdf.columns:
                    jdf["date"] = jdf["date"].astype(str)
                st.dataframe(jdf, use_container_width=True)
            else:
                st.caption("Aucun règlement saisi.")

    # ---------- SUPPRIMER ----------
    elif op == "Supprimer":
        if df_live.empty:
            st.info("Aucun client.")
        else:
            sel1, sel2 = st.columns(2)
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
            target_name = sel1.selectbox("Nom", [""]+names, index=0, key=f"del_nom_{SID}")
            target_id   = sel2.selectbox("ID_Client", [""]+ids, index=0, key=f"del_id_{SID}")

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({"Dossier N": row.get("Dossier N",""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                    df_new = df_live[~mask].copy()
                    _write_clients(df_new, clients_path)
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()

# =========================================================
# 📄 ONGLET : Visa (aperçu & édition)
# =========================================================
with tabs[4]:
    st.subheader("📄 Visa (aperçu & édition)")

    if not visa_path or not os.path.exists(visa_path):
        st.info("Charge un fichier Visa (ou un fichier unique avec onglet Visa).")
    else:
        dfv = read_visa_file(visa_path).copy()
        st.caption("Édite les colonnes d’options (1 = actif). Sauvegarde en bas.")
        edited = st.data_editor(dfv, use_container_width=True, key=f"visa_edit_{SID}")
        c1, c2 = st.columns([1,3])
        if c1.button("💾 Sauvegarder Visa", key=f"visa_save_{SID}"):
            try:
                # si fichier unique : réécrire les 2 onglets
                if clients_path == visa_path:
                    _write_clients(_read_clients(clients_path), clients_path)  # réécrit clients inchangé
                    # réécrit Visa aussi
                    with pd.ExcelWriter(visa_path, engine="openpyxl") as wr:
                        _read_clients(clients_path).to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                        edited.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                else:
                    # sinon, seulement le Visa
                    with pd.ExcelWriter(visa_path, engine="openpyxl") as wr:
                        edited.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                st.success("Visa sauvegardé.")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error("Erreur sauvegarde Visa : " + _safe_str(e))

        st.markdown("---")
        st.markdown("#### Aperçu filtres")
        # démonstration filtre cascade côté Visa
        cat_list = sorted([c for c in dfv.columns if _norm(c) in ("categorie","category")])
        sub_list = sorted([c for c in dfv.columns if "sous" in _norm(c)])
        cat_col = cat_list[0] if cat_list else "Categorie"
        sub_col = sub_list[0] if sub_list else "Sous-categorie"

        cats = sorted(dfv[cat_col].dropna().astype(str).unique().tolist()) if cat_col in dfv.columns else []
        cA, cB = st.columns(2)
        fcat = cA.selectbox("Catégorie", [""]+cats, index=0, key=f"v_cat_{SID}")
        fsub = ""
        if fcat:
            subs = sorted(dfv.loc[dfv[cat_col]==fcat, sub_col].dropna().astype(str).unique().tolist())
            fsub = cB.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"v_sub_{SID}")

        if fcat and fsub:
            row = dfv[(dfv[cat_col]==fcat) & (dfv[sub_col]==fsub)].head(1)
            if not row.empty:
                opt_cols = [c for c in dfv.columns if c not in (cat_col, sub_col)]
                opts = [c for c in opt_cols if _to_float(row.iloc[0].get(c, 0)) == 1]
                st.write("Options actives :", opts)
            else:
                st.info("Aucune ligne trouvée pour cette combinaison.")

# ====== FIN PARTIE 2 / 2 ======