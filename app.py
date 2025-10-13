# ===========================
# 🛂 Visa Manager — PARTIE 1/4
# ===========================
from __future__ import annotations

import os, json, re, zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple, Any

import pandas as pd
import streamlit as st


# ============================
# PARTIE 1 — Constantes & helpers & I/O
# ============================
import os, json, zipfile, unicodedata, re
from io import BytesIO
from datetime import date, datetime
import pandas as pd
import streamlit as st

# --- Noms de colonnes utilisés dans l’app
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"
DOSSIER_COL   = "Dossier N"
HONO          = "Montant honoraires (US $)"
AUTRE         = "Autres frais (US $)"
TOTAL         = "Total (US $)"

# --- Persistance des derniers chemins (fichier json local)
LAST_PATHS_FILE = ".cache_visamanager.json"

def _save_last_paths(clients_path: str|None, visa_path: str|None) -> None:
    try:
        data = {
            "clients_path": clients_path or "",
            "visa_path": visa_path or "",
        }
        with open(LAST_PATHS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _load_last_paths() -> tuple[str|None, str|None]:
    try:
        with open(LAST_PATHS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        cp = data.get("clients_path") or None
        vp = data.get("visa_path") or None
        if cp and not os.path.exists(cp): cp = None
        if vp and not os.path.exists(vp): vp = None
        return cp, vp
    except Exception:
        return None, None

# --- Normalisations/format
def _safe_str(x) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _safe_num_series(df: pd.DataFrame|pd.Series, col_or_series):
    s = df[col_or_series] if isinstance(df, pd.DataFrame) else df
    s = pd.to_numeric(s, errors="coerce")
    return s.fillna(0.0)

def _fmt_money_us(x: float|int) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(val):
    """Renvoie une date utilisable par st.date_input (ou None)."""
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return None
        return d.date()
    except Exception:
        return None

def _make_client_id(nom: str, d: date) -> str:
    base = re.sub(r"[^A-Za-z0-9]+", "-", _safe_str(nom)).strip("-").lower()
    if not isinstance(d, (date, datetime)):
        d = date.today()
    if isinstance(d, datetime):
        d = d.date()
    return f"{base}-{d:%Y%m%d}"

def _next_dossier(df_clients: pd.DataFrame, start: int = 13057) -> int:
    if DOSSIER_COL in df_clients.columns:
        vals = pd.to_numeric(df_clients[DOSSIER_COL], errors="coerce").dropna()
        if len(vals):
            return int(vals.max()) + 1
    return int(start)

# --- Lecture/écriture du fichier Clients (onglet "Clients")
def _read_clients(path: str|None) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        # DataFrame vide avec les colonnes attendues pour éviter toute erreur
        cols = [DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
                "Categorie", "Sous-categorie", "Visa",
                HONO, AUTRE, "Commentaires",
                TOTAL, "Payé", "Reste",
                "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"]
        return pd.DataFrame(columns=cols)
    try:
        # si le fichier a un onglet 'Clients', on le lit ; sinon on lit la première feuille
        xls = pd.ExcelFile(path)
        sh = SHEET_CLIENTS if SHEET_CLIENTS in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(path, sheet_name=sh)
        return df
    except Exception:
        return pd.DataFrame()

def _write_clients(df: pd.DataFrame, path: str|None) -> None:
    if not path:
        st.error("Aucun chemin de fichier Clients pour sauvegarder.")
        return
    try:
        # Écrit seulement l’onglet Clients (on ne touche pas au fichier Visa)
        with pd.ExcelWriter(path, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
    except Exception as e:
        st.error(f"Erreur écriture Clients: {e}")

# --- Lecture Visa brute (df_visa_raw doit exister plus bas)
def _read_visa(path: str|None) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        xls = pd.ExcelFile(path)
        sh = SHEET_VISA if SHEET_VISA in xls.sheet_names else xls.sheet_names[0]
        return pd.read_excel(path, sheet_name=sh)
    except Exception:
        return pd.DataFrame()

# --- Récupération des chemins depuis session_state ou depuis le fichier mémo
if "clients_path" not in st.session_state or "visa_path" not in st.session_state:
    last_c, last_v = _load_last_paths()
    st.session_state.setdefault("clients_path", last_c)
    st.session_state.setdefault("visa_path", last_v)

clients_path = st.session_state.get("clients_path")
visa_path    = st.session_state.get("visa_path")

# Si l’un manque, on tente de ne pas crasher : df_all / df_visa_raw deviennent vides
df_all = _read_clients(clients_path)
df_visa_raw = _read_visa(visa_path)

# ID de session pour clés uniques streamlit
SID = st.session_state.get("sid") or "S1"
st.session_state["sid"] = SID

# ---- Page & style ----
st.set_page_config(page_title="Visa Manager", page_icon="🛂", layout="wide")

st.markdown("""
<style>
.small-metrics .stMetric { padding: .25rem .5rem !important; }
.small-metrics .stMetric label, .small-metrics .stMetric span { font-size:.8rem !important; }
.compact-input .stTextInput input,
.compact-input .stNumberInput input,
.compact-input .stSelectbox div[data-baseweb="select"] { font-size:.85rem !important; height:2.0rem !important; }
.compact-textarea textarea { font-size:.9rem !important; }
.stButton>button { height:2.2rem }
</style>
""", unsafe_allow_html=True)

# ---- Constantes colonnes ----
DOSSIER_COL = "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"

SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

# ---- Persistance chemins récents ----
LAST_PATHS_FILE = ".cache_visamanager.json"

def _save_last_paths(clients_path: str|None, visa_path: str|None) -> None:
    try:
        data = {"clients_path": clients_path or "", "visa_path": visa_path or ""}
        with open(LAST_PATHS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _load_last_paths() -> tuple[str|None, str|None]:
    try:
        with open(LAST_PATHS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        cp = data.get("clients_path") or None
        vp = data.get("visa_path") or None
        if cp and not os.path.exists(cp): cp = None
        if vp and not os.path.exists(vp): vp = None
        return cp, vp
    except Exception:
        return None, None

# ---- Helpers safe ----
def _safe_str(x) -> str:
    try:
        if pd.isna(x): return ""
    except Exception:
        pass
    return str(x) if x is not None else ""

def _fmt_money_us(x: float) -> str:
    try:
        return "${:,.2f}".format(float(x))
    except Exception:
        return "$0.00"

def _safe_num_series(df: pd.DataFrame|pd.Series, col: str) -> pd.Series:
    """Force en float en nettoyant symboles/espaces; renvoie 0.0 si manquant."""
    if isinstance(df, pd.Series):
        s = df.copy()
    else:
        s = df.get(col, pd.Series([0.0]*len(df)))
    s = pd.to_numeric(
        pd.Series(s).astype(str)
        .str.replace(r"[^\d,\.\-]", "", regex=True)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0.0)
    return s

def _date_for_widget(val):
    """Retourne un objet date|None acceptable par st.date_input."""
    if isinstance(val, date): return val
    if isinstance(val, datetime): return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.notna(d): return d.date()
    except Exception:
        pass
    return None

def _norm_text(s: str) -> str:
    s = _safe_str(s)
    try:
        import unicodedata
        s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    except Exception:
        pass
    s = s.lower()
    s = s.replace("'", " ")
    s = re.sub(r"[^a-z0-9\+\/_\-\s]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

# ---- ID / Dossier ----
def _make_client_id(nom: str, d: date|datetime|str|None) -> str:
    base = _norm_text(nom).replace(" ", "-")
    if not base: base = "client"
    if isinstance(d, (date, datetime)):
        return f"{base}-{d:%Y%m%d}"
    d2 = _date_for_widget(d) or date.today()
    return f"{base}-{d2:%Y%m%d}"

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if df.empty or DOSSIER_COL not in df.columns:
        return start
    vals = pd.to_numeric(df[DOSSIER_COL], errors="coerce").dropna()
    if vals.empty: return start
    m = int(vals.max())
    return max(start, m+1)

# ---- Lecture/écriture Clients ----
@st.cache_data(show_spinner=False)
def read_clients(path: str) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        if path.lower().endswith(".xlsx"):
            try:
                # un seul fichier avec 2 onglets ?
                xls = pd.ExcelFile(path)
                if SHEET_CLIENTS in xls.sheet_names:
                    df = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
                else:
                    df = pd.read_excel(path)  # première feuille
            except Exception:
                df = pd.read_excel(path)
        else:
            return pd.DataFrame()
        return df
    except Exception:
        return pd.DataFrame()

def write_clients(df: pd.DataFrame, path: str) -> None:
    if not path:
        st.error("Chemin fichier Clients manquant.")
        return
    # si c'est un fichier bi-onglet, on préserve Visa si possible
    try:
        if os.path.exists(path):
            try:
                with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as wr:
                    pass
            except Exception:
                pass
        # On tente de lire l'autre onglet Visa si présent
        dfv = None
        try:
            xls = pd.ExcelFile(path)
            if SHEET_VISA in xls.sheet_names:
                dfv = pd.read_excel(path, sheet_name=SHEET_VISA)
        except Exception:
            dfv = None

        with pd.ExcelWriter(path, engine="openpyxl") as wr:
            df.to_excel(wr, index=False, sheet_name=SHEET_CLIENTS)
            if dfv is not None:
                dfv.to_excel(wr, index=False, sheet_name=SHEET_VISA)
    except Exception as e:
        st.error("Erreur d’écriture Clients : " + _safe_str(e))

# ---- Lecture Visa (structure + options via cases cochées = 1) ----
@st.cache_data(show_spinner=False)
def read_visa(path: str) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        # Accepte un fichier dédié ou l’onglet Visa d’un fichier unique
        xls = pd.ExcelFile(path)
        if SHEET_VISA in xls.sheet_names:
            df = pd.read_excel(path, sheet_name=SHEET_VISA)
        else:
            # si pas d’onglet "Visa", on lit la 1ère feuille
            df = pd.read_excel(path)
        return df
    except Exception:
        # fallback lecture simple
        try:
            return pd.read_excel(path)
        except Exception:
            return pd.DataFrame()

def _pick_col(df: pd.DataFrame, candidates: List[str]) -> str|None:
    cols = list(df.columns)
    norm_map = { _norm_text(c): c for c in cols }
    for cand in candidates:
        if cand in cols: return cand
        nc = _norm_text(cand)
        if nc in norm_map: return norm_map[nc]
    return None

def build_visa_map(df_visa_raw: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    Construit une structure:
    { 'Affaires/Tourisme': {
         'B-1': {'exclusive': ['COS','EOS'], 'options': ['... autres options selon ligne 1 = 1']},
         'B-2': {...}
      },
      'Etudiants': {
         'F-1': {'exclusive': ['COS','EOS'], 'options': [...]},
         'F-2': {...}
      },
      ...
    }
    Détection robuste des colonnes:
      - Catégorie = 'Categorie' / 'Category'
      - Sous-catégorie = 'Sous-categorie' / 'Sous-categories 1' / etc.
    Les options proviennent des en-têtes de colonnes (ligne 1) dont les cellules = 1 sur la ligne de la sous-catégorie.
    """
    if df_visa_raw is None or df_visa_raw.empty:
        return {}

    cat_col = _pick_col(df_visa_raw, ["Categorie", "Category"])
    sub_col = _pick_col(df_visa_raw, ["Sous-categorie", "Sous categories 1", "Sous-categories 1"])
    if sub_col is None:
        # essaye toute colonne qui ressemble à "Sous-categories X"
        poss = [c for c in df_visa_raw.columns if _norm_text(c).startswith("sous categories")]
        sub_col = poss[0] if poss else None

    if not cat_col or not sub_col:
        return {}

    # Les colonnes d'options = tout le reste sauf cat/sub, et on ne garde que celles qui ont des 1
    option_cols = [c for c in df_visa_raw.columns if c not in [cat_col, sub_col]]
    visa_map: Dict[str, Dict[str, Dict[str, Any]]] = {}

    for _, row in df_visa_raw.iterrows():
        cat = _safe_str(row.get(cat_col, "")).strip()
        sub = _safe_str(row.get(sub_col, "")).strip()
        if not cat or not sub:
            continue

        # options = entêtes où la cellule vaut 1 (case cochée)
        opts = []
        for oc in option_cols:
            val = row.get(oc, "")
            try:
                v = pd.to_numeric(val, errors="coerce")
            except Exception:
                v = pd.NA
            if pd.notna(v) and float(v) == 1.0:
                opts.append(_safe_str(oc).strip())

        # séparation COS/EOS si présent dans opts (exclusif)
        exclusive = []
        others = []
        for o in opts:
            o_norm = _norm_text(o)
            if o_norm in ("cos", "eos"):
                exclusive.append(o.strip())
            else:
                others.append(o.strip())

        if cat not in visa_map:
            visa_map[cat] = {}
        visa_map[cat][sub] = {
            "exclusive": exclusive,   # ex: ['COS','EOS'] (choix exclusif)
            "options": others         # autres cases multi-sélection
        }

    return visa_map

def apply_visa_selection_display(sel_cat: str, sel_sub: str, chosen_excl: str|None, chosen_multi: List[str]) -> str:
    """
    Construit le libellé Visa final: "<Sous-catégorie> <Exclusif>" si choisi,
    sinon "<Sous-catégorie>" seul.
    """
    base = _safe_str(sel_sub)
    if chosen_excl:
        return f"{base} {chosen_excl}"
    return base



# ===========================
# 🛂 Visa Manager — PARTIE 2/4
# ===========================

# ---- Chargement fichiers (1 ou 2 fichiers) ----
st.sidebar.header("📂 Fichiers")
mode = st.sidebar.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"], index=0)

last_clients, last_visa = _load_last_paths()

clients_path = st.sidebar.text_input("Chemin Clients (xlsx)", value=last_clients or "")
visa_path    = st.sidebar.text_input("Chemin Visa (xlsx)",    value=last_visa or "")

up_clients = None
up_visa    = None
up_single  = None

if mode == "Deux fichiers (Clients & Visa)":
    up_clients = st.sidebar.file_uploader("Clients (xlsx)", type=["xlsx"], key="up_clients")
    up_visa    = st.sidebar.file_uploader("Visa (xlsx)",    type=["xlsx"], key="up_visa")
    if up_clients is not None:
        # Sauvegarde dans un fichier temporaire persistant (en mémoire de session)
        clients_path = f"clients_{up_clients.name}"
        with open(clients_path, "wb") as f:
            f.write(up_clients.read())
    if up_visa is not None:
        visa_path = f"visa_{up_visa.name}"
        with open(visa_path, "wb") as f:
            f.write(up_visa.read())
else:
    up_single = st.sidebar.file_uploader("Fichier unique (2 onglets: Clients & Visa)", type=["xlsx"], key="up_single")
    if up_single is not None:
        # même fichier pour les deux
        single_path = f"single_{up_single.name}"
        with open(single_path, "wb") as f:
            f.write(up_single.read())
        clients_path = single_path
        visa_path    = single_path

# Mémoriser
_save_last_paths(clients_path, visa_path)

# ---- Lecture données ----
df_clients_raw = read_clients(clients_path) if clients_path else pd.DataFrame()
df_visa_raw    = read_visa(visa_path)       if visa_path    else pd.DataFrame()

# ---- Visa map (cat -> sub -> options) ----
visa_map = build_visa_map(df_visa_raw.copy()) if not df_visa_raw.empty else {}

# ---- Normalisation clients ----
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: 
        return pd.DataFrame()
    out = df.copy()

    # Colonnes minimales
    for c in [DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
              HONO, AUTRE, "Payé", "Reste", "Paiements", "Options",
              "Dossier envoyé", "Date d'envoi", "Dossier accepté", "Date d'acceptation",
              "Dossier refusé", "Date de refus", "Dossier annulé", "Date d'annulation",
              "RFE", "Commentaires autres frais", "Date", "Mois"]:
        if c not in out.columns:
            out[c] = pd.NA

    # Nombres
    for ncol in [HONO, AUTRE, "Payé", "Reste"]:
        out[ncol] = _safe_num_series(out, ncol)

    # Total
    if TOTAL not in out.columns:
        out[TOTAL] = out[HONO] + out[AUTRE]
    else:
        out[TOTAL] = _safe_num_series(out, TOTAL)
        # corrige si vide
        mask0 = out[TOTAL].isna() | (out[TOTAL] == 0)
        out.loc[mask0, TOTAL] = (out.loc[mask0, HONO] + out.loc[mask0, AUTRE])

    # Reste cohérent
    out["Reste"] = (out[TOTAL] - out["Payé"]).clip(lower=0.0)

    # Dates / Mois / Année
    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out["Mois"] = out["Mois"].apply(lambda x: f"{int(x):02d}" if _safe_str(x).strip().isdigit() else pd.NA)
    # Si Mois manquant, le déduire de la Date
    m_missing = out["Mois"].isna()
    out.loc[m_missing, "Mois"] = out.loc[m_missing, "Date"].dt.month.apply(lambda m: f"{int(m):02d}" if pd.notna(m) else pd.NA)
    out["_Année_"]   = out["Date"].dt.year.astype("Int64")
    out["_MoisNum_"] = pd.to_numeric(out["Mois"], errors="coerce").astype("Int64")

    # Nettoyage Options / Paiements (liste/dict)
    def _to_dict(x):
        if isinstance(x, dict): return x
        try:
            v = json.loads(_safe_str(x))
            return v if isinstance(v, dict) else {}
        except Exception:
            return {}
    def _to_list(x):
        if isinstance(x, list): return x
        try:
            v = json.loads(_safe_str(x))
            return v if isinstance(v, list) else []
        except Exception:
            return []

    out["Options"]   = out["Options"].apply(_to_dict)
    out["Paiements"] = out["Paiements"].apply(_to_list)

    # Booléens statut
    for c in ["Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"]:
        out[c] = out[c].apply(lambda v: 1 if str(v).strip() in ("1","True","true","TRUE","yes","oui") else 0)

    # Dates statut
    for c in ["Date d'envoi", "Date d'acceptation", "Date de refus", "Date d'annulation"]:
        out[c] = pd.to_datetime(out[c], errors="coerce")

    return out

df_all = normalize_clients(df_clients_raw)

# ---- Tabs ----
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "🧾 Gestion", "📄 Visa (aperçu)"])
SID = "vm"  # suffixe pour clés Streamlit uniques

# ==============================================
# 📊 ONGLET : Dashboard
# ==============================================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Charge un fichier Clients & Visa pour commencer.")
    else:
        years  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1,13)]
        cats   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        f1, f2, f3, f4, f5 = st.columns(5)
        dash_years = f1.multiselect("Année", years, default=[], key=f"dash_years_{SID}")
        dash_month = f2.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{SID}")
        dash_cats  = f3.multiselect("Catégorie", cats, default=[], key=f"dash_cats_{SID}")
        dash_subs  = f4.multiselect("Sous-catégorie", subs, default=[], key=f"dash_subs_{SID}")
        dash_visas = f5.multiselect("Visa", visas, default=[], key=f"dash_visas_{SID}")

        view = df_all.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in view.columns:
                view[c] = _safe_num_series(view, c)

        if dash_years: view = view[view["_Année_"].isin(dash_years)]
        if dash_month: view = view[view["Mois"].astype(str).isin(dash_month)]
        if dash_cats:  view = view[view["Categorie"].astype(str).isin(dash_cats)]
        if dash_subs:  view = view[view["Sous-categorie"].astype(str).isin(dash_subs)]
        if dash_visas: view = view[view["Visa"].astype(str).isin(dash_visas)]

        # KPI compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(view)}")
        k2.metric("Honoraires", _fmt_money_us(float(view[HONO].sum())))
        k3.metric("Payé",      _fmt_money_us(float(view["Payé"].sum())))
        k4.metric("Reste",     _fmt_money_us(float(view["Reste"].sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        # Tableau
        show_cols = [c for c in [
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
            "Date", "Mois", HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"
        ] if c in view.columns]

        sort_cols = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_cols) if sort_cols else view
        view_sorted = view_sorted.loc[:, ~view_sorted.columns.duplicated()].copy()

        st.dataframe(view_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_tbl_{SID}")



# ===========================
# 🛂 Visa Manager — PARTIE 3/4
# ===========================

# ==============================================
# 📈 ONGLET : Analyses
# ==============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1,13)]
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
        for coln in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if coln in dfA.columns:
                dfA[coln] = _safe_num_series(dfA, coln)

        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money_us(float(dfA.get(HONO, pd.Series(dtype=float)).sum())))
        k3.metric("Payé",      _fmt_money_us(float(dfA.get("Payé", pd.Series(dtype=float)).sum())))
        k4.metric("Reste",     _fmt_money_us(float(dfA.get("Reste", pd.Series(dtype=float)).sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("#### Répartition & %")
        cL, cR = st.columns(2)
        if not dfA.empty and "Categorie" in dfA.columns:
            vc = (dfA.groupby("Categorie", as_index=False)
                        .agg(N=("Categorie","size"), Honoraires=(HONO,"sum")))
            vc["% Dossiers"] = (vc["N"]/(vc["N"].sum() or 1)*100).round(1)
            cL.dataframe(vc.sort_values("N", ascending=False), use_container_width=True, height=260)
        if not dfA.empty and "Sous-categorie" in dfA.columns:
            vs = (dfA.groupby("Sous-categorie", as_index=False)
                        .agg(N=("Sous-categorie","size"), Honoraires=(HONO,"sum")))
            vs["% Dossiers"] = (vs["N"]/(vs["N"].sum() or 1)*100).round(1)
            cR.dataframe(vs.sort_values("N", ascending=False).head(25), use_container_width=True, height=260)

        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("#### 📊 Dossiers par catégorie")
            g1 = (dfA.groupby("Categorie", as_index=False).size().rename(columns={"size":"Nombre"}))
            st.bar_chart(g1.set_index("Categorie"))

        if not dfA.empty and "Mois" in dfA.columns and HONO in dfA.columns:
            st.markdown("#### 📈 Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = (tmp.groupby("Mois", as_index=False)[HONO].sum()
                       .reindex([f"{m:02d}" for m in range(1,13)], fill_value=0)
                       .sort_values("Mois"))
            st.line_chart(gm.set_index("Mois"))

        st.markdown("#### 🔁 Comparaison de périodes (A vs B)")
        ca1, ca2, cb1, cb2 = st.columns(4)
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=f"cmp_ya_{SID}")
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=f"cmp_ma_{SID}")
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=f"cmp_yb_{SID}")
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=f"cmp_mb_{SID}")

        def _filter_period(base, ys, ms):
            d = base.copy()
            for coln in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                if coln in d.columns:
                    d[coln] = _safe_num_series(d, coln)
            if ys: d = d[d["_Année_"].isin(ys)]
            if ms: d = d[d["Mois"].astype(str).isin(ms)]
            return d

        dfA_A = _filter_period(df_all, pa_years, pa_month)
        dfA_B = _filter_period(df_all, pb_years, pb_month)

        cpa, cpb = st.columns(2)
        cpa.metric("A — Dossiers", f"{len(dfA_A)}")
        cpa.metric("A — Honoraires", _fmt_money_us(float(dfA_A.get(HONO, pd.Series(dtype=float)).sum())))
        cpb.metric("B — Dossiers", f"{len(dfA_B)}")
        cpb.metric("B — Honoraires", _fmt_money_us(float(dfA_B.get(HONO, pd.Series(dtype=float)).sum())))

        st.markdown("##### Comparaison par mois")
        def _mk_month_series(d):
            if d.empty or "Mois" not in d.columns or HONO not in d.columns:
                return pd.DataFrame({"Mois": [], "Honoraires": []})
            t = d.copy()
            t["Mois"] = t["Mois"].astype(str)
            return (t.groupby("Mois", as_index=False)[HONO].sum()
                     .reindex([f"{m:02d}" for m in range(1,13)], fill_value=0)
                     .rename(columns={HONO:"Honoraires"}))

        A = _mk_month_series(dfA_A); A["Période"] = "A"
        B = _mk_month_series(dfA_B); B["Période"] = "B"
        comp = pd.concat([A,B], ignore_index=True)
        if not comp.empty:
            wide = comp.pivot_table(index="Mois", columns="Période", values="Honoraires", fill_value=0)
            st.bar_chart(wide)

        st.markdown("#### 🧾 Détails des dossiers filtrés")
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
            DOSSIER_COL,"ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in det.columns]
        sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_cols) if sort_cols else det
        det_sorted = det_sorted.loc[:, ~det_sorted.columns.duplicated()].copy()
        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"a_tbl_{SID}")

# ==============================================
# 🏦 ONGLET : Escrow
# ==============================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        for coln in [TOTAL, "Payé", "Reste"]:
            if coln in dfE.columns:
                dfE[coln] = _safe_num_series(dfE, coln)

        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money_us(float(dfE.get(TOTAL, pd.Series(dtype=float)).sum())))
        t2.metric("Payé",         _fmt_money_us(float(dfE.get("Payé", pd.Series(dtype=float)).sum())))
        t3.metric("Reste",        _fmt_money_us(float(dfE.get("Reste", pd.Series(dtype=float)).sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        agg = dfE.groupby("Categorie", as_index=False)[[c for c in [TOTAL,"Payé","Reste"] if c in dfE.columns]].sum()
        if TOTAL in agg.columns and "Payé" in agg.columns:
            agg["% Payé"] = ((agg["Payé"]/agg[TOTAL]).fillna(0.0)*100).round(1)
        st.dataframe(agg.sort_values(by=TOTAL if TOTAL in agg.columns else agg.columns[0], ascending=False),
                     use_container_width=True)

# ==============================================
# 📄 ONGLET : Visa (aperçu)
# ==============================================
with tabs[5]:
    st.subheader("📄 Visa (aperçu)")
    if not visa_map:
        st.info("Aucune structure Visa détectée.")
    else:
        cat_list = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""]+cat_list, index=0, key=f"v_cat_{SID}")
        if sel_cat:
            sub_list = sorted(list(visa_map[sel_cat].keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""]+sub_list, index=0, key=f"v_sub_{SID}")
            if sel_sub:
                opts = visa_map[sel_cat][sel_sub]
                excl = opts.get("exclusive", [])
                others = opts.get("options", [])
                st.write("**Options exclusives (au plus une)** :", ", ".join(excl) if excl else "—")
                st.write("**Autres options** :", ", ".join(others) if others else "—")



# ==============================================================
# 🔍  Sécurisation : détection dynamique des colonnes dans Visa
# ==============================================================
import unicodedata

def _normcol(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s = s.strip().lower().replace("-", " ").replace("_", " ")
    s = " ".join(s.split())
    return s

def pick_col(df, candidates):
    cols = list(df.columns)
    norm_map = {_normcol(c): c for c in cols}
    for cand in candidates:
        if cand in cols:
            return cand
        nc = _normcol(cand)
        if nc in norm_map:
            return norm_map[nc]
    return None

# Applique la détection au dataframe Visa chargé
cat_col  = pick_col(df_visa_raw, ["Categorie", "Category"])
sub_col  = pick_col(df_visa_raw, ["Sous-categorie", "Sous categorie", "Sous-categories 1", "Sous categories 1"])
visa_col = pick_col(df_visa_raw, ["Visa", "Type visa"])
if sub_col is None:
    poss = [c for c in df_visa_raw.columns if _normcol(c).startswith("sous categories")]
    sub_col = poss[0] if poss else None

cats  = sorted(df_visa_raw[cat_col].dropna().astype(str).unique().tolist()) if cat_col  else []
subs  = sorted(df_visa_raw[sub_col].dropna().astype(str).unique().tolist()) if sub_col  else []
visas = sorted(df_visa_raw[visa_col].dropna().astype(str).unique().tolist()) if visa_col else []


# ==============================================================
# 💾  Gestion CRUD Clients + Export global (Clients + Visa)
# ==============================================================
st.markdown("---")
st.subheader("🧾 Gestion des clients")

op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_op_{SID}")

df_live = _read_clients(clients_path)

# ---------------- AJOUT ----------------
if op == "Ajouter":
    st.markdown("### ➕ Ajouter un client")
    c1, c2, c3 = st.columns(3)
    nom  = c1.text_input("Nom", key=f"add_nom_{SID}")
    dt   = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
    mois = c3.selectbox("Mois", [f"{m:02d}" for m in range(1,13)],
                        index=int(date.today().month)-1, key=f"add_mois_{SID}")

    # Choix Visa
    st.markdown("#### 🎯 Choix Visa")
    sel_cat = st.selectbox("Catégorie", [""] + cats, key=f"add_cat_{SID}")
    sel_sub = st.selectbox("Sous-catégorie", [""] + subs, key=f"add_sub_{SID}")
    visa_final = st.selectbox("Visa", [""] + visas, key=f"add_visa_{SID}")

    f1, f2 = st.columns(2)
    honor = f1.number_input("Montant honoraires (US $)", 0.0, step=50.0, key=f"add_h_{SID}")
    other = f2.number_input("Autres frais (US $)", 0.0, step=20.0, key=f"add_o_{SID}")
    comment = st.text_area("Commentaires / détails autres frais", key=f"add_comm_{SID}")

    st.markdown("#### 📌 Statuts initiaux")
    s1, s2, s3, s4, s5 = st.columns(5)
    sent = s1.checkbox("Envoyé", key=f"add_sent_{SID}")
    acc  = s2.checkbox("Accepté", key=f"add_acc_{SID}")
    ref  = s3.checkbox("Refusé", key=f"add_ref_{SID}")
    ann  = s4.checkbox("Annulé", key=f"add_ann_{SID}")
    rfe  = s5.checkbox("RFE", key=f"add_rfe_{SID}")

    if rfe and not any([sent, acc, ref, ann]):
        st.warning("⚠️ RFE doit être associé à un statut envoyé / accepté / refusé / annulé.")

    if st.button("💾 Enregistrer le client", key=f"add_btn_{SID}"):
        if not nom:
            st.warning("Nom requis.")
            st.stop()
        total = honor + other
        reste = total
        did = _make_client_id(nom, dt)
        dossier_n = _next_dossier(df_live, start=13057)
        new_row = {
            "Dossier N": dossier_n,
            "ID_Client": did,
            "Nom": nom,
            "Date": dt,
            "Mois": mois,
            "Categorie": sel_cat,
            "Sous-categorie": sel_sub,
            "Visa": visa_final,
            "Montant honoraires (US $)": honor,
            "Autres frais (US $)": other,
            "Commentaires": comment,
            "Total (US $)": total,
            "Payé": 0.0,
            "Reste": reste,
            "Dossier envoyé": 1 if sent else 0,
            "Dossier accepté": 1 if acc else 0,
            "Dossier refusé": 1 if ref else 0,
            "Dossier annulé": 1 if ann else 0,
            "RFE": 1 if rfe else 0,
        }
        df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
        _write_clients(df_new, clients_path)
        st.success("Client ajouté ✅")
        st.cache_data.clear()
        st.rerun()

# ---------------- MODIFICATION ----------------
elif op == "Modifier":
    st.markdown("### ✏️ Modifier un client")
    if df_live.empty:
        st.info("Aucun client enregistré.")
    else:
        noms = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
        sel = st.selectbox("Choisir le client", [""] + noms, key=f"mod_sel_{SID}")
        if sel:
            idx = df_live[df_live["Nom"] == sel].index[0]
            row = df_live.loc[idx]
            c1, c2, c3 = st.columns(3)
            nom  = c1.text_input("Nom", row["Nom"], key=f"mod_nom_{SID}")
            dt   = c2.date_input("Date", _date_for_widget(row["Date"]), key=f"mod_date_{SID}")
            mois = c3.selectbox("Mois", [f"{m:02d}" for m in range(1,13)],
                                index=int(str(row["Mois"]))-1, key=f"mod_mois_{SID}")

            st.markdown("#### 🎯 Choix Visa")
            sel_cat = st.selectbox("Catégorie", [""] + cats,
                                   index=(cats.index(row["Categorie"])+1 if row["Categorie"] in cats else 0),
                                   key=f"mod_cat_{SID}")
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs,
                                   index=(subs.index(row["Sous-categorie"])+1 if row["Sous-categorie"] in subs else 0),
                                   key=f"mod_sub_{SID}")
            visa_final = st.selectbox("Visa", [""] + visas,
                                   index=(visas.index(row["Visa"])+1 if row["Visa"] in visas else 0),
                                   key=f"mod_visa_{SID}")

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", 0.0,
                                    value=float(row["Montant honoraires (US $)"]), step=50.0, key=f"mod_h_{SID}")
            other = f2.number_input("Autres frais (US $)", 0.0,
                                    value=float(row["Autres frais (US $)"]), step=20.0, key=f"mod_o_{SID}")
            comment = st.text_area("Commentaires / détails autres frais", value=_safe_str(row.get("Commentaires","")),
                                   key=f"mod_comm_{SID}")

            s1, s2, s3, s4, s5 = st.columns(5)
            sent = s1.checkbox("Envoyé", value=bool(row["Dossier envoyé"]), key=f"mod_sent_{SID}")
            acc  = s2.checkbox("Accepté", value=bool(row["Dossier accepté"]), key=f"mod_acc_{SID}")
            ref  = s3.checkbox("Refusé", value=bool(row["Dossier refusé"]), key=f"mod_ref_{SID}")
            ann  = s4.checkbox("Annulé", value=bool(row["Dossier annulé"]), key=f"mod_ann_{SID}")
            rfe  = s5.checkbox("RFE", value=bool(row["RFE"]), key=f"mod_rfe_{SID}")

            if st.button("💾 Sauvegarder", key=f"mod_save_{SID}"):
                total = honor + other
                reste = total - float(row.get("Payé", 0))
                for k,v in {
                    "Nom": nom, "Date": dt, "Mois": mois,
                    "Categorie": sel_cat, "Sous-categorie": sel_sub, "Visa": visa_final,
                    "Montant honoraires (US $)": honor, "Autres frais (US $)": other,
                    "Commentaires": comment, "Total (US $)": total, "Reste": reste,
                    "Dossier envoyé": 1 if sent else 0, "Dossier accepté": 1 if acc else 0,
                    "Dossier refusé": 1 if ref else 0, "Dossier annulé": 1 if ann else 0, "RFE": 1 if rfe else 0
                }.items():
                    df_live.at[idx, k] = v
                _write_clients(df_live, clients_path)
                st.success("Client modifié ✅")
                st.cache_data.clear()
                st.rerun()

# ---------------- SUPPRESSION ----------------
elif op == "Supprimer":
    st.markdown("### 🗑️ Supprimer un client")
    if df_live.empty:
        st.info("Aucun client.")
    else:
        noms = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
        sel = st.selectbox("Choisir le client", [""] + noms, key=f"del_sel_{SID}")
        if sel:
            if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                df_new = df_live[df_live["Nom"] != sel]
                _write_clients(df_new, clients_path)
                st.success("Client supprimé.")
                st.cache_data.clear()
                st.rerun()


# ==============================================================
# 💾  Export global ZIP (Clients + Visa)
# ==============================================================
st.markdown("---")
st.subheader("💾 Export global")
if st.button("Préparer archive ZIP", key=f"zip_{SID}"):
    try:
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            with BytesIO() as xbuf:
                with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                    df_live.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                zf.writestr("Clients.xlsx", xbuf.getvalue())
            zf.write(visa_path, "Visa.xlsx")
        st.session_state[f"zip_export_{SID}"] = buf.getvalue()
        st.success("Archive prête ✅")
    except Exception as e:
        st.error(f"Erreur export : {e}")

if st.session_state.get(f"zip_export_{SID}"):
    st.download_button(
        "⬇️ Télécharger Export (ZIP)",
        data=st.session_state[f"zip_export_{SID}"],
        file_name="Export_Visa_Manager.zip",
        mime="application/zip",
        key=f"zip_dl_{SID}"
    )