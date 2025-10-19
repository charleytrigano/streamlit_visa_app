# Visa Manager - app.py
# Streamlit app with robust Visa mapping: category -> sous-categories -> checkbox options
# Improved normalization and robust lookup for sous-categorie -> checkbox headers.
#
# Usage: streamlit run app.py
# Requirements: pandas, openpyxl

import os
import json
import re
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional

import pandas as pd
import numpy as np
import streamlit as st

# Try plotly (optional)
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    px = None
    HAS_PLOTLY = False

# -------------------------
# Config
# -------------------------
APP_TITLE = "🛂 Visa Manager"
COLS_CLIENTS = [
    "ID_Client", "Dossier N", "Nom", "Date",
    "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)",
    "Payé", "Solde", "Acompte 1", "Acompte 2",
    "RFE", "Dossiers envoyé", "Dossier approuvé",
    "Dossier refusé", "Dossier Annulé", "Commentaires"
]
MEMO_FILE = "_vmemory.json"
SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"
SID = "vmgr"
DEFAULT_START_CLIENT_ID = 13057

def skey(*parts: str) -> str:
    return f"{SID}_" + "_".join([p for p in parts if p])

# -------------------------
# Helpers (normalization, parsing)
# -------------------------
def normalize_header_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r'^\s+|\s+$', '', s)
    s = re.sub(r"\s+", " ", s)
    return s

def remove_accents(s: str) -> str:
    # simple replacements for common accents (sufficient for headers in French)
    if s is None:
        return ""
    s2 = str(s)
    s2 = s2.replace("é", "e").replace("è", "e").replace("ê", "e").replace("ë", "e")
    s2 = s2.replace("à", "a").replace("â", "a")
    s2 = s2.replace("î", "i").replace("ï", "i")
    s2 = s2.replace("ô", "o").replace("ö", "o")
    s2 = s2.replace("ù", "u").replace("û", "u").replace("ü", "u")
    s2 = s2.replace("ç", "c")
    return s2

def canonical_key(s: str) -> str:
    if s is None:
        return ""
    s2 = normalize_header_text(str(s)).lower()
    s2 = remove_accents(s2)
    s2 = re.sub(r"[^a-z0-9 ]", " ", s2)
    s2 = re.sub(r"\s+", " ", s2).strip()
    return s2

def money_to_float(x: Any) -> float:
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s == "" or s in ["-", "—", "–", "NA", "N/A"]:
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", s)
    if s == "":
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        if "," in s and s.count(",") == 1:
            if len(s.split(",")[-1]) == 2:
                s = s.replace(",", ".")
            else:
                s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        try:
            return float(re.sub(r"[^0-9.\-]", "", s))
        except Exception:
            return 0.0

# -------------------------
# Visa mapping builder
# -------------------------
def build_visa_map(dfv: pd.DataFrame) -> Dict[str, List[str]]:
    """
    Build category -> list of sous-categories mapping from Visa sheet.
    """
    vm: Dict[str, List[str]] = {}
    if dfv is None or dfv.empty:
        return vm
    df = dfv.copy()
    # Accept 'Categorie' variant
    if "Categories" not in df.columns and "Categorie" in df.columns:
        df = df.rename(columns={"Categorie": "Categories"})
    for _, row in df.iterrows():
        cat = str(row.get("Categories", "")).strip()
        sub = str(row.get("Sous-categorie", "")).strip()
        if not cat:
            continue
        vm.setdefault(cat, [])
        if sub and sub not in vm[cat]:
            vm[cat].append(sub)
    return vm

def build_sub_options_map_from_flags(dfv: pd.DataFrame) -> Dict[str, List[str]]:
    """
    For each row (Category, Sous-categorie), scan all other columns.
    If a cell is truthy (1, '1', 'x', 'oui', etc.), associate the column header as an option for that sous-categorie.
    Returns dict: sous_norm (lower, canonical) -> [option labels]
    """
    out: Dict[str, List[str]] = {}
    if dfv is None or dfv.empty:
        return out
    df = dfv.copy()
    # Normalize column names using current heuristic mapping if needed
    # We'll keep option labels as the header text exactly (trimmed).
    cols_to_skip = set(["Categories", "Categorie", "Sous-categorie"])
    # Determine columns that are possible flags
    cols_to_check = [c for c in df.columns if c not in cols_to_skip]
    for _, row in df.iterrows():
        sub_raw = str(row.get("Sous-categorie", "")).strip()
        if not sub_raw:
            continue
        sub_norm = canonical_key(sub_raw)
        for col in cols_to_check:
            val = row.get(col, "")
            truthy = False
            if pd.isna(val):
                truthy = False
            else:
                sval = str(val).strip().lower()
                if sval in ("1", "x", "t", "true", "oui", "yes", "y"):
                    truthy = True
                else:
                    # numeric values
                    try:
                        if float(sval) == 1.0:
                            truthy = True
                    except Exception:
                        truthy = False
            if truthy:
                label = str(col).strip()
                out.setdefault(sub_norm, [])
                if label not in out[sub_norm]:
                    out[sub_norm].append(label)
    return out

# -------------------------
# small UI helpers
# -------------------------
def debug_show(df, name="Data"):
    try:
        st.sidebar.markdown(f"**DEBUG — {name}**")
        if isinstance(df, dict):
            st.sidebar.write(df)
        else:
            st.sidebar.write(type(df))
    except Exception:
        pass

# -------------------------
# Reading helpers (robust)
# -------------------------
def try_read_excel_from_bytes(b: bytes, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
    bio = BytesIO(b)
    try:
        xls = pd.ExcelFile(bio, engine="openpyxl")
        sheets = xls.sheet_names
        # choose sheet logic
        if sheet_name and sheet_name in sheets:
            return pd.read_excel(BytesIO(b), sheet_name=sheet_name, engine="openpyxl")
        for cand in [SHEET_VISA, SHEET_CLIENTS, "Sheet1"]:
            if cand in sheets:
                try:
                    return pd.read_excel(BytesIO(b), sheet_name=cand, engine="openpyxl")
                except Exception:
                    continue
        # fallback first sheet
        return pd.read_excel(BytesIO(b), sheet_name=0, engine="openpyxl")
    except Exception:
        return None

def read_any_table(src: Any, sheet: Optional[str] = None, debug_prefix: str = "") -> Optional[pd.DataFrame]:
    if src is None:
        return None
    try:
        # file-like
        if hasattr(src, "read"):
            b = None
            try:
                b = src.getvalue()
            except Exception:
                try:
                    src.seek(0); b = src.read()
                except Exception:
                    b = None
            if b:
                df = try_read_excel_from_bytes(b, sheet)
                if df is not None:
                    return df
                try:
                    return pd.read_csv(BytesIO(b), sep=";", encoding="utf-8", on_bad_lines="skip")
                except Exception:
                    try:
                        return pd.read_csv(BytesIO(b), sep=",", encoding="utf-8", on_bad_lines="skip")
                    except Exception:
                        return None
        # path
        if isinstance(src, (str, os.PathLike)):
            p = str(src)
            if not os.path.exists(p):
                return None
            if p.lower().endswith(".csv"):
                try:
                    return pd.read_csv(p, sep=";", encoding="utf-8", on_bad_lines="skip")
                except Exception:
                    return pd.read_csv(p, sep=",", encoding="utf-8", on_bad_lines="skip")
            else:
                try:
                    return pd.read_excel(p, sheet_name=sheet or 0, engine="openpyxl")
                except Exception:
                    return None
    except Exception:
        return None
    return None

# -------------------------
# App start
# -------------------------
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title(APP_TITLE)

# Sidebar files input
st.sidebar.header("📂 Fichiers")
last_clients, last_visa, last_save_dir = ("", "", "")
try:
    if os.path.exists(MEMO_FILE):
        with open(MEMO_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
            last_clients = d.get("clients","")
            last_visa = d.get("visa","")
            last_save_dir = d.get("save_dir","")
except Exception:
    pass

mode = st.sidebar.radio("Mode de chargement", ["Un fichier (Clients)", "Deux fichiers (Clients & Visa)"], index=0, key=skey("mode"))
up_clients = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx","xls","csv"], key=skey("up_clients"))
up_visa = None
if mode == "Deux fichiers (Clients & Visa)":
    up_visa = st.sidebar.file_uploader("Visa (xlsx/csv)", type=["xlsx","xls","csv"], key=skey("up_visa"))

clients_path_in = st.sidebar.text_input("ou chemin local Clients (laisser vide si upload)", value=last_clients, key=skey("cli_path"))
visa_path_in = st.sidebar.text_input("ou chemin local Visa (laisser vide si upload)", value=(last_visa if mode!="Un fichier (Clients)" else ""), key=skey("vis_path"))
save_dir_in = st.sidebar.text_input("Dossier de sauvegarde (optionnel)", value=last_save_dir, key=skey("save_dir"))

if st.sidebar.button("📥 Sauvegarder chemins", key=skey("btn_save_paths")):
    try:
        with open(MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump({"clients": clients_path_in or "", "visa": visa_path_in or "", "save_dir": save_dir_in or ""}, f, ensure_ascii=False, indent=2)
        st.sidebar.success("Chemins sauvegardés.")
    except Exception:
        st.sidebar.error("Impossible de sauvegarder.")

# read files
clients_src = None
visa_src = None
if up_clients is not None:
    try:
        clients_bytes = up_clients.getvalue()
        clients_src = BytesIO(clients_bytes)
    except Exception:
        clients_src = up_clients
elif clients_path_in:
    clients_src = clients_path_in
elif last_clients:
    clients_src = last_clients

if mode == "Deux fichiers (Clients & Visa)":
    if up_visa is not None:
        try:
            visa_bytes = up_visa.getvalue()
            visa_src = BytesIO(visa_bytes)
        except Exception:
            visa_src = up_visa
    elif visa_path_in:
        visa_src = visa_path_in
    elif last_visa:
        visa_src = last_visa
else:
    visa_src = clients_src

df_clients_raw = None
df_visa_raw = None
try:
    df_clients_raw = read_any_table(clients_src, sheet=SHEET_CLIENTS, debug_prefix="[Clients] ")
except Exception:
    df_clients_raw = None
if df_clients_raw is None and clients_src is not None:
    df_clients_raw = read_any_table(clients_src, sheet=None, debug_prefix="[Clients fallback] ")

try:
    df_visa_raw = read_any_table(visa_src, sheet=SHEET_VISA, debug_prefix="[Visa] ")
except Exception:
    df_visa_raw = None
if df_visa_raw is None and visa_src is not None:
    df_visa_raw = read_any_table(visa_src, sheet=None, debug_prefix="[Visa fallback] ")

if df_clients_raw is None:
    df_clients_raw = pd.DataFrame()
if df_visa_raw is None:
    df_visa_raw = pd.DataFrame()

# show raw preview debug
if isinstance(df_clients_raw, pd.DataFrame) and not df_clients_raw.empty:
    debug_show(df_clients_raw.head(5).to_dict(), "Clients raw preview")
if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
    debug_show(df_visa_raw.head(8).to_dict(), "Visa raw preview")

# -------------------------
# Build visa maps (categories, sous-categories, per-sous options)
# -------------------------
visa_map = {}
visa_map_norm = {}
visa_categories = []
visa_sub_options_map = {}  # key: canonical_key(sous) -> list of option labels

if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
    try:
        # heuristic column mapping first (to ensure "Categories" and "Sous-categorie" recognized)
        df_visa_mapped, _ = map_columns_heuristic(df_visa_raw)
        try:
            # coerce names if needed
            df_visa_mapped = (df_visa_mapped.rename(columns={c:c for c in df_visa_mapped.columns}))
            df_visa_mapped = df_visa_mapped.copy()
        except Exception:
            pass
        # ensure we have desired columns names
        if "Categories" not in df_visa_mapped.columns and "Categorie" in df_visa_mapped.columns:
            df_visa_mapped = df_visa_mapped.rename(columns={"Categorie":"Categories"})
        # build basic maps
        raw_vm = build_visa_map(df_visa_mapped)
        visa_map = {k.strip(): [s.strip() for s in v] for k, v in raw_vm.items()}
        visa_map_norm = {canonical_key(k): v for k, v in visa_map.items()}
        visa_categories = sorted(list(visa_map.keys()))
        # build per-sub options map by scanning flag columns per row
        visa_sub_options_map = build_sub_options_map_from_flags(df_visa_mapped)
    except Exception as e:
        st.sidebar.error(f"Erreur build visa maps: {e}")
        visa_map = {}
        visa_map_norm = {}
        visa_categories = []
        visa_sub_options_map = {}

# Debug show maps
st.sidebar.markdown("DEBUG visa_map_norm (category_key -> subs):")
try:
    st.sidebar.write(visa_map_norm)
except Exception:
    pass
st.sidebar.markdown("DEBUG visa_sub_options_map (sous_key -> option headers):")
try:
    st.sidebar.write(visa_sub_options_map)
except Exception:
    pass

# -------------------------
# Helpers to lookup options for a sous-categorie robustly
# -------------------------
def get_sub_options_for(sub_value: str) -> List[str]:
    """
    Robust lookup:
    - try canonical_key(sub_value)
    - try lower/strip
    - try remove_accents + lower
    - try contains match among keys
    Returns the list of option labels (exact header text) or [].
    """
    if not sub_value or not isinstance(sub_value, str):
        return []
    tried = []
    candidates = []
    s_raw = sub_value.strip()
    s1 = canonical_key(s_raw)
    s2 = s_raw.strip().lower()
    s3 = remove_accents(s_raw).strip().lower()
    # try direct canonical key
    tried.append(("canonical", s1))
    if s1 in visa_sub_options_map:
        return visa_sub_options_map[s1][:]
    # try s2
    tried.append(("lower", s2))
    if s2 in visa_sub_options_map:
        return visa_sub_options_map[s2][:]
    # try s3
    tried.append(("no_accents", s3))
    if s3 in visa_sub_options_map:
        return visa_sub_options_map[s3][:]
    # try canonical of stored keys: iterate keys and compare canonical forms
    for k in list(visa_sub_options_map.keys()):
        if s1 == k:
            return visa_sub_options_map[k][:]
    # fallback: contains match (case-insensitive, accent-insensitive)
    s_match = remove_accents(s_raw).strip().lower()
    for k in visa_sub_options_map.keys():
        if s_match in remove_accents(k).lower() or remove_accents(k).lower() in s_match:
            candidates.extend(visa_sub_options_map.get(k, []))
    if candidates:
        # unique preserve order
        seen = set()
        out = []
        for c in candidates:
            if c not in seen:
                seen.add(c); out.append(c)
        tried.append(("contains_match", "found"))
        return out
    # debug: store tries in sidebar for investigation
    st.sidebar.markdown("DEBUG get_sub_options_for tries:")
    try:
        st.sidebar.write(tried)
    except Exception:
        pass
    return []

# -------------------------
# Normalize clients and prepare live DF
# -------------------------
def normalize_clients_for_live(df_clients_raw: pd.DataFrame) -> pd.DataFrame:
    if df_clients_raw is None or df_clients_raw.empty:
        return pd.DataFrame(columns=COLS_CLIENTS)
    df, _ = map_columns_heuristic(df_clients_raw)
    if "Date" in df.columns:
        try:
            df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
        except Exception:
            pass
    df = _ensure_columns(df, COLS_CLIENTS)
    # numeric normalization
    for col in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"]:
        if col in df.columns:
            df[col] = df[col].apply(money_to_float)
    # status normalization
    df = _normalize_status(df)
    # text columns
    for c in ["Nom", "Categories", "Sous-categorie", "Visa", "Commentaires"]:
        if c in df.columns:
            df[c] = df[c].astype(str).fillna("")
    # time features
    try:
        dser = pd.to_datetime(df["Date"], errors="coerce")
        df["_Année_"] = dser.dt.year.fillna(0).astype(int)
        df["_MoisNum_"] = dser.dt.month.fillna(0).astype(int)
        df["Mois"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) and m>0 else "")
    except Exception:
        df["_Année_"] = 0; df["_MoisNum_"] = 0; df["Mois"] = ""
    return df

df_all = normalize_clients_for_live(df_clients_raw)

DF_LIVE_KEY = skey("df_live")
if isinstance(df_all, pd.DataFrame) and not df_all.empty:
    st.session_state[DF_LIVE_KEY] = df_all.copy()
else:
    if DF_LIVE_KEY not in st.session_state:
        st.session_state[DF_LIVE_KEY] = pd.DataFrame(columns=COLS_CLIENTS)

def _get_df_live() -> pd.DataFrame:
    return st.session_state[DF_LIVE_KEY].copy()

def _set_df_live(df: pd.DataFrame) -> None:
    st.session_state[DF_LIVE_KEY] = df.copy()

# ensure flag columns exist in DF when creating/updating
def ensure_flag_columns(df: pd.DataFrame, flags: List[str]) -> None:
    for f in flags:
        if f not in df.columns:
            df[f] = 0

DEFAULT_FLAGS = ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]

# -------------------------
# UI - Tabs (Files / Dashboard / Analyses / Gestion / Export)
# -------------------------
tabs = st.tabs(["📄 Fichiers","📊 Dashboard","📈 Analyses","➕ / ✏️ / 🗑️ Gestion","💾 Export"])

# Fichiers tab
with tabs[0]:
    st.header("📂 Fichiers")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Clients")
        if up_clients is not None:
            st.write("Upload:", getattr(up_clients, "name", ""))
        elif isinstance(clients_src, str) and clients_src:
            st.write("Chargé depuis chemin local:", clients_src)
        if df_clients_raw is None or df_clients_raw.empty:
            st.warning("Lecture Clients : aucun tableau trouvé ou DataFrame vide.")
        else:
            st.success(f"Clients lus ({df_clients_raw.shape[0]} lignes, {df_clients_raw.shape[1]} colonnes)")
            st.dataframe(df_clients_raw.head(8), use_container_width=True, height=220)
    with c2:
        st.subheader("Visa")
        if mode == "Deux fichiers (Clients & Visa)":
            if up_visa is not None:
                st.write("Upload:", getattr(up_visa, "name", ""))
            elif isinstance(visa_src, str) and visa_src:
                st.write("Chargé depuis chemin local:", visa_src)
        else:
            st.write("Mode 'Un fichier' : Visa sera lu depuis le même fichier Clients si présent.")
        if df_visa_raw is None or df_visa_raw.empty:
            st.warning("Lecture Visa : aucun tableau trouvé ou DataFrame vide.")
        else:
            st.success(f"Visa lu ({df_visa_raw.shape[0]} lignes, {df_visa_raw.shape[1]} colonnes)")
            st.dataframe(df_visa_raw.head(8), use_container_width=True, height=220)
    st.markdown("---")
    if st.button("Réinitialiser la mémoire (recharger depuis fichiers)"):
        df_all2 = normalize_clients_for_live(df_clients_raw)
        _set_df_live(df_all2)
        st.success("Mémoire réinitialisée.")
        st.experimental_rerun()

# Dashboard tab (compact KPIs + recent)
with tabs[1]:
    st.subheader("📊 Dashboard")
    df_live_view = _get_df_live()
    if df_live_view is None or df_live_view.empty:
        st.info("Aucune donnée en mémoire.")
    else:
        cats = sorted(df_live_view["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_live_view.columns else []
        subs = sorted(df_live_view["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_live_view.columns else []
        years = sorted(pd.to_numeric(df_live_view["_Année_"], errors="coerce").dropna().astype(int).unique().tolist()) if "_Année_" in df_live_view.columns else []
        f1, f2, f3 = st.columns([1,1,1])
        sel_cat = f1.selectbox("Catégorie (filtre)", options=[""]+cats, index=0, key=skey("dash","cat"))
        sel_sub = f2.selectbox("Sous-catégorie (filtre)", options=[""]+subs, index=0, key=skey("dash","sub"))
        sel_year = f3.selectbox("Année (filtre)", options=[""]+ [str(y) for y in years], index=0, key=skey("dash","year"))
        view = df_live_view.copy()
        if sel_cat:
            view = view[view["Categories"].astype(str)==sel_cat]
        if sel_sub:
            view = view[view["Sous-categorie"].astype(str)==sel_sub]
        if sel_year:
            view = view[view["_Année_"].astype(str)==sel_year]
        total = (view.get("Montant honoraires (US $)",0).apply(_to_num) + view.get("Autres frais (US $)",0).apply(_to_num)).sum()
        paye = view.get("Payé",0).apply(_to_num).sum() if "Payé" in view.columns else 0.0
        solde = view.get("Solde",0).apply(_to_num).sum() if "Solde" in view.columns else 0.0
        kcols = st.columns([1,1,1])
        def small_metric(col, label, value):
            with col:
                st.markdown(f"<div style='font-size:14px;font-weight:600'>{label}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='font-size:16px;color:#0A6EBD;font-weight:700'>{value}</div>", unsafe_allow_html=True)
        small_metric(kcols[0], "Dossiers", f"{len(view):,}")
        small_metric(kcols[1], "Total facturé", f"${total:,.2f}")
        small_metric(kcols[2], "Solde total", f"${solde:,.2f}")
        st.markdown("---")
        st.subheader("Aperçu récent")
        recent = view.sort_values(by=["_Année_","_MoisNum_"], ascending=[False,False]).head(20).copy()
        display_cols = [c for c in ["Dossier N","ID_Client","Nom","Date","Categories","Sous-categorie","Visa","Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"] if c in recent.columns]
        for col in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]:
            if col in recent.columns:
                recent[col] = recent[col].apply(lambda x: f"${_to_num(x):,.2f}")
        if "Date" in recent.columns:
            try:
                recent["Date"] = pd.to_datetime(recent["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                recent["Date"] = recent["Date"].astype(str)
        st.dataframe(recent[display_cols].reset_index(drop=True), use_container_width=True)

# Analyses tab (kept simple)
with tabs[2]:
    st.subheader("📈 Analyses")
    st.info("Analyses disponibles dans la version complète.")

# Gestion tab: Add / Edit / Delete
with tabs[3]:
    st.subheader("➕ / ✏️ / 🗑️ Gestion")
    df_live = _get_df_live()
    # ensure columns
    for c in COLS_CLIENTS:
        if c not in df_live.columns:
            df_live[c] = "" if c not in ("Montant honoraires (US $)","Autres frais (US $)","Payé","Solde") else 0.0

    # categories source
    categories_options = visa_categories if visa_categories else sorted({str(x).strip() for x in df_live["Categories"].dropna().astype(str).tolist()})
    st.markdown("### Ajouter un dossier")
    st.write("Sélectionnez la catégorie (réactif):")
    categories_local = [""] + [c.strip() for c in categories_options]
    add_cat_sel = st.selectbox("Categories (réactif)", options=categories_local, index=0, key=skey("add","cat_sel"))

    # compute sous options for category
    add_sub_options = []
    if isinstance(add_cat_sel, str) and add_cat_sel.strip():
        cat_key = canonical_key(add_cat_sel)
        # find subs via visa_map_norm (category_key -> subs) where visa_map_norm keys are canonical keys
        # visa_map_norm was built earlier as canonical_key(category) -> subs (raw labels)
        if cat_key in visa_map_norm:
            add_sub_options = visa_map_norm.get(cat_key, [])[:]
        else:
            # fallback: try direct visa_map lookup by raw category name
            if add_cat_sel in visa_map:
                add_sub_options = visa_map.get(add_cat_sel, [])[:]
    if not add_sub_options:
        add_sub_options = sorted({str(x).strip() for x in df_live["Sous-categorie"].dropna().astype(str).tolist()})
    st.sidebar.write("DEBUG selected category:", repr(add_cat_sel))
    st.sidebar.write("DEBUG computed sous-categories:", add_sub_options)

    # Add form
    with st.form(key=skey("form_add")):
        col1, col2, col3 = st.columns(3)
        with col1:
            next_id = get_next_client_id(df_live)
            st.markdown(f"**ID_Client (automatique)**: {next_id}")
            add_dossier = st.text_input("Dossier N", value="", key=skey("add","dossier"))
            add_nom = st.text_input("Nom", value="", key=skey("add","nom"))
        with col2:
            add_date = st.date_input("Date", value=date.today(), key=skey("add","date"))
            st.markdown(f"Catégorie choisie: **{add_cat_sel}**")
            add_sub = st.selectbox("Sous-categorie", options=[""] + add_sub_options, index=0, key=skey("add","sub"))
            # robust lookup for checkbox labels
            specific_options = get_sub_options_for(add_sub)
            checkbox_options = specific_options if specific_options else DEFAULT_FLAGS
            st.sidebar.write("DEBUG add_sub selection:", repr(add_sub))
            st.sidebar.write("DEBUG options for selected sub:", specific_options)
            add_flags_state = {}
            for opt in checkbox_options:
                k = skey("add","flag", re.sub(r"\s+","_", opt))
                add_flags_state[opt] = st.checkbox(opt, value=False, key=k)
        with col3:
            add_visa = st.text_input("Visa", value="", key=skey("add","visa"))
            add_montant = st.text_input("Montant honoraires (US $)", value="0", key=skey("add","montant"))
            add_autres = st.text_input("Autres frais (US $)", value="0", key=skey("add","autres"))
        add_comments = st.text_area("Commentaires", value="", key=skey("add","comments"))
        submitted = st.form_submit_button("Ajouter")
        if submitted:
            try:
                new_row = {c: "" for c in df_live.columns}
                new_row["ID_Client"] = str(next_id)
                new_row["Dossier N"] = add_dossier
                new_row["Nom"] = add_nom
                new_row["Date"] = pd.to_datetime(add_date)
                new_row["Categories"] = add_cat_sel.strip() if isinstance(add_cat_sel, str) else add_cat_sel
                new_row["Sous-categorie"] = add_sub.strip() if isinstance(add_sub, str) else add_sub
                new_row["Visa"] = add_visa
                new_row["Montant honoraires (US $)"] = money_to_float(add_montant)
                new_row["Autres frais (US $)"] = money_to_float(add_autres)
                new_row["Payé"] = 0.0
                new_row["Solde"] = new_row["Montant honoraires (US $)"] + new_row["Autres frais (US $)"]
                new_row["Commentaires"] = add_comments
                flags_to_create = list(add_flags_state.keys())
                ensure_flag_columns(df_live, flags_to_create)
                for opt, val in add_flags_state.items():
                    new_row[opt] = 1 if val else 0
                # append and persist
                df_live = df_live.append(new_row, ignore_index=True)
                _set_df_live(df_live)
                st.success("Dossier ajouté.")
            except Exception as e:
                st.error(f"Erreur ajout: {e}")

    st.markdown("---")
    st.markdown("### Modifier un dossier")
    if df_live is None or df_live.empty:
        st.info("Aucun dossier à modifier.")
    else:
        choices = [f"{i} | {df_live.at[i,'Dossier N'] if 'Dossier N' in df_live.columns else ''} | {df_live.at[i,'Nom'] if 'Nom' in df_live.columns else ''}" for i in range(len(df_live))]
        sel = st.selectbox("Sélectionner ligne", options=[""]+choices, key=skey("edit","select"))
        if sel:
            idx = int(sel.split("|")[0].strip())
            row = df_live.loc[idx].copy()
            st.write("Modifier la catégorie (réactif) :")
            edit_cat_options = [""] + [c.strip() for c in categories_options]
            init_cat = str(row.get("Categories","")).strip()
            try:
                init_cat_index = edit_cat_options.index(init_cat)
            except Exception:
                init_cat_index = 0
            e_cat_sel = st.selectbox("Categories (réactif)", options=edit_cat_options, index=init_cat_index, key=skey("edit","cat_sel"))
            # compute sous options
            edit_sub_options = []
            if isinstance(e_cat_sel, str) and e_cat_sel.strip():
                cat_key = canonical_key(e_cat_sel)
                if cat_key in visa_map_norm:
                    edit_sub_options = visa_map_norm.get(cat_key, [])[:]
                else:
                    if e_cat_sel in visa_map:
                        edit_sub_options = visa_map.get(e_cat_sel, [])[:]
            if not edit_sub_options:
                edit_sub_options = sorted({str(x).strip() for x in df_live["Sous-categorie"].dropna().astype(str).tolist()})
            st.sidebar.write("DEBUG edit selected cat:", repr(e_cat_sel))
            st.sidebar.write("DEBUG edit_sub_options:", edit_sub_options)
            with st.form(key=skey("form_edit")):
                e_col1, e_col2 = st.columns(2)
                with e_col1:
                    st.markdown(f"**ID_Client :** {row.get('ID_Client','')}")
                    e_dossier = st.text_input("Dossier N", value=str(row.get("Dossier N","")), key=skey("edit","dossier"))
                    e_nom = st.text_input("Nom", value=str(row.get("Nom","")), key=skey("edit","nom"))
                with e_col2:
                    e_date = st.date_input("Date", value=date.today() if pd.isna(row.get("Date", pd.NaT)) else _date_for_widget(row.get("Date")), key=skey("edit","date"))
                    st.markdown(f"Category choisie: **{e_cat_sel}**")
                    init_sub = str(row.get("Sous-categorie","")).strip()
                    try:
                        init_sub_index = ([""] + edit_sub_options).index(init_sub)
                    except Exception:
                        init_sub_index = 0
                    e_sub = st.selectbox("Sous-categorie", options=[""] + edit_sub_options, index=init_sub_index, key=skey("edit","sub"))
                    # compute checkbox options for this e_sub
                    edit_specific = get_sub_options_for(e_sub)
                    checkbox_options_edit = edit_specific if edit_specific else DEFAULT_FLAGS
                    ensure_flag_columns(df_live, checkbox_options_edit)
                    edit_flags_state = {}
                    for opt in checkbox_options_edit:
                        initial_val = True if (opt in df_live.columns and _to_num(row.get(opt, 0))>0) else False
                        k = skey("edit","flag", re.sub(r"\s+","_", opt), str(idx))
                        edit_flags_state[opt] = st.checkbox(opt, value=initial_val, key=k)
                e_visa = st.text_input("Visa", value=str(row.get("Visa","")), key=skey("edit","visa"))
                e_montant = st.text_input("Montant honoraires (US $)", value=str(row.get("Montant honoraires (US $)",0)), key=skey("edit","montant"))
                e_autres = st.text_input("Autres frais (US $)", value=str(row.get("Autres frais (US $)",0)), key=skey("edit","autres"))
                e_paye = st.text_input("Payé", value=str(row.get("Payé",0)), key=skey("edit","paye"))
                e_comments = st.text_area("Commentaires", value=str(row.get("Commentaires","")), key=skey("edit","comments"))
                save = st.form_submit_button("Enregistrer modifications")
                if save:
                    try:
                        df_live.at[idx, "Dossier N"] = e_dossier
                        df_live.at[idx, "Nom"] = e_nom
                        df_live.at[idx, "Date"] = pd.to_datetime(e_date)
                        df_live.at[idx, "Categories"] = e_cat_sel.strip() if isinstance(e_cat_sel,str) else e_cat_sel
                        df_live.at[idx, "Sous-categorie"] = e_sub.strip() if isinstance(e_sub,str) else e_sub
                        df_live.at[idx, "Visa"] = e_visa
                        df_live.at[idx, "Montant honoraires (US $)"] = money_to_float(e_montant)
                        df_live.at[idx, "Autres frais (US $)"] = money_to_float(e_autres)
                        df_live.at[idx, "Payé"] = money_to_float(e_paye)
                        df_live.at[idx, "Solde"] = _to_num(df_live.at[idx, "Montant honoraires (US $)"]) + _to_num(df_live.at[idx, "Autres frais (US $)"]) - _to_num(df_live.at[idx, "Payé"])
                        df_live.at[idx, "Commentaires"] = e_comments
                        for opt, val in edit_flags_state.items():
                            df_live.at[idx, opt] = 1 if val else 0
                        _set_df_live(df_live)
                        st.success("Modifications enregistrées.")
                    except Exception as e:
                        st.error(f"Erreur enregistrement: {e}")

    st.markdown("---")
    st.markdown("### Supprimer des dossiers")
    if df_live is None or df_live.empty:
        st.info("Aucun dossier à supprimer.")
    else:
        choices_del = [f"{i} | {df_live.at[i,'Dossier N'] if 'Dossier N' in df_live.columns else ''} | {df_live.at[i,'Nom'] if 'Nom' in df_live.columns else ''}" for i in range(len(df_live))]
        selected_to_del = st.multiselect("Sélectionnez les lignes à supprimer", options=choices_del, key=skey("del","select"))
        if st.button("Supprimer sélection"):
            if selected_to_del:
                idxs = [int(s.split("|")[0].strip()) for s in selected_to_del]
                try:
                    df_live = df_live.drop(index=idxs).reset_index(drop=True)
                    _set_df_live(df_live)
                    st.success(f"{len(idxs)} ligne(s) supprimée(s).")
                except Exception as e:
                    st.error(f"Erreur suppression: {e}")
            else:
                st.warning("Aucune sélection pour suppression.")

# Export tab
with tabs[4]:
    st.header("💾 Export")
    df_live = _get_df_live()
    if df_live is None or df_live.empty:
        st.info("Aucune donnée à exporter.")
    else:
        st.write(f"Vue en mémoire: {df_live.shape[0]} lignes, {df_live.shape[1]} colonnes")
        col1, col2 = st.columns(2)
        with col1:
            csv_bytes = df_live.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Export CSV", data=csv_bytes, file_name="Clients_export.csv", mime="text/csv")
        with col2:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df_live.to_excel(writer, index=False, sheet_name="Clients")
            st.download_button("⬇️ Export XLSX", data=buf.getvalue(), file_name="Clients_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
