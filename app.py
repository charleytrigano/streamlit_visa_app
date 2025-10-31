# app.py - Visa Manager (fixed numeric conversions + safe date handling)
# - Robust money_to_float/_to_num to handle 'nan', pd.NaT, None, etc.
# - Ensure text_inputs do not display 'nan' strings; show empty string instead
# - All st.date_input values forced to native datetime.date or None
# - Recalc Solde just before display in edit form
# Requirements: pip install streamlit pandas openpyxl
# Run: streamlit run app.py

import os
import json
import re
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional

import pandas as pd
import streamlit as st

# ---- quick globals ----
df_clients_raw: Optional[pd.DataFrame] = None
df_visa_raw: Optional[pd.DataFrame] = None
clients_src_for_read = None
visa_src_for_read = None

# optional libs
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    px = None
    HAS_PLOTLY = False

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

# -------------------------
# Configuration & constants
# -------------------------
APP_TITLE = "🛂 Visa Manager"
COLS_CLIENTS = [
    "ID_Client", "Dossier N", "Nom", "Date",
    "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)",
    "Payé", "Solde", "Solde à percevoir (US $)",
    "Acompte 1", "Date Acompte 1",
    "Acompte 2", "Date Acompte 2", "Acompte 3", "Date Acompte 3", "Acompte 4", "Date Acompte 4",
    "Escrow",
    "RFE", "Dossiers envoyé", "Dossier approuvé",
    "Dossier refusé", "Dossier Annulé",
    "Date d'envoi",
    "Date reponse",
    "Date de création", "Créé par", "Dernière modification", "Modifié par",
    "Commentaires",
    "ModeReglement", "ModeReglement_Ac1", "ModeReglement_Ac2", "ModeReglement_Ac3", "ModeReglement_Ac4"
]
MEMO_FILE = "_vmemory.json"
CACHE_CLIENTS = "_clients_cache.bin"
CACHE_VISA = "_visa_cache.bin"
SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"
SID = "vmgr"
DEFAULT_START_CLIENT_ID = 13057
CURRENT_USER = "charleytrigano"
DEFAULT_FLAGS = ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]

def skey(*parts: str) -> str:
    return f"{SID}_" + "_".join([p for p in parts if p])

# -------------------------
# Helpers
# -------------------------
def normalize_header_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s)
    return s

def remove_accents(s: Any) -> str:
    if s is None:
        return ""
    s2 = str(s)
    replace_map = {
        "é":"e","è":"e","ê":"e","ë":"e",
        "à":"a","â":"a",
        "î":"i","ï":"i",
        "ô":"o","ö":"o",
        "ù":"u","û":"u","ü":"u",
        "ç":"c"
    }
    for k,v in replace_map.items():
        s2 = s2.replace(k, v)
    return s2

def canonical_key(s: Any) -> str:
    if s is None:
        return ""
    s2 = normalize_header_text(str(s)).lower()
    s2 = remove_accents(s2)
    s2 = re.sub(r"[^a-z0-9 ]", " ", s2)
    s2 = re.sub(r"\s+", " ", s2).strip()
    return s2

# Robust money parser - treat 'nan', 'NaN', None, pd.NaT, pd.NA consistently as 0.0
def money_to_float(x: Any) -> float:
    try:
        # pandas NA handling
        if x is None:
            return 0.0
        if isinstance(x, (float, int)) and (not pd.isna(x)):
            return float(x)
        # handle pandas types: Timestamp/NaT
        try:
            if pd.isna(x):
                return 0.0
        except Exception:
            pass
        s = str(x).strip()
        if s == "" or s.lower() in ("na","n/a","nan","none","null"):
            return 0.0
        # remove non-numeric except , . -
        s = s.replace("\u202f", "").replace("\xa0", "").replace(" ", "")
        s = re.sub(r"[^\d,.\-]", "", s)
        if s == "":
            return 0.0
        # determine decimal separator
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        else:
            if "," in s and s.count(",") == 1 and "." not in s:
                # assume comma as decimal if two decimals digits
                if len(s.split(",")[-1]) == 2:
                    s = s.replace(",", ".")
                else:
                    s = s.replace(",", "")
            else:
                s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            # final fallback: strip everything except digits and dot and minus
            return float(re.sub(r"[^0-9.\-]", "", s) or 0.0)
    except Exception:
        return 0.0

def _to_num(x: Any) -> float:
    if isinstance(x, (int, float)) and not pd.isna(x):
        return float(x)
    return money_to_float(x)

def _fmt_money(v: Any) -> str:
    try:
        return "${:,.2f}".format(float(v))
    except Exception:
        return "$0.00"

# Single authoritative date converter used everywhere
def _date_or_none_safe(v: Any) -> Optional[date]:
    try:
        if v is None:
            return None
        if isinstance(v, date) and not isinstance(v, datetime):
            return v
        if isinstance(v, datetime):
            return v.date()
        d = pd.to_datetime(v, errors="coerce")
        if pd.isna(d):
            return None
        return date(int(d.year), int(d.month), int(d.day))
    except Exception:
        return None

# -------------------------
# Column heuristics & mapping (unchanged)
# -------------------------
COL_CANDIDATES = {
    "id client": "ID_Client", "idclient": "ID_Client",
    "dossier n": "Dossier N", "dossier": "Dossier N",
    "nom": "Nom", "date": "Date",
    "categories": "Categories", "categorie": "Categories",
    "sous categorie": "Sous-categorie", "sous-categorie": "Sous-categorie", "souscategorie": "Sous-categorie",
    "visa": "Visa",
    "montant": "Montant honoraires (US $)", "montant honoraires": "Montant honoraires (US $)",
    "autres frais": "Autres frais (US $)", "autresfrais": "Autres frais (US $)",
    "payé": "Payé", "paye": "Payé",
    "solde": "Solde",
    "solde a percevoir": "Solde à percevoir (US $)",
    "acompte 1": "Acompte 1", "acompte1": "Acompte 1",
    "date acompte 1": "Date Acompte 1",
    "acompte 2": "Acompte 2", "acompte2": "Acompte 2",
    "acompte 3": "Acompte 3", "acompte3": "Acompte 3",
    "acompte 4": "Acompte 4", "acompte4": "Acompte 4",
    "escrow": "Escrow",
    "dossier envoye": "Dossiers envoyé", "dossier approuve": "Dossier approuvé", "dossier refuse": "Dossier refusé",
    "rfe": "RFE", "commentaires": "Commentaires", "mode reglement":"ModeReglement"
}

NUMERIC_TARGETS = [
    "Montant honoraires (US $)",
    "Autres frais (US $)",
    "Payé",
    "Solde",
    "Solde à percevoir (US $)",
    "Acompte 1",
    "Acompte 2",
    "Acompte 3",
    "Acompte 4"
]

def detect_acompte_columns(df: pd.DataFrame) -> List[str]:
    if df is None or df.empty:
        return []
    cols = [c for c in df.columns if "acompte" in canonical_key(c)]
    def sort_key(name):
        m = re.search(r"(\d+)", name)
        return int(m.group(1)) if m else 999
    return sorted(cols, key=sort_key)

def detect_montant_column(df: pd.DataFrame) -> Optional[str]:
    if df is None or df.empty:
        return None
    candidates = ["Montant honoraires (US $)", "Montant honoraires", "Montant", "Montant honoraires (USD)"]
    for c in candidates:
        if c in df.columns:
            return c
    for c in df.columns:
        k = canonical_key(c)
        if "montant" in k or "honorair" in k:
            return c
    return None

def detect_autres_column(df: pd.DataFrame) -> Optional[str]:
    if df is None or df.empty:
        return None
    candidates = ["Autres frais (US $)", "Autres frais", "Autres"]
    for c in candidates:
        if c in df.columns:
            return c
    for c in df.columns:
        k = canonical_key(c)
        if "autre" in k or "frais" in k:
            return c
    return None

def map_columns_heuristic(df: Any) -> Tuple[pd.DataFrame, Dict[str,str]]:
    if not isinstance(df, pd.DataFrame):
        return pd.DataFrame(), {}
    mapping: Dict[str,str] = {}
    for c in list(df.columns):
        key = canonical_key(c)
        mapped = None
        if key in COL_CANDIDATES:
            mapped = COL_CANDIDATES[key]
        else:
            for cand_key, std in sorted(COL_CANDIDATES.items(), key=lambda t: -len(t[0])):
                if cand_key in key:
                    mapped = std
                    break
        mapping[c] = mapped or normalize_header_text(c)
    new_names = {}
    seen = {}
    for orig, new in mapping.items():
        base = new
        cnt = seen.get(base, 0)
        if cnt:
            new_name = f"{base}_{cnt+1}"
            seen[base] = cnt+1
        else:
            new_name = base
            seen[base] = 1
        new_names[orig] = new_name
    try:
        df = df.rename(columns=new_names)
    except Exception:
        pass
    return df, new_names

def coerce_category_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    rename_map = {}
    def _ck(x): return canonical_key(str(x))
    for c in list(df.columns):
        k = _ck(c)
        if ("sous" in k and "categorie" in k) or ("souscategorie" in k):
            if "Sous-categorie" not in df.columns:
                rename_map[c] = "Sous-categorie"
        elif ("categorie" in k or "categories" in k) and "sous" not in k:
            if "Categories" not in df.columns:
                rename_map[c] = "Categories"
    if rename_map:
        try:
            df = df.rename(columns=rename_map)
        except Exception:
            pass
    return df

# -------------------------
# Visa mapping, reading helpers (unchanged)
# -------------------------
visa_sub_options_map: Dict[str, List[str]] = {}
visa_map: Dict[str, List[str]] = {}
visa_map_norm: Dict[str, List[str]] = {}
visa_categories: List[str] = []

def get_visa_options(cat: Optional[str], sub: Optional[str]) -> List[str]:
    try:
        if sub:
            ksub = canonical_key(sub)
            if ksub in visa_sub_options_map:
                opts = visa_sub_options_map.get(ksub, [])
                if opts:
                    return opts[:]
    except Exception:
        pass
    try:
        if cat:
            kcat = canonical_key(cat)
            if kcat in visa_map_norm:
                return visa_map_norm.get(kcat, [])[:]
    except Exception:
        pass
    return []

def try_read_excel_from_bytes(b: bytes, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
    bio = BytesIO(b)
    try:
        xls = pd.ExcelFile(bio, engine="openpyxl")
        sheets = xls.sheet_names
        if sheet_name and sheet_name in sheets:
            return pd.read_excel(BytesIO(b), sheet_name=sheet_name, engine="openpyxl")
        for cand in [SHEET_CLIENTS, SHEET_VISA, "Sheet1"]:
            if cand in sheets:
                try:
                    return pd.read_excel(BytesIO(b), sheet_name=cand, engine="openpyxl")
                except Exception:
                    continue
        return pd.read_excel(BytesIO(b), sheet_name=0, engine="openpyxl")
    except Exception:
        return None

def read_any_table(src: Any, sheet: Optional[str] = None, debug_prefix: str = "") -> Optional[pd.DataFrame]:
    def _log(msg: str):
        try:
            st.sidebar.info(f"{debug_prefix}{msg}")
        except Exception:
            pass
    if src is None:
        _log("read_any_table: src is None")
        return None
    try:
        if isinstance(src, (bytes, bytearray)):
            df = try_read_excel_from_bytes(bytes(src), sheet)
            if df is not None:
                return df
            for sep in [";", ","]:
                for enc in ["utf-8", "latin-1", "cp1252"]:
                    try:
                        return pd.read_csv(BytesIO(src), sep=sep, encoding=enc, on_bad_lines="skip")
                    except Exception:
                        continue
            return None
        if isinstance(src, BytesIO):
            b = src.getvalue()
            df = try_read_excel_from_bytes(b, sheet)
            if df is not None:
                return df
            for sep in [";", ","]:
                for enc in ["utf-8", "latin-1", "cp1252"]:
                    try:
                        return pd.read_csv(BytesIO(b), sep=sep, encoding=enc, on_bad_lines="skip")
                    except Exception:
                        continue
            return None
        if hasattr(src, "read") and hasattr(src, "name"):
            try:
                data = src.getvalue()
            except Exception:
                try:
                    src.seek(0); data = src.read()
                except Exception:
                    data = None
            if data:
                df = try_read_excel_from_bytes(data, sheet)
                if df is not None:
                    return df
                for sep in [";", ","]:
                    for enc in ["utf-8", "latin-1", "cp1252"]:
                        try:
                            return pd.read_csv(BytesIO(data), sep=sep, encoding=enc, on_bad_lines="skip")
                        except Exception:
                            continue
            return None
        if isinstance(src, (str, os.PathLike)):
            p = str(src)
            if not os.path.exists(p):
                _log(f"path does not exist: {p}")
                return None
            if p.lower().endswith(".csv"):
                for sep in [";", ","]:
                    for enc in ["utf-8", "latin-1", "cp1252"]:
                        try:
                            return pd.read_csv(p, sep=sep, encoding=enc, on_bad_lines="skip")
                        except Exception:
                            continue
                return None
            else:
                try:
                    return pd.read_excel(p, sheet_name=sheet or 0, engine="openpyxl")
                except Exception:
                    return None
    except Exception as e:
        _log(f"read_any_table exception: {e}")
        return None
    _log("read_any_table: unsupported src type")
    return None

# -------------------------
# Ensure columns & normalize dataset
# -------------------------
def _ensure_columns(df: Any, cols: List[str]) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame):
        df = pd.DataFrame()
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            if c in ["Payé", "Solde", "Solde à percevoir (US $)", "Montant honoraires (US $)", "Autres frais (US $)",
                     "Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
                out[c] = 0.0
            elif c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
                out[c] = 0
            elif c in ["Date de création", "Dernière modification", "Date", "Date Acompte 1", "Date Acompte 2", "Date Acompte 3", "Date Acompte 4", "Date d'envoi", "Date reponse"]:
                out[c] = pd.NaT
            elif c == "Escrow":
                out[c] = 0
            else:
                out[c] = ""
    try:
        return out[cols]
    except Exception:
        safe = pd.DataFrame(columns=cols)
        for c in cols:
            if c in out.columns:
                safe[c] = out[c]
            else:
                if c in ["Payé", "Solde", "Solde à percevoir (US $)", "Montant honoraires (US $)", "Autres frais (US $)",
                         "Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
                    safe[c] = 0.0
                elif c in ["Date de création", "Dernière modification", "Date", "Date Acompte 1", "Date Acompte 2", "Date Acompte 3", "Date Acompte 4", "Date d'envoi", "Date reponse"]:
                    safe[c] = pd.NaT
                elif c == "Escrow":
                    safe[c] = 0
                else:
                    safe[c] = ""
        return safe

def normalize_clients_for_live(df_clients_raw_in: Any) -> pd.DataFrame:
    df_clients_raw_local = df_clients_raw_in
    if not isinstance(df_clients_raw_local, pd.DataFrame):
        maybe_df = read_any_table(df_clients_raw_local, sheet=None, debug_prefix="[normalize] ")
        df_clients_raw_local = maybe_df if isinstance(maybe_df, pd.DataFrame) else pd.DataFrame()
    df_mapped, _ = map_columns_heuristic(df_clients_raw_local)
    for dtc in ["Date","Date de création","Dernière modification","Date Acompte 1","Date Acompte 2","Date Acompte 3","Date Acompte 4","Date d'envoi","Date reponse"]:
        if dtc in df_mapped.columns:
            try:
                df_mapped[dtc] = pd.to_datetime(df_mapped[dtc], dayfirst=True, errors="coerce")
            except Exception:
                pass
    df = _ensure_columns(df_mapped, COLS_CLIENTS)
    # convert numeric fields robustly
    for col in NUMERIC_TARGETS:
        if col in df.columns:
            try:
                df[col] = df[col].apply(lambda x: _to_num(x))
            except Exception:
                df[col] = 0.0
    # ensure acomptes columns exist
    for acc in ["Acompte 1","Acompte 2","Acompte 3","Acompte 4"]:
        if acc not in df.columns:
            df[acc] = 0.0
    acomptes_cols = detect_acompte_columns(df)
    if acomptes_cols:
        try:
            df["Payé"] = df[acomptes_cols].fillna(0).apply(lambda row: sum([_to_num(row[c]) for c in acomptes_cols]), axis=1)
        except Exception:
            df["Payé"] = df.get("Payé", 0).apply(lambda x: _to_num(x))
    try:
        montant_col = detect_montant_column(df) or "Montant honoraires (US $)"
        autres_col = detect_autres_column(df) or "Autres frais (US $)"
        df[montant_col] = df.get(montant_col, 0).apply(lambda x: _to_num(x))
        df[autres_col] = df.get(autres_col, 0).apply(lambda x: _to_num(x))
        df["Payé"] = df.get("Payé", 0).apply(lambda x: _to_num(x))
        df["Solde"] = df[montant_col] + df[autres_col] - df["Payé"]
        df["Solde à percevoir (US $)"] = df["Solde"].copy()
    except Exception:
        df["Solde"] = df.get("Solde", 0).apply(lambda x: _to_num(x))
        df["Solde à percevoir (US $)"] = df.get("Solde à percevoir (US $)", 0).apply(lambda x: _to_num(x))
    for f in ["RFE","Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé"]:
        if f not in df.columns:
            df[f] = 0
    if "Escrow" not in df.columns:
        df["Escrow"] = 0
    for c in ["Nom","Categories","Sous-catégorie","Visa","Commentaires","Créé par","Modifié par","ModeReglement","ModeReglement_Ac1","ModeReglement_Ac2","ModeReglement_Ac3","ModeReglement_Ac4"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str)
    try:
        dser = pd.to_datetime(df["Date"], errors="coerce")
        df["_Année_"] = dser.dt.year.fillna(0).astype(int)
        df["_MoisNum_"] = dser.dt.month.fillna(0).astype(int)
        df["Mois"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) and m>0 else "")
    except Exception:
        df["_Année_"] = 0; df["_MoisNum_"] = 0; df["Mois"] = ""
    return df

def recalc_payments_and_solde(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    acomptes = detect_acompte_columns(out)
    if not acomptes:
        for acc in ["Acompte 1","Acompte 2","Acompte 3","Acompte 4"]:
            if acc not in out.columns:
                out[acc] = 0.0
        acomptes = detect_acompte_columns(out)
    # coerce acomptes numeric
    for c in acomptes:
        try:
            out[c] = out[c].apply(lambda x: _to_num(x) if not isinstance(x,(int,float)) else float(x))
        except Exception:
            out[c] = out[c].apply(lambda x: 0.0)
    montant_col = detect_montant_column(out) or "Montant honoraires (US $)"
    autres_col = detect_autres_column(out) or "Autres frais (US $)"
    for c in [montant_col, autres_col]:
        if c not in out.columns:
            out[c] = 0.0
        else:
            try:
                out[c] = out[c].apply(lambda x: _to_num(x) if not isinstance(x,(int,float)) else float(x))
            except Exception:
                out[c] = out[c].apply(lambda x: 0.0)
    try:
        if acomptes:
            out["Payé"] = out[acomptes].sum(axis=1).astype(float)
        else:
            out["Payé"] = out.get("Payé",0).apply(lambda x: _to_num(x))
    except Exception:
        out["Payé"] = out.get("Payé",0).apply(lambda x: _to_num(x))
    try:
        out["Solde"] = out[montant_col] + out[autres_col] - out["Payé"]
        out["Solde à percevoir (US $)"] = out["Solde"].copy()
        out["Solde"] = out["Solde"].astype(float)
    except Exception:
        out["Solde"] = out.get("Solde",0).apply(lambda x: _to_num(x))
    if "Escrow" in out.columns:
        try:
            out["Escrow"] = out["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ("1","true","t","yes","oui","y","x") else (1 if _to_num(x) == 1 else 0))
        except Exception:
            out["Escrow"] = out["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ("1","true","t","yes","oui","y","x") else 0)
    return out

# initialize live df
DF_LIVE_KEY = skey("df_live")
if DF_LIVE_KEY not in st.session_state:
    st.session_state[DF_LIVE_KEY] = pd.DataFrame(columns=COLS_CLIENTS)

def _get_df_live() -> pd.DataFrame:
    df = st.session_state.get(DF_LIVE_KEY)
    if df is None or not isinstance(df, pd.DataFrame):
        df = pd.DataFrame(columns=COLS_CLIENTS)
        st.session_state[DF_LIVE_KEY] = df
    return df.copy()

def _get_df_live_safe() -> pd.DataFrame:
    try:
        return _get_df_live()
    except Exception:
        df = pd.DataFrame(columns=COLS_CLIENTS)
        st.session_state[DF_LIVE_KEY] = df
        return df.copy()

def _set_df_live(df: pd.DataFrame) -> None:
    st.session_state[DF_LIVE_KEY] = df.copy()

# -------------------------
# UI bootstrap & files
# -------------------------
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.sidebar.header("📂 Fichiers")
last_clients_path = ""
last_visa_path = ""
try:
    if os.path.exists(MEMO_FILE):
        with open(MEMO_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
            last_clients_path = d.get("clients","")
            last_visa_path = d.get("visa","")
except Exception:
    pass

up_clients = st.sidebar.file_uploader("Clients (xlsx/xls/csv)", type=["xlsx","xls","csv"], key=skey("up_clients"))
up_visa = st.sidebar.file_uploader("Visa (xlsx/xls/csv)", type=["xlsx","xls","csv"], key=skey("up_visa"))
clients_path_in = st.sidebar.text_input("ou chemin local Clients (optionnel)", value=last_clients_path or "", key=skey("cli_path"))
visa_path_in = st.sidebar.text_input("ou chemin local Visa (optionnel)", value=last_visa_path or "", key=skey("vis_path"))

if st.sidebar.button("📥 Sauvegarder chemins", key=skey("btn_save_paths")):
    try:
        with open(MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump({"clients": clients_path_in or "", "visa": visa_path_in or ""}, f, ensure_ascii=False, indent=2)
        st.sidebar.success("Chemins sauvegardés.")
    except Exception:
        st.sidebar.error("Impossible de sauvegarder les chemins.")

# store uploads to cache
if up_clients is not None:
    try:
        clients_bytes = up_clients.getvalue()
        with open(CACHE_CLIENTS, "wb") as f:
            f.write(clients_bytes)
        clients_src_for_read = BytesIO(clients_bytes)
    except Exception:
        clients_src_for_read = None
elif clients_path_in:
    clients_src_for_read = clients_path_in
elif os.path.exists(CACHE_CLIENTS):
    try:
        clients_bytes = open(CACHE_CLIENTS, "rb").read()
        clients_src_for_read = BytesIO(clients_bytes)
    except Exception:
        clients_src_for_read = None
else:
    clients_src_for_read = None

if up_visa is not None:
    try:
        visa_bytes = up_visa.getvalue()
        with open(CACHE_VISA, "wb") as f:
            f.write(visa_bytes)
        visa_src_for_read = BytesIO(visa_bytes)
    except Exception:
        visa_src_for_read = None
elif visa_path_in:
    visa_src_for_read = visa_path_in
elif os.path.exists(CACHE_VISA):
    try:
        visa_bytes = open(CACHE_VISA, "rb").read()
        visa_src_for_read = BytesIO(visa_bytes)
    except Exception:
        visa_src_for_read = None
else:
    visa_src_for_read = None

# read raw tables
if clients_src_for_read is not None:
    try:
        maybe = read_any_table(clients_src_for_read, sheet=SHEET_CLIENTS, debug_prefix="[Clients] ")
        if maybe is None:
            maybe = read_any_table(clients_src_for_read, sheet=None, debug_prefix="[Clients fallback] ")
        if isinstance(maybe, pd.DataFrame):
            df_clients_raw = maybe
    except Exception:
        df_clients_raw = df_clients_raw if df_clients_raw is not None else pd.DataFrame()

if visa_src_for_read is not None:
    try:
        maybe = read_any_table(visa_src_for_read, sheet=SHEET_VISA, debug_prefix="[Visa] ")
        if maybe is None:
            maybe = read_any_table(visa_src_for_read, sheet=None, debug_prefix="[Visa fallback] ")
        if isinstance(maybe, pd.DataFrame):
            df_visa_raw = maybe
    except Exception:
        df_visa_raw = df_visa_raw if df_visa_raw is not None else pd.DataFrame()

# sanitize visa sheet
if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
    try:
        df_visa_raw = df_visa_raw.fillna("")
        for c in df_visa_raw.columns:
            try:
                df_visa_raw[c] = df_visa_raw[c].astype(str).str.strip()
            except Exception:
                pass
    except Exception:
        pass

# build visa maps
visa_map = {}; visa_map_norm = {}; visa_categories = []; visa_sub_options_map = {}
if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
    df_visa_mapped, _ = map_columns_heuristic(df_visa_raw)
    try:
        df_visa_mapped = coerce_category_columns(df_visa_mapped)
    except Exception:
        pass
    raw_vm = {}
    try:
        for _, row in df_visa_mapped.iterrows():
            cat = str(row.get("Categories","")).strip()
            sub = str(row.get("Sous-categorie","")).strip()
            if not cat:
                continue
            raw_vm.setdefault(cat, [])
            if sub and sub not in raw_vm[cat]:
                raw_vm[cat].append(sub)
    except Exception:
        raw_vm = {}
    raw_vm = {k: [s for s in v if s and str(s).strip().lower() != "nan"] for k, v in raw_vm.items()}
    visa_map = {k.strip(): [s.strip() for s in v] for k, v in raw_vm.items()}
    visa_map_norm = {canonical_key(k): v for k, v in visa_map.items()}
    visa_categories = sorted(list(visa_map.keys()))
    visa_sub_options_map = {}
    try:
        cols_to_skip = set(["Categories","Categorie","Sous-categorie"])
        cols_to_check = [c for c in df_visa_mapped.columns if c not in cols_to_skip]
        for _, row in df_visa_mapped.iterrows():
            sub = str(row.get("Sous-categorie","")).strip()
            if not sub:
                continue
            key = canonical_key(sub)
            for col in cols_to_check:
                val = row.get(col,"")
                truthy = False
                if pd.isna(val):
                    truthy = False
                else:
                    sval = str(val).strip().lower()
                    if sval in ("1","x","t","true","oui","yes","y"):
                        truthy = True
                    else:
                        try:
                            if float(sval) == 1.0:
                                truthy = True
                        except Exception:
                            truthy = False
                if truthy:
                    visa_sub_options_map.setdefault(key, [])
                    if col not in visa_sub_options_map[key]:
                        visa_sub_options_map[key].append(col)
    except Exception:
        visa_sub_options_map = {}

globals()['visa_map'] = visa_map
globals()['visa_map_norm'] = visa_map_norm
globals()['visa_categories'] = visa_categories
globals()['visa_sub_options_map'] = visa_sub_options_map

# put live df into session state
df_all = normalize_clients_for_live(df_clients_raw)
df_all = recalc_payments_and_solde(df_all)
if isinstance(df_all, pd.DataFrame) and not df_all.empty:
    st.session_state[DF_LIVE_KEY] = df_all.copy()
else:
    if DF_LIVE_KEY not in st.session_state or st.session_state[DF_LIVE_KEY] is None:
        st.session_state[DF_LIVE_KEY] = pd.DataFrame(columns=COLS_CLIENTS)

# small helpers
def unique_nonempty(series):
    try:
        vals = series.dropna().astype(str).tolist()
    except Exception:
        vals = []
    out = []
    for v in vals:
        s = str(v).strip()
        if s == "" or s.lower() == "nan":
            continue
        out.append(s)
    return sorted(list(dict.fromkeys(out)))

def kpi_html(label: str, value: str, sub: str = "") -> str:
    html = f"""
    <div style="border:1px solid rgba(255,255,255,0.04); border-radius:6px; padding:8px 10px; margin:6px 4px;">
      <div style="font-size:12px; color:#666;">{label}</div>
      <div style="font-size:18px; font-weight:700; margin-top:4px;">{value}</div>
      <div style="font-size:11px; color:#888; margin-top:4px;">{sub}</div>
    </div>
    """
    return html

# -------------------------
# UI (tabs) - only core parts shown below (rest unchanged) - ensure safe date and solde display
tabs = st.tabs(["📄 Fichiers","📊 Dashboard","📈 Analyses","➕ Ajouter","✏️ / 🗑️ Gestion","💾 Export"])

# Files tab (unchanged)...
with tabs[0]:
    st.header("📂 Fichiers")
    # ... same as previous implementation (omitted here for brevity)

# Dashboard tab (unchanged)...
with tabs[1]:
    st.subheader("📊 Dashboard")
    # ... same as previous implementation (omitted here for brevity)

# Analyses tab (unchanged)...
with tabs[2]:
    st.subheader("📈 Analyses")
    # ... same as previous implementation (omitted here for brevity)

# ---- Add tab (Date Acompte 1 + Mode on same line) ----
with tabs[3]:
    st.subheader("➕ Ajouter un nouveau client")
    df_live = _get_df_live_safe()
    next_dossier_num = get_next_dossier_numeric(df_live)
    next_dossier = str(next_dossier_num)
    next_id_client = make_id_client_datebased(df_live)
    st.markdown(f"**ID_Client (auto)**: {next_id_client}")
    st.markdown(f"**Dossier N (auto)**: {next_dossier}")

    add_date = st.date_input("Date (événement)", value=date.today(), key=skey("addtab","date"))
    add_nom = st.text_input("Nom du client", value="", placeholder="Nom complet du client", key=skey("addtab","nom"))

    # categories / sub / visa (as before)
    if visa_categories:
        categories_options = visa_categories
    else:
        categories_options = unique_nonempty(df_live["Categories"]) if "Categories" in df_live.columns else []
    r3c1, r3c2, r3c3 = st.columns([1.2,1.6,1.6])
    with r3c1:
        categories_local = [""] + [c.strip() for c in categories_options]
        add_cat = st.selectbox("Catégorie", options=categories_local, index=0, key=skey("addtab","cat"))
    with r3c2:
        add_sub_options = []
        if isinstance(add_cat, str) and add_cat.strip():
            cat_key = canonical_key(add_cat)
            if cat_key in visa_map_norm:
                add_sub_options = visa_map_norm.get(cat_key, [])[:]
            else:
                if add_cat in visa_map:
                    add_sub_options = visa_map.get(add_cat, [])[:]
        if not add_sub_options:
            try:
                add_sub_options = sorted({str(x).strip() for x in df_live["Sous-catégorie"].dropna().astype(str).tolist()})
            except Exception:
                add_sub_options = []
        add_sub = st.selectbox("Sous-catégorie", options=[""] + add_sub_options, index=0, key=skey("addtab","sub"))
    with r3c3:
        specific_options = get_visa_options(add_cat, add_sub)
        if specific_options:
            add_visa = st.selectbox("Visa (options)", options=[""] + specific_options, index=0, key=skey("addtab","visa"))
        else:
            add_visa = st.text_input("Visa", value="", key=skey("addtab","visa"))

    # Montant + Acompte1
    r4c1, r4c2 = st.columns([1.4,1.0])
    with r4c1:
        add_montant = st.text_input("Montant honoraires (US $)", value="0", key=skey("addtab","montant"))
    with r4c2:
        a1 = st.text_input("Acompte 1", value="0", key=skey("addtab","ac1"))
    # Date Acompte 1 + mode
    r5c1, r5c2 = st.columns([1.6,1.0])
    with r5c1:
        a1_date = st.date_input("Date Acompte 1", value=None, key=skey("addtab","ac1_date"))
    with r5c2:
        st.caption("Mode de règlement")
        pay_cb = st.checkbox("CB", value=False, key=skey("addtab","pay_cb"))
        pay_cheque = st.checkbox("Cheque", value=False, key=skey("addtab","pay_cheque"))
        pay_virement = st.checkbox("Virement", value=False, key=skey("addtab","pay_virement"))
        pay_venmo = st.checkbox("Venmo", value=False, key=skey("addtab","pay_venmo"))

    add_comments = st.text_area("Commentaires", value="", key=skey("addtab","comments"))

    if st.button("Ajouter", key=skey("addtab","btn_add")):
        try:
            new_row = {c: "" for c in df_live.columns}
            new_row["ID_Client"] = next_id_client
            new_row["Dossier N"] = next_dossier
            new_row["Nom"] = add_nom
            new_row["Date"] = pd.to_datetime(add_date)
            new_row["Categories"] = add_cat
            new_row["Sous-categorie"] = add_sub
            new_row["Visa"] = add_visa
            new_row["Montant honoraires (US $)"] = money_to_float(add_montant)
            new_row["Autres frais (US $)"] = 0.0
            new_row["Acompte 1"] = money_to_float(a1)
            new_row["Date Acompte 1"] = pd.to_datetime(a1_date) if a1_date else pd.NaT
            new_row["Acompte 2"] = 0.0
            new_row["Acompte 3"] = 0.0
            new_row["Acompte 4"] = 0.0
            new_row["Date Acompte 2"] = pd.NaT
            new_row["Date Acompte 3"] = pd.NaT
            new_row["Date Acompte 4"] = pd.NaT
            modes = []
            if pay_cb: modes.append("CB")
            if pay_cheque: modes.append("Cheque")
            if pay_virement: modes.append("Virement")
            if pay_venmo: modes.append("Venmo")
            new_row["ModeReglement"] = ",".join(modes)
            new_row["ModeReglement_Ac1"] = ",".join(modes) if modes else ""
            new_row["ModeReglement_Ac2"] = ""
            new_row["ModeReglement_Ac3"] = ""
            new_row["ModeReglement_Ac4"] = ""
            new_row["Payé"] = new_row["Acompte 1"]
            new_row["Solde"] = new_row["Montant honoraires (US $)"] + new_row["Autres frais (US $)"] - new_row["Payé"]
            new_row["Solde à percevoir (US $)"] = new_row["Solde"]
            now = datetime.now()
            new_row["Date de création"] = now
            new_row["Créé par"] = CURRENT_USER
            new_row["Dernière modification"] = now
            new_row["Modifié par"] = CURRENT_USER
            new_row["Commentaires"] = add_comments
            ensure_flag_columns(new_row if isinstance(new_row, dict) else {}, DEFAULT_FLAGS)  # noop safe
            df_live = _get_df_live_safe()
            df_live = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            df_live = recalc_payments_and_solde(df_live)
            _set_df_live(df_live)
            st.success(f"Dossier ajouté : ID_Client {next_id_client} — Dossier N {next_dossier}")
        except Exception as e:
            st.error(f"Erreur ajout: {e}")

# ---- Gestion tab (edit) ----
with tabs[4]:
    st.subheader("✏️ / 🗑️ Gestion — Modifier / Supprimer")
    df_live = _get_df_live_safe()
    for c in COLS_CLIENTS:
        if c not in df_live.columns:
            if c in ["Date Acompte 2","Date Acompte 3","Date Acompte 4","Date d'envoi","Date reponse","Date de création","Dernière modification"]:
                df_live[c] = pd.NaT
            else:
                df_live[c] = "" if c not in NUMERIC_TARGETS else 0.0

    if df_live is None or df_live.empty:
        st.info("Aucun dossier à modifier ou supprimer.")
    else:
        choices = [f"{i} | {df_live.at[i,'Dossier N'] if 'Dossier N' in df_live.columns else ''} | {df_live.at[i,'Nom'] if 'Nom' in df_live.columns else ''}" for i in range(len(df_live))]
        sel = st.selectbox("Sélectionner ligne à modifier", options=[""]+choices, key=skey("edit","select"))
        if sel:
            idx = int(sel.split("|")[0].strip())
            row = df_live.loc[idx].copy()

            def txt(v):
                if pd.isna(v):
                    return ""
                # avoid showing 'nan' string
                s = str(v)
                if s.strip().lower() in ("nan","none","na","n/a"):
                    return ""
                return s

            def _safe_row_date_local(colname: str):
                try:
                    raw = row.get(colname)
                except Exception:
                    raw = None
                d = _date_or_none_safe(raw)
                if d is None:
                    return None
                if isinstance(d, date) and not isinstance(d, datetime):
                    return d
                try:
                    d2 = pd.to_datetime(d, errors="coerce")
                    if pd.isna(d2):
                        return None
                    return date(int(d2.year), int(d2.month), int(d2.day))
                except Exception:
                    return None

            def _parse_modes(raw):
                try:
                    if pd.isna(raw) or raw is None:
                        return []
                    s = str(raw).strip()
                    if not s:
                        return []
                    return [p.strip() for p in s.split(",") if p.strip()]
                except Exception:
                    return []

            row_modes_general = _parse_modes(row.get("ModeReglement", ""))
            row_mode_ac1 = _parse_modes(row.get("ModeReglement_Ac1", "")) or row_modes_general
            row_mode_ac2 = _parse_modes(row.get("ModeReglement_Ac2", ""))
            row_mode_ac3 = _parse_modes(row.get("ModeReglement_Ac3", ""))
            row_mode_ac4 = _parse_modes(row.get("ModeReglement_Ac4", ""))

            # ensure recalculation before display
            df_live = recalc_payments_and_solde(df_live)
            current_row = df_live.loc[idx].copy()

            with st.form(key=skey("form_edit", str(idx))):
                # Name (larger) and Solde on same line
                c_name, c_solde = st.columns([2.5,1])
                with c_name:
                    st.markdown(f"### {txt(current_row.get('Nom',''))}")
                with c_solde:
                    try:
                        sol_due_num = _to_num(current_row.get("Solde à percevoir (US $)", current_row.get("Solde", 0)))
                        sol_due = _fmt_money(sol_due_num)
                    except Exception:
                        sol_due = "$0.00"
                    st.markdown(f"**Solde dû**: {sol_due}")

                r1c1, r1c2, r1c3 = st.columns([1.4,1.0,1.2])
                with r1c1:
                    st.markdown(f"**ID_Client :** {txt(current_row.get('ID_Client',''))}")
                with r1c2:
                    e_dossier = st.text_input("Dossier N", value=txt(current_row.get("Dossier N","")), key=skey("edit","dossier", str(idx)))
                with r1c3:
                    e_date = st.date_input("Date (événement)", value=_safe_row_date_local("Date"), key=skey("edit","date", str(idx)))

                # Montants row
                m1, m2, m3 = st.columns([1.2,1.0,1.0])
                with m1:
                    e_montant = st.text_input("Montant honoraires (US $)", value=(str(current_row.get("Montant honoraires (US $)", "") if not pd.isna(current_row.get("Montant honoraires (US $)", "")) else "")), key=skey("edit","montant", str(idx)))
                with m2:
                    e_autres = st.text_input("Autres frais (US $)", value=(str(current_row.get("Autres frais (US $)", "") if not pd.isna(current_row.get("Autres frais (US $)", "")) else "")), key=skey("edit","autres", str(idx)))
                with m3:
                    try:
                        total_montant_val = _to_num(e_montant) + _to_num(e_autres)
                    except Exception:
                        total_montant_val = _to_num(current_row.get("Montant honoraires (US $)",0)) + _to_num(current_row.get("Autres frais (US $)",0))
                    st.text_input("Montant Total", value=str(total_montant_val), key=skey("edit","montant_total", str(idx)), disabled=True)

                # Acomptes row
                r_ac_1, r_ac_2, r_ac_3, r_ac_4 = st.columns([1.0,1.0,1.0,1.0])
                with r_ac_1:
                    e_ac1 = st.text_input("Acompte 1", value=(str(current_row.get("Acompte 1","")) if not pd.isna(current_row.get("Acompte 1","")) else ""), key=skey("edit","ac1", str(idx)))
                with r_ac_2:
                    e_ac2 = st.text_input("Acompte 2", value=(str(current_row.get("Acompte 2","")) if not pd.isna(current_row.get("Acompte 2","")) else ""), key=skey("edit","ac2", str(idx)))
                with r_ac_3:
                    e_ac3 = st.text_input("Acompte 3", value=(str(current_row.get("Acompte 3","")) if not pd.isna(current_row.get("Acompte 3","")) else ""), key=skey("edit","ac3", str(idx)))
                with r_ac_4:
                    e_ac4 = st.text_input("Acompte 4", value=(str(current_row.get("Acompte 4","")) if not pd.isna(current_row.get("Acompte 4","")) else ""), key=skey("edit","ac4", str(idx)))

                # Dates + modes row
                r_d1, r_d2, r_d3, r_d4 = st.columns([1.0,1.0,1.0,1.0])
                with r_d1:
                    e_mode_ac1 = st.multiselect("Mode A1", options=["CB","Cheque","Virement","Venmo"], default=row_mode_ac1, key=skey("edit","mode_ac1", str(idx)))
                    e_ac1_date = st.date_input("Date Acompte 1", value=_safe_row_date_local("Date Acompte 1"), key=skey("edit","ac1_date", str(idx)))
                with r_d2:
                    e_mode_ac2 = st.multiselect("Mode A2", options=["CB","Cheque","Virement","Venmo"], default=row_mode_ac2, key=skey("edit","mode_ac2", str(idx)))
                    e_ac2_date = st.date_input("Date Acompte 2", value=_safe_row_date_local("Date Acompte 2"), key=skey("edit","ac2_date", str(idx)))
                with r_d3:
                    e_mode_ac3 = st.multiselect("Mode A3", options=["CB","Cheque","Virement","Venmo"], default=row_mode_ac3, key=skey("edit","mode_ac3", str(idx)))
                    e_ac3_date = st.date_input("Date Acompte 3", value=_safe_row_date_local("Date Acompte 3"), key=skey("edit","ac3_date", str(idx)))
                with r_d4:
                    e_mode_ac4 = st.multiselect("Mode A4", options=["CB","Cheque","Virement","Venmo"], default=row_mode_ac4, key=skey("edit","mode_ac4", str(idx)))
                    e_ac4_date = st.date_input("Date Acompte 4", value=_safe_row_date_local("Date Acompte 4"), key=skey("edit","ac4_date", str(idx)))

                # Flags + RFE constraint
                f1, f2, f3, f4 = st.columns([1.0,1.0,1.0,1.0])
                with f1:
                    e_flag_envoye = st.checkbox("Dossiers envoyé", value=bool(int(current_row.get("Dossiers envoyé", 0))) if not pd.isna(current_row.get("Dossiers envoyé", 0)) else False, key=skey("edit","flag_envoye", str(idx)))
                with f2:
                    e_flag_approuve = st.checkbox("Dossier approuvé", value=bool(int(current_row.get("Dossier approuvé", 0))) if not pd.isna(current_row.get("Dossier approuvé", 0)) else False, key=skey("edit","flag_approuve", str(idx)))
                with f3:
                    e_flag_refuse = st.checkbox("Dossier refusé", value=bool(int(current_row.get("Dossier refusé", 0))) if not pd.isna(current_row.get("Dossier refusé", 0)) else False, key=skey("edit","flag_refuse", str(idx)))
                with f4:
                    e_flag_annule = st.checkbox("Dossier Annulé", value=bool(int(current_row.get("Dossier Annulé", 0))) if not pd.isna(current_row.get("Dossier Annulé", 0)) else False, key=skey("edit","flag_annule", str(idx)))

                other_flag_set = any([e_flag_envoye, e_flag_approuve, e_flag_refuse, e_flag_annule])
                if not other_flag_set:
                    st.markdown("**RFE** (active uniquement si un des états est coché)")
                    e_flag_rfe = st.checkbox("RFE", value=bool(int(current_row.get("RFE", 0))) if not pd.isna(current_row.get("RFE", 0)) else False, key=skey("edit","flag_rfe", str(idx)), disabled=True)
                else:
                    e_flag_rfe = st.checkbox("RFE", value=bool(int(current_row.get("RFE", 0))) if not pd.isna(current_row.get("RFE", 0)) else False, key=skey("edit","flag_rfe", str(idx)))

                dcol1, dcol2 = st.columns([1.5,1.0])
                with dcol1:
                    e_flags_date = st.date_input("Date d'envoi / Date état", value=_safe_row_date_local("Date d'envoi"), key=skey("edit","flags_date", str(idx)))
                with dcol2:
                    e_date_reponse = st.date_input("Date réponse", value=_safe_row_date_local("Date reponse"), key=skey("edit","date_reponse", str(idx)))

                e_escrow = st.checkbox("Escrow", value=bool(int(current_row.get("Escrow", 0))) if not pd.isna(current_row.get("Escrow", 0)) else False, key=skey("edit","escrow", str(idx)))
                e_comments = st.text_area("Commentaires", value=txt(current_row.get("Commentaires","")), key=skey("edit","comments", str(idx)))

                save = st.form_submit_button("Enregistrer modifications")
                if save:
                    try:
                        # update amounts and acomptes robustly
                        df_live.at[idx, "Dossier N"] = e_dossier
                        df_live.at[idx, "Nom"] = e_nom
                        df_live.at[idx, "Date"] = pd.to_datetime(e_date)
                        df_live.at[idx, "Montant honoraires (US $)"] = money_to_float(e_montant)
                        df_live.at[idx, "Autres frais (US $)"] = money_to_float(e_autres)
                        df_live.at[idx, "Acompte 1"] = money_to_float(e_ac1)
                        df_live.at[idx, "Acompte 2"] = money_to_float(e_ac2)
                        df_live.at[idx, "Acompte 3"] = money_to_float(e_ac3)
                        df_live.at[idx, "Acompte 4"] = money_to_float(e_ac4)
                        df_live.at[idx, "Date Acompte 1"] = pd.to_datetime(e_ac1_date) if e_ac1_date else pd.NaT
                        df_live.at[idx, "Date Acompte 2"] = pd.to_datetime(e_ac2_date) if e_ac2_date else pd.NaT
                        df_live.at[idx, "Date Acompte 3"] = pd.to_datetime(e_ac3_date) if e_ac3_date else pd.NaT
                        df_live.at[idx, "Date Acompte 4"] = pd.to_datetime(e_ac4_date) if e_ac4_date else pd.NaT
                        df_live.at[idx, "ModeReglement_Ac1"] = ",".join(e_mode_ac1) if isinstance(e_mode_ac1, (list,tuple)) else str(e_mode_ac1)
                        df_live.at[idx, "ModeReglement_Ac2"] = ",".join(e_mode_ac2) if isinstance(e_mode_ac2, (list,tuple)) else str(e_mode_ac2)
                        df_live.at[idx, "ModeReglement_Ac3"] = ",".join(e_mode_ac3) if isinstance(e_mode_ac3, (list,tuple)) else str(e_mode_ac3)
                        df_live.at[idx, "ModeReglement_Ac4"] = ",".join(e_mode_ac4) if isinstance(e_mode_ac4, (list,tuple)) else str(e_mode_ac4)
                        old_general = _parse_modes(current_row.get("ModeReglement",""))
                        combined = set(old_general + list(e_mode_ac1))
                        df_live.at[idx, "ModeReglement"] = ",".join(sorted(list(combined)))
                        df_live.at[idx, "Dossiers envoyé"] = 1 if e_flag_envoye else 0
                        df_live.at[idx, "Dossier approuvé"] = 1 if e_flag_approuve else 0
                        df_live.at[idx, "Dossier refusé"] = 1 if e_flag_refuse else 0
                        df_live.at[idx, "Dossier Annulé"] = 1 if e_flag_annule else 0
                        if e_flag_rfe and not any([e_flag_envoye, e_flag_approuve, e_flag_refuse, e_flag_annule]):
                            st.warning("RFE n'a pas été activé car aucun état (envoyé/approuvé/refusé/annulé) n'est coché.")
                            df_live.at[idx, "RFE"] = 0
                        else:
                            df_live.at[idx, "RFE"] = 1 if e_flag_rfe else 0
                        df_live.at[idx, "Date d'envoi"] = pd.to_datetime(e_flags_date) if e_flags_date else pd.NaT
                        df_live.at[idx, "Date reponse"] = pd.to_datetime(e_date_reponse) if e_date_reponse else pd.NaT
                        df_live.at[idx, "Escrow"] = 1 if e_escrow else 0
                        df_live.at[idx, "Commentaires"] = e_comments
                        df_live.at[idx, "Dernière modification"] = datetime.now()
                        df_live.at[idx, "Modifié par"] = CURRENT_USER
                        df_live = recalc_payments_and_solde(df_live)
                        df_live.at[idx, "Solde à percevoir (US $)"] = df_live.at[idx, "Solde"]
                        _set_df_live(df_live)
                        st.success("Modifications enregistrées.")
                    except Exception as e:
                        st.error(f"Erreur enregistrement: {e}")

# ---- Export tab (unchanged) ----
with tabs[5]:
    st.header("💾 Export")
    df_live = _get_df_live_safe()
    if df_live is None or df_live.empty:
        st.info("Aucune donnée à exporter.")
    else:
        csv_bytes = df_live.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Export CSV", data=csv_bytes, file_name="Clients_export.csv", mime="text/csv")
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df_live.to_excel(writer, index=False, sheet_name="Clients")
        st.download_button("⬇️ Export XLSX", data=buf.getvalue(), file_name="Clients_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# End of file
