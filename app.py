# Visa Manager - app.py
# Robust version: dynamic detection of Acompte columns and robust Solde recalculation,
# plus dashboard debug output for rows where Solde != Montant + Autres - sum(Acomptes).
#
# Usage: streamlit run app.py
# Requirements: pandas, openpyxl, streamlit; optional: plotly

import os
import json
import re
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional

import pandas as pd
import numpy as np
import streamlit as st

# Optional: plotly (if installed)
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    px = None
    HAS_PLOTLY = False

# For Excel export with formulas
try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

# =========================
# Configuration
# =========================
APP_TITLE = "🛂 Visa Manager"
COLS_CLIENTS = [
    "ID_Client", "Dossier N", "Nom", "Date",
    "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)",
    "Payé", "Solde", "Acompte 1", "Acompte 2",
    "Acompte 3", "Acompte 4",
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

# =========================
# Helpers
# =========================
def normalize_header_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r'^\s+|\s+$', '', s)
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

def money_to_float(x: Any) -> float:
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s == "" or s in ("-", "—", "–", "NA", "N/A"):
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

def _to_num(x: Any) -> float:
    return money_to_float(x) if not isinstance(x, (int, float)) else float(x)

def _fmt_money(v: Any) -> str:
    try:
        return "${:,.2f}".format(float(v))
    except Exception:
        return "$0.00"

def _date_for_widget(val: Any) -> date:
    if isinstance(val, date):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return date.today()
        return d.date()
    except Exception:
        return date.today()

# =========================
# Column heuristics
# =========================
COL_CANDIDATES = {
    "id client": "ID_Client", "idclient": "ID_Client",
    "dossier n": "Dossier N", "dossier": "Dossier N",
    "nom": "Nom", "date": "Date",
    "categories": "Categories", "categorie": "Categories",
    "sous categorie": "Sous-categorie", "sous-categorie": "Sous-categorie", "souscategorie": "Sous-categorie",
    "visa": "Visa",
    "montant": "Montant honoraires (US $)", "montant honoraires": "Montant honoraires (US $)", "honoraires": "Montant honoraires (US $)",
    "autres frais": "Autres frais (US $)", "autresfrais": "Autres frais (US $)",
    "payé": "Payé", "paye": "Payé",
    "solde": "Solde",
    "acompte 1": "Acompte 1", "acompte1": "Acompte 1",
    "acompte 2": "Acompte 2", "acompte2": "Acompte 2",
    "acompte 3": "Acompte 3", "acompte3": "Acompte 3",
    "acompte 4": "Acompte 4", "acompte4": "Acompte 4",
    "dossier envoye": "Dossiers envoyé", "dossier approuve": "Dossier approuvé", "dossier refuse": "Dossier refusé",
    "rfe": "RFE", "commentaires": "Commentaires"
}

NUMERIC_TARGETS = [
    "Montant honoraires (US $)",
    "Autres frais (US $)",
    "Payé",
    "Solde",
    "Acompte 1",
    "Acompte 2",
    "Acompte 3",
    "Acompte 4"
]

def map_columns_heuristic(df: Any) -> Tuple[pd.DataFrame, Dict[str,str]]:
    if not isinstance(df, pd.DataFrame):
        try:
            st.sidebar.warning("map_columns_heuristic: input is not a DataFrame — returning empty DataFrame.")
        except Exception:
            pass
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
        if mapped is None:
            mapped = normalize_header_text(c)
        mapping[c] = mapped
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
        try:
            st.sidebar.error("map_columns_heuristic: rename failed, returning original DataFrame.")
        except Exception:
            pass
        return df, new_names
    return df, new_names

def coerce_category_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    cols = list(df.columns)
    rename_map = {}
    def _ck(x): return canonical_key(str(x))
    for c in cols:
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

# =========================
# I/O helpers
# =========================
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
            try:
                return pd.read_csv(BytesIO(src), sep=";", encoding="utf-8", on_bad_lines="skip")
            except Exception:
                try:
                    return pd.read_csv(BytesIO(src), sep=",", encoding="utf-8", on_bad_lines="skip")
                except Exception:
                    return None
        if isinstance(src, (BytesIO,)):
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
                try:
                    return pd.read_csv(BytesIO(data), sep=";", encoding="utf-8", on_bad_lines="skip")
                except Exception:
                    try:
                        return pd.read_csv(BytesIO(data), sep=",", encoding="utf-8", on_bad_lines="skip")
                    except Exception:
                        return None
        if isinstance(src, (str, os.PathLike)):
            p = str(src)
            if not os.path.exists(p):
                _log(f"path does not exist: {p}")
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
    except Exception as e:
        _log(f"read_any_table exception: {e}")
        return None
    _log("read_any_table: unsupported src type")
    return None

# =========================
# Utilities to detect columns robustly
# =========================
def detect_acompte_columns(df: pd.DataFrame) -> List[str]:
    # find all columns whose canonical key contains "acompte" or starts with 'acompte'
    cols = []
    for c in df.columns:
        k = canonical_key(c)
        if "acompte" in k:
            cols.append(c)
    # sort by numeric suffix if possible (Acompte 1..4)
    def sort_key(name):
        m = re.search(r"(\d+)", name)
        return int(m.group(1)) if m else 999
    cols = sorted(cols, key=sort_key)
    return cols

def detect_montant_column(df: pd.DataFrame) -> Optional[str]:
    # prefer exact known name
    candidates = ["Montant honoraires (US $)", "Montant honoraires", "Montant", "Montant honoraires (USD)"]
    for c in candidates:
        if c in df.columns:
            return c
    # fallback: any column with 'montant' or 'honoraires' in canonical key
    for c in df.columns:
        k = canonical_key(c)
        if "montant" in k or "honorair" in k or "honoraires" in k:
            return c
    return None

def detect_autres_column(df: pd.DataFrame) -> Optional[str]:
    candidates = ["Autres frais (US $)", "Autres frais", "Autres"]
    for c in candidates:
        if c in df.columns:
            return c
    for c in df.columns:
        k = canonical_key(c)
        if "autre" in k or "autres" in k or "frais" in k:
            return c
    return None

# =========================
# Core normalization function (keeps prior behavior)
# =========================
def normalize_clients_for_live(df_clients_raw: Any) -> pd.DataFrame:
    if not isinstance(df_clients_raw, pd.DataFrame):
        try:
            if 'read_any_table' in globals() and callable(globals()['read_any_table']):
                maybe_df = read_any_table(df_clients_raw, sheet=None, debug_prefix="[normalize] ")
                if isinstance(maybe_df, pd.DataFrame):
                    df_clients_raw = maybe_df
                else:
                    df_clients_raw = pd.DataFrame()
        except Exception:
            df_clients_raw = pd.DataFrame()
    if df_clients_raw is None or not isinstance(df_clients_raw, pd.DataFrame):
        df_clients_raw = pd.DataFrame()

    try:
        df_mapped, _ = map_columns_heuristic(df_clients_raw)
    except Exception:
        df_mapped = df_clients_raw.copy() if isinstance(df_clients_raw, pd.DataFrame) else pd.DataFrame()

    if "Date" in df_mapped.columns:
        try:
            df_mapped["Date"] = pd.to_datetime(df_mapped["Date"], dayfirst=True, errors="coerce")
        except Exception:
            pass

    df = _ensure_columns(df_mapped, COLS_CLIENTS)

    # normalize numeric targets defensively
    for col in NUMERIC_TARGETS:
        if col in df.columns:
            try:
                df[col] = df[col].apply(lambda x: _to_num(x) if not isinstance(x,(int,float)) else float(x))
            except Exception:
                try:
                    df[col] = df[col].apply(lambda x: 0.0)
                except Exception:
                    pass

    # ensure acomptes exist
    for acc in ["Acompte 1","Acompte 2","Acompte 3","Acompte 4"]:
        if acc not in df.columns:
            df[acc] = 0.0

    # compute Payé from acomptes and Solde
    acomptes_cols = detect_acompte_columns(df)
    if acomptes_cols:
        try:
            df["Payé"] = df[acomptes_cols].fillna(0).apply(lambda row: sum([_to_num(row[c]) for c in acomptes_cols]), axis=1)
        except Exception:
            try:
                df["Payé"] = df.get("Payé", 0).apply(lambda x: _to_num(x))
            except Exception:
                pass

    try:
        montant_col = detect_montant_column(df) or "Montant honoraires (US $)"
        autres_col = detect_autres_column(df) or "Autres frais (US $)"
        df["Solde"] = df.get(montant_col,0).apply(_to_num) + df.get(autres_col,0).apply(_to_num) - df.get("Payé",0).apply(_to_num)
    except Exception:
        try:
            df["Solde"] = df.get("Solde",0).apply(lambda x: _to_num(x))
        except Exception:
            df["Solde"] = 0.0

    df = _normalize_status(df)

    for c in ["Nom", "Categories", "Sous-categorie", "Visa", "Commentaires"]:
        if c in df.columns:
            try:
                df[c] = df[c].astype(str).fillna("")
            except Exception:
                df[c] = df[c].fillna("").astype(str)

    try:
        dser = pd.to_datetime(df["Date"], errors="coerce")
        df["_Année_"] = dser.dt.year.fillna(0).astype(int)
        df["_MoisNum_"] = dser.dt.month.fillna(0).astype(int)
        df["Mois"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) and m>0 else "")
    except Exception:
        df["_Année_"] = 0; df["_MoisNum_"] = 0; df["Mois"] = ""

    return df

# =========================
# New robust recalc function
# =========================
def recalc_payments_and_solde(df: pd.DataFrame) -> pd.DataFrame:
    """
    Robust recalculation:
    - detect acomptes columns dynamically (any column whose canonical key contains 'acompte')
    - detect montant and autres columns dynamically
    - compute Payé = sum(acomptes) and Solde = Montant + Autres - Payé
    - return copy
    """
    if df is None or df.empty:
        return df
    out = df.copy()

    # detect acomptes (dynamic)
    acomptes = detect_acompte_columns(out)
    # ensure there is at least a placeholder for standard acomptes if none detected
    if not acomptes:
        for acc in ["Acompte 1","Acompte 2","Acompte 3","Acompte 4"]:
            if acc not in out.columns:
                out[acc] = 0.0
        acomptes = detect_acompte_columns(out)

    # cast acomptes to numeric
    for c in acomptes:
        try:
            out[c] = out[c].apply(lambda x: _to_num(x) if not isinstance(x,(int,float)) else float(x))
        except Exception:
            out[c] = out[c].apply(lambda x: 0.0)

    # detect montant and autres
    montant_col = detect_montant_column(out) or "Montant honoraires (US $)"
    autres_col = detect_autres_column(out) or "Autres frais (US $)"

    # ensure numeric
    for c in [montant_col, autres_col]:
        if c not in out.columns:
            out[c] = 0.0
        else:
            try:
                out[c] = out[c].apply(lambda x: _to_num(x) if not isinstance(x,(int,float)) else float(x))
            except Exception:
                out[c] = out[c].apply(lambda x: 0.0)

    # compute paid sum and store to Payé (overwrite)
    try:
        if acomptes:
            out["Payé"] = out[acomptes].sum(axis=1).astype(float)
        else:
            out["Payé"] = out.get("Payé",0).apply(lambda x: _to_num(x))
    except Exception:
        out["Payé"] = out.get("Payé",0).apply(lambda x: _to_num(x))

    # compute Solde and enforce float
    try:
        out["Solde"] = out[montant_col] + out[autres_col] - out["Payé"]
        out["Solde"] = out["Solde"].astype(float)
        out["Payé"] = out["Payé"].astype(float)
    except Exception:
        try:
            out["Solde"] = out.get("Solde",0).apply(lambda x: _to_num(x))
        except Exception:
            out["Solde"] = 0.0

    return out

# =========================
# Next ID & flags helpers
# =========================
def get_next_client_id(df: pd.DataFrame) -> int:
    if df is None or df.empty:
        return DEFAULT_START_CLIENT_ID
    vals = df.get("ID_Client", pd.Series([], dtype="object"))
    nums = []
    for v in vals.dropna().astype(str):
        m = re.search(r"(\d+)", v)
        if m:
            try:
                nums.append(int(m.group(1)))
            except Exception:
                pass
    if not nums:
        return DEFAULT_START_CLIENT_ID
    mx = max(nums)
    return max(DEFAULT_START_CLIENT_ID, mx) + 1

def ensure_flag_columns(df: pd.DataFrame, flags: List[str]) -> None:
    for f in flags:
        if f not in df.columns:
            df[f] = 0

DEFAULT_FLAGS = ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]

# =========================
# UI bootstrap / I/O / Tabs
# =========================
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.sidebar.header("📂 Fichiers")
last_clients, last_visa, last_save_dir = ("", "", "")
try:
    if os.path.exists(MEMO_FILE):
        with open(MEMO_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
            last_clients = d.get("clients",""); last_visa = d.get("visa",""); last_save_dir = d.get("save_dir","")
except Exception:
    pass

mode = st.sidebar.radio("Mode de chargement", ["Un fichier (Clients)", "Deux fichiers (Clients & Visa)"], index=0, key=skey("mode"))
up_clients = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx","xls","csv"], key=skey("up_clients"))
up_visa = None
if mode == "Deux fichiers (Clients & Visa)":
    up_visa = st.sidebar.file_uploader("Visa (xlsx/csv)", type=["xlsx","xls","csv"], key=skey("up_visa"))

clients_path_in = st.sidebar.text_input("ou chemin local Clients", value=last_clients or "", key=skey("cli_path"))
visa_path_in = st.sidebar.text_input("ou chemin local Visa", value=last_visa or "", key=skey("vis_path"))
save_dir_in = st.sidebar.text_input("Dossier de sauvegarde (optionnel)", value=last_save_dir or "", key=skey("save_dir"))

if st.sidebar.button("📥 Sauvegarder chemins", key=skey("btn_save_paths")):
    try:
        with open(MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump({"clients": clients_path_in or "", "visa": visa_path_in or "", "save_dir": save_dir_in or ""}, f, ensure_ascii=False, indent=2)
        st.sidebar.success("Chemins sauvegardés.")
    except Exception:
        st.sidebar.error("Impossible de sauvegarder les chemins.")

# Read uploaded files into bytes
clients_bytes = None
visa_bytes = None
if up_clients is not None:
    try:
        clients_bytes = up_clients.getvalue()
    except Exception:
        try:
            up_clients.seek(0); clients_bytes = up_clients.read()
        except Exception:
            clients_bytes = None
if up_visa is not None:
    try:
        visa_bytes = up_visa.getvalue()
    except Exception:
        try:
            up_visa.seek(0); visa_bytes = up_visa.read()
        except Exception:
            visa_bytes = None

if clients_bytes is not None:
    clients_src_for_read = BytesIO(clients_bytes)
elif clients_path_in:
    clients_src_for_read = clients_path_in
elif last_clients:
    clients_src_for_read = last_clients
else:
    clients_src_for_read = None

if mode == "Deux fichiers (Clients & Visa)":
    if visa_bytes is not None:
        visa_src_for_read = BytesIO(visa_bytes)
    elif visa_path_in:
        visa_src_for_read = visa_path_in
    elif last_visa:
        visa_src_for_read = last_visa
    else:
        visa_src_for_read = None
else:
    visa_src_for_read = clients_src_for_read

df_clients_raw = None
df_visa_raw = None
try:
    df_clients_raw = read_any_table(clients_src_for_read, sheet=SHEET_CLIENTS, debug_prefix="[Clients] ")
except Exception:
    df_clients_raw = None
if df_clients_raw is None and clients_src_for_read is not None:
    df_clients_raw = read_any_table(clients_src_for_read, sheet=None, debug_prefix="[Clients fallback] ")
if df_clients_raw is None:
    df_clients_raw = pd.DataFrame()

try:
    df_visa_raw = read_any_table(visa_src_for_read, sheet=SHEET_VISA, debug_prefix="[Visa] ")
except Exception:
    df_visa_raw = None
if df_visa_raw is None and visa_src_for_read is not None:
    df_visa_raw = read_any_table(visa_src_for_read, sheet=None, debug_prefix="[Visa fallback] ")
if df_visa_raw is None:
    df_visa_raw = pd.DataFrame()

# SANITIZE Visa raw dataframe
if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
    try:
        df_visa_raw = df_visa_raw.copy()
        df_visa_raw = df_visa_raw.fillna("")  # replace NaN with empty string
        for c in df_visa_raw.columns:
            try:
                df_visa_raw[c] = df_visa_raw[c].astype(str).str.strip()
                df_visa_raw[c] = df_visa_raw[c].replace(r'^\s*nan\s*$', "", regex=True, case=False)
            except Exception:
                pass
    except Exception:
        pass

# Build visa maps
visa_map = {}; visa_map_norm = {}; visa_categories = []; visa_sub_options_map = {}
if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
    try:
        df_visa_mapped, _ = map_columns_heuristic(df_visa_raw)
        try:
            df_visa_mapped = coerce_category_columns(df_visa_mapped)
        except Exception:
            pass
        raw_vm = build_visa_map(df_visa_mapped)
        raw_vm = {k: [s for s in v if s and str(s).strip().lower() != "nan"] for k, v in raw_vm.items() if k and str(k).strip() != "" and str(k).strip().lower() != "nan"}
        visa_map = {k.strip(): [s.strip() for s in v] for k, v in raw_vm.items()}
        visa_map_norm = {canonical_key(k): v for k, v in visa_map.items()}
        visa_categories = sorted(list(visa_map.keys()))
        visa_sub_options_map = build_sub_options_map_from_flags(df_visa_mapped)
        visa_sub_options_map = {k: [x for x in v if x and str(x).strip() != "" and str(x).strip().lower() != "nan"] for k, v in visa_sub_options_map.items() if k and str(k).strip() != "" and str(k).strip().lower() != "nan"}
    except Exception as e:
        st.sidebar.error(f"Erreur build visa maps: {e}")
        visa_map = {}; visa_map_norm = {}; visa_categories = []; visa_sub_options_map = {}
else:
    visa_map = {}; visa_map_norm = {}; visa_categories = []; visa_sub_options_map = {}

# Build live df and enforce canonical Payé/Solde
df_all = normalize_clients_for_live(df_clients_raw)
df_all = recalc_payments_and_solde(df_all)
DF_LIVE_KEY = skey("df_live")
if isinstance(df_all, pd.DataFrame) and not df_all.empty:
    st.session_state[DF_LIVE_KEY] = df_all.copy()
else:
    if DF_LIVE_KEY not in st.session_state or st.session_state[DF_LIVE_KEY] is None:
        st.session_state[DF_LIVE_KEY] = pd.DataFrame(columns=COLS_CLIENTS)

def _get_df_live() -> pd.DataFrame:
    return st.session_state[DF_LIVE_KEY].copy()

def _set_df_live(df: pd.DataFrame) -> None:
    st.session_state[DF_LIVE_KEY] = df.copy()

# Helper unique non-empty
def unique_nonempty(series):
    try:
        vals = series.dropna().astype(str).tolist()
    except Exception:
        vals = []
    out = []
    for v in vals:
        if v is None:
            continue
        s = str(v).strip()
        if s == "" or s.lower() == "nan":
            continue
        out.append(s)
    return sorted(list(dict.fromkeys(out)))

# Helper to build compact KPI HTML
def kpi_html(label: str, value: str, sub: str = "") -> str:
    html = f"""
    <div style="border:1px solid rgba(255,255,255,0.04); border-radius:6px; padding:8px 10px; margin:6px 4px; background:transparent;">
      <div style="font-size:12px; color:#a8b3c0;">{label}</div>
      <div style="font-size:18px; font-weight:700; margin-top:4px; color:#ffffff;">{value}</div>
      <div style="font-size:11px; color:#9aa9b7; margin-top:4px;">{sub}</div>
    </div>
    """
    return html

# =========================
# Tabs and UI (Dashboard includes extra debug for Solde mismatch)
# =========================
tabs = st.tabs(["📄 Fichiers","📊 Dashboard","📈 Analyses","➕ / ✏️ / 🗑️ Gestion","💾 Export"])

# ---- Files tab ----
with tabs[0]:
    st.header("📂 Fichiers")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Clients")
        if up_clients is not None:
            st.text(f"Upload: {getattr(up_clients,'name','')}")
        elif isinstance(clients_src_for_read, str) and clients_src_for_read:
            st.text(f"Chargé depuis: {clients_src_for_read}")
        if df_clients_raw is None or df_clients_raw.empty:
            st.warning("Aucun fichier Clients detecté.")
        else:
            st.success(f"Clients lus: {df_clients_raw.shape[0]} lignes")
            st.dataframe(df_clients_raw.head(8), use_container_width=True, height=240)
    with c2:
        st.subheader("Visa")
        if mode == "Deux fichiers (Clients & Visa)":
            if up_visa is not None:
                st.text(f"Upload: {getattr(up_visa,'name','')}")
            elif isinstance(visa_src_for_read, str) and visa_src_for_read:
                st.text(f"Chargé depuis: {visa_src_for_read}")
        else:
            st.info("Mode 'Un fichier' : Visa sera lu depuis le fichier Clients si présent.")
        if df_visa_raw is None or df_visa_raw.empty:
            st.warning("Aucun fichier Visa detecté.")
        else:
            st.success(f"Visa lu: {df_visa_raw.shape[0]} lignes, {df_visa_raw.shape[1]} colonnes")
            st.dataframe(df_visa_raw.head(8), use_container_width=True, height=240)
    st.markdown("---")
    col_a, col_b = st.columns([1,1])
    with col_a:
        if st.button("Réinitialiser mémoire (recharger)"):
            df_all2 = normalize_clients_for_live(df_clients_raw)
            df_all2 = recalc_payments_and_solde(df_all2)
            _set_df_live(df_all2)
            st.success("Mémoire réinitialisée.")
            try:
                st.experimental_rerun()
            except Exception:
                pass
    with col_b:
        if st.button("Actualiser la lecture"):
            try:
                st.experimental_rerun()
            except Exception:
                pass

# ---- Dashboard tab ----
with tabs[1]:
    st.subheader("📊 Dashboard (totaux et diagnostics)")
    df_live_view = recalc_payments_and_solde(_get_df_live())

    if df_live_view is None or df_live_view.empty:
        st.info("Aucune donnée en mémoire.")
    else:
        cats = unique_nonempty(df_live_view["Categories"]) if "Categories" in df_live_view.columns else []
        subs = unique_nonempty(df_live_view["Sous-categorie"]) if "Sous-categorie" in df_live_view.columns else []
        visas = unique_nonempty(df_live_view["Visa"]) if "Visa" in df_live_view.columns else []
        years = []
        if "_Année_" in df_live_view.columns:
            try:
                years_vals = pd.to_numeric(df_live_view["_Année_"], errors="coerce").dropna().unique().astype(int).tolist()
                years = sorted([int(y) for y in years_vals])
            except Exception:
                years = []

        # Filters
        f1, f2, f3, f4 = st.columns([1,1,1,1])
        sel_cat = f1.selectbox("Catégorie", options=[""]+cats, index=0, key=skey("dash","cat"))
        sel_sub = f2.selectbox("Sous-catégorie", options=[""]+subs, index=0, key=skey("dash","sub"))
        sel_visa = f3.selectbox("Visa", options=[""]+visas, index=0, key=skey("dash","visa"))
        year_options = ["Toutes les années"] + [str(y) for y in years]
        sel_year = f4.selectbox("Année", options=year_options, index=0, key=skey("dash","year"))

        # canonical copy then apply filters
        view = df_live_view.copy()
        if sel_cat:
            view = view[view["Categories"].astype(str) == sel_cat]
        if sel_sub:
            view = view[view["Sous-categorie"].astype(str) == sel_sub]
        if sel_visa:
            view = view[view["Visa"].astype(str) == sel_visa]
        if sel_year and sel_year != "Toutes les années":
            view = view[view["_Année_"].astype(str) == sel_year]

        # recompute canonical values on filtered view (safe)
        view = recalc_payments_and_solde(view)

        # prepare numeric fields
        def safe_num(x):
            try:
                return float(_to_num(x))
            except Exception:
                return 0.0

        montant_col = detect_montant_column(view) or "Montant honoraires (US $)"
        autres_col = detect_autres_column(view) or "Autres frais (US $)"
        acomptes_cols = detect_acompte_columns(view)

        view["_Montant_num_"] = view.get(montant_col, 0).apply(safe_num)
        view["_Autres_num_"] = view.get(autres_col, 0).apply(safe_num)
        for i, c in enumerate(acomptes_cols):
            view[f"_acompte_num_{i}"] = view.get(c, 0).apply(safe_num)
        view["_Acomptes_sum_"] = view[[f"_acompte_num_{i}" for i in range(len(acomptes_cols))]].sum(axis=1) if acomptes_cols else 0.0
        view["_Payé_num_"] = view.get("Payé", 0).apply(safe_num)
        # canonical solde from numeric parts
        canonical_solde_sum = (view["_Montant_num_"] + view["_Autres_num_"] - view["_Payé_num_"]).sum()

        # totals
        total_honoraires = view["_Montant_num_"].sum()
        total_autres = view["_Autres_num_"].sum()
        total_facture_calc = total_honoraires + total_autres
        total_paye = view["_Payé_num_"].sum()
        total_acomptes_sum = view["_Acomptes_sum_"].sum()
        total_solde_recorded = view.get("Solde", 0).apply(safe_num).sum()

        # KPIs
        cols_k = st.columns(4)
        cols_k[0].markdown(kpi_html("Dossiers (vue)", f"{len(view):,}"), unsafe_allow_html=True)
        cols_k[1].markdown(kpi_html("Montant honoraires", _fmt_money(total_honoraires)), unsafe_allow_html=True)
        cols_k[2].markdown(kpi_html("Autres frais", _fmt_money(total_autres)), unsafe_allow_html=True)
        cols_k[3].markdown(kpi_html("Total facturé (recalc)", _fmt_money(total_facture_calc)), unsafe_allow_html=True)

        st.markdown("---")
        cols_k2 = st.columns(2)
        cols_k2[0].markdown(kpi_html("Montant payé (Payé = somme acomptes)", _fmt_money(total_paye)), unsafe_allow_html=True)
        cols_k2[1].markdown(kpi_html("Solde total (recalc)", _fmt_money(canonical_solde_sum)), unsafe_allow_html=True)

        # Diagnostics if mismatch
        if abs(canonical_solde_sum - total_solde_recorded) > 0.005 or abs(total_paye - total_acomptes_sum) > 0.005:
            st.warning("Diagnostics : écart détecté entre valeurs recalculées et colonnes enregistrées.")
            st.write({
                "total_honoraires": float(total_honoraires),
                "total_autres": float(total_autres),
                "total_paye (col Payé)": float(total_paye),
                "total_acomptes_sum (sum Acompte cols)": float(total_acomptes_sum),
                "canonical_solde_sum (calc)": float(canonical_solde_sum),
                "total_solde_recorded (col Solde)": float(total_solde_recorded),
                "rows_shown": int(len(view)),
                "acompte_columns_detected": acomptes_cols,
                "montant_column": montant_col,
                "autres_column": autres_col
            })

        # EXPANDER: show rows where solde differs from calculated
        view["_Solde_calc_row_"] = view["_Montant_num_"] + view["_Autres_num_"] - view["_Payé_num_"]
        mismatches = view[(view.get("Solde",0).apply(safe_num) - view["_Solde_calc_row_"]).abs() > 0.005]
        with st.expander("DEBUG — Lignes où Solde != Montant + Autres − somme(Acomptes)"):
            if mismatches.empty:
                st.write("Aucune ligne en écart détectée.")
            else:
                # prepare display
                disp_cols = ["ID_Client","Dossier N","Nom","Date","Categories","Sous-categorie","Visa",
                             montant_col, autres_col] + acomptes_cols + ["Payé","Solde"]
                # Add computed numeric columns
                for c in ["_Montant_num_","_Autres_num_","_Acomptes_sum_","_Payé_num_","_Solde_calc_row_"]:
                    if c in mismatches.columns:
                        disp_cols.append(c)
                # show types & repr of problematic cells for first 20 rows
                mshow = mismatches.reset_index(drop=True)
                # Also add a column with raw values of acomptes concatenated for diagnosis
                try:
                    mshow["_acomptes_raw_concat_"] = mshow[acomptes_cols].astype(str).agg(" | ".join, axis=1)
                except Exception:
                    mshow["_acomptes_raw_concat_"] = ""
                st.dataframe(mshow[disp_cols + ["_acomptes_raw_concat_"]].head(200), use_container_width=True, height=360)
                st.markdown("Si vous voulez, téléchargez la vue filtrée (bouton Export) et envoyez-moi le fichier ou copiez quelques lignes du tableau ci‑dessous pour que j'analyse précisément.")

        # List clients for filters
        st.markdown("### Détails — clients correspondant aux filtres")
        display_df = view.copy()
        if "Date" in display_df.columns:
            try:
                display_df["Date"] = pd.to_datetime(display_df["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                display_df["Date"] = display_df["Date"].astype(str)
        money_cols = [montant_col, autres_col, "Payé","Solde"] + acomptes_cols
        for mc in money_cols:
            if mc in display_df.columns:
                try:
                    display_df[mc] = display_df[mc].apply(lambda x: _fmt_money(_to_num(x)))
                except Exception:
                    display_df[mc] = display_df[mc].astype(str)
        try:
            st.dataframe(display_df.reset_index(drop=True), use_container_width=True, height=360)
        except Exception:
            st.write("Impossible d'afficher la liste des clients (trop volumineuse). Utilisez l'export pour récupérer les données filtrées.")

# ---- Export and other tabs remain the same as before (omitted here for brevity) ----
# The rest of the app (Analyses / Gestion / Export) is unchanged from previous working version.
# If you want I can paste the full file including exports again, but main change is the robust recalc and debug above.
