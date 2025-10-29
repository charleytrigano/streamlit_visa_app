# Visa Manager - app.py
# Final updated app with corrections applied to Dashboard solde calculation,
# Payé enforced from Acompte 1..4, smaller KPI cards, Visa filter included,
# and recalc_payments_and_solde applied before filtering to ensure canonical totals.
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
# Small helpers (normalization / formatting)
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
# Column heuristics & helpers
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
# Visa mapping utilities
# =========================
def build_visa_map(dfv: pd.DataFrame) -> Dict[str, List[str]]:
    vm: Dict[str, List[str]] = {}
    if dfv is None or dfv.empty:
        return vm
    df = dfv.copy()
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
    out: Dict[str, List[str]] = {}
    if dfv is None or dfv.empty:
        return out
    df = dfv.copy()
    cols_to_skip = set(["Categories", "Categorie", "Sous-categorie"])
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

def get_sub_options_for(sub_value: str, visa_sub_options_map: Dict[str, List[str]]) -> List[str]:
    if not sub_value or not isinstance(sub_value, str):
        return []
    s_raw = sub_value.strip()
    s_can = canonical_key(s_raw)
    s_lower = s_raw.lower()
    s_noacc = remove_accents(s_raw).lower()
    if s_can in visa_sub_options_map:
        return visa_sub_options_map[s_can][:]
    if s_lower in visa_sub_options_map:
        return visa_sub_options_map[s_lower][:]
    if s_noacc in visa_sub_options_map:
        return visa_sub_options_map[s_noacc][:]
    s_match = remove_accents(s_raw).lower()
    candidates = []
    for k in visa_sub_options_map.keys():
        if s_match in remove_accents(k).lower() or remove_accents(k).lower() in s_match:
            candidates.extend(visa_sub_options_map.get(k, []))
    if candidates:
        seen = set(); out = []
        for c in candidates:
            if c not in seen:
                seen.add(c); out.append(c)
        return out
    try:
        st.sidebar.markdown("DEBUG get_sub_options_for tries (none matched):")
        st.sidebar.write([s_can, s_lower, s_noacc])
    except Exception:
        pass
    return []

# =========================
# Column/normalize helpers
# =========================
def _ensure_columns(df: Any, cols: List[str]) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame):
        try:
            st.sidebar.warning("_ensure_columns: input was not a DataFrame; coercing to empty DataFrame.")
        except Exception:
            pass
        df = pd.DataFrame()
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            if c in ["Payé", "Solde", "Montant honoraires (US $)", "Autres frais (US $)", "Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
                out[c] = 0.0
            elif c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
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
                safe[c] = 0.0 if c in ["Payé", "Solde", "Montant honoraires (US $)", "Autres frais (US $)", "Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"] else ""
        return safe

def _normalize_status(df: Any) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame):
        try:
            st.sidebar.warning("_normalize_status: input was not a DataFrame; coercing to empty DataFrame.")
        except Exception:
            pass
        df = pd.DataFrame()
    cols_status = ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]
    for c in cols_status:
        try:
            if c in df.columns:
                s = df[c]
                if not isinstance(s, pd.Series):
                    s = pd.Series([s] * len(df), index=df.index) if len(df) > 0 else pd.Series(dtype="float64")
                def _to_flag(v):
                    try:
                        if pd.isna(v):
                            return 0
                        vs = str(v).strip().lower()
                        if vs in ("1", "x", "t", "true", "oui", "o", "yes", "y"):
                            return 1
                        try:
                            if float(v) == 1.0:
                                return 1
                        except Exception:
                            pass
                        return 0
                    except Exception:
                        return 0
                df[c] = s.map(_to_flag).astype(int)
            else:
                df[c] = 0
        except Exception as ex:
            try:
                st.sidebar.error(f"_normalize_status: erreur sur colonne '{c}': {ex}")
            except Exception:
                pass
            df[c] = 0
    return df

def normalize_clients_for_live(df_clients_raw: Any) -> pd.DataFrame:
    """
    Coerce input to DataFrame, map headers, normalize numeric columns, ensure acomptes exist,
    compute Payé = sum(Acompte 1..4) and Solde = Montant + Autres - Payé.
    """
    if not isinstance(df_clients_raw, pd.DataFrame):
        try:
            if 'read_any_table' in globals() and callable(globals()['read_any_table']):
                maybe_df = read_any_table(df_clients_raw, sheet=None, debug_prefix="[normalize] ")
                if isinstance(maybe_df, pd.DataFrame):
                    df_clients_raw = maybe_df
                else:
                    df_clients_raw = pd.DataFrame()
            else:
                if isinstance(df_clients_raw, (bytes, bytearray)):
                    df_try = try_read_excel_from_bytes(bytes(df_clients_raw))
                    if isinstance(df_try, pd.DataFrame):
                        df_clients_raw = df_try
                    else:
                        df_clients_raw = pd.DataFrame()
                else:
                    df_clients_raw = pd.DataFrame()
        except Exception:
            df_clients_raw = pd.DataFrame()
    if df_clients_raw is None or not isinstance(df_clients_raw, pd.DataFrame):
        df_clients_raw = pd.DataFrame()

    try:
        df_mapped, _ = map_columns_heuristic(df_clients_raw)
        if df_mapped is None or not isinstance(df_mapped, pd.DataFrame):
            df_mapped = pd.DataFrame()
    except Exception:
        df_mapped = df_clients_raw.copy() if isinstance(df_clients_raw, pd.DataFrame) else pd.DataFrame()

    if "Date" in df_mapped.columns:
        try:
            df_mapped["Date"] = pd.to_datetime(df_mapped["Date"], dayfirst=True, errors="coerce")
        except Exception:
            pass

    df = _ensure_columns(df_mapped, COLS_CLIENTS)

    # Numeric normalization for known numeric targets including acomptes
    for col in NUMERIC_TARGETS:
        if col in df.columns:
            try:
                df[col] = df[col].apply(lambda x: _to_num(x) if not isinstance(x, (int, float)) else float(x))
            except Exception:
                try:
                    df[col] = df[col].apply(lambda x: 0.0)
                except Exception:
                    pass

    # Ensure acomptes exist (zeros if missing)
    for acc in ["Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
        if acc not in df.columns:
            df[acc] = 0.0

    # Compute Payé as sum of acomptes (enforce canonical meaning)
    try:
        df["Payé"] = df[["Acompte 1","Acompte 2","Acompte 3","Acompte 4"]].fillna(0).apply(lambda row: float(_to_num(row["Acompte 1"]) + _to_num(row["Acompte 2"]) + _to_num(row["Acompte 3"]) + _to_num(row["Acompte 4"])), axis=1)
    except Exception:
        # fallback: attempt column-wise sum
        try:
            df["Payé"] = _to_num(df.get("Acompte 1",0)) + _to_num(df.get("Acompte 2",0)) + _to_num(df.get("Acompte 3",0)) + _to_num(df.get("Acompte 4",0))
        except Exception:
            df["Payé"] = df.get("Payé",0).apply(lambda x: _to_num(x) if not isinstance(x,(int,float)) else float(x))

    # Recompute Solde as Montant + Autres - Payé
    try:
        df["Solde"] = df.get("Montant honoraires (US $)",0).apply(_to_num) + df.get("Autres frais (US $)",0).apply(_to_num) - df.get("Payé",0).apply(_to_num)
    except Exception:
        try:
            df["Solde"] = 0.0
        except Exception:
            pass

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

# --- New utility: ensure Payé and Solde canonical after any change ---
def recalc_payments_and_solde(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure acomptes numeric, compute Payé = sum(Acompte 1..4),
    and Solde = Montant honoraires + Autres frais - Payé.
    Returns a copied & updated DataFrame.
    """
    if df is None or df.empty:
        return df
    out = df.copy()

    # Ensure acomptes columns exist and cast to numeric
    for acc in ["Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
        if acc not in out.columns:
            out[acc] = 0.0
        else:
            out[acc] = out[acc].apply(lambda x: _to_num(x) if not isinstance(x, (int, float)) else float(x))

    # Ensure Montant / Autres are numeric
    for mc in ["Montant honoraires (US $)", "Autres frais (US $)"]:
        if mc not in out.columns:
            out[mc] = 0.0
        else:
            out[mc] = out[mc].apply(lambda x: _to_num(x) if not isinstance(x, (int, float)) else float(x))

    # Compute Payé from acomptes (overwrite to keep canonical)
    try:
        out["Payé"] = out[["Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]].sum(axis=1).astype(float)
    except Exception:
        out["Payé"] = out.get("Payé", 0).apply(lambda x: _to_num(x) if not isinstance(x, (int, float)) else float(x))

    # Recompute Solde
    out["Solde"] = out["Montant honoraires (US $)"] + out["Autres frais (US $)"] - out["Payé"]

    # enforce types
    try:
        out["Payé"] = out["Payé"].astype(float)
        out["Solde"] = out["Solde"].astype(float)
    except Exception:
        pass

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
# Read files and build maps / UI bootstrap
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

# Build visa maps with filtering of empty values
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

# Debug expander in sidebar
with st.sidebar.expander("DEBUG Visa / Maps", expanded=False):
    if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
        st.markdown("**Visa raw columns (preview):**")
        st.write(list(df_visa_raw.columns)[:80])
    else:
        st.write("Aucun Visa chargé.")
    st.markdown("**visa_map_norm (category key -> subs)**")
    st.write(visa_map_norm)
    st.markdown("**visa_sub_options_map (sous_key -> checkbox labels)**")
    st.write(visa_sub_options_map)

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
# Tabs and UI
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
    # Always start from canonical dataframe (recalc) to avoid drift
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

        # Filters: Category, Sous-categorie, Visa, Year
        f1, f2, f3, f4 = st.columns([1,1,1,1])
        sel_cat = f1.selectbox("Catégorie", options=[""]+cats, index=0, key=skey("dash","cat"))
        sel_sub = f2.selectbox("Sous-catégorie", options=[""]+subs, index=0, key=skey("dash","sub"))
        sel_visa = f3.selectbox("Visa", options=[""]+visas, index=0, key=skey("dash","visa"))
        year_options = ["Toutes les années"] + [str(y) for y in years]
        sel_year = f4.selectbox("Année", options=year_options, index=0, key=skey("dash","year"))

        # Start from a canonical copy then apply filters
        view = df_live_view.copy()
        # (recalc already applied to df_live_view above)
        if sel_cat:
            view = view[view["Categories"].astype(str) == sel_cat]
        if sel_sub:
            view = view[view["Sous-categorie"].astype(str) == sel_sub]
        if sel_visa:
            view = view[view["Visa"].astype(str) == sel_visa]
        if sel_year and sel_year != "Toutes les années":
            view = view[view["_Année_"].astype(str) == sel_year]

        # Defensive numeric conversions
        def safe_num(x):
            try:
                return float(_to_num(x))
            except Exception:
                return 0.0

        # Numeric prepared fields (guaranteed numeric)
        view["_Montant_num_"] = view.get("Montant honoraires (US $)", 0).apply(safe_num)
        view["_Autres_num_"] = view.get("Autres frais (US $)", 0).apply(safe_num)

        # Ensure acomptes present and numeric
        for acc in ["Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
            if acc not in view.columns:
                view[acc] = 0.0
        view["_Acompte1_"] = view.get("Acompte 1", 0).apply(safe_num)
        view["_Acompte2_"] = view.get("Acompte 2", 0).apply(safe_num)
        view["_Acompte3_"] = view.get("Acompte 3", 0).apply(safe_num)
        view["_Acompte4_"] = view.get("Acompte 4", 0).apply(safe_num)

        # Payé taken from canonical column (recalc_payments_and_solde enforced it)
        view["_Payé_num_"] = view.get("Payé", 0).apply(safe_num)

        # Totals: compute canonical totals from numeric columns
        total_honoraires = view["_Montant_num_"].sum()
        total_autres = view["_Autres_num_"].sum()
        total_facture_calc = total_honoraires + total_autres
        total_paye = view["_Payé_num_"].sum()

        # canonical solde computed from numeric columns (reliable)
        canonical_solde_sum = (view["_Montant_num_"] + view["_Autres_num_"] - view["_Payé_num_"]).sum()

        # For transparency compute acomptes sum
        view["_Acomptes_sum_row_"] = view[["_Acompte1_","_Acompte2_","_Acompte3_","_Acompte4_"]].sum(axis=1)
        total_acomptes_sum = view["_Acomptes_sum_row_"].sum()

        # Render KPIs in columns (compact)
        cols_k = st.columns(4)
        cols_k[0].markdown(kpi_html("Dossiers (vue)", f"{len(view):,}"), unsafe_allow_html=True)
        cols_k[1].markdown(kpi_html("Montant honoraires", _fmt_money(total_honoraires)), unsafe_allow_html=True)
        cols_k[2].markdown(kpi_html("Autres frais", _fmt_money(total_autres)), unsafe_allow_html=True)
        cols_k[3].markdown(kpi_html("Total facturé (recalc)", _fmt_money(total_facture_calc)), unsafe_allow_html=True)

        st.markdown("---")
        cols_k2 = st.columns(2)
        cols_k2[0].markdown(kpi_html("Montant payé (Payé = somme acomptes)", _fmt_money(total_paye)), unsafe_allow_html=True)
        cols_k2[1].markdown(kpi_html("Solde total (recalc)", _fmt_money(canonical_solde_sum)), unsafe_allow_html=True)

        # Check consistency vs stored Solde column
        total_solde_recorded = view.get("Solde", 0).apply(safe_num).sum()
        if abs(canonical_solde_sum - total_solde_recorded) > 0.005 or abs(total_paye - total_acomptes_sum) > 0.005:
            st.warning("Diagnostics (détail des totaux) :")
            st.write({
                "total_honoraires": float(total_honoraires),
                "total_autres": float(total_autres),
                "total_paye (col Payé)": float(total_paye),
                "total_acomptes_sum (sum Acompte1..4)": float(total_acomptes_sum),
                "canonical_solde_sum (calc)": float(canonical_solde_sum),
                "total_solde_recorded (col Solde)": float(total_solde_recorded),
                "rows_shown": int(len(view))
            })
        else:
            st.success("Solde stocké et solde recalculé sont cohérents.")

        # anomalies and checks
        view["écart_paye_acompte"] = (view["_Payé_num_"] - view["_Acomptes_sum_row_"]).abs()
        view["écart_solde"] = (view.get("Solde", 0).apply(safe_num) - (view["_Montant_num_"] + view["_Autres_num_"] - view["_Payé_num_"])).abs()
        anomalies = view[(view["_Montant_num_"]==0) & (view.get("Montant honoraires (US $)","")!="")]
        anomalies = pd.concat([anomalies, view[view["écart_solde"] > 0.01], view[view["écart_paye_acompte"] > 0.005]]).drop_duplicates()

        with st.expander("Lignes avec problèmes de conversion, écarts ou Payé ≠ somme acomptes"):
            if anomalies.empty:
                st.write("Aucune anomalie détectée.")
            else:
                display_cols = ["ID_Client","Dossier N","Nom","Date","Categories","Sous-categorie","Visa",
                                "Montant honoraires (US $)","Autres frais (US $)","Acompte 1","Acompte 2","Acompte 3","Acompte 4","Payé","Solde",
                                "_Montant_num_","_Autres_num_","_Acomptes_sum_row_","_Payé_num_","écart_solde","écart_paye_acompte"]
                cols = [c for c in display_cols if c in anomalies.columns]
                st.dataframe(anomalies[cols].reset_index(drop=True), use_container_width=True, height=300)

        # ----- list clients matching current filters -----
        st.markdown("### Détails — clients correspondant aux filtres")
        display_df = view.copy()
        if "Date" in display_df.columns:
            try:
                display_df["Date"] = pd.to_datetime(display_df["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                display_df["Date"] = display_df["Date"].astype(str)
        money_cols = ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde","Acompte 1","Acompte 2","Acompte 3","Acompte 4"]
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

# ---- Analyses tab ----
with tabs[2]:
    st.subheader("📈 Analyses")
    st.info("Graphiques et analyses (basics).")
    df_ = _get_df_live()
    if isinstance(df_, pd.DataFrame) and not df_.empty and "Categories" in df_.columns:
        cat_counts = df_["Categories"].value_counts().rename_axis("Categorie").reset_index(name="Nombre")
        if HAS_PLOTLY and px is not None:
            fig = px.pie(cat_counts, names="Categorie", values="Nombre", hole=0.4, title="Répartition par catégorie")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.bar_chart(cat_counts.set_index("Categorie")["Nombre"])

# ---- Gestion tab (Add / Edit / Delete) ----
with tabs[3]:
    st.subheader("➕ / ✏️ / 🗑️ Gestion")
    df_live = _get_df_live()
    for c in COLS_CLIENTS:
        if c not in df_live.columns:
            df_live[c] = "" if c not in NUMERIC_TARGETS else 0.0

    # Build categories options defensively
    if visa_categories:
        categories_options = visa_categories
    else:
        if "Categories" in df_live.columns:
            cats_series = df_live["Categories"].dropna().astype(str).apply(lambda s: s.strip())
            categories_options = sorted([c for c in dict.fromkeys(cats_series) if c and c.lower() != "nan"])
        else:
            categories_options = []

    st.markdown("### Ajouter un dossier")
    st.write("Dossier N | Nom | ID (ligne 1) — Catégorie | Sous-catégorie | Types (ligne 2) — Acomptes et Montants ensuite")

    with st.form(key=skey("form_add")):
        # Row 1
        r1c1, r1c2, r1c3 = st.columns([1.4,2.2,0.8])
        with r1c1:
            add_dossier = st.text_input("Dossier N", value="", placeholder="Ex: D12345", key=skey("add","dossier"))
        with r1c2:
            add_nom = st.text_input("Nom", value="", placeholder="Nom du client", key=skey("add","nom"))
        with r1c3:
            next_id = get_next_client_id(df_live)
            st.markdown(f"**ID_Client**\n{next_id}")

        # Row 2
        r2c1, r2c2, r2c3 = st.columns([1.4,1.8,2.2])
        with r2c1:
            categories_local = [""] + [c.strip() for c in categories_options]
            add_cat = st.selectbox("Catégorie", options=categories_local, index=0, key=skey("add","cat"))
        with r2c2:
            add_sub_options = []
            if isinstance(add_cat, str) and add_cat.strip():
                cat_key = canonical_key(add_cat)
                if cat_key in visa_map_norm:
                    add_sub_options = visa_map_norm.get(cat_key, [])[:]
                else:
                    if add_cat in visa_map:
                        add_sub_options = visa_map.get(add_cat, [])[:]
            if not add_sub_options:
                add_sub_options = sorted({str(x).strip() for x in df_live["Sous-categorie"].dropna().astype(str).tolist()})
            default_sub_index = 1 if add_sub_options else 0
            add_sub = st.selectbox("Sous-catégorie", options=[""] + add_sub_options, index=default_sub_index if default_sub_index < len([""]+add_sub_options) else 0, key=skey("add","sub"))
        with r2c3:
            specific_options = get_sub_options_for(add_sub, visa_sub_options_map)
            checkbox_options = specific_options if specific_options else DEFAULT_FLAGS
            st.markdown("Types / Flags")
            cols_chk = st.columns(2)
            add_flags_state = {}
            for i, opt in enumerate(checkbox_options):
                col_i = cols_chk[i % 2]
                k = skey("add","flag", re.sub(r"\s+","_", opt))
                add_flags_state[opt] = col_i.checkbox(opt, value=False, key=k)

        # Row 3: date/visa/montant
        r3c1, r3c2, r3c3 = st.columns([1.2,1.6,1.6])
        with r3c1:
            add_date = st.date_input("Date", value=date.today(), key=skey("add","date"))
        with r3c2:
            add_visa = st.text_input("Visa", value="", key=skey("add","visa"))
        with r3c3:
            add_montant = st.text_input("Montant honoraires (US $)", value="0", key=skey("add","montant"))

        # Row 4: autres frais / acomptes 1..4
        r4c1, r4c2 = st.columns([1.6,2.4])
        with r4c1:
            add_autres = st.text_input("Autres frais (US $)", value="0", key=skey("add","autres"))
        with r4c2:
            a1 = st.text_input("Acompte 1", value="0", key=skey("add","ac1"))
            a2 = st.text_input("Acompte 2", value="0", key=skey("add","ac2"))
            a3 = st.text_input("Acompte 3", value="0", key=skey("add","ac3"))
            a4 = st.text_input("Acompte 4", value="0", key=skey("add","ac4"))

        add_comments = st.text_area("Commentaires", value="", key=skey("add","comments"))

        submitted = st.form_submit_button("Ajouter")
        if submitted:
            try:
                new_row = {c: "" for c in df_live.columns}
                new_row["ID_Client"] = str(next_id)
                new_row["Dossier N"] = add_dossier
                new_row["Nom"] = add_nom
                new_row["Date"] = pd.to_datetime(add_date)
                new_row["Categories"] = add_cat.strip() if isinstance(add_cat, str) else add_cat
                new_row["Sous-categorie"] = add_sub.strip() if isinstance(add_sub, str) else add_sub
                new_row["Visa"] = add_visa
                new_row["Montant honoraires (US $)"] = money_to_float(add_montant)
                new_row["Autres frais (US $)"] = money_to_float(add_autres)
                # acomptes
                new_row["Acompte 1"] = money_to_float(a1)
                new_row["Acompte 2"] = money_to_float(a2)
                new_row["Acompte 3"] = money_to_float(a3)
                new_row["Acompte 4"] = money_to_float(a4)
                # payé = sum acomptes
                paid_sum = new_row["Acompte 1"] + new_row["Acompte 2"] + new_row["Acompte 3"] + new_row["Acompte 4"]
                new_row["Payé"] = paid_sum
                new_row["Solde"] = new_row["Montant honoraires (US $)"] + new_row["Autres frais (US $)"] - paid_sum
                new_row["Commentaires"] = add_comments
                flags_to_create = list(add_flags_state.keys())
                ensure_flag_columns(df_live, flags_to_create)
                for opt, val in add_flags_state.items():
                    new_row[opt] = 1 if val else 0
                # ensure acomptes exist in df_live
                for acc in ["Acompte 1","Acompte 2","Acompte 3","Acompte 4"]:
                    if acc not in df_live.columns:
                        df_live[acc] = 0.0
                df_live = df_live.append(new_row, ignore_index=True)

                # Recalculate canonical Payé and Solde for whole df and save
                df_live = recalc_payments_and_solde(df_live)
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
            with st.form(key=skey("form_edit")):
                ecol1, ecol2 = st.columns(2)
                with ecol1:
                    st.markdown(f"**ID_Client :** {row.get('ID_Client','')}")
                    e_dossier = st.text_input("Dossier N", value=str(row.get("Dossier N","")), key=skey("edit","dossier"))
                    e_nom = st.text_input("Nom", value=str(row.get("Nom","")), key=skey("edit","nom"))
                with ecol2:
                    e_date = st.date_input("Date", value=_date_for_widget(row.get("Date", date.today())), key=skey("edit","date"))
                    st.markdown(f"Category choisie: **{e_cat_sel}**")
                    init_sub = str(row.get("Sous-categorie","")).strip()
                    if init_sub == "" and edit_sub_options:
                        init_sub_index = 1
                    else:
                        try:
                            init_sub_index = ([""] + edit_sub_options).index(init_sub)
                        except Exception:
                            init_sub_index = 0
                    e_sub = st.selectbox("Sous-catégorie", options=[""] + edit_sub_options, index=init_sub_index, key=skey("edit","sub"))
                    edit_specific = get_sub_options_for(e_sub, visa_sub_options_map)
                    checkbox_options_edit = edit_specific if edit_specific else DEFAULT_FLAGS
                    ensure_flag_columns(df_live, checkbox_options_edit)
                    cols_chk = st.columns(2)
                    edit_flags_state = {}
                    for i, opt in enumerate(checkbox_options_edit):
                        col_i = cols_chk[i % 2]
                        initial_val = True if (opt in df_live.columns and _to_num(row.get(opt, 0))>0) else False
                        k = skey("edit","flag", re.sub(r"\s+","_", opt), str(idx))
                        edit_flags_state[opt] = col_i.checkbox(opt, value=initial_val, key=k)
                e_visa = st.text_input("Visa", value=str(row.get("Visa","")), key=skey("edit","visa"))
                e_montant = st.text_input("Montant honoraires (US $)", value=str(row.get("Montant honoraires (US $)",0)), key=skey("edit","montant"))
                e_autres = st.text_input("Autres frais (US $)", value=str(row.get("Autres frais (US $)",0)), key=skey("edit","autres"))
                e_ac1 = st.text_input("Acompte 1", value=str(row.get("Acompte 1",0)), key=skey("edit","ac1"))
                e_ac2 = st.text_input("Acompte 2", value=str(row.get("Acompte 2",0)), key=skey("edit","ac2"))
                e_ac3 = st.text_input("Acompte 3", value=str(row.get("Acompte 3",0)), key=skey("edit","ac3"))
                e_ac4 = st.text_input("Acompte 4", value=str(row.get("Acompte 4",0)), key=skey("edit","ac4"))
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
                        # acomptes
                        df_live.at[idx, "Acompte 1"] = money_to_float(e_ac1)
                        df_live.at[idx, "Acompte 2"] = money_to_float(e_ac2)
                        df_live.at[idx, "Acompte 3"] = money_to_float(e_ac3)
                        df_live.at[idx, "Acompte 4"] = money_to_float(e_ac4)
                        # Ensure acomptes exist if needed
                        for acc in ["Acompte 1","Acompte 2","Acompte 3","Acompte 4"]:
                            if acc not in df_live.columns:
                                df_live[acc] = 0.0
                        # Recalculate payments & solde for entire df
                        df_live = recalc_payments_and_solde(df_live)
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
                    df_live = recalc_payments_and_solde(df_live)
                    _set_df_live(df_live)
                    st.success(f"{len(idxs)} ligne(s) supprimée(s).")
                except Exception as e:
                    st.error(f"Erreur suppression: {e}")
            else:
                st.warning("Aucune sélection pour suppression.")

# ---- Export tab ----
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
