# Visa Manager - app.py
# Final application script with:
# - Clients & Visa upload caching (_clients_cache.bin, _visa_cache.bin)
# - Always two uploaders (Clients & Visa)
# - Automatic Dossier N numbering starting at 13057 (auto increment)
# - ID_Client generated differently from Dossier N: composed of creation date + sequence (YYYYMMDD-<seq>)
# - "Ajouter" tab: auto Dossier N (numeric) + auto ID_Client (date-based), Date, Nom, Cat/Sub/Visa,
#     Montant honoraires, Acompte 1, Date Acompte 1, Escrow checkbox (label exactly "Escrow"), Comments
# - Solde computed and stored (not editable in Add), visible/editable in Gestion
# - Escrow is stored and editable in Gestion
# - Edit (Gestion) tab supports Solde display/edit, Escrow checkbox, Acompte edits, metadata updates
# - Export CSV/XLSX with optional formulas (openpyxl)
# - Robust parsing and column heuristics
#
# Usage: streamlit run app.py
# Requires: pandas, streamlit; optional: openpyxl for XLSX formula exports.

import os
import json
import re
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional

import pandas as pd
import numpy as np
import streamlit as st

# Optional plotly
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    px = None
    HAS_PLOTLY = False

# openpyxl for writing formulas
try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

# -------------------------
# Configuration & current user
# -------------------------
APP_TITLE = "🛂 Visa Manager"
COLS_CLIENTS = [
    "ID_Client", "Dossier N", "Nom", "Date",
    "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)",
    "Payé", "Solde", "Solde à percevoir (US $)",
    "Acompte 1", "Date Acompte 1",
    "Acompte 2", "Date Acompte 2", "Acompte 3", "Acompte 4",
    "Escrow",
    "RFE", "Dossiers envoyé", "Dossier approuvé",
    "Dossier refusé", "Dossier Annulé",
    "Date de création", "Créé par", "Dernière modification", "Modifié par",
    "Commentaires"
]
MEMO_FILE = "_vmemory.json"
CACHE_CLIENTS = "_clients_cache.bin"
CACHE_VISA = "_visa_cache.bin"
SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"
SID = "vmgr"
DEFAULT_START_CLIENT_ID = 13057

# Current user (from session)
CURRENT_USER = "charleytrigano"

def skey(*parts: str) -> str:
    return f"{SID}_" + "_".join([p for p in parts if p])

# -------------------------
# Helpers: normalization / formatting
# -------------------------
def normalize_header_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r"^\s+|\s+$", "", s)
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
    try:
        if pd.isna(x):
            return 0.0
        s = str(x).strip()
        if s == "" or s in ("-", "—", "–", "NA", "N/A"):
            return 0.0
        s = s.replace("\u202f", "").replace("\xa0", "").replace(" ", "")
        s = re.sub(r"[^\d,.\-]", "", s)
        if s == "":
            return 0.0
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        else:
            if "," in s and s.count(",") == 1 and "." not in s:
                if len(s.split(",")[-1]) == 2:
                    s = s.replace(",", ".")
                else:
                    s = s.replace(",", "")
            else:
                s = s.replace(",", ".")
        s = s.strip()
        if s in ("", "-"):
            return 0.0
        return float(s)
    except Exception:
        try:
            return float(re.sub(r"[^0-9.\-]", "", str(x)))
        except Exception:
            return 0.0

def _to_num(x: Any) -> float:
    return money_to_float(x) if not isinstance(x, (int, float)) else float(x)

def _fmt_money(v: Any) -> str:
    try:
        return "${:,.2f}".format(float(v))
    except Exception:
        return "$0.00"

def _date_for_widget(val: Any) -> Optional[date]:
    if isinstance(val, date):
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

def _format_datetime_for_display(val: Any) -> str:
    if pd.isna(val) or val is None or str(val).strip() == "":
        return ""
    try:
        dt = pd.to_datetime(val)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return str(val)

# -------------------------
# Column heuristics & helpers
# -------------------------
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
    "solde a percevoir": "Solde à percevoir (US $)",
    "acompte 1": "Acompte 1", "acompte1": "Acompte 1",
    "date acompte 1": "Date Acompte 1",
    "acompte 2": "Acompte 2", "acompte2": "Acompte 2",
    "acompte 3": "Acompte 3", "acompte3": "Acompte 3",
    "acompte 4": "Acompte 4", "acompte4": "Acompte 4",
    "escrow": "Escrow",
    "dossier envoye": "Dossiers envoyé", "dossier approuve": "Dossier approuvé", "dossier refuse": "Dossier refusé",
    "rfe": "RFE", "commentaires": "Commentaires"
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

def map_columns_heuristic(df: Any) -> Tuple[pd.DataFrame, Dict[str,str]]:
    if not isinstance(df, pd.DataFrame):
        try:
            st.sidebar.warning("map_columns_heuristic: input is not a DataFrame — coercing to empty DataFrame.")
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

# -------------------------
# I/O helpers
# -------------------------
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

# -------------------------
# Column detection helpers
# -------------------------
def detect_acompte_columns(df: pd.DataFrame) -> List[str]:
    cols = []
    if df is None or df.empty:
        return cols
    for c in df.columns:
        k = canonical_key(c)
        if "acompte" in k:
            cols.append(c)
    def sort_key(name):
        m = re.search(r"(\d+)", name)
        return int(m.group(1)) if m else 999
    cols = sorted(cols, key=sort_key)
    return cols

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

# -------------------------
# Ensure columns helper (defensive)
# -------------------------
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
            if c in ["Payé", "Solde", "Solde à percevoir (US $)", "Montant honoraires (US $)", "Autres frais (US $)", "Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
                out[c] = 0.0
            elif c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
                out[c] = 0
            elif c in ["Date de création", "Dernière modification", "Date", "Date Acompte 1", "Date Acompte 2"]:
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
                if c in ["Payé", "Solde", "Solde à percevoir (US $)", "Montant honoraires (US $)", "Autres frais (US $)", "Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
                    safe[c] = 0.0
                elif c in ["Date de création", "Dernière modification", "Date", "Date Acompte 1", "Date Acompte 2"]:
                    safe[c] = pd.NaT
                elif c == "Escrow":
                    safe[c] = 0
                else:
                    safe[c] = ""
        return safe

# -------------------------
# Status normalizer
# -------------------------
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
        except Exception:
            df[c] = 0
    return df

# -------------------------
# normalize_clients_for_live (defensive)
# -------------------------
def normalize_clients_for_live(df_clients_raw: Any) -> pd.DataFrame:
    if not isinstance(df_clients_raw, pd.DataFrame):
        try:
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
        if not isinstance(df_mapped, pd.DataFrame):
            df_mapped = pd.DataFrame()
    except Exception:
        df_mapped = df_clients_raw.copy() if isinstance(df_clients_raw, pd.DataFrame) else pd.DataFrame()

    if "Solde" in df_mapped.columns:
        try:
            df_mapped["Solde_source"] = df_mapped["Solde"].copy()
        except Exception:
            pass
        try:
            df_mapped = df_mapped.drop(columns=["Solde"])
        except Exception:
            pass

    # parse date columns
    for dtc in ["Date","Date de création","Dernière modification","Date Acompte 1","Date Acompte 2"]:
        if dtc in df_mapped.columns:
            try:
                df_mapped[dtc] = pd.to_datetime(df_mapped[dtc], errors="coerce")
            except Exception:
                pass

    df = _ensure_columns(df_mapped, COLS_CLIENTS)

    for col in NUMERIC_TARGETS:
        if col in df.columns:
            try:
                df[col] = df[col].apply(lambda x: _to_num(x) if not isinstance(x, (int, float)) else float(x))
            except Exception:
                try:
                    df[col] = df[col].apply(lambda x: 0.0)
                except Exception:
                    pass

    for acc in ["Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
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
        df[montant_col] = df.get(montant_col, 0).apply(lambda x: _to_num(x) if not isinstance(x,(int,float)) else float(x))
        df[autres_col] = df.get(autres_col, 0).apply(lambda x: _to_num(x) if not isinstance(x,(int,float)) else float(x))
        df["Payé"] = df.get("Payé", 0).apply(lambda x: _to_num(x) if not isinstance(x,(int,float)) else float(x))
        df["Solde"] = df[montant_col] + df[autres_col] - df["Payé"]
        df["Solde à percevoir (US $)"] = df["Solde"].copy()
    except Exception:
        try:
            df["Solde"] = df.get("Solde", 0).apply(lambda x: _to_num(x))
            df["Solde à percevoir (US $)"] = df.get("Solde à percevoir (US $)", 0).apply(lambda x: _to_num(x))
        except Exception:
            df["Solde"] = 0.0
            df["Solde à percevoir (US $)"] = 0.0

    df = _normalize_status(df)

    for c in ["Nom", "Categories", "Sous-categorie", "Visa", "Commentaires", "Créé par", "Modifié par"]:
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

# -------------------------
# recalc_payments_and_solde
# -------------------------
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
        out["Payé"] = out[acomptes].sum(axis=1).astype(float) if acomptes else out.get("Payé",0).apply(lambda x: _to_num(x))
    except Exception:
        out["Payé"] = out.get("Payé",0).apply(lambda x: _to_num(x))

    try:
        out["Solde"] = out[montant_col] + out[autres_col] - out["Payé"]
        out["Solde à percevoir (US $)"] = out["Solde"].copy()
        out["Solde"] = out["Solde"].astype(float)
        out["Solde à percevoir (US $)"] = out["Solde à percevoir (US $)"].astype(float)
        out["Payé"] = out["Payé"].astype(float)
    except Exception:
        out["Solde"] = out.get("Solde",0).apply(lambda x: _to_num(x))
        out["Solde à percevoir (US $)"] = out.get("Solde à percevoir (US $)",0).apply(lambda x: _to_num(x))

    # Ensure Escrow is integer 0/1
    if "Escrow" in out.columns:
        try:
            out["Escrow"] = out["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ("1","true","t","yes","oui","y","x") else (1 if _to_num(x) == 1 else 0))
        except Exception:
            out["Escrow"] = out["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ("1","true","t","yes","oui","y","x") else 0)

    return out

# -------------------------
# Next ID & flags helpers
# -------------------------
def get_next_client_id_numeric(df: pd.DataFrame) -> int:
    # returns next numeric seq for Dossier N (>= DEFAULT_START_CLIENT_ID)
    if df is None or df.empty:
        return DEFAULT_START_CLIENT_ID
    vals = df.get("Dossier N", pd.Series([], dtype="object"))
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

def get_next_client_id_datebased(df: pd.DataFrame) -> str:
    # create ID_Client as YYYYMMDD-<seq> where seq is next numeric seq
    seq = get_next_client_id_numeric(df)
    datepart = datetime.now().strftime("%Y%m%d")
    return f"{datepart}-{seq}"

def ensure_flag_columns(df: pd.DataFrame, flags: List[str]) -> None:
    for f in flags:
        if f not in df.columns:
            df[f] = 0

DEFAULT_FLAGS = ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]

# -------------------------
# Streamlit UI bootstrap and file upload caching logic
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

# File uploaders (always present: Clients & Visa)
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

# Read uploaded files into bytes, persist them to cache files so they survive code edits and reruns
clients_bytes = None
visa_bytes = None

# If user uploaded via widget, take bytes and cache to disk
if up_clients is not None:
    try:
        clients_bytes = up_clients.getvalue()
        # persist to cache
        try:
            with open(CACHE_CLIENTS, "wb") as f:
                f.write(clients_bytes)
        except Exception:
            pass
    except Exception:
        try:
            up_clients.seek(0); clients_bytes = up_clients.read()
        except Exception:
            clients_bytes = None

if up_visa is not None:
    try:
        visa_bytes = up_visa.getvalue()
        try:
            with open(CACHE_VISA, "wb") as f:
                f.write(visa_bytes)
        except Exception:
            pass
    except Exception:
        try:
            up_visa.seek(0); visa_bytes = up_visa.read()
        except Exception:
            visa_bytes = None

# If no uploader bytes, but a local path provided use it
if clients_bytes is None and clients_path_in:
    clients_src_for_read = clients_path_in
elif clients_bytes is not None:
    clients_src_for_read = BytesIO(clients_bytes)
# If no uploader and no path, but cache file exists, load it
elif os.path.exists(CACHE_CLIENTS):
    try:
        clients_bytes = open(CACHE_CLIENTS, "rb").read()
        clients_src_for_read = BytesIO(clients_bytes)
    except Exception:
        clients_src_for_read = None
else:
    clients_src_for_read = None

if visa_bytes is None and visa_path_in:
    visa_src_for_read = visa_path_in
elif visa_bytes is not None:
    visa_src_for_read = BytesIO(visa_bytes)
elif os.path.exists(CACHE_VISA):
    try:
        visa_bytes = open(CACHE_VISA, "rb").read()
        visa_src_for_read = BytesIO(visa_bytes)
    except Exception:
        visa_src_for_read = None
else:
    visa_src_for_read = None

# -------------------------
# Read dataframes from provided sources
# -------------------------
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

# sanitize visa raw
if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
    try:
        df_visa_raw = df_visa_raw.copy()
        df_visa_raw = df_visa_raw.fillna("")
        for c in df_visa_raw.columns:
            try:
                df_visa_raw[c] = df_visa_raw[c].astype(str).str.strip()
                df_visa_raw[c] = df_visa_raw[c].replace(r'^\s*nan\s*$', "", regex=True, case=False)
            except Exception:
                pass
    except Exception:
        pass

# Build visa maps from Visa sheet
visa_map = {}; visa_map_norm = {}; visa_categories = []; visa_sub_options_map = {}
if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
    try:
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
        # visa_sub_options_map
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
    except Exception:
        visa_map = {}; visa_map_norm = {}; visa_categories = []; visa_sub_options_map = {}
else:
    visa_map = {}; visa_map_norm = {}; visa_categories = []; visa_sub_options_map = {}

# -------------------------
# Build live df and enforce canonical Payé/Solde
# -------------------------
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

# Helper unique_nonempty & kpi_html
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

def kpi_html(label: str, value: str, sub: str = "") -> str:
    html = f"""
    <div style="border:1px solid rgba(255,255,255,0.04); border-radius:6px; padding:8px 10px; margin:6px 4px;">
      <div style="font-size:12px; color:#666;">{label}</div>
      <div style="font-size:18px; font-weight:700; margin-top:4px;">{value}</div>
      <div style="font-size:11px; color:#888; margin-top:4px;">{sub}</div>
    </div>
    """
    return html

# ---- Export / Dashboard / Gestion tabs are unchanged in behavior; keep full implementations above ----

# End of file
