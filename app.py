# -*- coding: utf-8 -*-

# app.py - Visa Manager (complete)
# - Features:
#   * Import Clients/Visa (xlsx/csv), normalize columns, heuristic mapping
#   * Import single ComptaCli fiche and persist to cache so re-upload not required
#   * Session-backed clients table editable in "Gestion"
#   * Compta Client tab: select a row (index | Dossier N | Nom) like Gestion and export .xlsx
#   * Dashboard: filters by Category/Subcategory, Year, Month (with "Tous"), custom date range, and comparison between two periods
#   * Analyses: multiple charts (time series monthly, heatmap year x month, category treemap, top-N clients, comparison bars)
# Requirements: pip install streamlit pandas openpyxl plotly
# Run: streamlit run app.py

import os
import json
import re
from io import BytesIO
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# optional plotly for richer charts
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except Exception:
    PLOTLY_AVAILABLE = False

# -------------------------
# Configuration & constants
# -------------------------
APP_TITLE = "ðŸ›‚ Visa Manager"
COLS_CLIENTS = [
    "ID_Client",
    "Dossier N",
    "Nom",
    "Date",
    "Categories",
    "Sous-categorie",
    "Visa",
    "Montant honoraires (US $)",
    "Autres frais (US $)",
    "PayÃ©",
    "Solde",
    "Solde Ã  percevoir (US $)",
    "Acompte 1", "Date Acompte 1",
    "Acompte 2", "Date Acompte 2",
    "Acompte 3", "Date Acompte 3",
    "Acompte 4", "Date Acompte 4",
    "Escrow",
    "RFE",
    "Dossiers envoyÃ©",
    "Dossier approuvÃ©",
    "Dossier refusÃ©",
    "Dossier AnnulÃ©",
    "Commentaires",
    "ModeReglement",
    "ModeReglement_Ac1", "ModeReglement_Ac2", "ModeReglement_Ac3", "ModeReglement_Ac4"
]
MEMO_FILE = "_vmemory.json"
CACHE_CLIENTS = "_clients_cache.xlsx"
CACHE_VISA = "_visa_cache.xlsx"
SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"
SHEET_COMPTACLI = "ComptaCli"
SID = "vmgr"
DEFAULT_START_CLIENT_ID = 13057
DEFAULT_FLAGS = ["RFE", "Dossiers envoyÃ©", "Dossier approuvÃ©", "Dossier refusÃ©", "Dossier AnnulÃ©"]

def skey(*parts: str) -> str:
    return f"{SID}_" + "_".join([p for p in parts if p])

# -------------------------
# Basic helpers (parsing / formatting)
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
    replace_map = {"Ã©":"e","Ã¨":"e","Ãª":"e","Ã«":"e","Ã ":"a","Ã¢":"a","Ã®":"i","Ã¯":"i","Ã´":"o","Ã¶":"o","Ã¹":"u","Ã»":"u","Ã¼":"u","Ã§":"c"}
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
        if x is None:
            return 0.0
        try:
            if pd.isna(x):
                return 0.0
        except Exception:
            pass
        if isinstance(x, (int, float)):
            return float(x)
        s = str(x).strip()
        if s == "" or s.lower() in ("na","n/a","nan","none","null"):
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
        return float(s)
    except Exception:
        try:
            return float(re.sub(r"[^0-9.\-]", "", str(x)) or 0.0)
        except Exception:
            return 0.0

def _to_num(x: Any) -> float:
    if isinstance(x, (int, float)) and (not pd.isna(x)):
        return float(x)
    return money_to_float(x)

def _fmt_money(v: Any) -> str:
    try:
        return "${:,.2f}".format(float(v))
    except Exception:
        return "$0.00"

def _date_or_none_safe(v: Any) -> Optional[date]:
    try:
        if v is None:
            return None
        if isinstance(v, date) and not isinstance(v, datetime):
            return v
        if isinstance(v, datetime):
            return v.date()
        d = pd.to_datetime(v, dayfirst=True, errors="coerce")
        if pd.isna(d):
            return None
        return date(int(d.year), int(d.month), int(d.day))
    except Exception:
        return None

# -------------------------
# Column heuristics & detection
# -------------------------
COL_CANDIDATES = {
    "id client": "ID_Client", "idclient": "ID_Client",
    "dossier n": "Dossier N", "dossier": "Dossier N",
    "nom": "Nom", "date": "Date",
    "categories": "Categories", "categorie": "Categories",
    "sous categorie": "Sous-categorie", "sous-categorie": "Sous-categorie", "souscategorie": "Sous-categorie",
    "visa": "Visa",
    "montant": "Montant honoraires (US $)", "montant honoraires": "Montant honoraires (US $)",
    "autres frais": "Autres frais (US $)", "autre frais": "Autres frais (US $)",
    "payÃ©": "PayÃ©", "paye": "PayÃ©",
    "solde": "Solde",
    "mode reglement": "ModeReglement",
    "rfe": "RFE"
}
# extra variants
COL_CANDIDATES.update({
    "montant honoraires us": "Montant honoraires (US $)",
    "montant honoraires (us $)": "Montant honoraires (US $)",
    "montant total": "Montant honoraires (US $)",
    "autre frais": "Autres frais (US $)",
    "autres frais": "Autres frais (US $)",
    "mode reglement ac1": "ModeReglement_Ac1",
    "modereglement ac1": "ModeReglement_Ac1",
    "modereglement_ac1": "ModeReglement_Ac1",
    "mode reglement ac2": "ModeReglement_Ac2",
    "mode reglement ac3": "ModeReglement_Ac3",
    "mode reglement ac4": "ModeReglement_Ac4",
    "modereglement_ac2": "ModeReglement_Ac2",
    "modereglement_ac3": "ModeReglement_Ac3",
    "modereglement_ac4": "ModeReglement_Ac4",
    "escrow": "Escrow"
})

NUMERIC_TARGETS = [
    "Montant honoraires (US $)",
    "Autres frais (US $)",
    "PayÃ©",
    "Solde",
    "Solde Ã  percevoir (US $)",
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
    candidates = ["Montant honoraires (US $)", "Montant honoraires", "Montant", "Montant Total"]
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
    candidates = ["Autres frais (US $)", "Autres frais", "Autres", "Autre frais"]
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
# Visa maps init
# -------------------------
visa_map: Dict[str, List[str]] = {}
visa_map_norm: Dict[str, List[str]] = {}
visa_sub_options_map: Dict[str, List[str]] = {}
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

def parse_modes_global(raw: Any) -> List[str]:
    try:
        if pd.isna(raw) or raw is None:
            return []
        s = str(raw).strip()
        if not s:
            return []
        return [p.strip() for p in s.split(",") if p.strip()]
    except Exception:
        return []

# -------------------------
# Finance helpers
# -------------------------
def get_next_dossier_numeric(df: pd.DataFrame) -> int:
    try:
        if df is None or df.empty:
            return DEFAULT_START_CLIENT_ID
        vals = df.get("Dossier N", pd.Series([], dtype="object"))
        nums: List[int] = []
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
    except Exception:
        return DEFAULT_START_CLIENT_ID

def make_id_client_datebased(df: pd.DataFrame) -> str:
    seq = get_next_dossier_numeric(df)
    datepart = datetime.now().strftime("%Y%m%d")
    return f"{datepart}-{seq}"

def ensure_flag_columns(df_like: Any, flags: List[str]) -> None:
    if isinstance(df_like, pd.DataFrame):
        for f in flags:
            if f not in df_like.columns:
                df_like[f] = 0
    elif isinstance(df_like, dict):
        for f in flags:
            if f not in df_like:
                df_like[f] = 0

# -------------------------
# Normalize & recalc dataset
# -------------------------
def _ensure_columns(df: Any, cols: List[str]) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame):
        df = pd.DataFrame()
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            if c in NUMERIC_TARGETS:
                out[c] = 0.0
            elif c == "Escrow":
                out[c] = 0
            elif c in ["RFE", "Dossiers envoyÃ©", "Dossier approuvÃ©", "Dossier refusÃ©", "Dossier AnnulÃ©"]:
                out[c] = 0
            elif "Date" in c:
                out[c] = pd.NaT
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
                if c in NUMERIC_TARGETS:
                    safe[c] = 0.0
                elif c == "Escrow":
                    safe[c] = 0
                elif c in ["RFE", "Dossiers envoyÃ©", "Dossier approuvÃ©", "Dossier refusÃ©", "Dossier AnnulÃ©"]:
                    safe[c] = 0
                elif "Date" in c:
                    safe[c] = pd.NaT
                else:
                    safe[c] = ""
        return safe

def normalize_clients_for_live(raw: Any) -> pd.DataFrame:
    df_raw = raw
    if not isinstance(df_raw, pd.DataFrame):
        maybe = read_any_table(df_raw, sheet=None, debug_prefix="[normalize] ")
        df_raw = maybe if isinstance(maybe, pd.DataFrame) else pd.DataFrame()
    df_mapped, _ = map_columns_heuristic(df_raw)
    for dtc in [c for c in df_mapped.columns if "Date" in c]:
        try:
            df_mapped[dtc] = pd.to_datetime(df_mapped[dtc], dayfirst=True, errors="coerce")
        except Exception:
            pass
    df = _ensure_columns(df_mapped, COLS_CLIENTS)
    # numeric coercion
    for c in NUMERIC_TARGETS:
        if c in df.columns:
            try:
                df[c] = df[c].apply(lambda x: _to_num(x))
            except Exception:
                df[c] = 0.0
    # ensure acomptes exist
    for acc in ["Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
        if acc not in df.columns:
            df[acc] = 0.0
    acomptes_cols = detect_acompte_columns(df)
    if acomptes_cols:
        try:
            df["PayÃ©"] = df[acomptes_cols].fillna(0).apply(lambda row: sum([_to_num(row[c]) for c in acomptes_cols]), axis=1)
        except Exception:
            df["PayÃ©"] = df.get("PayÃ©", 0).apply(lambda x: _to_num(x))
    else:
        df["PayÃ©"] = df.get("PayÃ©", 0).apply(lambda x: _to_num(x))
    try:
        montant_col = detect_montant_column(df) or "Montant honoraires (US $)"
        autres_col = detect_autres_column(df) or "Autres frais (US $)"
        df[montant_col] = df.get(montant_col, 0).apply(lambda x: _to_num(x))
        df[autres_col] = df.get(autres_col, 0).apply(lambda x: _to_num(x))
        df["Solde"] = df[montant_col] + df[autres_col] - df["PayÃ©"]
        df["Solde Ã  percevoir (US $)"] = df["Solde"].copy()
    except Exception:
        df["Solde"] = df.get("Solde", 0).apply(lambda x: _to_num(x))
        df["Solde Ã  percevoir (US $)"] = df.get("Solde Ã  percevoir (US $)", 0).apply(lambda x: _to_num(x))
    for c in ["Nom","Categories","Sous-categorie","Visa","Commentaires","ModeReglement","ModeReglement_Ac1","ModeReglement_Ac2","ModeReglement_Ac3","ModeReglement_Ac4"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str)
    if "Escrow" not in df.columns:
        df["Escrow"] = 0
    return df

def recalc_payments_and_solde(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    acomptes = detect_acompte_columns(out)
    if not acomptes:
        for acc in ["Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"]:
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
        out["PayÃ©"] = out[acomptes].sum(axis=1).astype(float) if acomptes else out.get("PayÃ©",0).apply(lambda x: _to_num(x))
    except Exception:
        out["PayÃ©"] = out.get("PayÃ©",0).apply(lambda x: _to_num(x))
    try:
        out["Solde"] = out[montant_col] + out[autres_col] - out["PayÃ©"]
        out["Solde Ã  percevoir (US $)"] = out["Solde"].copy()
        out["Solde"] = out["Solde"].astype(float)
    except Exception:
        out["Solde"] = out.get("Solde",0).apply(lambda x: _to_num(x))
    if "Escrow" in out.columns:
        try:
            out["Escrow"] = out["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ("1","true","t","yes","oui","y","x") else (1 if _to_num(x) == 1 else 0))
        except Exception:
            out["Escrow"] = out["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ("1","true","t","yes","oui","y","x") else 0)
    return out

# -------------------------
# I/O helpers (Excel/CSV)
# -------------------------
def try_read_excel_from_bytes(b: bytes, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
    bio = BytesIO(b)
    try:
        xls = pd.ExcelFile(bio, engine="openpyxl")
        sheets = xls.sheet_names
        if sheet_name and sheet_name in sheets:
            return pd.read_excel(BytesIO(b), sheet_name=sheet_name, engine="openpyxl")
        for cand in [SHEET_CLIENTS, SHEET_VISA, SHEET_COMPTACLI, "Sheet1"]:
            if cand in sheets:
                try:
                    return pd.read_excel(BytesIO(b), sheet_name=cand, engine="openpyxl")
                except Exception:
                    continue
        return pd.read_excel(BytesIO(b), sheet_name=0, engine="openpyxl")
    except Exception:
        return None

def _normalize_incoming_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    new_cols = []
    for c in df.columns:
        s = str(c)
        s = re.sub(r"_\s+", "_", s)
        s = re.sub(r"\s+_", "_", s)
        s = re.sub(r"\s+", " ", s).strip()
        s = s.replace("\u00A0", " ").strip()
        new_cols.append(s)
    try:
        df = df.rename(columns=dict(zip(df.columns, new_cols)))
    except Exception:
        pass
    return df

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
                df = _normalize_incoming_columns(df)
                return df
            for sep in [";", ","]:
                for enc in ["utf-8", "latin-1", "cp1252"]:
                    try:
                        df = pd.read_csv(BytesIO(src), sep=sep, encoding=enc, on_bad_lines="skip")
                        df = _normalize_incoming_columns(df)
                        return df
                    except Exception:
                        continue
            return None
        if isinstance(src, BytesIO):
            b = src.getvalue()
            df = try_read_excel_from_bytes(b, sheet)
            if df is not None:
                df = _normalize_incoming_columns(df)
                return df
            for sep in [";", ","]:
                for enc in ["utf-8", "latin-1", "cp1252"]:
                    try:
                        df = pd.read_csv(BytesIO(b), sep=sep, encoding=enc, on_bad_lines="skip")
                        df = _normalize_incoming_columns(df)
                        return df
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
                    df = _normalize_incoming_columns(df)
                    return df
                for sep in [";", ","]:
                    for enc in ["utf-8", "latin-1", "cp1252"]:
                        try:
                            df = pd.read_csv(BytesIO(data), sep=sep, encoding=enc, on_bad_lines="skip")
                            df = _normalize_incoming_columns(df)
                            return df
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
                            df = pd.read_csv(p, sep=sep, encoding=enc, on_bad_lines="skip")
                            df = _normalize_incoming_columns(df)
                            return df
                        except Exception:
                            continue
                return None
            else:
                try:
                    df = pd.read_excel(p, sheet_name=sheet or 0, engine="openpyxl")
                    df = _normalize_incoming_columns(df)
                    return df
                except Exception:
                    return None
    except Exception as e:
        _log(f"read_any_table exception: {e}")
        return None
    _log("read_any_table: unsupported src type")
    return None

# -------------------------
# Parse ComptaCli sheet (single fiche)
# -------------------------
def parse_fiche_from_sheet(df_sheet: pd.DataFrame) -> Optional[pd.DataFrame]:
    try:
        if df_sheet is None or df_sheet.empty:
            return None
        df_sheet = _normalize_incoming_columns(df_sheet)
        df2 = df_sheet.fillna("").astype(str)
        lines = []
        for _, row in df2.iterrows():
            lines.append("; ".join([str(c).strip() for c in row.tolist() if str(c).strip() != ""]))
        out = {c: "" for c in COLS_CLIENTS}
        for n in ["Montant honoraires (US $)","Autres frais (US $)","Acompte 1","Acompte 2","Acompte 3","Acompte 4"]:
            out[n] = 0.0
        out["Escrow"] = 0
        for line in lines:
            l = line.lower()
            if "id_client" in l or "id client" in l or "iid_client" in l:
                m = re.search(r"([A-Za-z0-9\-\_]+)", line)
                if m:
                    out["ID_Client"] = m.group(1)
            if "dossier n" in l or "dossier" in l:
                parts = [p.strip() for p in re.split(r"[;:]", line) if p.strip()]
                for p in parts:
                    if re.search(r"\d", p):
                        out["Dossier N"] = p
                        break
            if "nom" in l and out.get("Nom","") == "":
                parts = [p.strip() for p in re.split(r"[;:]", line) if p.strip()]
                for p in parts:
                    if p.lower() not in ("nom","name"):
                        out["Nom"] = p
                        break
            if "categorie" in l or "sous" in l or "visa" in l:
                parts = [p.strip() for p in re.split(r"[;:]", line) if p.strip()]
                for p in parts:
                    pl = p.lower()
                    if any(k in pl for k in ("categorie","sous","visa")):
                        continue
                    if out["Categories"] == "":
                        out["Categories"] = p
                    elif out["Sous-categorie"] == "":
                        out["Sous-categorie"] = p
                    elif out["Visa"] == "":
                        out["Visa"] = p
            if "montant honoraires" in l or "montant total" in l:
                for token in re.split(r"[;:]", line):
                    v = money_to_float(token)
                    if v:
                        out["Montant honoraires (US $)"] = v
                        break
            if "autre" in l and "frais" in l:
                for token in re.split(r"[;:]", line):
                    v = money_to_float(token)
                    if v:
                        out["Autres frais (US $)"] = v
                        break
            if "date acompte 1" in l or "acompte 1" in l:
                for token in re.split(r"[;:]", line):
                    d = _date_or_none_safe(token)
                    if d:
                        out["Date Acompte 1"] = pd.to_datetime(d)
                    v = money_to_float(token)
                    if v and v != 0:
                        out["Acompte 1"] = v
            if "date acompte 2" in l or "acompte 2" in l:
                for token in re.split(r"[;:]", line):
                    d = _date_or_none_safe(token)
                    if d:
                        out["Date Acompte 2"] = pd.to_datetime(d)
                    v = money_to_float(token)
                    if v and v != 0:
                        out["Acompte 2"] = v
            if "date acompte 3" in l or "acompte 3" in l:
                for token in re.split(r"[;:]", line):
                    d = _date_or_none_safe(token)
                    if d:
                        out["Date Acompte 3"] = pd.to_datetime(d)
                    v = money_to_float(token)
                    if v and v != 0:
                        out["Acompte 3"] = v
            if "date acompte 4" in l or "acompte 4" in l:
                for token in re.split(r"[;:]", line):
                    d = _date_or_none_safe(token)
                    if d:
                        out["Date Acompte 4"] = pd.to_datetime(d)
                    v = money_to_float(token)
                    if v and v != 0:
                        out["Acompte 4"] = v
            if "escrow" in l:
                out["Escrow"] = 1 if ("1" in l or "oui" in l or "yes" in l or "true" in l) else 0
            if "comment" in l and out.get("Commentaires","") == "":
                m = re.split(r"commentaires?:", line, flags=re.I)
                if len(m) > 1:
                    out["Commentaires"] = m[1].strip()
        paye = sum([out.get("Acompte 1",0.0), out.get("Acompte 2",0.0), out.get("Acompte 3",0.0), out.get("Acompte 4",0.0)])
        out["Paye"] = paye
        out["Solde"] = out.get("Montant honoraires (US $)",0.0) + out.get("Autres frais (US $)",0.0) - paye
        out["Solde Ã  percevoir (US $)"] = out["Solde"]
        df_out = pd.DataFrame([out])
        for c in df_out.columns:
            if "Date" in c:
                try:
                    df_out[c] = pd.to_datetime(df_out[c], errors="coerce")
                except Exception:
                    pass
        return df_out
    except Exception:
        return None

# -------------------------
# Session-safe DataFrame
# -------------------------
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

def _persist_clients_cache(df: pd.DataFrame) -> None:
    try:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Clients")
        with open(CACHE_CLIENTS, "wb") as f:
            f.write(buf.getvalue())
    except Exception:
        pass

# -------------------------
# UI bootstrap (sidebar)
# -------------------------
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.sidebar.header("ðŸ“‚ Fichiers")
last_clients_path = ""
last_visa_path = ""
try:
    if os.path.exists(MEMO_FILE):
        with open(MEMO_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
            last_clients_path = d.get("clients", "")
            last_visa_path = d.get("visa", "")
except Exception:
    pass

up_clients = st.sidebar.file_uploader("Clients (xlsx/xls/xlsm/csv) â€” inclure feuille ComptaCli pour fiche", type=["xlsx","xls","xlsm","csv"], key=skey("up_clients"))
up_visa = st.sidebar.file_uploader("Visa (xlsx/xls/xlsm/csv)", type=["xlsx","xls","xlsm","csv"], key=skey("up_visa"))
clients_path_in = st.sidebar.text_input("ou chemin local Clients (optionnel)", value=last_clients_path or "", key=skey("cli_path"))
visa_path_in = st.sidebar.text_input("ou chemin local Visa (optionnel)", value=last_visa_path or "", key=skey("vis_path"))

if st.sidebar.button("ðŸ“¥ Sauvegarder chemins", key=skey("btn_save_paths")):
    try:
        with open(MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump({"clients": clients_path_in or "", "visa": visa_path_in or ""}, f, ensure_ascii=False, indent=2)
        st.sidebar.success("Chemins sauvegardÃ©s.")
    except Exception:
        st.sidebar.error("Impossible de sauvegarder les chemins.")

# Save uploaded bytes and try to detect ComptaCli sheet
clients_src_for_read = None
visa_src_for_read = None
uploaded_comptacli_df = None

# If cached clients exist, use it by default (persisted from previous imports/edits)
if os.path.exists(CACHE_CLIENTS) and up_clients is None and not clients_path_in:
    clients_src_for_read = CACHE_CLIENTS

if up_clients is not None:
    try:
        clients_bytes = up_clients.getvalue()
        try:
            xls_all = pd.read_excel(BytesIO(clients_bytes), sheet_name=None, engine="openpyxl")
        except Exception:
            xls_all = None
        if isinstance(xls_all, dict):
            for name, df_sheet in xls_all.items():
                df_sheet = _normalize_incoming_columns(df_sheet)
                if "comptacli" in canonical_key(name):
                    parsed = parse_fiche_from_sheet(df_sheet)
                    if parsed is not None:
                        uploaded_comptacli_df = parsed
            # persist clients/visa sheets to cache files
            for name, df_sheet in xls_all.items():
                df_sheet = _normalize_incoming_columns(df_sheet)
                if canonical_key(name) in (canonical_key(SHEET_CLIENTS), canonical_key("clients")) and isinstance(df_sheet, pd.DataFrame):
                    buf = BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        df_sheet.to_excel(writer, index=False, sheet_name="Clients")
                    with open(CACHE_CLIENTS, "wb") as f:
                        f.write(buf.getvalue())
                    clients_src_for_read = CACHE_CLIENTS
                elif canonical_key(name) in (canonical_key(SHEET_VISA), canonical_key("visa")) and isinstance(df_sheet, pd.DataFrame):
                    buf = BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        df_sheet.to_excel(writer, index=False, sheet_name="Visa")
                    with open(CACHE_VISA, "wb") as f:
                        f.write(buf.getvalue())
                    visa_src_for_read = CACHE_VISA
            if clients_src_for_read is None:
                with open(CACHE_CLIENTS, "wb") as f:
                    f.write(clients_bytes)
                clients_src_for_read = CACHE_CLIENTS
        else:
            with open(CACHE_CLIENTS, "wb") as f:
                f.write(clients_bytes)
            clients_src_for_read = CACHE_CLIENTS
    except Exception:
        clients_src_for_read = None
elif clients_path_in:
    clients_src_for_read = clients_path_in
elif os.path.exists(CACHE_CLIENTS) and clients_src_for_read is None:
    clients_src_for_read = CACHE_CLIENTS

if up_visa is not None:
    try:
        visa_bytes = up_visa.getvalue()
        with open(CACHE_VISA, "wb") as f:
            f.write(visa_bytes)
        visa_src_for_read = CACHE_VISA
    except Exception:
        visa_src_for_read = None
elif visa_path_in:
    visa_src_for_read = visa_path_in
elif os.path.exists(CACHE_VISA):
    visa_src_for_read = CACHE_VISA

# -------------------------
# Read raw Clients and Visa tables (if provided)
# -------------------------
df_clients_raw: Optional[pd.DataFrame] = None
df_visa_raw: Optional[pd.DataFrame] = None

try:
    if clients_src_for_read is not None:
        maybe = read_any_table(clients_src_for_read, sheet=SHEET_CLIENTS, debug_prefix="[Clients] ")
        if maybe is None:
            maybe = read_any_table(clients_src_for_read, sheet=None, debug_prefix="[Clients fallback] ")
        if isinstance(maybe, pd.DataFrame):
            df_clients_raw = _normalize_incoming_columns(maybe)
except Exception:
    df_clients_raw = None

try:
    if visa_src_for_read is not None:
        maybe = read_any_table(visa_src_for_read, sheet=SHEET_VISA, debug_prefix="[Visa] ")
        if maybe is None:
            maybe = read_any_table(visa_src_for_read, sheet=None, debug_prefix="[Visa fallback] ")
        if isinstance(maybe, pd.DataFrame):
            df_visa_raw = _normalize_incoming_columns(maybe)
except Exception:
    df_visa_raw = None

# Build visa maps if visa sheet present
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
    df_visa_mapped, _ = map_columns_heuristic(df_visa_raw)
    try:
        df_visa_mapped = coerce_category_columns(df_visa_mapped)
    except Exception:
        pass
    raw_vm = {}
    try:
        for _, r in df_visa_mapped.iterrows():
            cat = str(r.get("Categories","")).strip()
            sub = str(r.get("Sous-categorie","")).strip()
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
        for _, r in df_visa_mapped.iterrows():
            sub = str(r.get("Sous-categorie","")).strip()
            if not sub:
                continue
            key = canonical_key(sub)
            for col in cols_to_check:
                val = r.get(col,"")
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

globals().update({
    "visa_map": visa_map if 'visa_map' in locals() else {},
    "visa_map_norm": visa_map_norm if 'visa_map_norm' in locals() else {},
    "visa_categories": visa_categories if 'visa_categories' in locals() else [],
    "visa_sub_options_map": visa_sub_options_map if 'visa_sub_options_map' in locals() else {}
})

# -------------------------
# Initialize live df in session state
# -------------------------
df_all = normalize_clients_for_live(df_clients_raw if df_clients_raw is not None else None)
df_all = recalc_payments_and_solde(df_all)
if isinstance(df_all, pd.DataFrame) and not df_all.empty:
    st.session_state[DF_LIVE_KEY] = df_all.copy()
else:
    if DF_LIVE_KEY not in st.session_state or st.session_state[DF_LIVE_KEY] is None:
        st.session_state[DF_LIVE_KEY] = pd.DataFrame(columns=COLS_CLIENTS)

# -------------------------
# UI helpers
# -------------------------
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
# Tabs UI
# -------------------------
tabs = st.tabs(["ðŸ“„ Fichiers","ðŸ“Š Dashboard","ðŸ“ˆ Analyses","âž• Ajouter","âœï¸ / ðŸ—‘ï¸ Gestion","ðŸ’³ Compta Client","ðŸ’¾ Export"])

# ---- Files tab ----
with tabs[0]:
    st.header("ðŸ“‚ Fichiers")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Clients")
        if up_clients is not None:
            st.text(f"Upload: {getattr(up_clients,'name','')}")
        elif isinstance(clients_src_for_read, str) and clients_src_for_read:
            st.text(f"ChargÃ© depuis: {clients_src_for_read}")
        elif os.path.exists(CACHE_CLIENTS):
            st.text("ChargÃ© depuis le cache local")
        if df_clients_raw is None or (isinstance(df_clients_raw, pd.DataFrame) and df_clients_raw.empty):
            st.warning("Aucun fichier Clients detectÃ©.")
        else:
            st.success(f"Clients lus: {df_clients_raw.shape[0]} lignes")
            try:
                st.dataframe(df_clients_raw.head(100).reset_index(drop=True), use_container_width=True, height=360)
            except Exception:
                st.write(df_clients_raw.head(8))
    with c2:
        st.subheader("Visa")
        if up_visa is not None:
            st.text(f"Upload: {getattr(up_visa,'name','')}")
        elif isinstance(visa_src_for_read, str) and visa_src_for_read:
            st.text(f"ChargÃ© depuis: {visa_src_for_read}")
        elif os.path.exists(CACHE_VISA):
            st.text("ChargÃ© depuis le cache local")
        if df_visa_raw is None or (isinstance(df_visa_raw, pd.DataFrame) and df_visa_raw.empty):
            st.warning("Aucun fichier Visa detectÃ©.")
        else:
            st.success(f"Visa lu: {df_visa_raw.shape[0]} lignes")
            try:
                st.dataframe(df_visa_raw.head(100).reset_index(drop=True), use_container_width=True, height=360)
            except Exception:
                st.write(df_visa_raw.head(8))
    st.markdown("---")
    col_a, col_b = st.columns([1,1])
    with col_a:
        if st.button("RÃ©initialiser mÃ©moire (recharger)"):
            df_all2 = normalize_clients_for_live(df_clients_raw)
            df_all2 = recalc_payments_and_solde(df_all2)
            _set_df_live(df_all2)
            try:
                _persist_clients_cache(df_all2)
            except Exception:
                pass
            st.success("MÃ©moire rÃ©initialisÃ©e.")
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
    # If a ComptaCli was parsed on upload, offer import (and persist)
    if uploaded_comptacli_df is not None:
        st.markdown("---")
        st.info("Fiche ComptaCli dÃ©tectÃ©e dans l'xlsx uploadÃ©.")
        if st.button("Importer la fiche ComptaCli dÃ©tectÃ©e"):
            try:
                df_live = _get_df_live_safe()
                df_new = normalize_clients_for_live(uploaded_comptacli_df)
                df_new = recalc_payments_and_solde(df_new)
                df_live = pd.concat([df_live, df_new], ignore_index=True)
                df_live = recalc_payments_and_solde(df_live)
                _set_df_live(df_live)
                _persist_clients_cache(df_live)
                st.success("Fiche ComptaCli importÃ©e en mÃ©moire et persistÃ©e.")
                st.experimental_rerun()
            except Exception as e:
                st.error(f"Erreur import fiche: {e}")

# ---- Dashboard tab ----
with tabs[1]:
    st.subheader("ðŸ“Š Dashboard")
    df_live_view = recalc_payments_and_solde(_get_df_live_safe())
    if df_live_view is None or df_live_view.empty:
        st.info("Aucune donnÃ©e en mÃ©moire.")
    else:
        # derive years and months from Date column
        date_col_candidates = [c for c in df_live_view.columns if "date" in canonical_key(c)]
        date_col = "Date" if "Date" in df_live_view.columns else (date_col_candidates[0] if date_col_candidates else None)
        if date_col is None:
            st.warning("Aucune colonne Date dÃ©tectÃ©e â€” les filtres par annÃ©e/mois et graphiques temporels sont dÃ©sactivÃ©s.")
            years = []
        else:
            df_live_view[date_col] = pd.to_datetime(df_live_view[date_col], errors="coerce")
            df_live_view["_year_"] = df_live_view[date_col].dt.year
            df_live_view["_month_"] = df_live_view[date_col].dt.month
            years = sorted([int(y) for y in df_live_view["_year_"].dropna().unique().tolist()])

        # Filters: Category, Subcategory, Year, Month or custom range
        fcol1, fcol2, fcol3 = st.columns([1.2,1.2,1.6])
        with fcol1:
            cats = [""] + (unique_nonempty(df_live_view["Categories"]) if "Categories" in df_live_view.columns else [])
            sel_cat = st.selectbox("CatÃ©gorie", options=cats, index=0, key=skey("dash","cat"))
        with fcol2:
            subs = [""] + (unique_nonempty(df_live_view["Sous-categorie"]) if "Sous-categorie" in df_live_view.columns else [])
            sel_sub = st.selectbox("Sous-catÃ©gorie", options=subs, index=0, key=skey("dash","sub"))
        with fcol3:
            st.markdown("Filtrage temporel")
            timeframe_mode = st.selectbox("Mode temporel", options=["AnnÃ©e+Mois (rapide)","Plage libre (from/to)","Comparer deux pÃ©riodes"], index=0, key=skey("dash","tmode"))

        # helper to filter by year+month or range
        def apply_time_filter(df, mode="AnnÃ©e+Mois (rapide)", year=None, month=None, start=None, end=None):
            out = df.copy()
            if mode == "AnnÃ©e+Mois (rapide)":
                if year:
                    out = out[out["_year_"] == int(year)]
                    if month and month != "Tous":
                        out = out[out["_month_"] == int(month)]
            elif mode == "Plage libre (from/to)":
                if start:
                    out = out[out[date_col] >= pd.to_datetime(start)]
                if end:
                    out = out[out[date_col] <= pd.to_datetime(end)]
            return out

        # UI for each mode
        if timeframe_mode == "AnnÃ©e+Mois (rapide)":
            c1, c2 = st.columns([1,1])
            with c1:
                yrs = [""] + [str(y) for y in years]
                sel_year = st.selectbox("AnnÃ©e", options=yrs, index=0, key=skey("dash","year"))
            with c2:
                months = ["Tous"] + [str(i) for i in range(1,13)]
                sel_month = st.selectbox("Mois", options=months, index=0, key=skey("dash","month"))
            # filter view
            view = df_live_view.copy()
            if sel_cat:
                view = view[view["Categories"].astype(str) == sel_cat]
            if sel_sub:
                view = view[view["Sous-categorie"].astype(str) == sel_sub]
            if sel_year:
                view = apply_time_filter(view, mode="AnnÃ©e+Mois (rapide)", year=int(sel_year), month=(None if sel_month=="Tous" else int(sel_month)))
        elif timeframe_mode == "Plage libre (from/to)":
            c1, c2 = st.columns([1,1])
            with c1:
                start_date = st.date_input("Date dÃ©but", value=None, key=skey("dash","start"))
            with c2:
                end_date = st.date_input("Date fin", value=None, key=skey("dash","end"))
            view = df_live_view.copy()
            if sel_cat:
                view = view[view["Categories"].astype(str) == sel_cat]
            if sel_sub:
                view = view[view["Sous-categorie"].astype(str) == sel_sub]
            view = apply_time_filter(view, mode="Plage libre (from/to)", start=start_date, end=end_date)
        else:  # Compare two periods
            st.markdown("PÃ©riode A")
            a_col1, a_col2 = st.columns([1,1])
            with a_col1:
                yrs = [""] + [str(y) for y in years]
                a_year = st.selectbox("AnnÃ©e A", options=yrs, index=0, key=skey("dash","ayear"))
            with a_col2:
                a_months = ["Tous"] + [str(i) for i in range(1,13)]
                a_month = st.selectbox("Mois A", options=a_months, index=0, key=skey("dash","amonth"))
            st.markdown("PÃ©riode B")
            b_col1, b_col2 = st.columns([1,1])
            with b_col1:
                b_year = st.selectbox("AnnÃ©e B", options=yrs, index=0, key=skey("dash","byear"))
            with b_col2:
                b_month = st.selectbox("Mois B", options=a_months, index=0, key=skey("dash","bmonth"))
            # build views for A and B
            base = df_live_view.copy()
            if sel_cat:
                base = base[base["Categories"].astype(str) == sel_cat]
            if sel_sub:
                base = base[base["Sous-categorie"].astype(str) == sel_sub]
            viewA = base.copy()
            viewB = base.copy()
            if a_year:
                viewA = apply_time_filter(viewA, mode="AnnÃ©e+Mois (rapide)", year=int(a_year), month=(None if a_month=="Tous" else int(a_month)))
            if b_year:
                viewB = apply_time_filter(viewB, mode="AnnÃ©e+Mois (rapide)", year=int(b_year), month=(None if b_month=="Tous" else int(b_month)))
            # compute KPIs side-by-side
            def compute_kpis(df):
                montant_col = detect_montant_column(df) or "Montant honoraires (US $)"
                autres_col = detect_autres_column(df) or "Autres frais (US $)"
                total_honoraires = float(df.get(montant_col,0).apply(lambda x: _to_num(x)).sum())
                total_autres = float(df.get(autres_col,0).apply(lambda x: _to_num(x)).sum())
                acomptes_sum = 0.0
                for ac in detect_acompte_columns(df):
                    acomptes_sum += float(df.get(ac,0).apply(lambda x: _to_num(x)).sum())
                count = len(df)
                solde = total_honoraires + total_autres - acomptes_sum
                return {"count":count,"hon":total_honoraires,"autres":total_autres,"acomptes":acomptes_sum,"solde":solde}
            kpiA = compute_kpis(viewA)
            kpiB = compute_kpis(viewB)
            st.markdown("### Comparaison PÃ©riode A vs B")
            cA, cB = st.columns(2)
            with cA:
                st.markdown(f"**PÃ©riode A** â€” {a_year or 'â€”'} / {a_month or 'Tous'}")
                st.markdown(f"- Dossiers: {kpiA['count']}")
                st.markdown(f"- Honoraires: {_fmt_money(kpiA['hon'])}")
                st.markdown(f"- Acomptes: {_fmt_money(kpiA['acomptes'])}")
                st.markdown(f"- Solde: {_fmt_money(kpiA['solde'])}")
            with cB:
                st.markdown(f"**PÃ©riode B** â€” {b_year or 'â€”'} / {b_month or 'Tous'}")
                st.markdown(f"- Dossiers: {kpiB['count']}")
                st.markdown(f"- Honoraires: {_fmt_money(kpiB['hon'])}")
                st.markdown(f"- Acomptes: {_fmt_money(kpiB['acomptes'])}")
                st.markdown(f"- Solde: {_fmt_money(kpiB['solde'])}")
            # small bar chart comparison
            comp_df = pd.DataFrame([
                {"metric":"Dossiers","A":kpiA["count"],"B":kpiB["count"]},
                {"metric":"Honoraires","A":kpiA["hon"],"B":kpiB["hon"]},
                {"metric":"Acomptes","A":kpiA["acomptes"],"B":kpiB["acomptes"]},
                {"metric":"Solde","A":kpiA["solde"],"B":kpiB["solde"]}
            ])
            if PLOTLY_AVAILABLE:
                fig = go.Figure(data=[
                    go.Bar(name='PÃ©riode A', x=comp_df['metric'], y=comp_df['A']),
                    go.Bar(name='PÃ©riode B', x=comp_df['metric'], y=comp_df['B'])
                ])
                fig.update_layout(barmode='group', height=360, title="Comparaison mÃ©triques")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.write(comp_df)
            # show small tables of raw viewA/viewB if requested
            if st.checkbox("Voir listes PÃ©riode A / B", value=False):
                st.markdown("PÃ©riode A (extrait)")
                st.dataframe(viewA.reset_index(drop=True), use_container_width=True, height=200)
                st.markdown("PÃ©riode B (extrait)")
                st.dataframe(viewB.reset_index(drop=True), use_container_width=True, height=200)
            # skip default charts for compare mode
          

        # Default (non-compare) display of KPIs and small charts
        view = recalc_payments_and_solde(view)
        montant_col = detect_montant_column(view) or "Montant honoraires (US $)"
        autres_col = detect_autres_column(view) or "Autres frais (US $)"
        acomptes_cols = detect_acompte_columns(view)
        total_honoraires = float(view.get(montant_col,0).apply(lambda x: _to_num(x)).sum())
        total_autres = float(view.get(autres_col,0).apply(lambda x: _to_num(x)).sum())
        total_acomptes = 0.0
        for ac in acomptes_cols:
            total_acomptes += float(view.get(ac,0).apply(lambda x: _to_num(x)).sum())
        cols_k = st.columns(3)
        cols_k[0].markdown(kpi_html("Dossiers", f"{len(view):,}"), unsafe_allow_html=True)
        cols_k[1].markdown(kpi_html("Montant honoraires", _fmt_money(total_honoraires)), unsafe_allow_html=True)
        cols_k[2].markdown(kpi_html("Solde total", _fmt_money(total_honoraires + total_autres - total_acomptes)), unsafe_allow_html=True)

        st.markdown("### AperÃ§u clients (filtrÃ©)")
        try:
            display_df = view.copy()
            for mc in [montant_col, autres_col, "PayÃ©", "Solde"]:
                if mc in display_df.columns:
                    display_df[mc] = display_df[mc].apply(lambda x: _fmt_money(_to_num(x)))
            st.dataframe(display_df.reset_index(drop=True), use_container_width=True, height=360)
        except Exception:
            st.write("Impossible d'afficher le tableau.")

# ---- Analyses tab ----
with tabs[2]:
    st.subheader("ðŸ“ˆ Analyses")
    df_ = recalc_payments_and_solde(_get_df_live_safe())
    if df_ is None or df_.empty:
        st.info("Aucune donnÃ©e pour analyser.")
    else:
        # ensure date
        date_col_candidates = [c for c in df_.columns if "date" in canonical_key(c)]
        date_col = "Date" if "Date" in df_.columns else (date_col_candidates[0] if date_col_candidates else None)
        if date_col is None:
            st.warning("Aucune colonne 'Date' dÃ©tectÃ©e - analyses temporelles dÃ©sactivÃ©es.")
        else:
            df_[date_col] = pd.to_datetime(df_[date_col], errors="coerce")
            df_["_year_"] = df_[date_col].dt.year
            df_["_month_"] = df_[date_col].dt.month
            # Time series: honoraires per month
            monto = detect_montant_column(df_) or "Montant honoraires (US $)"
            ts = df_.groupby([df_[date_col].dt.to_period("M")])[monto].sum().reset_index()
            ts[date_col] = ts[date_col].dt.to_timestamp()
            st.markdown("#### SÃ©rie temporelle mensuelle - Montant honoraires")
            if PLOTLY_AVAILABLE:
                fig = px.line(ts, x=date_col, y=monto, markers=True, title="Montant honoraires par mois")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.line_chart(ts.set_index(date_col)[monto])
            # Heatmap year x month
            st.markdown("#### Heatmap AnnÃ©e x Mois (Montant honoraires)")
            pivot = df_.groupby([df_["_year_"], df_["_month_"]])[monto].sum().unstack(fill_value=0)
            if PLOTLY_AVAILABLE:
                fig = go.Figure(data=go.Heatmap(
                    z=pivot.values,
                    x=[f"{m:02d}" for m in pivot.columns],
                    y=[str(y) for y in pivot.index],
                    colorscale="Viridis"
                ))
                fig.update_layout(title="Heatmap Montants par Mois/AnnÃ©e", xaxis_title="Mois", yaxis_title="AnnÃ©e", height=420)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.dataframe(pivot)
            # Category treemap
            st.markdown("#### RÃ©partition par CatÃ©gorie")
            cat_col = "Categories" if "Categories" in df_.columns else None
            if cat_col:
                cat_agg = df_.groupby(cat_col)[monto].sum().reset_index().sort_values(monto, ascending=False)
                if PLOTLY_AVAILABLE:
                    fig = px.treemap(cat_agg, path=[cat_col], values=monto, title="Montant par CatÃ©gorie")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.bar_chart(cat_agg.set_index(cat_col)[monto])
            # Top N clients
            st.markdown("#### Top N clients par Montant honoraires")
            try:
                topn = int(st.slider("Top N", 5, 50, 10, key=skey("anal","topn")))
            except Exception:
                topn = 10
            top_clients = df_.groupby("Nom")[monto].sum().reset_index().sort_values(monto, ascending=False).head(topn)
            if PLOTLY_AVAILABLE:
                fig = px.bar(top_clients, x="Nom", y=monto, title=f"Top {topn} clients", labels={monto:"Montant"})
                fig.update_layout(xaxis_tickangle=-45, height=420)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.dataframe(top_clients)
            # Comparison A vs B quick UI (reuse dashboard choices)
            st.markdown("#### Comparaison rapide de deux pÃ©riodes (AnnÃ©e+Mois)")
            colA1, colA2, colB1, colB2 = st.columns([1,1,1,1])
            with colA1:
                years = sorted(df_["_year_"].dropna().unique().astype(int).tolist())
                a_year = st.selectbox("AnnÃ©e A", options=[""]+ [str(y) for y in years], index=0, key=skey("anal","ayear"))
            with colA2:
                a_month = st.selectbox("Mois A", options=["Tous"] + [str(i) for i in range(1,13)], index=0, key=skey("anal","amonth"))
            with colB1:
                b_year = st.selectbox("AnnÃ©e B", options=[""]+ [str(y) for y in years], index=0, key=skey("anal","byear"))
            with colB2:
                b_month = st.selectbox("Mois B", options=["Tous"] + [str(i) for i in range(1,13)], index=0, key=skey("anal","bmonth"))
            if st.button("Comparer maintenant", key=skey("anal","compare_btn")):
                base = df_.copy()
                def subset(df, y, m):
                    out = df.copy()
                    if y:
                        out = out[out["_year_"] == int(y)]
                    if m and m != "Tous":
                        out = out[out["_month_"] == int(m)]
                    return out
                A = subset(base, a_year, a_month)
                B = subset(base, b_year, b_month)
                def kpis(d):
                    hon = float(d.get(monto,0).apply(lambda x: _to_num(x)).sum())
                    acom = 0.0
                    for ac in detect_acompte_columns(d):
                        acom += float(d.get(ac,0).apply(lambda x: _to_num(x)).sum())
                    return {"count":len(d),"hon":hon,"acom":acom,"solde":hon - acom}
                ka = kpis(A); kb = kpis(B)
                st.markdown(f"PÃ©riode A: {a_year or 'â€”'} / {a_month or 'Tous'} â€” PÃ©riode B: {b_year or 'â€”'} / {b_month or 'Tous'}")
                c1,c2 = st.columns(2)
                with c1:
                    st.write(ka)
                with c2:
                    st.write(kb)
                comp_df = pd.DataFrame([
                    {"metric":"Dossiers","A":ka["count"],"B":kb["count"]},
                    {"metric":"Honoraires","A":ka["hon"],"B":kb["hon"]},
                    {"metric":"Acomptes","A":ka["acom"],"B":kb["acom"]},
                    {"metric":"Solde","A":ka["solde"],"B":kb["solde"]}
                ])
                if PLOTLY_AVAILABLE:
                    fig = go.Figure(data=[go.Bar(name='A', x=comp_df['metric'], y=comp_df['A']), go.Bar(name='B', x=comp_df['metric'], y=comp_df['B'])])
                    fig.update_layout(barmode='group', height=420)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.dataframe(comp_df)

# ---- Add tab ----
with tabs[3]:
    st.subheader("âž• Ajouter un nouveau client")
    df_live = _get_df_live_safe()
    next_dossier_num = get_next_dossier_numeric(df_live)
    next_dossier = str(next_dossier_num)
    next_id_client = make_id_client_datebased(df_live)
    st.markdown(f"**ID_Client (auto)**: {next_id_client}")
    st.markdown(f"**Dossier N (auto)**: {next_dossier}")

    add_date = st.date_input("Date (Ã©vÃ©nement)", value=date.today(), key=skey("addtab","date"))
    add_nom = st.text_input("Nom du client", value="", placeholder="Nom complet du client", key=skey("addtab","nom"))

    categories_options = visa_categories if visa_categories else (unique_nonempty(df_live["Categories"]) if "Categories" in df_live.columns else [])
    r3c1, r3c2, r3c3 = st.columns([1.2,1.6,1.6])
    with r3c1:
        add_cat = st.selectbox("CatÃ©gorie", options=[""] + categories_options, index=0, key=skey("addtab","cat"))
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
                add_sub_options = sorted({str(x).strip() for x in df_live["Sous-categorie"].dropna().astype(str).tolist()})
            except Exception:
                add_sub_options = []
        add_sub = st.selectbox("Sous-catÃ©gorie", options=[""] + add_sub_options, index=0, key=skey("addtab","sub"))
    with r3c3:
        specific_options = get_visa_options(add_cat, add_sub)
        if specific_options:
            add_visa = st.selectbox("Visa (options)", options=[""] + specific_options, index=0, key=skey("addtab","visa"))
        else:
            add_visa = st.text_input("Visa", value="", key=skey("addtab","visa"))

    r4c1, r4c2 = st.columns([1.4, 1.0])
    with r4c1:
        add_montant = st.text_input("Montant honoraires (US $)", value="0", key=skey("addtab","montant"))
    with r4c2:
        a1 = st.text_input("Acompte 1", value="0", key=skey("addtab","ac1"))
    r5c1, r5c2 = st.columns([1.6,1.0])
    with r5c1:
        a1_date = st.date_input("Date Acompte 1", value=None, key=skey("addtab","ac1_date"))
    with r5c2:
        st.caption("Mode de rÃ¨glement")
        pay_cb = st.checkbox("CB", value=False, key=skey("addtab","pay_cb"))
        pay_cheque = st.checkbox("Cheque", value=False, key=skey("addtab","pay_cheque"))
        pay_virement = st.checkbox("Virement", value=False, key=skey("addtab","pay_virement"))
        pay_venmo = st.checkbox("Venmo", value=False, key=skey("addtab","pay_venmo"))

    add_escrow = st.checkbox("Escrow", value=False, key=skey("addtab","escrow"))
    add_comments = st.text_area("Commentaires", value="", key=skey("addtab","comments"))

    if st.button("Ajouter", key=skey("addtab","btn_add")):
        try:
            new_row = {c: "" for c in COLS_CLIENTS}
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
            modes = []
            if st.session_state.get(skey("addtab","pay_cb"), False): modes.append("CB")
            if st.session_state.get(skey("addtab","pay_cheque"), False): modes.append("Cheque")
            if st.session_state.get(skey("addtab","pay_virement"), False): modes.append("Virement")
            if st.session_state.get(skey("addtab","pay_venmo"), False): modes.append("Venmo")
            new_row["ModeReglement"] = ",".join(modes)
            new_row["ModeReglement_Ac1"] = ",".join(modes) if modes else ""
            new_row["ModeReglement_Ac2"] = ""
            new_row["ModeReglement_Ac3"] = ""
            new_row["ModeReglement_Ac4"] = ""
            new_row["PayÃ©"] = new_row["Acompte 1"]
            new_row["Solde"] = new_row["Montant honoraires (US $)"] + new_row["Autres frais (US $)"] - new_row["PayÃ©"]
            new_row["Solde Ã  percevoir (US $)"] = new_row["Solde"]
            new_row["Escrow"] = 1 if st.session_state.get(skey("addtab","escrow"), False) else 0
            new_row["Commentaires"] = add_comments
            ensure_flag_columns(new_row, DEFAULT_FLAGS)
            for f in DEFAULT_FLAGS:
                new_row[f] = 0
            df_live = _get_df_live_safe()
            df_live = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            df_live = recalc_payments_and_solde(df_live)
            _set_df_live(df_live)
            _persist_clients_cache(df_live)
            st.success(f"Dossier ajoutÃ© : ID_Client {next_id_client} â€” Dossier N {next_dossier}")
        except Exception as e:
            st.error(f"Erreur ajout: {e}")

# ---- Gestion tab ----
with tabs[4]:
    st.subheader("âœï¸ / ðŸ—‘ï¸ Gestion â€” Modifier / Supprimer")
    df_live = _get_df_live_safe()
    # defensive ensure columns exist
    for c in COLS_CLIENTS:
        if c not in df_live.columns:
            if "Date" in c:
                df_live[c] = pd.NaT
            elif c in NUMERIC_TARGETS:
                df_live[c] = 0.0
            elif c == "Escrow":
                df_live[c] = 0
            elif c in ["RFE", "Dossiers envoyÃ©", "Dossier approuvÃ©", "Dossier refusÃ©", "Dossier AnnulÃ©"]:
                df_live[c] = 0
            else:
                df_live[c] = ""
    if df_live is None or df_live.empty:
        st.info("Aucun dossier Ã  modifier ou supprimer.")
    else:
        choices = [f"{i} | {df_live.at[i,'Dossier N']} | {df_live.at[i,'Nom']}" for i in range(len(df_live))]
        sel = st.selectbox("SÃ©lectionner ligne Ã  modifier", options=[""] + choices, key=skey("edit","select"))
        if sel:
            idx = int(sel.split("|")[0].strip())
            df_live = recalc_payments_and_solde(df_live)
            row = df_live.loc[idx].copy()

            def txt(v):
                if pd.isna(v):
                    return ""
                s = str(v)
                if s.strip().lower() in ("nan","none","na","n/a"):
                    return ""
                return s

            def _safe_row_date_local(colname: str):
                try:
                    raw = row.get(colname)
                except Exception:
                    raw = None
                return _date_or_none_safe(raw)

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

            row_modes_general = _parse_modes(row.get("ModeReglement",""))
            row_mode_ac1 = _parse_modes(row.get("ModeReglement_Ac1","")) or row_modes_general
            row_mode_ac2 = _parse_modes(row.get("ModeReglement_Ac2",""))
            row_mode_ac3 = _parse_modes(row.get("ModeReglement_Ac3",""))
            row_mode_ac4 = _parse_modes(row.get("ModeReglement_Ac4",""))

            with st.form(key=skey("form_edit", str(idx))):
                c_name, c_solde = st.columns([2.5,1])
                with c_name:
                    st.markdown(f"### {txt(row.get('Nom',''))}")
                with c_solde:
                    try:
                        sol_due_num = _to_num(row.get("Solde Ã  percevoir (US $)", row.get("Solde",0)))
                        st.markdown(f"**Solde dÃ»**: {_fmt_money(sol_due_num)}")
                    except Exception:
                        st.markdown("**Solde dÃ»**: $0.00")

                r1c1, r1c2, r1c3 = st.columns([1.4,1.0,1.2])
                with r1c1:
                    st.markdown(f"**ID_Client :** {txt(row.get('ID_Client',''))}")
                with r1c2:
                    e_dossier = st.text_input("Dossier N", value=txt(row.get("Dossier N","")), key=skey("edit","dossier", str(idx)))
                with r1c3:
                    e_date = st.date_input("Date (Ã©vÃ©nement)", value=_safe_row_date_local("Date"), key=skey("edit","date", str(idx)))

                # Category / Sous-categorie / Visa
                c_cat, c_sub, c_visa = st.columns([1.4,1.6,1.6])
                with c_cat:
                    cur_cat = txt(row.get("Categories",""))
                    edit_cat = st.text_input("CatÃ©gorie", value=cur_cat, key=skey("edit","cat", str(idx)))
                with c_sub:
                    cur_sub = txt(row.get("Sous-categorie",""))
                    edit_sub = st.text_input("Sous-catÃ©gorie", value=cur_sub, key=skey("edit","sub", str(idx)))
                with c_visa:
                    cur_visa = txt(row.get("Visa",""))
                    visa_opts = get_visa_options(edit_cat if 'edit_cat' in locals() else cur_cat, edit_sub if 'edit_sub' in locals() else cur_sub)
                    if visa_opts:
                        default_index = 0
                        if cur_visa in visa_opts:
                            default_index = visa_opts.index(cur_visa) + 1
                        edit_visa = st.selectbox("Visa", options=[""]+visa_opts, index=default_index, key=skey("edit","visa", str(idx)))
                    else:
                        edit_visa = st.text_input("Visa", value=cur_visa, key=skey("edit","visa", str(idx)))

                # Montants
                m1, m2, m3 = st.columns([1.2,1.0,1.0])
                with m1:
                    e_montant = st.text_input("Montant honoraires (US $)", value=txt(row.get("Montant honoraires (US $)","")), key=skey("edit","montant", str(idx)))
                with m2:
                    e_autres = st.text_input("Autres frais (US $)", value=txt(row.get("Autres frais (US $)","")), key=skey("edit","autres", str(idx)))
                with m3:
                    try:
                        total_montant_val = _to_num(e_montant) + _to_num(e_autres)
                    except Exception:
                        total_montant_val = _to_num(row.get("Montant honoraires (US $)",0)) + _to_num(row.get("Autres frais (US $)",0))
                    st.text_input("Montant Total", value=str(total_montant_val), key=skey("edit","montant_total", str(idx)), disabled=True)

                # Acomptes
                r_ac_1, r_ac_2, r_ac_3, r_ac_4 = st.columns([1.0,1.0,1.0,1.0])
                with r_ac_1:
                    e_ac1 = st.text_input("Acompte 1", value=txt(row.get("Acompte 1","")), key=skey("edit","ac1", str(idx)))
                with r_ac_2:
                    e_ac2 = st.text_input("Acompte 2", value=txt(row.get("Acompte 2","")), key=skey("edit","ac2", str(idx)))
                with r_ac_3:
                    e_ac3 = st.text_input("Acompte 3", value=txt(row.get("Acompte 3","")), key=skey("edit","ac3", str(idx)))
                with r_ac_4:
                    e_ac4 = st.text_input("Acompte 4", value=txt(row.get("Acompte 4","")), key=skey("edit","ac4", str(idx)))

                # Dates + modes
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
                    e_flag_envoye = st.checkbox("Dossiers envoyÃ©", value=bool(int(row.get("Dossiers envoyÃ©", 0))) if not pd.isna(row.get("Dossiers envoyÃ©", 0)) else False, key=skey("edit","flag_envoye", str(idx)))
                with f2:
                    e_flag_approuve = st.checkbox("Dossier approuvÃ©", value=bool(int(row.get("Dossier approuvÃ©", 0))) if not pd.isna(row.get("Dossier approuvÃ©", 0)) else False, key=skey("edit","flag_approuve", str(idx)))
                with f3:
                    e_flag_refuse = st.checkbox("Dossier refusÃ©", value=bool(int(row.get("Dossier refusÃ©", 0))) if not pd.isna(row.get("Dossier refusÃ©", 0)) else False, key=skey("edit","flag_refuse", str(idx)))
                with f4:
                    e_flag_annule = st.checkbox("Dossier AnnulÃ©", value=bool(int(row.get("Dossier AnnulÃ©", 0))) if not pd.isna(row.get("Dossier AnnulÃ©", 0)) else False, key=skey("edit","flag_annule", str(idx)))

                e_escrow = st.checkbox("Escrow", value=bool(int(row.get("Escrow", 0))) if not pd.isna(row.get("Escrow", 0)) else False, key=skey("edit","escrow", str(idx)))

                other_flag_set = any([e_flag_envoye, e_flag_approuve, e_flag_refuse, e_flag_annule])
                if not other_flag_set:
                    st.markdown("**RFE** (active uniquement si un des Ã©tats est cochÃ©)")
                    e_flag_rfe = st.checkbox("RFE", value=bool(int(row.get("RFE", 0))) if not pd.isna(row.get("RFE", 0)) else False, key=skey("edit","flag_rfe", str(idx)), disabled=True)
                else:
                    e_flag_rfe = st.checkbox("RFE", value=bool(int(row.get("RFE", 0))) if not pd.isna(row.get("RFE", 0)) else False, key=skey("edit","flag_rfe", str(idx)))

                e_comments = st.text_area("Commentaires", value=txt(row.get("Commentaires","")), key=skey("edit","comments", str(idx)))

                save = st.form_submit_button("Enregistrer modifications")
                if save:
                    try:
                        df_live.at[idx, "Dossier N"] = e_dossier
                        df_live.at[idx, "Nom"] = txt(row.get("Nom",""))
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
                        old_general = parse_modes_global(row.get("ModeReglement",""))
                        combined = set(old_general + list(e_mode_ac1))
                        df_live.at[idx, "ModeReglement"] = ",".join(sorted(list(combined)))
                        df_live.at[idx, "Dossiers envoyÃ©"] = 1 if e_flag_envoye else 0
                        df_live.at[idx, "Dossier approuvÃ©"] = 1 if e_flag_approuve else 0
                        df_live.at[idx, "Dossier refusÃ©"] = 1 if e_flag_refuse else 0
                        df_live.at[idx, "Dossier AnnulÃ©"] = 1 if e_flag_annule else 0
                        if e_flag_rfe and not any([e_flag_envoye, e_flag_approuve, e_flag_refuse, e_flag_annule]):
                            st.warning("RFE n'a pas Ã©tÃ© activÃ© car aucun Ã©tat (envoyÃ©/approuvÃ©/refusÃ©/annulÃ©) n'est cochÃ©.")
                            df_live.at[idx, "RFE"] = 0
                        else:
                            df_live.at[idx, "RFE"] = 1 if e_flag_rfe else 0
                        df_live.at[idx, "Escrow"] = 1 if e_escrow else 0
                        df_live.at[idx, "Categories"] = edit_cat if 'edit_cat' in locals() else cur_cat
                        df_live.at[idx, "Sous-categorie"] = edit_sub if 'edit_sub' in locals() else cur_sub
                        df_live.at[idx, "Visa"] = edit_visa if 'edit_visa' in locals() else cur_visa
                        df_live.at[idx, "Commentaires"] = e_comments
                        df_live = recalc_payments_and_solde(df_live)
                        df_live.at[idx, "Solde Ã  percevoir (US $)"] = df_live.at[idx, "Solde"]
                        _set_df_live(df_live)
                        _persist_clients_cache(df_live)
                        st.success("Modifications enregistrÃ©es.")
                    except Exception as e:
                        st.error(f"Erreur enregistrement: {e}")

    st.markdown("---")
    st.markdown("### Supprimer des dossiers")
    if df_live is None or df_live.empty:
        st.info("Aucun dossier Ã  supprimer.")
    else:
        choices_del = [f"{i} | {df_live.at[i,'Dossier N']} | {df_live.at[i,'Nom']}" for i in range(len(df_live))]
        selected_to_del = st.multiselect("SÃ©lectionnez les lignes Ã  supprimer", options=choices_del, key=skey("del","select"))
        if st.button("Supprimer sÃ©lection"):
            if selected_to_del:
                idxs = [int(s.split("|")[0].strip()) for s in selected_to_del]
                try:
                    df_live = df_live.drop(index=idxs).reset_index(drop=True)
                    df_live = recalc_payments_and_solde(df_live)
                    _set_df_live(df_live)
                    _persist_clients_cache(df_live)
                    st.success(f"{len(idxs)} ligne(s) supprimÃ©e(s).")
                except Exception as e:
                    st.error(f"Erreur suppression: {e}")
            else:
                st.warning("Aucune sÃ©lection pour suppression.")

# ---- Compta Client tab ----
with tabs[5]:
    st.subheader("ðŸ’³ Compta Client")
    df_live = recalc_payments_and_solde(_get_df_live_safe())
    if df_live is None or df_live.empty:
        st.info("Aucune donnÃ©e en mÃ©moire.")
    else:
        # detect Nom and Dossier columns
        col_nom = None
        col_dossier = None
        for c in df_live.columns:
            if c.strip().lower() == "nom":
                col_nom = c
            if c.strip().lower() in ("dossier n", "dossier", "dossier numÃ©ro", "dossier no", "dossier nÂ°"):
                col_dossier = c
        if col_nom is None:
            for c in df_live.columns:
                if "nom" in canonical_key(c) or "client" in canonical_key(c):
                    col_nom = c; break
        if col_dossier is None:
            for c in df_live.columns:
                if "dossier" in canonical_key(c) or "num" in canonical_key(c) or "numero" in canonical_key(c):
                    col_dossier = c; break
        if col_nom is None and col_dossier is None:
            st.warning("Impossible de trouver les colonnes 'Nom' ou 'Dossier N'. Colonnes disponibles :")
            st.write(list(df_live.columns))
        else:
            choices = []
            for i in range(len(df_live)):
                dn = str(df_live.at[i, col_dossier]) if col_dossier in df_live.columns else ""
                nm = str(df_live.at[i, col_nom]) if col_nom in df_live.columns else ""
                choices.append(f"{i} | {dn} | {nm}")
            sel = st.selectbox("SÃ©lectionner un client (par index | Dossier N | Nom)", options=[""] + choices, key=skey("compta","select"))
            if not sel:
                st.info("SÃ©lectionne une ligne pour afficher le relevÃ©.")
            else:
                idx = int(sel.split("|")[0].strip())
                df_live = recalc_payments_and_solde(df_live)
                if idx < 0 or idx >= len(df_live):
                    st.error("Index sÃ©lectionnÃ© invalide.")
                else:
                    row = df_live.loc[idx]
                    montant_col = detect_montant_column(df_live) or "Montant honoraires (US $)"
                    autres_col = detect_autres_column(df_live) or "Autres frais (US $)"
                    honoraires = _to_num(row.get(montant_col, 0))
                    autres = _to_num(row.get(autres_col, 0))
                    total_paye = _to_num(row.get("PayÃ©", 0))
                    solde = _to_num(row.get("Solde", honoraires + autres - total_paye))
                    st.markdown(f"### Fiche: {row.get(col_nom,'')} â€” Dossier {row.get(col_dossier,'')}")
                    st.markdown(f"- Montant honoraires : {_fmt_money(honoraires)}")
                    st.markdown(f"- Autres frais : {_fmt_money(autres)}")
                    st.markdown(f"- Total payÃ© : {_fmt_money(total_paye)}")
                    st.markdown(f"**Solde dÃ» : {_fmt_money(solde)}**")
                    st.markdown("---")
                    acomptes_cols = detect_acompte_columns(df_live)
                    data_ac = []
                    for ac in acomptes_cols:
                        val = _to_num(row.get(ac,0))
                        date_col = f"Date {ac}" if f"Date {ac}" in df_live.columns else ("Date Acompte 1" if ac=="Acompte 1" else "")
                        dval = row.get(date_col, "")
                        mode_col = "ModeReglement_Ac1" if ac=="Acompte 1" else f"ModeReglement_{ac.replace(' ','')}"
                        mode = row.get(mode_col, "")
                        if val and val != 0:
                            dd = ""
                            if isinstance(dval, pd.Timestamp):
                                dd = dval.date()
                            elif isinstance(dval, str) and dval.strip():
                                dd = dval
                            data_ac.append({"Acompte": ac, "Montant": _fmt_money(val), "Date": dd, "Mode": mode})
                    if data_ac:
                        st.table(data_ac)
                    else:
                        st.write("Aucun acompte enregistrÃ©.")
                    st.markdown("---")
                    if st.button("Exporter relevÃ© client (.xlsx)"):
                        try:
                            out_df = pd.DataFrame([{
                                "ID_Client": row.get("ID_Client",""),
                                "Dossier N": row.get(col_dossier,"") if col_dossier in row.index else "",
                                "Nom": row.get(col_nom,""),
                                "Montant honoraires": honoraires,
                                "Autres frais": autres,
                                "Total payÃ©": total_paye,
                                "Solde": solde
                            }])
                            buf = BytesIO()
                            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                                out_df.to_excel(writer, index=False, sheet_name="Releve")
                            buf.seek(0)
                            st.download_button(
                                "TÃ©lÃ©charger XLSX",
                                data=buf.getvalue(),
                                file_name=f"releve_client_{str(row.get(col_dossier,'')).replace('/','-')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as e:
                            st.error(f"Erreur export XLSX: {e}")



# app.py - Visa Manager (complete)
# - Features:
#   * Import Clients/Visa (xlsx/csv), normalize columns, heuristic mapping
#   * Import single ComptaCli fiche and persist to cache so re-upload not required
#   * Session-backed clients table editable in "Gestion"
#   * Compta Client tab: select a row (index | Dossier N | Nom) like Gestion and export .xlsx
#   * Dashboard: filters by Category/Subcategory, Year, Month (with "Tous"), custom date range, and comparison between two periods
#   * Analyses: multiple charts (time series monthly, heatmap year x month, category treemap, top-N clients, comparison bars)
# Requirements: pip install streamlit pandas openpyxl plotly
# Run: streamlit run app.py

import os
import json
import re
from io import BytesIO
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# optional plotly for richer charts
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_AVAILABLE = True
except Exception:
    PLOTLY_AVAILABLE = False

# [--- Reste des helpers, configs, mapping, etc. inchangÃ© ---]
# ... [aucune modification jusqu'Ã  la crÃ©ation des tabs]

# -------------------------
# Tabs UI (AJOUT DE L'ONGLET ESCROW)
# -------------------------
tabs = st.tabs([
    "ðŸ“„ Fichiers",
    "ðŸ“Š Dashboard",
    "ðŸ“ˆ Analyses",
    "âž• Ajouter",
    "âœï¸ / ðŸ—‘ï¸ Gestion",
    "ðŸ’³ Compta Client",
    "ðŸ’¾ Export",
    "ðŸ›¡ï¸ Escrow" # <-- AJOUT Escrow ici !
])

# ---- Files tab ----
with tabs[0]:
    st.header("ðŸ“‚ Fichiers")    
    # ... [bloc fichiers original inchangÃ©] ...

# ---- Dashboard tab ----
with tabs[1]:
    st.subheader("ðŸ“Š Dashboard")
    # ... [bloc dashboard original inchangÃ©] ...

# ---- Analyses tab ----
with tabs[2]:st.subheader("ðŸ“ˆ Analyses")    
    # ... [bloc analyses original inchangÃ©] ...

# ---- Ajouter tab ----
with tabs[3]:
    st.header("âž• Ajouter")
    # ... [bloc ajouter original inchangÃ©] ...

# ---- Gestion tab ----
with tabs[4]:
    st.header("âœï¸ / ðŸ—‘ï¸ Gestion")
    # ... [bloc gestion original inchangÃ©] ...

# ---- Compta Client tab ----
with tabs[5]:
    st.header("ðŸ’³ Compta Client")
    # ... [bloc compta client original inchangÃ©] ...

# ---- Export tab ----
with tabs[6]:
    st.header("ðŸ’¾ Export")
    # ... [bloc export original inchangÃ©] ...

# --- NOUVEAU ONGLET Escrow ---
with tabs[7]:
    st.subheader("ðŸ›¡ï¸ Escrow")
    df_live = _get_df_live_safe()
    if df_live is None or df_live.empty or "Escrow" not in df_live.columns:
        st.info("Aucun dossier Escrow dÃ©tectÃ©.")
    else:
        escrow_df = df_live[df_live["Escrow"] == 1].copy()
        colonnes_affichage = [
            "Dossier N",
            "Nom",
            "Date",
            "Acompte 1",
            "Date d'envoi"
        ]
        # Gestion de la colonne Date d'envoi (variante possible : "Date denvoi")
        if "Date d'envoi" not in escrow_df.columns and "Date denvoi" in escrow_df.columns:
            escrow_df = escrow_df.rename(columns={"Date denvoi": "Date d'envoi"})
        if "Date d'envoi" not in escrow_df.columns:
            escrow_df["Date d'envoi"] = pd.NaT

        colonnes_existantes = [c for c in colonnes_affichage if c in escrow_df.columns]

        st.markdown(f"**Nombre de dossiers Escrow : {len(escrow_df)}**")
        st.dataframe(escrow_df[colonnes_existantes].reset_index(drop=True), use_container_width=True)

        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            escrow_df[colonnes_existantes].to_excel(writer, index=False, sheet_name="Escrow")
        buf.seek(0)
        st.download_button(
            "TÃ©lÃ©charger XLSX (Escrow)",
            data=buf.getvalue(),
            file_name="Synthese_Escrow.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ... [fin du script inchangÃ©]

# ---- Export tab ----
with tabs[6]:
    st.header("ðŸ’¾ Export")
    df_live = _get_df_live_safe()
    if df_live is None or df_live.empty:
        st.info("Aucune donnÃ©e Ã  exporter.")
    else:
        st.write(f"Vue en mÃ©moire: {df_live.shape[0]} lignes, {df_live.shape[1]} colonnes")
        col1, col2 = st.columns(2)
        with col1:
            csv_bytes = df_live.to_csv(index=False).encode("utf-8")
            st.download_button("â¬‡ï¸ Export CSV", data=csv_bytes, file_name="Clients_export.csv", mime="text/csv")
        with col2:
            df_for_export = df_live.copy()
            try:
                montant_col = detect_montant_column(df_for_export) or "Montant honoraires (US $)"
                autres_col = detect_autres_column(df_for_export) or "Autres frais (US $)"
                acomptes_cols = detect_acompte_columns(df_for_export)
                df_for_export["_Montant_num_"] = df_for_export.get(montant_col,0).apply(lambda x: _to_num(x))
                df_for_export["_Autres_num_"] = df_for_export.get(autres_col,0).apply(lambda x: _to_num(x))
                for acc in acomptes_cols:
                    df_for_export[f"_num_{acc}"] = df_for_export.get(acc,0).apply(lambda x: _to_num(x))
                if acomptes_cols:
                    df_for_export["_Acomptes_sum_"] = df_for_export[[f"_num_{acc}" for acc in acomptes_cols]].sum(axis=1)
                else:
                    df_for_export["_Acomptes_sum_"] = 0.0
                df_for_export["Solde_formule"] = df_for_export["_Montant_num_"] + df_for_export["_Autres_num_"] - df_for_export["_Acomptes_sum_"]
                df_for_export["Solde Ã  percevoir (US $)"] = df_for_export["Solde_formule"]
            except Exception:
                df_for_export["Solde_formule"] = df_for_export.get("Solde",0).apply(lambda x: _to_num(x))
                df_for_export["Solde Ã  percevoir (US $)"] = df_for_export.get("Solde Ã  percevoir (US $)",0).apply(lambda x: _to_num(x))
            drop_cols = [c for c in df_for_export.columns if c.startswith("_num_") or c in ["_Montant_num_","_Autres_num_","_Acomptes_sum_"]]
            try:
                df_export_final = df_for_export.drop(columns=drop_cols)
            except Exception:
                df_export_final = df_for_export.copy()
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df_export_final.to_excel(writer, index=False, sheet_name="Clients")
            out_bytes = buf.getvalue()
            st.download_button("â¬‡ï¸ Export XLSX (avec colonne Solde_formule)", data=out_bytes, file_name="Clients_export_with_Solde_formule.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# End of file
