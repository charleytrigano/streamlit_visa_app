# app.py - Visa Manager (complete, conservative column pruning - Choice A)
# - Removed columns that were not used by Add/Gestion/Export UIs (conservative)
# - All helper functions defined before use
# - Robust numeric parsing (no "nan"), safe date handling
# - UI tabs: Fichiers, Dashboard, Analyses, Ajouter, Gestion, Export
# - Keep core columns used by UI; removed historical/log columns:
#   Removed: "Date de création","Créé par","Dernière modification","Modifié par","Date d'envoi","Date reponse","Escrow"
# Requirements: pip install streamlit pandas openpyxl
# Run: streamlit run app.py

import os
import json
import re
from io import BytesIO
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# -------------------------
# Configuration & constants
# -------------------------
APP_TITLE = "🛂 Visa Manager"
# Conservative columns: keep only those used in Add / Gestion / Export / Dashboard
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
    "Payé",
    "Solde",
    "Solde à percevoir (US $)",
    "Acompte 1","Date Acompte 1",
    "Acompte 2","Date Acompte 2",
    "Acompte 3","Date Acompte 3",
    "Acompte 4","Date Acompte 4",
    "RFE",
    "Dossiers envoyé",
    "Dossier approuvé",
    "Dossier refusé",
    "Dossier Annulé",
    "Commentaires",
    # Payment mode columns
    "ModeReglement",
    "ModeReglement_Ac1","ModeReglement_Ac2","ModeReglement_Ac3","ModeReglement_Ac4"
]
MEMO_FILE = "_vmemory.json"
CACHE_CLIENTS = "_clients_cache.bin"
CACHE_VISA = "_visa_cache.bin"
SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"
SID = "vmgr"
DEFAULT_START_CLIENT_ID = 13057
CURRENT_USER = "charleytrigano"
DEFAULT_FLAGS = ["RFE","Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé"]

def skey(*parts: str) -> str:
    return f"{SID}_" + "_".join([p for p in parts if p])

# -------------------------
# Parsing / formatting helpers
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
    replace_map = {"é":"e","è":"e","ê":"e","ë":"e","à":"a","â":"a","î":"i","ï":"i","ô":"o","ö":"o","ù":"u","û":"u","ü":"u","ç":"c"}
    for k,v in replace_map.items():
        s2 = s2.replace(k,v)
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
    # Robust conversion: treat None / NaN / "nan" as 0.0
    try:
        if x is None:
            return 0.0
        try:
            if pd.isna(x):
                return 0.0
        except Exception:
            pass
        if isinstance(x, (int,float)):
            return float(x)
        s = str(x).strip()
        if s == "" or s.lower() in ("na","n/a","nan","none","null"):
            return 0.0
        s = s.replace("\u202f","").replace("\xa0","").replace(" ","")
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
    if isinstance(x,(int,float)) and (not pd.isna(x)):
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
        d = pd.to_datetime(v, errors="coerce")
        if pd.isna(d):
            return None
        return date(int(d.year), int(d.month), int(d.day))
    except Exception:
        return None

# -------------------------
# Column heuristics and detection
# -------------------------
COL_CANDIDATES = {
    "id client":"ID_Client","idclient":"ID_Client",
    "dossier n":"Dossier N","dossier":"Dossier N",
    "nom":"Nom","date":"Date",
    "categories":"Categories","categorie":"Categories",
    "sous categorie":"Sous-categorie","sous-categorie":"Sous-categorie",
    "visa":"Visa","montant":"Montant honoraires (US $)","autres frais":"Autres frais (US $)",
    "paye":"Payé","payé":"Payé","solde":"Solde","rfe":"RFE","mode reglement":"ModeReglement"
}

NUMERIC_TARGETS = [
    "Montant honoraires (US $)","Autres frais (US $)","Payé","Solde","Solde à percevoir (US $)",
    "Acompte 1","Acompte 2","Acompte 3","Acompte 4"
]

def detect_acompte_columns(df: pd.DataFrame) -> List[str]:
    if df is None or df.empty:
        return []
    cols = [c for c in df.columns if "acompte" in canonical_key(c)]
    def keyfn(n):
        m = re.search(r"(\d+)", n)
        return int(m.group(1)) if m else 999
    return sorted(cols, key=keyfn)

def detect_montant_column(df: pd.DataFrame) -> Optional[str]:
    if df is None or df.empty:
        return None
    candidates = ["Montant honoraires (US $)","Montant honoraires","Montant"]
    for c in candidates:
        if c in df.columns:
            return c
    for c in df.columns:
        if "montant" in canonical_key(c) or "honorair" in canonical_key(c):
            return c
    return None

def detect_autres_column(df: pd.DataFrame) -> Optional[str]:
    if df is None or df.empty:
        return None
    candidates = ["Autres frais (US $)","Autres frais","Autres"]
    for c in candidates:
        if c in df.columns:
            return c
    for c in df.columns:
        if "autre" in canonical_key(c) or "frais" in canonical_key(c):
            return c
    return None

def map_columns_heuristic(df: Any) -> Tuple[pd.DataFrame, Dict[str,str]]:
    if not isinstance(df,pd.DataFrame):
        return pd.DataFrame(), {}
    mapping = {}
    for c in list(df.columns):
        k = canonical_key(c)
        mapped = None
        if k in COL_CANDIDATES:
            mapped = COL_CANDIDATES[k]
        else:
            for cand, std in COL_CANDIDATES.items():
                if cand in k:
                    mapped = std
                    break
        mapping[c] = mapped or normalize_header_text(c)
    newnames = {}
    seen = {}
    for orig, new in mapping.items():
        base = new
        cnt = seen.get(base,0)
        if cnt:
            new_name = f"{base}_{cnt+1}"
            seen[base] = cnt+1
        else:
            new_name = base
            seen[base] = 1
        newnames[orig] = new_name
    try:
        df = df.rename(columns=newnames)
    except Exception:
        pass
    return df, newnames

# -------------------------
# Visa maps initialization (safe before UI)
# -------------------------
visa_map: Dict[str,List[str]] = {}
visa_map_norm: Dict[str,List[str]] = {}
visa_sub_options_map: Dict[str,List[str]] = {}
visa_categories: List[str] = []

def get_visa_options(cat: Optional[str], sub: Optional[str]) -> List[str]:
    try:
        if sub:
            ksub = canonical_key(sub)
            if ksub in visa_sub_options_map:
                return visa_sub_options_map.get(ksub, [])[:]
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

# -------------------------
# Normalize / recalc functions
# -------------------------
def _ensure_columns(df: Any, cols: List[str]) -> pd.DataFrame:
    if not isinstance(df,pd.DataFrame):
        df = pd.DataFrame()
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            if c in NUMERIC_TARGETS:
                out[c] = 0.0
            elif c in ["RFE","Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé"]:
                out[c] = 0
            else:
                out[c] = "" if "Date" not in c else pd.NaT
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
                elif c in ["RFE","Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé"]:
                    safe[c] = 0
                elif "Date" in c:
                    safe[c] = pd.NaT
                else:
                    safe[c] = ""
        return safe

def normalize_clients_for_live(raw: Any) -> pd.DataFrame:
    df_raw = raw
    if not isinstance(df_raw,pd.DataFrame):
        maybe = read_any_table(df_raw, sheet=None, debug_prefix="[normalize] ")
        df_raw = maybe if isinstance(maybe,pd.DataFrame) else pd.DataFrame()
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
    # ensure acomptes
    for acc in ["Acompte 1","Acompte 2","Acompte 3","Acompte 4"]:
        if acc not in df.columns:
            df[acc] = 0.0
    try:
        acomptes = detect_acompte_columns(df)
        if acomptes:
            df["Payé"] = df[acomptes].fillna(0).apply(lambda row: sum([_to_num(row[c]) for c in acomptes]), axis=1)
        else:
            df["Payé"] = df.get("Payé",0).apply(lambda x: _to_num(x))
    except Exception:
        df["Payé"] = df.get("Payé",0).apply(lambda x: _to_num(x))
    try:
        montant_col = detect_montant_column(df) or "Montant honoraires (US $)"
        autres_col = detect_autres_column(df) or "Autres frais (US $)"
        df[montant_col] = df.get(montant_col,0).apply(lambda x: _to_num(x))
        df[autres_col] = df.get(autres_col,0).apply(lambda x: _to_num(x))
        df["Solde"] = df[montant_col] + df[autres_col] - df["Payé"]
        df["Solde à percevoir (US $)"] = df["Solde"].copy()
    except Exception:
        df["Solde"] = df.get("Solde",0).apply(lambda x: _to_num(x))
        df["Solde à percevoir (US $)"] = df.get("Solde à percevoir (US $)",0).apply(lambda x: _to_num(x))
    # textual columns safe
    for c in ["Nom","Categories","Sous-categorie","Visa","Commentaires","ModeReglement","ModeReglement_Ac1","ModeReglement_Ac2","ModeReglement_Ac3","ModeReglement_Ac4"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str)
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
    except Exception:
        out["Solde"] = out.get("Solde",0).apply(lambda x: _to_num(x))
    return out

# -------------------------
# I/O helpers (Excel/CSV)
# -------------------------
def try_read_excel_from_bytes(b: bytes, sheet_name: Optional[str]=None) -> Optional[pd.DataFrame]:
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

def read_any_table(src: Any, sheet: Optional[str]=None, debug_prefix: str="") -> Optional[pd.DataFrame]:
    def _log(msg: str):
        try:
            st.sidebar.info(f"{debug_prefix}{msg}")
        except Exception:
            pass
    if src is None:
        _log("read_any_table: src is None")
        return None
    try:
        if isinstance(src,(bytes,bytearray)):
            df = try_read_excel_from_bytes(bytes(src), sheet)
            if df is not None: return df
            for sep in [";",","]:
                for enc in ["utf-8","latin-1","cp1252"]:
                    try:
                        return pd.read_csv(BytesIO(src), sep=sep, encoding=enc, on_bad_lines="skip")
                    except Exception:
                        continue
            return None
        if isinstance(src, BytesIO):
            b = src.getvalue()
            df = try_read_excel_from_bytes(b, sheet)
            if df is not None: return df
            for sep in [";",","]:
                for enc in ["utf-8","latin-1","cp1252"]:
                    try:
                        return pd.read_csv(BytesIO(b), sep=sep, encoding=enc, on_bad_lines="skip")
                    except Exception:
                        continue
            return None
        if hasattr(src,"read") and hasattr(src,"name"):
            try:
                data = src.getvalue()
            except Exception:
                try:
                    src.seek(0); data = src.read()
                except Exception:
                    data = None
            if data:
                df = try_read_excel_from_bytes(data, sheet)
                if df is not None: return df
                for sep in [";",","]:
                    for enc in ["utf-8","latin-1","cp1252"]:
                        try:
                            return pd.read_csv(BytesIO(data), sep=sep, encoding=enc, on_bad_lines="skip")
                        except Exception:
                            continue
            return None
        if isinstance(src,(str,os.PathLike)):
            p = str(src)
            if not os.path.exists(p):
                _log(f"path does not exist: {p}")
                return None
            if p.lower().endswith(".csv"):
                for sep in [";",","]:
                    for enc in ["utf-8","latin-1","cp1252"]:
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
# Session-safe DataFrame in st.session_state
# -------------------------
DF_LIVE_KEY = skey("df_live")
if DF_LIVE_KEY not in st.session_state:
    st.session_state[DF_LIVE_KEY] = pd.DataFrame(columns=COLS_CLIENTS)

def _get_df_live() -> pd.DataFrame:
    df = st.session_state.get(DF_LIVE_KEY)
    if df is None or not isinstance(df,pd.DataFrame):
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
# UI bootstrap (sidebar)
# -------------------------
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.sidebar.header("📂 Fichiers")
last_clients_path = ""
last_visa_path = ""
try:
    if os.path.exists(MEMO_FILE):
        with open(MEMO_FILE,"r",encoding="utf-8") as f:
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
        with open(MEMO_FILE,"w",encoding="utf-8") as f:
            json.dump({"clients": clients_path_in or "", "visa": visa_path_in or ""}, f, ensure_ascii=False, indent=2)
        st.sidebar.success("Chemins sauvegardés.")
    except Exception:
        st.sidebar.error("Impossible de sauvegarder les chemins.")

# Handle upload caching
clients_src_for_read = None
visa_src_for_read = None
if up_clients is not None:
    try:
        clients_bytes = up_clients.getvalue()
        with open(CACHE_CLIENTS,"wb") as f:
            f.write(clients_bytes)
        clients_src_for_read = BytesIO(clients_bytes)
    except Exception:
        clients_src_for_read = None
elif clients_path_in:
    clients_src_for_read = clients_path_in
elif os.path.exists(CACHE_CLIENTS):
    try:
        clients_bytes = open(CACHE_CLIENTS,"rb").read()
        clients_src_for_read = BytesIO(clients_bytes)
    except Exception:
        clients_src_for_read = None

if up_visa is not None:
    try:
        visa_bytes = up_visa.getvalue()
        with open(CACHE_VISA,"wb") as f:
            f.write(visa_bytes)
        visa_src_for_read = BytesIO(visa_bytes)
    except Exception:
        visa_src_for_read = None
elif visa_path_in:
    visa_src_for_read = visa_path_in
elif os.path.exists(CACHE_VISA):
    try:
        visa_bytes = open(CACHE_VISA,"rb").read()
        visa_src_for_read = BytesIO(visa_bytes)
    except Exception:
        visa_src_for_read = None

# -------------------------
# Read raw tables (if provided)
# -------------------------
df_clients_raw: Optional[pd.DataFrame] = None
df_visa_raw: Optional[pd.DataFrame] = None

try:
    if clients_src_for_read is not None:
        maybe = read_any_table(clients_src_for_read, sheet=SHEET_CLIENTS, debug_prefix="[Clients] ")
        if maybe is None:
            maybe = read_any_table(clients_src_for_read, sheet=None, debug_prefix="[Clients fallback] ")
        if isinstance(maybe,pd.DataFrame):
            df_clients_raw = maybe
except Exception:
    df_clients_raw = None

try:
    if visa_src_for_read is not None:
        maybe = read_any_table(visa_src_for_read, sheet=SHEET_VISA, debug_prefix="[Visa] ")
        if maybe is None:
            maybe = read_any_table(visa_src_for_read, sheet=None, debug_prefix="[Visa fallback] ")
        if isinstance(maybe,pd.DataFrame):
            df_visa_raw = maybe
except Exception:
    df_visa_raw = None

# sanitize visa sheet and build maps if present
if isinstance(df_visa_raw,pd.DataFrame) and not df_visa_raw.empty:
    try:
        df_visa_raw = df_visa_raw.fillna("")
        for c in df_visa_raw.columns:
            try:
                df_visa_raw[c] = df_visa_raw[c].astype(str).str.strip()
            except Exception:
                pass
    except Exception:
        pass
    df_visa_mapped,_ = map_columns_heuristic(df_visa_raw)
    try:
        df_visa_mapped = df_visa_mapped.pipe(lambda d: d)  # noop but safe
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
    raw_vm = {k:[s for s in v if s and str(s).strip().lower()!="nan"] for k,v in raw_vm.items()}
    visa_map = {k.strip():[s.strip() for s in v] for k,v in raw_vm.items()}
    visa_map_norm = {canonical_key(k):v for k,v in visa_map.items()}
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

# expose globals
globals().update({
    "visa_map": visa_map,
    "visa_map_norm": visa_map_norm,
    "visa_categories": visa_categories,
    "visa_sub_options_map": visa_sub_options_map
})

# -------------------------
# Initialize live df in session state
# -------------------------
df_all = normalize_clients_for_live(df_clients_raw if df_clients_raw is not None else None)
df_all = recalc_payments_and_solde(df_all)
if isinstance(df_all,pd.DataFrame) and not df_all.empty:
    st.session_state[DF_LIVE_KEY] = df_all.copy()
else:
    if DF_LIVE_KEY not in st.session_state or st.session_state[DF_LIVE_KEY] is None:
        st.session_state[DF_LIVE_KEY] = pd.DataFrame(columns=COLS_CLIENTS)

# -------------------------
# Small UI helpers
# -------------------------
def unique_nonempty(series):
    try:
        vals = series.dropna().astype(str).tolist()
    except Exception:
        vals = []
    out = []
    for v in vals:
        s = str(v).strip()
        if s=="" or s.lower()=="nan":
            continue
        out.append(s)
    return sorted(list(dict.fromkeys(out)))

def kpi_html(label:str, value:str, sub:str="")->str:
    html = f"""
    <div style="border:1px solid rgba(255,255,255,0.04); border-radius:6px; padding:8px 10px; margin:6px 4px;">
      <div style="font-size:12px; color:#666;">{label}</div>
      <div style="font-size:18px; font-weight:700; margin-top:4px;">{value}</div>
      <div style="font-size:11px; color:#888; margin-top:4px;">{sub}</div>
    </div>
    """
    return html

# -------------------------
# Tabs UI (Files / Dashboard / Analyses / Add / Gestion / Export)
# -------------------------
tabs = st.tabs(["📄 Fichiers","📊 Dashboard","📈 Analyses","➕ Ajouter","✏️ / 🗑️ Gestion","💾 Export"])

# Files tab
with tabs[0]:
    st.header("📂 Fichiers")
    c1,c2 = st.columns(2)
    with c1:
        st.subheader("Clients")
        if up_clients is not None:
            st.text(f"Upload: {getattr(up_clients,'name','')}")
        elif isinstance(clients_src_for_read,str) and clients_src_for_read:
            st.text(f"Chargé depuis: {clients_src_for_read}")
        elif os.path.exists(CACHE_CLIENTS):
            st.text("Chargé depuis le cache local")
        if df_clients_raw is None or (isinstance(df_clients_raw,pd.DataFrame) and df_clients_raw.empty):
            st.warning("Aucun fichier Clients detecté.")
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
        elif isinstance(visa_src_for_read,str) and visa_src_for_read:
            st.text(f"Chargé depuis: {visa_src_for_read}")
        elif os.path.exists(CACHE_VISA):
            st.text("Chargé depuis le cache local")
        if df_visa_raw is None or (isinstance(df_visa_raw,pd.DataFrame) and df_visa_raw.empty):
            st.warning("Aucun fichier Visa detecté.")
        else:
            st.success(f"Visa lu: {df_visa_raw.shape[0]} lignes")
            try:
                st.dataframe(df_visa_raw.head(100).reset_index(drop=True), use_container_width=True, height=360)
            except Exception:
                st.write(df_visa_raw.head(8))
    st.markdown("---")
    col_a,col_b = st.columns([1,1])
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

# Dashboard tab
with tabs[1]:
    st.subheader("📊 Dashboard")
    df_live_view = recalc_payments_and_solde(_get_df_live_safe())
    if df_live_view is None or df_live_view.empty:
        st.info("Aucune donnée en mémoire.")
    else:
        cats = unique_nonempty(df_live_view["Categories"]) if "Categories" in df_live_view.columns else []
        subs = unique_nonempty(df_live_view["Sous-categorie"]) if "Sous-categorie" in df_live_view.columns else []
        f1,f2,f3 = st.columns([1,1,1])
        sel_cat = f1.selectbox("Catégorie", options=[""]+cats, index=0, key=skey("dash","cat"))
        sel_sub = f2.selectbox("Sous-catégorie", options=[""]+subs, index=0, key=skey("dash","sub"))
        view = df_live_view.copy()
        if sel_cat:
            view = view[view["Categories"].astype(str)==sel_cat]
        if sel_sub:
            view = view[view["Sous-catégorie"].astype(str)==sel_sub]
        view = recalc_payments_and_solde(view)
        montant_col = detect_montant_column(view) or "Montant honoraires (US $)"
        autres_col = detect_autres_column(view) or "Autres frais (US $)"
        acomptes_cols = detect_acompte_columns(view)
        total_honoraires = float(view.get(montant_col,0).apply(lambda x: _to_num(x)).sum())
        total_autres = float(view.get(autres_col,0).apply(lambda x: _to_num(x)).sum())
        total_acomptes = 0.0
        for ac in acomptes_cols:
            total_acomptes += float(view.get(ac,0).apply(lambda x:_to_num(x)).sum())
        cols_k = st.columns(3)
        cols_k[0].markdown(kpi_html("Dossiers", f"{len(view):,}"), unsafe_allow_html=True)
        cols_k[1].markdown(kpi_html("Montant honoraires", _fmt_money(total_honoraires)), unsafe_allow_html=True)
        cols_k[2].markdown(kpi_html("Solde total", _fmt_money(total_honoraires + total_autres - total_acomptes)), unsafe_allow_html=True)
        st.markdown("### Clients (aperçu)")
        try:
            display_df = view.copy()
            for mc in [montant_col,autres_col,"Payé","Solde"]:
                if mc in display_df.columns:
                    display_df[mc] = display_df[mc].apply(lambda x: _fmt_money(_to_num(x)))
            st.dataframe(display_df.reset_index(drop=True), use_container_width=True, height=360)
        except Exception:
            st.write("Impossible d'afficher le tableau.")

# Analyses tab
with tabs[2]:
    st.subheader("📈 Analyses")
    df_ = _get_df_live_safe()
    if isinstance(df_,pd.DataFrame) and not df_.empty and "Categories" in df_.columns:
        try:
            import plotly.express as px
            cnt = df_["Categories"].value_counts().reset_index()
            cnt.columns = ["Categorie","Nombre"]
            fig = px.pie(cnt, names="Categorie", values="Nombre", hole=0.4)
            st.plotly_chart(fig, use_container_width=True)
        except Exception:
            st.bar_chart(df_["Categories"].value_counts())

# Add tab
with tabs[3]:
    st.subheader("➕ Ajouter un nouveau client")
    df_live = _get_df_live_safe()
    next_dossier_num = get_next_dossier_numeric(df_live)
    next_dossier = str(next_dossier_num)
    next_id_client = make_id_client_datebased(df_live)
    st.markdown(f"**ID_Client (auto)**: {next_id_client}")
    st.markdown(f"**Dossier N (auto)**: {next_dossier}")

    add_date = st.date_input("Date (événement)", value=date.today(), key=skey("addtab","date"))
    add_nom = st.text_input("Nom du client", value="", key=skey("addtab","nom"))

    categories_options = visa_categories if visa_categories else unique_nonempty(df_live["Categories"]) if "Categories" in df_live.columns else []
    r3c1,r3c2,r3c3 = st.columns([1.2,1.6,1.6])
    with r3c1:
        add_cat = st.selectbox("Catégorie", options=[""]+categories_options, index=0, key=skey("addtab","cat"))
    with r3c2:
        add_sub_options=[]
        if isinstance(add_cat,str) and add_cat.strip():
            k = canonical_key(add_cat)
            if k in visa_map_norm:
                add_sub_options = visa_map_norm.get(k,[])[:]
            else:
                if add_cat in visa_map:
                    add_sub_options = visa_map.get(add_cat,[])[:]
        if not add_sub_options:
            try:
                add_sub_options = sorted({str(x).strip() for x in df_live["Sous-categorie"].dropna().astype(str).tolist()})
            except Exception:
                add_sub_options=[]
        add_sub = st.selectbox("Sous-catégorie", options=[""]+add_sub_options, index=0, key=skey("addtab","sub"))
    with r3c3:
        specific_opts = get_visa_options(add_cat, add_sub)
        if specific_opts:
            add_visa = st.selectbox("Visa (options)", options=[""]+specific_opts, index=0, key=skey("addtab","visa"))
        else:
            add_visa = st.text_input("Visa", value="", key=skey("addtab","visa"))

    r4c1,r4c2 = st.columns([1.4,1.0])
    with r4c1:
        add_montant = st.text_input("Montant honoraires (US $)", value="0", key=skey("addtab","montant"))
    with r4c2:
        a1 = st.text_input("Acompte 1", value="0", key=skey("addtab","ac1"))
    r5c1,r5c2 = st.columns([1.6,1.0])
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
            new_row = {c:"" for c in COLS_CLIENTS}
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
            for acc in ["Acompte 2","Acompte 3","Acompte 4"]:
                new_row[acc] = 0.0
            modes=[]
            if pay_cb: modes.append("CB")
            if pay_cheque: modes.append("Cheque")
            if pay_virement: modes.append("Virement")
            if pay_venmo: modes.append("Venmo")
            new_row["ModeReglement"] = ",".join(modes)
            new_row["ModeReglement_Ac1"] = ",".join(modes) if modes else ""
            new_row["Payé"] = new_row["Acompte 1"]
            new_row["Solde"] = new_row["Montant honoraires (US $)"] + new_row["Autres frais (US $)"] - new_row["Payé"]
            new_row["Solde à percevoir (US $)"] = new_row["Solde"]
            new_row["Commentaires"] = add_comments
            ensure_flag_columns(new_row, DEFAULT_FLAGS)
            for f in DEFAULT_FLAGS:
                new_row[f] = 0
            df_live = _get_df_live_safe()
            df_live = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            df_live = recalc_payments_and_solde(df_live)
            _set_df_live(df_live)
            st.success(f"Dossier ajouté : ID_Client {next_id_client} — Dossier N {next_dossier}")
        except Exception as e:
            st.error(f"Erreur ajout: {e}")

# Gestion tab
with tabs[4]:
    st.subheader("✏️ / 🗑️ Gestion — Modifier / Supprimer")
    df_live = _get_df_live_safe()
    for c in COLS_CLIENTS:
        if c not in df_live.columns:
            df_live[c] = pd.NaT if "Date" in c else (0.0 if c in NUMERIC_TARGETS else (0 if c in ["RFE","Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé"] else ""))
    if df_live is None or df_live.empty:
        st.info("Aucun dossier à modifier ou supprimer.")
    else:
        choices = [f"{i} | {df_live.at[i,'Dossier N']} | {df_live.at[i,'Nom']}" for i in range(len(df_live))]
        sel = st.selectbox("Sélectionner ligne à modifier", options=[""]+choices, key=skey("edit","select"))
        if sel:
            idx = int(sel.split("|")[0].strip())
            df_live = recalc_payments_and_solde(df_live)
            row = df_live.loc[idx].copy()
            def txt(v):
                if pd.isna(v): return ""
                s = str(v)
                if s.strip().lower() in ("nan","none","na","n/a"): return ""
                return s
            def _safe_row_date_local(col):
                try:
                    raw = row.get(col)
                except Exception:
                    raw = None
                return _date_or_none_safe(raw)
            def _parse_modes(raw):
                try:
                    if pd.isna(raw) or raw is None: return []
                    s = str(raw).strip()
                    if not s: return []
                    return [p.strip() for p in s.split(",") if p.strip()]
                except Exception:
                    return []
            row_modes_general = _parse_modes(row.get("ModeReglement",""))
            row_mode_ac1 = _parse_modes(row.get("ModeReglement_Ac1","")) or row_modes_general
            with st.form(key=skey("form_edit",str(idx))):
                c_name,c_solde = st.columns([2.5,1])
                with c_name:
                    st.markdown(f"### {txt(row.get('Nom',''))}")
                with c_solde:
                    try:
                        sol_num = _to_num(row.get("Solde à percevoir (US $)", row.get("Solde",0)))
                        st.markdown(f"**Solde dû**: {_fmt_money(sol_num)}")
                    except Exception:
                        st.markdown("**Solde dû**: $0.00")
                r1c1,r1c2,r1c3 = st.columns([1.4,1.0,1.2])
                with r1c1:
                    st.markdown(f"**ID_Client :** {txt(row.get('ID_Client',''))}")
                with r1c2:
                    e_dossier = st.text_input("Dossier N", value=txt(row.get("Dossier N","")), key=skey("edit","dossier",str(idx)))
                with r1c3:
                    e_date = st.date_input("Date (événement)", value=_safe_row_date_local("Date"), key=skey("edit","date",str(idx)))
                # amounts
                m1,m2,m3 = st.columns([1.2,1.0,1.0])
                with m1:
                    e_montant = st.text_input("Montant honoraires (US $)", value=txt(row.get("Montant honoraires (US $)","")), key=skey("edit","montant",str(idx)))
                with m2:
                    e_autres = st.text_input("Autres frais (US $)", value=txt(row.get("Autres frais (US $)","")), key=skey("edit","autres",str(idx)))
                with m3:
                    try:
                        total_val = _to_num(e_montant)+_to_num(e_autres)
                    except Exception:
                        total_val = _to_num(row.get("Montant honoraires (US $)",0))+_to_num(row.get("Autres frais (US $)",0))
                    st.text_input("Montant Total", value=str(total_val), key=skey("edit","montant_total",str(idx)), disabled=True)
                # acomptes
                r_ac_1,r_ac_2,r_ac_3,r_ac_4 = st.columns([1,1,1,1])
                with r_ac_1:
                    e_ac1 = st.text_input("Acompte 1", value=txt(row.get("Acompte 1","")), key=skey("edit","ac1",str(idx)))
                with r_ac_2:
                    e_ac2 = st.text_input("Acompte 2", value=txt(row.get("Acompte 2","")), key=skey("edit","ac2",str(idx)))
                with r_ac_3:
                    e_ac3 = st.text_input("Acompte 3", value=txt(row.get("Acompte 3","")), key=skey("edit","ac3",str(idx)))
                with r_ac_4:
                    e_ac4 = st.text_input("Acompte 4", value=txt(row.get("Acompte 4","")), key=skey("edit","ac4",str(idx)))
                # dates + modes
                r_d1,r_d2,r_d3,r_d4 = st.columns([1,1,1,1])
                with r_d1:
                    e_mode_ac1 = st.multiselect("Mode A1", options=["CB","Cheque","Virement","Venmo"], default=row_mode_ac1, key=skey("edit","mode_ac1",str(idx)))
                    e_ac1_date = st.date_input("Date Acompte 1", value=_safe_row_date_local("Date Acompte 1"), key=skey("edit","ac1_date",str(idx)))
                with r_d2:
                    e_mode_ac2 = st.multiselect("Mode A2", options=["CB","Cheque","Virement","Venmo"], default=_parse_modes(row.get("ModeReglement_Ac2","")), key=skey("edit","mode_ac2",str(idx)))
                    e_ac2_date = st.date_input("Date Acompte 2", value=_safe_row_date_local("Date Acompte 2"), key=skey("edit","ac2_date",str(idx)))
                with r_d3:
                    e_mode_ac3 = st.multiselect("Mode A3", options=["CB","Cheque","Virement","Venmo"], default=_parse_modes(row.get("ModeReglement_Ac3","")), key=skey("edit","mode_ac3",str(idx)))
                    e_ac3_date = st.date_input("Date Acompte 3", value=_safe_row_date_local("Date Acompte 3"), key=skey("edit","ac3_date",str(idx)))
                with r_d4:
                    e_mode_ac4 = st.multiselect("Mode A4", options=["CB","Cheque","Virement","Venmo"], default=_parse_modes(row.get("ModeReglement_Ac4","")), key=skey("edit","mode_ac4",str(idx)))
                    e_ac4_date = st.date_input("Date Acompte 4", value=_safe_row_date_local("Date Acompte 4"), key=skey("edit","ac4_date",str(idx)))
                # flags
                f1,f2,f3,f4 = st.columns([1,1,1,1])
                with f1:
                    e_flag_envoye = st.checkbox("Dossiers envoyé", value=bool(int(row.get("Dossiers envoyé",0))) if not pd.isna(row.get("Dossiers envoyé",0)) else False, key=skey("edit","flag_envoye",str(idx)))
                with f2:
                    e_flag_approuve = st.checkbox("Dossier approuvé", value=bool(int(row.get("Dossier approuvé",0))) if not pd.isna(row.get("Dossier approuvé",0)) else False, key=skey("edit","flag_approuve",str(idx)))
                with f3:
                    e_flag_refuse = st.checkbox("Dossier refusé", value=bool(int(row.get("Dossier refusé",0))) if not pd.isna(row.get("Dossier refusé",0)) else False, key=skey("edit","flag_refuse",str(idx)))
                with f4:
                    e_flag_annule = st.checkbox("Dossier Annulé", value=bool(int(row.get("Dossier Annulé",0))) if not pd.isna(row.get("Dossier Annulé",0)) else False, key=skey("edit","flag_annule",str(idx)))
                other_flag_set = any([e_flag_envoye,e_flag_approuve,e_flag_refuse,e_flag_annule])
                if not other_flag_set:
                    st.markdown("**RFE** (active uniquement si un des états est coché)")
                    e_flag_rfe = st.checkbox("RFE", value=bool(int(row.get("RFE",0))) if not pd.isna(row.get("RFE",0)) else False, key=skey("edit","flag_rfe",str(idx)), disabled=True)
                else:
                    e_flag_rfe = st.checkbox("RFE", value=bool(int(row.get("RFE",0))) if not pd.isna(row.get("RFE",0)) else False, key=skey("edit","flag_rfe",str(idx)))
                e_comments = st.text_area("Commentaires", value=txt(row.get("Commentaires","")), key=skey("edit","comments",str(idx)))
                save = st.form_submit_button("Enregistrer modifications")
                if save:
                    try:
                        df_live.at[idx,"Dossier N"] = e_dossier
                        df_live.at[idx,"Nom"] = txt(row.get("Nom",""))  # keep existing name if not changed via e_nom
                        df_live.at[idx,"Date"] = pd.to_datetime(e_date)
                        df_live.at[idx,"Montant honoraires (US $)"] = money_to_float(e_montant)
                        df_live.at[idx,"Autres frais (US $)"] = money_to_float(e_autres)
                        df_live.at[idx,"Acompte 1"] = money_to_float(e_ac1)
                        df_live.at[idx,"Acompte 2"] = money_to_float(e_ac2)
                        df_live.at[idx,"Acompte 3"] = money_to_float(e_ac3)
                        df_live.at[idx,"Acompte 4"] = money_to_float(e_ac4)
                        df_live.at[idx,"Date Acompte 1"] = pd.to_datetime(e_ac1_date) if e_ac1_date else pd.NaT
                        df_live.at[idx,"Date Acompte 2"] = pd.to_datetime(e_ac2_date) if e_ac2_date else pd.NaT
                        df_live.at[idx,"Date Acompte 3"] = pd.to_datetime(e_ac3_date) if e_ac3_date else pd.NaT
                        df_live.at[idx,"Date Acompte 4"] = pd.to_datetime(e_ac4_date) if e_ac4_date else pd.NaT
                        df_live.at[idx,"ModeReglement_Ac1"] = ",".join(e_mode_ac1) if isinstance(e_mode_ac1,(list,tuple)) else str(e_mode_ac1)
                        df_live.at[idx,"ModeReglement_Ac2"] = ",".join(e_mode_ac2) if isinstance(e_mode_ac2,(list,tuple)) else str(e_mode_ac2)
                        df_live.at[idx,"ModeReglement_Ac3"] = ",".join(e_mode_ac3) if isinstance(e_mode_ac3,(list,tuple)) else str(e_mode_ac3)
                        df_live.at[idx,"ModeReglement_Ac4"] = ",".join(e_mode_ac4) if isinstance(e_mode_ac4,(list,tuple)) else str(e_mode_ac4)
                        old_general = parse_modes_global(row.get("ModeReglement",""))
                        combined = set(old_general + list(e_mode_ac1))
                        df_live.at[idx,"ModeReglement"] = ",".join(sorted(list(combined)))
                        df_live.at[idx,"Dossiers envoyé"] = 1 if e_flag_envoye else 0
                        df_live.at[idx,"Dossier approuvé"] = 1 if e_flag_approuve else 0
                        df_live.at[idx,"Dossier refusé"] = 1 if e_flag_refuse else 0
                        df_live.at[idx,"Dossier Annulé"] = 1 if e_flag_annule else 0
                        if e_flag_rfe and not any([e_flag_envoye,e_flag_approuve,e_flag_refuse,e_flag_annule]):
                            st.warning("RFE n'a pas été activé car aucun état (envoyé/approuvé/refusé/annulé) n'est coché.")
                            df_live.at[idx,"RFE"] = 0
                        else:
                            df_live.at[idx,"RFE"] = 1 if e_flag_rfe else 0
                        df_live.at[idx,"Commentaires"] = e_comments
                        df_live = recalc_payments_and_solde(df_live)
                        df_live.at[idx,"Solde à percevoir (US $)"] = df_live.at[idx,"Solde"]
                        _set_df_live(df_live)
                        st.success("Modifications enregistrées.")
                    except Exception as e:
                        st.error(f"Erreur enregistrement: {e}")

    st.markdown("---")
    st.markdown("### Supprimer des dossiers")
    if df_live is None or df_live.empty:
        st.info("Aucun dossier à supprimer.")
    else:
        choices_del = [f"{i} | {df_live.at[i,'Dossier N']} | {df_live.at[i,'Nom']}" for i in range(len(df_live))]
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

# Export tab
with tabs[5]:
    st.header("💾 Export")
    df_live = _get_df_live_safe()
    if df_live is None or df_live.empty:
        st.info("Aucune donnée à exporter.")
    else:
        st.write(f"Vue en mémoire: {df_live.shape[0]} lignes, {df_live.shape[1]} colonnes")
        c1,c2 = st.columns(2)
        with c1:
            csv_bytes = df_live.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Export CSV", data=csv_bytes, file_name="Clients_export.csv", mime="text/csv")
        with c2:
            df_for_export = df_live.copy()
            try:
                montant_col = detect_montant_column(df_for_export) or "Montant honoraires (US $)"
                autres_col = detect_autres_column(df_for_export) or "Autres frais (US $)"
                acomptes_cols = detect_acompte_columns(df_for_export)
                df_for_export["_Montant_num_"] = df_for_export.get(montant_col,0).apply(lambda x:_to_num(x))
                df_for_export["_Autres_num_"] = df_for_export.get(autres_col,0).apply(lambda x:_to_num(x))
                for acc in acomptes_cols:
                    df_for_export[f"_num_{acc}"] = df_for_export.get(acc,0).apply(lambda x:_to_num(x))
                if acomptes_cols:
                    df_for_export["_Acomptes_sum_"] = df_for_export[[f"_num_{acc}" for acc in acomptes_cols]].sum(axis=1)
                else:
                    df_for_export["_Acomptes_sum_"] = 0.0
                df_for_export["Solde_formule"] = df_for_export["_Montant_num_"] + df_for_export["_Autres_num_"] - df_for_export["_Acomptes_sum_"]
                df_for_export["Solde à percevoir (US $)"] = df_for_export["Solde_formule"]
            except Exception:
                df_for_export["Solde_formule"] = df_for_export.get("Solde",0).apply(lambda x:_to_num(x))
                df_for_export["Solde à percevoir (US $)"] = df_for_export.get("Solde à percevoir (US $)",0).apply(lambda x:_to_num(x))
            drop_cols = [c for c in df_for_export.columns if c.startswith("_num_") or c in ["_Montant_num_","_Autres_num_","_Acomptes_sum_"]]
            try:
                df_export_final = df_for_export.drop(columns=drop_cols)
            except Exception:
                df_export_final = df_for_export.copy()
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df_export_final.to_excel(writer, index=False, sheet_name="Clients")
            out_bytes = buf.getvalue()
            st.download_button("⬇️ Export XLSX (avec colonne Solde_formule)", data=out_bytes, file_name="Clients_export_with_Solde_formule.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# End of file
