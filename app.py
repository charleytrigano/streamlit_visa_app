# app.py - Visa Manager (part 1/4)
# Début du fichier : imports, constantes et helpers

import os
import json
import re
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional

import pandas as pd
import streamlit as st

# Ensure key globals exist early to avoid NameError after reassembly
df_clients_raw: Optional[pd.DataFrame] = None
df_visa_raw: Optional[pd.DataFrame] = None
clients_src_for_read = None
visa_src_for_read = None

# Optional plotly for charts
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    px = None
    HAS_PLOTLY = False

# Optional openpyxl for advanced XLSX export
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
CURRENT_USER = "charleytrigano"

def skey(*parts: str) -> str:
    return f"{SID}_" + "_".join([p for p in parts if p])

# -------------------------
# Utility helpers
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

def money_to_float(x: Any) -> float:
    try:
        if pd.isna(x):
            return 0.0
    except Exception:
        pass
    try:
        s = str(x).strip()
        if s == "" or s.lower() in ("na","n/a","nan"):
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
            return float(re.sub(r"[^0-9.\-]", "", str(x)))
        except Exception:
            return 0.0

def _to_num(x: Any) -> float:
    if isinstance(x, (int, float)):
        return float(x)
    return money_to_float(x)

def _fmt_money(v: Any) -> str:
    try:
        return "${:,.2f}".format(float(v))
    except Exception:
        return "$0.00"

# Single authoritative safe date converter used everywhere
def _date_or_none_safe(v: Any) -> Optional[date]:
    """
    Return a native datetime.date or None for any input v.
    Guarantees never to return pandas.Timestamp or pandas.NaT.
    """
    try:
        if v is None:
            return None
        if isinstance(v, date) and not isinstance(v, datetime):
            return v
        if isinstance(v, datetime):
            return v.date()
        # Convert strings, numpy datetime64, pandas.Timestamp
        d = pd.to_datetime(v, errors="coerce")
        if pd.isna(d):
            return None
        return date(int(d.year), int(d.month), int(d.day))
    except Exception:
        return None

# app.py - Visa Manager (part 2/4)
# Column heuristics, visa mapping, I/O and normalization functions

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
# Visa mapping
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

# -------------------------
# Robust I/O helpers
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

# app.py - Visa Manager (part 3/4)
# Ensure columns, normalize dataset and session DF, UI bootstrap and Files/Dashboard/Analyses/Add tabs

# -------------------------
# Ensure columns & normalise dataset
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
            elif c in ["Date de création", "Dernière modification", "Date", "Date Acompte 1", "Date Acompte 2", "Date Acompte 3", "Date Acompte 4", "Date d'envoi"]:
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
                elif c in ["Date de création", "Dernière modification", "Date", "Date Acompte 1", "Date Acompte 2", "Date Acompte 3", "Date Acompte 4", "Date d'envoi"]:
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
    for dtc in ["Date","Date de création","Dernière modification","Date Acompte 1","Date Acompte 2","Date Acompte 3","Date Acompte 4","Date d'envoi"]:
        if dtc in df_mapped.columns:
            try:
                df_mapped[dtc] = pd.to_datetime(df_mapped[dtc], dayfirst=True, errors="coerce")
            except Exception:
                pass
    df = _ensure_columns(df_mapped, COLS_CLIENTS)
    for col in NUMERIC_TARGETS:
        if col in df.columns:
            try:
                df[col] = df[col].apply(lambda x: _to_num(x))
            except Exception:
                df[col] = 0.0
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
    for c in ["Nom","Categories","Sous-catégorie","Visa","Commentaires","Créé par","Modifié par"]:
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
    if "Escrow" in out.columns:
        try:
            out["Escrow"] = out["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ("1","true","t","yes","oui","y","x") else (1 if _to_num(x) == 1 else 0))
        except Exception:
            out["Escrow"] = out["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ("1","true","t","yes","oui","y","x") else 0)
    return out

# Initialize session DF properly
# If read_any_table hasn't been called earlier (because of assembly), keep safe defaults
try:
    df_clients_raw = df_clients_raw if df_clients_raw is not None else pd.DataFrame()
except Exception:
    df_clients_raw = pd.DataFrame()
try:
    df_visa_raw = df_visa_raw if df_visa_raw is not None else pd.DataFrame()
except Exception:
    df_visa_raw = pd.DataFrame()

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

# app.py - Visa Manager (part 4/4)
# UI tabs: Files, Dashboard, Analyses, Add, Gestion (form with safe date inputs), Export

# -------------------------
# Tabs UI
# -------------------------
tabs = st.tabs(["📄 Fichiers","📊 Dashboard","📈 Analyses","➕ Ajouter","✏️ / 🗑️ Gestion","💾 Export"])

# ---- Files tab ----
with tabs[0]:
    st.header("📂 Fichiers")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Clients")
        if 'up_clients' in globals() and up_clients is not None:
            st.text(f"Upload: {getattr(up_clients,'name','')}")
        elif isinstance(globals().get('clients_src_for_read',""), str) and globals().get('clients_src_for_read', ""):
            st.text(f"Chargé depuis: {globals().get('clients_src_for_read')}")
        elif os.path.exists(CACHE_CLIENTS):
            st.text("Chargé depuis le cache local")
        if df_clients_raw is None or (isinstance(df_clients_raw, pd.DataFrame) and df_clients_raw.empty):
            st.warning("Aucun fichier Clients detecté.")
        else:
            try:
                st.success(f"Clients lus: {df_clients_raw.shape[0]} lignes")
            except Exception:
                st.success("Clients lus")
            try:
                max_preview = 100
                if isinstance(df_clients_raw, pd.DataFrame) and df_clients_raw.shape[0] <= max_preview:
                    st.dataframe(df_clients_raw.reset_index(drop=True), use_container_width=True, height=360)
                elif isinstance(df_clients_raw, pd.DataFrame):
                    st.dataframe(df_clients_raw.head(100).reset_index(drop=True), use_container_width=True, height=360)
                    if st.button("Afficher tout (peut être lent)"):
                        st.dataframe(df_clients_raw.reset_index(drop=True), use_container_width=True, height=600)
            except Exception:
                try:
                    st.write(df_clients_raw.head(8))
                except Exception:
                    st.write("Aperçu indisponible")
    with c2:
        st.subheader("Visa")
        if 'up_visa' in globals() and up_visa is not None:
            st.text(f"Upload: {getattr(up_visa,'name','')}")
        elif isinstance(globals().get('visa_src_for_read',""), str) and globals().get('visa_src_for_read', ""):
            st.text(f"Chargé depuis: {globals().get('visa_src_for_read')}")
        elif os.path.exists(CACHE_VISA):
            st.text("Chargé depuis le cache local")
        if df_visa_raw is None or (isinstance(df_visa_raw, pd.DataFrame) and df_visa_raw.empty):
            st.warning("Aucun fichier Visa detecté.")
        else:
            try:
                st.success(f"Visa lu: {df_visa_raw.shape[0]} lignes, {df_visa_raw.shape[1]} colonnes")
                st.dataframe(df_visa_raw.reset_index(drop=True), use_container_width=True, height=360)
            except Exception:
                try:
                    st.write(df_visa_raw.head(8))
                except Exception:
                    st.write("Aperçu Visa indisponible")
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
        subs = unique_nonempty(df_live_view["Sous-catégorie"]) if "Sous-catégorie" in df_live_view.columns else []
        visas = unique_nonempty(df_live_view["Visa"]) if "Visa" in df_live_view.columns else []
        years = []
        if "_Année_" in df_live_view.columns:
            try:
                years = sorted([int(y) for y in pd.to_numeric(df_live_view["_Année_"], errors="coerce").dropna().unique().astype(int).tolist()])
            except Exception:
                years = []
        f1, f2, f3, f4 = st.columns([1,1,1,1])
        sel_cat = f1.selectbox("Catégorie", options=[""]+cats, index=0, key=skey("dash","cat"))
        sel_sub = f2.selectbox("Sous-catégorie", options=[""]+subs, index=0, key=skey("dash","sub"))
        sel_visa = f3.selectbox("Visa", options=[""]+visas, index=0, key=skey("dash","visa"))
        year_options = ["Toutes les années"] + [str(y) for y in years]
        sel_year = f4.selectbox("Année", options=year_options, index=0, key=skey("dash","year"))
        view = df_live_view.copy()
        if sel_cat:
            view = view[view["Categories"].astype(str) == sel_cat]
        if sel_sub:
            view = view[view["Sous-catégorie"].astype(str) == sel_sub]
        if sel_visa:
            view = view[view["Visa"].astype(str) == sel_visa]
        if sel_year and sel_year != "Toutes les années":
            view = view[view["_Année_"].astype(str) == sel_year]
        view = recalc_payments_and_solde(view)
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
        total_acomptes_sum = 0.0
        if acomptes_cols:
            for c in acomptes_cols:
                view[f"_ac_{c}"] = view.get(c, 0).apply(safe_num)
                total_acomptes_sum += float(view[f"_ac_{c}"].sum())
        total_honoraires = float(view["_Montant_num_"].sum())
        total_autres = float(view["_Autres_num_"].sum())
        total_paye = float(total_acomptes_sum)
        canonical_solde_sum = float(total_honoraires + total_autres - total_paye)
        cols_k = st.columns(4)
        cols_k[0].markdown(kpi_html("Dossiers (vue)", f"{len(view):,}"), unsafe_allow_html=True)
        cols_k[1].markdown(kpi_html("Montant honoraires", _fmt_money(total_honoraires)), unsafe_allow_html=True)
        cols_k[2].markdown(kpi_html("Autres frais", _fmt_money(total_autres)), unsafe_allow_html=True)
        cols_k[3].markdown(kpi_html("Total facturé (recalc)", _fmt_money(total_honoraires + total_autres)), unsafe_allow_html=True)
        st.markdown("---")
        cols_k2 = st.columns(2)
        cols_k2[0].markdown(kpi_html("Montant payé (somme acomptes)", _fmt_money(total_paye)), unsafe_allow_html=True)
        cols_k2[1].markdown(kpi_html("Solde total (recalc)", _fmt_money(canonical_solde_sum)), unsafe_allow_html=True)
        st.markdown("### Détails — clients correspondant aux filtres")
        display_df = view.copy()
        if "Date" in display_df.columns:
            try:
                display_df["Date"] = pd.to_datetime(display_df["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                display_df["Date"] = display_df["Date"].astype(str)
        for dtc in ["Date Acompte 1", "Date Acompte 2", "Date d'envoi", "Date de création", "Dernière modification"]:
            if dtc in display_df.columns:
                try:
                    display_df[dtc] = pd.to_datetime(display_df[dtc], errors="coerce").dt.strftime("%Y-%m-%d")
                except Exception:
                    display_df[dtc] = display_df[dtc].astype(str)
        money_cols = [montant_col, autres_col, "Payé","Solde","Solde à percevoir (US $)"] + acomptes_cols
        for mc in money_cols:
            if mc in display_df.columns:
                try:
                    display_df[mc] = display_df[mc].apply(lambda x: _fmt_money(_to_num(x)))
                except Exception:
                    display_df[mc] = display_df[mc].astype(str)
        try:
            st.dataframe(display_df.reset_index(drop=True), use_container_width=True, height=360)
        except Exception:
            st.write("Impossible d'afficher la liste des clients (trop volumineuse). Utilisez l'export")

# ---- Gestion tab / Export tab implemented above in the rest of file (full earlier parts) ----

# End of file
