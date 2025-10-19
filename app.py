# Visa Manager - app.py
# Streamlit app with robust reading/normalization, Dashboard, Analyses and Management (Add/Edit/Delete).
# Fix: category select is reactive and updates sub-category options using a normalized visa_map (visa_map_norm).
#
# Usage: streamlit run app.py
# Requirements: pandas, openpyxl; optional: plotly

import os
import json
import re
import io
from io import BytesIO
from datetime import date, datetime, timedelta
from typing import Tuple, Dict, Any, List, Optional

import pandas as pd
import numpy as np
import streamlit as st

# Try plotly
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    px = None
    HAS_PLOTLY = False

# Config
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

# --- plotting helpers (unchanged) ---
def plot_pie(df_counts, names_col: str = "Categorie", value_col: str = "Nombre", title: str = ""):
    if HAS_PLOTLY and px is not None:
        fig = px.pie(df_counts, names=names_col, values=value_col, hole=0.45, title=title)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.write(title)
        if value_col in df_counts.columns and names_col in df_counts.columns:
            st.bar_chart(df_counts.set_index(names_col)[value_col])

def plot_barh(df_bar, x: str, y: str, title: str = ""):
    if HAS_PLOTLY and px is not None:
        fig = px.bar(df_bar, x=x, y=y, orientation="h", title=title, text=x)
        fig.update_layout(yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.write(title)
        if x in df_bar.columns and y in df_bar.columns:
            st.bar_chart(df_bar.set_index(y)[x])

def plot_line(df_line, x: str, y: str, title: str = "", x_title: str = "", y_title: str = ""):
    if HAS_PLOTLY and px is not None:
        fig = px.line(df_line, x=x, y=y, markers=True, title=title)
        fig.update_layout(xaxis_title=x_title, yaxis_title=y_title)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.write(title)
        if x in df_line.columns and y in df_line.columns:
            st.line_chart(df_line.set_index(x)[y])

# --- normalization helpers (unchanged) ---
def normalize_header_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r'^\s+|\s+$', '', s)
    s = re.sub(r"\s+", " ", s)
    return s

def canonical_key(s: str) -> str:
    if s is None:
        return ""
    s2 = normalize_header_text(s).lower()
    s2 = s2.replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("ç", "c").replace("ô","o").replace("ù","u")
    s2 = re.sub(r"[^a-z0-9 ]", "", s2)
    s2 = s2.replace("_", " ").replace("-", " ").strip()
    s2 = re.sub(r"\s+", " ", s2)
    return s2

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
    "dossier envoye": "Dossiers envoyé", "dossier approuve": "Dossier approuvé", "dossier refuse": "Dossier refusé",
    "rfe": "RFE", "commentaires": "Commentaires"
}

NUMERIC_TARGETS = [
    "Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde",
    "Acompte 1", "Acompte 2", "Acompte 3", "Acompte 4"
]

def map_columns_heuristic(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str,str]]:
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
    df = df.rename(columns=new_names)
    return df, new_names

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
        s2 = re.sub(r"[^0-9.\-]", "", s)
        try:
            return float(s2)
        except Exception:
            return 0.0

# Visa mapping builder
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

# Basic features/helpers
def _to_num(x: Any) -> float:
    return money_to_float(x) if not isinstance(x, (int, float)) else float(x)

def _fmt_money(v: float) -> str:
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

def _ensure_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            if c in ["Payé", "Solde", "Montant honoraires (US $)", "Autres frais (US $)", "Acompte 1", "Acompte 2"]:
                out[c] = 0.0
            elif c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
                out[c] = 0
            else:
                out[c] = ""
    return out[cols]

def _normalize_clients_numeric(df: pd.DataFrame) -> pd.DataFrame:
    for c in NUMERIC_TARGETS:
        if c in df.columns:
            df[c] = df[c].apply(_to_num)
    if "Montant honoraires (US $)" in df.columns and "Autres frais (US $)" in df.columns:
        total = df["Montant honoraires (US $)"].fillna(0) + df["Autres frais (US $)"].fillna(0)
        if "Payé" in df.columns:
            df["Solde"] = (total - df["Payé"].fillna(0))
        else:
            df["Solde"] = total
    return df

def _normalize_status(df: pd.DataFrame) -> pd.DataFrame:
    for c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: 1 if str(x).strip().lower() in ["1", "true", "oui", "o", "x", "yes"] else 0)
        else:
            df[c] = 0
    return df

def normalize_clients(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS_CLIENTS)
    df = df.copy()
    df, _mapping = map_columns_heuristic(df)
    if "Date" in df.columns:
        try:
            df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce")
        except Exception:
            pass
    df = _ensure_columns(df, COLS_CLIENTS)
    df = _normalize_clients_numeric(df)
    df = _normalize_status(df)
    for c in ["Nom", "Categories", "Sous-categorie", "Visa", "Commentaires"]:
        if c in df.columns:
            df[c] = df[c].astype(str).fillna("")
    try:
        dser = pd.to_datetime(df["Date"], errors="coerce")
        df["_Année_"] = dser.dt.year.fillna(0).astype(int)
        df["_MoisNum_"] = dser.dt.month.fillna(0).astype(int)
        df["Mois"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) and m>0 else "")
    except Exception:
        df["_Année_"] = 0
        df["_MoisNum_"] = 0
        df["Mois"] = ""
    return df

def _ensure_time_features(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()
    if "Date" in df.columns:
        try:
            dd = pd.to_datetime(df["Date"], errors="coerce")
        except Exception:
            dd = pd.to_datetime(pd.Series([], dtype="datetime64[ns]"))
        df["_Année_"] = dd.dt.year
        df["_MoisNum_"] = dd.dt.month
        df["Mois"] = dd.dt.month.apply(lambda m: f"{int(m):02d}" if pd.notna(m) else "")
    else:
        if "_Année_" not in df.columns:
            df["_Année_"] = pd.NA
        if "_MoisNum_" not in df.columns:
            df["_MoisNum_"] = pd.NA
        if "Mois" not in df.columns:
            df["Mois"] = ""
    return df

def safe_rerun():
    try:
        rerun_fn = getattr(st, "experimental_rerun", None)
        if callable(rerun_fn):
            rerun_fn()
            return
        rerun_fn2 = getattr(st, "rerun", None)
        if callable(rerun_fn2):
            rerun_fn2()
            return
    except Exception:
        pass
    try:
        st.sidebar.info("Rerun non disponible ; session_state mis à jour.")
        st.session_state.setdefault("_need_rerun", True)
    except Exception:
        pass

# Robust readers (same as before)
def try_read_excel_from_bytes(b: bytes, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
    bio = BytesIO(b)
    try:
        xls = pd.ExcelFile(bio, engine="openpyxl")
        sheets = xls.sheet_names
        try:
            st.sidebar.info(f"Excel file detected; sheets: {sheets}")
        except Exception:
            pass
        candidates: List[str] = []
        if sheet_name and sheet_name in sheets:
            candidates.append(sheet_name)
        for cand in [SHEET_CLIENTS, SHEET_VISA, "Sheet1"]:
            if cand in sheets and cand not in candidates:
                candidates.append(cand)
        for s in sheets:
            if s not in candidates:
                candidates.append(s)
        best_df: Optional[pd.DataFrame] = None
        best_non_null = -1
        HEADER_SCAN_ROWS = 8
        for cand in candidates:
            try:
                bio2 = BytesIO(b)
                df_raw = pd.read_excel(bio2, sheet_name=cand, header=None, engine="openpyxl")
                if df_raw is None:
                    continue
                topn = min(HEADER_SCAN_ROWS, len(df_raw))
                row_non_null_counts = [(i, df_raw.iloc[i].count()) for i in range(topn)]
                if row_non_null_counts:
                    best_row_idx, max_non_null = max(row_non_null_counts, key=lambda x: x[1])
                else:
                    best_row_idx, max_non_null = (0, 0)
                header_row = best_row_idx if max_non_null >= 2 else (0 if len(df_raw) > 0 else None)
                try:
                    bio3 = BytesIO(b)
                    if header_row is not None:
                        df_try = pd.read_excel(bio3, sheet_name=cand, header=header_row, engine="openpyxl")
                    else:
                        df_try = pd.read_excel(bio3, sheet_name=cand, engine="openpyxl")
                except Exception:
                    bio3 = BytesIO(b)
                    df_try = pd.read_excel(bio3, sheet_name=cand, header=None, engine="openpyxl")
                    for i in range(min(HEADER_SCAN_ROWS, len(df_try))):
                        if df_try.iloc[i].count() >= 2:
                            cols = df_try.iloc[i].astype(str).fillna("").tolist()
                            df_try = df_try.iloc[i+1:].copy()
                            df_try.columns = cols
                            break
                if df_try is None:
                    continue
                non_null_rows = df_try.dropna(how="all").shape[0]
                if non_null_rows > 0:
                    try:
                        st.sidebar.info(f"Selected sheet '{cand}' with {non_null_rows} data rows (header_row={header_row}).")
                    except Exception:
                        pass
                    return df_try
                if non_null_rows > best_non_null:
                    best_non_null = non_null_rows
                    best_df = df_try
            except Exception:
                continue
        return best_df
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
    if isinstance(src, (bytes, bytearray)):
        _log("read_any_table: src is raw bytes")
        df = try_read_excel_from_bytes(bytes(src), sheet)
        if df is not None:
            return df
        try:
            return pd.read_csv(BytesIO(src), sep=";", encoding="utf-8")
        except Exception:
            return None
    if isinstance(src, (io.BytesIO, BytesIO)):
        try:
            b = src.getvalue()
        except Exception:
            try:
                src.seek(0); b = src.read()
            except Exception:
                b = None
        if not b:
            _log("BytesIO: no data")
            return None
        df = try_read_excel_from_bytes(b, sheet)
        if df is not None:
            return df
        try:
            return pd.read_csv(BytesIO(b), sep=";", encoding="utf-8")
        except Exception:
            return None
    if hasattr(src, "read") and hasattr(src, "name"):
        name = getattr(src, "name", "")
        _log(f"Uploaded file name: {name}")
        data = None
        try:
            data = src.getvalue()
        except Exception:
            try:
                src.seek(0); data = src.read()
            except Exception:
                data = None
        if not data:
            _log("Uploaded file: no bytes extracted")
            return None
        lname = name.lower()
        if lname.endswith(".csv"):
            try:
                return pd.read_csv(BytesIO(data), sep=";", encoding="utf-8", on_bad_lines="skip")
            except Exception:
                try:
                    return pd.read_csv(BytesIO(data), sep=";", encoding="latin1", on_bad_lines="skip")
                except Exception:
                    return None
        df = try_read_excel_from_bytes(data, sheet)
        if df is not None:
            return df
        try:
            return pd.read_csv(BytesIO(data), sep=";", encoding="utf-8", on_bad_lines="skip")
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
                try:
                    return pd.read_csv(p, sep=";", encoding="latin1", on_bad_lines="skip")
                except Exception:
                    return None
        try:
            xls = pd.ExcelFile(p)
            sheets = xls.sheet_names
            _log(f"Excel path sheets: {sheets}")
            if sheet and sheet in sheets:
                return pd.read_excel(p, sheet_name=sheet, engine="openpyxl")
            for candidate in [SHEET_CLIENTS, SHEET_VISA, "Sheet1"]:
                if candidate in sheets:
                    return pd.read_excel(p, sheet_name=candidate, engine="openpyxl")
            return pd.read_excel(p, sheet_name=0, engine="openpyxl")
        except Exception:
            try:
                return pd.read_csv(p, sep=";", encoding="utf-8", on_bad_lines="skip")
            except Exception:
                return None
    _log("read_any_table: unsupported src type")
    return None

# --- simple debug preview helper ---
def debug_show_columns_preview(df, name="Data"):
    try:
        cols = list(df.columns) if isinstance(df, pd.DataFrame) else []
        st.sidebar.markdown(f"**DEBUG — colonnes {name} ({len(cols)}) :**")
        if cols:
            st.sidebar.write(cols)
            for c in cols[:8]:
                try:
                    vals = df[c].dropna().astype(str).unique()[:6].tolist()
                    st.sidebar.write(f"- {c}: {vals}")
                except Exception:
                    pass
        else:
            st.sidebar.write("Aucune colonne détectée.")
    except Exception as e:
        st.sidebar.write(f"DEBUG erreur: {e}")

# next client id
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

# -------------------------
# UI start
# -------------------------
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title(APP_TITLE)

# Sidebar file controls
st.sidebar.header("📂 Fichiers")
last_clients, last_visa, last_save_dir = ("", "", "")
try:
    if os.path.exists(MEMO_FILE):
        with open(MEMO_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
            last_clients, last_visa, last_save_dir = d.get("clients", ""), d.get("visa", ""), d.get("save_dir", "")
except Exception:
    last_clients, last_visa, last_save_dir = ("", "", "")

mode = st.sidebar.radio("Mode de chargement", ["Un fichier (Clients)", "Deux fichiers (Clients & Visa)"], index=0, key=skey("mode"))
up_clients = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx", "xls", "csv"], key=skey("up_clients"))
up_visa = None
if mode == "Deux fichiers (Clients & Visa)":
    up_visa = st.sidebar.file_uploader("Visa (xlsx/csv)", type=["xlsx", "xls", "csv"], key=skey("up_visa"))

clients_path_in = st.sidebar.text_input("ou chemin local Clients (laisser vide si upload)", value=last_clients, key=skey("cli_path"))
visa_path_in = st.sidebar.text_input("ou chemin local Visa (laisser vide si upload)", value=(last_visa if mode != "Un fichier (Clients)" else ""), key=skey("vis_path"))
save_dir_in = st.sidebar.text_input("Dossier de sauvegarde (optionnel)", value=last_save_dir, key=skey("save_dir"))

if st.sidebar.button("📥 Sauvegarder chemins", key=skey("btn_load")):
    try:
        with open(MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump({"clients": clients_path_in or "", "visa": visa_path_in or "", "save_dir": save_dir_in or ""}, f, ensure_ascii=False, indent=2)
        st.sidebar.success("Chemins mémorisés.")
    except Exception:
        st.sidebar.error("Impossible de sauvegarder les chemins.")
    safe_rerun()

# read uploaded files into bytes
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

# sources
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

# Read tables
df_clients_raw = None
df_visa_raw = None
try:
    df_clients_raw = read_any_table(clients_src_for_read, sheet=SHEET_CLIENTS, debug_prefix="[Clients] ")
except Exception as e:
    st.sidebar.error(f"[Clients] Exception primary read: {e}")
if df_clients_raw is None:
    try:
        df_clients_raw = robust_read_clients(clients_src_for_read)
    except Exception:
        pass

try:
    df_visa_raw = read_any_table(visa_src_for_read, sheet=SHEET_VISA, debug_prefix="[Visa] ")
except Exception as e:
    st.sidebar.error(f"[Visa] Exception primary read: {e}")
if df_visa_raw is None:
    try:
        df_visa_raw = read_any_table(visa_src_for_read, sheet=None, debug_prefix="[Visa fallback] ")
    except Exception:
        pass

if df_visa_raw is None:
    df_visa_raw = pd.DataFrame()
if df_clients_raw is None:
    df_clients_raw = pd.DataFrame()

# Debug raw
if isinstance(df_clients_raw, pd.DataFrame):
    debug_show_columns_preview(df_clients_raw, "Clients (raw)")
if isinstance(df_visa_raw, pd.DataFrame):
    debug_show_columns_preview(df_visa_raw, "Visa (raw)")

# Build visa_map and normalized map
visa_map = {}
visa_map_norm = {}
visa_categories = []
if isinstance(df_visa_raw, pd.DataFrame) and not df_visa_raw.empty:
    try:
        df_visa_raw, _vm = map_columns_heuristic(df_visa_raw)
        df_visa_raw = coerce_category_columns(df_visa_raw) if 'coerce_category_columns' in globals() else df_visa_raw
        raw_vm = build_visa_map(df_visa_raw)
        visa_map = {str(k).strip(): [str(s).strip() for s in v if str(s).strip()] for k, v in raw_vm.items() if str(k).strip()}
        visa_map_norm = {k.strip().lower(): [s.strip() for s in v] for k, v in visa_map.items()}
        visa_categories = sorted(list(visa_map.keys()))
    except Exception:
        visa_map = {}
        visa_map_norm = {}
        visa_categories = []
else:
    visa_map = {}
    visa_map_norm = {}
    visa_categories = []

# show normalized visa_map in sidebar for debugging
st.sidebar.markdown("DEBUG visa_map_norm:")
try:
    st.sidebar.write(visa_map_norm)
except Exception:
    pass

# Normalize clients and persist working df_all
if isinstance(df_clients_raw, pd.DataFrame) and not df_clients_raw.empty:
    try:
        df_clients_raw, _map = map_columns_heuristic(df_clients_raw)
        for col in NUMERIC_TARGETS:
            if col in df_clients_raw.columns:
                df_clients_raw[col] = df_clients_raw[col].apply(money_to_float)
        if "Date" in df_clients_raw.columns:
            try:
                df_clients_raw["Date"] = pd.to_datetime(df_clients_raw["Date"], dayfirst=True, errors="coerce")
            except Exception:
                pass
        st.sidebar.success("Clients : colonnes normalisées automatiquement.")
    except Exception as e:
        st.sidebar.error(f"Erreur normalisation Clients: {e}")

try:
    df_all = _ensure_time_features(normalize_clients(df_clients_raw))
    # coerce category columns if function exists
    if 'coerce_category_columns' in globals():
        try:
            df_all = coerce_category_columns(df_all)
        except Exception:
            pass
    debug_show_columns_preview(df_all, "Clients (normalized)")
except Exception as e:
    st.sidebar.error(f"Erreur préparation df_all: {e}")
    df_all = pd.DataFrame(columns=COLS_CLIENTS)

# session state for live DF
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

# Tabs
tabs = st.tabs(["📄 Fichiers","📊 Dashboard","📈 Analyses","➕ / ✏️ / 🗑️ Gestion","💾 Export"])

# Fichiers tab
with tabs[0]:
    st.header("📂 Fichiers")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Clients")
        if up_clients is not None:
            st.write("Upload:", up_clients.name)
            try: st.write(f"Taille: {len(clients_bytes)} bytes")
            except Exception: pass
        elif isinstance(clients_src_for_read, str) and clients_src_for_read:
            st.write("Chargé depuis chemin local:", clients_src_for_read)
        else:
            st.info("Aucun fichier Clients sélectionné.")
        if df_clients_raw is None or df_clients_raw.empty:
            st.warning("Lecture Clients : aucun tableau trouvé ou DataFrame vide.")
        else:
            st.success(f"Clients lus ({df_clients_raw.shape[0]} lignes, {df_clients_raw.shape[1]} colonnes)")
            try: st.dataframe(df_clients_raw.head(8), use_container_width=True, height=220)
            except Exception: st.write("Aperçu indisponible.")
    with c2:
        st.subheader("Visa")
        if mode == "Deux fichiers (Clients & Visa)":
            if up_visa is not None:
                st.write("Upload:", up_visa.name)
                try: st.write(f"Taille: {len(visa_bytes)} bytes")
                except Exception: pass
            elif isinstance(visa_src_for_read, str) and visa_src_for_read:
                st.write("Chargé depuis chemin local:", visa_src_for_read)
            else:
                st.info("Aucun fichier Visa sélectionné.")
        else:
            st.write("Mode 'Un fichier' : Visa sera lu depuis le même fichier Clients si présent.")
        if df_visa_raw is None or df_visa_raw.empty:
            st.warning("Lecture Visa : aucun tableau trouvé ou DataFrame vide.")
        else:
            st.success(f"Visa lu ({df_visa_raw.shape[0]} lignes, {df_visa_raw.shape[1]} colonnes)")
            try: st.dataframe(df_visa_raw.head(8), use_container_width=True, height=220)
            except Exception: st.write("Aperçu Visa indisponible.")
    st.markdown("---")
    a1, a2 = st.columns([1,1])
    with a1:
        if st.button("Réinitialiser la mémoire (recharger depuis fichiers)"):
            df_all2 = _ensure_time_features(normalize_clients(df_clients_raw))
            if 'coerce_category_columns' in globals():
                df_all2 = coerce_category_columns(df_all2)
            _set_df_live(df_all2)
            st.success("Mémoire réinitialisée.")
            safe_rerun()
    with a2:
        if st.button("Actualiser la lecture"):
            safe_rerun()

# Dashboard tab (compact KPIs)
with tabs[1]:
    st.subheader("📊 Dashboard")
    df_all_current = _get_df_live()
    if df_all_current is None or df_all_current.empty:
        st.info("Aucune donnée cliente en mémoire.")
    else:
        cats = sorted(df_all_current["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all_current.columns else []
        subs = sorted(df_all_current["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all_current.columns else []
        visas = sorted(df_all_current["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all_current.columns else []
        years = sorted(pd.to_numeric(df_all_current["_Année_"], errors="coerce").dropna().astype(int).unique().tolist()) if "_Année_" in df_all_current.columns else []

        f1, f2, f3, f4 = st.columns([1,1,1,1])
        sel_cat = f1.multiselect("Catégories", cats, default=[], key=skey("dash","cats"))
        sel_sub = f2.multiselect("Sous-catégories", subs, default=[], key=skey("dash","subs"))
        sel_visa = f3.multiselect("Visa", visas, default=[], key=skey("dash","visas"))
        sel_year = f4.multiselect("Année", years, default=[], key=skey("dash","years"))

        view = df_all_current.copy()
        if sel_cat: view = view[view["Categories"].astype(str).isin(sel_cat)]
        if sel_sub: view = view[view["Sous-categorie"].astype(str).isin(sel_sub)]
        if sel_visa: view = view[view["Visa"].astype(str).isin(sel_visa)]
        if sel_year: view = view[view["_Année_"].isin(sel_year)]

        total = (view.get("Montant honoraires (US $)", 0).apply(_to_num) + view.get("Autres frais (US $)", 0).apply(_to_num)).sum()
        paye = view.get("Payé", 0).apply(_to_num).sum() if "Payé" in view.columns else 0.0
        solde = view.get("Solde", 0).apply(_to_num).sum() if "Solde" in view.columns else 0.0
        avg = (total / len(view)) if len(view) else 0.0
        n_dossiers = len(view)

        kcols = st.columns([1,1,1,1])
        def small_metric(col, label, value):
            with col:
                st.markdown(f"<div style='font-size:14px; font-weight:600'>{label}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='font-size:16px; color:#0A6EBD; font-weight:700'>{value}</div>", unsafe_allow_html=True)
        small_metric(kcols[0], "Dossiers", f"{n_dossiers:,}")
        small_metric(kcols[1], "Total facturé", _fmt_money(total))
        small_metric(kcols[2], "Total reçu", _fmt_money(paye))
        small_metric(kcols[3], "Solde total", _fmt_money(solde))
        st.markdown("---")
        st.subheader("Aperçu récent des dossiers")
        recent = view.sort_values(by=["_Année_", "_MoisNum_"], ascending=[False, False]).head(20).copy()
        display_cols = [c for c in ["Dossier N","ID_Client","Nom","Date","Categories","Sous-categorie","Visa","Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"] if c in recent.columns]
        for col in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]:
            if col in recent.columns:
                recent[col] = recent[col].apply(lambda x: _fmt_money(_to_num(x)))
        if "Date" in recent.columns:
            try:
                recent["Date"] = pd.to_datetime(recent["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                recent["Date"] = recent["Date"].astype(str)
        st.dataframe(recent[display_cols].reset_index(drop=True), use_container_width=True)

# Analyses tab
with tabs[2]:
    st.subheader("📈 Analyses")
    df_all_current = _get_df_live()
    if df_all_current is None or df_all_current.empty:
        st.info("Aucune donnée pour les analyses.")
    else:
        st.markdown("### Répartition par Catégorie")
        if "Categories" in df_all_current.columns:
            cat_counts = df_all_current["Categories"].value_counts().rename_axis("Categorie").reset_index(name="Nombre")
            if not cat_counts.empty:
                plot_pie(cat_counts, names_col="Categorie", value_col="Nombre", title="Catégories")
            else:
                st.info("Pas de catégories.")
        else:
            st.info("Colonne 'Categories' introuvable.")
        st.markdown("---")
        st.markdown("### Évolution Mensuelle (Total US)")
        tmp = df_all_current.copy()
        if "Montant honoraires (US $)" in tmp.columns and "Autres frais (US $)" in tmp.columns:
            tmp["Total_US"] = tmp["Montant honoraires (US $)"].apply(_to_num) + tmp["Autres frais (US $)"].apply(_to_num)
        else:
            tmp["Total_US"] = 0.0
        if "_Année_" in tmp.columns and "Mois" in tmp.columns:
            tmp = tmp.dropna(subset=["_Année_","Mois"])
            if not tmp.empty:
                tmp["YearMonth"] = tmp["_Année_"].astype(int).astype(str) + "-" + tmp["Mois"].astype(str)
                g = tmp.groupby("YearMonth", as_index=False)["Total_US"].sum().sort_values("YearMonth")
                if not g.empty:
                    plot_line(g, x="YearMonth", y="Total_US", title="Évolution Mensuelle", x_title="Période", y_title="Montant (US$)")
                else:
                    st.info("Pas assez de données.")
            else:
                st.info("Pas de données temporelles.")
        else:
            st.info("Colonnes temporelles manquantes.")

# Gestion tab (Add/Edit/Delete) - IMPORTANT: category select is outside the form for reactivity
with tabs[3]:
    st.subheader("➕ / ✏️ / 🗑️ Gestion")
    df_live = _get_df_live()
    if df_live is None:
        df_live = pd.DataFrame(columns=COLS_CLIENTS)
    for c in COLS_CLIENTS:
        if c not in df_live.columns:
            df_live[c] = "" if c not in NUMERIC_TARGETS else 0.0

    # categories source (trimmed)
    categories_options = visa_categories if visa_categories else sorted({str(x).strip() for x in df_live["Categories"].dropna().astype(str).tolist()})
    # Category selector (outside form) - reactive
    st.markdown("### Ajouter un dossier")
    st.write("Sélectionnez la catégorie (déclenche la liste des sous-catégories) :")
    categories_options_local = [""] + [c.strip() for c in categories_options]
    add_cat_sel = st.selectbox("Categories (réactif)", options=categories_options_local, index=0, key=skey("add","cat_sel"))

    # compute sub options from normalized visa_map_norm (keys lower)
    add_sub_options = []
    if isinstance(add_cat_sel, str) and add_cat_sel.strip():
        add_sub_options = visa_map_norm.get(add_cat_sel.strip().lower(), [])
    if not add_sub_options:
        add_sub_options = sorted({str(x).strip() for x in df_live["Sous-categorie"].dropna().astype(str).tolist()})
    # debug show selected category & computed suboptions
    st.sidebar.write("DEBUG selected add_cat:", repr(add_cat_sel))
    st.sidebar.write("DEBUG computed add_sub_options:", add_sub_options)

    # Now the form (uses the previously selected category)
    with st.form(key=skey("form_add")):
        col_a1, col_a2, col_a3 = st.columns(3)
        with col_a1:
            next_id = get_next_client_id(df_live)
            st.markdown(f"**ID_Client (automatique)**: {next_id}")
            add_dossier = st.text_input("Dossier N", value="", key=skey("add","dossier"))
            add_nom = st.text_input("Nom", value="", key=skey("add","nom"))
        with col_a2:
            add_date = st.date_input("Date", value=date.today(), key=skey("add","date"))
            # category is taken from add_cat_sel
            st.markdown(f"Category choisie: **{add_cat_sel}**")
            # subcategory select uses computed options
            add_sub = st.selectbox("Sous-categorie", options=[""] + add_sub_options, index=0, key=skey("add","sub"))
        with col_a3:
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
        sel = st.selectbox("Sélectionner ligne", options=[""] + choices, key=skey("edit","select"))
        if sel:
            idx = int(sel.split("|")[0].strip())
            row = df_live.loc[idx].copy()
            # For edit: allow category selection outside form for reactivity
            st.write("Modifier la catégorie (réactif) :")
            edit_cat_options = [""] + ([c.strip() for c in categories_options] if categories_options else sorted({str(x).strip() for x in df_live["Categories"].dropna().astype(str).tolist()}))
            init_cat = str(row.get("Categories","")).strip()
            try:
                init_index = edit_cat_options.index(init_cat)
            except Exception:
                init_index = 0
            e_cat_sel = st.selectbox("Categories (réactif)", options=edit_cat_options, index=init_index, key=skey("edit","cat_sel"))
            # compute sub options for edit
            if isinstance(e_cat_sel, str) and e_cat_sel.strip() and visa_map_norm:
                edit_sub_options = visa_map_norm.get(e_cat_sel.strip().lower(), [])
            else:
                edit_sub_options = sorted({str(x).strip() for x in df_live["Sous-categorie"].dropna().astype(str).tolist()})
            st.sidebar.write("DEBUG edit selected cat:", repr(e_cat_sel))
            st.sidebar.write("DEBUG edit_sub_options:", edit_sub_options)
            with st.form(key=skey("form_edit")):
                ecol1, ecol2 = st.columns(2)
                with ecol1:
                    st.markdown(f"**ID_Client :** {row.get('ID_Client','')}")
                    e_dossier = st.text_input("Dossier N", value=str(row.get("Dossier N","")), key=skey("edit","dossier"))
                    e_nom = st.text_input("Nom", value=str(row.get("Nom","")), key=skey("edit","nom"))
                with ecol2:
                    e_date = st.date_input("Date", value=_date_for_widget(row.get("Date", date.today())), key=skey("edit","date"))
                    # show category chosen
                    st.markdown(f"Category choisie: **{e_cat_sel}**")
                    # sub select inside form based on computed options
                    init_sub = str(row.get("Sous-categorie","")).strip()
                    try:
                        init_sub_index = ([""] + edit_sub_options).index(init_sub)
                    except Exception:
                        init_sub_index = 0
                    e_sub = st.selectbox("Sous-categorie", options=[""] + edit_sub_options, index=init_sub_index, key=skey("edit","sub"))
                e_visa = st.text_input("Visa", value=str(row.get("Visa","")), key=skey("edit","visa_2"))
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
                        df_live.at[idx, "Categories"] = e_cat_sel.strip() if isinstance(e_cat_sel, str) else e_cat_sel
                        df_live.at[idx, "Sous-categorie"] = e_sub.strip() if isinstance(e_sub, str) else e_sub
                        df_live.at[idx, "Visa"] = e_visa
                        df_live.at[idx, "Montant honoraires (US $)"] = money_to_float(e_montant)
                        df_live.at[idx, "Autres frais (US $)"] = money_to_float(e_autres)
                        df_live.at[idx, "Payé"] = money_to_float(e_paye)
                        df_live.at[idx, "Solde"] = _to_num(df_live.at[idx, "Montant honoraires (US $)"]) + _to_num(df_live.at[idx, "Autres frais (US $)"]) - _to_num(df_live.at[idx, "Payé"])
                        df_live.at[idx, "Commentaires"] = e_comments
                        _set_df_live(df_live)
                        st.success("Modifications enregistrées.")
                    except Exception as e:
                        st.error(f"Erreur enregistrement: {e}")

    st.markdown("---")
    st.markdown("### Supprimer des dossiers")
    if df_live is None or df_live.empty:
        st.info("Aucun dossier en mémoire à supprimer.")
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
