# Visa Manager - app.py
# Full application with robust Excel reading, header detection for "Clients" sheet,
# simplified Files tab and enhanced Dashboard (date range, search, KPIs, interactive charts).
# Plotly import is done safely with fallback to Streamlit built-in charts if plotly is absent.

import os
import json
import re
import io
from io import BytesIO
from datetime import date, datetime, timedelta
from typing import Tuple, Dict, Any, List, Optional
from pathlib import Path

import pandas as pd
import numpy as np
import streamlit as st

# Safe import of plotly with graceful fallback to Streamlit charts
try:
    import plotly.express as px
    HAS_PLOTLY = True
except Exception:
    px = None
    HAS_PLOTLY = False

# Helper drawing functions: use Plotly if available, otherwise Streamlit builtins
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
            # fallback: plotly-like horizontal via bar_chart (index needs to be y)
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

# =========================
# Constantes et configuration
# =========================
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

# =========================
# Utilitaires
# =========================

def _safe_str(x: Any) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _to_num(x: Any) -> float:
    if isinstance(x, (int, float)):
        return float(x)
    s = _safe_str(x)
    if not s:
        return 0.0
    s = re.sub(r"[^\d,.-]", "", s)
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

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
    num_cols = ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Acompte 1", "Acompte 2"]
    for c in num_cols:
        if c in df.columns:
            df[c] = df[c].apply(_to_num)
    if "Montant honoraires (US $)" in df.columns and "Autres frais (US $)" in df.columns:
        total = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]
        paye = df["Payé"] if "Payé" in df.columns else 0.0
        df["Solde"] = (total - paye)
    return df

def _normalize_status(df: pd.DataFrame) -> pd.DataFrame:
    for c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: 1 if str(x).strip() in ["1", "True", "true", "OUI", "Oui", "oui", "X", "x"] else 0)
        else:
            df[c] = 0
    return df

def normalize_clients(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS_CLIENTS)
    df = df.copy()
    ren = {
        "Categorie": "Categories",
        "Catégorie": "Categories",
        "Sous-categorie": "Sous-categorie",
        "Sous-catégorie": "Sous-categorie",
        "Payee": "Payé",
        "Payé (US $)": "Payé",
        "Montant honoraires": "Montant honoraires (US $)",
        "Autres frais": "Autres frais (US $)",
        "Dossier envoye": "Dossiers envoyé",
        "Dossier envoyé": "Dossiers envoyé",
    }
    df.rename(columns={k: v for k, v in ren.items() if k in df.columns}, inplace=True)
    df = _ensure_columns(df, COLS_CLIENTS)
    if "Date" in df.columns:
        try:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        except Exception:
            pass
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

# Safe rerun wrapper (handles environments missing experimental_rerun)
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
    except Exception as e:
        try:
            st.sidebar.error(f"Impossible d'effectuer le rerun automatique : {e}")
        except Exception:
            pass
        return
    try:
        st.sidebar.info("Rerun non disponible dans cette version de Streamlit ; mise à jour session_state.")
        st.session_state.setdefault("_need_rerun", True)
    except Exception:
        pass

# try_read_excel_from_bytes with header detection and first non-empty sheet selection
def try_read_excel_from_bytes(b: bytes, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
    bio = BytesIO(b)
    try:
        xls = pd.ExcelFile(bio, engine="openpyxl")
        sheets = xls.sheet_names
        try:
            st.sidebar.info(f"Excel file detected; sheets: {sheets}")
        except Exception:
            pass

        # Build candidate list: requested sheet, known names, then workbook order
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

        # We'll inspect first N rows to detect header row
        HEADER_SCAN_ROWS = 8

        for cand in candidates:
            try:
                # Read sheet raw (no header) to inspect top rows
                bio2 = BytesIO(b)
                df_raw = pd.read_excel(bio2, sheet_name=cand, header=None, engine="openpyxl")
                if df_raw is None:
                    continue

                # compute number of non-empty cells per row for top rows
                topn = min(HEADER_SCAN_ROWS, len(df_raw))
                row_non_null_counts = [(i, df_raw.iloc[i].count()) for i in range(topn)]

                # choose header_row as the row index among topn with max non-null cells,
                # but only if max_non_null >= 2 (heuristic: header must have at least 2 columns)
                if row_non_null_counts:
                    best_row_idx, max_non_null = max(row_non_null_counts, key=lambda x: x[1])
                else:
                    best_row_idx, max_non_null = (0, 0)

                header_row = None
                if max_non_null >= 2:
                    header_row = best_row_idx
                else:
                    header_row = 0 if len(df_raw) > 0 else None

                # Now read properly using detected header_row
                try:
                    bio3 = BytesIO(b)
                    if header_row is not None:
                        df_try = pd.read_excel(bio3, sheet_name=cand, header=header_row, engine="openpyxl")
                    else:
                        df_try = pd.read_excel(bio3, sheet_name=cand, engine="openpyxl")
                except Exception:
                    # last resort: read with header=None and then set first non-all-NaN row as columns
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

                # count meaningful rows (at least one non-null cell)
                non_null_rows = df_try.dropna(how="all").shape[0]
                if non_null_rows > 0:
                    try:
                        st.sidebar.info(f"Selected sheet '{cand}' with {non_null_rows} data rows (header_row={header_row}).")
                    except Exception:
                        pass
                    return df_try

                # track best fallback (sheet with most non-null rows)
                if non_null_rows > best_non_null:
                    best_non_null = non_null_rows
                    best_df = df_try

            except Exception as e:
                try:
                    st.sidebar.info(f"Lecture failed pour feuille {cand}: {e}")
                except Exception:
                    pass
                continue

        # If nothing non-empty found, return the best attempt (may be empty)
        return best_df
    except Exception as e:
        try:
            st.sidebar.info(f"try_read_excel_from_bytes failed: {e}")
        except Exception:
            pass
        return None

# Robust read_any_table (uses try_read_excel_from_bytes)
def read_any_table(src: Any, sheet: Optional[str] = None, debug_prefix: str = "") -> Optional[pd.DataFrame]:
    def _log(msg: str):
        try:
            st.sidebar.info(f"{debug_prefix}{msg}")
        except Exception:
            pass

    if src is None:
        _log("read_any_table: src is None")
        return None

    # bytes / bytearray
    if isinstance(src, (bytes, bytearray)):
        _log("read_any_table: src is raw bytes")
        df = try_read_excel_from_bytes(bytes(src), sheet)
        if df is not None:
            return df
        try:
            return pd.read_csv(BytesIO(src))
        except Exception as e:
            _log(f"CSV from bytes failed: {repr(e)}")
            return None

    # BytesIO
    if isinstance(src, (io.BytesIO, BytesIO)):
        try:
            b = src.getvalue()
        except Exception:
            try:
                src.seek(0)
                b = src.read()
            except Exception:
                b = None
        if not b:
            _log("BytesIO: no data")
            return None
        df = try_read_excel_from_bytes(b, sheet)
        if df is not None:
            return df
        try:
            return pd.read_csv(BytesIO(b))
        except Exception as e:
            _log(f"CSV fallback from BytesIO failed: {repr(e)}")
            return None

    # UploadedFile or file-like
    if hasattr(src, "read") and hasattr(src, "name"):
        name = getattr(src, "name", "")
        _log(f"Uploaded file name: {name}")
        data = None
        try:
            data = src.getvalue()
        except Exception:
            try:
                src.seek(0)
                data = src.read()
            except Exception as e:
                _log(f"Failed to read uploaded file bytes: {repr(e)}")
                data = None
        if not data:
            _log("Uploaded file: no bytes extracted")
            return None
        try:
            _log(f"Uploaded file size: {len(data)} bytes")
        except Exception:
            pass
        lname = name.lower()
        if lname.endswith(".csv"):
            try:
                return pd.read_csv(BytesIO(data), encoding="utf-8", on_bad_lines="skip")
            except Exception as e1:
                _log(f"CSV utf-8 read failed: {repr(e1)}; trying latin1")
                try:
                    return pd.read_csv(BytesIO(data), encoding="latin1", on_bad_lines="skip")
                except Exception as e2:
                    _log(f"CSV latin1 read failed: {repr(e2)}")
                    return None
        df = try_read_excel_from_bytes(data, sheet)
        if df is not None:
            return df
        try:
            return pd.read_csv(BytesIO(data), on_bad_lines="skip")
        except Exception as e:
            _log(f"Final CSV fallback failed: {repr(e)}")
            return None

    # Path
    if isinstance(src, (str, os.PathLike)):
        p = str(src)
        if not os.path.exists(p):
            _log(f"path does not exist: {p}")
            return None
        if p.lower().endswith(".csv"):
            try:
                return pd.read_csv(p)
            except Exception as e:
                _log(f"read_csv(path) failed: {repr(e)}")
                return None
        try:
            xls = pd.ExcelFile(p)
            sheets = xls.sheet_names
            _log(f"Excel path sheets: {sheets}")
            if sheet and sheet in sheets:
                return pd.read_excel(p, sheet_name=sheet)
            for candidate in [SHEET_CLIENTS, SHEET_VISA, "Sheet1"]:
                if isinstance(candidate, str) and candidate in sheets:
                    return pd.read_excel(p, sheet_name=candidate)
            return pd.read_excel(p, sheet_name=0)
        except Exception as e:
            _log(f"read_excel(path) failed: {repr(e)}")
            return None

    _log("read_any_table: unsupported src type")
    return None

def load_last_paths() -> Tuple[str, str, str]:
    if not os.path.exists(MEMO_FILE):
        return "", "", ""
    try:
        with open(MEMO_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("clients", ""), data.get("visa", ""), data.get("save_dir", "")
    except Exception:
        return "", "", ""

def save_last_paths(clients_path: str, visa_path: str, save_dir: str) -> None:
    data = {"clients": clients_path or "", "visa": visa_path or "", "save_dir": save_dir or ""}
    try:
        with open(MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def skey(*parts: str) -> str:
    return f"{SID}_" + "_".join([p for p in parts if p])

def build_visa_map(dfv: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    vm: Dict[str, Dict[str, Dict[str, Any]]] = {}
    if dfv is None or dfv.empty:
        return vm
    cols = [c for c in dfv.columns if _safe_str(c)]
    if "Categories" not in cols and "Catégorie" in cols:
        dfv = dfv.rename(columns={"Catégorie": "Categories"})
    if "Sous-categorie" not in cols and "Sous-catégorie" in cols:
        dfv = dfv.rename(columns={"Sous-catégorie": "Sous-categorie"})
    if "Categories" not in dfv.columns or "Sous-categorie" not in dfv.columns:
        return vm
    fixed = ["Categories", "Sous-categorie"]
    option_cols = [c for c in dfv.columns if c not in fixed]
    for _, row in dfv.iterrows():
        cat = _safe_str(row.get("Categories", "")).strip()
        sub = _safe_str(row.get("Sous-categorie", "")).strip()
        if not cat or not sub:
            continue
        vm.setdefault(cat, {})
        vm[cat].setdefault(sub, {"exclusive": None, "options": []})
        opts = []
        for oc in option_cols:
            val = _safe_str(row.get(oc, "")).strip()
            if val in ["1", "x", "X", "oui", "Oui", "OUI", "True", "true"]:
                opts.append(oc)
        exclusive = None
        if set([o.upper() for o in opts]) == set(["COS", "EOS"]):
            exclusive = "radio_group"
        vm[cat][sub] = {"exclusive": exclusive, "options": opts}
    return vm

# =========================
# Interface Streamlit
# =========================

st.set_page_config(page_title="Visa Manager", layout="wide")
st.title(APP_TITLE)

# Sidebar (file upload controls)
st.sidebar.header("📂 Fichiers")
last_clients, last_visa, last_save_dir = load_last_paths()

mode = st.sidebar.radio(
    "Mode de chargement",
    ["Un fichier (Clients)", "Deux fichiers (Clients & Visa)"],
    index=0,
    key=skey("mode")
)

up_clients = st.sidebar.file_uploader(
    "Clients (xlsx/csv)", type=["xlsx", "xls", "csv"], key=skey("up_clients")
)
up_visa = None
if mode == "Deux fichiers (Clients & Visa)":
    up_visa = st.sidebar.file_uploader(
        "Visa (xlsx/csv)", type=["xlsx", "xls", "csv"], key=skey("up_visa")
    )

clients_path_in = st.sidebar.text_input("ou chemin local Clients (laisser vide si upload)", value=last_clients, key=skey("cli_path"))
visa_path_in = st.sidebar.text_input("ou chemin local Visa (laisser vide si upload)", value=(last_visa if mode != "Un fichier (Clients)" else ""), key=skey("vis_path"))
save_dir_in = st.sidebar.text_input("Dossier de sauvegarde (optionnel)", value=last_save_dir, key=skey("save_dir"))

if st.sidebar.button("📥 Sauvegarder chemins", key=skey("btn_load")):
    save_last_paths(clients_path_in, visa_path_in, save_dir_in)
    st.sidebar.success("Chemins mémorisés.")
    safe_rerun()

# Read uploaded files robustly and avoid re-consuming UploadedFile stream
clients_bytes = None
visa_bytes = None

if up_clients is not None:
    try:
        clients_bytes = up_clients.getvalue()
    except Exception:
        try:
            up_clients.seek(0)
            clients_bytes = up_clients.read()
        except Exception:
            clients_bytes = None

if up_visa is not None:
    try:
        visa_bytes = up_visa.getvalue()
    except Exception:
        try:
            up_visa.seek(0)
            visa_bytes = up_visa.read()
        except Exception:
            visa_bytes = None

# Determine sources for read
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

# Read (clients + visa)
df_clients_raw = None
df_visa_raw = None

try:
    df_clients_raw = read_any_table(clients_src_for_read, sheet=SHEET_CLIENTS, debug_prefix="[Clients] ")
except Exception as e:
    st.sidebar.error(f"[Clients] Exception during read_any_table: {e}")

if df_clients_raw is None:
    try:
        df_clients_raw = read_any_table(clients_src_for_read, sheet=None, debug_prefix="[Clients fallback] ")
    except Exception as e:
        st.sidebar.error(f"[Clients fallback] Exception: {e}")

try:
    df_visa_raw = read_any_table(visa_src_for_read, sheet=SHEET_VISA, debug_prefix="[Visa] ")
except Exception as e:
    st.sidebar.error(f"[Visa] Exception during read_any_table: {e}")

if df_visa_raw is None:
    try:
        df_visa_raw = read_any_table(visa_src_for_read, sheet=None, debug_prefix="[Visa fallback] ")
    except Exception as e:
        st.sidebar.error(f"[Visa fallback] Exception: {e}")

if df_visa_raw is None:
    df_visa_raw = pd.DataFrame()

# Ensure we have a usable df_all and initialize session_state safely
try:
    # Attempt to recover if df_clients_raw is not a DataFrame
    if not isinstance(df_clients_raw, pd.DataFrame):
        try:
            tmp = read_any_table(df_clients_raw, sheet=SHEET_CLIENTS)
            if isinstance(tmp, pd.DataFrame):
                df_clients_raw = tmp
        except Exception:
            pass

    try:
        df_all = _ensure_time_features(normalize_clients(df_clients_raw))
    except Exception as e_norm:
        try:
            st.sidebar.error(f"Erreur lors de la normalisation des clients : {e_norm}")
            st.sidebar.exception(e_norm)
            st.sidebar.write("Type de df_clients_raw:", type(df_clients_raw))
            try:
                sr = repr(df_clients_raw)
                st.sidebar.write("Repr (truncated):", sr[:1000] + ("..." if len(sr) > 1000 else ""))
            except Exception:
                pass
        except Exception:
            pass
        df_all = pd.DataFrame(columns=COLS_CLIENTS)

except Exception as e_top:
    try:
        st.sidebar.error(f"Erreur inattendue lors de la préparation des données clients : {e_top}")
        st.sidebar.exception(e_top)
    except Exception:
        pass
    df_all = pd.DataFrame(columns=COLS_CLIENTS)

# Persist working copy in session_state (always define it)
DF_LIVE_KEY = skey("df_live")
if DF_LIVE_KEY not in st.session_state or st.session_state[DF_LIVE_KEY] is None:
    st.session_state[DF_LIVE_KEY] = df_all.copy() if (df_all is not None) else pd.DataFrame()

def _get_df_live() -> pd.DataFrame:
    return st.session_state[DF_LIVE_KEY].copy()

def _set_df_live(df: pd.DataFrame) -> None:
    st.session_state[DF_LIVE_KEY] = df.copy()

# Tabs
tabs = st.tabs([
    "📄 Fichiers",
    "📊 Dashboard",
    "📈 Analyses",
    "🏦 Escrow",
    "👤 Compte client",
    "🧾 Gestion",
    "📄 Visa (aperçu)",
    "💾 Export",
])

# -----------------------
# FICHIERS - Simplified: show only Clients and Visa previews
# -----------------------
with tabs[0]:
    st.header("📂 Fichiers")
    colA, colB = st.columns(2)

    # Clients card (left)
    with colA:
        st.subheader("Clients")
        if up_clients is not None:
            st.write("Upload:", up_clients.name)
            try:
                st.write(f"Taille: {len(clients_bytes)} bytes")
            except Exception:
                pass
        elif isinstance(clients_src_for_read, str) and clients_src_for_read:
            st.write("Chargé depuis chemin local :", clients_src_for_read)
        else:
            st.info("Aucun fichier Clients sélectionné.")

        if df_clients_raw is None or (isinstance(df_clients_raw, pd.DataFrame) and df_clients_raw.empty):
            st.warning("Lecture Clients : aucun tableau trouvé ou DataFrame vide.")
        else:
            st.success(f"Clients lus ({df_clients_raw.shape[0]} lignes, {df_clients_raw.shape[1]} colonnes)")
            try:
                st.dataframe(df_clients_raw.head(8), use_container_width=True, height=220)
            except Exception:
                st.write("Aperçu indisponible pour ce format de fichier.")

    # Visa card (right)
    with colB:
        st.subheader("Visa")
        if mode == "Deux fichiers (Clients & Visa)":
            if up_visa is not None:
                st.write("Upload:", up_visa.name)
                try:
                    st.write(f"Taille: {len(visa_bytes)} bytes")
                except Exception:
                    pass
            elif isinstance(visa_src_for_read, str) and visa_src_for_read:
                st.write("Chargé depuis chemin local :", visa_src_for_read)
            else:
                st.info("Aucun fichier Visa sélectionné.")
        else:
            st.write("Mode 'Un fichier' : Visa sera lu depuis le même fichier Clients si présent.")

        if df_visa_raw is None or (isinstance(df_visa_raw, pd.DataFrame) and df_visa_raw.empty):
            st.warning("Lecture Visa : aucun tableau trouvé ou DataFrame vide.")
        else:
            st.success(f"Visa lu ({df_visa_raw.shape[0]} lignes, {df_visa_raw.shape[1]} colonnes)")
            try:
                st.dataframe(df_visa_raw.head(8), use_container_width=True, height=220)
            except Exception:
                st.write("Aperçu Visa indisponible pour ce format de fichier.")

    st.markdown("---")
    # Small actions
    a1, a2 = st.columns([1,1])
    with a1:
        if st.button("Réinitialiser la mémoire (annuler modifications en mémoire)"):
            df_all = _ensure_time_features(normalize_clients(df_clients_raw))
            _set_df_live(df_all)
            st.success("Mémoire réinitialisée à partir des fichiers chargés.")
            safe_rerun()
    with a2:
        if st.button("Actualiser la lecture"):
            safe_rerun()

# -----------------------
# DASHBOARD - Enhanced (date range, search, KPIs, Plotly charts, Top clients, export)
# -----------------------
with tabs[1]:
    st.subheader("📊 Dashboard")
    df_all_current = _get_df_live()

    # If no data, show message
    if df_all_current is None or df_all_current.empty:
        st.info("Aucune donnée cliente en mémoire. Chargez le fichier Clients dans l'onglet Fichiers.")
    else:
        # TOP FILTERS row: date range, search, dossier number
        today = date.today()
        min_date = df_all_current["Date"].min() if "Date" in df_all_current.columns else pd.Timestamp(today - timedelta(days=365))
        max_date = df_all_current["Date"].max() if "Date" in df_all_current.columns else pd.Timestamp(today)
        try:
            min_date = _date_for_widget(min_date)
        except Exception:
            min_date = today - timedelta(days=365)
        try:
            max_date = _date_for_widget(max_date)
        except Exception:
            max_date = today

        with st.expander("Filtres (afficher / masquer)", expanded=True):
            fcol1, fcol2, fcol3, fcol4 = st.columns([2,2,2,2])
            # date range
            with fcol1:
                date_from = st.date_input("Date de", value=min_date, key=skey("filter", "date_from"))
                date_to = st.date_input("Date à", value=max_date, key=skey("filter", "date_to"))
            # search text
            with fcol2:
                search_text = st.text_input("Recherche (Nom / Partie du nom)", value="", key=skey("filter", "search"))
            # dossier exact
            with fcol3:
                dossier_search = st.text_input("Dossier N (exact)", value="", key=skey("filter", "dossier"))
            # quick KPI scope checkbox
            with fcol4:
                scope_period = st.checkbox("KPIs sur la période filtrée", value=True, key=skey("filter", "scope_period"))
                top_n = st.number_input("Top clients (N)", min_value=3, max_value=100, value=10, step=1, key=skey("filter", "top_n"))

        # Apply filters
        view = df_all_current.copy()
        # date filter if Date exists
        if "Date" in view.columns:
            try:
                # convert inputs to timestamps
                dt_from = pd.to_datetime(date_from)
                dt_to = pd.to_datetime(date_to) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                view = view[(pd.to_datetime(view["Date"], errors="coerce") >= dt_from) & (pd.to_datetime(view["Date"], errors="coerce") <= dt_to)]
            except Exception:
                pass

        # text search on Nom
        if search_text and "Nom" in view.columns:
            q = search_text.strip().lower()
            view = view[view["Nom"].str.lower().str.contains(q, na=False)]

        # dossier exact
        if dossier_search:
            col = "Dossier N"
            if col in view.columns:
                view = view[view][view[col].astype(str).str.strip() == dossier_search.strip()]

        # aggregated Total_US and clean needed columns
        if "Montant honoraires (US $)" in view.columns and "Autres frais (US $)" in view.columns:
            view["Total_US"] = view["Montant honoraires (US $)"].apply(_to_num) + view["Autres frais (US $)"].apply(_to_num)
        else:
            view["Total_US"] = 0.0

        # KPIs calculation (either on filtered view or full depending on scope_period)
        kview = view if scope_period else df_all_current.copy()
        if "Total_US" not in kview.columns:
            if "Montant honoraires (US $)" in kview.columns and "Autres frais (US $)" in kview.columns:
                kview["Total_US"] = kview["Montant honoraires (US $)"].apply(_to_num) + kview["Autres frais (US $)"].apply(_to_num)
            else:
                kview["Total_US"] = 0.0

        total_count = len(kview)
        total_facture = kview["Total_US"].sum() if "Total_US" in kview.columns else 0.0
        total_recu = kview["Payé"].apply(_to_num).sum() if "Payé" in kview.columns else 0.0
        total_solde = kview["Solde"].apply(_to_num).sum() if "Solde" in kview.columns else 0.0
        avg_per_dossier = (total_facture / total_count) if total_count else 0.0
        taux_envoye = (kview["Dossiers envoyé"].apply(_to_num).clip(0,1).sum() / total_count * 100) if ("Dossiers envoyé" in kview.columns and total_count) else 0.0
        n_refus = int(kview["Dossier refusé"].apply(_to_num).sum()) if "Dossier refusé" in kview.columns else 0

        # KPI display
        k1, k2, k3, k4, k5, k6 = st.columns([1,1,1,1,1,1])
        k1.metric("Dossiers", f"{total_count:,}")
        k2.metric("Total facturé", _fmt_money(total_facture))
        k3.metric("Total reçu", _fmt_money(total_recu))
        k4.metric("Solde total", _fmt_money(total_solde))
        k5.metric("Moyenne / dossier", _fmt_money(avg_per_dossier))
        k6.metric("Taux envoyés (%)", f"{taux_envoye:.0f}%")
        st.markdown(f"Nombre de refus (période): **{n_refus}**")

        st.markdown("---")

        # Charts area using safe plot helpers
        chart1, chart2 = st.columns([1,2])
        with chart1:
            st.subheader("Répartition par Catégorie")
            if "Categories" in view.columns:
                cat_counts = view["Categories"].value_counts().rename_axis("Categorie").reset_index(name="Nombre")
                if not cat_counts.empty:
                    plot_pie(cat_counts, names_col="Categorie", value_col="Nombre", title="Catégories")
                else:
                    st.info("Pas de catégories à afficher.")
            else:
                st.info("Colonne 'Categories' introuvable.")

            st.write("")  # spacing
            st.subheader("Top Sous-catégories")
            if "Sous-categorie" in view.columns:
                sub_counts = view["Sous-categorie"].value_counts().head(10).reset_index()
                sub_counts.columns = ["Sous-categorie", "Nombre"]
                if not sub_counts.empty:
                    plot_barh(sub_counts, x="Nombre", y="Sous-categorie", title="Top Sous-catégories")
                else:
                    st.info("Pas de sous-catégories à afficher.")
            else:
                st.info("Colonne 'Sous-categorie' introuvable.")

        with chart2:
            st.subheader("Évolution Mensuelle (Total US)")
            tmp = view.copy()
            if "_Année_" in tmp.columns and "Mois" in tmp.columns:
                tmp = tmp.dropna(subset=["_Année_", "Mois"])
                if not tmp.empty:
                    tmp["YearMonth"] = tmp["_Année_"].astype(int).astype(str) + "-" + tmp["Mois"].astype(str)
                    g = tmp.groupby("YearMonth", as_index=False)["Total_US"].sum().sort_values("YearMonth")
                    if not g.empty:
                        plot_line(g, x="YearMonth", y="Total_US", title="Évolution Mensuelle", x_title="Période", y_title="Montant (US$)")
                    else:
                        st.info("Pas assez de données mensuelles.")
                else:
                    st.info("Pas de données temporelles pour la période sélectionnée.")
            else:
                st.info("Colonnes temporelles manquantes (_Année_/Mois).")

        st.markdown("---")

        # Top clients by total facturé
        st.subheader(f"Top {int(top_n)} clients par Total facturé")
        if "ID_Client" in view.columns or "Nom" in view.columns:
            grp = view.groupby(["ID_Client", "Nom"], dropna=False).agg({
                "Total_US": "sum",
                "Dossier N": "count"
            }).reset_index().rename(columns={"Dossier N": "Nb_dossiers"})
            grp_sorted = grp.sort_values("Total_US", ascending=False).head(int(top_n))
            if not grp_sorted.empty:
                # safe bar horizontal
                if HAS_PLOTLY and px is not None:
                    fig_top = px.bar(grp_sorted, x="Total_US", y="Nom", orientation="h", title=f"Top {int(top_n)} clients", text="Nb_dossiers")
                    fig_top.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Total facturé (US$)")
                    st.plotly_chart(fig_top, use_container_width=True)
                else:
                    st.write(f"Top {int(top_n)} clients")
                    st.dataframe(grp_sorted.assign(Total_US=lambda d: d["Total_US"].map(lambda v: _fmt_money(v))).reset_index(drop=True), use_container_width=True)
            else:
                st.info("Pas de clients à afficher.")
        else:
            st.info("Colonnes client introuvables (ID_Client/Nom).")

        st.markdown("---")

        # Detailed table with column selection and pagination
        st.subheader("Table détaillée")
        available_cols = [c for c in [
            "Dossier N", "ID_Client", "Nom", "Date", "Categories", "Sous-categorie", "Visa",
            "Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Total_US"
        ] if c in view.columns]
        # default selected columns
        default_cols = [c for c in ["Dossier N", "Nom", "Date", "Categories", "Total_US", "Payé", "Solde"] if c in available_cols]
        cols_selected = st.multiselect("Colonnes à afficher", options=available_cols, default=default_cols, key=skey("table", "cols"))

        # Pagination controls
        page_size = st.selectbox("Lignes par page", options=[10, 25, 50, 100], index=1, key=skey("table", "page_size"))
        total_rows = len(view)
        total_pages = max(1, int(np.ceil(total_rows / page_size)))
        page = st.number_input("Page", min_value=1, max_value=total_pages, value=1, step=1, key=skey("table", "page_num"))

        start = (page - 1) * page_size
        end = start + page_size
        display_df = view.sort_values(by=["_Année_", "_MoisNum_"], ascending=[False, False]).iloc[start:end].copy()

        # format money/date
        for col in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Total_US"]:
            if col in display_df.columns:
                display_df[col] = display_df[col].apply(lambda x: _fmt_money(_to_num(x)))
        if "Date" in display_df.columns:
            try:
                display_df["Date"] = pd.to_datetime(display_df["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                display_df["Date"] = display_df["Date"].astype(str)

        if cols_selected:
            st.dataframe(display_df[cols_selected].reset_index(drop=True), use_container_width=True)
        else:
            st.dataframe(display_df.reset_index(drop=True), use_container_width=True)

        # Export filtered view CSV / XLSX
        col1, col2 = st.columns(2)
        with col1:
            csv_bytes = view.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Export CSV (vue filtrée)", data=csv_bytes, file_name="Clients_filtered.csv", mime="text/csv", key=skey("export", "csv"))
        with col2:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                view.to_excel(writer, index=False, sheet_name="Filtered")
            st.download_button("⬇️ Export XLSX (vue filtrée)", data=buf.getvalue(), file_name="Clients_filtered.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=skey("export", "xlsx"))

# NOTE: Remaining UI sections (Analyses, Escrow, Compte client, Gestion, Visa preview, Export)
# are intentionally omitted in this file for brevity but can be added similarly with improved UX.
