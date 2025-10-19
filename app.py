# url=https://github.com/charleytrigano/streamlit_visa_app/blob/main/app.py
# Visa Manager - app.py
# Full application with robust Excel reading and header detection for "Clients" sheet.
# Fichiers tab simplified: only shows Clients and Visa read / previews.
import os
import json
import re
import io
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional
from pathlib import Path

import pandas as pd
import numpy as np
import streamlit as st

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

# Sidebar (keep file upload controls here for convenience)
st.sidebar.header("📂 Fichiers")
last_clients, last_visa, last_save_dir = load_last_paths()

mode = st.sidebar.radio(
    "Mode de chargement",
    ["Un fichier (Clients)", "Deux fichiers (Clients & Visa)"],
    index=0,
    key=skey("mode")
)

# Keep uploaders in sidebar; Fichiers tab will only show read previews
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

# Persist working copy in session_state
DF_LIVE_KEY = skey("df_live")
if DF_LIVE_KEY not in st.session_state or st.session_state[DF_LIVE_KEY] is None:
    df_all = _ensure_time_features(normalize_clients(df_clients_raw))
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
            # show compact preview (first 8 rows) and let user expand to see full head if needed
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
            # reset session df_live to current read
            df_all = _ensure_time_features(normalize_clients(df_clients_raw))
            _set_df_live(df_all)
            st.success("Mémoire réinitialisée à partir des fichiers chargés.")
            safe_rerun()
    with a2:
        if st.button("Actualiser la lecture"):
            # re-run read (simple feedback)
            st.experimental_rerun()

# -----------------------
# Dashboard (unchanged structure, improved presentation)
# -----------------------
with tabs[1]:
    st.subheader("📊 Dashboard")
    df_all_current = _get_df_live()

    # FILTERS area (top) - compact and consistent
    with st.container():
        f1, f2, f3, f4 = st.columns([2,2,2,3])
        cats = sorted(df_all_current["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all_current.columns else []
        subs = sorted(df_all_current["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all_current.columns else []
        visas = sorted(df_all_current["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all_current.columns else []
        years = sorted(pd.to_numeric(df_all_current["_Année_"], errors="coerce").dropna().astype(int).unique().tolist()) if "_Année_" in df_all_current.columns else []

        sel_cat = f1.multiselect("Catégories", options=cats, default=[], key=skey("dash", "cats"))
        sel_sub = f2.multiselect("Sous-catégories", options=subs, default=[], key=skey("dash", "subs"))
        sel_visa = f3.multiselect("Visa", options=visas, default=[], key=skey("dash", "visas"))
        sel_year = f4.multiselect("Années", options=years, default=[], key=skey("dash", "years"))

    # Apply filters
    view = df_all_current.copy() if df_all_current is not None else pd.DataFrame()
    if sel_cat:
        view = view[view["Categories"].astype(str).isin(sel_cat)]
    if sel_sub:
        view = view[view["Sous-categorie"].astype(str).isin(sel_sub)]
    if sel_visa:
        view = view[view["Visa"].astype(str).isin(sel_visa)]
    if sel_year:
        view = view[view["_Année_"].isin(sel_year)]

    if view is None or view.empty:
        st.warning("Aucune donnée correspondant aux filtres.")
    else:
        # KPIs row
        total_clients = len(view)
        total_honoraires = (view["Montant honoraires (US $)"].apply(_to_num) + view["Autres frais (US $)"].apply(_to_num)).sum()
        total_paye = view["Payé"].apply(_to_num).sum()
        total_solde = view["Solde"].apply(_to_num).sum()

        kcol1, kcol2, kcol3, kcol4 = st.columns([1.4,1.4,1.4,1.4])
        kcol1.metric("Dossiers", f"{total_clients:,}")
        kcol2.metric("Total Facturé", _fmt_money(total_honoraires))
        kcol3.metric("Total Reçu", _fmt_money(total_paye))
        kcol4.metric("Solde Total", _fmt_money(total_solde))

        st.markdown("---")

        # Charts row
        c1, c2 = st.columns([1, 2])
        with c1:
            st.subheader("Répartition par Catégorie")
            if "Categories" in view.columns:
                cat_counts = view["Categories"].value_counts().rename_axis("Categorie").reset_index(name="Nombre")
                if not cat_counts.empty:
                    st.bar_chart(cat_counts.set_index("Categorie")["Nombre"])
                else:
                    st.info("Pas de catégories à afficher.")
            else:
                st.info("Colonne 'Categories' introuvable.")

            st.subheader("Top 10 Sous-catégories")
            if "Sous-categorie" in view.columns:
                sub_counts = view["Sous-categorie"].value_counts().head(10)
                st.bar_chart(sub_counts)
            else:
                st.info("Colonne 'Sous-categorie' introuvable.")

        with c2:
            st.subheader("Évolution Mensuelle (Honoraires + Frais)")
            tmp = view.copy()
            tmp["Total_US"] = tmp["Montant honoraires (US $)"].apply(_to_num) + tmp["Autres frais (US $)"].apply(_to_num)
            if "_Année_" in tmp.columns and "Mois" in tmp.columns:
                tmp["YearMonth"] = tmp["_Année_"].astype(str) + "-" + tmp["Mois"].astype(str)
                g = tmp.groupby("YearMonth", as_index=False)["Total_US"].sum().sort_values("YearMonth")
                if not g.empty:
                    g = g.set_index("YearMonth")
                    st.line_chart(g)
                else:
                    st.info("Pas assez de données mensuelles.")
            else:
                st.info("Colonnes temporelles manquantes (_Année_/Mois).")

        st.markdown("---")

        # Recent table
        st.subheader("Aperçu récent des dossiers")
        recent = view.sort_values(by=["_Année_", "_MoisNum_"], ascending=[False, False]).head(20).copy()
        display_cols = [c for c in [
            "Dossier N", "ID_Client", "Nom", "Date", "Categories", "Sous-categorie", "Visa",
            "Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"
        ] if c in recent.columns]
        for col in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"]:
            if col in recent.columns:
                recent[col] = recent[col].apply(lambda x: _fmt_money(_to_num(x)))
        if "Date" in recent.columns:
            try:
                recent["Date"] = pd.to_datetime(recent["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                recent["Date"] = recent["Date"].astype(str)
        st.dataframe(recent[display_cols].reset_index(drop=True), use_container_width=True)

# NOTE: Remaining UI sections omitted for brevity; unchanged.
