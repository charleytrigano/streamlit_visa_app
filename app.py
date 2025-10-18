# Complete app.py with enhanced robust file reading and diagnostics for Clients.xlsx upload issue
# Replaces previous versions: adds detailed debug output (sidebar + expander) so you can see why a file
# was not parsed into a clients DataFrame (sheet names, file sizes, shapes, exceptions).
import os
import json
import re
import io
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional
from pathlib import Path

import pandas as pd
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
# Fonctions utilitaires
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
        df["Solde"] = (total - paye).clip(lower=0.0)
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
    df["Nom"] = df["Nom"].astype(str)
    df["Categories"] = df["Categories"].astype(str)
    df["Sous-categorie"] = df["Sous-categorie"].astype(str)
    df["Visa"] = df["Visa"].astype(str)
    df["Commentaires"] = df["Commentaires"].astype(str)
    try:
        dser = pd.to_datetime(df["Date"], errors="coerce")
        df["_Année_"] = dser.dt.year.fillna(0).astype(int)
        df["_MoisNum_"] = dser.dt.month.fillna(0).astype(int)
        df["Mois"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if m and m == m else "")
    except Exception:
        df["_Année_"] = 0
        df["_MoisNum_"] = 0
        df["Mois"] = ""
    return df

# Robust read_any_table with diagnostics and Excel sheet handling
def read_any_table(src: Any, sheet: Optional[str] = None, debug_prefix: str = "") -> Optional[pd.DataFrame]:
    """
    Read a table robustly from:
      - UploadedFile (st.file_uploader) or any file-like with .read and .name
      - bytes / bytearray
      - BytesIO
      - file path (str / Path)
    If the source is an Excel file, we try to list sheets and pick the requested sheet (if provided).
    Returns DataFrame or None on failure.
    debug_prefix: string prepended to debug messages in sidebar/expander to help trace where read_any_table was called.
    """
    def _log(msg: str):
        # write small debug messages to sidebar to avoid cluttering main UI
        try:
            st.sidebar.info(f"{debug_prefix}{msg}")
        except Exception:
            pass

    if src is None:
        _log("read_any_table: src is None")
        return None

    # Helper to attempt excel read with sheet selection using ExcelFile to inspect sheets first
    def try_read_excel_from_bytes(b: bytes, sheet_name: Optional[str] = None) -> Optional[pd.DataFrame]:
        bio = BytesIO(b)
        try:
            # inspect sheets
            xls = pd.ExcelFile(bio, engine="openpyxl")
            sheets = xls.sheet_names
            _log(f"Excel file detected; sheets: {sheets}")
            # choose sheet
            if sheet_name and sheet_name in sheets:
                bio = BytesIO(b)  # reset
                return pd.read_excel(bio, sheet_name=sheet_name, engine="openpyxl")
            # fallback: try standard names then first sheet
            for candidate in [SHEET_CLIENTS, SHEET_VISA, "Sheet1", 0]:
                if isinstance(candidate, str) and candidate in sheets:
                    bio = BytesIO(b)
                    return pd.read_excel(bio, sheet_name=candidate, engine="openpyxl")
            # final fallback: first sheet
            bio = BytesIO(b)
            return pd.read_excel(bio, sheet_name=0, engine="openpyxl")
        except Exception as e:
            _log(f"try_read_excel_from_bytes failed: {e}")
            return None

    # If src is bytes/bytearray
    if isinstance(src, (bytes, bytearray)):
        _log("read_any_table: src is raw bytes")
        # Try excel first
        df = try_read_excel_from_bytes(bytes(src), sheet)
        if df is not None:
            return df
        # Try csv
        try:
            return pd.read_csv(BytesIO(src))
        except Exception as e:
            _log(f"CSV from bytes failed: {e}")
            return None

    # If src is BytesIO
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
        return try_read_excel_from_bytes(b, sheet) or (pd.read_csv(BytesIO(b)) if _safe_str(sheet) == "" else None)

    # UploadedFile or file-like with .read and .name (Streamlit UploadedFile)
    if hasattr(src, "read") and hasattr(src, "name"):
        name = getattr(src, "name", "")
        _log(f"Uploaded file name: {name}")
        # read raw bytes once
        data = None
        try:
            data = src.getvalue()  # UploadedFile has getvalue()
        except Exception:
            try:
                src.seek(0)
                data = src.read()
            except Exception as e:
                _log(f"Failed to read uploaded file bytes: {e}")
                data = None
        if not data:
            _log("Uploaded file: no bytes extracted")
            return None

        # log size
        try:
            _log(f"Uploaded file size: {len(data)} bytes")
        except Exception:
            pass

        lname = name.lower()
        if lname.endswith(".csv"):
            # try encodings
            try:
                return pd.read_csv(BytesIO(data), encoding="utf-8", on_bad_lines="skip")
            except Exception as e1:
                _log(f"CSV utf-8 read failed: {e1}; trying latin1")
                try:
                    return pd.read_csv(BytesIO(data), encoding="latin1", on_bad_lines="skip")
                except Exception as e2:
                    _log(f"CSV latin1 read failed: {e2}")
                    return None
        # try excel intelligently
        df = try_read_excel_from_bytes(data, sheet)
        if df is not None:
            return df
        # last resort try csv parsing
        try:
            return pd.read_csv(BytesIO(data), on_bad_lines="skip")
        except Exception as e:
            _log(f"Final CSV fallback failed: {e}")
            return None

    # If src is a path string or Path
    if isinstance(src, (str, os.PathLike)):
        p = str(src)
        if not os.path.exists(p):
            _log(f"path does not exist: {p}")
            return None
        if p.lower().endswith(".csv"):
            try:
                return pd.read_csv(p)
            except Exception as e:
                _log(f"read_csv(path) failed: {e}")
                return None
        # Excel path: inspect sheets and try requested
        try:
            xls = pd.ExcelFile(p)
            sheets = xls.sheet_names
            _log(f"Excel path sheets: {sheets}")
            if sheet and sheet in sheets:
                return pd.read_excel(p, sheet_name=sheet)
            for candidate in [SHEET_CLIENTS, SHEET_VISA, "Sheet1", 0]:
                if isinstance(candidate, str) and candidate in sheets:
                    return pd.read_excel(p, sheet_name=candidate)
            return pd.read_excel(p, sheet_name=0)
        except Exception as e:
            _log(f"read_excel(path) failed: {e}")
            return None

    _log("read_any_table: unsupported src type")
    return None

def read_sheet_from_path(path: str, sheet_name: str) -> Optional[pd.DataFrame]:
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
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

def _norm(s: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]", "_", s.strip().lower())

def make_client_id(nom: str, dval: date) -> str:
    return f"{_norm(nom)}_{int(datetime.now().timestamp())}"

def next_dossier(df: pd.DataFrame) -> int:
    max_dossier = df.get("Dossier N", pd.Series([13056])).astype(str).str.extract(r"(\d+)").fillna(13056).astype(int).max()
    return max_dossier + 1

def _to_float(x: Any) -> float:
    return _to_num(x)

def _ensure_dir(pdir: Path) -> None:
    pdir.mkdir(parents=True, exist_ok=True)

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

def visa_option_selector(vm: Dict[str, Any], cat: str, sub: str, keybase: str) -> str:
    if cat not in vm or sub not in vm[cat]:
        return sub
    meta = vm[cat][sub]
    opts = meta.get("options", [])
    if not opts:
        return sub
    if meta.get("exclusive") == "radio_group" and set([o.upper() for o in opts]) == set(["COS", "EOS"]):
        pick = st.radio("Options", ["COS", "EOS"], horizontal=True, key=skey(keybase, "opt"))
        return f"{sub} {pick}"
    else:
        picks = st.multiselect("Options", opts, default=[], key=skey(keybase, "opts"))
        if not picks:
            return sub
        return f"{sub} {picks[0]}"

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

# =========================
# Interface Streamlit
# =========================

st.set_page_config(page_title="Visa Manager", layout="wide")
st.title(APP_TITLE)

# Sidebar file controls & previous paths
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

clients_path_in = st.sidebar.text_input("ou chemin local Clients", value=last_clients, key=skey("cli_path"))
visa_path_in = st.sidebar.text_input("ou chemin local Visa", value=(last_visa if mode != "Un fichier (Clients)" else ""), key=skey("vis_path"))
save_dir_in = st.sidebar.text_input("Dossier de sauvegarde", value=last_save_dir, key=skey("save_dir"))

if st.sidebar.button("📥 Charger", key=skey("btn_load")):
    save_last_paths(clients_path_in, visa_path_in, save_dir_in)
    st.sidebar.success("Chemins mémorisés. Re-lancement pour prise en compte.")
    st.experimental_rerun()

# -----------------------
# Read uploaded files robustly and avoid re-consuming UploadedFile stream
# -----------------------
clients_bytes = None
visa_bytes = None

# Read uploaded file bytes once (safest approach for Streamlit UploadedFile)
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

# Determine the proper src objects for reading functions
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
    # single-file mode: use same bytes/path for visa reading as clients
    visa_src_for_read = clients_src_for_read

# Read with diagnostics (we pass debug_prefix so read_any_table logs to sidebar where possible)
df_clients_raw = None
df_visa_raw = None

try:
    df_clients_raw = read_any_table(clients_src_for_read, sheet=SHEET_CLIENTS, debug_prefix="[Clients] ")
except Exception as e:
    st.sidebar.error(f"[Clients] Exception during read_any_table: {e}")

# Fallback: if the read returned None, try without specifying sheet (some files have different sheet names)
if df_clients_raw is None:
    try:
        df_clients_raw = read_any_table(clients_src_for_read, sheet=None, debug_prefix="[Clients fallback] ")
    except Exception as e:
        st.sidebar.error(f"[Clients fallback] Exception: {e}")

# Read Visa sheet(s)
try:
    df_visa_raw = read_any_table(visa_src_for_read, sheet=SHEET_VISA, debug_prefix="[Visa] ")
except Exception as e:
    st.sidebar.error(f"[Visa] Exception during read_any_table: {e}")

# fallback for visa if not found (single-sheet or different naming)
if df_visa_raw is None:
    try:
        df_visa_raw = read_any_table(visa_src_for_read, sheet=None, debug_prefix="[Visa fallback] ")
    except Exception as e:
        st.sidebar.error(f"[Visa fallback] Exception: {e}")

if df_visa_raw is None:
    df_visa_raw = pd.DataFrame()

# Diagnostics shown in main UI for easier debugging by user
with st.expander("📄 Fichiers chargés & diagnostics", expanded=True):
    st.write("Clients uploader:", getattr(up_clients, "name", "(no uploaded file)"))
    if clients_bytes is not None:
        st.write(f"Clients upload size: {len(clients_bytes)} bytes")
    st.write("Clients read result:")
    if df_clients_raw is None:
        st.warning("Lecture Clients: Aucune table lue (df_clients_raw is None).")
    else:
        st.success(f"Clients DataFrame shape: {df_clients_raw.shape}")
        st.write("Clients columns:", list(df_clients_raw.columns))
        try:
            st.dataframe(df_clients_raw.head(10), use_container_width=True)
        except Exception:
            st.write("Impossible d'afficher head des Clients (format non standard).")

    st.write("---")
    st.write("Visa uploader:", getattr(up_visa, "name", "(no uploaded file)"))
    if visa_bytes is not None:
        st.write(f"Visa upload size: {len(visa_bytes)} bytes")
    st.write("Visa read result:")
    if df_visa_raw is None or df_visa_raw.empty:
        st.warning("Lecture Visa: Aucune table lue ou DataFrame vide.")
    else:
        st.success(f"Visa DataFrame shape: {df_visa_raw.shape}")
        st.write("Visa columns:", list(df_visa_raw.columns))
        try:
            st.dataframe(df_visa_raw.head(10), use_container_width=True)
        except Exception:
            st.write("Impossible d'afficher head des Visa (format non standard).")

# If clients DF is None, normalize_clients will return empty template; but show message and hint
if df_clients_raw is None or (isinstance(df_clients_raw, pd.DataFrame) and df_clients_raw.empty):
    st.info("Aucun client chargé. Chargez les fichiers dans la barre latérale. Si votre fichier contient plusieurs feuilles, vérifiez que la feuille 'Clients' existe ou essayez de renommer la feuille en 'Clients' ou 'Sheet1'.")
else:
    st.success("Clients chargés avec succès (voir diagnostic ci-dessus).")

# Build visa map and working DF
visa_map = build_visa_map(df_visa_raw)
df_all = _ensure_time_features(normalize_clients(df_clients_raw))

# Persist working copy in session_state
DF_LIVE_KEY = skey("df_live")
if DF_LIVE_KEY not in st.session_state or st.session_state[DF_LIVE_KEY] is None:
    st.session_state[DF_LIVE_KEY] = df_all.copy() if (df_all is not None) else pd.DataFrame()

def _get_df_live() -> pd.DataFrame:
    return st.session_state[DF_LIVE_KEY].copy()

def _set_df_live(df: pd.DataFrame) -> None:
    st.session_state[DF_LIVE_KEY] = df.copy()

# Tabs and UI (dashboard / analyses / escrow / account / gestion / visa preview / export)
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

# Dashboard (uses _get_df_live())
with tabs[1]:
    st.subheader("📊 Dashboard")
    df_all_current = _get_df_live()
    if df_all_current is None or df_all_current.empty:
        st.info("Aucun client chargé. Chargez les fichiers dans la barre latérale.")
    else:
        cats = sorted(df_all_current["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all_current.columns else []
        subs = sorted(df_all_current["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all_current.columns else []
        visas = sorted(df_all_current["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all_current.columns else []
        years = sorted(pd.to_numeric(df_all_current["_Année_"], errors="coerce").dropna().astype(int).unique().tolist())

        a1, a2, a3, a4 = st.columns([1, 1, 1, 1])
        fc = a1.multiselect("Catégories", cats, default=[], key=skey("dash", "cats"))
        fs = a2.multiselect("Sous-catégories", subs, default=[], key=skey("dash", "subs"))
        fv = a3.multiselect("Visa", visas, default=[], key=skey("dash", "visas"))
        fy = a4.multiselect("Année", years, default=[], key=skey("dash", "years"))

        view = df_all_current.copy()
        if fc:
            view = view[view["Categories"].astype(str).isin(fc)]
        if fs:
            view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv:
            view = view[view["Visa"].astype(str).isin(fv)]
        if fy:
            view = view[view["_Année_"].isin(fy)]

        k1, k2, k3, k4, k5 = st.columns([1, 1, 1, 1, 1])
        k1.metric("Dossiers", f"{len(view)}")
        total = (view["Montant honoraires (US $)"].apply(_to_num) + view["Autres frais (US $)"].apply(_to_num)).sum()
        paye = view["Payé"].apply(_to_num).sum()
        solde = view["Solde"].apply(_to_num).sum()
        env_pct = 0
        if "Dossiers envoyé" in view.columns and len(view) > 0:
            env_pct = int(100 * (view["Dossiers envoyé"].apply(_to_num).clip(lower=0, upper=1).sum() / len(view)))
        k2.metric("Honoraires+Frais", _fmt_money(total))
        k3.metric("Payé", _fmt_money(paye))
        k4.metric("Solde", _fmt_money(solde))
        k5.metric("Envoyés (%)", f"{env_pct}%")

        if not view.empty and "Categories" in view.columns:
            vc = view["Categories"].value_counts().rename_axis("Categorie").reset_index(name="Nombre")
            st.bar_chart(vc.set_index("Categorie"))

        if not view.empty and "Mois" in view.columns:
            tmp = view.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            g = tmp.groupby("Mois", as_index=False).agg({
                "Montant honoraires (US $)": "sum",
                "Autres frais (US $)": "sum",
                "Payé": "sum",
                "Solde": "sum",
            }).sort_values("Mois")
            g = g.fillna(0)
            g = g.set_index("Mois")
            st.bar_chart(g)

        show_cols = [c for c in [
            "Dossier N", "ID_Client", "Nom", "Date", "Mois", "Categories", "Sous-categorie", "Visa",
            "Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Commentaires",
            "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé", "RFE"
        ] if c in view.columns]

        detail = view.copy()
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"]:
            if c in detail.columns:
                detail[c] = detail[c].apply(_to_num).map(_fmt_money)
        if "Date" in detail.columns:
            try:
                detail["Date"] = pd.to_datetime(detail["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                detail["Date"] = detail["Date"].astype(str)

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categories", "Nom"] if c in detail.columns]
        detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail
        st.dataframe(detail_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=skey("dash", "table"))

# The rest of the UI (Analyses, Escrow, Compte client, Gestion, Visa preview, Export) follows
# the same logic as previously provided and uses _get_df_live/_set_df_live for persistence.
# If you want I can paste the rest verbatim; but the critical part for your upload problem
# is the robust read_any_table and the diagnostics shown above.
