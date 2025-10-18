# url=https://github.com/charleytrigano/streamlit_visa_app/blob/main/app.py
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

def read_any_table(src: Any, sheet: Optional[str] = None) -> Optional[pd.DataFrame]:
    """
    Robust reader: accepts UploadedFile, bytes/BytesIO, or file path.
    For UploadedFile we read bytes once into memory and create BytesIO copies to avoid EOF issues.
    """
    if src is None:
        return None

    # If src is a BytesIO or bytes, make sure we have a fresh BytesIO
    if isinstance(src, (bytes, bytearray)):
        bio = BytesIO(src)
        try:
            return pd.read_excel(bio, sheet_name=(sheet if sheet else 0), engine="openpyxl")
        except Exception:
            bio.seek(0)
            try:
                return pd.read_csv(bio)
            except Exception:
                return None
    if isinstance(src, (io.BytesIO, BytesIO)):
        try:
            bio2 = BytesIO(src.getvalue())
            return pd.read_excel(bio2, sheet_name=(sheet if sheet else 0), engine="openpyxl")
        except Exception:
            try:
                src.seek(0)
                return pd.read_csv(src)
            except Exception:
                return None

    # UploadedFile (Streamlit) or file-like with .name and .read
    if hasattr(src, "read") and hasattr(src, "name"):
        # obtain bytes safely (getbuffer/getvalue when available)
        try:
            data = src.getvalue()  # works on UploadedFile
        except Exception:
            try:
                src.seek(0)
            except Exception:
                pass
            try:
                data = src.read()
            except Exception:
                data = None

        if not data:
            return None

        # Always work from a BytesIO (we'll create copies if needed elsewhere)
        bio = BytesIO(data)

        name = getattr(src, "name", "").lower()
        if name.endswith(".csv"):
            # try utf-8 then latin1
            try:
                return pd.read_csv(BytesIO(data), encoding="utf-8", on_bad_lines="skip")
            except Exception:
                try:
                    return pd.read_csv(BytesIO(data), encoding="latin1", on_bad_lines="skip")
                except Exception:
                    return None
        # excel attempt with openpyxl (xlsx/xlsm/xltx etc.), fallback to generic read_excel
        try:
            return pd.read_excel(BytesIO(data), sheet_name=(sheet if sheet else 0), engine="openpyxl")
        except Exception:
            try:
                return pd.read_excel(BytesIO(data), sheet_name=(sheet if sheet else 0))
            except Exception:
                # last resort: try csv parse
                try:
                    return pd.read_csv(BytesIO(data), on_bad_lines="skip")
                except Exception:
                    return None

    # Path on disk
    if isinstance(src, (str, os.PathLike)):
        p = str(src)
        if not os.path.exists(p):
            return None
        if p.lower().endswith(".csv"):
            try:
                return pd.read_csv(p)
            except Exception:
                return None
        try:
            return pd.read_excel(p, sheet_name=(sheet if sheet else 0))
        except Exception:
            return None

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

# Barre latérale
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
    st.success("Chemins mémorisés. Re-lancement pour prise en compte.")
    st.experimental_rerun()

# -----------------------
# Read uploaded files robustly and avoid re-consuming UploadedFile stream
# -----------------------
# If up_clients provided: read bytes once and reuse copies for clients and visa reading when needed
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

# Determine clients_src for reading
if clients_bytes is not None:
    # use BytesIO copy
    clients_src_for_read = BytesIO(clients_bytes)
elif clients_path_in:
    clients_src_for_read = clients_path_in
elif last_clients:
    clients_src_for_read = last_clients
else:
    clients_src_for_read = None

# Determine visa_src for reading
if mode == "Deux fichiers (Clients & Visa)":
    if visa_bytes is not None:
        visa_src_for_read = BytesIO(visa_bytes)
    elif visa_path_in:
        visa_src_for_read = visa_path_in
    else:
        visa_src_for_read = None
else:
    # single file mode: use same bytes as clients (if any)
    if clients_bytes is not None:
        visa_src_for_read = BytesIO(clients_bytes)
    elif clients_path_in:
        visa_src_for_read = clients_path_in
    else:
        visa_src_for_read = None

# Now read
df_clients_raw = normalize_clients(read_any_table(clients_src_for_read))
df_visa_raw = read_any_table(visa_src_for_read, sheet=SHEET_VISA)
if df_visa_raw is None:
    df_visa_raw = read_any_table(visa_src_for_read)
if df_visa_raw is None:
    df_visa_raw = pd.DataFrame()

# Affichage des fichiers chargés
with st.expander("📄 Fichiers chargés", expanded=True):
    st.write("**Clients** :", ("(aucun)" if (df_clients_raw is None or df_clients_raw.empty) else (getattr(up_clients, 'name', str(clients_src_for_read)))))
    st.write("**Visa** :", ("(aucun)" if (df_visa_raw is None or df_visa_raw.empty) else (getattr(up_visa, 'name', str(visa_src_for_read)))))

# Construction de la carte Visa
visa_map = build_visa_map(df_visa_raw)

# Normalisation et time features
df_all = _ensure_time_features(df_clients_raw)

# Persist working copy in session_state
DF_LIVE_KEY = skey("df_live")
if DF_LIVE_KEY not in st.session_state or st.session_state[DF_LIVE_KEY] is None:
    st.session_state[DF_LIVE_KEY] = df_all.copy() if (df_all is not None) else pd.DataFrame()

def _get_df_live() -> pd.DataFrame:
    return st.session_state[DF_LIVE_KEY].copy()

def _set_df_live(df: pd.DataFrame) -> None:
    st.session_state[DF_LIVE_KEY] = df.copy()

# UI tabs (rest of the app unchanged...)
# For brevity in this response the rest of the UI code (dashboard, analyses, escrow, account,
# gestion, visa preview, export) is unchanged and should be appended here exactly as in your app.
# (In your local file, keep the full UI code; only the file-reading section above needed the fix.)
#
# Note: If you want I will paste the whole UI section again (unchanged from your working version).
# For now I keep it out to focus this patch on the root cause (upload/reading).
