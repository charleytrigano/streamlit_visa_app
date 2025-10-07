# =========================
# VISA APP — PARTIE 1/2
# =========================
from __future__ import annotations
import json, io, os, re
from pathlib import Path
from datetime import date, datetime
from typing import Tuple, List, Dict, Any

import pandas as pd
import streamlit as st
import altair as alt

# ---------- Constantes colonnes / libellés ----------
DOSSIER_COL = "Dossier N"
HONO = "Montant honoraires (US $)"
AUTRE = "Autres Frais (US $)"
TOTAL = "Total (US $)"

# Statuts + dates associées
S_ENVOYE, D_ENVOYE = "Dossier envoyé", "Date envoyé"
S_APPROUVE, D_APPROUVE = "Dossier approuvé", "Date approuvé"
S_RFE, D_RFE = "RFE", "Date RFE"
S_REFUSE, D_REFUSE = "Dossier refusé", "Date refusé"
S_ANNULE, D_ANNULE = "Dossier annulé", "Date annulé"
STATUS_COLS  = [S_ENVOYE, S_APPROUVE, S_RFE, S_REFUSE, S_ANNULE]
STATUS_DATES = [D_ENVOYE, D_APPROUVE, D_RFE, D_REFUSE, D_ANNULE]

# ESCROW
ESC_TR = "ESCROW transféré (US $)"
ESC_JR = "Journal ESCROW"   # JSON list [{"ts": "...", "amount": float, "note": ""}, ...]

# Démarrage numérotation dossier
DOSSIER_START = 13057

# ---------- État persistant (dernier fichier utilisé) ----------
STATE_FILE = Path(".visa_app_state.json")

def _load_last_path() -> Path | None:
    try:
        if STATE_FILE.exists():
            data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            p = Path(data.get("last_path",""))
            return p if p.exists() else None
    except Exception:
        pass
    return None

def _save_last_path(p: Path) -> None:
    try:
        STATE_FILE.write_text(json.dumps({"last_path": str(p)}, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

def save_workspace_path(p: Path) -> None:
    st.session_state["current_path"] = str(p)
    _save_last_path(p)

def current_file_path() -> Path | None:
    p = st.session_state.get("current_path")
    if p:
        pth = Path(p)
        if pth.exists():
            return pth
    return _load_last_path()

# ---------- Format & conversions ----------
def _safe_str(x) -> str:
    try:
        s = "" if x is None else str(x)
        return s.strip()
    except Exception:
        return ""

def _fmt_money_us(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def _to_num(s: pd.Series) -> pd.Series:
    """Nettoie les nombres (US/EU) -> float"""
    if s is None:
        return pd.Series(dtype=float)
    s = s.astype(str)
    # enlève $ et espaces
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)
    # si on détecte 2 séparateurs, assume que le point est décimal US
    def _clean_one(v: str) -> float:
        if v == "" or v == "-":
            return 0.0
        # cas 1 234,56 (EU)
        if v.count(",")==1 and v.count(".")==0:
            v = v.replace(",", ".")
        # cas 1,234.56 (US) -> retire les virgules milliers
        if v.count(".")==1 and v.count(",")>=1:
            v = v.replace(",", "")
        try:
            return float(v)
        except Exception:
            return 0.0
    return s.map(_clean_one)

def _to_int(s: pd.Series) -> pd.Series:
    try:
        return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)
    except Exception:
        return pd.Series([0]*len(s), dtype=int)

# ---------- Paiements (JSON en cellule) ----------
def _parse_json_list(val: str | list | None) -> list:
    if val is None:
        return []
    if isinstance(val, list):
        return val
    try:
        out = json.loads(val)
        return out if isinstance(out, list) else []
    except Exception:
        return []

def _sum_payments(lst: list[dict]) -> float:
    total = 0.0
    for e in lst:
        try:
            total += float(e.get("amount", 0.0))
        except Exception:
            pass
    return total

# ---------- IO Excel ----------
def list_sheets(path: Path) -> list[str]:
    try:
        xls = pd.ExcelFile(path)
        return xls.sheet_names
    except Exception:
        return []

def read_sheet(path: Path, sheet: str, normalize: bool = False, visa_ref: pd.DataFrame | None = None) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()
    if normalize:
        return normalize_dataframe(df, visa_ref=visa_ref)
    return df

def write_sheet_inplace(path: Path, sheet: str, df: pd.DataFrame):
    """Écrit la feuille sheet en conservant les autres feuilles ; si sheet n'existe pas, elle est créée."""
    path = Path(path)
    try:
        if path.exists():
            book = pd.ExcelFile(path)
            sheets = {sn: pd.read_excel(path, sheet_name=sn) for sn in book.sheet_names}
        else:
            sheets = {}
        sheets[sheet] = df
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for sn, sdf in sheets.items():
                sdf.to_excel(writer, sheet_name=sn, index=False)
    except Exception as e:
        st.error(f"Erreur à l’écriture: {e}")
        raise

def ensure_dossier_numbers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if DOSSIER_COL not in df.columns:
        df[DOSSIER_COL] = 0
    nums = _to_int(df[DOSSIER_COL])
    # Si tableau semble neuf, démarre à DOSSIER_START
    if (nums == 0).all():
        start = DOSSIER_START
        df[DOSSIER_COL] = [start + i for i in range(len(df))]
        return df
    # Sinon, remplit les manquants à partir du max existant
    maxn = int(nums.max())
    for i in range(len(df)):
        if nums.iat[i] <= 0:
            maxn += 1
            df.at[i, DOSSIER_COL] = maxn
    return df

def next_dossier_number(df: pd.DataFrame) -> int:
    if df is None or df.empty or DOSSIER_COL not in df.columns:
        return DOSSIER_START
    nums = _to_int(df[DOSSIER_COL])
    m = int(nums.max()) if len(nums) else DOSSIER_START - 1
    if m < DOSSIER_START - 1:
        m = DOSSIER_START - 1
    return m + 1

def set_current_file_from_upload(up_file) -> Path | None:
    """Sauvegarde un upload en fichier physique à côté de l’app et le sélectionne."""
    if up_file is None:
        return None
    name = up_file.name or "donnees_visa_clients.xlsx"
    buf = up_file.getvalue() if hasattr(up_file, "getvalue") else up_file.read()
    path = Path(name).resolve()
    try:
        with open(path, "wb") as f:
            f.write(buf)
        save_workspace_path(path)
        return path
    except Exception as e:
        st.error(f"Impossible d’enregistrer le fichier uploadé: {e}")
        return None

# ---------- Normalisation / mapping Visa ----------
def read_visa_reference(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=["Catégorie","Visa"])
    # Min: 2 colonnes
    for col in ["Catégorie","Visa"]:
        if col not in df.columns:
            df[col] = ""
    df["Catégorie"] = df["Catégorie"].fillna("").astype(str).str.strip()
    df["Visa"] = df["Visa"].fillna("").astype(str).str.strip()
    return df[["Catégorie","Visa"]].copy()

def map_category_from_ref(df_ref: pd.DataFrame, visa: str) -> str:
    if df_ref is None or df_ref.empty:
        return ""
    v = _safe_str(visa)
    row = df_ref[df_ref["Visa"].astype(str).str.lower() == v.lower()]
    if len(row) == 0:
        return ""
    return _safe_str(row.iloc[0]["Catégorie"])

# Détection d’une feuille de référence (Visa)
def looks_like_reference(df: pd.DataFrame) -> bool:
    if df is None or df.empty:
        return False
    cols = [c.lower() for c in df.columns.astype(str)]
    return ("catégorie" in cols or "categorie" in cols) and ("visa" in cols)

def _make_client_id_from_row(row: dict) -> str:
    # ID client basé sur Nom + Date + hash court
    nom = _safe_str(row.get("Nom"))
    try:
        d = pd.to_datetime(row.get("Date")).date()
    except Exception:
        d = date.today()
    tel = _safe_str(row.get("Téléphone",""))  # (ancien champ optionnel, ignoré si absent)
    base = f"{nom}-{tel}-{d.strftime('%Y%m%d')}"
    # simplifie l’ID (alphanum & tirets)
    base = re.sub(r"[^A-Za-z0-9\-]+", "", base.replace(" ", "-"))
    return base.lower()

def normalize_dataframe(df: pd.DataFrame, visa_ref: pd.DataFrame | None = None) -> pd.DataFrame:
    """Nettoie les champs, calcule Total/Payé/Reste, Date/Mois (MM), map Catégorie si vide."""
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    # Renommages souples (compat rétro)
    rename = {}
    for c in df.columns:
        lc = str(c).lower().strip()
        if lc in ("montant honoraires", "montant honoraires (us $)", "honoraires", "montant"):
            rename[c] = HONO
        elif lc in ("autres frais", "autres frais (us $)", "autres"):
            rename[c] = AUTRE
        elif lc in ("total", "total (us $)"):
            rename[c] = TOTAL
        elif lc == "dossier n" or lc == "dossier":
            rename[c] = DOSSIER_COL
        elif lc == "reste (us $)" or lc == "solde (us $)" or lc == "solde":
            rename[c] = "Reste"
        elif lc == "paye (us $)" or lc == "payé (us $)" or lc == "paye" or lc == "payé":
            rename[c] = "Payé"
    if rename:
        df = df.rename(columns=rename)

    # Colonnes minimales
    for c in [DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", HONO, AUTRE, TOTAL, "Payé", "Reste", "Paiements", "Date", "Mois"]:
        if c not in df.columns:
            if c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                df[c] = 0.0
            elif c == "Paiements":
                df[c] = ""
            else:
                df[c] = ""

    # Nettoyage numériques
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = _to_num(df[c])

    # Date -> date + Mois (MM)
    def _to_date(x):
        try:
            if pd.isna(x) or x == "":
                return pd.NaT
            return pd.to_datetime(x).date()
        except Exception:
            return pd.NaT
    df["Date"] = df["Date"].map(_to_date)
    df["Mois"] = df["Date"].apply(lambda d: f"{d.month:02d}" if pd.notna(d) else pd.NA)

    # Total = honoraires + autres
    df[TOTAL] = _to_num(df.get(HONO, 0.0)) + _to_num(df.get(AUTRE, 0.0))

    # Payé à partir des paiements JSON s'il existe
    paid_from_json = []
    for _, r in df.iterrows():
        plist = _parse_json_list(r.get("Paiements", ""))
        paid_from_json.append(_sum_payments(plist))
    paid_from_json = pd.Series(paid_from_json, index=df.index, dtype=float)
    # Si "Payé" existant est 0 et JSON > 0 => prend JSON
    df["Payé"] = pd.Series([max(a, b) for a, b in zip(_to_num(df["Payé"]), paid_from_json)], index=df.index)

    # Reste
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # Catégorie depuis référentiel si manquante
    if visa_ref is not None and not visa_ref.empty:
        mask_cat_missing = (df["Catégorie"].astype(str).str.strip() == "")
        if mask_cat_missing.any():
            df.loc[mask_cat_missing, "Catégorie"] = df.loc[mask_cat_missing, "Visa"].apply(lambda v: map_category_from_ref(visa_ref, v))

    # Statuts: s'assure présence
    for b in STATUS_COLS:
        if b not in df.columns:
            df[b] = False
        else:
            df[b] = df[b].astype(bool)
    for dcol in STATUS_DATES:
        if dcol not in df.columns:
            df[dcol] = ""

    # ESCROW
    if ESC_TR not in df.columns:
        df[ESC_TR] = 0.0
    df[ESC_TR] = _to_num(df[ESC_TR])
    if ESC_JR not in df.columns:
        df[ESC_JR] = ""

    # Dossier N
    df = ensure_dossier_numbers(df)

    return df

# ---------- Hiérarchie Catégorie / Sous-catégories / Visa ----------
def read_visa_reference_hier(path: Path) -> pd.DataFrame:
    """Lit la feuille Visa telle quelle, en conservant toutes les colonnes hiérarchiques."""
    try:
        dfv = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=["Catégorie","Visa"])
    def _norm(c):
        c = str(c).strip()
        c_low = (c.lower()
                   .replace("é","e").replace("è","e").replace("ê","e")
                   .replace("à","a").replace("ô","o").replace("ï","i").replace("ç","c"))
        return c, c_low
    rename = {}
    for c in dfv.columns:
        raw, low = _norm(c)
        if low in ("categorie","catégorie"): rename[c] = "Catégorie"
        elif low.startswith("sous-categorie") or low.startswith("sous-catégorie") or low.startswith("sous categorie"):
            rename[c] = " ".join([w.capitalize() for w in raw.split()])  # ex: "Sous-catégorie 2"
        elif low == "visa":
            rename[c] = "Visa"
    if rename:
        dfv = dfv.rename(columns=rename)
    for c in dfv.columns:
        if dfv[c].dtype == "object":
            dfv[c] = dfv[c].fillna("").astype(str).str.strip()
    return dfv

def get_hierarchy_columns(df_ref: pd.DataFrame) -> list[str]:
    """Retourne l’ordre des colonnes hiérarchiques : [Catégorie, Sous-catégorie 1..n, Visa?]."""
    if df_ref is None or df_ref.empty:
        return ["Catégorie","Visa"]
    cols = [c for c in df_ref.columns if isinstance(c, str)]
    out = []
    if "Catégorie" in cols: out.append("Catégorie")
    subs = []
    for c in cols:
        low = c.lower().replace("é","e").replace("è","e")
        if low.startswith("sous-categorie"):
            subs.append(c)
    def _num_key(c):
        m = re.search(r"(\d+)", c)
        return int(m.group(1)) if m else 999
    subs = sorted(subs, key=_num_key)
    out.extend(subs)
    if "Visa" in cols: out.append("Visa")
    return out

def cascading_visa_picker(df_ref: pd.DataFrame, key_prefix: str, init: dict | None = None) -> dict:
    """Affiche les selectbox en cascade. Renvoie {col: valeur}. 'Visa' peut être '' si la branche n'a pas de visa."""
    if df_ref is None or df_ref.empty:
        st.info("Référentiel Visa vide."); return {"Catégorie":"", "Visa":""}
    cols = get_hierarchy_columns(df_ref)
    sel = {}
    df_work = df_ref.copy()
    for col in cols:
        # filtre sur les niveaux précédents
        for prev_col, prev_val in sel.items():
            if prev_val != "":
                df_work = df_work[df_work[prev_col].astype(str) == prev_val]
        options = sorted([v for v in df_work[col].astype(str).unique() if v != ""])
        if not options:
            if col == "Visa": sel[col] = ""
            break
        default_idx = 0
        if init and init.get(col, "") in options:
            default_idx = options.index(init[col])
        sel[col] = st.selectbox(col, [""] + options, index=default_idx+1 if options else 0, key=f"{key_prefix}_{col}")
        if sel[col] == "":
            for c2 in cols[cols.index(col)+1:]:
                sel[c2] = ""
            break
    for c in cols:
        sel.setdefault(c, "")
    sel.setdefault("Catégorie",""); sel.setdefault("Visa","")
    return sel

def visas_autorises_depuis_cascade(df_ref_full: pd.DataFrame, sel_path: dict) -> list[str]:
    """Calcule la liste des visas autorisés à partir de la sélection hiérarchique."""
    if df_ref_full is None or df_ref_full.empty:
        return []
    cols = get_hierarchy_columns(df_ref_full)
    dfw = df_ref_full.copy()
    for c in cols:
        val = sel_path.get(c, "")
        if val:
            dfw = dfw[dfw[c].astype(str) == val]
    if "Visa" not in dfw.columns:
        return []
    visas = sorted([v for v in dfw["Visa"].astype(str).unique() if v != ""])
    return visas

# ---------- ESCROW helpers ----------
def escrow_available_from_row(row: pd.Series) -> float:
    """Disponible à transférer depuis ESCROW (honoraires payés - déjà transféré)."""
    hono = float(pd.to_numeric(pd.Series([row.get(HONO, 0.0)]), errors="coerce").fillna(0.0).iloc[0])
    paid = float(pd.to_numeric(pd.Series([row.get("Payé", 0.0)]), errors="coerce").fillna(0.0).iloc[0])
    transferred = float(pd.to_numeric(pd.Series([row.get(ESC_TR, 0.0)]), errors="coerce").fillna(0.0).iloc[0])
    dispo = max(min(paid, hono) - transferred, 0.0)
    return float(dispo)

def append_escrow_journal(row: pd.Series, amount: float, note: str = "") -> str:
    lst = _parse_json_list(row.get(ESC_JR, ""))
    lst.append({"ts": datetime.now().isoformat(timespec="seconds"), "amount": float(amount), "note": _safe_str(note)})
    return json.dumps(lst, ensure_ascii=False)


# =========================
# VISA APP — PARTIE 1/2
# =========================
from __future__ import annotations
import json, io, os, re
from pathlib import Path
from datetime import date, datetime
from typing import Tuple, List, Dict, Any

import pandas as pd
import streamlit as st
import altair as alt

# ---------- Constantes colonnes / libellés ----------
DOSSIER_COL = "Dossier N"
HONO = "Montant honoraires (US $)"
AUTRE = "Autres Frais (US $)"
TOTAL = "Total (US $)"

# Statuts + dates associées
S_ENVOYE, D_ENVOYE = "Dossier envoyé", "Date envoyé"
S_APPROUVE, D_APPROUVE = "Dossier approuvé", "Date approuvé"
S_RFE, D_RFE = "RFE", "Date RFE"
S_REFUSE, D_REFUSE = "Dossier refusé", "Date refusé"
S_ANNULE, D_ANNULE = "Dossier annulé", "Date annulé"
STATUS_COLS  = [S_ENVOYE, S_APPROUVE, S_RFE, S_REFUSE, S_ANNULE]
STATUS_DATES = [D_ENVOYE, D_APPROUVE, D_RFE, D_REFUSE, D_ANNULE]

# ESCROW
ESC_TR = "ESCROW transféré (US $)"
ESC_JR = "Journal ESCROW"   # JSON list [{"ts": "...", "amount": float, "note": ""}, ...]

# Démarrage numérotation dossier
DOSSIER_START = 13057

# ---------- État persistant (dernier fichier utilisé) ----------
STATE_FILE = Path(".visa_app_state.json")

def _load_last_path() -> Path | None:
    try:
        if STATE_FILE.exists():
            data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            p = Path(data.get("last_path",""))
            return p if p.exists() else None
    except Exception:
        pass
    return None

def _save_last_path(p: Path) -> None:
    try:
        STATE_FILE.write_text(json.dumps({"last_path": str(p)}, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

def save_workspace_path(p: Path) -> None:
    st.session_state["current_path"] = str(p)
    _save_last_path(p)

def current_file_path() -> Path | None:
    p = st.session_state.get("current_path")
    if p:
        pth = Path(p)
        if pth.exists():
            return pth
    return _load_last_path()

# ---------- Format & conversions ----------
def _safe_str(x) -> str:
    try:
        s = "" if x is None else str(x)
        return s.strip()
    except Exception:
        return ""

def _fmt_money_us(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def _to_num(s: pd.Series) -> pd.Series:
    """Nettoie les nombres (US/EU) -> float"""
    if s is None:
        return pd.Series(dtype=float)
    s = s.astype(str)
    # enlève $ et espaces
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)
    # si on détecte 2 séparateurs, assume que le point est décimal US
    def _clean_one(v: str) -> float:
        if v == "" or v == "-":
            return 0.0
        # cas 1 234,56 (EU)
        if v.count(",")==1 and v.count(".")==0:
            v = v.replace(",", ".")
        # cas 1,234.56 (US) -> retire les virgules milliers
        if v.count(".")==1 and v.count(",")>=1:
            v = v.replace(",", "")
        try:
            return float(v)
        except Exception:
            return 0.0
    return s.map(_clean_one)

def _to_int(s: pd.Series) -> pd.Series:
    try:
        return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)
    except Exception:
        return pd.Series([0]*len(s), dtype=int)

# ---------- Paiements (JSON en cellule) ----------
def _parse_json_list(val: str | list | None) -> list:
    if val is None:
        return []
    if isinstance(val, list):
        return val
    try:
        out = json.loads(val)
        return out if isinstance(out, list) else []
    except Exception:
        return []

def _sum_payments(lst: list[dict]) -> float:
    total = 0.0
    for e in lst:
        try:
            total += float(e.get("amount", 0.0))
        except Exception:
            pass
    return total

# ---------- IO Excel ----------
def list_sheets(path: Path) -> list[str]:
    try:
        xls = pd.ExcelFile(path)
        return xls.sheet_names
    except Exception:
        return []

def read_sheet(path: Path, sheet: str, normalize: bool = False, visa_ref: pd.DataFrame | None = None) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()
    if normalize:
        return normalize_dataframe(df, visa_ref=visa_ref)
    return df

def write_sheet_inplace(path: Path, sheet: str, df: pd.DataFrame):
    """Écrit la feuille sheet en conservant les autres feuilles ; si sheet n'existe pas, elle est créée."""
    path = Path(path)
    try:
        if path.exists():
            book = pd.ExcelFile(path)
            sheets = {sn: pd.read_excel(path, sheet_name=sn) for sn in book.sheet_names}
        else:
            sheets = {}
        sheets[sheet] = df
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for sn, sdf in sheets.items():
                sdf.to_excel(writer, sheet_name=sn, index=False)
    except Exception as e:
        st.error(f"Erreur à l’écriture: {e}")
        raise

def ensure_dossier_numbers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if DOSSIER_COL not in df.columns:
        df[DOSSIER_COL] = 0
    nums = _to_int(df[DOSSIER_COL])
    # Si tableau semble neuf, démarre à DOSSIER_START
    if (nums == 0).all():
        start = DOSSIER_START
        df[DOSSIER_COL] = [start + i for i in range(len(df))]
        return df
    # Sinon, remplit les manquants à partir du max existant
    maxn = int(nums.max())
    for i in range(len(df)):
        if nums.iat[i] <= 0:
            maxn += 1
            df.at[i, DOSSIER_COL] = maxn
    return df

def next_dossier_number(df: pd.DataFrame) -> int:
    if df is None or df.empty or DOSSIER_COL not in df.columns:
        return DOSSIER_START
    nums = _to_int(df[DOSSIER_COL])
    m = int(nums.max()) if len(nums) else DOSSIER_START - 1
    if m < DOSSIER_START - 1:
        m = DOSSIER_START - 1
    return m + 1

def set_current_file_from_upload(up_file) -> Path | None:
    """Sauvegarde un upload en fichier physique à côté de l’app et le sélectionne."""
    if up_file is None:
        return None
    name = up_file.name or "donnees_visa_clients.xlsx"
    buf = up_file.getvalue() if hasattr(up_file, "getvalue") else up_file.read()
    path = Path(name).resolve()
    try:
        with open(path, "wb") as f:
            f.write(buf)
        save_workspace_path(path)
        return path
    except Exception as e:
        st.error(f"Impossible d’enregistrer le fichier uploadé: {e}")
        return None

# ---------- Normalisation / mapping Visa ----------
def read_visa_reference(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=["Catégorie","Visa"])
    # Min: 2 colonnes
    for col in ["Catégorie","Visa"]:
        if col not in df.columns:
            df[col] = ""
    df["Catégorie"] = df["Catégorie"].fillna("").astype(str).str.strip()
    df["Visa"] = df["Visa"].fillna("").astype(str).str.strip()
    return df[["Catégorie","Visa"]].copy()

def map_category_from_ref(df_ref: pd.DataFrame, visa: str) -> str:
    if df_ref is None or df_ref.empty:
        return ""
    v = _safe_str(visa)
    row = df_ref[df_ref["Visa"].astype(str).str.lower() == v.lower()]
    if len(row) == 0:
        return ""
    return _safe_str(row.iloc[0]["Catégorie"])

# Détection d’une feuille de référence (Visa)
def looks_like_reference(df: pd.DataFrame) -> bool:
    if df is None or df.empty:
        return False
    cols = [c.lower() for c in df.columns.astype(str)]
    return ("catégorie" in cols or "categorie" in cols) and ("visa" in cols)

def _make_client_id_from_row(row: dict) -> str:
    # ID client basé sur Nom + Date + hash court
    nom = _safe_str(row.get("Nom"))
    try:
        d = pd.to_datetime(row.get("Date")).date()
    except Exception:
        d = date.today()
    tel = _safe_str(row.get("Téléphone",""))  # (ancien champ optionnel, ignoré si absent)
    base = f"{nom}-{tel}-{d.strftime('%Y%m%d')}"
    # simplifie l’ID (alphanum & tirets)
    base = re.sub(r"[^A-Za-z0-9\-]+", "", base.replace(" ", "-"))
    return base.lower()

def normalize_dataframe(df: pd.DataFrame, visa_ref: pd.DataFrame | None = None) -> pd.DataFrame:
    """Nettoie les champs, calcule Total/Payé/Reste, Date/Mois (MM), map Catégorie si vide."""
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    # Renommages souples (compat rétro)
    rename = {}
    for c in df.columns:
        lc = str(c).lower().strip()
        if lc in ("montant honoraires", "montant honoraires (us $)", "honoraires", "montant"):
            rename[c] = HONO
        elif lc in ("autres frais", "autres frais (us $)", "autres"):
            rename[c] = AUTRE
        elif lc in ("total", "total (us $)"):
            rename[c] = TOTAL
        elif lc == "dossier n" or lc == "dossier":
            rename[c] = DOSSIER_COL
        elif lc == "reste (us $)" or lc == "solde (us $)" or lc == "solde":
            rename[c] = "Reste"
        elif lc == "paye (us $)" or lc == "payé (us $)" or lc == "paye" or lc == "payé":
            rename[c] = "Payé"
    if rename:
        df = df.rename(columns=rename)

    # Colonnes minimales
    for c in [DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", HONO, AUTRE, TOTAL, "Payé", "Reste", "Paiements", "Date", "Mois"]:
        if c not in df.columns:
            if c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                df[c] = 0.0
            elif c == "Paiements":
                df[c] = ""
            else:
                df[c] = ""

    # Nettoyage numériques
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = _to_num(df[c])

    # Date -> date + Mois (MM)
    def _to_date(x):
        try:
            if pd.isna(x) or x == "":
                return pd.NaT
            return pd.to_datetime(x).date()
        except Exception:
            return pd.NaT
    df["Date"] = df["Date"].map(_to_date)
    df["Mois"] = df["Date"].apply(lambda d: f"{d.month:02d}" if pd.notna(d) else pd.NA)

    # Total = honoraires + autres
    df[TOTAL] = _to_num(df.get(HONO, 0.0)) + _to_num(df.get(AUTRE, 0.0))

    # Payé à partir des paiements JSON s'il existe
    paid_from_json = []
    for _, r in df.iterrows():
        plist = _parse_json_list(r.get("Paiements", ""))
        paid_from_json.append(_sum_payments(plist))
    paid_from_json = pd.Series(paid_from_json, index=df.index, dtype=float)
    # Si "Payé" existant est 0 et JSON > 0 => prend JSON
    df["Payé"] = pd.Series([max(a, b) for a, b in zip(_to_num(df["Payé"]), paid_from_json)], index=df.index)

    # Reste
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # Catégorie depuis référentiel si manquante
    if visa_ref is not None and not visa_ref.empty:
        mask_cat_missing = (df["Catégorie"].astype(str).str.strip() == "")
        if mask_cat_missing.any():
            df.loc[mask_cat_missing, "Catégorie"] = df.loc[mask_cat_missing, "Visa"].apply(lambda v: map_category_from_ref(visa_ref, v))

    # Statuts: s'assure présence
    for b in STATUS_COLS:
        if b not in df.columns:
            df[b] = False
        else:
            df[b] = df[b].astype(bool)
    for dcol in STATUS_DATES:
        if dcol not in df.columns:
            df[dcol] = ""

    # ESCROW
    if ESC_TR not in df.columns:
        df[ESC_TR] = 0.0
    df[ESC_TR] = _to_num(df[ESC_TR])
    if ESC_JR not in df.columns:
        df[ESC_JR] = ""

    # Dossier N
    df = ensure_dossier_numbers(df)

    return df

# ---------- Hiérarchie Catégorie / Sous-catégories / Visa ----------
def read_visa_reference_hier(path: Path) -> pd.DataFrame:
    """Lit la feuille Visa telle quelle, en conservant toutes les colonnes hiérarchiques."""
    try:
        dfv = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=["Catégorie","Visa"])
    def _norm(c):
        c = str(c).strip()
        c_low = (c.lower()
                   .replace("é","e").replace("è","e").replace("ê","e")
                   .replace("à","a").replace("ô","o").replace("ï","i").replace("ç","c"))
        return c, c_low
    rename = {}
    for c in dfv.columns:
        raw, low = _norm(c)
        if low in ("categorie","catégorie"): rename[c] = "Catégorie"
        elif low.startswith("sous-categorie") or low.startswith("sous-catégorie") or low.startswith("sous categorie"):
            rename[c] = " ".join([w.capitalize() for w in raw.split()])  # ex: "Sous-catégorie 2"
        elif low == "visa":
            rename[c] = "Visa"
    if rename:
        dfv = dfv.rename(columns=rename)
    for c in dfv.columns:
        if dfv[c].dtype == "object":
            dfv[c] = dfv[c].fillna("").astype(str).str.strip()
    return dfv

def get_hierarchy_columns(df_ref: pd.DataFrame) -> list[str]:
    """Retourne l’ordre des colonnes hiérarchiques : [Catégorie, Sous-catégorie 1..n, Visa?]."""
    if df_ref is None or df_ref.empty:
        return ["Catégorie","Visa"]
    cols = [c for c in df_ref.columns if isinstance(c, str)]
    out = []
    if "Catégorie" in cols: out.append("Catégorie")
    subs = []
    for c in cols:
        low = c.lower().replace("é","e").replace("è","e")
        if low.startswith("sous-categorie"):
            subs.append(c)
    def _num_key(c):
        m = re.search(r"(\d+)", c)
        return int(m.group(1)) if m else 999
    subs = sorted(subs, key=_num_key)
    out.extend(subs)
    if "Visa" in cols: out.append("Visa")
    return out

def cascading_visa_picker(df_ref: pd.DataFrame, key_prefix: str, init: dict | None = None) -> dict:
    """Affiche les selectbox en cascade. Renvoie {col: valeur}. 'Visa' peut être '' si la branche n'a pas de visa."""
    if df_ref is None or df_ref.empty:
        st.info("Référentiel Visa vide."); return {"Catégorie":"", "Visa":""}
    cols = get_hierarchy_columns(df_ref)
    sel = {}
    df_work = df_ref.copy()
    for col in cols:
        # filtre sur les niveaux précédents
        for prev_col, prev_val in sel.items():
            if prev_val != "":
                df_work = df_work[df_work[prev_col].astype(str) == prev_val]
        options = sorted([v for v in df_work[col].astype(str).unique() if v != ""])
        if not options:
            if col == "Visa": sel[col] = ""
            break
        default_idx = 0
        if init and init.get(col, "") in options:
            default_idx = options.index(init[col])
        sel[col] = st.selectbox(col, [""] + options, index=default_idx+1 if options else 0, key=f"{key_prefix}_{col}")
        if sel[col] == "":
            for c2 in cols[cols.index(col)+1:]:
                sel[c2] = ""
            break
    for c in cols:
        sel.setdefault(c, "")
    sel.setdefault("Catégorie",""); sel.setdefault("Visa","")
    return sel

def visas_autorises_depuis_cascade(df_ref_full: pd.DataFrame, sel_path: dict) -> list[str]:
    """Calcule la liste des visas autorisés à partir de la sélection hiérarchique."""
    if df_ref_full is None or df_ref_full.empty:
        return []
    cols = get_hierarchy_columns(df_ref_full)
    dfw = df_ref_full.copy()
    for c in cols:
        val = sel_path.get(c, "")
        if val:
            dfw = dfw[dfw[c].astype(str) == val]
    if "Visa" not in dfw.columns:
        return []
    visas = sorted([v for v in dfw["Visa"].astype(str).unique() if v != ""])
    return visas

# ---------- ESCROW helpers ----------
def escrow_available_from_row(row: pd.Series) -> float:
    """Disponible à transférer depuis ESCROW (honoraires payés - déjà transféré)."""
    hono = float(pd.to_numeric(pd.Series([row.get(HONO, 0.0)]), errors="coerce").fillna(0.0).iloc[0])
    paid = float(pd.to_numeric(pd.Series([row.get("Payé", 0.0)]), errors="coerce").fillna(0.0).iloc[0])
    transferred = float(pd.to_numeric(pd.Series([row.get(ESC_TR, 0.0)]), errors="coerce").fillna(0.0).iloc[0])
    dispo = max(min(paid, hono) - transferred, 0.0)
    return float(dispo)

def append_escrow_journal(row: pd.Series, amount: float, note: str = "") -> str:
    lst = _parse_json_list(row.get(ESC_JR, ""))
    lst.append({"ts": datetime.now().isoformat(timespec="seconds"), "amount": float(amount), "note": _safe_str(note)})
    return json.dumps(lst, ensure_ascii=False)
