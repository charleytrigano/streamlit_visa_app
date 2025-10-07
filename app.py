# =========================
# VISA APP — PARTIE 1/3
# =========================

import json, re
from pathlib import Path
from datetime import date, datetime
from typing import Any

import pandas as pd
import streamlit as st
import altair as alt

# ---------- Constantes colonnes / libellés ----------
DOSSIER_COL = "Dossier N"
HONO = "Montant honoraires (US $)"
AUTRE = "Autres Frais (US $)"
TOTAL = "Total (US $)"

# Statuts + dates associées (ordre demandé)
S_ENVOYE, D_ENVOYE = "Dossier envoyé", "Date envoyé"
S_APPROUVE, D_APPROUVE = "Dossier approuvé", "Date approuvé"
S_RFE, D_RFE = "RFE", "Date RFE"
S_REFUSE, D_REFUSE = "Dossier refusé", "Date refusé"
S_ANNULE, D_ANNULE = "Dossier annulé", "Date annulé"
STATUS_COLS  = [S_ENVOYE, S_APPROUVE, S_RFE, S_REFUSE, S_ANNULE]
STATUS_DATES = [D_ENVOYE, D_APPROUVE, D_RFE, D_REFUSE, D_ANNULE]

# ESCROW
ESC_TR = "ESCROW transféré (US $)"
ESC_JR = "Journal ESCROW"   # JSON [{"ts": "...", "amount": float, "note": ""}, ...]

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
    """Convertit une Series (ou 1ère colonne d’un DataFrame) en float propre.
       Gère les cas où df[col] retourne un DataFrame (colonnes dupliquées)."""
    if s is None:
        return pd.Series(dtype=float)
    if isinstance(s, pd.DataFrame):
        if s.shape[1] == 0:
            return pd.Series(dtype=float, index=s.index if hasattr(s, "index") else None)
        s = s.iloc[:, 0]  # prend la 1ère colonne
    s = s.astype(str)
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)  # supprime symboles
    def _clean_one(v: str) -> float:
        if v == "" or v == "-":
            return 0.0
        # format EU 1 234,56 -> 1234.56
        if v.count(",")==1 and v.count(".")==0:
            v = v.replace(",", ".")
        # format US 1,234.56 -> 1234.56
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
def _parse_json_list(val: Any) -> list:
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

def set_current_file_from_upload(up_file) -> Path | None:
    """Sauvegarde un upload en fichier physique et le sélectionne comme fichier courant."""
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

# ---------- Normalisation / mapping Visa simple (Catégorie-Visa) ----------
def read_visa_reference(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=["Catégorie","Visa"])
    for col in ["Catégorie","Visa"]:
        if col not in df.columns:
            df[col] = ""
    df["Catégorie"] = df["Catégorie"].fillna("").astype(str).str.strip()
    df["Visa"] = df["Visa"].fillna("").astype(str).str.strip()
    return df[["Catégorie","Visa"]].copy()

def looks_like_reference(df: pd.DataFrame) -> bool:
    if df is None or df.empty:
        return False
    cols = [c.lower() for c in df.columns.astype(str)]
    return ("catégorie" in cols or "categorie" in cols) and ("visa" in cols)

def map_category_from_ref(df_ref: pd.DataFrame, visa: str) -> str:
    if df_ref is None or df_ref.empty:
        return ""
    v = _safe_str(visa)
    row = df_ref[df_ref["Visa"].astype(str).str.lower() == v.lower()]
    if len(row) == 0:
        return ""
    return _safe_str(row.iloc[0]["Catégorie"])

def ensure_dossier_numbers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if DOSSIER_COL not in df.columns:
        df[DOSSIER_COL] = 0
    nums = _to_int(df[DOSSIER_COL])
    if (nums == 0).all():
        start = DOSSIER_START
        df[DOSSIER_COL] = [start + i for i in range(len(df))]
        return df
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

def _make_client_id_from_row(row: dict) -> str:
    # ID client basé sur Nom + Date
    nom = _safe_str(row.get("Nom"))
    try:
        d = pd.to_datetime(row.get("Date")).date()
    except Exception:
        d = date.today()
    base = f"{nom}-{d.strftime('%Y%m%d')}"
    base = re.sub(r"[^A-Za-z0-9\-]+", "", base.replace(" ", "-"))
    return base.lower()

# ---------- Fusion des colonnes dupliquées ----------
def _collapse_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Fusionne les colonnes dupliquées (même nom).
       Numériques -> somme ligne par ligne. Sinon -> 1ère valeur non vide."""
    if df is None or df.empty:
        return df
    cols = df.columns.astype(str)
    if not cols.duplicated().any():
        return df

    out = pd.DataFrame(index=df.index)
    for col in pd.unique(cols):
        same = df.loc[:, cols == col]
        if same.shape[1] == 1:
            out[col] = same.iloc[:, 0]
            continue
        # Essai: tout convertir en numérique et sommer si possible
        try:
            same_num = same.apply(pd.to_numeric, errors="coerce")
            if same_num.notna().any().any():
                out[col] = same_num.sum(axis=1, skipna=True)
                continue
        except Exception:
            pass
        # Sinon, 1ère valeur non vide
        def _first_non_empty(row):
            for v in row:
                if pd.notna(v) and str(v).strip() != "":
                    return v
            return ""
        out[col] = same.apply(_first_non_empty, axis=1)

    return out

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
        elif lc in ("reste (us $)", "solde (us $)", "solde"):
            rename[c] = "Reste"
        elif lc in ("paye (us $)","payé (us $)","paye","payé"):
            rename[c] = "Payé"
        elif lc == "sous-type" or lc == "soustype" or lc == "sous type":
            rename[c] = "Sous-type"
    if rename:
        df = df.rename(columns=rename)

    # ⚠️ Écrase les colonnes dupliquées après renommage
    df = _collapse_duplicate_columns(df)

    # Colonnes minimales
    base_cols = [DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", "Sous-type",
                 HONO, AUTRE, TOTAL, "Payé", "Reste", "Paiements", "Date", "Mois"]
    for c in base_cols:
        if c not in df.columns:
            if c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                df[c] = 0.0
            elif c == "Paiements":
                df[c] = ""
            else:
                df[c] = ""

    # Numériques
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = _to_num(df[c])

    # Date & Mois
    def _to_date(x):
        try:
            if pd.isna(x) or x == "":
                return pd.NaT
            return pd.to_datetime(x).date()
        except Exception:
            return pd.NaT
    df["Date"] = df["Date"].map(_to_date)
    df["Mois"] = df["Date"].apply(lambda d: f"{d.month:02d}" if pd.notna(d) else pd.NA)

    # Total
    df[TOTAL] = _to_num(df.get(HONO, 0.0)) + _to_num(df.get(AUTRE, 0.0))

    # Payé via JSON si disponible
    paid_from_json = []
    for _, r in df.iterrows():
        plist = _parse_json_list(r.get("Paiements", ""))
        paid_from_json.append(_sum_payments(plist))
    paid_from_json = pd.Series(paid_from_json, index=df.index, dtype=float)
    df["Payé"] = pd.Series([max(a, b) for a, b in zip(_to_num(df["Payé"]), paid_from_json)], index=df.index)

    # Reste
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # Catégorie depuis ref si manquante
    if visa_ref is not None and not visa_ref.empty:
        mask_cat_missing = (df["Catégorie"].astype(str).str.strip() == "")
        if mask_cat_missing.any():
            df.loc[mask_cat_missing, "Catégorie"] = df.loc[mask_cat_missing, "Visa"].apply(lambda v: map_category_from_ref(visa_ref, v))

    # Statuts & dates
    for b in STATUS_COLS:
        if b not in df.columns:
            df[b] = False
        else:
            df[b] = df[b].astype(bool)
    for dcol in STATUS_DATES:
        if dcol not in df.columns:
            df[dcol] = ""

    # ESCROW
    if ESC_TR not in df.columns: df[ESC_TR] = 0.0
    df[ESC_TR] = _to_num(df[ESC_TR])
    if ESC_JR not in df.columns: df[ESC_JR] = ""

    # Dossier N
    df = ensure_dossier_numbers(df)

    return df

# ---------- HIERARCHIE ARBORESCENTE (Catégorie -> Visa -> Sous-type COS/EOS ou "VISA final") ----------
TREE_COLS = ["Catégorie", "Visa", "Sous-type"]

def _norm_header_map(cols: list[str]) -> dict:
    """Essaie de reconnaître Catégorie / Visa / Sous-type / VISA (final 'B-1 COS')."""
    m = {}
    for c in cols:
        raw = str(c).strip()
        low = (raw.lower()
                 .replace("é","e").replace("è","e").replace("ê","e")
                 .replace("à","a").replace("ô","o").replace("ï","i").replace("ç","c"))
        if low in ("categorie","categories","catégorie"):
            m[c] = "Catégorie"
        elif low == "visa":
            # peut être soit le code (B-1), soit la finale ("B-1 COS")
            m[c] = "VISA_FINAL"
        elif low in ("sous-type","soustype","sous type","type","subtype"):
            m[c] = "Sous-type"
    return m

def read_visa_reference_tree(path: Path) -> pd.DataFrame:
    """
    Lit la feuille 'Visa' :
    - soit colonnes Catégorie / Visa / Sous-type (COS/EOS)
    - soit colonnes Catégorie / (sous-catégories...) / VISA (ex: "B-1 COS")
    Retourne toujours un DF normalisé: Catégorie, Visa, Sous-type.
    """
    try:
        dfv = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=TREE_COLS)

    # 1) normaliser les en-têtes clés
    hdr_map = _norm_header_map(list(dfv.columns))
    if hdr_map:
        dfv = dfv.rename(columns=hdr_map)

    # 2) Catégorie (ffill si organigramme en escalier)
    if "Catégorie" not in dfv.columns:
        dfv["Catégorie"] = ""
    dfv["Catégorie"] = dfv["Catégorie"].fillna("").astype(str).str.strip()
    if (dfv["Catégorie"] == "").any():
        dfv["Catégorie"] = dfv["Catégorie"].replace("", pd.NA).ffill().fillna("")

    # 3) Deux cas :
    out_rows = []
    if "VISA_FINAL" in dfv.columns:
        series = dfv["VISA_FINAL"].fillna("").astype(str).str.strip()
        for _, r in dfv.iterrows():
            cat = str(r.get("Catégorie","")).strip()
            final = str(r.get("VISA_FINAL","")).strip()
            if final == "":
                continue
            parts = final.split()
            if len(parts) == 1:
                visa_code = parts[0]
                sous = ""
            else:
                visa_code = " ".join(parts[:-1])
                sous = parts[-1].upper()
                if sous not in {"COS","EOS"}:
                    visa_code = final
                    sous = ""
            out_rows.append({"Catégorie": cat, "Visa": visa_code, "Sous-type": sous})
    else:
        if "Visa" not in dfv.columns:
            dfv["Visa"] = ""
        if "Sous-type" not in dfv.columns:
            dfv["Sous-type"] = ""
        for _, r in dfv.iterrows():
            cat  = str(r.get("Catégorie","")).strip()
            visa = str(r.get("Visa","")).strip()
            sous = str(r.get("Sous-type","")).strip().upper()
            if visa == "" and sous == "":
                continue
            if sous not in {"","COS","EOS"}:
                sous = re.sub(r"\s+", "", sous).upper()
                sous = {"C0S":"COS","E0S":"EOS"}.get(sous, sous)
            out_rows.append({"Catégorie": cat, "Visa": visa, "Sous-type": sous})

    df = pd.DataFrame(out_rows, columns=TREE_COLS).fillna("")
    for c in TREE_COLS:
        df[c] = df[c].astype(str).str.strip()
    df = df.drop_duplicates().reset_index(drop=True)
    return df

def cascading_visa_picker_tree(df_ref: pd.DataFrame, key_prefix: str, init: dict | None = None) -> dict:
    """
    3 niveaux : Catégorie -> Visa -> Sous-type (COS/EOS optionnel).
    Retourne {"Catégorie":..., "Visa":..., "Sous-type":...}
    """
    result = {"Catégorie":"", "Visa":"", "Sous-type":""}
    if df_ref is None or df_ref.empty:
        st.info("Référentiel Visa vide.")
        return result

    # 1) Catégorie
    dfC = df_ref.copy()
    cats = sorted([v for v in dfC["Catégorie"].unique() if v])
    idxC = 0
    if init and init.get("Catégorie","") in cats: idxC = cats.index(init["Catégorie"])+1
    result["Catégorie"] = st.selectbox("Catégorie", [""]+cats, index=idxC, key=f"{key_prefix}_cat")
    if result["Catégorie"]:
        dfC = dfC[dfC["Catégorie"] == result["Catégorie"]]

    # 2) Visa filtré par Catégorie
    visas = sorted([v for v in dfC["Visa"].unique() if v])
    idxV = 0
    if init and init.get("Visa","") in visas: idxV = visas.index(init["Visa"])+1
    result["Visa"] = st.selectbox("Visa", [""]+visas, index=idxV, key=f"{key_prefix}_visa")
    if result["Visa"]:
        dfV = dfC[dfC["Visa"] == result["Visa"]]
    else:
        dfV = dfC.copy()

    # 3) Sous-type (COS/EOS) si dispo
    sous = sorted([v for v in dfV["Sous-type"].unique() if v])
    idxS = 0
    if init and init.get("Sous-type","") in sous: idxS = sous.index(init["Sous-type"])+1
    result["Sous-type"] = st.selectbox("Sous-type (COS/EOS)", [""]+sous, index=idxS, key=f"{key_prefix}_soustype")

    # feedback
    if not visas:
        st.caption("Visa : (aucun pour cette catégorie)")
    elif result["Visa"] and not sous:
        st.caption(f"Visa : **{result['Visa']}** (pas de sous-type)")
    elif result["Visa"] and sous and not result["Sous-type"]:
        st.caption(f"Visa **{result['Visa']}** — sous-types possibles : {', '.join(sous)}")

    return result

def visas_autorises_from_tree(df_ref: pd.DataFrame, sel: dict) -> list[str]:
    """Liste des visas compatibles avec la sélection (pour filtrer les dossiers)."""
    if df_ref is None or df_ref.empty:
        return []
    dfw = df_ref.copy()
    cat = _safe_str(sel.get("Catégorie","")); vis = _safe_str(sel.get("Visa","")); stype = _safe_str(sel.get("Sous-type","")).upper()
    if cat:   dfw = dfw[dfw["Catégorie"]==cat]
    if vis:   dfw = dfw[dfw["Visa"]==vis]
    if stype: dfw = dfw[dfw["Sous-type"]==stype]
    return sorted([v for v in dfw["Visa"].unique() if v])

# ---------- ESCROW helpers ----------
def escrow_available_from_row(row: pd.Series) -> float:
    """Disponible à transférer depuis ESCROW (honoraires payés - déjà transféré)."""
    hono = float(pd.to_numeric(pd.Series([row.get(HONO, 0.0)]), errors="coerce").fillna(0.0).iloc[0])
    paid = float(pd.to_numeric(pd.Series([row.get("Payé", 0.0)]), errors="coerce").fillna(0.0).iloc[0])
    transferred = float(pd.to_numeric(pd.Series([row.get(ESC_TR, 0.0)]), errors="coerce").fillna(0.0).iloc[0])
    return float(max(min(paid, hono) - transferred, 0.0))

def append_escrow_journal(row: pd.Series, amount: float, note: str = "") -> str:
    lst = _parse_json_list(row.get(ESC_JR, ""))
    lst.append({"ts": datetime.now().isoformat(timespec="seconds"), "amount": float(amount), "note": _safe_str(note)})
    return json.dumps(lst, ensure_ascii=False)


# =========================
# VISA APP — PARTIE 2/3
# =========================

st.set_page_config(
    page_title="Visa Manager",
    page_icon="🛂",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🧳 Gestion des Dossiers Visa")

# Chargement du dernier fichier utilisé
current_path = current_file_path()
visa_ref_tree = None
live_all = None
sheet_choice = "Clients"

# --- Sidebar : source du fichier ---
st.sidebar.header("📂 Données")
src_choice = st.sidebar.radio("Source :", ["Fichier par défaut", "Importer un Excel"], horizontal=False)

if src_choice == "Importer un Excel":
    up = st.sidebar.file_uploader("Dépose un fichier Excel (.xlsx, .xls)", type=["xlsx", "xls"])
    if up:
        current_path = set_current_file_from_upload(up)
        st.sidebar.success(f"Fichier chargé : {current_path.name}")
else:
    if current_path and current_path.exists():
        st.sidebar.info(f"Fichier courant : {current_path.name}")
    else:
        st.sidebar.warning("Aucun fichier sélectionné. Veuillez importer un Excel.")
        st.stop()

# Sauvegarde du chemin
if current_path:
    save_workspace_path(current_path)

# Lecture des feuilles disponibles
if not current_path or not current_path.exists():
    st.warning("Aucun fichier Excel valide trouvé.")
    st.stop()

sheets = list_sheets(current_path)
if not sheets:
    st.error("Aucune feuille trouvée dans ce fichier Excel.")
    st.stop()

# Choix de la feuille à manipuler
sheet_choice = st.sidebar.selectbox("Feuille à afficher :", sheets, index=0)

# Lecture des données
visa_ref_tree = read_visa_reference_tree(current_path)
live_all = read_sheet(current_path, sheet_choice, normalize=True, visa_ref=visa_ref_tree)

if live_all.empty:
    st.error("Aucune donnée à afficher.")
    st.stop()

# --- Section principale : Tableau de bord ---
st.markdown("## 📊 Tableau de bord général")

# Filtres dynamiques
st.subheader("🔍 Filtres")

colF1, colF2, colF3 = st.columns(3)
with colF1:
    sel_path = cascading_visa_picker_tree(visa_ref_tree, key_prefix="dashflt")
with colF2:
    yearsA = sorted(live_all["Date"].dropna().apply(lambda d: d.year).unique().tolist())
    sel_years = st.multiselect("Année", yearsA, default=[])
with colF3:
    monthsA = sorted(live_all["Mois"].dropna().unique().tolist())
    sel_months = st.multiselect("Mois (MM)", monthsA, default=[])

# Application des filtres
mask = pd.Series(True, index=live_all.index)
if sel_years:
    mask &= live_all["Date"].apply(lambda d: d.year if pd.notna(d) else None).isin(sel_years)
if sel_months:
    mask &= live_all["Mois"].isin(sel_months)

visas_aut = visas_autorises_from_tree(visa_ref_tree, sel_path)
if visas_aut:
    mask &= live_all["Visa"].isin(visas_aut)
elif sel_path.get("Catégorie"):
    mask &= live_all["Catégorie"] == sel_path["Catégorie"]

filtered = live_all.loc[mask].copy()

# --- KPIs ---
st.markdown("### 📈 Indicateurs clés")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Dossiers total", len(filtered))
c2.metric("Montant honoraires total", _fmt_money_us(filtered[HONO].sum()))
c3.metric("Total payé", _fmt_money_us(filtered["Payé"].sum()))
c4.metric("Solde à encaisser", _fmt_money_us(filtered["Reste"].sum()))

# --- Alerte ESCROW ---
st.markdown("### ⚠️ Comptes ESCROW à solder avant transfert")
esc_alerts = []
for _, r in filtered.iterrows():
    avail = escrow_available_from_row(r)
    if avail > 0 and r.get(S_ENVOYE, False):
        esc_alerts.append((r.get("Nom", ""), _fmt_money_us(avail)))

if esc_alerts:
    st.warning("Certains dossiers envoyés ont encore des montants à transférer sur le compte ordinaire :")
    for nom, val in esc_alerts:
        st.write(f"- **{nom}** → {val}")
else:
    st.success("Aucun transfert ESCROW en attente pour les dossiers envoyés ✅")

# --- Tableau principal ---
st.markdown("### 📋 Liste des dossiers")
show_cols = [DOSSIER_COL, "Nom", "Catégorie", "Visa", "Sous-type", "Date", HONO, AUTRE, TOTAL, "Payé", "Reste"]
show_cols = [c for c in show_cols if c in filtered.columns]

if not filtered.empty:
    df_show = filtered[show_cols].copy()
    df_show[HONO] = df_show[HONO].map(_fmt_money_us)
    df_show[AUTRE] = df_show[AUTRE].map(_fmt_money_us)
    df_show[TOTAL] = df_show[TOTAL].map(_fmt_money_us)
    df_show["Payé"] = df_show["Payé"].map(_fmt_money_us)
    df_show["Reste"] = df_show["Reste"].map(_fmt_money_us)
    st.dataframe(df_show, use_container_width=True)
else:
    st.info("Aucun dossier trouvé pour ces filtres.")


# =========================
# VISA APP — PARTIE 3/3
# =========================

# ---- Préparation : retrouver la feuille "clients" par défaut
def _guess_clients_sheet(path: Path, fallback: str | None = None) -> str:
    sheets = list_sheets(path)
    # priorité aux noms usuels
    for name in ["Clients", "Client", "Données", "Data"]:
        if name in sheets:
            return name
    # sinon, 1ère feuille qui contient Nom & Visa
    for sn in sheets:
        try:
            t = pd.read_excel(path, sheet_name=sn, nrows=2)
            cols = {c.lower() for c in t.columns.astype(str)}
            if "nom" in cols and "visa" in cols:
                return sn
        except Exception:
            pass
    # fallback -> actuelle
    return fallback or sheets[0]

clients_sheet = _guess_clients_sheet(current_path, sheet_choice)

tabs = st.tabs(["👤 Clients (CRUD + paiements)", "📊 Analyses", "🏦 ESCROW"])

# ======================================================================
# 👤 CLIENTS (CRUD + paiements)
# ======================================================================
with tabs[0]:
    st.subheader("Clients — créer, modifier, supprimer (écriture directe dans l’Excel)")
    if st.button("🔄 Recharger", key="reload_clients"):
        st.rerun()

    live_raw = read_sheet(current_path, clients_sheet, normalize=False).copy()
    live_norm = normalize_dataframe(live_raw, visa_ref=visa_ref_tree).copy()

    # ID technique pour localiser la ligne
    live_raw["_RowID"] = range(len(live_raw))

    action = st.radio("Action", ["Créer", "Modifier", "Supprimer"], horizontal=True, key="crud_action")

    # ---------------------- CRÉER ----------------------
    if action == "Créer":
        st.markdown("### ➕ Nouveau client")
        next_num = next_dossier_number(live_raw)

        with st.form("create_form", clear_on_submit=False):
            c0, c1, c2 = st.columns([1,1,1])
            c0.metric("Prochain Dossier N", f"{next_num}")
            nom_in = c1.text_input("Nom *")
            d      = c2.date_input("Date", value=date.today())

            st.caption("Sélection hiérarchique (Catégorie → Visa → Sous-type)")
            sel_path = cascading_visa_picker_tree(visa_ref_tree, key_prefix="create_tree")
            cat  = sel_path.get("Catégorie",""); visa = sel_path.get("Visa",""); stype = sel_path.get("Sous-type","")

            c5,c6 = st.columns(2)
            honoraires = c5.number_input(HONO, value=0.0, step=10.0, format="%.2f")
            autres     = c6.number_input(AUTRE, value=0.0, step=10.0, format="%.2f")

            st.markdown("#### État du dossier (avec dates)")
            def _dtdefault(): return date.today()
            r1c1, r1c2 = st.columns(2)
            v_env = r1c1.checkbox(S_ENVOYE, value=False);   dt_env = r1c2.date_input(D_ENVOYE, value=_dtdefault(), disabled=not v_env)
            r2c1, r2c2 = st.columns(2)
            v_app = r2c1.checkbox(S_APPROUVE, value=False); dt_app = r2c2.date_input(D_APPROUVE, value=_dtdefault(), disabled=not v_app)
            r3c1, r3c2 = st.columns(2)
            v_rfe = r3c1.checkbox(S_RFE, value=False);      dt_rfe = r3c2.date_input(D_RFE, value=_dtdefault(), disabled=not v_rfe)
            r4c1, r4c2 = st.columns(2)
            v_ref = r4c1.checkbox(S_REFUSE, value=False);   dt_ref = r4c2.date_input(D_REFUSE, value=_dtdefault(), disabled=not v_ref)
            r5c1, r5c2 = st.columns(2)
            v_ann = r5c1.checkbox(S_ANNULE, value=False);   dt_ann = r5c2.date_input(D_ANNULE, value=_dtdefault(), disabled=not v_ann)

            ok = st.form_submit_button("💾 Sauvegarder (dans le fichier)", type="primary")

        if ok:
            if not _safe_str(nom_in):
                st.error("Le nom est obligatoire."); st.stop()
            if v_rfe and not (v_env or v_ref or v_ann):
                st.error("RFE autorisé seulement si Envoyé/Refusé/Annulé est coché."); st.stop()

            # nom unique si doublon : "xxxx-0", "xxxx-1", ...
            existing = set(live_raw.get("Nom","").astype(str))
            use_name = _safe_str(nom_in)
            if use_name in existing:
                k = 0
                while f"{use_name}-{k}" in existing: k += 1
                use_name = f"{use_name}-{k}"

            gen_id = _make_client_id_from_row({"Nom": use_name, "Date": d})
            exist_ids = set(live_raw.get("ID_Client","").astype(str))
            new_id = gen_id; n=1
            while new_id in exist_ids:
                n += 1; new_id = f"{gen_id}-{n:02d}"

            total = float((honoraires or 0.0)+(autres or 0.0))
            new_row = {
                DOSSIER_COL: int(next_num),
                "ID_Client": new_id,
                "Nom": use_name,
                "Date": str(d),
                "Mois": f"{d.month:02d}",
                "Catégorie": _safe_str(cat),
                "Visa": _safe_str(visa),
                "Sous-type": _safe_str(stype).upper(),
                HONO: float(honoraires or 0.0),
                AUTRE: float(autres or 0.0),
                TOTAL: total,
                "Payé": 0.0,
                "Reste": max(total, 0.0),
                "Paiements": "",
                S_ENVOYE: bool(v_env),   D_ENVOYE: (str(dt_env) if v_env else ""),
                S_APPROUVE: bool(v_app), D_APPROUVE: (str(dt_app) if v_app else ""),
                S_RFE: bool(v_rfe),      D_RFE: (str(dt_rfe) if v_rfe else ""),
                S_REFUSE: bool(v_ref),   D_REFUSE: (str(dt_ref) if v_ref else ""),
                S_ANNULE: bool(v_ann),   D_ANNULE: (str(dt_ann) if v_ann else ""),
                ESC_TR: 0.0,
                ESC_JR: ""
            }

            live_after = pd.concat([live_raw.drop(columns=["_RowID"]), pd.DataFrame([new_row])], ignore_index=True)
            live_after = ensure_dossier_numbers(live_after)
            write_sheet_inplace(current_path, clients_sheet, live_after)
            st.success("Client créé ✅"); st.rerun()

    # ---------------------- MODIFIER ----------------------
    if action == "Modifier":
        st.markdown("### ✏️ Modifier une fiche & gérer les paiements")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            options = [(int(r["_RowID"]),
                        f'{int(r.get(DOSSIER_COL,0))} — { _safe_str(r.get("ID_Client")) } — { _safe_str(r.get("Nom")) }')
                       for _, r in live_raw.iterrows()]
            labels = [lab for _, lab in options]
            sel_lab = st.selectbox("Sélection", labels, key="edit_pick")
            sel_rowid = [rid for rid, lab in options if lab==sel_lab][0]
            idx = live_raw.index[live_raw["_RowID"]==sel_rowid][0]
            init = live_norm.loc[idx].to_dict()

            with st.form(f"edit_form_{sel_rowid}", clear_on_submit=False):
                c0, c1, c2 = st.columns([1,1,1])
                c0.metric("Dossier N", f'{int(init.get(DOSSIER_COL,0))}')
                nom = c1.text_input("Nom", value=_safe_str(init.get("Nom")), key=f"nom_{sel_rowid}")
                try:
                    d_init = pd.to_datetime(init.get("Date")).date() if _safe_str(init.get("Date")) else date.today()
                except Exception:
                    d_init = date.today()
                d = c2.date_input("Date", value=d_init, key=f"date_{sel_rowid}")

                st.caption("Sélection hiérarchique (Catégorie → Visa → Sous-type)")
                init_path = {
                    "Catégorie": _safe_str(init.get("Catégorie")),
                    "Visa": _safe_str(init.get("Visa")),
                    "Sous-type": _safe_str(init.get("Sous-type")).upper()
                }
                sel_path = cascading_visa_picker_tree(visa_ref_tree, key_prefix=f"edit_tree_{sel_rowid}", init=init_path)
                cat  = sel_path.get("Catégorie",""); visa = sel_path.get("Visa",""); stype = sel_path.get("Sous-type","")

                def _f(v, alt=0.0):
                    try: return float(v)
                    except Exception: return float(alt)
                hono0  = _f(init.get(HONO, init.get("Montant", 0.0)))
                autre0 = _f(init.get(AUTRE, 0.0))
                paye0  = _f(init.get("Payé", 0.0))

                c5,c6 = st.columns(2)
                honoraires = c5.number_input(HONO, value=hono0, step=10.0, format="%.2f", key=f"hono_{sel_rowid}")
                autres     = c6.number_input(AUTRE, value=autre0, step=10.0, format="%.2f", key=f"autre_{sel_rowid}")

                c7,c8 = st.columns(2)
                total_preview = float(honoraires + autres)
                c7.metric("Total (US $)", _fmt_money_us(total_preview))
                c8.metric("Solde après sauvegarde", _fmt_money_us(max(total_preview - paye0, 0.0)))

                st.markdown("#### État du dossier (avec dates)")
                def _get_dt(key, default_today=True):
                    v = _safe_str(init.get(key, ""))
                    try:
                        return pd.to_datetime(v).date() if v else (date.today() if default_today else None)
                    except Exception:
                        return date.today() if default_today else None

                r1c1, r1c2 = st.columns(2)
                v_env = r1c1.checkbox(S_ENVOYE, value=bool(init.get(S_ENVOYE)));   dt_env = r1c2.date_input(D_ENVOYE, value=_get_dt(D_ENVOYE), disabled=not v_env)
                r2c1, r2c2 = st.columns(2)
                v_app = r2c1.checkbox(S_APPROUVE, value=bool(init.get(S_APPROUVE))); dt_app = r2c2.date_input(D_APPROUVE, value=_get_dt(D_APPROUVE), disabled=not v_app)
                r3c1, r3c2 = st.columns(2)
                v_rfe = r3c1.checkbox(S_RFE, value=bool(init.get(S_RFE)));         dt_rfe = r3c2.date_input(D_RFE, value=_get_dt(D_RFE), disabled=not v_rfe)
                r4c1, r4c2 = st.columns(2)
                v_ref = r4c1.checkbox(S_REFUSE, value=bool(init.get(S_REFUSE)));   dt_ref = r4c2.date_input(D_REFUSE, value=_get_dt(D_REFUSE), disabled=not v_ref)
                r5c1, r5c2 = st.columns(2)
                v_ann = r5c1.checkbox(S_ANNULE, value=bool(init.get(S_ANNULE)));   dt_ann = r5c2.date_input(D_ANNULE, value=_get_dt(D_ANNULE), disabled=not v_ann)

                ok_fiche = st.form_submit_button("💾 Enregistrer la fiche", type="primary")

            if ok_fiche:
                if v_rfe and not (v_env or v_ref or v_ann):
                    st.error("RFE autorisé seulement si Envoyé/Refusé/Annulé est coché."); st.stop()

                live_w = read_sheet(current_path, clients_sheet, normalize=False).copy()

                # Re-trouve la ligne par ID_Client, sinon par Dossier N, sinon position
                t_idx = None
                key_id = _safe_str(init.get("ID_Client"))
                if key_id and "ID_Client" in live_w.columns:
                    hits = live_w.index[live_w["ID_Client"].astype(str) == key_id]
                    if len(hits)>0: t_idx = hits[0]
                if t_idx is None and DOSSIER_COL in live_w.columns:
                    try:
                        num = int(_to_int(pd.Series([init.get(DOSSIER_COL)])).iloc[0])
                        hits = live_w.index[_to_int(live_w[DOSSIER_COL]) == num]
                        if len(hits)>0: t_idx = hits[0]
                    except Exception:
                        pass
                if t_idx is None and 0 <= int(idx) < len(live_w):
                    t_idx = int(idx)
                if t_idx is None:
                    st.error("Ligne introuvable."); st.stop()

                # Assure colonnes
                for c in [HONO, AUTRE, TOTAL, "Payé","Reste","Paiements",
                          "Catégorie","Visa","Sous-type","Nom","Date","Mois",
                          S_ENVOYE, D_ENVOYE, S_APPROUVE, D_APPROUVE, S_RFE, D_RFE, S_REFUSE, D_REFUSE, S_ANNULE, D_ANNULE,
                          ESC_TR, ESC_JR, DOSSIER_COL]:
                    if c not in live_w.columns:
                        live_w[c] = 0.0 if c in [HONO, AUTRE, TOTAL, "Payé","Reste", ESC_TR] else ""

                # Écrit les champs
                live_w.at[t_idx,"Nom"] = _safe_str(nom)
                live_w.at[t_idx,"Date"] = str(d)
                live_w.at[t_idx,"Mois"] = f"{d.month:02d}"
                live_w.at[t_idx,"Catégorie"] = _safe_str(cat)
                live_w.at[t_idx,"Visa"] = _safe_str(visa)
                live_w.at[t_idx,"Sous-type"] = _safe_str(stype).upper()
                live_w.at[t_idx, HONO] = float(honoraires or 0.0)
                live_w.at[t_idx, AUTRE] = float(autres or 0.0)

                # Statuts + dates
                live_w.at[t_idx, S_ENVOYE] = bool(v_env);  live_w.at[t_idx, D_ENVOYE] = (str(dt_env) if v_env else "")
                live_w.at[t_idx, S_APPROUVE] = bool(v_app); live_w.at[t_idx, D_APPROUVE] = (str(dt_app) if v_app else "")
                live_w.at[t_idx, S_RFE] = bool(v_rfe);      live_w.at[t_idx, D_RFE] = (str(dt_rfe) if v_rfe else "")
                live_w.at[t_idx, S_REFUSE] = bool(v_ref);   live_w.at[t_idx, D_REFUSE] = (str(dt_ref) if v_ref else "")
                live_w.at[t_idx, S_ANNULE] = bool(v_ann);   live_w.at[t_idx, D_ANNULE] = (str(dt_ann) if v_ann else "")

                # Recalcul total / payé / reste
                pay_json = live_w.at[t_idx, "Paiements"]
                paid = _sum_payments(_parse_json_list(pay_json))
                total = float((honoraires or 0.0) + (autres or 0.0))
                live_w.at[t_idx, "Payé"]  = float(paid)
                live_w.at[t_idx, TOTAL]   = total
                live_w.at[t_idx, "Reste"] = max(total - float(paid), 0.0)

                live_w = ensure_dossier_numbers(live_w)
                write_sheet_inplace(current_path, clients_sheet, live_w)
                st.success("Fiche enregistrée ✅"); st.rerun()

            # ---------------- Paiements ----------------
            st.markdown("#### 💳 Historique & ajout d’un règlement")
            live_now = read_sheet(current_path, clients_sheet, normalize=False)
            ixs = live_now.index[live_now.get("ID_Client","").astype(str) == _safe_str(init.get("ID_Client"))]
            if len(ixs)==0:
                st.info("Ligne introuvable pour les paiements.")
            else:
                i = ixs[0]
                if "Paiements" not in live_now.columns: live_now["Paiements"] = ""
                plist = _parse_json_list(live_now.at[i,"Paiements"])

                # Tableau historique
                if plist:
                    dfp = pd.DataFrame(plist)
                    for col in ["date","mode","note","amount"]:
                        if col not in dfp.columns: dfp[col] = "" if col!="amount" else 0.0
                    dfp["date"] = pd.to_datetime(dfp["date"], errors="coerce").dt.date.astype(str)
                    dfp["Montant ($)"] = dfp["amount"].apply(lambda x: _fmt_money_us(float(x) if pd.notna(x) else 0.0))
                    st.table(dfp[["date","mode","Montant ($)","note"]])
                else:
                    st.caption("Aucun paiement enregistré pour ce client.")

                cA, cB, cC, cD = st.columns([1,1,1,2])
                pay_date = cA.date_input("Date", value=date.today(), key=f"pay_date_{i}")
                pay_mode = cB.selectbox("Mode", ["CB","Chèque","Espèces","Virement","Venmo","Autre"], key=f"pay_mode_{i}")
                pay_amt  = cC.number_input("Montant ($)", min_value=0.0, step=10.0, format="%.2f", key=f"pay_amt_{i}")
                pay_note = cD.text_input("Note", "", key=f"pay_note_{i}")

                if st.button("💾 Enregistrer ce règlement", key=f"pay_add_btn_{i}"):
                    try:
                        add = float(pay_amt or 0.0)
                        if add <= 0: st.warning("Le montant doit être > 0."); st.stop()
                        # plafond = reste courant
                        norm = normalize_dataframe(live_now.copy(), visa_ref=read_visa_reference(current_path))
                        mask_id = (norm["ID_Client"].astype(str) == _safe_str(init.get("ID_Client")))
                        reste_curr = float(norm.loc[mask_id, "Reste"].sum()) if mask_id.any() else 0.0
                        if add > reste_curr + 1e-9:
                            add = reste_curr

                        plist.append({"date": str(pay_date), "amount": float(add), "mode": pay_mode, "note": pay_note})
                        live_now.at[i,"Paiements"] = json.dumps(plist, ensure_ascii=False)

                        total_paid = _sum_payments(plist)
                        hono = _to_num(pd.Series([live_now.at[i, HONO] if HONO in live_now.columns else 0.0])).iloc[0]
                        autr = _to_num(pd.Series([live_now.at[i, AUTRE] if AUTRE in live_now.columns else 0.0])).iloc[0]
                        total = float(hono + autr)
                        live_now.at[i,"Payé"]  = float(total_paid)
                        live_now.at[i,"Reste"] = max(total - float(total_paid), 0.0)
                        live_now.at[i,TOTAL]   = total

                        write_sheet_inplace(current_path, clients_sheet, live_now)
                        st.success("Règlement ajouté ✅"); st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")

                # Suppression d’une ligne de paiement
                if plist:
                    del_idx = st.number_input("Supprimer la ligne n° (1..n)", min_value=1, max_value=len(plist), value=1, step=1, key=f"del_pay_idx_{i}")
                    if st.button("🗑️ Supprimer cette ligne", key=f"del_pay_btn_{i}"):
                        try:
                            del plist[int(del_idx)-1]
                            live_now.at[i,"Paiements"] = json.dumps(plist, ensure_ascii=False)
                            total_paid = _sum_payments(plist)
                            hono = _to_num(pd.Series([live_now.at[i, HONO] if HONO in live_now.columns else 0.0])).iloc[0]
                            autr = _to_num(pd.Series([live_now.at[i, AUTRE] if AUTRE in live_now.columns else 0.0])).iloc[0]
                            total = float(hono + autr)
                            live_now.at[i,"Payé"]  = float(total_paid)
                            live_now.at[i,"Reste"] = max(total - float(total_paid), 0.0)
                            live_now.at[i,TOTAL]   = total
                            write_sheet_inplace(current_path, clients_sheet, live_now)
                            st.success("Ligne supprimée ✅"); st.rerun()
                        except Exception as e:
                            st.error(f"Erreur suppression : {e}")

    # ---------------------- SUPPRIMER ----------------------
    if action == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            options = [(int(r["_RowID"]),
                        f'{int(r.get(DOSSIER_COL,0))} — { _safe_str(r.get("ID_Client")) } — { _safe_str(r.get("Nom")) }')
                       for _, r in live_raw.iterrows()]
            labels = [lab for _, lab in options]
            sel_lab = st.selectbox("Sélection", labels, key="del_pick")
            sel_rowid = [rid for rid, lab in options if lab==sel_lab][0]
            idx = live_raw.index[live_raw["_RowID"]==sel_rowid][0]
            st.error("⚠️ Action irréversible")
            if st.button("Supprimer définitivement", key="del_btn"):
                live_w = live_raw.drop(columns=["_RowID"]).copy()
                key = _safe_str(live_raw.at[idx, "ID_Client"])
                if key and "ID_Client" in live_w.columns:
                    live_w = live_w[live_w["ID_Client"].astype(str)!=key].reset_index(drop=True)
                else:
                    nom = _safe_str(live_raw.at[idx,"Nom"]); dat = _safe_str(live_raw.at[idx,"Date"])
                    live_w = live_w[~((live_w.get("Nom","").astype(str)==nom)&(live_w.get("Date","").astype(str)==dat))].reset_index(drop=True)
                live_w = ensure_dossier_numbers(live_w)
                write_sheet_inplace(current_path, clients_sheet, live_w)
                st.success("Client supprimé ✅"); st.rerun()

# ======================================================================
# 📊 ANALYSES
# ======================================================================
with tabs[1]:
    st.subheader("Analyses — volumes & financier")

    dfA = normalize_dataframe(read_sheet(current_path, clients_sheet, normalize=False), visa_ref=visa_ref_tree)
    if dfA.empty:
        st.info("Aucune donnée.")
        st.stop()

    # Filtres
    cL, cR = st.columns([1,2])
    with cL:
        st.caption("Filtres (Catégorie → Visa → Sous-type)")
        sel_path_a = cascading_visa_picker_tree(visa_ref_tree, key_prefix="anal_tree")
        visas_aut_a = visas_autorises_from_tree(visa_ref_tree, sel_path_a)
    with cR:
        cR1, cR2, cR3 = st.columns(3)
        yearsA  = sorted({d.year for d in dfA["Date"] if pd.notna(d)})
        monthsA = [f"{m:02d}" for m in range(1,13)]
        sel_years  = cR1.multiselect("Année", yearsA, default=[])
        sel_months = cR2.multiselect("Mois (MM)", monthsA, default=[])

    fA = dfA.copy()
    for col in ["Catégorie","Visa","Sous-type"]:
        val = _safe_str(sel_path_a.get(col,""))
        if val:
            fA = fA[fA[col].astype(str)==val]
    if visas_aut_a:
        fA = fA[fA["Visa"].astype(str).isin(visas_aut_a)]
    if sel_years:
        fA = fA[fA["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))]
    if sel_months:
        fA = fA[fA["Mois"].isin(sel_months)]

    # KPI
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(fA)}")
    k2.metric("Total (US $)", _fmt_money_us(float(fA.get(TOTAL, pd.Series(dtype=float)).sum())))
    k3.metric("Payé (US $)", _fmt_money_us(float(fA.get("Payé", pd.Series(dtype=float)).sum())))
    k4.metric("Solde (US $)", _fmt_money_us(float(fA.get("Reste", pd.Series(dtype=float)).sum())))
    st.markdown('</div>', unsafe_allow_html=True)

    # Période
    fA["Année"]  = fA["Date"].apply(lambda x: x.year if pd.notna(x) else pd.NA)
    fA["MoisNum"]= fA["Date"].apply(lambda x: int(x.month) if pd.notna(x) else pd.NA)
    fA["Periode"]= fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")

    # Graph volumes / montants
    st.markdown("### 📈 Volumes par période")
    vol = fA.groupby("Periode").size().reset_index(name="Volume")
    if not vol.empty:
        try:
            st.altair_chart(
                alt.Chart(vol).mark_line(point=True).encode(
                    x=alt.X("Periode:N", sort=None),
                    y="Volume:Q",
                    tooltip=["Periode","Volume"]
                ).properties(height=260),
                use_container_width=True
            )
        except Exception:
            st.dataframe(vol, use_container_width=True)

    st.markdown("### 💵 Montants par année")
    by_year = fA.dropna(subset=["Année"]).groupby("Année").agg(
        Dossiers=("Nom","count"),
        Honoraires=(HONO,"sum"),
        Autres=(AUTRE,"sum"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index().sort_values("Année")

    c1, c2 = st.columns(2)
    if not by_year.empty:
        try:
            c1.altair_chart(
                alt.Chart(by_year).mark_bar().encode(
                    x="Année:N",
                    y="Dossiers:Q",
                    tooltip=["Année","Dossiers"]
                ).properties(height=260),
                use_container_width=True
            )
            melted = by_year.melt("Année", ["Honoraires","Autres","Total","Payé","Reste"], var_name="Indicateur", value_name="Montant")
            c2.altair_chart(
                alt.Chart(melted).mark_bar().encode(
                    x="Année:N", y="Montant:Q", color="Indicateur:N",
                    tooltip=["Année","Indicateur", alt.Tooltip("Montant:Q", format="$.2f")]
                ).properties(height=260),
                use_container_width=True
            )
        except Exception:
            c1.dataframe(by_year[["Année","Dossiers"]], use_container_width=True)
            c2.dataframe(by_year.drop(columns=["Dossiers"]), use_container_width=True)

    st.markdown("### 🔎 Détails (clients filtrés)")
    details_cols = [c for c in ["Periode", DOSSIER_COL, "ID_Client", "Nom", "Catégorie","Visa","Sous-type",
                                "Date", HONO, AUTRE, TOTAL, "Payé","Reste","Année","MoisNum"]
                    if c in fA.columns]
    det = fA[details_cols].copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in det.columns:
            det[col] = det[col].apply(lambda x: _fmt_money_us(x) if pd.notna(x) else "")
    st.dataframe(det.sort_values(["Année","MoisNum","Catégorie","Nom"], na_position="last"),
                 use_container_width=True)

# ======================================================================
# 🏦 ESCROW
# ======================================================================
with tabs[2]:
    st.subheader("ESCROW — dépôts sur honoraires & transferts")

    live = normalize_dataframe(read_sheet(current_path, clients_sheet, normalize=False), visa_ref=visa_ref_tree).copy()
    if live.empty:
        st.info("Aucune donnée.")
        st.stop()
    if ESC_TR not in live.columns: live[ESC_TR] = 0.0
    else: live[ESC_TR] = pd.to_numeric(live[ESC_TR], errors="coerce").fillna(0.0)
    live["ESCROW dispo"] = live.apply(escrow_available_from_row, axis=1)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Dossiers", f"{len(live)}")
    c2.metric("ESCROW total dispo", _fmt_money_us(float(live["ESCROW dispo"].sum())))
    envoyes = live[(live[S_ENVOYE]==True)]
    a_transferer = envoyes[envoyes["ESCROW dispo"]>0.004]
    c3.metric("Envoyés à réclamer", f"{len(a_transferer)}")
    c4.metric("Montant à réclamer", _fmt_money_us(float(a_transferer["ESCROW dispo"].sum())))

    st.divider()
    st.markdown("### 📌 À transférer (dossiers **envoyés**)")
    if a_transferer.empty:
        st.success("Aucun transfert en attente.")
    else:
        for _, r in a_transferer.sort_values("Date").iterrows():
            with st.expander(f'🔔 {r[DOSSIER_COL]} — {r["ID_Client"]} — {r["Nom"]} — {r.get("Catégorie","")} / {r["Visa"]} — ESCROW dispo: {_fmt_money_us(r["ESCROW dispo"])}'):
                cA, cB, cC = st.columns(3)
                cA.metric("Honoraires", _fmt_money_us(float(r.get(HONO,0.0))))
                cB.metric("Déjà transféré", _fmt_money_us(float(r.get(ESC_TR,0.0))))
                cC.metric("Payé", _fmt_money_us(float(r.get("Payé",0.0))))
                amt = st.number_input("Montant à marquer comme transféré (US $)",
                                      min_value=0.0, value=float(r["ESCROW dispo"]),
                                      step=10.0, format="%.2f", key=f"esc_amt_{r['ID_Client']}")
                note = st.text_input("Note (facultatif)", "", key=f"esc_note_{r['ID_Client']}")
                if st.button("✅ Marquer transféré", key=f"esc_btn_{r['ID_Client']}"):
                    try:
                        live_w = read_sheet(current_path, clients_sheet, normalize=False).copy()
                        for c in [ESC_TR, ESC_JR]:
                            if c not in live_w.columns: live_w[c] = 0.0 if c==ESC_TR else ""
                        idxs = live_w.index[live_w.get("ID_Client","").astype(str)==str(r["ID_Client"])]
                        if len(idxs)==0: st.error("Ligne introuvable."); st.stop()
                        i = idxs[0]
                        tmp = normalize_dataframe(live_w.copy(), visa_ref=read_visa_reference(current_path))
                        disp = float(tmp.loc[tmp["ID_Client"].astype(str)==str(r["ID_Client"]), :].apply(escrow_available_from_row, axis=1).iloc[0])
                        add = float(min(max(amt,0.0), disp))
                        live_w.at[i, ESC_TR] = float(pd.to_numeric(pd.Series([live_w.at[i, ESC_TR]]), errors="coerce").fillna(0.0).iloc[0] + add)
                        # Journal
                        j = _parse_json_list(live_w.at[i, ESC_JR] if ESC_JR in live_w.columns else "")
                        j.append({"ts": datetime.now().isoformat(timespec="seconds"), "amount": add, "note": _safe_str(note)})
                        live_w.at[i, ESC_JR] = json.dumps(j, ensure_ascii=False)
                        live_w = ensure_dossier_numbers(live_w)
                        write_sheet_inplace(current_path, clients_sheet, live_w)
                        st.success("Transfert ESCROW enregistré ✅"); st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")

    st.divider()
    st.markdown("### 🧾 Historique des transferts")
    has_journal = live[live[ESC_JR].astype(str).str.len() > 0]
    if has_journal.empty:
        st.caption("Aucun journal de transfert.")
    else:
        rows = []
        for _, r in has_journal.iterrows():
            entries = _parse_json_list(r.get(ESC_JR, ""))
            for e in entries:
                rows.append({
                    DOSSIER_COL: r.get(DOSSIER_COL, ""),
                    "ID_Client": r.get("ID_Client", ""),
                    "Nom": r.get("Nom", ""),
                    "Visa": r.get("Visa", ""),
                    "Date dossier": r.get("Date", ""),
                    "Horodatage": e.get("ts", ""),
                    "Montant (US $)": float(e.get("amount", 0.0)),
                    "Note": e.get("note", "")
                })
        jdf = pd.DataFrame(rows)
        if not jdf.empty:
            jdf["Montant (US $)"] = jdf["Montant (US $)"].apply(lambda x: _fmt_money_us(float(x) if pd.notna(x) else 0.0))
        st.dataframe(jdf.sort_values("Horodatage"), use_container_width=True)