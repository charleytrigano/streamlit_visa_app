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
S_ENVOYE, D_ENVOYE   = "Dossier envoyé",  "Date envoyé"
S_APPROUVE, D_APPROUVE = "Dossier approuvé","Date approuvé"
S_RFE, D_RFE         = "RFE",             "Date RFE"
S_REFUSE, D_REFUSE   = "Dossier refusé",  "Date refusé"
S_ANNULE, D_ANNULE   = "Dossier annulé",  "Date annulé"
STATUS_COLS  = [S_ENVOYE, S_APPROUVE, S_RFE, S_REFUSE, S_ANNULE]
STATUS_DATES = [D_ENVOYE, D_APPROUVE, D_RFE, D_REFUSE, D_ANNULE]

# ESCROW
ESC_TR = "ESCROW transféré (US $)"
ESC_JR = "Journal ESCROW"   # JSON [{"ts": "...", "amount": float, "note": ""}]

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
       Gère also df[col] -> DataFrame quand colonnes dupliquées."""
    if s is None:
        return pd.Series(dtype=float)
    if isinstance(s, pd.DataFrame):
        if s.shape[1] == 0:
            return pd.Series(dtype=float, index=s.index if hasattr(s, "index") else None)
        s = s.iloc[:, 0]
    s = s.astype(str)
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)
    def _clean_one(v: str) -> float:
        if v == "" or v == "-":
            return 0.0
        # EU -> point
        if v.count(",")==1 and v.count(".")==0:
            v = v.replace(",", ".")
        # US -> enlève séparateurs
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
    """Écrit la feuille `sheet` en conservant les autres feuilles."""
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

# ---------- Fusion de colonnes dupliquées ----------
def _collapse_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
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
        try:
            same_num = same.apply(pd.to_numeric, errors="coerce")
            if same_num.notna().any().any():
                out[col] = same_num.sum(axis=1, skipna=True); continue
        except Exception:
            pass
        def _first_non_empty(row):
            for v in row:
                if pd.notna(v) and str(v).strip() != "":
                    return v
            return ""
        out[col] = same.apply(_first_non_empty, axis=1)
    return out

# ---------- Normalisation / mapping Visa ----------
def read_visa_reference(path: Path) -> pd.DataFrame:
    """Version simple Catégorie/Visa (sans sous-type) pour mapping catégorie si manquante."""
    try:
        df = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=["Catégorie","Visa"])
    rename = {}
    for c in df.columns:
        lc = str(c).lower().strip()
        if lc in ("categorie","catégorie"): rename[c] = "Catégorie"
        elif lc == "visa": rename[c] = "Visa"
    if rename: df = df.rename(columns=rename)
    for col in ["Catégorie","Visa"]:
        if col not in df.columns: df[col] = ""
    df["Catégorie"] = df["Catégorie"].fillna("").astype(str).str.strip()
    df["Visa"]       = df["Visa"].fillna("").astype(str).str.strip()
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
    nom = _safe_str(row.get("Nom"))
    try:
        d = pd.to_datetime(row.get("Date")).date()
    except Exception:
        d = date.today()
    base = f"{nom}-{d.strftime('%Y%m%d')}"
    base = re.sub(r"[^A-Za-z0-9\-]+", "", base.replace(" ", "-"))
    return base.lower()

# ---------- Normalisation des données clients ----------
def normalize_dataframe(df: pd.DataFrame, visa_ref: pd.DataFrame | None = None) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()

    # Renommages souples
    rename = {}
    for c in df.columns:
        lc = str(c).lower().strip()
        if lc in ("montant honoraires", "montant honoraires (us $)", "honoraires", "montant"):
            rename[c] = HONO
        elif lc in ("autres frais", "autres frais (us $)", "autres"):
            rename[c] = AUTRE
        elif lc in ("total", "total (us $)"):
            rename[c] = TOTAL
        elif lc in ("dossier n","dossier"):
            rename[c] = DOSSIER_COL
        elif lc in ("reste (us $)","solde (us $)","solde"):
            rename[c] = "Reste"
        elif lc in ("paye (us $)","payé (us $)","paye","payé"):
            rename[c] = "Payé"
        elif lc in ("sous-type","soustype","sous type","type","subtype"):
            rename[c] = "Sous-type"
    if rename:
        df = df.rename(columns=rename)

    # Fusion colonnes dupliquées
    df = _collapse_duplicate_columns(df)

    # Colonnes minimales
    for c in [DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Sous-type",
              HONO,AUTRE,TOTAL,"Payé","Reste","Paiements","Date","Mois"]:
        if c not in df.columns:
            if c in [HONO,AUTRE,TOTAL,"Payé","Reste"]:
                df[c] = 0.0
            elif c == "Paiements":
                df[c] = ""
            else:
                df[c] = ""

    # Numériques
    for c in [HONO,AUTRE,TOTAL,"Payé","Reste"]:
        df[c] = _to_num(df[c])

    # Date & Mois (MM uniquement)
    def _to_date(x):
        try:
            if pd.isna(x) or x == "":
                return pd.NaT
            return pd.to_datetime(x).date()
        except Exception:
            return pd.NaT
    df["Date"] = df["Date"].map(_to_date)
    df["Mois"] = df["Date"].apply(lambda d: f"{d.month:02d}" if pd.notna(d) else pd.NA)

    # Total / Payé / Reste
    df[TOTAL] = _to_num(df.get(HONO,0.0)) + _to_num(df.get(AUTRE,0.0))
    paid_from_json = []
    for _, r in df.iterrows():
        paid_from_json.append(_sum_payments(_parse_json_list(r.get("Paiements",""))))
    paid_from_json = pd.Series(paid_from_json, index=df.index, dtype=float)
    df["Payé"] = pd.Series([max(a,b) for a,b in zip(_to_num(df["Payé"]), paid_from_json)], index=df.index)
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # Catégorie depuis ref si manquante
    if visa_ref is not None and not visa_ref.empty:
        mask_cat_missing = (df["Catégorie"].astype(str).str.strip() == "")
        if mask_cat_missing.any():
            df.loc[mask_cat_missing, "Catégorie"] = df.loc[mask_cat_missing, "Visa"].apply(
                lambda v: map_category_from_ref(visa_ref, v)
            )

    # Statuts & dates
    for b in STATUS_COLS:
        if b not in df.columns: df[b] = False
        else: df[b] = df[b].astype(bool)
    for dcol in STATUS_DATES:
        if dcol not in df.columns: df[dcol] = ""

    # ESCROW
    if ESC_TR not in df.columns: df[ESC_TR] = 0.0
    df[ESC_TR] = _to_num(df[ESC_TR])
    if ESC_JR not in df.columns: df[ESC_JR] = ""

    # Dossier N
    df = ensure_dossier_numbers(df)
    return df

# ---------- HIERARCHIE (Catégorie -> Visa -> Sous-type COS/EOS ou dernière colonne VISA "B-1 COS") ----------
TREE_COLS = ["Catégorie", "Visa", "Sous-type"]

def _norm_header_map(cols: list[str]) -> dict:
    m = {}
    for c in cols:
        raw = str(c).strip()
        low = (raw.lower()
                 .replace("é","e").replace("è","e").replace("ê","e")
                 .replace("à","a").replace("ô","o").replace("ï","i").replace("ç","c"))
        if low in ("categorie","categories","catégorie"): m[c] = "Catégorie"
        elif low == "visa": m[c] = "VISA_FINAL"  # peut contenir "B-1 COS"
        elif low in ("sous-type","soustype","sous type","type","subtype"): m[c] = "Sous-type"
    return m

def read_visa_reference_tree(path: Path) -> pd.DataFrame:
    try:
        dfv = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=TREE_COLS)
    hdr_map = _norm_header_map(list(dfv.columns))
    if hdr_map: dfv = dfv.rename(columns=hdr_map)

    if "Catégorie" not in dfv.columns: dfv["Catégorie"] = ""
    dfv["Catégorie"] = dfv["Catégorie"].fillna("").astype(str).str.strip()
    if (dfv["Catégorie"] == "").any():
        dfv["Catégorie"] = dfv["Catégorie"].replace("", pd.NA).ffill().fillna("")

    out_rows = []
    if "VISA_FINAL" in dfv.columns:
        series = dfv["VISA_FINAL"].fillna("").astype(str).str.strip()
        for _, r in dfv.iterrows():
            cat = str(r.get("Catégorie","")).strip()
            final = str(r.get("VISA_FINAL","")).strip()
            if final == "": continue
            parts = final.split()
            if len(parts) == 1:
                visa_code = parts[0]; sous = ""
            else:
                visa_code = " ".join(parts[:-1]); sous = parts[-1].upper()
                if sous not in {"COS","EOS"}:
                    visa_code = final; sous = ""
            out_rows.append({"Catégorie":cat, "Visa":visa_code, "Sous-type":sous})
    else:
        if "Visa" not in dfv.columns: dfv["Visa"] = ""
        if "Sous-type" not in dfv.columns: dfv["Sous-type"] = ""
        for _, r in dfv.iterrows():
            cat  = str(r.get("Catégorie","")).strip()
            visa = str(r.get("Visa","")).strip()
            sous = str(r.get("Sous-type","")).strip().upper()
            if visa == "" and sous == "": continue
            if sous not in {"","COS","EOS"}:
                sous = re.sub(r"\s+","",sous).upper()
            out_rows.append({"Catégorie":cat, "Visa":visa, "Sous-type":sous})

    df = pd.DataFrame(out_rows, columns=TREE_COLS).fillna("")
    for c in TREE_COLS: df[c] = df[c].astype(str).str.strip()
    df = df.drop_duplicates().reset_index(drop=True)
    return df

def cascading_visa_picker_tree(df_ref: pd.DataFrame, key_prefix: str, init: dict | None = None) -> dict:
    """3 niveaux : Catégorie -> Visa -> Sous-type (COS/EOS)"""
    result = {"Catégorie":"", "Visa":"", "Sous-type":""}
    if df_ref is None or df_ref.empty:
        st.info("Référentiel Visa vide."); return result

    # 1) Catégorie
    dfC = df_ref.copy()
    cats = sorted([v for v in dfC["Catégorie"].unique() if v])
    idxC = 0
    if init and init.get("Catégorie","") in cats: idxC = cats.index(init["Catégorie"])+1
    result["Catégorie"] = st.selectbox("Catégorie", [""]+cats, index=idxC, key=f"{key_prefix}_cat")
    if result["Catégorie"]:
        dfC = dfC[dfC["Catégorie"] == result["Catégorie"]]

    # 2) Visa
    visas = sorted([v for v in dfC["Visa"].unique() if v])
    idxV = 0
    if init and init.get("Visa","") in visas: idxV = visas.index(init["Visa"])+1
    result["Visa"] = st.selectbox("Visa", [""]+visas, index=idxV, key=f"{key_prefix}_visa")
    dfV = dfC[dfC["Visa"] == result["Visa"]] if result["Visa"] else dfC.copy()

    # 3) Sous-type (COS/EOS) — optionnel
    sous = sorted([v for v in dfV["Sous-type"].unique() if v])
    idxS = 0
    if init and init.get("Sous-type","") in sous: idxS = sous.index(init["Sous-type"])+1
    result["Sous-type"] = st.selectbox("Sous-type (COS/EOS)", [""]+sous, index=idxS, key=f"{key_prefix}_soustype")

    if not visas:
        st.caption("Visa : (aucun pour cette catégorie)")
    elif result["Visa"] and not sous:
        st.caption(f"Visa : **{result['Visa']}** (pas de sous-type)")
    elif result["Visa"] and sous and not result["Sous-type"]:
        st.caption(f"Visa **{result['Visa']}** — sous-types possibles : {', '.join(sous)}")

    return result

def visas_autorises_from_tree(df_ref: pd.DataFrame, sel: dict) -> list[str]:
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
    """Disponible à transférer (honoraires payés - déjà transféré)."""
    hono = float(pd.to_numeric(pd.Series([row.get(HONO,0.0)]), errors="coerce").fillna(0.0).iloc[0])
    paid = float(pd.to_numeric(pd.Series([row.get("Payé",0.0)]), errors="coerce").fillna(0.0).iloc[0])
    transferred = float(pd.to_numeric(pd.Series([row.get(ESC_TR,0.0)]), errors="coerce").fillna(0.0).iloc[0])
    return float(max(min(paid, hono) - transferred, 0.0))

def append_escrow_journal(row: pd.Series, amount: float, note: str = "") -> str:
    lst = _parse_json_list(row.get(ESC_JR, ""))
    lst.append({"ts": datetime.now().isoformat(timespec="seconds"), "amount": float(amount), "note": _safe_str(note)})
    return json.dumps(lst, ensure_ascii=False)


# =========================
# VISA APP — PARTIE 2/3
# =========================

st.set_page_config(page_title="Visa Manager", layout="wide")

# ---------- Barre latérale : fichier ----------
st.sidebar.header("📂 Fichier Excel")
uploaded = st.sidebar.file_uploader("Charger/Remplacer fichier (.xlsx)", type=["xlsx"], key="uploader")
if uploaded is not None:
    p = set_current_file_from_upload(uploaded)
    if p: st.sidebar.success(f"Fichier chargé: {p.name}")

path_text = st.sidebar.text_input("Ou saisir un chemin existant", value=st.session_state.get("current_path",""))
colB1, colB2 = st.sidebar.columns(2)
if colB1.button("📄 Ouvrir ce fichier", key="open_file_btn"):
    p = Path(path_text)
    if p.exists():
        save_workspace_path(p); st.sidebar.success(f"Ouvert: {p.name}"); st.rerun()
    else:
        st.sidebar.error("Chemin invalide.")
if colB2.button("♻️ Reprendre le dernier", key="resume_last_btn"):
    p = _load_last_path()
    if p:
        save_workspace_path(p); st.sidebar.success(f"Repris: {p.name}"); st.rerun()
    else:
        st.sidebar.info("Aucun fichier précédent.")

current_path = current_file_path()
if current_path is None:
    st.warning("Aucun fichier sélectionné. Charge un .xlsx ou fournis un chemin valide."); st.stop()

# ---------- Feuilles ----------
sheets = list_sheets(current_path)
if not sheets:
    st.error("Impossible de lire le classeur."); st.stop()

st.sidebar.markdown("---"); st.sidebar.write("**Feuilles détectées :**")
for i, sn in enumerate(sheets): st.sidebar.write(f"- {i+1}. {sn}")

# feuille clients par défaut
client_target_sheet = None
for sn in sheets:
    df_try = read_sheet(current_path, sn, normalize=False)
    if {"Nom","Visa"}.issubset(set(df_try.columns.astype(str))):
        client_target_sheet = sn; break

sheet_choice = st.sidebar.selectbox(
    "Feuille à afficher sur le Dashboard :",
    sheets,
    index=max(0, sheets.index(client_target_sheet) if client_target_sheet in sheets else 0),
    key="sheet_choice"
)

# ---------- Titre & tabs ----------
st.title("🛂 Visa Manager — US $")
tabs = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"])

# ---------- Référentiel Visa (arbre 3 niveaux) ----------
visa_ref_tree   = read_visa_reference_tree(current_path)  # Catégorie / Visa / Sous-type
visa_ref_simple = read_visa_reference(current_path)       # Catégorie / Visa (pour mapping)

# ================= DASHBOARD =================
with tabs[0]:
    df_raw = read_sheet(current_path, sheet_choice, normalize=False)

    # Si c'est la feuille Visa, on montre la référence telle quelle
    if looks_like_reference(df_raw) and sheet_choice == "Visa":
        st.subheader("📄 Référentiel — Catégorie / Visa / Sous-type")
        st.dataframe(visa_ref_tree, use_container_width=True)
        st.stop()

    df = read_sheet(current_path, sheet_choice, normalize=True, visa_ref=visa_ref_simple)

    # Filtres hiérarchiques + dates (keys uniques 'dash_*')
    st.markdown("### 🔎 Filtres (Catégorie → Visa → Sous-type)")
    with st.container():
        cTopL, cTopR = st.columns([1,2])
        show_all = cTopL.checkbox("Afficher tous les dossiers", value=False, key="dash_show_all")
        cTopL.caption("Sélection hiérarchique")
        with cTopL:
            sel_path_dash = cascading_visa_picker_tree(visa_ref_tree, key_prefix="dash_tree")
        visas_aut = visas_autorises_from_tree(visa_ref_tree, sel_path_dash)

        cR1, cR2, cR3 = cTopR.columns(3)
        years  = sorted({d.year for d in df["Date"] if pd.notna(d)}) if "Date" in df.columns else []
        months = sorted(df["Mois"].dropna().unique()) if "Mois" in df.columns else []
        sel_years  = cR1.multiselect("Année", years, default=[], key="dash_years")
        sel_months = cR2.multiselect("Mois (MM)", months, default=[], key="dash_months")
        include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True, key="dash_na")

    f = df.copy()
    if not show_all:
        for col in ["Catégorie","Visa","Sous-type"]:
            if col in f.columns:
                val = _safe_str(sel_path_dash.get(col,""))
                if val:
                    f = f[f[col].astype(str) == val]
        if "Visa" in f.columns and visas_aut:
            f = f[f["Visa"].astype(str).isin(visas_aut)]

    if "Date" in f.columns and sel_years:
        mask = f["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask |= f["Date"].isna()
        f = f[mask]
    if "Mois" in f.columns and sel_months:
        mask = f["Mois"].isin(sel_months)
        if include_na_dates: mask |= f["Mois"].isna()
        f = f[mask]

    hidden = len(df) - len(f)
    if hidden > 0: st.caption(f"🔎 {hidden} ligne(s) masquée(s) par les filtres.")

    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Total (US $)", _fmt_money_us(float(f.get(TOTAL, pd.Series(dtype=float)).sum())) )
    k3.metric("Payé (US $)", _fmt_money_us(float(f.get("Payé", pd.Series(dtype=float)).sum())) )
    k4.metric("Solde (US $)", _fmt_money_us(float(f.get("Reste", pd.Series(dtype=float)).sum())) )
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    st.subheader("📋 Données (aperçu)")
    cols_show = [c for c in [
        DOSSIER_COL,"ID_Client","Nom","Date","Mois",
        "Catégorie","Visa","Sous-type",
        HONO, AUTRE, TOTAL, "Payé","Reste",
        S_ENVOYE, D_ENVOYE, S_APPROUVE, D_APPROUVE, S_RFE, D_RFE, S_REFUSE, D_REFUSE, S_ANNULE, D_ANNULE
    ] if c in f.columns]
    view = f.copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in view.columns: view[col] = view[col].map(_fmt_money_us)
    if "Date" in view.columns: view["Date"] = view["Date"].astype(str)
    st.dataframe(view[cols_show], use_container_width=True)


# ==========================
# 📊 ANALYSES — Volume & Financier
# ==========================
st.markdown("## 📊 Analyses — Volumes & Financier")

if "Visa" not in live.columns:
    st.warning("Aucune donnée Visa disponible.")
else:
    st.divider()
    st.subheader("Filtres d’analyse")

    cL, cR = st.columns([1,2])
    show_all_A = cL.checkbox("Afficher tous les dossiers", value=False, key="anal_show_all")

    # cascade Catégorie → Visa → Sous-type
    cL.caption("Sélection hiérarchique (Catégorie → Visa → Sous-type)")
    with cL:
        sel_path_anal = cascading_visa_picker_tree(visa_ref_tree, key_prefix="anal_tree")
    visas_aut_A = visas_autorises_from_tree(visa_ref_tree, sel_path_anal)

    # Filtres temporels
    cR1, cR2, cR3 = cR.columns(3)
    if "Date" in live.columns:
        yearsA  = sorted({d.year for d in pd.to_datetime(live["Date"], errors="coerce").dropna()})
    else:
        yearsA = []
    monthsA = [f"{m:02d}" for m in range(1,13)]

    sel_years  = cR1.multiselect("Année", yearsA, default=[], key="anal_years")
    sel_months = cR2.multiselect("Mois (MM)", monthsA, default=[], key="anal_months")
    include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True, key="anal_na")

    # filtrage du jeu
    dfA = live.copy()
    if visas_aut_A and not show_all_A:
        dfA = dfA[dfA["Visa"].isin(visas_aut_A)]
    if sel_years:
        dfA = dfA[dfA["Date"].apply(lambda d: pd.to_datetime(d, errors="coerce").year if pd.notna(d) else None).isin(sel_years)]
    if sel_months:
        dfA = dfA[dfA["Date"].apply(lambda d: f"{pd.to_datetime(d, errors='coerce').month:02d}" if pd.notna(d) else None).isin(sel_months)]

    st.divider()
    st.subheader("🔢 Indicateurs clés")

    if dfA.empty:
        st.info("Aucune donnée pour la période ou les filtres choisis.")
    else:
        tot_dossiers = len(dfA)
        solde_zero = (dfA["Reste"] <= 0.01).sum() if "Reste" in dfA.columns else 0
        reste_total = dfA["Reste"].sum() if "Reste" in dfA.columns else 0.0
        total_paye  = dfA["Payé"].sum() if "Payé" in dfA.columns else 0.0
        total_hono  = dfA["Montant honoraires (US $)"].sum() if "Montant honoraires (US $)" in dfA.columns else 0.0
        taux_solde  = (solde_zero / tot_dossiers * 100) if tot_dossiers else 0

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total dossiers", tot_dossiers)
        c2.metric("Dossiers soldés", solde_zero, f"{taux_solde:.0f}%")
        c3.metric("Honoraires totaux", _fmt_money_us(total_hono))
        c4.metric("Encaissements", _fmt_money_us(total_paye))
        c5.metric("Solde restant", _fmt_money_us(reste_total))

        st.divider()
        st.subheader("📅 Dossiers par mois et par visa")

        if "Date" in dfA.columns:
            dfA["Mois"] = pd.to_datetime(dfA["Date"], errors="coerce").dt.to_period("M").astype(str)
            grp = dfA.groupby(["Mois", "Visa"]).size().reset_index(name="Nb dossiers")

            chart_vol = alt.Chart(grp).mark_bar().encode(
                x=alt.X("Mois:N", title="Mois"),
                y=alt.Y("Nb dossiers:Q", title="Nombre de dossiers"),
                color="Visa:N",
                tooltip=["Mois", "Visa", "Nb dossiers"]
            ).properties(height=350)
            st.altair_chart(chart_vol, use_container_width=True)

        st.divider()
        st.subheader("💼 Détail des dossiers analysés")

        st.dataframe(
            dfA[["ID_Client", "Nom", "Visa", "Date", "Montant honoraires (US $)", "Payé", "Reste"]],
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            "⬇️ Télécharger les données filtrées (Excel)",
            data=_to_excel_bytes(dfA),
            file_name="analyse_filtrée.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ==========================
# 💾 SAUVEGARDE AUTOMATIQUE
# ==========================
st.divider()
st.markdown("### 💾 Sauvegarde et restauration automatique")

if "current_path" in locals() and current_path.exists():
    auto_save_path = Path("last_used_file.txt")
    try:
        auto_save_path.write_text(str(current_path), encoding="utf-8")
        st.caption(f"✅ Chemin du dernier fichier sauvegardé : `{current_path.name}`")
    except Exception as e:
        st.error(f"Erreur lors de la sauvegarde du chemin : {e}")
else:
    st.caption("⚠️ Aucun fichier actif à sauvegarder.")

st.divider()
st.success("✔️ Application Visa — Tous les modules sont chargés (Clients, Visa, ESCROW, Analyses).")