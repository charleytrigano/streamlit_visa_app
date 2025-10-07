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




# ================================
# 👤 ONGLET 2 — CLIENTS (CRUD + paiements)
# ================================
with tabs[1]:
    st.subheader("👤 Clients — Créer / Modifier / Supprimer (écriture directe)")

    if client_target_sheet is None:
        st.warning("Aucune feuille *Clients* valide (au minimum colonnes Nom & Visa).")
        st.stop()

    # Rechargement rapide
    if st.button("🔄 Recharger le fichier", key="crud_reload"):
        st.rerun()

    live_raw = read_sheet(current_path, client_target_sheet, normalize=False).copy()
    live_raw = ensure_dossier_numbers(live_raw)
    live_raw["_RowID"] = range(len(live_raw))  # identifiant interne pour la sélection
    action = st.radio("Action", ["Créer", "Modifier", "Supprimer"], horizontal=True, key="crud_action")

    # ---------- CRÉER ----------
    if action == "Créer":
        st.markdown("### ➕ Nouveau client")

        # Colonnes minimales garanties
        needed = [DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
                  "Catégorie", "Visa", "Sous-type",
                  HONO, AUTRE, TOTAL, "Payé", "Reste",
                  "Paiements", ESC_TR, ESC_JR] + STATUS_COLS + STATUS_DATES
        for c in needed:
            if c not in live_raw.columns:
                if c in [HONO, AUTRE, TOTAL, "Payé", "Reste", ESC_TR]: live_raw[c] = 0.0
                elif c in STATUS_COLS: live_raw[c] = False
                else: live_raw[c] = ""

        next_num = next_dossier_number(live_raw)

        with st.form("create_form", clear_on_submit=False):
            c0, c1, c2 = st.columns([1, 1, 1])
            c0.metric("Prochain Dossier N", f"{next_num}")
            nom_in = c1.text_input("Nom")
            d = c2.date_input("Date", value=date.today())

            st.caption("Sélection hiérarchique (Catégorie → Visa → Sous-type)")
            sel_path = cascading_visa_picker_tree(visa_ref_tree, key_prefix="create_tree")
            cat = sel_path.get("Catégorie", "")
            visa = sel_path.get("Visa", "")
            stype = sel_path.get("Sous-type", "")

            c5, c6 = st.columns(2)
            honoraires = c5.number_input(HONO, value=0.0, step=10.0, format="%.2f")
            autres = c6.number_input(AUTRE, value=0.0, step=10.0, format="%.2f")

            st.markdown("#### État du dossier (avec dates)")
            r1c1, r1c2 = st.columns(2)
            v_env = r1c1.checkbox(S_ENVOYE, value=False, key="cre_env")
            dt_env = r1c2.date_input(D_ENVOYE, value=date.today(), disabled=not v_env, key="cre_dt_env")
            r2c1, r2c2 = st.columns(2)
            v_app = r2c1.checkbox(S_APPROUVE, value=False, key="cre_app")
            dt_app = r2c2.date_input(D_APPROUVE, value=date.today(), disabled=not v_app, key="cre_dt_app")
            r3c1, r3c2 = st.columns(2)
            v_rfe = r3c1.checkbox(S_RFE, value=False, key="cre_rfe")
            dt_rfe = r3c2.date_input(D_RFE, value=date.today(), disabled=not v_rfe, key="cre_dt_rfe")
            r4c1, r4c2 = st.columns(2)
            v_ref = r4c1.checkbox(S_REFUSE, value=False, key="cre_ref")
            dt_ref = r4c2.date_input(D_REFUSE, value=date.today(), disabled=not v_ref, key="cre_dt_ref")
            r5c1, r5c2 = st.columns(2)
            v_ann = r5c1.checkbox(S_ANNULE, value=False, key="cre_ann")
            dt_ann = r5c2.date_input(D_ANNULE, value=date.today(), disabled=not v_ann, key="cre_dt_ann")

            ok = st.form_submit_button("💾 Sauvegarder (dans le fichier)", type="primary")

        if ok:
            if v_rfe and not (v_env or v_ref or v_ann):
                st.error("RFE ⇢ seulement si *Dossier envoyé* ou *refusé/annulé* est coché.")
                st.stop()

            # Nom unique (suffixes -0, -1, … si besoin)
            existing_names = set(live_raw["Nom"].dropna().astype(str))
            base_name = _safe_str(nom_in)
            use_name = base_name
            if base_name in existing_names:
                k = 0
                while f"{base_name}-{k}" in existing_names:
                    k += 1
                use_name = f"{base_name}-{k}"

            # ID client unique (Nom + Date)
            gen_id = _make_client_id_from_row({"Nom": use_name, "Date": d})
            existing_ids = set(live_raw["ID_Client"].astype(str)) if "ID_Client" in live_raw.columns else set()
            new_id, n = gen_id, 1
            while new_id in existing_ids:
                n += 1
                new_id = f"{gen_id}-{n:02d}"

            total = float((honoraires or 0.0) + (autres or 0.0))
            new_row = {
                DOSSIER_COL: int(next_num),
                "ID_Client": new_id,
                "Nom": use_name,
                "Date": str(d),
                "Mois": f"{d.month:02d}",
                "Catégorie": _safe_str(cat),
                "Visa": _safe_str(visa),
                "Sous-type": _safe_str(stype),
                HONO: float(honoraires or 0.0),
                AUTRE: float(autres or 0.0),
                TOTAL: total,
                "Payé": 0.0,
                "Reste": max(total, 0.0),
                "Paiements": "",
                ESC_TR: 0.0,
                ESC_JR: "",
                S_ENVOYE: bool(v_env),   D_ENVOYE:   (str(dt_env) if v_env else ""),
                S_APPROUVE: bool(v_app), D_APPROUVE: (str(dt_app) if v_app else ""),
                S_RFE: bool(v_rfe),      D_RFE:      (str(dt_rfe) if v_rfe else ""),
                S_REFUSE: bool(v_ref),   D_REFUSE:   (str(dt_ref) if v_ref else ""),
                S_ANNULE: bool(v_ann),   D_ANNULE:   (str(dt_ann) if v_ann else "")
            }

            live_after = pd.concat([live_raw.drop(columns=["_RowID"]), pd.DataFrame([new_row])], ignore_index=True)
            live_after = ensure_dossier_numbers(live_after)
            write_sheet_inplace(current_path, client_target_sheet, live_after)
            save_workspace_path(current_path)
            st.success(f"Client créé **dans le fichier** (Dossier N {next_num}). ✅")
            st.rerun()

    # ---------- MODIFIER ----------
    if action == "Modifier":
        st.markdown("### ✏️ Modifier un client — fiche, statuts & paiements")

        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            # Sélection d'une ligne
            options = [(int(r["_RowID"]),
                        int(r.name),
                        f'{int(r.get(DOSSIER_COL,0))} — { _safe_str(r.get("ID_Client")) } — { _safe_str(r.get("Nom")) }')
                       for _, r in live_raw.iterrows()]
            labels = [lab for _, __, lab in options]
            sel_lab = st.selectbox("Sélection", labels, key="edit_sel_label")
            sel_rowid, orig_pos, _ = [t for t in options if t[2] == sel_lab][0]
            idx = live_raw.index[live_raw["_RowID"] == sel_rowid][0]
            init = live_raw.loc[idx].to_dict()

            # Formulaire de modif
            with st.form(f"edit_form_{sel_rowid}", clear_on_submit=False):
                c0, c1, c2 = st.columns([1, 1, 1])
                c0.metric("Dossier N", f'{int(init.get(DOSSIER_COL, 0))}')
                nom = c1.text_input("Nom", value=_safe_str(init.get("Nom")), key=f"edit_nom_{sel_rowid}")
                try:
                    d_init = pd.to_datetime(init.get("Date")).date() if _safe_str(init.get("Date")) else date.today()
                except Exception:
                    d_init = date.today()
                d = c2.date_input("Date", value=d_init, key=f"edit_date_{sel_rowid}")

                st.caption("Sélection hiérarchique (Catégorie → Visa → Sous-type)")
                init_path = {
                    "Catégorie": _safe_str(init.get("Catégorie")),
                    "Visa": _safe_str(init.get("Visa")),
                    "Sous-type": _safe_str(init.get("Sous-type")).upper(),
                }
                sel_path = cascading_visa_picker_tree(visa_ref_tree, key_prefix=f"edit_tree_{sel_rowid}", init=init_path)
                cat = sel_path.get("Catégorie", "")
                visa = sel_path.get("Visa", "")
                stype = sel_path.get("Sous-type", "")

                def _flt(x, alt=0.0):
                    try:
                        return float(x)
                    except Exception:
                        return float(alt)

                hono0 = _flt(init.get(HONO, init.get("Montant", 0.0)))
                autre0 = _flt(init.get(AUTRE, 0.0))
                paye0 = _flt(init.get("Payé", 0.0))

                c5, c6 = st.columns(2)
                honoraires = c5.number_input(HONO, value=hono0, step=10.0, format="%.2f", key=f"edit_hono_{sel_rowid}")
                autres = c6.number_input(AUTRE, value=autre0, step=10.0, format="%.2f", key=f"edit_autre_{sel_rowid}")
                total_preview = float(honoraires + autres)
                st.caption(f"Payé actuel : {_fmt_money_us(paye0)} — Solde à venir : {_fmt_money_us(max(total_preview - paye0, 0.0))}")

                st.markdown("#### État du dossier (avec dates)")
                def _get_dt(key):
                    v = _safe_str(init.get(key))
                    try:
                        return pd.to_datetime(v).date() if v else date.today()
                    except Exception:
                        return date.today()

                r1c1, r1c2 = st.columns(2)
                v_env = r1c1.checkbox(S_ENVOYE, value=bool(init.get(S_ENVOYE)), key=f"edit_env_{sel_rowid}")
                dt_env = r1c2.date_input(D_ENVOYE, value=_get_dt(D_ENVOYE), disabled=not v_env, key=f"edit_dt_env_{sel_rowid}")

                r2c1, r2c2 = st.columns(2)
                v_app = r2c1.checkbox(S_APPROUVE, value=bool(init.get(S_APPROUVE)), key=f"edit_app_{sel_rowid}")
                dt_app = r2c2.date_input(D_APPROUVE, value=_get_dt(D_APPROUVE), disabled=not v_app, key=f"edit_dt_app_{sel_rowid}")

                r3c1, r3c2 = st.columns(2)
                v_rfe = r3c1.checkbox(S_RFE, value=bool(init.get(S_RFE)), key=f"edit_rfe_{sel_rowid}")
                dt_rfe = r3c2.date_input(D_RFE, value=_get_dt(D_RFE), disabled=not v_rfe, key=f"edit_dt_rfe_{sel_rowid}")

                r4c1, r4c2 = st.columns(2)
                v_ref = r4c1.checkbox(S_REFUSE, value=bool(init.get(S_REFUSE)), key=f"edit_ref_{sel_rowid}")
                dt_ref = r4c2.date_input(D_REFUSE, value=_get_dt(D_REFUSE), disabled=not v_ref, key=f"edit_dt_ref_{sel_rowid}")

                r5c1, r5c2 = st.columns(2)
                v_ann = r5c1.checkbox(S_ANNULE, value=bool(init.get(S_ANNULE)), key=f"edit_ann_{sel_rowid}")
                dt_ann = r5c2.date_input(D_ANNULE, value=_get_dt(D_ANNULE), disabled=not v_ann, key=f"edit_dt_ann_{sel_rowid}")

                ok_fiche = st.form_submit_button("💾 Enregistrer la fiche (dans le fichier)", type="primary")

            if ok_fiche:
                if v_rfe and not (v_env or v_ref or v_ann):
                    st.error("RFE ⇢ seulement si *Dossier envoyé* ou *refusé/annulé* est coché.")
                    st.stop()

                live = read_sheet(current_path, client_target_sheet, normalize=False).copy()

                # Retrouver la ligne à mettre à jour (ID_Client prioritaire → Dossier N → fallback pos)
                t_idx = None
                key_id = _safe_str(init.get("ID_Client"))
                if key_id and "ID_Client" in live.columns:
                    hits = live.index[live["ID_Client"].astype(str) == key_id]
                    if len(hits) > 0:
                        t_idx = hits[0]
                if t_idx is None and (DOSSIER_COL in live.columns) and (init.get(DOSSIER_COL) not in [None, ""]):
                    try:
                        num = int(_to_int(pd.Series([init.get(DOSSIER_COL)])).iloc[0])
                        hits = live.index[_to_int(live[DOSSIER_COL]) == num]
                        if len(hits) > 0:
                            t_idx = hits[0]
                    except Exception:
                        pass
                if t_idx is None and (orig_pos is not None) and 0 <= int(orig_pos) < len(live):
                    t_idx = int(orig_pos)
                if t_idx is None:
                    st.error("Ligne introuvable.")
                    st.stop()

                total = float((honoraires or 0.0) + (autres or 0.0))
                for c in [HONO, AUTRE, TOTAL, "Payé", "Reste", "Paiements", ESC_TR, ESC_JR,
                          "Nom", "Date", "Mois", "Catégorie", "Visa", "Sous-type"] + STATUS_COLS + STATUS_DATES + [DOSSIER_COL]:
                    if c not in live.columns:
                        live[c] = 0.0 if c in [HONO, AUTRE, TOTAL, "Payé", "Reste", ESC_TR] else (False if c in STATUS_COLS else "")

                # MAJ valeurs
                live.at[t_idx, "Nom"] = _safe_str(nom)
                live.at[t_idx, "Date"] = str(d)
                live.at[t_idx, "Mois"] = f"{d.month:02d}"
                live.at[t_idx, "Catégorie"] = _safe_str(cat)
                live.at[t_idx, "Visa"] = _safe_str(visa)
                live.at[t_idx, "Sous-type"] = _safe_str(stype)
                live.at[t_idx, HONO] = float(honoraires or 0.0)
                live.at[t_idx, AUTRE] = float(autres or 0.0)

                # Statuts & dates
                live.at[t_idx, S_ENVOYE] = bool(v_env);   live.at[t_idx, D_ENVOYE] = (str(dt_env) if v_env else "")
                live.at[t_idx, S_APPROUVE] = bool(v_app); live.at[t_idx, D_APPROUVE] = (str(dt_app) if v_app else "")
                live.at[t_idx, S_RFE] = bool(v_rfe);      live.at[t_idx, D_RFE] = (str(dt_rfe) if v_rfe else "")
                live.at[t_idx, S_REFUSE] = bool(v_ref);   live.at[t_idx, D_REFUSE] = (str(dt_ref) if v_ref else "")
                live.at[t_idx, S_ANNULE] = bool(v_ann);   live.at[t_idx, D_ANNULE] = (str(dt_ann) if v_ann else "")

                # Recalc payé / reste via JSON paiements
                pay_json = live.at[t_idx, "Paiements"]
                paid = _sum_payments(_parse_json_list(pay_json))
                live.at[t_idx, "Payé"] = float(paid)
                live.at[t_idx, TOTAL] = total
                live.at[t_idx, "Reste"] = max(total - float(paid), 0.0)

                live = ensure_dossier_numbers(live)
                write_sheet_inplace(current_path, client_target_sheet, live)
                save_workspace_path(current_path)
                st.success("Fiche enregistrée **dans le fichier**. ✅")
                st.rerun()

            # ----- Paiements / Historique -----
            live_now = read_sheet(current_path, client_target_sheet, normalize=False)
            ixs = live_now.index[live_now.get("ID_Client", "").astype(str) == _safe_str(init.get("ID_Client"))]
            st.markdown("#### 💳 Historique & nouveaux règlements")
            if len(ixs) == 0:
                st.info("Ligne introuvable pour les paiements.")
            else:
                i = ixs[0]
                if "Paiements" not in live_now.columns:
                    live_now["Paiements"] = ""
                plist = _parse_json_list(live_now.at[i, "Paiements"])

                if plist:
                    dfp = pd.DataFrame(plist)
                    if "date" in dfp.columns:
                        dfp["date"] = pd.to_datetime(dfp["date"], errors="coerce").dt.date.astype(str)
                    if "amount" in dfp.columns:
                        dfp["Montant ($)"] = dfp["amount"].apply(lambda x: _fmt_money_us(float(x) if pd.notna(x) else 0.0))
                    for col in ["mode", "note"]:
                        if col not in dfp.columns:
                            dfp[col] = ""
                    show = dfp[["date", "mode", "Montant ($)", "note"]] if set(["date", "mode", "note"]).issubset(dfp.columns) else dfp
                    with st.expander("Historique des règlements", expanded=True):
                        st.table(show.rename(columns={"date": "Date", "mode": "Mode", "note": "Note"}))
                        if len(plist) > 0:
                            del_idx = st.number_input("Supprimer la ligne n° (1..n)", min_value=1, max_value=len(plist), value=1, step=1, key=f"del_pay_idx_{i}")
                            if st.button("🗑️ Supprimer cette ligne", key=f"del_pay_btn_{i}"):
                                try:
                                    del plist[int(del_idx) - 1]
                                    live_now.at[i, "Paiements"] = json.dumps(plist, ensure_ascii=False)
                                    total_paid = _sum_payments(plist)
                                    hono = _to_num(pd.Series([live_now.at[i, HONO] if HONO in live_now.columns else 0.0])).iloc[0]
                                    autr = _to_num(pd.Series([live_now.at[i, AUTRE] if AUTRE in live_now.columns else 0.0])).iloc[0]
                                    total = float(hono + autr)
                                    live_now.at[i, "Payé"] = float(total_paid)
                                    live_now.at[i, "Reste"] = max(total - float(total_paid), 0.0)
                                    live_now.at[i, TOTAL] = total
                                    write_sheet_inplace(current_path, client_target_sheet, live_now)
                                    st.success("Ligne supprimée et soldes recalculés. ✅")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Erreur suppression : {e}")
                else:
                    st.caption("Aucun paiement enregistré pour ce client.")

                cA, cB, cC, cD = st.columns([1, 1, 1, 2])
                pay_date = cA.date_input("Date", value=date.today(), key=f"pay_date_{i}")
                pay_mode = cB.selectbox("Mode", ["CB", "Chèque", "Espèces", "Virement", "Venmo", "Autre"], key=f"pay_mode_{i}")
                pay_amt = cC.number_input("Montant ($)", min_value=0.0, step=10.0, format="%.2f", key=f"pay_amt_{i}")
                pay_note = cD.text_input("Note", "", key=f"pay_note_{i}")
                if st.button("💾 Enregistrer ce règlement (dans le fichier)", key=f"pay_add_btn_{i}"):
                    try:
                        add = float(pay_amt or 0.0)
                        if add <= 0:
                            st.warning("Le montant doit être > 0.")
                            st.stop()
                        norm = normalize_dataframe(live_now.copy(), visa_ref=read_visa_reference(current_path))
                        mask_id = norm["ID_Client"].astype(str) == _safe_str(init.get("ID_Client"))
                        reste_curr = float(norm.loc[mask_id, "Reste"].sum()) if mask_id.any() else 0.0
                        if add > reste_curr + 1e-9:
                            add = reste_curr
                        plist.append({"date": str(pay_date), "amount": float(add), "mode": pay_mode, "note": pay_note})
                        live_now.at[i, "Paiements"] = json.dumps(plist, ensure_ascii=False)
                        total_paid = _sum_payments(plist)
                        hono = _to_num(pd.Series([live_now.at[i, HONO] if HONO in live_now.columns else 0.0])).iloc[0]
                        autr = _to_num(pd.Series([live_now.at[i, AUTRE] if AUTRE in live_now.columns else 0.0])).iloc[0]
                        total = float(hono + autr)
                        live_now.at[i, "Payé"] = float(total_paid)
                        live_now.at[i, "Reste"] = max(total - float(total_paid), 0.0)
                        live_now.at[i, TOTAL] = total
                        write_sheet_inplace(current_path, client_target_sheet, live_now)
                        st.success("Règlement ajouté. ✅")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")

    # ---------- SUPPRIMER ----------
    if action == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            options = [(int(r["_RowID"]),
                        f'{int(r.get(DOSSIER_COL,0))} — { _safe_str(r.get("ID_Client")) } — { _safe_str(r.get("Nom")) }')
                       for _, r in live_raw.iterrows()]
            labels = [lab for _, lab in options]
            sel_lab = st.selectbox("Sélection", labels, key="del_select")
            sel_rowid = [rid for rid, lab in options if lab == sel_lab][0]
            idx = live_raw.index[live_raw["_RowID"] == sel_rowid][0]
            st.error("⚠️ Action irréversible.")
            if st.button("Supprimer (dans le fichier)", key="del_btn"):
                live = live_raw.drop(columns=["_RowID"]).copy()
                key = _safe_str(live_raw.at[idx, "ID_Client"])
                if key and "ID_Client" in live.columns:
                    live = live[live["ID_Client"].astype(str) != key].reset_index(drop=True)
                else:
                    nom = _safe_str(live_raw.at[idx, "Nom"])
                    dat = _safe_str(live_raw.at[idx, "Date"])
                    live = live[~((live.get("Nom", "").astype(str) == nom) &
                                  (live.get("Date", "").astype(str) == dat))].reset_index(drop=True)
                live = ensure_dossier_numbers(live)
                write_sheet_inplace(current_path, client_target_sheet, live)
                save_workspace_path(current_path)
                st.success("Client supprimé **dans le fichier**. ✅")
                st.rerun()

# ====================================
# 📊 ONGLET 3 — ANALYSES (filtres + charts)
# ====================================
with tabs[2]:
    st.subheader("📊 Analyses — Volumes & Financier")

    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa).")
        st.stop()

    visa_ref_simple = read_visa_reference(current_path)
    dfA_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    dfA = normalize_dataframe(dfA_raw, visa_ref=visa_ref_simple).copy()
    if dfA.empty:
        st.info("Aucune donnée pour analyser.")
        st.stop()

    # Filtres (Catégorie → Visa → Sous-type) + Année / Mois
    with st.container():
        cL, cR = st.columns([1, 2])
        show_all_A = cL.checkbox("Afficher tous les dossiers", value=False, key="anal_show_all")

        cL.caption("Sélection hiérarchique (Catégorie → Visa → Sous-type)")
        with cL:
            sel_path_anal = cascading_visa_picker_tree(visa_ref_tree, key_prefix="anal_tree")
        visas_aut_A = visas_autorises_from_tree(visa_ref_tree, sel_path_anal)

        cR1, cR2, cR3 = cR.columns(3)
        yearsA = sorted({d.year for d in dfA["Date"] if pd.notna(d)}) if "Date" in dfA.columns else []
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        sel_years = cR1.multiselect("Année", yearsA, default=[], key="anal_years")
        sel_months = cR2.multiselect("Mois (MM)", monthsA, default=[], key="anal_months")
        include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True, key="anal_na")

    # Application des filtres
    fA = dfA.copy()
    if not show_all_A:
        for col in ["Catégorie", "Visa", "Sous-type"]:
            if col in fA.columns:
                val = _safe_str(sel_path_anal.get(col, ""))
                if val:
                    fA = fA[fA[col].astype(str) == val]
        if "Visa" in fA.columns and visas_aut_A:
            fA = fA[fA["Visa"].astype(str).isin(visas_aut_A)]

    if "Date" in fA.columns and sel_years:
        mask_year = fA["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates:
            mask_year |= fA["Date"].isna()
        fA = fA[mask_year]
    if "Mois" in fA.columns and sel_months:
        mask_month = fA["Mois"].isin(sel_months)
        if include_na_dates:
            mask_month |= fA["Mois"].isna()
        fA = fA[mask_month]

    # KPI
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.1rem}
           .small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(fA)}")
    k2.metric("Total (US $)", _fmt_money_us(float(fA.get(TOTAL, pd.Series(dtype=float)).sum())))
    k3.metric("Payé (US $)", _fmt_money_us(float(fA.get("Payé", pd.Series(dtype=float)).sum())))
    k4.metric("Solde (US $)", _fmt_money_us(float(fA.get("Reste", pd.Series(dtype=float)).sum())))
    st.markdown('</div>', unsafe_allow_html=True)

    # Volumes / période (création)
    st.markdown("### 📈 Volumes de créations")
    fA["Periode"] = fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")
    vol_crees = fA.groupby("Periode").size().reset_index(name="Créés")
    if not vol_crees.empty:
        try:
            st.altair_chart(
                alt.Chart(vol_crees).mark_line(point=True).encode(
                    x=alt.X("Periode:N", sort=None, title="Période"),
                    y=alt.Y("Créés:Q"),
                    tooltip=["Periode", "Créés"]
                ).properties(height=260), use_container_width=True
            )
        except Exception:
            st.dataframe(vol_crees, use_container_width=True)

    st.divider()
    st.markdown("### 🔎 Détails (clients)")
    cols = [c for c in ["Periode", DOSSIER_COL, "ID_Client", "Nom",
                        "Catégorie", "Visa", "Sous-type",
                        "Date", HONO, AUTRE, TOTAL, "Payé", "Reste"] if c in fA.columns]
    details = fA[cols].copy()
    for col in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if col in details.columns:
            details[col] = details[col].apply(lambda x: _fmt_money_us(x) if pd.notna(x) else "")
    st.dataframe(details.sort_values(["Periode", "Catégorie", "Nom"]), use_container_width=True)

# =================================
# 🏦 ONGLET 4 — ESCROW (transferts)
# =================================
with tabs[3]:
    st.subheader("🏦 ESCROW — dépôts honoraires & transferts")

    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide.")
        st.stop()

    live_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    live = normalize_dataframe(live_raw, visa_ref=read_visa_reference(current_path)).copy()
    if ESC_TR not in live.columns:
        live[ESC_TR] = 0.0
    else:
        live[ESC_TR] = pd.to_numeric(live[ESC_TR], errors="coerce").fillna(0.0)
    live["ESCROW dispo"] = live.apply(escrow_available_from_row, axis=1)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Dossiers", f"{len(live)}")
    c2.metric("ESCROW total dispo", _fmt_money_us(float(live["ESCROW dispo"].sum())))
    envoyes = live[(live[S_ENVOYE] == True)]
    a_transferer = envoyes[envoyes["ESCROW dispo"] > 0.004]
    c3.metric("Dossiers envoyés (à réclamer)", f"{len(a_transferer)}")
    c4.metric("Montant à réclamer", _fmt_money_us(float(a_transferer["ESCROW dispo"].sum())))

    st.divider()
    st.markdown("### 📌 À transférer (dossiers **envoyés**)")
    if a_transferer.empty:
        st.success("Aucun transfert en attente pour des dossiers envoyés.")
    else:
        for _, r in a_transferer.sort_values("Date").iterrows():
            with st.expander(f'🔔 {r[DOSSIER_COL]} — {r["ID_Client"]} — {r["Nom"]} — {r.get("Catégorie","")} / {r["Visa"]} — ESCROW dispo: {_fmt_money_us(r["ESCROW dispo"])}'):
                cA, cB, cC = st.columns(3)
                cA.metric("Honoraires", _fmt_money_us(float(r.get(HONO, 0.0))))
                cB.metric("Déjà transféré", _fmt_money_us(float(r.get(ESC_TR, 0.0))))
                cC.metric("Payé", _fmt_money_us(float(r.get("Payé", 0.0))))
                amt = st.number_input("Montant à marquer comme transféré (US $)",
                                      min_value=0.0, value=float(r["ESCROW dispo"]),
                                      step=10.0, format="%.2f", key=f"esc_amt_{r['ID_Client']}")
                note = st.text_input("Note (facultatif)", "", key=f"esc_note_{r['ID_Client']}")
                if st.button("✅ Marquer transféré (écrit dans le fichier)", key=f"esc_btn_{r['ID_Client']}"):
                    try:
                        live_w = read_sheet(current_path, client_target_sheet, normalize=False).copy()
                        for c in [ESC_TR, ESC_JR]:
                            if c not in live_w.columns:
                                live_w[c] = 0.0 if c == ESC_TR else ""
                        idxs = live_w.index[live_w.get("ID_Client", "").astype(str) == str(r["ID_Client"])]
                        if len(idxs) == 0:
                            st.error("Ligne introuvable.")
                            st.stop()
                        i = idxs[0]
                        tmp = normalize_dataframe(live_w.copy(), visa_ref=read_visa_reference(current_path))
                        disp = float(tmp.loc[tmp["ID_Client"].astype(str) == str(r["ID_Client"]), :].apply(escrow_available_from_row, axis=1).iloc[0])
                        add = float(min(max(amt, 0.0), disp))
                        live_w.at[i, ESC_TR] = float(pd.to_numeric(pd.Series([live_w.at[i, ESC_TR]]), errors="coerce").fillna(0.0).iloc[0] + add)
                        live_w.at[i, ESC_JR] = append_escrow_journal(live_w.loc[i], add, note)
                        live_w = ensure_dossier_numbers(live_w)
                        write_sheet_inplace(current_path, client_target_sheet, live_w)
                        st.success("Transfert ESCROW enregistré **dans le fichier**. ✅")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")

    st.divider()
    st.markdown("### 📥 En cours d’alimentation (dossiers **non envoyés**)")
    non_env = live[(live[S_ENVOYE] != True) & (live["ESCROW dispo"] > 0.004)].copy()
    if non_env.empty:
        st.info("Rien en attente côté dossiers non envoyés.")
    else:
        show = non_env[[DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", "Date", HONO, "Payé", ESC_TR, "ESCROW dispo"]].copy()
        for col in [HONO, "Payé", ESC_TR, "ESCROW dispo"]:
            show[col] = show[col].map(_fmt_money_us)
        st.dataframe(show, use_container_width=True)

    st.divider()
    st.markdown("### 🧾 Historique des transferts (journal)")
    has_journal = live[live[ESC_JR].astype(str).str.len() > 0]
    if has_journal.empty:
        st.caption("Aucun journal de transfert pour le moment.")
    else:
        rows = []
        for _, rr in has_journal.iterrows():
            entries = _parse_json_list(rr.get(ESC_JR, ""))
            for e in entries:
                rows.append({
                    DOSSIER_COL: rr.get(DOSSIER_COL, ""),
                    "ID_Client": rr.get("ID_Client", ""),
                    "Nom": rr.get("Nom", ""),
                    "Visa": rr.get("Visa", ""),
                    "Date": rr.get("Date", ""),
                    "Horodatage": e.get("ts", ""),
                    "Montant (US $)": float(e.get("amount", 0.0)),
                    "Note": e.get("note", "")
                })
        jdf = pd.DataFrame(rows)
        if not jdf.empty:
            jdf["Montant (US $)"] = jdf["Montant (US $)"].apply(lambda x: _fmt_money_us(float(x) if pd.notna(x) else 0.0))
        st.dataframe(jdf, use_container_width=True)