
# =========================
# VISA APP — PARTIE 1/5
# =========================

from __future__ import annotations

import json, re, unicodedata
from pathlib import Path
from datetime import date, datetime
from typing import Any

import pandas as pd
import numpy as np
import streamlit as st

# ---------- Constantes colonnes / libellés ----------
DOSSIER_COL = "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"
PAY_JSON = "Paiements"   # JSON [{"date":"YYYY-MM-DD","mode":"CB","amount":123.45}, ...]

# Statuts + dates associées (ordre demandé)
S_ENVOYE,   D_ENVOYE   = "Dossier envoyé",  "Date envoyé"
S_APPROUVE, D_APPROUVE = "Dossier approuvé","Date approuvé"
S_RFE,      D_RFE      = "RFE",             "Date RFE"
S_REFUSE,   D_REFUSE   = "Dossier refusé",  "Date refusé"
S_ANNULE,   D_ANNULE   = "Dossier annulé",  "Date annulé"
STATUS_COLS  = [S_ENVOYE, S_APPROUVE, S_RFE, S_REFUSE, S_ANNULE]
STATUS_DATES = [D_ENVOYE, D_APPROUVE, D_RFE, D_REFUSE, D_ANNULE]

# ESCROW
ESC_TR = "ESCROW transféré (US $)"     # somme des transferts de l'escrow vers compte ordinaire
ESC_JR = "Journal ESCROW"              # JSON [{"ts": "...", "amount": float, "note": ""}, ...]

# Numérotation dossier initiale
DOSSIER_START = 13057

# ---------- Persistance du dernier fichier ----------
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

def set_current_file_from_upload(up_file) -> Path | None:
    """Enregistre l'upload sur disque et le sélectionne comme fichier courant."""
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

# ---------- Formats & conversions ----------
def _safe_str(x) -> str:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip()
    except Exception:
        return ""

def _fmt_money_us(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def _to_num(s: Any) -> pd.Series:
    """Convertit une Series (ou la 1ère col d’un DataFrame) en float propre."""
    if s is None:
        return pd.Series(dtype=float)
    if isinstance(s, pd.DataFrame):
        if s.shape[1] == 0:
            return pd.Series(dtype=float, index=s.index if hasattr(s, "index") else None)
        s = s.iloc[:, 0]
    s = pd.Series(s)
    s = s.astype(str)
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)
    # 1 234,56 -> 1234.56 ; 1,234.56 -> 1234.56
    def _clean_one(v: str) -> float:
        if v == "" or v == "-":
            return 0.0
        if v.count(",")==1 and v.count(".")==0:
            v = v.replace(",", ".")
        if v.count(".")==1 and v.count(",")>=1:
            v = v.replace(",", "")
        try:
            return float(v)
        except Exception:
            return 0.0
    return s.map(_clean_one)

def _to_int(s: Any) -> pd.Series:
    try:
        return pd.to_numeric(pd.Series(s), errors="coerce").fillna(0).astype(int)
    except Exception:
        return pd.Series([0]*len(pd.Series(s)), dtype=int)

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

def read_sheet(path: Path, sheet: str, normalize: bool = False) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()
    if normalize:
        return normalize_dataframe(df)
    return df

def write_sheet_inplace(path: Path, sheet: str, df: pd.DataFrame):
    """Écrit la feuille sheet en conservant les autres feuilles ; crée la feuille si absente."""
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

# ---------- Numérotation / IDs ----------
def ensure_dossier_numbers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if DOSSIER_COL not in df.columns:
        df[DOSSIER_COL] = 0
    nums = _to_int(df[DOSSIER_COL])
    if (nums == 0).all():  # si tout est vide -> initialise
        start = DOSSIER_START
        df[DOSSIER_COL] = [start + i for i in range(len(df))]
        return df
    maxn = int(nums.max()) if len(nums) else (DOSSIER_START - 1)
    for i in range(len(df)):
        if int(nums.iat[i]) <= 0:
            maxn += 1
            df.at[i, DOSSIER_COL] = maxn
    return df

def next_dossier_number(df: pd.DataFrame) -> int:
    if df is None or df.empty or DOSSIER_COL not in df.columns:
        return DOSSIER_START
    nums = _to_int(df[DOSSIER_COL])
    m = int(nums.max()) if len(nums) else (DOSSIER_START - 1)
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
       Numériques -> somme ; sinon -> 1ère valeur non vide."""
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
        # tentative: conversion num puis somme
        try:
            same_num = same.apply(pd.to_numeric, errors="coerce")
            if same_num.notna().any().any():
                out[col] = same_num.sum(axis=1, skipna=True)
                continue
        except Exception:
            pass
        # sinon 1ère non vide
        def _first_non_empty(row):
            for v in row:
                if pd.notna(v) and str(v).strip() != "":
                    return v
            return ""
        out[col] = same.apply(_first_non_empty, axis=1)
    return out

# ---------- Normalisation principale (clients) ----------
def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie champs, calcule Total/Payé/Reste, Date/Mois (MM)."""
    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    # Renommages souples (compat retro)
    rename = {}
    for c in df.columns:
        lc = str(c).lower().strip()
        if lc in ("montant honoraires", "montant honoraires (us $)", "honoraires", "montant"):
            rename[c] = HONO
        elif lc in ("autres frais", "autres frais (us $)", "autres", "autres frais (us$)"):
            rename[c] = AUTRE
        elif lc in ("total", "total (us $)", "total (us$)"):
            rename[c] = TOTAL
        elif lc in ("dossier n", "dossier"):
            rename[c] = DOSSIER_COL
        elif lc in ("reste (us $)", "solde (us $)", "solde", "reste"):
            rename[c] = "Reste"
        elif lc in ("paye (us $)","payé (us $)","paye","payé","paye ($)","payé ($)"):
            rename[c] = "Payé"
        elif lc in ("categorie","catégorie","category"):
            rename[c] = "Catégorie"
        elif lc == "visa":  # on laisse "Visa"
            pass
    if rename:
        df = df.rename(columns=rename)

    # Écrase les colonnes dupliquées après renommage
    df = _collapse_duplicate_columns(df)

    # Colonnes minimales
    for c in [DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa",
              HONO, AUTRE, TOTAL, "Payé", "Reste", PAY_JSON, "Date", "Mois",
              ESC_TR, ESC_JR] + STATUS_COLS + STATUS_DATES:
        if c not in df.columns:
            if c in [HONO, AUTRE, TOTAL, "Payé", "Reste", ESC_TR]:
                df[c] = 0.0
            elif c == PAY_JSON or c == ESC_JR:
                df[c] = ""
            elif c in STATUS_COLS:
                df[c] = False
            elif c in STATUS_DATES:
                df[c] = ""
            else:
                df[c] = ""

    # Numériques
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste", ESC_TR]:
        df[c] = _to_num(df[c])

    # Date & Mois (MM)
    def _to_date(x):
        try:
            if pd.isna(x) or x == "":
                return pd.NaT
            return pd.to_datetime(x).date()
        except Exception:
            return pd.NaT
    df["Date"] = df["Date"].map(_to_date)
    df["Mois"] = df["Date"].apply(lambda d: f"{d.month:02d}" if pd.notna(d) else pd.NA)

    # Calcul Total
    df[TOTAL] = _to_num(df.get(HONO, 0.0)) + _to_num(df.get(AUTRE, 0.0))

    # Payé via JSON si dispo (prend le max entre colonne Payé et somme JSON)
    paid_from_json = []
    for _, r in df.iterrows():
        plist = _parse_json_list(r.get(PAY_JSON, ""))
        paid_from_json.append(_sum_payments(plist))
    paid_from_json = pd.Series(paid_from_json, index=df.index, dtype=float)
    df["Payé"] = pd.Series([max(a, b) for a, b in zip(_to_num(df["Payé"]), paid_from_json)], index=df.index)

    # Reste (>= 0)
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # Statuts & dates (types)
    for b in STATUS_COLS:
        df[b] = df[b].astype(bool)
    for dcol in STATUS_DATES:
        df[dcol] = df[dcol].astype(str)

    # ESCROW types
    df[ESC_TR] = _to_num(df[ESC_TR])

    # Dossier N auto
    df = ensure_dossier_numbers(df)

    return df

# ====== RÉFÉRENTIEL VISA (Catégorie -> SC1 -> SC2 -> SC3 -> SC4 -> Visa) ======
REF_COLS = ["Catégorie","SC1","SC2","SC3","SC4","Visa"]

def _norm_txt(x: str) -> str:
    """Normalise pour comparaison: sans accents, minuscules, fusion espaces, neutralise '/' et '-'. """
    s = "" if x is None else str(x)
    s = s.strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s*[/\-]\s*", " ", s)
    s = re.sub(r"[^a-zA-Z0-9\s]+", " ", s)
    s = " ".join(s.lower().split())
    return s

def _find_col(df: pd.DataFrame, targets: list[str]) -> str | None:
    """Retrouve le nom réel d’une colonne par liste de candidats (tolérant accents/casse)."""
    if df is None or df.empty:
        return None
    m = { _norm_txt(c): str(c) for c in df.columns.astype(str) }
    for t in targets:
        nt = _norm_txt(t)
        if nt in m:
            return m[nt]
    # fallback partiel
    for t in targets:
        nt = _norm_txt(t)
        for k, orig in m.items():
            if nt in k:
                return orig
    return None

def read_visa_reference_tree(path: Path) -> pd.DataFrame:
    """
    Lit la feuille 'Visa' du classeur avec colonnes:
      A: Catégorie, B: Sous-categorie 1, C: Sous-categorie 2, D: Sous-categorie 3,
      E: Sous-categorie 4, F: Visa
    Retourne un DataFrame normalisé REF_COLS.
    """
    try:
        base = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=REF_COLS)

    col_cat = _find_col(base, ["Catégorie","Categorie","Category"])
    col_s1  = _find_col(base, ["Sous-categorie 1","Sous catégorie 1","SC1","Subcategory 1"])
    col_s2  = _find_col(base, ["Sous-categorie 2","Sous catégorie 2","SC2","Subcategory 2"])
    col_s3  = _find_col(base, ["Sous-categorie 3","Sous catégorie 3","SC3","Subcategory 3"])
    col_s4  = _find_col(base, ["Sous-categorie 4","Sous catégorie 4","SC4","Subcategory 4"])
    col_v   = _find_col(base, ["Visa"])

    df = pd.DataFrame()
    df["Catégorie"] = base[col_cat] if col_cat else ""
    df["SC1"] = base[col_s1] if col_s1 else ""
    df["SC2"] = base[col_s2] if col_s2 else ""
    df["SC3"] = base[col_s3] if col_s3 else ""
    df["SC4"] = base[col_s4] if col_s4 else ""
    df["Visa"] = base[col_v] if col_v else ""

    for c in REF_COLS:
        df[c] = df[c].fillna("").astype(str).str.strip()

    # Propagation douce pour stabiliser l'arbre
    df["Catégorie"] = df["Catégorie"].replace("", pd.NA).ffill().fillna("")
    df["SC1"] = df["SC1"].replace("", pd.NA).ffill().fillna("")
    df["SC2"] = df["SC2"].replace("", pd.NA).ffill().fillna("")
    df["SC3"] = df["SC3"].replace("", pd.NA).ffill().fillna("")
    df["SC4"] = df["SC4"].replace("", pd.NA).ffill().fillna("")

    df = df[REF_COLS].drop_duplicates().reset_index(drop=True)
    return df

def cascading_visa_picker_tree(df_ref: pd.DataFrame, key_prefix: str, init: dict | None = None) -> dict:
    """
    Affiche 6 selectbox en cascade. Retourne:
    {"Catégorie":..., "SC1":..., "SC2":..., "SC3":..., "SC4":..., "Visa":...}
    """
    res = {"Catégorie":"", "SC1":"", "SC2":"", "SC3":"", "SC4":"", "Visa":""}
    if df_ref is None or df_ref.empty:
        st.info("Référentiel Visa vide.")
        return res

    # 1) Catégorie
    cats = sorted([v for v in df_ref["Catégorie"].unique() if v])
    idxC = 0
    if init and init.get("Catégorie") in cats: idxC = cats.index(init["Catégorie"])+1
    res["Catégorie"] = st.selectbox("Catégorie", [""]+cats, index=idxC, key=f"{key_prefix}_cat")
    sub = df_ref.copy()
    if res["Catégorie"]:
        sub = sub[sub["Catégorie"] == res["Catégorie"]]

    # 2) SC1
    sc1s = sorted([v for v in sub["SC1"].unique() if v])
    idx1 = 0
    if init and init.get("SC1") in sc1s: idx1 = sc1s.index(init["SC1"])+1
    res["SC1"] = st.selectbox("Sous-catégorie 1", [""]+sc1s, index=idx1, key=f"{key_prefix}_sc1")
    if res["SC1"]:
        sub = sub[sub["SC1"] == res["SC1"]]

    # 3) SC2
    sc2s = sorted([v for v in sub["SC2"].unique() if v])
    idx2 = 0
    if init and init.get("SC2") in sc2s: idx2 = sc2s.index(init["SC2"])+1
    res["SC2"] = st.selectbox("Sous-catégorie 2", [""]+sc2s, index=idx2, key=f"{key_prefix}_sc2")
    if res["SC2"]:
        sub = sub[sub["SC2"] == res["SC2"]]

    # 4) SC3
    sc3s = sorted([v for v in sub["SC3"].unique() if v])
    idx3 = 0
    if init and init.get("SC3") in sc3s: idx3 = sc3s.index(init["SC3"])+1
    res["SC3"] = st.selectbox("Sous-catégorie 3", [""]+sc3s, index=idx3, key=f"{key_prefix}_sc3")
    if res["SC3"]:
        sub = sub[sub["SC3"] == res["SC3"]]

    # 5) SC4
    sc4s = sorted([v for v in sub["SC4"].unique() if v])
    idx4 = 0
    if init and init.get("SC4") in sc4s: idx4 = sc4s.index(init["SC4"])+1
    res["SC4"] = st.selectbox("Sous-catégorie 4", [""]+sc4s, index=idx4, key=f"{key_prefix}_sc4")
    if res["SC4"]:
        sub = sub[sub["SC4"] == res["SC4"]]

    # 6) Visa
    visas = sorted([v for v in sub["Visa"].unique() if v])
    idxV = 0
    if init and init.get("Visa") in visas: idxV = visas.index(init["Visa"])+1
    res["Visa"] = st.selectbox("Visa", [""]+visas, index=idxV, key=f"{key_prefix}_visa")

    if not visas:
        st.caption("Aucun visa à ce niveau. Continue d’affiner ou laisse vide pour voir tous les dossiers correspondants.")
    return res

def visas_autorises_from_tree(df_ref: pd.DataFrame, sel: dict) -> list[str]:
    if df_ref is None or df_ref.empty:
        return []
    sub = df_ref.copy()
    for key in ["Catégorie","SC1","SC2","SC3","SC4"]:
        val = (sel.get(key) or "").strip()
        if val:
            sub = sub[sub[key] == val]
    if (sel.get("Visa") or "").strip():
        sub = sub[sub["Visa"] == sel["Visa"]]
    return sorted([v for v in sub["Visa"].unique() if v])

def filter_by_selection(df: pd.DataFrame, sel: dict, df_ref_tree: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    Filtre les dossiers clients par chemin (Catégorie -> SC1..SC4 -> Visa).
    - Tolère accents/casse/espaces (normalisation).
    - Si df_ref_tree est fourni, on utilise sa whitelist de Visa pour éviter les faux négatifs.
    """
    if df is None or df.empty:
        return df

    f = df.copy()
    col_cat  = _find_col(f, ["Catégorie","Categorie","Category"])
    col_visa = _find_col(f, ["Visa"])

    f["__norm_cat"]  = f[col_cat].astype(str).map(_norm_txt) if col_cat else ""
    f["__norm_visa"] = f[col_visa].astype(str).map(_norm_txt) if col_visa else ""

    want_cat  = _norm_txt(sel.get("Catégorie",""))
    want_visa = _norm_txt(sel.get("Visa",""))

    if want_cat:
        f = f[f["__norm_cat"] == want_cat]

    if want_visa:
        f = f[f["__norm_visa"] == want_visa]
    else:
        if df_ref_tree is not None:
            visas_aut = visas_autorises_from_tree(df_ref_tree, sel)
            if visas_aut:
                visas_norm = {_norm_txt(v) for v in visas_aut}
                f = f[f["__norm_visa"].isin(visas_norm)]

    return f.drop(columns=[c for c in f.columns if c.startswith("__norm_")], errors="ignore")


# =========================
# VISA APP — PARTIE 2/5
# =========================

st.set_page_config(page_title="Visa Manager", layout="wide")

# ---------- Barre latérale : gestion du fichier ----------
st.sidebar.header("📂 Fichier Excel")
uploaded = st.sidebar.file_uploader("Charger/Remplacer fichier (.xlsx)", type=["xlsx"], key="uploader")
if uploaded is not None:
    p = set_current_file_from_upload(uploaded)
    if p:
        st.sidebar.success(f"Fichier chargé: {p.name}")

path_text = st.sidebar.text_input("Ou saisir le chemin d’un fichier existant", value=st.session_state.get("current_path", ""))
colB1, colB2 = st.sidebar.columns(2)
if colB1.button("📄 Ouvrir ce fichier", key="open_manual"):
    p = Path(path_text)
    if p.exists():
        save_workspace_path(p)
        st.sidebar.success(f"Ouvert: {p.name}")
        st.rerun()
    else:
        st.sidebar.error("Chemin invalide.")
if colB2.button("♻️ Reprendre le dernier fichier", key="open_last"):
    p = _load_last_path()
    if p:
        save_workspace_path(p)
        st.sidebar.success(f"Repris: {p.name}")
        st.rerun()
    else:
        st.sidebar.info("Aucun fichier précédemment enregistré.")

current_path = current_file_path()
if current_path is None:
    st.warning("Aucun fichier sélectionné. Charge un .xlsx ou choisis un chemin valide.")
    st.stop()

# ---------- Feuilles disponibles ----------
sheets = list_sheets(current_path)
if not sheets:
    st.error("Impossible de lire le classeur. Assure-toi que le fichier est un .xlsx valide.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.write("**Feuilles détectées :**")
for i, sn in enumerate(sheets):
    st.sidebar.write(f"- {i+1}. {sn}")

# Détection d’une feuille « clients »
client_target_sheet = None
for sn in sheets:
    df_try = read_sheet(current_path, sn, normalize=False)
    if {"Nom", "Visa"}.issubset(set(df_try.columns.astype(str))):
        client_target_sheet = sn
        break

sheet_choice = st.sidebar.selectbox(
    "Feuille à afficher sur le Dashboard :",
    sheets,
    index=max(0, sheets.index(client_target_sheet) if client_target_sheet in sheets else 0),
    key="sheet_choice_select"
)

# ---------- Titre & onglets ----------
st.title("🛂 Visa Manager — US $")

tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs(
    ["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"]
)

# ---------- Référentiel Visa (6 niveaux) ----------
visa_ref_tree = read_visa_reference_tree(current_path)

# ================= DASHBOARD =================
with tab_dash:
    df_raw = read_sheet(current_path, sheet_choice, normalize=False)
    df = read_sheet(current_path, sheet_choice, normalize=True)

    # --- Filtres (clés uniques dash_*) ---
    st.markdown("### 🔎 Filtres (Catégorie → SC1 → SC2 → SC3 → SC4 → Visa)")
    with st.container():
        cTopL, cTopR = st.columns([1,2])
        show_all = cTopL.checkbox("Afficher tous les dossiers", value=False, key=f"dash_show_all_{sheet_choice}")
        cTopL.caption("Sélection hiérarchique")

        with cTopL:
            sel_path_dash = cascading_visa_picker_tree(visa_ref_tree, key_prefix=f"dash_tree_{sheet_choice}")

        cR1, cR2, cR3 = cTopR.columns(3)
        years  = sorted({d.year for d in df["Date"] if pd.notna(d)}) if "Date" in df.columns else []
        months = sorted([m for m in df["Mois"].dropna().unique()]) if "Mois" in df.columns else []
        sel_years  = cR1.multiselect("Année", years, default=[], key=f"dash_years_{sheet_choice}")
        sel_months = cR2.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{sheet_choice}")
        include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True, key=f"dash_na_{sheet_choice}")

    # ---------- Application des filtres ----------
    f = df.copy()
    if not show_all:
        f = filter_by_selection(f, sel_path_dash, df_ref_tree=visa_ref_tree)

    if "Date" in f.columns and sel_years:
        mask = f["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask |= f["Date"].isna()
        f = f[mask]
    if "Mois" in f.columns and sel_months:
        mask = f["Mois"].isin(sel_months)
        if include_na_dates: mask |= f["Mois"].isna()
        f = f[mask]

    hidden = len(df) - len(f)
    if hidden > 0:
        st.caption(f"🔎 {hidden} ligne(s) masquée(s) par les filtres.")

    # KPI compacts
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Total (US $)", _fmt_money_us(float(f.get(TOTAL, pd.Series(dtype=float)).sum())))
    k3.metric("Payé (US $)", _fmt_money_us(float(f.get("Payé", pd.Series(dtype=float)).sum())))
    k4.metric("Solde (US $)", _fmt_money_us(float(f.get("Reste", pd.Series(dtype=float)).sum())))
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    st.subheader("📋 Données (aperçu)")
    cols_show = [c for c in [
        DOSSIER_COL,"ID_Client","Nom","Date","Mois",
        "Catégorie","Visa",
        HONO, AUTRE, TOTAL, "Payé","Reste",
        S_ENVOYE, D_ENVOYE, S_APPROUVE, D_APPROUVE, S_RFE, D_RFE, S_REFUSE, D_REFUSE, S_ANNULE, D_ANNULE
    ] if c in f.columns]
    view = f.copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in view.columns: view[col] = pd.to_numeric(view[col], errors="coerce").fillna(0.0).map(_fmt_money_us)
    if "Date" in view.columns: view["Date"] = view["Date"].astype(str)
    st.dataframe(view[cols_show], use_container_width=True)


# =========================
# VISA APP — PARTIE 3/5
# =========================

with tab_clients:
    st.subheader("👥 Clients — Créer / Modifier / Supprimer")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa)."); st.stop()

    live_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    live = read_sheet(current_path, client_target_sheet, normalize=True)

    # ---------- Sélection client ----------
    left, right = st.columns([1,1])
    with left:
        st.markdown("### 🔎 Rechercher / Sélectionner")
        names = live["Nom"].fillna("").astype(str).tolist() if "Nom" in live.columns else []
        ids   = live["ID_Client"].fillna("").astype(str).tolist() if "ID_Client" in live.columns else []
        display = [f"{n}  —  {i}" for n,i in zip(names,ids)]
        sel_idx = st.selectbox("Client existant", options=list(range(len(display))), format_func=lambda i: display[i] if i is not None and i < len(display) else "", key="cli_sel_idx")
        sel_row = live.iloc[sel_idx] if len(live) and sel_idx is not None else None

    with right:
        st.markdown("### ➕ Nouveau client")
        new_name = st.text_input("Nom", key="new_nom")
        new_date = st.date_input("Date création", value=date.today(), key="new_date")
        # Sélection Visa via référentiel
        st.caption("Choisis le Visa pour ce dossier :")
        sel_path_new = cascading_visa_picker_tree(visa_ref_tree, key_prefix="new_cli")
        new_visa = sel_path_new.get("Visa","")
        new_cat  = sel_path_new.get("Catégorie","")

        new_hono = st.number_input(HONO, min_value=0.0, step=10.0, format="%.2f", key="new_hono")
        new_autr = st.number_input(AUTRE, min_value=0.0, step=10.0, format="%.2f", key="new_autre")

        if st.button("💾 Créer ce client", key="btn_create_client"):
            if not new_name:
                st.warning("Renseigne le **Nom**.")
            elif not new_visa:
                st.warning("Sélectionne un **Visa**.")
            else:
                base = live.copy()
                next_dos = next_dossier_number(base)
                # ID client unique basé sur Nom + Date
                client_id = _make_client_id_from_row({"Nom": new_name, "Date": new_date})
                # éviter collision
                i = 0
                orig = client_id
                while (base["ID_Client"].astype(str) == client_id).any():
                    i += 1
                    client_id = f"{orig}-{i}"

                row = {
                    DOSSIER_COL: next_dos,
                    "ID_Client": client_id,
                    "Nom": new_name,
                    "Date": pd.to_datetime(new_date).date(),
                    "Mois": f"{new_date.month:02d}",
                    "Catégorie": new_cat,
                    "Visa": new_visa,
                    HONO: float(new_hono),
                    AUTRE: float(new_autr),
                    TOTAL: float(new_hono) + float(new_autr),
                    "Payé": 0.0,
                    "Reste": float(new_hono) + float(new_autr),
                    PAY_JSON: "[]",
                    ESC_TR: 0.0,
                    ESC_JR: "[]",
                    S_ENVOYE: False, S_APPROUVE: False, S_RFE: False, S_REFUSE: False, S_ANNULE: False,
                    D_ENVOYE: "", D_APPROUVE: "", D_RFE: "", D_REFUSE: "", D_ANNULE: ""
                }
                base = pd.concat([base, pd.DataFrame([row])], ignore_index=True)
                base = normalize_dataframe(base)
                write_sheet_inplace(current_path, client_target_sheet, base)
                st.success("Client créé et sauvegardé.")
                st.rerun()

    st.markdown("---")
    st.markdown("### ✏️ Modifier / Supprimer / Paiements")
    if sel_row is None or len(live)==0:
        st.info("Sélectionne un client à gauche ou crée un nouveau client.")
        st.stop()

    idx = sel_idx
    ed = live.loc[idx].to_dict()

    c1, c2, c3 = st.columns(3)
    with c1:
        ed_nom  = st.text_input("Nom", value=_safe_str(ed.get("Nom","")), key="ed_nom")
        ed_date = st.date_input("Date création", value=(pd.to_datetime(ed.get("Date")).date() if pd.notna(ed.get("Date")) else date.today()), key="ed_date")
    with c2:
        st.caption("Visa (mise à jour via référentiel)")
        init_path = {"Catégorie": _safe_str(ed.get("Catégorie","")), "SC1":"", "SC2":"", "SC3":"", "SC4":"", "Visa": _safe_str(ed.get("Visa",""))}
        ed_sel    = cascading_visa_picker_tree(visa_ref_tree, key_prefix=f"edit_{idx}", init=init_path)
        ed_cat    = ed_sel.get("Catégorie","")
        ed_visa   = ed_sel.get("Visa","")
    with c3:
        ed_hono = st.number_input(HONO, min_value=0.0, value=float(ed.get(HONO,0.0)), step=10.0, format="%.2f", key=f"ed_hono_{idx}")
        ed_autr = st.number_input(AUTRE, min_value=0.0, value=float(ed.get(AUTRE,0.0)), step=10.0, format="%.2f", key=f"ed_autre_{idx}")

    # Statuts
    s1,s2,s3,s4,s5 = st.columns(5)
    with s1:
        b_env = st.checkbox(S_ENVOYE, value=bool(ed.get(S_ENVOYE, False)), key=f"ed_env_{idx}")
        d_env = st.date_input(D_ENVOYE, value=(pd.to_datetime(ed.get(D_ENVOYE)).date() if _safe_str(ed.get(D_ENVOYE)) else date.today()), key=f"ed_denvoi_{idx}") if b_env else ""
    with s2:
        b_app = st.checkbox(S_APPROUVE, value=bool(ed.get(S_APPROUVE, False)), key=f"ed_app_{idx}")
        d_app = st.date_input(D_APPROUVE, value=(pd.to_datetime(ed.get(D_APPROUVE)).date() if _safe_str(ed.get(D_APPROUVE)) else date.today()), key=f"ed_dappr_{idx}") if b_app else ""
    with s3:
        b_rfe = st.checkbox(S_RFE, value=bool(ed.get(S_RFE, False)), key=f"ed_rfe_{idx}")
        d_rfe = st.date_input(D_RFE, value=(pd.to_datetime(ed.get(D_RFE)).date() if _safe_str(ed.get(D_RFE)) else date.today()), key=f"ed_drfe_{idx}") if b_rfe else ""
    with s4:
        b_ref = st.checkbox(S_REFUSE, value=bool(ed.get(S_REFUSE, False)), key=f"ed_ref_{idx}")
        d_ref = st.date_input(D_REFUSE, value=(pd.to_datetime(ed.get(D_REFUSE)).date() if _safe_str(ed.get(D_REFUSE)) else date.today()), key=f"ed_dref_{idx}") if b_ref else ""
    with s5:
        b_ann = st.checkbox(S_ANNULE, value=bool(ed.get(S_ANNULE, False)), key=f"ed_ann_{idx}")
        d_ann = st.date_input(D_ANNULE, value=(pd.to_datetime(ed.get(D_ANNULE)).date() if _safe_str(ed.get(D_ANNULE)) else date.today()), key=f"ed_dann_{idx}") if b_ann else ""

    st.markdown("#### 💳 Paiements (multi-acomptes)")
    pay_modes = ["CB","Chèque","Cash","Virement","Venmo"]
    pcol1, pcol2, pcol3, pcol4 = st.columns([1,1,1,2])
    with pcol1:
        p_date = st.date_input("Date paiement", value=date.today(), key=f"p_date_{idx}")
    with pcol2:
        p_mode = st.selectbox("Mode", pay_modes, index=0, key=f"p_mode_{idx}")
    with pcol3:
        p_amt  = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"p_amt_{idx}")
    with pcol4:
        if st.button("➕ Ajouter ce paiement", key=f"btn_addpay_{idx}"):
            base = live.copy()
            # Recalcule solde avant ajout
            base_norm = normalize_dataframe(base.copy())
            reste_curr = float(base_norm.loc[idx, "Reste"])
            if float(p_amt) <= 0:
                st.warning("Le montant doit être > 0.")
            elif reste_curr <= 0:
                st.info("Dossier déjà soldé.")
            else:
                row = base.loc[idx].to_dict()
                plist = _parse_json_list(row.get(PAY_JSON,""))
                plist.append({"date": str(p_date), "mode": p_mode, "amount": float(p_amt)})
                row[PAY_JSON] = json.dumps(plist, ensure_ascii=False)
                base.loc[idx] = row
                base = normalize_dataframe(base)
                write_sheet_inplace(current_path, client_target_sheet, base)
                st.success("Paiement ajouté et sauvegardé.")
                st.rerun()

    # Historique paiements
    try:
        plist = _parse_json_list(ed.get(PAY_JSON,""))
    except Exception:
        plist = []
    st.write("**Historique des paiements**")
    if not plist:
        st.caption("Aucun paiement saisi.")
    else:
        hist = pd.DataFrame(plist)
        if "amount" in hist.columns:
            hist = hist.sort_values(by="date", ascending=True)
            hist["amount"] = hist["amount"].map(_fmt_money_us)
        st.dataframe(hist, use_container_width=True)

    # Actions enregistrer / supprimer
    ac1, ac2 = st.columns([1,1])
    if ac1.button("💾 Sauvegarder les modifications", key=f"btn_save_{idx}"):
        base = live.copy()
        row = base.loc[idx].to_dict()
        row["Nom"]  = ed_nom
        row["Date"] = pd.to_datetime(ed_date).date()
        row["Mois"] = f"{ed_date.month:02d}"
        row["Catégorie"] = ed_cat or row.get("Catégorie","")
        row["Visa"]      = ed_visa or row.get("Visa","")
        row[HONO] = float(ed_hono)
        row[AUTRE]= float(ed_autr)
        row[TOTAL]= float(ed_hono) + float(ed_autr)

        # statuts
        row[S_ENVOYE]= bool(b_env); row[D_ENVOYE]= str(d_env) if b_env else ""
        row[S_APPROUVE]= bool(b_app); row[D_APPROUVE]= str(d_app) if b_app else ""
        row[S_RFE]= bool(b_rfe); row[D_RFE]= str(d_rfe) if b_rfe else ""
        row[S_REFUSE]= bool(b_ref); row[D_REFUSE]= str(d_ref) if b_ref else ""
        row[S_ANNULE]= bool(b_ann); row[D_ANNULE]= str(d_ann) if b_ann else ""

        base.loc[idx] = row
        base = normalize_dataframe(base)
        write_sheet_inplace(current_path, client_target_sheet, base)
        st.success("Modifications sauvegardées.")
        st.rerun()

    if ac2.button("🗑️ Supprimer ce client", key=f"btn_del_{idx}"):
        base = live.copy()
        base = base.drop(index=idx).reset_index(drop=True)
        base = normalize_dataframe(base)
        write_sheet_inplace(current_path, client_target_sheet, base)
        st.success("Client supprimé.")
        st.rerun()


# =========================
# VISA APP — PARTIE 4/5
# =========================

try:
    import altair as alt
except Exception:
    alt = None

with tab_analyses:
    st.subheader("📊 Analyses — Volumes, Financier & Comparaisons")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa)."); st.stop()

    dfA_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    dfA = normalize_dataframe(dfA_raw).copy()
    if dfA.empty: st.info("Aucune donnée pour analyser."); st.stop()

    # Filtres (clés uniques anal_*)
    with st.container():
        cL, cR = st.columns([1,2])
        show_all_A = cL.checkbox("Afficher tous les dossiers", value=False, key="anal_show_all")
        cL.caption("Sélection (Catégorie → SC1 → SC2 → SC3 → SC4 → Visa)")
        with cL:
            sel_path_anal = cascading_visa_picker_tree(visa_ref_tree, key_prefix="anal_tree")

        cR1, cR2, cR3 = cR.columns(3)
        yearsA  = sorted({d.year for d in dfA["Date"] if pd.notna(d)}) if "Date" in dfA.columns else []
        monthsA = sorted([m for m in dfA["Mois"].dropna().unique()]) if "Mois" in dfA.columns else []
        sel_years  = cR1.multiselect("Année", yearsA, default=[], key="anal_years")
        sel_months = cR2.multiselect("Mois (MM)", monthsA, default=[], key="anal_months")
        include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True, key="anal_na")

    # ---------- Application des filtres ----------
    fA = dfA.copy()
    if not show_all_A:
        fA = filter_by_selection(fA, sel_path_anal, df_ref_tree=visa_ref_tree)

    if "Date" in fA.columns and sel_years:
        mask_year = fA["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask_year |= fA["Date"].isna()
        fA = fA[mask_year]
    if "Mois" in fA.columns and sel_months:
        mask_month = fA["Mois"].isin(sel_months)
        if include_na_dates: mask_month |= fA["Mois"].isna()
        fA = fA[mask_month]

    # Enrichissements
    fA["Année"] = fA["Date"].apply(lambda x: x.year if pd.notna(x) else pd.NA)
    fA["MoisNum"] = fA["Date"].apply(lambda x: int(x.month) if pd.notna(x) else pd.NA)
    fA["Periode"] = fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")

    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in fA.columns: fA[col] = pd.to_numeric(fA[col], errors="coerce").fillna(0.0)

    # KPI
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(fA)}")
    k2.metric("Total (US $)", _fmt_money_us(float(fA.get(TOTAL, pd.Series(dtype=float)).sum())) )
    k3.metric("Payé (US $)", _fmt_money_us(float(fA.get("Payé", pd.Series(dtype=float)).sum())) )
    k4.metric("Solde (US $)", _fmt_money_us(float(fA.get("Reste", pd.Series(dtype=float)).sum())) )
    st.markdown('</div>', unsafe_allow_html=True)

    # Volumes créations
    st.markdown("### 📈 Volumes de créations")
    vol_crees = fA.groupby("Periode").size().reset_index(name="Créés")
    df_vol = vol_crees.rename(columns={"Créés":"Volume"}).assign(Indic="Créés")

    if alt is not None and not df_vol.empty:
        try:
            st.altair_chart(
                alt.Chart(df_vol).mark_line(point=True).encode(
                    x=alt.X("Periode:N", sort=None, title="Période"),
                    y=alt.Y("Volume:Q"),
                    color=alt.Color("Indic:N", legend=alt.Legend(title="")),
                    tooltip=["Periode","Indic","Volume"]
                ).properties(height=260), use_container_width=True
            )
        except Exception:
            st.dataframe(df_vol, use_container_width=True)
    else:
        st.dataframe(df_vol, use_container_width=True)

    st.divider()

    # Comparaisons YoY & MoM
    st.markdown("## 🔁 Comparaisons (YoY & MoM)")

    by_year = fA.dropna(subset=["Année"]).groupby("Année").agg(
        Dossiers=("Nom","count"),
        Honoraires=(HONO,"sum"),
        Autres=(AUTRE,"sum"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index().sort_values("Année")

    c1, c2 = st.columns(2)
    if alt is not None and not by_year.empty:
        try:
            c1.altair_chart(
                alt.Chart(by_year.melt("Année", ["Dossiers"])).mark_bar().encode(
                    x=alt.X("Année:N"), y=alt.Y("value:Q", title="Volume"),
                    color=alt.Color("variable:N", legend=None),
                    tooltip=["Année","value"]
                ).properties(title="Nombre de dossiers", height=260), use_container_width=True
            )
        except Exception:
            c1.dataframe(by_year[["Année","Dossiers"]], use_container_width=True)
        try:
            metric_vars = ["Honoraires","Autres","Total","Payé","Reste"]
            yo = by_year.melt("Année", metric_vars, var_name="Indicateur", value_name="Montant")
            c2.altair_chart(
                alt.Chart(yo).mark_bar().encode(
                    x=alt.X("Année:N"),
                    y=alt.Y("Montant:Q"),
                    color=alt.Color("Indicateur:N"),
                    tooltip=["Année","Indicateur", alt.Tooltip("Montant:Q", format="$.2f")]
                ).properties(title="Montants par année", height=260), use_container_width=True
            )
        except Exception:
            c2.dataframe(by_year.drop(columns=["Dossiers"]), use_container_width=True)
    else:
        c1.dataframe(by_year[["Année","Dossiers"]], use_container_width=True)
        c2.dataframe(by_year.drop(columns=["Dossiers"]), use_container_width=True)

    st.markdown("### 📅 Mois (1..12) — Année sur année")
    by_year_month = fA.dropna(subset=["Année","MoisNum"]).groupby(["Année","MoisNum"]).agg(
        Dossiers=("Nom","count"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index()

    c3, c4 = st.columns(2)
    if alt is not None and not by_year_month.empty:
        try:
            c3.altair_chart(
                alt.Chart(by_year_month).mark_line(point=True).encode(
                    x=alt.X("MoisNum:O", title="Mois"),
                    y=alt.Y("Dossiers:Q"),
                    color=alt.Color("Année:N"),
                    tooltip=["Année","MoisNum","Dossiers"]
                ).properties(title="Dossiers par mois (YoY)", height=260), use_container_width=True
            )
        except Exception:
            c3.dataframe(by_year_month.pivot(index="MoisNum", columns="Année", values="Dossiers"), use_container_width=True)
        try:
            c4.altair_chart(
                alt.Chart(by_year_month.melt(["Année","MoisNum"], ["Total","Payé","Reste"],
                                             var_name="Indicateur", value_name="Montant")
                ).mark_line(point=True).encode(
                    x=alt.X("MoisNum:O", title="Mois"),
                    y=alt.Y("Montant:Q"),
                    color=alt.Color("Année:N"),
                    tooltip=["Année","MoisNum","Indicateur", alt.Tooltip("Montant:Q", format="$.2f")]
                ).properties(title="Montants par mois (YoY)", height=260),
                use_container_width=True
            )
        except Exception:
            c4.dataframe(by_year_month.pivot_table(index="MoisNum", columns="Année", values="Total"), use_container_width=True)
    else:
        c3.dataframe(by_year_month.pivot(index="MoisNum", columns="Année", values="Dossiers"), use_container_width=True)
        c4.dataframe(by_year_month.pivot_table(index="MoisNum", columns="Année", values="Total"), use_container_width=True)

    st.markdown("### 🛂 Par type de visa — Année sur année")
    topN = st.slider("Top N visas (par Total)", 3, 20, 10, 1, key="cmp_topn")
    metric_cmp = st.selectbox("Indicateur", ["Dossiers","Total","Payé","Reste","Honoraires","Autres"], index=1, key="cmp_metric")

    by_year_visa = fA.dropna(subset=["Année"]).groupby(["Année","Visa"]).agg(
        Dossiers=("Nom","count"),
        Honoraires=(HONO,"sum"),
        Autres=(AUTRE,"sum"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index()

    top_visas = (by_year_visa.groupby("Visa")["Total"].sum()
                 .sort_values(ascending=False).head(topN).index.tolist())
    by_year_visa_top = by_year_visa[by_year_visa["Visa"].isin(top_visas)].copy()

    if alt is not None and not by_year_visa_top.empty:
        try:
            st.altair_chart(
                alt.Chart(by_year_visa_top).mark_bar().encode(
                    x=alt.X("Visa:N", sort=top_visas),
                    y=alt.Y(f"{metric_cmp}:Q"),
                    color=alt.Color("Année:N"),
                    tooltip=["Visa","Année", alt.Tooltip(f"{metric_cmp}:Q", format="$.2f" if metric_cmp!="Dossiers" else "")],
                ).properties(height=300), use_container_width=True
            )
        except Exception:
            st.dataframe(by_year_visa_top.pivot_table(index="Visa", columns="Année", values=metric_cmp, aggfunc="sum"),
                         use_container_width=True)
    else:
        st.dataframe(by_year_visa_top.pivot_table(index="Visa", columns="Année", values=metric_cmp, aggfunc="sum"),
                     use_container_width=True)

    st.divider()
    st.markdown("### 🔎 Détails (clients)")
    details_cols = [c for c in ["Periode",DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Date",
                                HONO, AUTRE, TOTAL, "Payé","Reste","Année","MoisNum"] if c in fA.columns]
    details = fA.copy()
    details["Periode"] = details["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in details.columns: details[col] = details[col].apply(lambda x: _fmt_money_us(x) if pd.notna(x) else "")
    st.dataframe(details[details_cols].sort_values(["Année","MoisNum","Catégorie","Nom"]), use_container_width=True)


# =========================
# VISA APP — PARTIE 5/5
# =========================

with tab_escrow:
    st.subheader("🏦 ESCROW — suivi & transferts")

    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide."); st.stop()

    dfE = read_sheet(current_path, client_target_sheet, normalize=True)
    if dfE.empty:
        st.info("Aucun dossier."); st.stop()

    # Calcul dispo ESCROW par dossier = min(Payé, Honoraires) - déjà Transféré
    dfE["Dispo ESCROW"] = (dfE["Payé"].clip(upper=dfE[HONO]) - dfE[ESC_TR]).clip(lower=0.0)

    # Alerte : dossiers "envoyés" avec dispo > 0 => à réclamer/transferer
    to_claim = dfE[(dfE[S_ENVOYE]==True) & (dfE["Dispo ESCROW"]>0.0)]
    if len(to_claim):
        st.warning(f"⚠️ {len(to_claim)} dossier(s) envoyé(s) ont de l’ESCROW à transférer.")
        st.dataframe(to_claim[[DOSSIER_COL,"ID_Client","Nom","Visa",HONO,"Payé","Dispo ESCROW"]], use_container_width=True)

    st.divider()
    st.markdown("### 🔁 Marquer un transfert d’ESCROW → Compte ordinaire")
    for i, r in dfE.iterrows():
        dispo = float(r["Dispo ESCROW"])
        if dispo <= 0:
            continue
        with st.expander(f"{r[DOSSIER_COL]} — {r['Nom']} — Visa {r['Visa']} — Dispo: {_fmt_money_us(dispo)}", expanded=False):
            amt = st.number_input("Montant à marquer comme transféré (US $)",
                                  min_value=0.0, value=float(dispo),
                                  step=10.0, format="%.2f", key=f"esc_amt_{i}")
            note = st.text_input("Note (optionnelle)", key=f"esc_note_{i}")
            if st.button("💾 Enregistrer le transfert", key=f"esc_save_{i}"):
                base = read_sheet(current_path, client_target_sheet, normalize=False)
                # mettre à jour ligne i
                row = base.loc[i].to_dict()
                journal = _parse_json_list(row.get(ESC_JR,""))
                journal.append({"ts": datetime.now().isoformat(timespec="seconds"), "amount": float(amt), "note": _safe_str(note)})
                row[ESC_JR] = json.dumps(journal, ensure_ascii=False)
                row[ESC_TR] = float(_safe_str(row.get(ESC_TR,0.0)) or 0.0) + float(amt)
                base.loc[i] = row
                base = normalize_dataframe(base)
                write_sheet_inplace(current_path, client_target_sheet, base)
                st.success("Transfert enregistré.")
                st.rerun()

    st.divider()
    st.markdown("### 📒 Journal ESCROW (tous dossiers)")
    rows = []
    for i, r in dfE.iterrows():
        jr = _parse_json_list(r.get(ESC_JR,""))
        for ent in jr:
            rows.append({
                "Horodatage": ent.get("ts",""),
                DOSSIER_COL: r.get(DOSSIER_COL,""),
                "ID_Client": r.get("ID_Client",""),
                "Nom": r.get("Nom",""),
                "Visa": r.get("Visa",""),
                "Montant": float(ent.get("amount",0.0)),
                "Note": ent.get("note","")
            })
    if rows:
        jdf = pd.DataFrame(rows).sort_values("Horodatage")
        jdf["Montant"] = jdf["Montant"].map(_fmt_money_us)
        st.dataframe(jdf, use_container_width=True)
    else:
        st.caption("Pas encore de transferts enregistrés.")