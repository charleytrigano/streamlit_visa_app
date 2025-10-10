# ============================================
# VISA APP — PARTIE 1/5
# Imports, constantes, helpers, lecture fichiers, onglets
# ============================================

from __future__ import annotations

import json
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Tuple, List, Dict, Any

import numpy as np
import pandas as pd
import streamlit as st

# ---------- Constantes colonnes ----------
HONO   = "Montant honoraires (US $)"
AUTRE  = "Autres frais (US $)"
TOTAL  = "Total (US $)"
PAY_JSON = "Paiements"
DOSSIER_COL = "Dossier N"

REF_LEVELS = ["Catégorie"] + [f"Sous-categories {i}" for i in range(1,9)]
DOSSIER_START = 13057

# ---------- Utils génériques ----------
def _fmt_money_us(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _safe_str(x) -> str:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip()
    except Exception:
        return ""

def _norm_txt(x: str) -> str:
    s = _safe_str(x)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s*[/\-]\s*", " ", s)
    s = re.sub(r"[^a-zA-Z0-9\s]+", " ", s)
    return " ".join(s.lower().split())

def _to_num_series(s_like) -> pd.Series:
    """Convertit une colonne (Series/DataFrame/list) en Series float robuste."""
    if isinstance(s_like, pd.DataFrame):
        if s_like.shape[1] == 0:
            return pd.Series([], dtype=float)
        s = s_like.iloc[:, 0]
    else:
        s = pd.Series(s_like)
    s = s.astype(str).str.replace(r"[^\d,.\-]", "", regex=True)
    def _one(x: str) -> float:
        if x == "" or x == "-":
            return 0.0
        if x.count(",") == 1 and x.count(".") == 0:
            x = x.replace(",", ".")
        if x.count(".") == 1 and x.count(",") >= 1:
            x = x.replace(",", "")
        try:
            return float(x)
        except Exception:
            return 0.0
    return s.map(_one)

def _collapse_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Fusionne les colonnes dupliquées (somme si numérique, sinon première non vide)."""
    cols = df.columns.astype(str)
    if not cols.duplicated().any():
        return df
    out = pd.DataFrame(index=df.index)
    for col in pd.unique(cols):
        same = df.loc[:, cols == col]
        if same.shape[1] == 1:
            out[col] = same.iloc[:, 0]
            continue
        # Essai somme numérique
        try:
            same_num = same.apply(pd.to_numeric, errors="coerce")
            if same_num.notna().any().any():
                out[col] = same_num.sum(axis=1, skipna=True)
                continue
        except Exception:
            pass
        # Sinon, première non vide
        def _first_non_empty(row):
            for v in row:
                if pd.notna(v) and str(v).strip() != "":
                    return v
            return ""
        out[col] = same.apply(_first_non_empty, axis=1)
    return out

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    """Retourne une Series numérique pour 'col', même si dupliquée ou absente."""
    if col not in df.columns:
        return pd.Series([], dtype=float)
    s = df[col]
    if isinstance(s, pd.DataFrame):  # colonnes dupliquées
        s = s.iloc[:, 0]
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _visa_code_only(v: str) -> str:
    s = _safe_str(v)
    if not s:
        return ""
    parts = s.split()
    if len(parts) >= 2 and parts[-1].upper() in {"COS", "EOS"}:
        return " ".join(parts[:-1]).strip()
    return s.strip()

def next_dossier_number(df: pd.DataFrame) -> int:
    if df is None or df.empty or DOSSIER_COL not in df.columns:
        return DOSSIER_START
    try:
        nums = pd.to_numeric(df[DOSSIER_COL], errors="coerce")
        m = int(nums.max()) if nums.notna().any() else DOSSIER_START - 1
    except Exception:
        m = DOSSIER_START - 1
    return max(m, DOSSIER_START - 1) + 1

def _make_client_id_from_row(row: dict) -> str:
    nom = _safe_str(row.get("Nom"))
    d = row.get("Date")
    try:
        d = pd.to_datetime(d).date() if pd.notna(d) else date.today()
    except Exception:
        d = date.today()
    base = f"{nom}-{d.strftime('%Y%m%d')}"
    base = re.sub(r"[^A-Za-z0-9\-]+", "", base.replace(" ", "-")).lower()
    return base

# ---------- Normalisation CLIENTS ----------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = _collapse_duplicate_columns(df.copy())

    # mapping souple
    ren = {}
    for c in df.columns:
        lc = _norm_txt(c)
        if "montant honoraires" in lc or lc == "honoraires":
            ren[c] = HONO
        elif "autres frais" in lc or lc == "autres":
            ren[c] = AUTRE
        elif lc.startswith("total"):
            ren[c] = TOTAL
        elif lc in {"reste", "solde"}:
            ren[c] = "Reste"
        elif "paye" in lc or "payé" in lc:
            ren[c] = "Payé"
        elif "categorie" in lc:
            ren[c] = "Catégorie"
        elif lc == "visa":
            ren[c] = "Visa"
        elif lc in {"dossier n", "dossier"}:
            ren[c] = DOSSIER_COL
        elif lc == "id_client":
            ren[c] = "ID_Client"
        elif lc == "nom":
            ren[c] = "Nom"
        elif lc == "date":
            ren[c] = "Date"
        elif lc == "mois":
            ren[c] = "Mois"
        elif lc == "paiements":
            ren[c] = PAY_JSON
    if ren:
        df = df.rename(columns=ren)

    # colonnes requises
    required = [
        DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois", "Catégorie", "Visa",
        HONO, AUTRE, TOTAL, "Payé", "Reste", PAY_JSON,
        "ESCROW transféré (US $)", "Journal ESCROW",
        "Dossier envoyé", "Date envoyé", "Dossier approuvé", "Date approuvé",
        "RFE", "Date RFE", "Dossier refusé", "Date refusé", "Dossier annulé", "Date annulé",
    ]
    for c in required:
        if c not in df.columns:
            if c in [HONO, AUTRE, TOTAL, "Payé", "Reste", "ESCROW transféré (US $)"]:
                df[c] = 0.0
            elif c in [PAY_JSON, "Journal ESCROW", "ID_Client", "Nom", "Catégorie", "Visa", "Mois", "Date"]:
                df[c] = ""
            elif c in ["Dossier envoyé", "Dossier approuvé", "RFE", "Dossier refusé", "Dossier annulé"]:
                df[c] = False
            else:
                df[c] = ""

    # Nettoyage Visa/Catégorie
    df["Visa"] = df["Visa"].astype(str).map(_visa_code_only)
    df["Catégorie"] = df["Catégorie"].replace("", pd.NA).fillna(df["Visa"]).astype(str)

    # Numériques
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste", "ESCROW transféré (US $)"]:
        df[c] = _to_num_series(df[c])

    # Dates + Mois
    def _to_date(x):
        try:
            if x == "" or pd.isna(x):
                return pd.NaT
            return pd.to_datetime(x).date()
        except Exception:
            return pd.NaT
    df["Date"] = df["Date"].map(_to_date)
    df["Mois"] = df["Date"].apply(lambda d: f"{d.month:02d}" if pd.notna(d) else pd.NA)

    # Totaux
    df[TOTAL] = df[HONO] + df[AUTRE]
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # N° de dossier
    if DOSSIER_COL in df.columns:
        nums = pd.to_numeric(df[DOSSIER_COL], errors="coerce")
        maxn = int(nums.max()) if nums.notna().any() else DOSSIER_START - 1
        for i in range(len(df)):
            if pd.isna(nums.iat[i]) or int(nums.iat[i]) <= 0:
                maxn += 1
                df.at[i, DOSSIER_COL] = maxn
        try:
            df[DOSSIER_COL] = df[DOSSIER_COL].astype(int)
        except Exception:
            pass

    # ID client si manquant
    for i, r in df.iterrows():
        if not _safe_str(r.get("ID_Client", "")):
            base = _make_client_id_from_row(r.to_dict())
            cand = base
            j = 0
            while (df["ID_Client"].astype(str) == cand).any():
                j += 1
                cand = f"{base}-{j}"
            df.at[i, "ID_Client"] = cand

    # Tri
    try:
        df = df.sort_values(["Date", "Nom"], na_position="last").reset_index(drop=True)
    except Exception:
        pass
    return df

# ---------- Normalisation VISA (référentiel) ----------
def _ensure_visa_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et structure le fichier Visa (Catégorie + Sous-categories 1..8 + Actif)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=REF_LEVELS + ["Actif"])
    out = df.copy()

    # normalisation des noms
    norm = {re.sub(r"[^a-z0-9]+", "", str(c).lower()): str(c) for c in out.columns}
    def find_col(*cands):
        for cand in cands:
            key = re.sub(r"[^a-z0-9]+", "", cand.lower())
            if key in norm:
                return norm[key]
        for cand in cands:
            key = re.sub(r"[^a-z0-9]+", "", cand.lower())
            for k, orig in norm.items():
                if key in k:
                    return orig
        return None

    cat = find_col("Catégorie", "Categorie", "Category")
    out = out.rename(columns={cat: "Catégorie"}) if cat else out.assign(**{"Catégorie": ""})
    for i in range(1, 9):
        sc = find_col(f"Sous-categories {i}", f"Sous categorie {i}", f"SC{i}")
        if sc:
            out = out.rename(columns={sc: f"Sous-categories {i}"})
        else:
            out[f"Sous-categories {i}"] = ""
    act = find_col("Actif", "Active", "Inclure", "Afficher")
    out = out.rename(columns={act: "Actif"}) if act else out.assign(**{"Actif": 1})

    out = out.reindex(columns=REF_LEVELS + ["Actif"])
    for c in REF_LEVELS + ["Actif"]:
        out[c] = out[c].fillna("").astype(str).str.strip()
    out["Catégorie"] = out["Catégorie"].replace("", pd.NA).ffill().fillna("")
    try:
        out["Actif_num"] = pd.to_numeric(out["Actif"], errors="coerce").fillna(0).astype(int)
        out = out[out["Actif_num"] == 1].drop(columns=["Actif_num"])
    except Exception:
        pass
    mask = out[REF_LEVELS].apply(lambda r: "".join(r.values), axis=1) != ""
    return out[mask].reset_index(drop=True)

def _slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(s).lower()).strip("_")

def _multi_bool_inputs(options: List[str], label: str, keyprefix: str, as_toggle: bool=False) -> List[str]:
    """Affiche une sélection multi-case (checkbox ou toggle)"""
    options = [o for o in options if str(o).strip() != ""]
    if not options:
        st.caption(f"Aucune option pour **{label}**.")
        return []
    with st.expander(label, expanded=False):
        c1, c2 = st.columns(2)
        all_on  = c1.toggle("Tout sélectionner", value=False, key=f"{keyprefix}_all")
        none_on = c2.toggle("Tout désélectionner", value=False, key=f"{keyprefix}_none")
        selected = []
        cols = st.columns(3 if len(options) > 6 else 2)
        for i, opt in enumerate(sorted(options)):
            k = f"{keyprefix}_{i}"
            if all_on:
                st.session_state[k] = True
            if none_on:
                st.session_state[k] = False
            with cols[i % len(cols)]:
                val = st.toggle(opt, value=st.session_state.get(k, False), key=k) if as_toggle \
                      else st.checkbox(opt, value=st.session_state.get(k, False), key=k)
                if val:
                    selected.append(opt)
    return selected

def build_checkbox_filters_grouped(df_ref_in: pd.DataFrame, keyprefix: str, as_toggle: bool=False) -> dict:
    """Construit l’arborescence dynamique de filtres (Catégorie → SC1..SC8), renvoie une whitelist par Catégorie."""
    df_ref = _ensure_visa_columns(df_ref_in)
    res = {"Catégorie": [], "SC_map": {}, "__whitelist_visa__": []}
    if df_ref.empty:
        st.info("Référentiel Visa vide ou invalide.")
        return res

    cats = sorted([v for v in df_ref["Catégorie"].unique() if str(v).strip() != ""])
    sel_cats = _multi_bool_inputs(cats, "Catégories", f"{keyprefix}_cat", as_toggle=as_toggle)
    res["Catégorie"] = sel_cats

    whitelist = set()
    for cat in sel_cats:
        sub = df_ref[df_ref["Catégorie"] == cat].copy()
        res["SC_map"][cat] = {}
        st.markdown(f"#### 🧭 {cat}")
        for i in range(1, 9):
            col = f"Sous-categories {i}"
            options = sorted([v for v in sub[col].unique() if str(v).strip() != ""])
            picked = _multi_bool_inputs(options, f"{cat} — {col}", f"{keyprefix}_{_slug(cat)}_sc{i}", as_toggle=as_toggle)
            res["SC_map"][cat][col] = picked
            if picked:
                sub = sub[sub[col].isin(picked)]
        # Dans ta logique, le « code de base » à filtrer côté Clients = Catégorie
        whitelist.add(cat)

    res["__whitelist_visa__"] = sorted(whitelist)
    return res

def filter_clients_by_ref(df_clients: pd.DataFrame, sel: dict) -> pd.DataFrame:
    """Applique le filtre de sélection (whitelist Catégorie) au tableau des clients."""
    if df_clients is None or df_clients.empty:
        return df_clients
    f = df_clients.copy()
    wl = set(sel.get("__whitelist_visa__", []))
    if wl and "Catégorie" in f.columns:
        f = f[f["Catégorie"].astype(str).isin(wl)]
    return f

# ---------- I/O Excel ----------
def list_sheets(xlsx_path: Path) -> List[str]:
    try:
        return pd.ExcelFile(xlsx_path).sheet_names
    except Exception:
        return []

def read_sheet(xlsx_path: Path, sheet_name: str, normalize: bool=True) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    return normalize_clients(df) if normalize else df

def write_sheet_inplace(xlsx_path: Path, sheet_name: str, df: pd.DataFrame) -> None:
    """Écrit df dans sheet_name en conservant les autres feuilles."""
    try:
        xls = pd.ExcelFile(xlsx_path)
        sheets = xls.sheet_names
    except Exception:
        sheets = []
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as w:
        wrote = False
        for sn in sheets:
            if sn == sheet_name:
                (df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)).to_excel(w, sheet_name=sn, index=False)
                wrote = True
            else:
                pd.read_excel(xlsx_path, sheet_name=sn).to_excel(w, sheet_name=sn, index=False)
        if not wrote:
            df.to_excel(w, sheet_name=sheet_name, index=False)

# ---------- Persistance des derniers chemins ----------
STATE_FILE = Path(".visa_app_state.json")

def _save_last_paths(clients: Path | None = None, visa: Path | None = None) -> None:
    data = {}
    if STATE_FILE.exists():
        try:
            data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    if clients is not None:
        data["clients_path"] = str(clients)
    if visa is not None:
        data["visa_path"] = str(visa)
    STATE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def _load_last_paths() -> tuple[Path | None, Path | None]:
    if not STATE_FILE.exists():
        return None, None
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        cp = Path(data.get("clients_path", "")) if data.get("clients_path") else None
        vp = Path(data.get("visa_path", "")) if data.get("visa_path") else None
        if cp is not None and not cp.exists():
            cp = None
        if vp is not None and not vp.exists():
            vp = None
        return cp, vp
    except Exception:
        return None, None

# ---------- UI : Fichiers (barre latérale) ----------
st.sidebar.header("📁 Fichiers")

last_clients, last_visa = _load_last_paths()

up_clients = st.sidebar.file_uploader("Classeur **Clients** (.xlsx)", type=["xlsx"], key="up_clients")
up_visa    = st.sidebar.file_uploader("Référentiel **Visa** (.xlsx)", type=["xlsx"], key="up_visa")

clients_text = st.sidebar.text_input("Chemin Clients", value=str(last_clients) if last_clients else "")
visa_text    = st.sidebar.text_input("Chemin Visa", value=str(last_visa) if last_visa else "")

clients_path: Path | None = None
visa_path: Path | None = None

# Sauvegarde des uploads localement + mémorisation
if up_clients is not None:
    p = Path(up_clients.name).resolve()
    p.write_bytes(up_clients.getvalue())
    clients_path = p
    _save_last_paths(clients=p)

if up_visa is not None:
    p = Path(up_visa.name).resolve()
    p.write_bytes(up_visa.getvalue())
    visa_path = p
    _save_last_paths(visa=p)

if clients_path is None and clients_text:
    p = Path(clients_text)
    if p.exists():
        clients_path = p

if visa_path is None and visa_text:
    p = Path(visa_text)
    if p.exists():
        visa_path = p

# Lecture différée sécurisée
if clients_path is None or not clients_path.exists():
    st.warning("Charge/indique d’abord le **classeur Clients** (.xlsx).")
    st.stop()

# Heuristique de feuille Clients
sheets = list_sheets(clients_path)
if "Clients" in sheets:
    sheet_choice = "Clients"
elif "Clients_normalises" in sheets:
    sheet_choice = "Clients_normalises"
else:
    sheet_choice = sheets[0] if sheets else None

if sheet_choice is None:
    st.error("Aucune feuille exploitable dans le classeur Clients.")
    st.stop()

df_clients = read_sheet(clients_path, sheet_choice, normalize=True)

# Référentiel Visa : on accepte un fichier séparé ou une feuille dans le classeur Clients
if visa_path and visa_path.exists():
    try:
        xlv = pd.ExcelFile(visa_path)
        visa_sn = "Visa" if "Visa" in xlv.sheet_names else ("Visa_normalise" if "Visa_normalise" in xlv.sheet_names else xlv.sheet_names[0])
        df_visa = pd.read_excel(visa_path, sheet_name=visa_sn)
    except Exception:
        df_visa = pd.DataFrame()
else:
    # Cherche dans le classeur des clients
    if "Visa" in sheets:
        df_visa = pd.read_excel(clients_path, sheet_name="Visa")
    elif "Visa_normalise" in sheets:
        df_visa = pd.read_excel(clients_path, sheet_name="Visa_normalise")
    else:
        df_visa = pd.DataFrame()

df_visa = _ensure_visa_columns(df_visa)

# ---------- Création des onglets ----------
tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"])


# ============================================
# VISA APP — PARTIE 2/5
# Dashboard : filtres hiérarchiques + KPI + tableau
# ============================================

with tab_dash:
    st.subheader("📊 Dashboard")

    # --- 1) Filtres VISA (hiérarchiques) ---
    df_visa_safe = _ensure_visa_columns(df_visa)
    if df_visa_safe.empty:
        st.warning("⚠️ Le référentiel Visa est vide ou mal formé. Charge un fichier Visa valide.")
        sel = {"__whitelist_visa__": [], "Catégorie": []}
        f = df_clients.copy()
    else:
        sel = build_checkbox_filters_grouped(
            df_visa_safe,
            keyprefix=f"flt_dash_{sheet_choice}",
            as_toggle=False,  # passer à True pour des boutons bascule
        )
        f = filter_clients_by_ref(df_clients, sel)

    # --- 2) Filtres additionnels (Année / Mois / Solde / Recherche) ---
    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])

    # Calcul des champs dérivés nécessaires aux filtres
    f["_Année_"] = f["Date"].apply(lambda x: x.year if pd.notna(x) else pd.NA)
    f["_MoisNum_"] = f["Date"].apply(lambda x: int(x.month) if pd.notna(x) else pd.NA)

    yearsA = sorted([int(y) for y in f["_Année_"].dropna().unique()]) if not f.empty else []
    monthsA = [f"{m:02d}" for m in sorted([int(m) for m in f["_MoisNum_"].dropna().unique()])] if not f.empty else []

    with c1:
        sel_years = st.multiselect("Année", yearsA, default=[], key=f"yr_{sheet_choice}")
    with c2:
        sel_months = st.multiselect("Mois (MM)", monthsA, default=[], key=f"mo_{sheet_choice}")
    with c3:
        solde_mode = st.selectbox(
            "Solde",
            ["Tous", "Soldé (Reste = 0)", "Non soldé (Reste > 0)"],
            index=0,
            key=f"solde_{sheet_choice}",
        )
    with c4:
        q = st.text_input("Recherche (nom, ID, visa…)", "", key=f"q_{sheet_choice}")

    # Application des filtres additionnels
    ff = f.copy()
    if sel_years:
        ff = ff[ff["_Année_"].isin(sel_years)]
    if sel_months:
        ff = ff[ff["Mois"].astype(str).isin(sel_months)]
    if solde_mode == "Soldé (Reste = 0)":
        ff = ff[_safe_num_series(ff, "Reste") <= 0.0000001]
    elif solde_mode == "Non soldé (Reste > 0)":
        ff = ff[_safe_num_series(ff, "Reste") > 0.0000001]
    if q:
        qn = q.lower().strip()
        def _row_match(r):
            hay = " ".join([
                _safe_str(r.get("Nom","")),
                _safe_str(r.get("ID_Client","")),
                _safe_str(r.get("Catégorie","")),
                _safe_str(r.get("Visa","")),
                str(r.get(DOSSIER_COL,"")),
            ]).lower()
            return qn in hay
        ff = ff[ff.apply(_row_match, axis=1)]

    # --- 3) Bouton réinitialiser les filtres de la section ---
    if st.button("🔄 Réinitialiser les filtres", key=f"reset_dash_{sheet_choice}"):
        # On nettoie uniquement les clés de ce dashboard
        for k in list(st.session_state.keys()):
            if k.startswith(f"flt_dash_{sheet_choice}") or \
               k in {f"yr_{sheet_choice}", f"mo_{sheet_choice}", f"solde_{sheet_choice}", f"q_{sheet_choice}"}:
                del st.session_state[k]
        st.rerun()

    # --- 4) KPI (sécurisés contre colonnes dupliquées/vides) ---
    st.markdown("""
    <style>
    .small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}
    .small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(ff)}")
    k2.metric("Honoraires", _fmt_money_us(_safe_num_series(ff, HONO).sum()))
    k3.metric("Payé",      _fmt_money_us(_safe_num_series(ff, "Payé").sum()))
    k4.metric("Solde",     _fmt_money_us(_safe_num_series(ff, "Reste").sum()))
    st.markdown('</div>', unsafe_allow_html=True)

    # --- 5) Tableau (montants formatés) ---
view = ff.copy()
for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
    if c in view.columns:
        view[c] = _safe_num_series(view, c).map(_fmt_money_us)
if "Date" in view.columns:
    view["Date"] = view["Date"].astype(str)

show_cols = [c for c in [
    DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", "Date", "Mois",
    HONO, AUTRE, TOTAL, "Payé", "Reste"
] if c in view.columns]

# ✅ d'abord trier sur le DataFrame complet (qui contient les colonnes dérivées),
# puis ne sélectionner que les colonnes d'affichage
sort_keys = [c for c in ["_Année_", "_MoisNum_", "Catégorie", "Nom"] if c in view.columns]
view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

st.dataframe(
    view_sorted[show_cols].reset_index(drop=True),
    use_container_width=True,
)

    # Petit rappel des filtres actifs
    with st.expander("🧾 Filtres actifs", expanded=False):
        st.write({
            "Catégories": sel.get("Catégorie", []),
            "Années": sel_years,
            "Mois": sel_months,
            "Solde": solde_mode,
            "Recherche": q,
        })

# ============================================
# VISA APP — PARTIE 3/5
# Clients : créer / modifier / supprimer / paiements multiples
# ============================================

with tab_clients:
    st.subheader("👥 Clients — créer / modifier / supprimer / paiements")

    # Sécu: vérifier que le classeur est chargé (fait en PARTIE 1)
    if df_clients is None or df_clients.empty:
        st.info("Aucun client pour le moment. Crée ton premier client ➕ à droite.")
    live = df_clients.copy()  # vue de travail normalisée

    # ---------- Sélection d’un client existant ----------
    cL, cR = st.columns([1, 1])

    with cL:
        st.markdown("### 🔎 Sélection")
        if live.empty:
            sel_idx = None
            sel_row = None
            st.caption("Aucun client.")
        else:
            labels = (live.get("Nom", "").astype(str) + " — " + live.get("ID_Client", "").astype(str)).fillna("")
            sel_idx = st.selectbox(
                "Client",
                options=list(live.index),
                format_func=lambda i: labels.iloc[i],
                key=f"cli_sel_{sheet_choice}",
            )
            sel_row = live.loc[sel_idx] if sel_idx is not None else None

    # ---------- Création d’un nouveau client ----------
    with cR:
        st.markdown("### ➕ Nouveau client")
        new_name = st.text_input("Nom", key=f"new_nom_{sheet_choice}")
        new_date = st.date_input("Date de création", value=date.today(), key=f"new_date_{sheet_choice}")

        # Choix du visa (code) — basé sur le référentiel si présent
        if 'df_visa' in globals() and isinstance(df_visa, pd.DataFrame) and not df_visa.empty:
            codes = sorted(df_visa["Catégorie"].dropna().astype(str).unique().tolist())
        else:
            codes = sorted(live.get("Visa", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
        new_visa = st.selectbox("Visa (code)", options=[""] + codes, index=0, key=f"new_visa_{sheet_choice}")

        new_hono = st.number_input(HONO, min_value=0.0, step=10.0, format="%.2f", key=f"new_hono_{sheet_choice}")
        new_autr = st.number_input(AUTRE, min_value=0.0, step=10.0, format="%.2f", key=f"new_autr_{sheet_choice}")

        if st.button("💾 Créer", key=f"btn_new_{sheet_choice}"):
            if not new_name:
                st.warning("Renseigne le **Nom**.")
            elif not new_visa:
                st.warning("Choisis un **Visa**.")
            else:
                base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)  # non normalisé pour préserver l’ordre
                base_norm = normalize_clients(base_raw.copy())

                dossier = next_dossier_number(base_norm)
                client_id = _make_client_id_from_row({"Nom": new_name, "Date": new_date})
                # éviter collision d’ID
                origin = client_id
                i = 0
                if "ID_Client" in base_norm.columns:
                    while (base_norm["ID_Client"].astype(str) == client_id).any():
                        i += 1
                        client_id = f"{origin}-{i}"

                total = float(new_hono) + float(new_autr)
                new_row = {
                    DOSSIER_COL: dossier,
                    "ID_Client": client_id,
                    "Nom": new_name,
                    "Date": pd.to_datetime(new_date).date(),
                    "Mois": f"{new_date.month:02d}",
                    "Catégorie": new_visa,                 # ta logique: Catégorie = code racine
                    "Visa": _visa_code_only(new_visa),      # code de base (sans COS/EOS)
                    HONO: float(new_hono),
                    AUTRE: float(new_autr),
                    TOTAL: total,
                    "Payé": 0.0,
                    "Reste": total,
                    PAY_JSON: "[]",
                    "ESCROW transféré (US $)": 0.0,
                    "Journal ESCROW": "[]",
                    "Dossier envoyé": False,
                    "Date envoyé": "",
                    "Dossier approuvé": False,
                    "Date approuvé": "",
                    "RFE": False,
                    "Date RFE": "",
                    "Dossier refusé": False,
                    "Date refusé": "",
                    "Dossier annulé": False,
                    "Date annulé": "",
                }

                base_raw = pd.concat([base_raw, pd.DataFrame([new_row])], ignore_index=True)
                # normalise et écrit
                base_norm = normalize_clients(base_raw)
                write_sheet_inplace(clients_path, sheet_choice, base_norm)
                st.success("✅ Client créé.")
                st.rerun()

    st.markdown("---")

    # Si aucun client sélectionné on s’arrête ici
    if sel_row is None:
        st.stop()

    # ---------- Formulaire d’édition ----------
    idx = sel_idx
    ed = sel_row.to_dict()

    e1, e2, e3 = st.columns(3)
    with e1:
        ed_nom = st.text_input("Nom", value=_safe_str(ed.get("Nom", "")), key=f"ed_nom_{idx}_{sheet_choice}")
        ed_date = st.date_input(
            "Date de création",
            value=(pd.to_datetime(ed.get("Date")).date() if pd.notna(ed.get("Date")) else date.today()),
            key=f"ed_date_{idx}_{sheet_choice}",
        )
    with e2:
        # Visa depuis référentiel si dispo
        codes_all = sorted(df_visa["Catégorie"].dropna().astype(str).unique().tolist()) if 'df_visa' in globals() and not df_visa.empty else sorted(live["Visa"].dropna().astype(str).unique().tolist())
        current_code = _visa_code_only(ed.get("Visa", ""))
        ed_visa = st.selectbox(
            "Visa (code)",
            options=[""] + codes_all,
            index=(codes_all.index(current_code) + 1 if current_code in codes_all else 0),
            key=f"ed_visa_{idx}_{sheet_choice}",
        )
    with e3:
        ed_hono = st.number_input(HONO, min_value=0.0, value=float(ed.get(HONO, 0.0)), step=10.0, format="%.2f", key=f"ed_hono_{idx}_{sheet_choice}")
        ed_autr = st.number_input(AUTRE, min_value=0.0, value=float(ed.get(AUTRE, 0.0)), step=10.0, format="%.2f", key=f"ed_autr_{idx}_{sheet_choice}")

    st.markdown("#### 🧾 Statuts du dossier")
    s1, s2, s3 = st.columns(3)
    with s1:
        ed_env = st.checkbox("Dossier envoyé", value=bool(ed.get("Dossier envoyé", False)), key=f"ed_env_{idx}_{sheet_choice}")
        ed_app = st.checkbox("Dossier approuvé", value=bool(ed.get("Dossier approuvé", False)), key=f"ed_app_{idx}_{sheet_choice}")
    with s2:
        ed_rfe = st.checkbox("RFE", value=bool(ed.get("RFE", False)), key=f"ed_rfe_{idx}_{sheet_choice}")
        ed_ref = st.checkbox("Dossier refusé", value=bool(ed.get("Dossier refusé", False)), key=f"ed_ref_{idx}_{sheet_choice}")
    with s3:
        ed_ann = st.checkbox("Dossier annulé", value=bool(ed.get("Dossier annulé", False)), key=f"ed_ann_{idx}_{sheet_choice}")

    # Contrainte business rappelée (si tu veux la forcer : on peut renforcer ici)
    st.caption("💡 Rappel : RFE ne peut être activé que si l’un des statuts **Envoyé / Approuvé / Refusé / Annulé** est vrai.")

    # ---------- Paiements (acomptes multiples) ----------
    st.markdown("### 💳 Paiements (acomptes)")

    p1, p2, p3, p4 = st.columns([1, 1, 1, 2])
    with p1:
        p_date = st.date_input("Date paiement", value=date.today(), key=f"p_date_{idx}_{sheet_choice}")
    with p2:
        p_mode = st.selectbox("Mode", ["CB", "Chèque", "Cash", "Virement", "Venmo"], key=f"p_mode_{idx}_{sheet_choice}")
    with p3:
        p_amt = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"p_amt_{idx}_{sheet_choice}")
    with p4:
        if st.button("➕ Ajouter paiement", key=f"btn_addpay_{idx}_{sheet_choice}"):
            base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
            base_norm = normalize_clients(base_raw.copy())

            # contrôle solde
            if float(p_amt) <= 0:
                st.warning("Le montant doit être > 0.")
            else:
                # récupère la ligne réelle via ID_Client si possible
                idc = _safe_str(ed.get("ID_Client", ""))
                if idc and "ID_Client" in base_raw.columns:
                    try:
                        real_idx = base_raw.index[base_raw["ID_Client"].astype(str) == idc][0]
                    except Exception:
                        real_idx = idx
                else:
                    real_idx = idx

                row = base_raw.loc[real_idx].to_dict()
                # parse JSON paiements
                try:
                    plist = json.loads(row.get(PAY_JSON, "[]"))
                    if not isinstance(plist, list):
                        plist = []
                except Exception:
                    plist = []
                plist.append({"date": str(p_date), "mode": p_mode, "amount": float(p_amt)})
                row[PAY_JSON] = json.dumps(plist, ensure_ascii=False)

                base_raw.loc[real_idx] = row
                base_norm = normalize_clients(base_raw.copy())
                write_sheet_inplace(clients_path, sheet_name=sheet_choice, df=base_norm)
                st.success("Paiement ajouté.")
                st.rerun()

    # Historique paiements
    try:
        hist = json.loads(_safe_str(sel_row.get(PAY_JSON, "[]")))
        if not isinstance(hist, list):
            hist = []
    except Exception:
        hist = []
    st.write("**Historique des paiements**")
    if hist:
        h = pd.DataFrame(hist)
        if "amount" in h.columns:
            h["amount"] = h["amount"].astype(float).map(_fmt_money_us)
        st.dataframe(h, use_container_width=True)
    else:
        st.caption("Aucun paiement saisi.")

    st.markdown("---")

    # ---------- Boutons : Sauvegarder / Supprimer ----------
    a1, a2 = st.columns([1, 1])

    if a1.button("💾 Sauvegarder les modifications", key=f"btn_save_{idx}_{sheet_choice}"):
        base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
        # retrouver la ligne réelle via ID si possible
        idc = _safe_str(ed.get("ID_Client", ""))
        if idc and "ID_Client" in base_raw.columns:
            idx_real_list = base_raw.index[base_raw["ID_Client"].astype(str) == idc].tolist()
            real_idx = idx_real_list[0] if idx_real_list else None
        else:
            real_idx = idx

        if real_idx is None or not (0 <= real_idx < len(base_raw)):
            st.error("Ligne introuvable.")
        else:
            row = base_raw.loc[real_idx].to_dict()

            # maj champs
            row["Nom"] = ed_nom
            row["Date"] = pd.to_datetime(ed_date).date()
            row["Mois"] = f"{ed_date.month:02d}"
            if ed_visa:
                row["Catégorie"] = ed_visa
                row["Visa"] = _visa_code_only(ed_visa)
            row[HONO] = float(ed_hono)
            row[AUTRE] = float(ed_autr)
            row[TOTAL] = float(ed_hono) + float(ed_autr)

            # statuts
            row["Dossier envoyé"] = bool(ed_env)
            row["Dossier approuvé"] = bool(ed_app)
            row["RFE"] = bool(ed_rfe)
            row["Dossier refusé"] = bool(ed_ref)
            row["Dossier annulé"] = bool(ed_ann)

            base_raw.loc[real_idx] = row
            base_norm = normalize_clients(base_raw.copy())
            write_sheet_inplace(clients_path, sheet_choice, base_norm)
            st.success("✅ Modifications sauvegardées.")
            st.rerun()

    if a2.button("🗑️ Supprimer ce client", key=f"btn_del_{idx}_{sheet_choice}"):
        base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
        # supprimer via ID si possible
        idc = _safe_str(ed.get("ID_Client", ""))
        if idc and "ID_Client" in base_raw.columns:
            mask = base_raw["ID_Client"].astype(str) == idc
            base_raw = base_raw.loc[~mask].reset_index(drop=True)
        else:
            if 0 <= idx < len(base_raw):
                base_raw = base_raw.drop(index=idx).reset_index(drop=True)
            else:
                st.error("Ligne introuvable."); st.stop()

        base_norm = normalize_clients(base_raw.copy())
        write_sheet_inplace(clients_path, sheet_choice, base_norm)
        st.success("🗑️ Client supprimé.")
        st.rerun()

# ============================================
# VISA APP — PARTIE 4/5
# Analyses : filtres + KPI + comparaisons + détails
# ============================================

with tab_analyses:
    st.subheader("📈 Analyses — Volumes & Financier")

    # --- 1) Filtres VISA hiérarchiques (réutilise le référentiel) ---
    df_visa_safe = _ensure_visa_columns(df_visa)
    if df_visa_safe.empty:
        st.warning("⚠️ Le référentiel Visa est vide ou mal formé. Les filtres de catégories sont désactivés.")
        sel = {"__whitelist_visa__": [], "Catégorie": []}
        base = df_clients.copy()
    else:
        sel = build_checkbox_filters_grouped(
            df_visa_safe,
            keyprefix=f"flt_ana_{sheet_choice}",
            as_toggle=False,
        )
        base = filter_clients_by_ref(df_clients, sel)

    # Champs dérivés année/mois
    base = base.copy()
    base["_Année_"] = base["Date"].apply(lambda x: x.year if pd.notna(x) else pd.NA)
    base["_MoisNum_"] = base["Date"].apply(lambda x: int(x.month) if pd.notna(x) else pd.NA)
    base["_Mois_"] = base["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) else pd.NA)

    # --- 2) Filtres additionnels ---
    cR1, cR2, cR3, cR4 = st.columns([1, 1, 1, 2])
    yearsA  = sorted([int(y) for y in base["_Année_"].dropna().unique()]) if not base.empty else []
    monthsA = [f"{m:02d}" for m in sorted([int(m) for m in base["_MoisNum_"].dropna().unique()])] if not base.empty else []

    with cR1:
        sel_years  = st.multiselect("Année", yearsA, default=[], key=f"ana_year_{sheet_choice}")
    with cR2:
        sel_months = st.multiselect("Mois (MM)", monthsA, default=[], key=f"ana_month_{sheet_choice}")
    with cR3:
        solde_mode = st.selectbox("Solde",
                                  ["Tous", "Soldé (Reste = 0)", "Non soldé (Reste > 0)"],
                                  index=0, key=f"ana_solde_{sheet_choice}")
    with cR4:
        q = st.text_input("Recherche (nom, ID, visa…)", "", key=f"ana_q_{sheet_choice}")

    ff = base.copy()
    if sel_years:
        ff = ff[ff["_Année_"].isin(sel_years)]
    if sel_months:
        ff = ff[ff["_Mois_"].astype(str).isin(sel_months)]
    if solde_mode == "Soldé (Reste = 0)":
        ff = ff[_safe_num_series(ff, "Reste") <= 0.0000001]
    elif solde_mode == "Non soldé (Reste > 0)":
        ff = ff[_safe_num_series(ff, "Reste") > 0.0000001]
    if q:
        qn = q.lower().strip()
        def _row_match(r):
            hay = " ".join([
                _safe_str(r.get("Nom","")),
                _safe_str(r.get("ID_Client","")),
                _safe_str(r.get("Catégorie","")),
                _safe_str(r.get("Visa","")),
                str(r.get(DOSSIER_COL,"")),
            ]).lower()
            return qn in hay
        ff = ff[ff.apply(_row_match, axis=1)]

    # --- 3) KPI globaux sur le périmètre filtré ---
    st.markdown("""
    <style>
    .small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}
    .small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}
    </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(ff)}")
    k2.metric("Honoraires", _fmt_money_us(_safe_num_series(ff, HONO).sum()))
    k3.metric("Encaissements (Payé)", _fmt_money_us(_safe_num_series(ff, "Payé").sum()))
    k4.metric("Solde à encaisser", _fmt_money_us(_safe_num_series(ff, "Reste").sum()))
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # --- 4) Comparaison Année → Année (volumes + financier) ---
    st.markdown("### 📆 Comparaison Année → Année")
    if ff["_Année_"].dropna().empty:
        st.info("Aucune date exploitable pour la comparaison annuelle.")
    else:
        grpY = ff.groupby("_Année_", dropna=True).agg(
            Dossiers = ("ID_Client", "count"),
            Honoraires = (HONO, lambda s: _safe_num_series(ff.loc[s.index], HONO).sum()),
            Paye = ("Payé",   lambda s: _safe_num_series(ff.loc[s.index], "Payé").sum()),
            Reste = ("Reste", lambda s: _safe_num_series(ff.loc[s.index], "Reste").sum()),
        ).reset_index().rename(columns={"_Année_":"Année"})
        grpY = grpY.sort_values("Année")

        st.dataframe(grpY, use_container_width=True)

        # Barres volumes
        try:
            import altair as alt
            ch1 = alt.Chart(grpY).mark_bar().encode(
                x=alt.X("Année:O", sort=None),
                y=alt.Y("Dossiers:Q")
            ).properties(height=240)
            st.altair_chart(ch1, use_container_width=True)
        except Exception:
            pass

        # Lignes financier
        try:
            import altair as alt
            g_long = grpY.melt(id_vars=["Année"], value_vars=["Honoraires","Paye","Reste"],
                               var_name="Type", value_name="Montant")
            ch2 = alt.Chart(g_long).mark_line(point=True).encode(
                x=alt.X("Année:O", sort=None),
                y=alt.Y("Montant:Q"),
                color="Type:N"
            ).properties(height=260)
            st.altair_chart(ch2, use_container_width=True)
        except Exception:
            pass

    st.markdown("---")

    # --- 5) Comparaison par Mois (toutes années confondues) ---
    st.markdown("### 🗓️ Par mois (toutes années)")
    if ff["_Mois_"].dropna().empty:
        st.info("Aucun mois exploitable.")
    else:
        grpM = ff.groupby("_Mois_", dropna=True).agg(
            Dossiers = ("ID_Client", "count"),
            Honoraires = (HONO, lambda s: _safe_num_series(ff.loc[s.index], HONO).sum()),
            Paye = ("Payé",   lambda s: _safe_num_series(ff.loc[s.index], "Payé").sum()),
            Reste = ("Reste", lambda s: _safe_num_series(ff.loc[s.index], "Reste").sum()),
        ).reset_index().rename(columns={"_Mois_":"Mois"})
        grpM = grpM.sort_values("Mois")

        st.dataframe(grpM, use_container_width=True)

        try:
            import altair as alt
            ch3 = alt.Chart(grpM).mark_bar().encode(
                x=alt.X("Mois:O", sort=None),
                y=alt.Y("Dossiers:Q")
            ).properties(height=220)
            st.altair_chart(ch3, use_container_width=True)
        except Exception:
            pass

    st.markdown("---")


# --- 6) Détails des dossiers correspondants (liste clients) ---
st.markdown("### 📋 Détails des dossiers filtrés")
detail = ff.copy()
for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
    if c in detail.columns:
        detail[c] = _safe_num_series(detail, c).map(_fmt_money_us)
if "Date" in detail.columns:
    detail["Date"] = detail["Date"].astype(str)

show_cols = [c for c in [
    DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", "Date", "Mois",
    HONO, AUTRE, TOTAL, "Payé", "Reste",
    "Dossier envoyé", "Dossier approuvé", "RFE", "Dossier refusé", "Dossier annulé"
] if c in detail.columns]

# ✅ trier avant de sélectionner les colonnes
sort_keys = [c for c in ["_Année_", "_MoisNum_", "Catégorie", "Nom"] if c in detail.columns]
detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail

st.dataframe(detail_sorted[show_cols].reset_index(drop=True), use_container_width=True)

    with st.expander("🧾 Filtres actifs", expanded=False):
        st.write({
            "Catégories": sel.get("Catégorie", []),
            "Années": sel_years,
            "Mois": sel_months,
            "Solde": solde_mode,
            "Recherche": q,
        })

# ============================================
# VISA APP — PARTIE 5/5
# ESCROW : calculs, transferts, journal & alertes
# ============================================

with tab_escrow:
    st.subheader("🏦 ESCROW — dépôts, transferts & alertes")

    # ------- helpers locaux -------
    def _sum_payments_from_json(js: str) -> float:
        try:
            lst = json.loads(_safe_str(js) or "[]")
            if not isinstance(lst, list):
                return 0.0
            s = 0.0
            for it in lst:
                try:
                    s += float(it.get("amount", 0.0))
                except Exception:
                    pass
            return max(0.0, s)
        except Exception:
            return 0.0

    def _escrow_row_metrics(r: pd.Series) -> dict:
        """Calcule les éléments financiers ESCROW pour une ligne client."""
        hono = float(r.get(HONO, 0.0) or 0.0)
        pay_decl = float(r.get("Payé", 0.0) or 0.0)
        pay_js = _sum_payments_from_json(r.get(PAY_JSON, "[]"))
        pay = max(pay_decl, pay_js)  # tolérant: prend le plus grand des 2
        transf = float(r.get("ESCROW transféré (US $)", 0.0) or 0.0)
        dispo = max(min(pay, hono) - transf, 0.0)
        return {
            "honoraires": hono,
            "paye": pay,
            "transfere": transf,
            "dispo": dispo,
        }

    # Vue normalisée
    live = df_clients.copy()

    # ------- filtres rapides -------
    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    with c1:
        only_with_dispo = st.toggle("Uniquement dossiers avec ESCROW disponible", value=True, key="esc_onlydispo")
    with c2:
        only_sent = st.toggle("Uniquement dossiers envoyés", value=False, key="esc_onlysent")
    with c3:
        order_by_dispo = st.toggle("Trier par ESCROW disponible", value=True, key="esc_sortdispo")
    with c4:
        q = st.text_input("Recherche (Nom / ID / Dossier N / Visa)", "", key="esc_q")

    # ------- calculs par ligne -------
    rows = []
    for i, r in live.iterrows():
        m = _escrow_row_metrics(r)
        row = {
            "idx": i,
            DOSSIER_COL: r.get(DOSSIER_COL, ""),
            "ID_Client": r.get("ID_Client", ""),
            "Nom": r.get("Nom", ""),
            "Catégorie": r.get("Catégorie", ""),
            "Visa": r.get("Visa", ""),
            "Dossier envoyé": bool(r.get("Dossier envoyé", False)),
            HONO: m["honoraires"],
            "Payé_calc": m["paye"],
            "ESCROW transféré (US $)": m["transfere"],
            "ESCROW dispo": m["dispo"],
            "Journal ESCROW": _safe_str(r.get("Journal ESCROW", "[]")),
        }
        rows.append(row)

    jdf = pd.DataFrame(rows)
    if only_with_dispo:
        jdf = jdf[jdf["ESCROW dispo"] > 0.0]
    if only_sent:
        if "Dossier envoyé" in jdf.columns:
            jdf = jdf[jdf["Dossier envoyé"] == True]
    if q:
        qn = q.lower().strip()
        def _m(row):
            hay = " ".join([
                _safe_str(row.get("Nom","")),
                _safe_str(row.get("ID_Client","")),
                str(row.get(DOSSIER_COL,"")),
                _safe_str(row.get("Visa","")),
                _safe_str(row.get("Catégorie","")),
            ]).lower()
            return qn in hay
        jdf = jdf[jdf.apply(_m, axis=1)]

    if order_by_dispo and not jdf.empty:
        jdf = jdf.sort_values("ESCROW dispo", ascending=False)

    # ------- KPI globaux -------
    st.markdown("""
    <style>
      .small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}
      .small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}
    </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers listés", f"{len(jdf)}")
    k2.metric("Honoraires (périmètre)", _fmt_money_us(float(jdf[HONO].sum()) if not jdf.empty else 0.0))
    k3.metric("ESCROW transféré (périmètre)", _fmt_money_us(float(jdf["ESCROW transféré (US $)"].sum()) if not jdf.empty else 0.0))
    k4.metric("ESCROW dispo (périmètre)", _fmt_money_us(float(jdf["ESCROW dispo"].sum()) if not jdf.empty else 0.0))
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # ------- tableau & actions -------
    if jdf.empty:
        st.info("Aucun dossier à afficher avec les filtres actuels.")
        st.stop()

    # affichage tableau lisible
    show = jdf.copy()
    show[HONO] = show[HONO].map(_fmt_money_us)
    show["Payé_calc"] = show["Payé_calc"].map(_fmt_money_us)
    show["ESCROW transféré (US $)"] = show["ESCROW transféré (US $)"].map(_fmt_money_us)
    show["ESCROW dispo"] = show["ESCROW dispo"].map(_fmt_money_us)
    st.dataframe(
        show[[DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", HONO, "Payé_calc", "ESCROW transféré (US $)", "ESCROW dispo", "Dossier envoyé"]]
        .reset_index(drop=True),
        use_container_width=True
    )

    st.markdown("### ↗️ Enregistrer un transfert ESCROW")
    st.caption("Rappels : l’ESCROW disponible = min(Payé, Honoraires) − déjà transféré. Seule la partie **honoraires** est transférée.")

    for _, r in jdf.iterrows():
        with st.expander(f"{r['Nom']} — {r['ID_Client']} — Dossier {r[DOSSIER_COL]}", expanded=False):
            cA, cB, cC, cD = st.columns([1, 1, 1, 2])
            dispo = float(r["ESCROW dispo"])
            with cA:
                st.write("**ESCROW disponible**")
                st.write(_fmt_money_us(dispo))
            with cB:
                t_date = st.date_input("Date transfert", value=date.today(), key=f"esc_dt_{r['ID_Client']}")
            with cC:
                amt = st.number_input("Montant à transférer (US $)",
                                      min_value=0.0, value=float(dispo),
                                      step=10.0, format="%.2f",
                                      key=f"esc_amt_{r['ID_Client']}")
            with cD:
                note = st.text_input("Note (optionnel)", value="", key=f"esc_note_{r['ID_Client']}")

            can_transfer = dispo > 0 and amt > 0 and amt <= dispo + 1e-9
            if st.button("💸 Enregistrer transfert", key=f"esc_btn_{r['ID_Client']}"):
                if not can_transfer:
                    st.warning("Montant invalide (doit être > 0 et ≤ ESCROW disponible).")
                    st.stop()
                # écriture dans le classeur
                base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
                # retrouver la ligne par ID_Client si possible
                if "ID_Client" in base_raw.columns:
                    idxs = base_raw.index[base_raw["ID_Client"].astype(str) == str(r["ID_Client"])].tolist()
                    real_idx = idxs[0] if idxs else None
                else:
                    real_idx = int(r["idx"])
                if real_idx is None or not (0 <= real_idx < len(base_raw)):
                    st.error("Ligne introuvable pour ce client.")
                    st.stop()

                row = base_raw.loc[real_idx].to_dict()
                # maj transféré
                curr_tr = float(row.get("ESCROW transféré (US $)", 0.0) or 0.0)
                row["ESCROW transféré (US $)"] = curr_tr + float(amt)

                # append journal
                try:
                    jlist = json.loads(_safe_str(row.get("Journal ESCROW", "[]")) or "[]")
                    if not isinstance(jlist, list):
                        jlist = []
                except Exception:
                    jlist = []
                jlist.append({
                    "ts": datetime.now().isoformat(timespec="seconds"),
                    "date": str(t_date),
                    "amount": float(amt),
                    "note": note,
                })
                row["Journal ESCROW"] = json.dumps(jlist, ensure_ascii=False)

                base_raw.loc[real_idx] = row
                base_norm = normalize_clients(base_raw.copy())
                write_sheet_inplace(clients_path, sheet_choice, base_norm)
                st.success("✅ Transfert enregistré.")
                st.rerun()

    st.markdown("---")

    # ------- alertes : dossiers envoyés à réclamer -------
    st.markdown("### 🚨 Alertes — dossiers envoyés à réclamer")
    alert_df = jdf[(jdf["Dossier envoyé"] == True) & (jdf["ESCROW dispo"] > 0.0)] if "Dossier envoyé" in jdf.columns else pd.DataFrame()
    if alert_df.empty:
        st.success("Aucune alerte : tous les dossiers envoyés ont leur ESCROW transféré ✅.")
    else:
        alert_view = alert_df.copy()
        alert_view["ESCROW dispo"] = alert_view["ESCROW dispo"].map(_fmt_money_us)
        st.dataframe(alert_view[[DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", "ESCROW dispo"]],
                     use_container_width=True)

    st.markdown("---")

    # ------- Journal ESCROW global (optionnel) -------
    st.markdown("### 📚 Journal ESCROW — global")
    logs = []
    for _, r in jdf.iterrows():
        try:
            jlist = json.loads(_safe_str(r.get("Journal ESCROW", "[]")) or "[]")
            if not isinstance(jlist, list):
                continue
            for it in jlist:
                logs.append({
                    "Dossier N": r.get(DOSSIER_COL, ""),
                    "ID_Client": r.get("ID_Client", ""),
                    "Nom": r.get("Nom", ""),
                    "Date": it.get("date",""),
                    "Horodatage": it.get("ts",""),
                    "Montant": float(it.get("amount", 0.0) or 0.0),
                    "Note": it.get("note",""),
                })
        except Exception:
            continue

    if logs:
        j = pd.DataFrame(logs).sort_values(["Horodatage","Date"], na_position="last").reset_index(drop=True)
        j["Montant"] = j["Montant"].map(_fmt_money_us)
        st.dataframe(j, use_container_width=True)
    else:
        st.caption("Aucun mouvement dans le journal ESCROW.")

