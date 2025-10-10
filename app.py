# ========================
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

# =============================
# 🧩 Filtres contextuels VISA — version robuste
# =============================
import re
import pandas as pd
import streamlit as st

# --- Liste hiérarchique des niveaux de catégories ---
REF_LEVELS = ["Catégorie"] + [f"Sous-categories {i}" for i in range(1,9)]

def _ensure_visa_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et structure le fichier Visa pour éviter les erreurs."""
    if df is None or df.empty:
        return pd.DataFrame(columns=REF_LEVELS + ["Actif"])
    out = df.copy()

    # normalisation des noms de colonnes
    norm = {re.sub(r"[^a-z0-9]+","", str(c).lower()): str(c) for c in out.columns}
    def find_col(*cands):
        for cand in cands:
            key = re.sub(r"[^a-z0-9]+","", cand.lower())
            if key in norm:
                return norm[key]
            for k, orig in norm.items():
                if key in k:
                    return orig
        return None

    # renommer les colonnes
    cat = find_col("Catégorie","Categorie","Category")
    out = out.rename(columns={cat: "Catégorie"}) if cat else out.assign(**{"Catégorie": ""})
    for i in range(1,9):
        sc = find_col(f"Sous-categories {i}", f"Sous categorie {i}", f"SC{i}")
        if sc:
            out = out.rename(columns={sc: f"Sous-categories {i}"})
        else:
            out[f"Sous-categories {i}"] = ""
    act = find_col("Actif","Active","Inclure","Afficher")
    out = out.rename(columns={act: "Actif"}) if act else out.assign(**{"Actif": 1})

    # ne garder que les colonnes utiles
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
    out = out[mask].reset_index(drop=True)
    return out

def _slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(s).lower()).strip("_")

def _multi_bool_inputs(options: list[str], label: str, keyprefix: str, as_toggle: bool=False) -> list[str]:
    """Affiche une sélection multi-case (checkbox ou toggle)"""
    options = [o for o in options if str(o).strip() != ""]
    if not options:
        st.caption(f"Aucune option pour **{label}**.")
        return []
    with st.expander(label, expanded=False):
        c1, c2 = st.columns(2)
        all_on = c1.toggle("Tout sélectionner", value=False, key=f"{keyprefix}_all")
        none_on = c2.toggle("Tout désélectionner", value=False, key=f"{keyprefix}_none")
        selected = []
        cols = st.columns(3 if len(options) > 6 else 2)
        for i, opt in enumerate(sorted(options)):
            k = f"{keyprefix}_{i}"
            if all_on: st.session_state[k] = True
            if none_on: st.session_state[k] = False
            with cols[i % len(cols)]:
                val = st.toggle(opt, value=st.session_state.get(k, False), key=k) if as_toggle \
                      else st.checkbox(opt, value=st.session_state.get(k, False), key=k)
                if val:
                    selected.append(opt)
    return selected

def build_checkbox_filters_grouped(df_ref_in: pd.DataFrame, keyprefix: str, as_toggle: bool=False) -> dict:
    """Construit l’arborescence dynamique de filtres"""
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
        for i in range(1,9):
            col = f"Sous-categories {i}"
            options = sorted([v for v in sub[col].unique() if str(v).strip() != ""])
            picked = _multi_bool_inputs(options, f"{cat} — {col}", f"{keyprefix}_{_slug(cat)}_sc{i}", as_toggle=as_toggle)
            res["SC_map"][cat][col] = picked
            if picked:
                sub = sub[sub[col].isin(picked)]
        whitelist.add(cat)

    res["__whitelist_visa__"] = sorted(whitelist)
    return res

def filter_clients_by_ref(df_clients: pd.DataFrame, sel: dict) -> pd.DataFrame:
    """Applique le filtre de sélection au tableau des clients"""
    if df_clients is None or df_clients.empty:
        return df_clients
    f = df_clients.copy()
    wl = set(sel.get("__whitelist_visa__", []))
    if wl and "Catégorie" in f.columns:
        f = f[f["Catégorie"].astype(str).isin(wl)]
    return f

# --- Helper additionnel pour colonnes dupliquées (sécurité KPI) ---
def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([], dtype=float)
    s = df[col]
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0]
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

# ---------- Constantes colonnes ----------
DOSSIER_COL = "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"
PAY_JSON = "Paiements"

S_ENVOYE,   D_ENVOYE   = "Dossier envoyé",  "Date envoyé"
S_APPROUVE, D_APPROUVE = "Dossier approuvé","Date approuvé"
S_RFE,      D_RFE      = "RFE",             "Date RFE"
S_REFUSE,   D_REFUSE   = "Dossier refusé",  "Date refusé"
S_ANNULE,   D_ANNULE   = "Dossier annulé",  "Date annulé"
STATUS_COLS  = [S_ENVOYE, S_APPROUVE, S_RFE, S_REFUSE, S_ANNULE]
STATUS_DATES = [D_ENVOYE, D_APPROUVE, D_RFE, D_REFUSE, D_ANNULE]

ESC_TR = "ESCROW transféré (US $)"
ESC_JR = "Journal ESCROW"

DOSSIER_START = 13057

# ---------- Persistance derniers chemins ----------
STATE_FILE = Path(".visa_app_state.json")

def _save_last_paths(clients: Path|None=None, visa: Path|None=None):
    data = {}
    if STATE_FILE.exists():
        try: data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception: data = {}
    if clients is not None: data["clients_path"] = str(clients)
    if visa is not None:    data["visa_path"]    = str(visa)
    STATE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def _load_last_paths() -> tuple[Path|None, Path|None]:
    if not STATE_FILE.exists(): return None, None
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        c = Path(data.get("clients_path","")) if data.get("clients_path") else None
        v = Path(data.get("visa_path",""))    if data.get("visa_path") else None
        if c and not c.exists(): c = None
        if v and not v.exists(): v = None
        return c, v
    except Exception:
        return None, None

# ---------- Helpers ----------
def _safe_str(x) -> str:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)): return ""
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

def _visa_code_only(v: str) -> str:
    s = _safe_str(v)
    if not s: return ""
    parts = s.split()
    if len(parts) >= 2 and parts[-1].upper() in {"COS","EOS"}:
        return " ".join(parts[:-1]).strip()
    return s.strip()

def _to_num(s: Any) -> pd.Series:
    if s is None: return pd.Series(dtype=float)
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:,0] if s.shape[1] else pd.Series(dtype=float)
    s = pd.Series(s).astype(str).str.replace(r"[^\d,.\-]", "", regex=True)
    def _one(x):
        if x=="" or x=="-": return 0.0
        if x.count(",")==1 and x.count(".")==0: x=x.replace(",",".")
        if x.count(".")==1 and x.count(",")>=1: x=x.replace(",","")
        try: return float(x)
        except: return 0.0
    return s.map(_one)

def _to_int(s: Any) -> pd.Series:
    try: return pd.to_numeric(pd.Series(s), errors="coerce").fillna(0).astype(int)
    except Exception: return pd.Series([0]*len(pd.Series(s)), dtype=int)

def _fmt_money_us(v: float) -> str:
    try: return f"${v:,.2f}"
    except: return "$0.00"

def _parse_json_list(val: Any) -> list:
    if val is None: return []
    if isinstance(val, list): return val
    try:
        out = json.loads(val)
        return out if isinstance(out, list) else []
    except Exception:
        return []

def _sum_payments(lst: list[dict]) -> float:
    return sum(float(e.get("amount",0.0) or 0.0) for e in lst)

# ---------- Excel utils ----------
def list_sheets(path: Path) -> list[str]:
    try: return pd.ExcelFile(path).sheet_names
    except Exception: return []

def read_sheet(path: Path, sheet: str, normalize: bool=False) -> pd.DataFrame:
    try: df = pd.read_excel(path, sheet_name=sheet)
    except Exception: return pd.DataFrame()
    return normalize_clients(df) if normalize else df

def write_sheet_inplace(path: Path, sheet: str, df: pd.DataFrame):
    path = Path(path)
    try:
        if path.exists():
            book = pd.ExcelFile(path)
            sheets = {sn: pd.read_excel(path, sheet_name=sn) for sn in book.sheet_names}
        else:
            sheets = {}
        sheets[sheet] = df
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            for sn, sdf in sheets.items():
                sdf.to_excel(w, sheet_name=sn, index=False)
    except Exception as e:
        st.error(f"Erreur écriture Excel: {e}")
        raise

# ---------- Dossiers / IDs ----------
def ensure_dossier_numbers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if DOSSIER_COL not in df.columns: df[DOSSIER_COL] = 0
    nums = _to_int(df[DOSSIER_COL])
    if (nums == 0).all():
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
    if df is None or df.empty or DOSSIER_COL not in df.columns: return DOSSIER_START
    nums = _to_int(df[DOSSIER_COL])
    m = int(nums.max()) if len(nums) else (DOSSIER_START - 1)
    return max(m, DOSSIER_START-1) + 1

def _make_client_id_from_row(row: dict) -> str:
    nom = _safe_str(row.get("Nom"))
    try: d = pd.to_datetime(row.get("Date")).date()
    except Exception: d = date.today()
    base = f"{nom}-{d.strftime('%Y%m%d')}"
    base = re.sub(r"[^A-Za-z0-9\-]+", "", base.replace(" ", "-"))
    return base.lower()

def _collapse_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return df
    cols = df.columns.astype(str)
    if not cols.duplicated().any(): return df
    out = pd.DataFrame(index=df.index)
    for col in pd.unique(cols):
        same = df.loc[:, cols == col]
        if same.shape[1] == 1:
            out[col] = same.iloc[:, 0]; continue
        try:
            same_num = same.apply(pd.to_numeric, errors="coerce")
            if same_num.notna().any().any():
                out[col] = same_num.sum(axis=1, skipna=True); continue
        except Exception: pass
        def _first_non_empty(row):
            for v in row:
                if pd.notna(v) and str(v).strip() != "": return v
            return ""
        out[col] = same.apply(_first_non_empty, axis=1)
    return out

# ---------- Normalisation Clients ----------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame()
    df = _collapse_duplicate_columns(df.copy())

    # Renommages souples
    ren = {}
    for c in df.columns:
        lc = _norm_txt(c)
        if "montant honoraires" in lc or lc=="honoraires": ren[c]=HONO
        elif "autres frais" in lc or lc=="autres": ren[c]=AUTRE
        elif lc.startswith("total"): ren[c]=TOTAL
        elif lc in {"reste","solde"}: ren[c]="Reste"
        elif "paye" in lc or "payé" in lc: ren[c]="Payé"
        elif "categorie" in lc: ren[c]="Catégorie"
        elif lc in {"visa"}: ren[c]="Visa"
        elif lc in {"dossier n","dossier"}: ren[c]=DOSSIER_COL
        elif lc == "id_client": ren[c] = "ID_Client"
        elif lc == "nom": ren[c] = "Nom"
        elif lc == "date": ren[c] = "Date"
        elif lc == "mois": ren[c] = "Mois"
    if ren: df = df.rename(columns=ren)

    # Colonnes minimales
    for c in [DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Date","Mois",
              HONO,AUTRE,TOTAL,"Payé","Reste",PAY_JSON,ESC_TR,ESC_JR] + STATUS_COLS + STATUS_DATES:
        if c not in df.columns:
            if c in [HONO,AUTRE,TOTAL,"Payé","Reste",ESC_TR]:
                df[c] = 0.0
            elif c in [PAY_JSON,ESC_JR,"ID_Client","Nom","Catégorie","Visa","Date","Mois"]:
                df[c] = ""
            elif c in STATUS_COLS:
                df[c] = False
            elif c in STATUS_DATES:
                df[c] = ""

    # Canoniser Visa/Catégorie
    df["Visa"] = df["Visa"].map(_visa_code_only)
    df["Catégorie"] = df["Catégorie"].replace("", pd.NA).fillna(df["Visa"]).astype(str)

    # Numériques
    for c in [HONO,AUTRE,TOTAL,"Payé","Reste",ESC_TR]:
        df[c] = _to_num(df[c])

    # Date & Mois
    def _to_date(x):
        try:
            if x=="" or pd.isna(x): return pd.NaT
            return pd.to_datetime(x).date()
        except: return pd.NaT
    df["Date"] = df["Date"].map(_to_date)
    df["Mois"] = df["Date"].apply(lambda d: f"{d.month:02d}" if pd.notna(d) else pd.NA)

    # Payé depuis JSON (si existant)
    paid_from_json = []
    for _, r in df.iterrows():
        plist = _parse_json_list(r.get(PAY_JSON, ""))
        paid_from_json.append(_sum_payments(plist))
    paid_from_json = pd.Series(paid_from_json, index=df.index, dtype=float)
    df["Payé"] = pd.Series([max(a, b) for a, b in zip(_to_num(df["Payé"]), paid_from_json)], index=df.index)

    # Totaux
    df[TOTAL] = _to_num(df.get(HONO, 0.0)) + _to_num(df.get(AUTRE, 0.0))
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # Numéros
    df = ensure_dossier_numbers(df)
    return df

# ---------- Référentiel VISA.xlsx : Catégorie + Sous-catégories 1..8 ----------
REF_LEVELS = ["Catégorie"] + [f"Sous-categories {i}" for i in range(1,9)]

def _find_col(df: pd.DataFrame, candidates: list[str]) -> str|None:
    if df is None or df.empty: return None
    m = {_norm_txt(c): str(c) for c in df.columns.astype(str)}
    for t in candidates:
        nt = _norm_txt(t)
        if nt in m: return m[nt]
    for t in candidates:
        nt = _norm_txt(t)
        for k,orig in m.items():
            if nt in k: return orig
    return None

def read_visa_matrix(visa_path: Path) -> pd.DataFrame:
    """Lit Visa.xlsx (onglet 'Visa' ou 'Visa_normalise')"""
    if visa_path is None or not Path(visa_path).exists():
        return pd.DataFrame(columns=REF_LEVELS)
    try:
        xls = pd.ExcelFile(visa_path)
        sn = "Visa" if "Visa" in xls.sheet_names else ("Visa_normalise" if "Visa_normalise" in xls.sheet_names else xls.sheet_names[0])
        base = pd.read_excel(visa_path, sheet_name=sn)
    except Exception:
        return pd.DataFrame(columns=REF_LEVELS)

    cols = {}
    for lvl in REF_LEVELS:
        col = _find_col(base, [lvl, lvl.replace("categories","catégories"), lvl.replace("categories","categorie")])
        cols[lvl] = col

    out = pd.DataFrame()
    for lvl in REF_LEVELS:
        out[lvl] = base[cols[lvl]] if cols[lvl] else ""

    for c in REF_LEVELS:
        out[c] = out[c].fillna("").astype(str).str.strip()

    # drop lignes vides, ffill Catégorie
    out = out[~(out.apply(lambda r: "".join(r.values), axis=1)=="")].reset_index(drop=True)
    out["Catégorie"] = out["Catégorie"].replace("", pd.NA).ffill().fillna("")

    # VisaCode = code de la première colonne (ex. E-2, B-1…)
    out["VisaCode"] = out["Catégorie"].apply(_visa_code_only)

    # Chemin lisible
    def path_str(row):
        parts = [row["Catégorie"]] + [row[f"Sous-categories {i}"] for i in range(1,9)]
        parts = [p for p in parts if _safe_str(p)]
        return " > ".join(parts)
    out["Path"] = out.apply(path_str, axis=1)
    return out


# ================= DASHBOARD =================
with tab_dash:
    st.subheader("📊 Dashboard")

    # Charger le référentiel Visa
    df_visa_safe = _ensure_visa_columns(df_visa if 'df_visa' in globals() else pd.DataFrame())
    if df_visa_safe.empty:
        st.warning("⚠️ Le référentiel Visa est vide ou mal formé. Charge d'abord ton fichier Visa.xlsx.")
        sel = {"__whitelist_visa__": [], "Catégorie": []}
        f = df_clients.copy()
    else:
        # Construction dynamique des filtres (checkboxes hiérarchiques)
        sel = build_checkbox_filters_grouped(df_visa_safe, keyprefix=f"flt_dash_{sheet_choice}", as_toggle=False)
        f = filter_clients_by_ref(df_clients, sel)

    # --- KPI principaux ---
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Honoraires", _fmt_money_us(_safe_num_series(f, "Montant honoraires (US $)").sum()))
    k3.metric("Payé",      _fmt_money_us(_safe_num_series(f, "Payé").sum()))
    k4.metric("Solde",     _fmt_money_us(_safe_num_series(f, "Reste").sum()))
    st.markdown('</div>', unsafe_allow_html=True)

    # --- Table des dossiers filtrés ---
    if not f.empty:
        st.dataframe(f.reset_index(drop=True))
    else:
        st.info("Aucun dossier ne correspond à la sélection actuelle.")


# =========================
# VISA APP — PARTIE 3/5
# =========================

with tab_clients:
    st.subheader("👥 Clients — Créer / Modifier / Supprimer")

    live_raw = read_sheet(clients_path, sheet_choice, normalize=False)
    live     = read_sheet(clients_path, sheet_choice, normalize=True)

    left, right = st.columns([1,1])

    # ---------- Sélection ----------
    with left:
        st.markdown("### 🔎 Rechercher / Sélectionner")
        if live.empty:
            st.caption("Aucun client pour le moment.")
            sel_idx = None
            sel_row = None
        else:
            names = live["Nom"].fillna("").astype(str).tolist()
            ids   = live["ID_Client"].fillna("").astype(str).tolist()
            labels = [f"{n}  —  {i}" for n, i in zip(names, ids)]
            sel_idx = st.selectbox(
                "Client existant",
                options=list(range(len(labels))),
                format_func=lambda i: labels[i] if i is not None and i < len(labels) else "",
                key=f"cli_sel_idx_{sheet_choice}"
            )
            sel_row = live.iloc[sel_idx] if live.shape[0] and sel_idx is not None else None

    # ---------- Création ----------
    with right:
        st.markdown("### ➕ Nouveau client")
        new_name = st.text_input("Nom", key=f"new_nom_{sheet_choice}")
        new_date = st.date_input("Date création", value=date.today(), key=f"new_date_{sheet_choice}")

        visa_codes = sorted(df_visa["VisaCode"].dropna().unique().tolist())
        new_visa = st.selectbox("Visa (code)", options=[""]+visa_codes, index=0, key=f"new_visa_{sheet_choice}")
        new_cat  = new_visa

        new_hono = st.number_input(HONO, min_value=0.0, step=10.0, format="%.2f", key=f"new_hono_{sheet_choice}")
        new_autr = st.number_input(AUTRE, min_value=0.0, step=10.0, format="%.2f", key=f"new_autre_{sheet_choice}")

        if st.button("💾 Créer ce client", key=f"btn_create_client_{sheet_choice}"):
            if not new_name:
                st.warning("Renseigne le **Nom**.")
            elif not new_visa:
                st.warning("Sélectionne un **Visa**.")
            else:
                base = read_sheet(clients_path, sheet_choice, normalize=True).copy()
                next_dos = next_dossier_number(base)
                client_id = _make_client_id_from_row({"Nom": new_name, "Date": new_date})
                i = 0; origin = client_id
                while (base["ID_Client"].astype(str) == client_id).any():
                    i += 1; client_id = f"{origin}-{i}"

                new_row = {
                    DOSSIER_COL: next_dos,
                    "ID_Client": client_id,
                    "Nom": new_name,
                    "Date": pd.to_datetime(new_date).date(),
                    "Mois": f"{new_date.month:02d}",
                    "Catégorie": new_cat,
                    "Visa": _visa_code_only(new_visa),
                    HONO: float(new_hono), AUTRE: float(new_autr),
                    TOTAL: float(new_hono) + float(new_autr),
                    "Payé": 0.0, "Reste": float(new_hono) + float(new_autr),
                    PAY_JSON: "[]", ESC_TR: 0.0, ESC_JR: "[]",
                    S_ENVOYE: False, S_APPROUVE: False, S_RFE: False, S_REFUSE: False, S_ANNULE: False,
                    D_ENVOYE: "", D_APPROUVE: "", D_RFE: "", D_REFUSE: "", D_ANNULE: ""
                }

                base_raw = read_sheet(clients_path, sheet_choice, normalize=False).copy()
                base_raw = pd.concat([base_raw, pd.DataFrame([new_row])], ignore_index=True)
                base_raw = normalize_clients(base_raw)
                write_sheet_inplace(clients_path, sheet_choice, base_raw)
                st.success("Client créé et sauvegardé.")
                st.rerun()

    st.markdown("---")
    st.markdown("### ✏️ Modifier / Supprimer / Paiements")
    if sel_row is None:
        st.info("Sélectionne un client à gauche ou crée un nouveau client.")
        st.stop()

    idx = sel_idx
    ed = live.loc[idx].to_dict()

    # ---------- Formulaire modification ----------
    c1, c2, c3 = st.columns(3)
    with c1:
        ed_nom  = st.text_input("Nom", value=_safe_str(ed.get("Nom","")), key=f"ed_nom_{idx}_{sheet_choice}")
        ed_date_val = pd.to_datetime(ed.get("Date")).date() if pd.notna(ed.get("Date")) else date.today()
        ed_date = st.date_input("Date création", value=ed_date_val, key=f"ed_date_{idx}_{sheet_choice}")
    with c2:
        visa_codes = sorted(df_visa["VisaCode"].dropna().unique().tolist())
        current_code = _visa_code_only(ed.get("Visa",""))
        ed_visa = st.selectbox("Visa (code)", options=[""]+visa_codes,
                               index=(visa_codes.index(current_code)+1 if current_code in visa_codes else 0),
                               key=f"ed_visa_{idx}_{sheet_choice}")
        ed_cat  = ed_visa
    with c3:
        ed_hono = st.number_input(HONO, min_value=0.0, value=float(ed.get(HONO,0.0)), step=10.0, format="%.2f", key=f"ed_hono_{idx}_{sheet_choice}")
        ed_autr = st.number_input(AUTRE, min_value=0.0, value=float(ed.get(AUTRE,0.0)), step=10.0, format="%.2f", key=f"ed_autr_{idx}_{sheet_choice}")

    # ---------- Statuts ----------
    s1,s2,s3,s4,s5 = st.columns(5)
    with s1:
        b_env = st.checkbox(S_ENVOYE, value=bool(ed.get(S_ENVOYE, False)), key=f"ed_env_{idx}_{sheet_choice}")
        d_env = st.date_input(D_ENVOYE, value=(pd.to_datetime(ed.get(D_ENVOYE)).date() if _safe_str(ed.get(D_ENVOYE)) else date.today()),
                              key=f"ed_denvoi_{idx}_{sheet_choice}") if b_env else ""
    with s2:
        b_app = st.checkbox(S_APPROUVE, value=bool(ed.get(S_APPROUVE, False)), key=f"ed_app_{idx}_{sheet_choice}")
        d_app = st.date_input(D_APPROUVE, value=(pd.to_datetime(ed.get(D_APPROUVE)).date() if _safe_str(ed.get(D_APPROUVE)) else date.today()),
                              key=f"ed_dappr_{idx}_{sheet_choice}") if b_app else ""
    with s3:
        b_rfe = st.checkbox(S_RFE, value=bool(ed.get(S_RFE, False)), key=f"ed_rfe_{idx}_{sheet_choice}")
        d_rfe = st.date_input(D_RFE, value=(pd.to_datetime(ed.get(D_RFE)).date() if _safe_str(ed.get(D_RFE)) else date.today()),
                              key=f"ed_drfe_{idx}_{sheet_choice}") if b_rfe else ""
    with s4:
        b_ref = st.checkbox(S_REFUSE, value=bool(ed.get(S_REFUSE, False)), key=f"ed_ref_{idx}_{sheet_choice}")
        d_ref = st.date_input(D_REFUSE, value=(pd.to_datetime(ed.get(D_REFUSE)).date() if _safe_str(ed.get(D_REFUSE)) else date.today()),
                              key=f"ed_dref_{idx}_{sheet_choice}") if b_ref else ""
    with s5:
        b_ann = st.checkbox(S_ANNULE, value=bool(ed.get(S_ANNULE, False)), key=f"ed_ann_{idx}_{sheet_choice}")
        d_ann = st.date_input(D_ANNULE, value=(pd.to_datetime(ed.get(D_ANNULE)).date() if _safe_str(ed.get(D_ANNULE)) else date.today()),
                              key=f"ed_dann_{idx}_{sheet_choice}") if b_ann else ""

    # ---------- Paiements ----------
    st.markdown("#### 💳 Paiements (multi-acomptes)")
    pay_modes = ["CB","Chèque","Cash","Virement","Venmo"]
    pcol1, pcol2, pcol3, pcol4 = st.columns([1,1,1,2])
    with pcol1:
        p_date = st.date_input("Date paiement", value=date.today(), key=f"p_date_{idx}_{sheet_choice}")
    with pcol2:
        p_mode = st.selectbox("Mode", pay_modes, index=0, key=f"p_mode_{idx}_{sheet_choice}")
    with pcol3:
        p_amt  = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"p_amt_{idx}_{sheet_choice}")
    with pcol4:
        if st.button("➕ Ajouter ce paiement", key=f"btn_addpay_{idx}_{sheet_choice}"):
            base = read_sheet(clients_path, sheet_choice, normalize=True).copy()
            reste_curr = float(base.loc[idx, "Reste"])
            if float(p_amt) <= 0:
                st.warning("Le montant doit être > 0.")
            elif reste_curr <= 0:
                st.info("Dossier déjà soldé.")
            else:
                base_raw = read_sheet(clients_path, sheet_choice, normalize=False).copy()
                row = base_raw.loc[idx].to_dict()
                plist = _parse_json_list(row.get(PAY_JSON,""))
                plist.append({"date": str(p_date), "mode": p_mode, "amount": float(p_amt)})
                row[PAY_JSON] = json.dumps(plist, ensure_ascii=False)
                base_raw.loc[idx] = row
                base_raw = normalize_clients(base_raw)
                write_sheet_inplace(clients_path, sheet_choice, base_raw)
                st.success("Paiement ajouté et sauvegardé.")
                st.rerun()

    try:
        plist = _parse_json_list(live_raw.loc[sel_idx].get(PAY_JSON,"") if live_raw.shape[0] else "[]")
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

    # ---------- Enregistrer / Supprimer ----------
    ac1, ac2 = st.columns([1,1])

    if ac1.button("💾 Sauvegarder les modifications", key=f"btn_save_{idx}_{sheet_choice}"):
        base_raw = read_sheet(clients_path, sheet_choice, normalize=False).copy()
        if idx >= len(base_raw):
            st.error("Ligne introuvable."); st.stop()

        row = base_raw.loc[idx].to_dict()
        row["Nom"]  = ed_nom
        row["Date"] = pd.to_datetime(ed_date).date()
        row["Mois"] = f"{ed_date.month:02d}"
        row["Catégorie"] = ed_cat
        row["Visa"] = _visa_code_only(ed_visa)
        row[HONO] = float(ed_hono)
        row[AUTRE]= float(ed_autr)
        row[TOTAL]= float(ed_hono) + float(ed_autr)
        row[S_ENVOYE]= bool(b_env); row[D_ENVOYE]= str(d_env) if b_env else ""
        row[S_APPROUVE]= bool(b_app); row[D_APPROUVE]= str(d_app) if b_app else ""
        row[S_RFE]= bool(b_rfe); row[D_RFE]= str(d_rfe) if b_rfe else ""
        row[S_REFUSE]= bool(b_ref); row[D_REFUSE]= str(d_ref) if b_ref else ""
        row[S_ANNULE]= bool(b_ann); row[D_ANNULE]= str(d_ann) if b_ann else ""

        base_raw.loc[idx] = row
        base_raw = normalize_clients(base_raw)
        write_sheet_inplace(clients_path, sheet_choice, base_raw)
        st.success("Modifications sauvegardées.")
        st.rerun()

    if a2 := ac2.button("🗑️ Supprimer ce client", key=f"btn_del_{idx}_{sheet_choice}"):
        base_raw = read_sheet(clients_path, sheet_choice, normalize=False).copy()
        if 0 <= idx < len(base_raw):
            base_raw = base_raw.drop(index=idx).reset_index(drop=True)
            base_raw = normalize_clients(base_raw)
            write_sheet_inplace(clients_path, sheet_choice, base_raw)
            st.success("Client supprimé.")
            st.rerun()
        else:
            st.error("Ligne introuvable.")



# =========================
# VISA APP — PARTIE 4/5
# =========================
try:
    import altair as alt
except Exception:
    alt = None

with tab_analyses:
    st.subheader("📈 Analyses — Volumes & Financier")
    dfA_raw = read_sheet(clients_path, sheet_choice, normalize=False)
    dfA = normalize_clients(dfA_raw)
    if dfA.empty:
        st.info("Aucune donnée pour analyser."); st.stop()

    # Filtres contextuels identiques au Dashboard
    selA = build_checkbox_filters_grouped(df_visa, keyprefix=f"flt_anal_{sheet_choice}", as_toggle=False)
    fA = filter_clients_by_ref(dfA, selA)

    # Filtres date
    cR1, cR2, cR3 = st.columns(3)
    yearsA  = sorted({d.year for d in fA["Date"] if pd.notna(d)}) if "Date" in fA.columns else []
    monthsA = sorted([m for m in fA["Mois"].dropna().unique()]) if "Mois" in fA.columns else []
    sel_years  = cR1.multiselect("Année", yearsA, default=[], key=f"anal_years_{sheet_choice}")
    sel_months = cR2.multiselect("Mois (MM)", monthsA, default=[], key=f"anal_months_{sheet_choice}")
    include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True, key=f"anal_na_{sheet_choice}")

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

    # KPI (robustes)
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(fA)}")
    k2.metric("Honoraires", _fmt_money_us(_safe_num_series(fA, HONO).sum()))
    k3.metric("Payé",      _fmt_money_us(_safe_num_series(fA, "Payé").sum()))
    k4.metric("Solde",     _fmt_money_us(_safe_num_series(fA, "Reste").sum()))
    st.markdown('</div>', unsafe_allow_html=True)

    # Volumes par période
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
                ).properties(height=260),
                use_container_width=True
            )
        except Exception:
            st.dataframe(df_vol, use_container_width=True)
    else:
        st.dataframe(df_vol, use_container_width=True)

    st.divider()

    # Comparaisons YoY & par mois
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in fA.columns:
            fA[col] = _safe_num_series(fA, col)  # s'assure du type numérique

    st.markdown("## 🔁 Comparaisons (YoY & Mois)")
    by_year = fA.dropna(subset=["Année"]).groupby("Année").agg(
        Dossiers=("Nom","count"),
        Honoraires=(HONO,"sum"),
        Autres=(AUTRE,"sum"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index().sort_values("Année")

    c1, c2 = st.columns(2)
    c1.dataframe(by_year, use_container_width=True)

    by_year_month = fA.dropna(subset=["Année","MoisNum"]).groupby(["Année","MoisNum"]).agg(
        Dossiers=("Nom","count"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index()

    c2.dataframe(by_year_month, use_container_width=True)

    st.divider()
    st.markdown("### 🔎 Détails (clients)")
    details = fA.copy()
    details["Periode"] = details["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")
    details_cols = [c for c in [
        "Periode", DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", "Date",
        HONO, AUTRE, TOTAL, "Payé", "Reste", "Année", "MoisNum"
    ] if c in details.columns]
    for col in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if col in details.columns:
            details[col] = pd.to_numeric(details[col], errors="coerce").fillna(0.0).map(_fmt_money_us)
    if "Date" in details.columns: details["Date"] = details["Date"].astype(str)
    st.dataframe(details[details_cols].sort_values(["Année","MoisNum","Catégorie","Nom"]), use_container_width=True)


# =========================
# VISA APP — PARTIE 5/5
# =========================

with tab_escrow:
    st.subheader("🏦 ESCROW — suivi & transferts")

    dfE = read_sheet(clients_path, sheet_choice, normalize=True)
    if dfE.empty:
        st.info("Aucun dossier."); st.stop()

    # Dispo ESCROW = min(Payé, Honoraires) - ESCROW transféré
    dfE["Dispo ESCROW"] = (dfE["Payé"].clip(upper=dfE[HONO]) - dfE[ESC_TR]).clip(lower=0.0)

    # Alerte
    to_claim = dfE[(dfE["Dispo ESCROW"] > 0.0)]
    if len(to_claim):
        st.warning(f"⚠️ {len(to_claim)} dossier(s) ont de l’ESCROW à transférer.")
        show_cols = [c for c in [DOSSIER_COL,"ID_Client","Nom","Visa",HONO,"Payé","Dispo ESCROW"] if c in to_claim.columns]
        tmp = to_claim[show_cols].copy()
        for col in [HONO,"Payé","Dispo ESCROW"]:
            if col in tmp.columns:
                tmp[col] = pd.to_numeric(tmp[col], errors="coerce").fillna(0.0).map(_fmt_money_us)
        st.dataframe(tmp, use_container_width=True)

    st.divider()
    st.markdown("### 🔁 Marquer un transfert d’ESCROW → Compte ordinaire")

    df_with_dispo = dfE[dfE["Dispo ESCROW"] > 0.0].reset_index(drop=True)
    if df_with_dispo.empty:
        st.caption("Aucun dossier avec ESCROW disponible pour transfert.")
    else:
        for i, r in df_with_dispo.iterrows():
            dispo = float(r["Dispo ESCROW"])
            header = f"{r.get(DOSSIER_COL,'')} — {r.get('Nom','')} — Visa {r.get('Visa','')} — Dispo: {_fmt_money_us(dispo)}"
            with st.expander(header, expanded=False):
                amt = st.number_input(
                    "Montant à marquer comme transféré (US $)",
                    min_value=0.0, value=float(dispo),
                    step=10.0, format="%.2f",
                    key=f"esc_amt_{sheet_choice}_{i}"
                )
                note = st.text_input("Note (optionnelle)", key=f"esc_note_{sheet_choice}_{i}")

                if st.button("💾 Enregistrer le transfert", key=f"esc_save_{sheet_choice}_{i}"):
                    base_raw = read_sheet(clients_path, sheet_choice, normalize=False).copy()

                    # On retrouve la ligne via ID_Client (plus robuste)
                    idc = str(r.get("ID_Client",""))
                    if "ID_Client" in base_raw.columns and idc:
                        try:
                            real_idx = base_raw.index[base_raw["ID_Client"].astype(str) == idc][0]
                        except Exception:
                            real_idx = None
                    else:
                        real_idx = None

                    if real_idx is None:
                        st.error("Ligne introuvable pour ce dossier.")
                    else:
                        row = base_raw.loc[real_idx].to_dict()
                        journal = _parse_json_list(row.get(ESC_JR,""))
                        journal.append({
                            "ts": datetime.now().isoformat(timespec="seconds"),
                            "amount": float(amt),
                            "note": _safe_str(note)
                        })
                        row[ESC_JR] = json.dumps(journal, ensure_ascii=False)
                        try:
                            curr_tr = float(row.get(ESC_TR, 0.0) or 0.0)
                        except Exception:
                            curr_tr = 0.0
                        row[ESC_TR] = curr_tr + float(amt)

                        base_raw.loc[real_idx] = row
                        base_raw = normalize_clients(base_raw)
                        write_sheet_inplace(clients_path, sheet_choice, base_raw)
                        st.success("Transfert enregistré.")
                        st.rerun()

    st.divider()
    st.markdown("### 📒 Journal ESCROW (tous dossiers)")
    rows = []
    base_for_journal = read_sheet(clients_path, sheet_choice, normalize=False).copy()
    for j, r in base_for_journal.iterrows():
        jr = _parse_json_list(r.get(ESC_JR, ""))
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
        jdf = pd.DataFrame(rows)
        try:
            jdf["Horodatage_dt"] = pd.to_datetime(jdf["Horodatage"], errors="coerce")
            jdf = jdf.sort_values("Horodatage_dt").drop(columns=["Horodatage_dt"])
        except Exception:
            jdf = jdf.sort_values("Horodatage")
        jdf["Montant"] = jdf["Montant"].map(_fmt_money_us)
        st.dataframe(jdf, use_container_width=True)
    else:
        st.caption("Pas encore de transferts enregistrés.")
