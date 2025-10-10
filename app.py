from __future__ import annotations

# ============================================
# VISA MANAGER — APP COMPLETE
# - Fichier Clients : donnees_visa_clients1_adapte.xlsx (feuille "Clients")
# - Fichier Visa    : donnees_visa_clients1.xlsx        (feuille "Visa")
# ============================================

import json
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import streamlit as st
from io import BytesIO
import openpyxl

# Colonnes minimales pour la feuille Clients
CLIENTS_COLUMNS = [
    "Dossier N","ID_Client","Nom","Date","Mois",
    "Categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
    "Paiements","ESCROW transféré (US $)","Journal ESCROW",
    "Dossier envoyé","Date envoyé",
    "Dossier approuvé","Date approuvé",
    "RFE","Date RFE",
    "Dossier refusé","Date refusé",
    "Dossier annulé","Date annulé",
]

# Colonnes minimales pour la feuille Visa (sans accents)
VISA_COLUMNS = [
    "Categorie","Sous-categorie 1","COS","EOS"  # tu peux en rajouter (Premium, etc.)
]

def _create_clients_template(path: str|Path, sheet_name: str="Clients") -> None:
    df = pd.DataFrame(columns=CLIENTS_COLUMNS)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        df.to_excel(wr, sheet_name=sheet_name, index=False)

def _create_visa_template(path: str|Path, sheet_name: str="Visa") -> None:
    # Exemple de lignes : adapte si besoin
    rows = [
        {"Categorie":"B-1","Sous-categorie 1":"Affaires","COS":"x","EOS":""},
        {"Categorie":"B-2","Sous-categorie 1":"Tourisme","COS":"x","EOS":"x"},
        {"Categorie":"F-1","Sous-categorie 1":"Etudiant","COS":"x","EOS":"x"},
    ]
    df = pd.DataFrame(rows, columns=VISA_COLUMNS)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        df.to_excel(wr, sheet_name=sheet_name, index=False)

def ensure_files_exist(clients_path: str|Path, visa_path: str|Path) -> None:
    clients_path = Path(clients_path)
    visa_path = Path(visa_path)
    if not clients_path.exists():
        _create_clients_template(clients_path)
    if not visa_path.exists():
        _create_visa_template(visa_path)

def safe_excel_first_sheet(path: str|Path, preferred: str|None=None) -> str:
    """Retourne le nom de feuille à lire : `preferred` si présent, sinon la 1re."""
    with pd.ExcelFile(path) as xls:
        sheets = xls.sheet_names
    if preferred and preferred in sheets:
        return preferred
    return sheets[0] if sheets else "Clients"  # fallback

# ------------------ CONFIG -------------------
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")
st.caption("Gestion des clients, visas, acomptes, ESCROW et analyses — Excel en source unique")

# ------------------ CONSTANTES ----------------
DEFAULT_CLIENTS = "donnees_visa_clients1_adapte.xlsx"
DEFAULT_VISA    = "donnees_visa_clients1.xlsx"

DOSSIER_COL = "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"
PAY_JSON = "Paiements"  # JSON de paiements (liste de dicts {date, mode, amount})

VISA_LEVELS = [
    "Categorie",
    "Sous-categorie 1", "Sous-categorie 2", "Sous-categorie 3", "Sous-categorie 4",
    "Sous-categorie 5", "Sous-categorie 6", "Sous-categorie 7", "Sous-categorie 8",
]

# ------------------ HELPERS -------------------
def _safe_str(x) -> str:
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)

def _to_num(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.replace(r"[^\d,.\-]", "", regex=True).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df), index=df.index, dtype=float)
    v = df[col]
    if pd.api.types.is_numeric_dtype(v):
        return v.fillna(0.0).astype(float)
    return _to_num(v)

def _fmt_money_us(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _uniquify_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(map(str, df.columns))
    seen = {}
    new_cols = []
    for c in cols:
        if c not in seen:
            seen[c] = 1
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
    out = df.copy()
    out.columns = new_cols
    return out

# ------------------ VISA REF ------------------
def _ensure_visa_columns(df_visa: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df_visa, pd.DataFrame) or df_visa.empty:
        return pd.DataFrame(columns=VISA_LEVELS)
    df = df_visa.copy()
    for c in VISA_LEVELS:
        if c not in df.columns:
            df[c] = ""
        else:
            df[c] = df[c].fillna("").astype(str)
    return df[VISA_LEVELS]

@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str = "Visa") -> dict[str, list[str]]:
    """
    Lit la feuille Visa et renvoie par Categorie : la liste d'intitulés Visa
    au format '<Sous-categorie> <NomCaseCochee>'.
    - Colonnes attendues : 'Categorie', 'Sous-categorie 1' (ou vide), puis plusieurs colonnes "cases".
    - Une case est considérée cochée si valeur ∈ {1, x, ✓, true, oui, yes, y, o} (case insensitive).
    """
    import numpy as np
    try:
        dfv = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    except Exception:
        return {}
    dfv = dfv.copy()
    dfv = _uniquify_columns(dfv)
    dfv.columns = dfv.columns.map(str).str.strip()

    # détection stricte sans accents :
    cat_col = "Categorie" if "Categorie" in dfv.columns else None
    # sous-cat : accepte différents libellés sans accents
    sub_col = None
    for c in ["Sous-categorie 1", "Sous-categorie", "Sous-categories 1", "Sous-categories"]:
        if c in dfv.columns:
            sub_col = c
            break
    if not cat_col:
        return {}
    if not sub_col:
        dfv["_Sous_"] = ""
        sub_col = "_Sous_"

    checkbox_cols = [c for c in dfv.columns if c not in {cat_col, sub_col}]

    def _is_checked(v) -> bool:
        if pd.isna(v): return False
        if isinstance(v, (int, float)): return float(v) != 0.0
        s = str(v).strip().lower()
        return s in {"1", "x", "✓", "true", "vrai", "oui", "yes", "y", "o"}

    out: dict[str, list[str]] = {}
    for _, row in dfv.iterrows():
        cat = _safe_str(row.get(cat_col, "")).strip()
        if not cat:
            continue
        sous = _safe_str(row.get(sub_col, "")).strip()
        for col in checkbox_cols:
            if _is_checked(row.get(col, np.nan)):
                label = f"{sous} {col}".strip()
                out.setdefault(cat, []).append(label)

    # unique trié
    for k, v in out.items():
        out[k] = sorted(set(v))
    return out

# ------------------ CLIENTS I/O ----------------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # colonnes minimales
    for c in [
        DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
        "Categorie", "Visa",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        PAY_JSON,
        "ESCROW transféré (US $)", "Journal ESCROW",
        "Dossier envoyé", "Date envoyé",
        "Dossier approuvé", "Date approuvé",
        "RFE", "Date RFE",
        "Dossier refusé", "Date refusé",
        "Dossier annulé", "Date annulé",
    ]:
        if c not in df.columns:
            df[c] = None

    # dates / mois
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df["Mois"] = df.apply(
        lambda r: f"{pd.to_datetime(r['Date']).month:02d}" if pd.notna(r["Date"]) else (_safe_str(r.get("Mois",""))[:2] or None),
        axis=1,
    )

    # montants
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste", "ESCROW transféré (US $)"]:
        df[c] = _safe_num_series(df, c)
    df[TOTAL] = df[HONO] + df[AUTRE]

    # paiements (JSON)
    def _norm_p(x):
        try:
            j = json.loads(_safe_str(x) or "[]")
            return j if isinstance(j, list) else []
        except Exception:
            return []
    df[PAY_JSON] = df[PAY_JSON].apply(_norm_p)

    # payé = max(col Payé, somme JSON)
    def _sum_json(lst):
        try:
            return float(sum(float(it.get("amount", 0.0) or 0.0) for it in (lst or [])))
        except Exception:
            return 0.0
    js_paid = df[PAY_JSON].apply(_sum_json)
    df["Payé"] = df["Payé"].fillna(0.0).astype(float)
    df["Payé"] = pd.concat([df["Payé"], js_paid], axis=1).max(axis=1)

    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # dérivées pour filtres
    df["_Année_"] = df["Date"].apply(lambda d: d.year if pd.notna(d) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda d: d.month if pd.notna(d) else pd.NA)
    df["_Mois_"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) else pd.NA)

    # booleans
    for c in ["Dossier envoyé", "Dossier approuvé", "RFE", "Dossier refusé", "Dossier annulé"]:
        df[c] = df[c].apply(lambda v: bool(v) if pd.notna(v) else False)

    # journal escrow
    def _norm_j(x):
        try:
            j = json.loads(_safe_str(x) or "[]")
            return j if isinstance(j, list) else []
        except Exception:
            return []
    df["Journal ESCROW"] = df["Journal ESCROW"].apply(_norm_j)

    return _uniquify_columns(df)

@st.cache_data(show_spinner=False)
def read_sheet(xlsx_path: str | Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    return normalize_clients(_uniquify_columns(df))

def write_sheet_inplace(xlsx_path: str | Path, sheet_name: str, df: pd.DataFrame) -> None:
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="w") as wr:
            df.to_excel(wr, sheet_name=sheet_name, index=False)
        return
    # supprimer puis réécrire proprement la feuille
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as wr:
        try:
            book = wr.book
            if sheet_name in book.sheetnames:
                sh = book[sheet_name]; book.remove(sh); book.create_sheet(sheet_name)
        except Exception:
            pass
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
        df.to_excel(wr, sheet_name=sheet_name, index=False)

# ------------------ ID & DOSSIER --------------
def _make_client_id_from_row(r: dict) -> str:
    nom = _safe_str(r.get("Nom","")).strip().replace(" ", "_")
    dt = r.get("Date", date.today())
    if isinstance(dt, str):
        try:
            dt = pd.to_datetime(dt).date()
        except Exception:
            dt = date.today()
    stamp = f"{dt:%Y%m%d}"
    return f"{nom}-{stamp}"

def next_dossier_number(df: pd.DataFrame, start: int = 13057) -> int:
    if DOSSIER_COL in df.columns and pd.to_numeric(df[DOSSIER_COL], errors="coerce").notna().any():
        return int(pd.to_numeric(df[DOSSIER_COL], errors="coerce").max()) + 1
    return int(start)

# ------------------ BARRE LATERALE ------------
st.sidebar.header("📂 Fichiers")
clients_path = st.sidebar.text_input("Fichier Clients (.xlsx)", value=DEFAULT_CLIENTS, key="cli_path")
visa_path    = st.sidebar.text_input("Fichier Visa (.xlsx)",    value=DEFAULT_VISA,    key="visa_path")

# Crée les fichiers modèles si absents
ensure_files_exist(clients_path, visa_path)

# Déterminer une feuille valide
try:
    sheet_choice = safe_excel_first_sheet(clients_path, preferred="Clients")
except Exception as e:
    st.error(f"Impossible d’ouvrir {clients_path} : {e}")
    # on recrée un modèle propre et on retente
    _create_clients_template(clients_path)
    sheet_choice = "Clients"

# Charger les données clients (avec normalisation)
try:
    df_clients = read_sheet(clients_path, sheet_choice)
except FileNotFoundError:
    _create_clients_template(clients_path)
    df_clients = read_sheet(clients_path, "Clients")
except Exception as e:
    st.error(f"Erreur lecture Clients: {e}")
    _create_clients_template(clients_path)
    df_clients = read_sheet(clients_path, "Clients")

# Charger les choix VISA depuis l’onglet Visa
try:
    visa_choices_by_category = parse_visa_sheet(visa_path, sheet_name="Visa")
except FileNotFoundError:
    _create_visa_template(visa_path)
    visa_choices_by_category = parse_visa_sheet(visa_path, sheet_name="Visa")
except Exception as e:
    st.warning(f"Référentiel Visa non lisible ({e}). Utilisation d’un modèle.")
    _create_visa_template(visa_path)
    visa_choices_by_category = parse_visa_sheet(visa_path, sheet_name="Visa")

# Affichage court
with st.sidebar.expander("🔎 Vérif Visa détectés", expanded=False):
    if visa_choices_by_category:
        for cat, vals in visa_choices_by_category.items():
            st.write(f"**{cat}** → {', '.join(vals)}")
    else:
        st.caption("Aucune structure Visa détectée (onglet ‘Visa’ vide).")
# ------------------ TABS ----------------------
tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs([
    "Dashboard", "Clients", "Analyses", "ESCROW"
])

# ================== DASHBOARD =================
with tab_dash:
    st.subheader("📊 Tableau de bord — Synthèse")
    # Filtres hiérarchie (Catégorie + Visa depuis listes)
    cats_all = sorted(df_clients["Categorie"].dropna().astype(str).unique().tolist())
    visas_all = sorted(df_clients["Visa"].dropna().astype(str).unique().tolist())

    c1, c2, c3, c4 = st.columns([1,1,1,2])
    with c1:
        sel_cats = st.multiselect("Catégories", cats_all, default=[])
    with c2:
        sel_visas = st.multiselect("Visa", visas_all, default=[])
    with c3:
        sel_solde = st.selectbox("Solde", ["Tous", "Soldé (Reste = 0)", "Non soldé (Reste > 0)"], index=0)
    with c4:
        q = st.text_input("Recherche (Nom / ID / Dossier / Visa)", "")

    ff = df_clients.copy()
    if sel_cats:
        ff = ff[ff["Categorie"].astype(str).isin(sel_cats)]
    if sel_visas:
        ff = ff[ff["Visa"].astype(str).isin(sel_visas)]
    if sel_solde == "Soldé (Reste = 0)":
        ff = ff[_safe_num_series(ff, "Reste") <= 1e-9]
    elif sel_solde == "Non soldé (Reste > 0)":
        ff = ff[_safe_num_series(ff, "Reste") > 1e-9]
    if q:
        qn = q.lower().strip()
        def _m(r):
            hay = " ".join([
                _safe_str(r.get("Nom","")),
                _safe_str(r.get("ID_Client","")),
                _safe_str(r.get("Categorie","")),
                _safe_str(r.get("Visa","")),
                str(r.get(DOSSIER_COL,""))
            ]).lower()
            return qn in hay
        ff = ff[ff.apply(_m, axis=1)]

    st.info(f"**{len(ff)} dossiers** avec les filtres.")

    # KPI
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(ff)}")
    k2.metric("Honoraires", _fmt_money_us(_safe_num_series(ff,HONO).sum()))
    k3.metric("Payé", _fmt_money_us(_safe_num_series(ff,"Payé").sum()))
    k4.metric("Reste", _fmt_money_us(_safe_num_series(ff,"Reste").sum()))

    st.markdown("---")

    # Tableau
    view = ff.copy()
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c in view.columns:
            view[c] = _safe_num_series(view,c).map(_fmt_money_us)
    if "Date" in view.columns:
        view["Date"] = view["Date"].astype(str)

    show_cols = [c for c in [
        DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Visa", "Date", "Mois",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Dossier envoyé", "Dossier approuvé", "RFE", "Dossier refusé", "Dossier annulé"
    ] if c in view.columns]

    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
    view = view.sort_values(by=sort_cols) if sort_cols else view

    st.dataframe(_uniquify_columns(view[show_cols].reset_index(drop=True)), use_container_width=True)

# ================== CLIENTS (CRUD) ===========
with tab_clients:
    st.subheader("👥 Clients — Créer / Modifier / Supprimer / Paiements")

    live = df_clients.copy()

    # Sélection
    cL, cR = st.columns([1,1])
    with cL:
        st.markdown("### 🔎 Sélection")
        if live.empty:
            st.caption("Aucun client.")
            sel_idx = None
            sel_row = None
        else:
            labels = (live.get("Nom","").astype(str) + " — " + live.get("ID_Client","").astype(str)).fillna("")
            sel_idx = st.selectbox("Client", options=list(live.index),
                                   format_func=lambda i: labels.iloc[i],
                                   key=f"cli_sel")
            sel_row = live.loc[sel_idx] if sel_idx is not None else None

    # Création
    with cR:
        st.markdown("### ➕ Nouveau client")
        new_name = st.text_input("Nom")
        new_date = st.date_input("Date de création", value=date.today())
        cats = sorted(visa_choices_by_category.keys())
        new_cat = st.selectbox("Categorie", options=[""]+cats, index=0)
        visa_opts = visa_choices_by_category.get(new_cat, [])
        new_visa = st.selectbox("Visa (auto depuis onglet Visa)", options=[""]+visa_opts, index=0)

        new_hono = st.number_input(HONO, min_value=0.0, step=10.0, format="%.2f")
        new_autr = st.number_input(AUTRE, min_value=0.0, step=10.0, format="%.2f")

        if st.button("💾 Créer"):
            if not new_name:
                st.warning("Nom obligatoire."); st.stop()
            if not new_cat:
                st.warning("Categorie obligatoire."); st.stop()
            if not new_visa:
                st.warning("Visa obligatoire (dépend de la Categorie)."); st.stop()

            base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
            base_norm = normalize_clients(base_raw.copy())

            dossier = next_dossier_number(base_norm)
            client_id = _make_client_id_from_row({"Nom": new_name, "Date": new_date})
            # collision ID_Client → suffixes -1, -2...
            origin = client_id
            i = 0
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
                "Categorie": new_cat,
                "Visa": new_visa,
                HONO: float(new_hono),
                AUTRE: float(new_autr),
                TOTAL: total,
                "Payé": 0.0,
                "Reste": total,
                PAY_JSON: "[]",
                "ESCROW transféré (US $)": 0.0,
                "Journal ESCROW": "[]",
                "Dossier envoyé": False, "Date envoyé": "",
                "Dossier approuvé": False, "Date approuvé": "",
                "RFE": False, "Date RFE": "",
                "Dossier refusé": False, "Date refusé": "",
                "Dossier annulé": False, "Date annulé": "",
            }
            base_raw = pd.concat([base_raw, pd.DataFrame([new_row])], ignore_index=True)
            base_norm = normalize_clients(base_raw)
            write_sheet_inplace(clients_path, sheet_choice, base_norm)
            st.success("✅ Client créé.")
            st.rerun()

    st.markdown("---")

    if sel_row is None:
        st.stop()

    # Edition
    idx = sel_idx
    ed = sel_row.to_dict()

    e1,e2,e3 = st.columns(3)
    with e1:
        ed_nom = st.text_input("Nom", value=_safe_str(ed.get("Nom","")))
        ed_date = st.date_input("Date de création",
                                value=(pd.to_datetime(ed.get("Date")).date() if pd.notna(ed.get("Date")) else date.today()))
    with e2:
        cats = sorted(visa_choices_by_category.keys())
        curr_cat = _safe_str(ed.get("Categorie",""))
        ed_cat = st.selectbox("Categorie", options=[""]+cats,
                              index=(cats.index(curr_cat)+1 if curr_cat in cats else 0))
        visa_opts = visa_choices_by_category.get(ed_cat, [])
        curr_visa = _safe_str(ed.get("Visa",""))
        ed_visa = st.selectbox("Visa (auto depuis onglet Visa)", options=[""]+visa_opts,
                               index=(visa_opts.index(curr_visa)+1 if curr_visa in visa_opts else 0))
    with e3:
        ed_hono = st.number_input(HONO, min_value=0.0, value=float(ed.get(HONO,0.0)), step=10.0, format="%.2f")
        ed_autr = st.number_input(AUTRE, min_value=0.0, value=float(ed.get(AUTRE,0.0)), step=10.0, format="%.2f")

    st.markdown("#### 🧾 Statuts du dossier")
    s1,s2,s3 = st.columns(3)
    with s1:
        ed_env = st.checkbox("Dossier envoyé", value=bool(ed.get("Dossier envoyé",False)))
        ed_app = st.checkbox("Dossier approuvé", value=bool(ed.get("Dossier approuvé",False)))
    with s2:
        ed_rfe = st.checkbox("RFE", value=bool(ed.get("RFE",False)))
        ed_ref = st.checkbox("Dossier refusé", value=bool(ed.get("Dossier refusé",False)))
    with s3:
        ed_ann = st.checkbox("Dossier annulé", value=bool(ed.get("Dossier annulé",False)))

    st.caption("💡 RFE n’a de sens que si le dossier a un statut (Envoyé/Approuvé/Refusé/Annulé).")

    # Paiements
    st.markdown("### 💳 Paiements (acomptes)")
    p1,p2,p3,p4 = st.columns([1,1,1,2])
    with p1:
        p_date = st.date_input("Date paiement", value=date.today())
    with p2:
        p_mode = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"])
    with p3:
        p_amt = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f")
    with p4:
        if st.button("➕ Ajouter paiement"):
            base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
            # retrouver ligne via ID_Client si possible
            idc = _safe_str(ed.get("ID_Client",""))
            if idc and "ID_Client" in base_raw.columns:
                idxs = base_raw.index[base_raw["ID_Client"].astype(str)==idc].tolist()
                real_idx = idxs[0] if idxs else idx
            else:
                real_idx = idx

            if float(p_amt) <= 0:
                st.warning("Montant > 0 requis.")
            else:
                row = base_raw.loc[real_idx].to_dict()
                try:
                    plist = json.loads(_safe_str(row.get(PAY_JSON,"[]")) or "[]")
                    if not isinstance(plist,list): plist=[]
                except Exception:
                    plist=[]
                plist.append({"date": str(p_date), "mode": p_mode, "amount": float(p_amt)})
                row[PAY_JSON] = json.dumps(plist, ensure_ascii=False)
                base_raw.loc[real_idx] = row
                base_norm = normalize_clients(base_raw.copy())
                write_sheet_inplace(clients_path, sheet_choice, base_norm)
                st.success("Paiement ajouté.")
                st.rerun()

    # Historique
    try:
        hist = json.loads(_safe_str(sel_row.get(PAY_JSON,"[]")) or "[]")
        if not isinstance(hist,list): hist=[]
    except Exception:
        hist=[]
    st.write("**Historique paiements**")
    if hist:
        h = pd.DataFrame(hist)
        if "amount" in h.columns:
            h["amount"] = h["amount"].astype(float).map(_fmt_money_us)
        st.dataframe(h, use_container_width=True)
    else:
        st.caption("Aucun paiement saisi.")

    st.markdown("---")

    a1,a2 = st.columns([1,1])
    if a1.button("💾 Sauvegarder les modifications"):
        base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
        idc = _safe_str(ed.get("ID_Client",""))
        if idc and "ID_Client" in base_raw.columns:
            idxs = base_raw.index[base_raw["ID_Client"].astype(str)==idc].tolist()
            real_idx = idxs[0] if idxs else idx
        else:
            real_idx = idx
        if real_idx is None or not (0 <= real_idx < len(base_raw)):
            st.error("Ligne introuvable."); st.stop()

        row = base_raw.loc[real_idx].to_dict()
        row["Nom"] = ed_nom
        row["Date"] = pd.to_datetime(ed_date).date()
        row["Mois"] = f"{ed_date.month:02d}"
        row["Categorie"] = ed_cat
        row["Visa"] = ed_visa
        row[HONO] = float(ed_hono)
        row[AUTRE] = float(ed_autr)
        row[TOTAL] = float(ed_hono) + float(ed_autr)
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

    if a2.button("🗑️ Supprimer ce client"):
        base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
        idc = _safe_str(ed.get("ID_Client",""))
        if idc and "ID_Client" in base_raw.columns:
            base_raw = base_raw.loc[base_raw["ID_Client"].astype(str)!=idc].reset_index(drop=True)
        else:
            base_raw = base_raw.drop(index=idx).reset_index(drop=True)
        base_norm = normalize_clients(base_raw.copy())
        write_sheet_inplace(clients_path, sheet_choice, base_norm)
        st.success("🗑️ Client supprimé.")
        st.rerun()

# ================== ANALYSES ==================
with tab_analyses:
    st.subheader("📈 Analyses — volumes & financier")

    yearsA  = sorted([int(y) for y in df_clients["_Année_"].dropna().unique()]) if not df_clients.empty else []
    monthsA = [f"{m:02d}" for m in sorted([int(m) for m in df_clients["_MoisNum_"].dropna().unique()])] if not df_clients.empty else []
    catsA   = sorted(df_clients["Categorie"].dropna().astype(str).unique().tolist())
    visasA  = sorted(df_clients["Visa"].dropna().astype(str).unique().tolist())

    c1,c2,c3,c4 = st.columns([1,1,1,2])
    with c1:
        sel_years  = st.multiselect("Année", yearsA, default=[])
    with c2:
        sel_months = st.multiselect("Mois (MM)", monthsA, default=[])
    with c3:
        sel_cats   = st.multiselect("Catégories", catsA, default=[])
    with c4:
        sel_visas  = st.multiselect("Visa", visasA, default=[])

    ff = df_clients.copy()
    if sel_years:
        ff = ff[ff["_Année_"].isin(sel_years)]
    if sel_months:
        ff = ff[ff["_Mois_"].astype(str).isin(sel_months)]
    if sel_cats:
        ff = ff[ff["Categorie"].astype(str).isin(sel_cats)]
    if sel_visas:
        ff = ff[ff["Visa"].astype(str).isin(sel_visas)]

    st.info(f"{len(ff)} dossiers dans le périmètre.")

    # KPI
    sk1,sk2,sk3,sk4 = st.columns(4)
    sk1.metric("Dossiers", f"{len(ff)}")
    sk2.metric("Honoraires", _fmt_money_us(_safe_num_series(ff,HONO).sum()))
    sk3.metric("Payé", _fmt_money_us(_safe_num_series(ff,"Payé").sum()))
    sk4.metric("Reste", _fmt_money_us(_safe_num_series(ff,"Reste").sum()))

    st.markdown("---")

    # par année
    st.markdown("#### 📆 Année → synthèse")
    if not ff.empty and ff["_Année_"].notna().any():
        def _sum_col(df_loc, col): return _safe_num_series(df_loc,col).sum()
        grpY = ff.groupby("_Année_", dropna=True).apply(
            lambda g: pd.Series({
                "Dossiers": int(g.shape[0]),
                "Honoraires": _sum_col(g,HONO),
                "Paye": _sum_col(g,"Payé"),
                "Reste": _sum_col(g,"Reste")
            })
        ).reset_index().rename(columns={"_Année_":"Année"}).sort_values("Année")
        st.dataframe(_uniquify_columns(grpY), use_container_width=True)
    else:
        st.caption("Aucune année exploitable.")

    st.markdown("#### 🗓️ Mois (toutes années)")
    if not ff.empty and ff["_Mois_"].notna().any():
        def _sum_col(df_loc, col): return _safe_num_series(df_loc,col).sum()
        grpM = ff.groupby("_Mois_", dropna=True).apply(
            lambda g: pd.Series({
                "Dossiers": int(g.shape[0]),
                "Honoraires": _sum_col(g,HONO),
                "Paye": _sum_col(g,"Payé"),
                "Reste": _sum_col(g,"Reste")
            })
        ).reset_index().rename(columns={"_Mois_":"Mois"}).sort_values("Mois")
        st.dataframe(_uniquify_columns(grpM), use_container_width=True)
    else:
        st.caption("Aucun mois exploitable.")

    st.markdown("#### 📋 Détails filtrés")
    det = ff.copy()
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c in det.columns:
            det[c] = _safe_num_series(det,c).map(_fmt_money_us)
    if "Date" in det.columns:
        det["Date"] = det["Date"].astype(str)
    show_cols = [c for c in [
        DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Visa", "Date", "Mois",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Dossier envoyé", "Dossier approuvé", "RFE", "Dossier refusé", "Dossier annulé"
    ] if c in det.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
    det = det.sort_values(by=sort_cols) if sort_cols else det
    st.dataframe(_uniquify_columns(det[show_cols].reset_index(drop=True)), use_container_width=True)

# ================== ESCROW ====================
with tab_escrow:
    st.subheader("🏦 ESCROW — dépôts, transferts & alertes")

    # construire métriques escrow par ligne
    rows = []
    for _, r in df_clients.iterrows():
        hono = float(r.get(HONO,0.0) or 0.0)
        # payé = max(col Payé, somme JSON)
        try:
            plist = json.loads(_safe_str(r.get(PAY_JSON,"[]")) or "[]"); 
            if not isinstance(plist,list): plist=[]
        except Exception:
            plist=[]
        js_sum = sum(float(it.get("amount",0.0) or 0.0) for it in plist)
        pay = max(float(r.get("Payé",0.0) or 0.0), js_sum)
        transf = float(r.get("ESCROW transféré (US $)",0.0) or 0.0)
        dispo = max(min(pay, hono) - transf, 0.0)
        rows.append({
            "idx": _,
            DOSSIER_COL: r.get(DOSSIER_COL,""),
            "ID_Client": r.get("ID_Client",""),
            "Nom": r.get("Nom",""),
            "Categorie": r.get("Categorie",""),
            "Visa": r.get("Visa",""),
            "Dossier envoyé": bool(r.get("Dossier envoyé",False)),
            HONO: hono,
            "Payé_calc": pay,
            "ESCROW transféré (US $)": transf,
            "ESCROW dispo": dispo,
            "Journal ESCROW": _safe_str(r.get("Journal ESCROW","[]"))
        })
    jdf = pd.DataFrame(rows)

    c1,c2,c3,c4 = st.columns([1,1,1,2])
    with c1:
        only_dispo = st.toggle("Uniquement ESCROW disponible", value=True)
    with c2:
        only_sent  = st.toggle("Uniquement dossiers envoyés", value=False)
    with c3:
        order_dispo = st.toggle("Trier par dispo", value=True)
    with c4:
        q = st.text_input("Recherche (Nom/ID/Dossier/Visa)", "")

    if only_dispo:
        jdf = jdf[jdf["ESCROW dispo"] > 0.0]
    if only_sent:
        jdf = jdf[jdf["Dossier envoyé"]==True]
    if q:
        qn = q.lower().strip()
        def _m(row):
            hay = " ".join([
                _safe_str(row.get("Nom","")), _safe_str(row.get("ID_Client","")),
                str(row.get(DOSSIER_COL,"")), _safe_str(row.get("Visa","")), _safe_str(row.get("Categorie","")),
            ]).lower()
            return qn in hay
        jdf = jdf[jdf.apply(_m, axis=1)]
    if order_dispo and not jdf.empty:
        jdf = jdf.sort_values("ESCROW dispo", ascending=False)

    # KPI
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.1rem}
    .small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(jdf)}")
    k2.metric("Honoraires périmètre", _fmt_money_us(float(jdf[HONO].sum()) if not jdf.empty else 0.0))
    k3.metric("ESCROW transféré", _fmt_money_us(float(jdf["ESCROW transféré (US $)"].sum()) if not jdf.empty else 0.0))
    k4.metric("ESCROW dispo", _fmt_money_us(float(jdf["ESCROW dispo"].sum()) if not jdf.empty else 0.0))
    st.markdown('</div>', unsafe_allow_html=True)

    if jdf.empty:
        st.info("Aucun dossier à afficher avec les filtres.")
        st.stop()

    show = jdf.copy()
    show[HONO] = show[HONO].map(_fmt_money_us)
    show["Payé_calc"] = show["Payé_calc"].map(_fmt_money_us)
    show["ESCROW transféré (US $)"] = show["ESCROW transféré (US $)"].map(_fmt_money_us)
    show["ESCROW dispo"] = show["ESCROW dispo"].map(_fmt_money_us)
    st.dataframe(
        show[[DOSSIER_COL,"ID_Client","Nom","Categorie","Visa",HONO,"Payé_calc","ESCROW transféré (US $)","ESCROW dispo","Dossier envoyé"]]
        .reset_index(drop=True),
        use_container_width=True
    )

    st.markdown("### ↗️ Enregistrer un transfert ESCROW")
    st.caption("L’ESCROW disponible = min(Payé, Honoraires) − Transféré.")

    for _, r in jdf.iterrows():
        st.markdown(f"**{r['Nom']} — {r['ID_Client']} — Dossier {r[DOSSIER_COL]}**")
        cA,cB,cC,cD = st.columns([1,1,1,2])
        with cA: st.write("Disponible :", _fmt_money_us(float(r["ESCROW dispo"])))
        with cB: t_date = st.date_input("Date transfert", value=date.today(), key=f"esc_dt_{r['ID_Client']}")
        with cC: amt = st.number_input("Montant (US $)", min_value=0.0, value=float(r["ESCROW dispo"]), step=10.0, format="%.2f", key=f"esc_amt_{r['ID_Client']}")
        with cD: note = st.text_input("Note (optionnel)", value="", key=f"esc_note_{r['ID_Client']}")

        ok = float(r["ESCROW dispo"]) > 0 and float(amt) > 0 and float(amt) <= float(r["ESCROW dispo"]) + 1e-9
        if st.button("💸 Enregistrer transfert", key=f"esc_btn_{r['ID_Client']}"):
            if not ok:
                st.warning("Montant invalide.")
                st.stop()

            base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
            # recherche par ID_Client prioritaire
            idc = _safe_str(r.get("ID_Client",""))
            if idc and "ID_Client" in base_raw.columns:
                idxs = base_raw.index[base_raw["ID_Client"].astype(str) == idc].tolist()
                real_idx = idxs[0] if idxs else int(r["idx"])
            else:
                real_idx = int(r["idx"])

            if real_idx is None or not (0 <= real_idx < len(base_raw)):
                st.error("Ligne introuvable."); st.stop()

            row = base_raw.loc[real_idx].to_dict()
            curr_tr = float(row.get("ESCROW transféré (US $)",0.0) or 0.0)
            row["ESCROW transféré (US $)"] = curr_tr + float(amt)

            try:
                jlog = json.loads(_safe_str(row.get("Journal ESCROW","[]")) or "[]")
                if not isinstance(jlog,list): jlog=[]
            except Exception:
                jlog=[]
            jlog.append({"ts": datetime.now().isoformat(timespec="seconds"), "date": str(t_date), "amount": float(amt), "note": note})
            row["Journal ESCROW"] = json.dumps(jlog, ensure_ascii=False)

            base_raw.loc[real_idx] = row
            base_norm = normalize_clients(base_raw.copy())
            write_sheet_inplace(clients_path, sheet_choice, base_norm)
            st.success("✅ Transfert enregistré.")
            st.rerun()

    st.markdown("---")
    st.markdown("### 🚨 Alertes — dossiers envoyés non transférés")
    alert = jdf[(jdf["Dossier envoyé"]==True) & (jdf["ESCROW dispo"]>0.0)] if "Dossier envoyé" in jdf.columns else pd.DataFrame()
    if alert.empty:
        st.success("Aucune alerte : tous les dossiers envoyés ont leur ESCROW transféré ✅.")
    else:
        a = alert.copy()
        a["ESCROW dispo"] = a["ESCROW dispo"].map(_fmt_money_us)
        st.dataframe(a[[DOSSIER_COL,"ID_Client","Nom","Categorie","Visa","ESCROW dispo"]], use_container_width=True)