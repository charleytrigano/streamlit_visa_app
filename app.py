# ===============================
# ======   PARTIE 1 / 2    ======
# ===============================
from __future__ import annotations

import os, json, zipfile, re
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple, Any

import pandas as pd
import streamlit as st

# -------------------------------------------------
# Constantes / noms de colonnes normalisés (FR/US$)
# -------------------------------------------------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

DOSSIER_COL = "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"

# -------------------------------------------------
# CSS pour compacter les KPI
# -------------------------------------------------
st.set_page_config(page_title="Visa Manager", layout="wide")
st.markdown("""
<style>
.small-metrics .stMetric { padding: 0.25rem 0.5rem !important; }
.small-metrics .st-emotion-cache-1xarl3l { font-size: 0.8rem !important; }
.small-metrics .st-emotion-cache-ocqkz7 { font-size: 0.9rem !important; }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# Outils sûrs de conversion / formatage
# -------------------------------------------------
def _safe_str(x: Any) -> str:
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)

def _to_float(x: Any) -> float:
    s = _safe_str(x).strip()
    if not s:
        return 0.0
    # enlever tout sauf chiffres, . , - 
    s = re.sub(r"[^0-9\.\,\-]", "", s)
    # gérer , comme séparateur décimal
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df), index=df.index, dtype=float)
    return df[col].apply(_to_float)

def _fmt_money_us(v: float | int | str) -> str:
    try:
        f = float(v)
    except Exception:
        f = 0.0
    return f"${f:,.2f}"

def _norm(s: str) -> str:
    """Normalisation simple ASCII, sans accents (supposés absents),
    lower, et nettoyage caractères spéciaux (pattern sûr)."""
    s = _safe_str(s).lower()
    # placer '-' à la fin pour éviter les surprises en classe de caractères
    s = re.sub(r"[^a-z0-9+/_ ]+-?", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _date_for_widget(val: Any, default: date | None = None) -> date | None:
    """Retourne une date python.date sûre pour date_input (None autorisé)."""
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return default
        return d.date()
    except Exception:
        return default

def _month_index(m: Any) -> int:
    """Retourne un index 0..11 à partir d'un mois 'MM'/'M'/int."""
    s = _safe_str(m)
    if not s:
        return 0
    try:
        mm = int(s)
        return max(0, min(11, mm-1))
    except Exception:
        return 0

def make_client_id(nom: str, d: date) -> str:
    base = _safe_str(nom).strip().replace(" ", "_")
    return f"{base}-{d:%Y%m%d}"

def next_dossier_number(df: pd.DataFrame, start: int = 13057) -> int:
    if DOSSIER_COL not in df.columns or df.empty:
        return start
    try:
        mx = pd.to_numeric(df[DOSSIER_COL], errors="coerce").dropna().astype(int).max()
        return int(mx) + 1 if pd.notna(mx) else start
    except Exception:
        return start

# -------------------------------------------------
# Lecture / écriture fichiers (clients/visa)
# -------------------------------------------------
@st.cache_data(show_spinner=False)
def read_clients_file(path: str) -> pd.DataFrame:
    """Lit la table clients depuis:
       - un fichier unique (2 onglets: Clients & Visa), ou
       - un fichier Clients dédié (onglet Clients implicite)."""
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        # essayer l’onglet dédié
        return pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        # sinon lire la première feuille
        try:
            return pd.read_excel(path)
        except Exception:
            return pd.DataFrame()

@st.cache_data(show_spinner=False)
def read_visa_file(path: str) -> pd.DataFrame:
    """Lit la table visa (structure options)."""
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name=SHEET_VISA)
    except Exception:
        try:
            return pd.read_excel(path)
        except Exception:
            return pd.DataFrame()

def write_workbook(df_clients: pd.DataFrame,
                   clients_path: str | None,
                   df_visa: pd.DataFrame,
                   visa_path: str | None) -> Tuple[bool, str]:
    """Écrit selon la configuration :
       - 1 seul fichier (même chemin) => 2 onglets.
       - 2 fichiers séparés."""
    try:
        if clients_path and visa_path and os.path.abspath(clients_path) == os.path.abspath(visa_path):
            # un seul fichier (2 onglets)
            with pd.ExcelWriter(clients_path, engine="openpyxl") as wr:
                df_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                df_visa.to_excel(wr, sheet_name=SHEET_VISA, index=False)
        else:
            if clients_path:
                # écrire clients
                if os.path.exists(clients_path):
                    try:
                        with pd.ExcelWriter(clients_path, engine="openpyxl", mode="w") as wr:
                            df_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                    except Exception:
                        # fallback: sans nom d’onglet
                        with pd.ExcelWriter(clients_path, engine="openpyxl", mode="w") as wr:
                            df_clients.to_excel(wr, index=False)
                else:
                    with pd.ExcelWriter(clients_path, engine="openpyxl", mode="w") as wr:
                        df_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
            if visa_path:
                # écrire visa
                if os.path.exists(visa_path):
                    try:
                        with pd.ExcelWriter(visa_path, engine="openpyxl", mode="w") as wr:
                            df_visa.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                    except Exception:
                        with pd.ExcelWriter(visa_path, engine="openpyxl", mode="w") as wr:
                            df_visa.to_excel(wr, index=False)
                else:
                    with pd.ExcelWriter(visa_path, engine="openpyxl", mode="w") as wr:
                        df_visa.to_excel(wr, sheet_name=SHEET_VISA, index=False)
        return True, ""
    except Exception as e:
        return False, _safe_str(e)

# -------------------------------------------------
# Normalisations Clients & Visa
# -------------------------------------------------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            DOSSIER_COL,"ID_Client","Nom","Date","Mois","Categorie","Sous-categorie","Visa",
            HONO,AUTRE,TOTAL,"Payé","Reste","Paiements","Options",
            "Dossier envoyé","Date d'envoi","Dossier accepté","Date d'acceptation",
            "Dossier refusé","Date de refus","Dossier annulé","Date d'annulation","RFE",
        ])
    df = df.copy()
    # colonnes minimales
    for c in [DOSSIER_COL,"ID_Client","Nom","Date","Mois","Categorie","Sous-categorie","Visa",
              HONO,AUTRE,TOTAL,"Payé","Reste","Paiements","Options",
              "Dossier envoyé","Date d'envoi","Dossier accepté","Date d'acceptation",
              "Dossier refusé","Date de refus","Dossier annulé","Date d'annulation","RFE"]:
        if c not in df.columns:
            df[c] = None

    # numériques
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = df[c].apply(_to_float)

    # total si manquant
    df[TOTAL] = df[HONO] + df[AUTRE]
    # reste si manquant
    df["Reste"] = (df[TOTAL] - df["Payé"]).apply(lambda x: max(0.0, x))

    # Date → str (on gardera des vraie date dans les widgets)
    # mais on crée colonnes _Année_/_MoisNum_ pour tri
    try:
        dser = pd.to_datetime(df["Date"], errors="coerce")
        df["_Année_"]   = dser.dt.year
        df["_MoisNum_"] = dser.dt.month
    except Exception:
        df["_Année_"] = pd.NA
        df["_MoisNum_"] = pd.NA

    # Paiements & Options en types sûrs
    def _json_to_list(x):
        if isinstance(x, list):
            return x
        try:
            j = json.loads(_safe_str(x) or "[]")
            return j if isinstance(j, list) else []
        except Exception:
            return []

    def _json_to_any(x):
        # Options = liste de strings ou JSON list
        if isinstance(x, list):
            return x
        try:
            j = json.loads(_safe_str(x) or "[]")
            return j if isinstance(j, list) else []
        except Exception:
            return []

    df["Paiements"] = df["Paiements"].apply(_json_to_list)
    df["Options"]   = df["Options"].apply(_json_to_any)

    # Mois au format MM
    df["Mois"] = df["Mois"].apply(lambda m: f"{int(_safe_str(m) or '1'):02d}" if _safe_str(m) else "")
    return df

def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """Construit une map {Categorie: {Sous-categorie: {'options':[...]} } }
       À partir de la feuille Visa:
       - colonnes Catégorie / Sous-catégorie détectées
       - toutes les autres colonnes = colonnes d’options (0/1)
    """
    vm: Dict[str, Dict[str, Dict[str, Any]]] = {}
    if df_visa is None or df_visa.empty:
        return vm
    cols = list(df_visa.columns)
    # détecter colonnes
    cat_col = None
    sub_col = None
    for c in cols:
        n = _norm(c)
        if n in ("categorie", "category") and cat_col is None:
            cat_col = c
        if "sous" in n and "categor" in n and sub_col is None:
            sub_col = c
    if cat_col is None:
        cat_col = "Categorie" if "Categorie" in df_visa.columns else cols[0]
    if sub_col is None:
        if "Sous-categorie" in df_visa.columns:
            sub_col = "Sous-categorie"
        else:
            # 1re autre colonne
            cand = [c for c in cols if c != cat_col]
            sub_col = cand[0] if cand else cols[0]

    option_cols = [c for c in cols if c not in (cat_col, sub_col)]

    for _, r in df_visa.iterrows():
        cat = _safe_str(r.get(cat_col, "")).strip()
        sub = _safe_str(r.get(sub_col, "")).strip()
        if not cat or not sub:
            continue
        opts: List[str] = []
        for oc in option_cols:
            v = _to_float(r.get(oc, 0))
            if v == 1.0:
                opts.append(oc)
        vm.setdefault(cat, {})
        vm[cat].setdefault(sub, {"options": []})
        vm[cat][sub]["options"] = sorted(list(set(opts)))
    return vm

def compute_visa_string(sub: str, selected_opts: List[str]) -> str:
    """Concatène 'Sous-categorie' + première option (ou toutes séparées).
       Spéc: résultat = 'sub' + ' ' + 'opt1' si une seule, sinon 'sub opt1 / opt2 / ...'"""
    sub = _safe_str(sub)
    if not selected_opts:
        return sub
    if len(selected_opts) == 1:
        return f"{sub} {selected_opts[0]}"
    return f"{sub} " + " / ".join(selected_opts)

def render_option_checkboxes(option_labels: List[str], keyprefix: str, preselected: List[str]) -> List[str]:
    if not option_labels:
        st.info("Aucune option disponible pour cette sous-catégorie.")
        return []
    cols = st.columns(min(4, max(1, len(option_labels))))
    chosen: List[str] = []
    for i, lab in enumerate(option_labels):
        default = lab in (preselected or [])
        with cols[i % len(cols)]:
            if st.checkbox(lab, value=default, key=f"{keyprefix}_{i}"):
                chosen.append(lab)
    return chosen

# -------------------------------------------------
# Entête UI
# -------------------------------------------------
st.title("🛂 Visa Manager")

# -------------------------------------------------
# Sélecteur de mode de chargement & mémorisation des chemins
# -------------------------------------------------
LAST_FILE_JSON = os.path.join(".", ".last_paths.json")
def _save_last_paths(clients_path: str|None, visa_path: str|None):
    try:
        with open(LAST_FILE_JSON, "w", encoding="utf-8") as f:
            json.dump({"clients": clients_path, "visa": visa_path}, f)
    except Exception:
        pass

def _load_last_paths() -> Tuple[str|None, str|None]:
    try:
        with open(LAST_FILE_JSON, "r", encoding="utf-8") as f:
            d = json.load(f)
            return d.get("clients"), d.get("visa")
    except Exception:
        return None, None

st.markdown("## 📂 Fichiers")
mode = st.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"], horizontal=True)

up_clients = None
up_visa    = None
single_file = None

if mode == "Deux fichiers (Clients & Visa)":
    c1, c2 = st.columns(2)
    with c1:
        up_clients = st.file_uploader("Clients (xlsx)", type=["xlsx"], key="u_clients")
    with c2:
        up_visa    = st.file_uploader("Visa (xlsx)", type=["xlsx"], key="u_visa")
else:
    single_file = st.file_uploader("Fichier unique (2 onglets: Clients & Visa)", type=["xlsx"], key="u_single")

last_clients_path, last_visa_path = _load_last_paths()

# Construire chemins (si upload → écrire en mémoire disque)
def _persist_upload(up, fallback_name: str) -> str|None:
    if up is None:
        return None
    try:
        data = up.read()
        path = os.path.join(".", fallback_name)
        with open(path, "wb") as f:
            f.write(data)
        return path
    except Exception:
        return None

clients_path = None
visa_path    = None

if mode == "Deux fichiers (Clients & Visa)":
    if up_clients:
        clients_path = _persist_upload(up_clients, "clients.xlsx")
    else:
        clients_path = last_clients_path
    if up_visa:
        visa_path = _persist_upload(up_visa, "visa.xlsx")
    else:
        visa_path = last_visa_path
else:
    if single_file:
        p = _persist_upload(single_file, "workbook.xlsx")
        clients_path = p
        visa_path    = p
    else:
        # dernier fichier unique si le cas
        if last_clients_path and last_visa_path and last_clients_path == last_visa_path:
            clients_path = last_clients_path
            visa_path    = last_visa_path

# Sauver mémorisation
_save_last_paths(clients_path, visa_path)

# -------------------------------------------------
# Lire dataframes et normaliser
# -------------------------------------------------
df_clients_raw = read_clients_file(clients_path) if clients_path else pd.DataFrame()
df_visa_raw    = read_visa_file(visa_path) if visa_path else pd.DataFrame()

df_clients = normalize_clients(df_clients_raw)
df_all = df_clients.copy()

# construire visa_map pour cascades & options
visa_map = build_visa_map(df_visa_raw.copy()) if not df_visa_raw.empty else {}

# -------------------------------------------------
# Tabs, dans l’ordre demandé (ancienne présentation)
# -------------------------------------------------
tabs = st.tabs([
    "📊 Dashboard",        # 0
    "📈 Analyses",         # 1
    "🏦 Escrow",           # 2
    "👤 Clients",          # 3 (CRUD + paiements)
    "🧾 Gestion",          # 4 (éditeur Visa)
    "📄 Visa (aperçu)"     # 5
])

SID = "vm"  # suffixe de clés uniques

# ==============================================
# 📊 ONGLET : Dashboard (filtres + KPI + tableau)
# ==============================================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucune donnée client chargée.")
    else:
        years  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1,13)]
        cats   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        f1, f2, f3, f4, f5 = st.columns([1,1,1,1,2])
        fy = f1.multiselect("Année", years, default=[], key=f"dash_years_{SID}")
        fm = f2.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{SID}")
        fc = f3.multiselect("Catégorie", cats, default=[], key=f"dash_cats_{SID}")
        fs = f4.multiselect("Sous-catégorie", subs, default=[], key=f"dash_subs_{SID}")
        fv = f5.multiselect("Visa", visas, default=[], key=f"dash_visas_{SID}")

        view = df_all.copy()
        if fy: view = view[view["_Année_"].isin(fy)]
        if fm: view = view[view["Mois"].astype(str).isin(fm)]
        if fc: view = view[view["Categorie"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        # KPI compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(view)}")
        k2.metric("Honoraires", _fmt_money_us(float(_safe_num_series(view, HONO).sum())))
        k3.metric("Payé",      _fmt_money_us(float(_safe_num_series(view, "Payé").sum())))
        k4.metric("Reste",     _fmt_money_us(float(_safe_num_series(view, "Reste").sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        # Tableau
        detail = view.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in detail.columns:
                detail[c] = _safe_num_series(detail, c).map(_fmt_money_us)
        if "Date" in detail.columns:
            try:
                detail["Date"] = pd.to_datetime(detail["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                detail["Date"] = detail["Date"].astype(str)

        show_cols = [c for c in [
            DOSSIER_COL,"ID_Client","Nom","Categorie","Sous-categorie","Visa",
            "Date","Mois", HONO, AUTRE, TOTAL,"Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in detail.columns]

        sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in detail.columns]
        detail_sorted = detail.sort_values(by=sort_cols) if sort_cols else detail
        detail_sorted = detail_sorted.loc[:, ~detail_sorted.columns.duplicated()].copy()

        st.dataframe(detail_sorted[show_cols].reset_index(drop=True),
                     use_container_width=True, key=f"dash_tbl_{SID}")



# ==============================================
# 📈 ONGLET : Analyses (filtres + KPI + graphs + détail)
# ==============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_years_{SID}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"a_months_{SID}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"a_cats_{SID}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"a_subs_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_visas_{SID}")

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPI compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money_us(float(_safe_num_series(dfA, HONO).sum())))
        k3.metric("Payé",      _fmt_money_us(float(_safe_num_series(dfA, "Payé").sum())))
        k4.metric("Reste",     _fmt_money_us(float(_safe_num_series(dfA, "Reste").sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        # Graph 1 : dossiers par catégorie
        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("#### Dossiers par catégorie")
            vc = dfA["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie","Nombre"]
            st.bar_chart(vc.set_index("Categorie"))

        # Graph 2 : honoraires par mois
        if not dfA.empty and HONO in dfA.columns and "Mois" in dfA.columns:
            st.markdown("#### Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois")
            st.line_chart(gm.set_index("Mois"))

        # Détails
        st.markdown("#### Détails filtrés")
        det = dfA.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).map(_fmt_money_us)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            DOSSIER_COL,"ID_Client","Nom","Categorie","Sous-categorie","Visa",
            "Date","Mois", HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in det.columns]

        sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_cols) if sort_cols else det
        det_sorted = det_sorted.loc[:, ~det_sorted.columns.duplicated()].copy()

        st.dataframe(det_sorted[show_cols].reset_index(drop=True),
                     use_container_width=True, key=f"a_tbl_{SID}")


# ==============================================
# 🏦 ONGLET : Escrow — synthèse simple
# ==============================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Payé"]  = _safe_num_series(dfE, "Payé")
        dfE["Reste"] = _safe_num_series(dfE, "Reste")
        dfE[TOTAL]   = _safe_num_series(dfE, TOTAL)

        agg = dfE.groupby("Categorie", as_index=False)[[TOTAL, "Payé", "Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg[TOTAL]).replace([pd.NA, pd.NaT], 0).fillna(0.0) * 100
        st.dataframe(agg, use_container_width=True, key=f"esc_agg_{SID}")

        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money_us(float(dfE[TOTAL].sum())))
        t2.metric("Payé",         _fmt_money_us(float(dfE["Payé"].sum())))
        t3.metric("Reste",        _fmt_money_us(float(dfE["Reste"].sum())))
        st.caption("NB : on peut affiner l’Escrow si besoin (transferts après « Dossier envoyé »…).")


# ==============================================
# 👤 ONGLET : Clients — CRUD + paiements + statuts
# ==============================================
def _read_clients(path: str|None) -> pd.DataFrame:
    return normalize_clients(read_clients_file(path)) if path else pd.DataFrame()

def _write_clients(df: pd.DataFrame, path: str|None):
    if not path:
        st.error("Aucun fichier Clients défini.")
        return
    # si 1 seul fichier (clients_path == visa_path), écrire les deux
    if clients_path and visa_path and os.path.abspath(clients_path) == os.path.abspath(visa_path):
        ok, msg = write_workbook(df, clients_path, df_visa_raw, visa_path)
    else:
        ok, msg = write_workbook(df, clients_path, df_visa_raw, None)
    if not ok:
        st.error("Erreur d’écriture : " + msg)

def build_visa_option_selector(visa_map: Dict[str,Any], cat: str, sub: str,
                               keyprefix: str, preselected: Dict[str,Any]|List[str]|None=None) -> Tuple[str, Dict[str,Any], str]:
    """Affiche les cases à cocher pour les options de la sous-catégorie choisie.
       Retourne: (visa_str, options_dict, info_message)
    """
    info = ""
    if cat not in visa_map or sub not in visa_map.get(cat, {}):
        st.info("Aucune option disponible pour cette sous-catégorie.")
        return sub, {"options":[]}, info

    row_opts = visa_map[cat][sub].get("options", [])
    # préselection
    preset_list: List[str] = []
    if isinstance(preselected, dict):
        preset_list = preselected.get("options", []) if isinstance(preselected.get("options"), list) else []
    elif isinstance(preselected, list):
        preset_list = preselected

    st.markdown("Options disponibles :")
    chosen = render_option_checkboxes(row_opts, keyprefix=keyprefix, preselected=preset_list)
    visa_str = compute_visa_string(sub, chosen)
    return visa_str, {"options": chosen}, info

with tabs[3]:
    st.subheader("👤 Clients — Ajouter / Modifier / Supprimer / Paiements")

    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_op_{SID}")
    df_live = _read_clients(clients_path)

    # ---------- AJOUT ----------
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")
        d1, d2, d3 = st.columns(3)
        nom  = d1.text_input("Nom", "", key=f"add_nom_{SID}")
        dt   = d2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=date.today().month-1, key=f"add_mois_{SID}")

        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""]+cats, index=0, key=f"add_cat_{SID}")
        sel_sub = ""
        visa_final, opts_dict, info_msg = "", {"options":[]}, ""
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"add_sub_{SID}")
            if sel_sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=f"add_opts_{SID}", preselected=[]
                )
        if info_msg:
            st.info(info_msg)

        f1, f2 = st.columns(2)
        honor = f1.number_input(HONO, min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
        other = f2.number_input(AUTRE, min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")

        st.markdown("#### 📌 Statuts initiaux")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent   = s1.checkbox("Dossier envoyé", key=f"add_sent_{SID}")
        sent_d = s1.date_input("Date d'envoi", value=None, key=f"add_sentd_{SID}")
        acc    = s2.checkbox("Dossier accepté", key=f"add_acc_{SID}")
        acc_d  = s2.date_input("Date d'acceptation", value=None, key=f"add_accd_{SID}")
        ref    = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
        ref_d  = s3.date_input("Date de refus", value=None, key=f"add_refd_{SID}")
        ann    = s4.checkbox("Dossier annulé", key=f"add_ann_{SID}")
        ann_d  = s4.date_input("Date d'annulation", value=None, key=f"add_annd_{SID}")
        rfe    = s5.checkbox("RFE", key=f"add_rfe_{SID}")
        if rfe and not any([sent, acc, ref, ann]):
            st.warning("⚠️ RFE doit être coché avec un autre statut (envoyé/accepté/refusé/annulé).")

        if st.button("💾 Enregistrer le client", key=f"btn_add_{SID}"):
            if not nom:
                st.warning("Veuillez saisir le nom.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Veuillez choisir la catégorie et la sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            paye  = 0.0
            reste = max(0.0, total - paye)
            did   = make_client_id(nom, dt)
            dossier_n = next_dossier_number(df_live, start=13057)

            new_row = {
                DOSSIER_COL: dossier_n,
                "ID_Client": did,
                "Nom": nom,
                "Date": dt,
                "Mois": f"{int(mois):02d}",
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": visa_final if visa_final else sel_sub,
                HONO: float(honor),
                AUTRE: float(other),
                TOTAL: total,
                "Payé": 0.0,
                "Reste": reste,
                "Paiements": [],
                "Options": opts_dict,
                "Dossier envoyé": 1 if sent else 0,
                "Date d'envoi": sent_d if sent_d else (dt if sent else None),
                "Dossier accepté": 1 if acc else 0,
                "Date d'acceptation": acc_d,
                "Dossier refusé": 1 if ref else 0,
                "Date de refus": ref_d,
                "Dossier annulé": 1 if ann else 0,
                "Date d'annulation": ann_d,
                "RFE": 1 if rfe else 0,
            }
            df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            _write_clients(df_new, clients_path)
            st.success("Client ajouté.")
            st.cache_data.clear()
            st.rerun()

    # ---------- MODIFICATION / PAIEMENTS ----------
    if op == "Modifier":
        st.markdown("### ✏️ Modifier un client & gérer les paiements")
        if df_live.empty:
            st.info("Aucun client à modifier.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            s1, s2 = st.columns(2)
            target_name = s1.selectbox("Nom", [""]+names, index=0, key=f"mod_nom_{SID}")
            target_id   = s2.selectbox("ID_Client", [""]+ids, index=0, key=f"mod_id_{SID}")

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is None or not mask.any():
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()

            d1, d2, d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nomv_{SID}")
            dt   = d2.date_input("Date de création", value=_date_for_widget(row.get("Date"), default=date.today()), key=f"mod_date_{SID}")
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=_month_index(row.get("Mois","01")), key=f"mod_mois_{SID}")

            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""]+cats,
                                   index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                                   key=f"mod_cat_{SID}")

            sel_sub = _safe_str(row.get("Sous-categorie",""))
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""]+subs,
                                       index=(subs.index(sel_sub)+1 if sel_sub in subs else 0),
                                       key=f"mod_sub_{SID}")

            # Options déjà enregistrées
            preset_opts = row.get("Options", [])
            if not isinstance(preset_opts, (list, dict)):
                try:
                    preset_opts = json.loads(_safe_str(preset_opts) or "[]")
                except Exception:
                    preset_opts = []
            visa_final, opts_dict, info_msg = "", {"options":[]}, ""
            if sel_cat and sel_sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=f"mod_opts_{SID}", preselected=preset_opts
                )
            if info_msg:
                st.info(info_msg)

            f1, f2 = st.columns(2)
            h_val = float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0])
            o_val = float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0])
            honor = f1.number_input(HONO, min_value=0.0, value=h_val, step=50.0, format="%.2f", key=f"mod_h_{SID}")
            other = f2.number_input(AUTRE, min_value=0.0, value=o_val, step=20.0, format="%.2f", key=f"mod_o_{SID}")

            st.markdown("#### 📌 Statuts & dates")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent = s1.checkbox("Dossier envoyé", value=bool(row.get("Dossier envoyé")), key=f"mod_sent_{SID}")
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
            acc  = s2.checkbox("Dossier accepté", value=bool(row.get("Dossier accepté")), key=f"mod_acc_{SID}")
            acc_d = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
            ref  = s3.checkbox("Dossier refusé", value=bool(row.get("Dossier refusé")), key=f"mod_ref_{SID}")
            ref_d = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=f"mod_refd_{SID}")
            ann  = s4.checkbox("Dossier annulé", value=bool(row.get("Dossier annulé")), key=f"mod_ann_{SID}")
            ann_d = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
            rfe  = s5.checkbox("RFE", value=bool(row.get("RFE")), key=f"mod_rfe_{SID}")
            if rfe and not any([sent, acc, ref, ann]):
                st.warning("⚠️ RFE doit être coché avec un autre statut (envoyé/accepté/refusé/annulé).")

            # Paiements
            st.markdown("#### 💵 Paiements")
            pay_hist = row.get("Paiements", [])
            if not isinstance(pay_hist, list):
                try:
                    pay_hist = json.loads(_safe_str(pay_hist) or "[]")
                except Exception:
                    pay_hist = []

            if pay_hist:
                st.write("Historique :")
                for i, p in enumerate(pay_hist):
                    st.write(f"• {p.get('date','')} — {p.get('mode','')} — {_fmt_money_us(_to_float(p.get('montant',0)))}")

            cpay1, cpay2, cpay3 = st.columns(3)
            pay_date = cpay1.date_input("Date paiement", value=date.today(), key=f"mod_paydate_{SID}")
            pay_mode = cpay2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], index=0, key=f"mod_paymode_{SID}")
            pay_amt  = cpay3.number_input("Montant (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"mod_payamt_{SID}")
            add_pay  = st.button("➕ Ajouter paiement", key=f"btn_addpay_{SID}")

            total = float(honor) + float(other)
            paye  = float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0])
            reste = max(0.0, total - paye)

            if add_pay:
                if pay_amt <= 0:
                    st.warning("Montant de paiement doit être > 0.")
                    st.stop()
                # ne pas dépasser le reste négatif
                if pay_amt > reste + 1e-9:
                    st.warning("Le paiement dépasse le reste à payer.")
                    st.stop()
                pay_hist.append({"date": str(pay_date), "mode": pay_mode, "montant": float(pay_amt)})
                paye2  = paye + float(pay_amt)
                reste2 = max(0.0, total - paye2)
                df_live.at[idx, "Paiements"] = pay_hist
                df_live.at[idx, "Payé"] = paye2
                df_live.at[idx, "Reste"] = reste2
                _write_clients(df_live, clients_path)
                st.success("Paiement ajouté.")
                st.cache_data.clear()
                st.rerun()

            # Enregistrer les modifs principales
            if st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}"):
                if not nom:
                    st.warning("Nom requis.")
                    st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisissez catégorie et sous-catégorie.")
                    st.stop()
                total = float(honor) + float(other)
                paye  = float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0])
                reste = max(0.0, total - paye)

                df_live.at[idx, "Nom"] = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(mois):02d}"
                df_live.at[idx, "Categorie"] = sel_cat
                df_live.at[idx, "Sous-categorie"] = sel_sub
                df_live.at[idx, "Visa"] = (visa_final if visa_final else sel_sub)
                df_live.at[idx, HONO] = float(honor)
                df_live.at[idx, AUTRE] = float(other)
                df_live.at[idx, TOTAL] = total
                df_live.at[idx, "Reste"] = reste
                df_live.at[idx, "Options"] = opts_dict

                df_live.at[idx, "Dossier envoyé"] = 1 if sent else 0
                df_live.at[idx, "Date d'envoi"] = sent_d if sent_d else (dt if sent else None)
                df_live.at[idx, "Dossier accepté"] = 1 if acc else 0
                df_live.at[idx, "Date d'acceptation"] = acc_d
                df_live.at[idx, "Dossier refusé"] = 1 if ref else 0
                df_live.at[idx, "Date de refus"] = ref_d
                df_live.at[idx, "Dossier annulé"] = 1 if ann else 0
                df_live.at[idx, "Date d'annulation"] = ann_d
                df_live.at[idx, "RFE"] = 1 if rfe else 0

                _write_clients(df_live, clients_path)
                st.success("Modifications enregistrées.")
                st.cache_data.clear()
                st.rerun()

    # ---------- SUPPRESSION ----------
    if op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client à supprimer.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            s1, s2 = st.columns(2)
            target_name = s1.selectbox("Nom", [""]+names, index=0, key=f"del_nom_{SID}")
            target_id   = s2.selectbox("ID_Client", [""]+ids, index=0, key=f"del_id_{SID}")

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({DOSSIER_COL: row.get(DOSSIER_COL,""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                    df_new = df_live[~mask].copy()
                    _write_clients(df_new, clients_path)
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()


# ==============================================
# 🧾 ONGLET : Gestion — éditeur Visa (cases à cocher)
# ==============================================
with tabs[4]:
    st.subheader("🧾 Gestion — Feuille Visa (options 0/1)")
    dfv = df_visa_raw.copy()

    if dfv.empty:
        st.info("Aucune donnée Visa chargée.")
    else:
        st.caption("Chaque ligne = (Catégorie, Sous-catégorie, ...options=0/1). Les colonnes hors Catégorie/Sous-catégorie sont des cases à cocher (1=coché).")
        st.dataframe(dfv, use_container_width=True, key=f"visa_view_{SID}")

        st.markdown("#### Modifier une ligne existante")
        # choix ligne par couple (Catégorie, Sous-catégorie)
        cat_col = "Categorie" if "Categorie" in dfv.columns else dfv.columns[0]
        # heuristique sous-cat
        sub_col = "Sous-catégorie" if "Sous-catégorie" in dfv.columns else None
        if not sub_col:
            for c in dfv.columns:
                if "sous" in c.lower() and "cat" in c.lower():
                    sub_col = c
                    break
        if not sub_col:
            # fallback: 2e colonne
            sub_col = dfv.columns[1] if len(dfv.columns) > 1 else dfv.columns[0]

        cats_v = sorted(dfv[cat_col].dropna().astype(str).unique().tolist())
        selC = st.selectbox("Catégorie", [""]+cats_v, index=0, key=f"vm_ed_cat_{SID}")
        if selC:
            subs_v = sorted(dfv[dfv[cat_col].astype(str)==selC][sub_col].dropna().astype(str).unique().tolist())
            selS = st.selectbox("Sous-catégorie", [""]+subs_v, index=0, key=f"vm_ed_sub_{SID}")
        else:
            selS = ""

        if selC and selS:
            row_mask = (dfv[cat_col].astype(str)==selC) & (dfv[sub_col].astype(str)==selS)
            opt_cols = [c for c in dfv.columns if c not in (cat_col, sub_col)]
            # affichage cases à cocher
            cols_opt = st.columns(min(4, len(opt_cols) if opt_cols else 1))
            new_vals = {}
            for i, oc in enumerate(opt_cols):
                val = int(_to_float(dfv.loc[row_mask, oc].iloc[0])) if row_mask.any() else 0
                with cols_opt[i % len(cols_opt)]:
                    new_vals[oc] = 1 if st.checkbox(oc, value=bool(val), key=f"visa_ck_{SID}_{i}") else 0

            if st.button("💾 Enregistrer ligne Visa", key=f"visa_save_line_{SID}"):
                for oc, v in new_vals.items():
                    dfv.loc[row_mask, oc] = v
                # écrire
                if clients_path and visa_path and os.path.abspath(clients_path) == os.path.abspath(visa_path):
                    ok, msg = write_workbook(df_clients, clients_path, dfv, visa_path)
                else:
                    ok, msg = write_workbook(df_clients, clients_path, dfv, visa_path)
                if ok:
                    st.success("Ligne Visa enregistrée.")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error("Erreur écriture : " + msg)

        st.markdown("---")
        st.markdown("#### ➕ Ajouter une nouvelle ligne Visa")
        n1, n2 = st.columns(2)
        addC = n1.text_input("Catégorie", "", key=f"visa_add_cat_{SID}")
        addS = n2.text_input("Sous-catégorie", "", key=f"visa_add_sub_{SID}")
        # proposer colonnes options existantes
        opt_cols = [c for c in dfv.columns if c not in (cat_col, sub_col)]
        cols_opt2 = st.columns(min(4, len(opt_cols) if opt_cols else 1))
        new_vals2 = {}
        for i, oc in enumerate(opt_cols):
            with cols_opt2[i % len(cols_opt2)]:
                new_vals2[oc] = 1 if st.checkbox(f"{oc} (nouvelle)", value=False, key=f"visa_add_ck_{SID}_{i}") else 0

        if st.button("➕ Ajouter la ligne Visa", key=f"visa_add_row_{SID}"):
            if not addC or not addS:
                st.warning("Catégorie et Sous-catégorie sont requises.")
            else:
                row_data = {cat_col: addC, sub_col: addS}
                for oc in opt_cols:
                    row_data[oc] = new_vals2.get(oc, 0)
                dfv = pd.concat([dfv, pd.DataFrame([row_data])], ignore_index=True)

                # écrire
                if clients_path and visa_path and os.path.abspath(clients_path) == os.path.abspath(visa_path):
                    ok, msg = write_workbook(df_clients, clients_path, dfv, visa_path)
                else:
                    ok, msg = write_workbook(df_clients, clients_path, dfv, visa_path)
                if ok:
                    st.success("Nouvelle ligne Visa ajoutée.")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error("Erreur écriture : " + msg)


# ==============================================
# 📄 ONGLET : Visa (aperçu — filtres)
# ==============================================
with tabs[5]:
    st.subheader("📄 Visa — aperçu et filtres")

    if df_visa_raw.empty:
        st.info("Aucune donnée Visa chargée.")
    else:
        # colonnes Catégorie / Sous-catégorie
        cat_col = "Categorie" if "Categorie" in df_visa_raw.columns else df_visa_raw.columns[0]
        sub_col = "Sous-catégorie" if "Sous-catégorie" in df_visa_raw.columns else None
        if not sub_col:
            for c in df_visa_raw.columns:
                if "sous" in c.lower() and "cat" in c.lower():
                    sub_col = c
                    break
        if not sub_col:
            sub_col = df_visa_raw.columns[1] if len(df_visa_raw.columns) > 1 else df_visa_raw.columns[0]

        cats_v = sorted(df_visa_raw[cat_col].dropna().astype(str).unique().tolist())
        v1, v2 = st.columns(2)
        fc = v1.multiselect("Catégorie", cats_v, default=[], key=f"v_c_{SID}")
        if fc:
            subs_v = sorted(df_visa_raw[df_visa_raw[cat_col].astype(str).isin(fc)][sub_col].dropna().astype(str).unique().tolist())
        else:
            subs_v = sorted(df_visa_raw[sub_col].dropna().astype(str).unique().tolist())
        fs = v2.multiselect("Sous-catégorie", subs_v, default=[], key=f"v_s_{SID}")

        viewv = df_visa_raw.copy()
        if fc:
            viewv = viewv[viewv[cat_col].astype(str).isin(fc)]
        if fs:
            viewv = viewv[viewv[sub_col].astype(str).isin(fs)]

        st.dataframe(viewv, use_container_width=True, key=f"visa_tbl_{SID}")


# ==============================================
# 💾 Export global (Clients + Visa) — ZIP
# ==============================================
st.markdown("---")
st.markdown("### 💾 Export global (Clients + Visa)")
colz1, colz2 = st.columns([1,3])

with colz1:
    if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID}"):
        try:
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients
                df_export = _read_clients(clients_path)
                with BytesIO() as xbuf:
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                        df_export.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                    zf.writestr("Clients.xlsx", xbuf.getvalue())
                # Visa
                try:
                    # si fichier unique : on joint tel quel
                    if clients_path and visa_path and os.path.abspath(clients_path) == os.path.abspath(visa_path):
                        zf.write(visa_path, "Workbook.xlsx")
                    else:
                        with BytesIO() as vb:
                            with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                df_visa_raw.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                            zf.writestr("Visa.xlsx", vb.getvalue())
                except Exception:
                    pass
            st.session_state[f"zip_export_{SID}"] = buf.getvalue()
            st.success("Archive prête.")
        except Exception as e:
            st.error("Erreur de préparation : " + _safe_str(e))

with colz2:
    if st.session_state.get(f"zip_export_{SID}"):
        st.download_button(
            label="⬇️ Télécharger l’export (ZIP)",
            data=st.session_state[f"zip_export_{SID}"],
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
            key=f"zip_dl_{SID}",
        )
