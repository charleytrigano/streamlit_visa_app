# ===============================
# ======   PARTIE 1 / 2    ======
# ===============================

from __future__ import annotations

import os, io, json, zipfile, hashlib, unicodedata, re
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple, Any

import pandas as pd
import streamlit as st

# ---------------------------
# Configuration générale UI
# ---------------------------
st.set_page_config(page_title="🛂 Visa Manager", layout="wide")

st.markdown("""
<style>
/* KPI compacts */
.small-metrics .stMetric {
  padding: 0.2rem 0.4rem !important;
}
.small-metrics [data-testid="stMetricValue"] {
  font-size: 0.9rem !important;
}
.small-metrics [data-testid="stMetricDelta"] {
  font-size: 0.7rem !important;
}
/* Sélecteurs compacts */
.stSelectbox, .stMultiSelect {
  font-size: 0.9rem !important;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Constantes des colonnes
# ---------------------------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"
DOSSIER_COL = "Dossier N"

# ---------------------------
# Utilitaires sûrs / robustes
# ---------------------------
def _safe_str(x: Any) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _to_float(x: Any) -> float:
    try:
        if isinstance(x, (int, float)): return float(x)
        s = _safe_str(x)
        s = s.replace(" ", "").replace("\xa0","").replace(",", ".")
        s = re.sub(r"[^0-9.\-]", "", s)
        return float(s) if s not in ("", "-", ".", "-.") else 0.0
    except Exception:
        return 0.0

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df), dtype=float)
    return pd.to_numeric(df[col].apply(_to_float), errors="coerce").fillna(0.0)

def _fmt_money_us(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(val, fallback: date | None = None) -> date | None:
    """Assure une valeur compatible st.date_input (évite NaT)."""
    if isinstance(val, date): return val
    if isinstance(val, datetime): return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d): return fallback
        return d.date()
    except Exception:
        return fallback

def _norm(s: str) -> str:
    """Normalise: ascii, lower, garde a-z0-9 + / _ espace et - (robuste aux accents)."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = s.encode("ascii", "ignore").decode("ascii")
    s = s.lower().strip()
    # IMPORTANT : '-' placé avant la fin/fin de classe pour éviter PatternError
    s = re.sub(r"[^a-z0-9+/_ -]+", " ", s)   # <- le '-' est en fin de classe
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def _month_index(m) -> int:
    try:
        mm = int(_safe_str(m))
        return max(0, min(11, mm-1))
    except Exception:
        return 0

def _best_index(opts: List[str], value: str) -> int:
    if value in opts:
        return opts.index(value) + 1  # +1 car on a une option "" devant
    return 0

# ---------------------------
# Cache — lecture fichiers
# ---------------------------
@st.cache_data(show_spinner=False)
def read_clients_file(path: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=SHEET_CLIENTS)

@st.cache_data(show_spinner=False)
def read_visa_file(path: str) -> pd.DataFrame:
    # Feuille Visa : entêtes = Catégorie, Sous-catégorie, puis colonnes d’options (1 = actif)
    df = pd.read_excel(path, sheet_name=SHEET_VISA)
    return df

@st.cache_data(show_spinner=False)
def write_workbook(df_clients: pd.DataFrame, clients_path: str,
                   df_visa: pd.DataFrame | None, visa_path: str | None) -> Tuple[bool, str]:
    try:
        if clients_path and visa_path and clients_path == visa_path:
            # fichier unique avec 2 onglets
            with pd.ExcelWriter(clients_path, engine="openpyxl") as wr:
                df_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                if df_visa is not None:
                    df_visa.to_excel(wr, sheet_name=SHEET_VISA, index=False)
            return True, ""
        else:
            # deux fichiers séparés
            with pd.ExcelWriter(clients_path, engine="openpyxl") as wr:
                df_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
            if visa_path and df_visa is not None:
                with pd.ExcelWriter(visa_path, engine="openpyxl") as wr:
                    df_visa.to_excel(wr, sheet_name=SHEET_VISA, index=False)
            return True, ""
    except Exception as e:
        return False, _safe_str(e)

# ---------------------------
# Normalisation Clients
# ---------------------------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            DOSSIER_COL,"ID_Client","Nom","Date","Mois","Categorie","Sous-categorie","Visa",
            HONO,AUTRE,TOTAL,"Payé","Reste","Paiements","Options","Notes",
            "Dossier envoyé","Date d'envoi","Dossier accepté","Date d'acceptation",
            "Dossier refusé","Date de refus","Dossier annulé","Date d'annulation","RFE","EscrowTransfers"
        ])

    out = df.copy()

    # Colonnes chiffrées
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c in out.columns:
            out[c] = _safe_num_series(out, c)

    # Totaux si manquants
    if TOTAL in out.columns:
        out[TOTAL] = _safe_num_series(out, HONO) + _safe_num_series(out, AUTRE)
    else:
        out[TOTAL] = _safe_num_series(out, HONO) + _safe_num_series(out, AUTRE)

    # Payé/Reste par défaut
    if "Payé" not in out.columns:
        out["Payé"] = 0.0
    if "Reste" not in out.columns:
        out["Reste"] = (out[TOTAL] - out["Payé"]).clip(lower=0)

    # Paiements & Options en listes
    for jscol in ["Paiements", "Options", "EscrowTransfers"]:
        if jscol not in out.columns:
            out[jscol] = [[] for _ in range(len(out))]
        else:
            out[jscol] = out[jscol].apply(lambda v: v if isinstance(v, list)
                                          else (json.loads(_safe_str(v) or "[]")
                                                if _safe_str(v).startswith("[") else []))

    # Bool/entiers statuts
    for b in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        if b not in out.columns:
            out[b] = 0
        else:
            out[b] = out[b].apply(lambda x: 1 if _to_float(x) != 0 else 0)

    # Dates statut : garder tel quel (converties à l’affichage)
    for dc in ["Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation"]:
        if dc not in out.columns:
            out[dc] = None

    # Date / Mois / Année techniques
    if "Date" in out.columns:
        try:
            dd = pd.to_datetime(out["Date"], errors="coerce")
            out["_Année_"] = dd.dt.year
            out["_MoisNum_"] = dd.dt.month
        except Exception:
            out["_Année_"] = None
            out["_MoisNum_"] = None
    if "Mois" in out.columns:
        out["Mois"] = out["Mois"].astype(str).str.zfill(2)
    else:
        out["Mois"] = out["_MoisNum_"].fillna(1).astype(int).astype(str).str.zfill(2)

    return out

# ---------------------------
# Construction visa_map depuis la feuille Visa
# ---------------------------
def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    Construit une structure :
    {
      "Affaires/Tourisme": {
         "B-1": {"options": ["COS","EOS", ...]},
         "B-2": {"options": [...]},
      },
      "Etudiants": {
         "F-1": {...}, ...
      }
    }
    Règle : toute colonne (hors Catégorie/Sous-catégorie) où la cellule vaut 1 => option active.
    Les libellés d’options viennent EXACTEMENT de la ligne d’entête.
    """
    if df_visa is None or df_visa.empty:
        return {}

    # détecter les colonnes Catégorie / Sous-catégorie par normalisation
    cols = list(df_visa.columns)
    cat_col = None
    sub_col = None
    for c in cols:
        n = _norm(c)
        if n in ("categorie", "category") and cat_col is None:
            cat_col = c
        if ("sous" in n and "categor" in n) and sub_col is None:
            sub_col = c
    # fallback
    if cat_col is None: cat_col = "Categorie" if "Categorie" in df_visa.columns else cols[0]
    if sub_col is None:
        sub_candidates = [c for c in cols if c != cat_col]
        sub_col = "Sous-categorie" if "Sous-categorie" in df_visa.columns else (sub_candidates[0] if sub_candidates else cols[0])

    option_cols = [c for c in cols if c not in (cat_col, sub_col)]

    visa_map: Dict[str, Dict[str, Dict[str, Any]]] = {}
    for _, row in df_visa.iterrows():
        cat = _safe_str(row.get(cat_col, "")).strip()
        sub = _safe_str(row.get(sub_col, "")).strip()
        if not cat or not sub:
            continue
        opts: List[str] = []
        for oc in option_cols:
            v = row.get(oc, 0)
            if _to_float(v) == 1:
                opts.append(_safe_str(oc).strip())

        visa_map.setdefault(cat, {}).setdefault(sub, {})["options"] = opts

    return visa_map

def render_option_checkboxes(options: List[str], keyprefix: str, preselected: List[str] | None = None) -> List[str]:
    """Affiche des cases à cocher pour chaque option (1ère ligne du Visa). Retourne la liste cochée."""
    if not options:
        return []
    pre = set(preselected or [])
    sel: List[str] = []
    cols = st.columns(min(4, max(1, len(options))))
    for i, opt in enumerate(options):
        with cols[i % len(cols)]:
            is_on = st.checkbox(opt, value=(opt in pre), key=f"{keyprefix}_{i}")
        if is_on:
            sel.append(opt)
    return sel

def compute_visa_string(sub: str, options: List[str]) -> str:
    """Affiche 'Sous-catégorie + (options triées)'. Ex: 'B-1 COS' ou 'F-1 COS EOS'."""
    sub = _safe_str(sub)
    if not options:
        return sub
    opt_txt = " ".join(sorted([_safe_str(o) for o in options]))
    return f"{sub} {opt_txt}".strip()

# ---------------------------
# ID et Dossier N
# ---------------------------
def make_client_id(nom: str, dt: date | datetime | None) -> str:
    base = _norm(nom).replace(" ", "")
    if not base:
        base = "client"
    d = _date_for_widget(dt, date.today())
    return f"{base}-{d:%Y%m%d}"

def next_dossier_number(df: pd.DataFrame, start: int = 13057) -> int:
    if DOSSIER_COL in df.columns:
        try:
            mx = pd.to_numeric(df[DOSSIER_COL], errors="coerce").max()
            return int(mx) + 1 if pd.notna(mx) else start
        except Exception:
            return start
    return start

# ---------------------------
# Mémoire de chemins (session)
# ---------------------------
if "last_clients_path" not in st.session_state: st.session_state.last_clients_path = ""
if "last_visa_path" not in st.session_state: st.session_state.last_visa_path = ""
if "zip_export_blob" not in st.session_state: st.session_state.zip_export_blob = None

# ---------------------------
# Zone chargement fichiers
# ---------------------------
st.markdown("## 📂 Fichiers")
mode = st.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"], horizontal=False)

up_clients = None
up_visa = None
single_file = None

if mode == "Deux fichiers (Clients & Visa)":
    up_clients = st.file_uploader("Clients (xlsx)", type=["xlsx"])
    up_visa    = st.file_uploader("Visa (xlsx)", type=["xlsx"])
else:
    single_file = st.file_uploader("Fichier unique (2 onglets : Clients & Visa)", type=["xlsx"])

# Appliquer chargements
clients_path = st.session_state.last_clients_path
visa_path    = st.session_state.last_visa_path

# Cas 2 fichiers
if up_clients is not None:
    clients_path = "clients_uploaded.xlsx"
    with open(clients_path, "wb") as f:
        f.write(up_clients.read())
    st.session_state.last_clients_path = clients_path

if up_visa is not None:
    visa_path = "visa_uploaded.xlsx"
    with open(visa_path, "wb") as f:
        f.write(up_visa.read())
    st.session_state.last_visa_path = visa_path

# Cas 1 fichier
if single_file is not None:
    both_path = "workbook_uploaded.xlsx"
    with open(both_path, "wb") as f:
        f.write(single_file.read())
    # on considère ce fichier comme source des 2 feuilles
    clients_path = both_path
    visa_path    = both_path
    st.session_state.last_clients_path = clients_path
    st.session_state.last_visa_path    = visa_path

# Si rien uploadé, tenter de retrouver derniers chemins (si existent sur disque)
if clients_path and not os.path.exists(clients_path):
    clients_path = ""
if visa_path and not os.path.exists(visa_path):
    visa_path = ""

# Alerte si manquants
if not clients_path:
    st.warning("Aucun fichier **Clients** chargé pour le moment.")
if not visa_path:
    st.warning("Aucun fichier **Visa** chargé pour le moment.")

st.markdown("# 🛂 Visa Manager")

# Charger les données actuelles
df_clients_raw = pd.DataFrame()
df_visa_raw    = pd.DataFrame()
if clients_path:
    try:
        df_clients_raw = read_clients_file(clients_path)
    except Exception as e:
        st.error("Impossible de lire la feuille Clients : " + _safe_str(e))
if visa_path:
    try:
        df_visa_raw = read_visa_file(visa_path)
    except Exception as e:
        st.error("Impossible de lire la feuille Visa : " + _safe_str(e))

df_all = normalize_clients(df_clients_raw)
visa_map = build_visa_map(df_visa_raw.copy()) if not df_visa_raw.empty else {}

# Clé de widgets stable selon fichiers
SID = hashlib.md5(f"{clients_path}|{visa_path}".encode()).hexdigest()[:6] if (clients_path or visa_path) else "base"

# ---------------------------
# Barre latérale — navigation
# ---------------------------
with st.sidebar:
    st.markdown("## Navigation")
    st.write("• 📊 Dashboard\n• 📈 Analyses\n• 🏦 Escrow\n• 👤 Clients\n• 📄 Visa (aperçu)")
    st.markdown("---")
    if clients_path:
        st.caption(f"Clients: `{os.path.basename(clients_path)}`")
    if visa_path:
        st.caption(f"Visa: `{os.path.basename(visa_path)}`")

# ---------------------------
# Tabs principaux
# ---------------------------
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "📄 Visa (aperçu)"])

# =========================================================
# 📊 ONGLET : Dashboard (liste + KPI + filtres)
# =========================================================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        years  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1, 13)]
        cats   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        f1, f2, f3, f4, f5 = st.columns([1,1,1,1,1])
        fy = f1.multiselect("Année", years, default=[], key=f"dash_years_{SID}")
        fm = f2.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{SID}")
        fc = f3.multiselect("Catégorie", cats, default=[], key=f"dash_cats_{SID}")
        fs = f4.multiselect("Sous-catégorie", subs, default=[], key=f"dash_subs_{SID}")
        fv = f5.multiselect("Visa", visas, default=[], key=f"dash_visas_{SID}")

        ff = df_all.copy()
        if fy: ff = ff[ff["_Année_"].isin(fy)]
        if fm: ff = ff[ff["Mois"].astype(str).isin(fm)]
        if fc: ff = ff[ff["Categorie"].astype(str).isin(fc)]
        if fs: ff = ff[ff["Sous-categorie"].astype(str).isin(fs)]
        if fv: ff = ff[ff["Visa"].astype(str).isin(fv)]

        # KPI compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Dossiers", f"{len(ff)}")
        k2.metric("Honoraires", _fmt_money_us(float(_safe_num_series(ff, HONO).sum())))
        k3.metric("Autres frais", _fmt_money_us(float(_safe_num_series(ff, AUTRE).sum())))
        k4.metric("Payé", _fmt_money_us(float(_safe_num_series(ff, "Payé").sum())))
        k5.metric("Reste", _fmt_money_us(float(_safe_num_series(ff, "Reste").sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        # Table
        view = ff.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in view.columns:
                view[c] = _safe_num_series(view, c).map(_fmt_money_us)
        if "Date" in view.columns:
            try:
                view["Date"] = pd.to_datetime(view["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                view["Date"] = view["Date"].astype(str)

        show_cols = [c for c in [
            DOSSIER_COL,"ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in view.columns]

        sort_cols = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        v = view.copy()
        if sort_cols:
            v = v.sort_values(by=sort_cols)
        v = v.loc[:, ~v.columns.duplicated()].copy()

        st.dataframe(v[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_tbl_{SID}")




# ===============================
# ======   PARTIE 2 / 2    ======
# ===============================

# =========================================================
# 📈 ONGLET : Analyses (filtres + KPI + graph + détails)
# =========================================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
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

        # Graphiques simples (barres)
        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("#### Dossiers par catégorie")
            vc = dfA["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie", "Nombre"]
            st.bar_chart(vc.set_index("Categorie"))

        if not dfA.empty and HONO in dfA.columns and "Mois" in dfA.columns:
            st.markdown("#### Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois")
            st.line_chart(gm.set_index("Mois"))

        # Détails formatés
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
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
            "Date", "Mois", HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"
        ] if c in det.columns]

        sort_cols = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_cols) if sort_cols else det
        det_sorted = det_sorted.loc[:, ~det_sorted.columns.duplicated()].copy()

        st.dataframe(det_sorted[show_cols].reset_index(drop=True),
                     use_container_width=True, key=f"a_det_{SID}")


# =========================================================
# 🏦 ONGLET : Escrow (synthèse)
# =========================================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")

    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE[TOTAL] = _safe_num_series(dfE, TOTAL)
        dfE["Payé"] = _safe_num_series(dfE, "Payé")
        dfE["Reste"] = _safe_num_series(dfE, "Reste")

        agg = dfE.groupby("Categorie", as_index=False)[[TOTAL, "Payé", "Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg[TOTAL]).replace([pd.NA, pd.NaT], 0).fillna(0) * 100.0
        st.dataframe(agg, use_container_width=True, key=f"esc_agg_{SID}")

        k1, k2, k3 = st.columns(3)
        k1.metric("Total", _fmt_money_us(float(dfE[TOTAL].sum())))
        k2.metric("Payé",  _fmt_money_us(float(dfE["Payé"].sum())))
        k3.metric("Reste", _fmt_money_us(float(dfE["Reste"].sum())))

        st.caption("NB : si vous tenez un compte ESCROW par dossier, vous pouvez utiliser la colonne 'Paiements' pour distinguer les acomptes et déclencher vos transferts une fois « Dossier envoyé » coché.")


# =========================================================
# 👤 ONGLET : Clients — CRUD + Paiements + Statuts
# =========================================================
with tabs[3]:
    st.subheader("👤 Clients — Ajouter / Modifier / Supprimer")

    # Relecture “raw” pour éditer/écrire
    df_live = pd.DataFrame()
    if clients_path:
        try:
            df_live = read_clients_file(clients_path)
        except Exception as e:
            st.error("Lecture Clients impossible : " + _safe_str(e))
    df_live = normalize_clients(df_live)

    action = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"cli_action_{SID}")

    # --------- AJOUT ---------
    if action == "Ajouter":
        st.markdown("### ➕ Ajouter un client")
        c1, c2, c3 = st.columns(3)
        nom  = c1.text_input("Nom", "", key=f"add_nom_{SID}")
        dt   = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=date.today().month-1, key=f"add_mois_{SID}")

        st.markdown("#### 🎯 Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""]+cats, index=0, key=f"add_cat_{SID}")
        sel_sub = ""
        selected_opts: List[str] = []
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"add_sub_{SID}")
            if sel_sub:
                opt_list = visa_map.get(sel_cat, {}).get(sel_sub, {}).get("options", [])
                selected_opts = render_option_checkboxes(opt_list, keyprefix=f"add_opt_{SID}", preselected=[])

        final_visa = compute_visa_string(sel_sub, selected_opts)

        f1, f2 = st.columns(2)
        honor = f1.number_input(HONO, min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
        other = f2.number_input(AUTRE, min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")

        st.markdown("#### 📌 Statuts (avec dates)")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent = s1.checkbox("Dossier envoyé", key=f"add_sent_{SID}")
        sent_d = s1.date_input("Date d'envoi", value=None, key=f"add_sentd_{SID}")
        acc = s2.checkbox("Dossier accepté", key=f"add_acc_{SID}")
        acc_d = s2.date_input("Date d'acceptation", value=None, key=f"add_accd_{SID}")
        ref = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
        ref_d = s3.date_input("Date de refus", value=None, key=f"add_refd_{SID}")
        ann = s4.checkbox("Dossier annulé", key=f"add_ann_{SID}")
        ann_d = s4.date_input("Date d'annulation", value=None, key=f"add_annd_{SID}")
        rfe = s5.checkbox("RFE", key=f"add_rfe_{SID}")
        if rfe and not any([sent, acc, ref, ann]):
            st.warning("RFE doit être coché avec un autre statut (envoyé/accepté/refusé/annulé).")

        if st.button("💾 Enregistrer", key=f"btn_add_{SID}"):
            if not nom:
                st.warning("Le nom est requis.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Choisissez la catégorie et la sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            paye  = 0.0
            reste = total - paye

            new_row = {
                DOSSIER_COL: next_dossier_number(df_live),
                "ID_Client": make_client_id(nom, dt),
                "Nom": nom,
                "Date": dt,
                "Mois": f"{int(_safe_str(mois)):02d}",
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": final_visa if final_visa else sel_sub,
                HONO: float(honor),
                AUTRE: float(other),
                TOTAL: total,
                "Payé": paye,
                "Reste": reste,
                "Paiements": [],
                "Options": selected_opts,
                "Dossier envoyé": 1 if sent else 0,
                "Date d'envoi": sent_d if sent_d else None,
                "Dossier accepté": 1 if acc else 0,
                "Date d'acceptation": acc_d if acc_d else None,
                "Dossier refusé": 1 if ref else 0,
                "Date de refus": ref_d if ref_d else None,
                "Dossier annulé": 1 if ann else 0,
                "Date d'annulation": ann_d if ann_d else None,
                "RFE": 1 if rfe else 0,
            }

            df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            ok, err = write_workbook(df_new, clients_path, df_visa_raw, visa_path)
            if ok:
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()
            else:
                st.error("Échec écriture : " + err)

    # --------- MODIFICATION ---------
    if action == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if df_live.empty:
            st.info("Aucun client à modifier.")
        else:
            # Sélection du client
            ids = df_live["ID_Client"].dropna().astype(str).tolist()
            target_id = st.selectbox("ID_Client", [""] + sorted(ids), index=0, key=f"mod_id_{SID}")
            if not target_id:
                st.stop()

            mask = (df_live["ID_Client"].astype(str) == target_id)
            if not mask.any():
                st.warning("Ligne introuvable.")
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()

            # En-tête : identifiants & info
            st.write({DOSSIER_COL: row.get(DOSSIER_COL,""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})

            d1, d2, d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nomv_{SID}")
            dt   = d2.date_input("Date de création", value=_date_for_widget(row.get("Date"), date.today()), key=f"mod_date_{SID}")
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=_month_index(row.get("Mois","01")), key=f"mod_mois_{SID}")

            # Visa cascade + options
            st.markdown("#### 🎯 Visa")
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""]+cats,
                                   index=_best_index(cats, preset_cat),
                                   key=f"mod_cat_{SID}")

            subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
            preset_sub = _safe_str(row.get("Sous-categorie",""))
            sel_sub = st.selectbox("Sous-catégorie", [""]+subs,
                                   index=_best_index(subs, preset_sub),
                                   key=f"mod_sub_{SID}")

            preset_opts = row.get("Options", [])
            if not isinstance(preset_opts, list):
                try:
                    preset_opts = json.loads(_safe_str(preset_opts) or "[]")
                    if not isinstance(preset_opts, list):
                        preset_opts = []
                except Exception:
                    preset_opts = []

            selected_opts: List[str] = []
            if sel_cat and sel_sub:
                opt_list = visa_map.get(sel_cat, {}).get(sel_sub, {}).get("options", [])
                selected_opts = render_option_checkboxes(opt_list, keyprefix=f"mod_opt_{SID}", preselected=preset_opts)
            final_visa = compute_visa_string(sel_sub, selected_opts)

            # Montants
            f1, f2, f3, f4 = st.columns(4)
            honor = f1.number_input(HONO, min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0]),
                                    step=50.0, format="%.2f", key=f"mod_h_{SID}")
            other = f2.number_input(AUTRE, min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0]),
                                    step=20.0, format="%.2f", key=f"mod_o_{SID}")
            paye  = float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0])
            reste = max(0.0, float(honor) + float(other) - paye)
            f3.metric("Payé", _fmt_money_us(paye))
            f4.metric("Reste", _fmt_money_us(reste))

            # Paiements
            st.markdown("#### 💵 Paiements")
            pay_list = row.get("Paiements", [])
            if not isinstance(pay_list, list):
                try:
                    pay_list = json.loads(_safe_str(pay_list) or "[]")
                    if not isinstance(pay_list, list):
                        pay_list = []
                except Exception:
                    pay_list = []
            # affichage historique
            if pay_list:
                disp = pd.DataFrame(pay_list)
                if "Montant" in disp.columns:
                    disp["Montant"] = disp["Montant"].apply(_to_float).map(_fmt_money_us)
                st.dataframe(disp, use_container_width=True, key=f"pay_hist_{SID}")
            else:
                st.caption("Aucun paiement enregistré.")

            # Ajout paiement si reste > 0
            if reste > 0.0001:
                ap1, ap2, ap3, ap4 = st.columns([1,1,1,1])
                pay_date = ap1.date_input("Date paiement", value=date.today(), key=f"pay_date_{SID}")
                pay_mode = ap2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"pay_mode_{SID}")
                pay_amt  = ap3.number_input("Montant (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"pay_amt_{SID}")
                if ap4.button("Ajouter paiement", key=f"btn_addpay_{SID}"):
                    add = float(pay_amt or 0.0)
                    if add <= 0:
                        st.warning("Le montant doit être > 0.")
                        st.stop()
                    pay_list.append({"Date": _safe_str(pay_date), "Mode": pay_mode, "Montant": add})
                    paye_new  = paye + add
                    total_new = float(honor) + float(other)
                    reste_new = max(0.0, total_new - paye_new)

                    # mise à jour row / df_live
                    df_live.at[idx, "Paiements"] = pay_list
                    df_live.at[idx, "Payé"] = paye_new
                    df_live.at[idx, "Reste"] = reste_new

                    ok, err = write_workbook(df_live, clients_path, df_visa_raw, visa_path)
                    if ok:
                        st.success("Paiement ajouté.")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("Erreur écriture : " + err)

            # Statuts + dates
            st.markdown("#### 📌 Statuts & dates")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent  = s1.checkbox("Dossier envoyé", value=bool(row.get("Dossier envoyé")), key=f"mod_sent_{SID}")
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
            acc   = s2.checkbox("Dossier accepté", value=bool(row.get("Dossier accepté")), key=f"mod_acc_{SID}")
            acc_d = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
            ref   = s3.checkbox("Dossier refusé", value=bool(row.get("Dossier refusé")), key=f"mod_ref_{SID}")
            ref_d = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=f"mod_refd_{SID}")
            ann   = s4.checkbox("Dossier annulé", value=bool(row.get("Dossier annulé")), key=f"mod_ann_{SID}")
            ann_d = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
            rfe   = s5.checkbox("RFE", value=bool(row.get("RFE")), key=f"mod_rfe_{SID}")

            if rfe and not any([sent, acc, ref, ann]):
                st.warning("RFE doit être coché avec un autre statut (envoyé/accepté/refusé/annulé).")

            if st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}"):
                if not nom:
                    st.warning("Le nom est requis.")
                    st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisissez Catégorie et Sous-catégorie.")
                    st.stop()

                total_new = float(honor) + float(other)
                # paye/reste déjà recalculés lors d’un ajout paiement ; sinon recalcul ici
                paye_now  = float(_safe_num_series(pd.DataFrame([df_live.loc[idx]]), "Payé").iloc[0])
                reste_now = max(0.0, total_new - paye_now)

                df_live.at[idx, "Nom"] = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(_safe_str(mois)):02d}"
                df_live.at[idx, "Categorie"] = sel_cat
                df_live.at[idx, "Sous-categorie"] = sel_sub
                df_live.at[idx, "Visa"] = (final_visa if final_visa else sel_sub)
                df_live.at[idx, HONO] = float(honor)
                df_live.at[idx, AUTRE] = float(other)
                df_live.at[idx, TOTAL] = total_new
                df_live.at[idx, "Reste"] = reste_now
                df_live.at[idx, "Options"] = selected_opts
                df_live.at[idx, "Dossier envoyé"] = 1 if sent else 0
                df_live.at[idx, "Date d'envoi"] = sent_d if sent_d else None
                df_live.at[idx, "Dossier accepté"] = 1 if acc else 0
                df_live.at[idx, "Date d'acceptation"] = acc_d if acc_d else None
                df_live.at[idx, "Dossier refusé"] = 1 if ref else 0
                df_live.at[idx, "Date de refus"] = ref_d if ref_d else None
                df_live.at[idx, "Dossier annulé"] = 1 if ann else 0
                df_live.at[idx, "Date d'annulation"] = ann_d if ann_d else None
                df_live.at[idx, "RFE"] = 1 if rfe else 0

                ok, err = write_workbook(df_live, clients_path, df_visa_raw, visa_path)
                if ok:
                    st.success("Modifications enregistrées.")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error("Erreur écriture : " + err)

    # --------- SUPPRESSION ---------
    if action == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client à supprimer.")
        else:
            ids = df_live["ID_Client"].dropna().astype(str).tolist()
            target_id = st.selectbox("ID_Client", [""]+sorted(ids), index=0, key=f"del_id_{SID}")
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
                if mask.any():
                    row = df_live[mask].iloc[0]
                    st.write({DOSSIER_COL: row.get(DOSSIER_COL,""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                    if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                        df_new = df_live[~mask].copy()
                        ok, err = write_workbook(df_new, clients_path, df_visa_raw, visa_path)
                        if ok:
                            st.success("Client supprimé.")
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error("Erreur écriture : " + err)


# =========================================================
# 📄 ONGLET : Visa (aperçu & édition par cases à cocher)
# =========================================================
with tabs[4]:
    st.subheader("📄 Visa (aperçu & édition)")

    if df_visa_raw.empty:
        st.info("Aucune feuille Visa chargée.")
    else:
        # Détecter colonnes catégorie / sous-catégorie / options
        cols = list(df_visa_raw.columns)
        cat_col = None
        sub_col = None
        for c in cols:
            n = _norm(c)
            if n in ("categorie", "category") and cat_col is None:
                cat_col = c
            if ("sous" in n and "categor" in n) and sub_col is None:
                sub_col = c
        if cat_col is None: cat_col = "Categorie" if "Categorie" in df_visa_raw.columns else cols[0]
        if sub_col is None:
            candidates = [c for c in cols if c != cat_col]
            sub_col = "Sous-categorie" if "Sous-categorie" in df_visa_raw.columns else (candidates[0] if candidates else cols[0])

        option_cols = [c for c in cols if c not in (cat_col, sub_col)]

        # Filtre pour aperçu
        f1, f2 = st.columns(2)
        cats = sorted(df_visa_raw[cat_col].dropna().astype(str).unique().tolist())
        sel_cat_v = f1.selectbox("Catégorie", [""]+cats, index=0, key=f"vz_cat_{SID}")
        subs = sorted(df_visa_raw[df_visa_raw[cat_col].astype(str)==sel_cat_v][sub_col].dropna().astype(str).unique().tolist()) if sel_cat_v else []
        sel_sub_v = f2.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"vz_sub_{SID}")

        # Affichage des options (cases) et édition
        df_edit = df_visa_raw.copy()
        if sel_cat_v and sel_sub_v:
            mask = (df_edit[cat_col].astype(str)==sel_cat_v) & (df_edit[sub_col].astype(str)==sel_sub_v)
            if mask.any():
                rix = df_edit[mask].index[0]
                st.markdown("#### Options actives (cocher = 1)")
                cols_opt = st.columns(min(4, max(1, len(option_cols))))
                new_vals = {}
                for i, oc in enumerate(option_cols):
                    cur = _to_float(df_edit.at[rix, oc]) == 1.0
                    with cols_opt[i % len(cols_opt)]:
                        val = st.checkbox(oc, value=cur, key=f"vz_opt_{SID}_{i}")
                    new_vals[oc] = 1 if val else 0
                if st.button("💾 Enregistrer Visa", key=f"btn_vz_save_{SID}"):
                    for oc, v in new_vals.items():
                        df_edit.at[rix, oc] = v
                    # Écrire
                    ok, err = write_workbook(read_clients_file(clients_path) if clients_path else df_clients_raw,
                                             clients_path,
                                             df_edit,
                                             visa_path)
                    if ok:
                        st.success("Feuille Visa enregistrée.")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("Erreur écriture : " + err)

        # Aperçu complet
        st.markdown("#### Aperçu de la table Visa")
        st.dataframe(df_visa_raw, use_container_width=True, key=f"vz_tbl_{SID}")


# =========================================================
# 💽 Export global (ZIP)
# =========================================================
st.markdown("---")
st.markdown("### 💽 Export global (Clients + Visa)")

colz1, colz2 = st.columns([1,3])
with colz1:
    if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID}"):
        try:
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients
                if clients_path and os.path.exists(clients_path):
                    # si fichier unique => on le met tel quel
                    if clients_path == visa_path:
                        zf.write(clients_path, arcname=os.path.basename(clients_path))
                    else:
                        # écriture clients propre
                        with BytesIO() as cb:
                            with pd.ExcelWriter(cb, engine="openpyxl") as wr:
                                read_clients_file.clear()  # pour recharger proprement
                                dfc = read_clients_file(clients_path)
                                dfc.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                            zf.writestr("Clients.xlsx", cb.getvalue())
                # Visa
                if visa_path and os.path.exists(visa_path) and clients_path != visa_path:
                    zf.write(visa_path, arcname=os.path.basename(visa_path))
            st.session_state.zip_export_blob = buf.getvalue()
            st.success("Archive prête.")
        except Exception as e:
            st.error("Erreur de préparation : " + _safe_str(e))

with colz2:
    if st.session_state.get("zip_export_blob"):
        st.download_button(
            label="⬇️ Télécharger l’export (ZIP)",
            data=st.session_state.zip_export_blob,
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
            key=f"zip_dl_{SID}",
        )