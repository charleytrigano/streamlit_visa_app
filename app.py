# ==============================================
# PARTIE 1 / 4 — Imports • Constantes • Helpers • Chargement fichiers
# ==============================================
from __future__ import annotations
import os, json, re, zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple
import pandas as pd
import streamlit as st

# ------------ Constantes colonnes / feuilles
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"
DOSSIER_COL   = "Dossier N"
HONO          = "Montant honoraires (US $)"
AUTRE         = "Autres frais (US $)"
TOTAL         = "Total (US $)"

# Fichier local pour mémoriser les derniers chemins
LAST_PATHS_FILE = ".cache_visamanager.json"

# ------------ Utils sûrs
def _safe_str(x) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _safe_num_series(df_or_series, col=None) -> pd.Series:
    s = df_or_series[col] if isinstance(df_or_series, pd.DataFrame) else df_or_series
    s = pd.to_numeric(s, errors="coerce")
    return s.fillna(0.0)

def _fmt_money_us(x: float|int) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(val):
    """Renvoie une date (ou None) compatible st.date_input."""
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        return None if pd.isna(d) else d.date()
    except Exception:
        return None

def _make_client_id(nom: str, d: date) -> str:
    base = re.sub(r"[^A-Za-z0-9]+", "-", _safe_str(nom)).strip("-").lower()
    if not isinstance(d, (date, datetime)):
        d = date.today()
    if isinstance(d, datetime):
        d = d.date()
    return f"{base}-{d:%Y%m%d}"

def _next_dossier(df_clients: pd.DataFrame, start: int = 13057) -> int:
    if DOSSIER_COL in df_clients.columns:
        vals = pd.to_numeric(df_clients[DOSSIER_COL], errors="coerce").dropna()
        if len(vals):
            return int(vals.max()) + 1
    return int(start)

# ------------ Mémoire des derniers chemins
def _save_last_paths(clients_path: str|None, visa_path: str|None) -> None:
    try:
        with open(LAST_PATHS_FILE, "w", encoding="utf-8") as f:
            json.dump({"clients_path": clients_path or "", "visa_path": visa_path or ""}, f)
    except Exception:
        pass

def _load_last_paths() -> Tuple[str|None, str|None]:
    try:
        with open(LAST_PATHS_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        cp = d.get("clients_path") or None
        vp = d.get("visa_path") or None
        if cp and not os.path.exists(cp): cp = None
        if vp and not os.path.exists(vp): vp = None
        return cp, vp
    except Exception:
        return None, None

# ------------ Lecture Excel générique
def _read_excel_any(path_or_buf, sheet_hint: str|None) -> Tuple[pd.DataFrame, str|None]:
    if not path_or_buf:
        return pd.DataFrame(), None
    try:
        xls = pd.ExcelFile(path_or_buf)
        sh = sheet_hint if (sheet_hint and sheet_hint in xls.sheet_names) else xls.sheet_names[0]
        return pd.read_excel(path_or_buf, sheet_name=sh), sh
    except Exception as e:
        st.warning(f"Impossible de lire le fichier : {e}")
        return pd.DataFrame(), None

def _read_clients(path: str|None) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        cols = [DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
                "Categorie", "Sous-categorie", "Visa",
                HONO, AUTRE, "Commentaires",
                TOTAL, "Payé", "Reste",
                "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"]
        return pd.DataFrame(columns=cols)
    try:
        xls = pd.ExcelFile(path)
        sh = SHEET_CLIENTS if SHEET_CLIENTS in xls.sheet_names else xls.sheet_names[0]
        return pd.read_excel(path, sheet_name=sh)
    except Exception:
        return pd.DataFrame()

def _write_clients(df: pd.DataFrame, path: str|None) -> None:
    if not path:
        st.error("Aucun chemin Clients pour sauvegarder.")
        return
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
    except Exception as e:
        st.error(f"Erreur écriture Clients : {e}")

def _read_visa(path: str|None) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        xls = pd.ExcelFile(path)
        sh = SHEET_VISA if SHEET_VISA in xls.sheet_names else xls.sheet_names[0]
        return pd.read_excel(path, sheet_name=sh)
    except Exception:
        return pd.DataFrame()

# ------------ Construction de la carte Visa (catégorie -> sous-catégorie -> options)
def build_visa_map(df_visa_raw: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, List[str]]]]:
    """
    Le fichier Visa (onglet 'Visa') contient :
      - colonnes : 'Categorie', 'Sous-categorie', puis des colonnes d'options (cases à cocher)
      - dans les colonnes d'options, la valeur 1 indique que l’option est disponible pour la sous-catégorie
    On construit un dict : {Categorie: {Sous-categorie: {"options": [labels]}}}
    Les 'labels' d’option sont simplement les noms de colonnes (ligne d’en-tête).
    """
    if df_visa_raw.empty:
        return {}

    # Harmoniser noms colonnes usuelles (tolère variantes)
    cols = {c.strip(): c for c in df_visa_raw.columns}
    c_cat = cols.get("Categorie") or cols.get("Catégorie") or "Categorie"
    c_sub = cols.get("Sous-categorie") or cols.get("Sous-categories") or "Sous-categorie"

    if c_cat not in df_visa_raw.columns or c_sub not in df_visa_raw.columns:
        # On tente de deviner : premières colonnes = categorie / sous-categorie
        df = df_visa_raw.copy()
        df.columns = [str(c).strip() for c in df.columns]
        if len(df.columns) >= 2:
            df.rename(columns={df.columns[0]: "Categorie", df.columns[1]: "Sous-categorie"}, inplace=True)
            c_cat, c_sub = "Categorie", "Sous-categorie"
            df_visa_raw = df

    option_cols = [c for c in df_visa_raw.columns if c not in [c_cat, c_sub]]

    vmap: Dict[str, Dict[str, Dict[str, List[str]]]] = {}
    for _, r in df_visa_raw.iterrows():
        cat = _safe_str(r.get(c_cat)).strip()
        sub = _safe_str(r.get(c_sub)).strip()
        if not cat or not sub:
            continue
        opts: List[str] = []
        for oc in option_cols:
            val = r.get(oc)
            # 1, True, "1", "x", etc. => on retient "1" strictement et True
            if str(val).strip() in ("1", "True", "true") or val is True:
                opts.append(str(oc))
        vmap.setdefault(cat, {})[sub] = {"options": opts}
    return vmap

# ------------ UI fichiers + mémoire des chemins
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

st.session_state.setdefault("clients_path", None)
st.session_state.setdefault("visa_path", None)

# Restauration des derniers chemins
last_c, last_v = _load_last_paths()
if st.session_state["clients_path"] is None and last_c:
    st.session_state["clients_path"] = last_c
if st.session_state["visa_path"] is None and last_v:
    st.session_state["visa_path"] = last_v

clients_path = st.session_state.get("clients_path")
visa_path    = st.session_state.get("visa_path")

st.sidebar.markdown("## 📂 Fichiers")
mode = st.sidebar.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"], index=0, key="file_mode")

up_clients = None
up_visa    = None
up_single  = None

if mode == "Deux fichiers (Clients & Visa)":
    up_clients = st.sidebar.file_uploader("Clients (xlsx)", type=["xlsx"], key="up_clients")
    up_visa    = st.sidebar.file_uploader("Visa (xlsx)",     type=["xlsx"], key="up_visa")
else:
    up_single  = st.sidebar.file_uploader("Fichier unique (2 onglets)", type=["xlsx"], key="up_single")

# Lecture selon le mode
df_all      = pd.DataFrame()
df_visa_raw = pd.DataFrame()

if up_single is not None:
    df_all, _      = _read_excel_any(up_single, SHEET_CLIENTS)
    df_visa_raw, _ = _read_excel_any(up_single, SHEET_VISA)
    tmp_path = "last_single.xlsx"
    with open(tmp_path, "wb") as f: f.write(up_single.getbuffer())
    st.session_state["clients_path"] = tmp_path
    st.session_state["visa_path"]    = tmp_path
    clients_path, visa_path = tmp_path, tmp_path
    _save_last_paths(tmp_path, tmp_path)
else:
    if up_clients is not None:
        df_all, _ = _read_excel_any(up_clients, SHEET_CLIENTS)
        tmp_c = "last_clients.xlsx"
        with open(tmp_c, "wb") as f: f.write(up_clients.getbuffer())
        st.session_state["clients_path"] = tmp_c
        clients_path = tmp_c
    else:
        if clients_path:
            df_all, _ = _read_excel_any(clients_path, SHEET_CLIENTS)

    if up_visa is not None:
        df_visa_raw, _ = _read_excel_any(up_visa, SHEET_VISA)
        tmp_v = "last_visa.xlsx"
        with open(tmp_v, "wb") as f: f.write(up_visa.getbuffer())
        st.session_state["visa_path"] = tmp_v
        visa_path = tmp_v
    else:
        if visa_path:
            df_visa_raw, _ = _read_excel_any(visa_path, SHEET_VISA)
    _save_last_paths(st.session_state.get("clients_path"), st.session_state.get("visa_path"))

# État dans la sidebar
def _count(df):
    try: return len(df)
    except: return 0

st.sidebar.caption("### État du chargement")
st.sidebar.write({
    "Clients": {"source": clients_path or ("upload unique" if up_single else "—"), "lignes": _count(df_all)},
    "Visa":    {"source": visa_path or ("upload unique" if up_single else "—"), "lignes": _count(df_visa_raw)},
})
st.sidebar.info("Les fichiers chargés sont mémorisés et rechargés automatiquement au prochain démarrage.")

# ------------ Normalisation Clients (dérivés utiles : Total, Reste, Année, MoisNum)
if not df_all.empty:
    # Assurer présence colonnes clés
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c not in df_all.columns:
            df_all[c] = 0.0

    df_all[HONO] = _safe_num_series(df_all, HONO)
    df_all[AUTRE]= _safe_num_series(df_all, AUTRE)
    if TOTAL in df_all.columns:
        df_all[TOTAL] = _safe_num_series(df_all, TOTAL)
    else:
        df_all[TOTAL] = df_all[HONO] + df_all[AUTRE]

    if "Payé" in df_all.columns:
        df_all["Payé"] = _safe_num_series(df_all, "Payé")
    else:
        df_all["Payé"] = 0.0

    df_all["Reste"] = (df_all[TOTAL] - df_all["Payé"]).clip(lower=0.0)

    # Année / MoisNum à partir de 'Date' si possible, sinon de 'Mois'
    if "Date" in df_all.columns:
        dts = pd.to_datetime(df_all["Date"], errors="coerce")
        df_all["_Année_"]  = dts.dt.year
        df_all["_MoisNum_"] = dts.dt.month
    if "_Année_" not in df_all.columns:
        df_all["_Année_"] = pd.to_numeric(df_all.get("Année", pd.Series(dtype=float)), errors="coerce")
    if "_MoisNum_" not in df_all.columns:
        df_all["_MoisNum_"] = pd.to_numeric(df_all.get("Mois", pd.Series(dtype=float)), errors="coerce")

    # Mois affichable (MM)
    if "Mois" in df_all.columns:
        df_all["Mois"] = df_all["Mois"].apply(lambda x: f"{int(x):02d}" if pd.notna(pd.to_numeric(x, errors="coerce")) else _safe_str(x))
    else:
        df_all["Mois"] = df_all["_MoisNum_"].apply(lambda x: f"{int(x):02d}" if pd.notna(x) else "")

# ------------ Carte Visa (cat -> sous-cat -> options)
visa_map: Dict[str, Dict[str, Dict[str, List[str]]]] = build_visa_map(df_visa_raw.copy()) if not df_visa_raw.empty else {}

# ------------ Générateur de clés Streamlit robustes (évite collisions)
def _ns() -> str:
    cp = _safe_str(st.session_state.get("clients_path") or "nocl")
    vp = _safe_str(st.session_state.get("visa_path") or "novi")
    return f"{abs(hash((cp, vp)))%10**8}"

def skey(*parts) -> str:
    return "k_" + "_".join([_ns(), *[str(p) for p in parts]])

# ------------ CSS (KPI compacts)
st.markdown("""
<style>
.small-metrics .stMetric {
  padding: 0.2rem 0.5rem;
}
.small-metrics [data-testid="stMetricValue"] {
  font-size: 1.0rem;
}
.small-metrics [data-testid="stMetricDelta"] {
  font-size: 0.8rem;
}
</style>
""", unsafe_allow_html=True)

# ------------ Tabs principaux
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "🧾 Gestion", "📄 Visa (aperçu)"])




# ==============================================
# PARTIE 2 / 4 — 📊 ONGLET : Dashboard (filtres + KPI + liste + export)
# ==============================================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucune donnée client chargée (onglet Clients du fichier Excel manquant ou vide).")
    else:
        # --- Listes pour filtres
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all.get("_Année_", pd.Series(dtype=float)),
                                                        errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_all.get("Categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        subsA   = sorted(df_all.get("Sous-categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        visasA  = sorted(df_all.get("Visa", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

        # --- Filtres
        c1, c2, c3, c4, c5 = st.columns([1,1,1,1,1])
        fy = c1.multiselect("Année", yearsA, default=[], key=skey("dash","years"))
        fm = c2.multiselect("Mois (MM)", monthsA, default=[], key=skey("dash","months"))
        fc = c3.multiselect("Catégorie", catsA, default=[], key=skey("dash","cats"))
        fs = c4.multiselect("Sous-catégorie", subsA, default=[], key=skey("dash","subs"))
        fv = c5.multiselect("Visa", visasA, default=[], key=skey("dash","visas"))

        # Filtres statut (optionnels, s'ils existent)
        st.markdown("###### Statuts (facultatif)")
        sA, sB, sC, sD, sE = st.columns(5)
        f_sent = sA.selectbox("Dossier envoyé", ["—","Oui","Non"], index=0, key=skey("dash","sent"))
        f_acc  = sB.selectbox("Dossier accepté", ["—","Oui","Non"], index=0, key=skey("dash","acc"))
        f_ref  = sC.selectbox("Dossier refusé", ["—","Oui","Non"], index=0, key=skey("dash","ref"))
        f_ann  = sD.selectbox("Dossier annulé", ["—","Oui","Non"], index=0, key=skey("dash","ann"))
        f_rfe  = sE.selectbox("RFE", ["—","Oui","Non"], index=0, key=skey("dash","rfe"))

        # --- Appliquer filtres
        view = df_all.copy()

        if fy: view = view[view.get("_Année_", "").isin(fy)]
        if fm: view = view[view.get("Mois", "").astype(str).isin(fm)]
        if fc: view = view[view.get("Categorie", "").astype(str).isin(fc)]
        if fs: view = view[view.get("Sous-categorie", "").astype(str).isin(fs)]
        if fv: view = view[view.get("Visa", "").astype(str).isin(fv)]

        def _apply_flag(df: pd.DataFrame, col: str, val: str) -> pd.DataFrame:
            if col not in df.columns or val == "—":
                return df
            if val == "Oui":
                return df[df[col].fillna(0).astype(int) == 1]
            else:
                return df[df[col].fillna(0).astype(int) == 0]

        for col_flag, choice in [
            ("Dossier envoyé",  f_sent),
            ("Dossier accepté", f_acc),
            ("Dossier refusé",  f_ref),
            ("Dossier annulé",  f_ann),
            ("RFE",             f_rfe),
        ]:
            view = _apply_flag(view, col_flag, choice)

        # --- KPI compacts
        for coln in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if coln in view.columns:
                view[coln] = _safe_num_series(view, coln)

        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(view)}")
        k2.metric("Honoraires", _fmt_money_us(float(view.get(HONO, pd.Series(dtype=float)).sum())))
        k3.metric("Payé",      _fmt_money_us(float(view.get("Payé", pd.Series(dtype=float)).sum())))
        k4.metric("Reste",     _fmt_money_us(float(view.get("Reste", pd.Series(dtype=float)).sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        # --- Tableau
        # Mise en forme affichage (argent, date)
        show = view.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in show.columns:
                show[c] = _safe_num_series(show, c).map(_fmt_money_us)
        if "Date" in show.columns:
            try:
                show["Date"] = pd.to_datetime(show["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                show["Date"] = show["Date"].astype(str)

        show_cols = [c for c in [
            DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
            "Categorie", "Sous-categorie", "Visa",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"
        ] if c in show.columns]

        # Tri doux si colonnes présentes
        sort_by = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_by) if sort_by else view

        # éviter les doublons de colonnes (pyarrow n'aime pas)
        view_sorted = view_sorted.loc[:, ~view_sorted.columns.duplicated()].copy()
        st.dataframe(
            view_sorted[show_cols].reset_index(drop=True),
            use_container_width=True,
            key=skey("dash","table")
        )

        # --- Export Excel (vue filtrée)
        st.markdown("##### Export de la vue filtrée")
        try:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as wr:
                # version numérique (utile pour retraitement)
                view_num = df_all.copy()
                for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                    if c in view_num.columns:
                        view_num[c] = _safe_num_series(view_num, c)
                view_num = view_num.loc[view_sorted.index]  # aligner au tri/filtre
                view_num[show_cols].to_excel(wr, sheet_name="Dashboard", index=False)
            st.download_button(
                "⬇️ Télécharger (Excel)",
                data=buf.getvalue(),
                file_name="Dashboard_filtre.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=skey("dash","dl")
            )
        except Exception as e:
            st.warning(f"Export non disponible : {_safe_str(e)}")




# ==============================================
# PARTIE 3 / 4 — 📈 Analyses + 🏦 Escrow
# ==============================================

# --------- Helpers clés uniques déjà définis en P1 :
# def skey(*parts) -> str: ...

with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        # Colonnes numériques normalisées
        for coln in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if coln in df_all.columns:
                df_all[coln] = _safe_num_series(df_all, coln)

        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all.get("_Année_", pd.Series(dtype=float)),
                                                        errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_all.get("Categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        subsA   = sorted(df_all.get("Sous-categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        visasA  = sorted(df_all.get("Visa", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

        # --- Filtres Analyse
        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=skey("an","years"))
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=skey("an","months"))
        fc = a3.multiselect("Catégorie", catsA, default=[], key=skey("an","cats"))
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=skey("an","subs"))
        fv = a5.multiselect("Visa", visasA, default=[], key=skey("an","visas"))

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA.get("_Année_", "").isin(fy)]
        if fm: dfA = dfA[dfA.get("Mois", "").astype(str).isin(fm)]
        if fc: dfA = dfA[dfA.get("Categorie", "").astype(str).isin(fc)]
        if fs: dfA = dfA[dfA.get("Sous-categorie", "").astype(str).isin(fs)]
        if fv: dfA = dfA[dfA.get("Visa", "").astype(str).isin(fv)]

        # --- KPI compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money_us(float(dfA.get(HONO, pd.Series(dtype=float)).sum())))
        k3.metric("Payé",      _fmt_money_us(float(dfA.get("Payé", pd.Series(dtype=float)).sum())))
        k4.metric("Reste",     _fmt_money_us(float(dfA.get("Reste", pd.Series(dtype=float)).sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        # --- Répartition et %
        st.markdown("#### Répartition par Catégorie & Sous-catégorie")
        cL, cR = st.columns(2)

        if not dfA.empty and "Categorie" in dfA.columns:
            vc = (dfA.groupby("Categorie", as_index=False)
                    .agg(N=("Categorie","size"),
                         Honoraires=(HONO,"sum") if HONO in dfA.columns else ("Categorie","size")))
            totN = int(vc["N"].sum() or 1)
            vc["% Dossiers"] = (vc["N"]/totN*100).round(1)
            with cL:
                st.dataframe(vc.sort_values("N", ascending=False),
                             use_container_width=True, height=260, key=skey("an","tab","vc"))

        if not dfA.empty and "Sous-categorie" in dfA.columns:
            vs = (dfA.groupby("Sous-categorie", as_index=False)
                    .agg(N=("Sous-categorie","size"),
                         Honoraires=(HONO,"sum") if HONO in dfA.columns else ("Sous-categorie","size")))
            totNs = int(vs["N"].sum() or 1)
            vs["% Dossiers"] = (vs["N"]/totNs*100).round(1)
            with cR:
                st.dataframe(vs.sort_values("N", ascending=False).head(25),
                             use_container_width=True, height=260, key=skey("an","tab","vs"))

        # --- Graphiques
        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("#### 📊 Dossiers par catégorie")
            g1 = (dfA.groupby("Categorie", as_index=False).size().rename(columns={"size":"Nombre"}))
            st.bar_chart(g1.set_index("Categorie"))

        if not dfA.empty and "Mois" in dfA.columns and HONO in dfA.columns:
            st.markdown("#### 📈 Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = (tmp.groupby("Mois", as_index=False)[HONO].sum()
                    .reindex([f"{m:02d}" for m in range(1,13)], fill_value=0)
                    .sort_values("Mois"))
            st.line_chart(gm.set_index("Mois"))

        # --- Comparaison A vs B
        st.markdown("#### 🔁 Comparaison de périodes (A vs B)")
        ca1, ca2, cb1, cb2 = st.columns(4)
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=skey("cmp","ya"))
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=skey("cmp","ma"))
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=skey("cmp","yb"))
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=skey("cmp","mb"))

        def _filter_period(base: pd.DataFrame, ys, ms) -> pd.DataFrame:
            d = base.copy()
            for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                if c in d.columns:
                    d[c] = _safe_num_series(d, c)
            if ys: d = d[d.get("_Année_", "").isin(ys)]
            if ms: d = d[d.get("Mois", "").astype(str).isin(ms)]
            return d

        dfA_A = _filter_period(df_all, pa_years, pa_month)
        dfA_B = _filter_period(df_all, pb_years, pb_month)

        cpa, cpb = st.columns(2)
        with cpa:
            st.metric("A — Dossiers", f"{len(dfA_A)}")
            st.metric("A — Honoraires", _fmt_money_us(float(dfA_A.get(HONO, pd.Series(dtype=float)).sum())))
        with cpb:
            st.metric("B — Dossiers", f"{len(dfA_B)}")
            st.metric("B — Honoraires", _fmt_money_us(float(dfA_B.get(HONO, pd.Series(dtype=float)).sum())))

        st.markdown("##### Comparaison par mois")
        def _mk_month_series(d: pd.DataFrame) -> pd.DataFrame:
            if d.empty or "Mois" not in d.columns or HONO not in d.columns:
                return pd.DataFrame({"Mois": [], "Honoraires": []})
            t = d.copy()
            t["Mois"] = t["Mois"].astype(str)
            return (t.groupby("Mois", as_index=False)[HONO].sum()
                     .reindex([f"{m:02d}" for m in range(1,13)], fill_value=0)
                     .rename(columns={HONO:"Honoraires"}))

        A = _mk_month_series(dfA_A); A["Période"] = "A"
        B = _mk_month_series(dfA_B); B["Période"] = "B"
        comp = pd.concat([A, B], ignore_index=True)
        if not comp.empty:
            wide = comp.pivot_table(index="Mois", columns="Période", values="Honoraires", fill_value=0)
            st.bar_chart(wide)

        # --- Détails filtrés
        st.markdown("#### 🧾 Détails des dossiers filtrés")
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
            DOSSIER_COL,"ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in det.columns]

        sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_cols) if sort_cols else det
        det_sorted = det_sorted.loc[:, ~det_sorted.columns.duplicated()].copy()

        st.dataframe(det_sorted[show_cols].reset_index(drop=True),
                     use_container_width=True, key=skey("an","detail"))

# ==============================================
# 🏦 Escrow — synthèse compacte
# ==============================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        for coln in [TOTAL, "Payé", "Reste"]:
            if coln in dfE.columns:
                dfE[coln] = _safe_num_series(dfE, coln)

        # KPI compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money_us(float(dfE.get(TOTAL, pd.Series(dtype=float)).sum())))
        t2.metric("Payé",         _fmt_money_us(float(dfE.get("Payé", pd.Series(dtype=float)).sum())))
        t3.metric("Reste",        _fmt_money_us(float(dfE.get("Reste", pd.Series(dtype=float)).sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        agg_cols = [c for c in [TOTAL, "Payé", "Reste"] if c in dfE.columns]
        if agg_cols:
            agg = dfE.groupby("Categorie", as_index=False)[agg_cols].sum()
            if TOTAL in agg and "Payé" in agg:
                agg["% Payé"] = ((agg["Payé"] / agg[TOTAL]).replace([pd.NA, pd.NaT], 0).fillna(0.0)*100).round(1)
            st.dataframe(agg.sort_values(by=TOTAL if TOTAL in agg else agg_cols[0], ascending=False),
                         use_container_width=True, key=skey("esc","agg"))
        else:
            st.info("Colonnes financières manquantes pour l’agrégation.")




# ==============================================
# PARTIE 4 / 4 — 👤 Clients • 🧾 Gestion (CRUD) • 📄 Visa (aperçu) • Exports
# (Colle ce bloc à la fin de ton app.py)
# Dépend de helpers/constantes déjà définis en amont :
#  - skey, _safe_str, _safe_num_series, _fmt_money_us, _read_clients, _write_clients
#  - _make_client_id, _next_dossier, build_visa_option_selector
#  - visa_map, df_all, df_visa_raw, clients_path, visa_path, tabs
#  - DOSSIER_COL, HONO, AUTRE, TOTAL
# ==============================================

from datetime import date, datetime
import json, os, zipfile
from io import BytesIO
from typing import List

# ---------- petit helper pour composer "Visa" à partir des options cochées
def compose_visa_label(sub: str, options: List[str]) -> str:
    sub = _safe_str(sub)
    if not options:
        return sub
    if len(options) == 1:
        return f"{sub} {options[0]}"
    return f"{sub} {'+'.join(options)}"

# ---------- rendu des options (cases à cocher) selon visa_map
def render_options_for(cat: str, sub: str, keyprefix: str, preset: List[str] | None = None) -> List[str]:
    opts = visa_map.get(cat, {}).get(sub, {}).get("options", []) if visa_map else []
    if not opts:
        st.info("Aucune option définie pour cette sous-catégorie.")
        return []
    sel: List[str] = []
    cols = st.columns(min(4, max(1, len(opts))))
    for i, oc in enumerate(opts):
        default_checked = (preset is not None and oc in preset)
        checked = cols[i % len(cols)].checkbox(oc, value=default_checked, key=skey(keyprefix, "opt", i, oc))
        if checked:
            sel.append(oc)
    return sel

# ---------- helper date sûre pour widget
def _date_for_widget(val):
    if isinstance(val, date):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d2 = pd.to_datetime(val, errors="coerce")
        return d2.date() if pd.notna(d2) else None
    except Exception:
        return None


# ==============================================
# 👤 ONGLET : Clients — suivi & paiements (corrigé)
# ==============================================
with tabs[3]:
    st.subheader("👤 Clients — suivi & paiements")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        names = sorted(df_all.get("Nom", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        sel_name = st.selectbox("Nom du client", [""] + names, index=0, key=skey("cl","selname"))
        if sel_name:
            rows = df_all[df_all["Nom"].astype(str) == sel_name].copy()
            if rows.empty:
                st.warning("Client introuvable.")
            else:
                ids = rows.get("ID_Client", pd.Series(dtype=str)).astype(str).tolist()
                sel_id = st.selectbox("ID_Client", [""] + ids, index=(1 if len(ids) else 0), key=skey("cl","selid"))

                if sel_id:
                    row = rows[rows["ID_Client"].astype(str) == sel_id]
                    row = row.iloc[0].copy() if not row.empty else pd.Series(dtype=object)
                else:
                    row = rows.iloc[0].copy()

                # --- Sécurisation : convertir row en dict pour .get(...)
                if isinstance(row, pd.Series):
                    row = row.to_dict()
                elif not isinstance(row, dict):
                    row = {}

                tot_val   = float(_safe_num_series(pd.Series([row.get(TOTAL, 0.0)]), 0).iloc[0])
                paye_val  = float(_safe_num_series(pd.Series([row.get("Payé", 0.0)]), 0).iloc[0])
                reste_val = float(_safe_num_series(pd.Series([row.get("Reste", 0.0)]), 0).iloc[0])

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total", _fmt_money_us(tot_val))
                c2.metric("Payé",  _fmt_money_us(paye_val))
                c3.metric("Reste", _fmt_money_us(reste_val))
                c4.metric("Visa",  _safe_str(row.get("Visa","—")))

                st.markdown("#### Détails")
                dL, dR = st.columns([1,2])
                with dL:
                    st.write({
                        "Dossier N": row.get(DOSSIER_COL,""),
                        "Nom": row.get("Nom",""),
                        "Catégorie": row.get("Categorie",""),
                        "Sous-catégorie": row.get("Sous-categorie",""),
                        "Visa": row.get("Visa",""),
                        "Date": _safe_str(row.get("Date","")),
                        "Mois": _safe_str(row.get("Mois","")),
                    })
                with dR:
                    st.write({
                        "Dossier envoyé":       int(row.get("Dossier envoyé",0) or 0),
                        "Date d'envoi":         _safe_str(row.get("Date d'envoi","")),
                        "Dossier accepté":      int(row.get("Dossier accepté",0) or 0),
                        "Date d'acceptation":   _safe_str(row.get("Date d'acceptation","")),
                        "Dossier refusé":       int(row.get("Dossier refusé",0) or 0),
                        "Date de refus":        _safe_str(row.get("Date de refus","")),
                        "Dossier annulé":       int(row.get("Dossier annulé",0) or 0),
                        "Date d'annulation":    _safe_str(row.get("Date d'annulation","")),
                        "RFE":                  int(row.get("RFE",0) or 0),
                    })

                st.markdown("#### Paiements")
                paiements = row.get("Paiements", [])
                if isinstance(paiements, str):
                    try:
                        paiements = json.loads(paiements)
                    except Exception:
                        paiements = []
                if not isinstance(paiements, list):
                    paiements = []

                if paiements:
                    dfp = pd.DataFrame(paiements)
                    if "montant" in dfp.columns:
                        dfp["montant"] = pd.to_numeric(dfp["montant"], errors="coerce").fillna(0.0)
                    st.dataframe(dfp, use_container_width=True, key=skey("cl","payhist"))
                else:
                    st.info("Aucun paiement enregistré.")

                if reste_val > 0:
                    st.markdown("##### Ajouter un paiement")
                    p1, p2, p3 = st.columns([1,1,1])
                    p_date = p1.date_input("Date", value=date.today(), key=skey("cl","paydate"))
                    p_mode = p2.selectbox("Mode", ["Cash","Chèque","CB","Virement","Venmo"], key=skey("cl","paymode"))
                    p_amt  = p3.number_input("Montant (US $)", min_value=0.0, step=10.0, value=0.0, format="%.2f",
                                             key=skey("cl","payamt"))
                    if st.button("💾 Ajouter paiement", key=skey("cl","payadd")):
                        if p_amt <= 0:
                            st.warning("Montant invalide.")
                        else:
                            paiements.append({
                                "date": _safe_str(p_date),
                                "mode": _safe_str(p_mode),
                                "montant": float(p_amt)
                            })
                            df_live = _read_clients(clients_path)
                            sel_client_id = _safe_str(row.get("ID_Client",""))
                            if "ID_Client" in df_live.columns and sel_client_id:
                                mask = (df_live["ID_Client"].astype(str) == sel_client_id)
                                if mask.any():
                                    i = df_live[mask].index[0]
                                    paid_prev  = float(_safe_num_series(pd.Series([df_live.at[i, "Payé"] if "Payé" in df_live.columns else 0]), 0).iloc[0])
                                    total_prev = float(_safe_num_series(pd.Series([df_live.at[i, TOTAL] if TOTAL in df_live.columns else 0]), 0).iloc[0])
                                    paid_new   = paid_prev + float(p_amt)
                                    reste_new  = max(0.0, total_prev - paid_new)

                                    df_live.at[i, "Paiements"] = json.dumps(paiements, ensure_ascii=False)
                                    df_live.at[i, "Payé"] = paid_new
                                    df_live.at[i, "Reste"] = reste_new

                                    _write_clients(df_live, clients_path)
                                    st.success("Paiement ajouté.")
                                    st.cache_data.clear()
                                    st.rerun()
                                else:
                                    st.error("Impossible de retrouver le client dans le fichier pour enregistrer le paiement.")
                            else:
                                st.error("ID_Client manquant — ajout impossible.")


# ==============================================
# 🧾 ONGLET : Gestion — Ajouter / Modifier / Supprimer
# ==============================================
with tabs[4]:
    st.subheader("🧾 Gestion des clients (Ajouter / Modifier / Supprimer)")
    op = st.radio("Action", ["Ajouter","Modifier","Supprimer"], horizontal=True, key=skey("crud","op"))

    df_live = _read_clients(clients_path)

    # ---------- AJOUT ----------
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")

        c1, c2, c3 = st.columns([2,1,1])
        nom  = c1.text_input("Nom", "", key=skey("add","nom"))
        dt   = c2.date_input("Date de création", value=date.today(), key=skey("add","date"))
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=int(date.today().month)-1, key=skey("add","mois"))

        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=skey("add","cat"))
        sel_sub = ""
        sel_opts: List[str] = []
        visa_final = ""
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=skey("add","sub"))
            if sel_sub:
                # options visuelles selon la structure
                preset = []
                sel_opts = render_options_for(sel_cat, sel_sub, keyprefix="add_opts", preset=preset)
                visa_final = compose_visa_label(sel_sub, sel_opts)

        f1, f2 = st.columns(2)
        honor = f1.number_input(HONO, min_value=0.0, value=0.0, step=50.0, format="%.2f", key=skey("add","honor"))
        other = f2.number_input(AUTRE, min_value=0.0, value=0.0, step=20.0, format="%.2f", key=skey("add","other"))
        comment_autre = st.text_area("Commentaires (Autres frais)", "", key=skey("add","autre_comment"))

        st.markdown("#### 📌 Statuts initiaux")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent   = s1.checkbox("Dossier envoyé", key=skey("add","sent"))
        sent_d = s1.date_input("Date d'envoi", value=None, key=skey("add","sentd"))
        acc    = s2.checkbox("Dossier accepté", key=skey("add","acc"))
        acc_d  = s2.date_input("Date d'acceptation", value=None, key=skey("add","accd"))
        ref    = s3.checkbox("Dossier refusé", key=skey("add","ref"))
        ref_d  = s3.date_input("Date de refus", value=None, key=skey("add","refd"))
        ann    = s4.checkbox("Dossier annulé", key=skey("add","ann"))
        ann_d  = s4.date_input("Date d'annulation", value=None, key=skey("add","annd"))
        rfe    = s5.checkbox("RFE", key=skey("add","rfe"))
        if rfe and not any([sent, acc, ref, ann]):
            st.warning("⚠️ La case RFE ne peut être cochée qu’avec un autre statut.")

        if st.button("💾 Enregistrer le client", key=skey("add","save")):
            if not nom:
                st.warning("Veuillez saisir le nom.")
                st.stop()
            if not (sel_cat and sel_sub):
                st.warning("Choisissez Catégorie et Sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            paye  = 0.0
            reste = max(0.0, total - paye)
            did = _make_client_id(nom, dt)
            dossier_n = _next_dossier(df_live, start=13057)

            new_row = {
                DOSSIER_COL: dossier_n,
                "ID_Client": did,
                "Nom": nom,
                "Date": dt,
                "Mois": f"{int(mois):02d}",
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": (visa_final if visa_final else sel_sub),
                HONO: float(honor),
                AUTRE: float(other),
                TOTAL: total,
                "Commentaires autres": comment_autre,
                "Payé": paye,
                "Reste": reste,
                "Paiements": [],
                "Options": sel_opts,  # on sauvegarde la liste simple
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

    # ---------- MODIFICATION ----------
    elif op == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if df_live.empty:
            st.info("Aucun client à modifier.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            sel1, sel2 = st.columns(2)
            target_name = sel1.selectbox("Nom", [""]+names, index=0, key=skey("mod","nom"))
            target_id   = sel2.selectbox("ID_Client", [""]+ids, index=0, key=skey("mod","id"))

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if not (mask is not None and mask.any()):
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()
            if isinstance(row, pd.Series):
                row = row.to_dict()

            d1, d2, d3 = st.columns([2,1,1])
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=skey("mod","nomv"))
            dt   = d2.date_input("Date de création", value=_date_for_widget(row.get("Date")), key=skey("mod","date"))
            mois_val = _safe_str(row.get("Mois","01"))
            try:
                mois_idx = max(0, min(11, int(mois_val) - 1))
            except Exception:
                mois_idx = 0
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=mois_idx, key=skey("mod","mois"))

            # Visa cascade
            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""] + cats,
                                   index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                                   key=skey("mod","cat"))

            subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
            preset_sub = _safe_str(row.get("Sous-categorie",""))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs,
                                   index=(subs.index(preset_sub)+1 if preset_sub in subs else 0),
                                   key=skey("mod","sub"))

            preset_opts = row.get("Options", [])
            if isinstance(preset_opts, str):
                try:
                    preset_opts = json.loads(preset_opts)
                except Exception:
                    preset_opts = []
            if not isinstance(preset_opts, list):
                preset_opts = []

            sel_opts = render_options_for(sel_cat, sel_sub, keyprefix="mod_opts", preset=preset_opts) if (sel_cat and sel_sub) else []
            visa_final = compose_visa_label(sel_sub, sel_opts) if sel_sub else ""

            f1, f2 = st.columns(2)
            honor = f1.number_input(HONO, min_value=0.0,
                                    value=float(_safe_num_series(pd.Series([row.get(HONO,0.0)]),0).iloc[0]),
                                    step=50.0, format="%.2f", key=skey("mod","honor"))
            other = f2.number_input(AUTRE, min_value=0.0,
                                    value=float(_safe_num_series(pd.Series([row.get(AUTRE,0.0)]),0).iloc[0]),
                                    step=20.0, format="%.2f", key=skey("mod","other"))
            comment_autre = st.text_area("Commentaires (Autres frais)",
                                         _safe_str(row.get("Commentaires autres","")),
                                         key=skey("mod","autre_comment"))

            st.markdown("#### 📌 Statuts")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent   = s1.checkbox("Dossier envoyé", value=bool(row.get("Dossier envoyé")), key=skey("mod","sent"))
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=skey("mod","sentd"))
            acc    = s2.checkbox("Dossier accepté", value=bool(row.get("Dossier accepté")), key=skey("mod","acc"))
            acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=skey("mod","accd"))
            ref    = s3.checkbox("Dossier refusé", value=bool(row.get("Dossier refusé")), key=skey("mod","ref"))
            ref_d  = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=skey("mod","refd"))
            ann    = s4.checkbox("Dossier annulé", value=bool(row.get("Dossier annulé")), key=skey("mod","ann"))
            ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=skey("mod","annd"))
            rfe    = s5.checkbox("RFE", value=bool(row.get("RFE")), key=skey("mod","rfe"))

            if st.button("💾 Enregistrer les modifications", key=skey("mod","save")):
                if not nom:
                    st.warning("Le nom est requis.")
                    st.stop()
                if not (sel_cat and sel_sub):
                    st.warning("Choisissez Catégorie et Sous-catégorie.")
                    st.stop()

                total = float(honor) + float(other)
                paye  = float(_safe_num_series(pd.Series([row.get("Payé",0.0)]),0).iloc[0])
                reste = max(0.0, total - paye)

                df_live.at[idx, "Nom"] = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(mois):02d}"
                df_live.at[idx, "Categorie"] = sel_cat
                df_live.at[idx, "Sous-categorie"] = sel_sub
                df_live.at[idx, "Visa"] = (visa_final if visa_final else sel_sub)
                df_live.at[idx, HONO] = float(honor)
                df_live.at[idx, AUTRE] = float(other)
                df_live.at[idx, TOTAL] = float(total)
                df_live.at[idx, "Commentaires autres"] = comment_autre
                df_live.at[idx, "Reste"] = reste
                df_live.at[idx, "Options"] = json.dumps(sel_opts, ensure_ascii=False)

                df_live.at[idx, "Dossier envoyé"] = 1 if sent else 0
                df_live.at[idx, "Date d'envoi"] = sent_d
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
    elif op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client à supprimer.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            sel1, sel2 = st.columns(2)
            target_name = sel1.selectbox("Nom", [""]+names, index=0, key=skey("del","nom"))
            target_id   = sel2.selectbox("ID_Client", [""]+ids, index=0, key=skey("del","id"))

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({"Dossier N": row.get(DOSSIER_COL,""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=skey("del","ok")):
                    df_new = df_live[~mask].copy()
                    _write_clients(df_new, clients_path)
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()


# ==============================================
# 📄 ONGLET : Visa (aperçu)
# ==============================================
with tabs[5]:
    st.subheader("📄 Visa — aperçu & test de sélection")
    if df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True, key=skey("visa","raw"))
        cats = sorted(list(visa_map.keys()))
        cat = st.selectbox("Catégorie", [""]+cats, index=0, key=skey("visa","cat"))
        if cat:
            subs = sorted(list(visa_map.get(cat, {}).keys()))
            sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=skey("visa","sub"))
            if sub:
                st.caption("Options (cases à cocher)")
                sel_opts = render_options_for(cat, sub, keyprefix="visa_test", preset=[])
                st.write("Résultat Visa :", compose_visa_label(sub, sel_opts))


# ==============================================
# 💾 Export global (Clients + Visa) & Rappel fichiers
# ==============================================
st.markdown("---")
st.markdown("### 💾 Export & Rappel fichiers")
cL, cR = st.columns([1,1])

with cL:
    if clients_path and os.path.exists(clients_path):
        with open(clients_path, "rb") as f:
            st.download_button("⬇️ Télécharger Clients.xlsx", f.read(),
                               file_name=os.path.basename(clients_path),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key=skey("dl","clients"))
    else:
        st.caption("Aucun fichier Clients en mémoire.")

with cR:
    if visa_path and os.path.exists(visa_path):
        with open(visa_path, "rb") as f:
            st.download_button("⬇️ Télécharger Visa.xlsx", f.read(),
                               file_name=os.path.basename(visa_path),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key=skey("dl","visa"))
    else:
        st.caption("Aucun fichier Visa en mémoire.")