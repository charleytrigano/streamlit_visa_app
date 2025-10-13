# ================================
# 🛂 VISA MANAGER — PARTIE 1/4
# ================================
from __future__ import annotations

import os, json, re, zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple, Any

import pandas as pd
import streamlit as st

# ---------- Config & styles ----------
st.set_page_config(page_title="Visa Manager", page_icon="🛂", layout="wide")

st.markdown("""
<style>
.small-metrics .stMetric { padding: 0.25rem 0.5rem !important; }
.small-metrics .stMetric label, .small-metrics .stMetric span { font-size: 0.8rem !important; }
.compact-input .stTextInput input,
.compact-input .stNumberInput input,
.compact-input .stSelectbox div[data-baseweb="select"] { font-size: 0.85rem !important; height: 2.0rem !important; }
.compact-textarea textarea { font-size: 0.9rem !important; }
</style>
""", unsafe_allow_html=True)

SID = "vm"  # suffixe de clés Streamlit pour éviter collisions

# ---------- Constantes colonnes ----------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

DOSSIER_COL = "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"

PAY_MODES = ["Chèque", "CB", "Cash", "Virement", "Venmo"]

# ---------- Helpers sûrs ----------
def _safe_str(x: Any) -> str:
    try:
        if pd.isna(x): return ""
    except Exception:
        pass
    return str(x)

def _fmt_money_us(v: float | int | str) -> str:
    try:
        f = float(v)
    except Exception:
        f = 0.0
    return f"${f:,.2f}"

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index, dtype=float)
    s = df[col]
    if pd.api.types.is_numeric_dtype(s):
        return s.fillna(0.0).astype(float)
    # nettoie toute chaîne ($, espaces, etc.)
    s = s.astype(str).str.replace(r"[^\d,.\-]", "", regex=True)
    # virgule -> point si pas déjà
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _date_for_widget(v: Any) -> date | None:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.isna(dt): return None
        return dt.date()
    except Exception:
        return None

def _make_client_id(name: str, d: date) -> str:
    base = re.sub(r"[^a-z0-9\-]+", "-", _safe_str(name).lower()).strip("-")
    if not base:
        base = "client"
    return f"{base}-{d:%Y%m%d}"

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if df.empty or DOSSIER_COL not in df.columns:
        return start
    vals = pd.to_numeric(df[DOSSIER_COL], errors="coerce")
    m = pd.to_numeric(vals[~vals.isna()], errors="coerce")
    if m.empty:
        return start
    try:
        return max(int(m.max()) + 1, start)
    except Exception:
        return start

# ---------- Persistance des derniers chemins ----------
LAST_PATHS_FILE = ".cache_visamanager.json"

def _save_last_paths(clients_path: str | None, visa_path: str | None) -> None:
    try:
        data = {"clients_path": clients_path or "", "visa_path": visa_path or ""}
        with open(LAST_PATHS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _load_last_paths() -> tuple[str | None, str | None]:
    try:
        with open(LAST_PATHS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        cp = data.get("clients_path") or None
        vp = data.get("visa_path") or None
        if cp and not os.path.exists(cp): cp = None
        if vp and not os.path.exists(vp): vp = None
        return cp, vp
    except Exception:
        return None, None

# ---------- Lecture/écriture Excel ----------
@st.cache_data(show_spinner=False)
def read_excel_file(path: str, sheet: str | None = None) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet) if sheet else pd.read_excel(path)

def write_clients_excel(path: str, df: pd.DataFrame, visa_keep_path: str | None = None) -> None:
    # Sauvegarde le DF Clients, et conserve Visa si path unique
    if visa_keep_path is None:
        # fichier Clients seul
        with pd.ExcelWriter(path, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        return

    # un seul fichier (2 onglets)
    try:
        v = pd.read_excel(visa_keep_path, sheet_name=SHEET_VISA)
    except Exception:
        v = pd.DataFrame()
    with pd.ExcelWriter(path, engine="openpyxl") as wr:
        df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        if not v.empty:
            v.to_excel(wr, sheet_name=SHEET_VISA, index=False)

def ensure_clients_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Construit les colonnes si manquantes
    base_cols = [
        DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
        "Categorie", "Sous-categorie", "Visa",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Paiements", "Options",
        "Commentaires autres frais",
        "Dossier envoyé", "Date d'envoi",
        "Dossier accepté", "Date d'acceptation",
        "Dossier refusé", "Date de refus",
        "Dossier annulé", "Date d'annulation",
        "RFE",
    ]
    for c in base_cols:
        if c not in df.columns:
            df[c] = pd.NA

    # numérise montants
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = _safe_num_series(df, c)

    # normalise Mois (MM)
    if "Mois" in df.columns:
        df["Mois"] = df["Mois"].apply(lambda x: f"{int(_safe_str(x) or '1'):02d}" if _safe_str(x) else "")
    # _Année_ / _MoisNum_ pour tri
    if "Date" in df.columns:
        dd = pd.to_datetime(df["Date"], errors="coerce")
        df["_Année_"] = dd.dt.year
        df["_MoisNum_"] = dd.dt.month
    else:
        df["_Année_"] = pd.NA
        df["_MoisNum_"] = pd.NA

    # total si absent
    df[TOTAL] = df[HONO] + df[AUTRE]
    # reste si absent
    if "Payé" not in df.columns: df["Payé"] = 0.0
    df["Reste"] = (df[TOTAL] - _safe_num_series(df, "Payé")).clip(lower=0.0)
    return df

# ---------- Parsing onglet Visa → carte {Categorie: {Sous-categorie: {options: [...]}}} ----------
@st.cache_data(show_spinner=False)
def build_visa_map(visa_df: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    vm: Dict[str, Dict[str, Dict[str, Any]]] = {}
    if visa_df.empty:
        return vm

    # colonnes minimales
    cols = [c for c in visa_df.columns]
    # On attend au minimum "Categorie" et "Sous-categorie"
    if "Categorie" not in cols or "Sous-categorie" not in cols:
        return vm

    option_cols = [c for c in cols if c not in ("Categorie", "Sous-categorie")]

    for _, row in visa_df.iterrows():
        cat = _safe_str(row.get("Categorie")).strip()
        sub = _safe_str(row.get("Sous-categorie")).strip()
        if not cat or not sub:
            continue

        opts_available: List[str] = []
        for oc in option_cols:
            val = row.get(oc)
            try:
                ok = float(val) == 1.0
            except Exception:
                ok = str(val).strip() == "1"
            if ok:
                opts_available.append(oc)

        if cat not in vm:
            vm[cat] = {}
        vm[cat][sub] = {
            "options": opts_available,      # ex: ["COS","EOS", ...]
        }
    return vm

def render_options_for_sub(vm: Dict[str, Dict[str, Dict[str, Any]]],
                           cat: str, sub: str, keyprefix: str,
                           preset: Dict[str, Any] | None = None) -> Tuple[str, Dict[str, Any], str]:
    """
    Affiche dynamiquement les cases à cocher des options correspondant à la (cat, sub).
    Retourne (visa_final, options_dict, info)
    - visa_final : libellé Visa = f"{sub} {opt}" si choix exclusif (COS/EOS), sinon sub
    - options_dict : {"exclusive": "... ou None", "options": ["opt1","opt2",...]}
    """
    info = ""
    options_dict = {"exclusive": None, "options": []}
    visa_final = sub

    if cat not in vm or sub not in vm[cat]:
        st.warning("Aucune option disponible pour cette sous-catégorie.")
        return visa_final, options_dict, info

    opts = vm[cat][sub]["options"] or []
    if not opts:
        st.info("Cette sous-catégorie n’a pas d’options supplémentaires.")
        return visa_final, options_dict, info

    # Exclusif si le duo COS/EOS est présent → on force radio pour ce duo
    exclusive_choice = None
    if "COS" in opts or "EOS" in opts:
        found = [o for o in ["COS", "EOS"] if o in opts]
        if len(found) >= 1:
            preset_ex = None
            if preset and isinstance(preset, dict):
                preset_ex = preset.get("exclusive")
            exclusive_choice = st.radio(
                "Choix exclusif (COS / EOS)", found,
                index=(found.index(preset_ex) if preset_ex in found else 0),
                key=f"{keyprefix}_excl"
            )
            options_dict["exclusive"] = exclusive_choice
            visa_final = f"{sub} {exclusive_choice}"

    # Les autres options (en cases à cocher)
    other_opts = [o for o in opts if o not in (["COS","EOS"])]
    chosen_list: List[str] = []
    if other_opts:
        st.markdown("Options supplémentaires :")
        for oc in other_opts:
            preset_on = False
            if preset and isinstance(preset, dict):
                p_opts = preset.get("options") or []
                preset_on = oc in p_opts
            on = st.checkbox(oc, value=preset_on, key=f"{keyprefix}_{oc}")
            if on: chosen_list.append(oc)
    options_dict["options"] = chosen_list

    return visa_final, options_dict, info



# ================================
# 🛂 VISA MANAGER — PARTIE 2/4
# ================================

st.title("🛂 Visa Manager")

# --------- Chargement fichiers ---------
st.markdown("## 📂 Fichiers")

mode = st.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"],
                horizontal=True, key=f"mode_{SID}")

last_clients, last_visa = _load_last_paths()

clients_path: str | None = None
visa_path: str | None = None

c1, c2 = st.columns(2)
with c1:
    if mode == "Deux fichiers (Clients & Visa)":
        up_c = st.file_uploader("Clients (xlsx)", type=["xlsx"], key=f"fu_clients_{SID}")
        if up_c is not None:
            clients_path = os.path.join(".", f"_clients_{up_c.name}")
            with open(clients_path, "wb") as f:
                f.write(up_c.read())
        elif last_clients:
            st.caption(f"Dernier Clients utilisé : {last_clients}")
            clients_path = last_clients
    else:
        up_one = st.file_uploader("Fichier unique (2 onglets)", type=["xlsx"], key=f"fu_one_{SID}")
        if up_one is not None:
            one_path = os.path.join(".", f"_both_{up_one.name}")
            with open(one_path, "wb") as f:
                f.write(up_one.read())
            clients_path = one_path  # même fichier
            visa_path    = one_path
        elif last_clients and last_visa and (last_clients == last_visa):
            st.caption(f"Dernier fichier unique : {last_clients}")
            clients_path = last_clients
            visa_path    = last_visa

with c2:
    if mode == "Deux fichiers (Clients & Visa)":
        up_v = st.file_uploader("Visa (xlsx)", type=["xlsx"], key=f"fu_visa_{SID}")
        if up_v is not None:
            visa_path = os.path.join(".", f"_visa_{up_v.name}")
            with open(visa_path, "wb") as f:
                f.write(up_v.read())
        elif last_visa:
            st.caption(f"Dernier Visa utilisé : {last_visa}")
            visa_path = last_visa

# Mémorise chemins
if clients_path or visa_path:
    _save_last_paths(clients_path, visa_path)

# --------- Lecture des données ----------
df_clients_raw = pd.DataFrame()
df_visa_raw    = pd.DataFrame()

if clients_path:
    try:
        if visa_path and (visa_path == clients_path):
            # un fichier, 2 onglets
            df_clients_raw = read_excel_file(clients_path, SHEET_CLIENTS)
            df_visa_raw    = read_excel_file(clients_path, SHEET_VISA)
        else:
            # deux fichiers
            df_clients_raw = read_excel_file(clients_path, SHEET_CLIENTS) if mode != "Un seul fichier (2 onglets)" else read_excel_file(clients_path, SHEET_CLIENTS)
            if visa_path:
                df_visa_raw = read_excel_file(visa_path, SHEET_VISA)
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")

# Normalisation clients
df_all = pd.DataFrame()
if not df_clients_raw.empty:
    df_all = ensure_clients_columns(df_clients_raw.copy())

# Carte Visa
visa_map: Dict[str, Dict[str, Dict[str, Any]]] = {}
if not df_visa_raw.empty:
    # on s’assure que les deux colonnes existent
    if "Categorie" in df_visa_raw.columns and "Sous-categorie" in df_visa_raw.columns:
        visa_map = build_visa_map(df_visa_raw.copy())

# Tabs principaux
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "📄 Visa (aperçu)"])

# --------- Visa aperçu ---------
with tabs[4]:
    st.subheader("📄 Visa (aperçu)")
    if df_visa_raw.empty:
        st.info("Aucune donnée Visa.")
    else:
        # Filtres simples
        cats = sorted(df_visa_raw["Categorie"].dropna().astype(str).unique().tolist())
        sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=f"v_cat_{SID}")
        if sel_cat:
            subs = sorted(df_visa_raw.loc[df_visa_raw["Categorie"].astype(str) == sel_cat, "Sous-categorie"].dropna().astype(str).unique().tolist())
        else:
            subs = sorted(df_visa_raw["Sous-categorie"].dropna().astype(str).unique().tolist())
        sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=f"v_sub_{SID}")

        view = df_visa_raw.copy()
        if sel_cat:
            view = view[view["Categorie"].astype(str) == sel_cat]
        if sel_sub:
            view = view[view["Sous-categorie"].astype(str) == sel_sub]

        st.dataframe(view.reset_index(drop=True), use_container_width=True)

# --------- Dashboard ---------
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        # listes de filtres
        years  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1, 13)]
        cats   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        cA, cB, cC, cD, cE = st.columns(5)
        f_years = cA.multiselect("Année", years, default=[], key=f"d_years_{SID}")
        f_month = cB.multiselect("Mois (MM)", months, default=[], key=f"d_months_{SID}")
        f_cat   = cC.multiselect("Catégorie", cats, default=[], key=f"d_cat_{SID}")
        f_sub   = cD.multiselect("Sous-catégorie", subs, default=[], key=f"d_sub_{SID}")
        f_visa  = cE.multiselect("Visa", visas, default=[], key=f"d_visa_{SID}")

        dfD = df_all.copy()
        for coln in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if coln in dfD.columns:
                dfD[coln] = _safe_num_series(dfD, coln)

        if f_years: dfD = dfD[dfD["_Année_"].isin(f_years)]
        if f_month: dfD = dfD[dfD["Mois"].astype(str).isin(f_month)]
        if f_cat:   dfD = dfD[dfD["Categorie"].astype(str).isin(f_cat)]
        if f_sub:   dfD = dfD[dfD["Sous-categorie"].astype(str).isin(f_sub)]
        if f_visa:  dfD = dfD[dfD["Visa"].astype(str).isin(f_visa)]

        # KPI compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfD)}")
        k2.metric("Honoraires", _fmt_money_us(float(dfD[HONO].sum())) if HONO in dfD else "$0")
        k3.metric("Payé",      _fmt_money_us(float(dfD["Payé"].sum())) if "Payé" in dfD else "$0")
        k4.metric("Reste",     _fmt_money_us(float(dfD["Reste"].sum())) if "Reste" in dfD else "$0")
        st.markdown('</div>', unsafe_allow_html=True)

        # Détails
        view = dfD.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in view.columns:
                view[c] = _safe_num_series(view, c).map(_fmt_money_us)
        if "Date" in view.columns:
            try:
                view["Date"] = pd.to_datetime(view["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                view["Date"] = view["Date"].astype(str)

        show_cols = [c for c in [
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
            "Date", "Mois", HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in view.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_keys) if sort_keys else view
        # enlève doublons éventuels
        view_sorted = view_sorted.loc[:, ~view_sorted.columns.duplicated()].copy()

        st.dataframe(
            view_sorted[show_cols].reset_index(drop=True),
            use_container_width=True,
            key=f"d_tbl_{SID}"
        )



# ================================
# 🛂 VISA MANAGER — PARTIE 3/4
# ================================

# --------- Analyses ---------
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
        for coln in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if coln in dfA.columns:
                dfA[coln] = _safe_num_series(dfA, coln)

        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPI compacts
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money_us(float(dfA[HONO].sum())) if HONO in dfA else "$0")
        k3.metric("Payé",      _fmt_money_us(float(dfA["Payé"].sum())) if "Payé" in dfA else "$0")
        k4.metric("Reste",     _fmt_money_us(float(dfA["Reste"].sum())) if "Reste" in dfA else "$0")
        st.markdown('</div>', unsafe_allow_html=True)

        # % par catégorie / sous-cat
        st.markdown("#### Répartition & %")
        cL, cR = st.columns(2)
        if not dfA.empty and "Categorie" in dfA.columns:
            vc = (dfA.groupby("Categorie", as_index=False)
                        .agg(N=("Categorie","size"), Honoraires=(HONO,"sum")))
            vc["% Dossiers"] = (vc["N"] / (vc["N"].sum() or 1) * 100).round(1)
            with cL:
                st.dataframe(vc.sort_values("N", ascending=False), use_container_width=True, height=260)
        if not dfA.empty and "Sous-categorie" in dfA.columns:
            vs = (dfA.groupby("Sous-categorie", as_index=False)
                        .agg(N=("Sous-categorie","size"), Honoraires=(HONO,"sum")))
            vs["% Dossiers"] = (vs["N"] / (vs["N"].sum() or 1) * 100).round(1)
            with cR:
                st.dataframe(vs.sort_values("N", ascending=False).head(25), use_container_width=True, height=260)

        # Graph : dossiers par catégorie
        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("#### 📊 Dossiers par catégorie")
            g1 = (dfA.groupby("Categorie", as_index=False).size().rename(columns={"size":"Nombre"}))
            st.bar_chart(g1.set_index("Categorie"))

        # Graph : honoraires par mois
        if not dfA.empty and "Mois" in dfA.columns and HONO in dfA.columns:
            st.markdown("#### 📈 Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = (tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois"))
            st.line_chart(gm.set_index("Mois"))

        # Comparaison périodes A vs B
        st.markdown("#### 🔁 Comparaison de périodes (A vs B)")
        ca1, ca2, cb1, cb2 = st.columns(4)
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=f"cmp_ya_{SID}")
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=f"cmp_ma_{SID}")
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=f"cmp_yb_{SID}")
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=f"cmp_mb_{SID}")

        def _filter_period(base, ys, ms):
            d = base.copy()
            if ys: d = d[d["_Année_"].isin(ys)]
            if ms: d = d[d["Mois"].astype(str).isin(ms)]
            return d

        A = _filter_period(df_all, pa_years, pa_month)
        B = _filter_period(df_all, pb_years, pb_month)
        for ddf in (A, B):
            for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                if c in ddf.columns: ddf[c] = _safe_num_series(ddf, c)

        cpa, cpb = st.columns(2)
        with cpa:
            st.metric("A — Dossiers", f"{len(A)}")
            st.metric("A — Honoraires", _fmt_money_us(float(A[HONO].sum())) if HONO in A else "$0")
        with cpb:
            st.metric("B — Dossiers", f"{len(B)}")
            st.metric("B — Honoraires", _fmt_money_us(float(B[HONO].sum())) if HONO in B else "$0")

        if not (A.empty and B.empty):
            st.markdown("##### Comparaison par mois")
            def _mk_month_series(d):
                if d.empty or "Mois" not in d.columns or HONO not in d.columns:
                    return pd.DataFrame({"Mois": [], "Honoraires":[]})
                t = d.copy()
                t["Mois"] = t["Mois"].astype(str)
                return (t.groupby("Mois", as_index=False)[HONO].sum()
                          .reindex([f"{m:02d}" for m in range(1,13)], fill_value=0)
                          .rename(columns={HONO:"Honoraires"}))
            AA = _mk_month_series(A); AA["Periode"]="A"
            BB = _mk_month_series(B); BB["Periode"]="B"
            comp = pd.concat([AA, BB], ignore_index=True)
            if not comp.empty:
                wide = comp.pivot_table(index="Mois", columns="Periode", values="Honoraires", fill_value=0)
                st.bar_chart(wide)

        # Détails
        st.markdown("#### 🧾 Détails des dossiers filtrés")
        det = dfA.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in det.columns: det[c] = _safe_num_series(det, c).map(_fmt_money_us)
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
        det_sorted = det_sorted.loc[:, ~det_sorted.columns.duplicated()]
        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"a_tbl_{SID}")

# --------- Escrow ---------
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        for c in [TOTAL,"Payé","Reste"]:
            if c in dfE.columns: dfE[c] = _safe_num_series(dfE, c)

        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money_us(float(dfE[TOTAL].sum())))
        t2.metric("Payé",         _fmt_money_us(float(dfE["Payé"].sum())))
        t3.metric("Reste",        _fmt_money_us(float(dfE["Reste"].sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        agg = dfE.groupby("Categorie", as_index=False)[[TOTAL,"Payé","Reste"]].sum()
        agg["% Payé"] = ((agg["Payé"] / agg[TOTAL]).fillna(0.0) * 100).round(1)
        st.dataframe(agg.sort_values(TOTAL, ascending=False), use_container_width=True)




# ==============================================
# 📈 ONGLET : Analyses (filtres + KPI + graphs + comparaison + détail)
# ==============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        # Uniques
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all.get("_Année_", pd.Series(dtype=float)), errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_all.get("Categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        subsA   = sorted(df_all.get("Sous-categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        visasA  = sorted(df_all.get("Visa", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_years_{SID}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"a_months_{SID}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"a_cats_{SID}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"a_subs_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_visas_{SID}")

        dfA = df_all.copy()
        # Numériques sûrs
        for coln in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if coln in dfA.columns:
                dfA[coln] = _safe_num_series(dfA, coln)

        # Filtres
        if fy: dfA = dfA[dfA.get("_Année_", "").isin(fy)]
        if fm: dfA = dfA[dfA.get("Mois", "").astype(str).isin(fm)]
        if fc: dfA = dfA[dfA.get("Categorie", "").astype(str).isin(fc)]
        if fs: dfA = dfA[dfA.get("Sous-categorie", "").astype(str).isin(fs)]
        if fv: dfA = dfA[dfA.get("Visa", "").astype(str).isin(fv)]

        # KPI (compacts si CSS ajouté en amont)
        st.markdown('<div class="small-metrics">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money_us(float(dfA[HONO].sum())) if HONO in dfA else "$0")
        k3.metric("Payé",      _fmt_money_us(float(dfA.get("Payé", pd.Series(dtype=float)).sum())))
        k4.metric("Reste",     _fmt_money_us(float(dfA.get("Reste", pd.Series(dtype=float)).sum())))
        st.markdown('</div>', unsafe_allow_html=True)

        # % par catégorie / sous-catégorie
        st.markdown("#### Répartition & %")
        cL, cR = st.columns(2)
        if not dfA.empty and "Categorie" in dfA.columns:
            vc = (dfA.groupby("Categorie", as_index=False)
                    .agg(N=("Categorie","size"),
                         Honoraires=(HONO,"sum") if HONO in dfA else (("Categorie","size"))))
            totN = vc["N"].sum() or 1
            vc["% Dossiers"] = (vc["N"]/totN*100).round(1)
            with cL:
                st.dataframe(vc.sort_values("N", ascending=False), use_container_width=True, height=260)
        if not dfA.empty and "Sous-categorie" in dfA.columns:
            vs = (dfA.groupby("Sous-categorie", as_index=False)
                    .agg(N=("Sous-categorie","size"),
                         Honoraires=(HONO,"sum") if HONO in dfA else (("Sous-categorie","size"))))
            totNs = vs["N"].sum() or 1
            vs["% Dossiers"] = (vs["N"]/totNs*100).round(1)
            with cR:
                st.dataframe(vs.sort_values("N", ascending=False).head(25), use_container_width=True, height=260)

        # Graph : Dossiers par catégorie
        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("#### 📊 Dossiers par catégorie")
            g1 = (dfA.groupby("Categorie", as_index=False).size().rename(columns={"size":"Nombre"}))
            st.bar_chart(g1.set_index("Categorie"))

        # Graph : Honoraires par mois
        if not dfA.empty and "Mois" in dfA.columns and HONO in dfA.columns:
            st.markdown("#### 📈 Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = (tmp.groupby("Mois", as_index=False)[HONO].sum()
                    .reindex([f"{m:02d}" for m in range(1,13)], fill_value=0)
                    .sort_values("Mois"))
            st.line_chart(gm.set_index("Mois"))

        # Comparaison période A vs B
        st.markdown("#### 🔁 Comparaison de périodes (A vs B)")
        ca1, ca2, cb1, cb2 = st.columns(4)
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=f"cmp_ya_{SID}")
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=f"cmp_ma_{SID}")
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=f"cmp_yb_{SID}")
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=f"cmp_mb_{SID}")

        def _filter_period(base, ys, ms):
            d = base.copy()
            for coln in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                if coln in d.columns:
                    d[coln] = _safe_num_series(d, coln)
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

        # Comparaison par mois
        st.markdown("##### Comparaison par mois")
        def _mk_month_series(d):
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

        # Détails filtrés
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
                     use_container_width=True, key=f"a_tbl_{SID}")


# ==============================================
# 🏦 ONGLET : Escrow — synthèse compacte
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

        agg = dfE.groupby("Categorie", as_index=False)[[c for c in [TOTAL,"Payé","Reste"] if c in dfE.columns]].sum()
        if TOTAL in agg and "Payé" in agg:
            agg["% Payé"] = ((agg["Payé"] / agg[TOTAL]).fillna(0.0)*100).round(1)
        st.dataframe(agg.sort_values(by=TOTAL if TOTAL in agg else agg.columns[1], ascending=False),
                     use_container_width=True)




