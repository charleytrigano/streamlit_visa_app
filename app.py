# app.py
from __future__ import annotations

import json
import zipfile
from io import BytesIO
from uuid import uuid4
from datetime import date, datetime

import pandas as pd
import streamlit as st

# ==============================
# Constantes
# ==============================
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"

STATUS_FIELDS = [
    ("Dossier envoyé",      "Date d'envoi"),
    ("Dossier accepté",     "Date d'acceptation"),
    ("Dossier refusé",      "Date de refus"),
    ("Dossier annulé",      "Date d'annulation"),
]
RFE_FIELD = "RFE"

REQUIRED_CLIENTS_COLS = [
    "Dossier N", "ID_Client", "Nom", "Date", "Mois",
    "Categorie", "Sous-categorie", "Visa",
    HONO, AUTRE, TOTAL, "Payé", "Reste", "Paiements", "Options", "Notes",
] + [s for (s, _) in STATUS_FIELDS] + [d for (_, d) in STATUS_FIELDS] + [RFE_FIELD]

# ==============================
# Helpers sûrs
# ==============================
def _safe_str(x) -> str:
    try:
        return "" if pd.isna(x) else str(x)
    except Exception:
        return ""

def _to_float(x, default=0.0) -> float:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return float(default)
        if isinstance(x, (int, float)):
            return float(x)
        s = str(x)
        s = s.replace("\u202f", "").replace(" ", "").replace(",", ".")
        s = "".join(ch for ch in s if ch.isdigit() or ch in ".-")
        return float(s) if s else float(default)
    except Exception:
        return float(default)

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df), index=df.index, dtype=float)
    return df[col].apply(_to_float)

def _fmt_money(v: float) -> str:
    try:
        return f"${v:,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(val) -> date:
    """Toujours une vraie date pour st.date_input (fallback=aujourd'hui)."""
    try:
        if val is None:
            return date.today()
        if isinstance(val, date) and not isinstance(val, datetime):
            return val
        if isinstance(val, datetime):
            return val.date()
        ts = pd.to_datetime(val, errors="coerce")
        if pd.isna(ts):
            return date.today()
        return ts.to_pydatetime().date()
    except Exception:
        return date.today()

def _date_or_none(val):
    """date/datetime → date ; chaîne → date ; sinon None"""
    try:
        if val is None:
            return None
        if isinstance(val, date) and not isinstance(val, datetime):
            return val
        if isinstance(val, datetime):
            return val.date()
        ts = pd.to_datetime(val, errors="coerce")
        if pd.isna(ts):
            return None
        return ts.to_pydatetime().date()
    except Exception:
        return None

def _json_loads_or(v, fallback):
    try:
        if isinstance(v, (list, dict)):
            return v
        s = _safe_str(v)
        if not s:
            return fallback
        return json.loads(s)
    except Exception:
        return fallback

def _month_index(val) -> int:
    """
    Convertit une valeur quelconque ('07', 7, '', NaN, '2025-07-01') en index 0..11 pour le selectbox Mois.
    Repli sûr sur 0 (→ '01').
    """
    try:
        s = _safe_str(val).strip()
        m = None
        if s.isdigit():
            m = int(s)
        else:
            ts = pd.to_datetime(s, errors="coerce")
            if isinstance(ts, pd.Timestamp) and not pd.isna(ts):
                m = int(ts.month)
        if m is None:
            m = int(pd.to_numeric(s, errors="coerce"))
        if not (1 <= m <= 12):
            m = 1
        return m - 1
    except Exception:
        return 0

def next_dossier_number(df: pd.DataFrame, start=13057) -> int:
    try:
        if "Dossier N" not in df.columns or df.empty:
            return start
        v = pd.to_numeric(df["Dossier N"], errors="coerce").dropna()
        if v.empty:
            return start
        return int(v.max()) + 1
    except Exception:
        return start

def make_client_id(nom: str, d: date) -> str:
    base = _safe_str(nom).strip().replace(" ", "").replace("/", "-")
    if not base:
        base = "CLIENT"
    return f"{base}-{d:%Y%m%d}"

# ==============================
# Mémoire fichiers (restaurer dernier choix)
# ==============================
SID = st.session_state.get("SID") or str(uuid4())[:8]
st.session_state["SID"] = SID
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

# chemins par défaut + restauration
clients_path = st.session_state.get("clients_path", "donnees_visa_clients1_adapte.xlsx")
visa_path    = st.session_state.get("visa_path",    "donnees_visa_clients1.xlsx")

# ==============================
# Sidebar : chargement fichiers & mémoire
# ==============================
with st.sidebar:
    st.header("📂 Fichiers")
    st.caption("Les fichiers choisis ici sont mémorisés tant que la session est ouverte.")
    c1, c2 = st.columns(2)
    with c1:
        st.text("Clients")
        upC = st.file_uploader(" ", type=["xlsx"], key=f"upC_{SID}", label_visibility="collapsed")
    with c2:
        st.text("Visa")
        upV = st.file_uploader("   ", type=["xlsx"], key=f"upV_{SID}", label_visibility="collapsed")

    if upC is not None:
        st.session_state["clients_bin"] = upC.read()
        st.session_state["clients_name"] = upC.name
        clients_path = upC.name
        st.session_state["clients_path"] = clients_path

    if upV is not None:
        st.session_state["visa_bin"] = upV.read()
        st.session_state["visa_name"] = upV.name
        visa_path = upV.name
        st.session_state["visa_path"] = visa_path

    if st.button("🧹 Oublier les fichiers chargés", key=f"clr_{SID}"):
        for k in ["clients_bin","clients_name","clients_path","visa_bin","visa_name","visa_path"]:
            st.session_state.pop(k, None)
        clients_path = "donnees_visa_clients1_adapte.xlsx"
        visa_path    = "donnees_visa_clients1.xlsx"
        st.success("Mémoire nettoyée (noms par défaut).")
        st.rerun()

# ==============================
# Lecture/écriture Excel (depuis binaire si présent)
# ==============================
@st.cache_data(show_spinner=False)
def read_excel_maybe_bin(bin_bytes: bytes | None, fallback_path: str, sheet: str) -> pd.DataFrame:
    if bin_bytes:
        return pd.read_excel(BytesIO(bin_bytes), sheet_name=sheet)
    else:
        return pd.read_excel(fallback_path, sheet_name=sheet)

def write_clients_maybe_bin(df: pd.DataFrame):
    """Écrit la feuille Clients dans le binaire si présent, sinon sur le disque (même nom)."""
    if st.session_state.get("clients_bin") is not None:
        try:
            existing = {}
            try:
                with BytesIO(st.session_state["clients_bin"]) as bio:
                    xls = pd.ExcelFile(bio)
                    for sh in xls.sheet_names:
                        existing[sh] = pd.read_excel(BytesIO(st.session_state["clients_bin"]), sheet_name=sh)
            except Exception:
                pass
            existing[SHEET_CLIENTS] = df.copy()
            with BytesIO() as outb:
                with pd.ExcelWriter(outb, engine="openpyxl") as wr:
                    for sh, d in existing.items():
                        d.to_excel(wr, sheet_name=sh, index=False)
                st.session_state["clients_bin"] = outb.getvalue()
            return True, None
        except Exception as e:
            return False, str(e)
    else:
        try:
            existing = {}
            try:
                xls = pd.ExcelFile(clients_path)
                for sh in xls.sheet_names:
                    existing[sh] = pd.read_excel(clients_path, sheet_name=sh)
            except Exception:
                pass
            existing[SHEET_CLIENTS] = df.copy()
            with pd.ExcelWriter(clients_path, engine="openpyxl") as wr:
                for sh, d in existing.items():
                    d.to_excel(wr, sheet_name=sh, index=False)
            return True, None
        except Exception as e:
            return False, str(e)

@st.cache_data(show_spinner=False)
def read_clients() -> pd.DataFrame:
    try:
        return read_excel_maybe_bin(st.session_state.get("clients_bin"), clients_path, SHEET_CLIENTS)
    except Exception:
        return pd.DataFrame(columns=REQUIRED_CLIENTS_COLS)

@st.cache_data(show_spinner=False)
def read_visa() -> pd.DataFrame:
    return read_excel_maybe_bin(st.session_state.get("visa_bin"), visa_path, SHEET_VISA)

def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    for col in REQUIRED_CLIENTS_COLS:
        if col not in df.columns:
            df[col] = None
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = _safe_num_series(df, c)
    df["Paiements"] = df["Paiements"].apply(lambda x: _json_loads_or(x, []))
    df["Options"]   = df["Options"].apply(lambda x: _json_loads_or(x, {"options": [], "exclusive": None}))
    try:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    except Exception:
        pass
    df["_Année_"]   = pd.to_datetime(df["Date"], errors="coerce").dt.year.astype("Int64")
    df["_MoisNum_"] = pd.to_datetime(df["Date"], errors="coerce").dt.month.astype("Int64")
    df["Mois"] = df["Mois"].apply(lambda m: f"{int(m):02d}" if _safe_str(m).strip().isdigit() else _safe_str(m))
    # recalcule total / reste si manquants
    df[TOTAL] = _safe_num_series(df, HONO) + _safe_num_series(df, AUTRE)
    df["Reste"] = (_safe_num_series(df, TOTAL) - _safe_num_series(df, "Payé")).clip(lower=0)
    return df

def build_visa_map(df_visa: pd.DataFrame) -> dict:
    """
    Construit la structure:
    visa_map[cat][sub] = {
        "options": [liste des intitulés d'options disponibles (cellule == 1)],
        "all_options": [toutes les colonnes options existantes]
    }
    Les colonnes 'Categorie' et 'Sous-categorie' sont obligatoires.
    """
    if df_visa.empty:
        return {}
    if "Categorie" not in df_visa.columns or "Sous-categorie" not in df_visa.columns:
        return {}
    option_cols = [c for c in df_visa.columns if c not in ["Categorie", "Sous-categorie"]]
    result = {}
    for _, row in df_visa.iterrows():
        cat  = _safe_str(row["Categorie"])
        sub  = _safe_str(row["Sous-categorie"])
        if not cat or not sub:
            continue
        opts_available = []
        for oc in option_cols:
            val = row.get(oc, None)
            if _to_float(val, 0.0) == 1.0:
                opts_available.append(oc)
        result.setdefault(cat, {})
        result[cat][sub] = {"options": sorted(opts_available), "all_options": option_cols}
    return result

def render_option_checkboxes(options: list[str], keyprefix: str, preselected: list[str] | None = None) -> list[str]:
    sel = []
    pre = set(preselected or [])
    n = max(1, min(4, len(options))) if options else 1
    cols = st.columns(n) if options else [st]
    for i, opt in enumerate(options):
        col = cols[i % n]
        with col:
            v = st.checkbox(opt, value=(opt in pre), key=f"{keyprefix}_{i}")
        if v:
            sel.append(opt)
    return sel

def compute_visa_string(sub: str, options_sel: list[str]) -> str:
    if options_sel:
        return f"{sub} " + " ".join(options_sel)
    return sub

# ==============================
# Chargements initiaux
# ==============================
df_clients_raw = read_clients()
df_clients = normalize_clients(df_clients_raw.copy())

try:
    df_visa = read_visa()
except Exception as e:
    st.error(f"Impossible de lire le fichier Visa : {e}")
    df_visa = pd.DataFrame(columns=["Categorie","Sous-categorie"])

visa_map = build_visa_map(df_visa)

# ==============================
# Téléchargements rapides (sidebar)
# ==============================
with st.sidebar:
    st.header("📥 Téléchargements")
    # Clients actuel
    exp_clients = df_clients.copy()
    with BytesIO() as outc:
        with pd.ExcelWriter(outc, engine="openpyxl") as wr:
            exp_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        st.download_button("⬇️ Clients.xlsx", outc.getvalue(),
                           file_name="Clients.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key=f"dlC_{SID}")
    # Visa actuel
    visa_bytes = st.session_state.get("visa_bin", None)
    if visa_bytes:
        st.download_button("⬇️ Visa.xlsx", visa_bytes,
                           file_name="Visa.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key=f"dlV_{SID}")
    else:
        try:
            with open(visa_path, "rb") as f:
                st.download_button("⬇️ Visa.xlsx", f.read(),
                                   file_name="Visa.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key=f"dlVd_{SID}")
        except Exception:
            st.caption("Visa.xlsx introuvable sur disque.")

# ==============================
# TABS
# ==============================
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "📄 Visa (aperçu)"])

# ==============================================
# 📊 Dashboard
# ==============================================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_clients.empty:
        st.info("Aucune donnée client.")
    else:
        years  = sorted([int(y) for y in pd.to_numeric(df_clients["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1, 13)]
        cats   = sorted(df_clients["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_clients.columns else []
        subs   = sorted(df_clients["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_clients.columns else []
        visas  = sorted(df_clients["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_clients.columns else []

        f1, f2, f3, f4, f5 = st.columns(5)
        fy = f1.multiselect("Année", years, default=[], key=f"dash_years_{SID}")
        fm = f2.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{SID}")
        fc = f3.multiselect("Catégorie", cats, default=[], key=f"dash_cats_{SID}")
        fs = f4.multiselect("Sous-catégorie", subs, default=[], key=f"dash_subs_{SID}")
        fv = f5.multiselect("Visa", visas, default=[], key=f"dash_visas_{SID}")

        ff = df_clients.copy()
        if fy: ff = ff[ff["_Année_"].isin(fy)]
        if fm: ff = ff[ff["Mois"].astype(str).isin(fm)]
        if fc: ff = ff[ff["Categorie"].astype(str).isin(fc)]
        if fs: ff = ff[ff["Sous-categorie"].astype(str).isin(fs)]
        if fv: ff = ff[ff["Visa"].astype(str).isin(fv)]

        # KPI
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(ff)}")
        k2.metric("Honoraires", _fmt_money(float(_safe_num_series(ff, HONO).sum())))
        k3.metric("Payé", _fmt_money(float(_safe_num_series(ff, "Payé").sum())))
        k4.metric("Reste", _fmt_money(float(_safe_num_series(ff, "Reste").sum())))

        # Table
        view = ff.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in view.columns:
                view[c] = _safe_num_series(view, c).apply(_fmt_money)

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in view_sorted.columns]
        show_cols = list(dict.fromkeys(show_cols))  # éviter doublons

        st.dataframe(
            view_sorted[show_cols].reset_index(drop=True),
            use_container_width=True,
            key=f"dash_tbl_{SID}"
        )

# ==============================================
# 📈 Analyses
# ==============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_clients.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_clients["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_clients["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_clients.columns else []
        subsA   = sorted(df_clients["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_clients.columns else []
        visasA  = sorted(df_clients["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_clients.columns else []

        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_years_{SID}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"a_months_{SID}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"a_cats_{SID}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"a_subs_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_visas_{SID}")

        dfA = df_clients.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(float(_safe_num_series(dfA, HONO).sum())))
        k3.metric("Payé", _fmt_money(float(_safe_num_series(dfA, "Payé").sum())))
        k4.metric("Reste", _fmt_money(float(_safe_num_series(dfA, "Reste").sum())))

        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("### 📊 Dossiers par catégorie")
            vc = dfA["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie", "Nombre"]
            st.bar_chart(vc.set_index("Categorie"))

        if not dfA.empty and "Mois" in dfA.columns:
            st.markdown("### 📈 Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois")
            st.line_chart(gm.set_index("Mois"))

        st.markdown("### 🧾 Détails des dossiers filtrés")
        det = dfA.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).apply(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det
        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"a_detail_{SID}")

# ==============================================
# 🏦 Escrow (synthèse simple)
# ==============================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_clients.empty:
        st.info("Aucun client.")
    else:
        dfE = df_clients.copy()
        dfE["Payé"]  = _safe_num_series(dfE, "Payé")
        dfE["Reste"] = _safe_num_series(dfE, "Reste")
        dfE[TOTAL]   = _safe_num_series(dfE, TOTAL)

        agg = dfE.groupby("Categorie", as_index=False)[[TOTAL, "Payé", "Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg[TOTAL]).replace([pd.NA, pd.NaT], 0).fillna(0.0) * 100
        st.dataframe(agg, use_container_width=True, key=f"esc_agg_{SID}")

        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money(float(dfE[TOTAL].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Reste", _fmt_money(float(dfE["Reste"].sum())))

# ==============================================
# 👤 Clients (CRUD + paiements)
# ==============================================
with tabs[3]:
    st.subheader("👤 Clients — Gestion & Suivi")

    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_{SID}")

    live = read_clients()
    live = normalize_clients(live)

    # ------- Ajouter -------
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")

        c1, c2, c3 = st.columns(3)
        nom  = c1.text_input("Nom", key=f"add_nom_{SID}")
        dt   = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=date.today().month-1, key=f"add_mois_{SID}")

        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=f"add_cat_{SID}")
        sel_sub = ""
        options_available = []
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=f"add_sub_{SID}")
            if sel_sub:
                options_available = visa_map[sel_cat][sel_sub]["options"]

        opts_sel = []
        if options_available:
            st.caption("Options disponibles pour cette sous-catégorie :")
            opts_sel = render_option_checkboxes(options_available, keyprefix=f"add_opts_{SID}")

        visa_final = compute_visa_string(sel_sub, opts_sel) if sel_sub else ""

        f1, f2 = st.columns(2)
        honor = f1.number_input(HONO, min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
        other = f2.number_input(AUTRE, min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")

        st.markdown("#### 📌 Statuts initiaux")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent = s1.checkbox("Dossier envoyé", key=f"add_sent_{SID}")
        sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(None), key=f"add_sentd_{SID}")
        acc  = s2.checkbox("Dossier accepté", key=f"add_acc_{SID}")
        acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(None), key=f"add_accd_{SID}")
        ref  = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
        ref_d  = s3.date_input("Date de refus", value=_date_for_widget(None), key=f"add_refd_{SID}")
        ann  = s4.checkbox("Dossier annulé", key=f"add_ann_{SID}")
        ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(None), key=f"add_annd_{SID}")
        rfe  = s5.checkbox("RFE", key=f"add_rfe_{SID}")
        if rfe and not any([sent, acc, ref, ann]):
            st.warning("⚠️ RFE ne peut être coché qu’avec un autre statut (envoyé/accepté/refusé/annulé).")

        note = st.text_area("Notes", key=f"add_note_{SID}")

        if st.button("💾 Enregistrer le client", key=f"btn_add_{SID}"):
            if not nom or not sel_cat or not sel_sub:
                st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                st.stop()

            total = float(honor) + float(other)
            paye  = 0.0
            reste = total

            new_row = {
                "Dossier N": next_dossier_number(live, start=13057),
                "ID_Client": make_client_id(nom, dt),
                "Nom": nom,
                "Date": dt,
                "Mois": mois,
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": visa_final or sel_sub,
                HONO: float(honor),
                AUTRE: float(other),
                TOTAL: total,
                "Payé": paye,
                "Reste": reste,
                "Paiements": [],
                "Options": {"options": opts_sel, "exclusive": None},
                "Notes": note,
                "Dossier envoyé": 1 if sent else 0,
                "Date d'envoi": (dt if sent else None) if not sent_d else sent_d,
                "Dossier accepté": 1 if acc else 0,
                "Date d'acceptation": acc_d if acc else None,
                "Dossier refusé": 1 if ref else 0,
                "Date de refus": ref_d if ref else None,
                "Dossier annulé": 1 if ann else 0,
                "Date d'annulation": ann_d if ann else None,
                "RFE": 1 if rfe else 0,
            }
            out = pd.concat([live, pd.DataFrame([new_row])], ignore_index=True)
            ok, err = write_clients_maybe_bin(out)
            if ok:
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()
            else:
                st.error(f"Erreur d’écriture : {err}")

    # ------- Modifier -------
    elif op == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if live.empty:
            st.info("Aucun client.")
        else:
            names = sorted(live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(live["ID_Client"].dropna().astype(str).unique().tolist())
            c1, c2 = st.columns(2)
            sel_name = c1.selectbox("Nom", [""]+names, index=0, key=f"mod_n_{SID}")
            sel_id   = c2.selectbox("ID_Client", [""]+ids, index=0, key=f"mod_i_{SID}")

            mask = None
            if sel_id:
                mask = (live["ID_Client"].astype(str) == sel_id)
            elif sel_name:
                mask = (live["Nom"].astype(str) == sel_name)

            if mask is None or not mask.any():
                st.stop()

            idx = live[mask].index[0]
            row = live.loc[idx].copy()

            d1, d2, d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nom_{SID}")
            dt   = d2.date_input("Date de création", value=_date_for_widget(row.get("Date")), key=f"mod_date_{SID}")
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=_month_index(row.get("Mois")), key=f"mod_mois_{SID}")

            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""] + cats,
                                   index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                                   key=f"mod_cat_{SID}")
            subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
            preset_sub = _safe_str(row.get("Sous-categorie",""))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs,
                                   index=(subs.index(preset_sub)+1 if preset_sub in subs else 0),
                                   key=f"mod_sub_{SID}")

            options_available = visa_map[sel_cat][sel_sub]["options"] if sel_cat and sel_sub and sel_cat in visa_map and sel_sub in visa_map[sel_cat] else []
            preset_opts = _json_loads_or(row.get("Options"), {"options": [], "exclusive": None})
            preset_list = preset_opts.get("options", []) if isinstance(preset_opts, dict) else []
            opts_sel = render_option_checkboxes(options_available, keyprefix=f"mod_opts_{SID}", preselected=preset_list)
            visa_final = compute_visa_string(sel_sub, opts_sel) if sel_sub else _safe_str(row.get("Visa",""))

            f1, f2 = st.columns(2)
            honor = f1.number_input(HONO, min_value=0.0,
                                    value=float(_to_float(row.get(HONO, 0.0))),
                                    step=50.0, format="%.2f", key=f"mod_h_{SID}")
            other = f2.number_input(AUTRE, min_value=0.0,
                                    value=float(_to_float(row.get(AUTRE, 0.0))),
                                    step=20.0, format="%.2f", key=f"mod_o_{SID}")

            st.markdown("#### 📌 Statuts & dates")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent = s1.checkbox("Dossier envoyé", value=bool(row.get("Dossier envoyé")), key=f"mod_sent_{SID}")
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
            acc  = s2.checkbox("Dossier accepté", value=bool(row.get("Dossier accepté")), key=f"mod_acc_{SID}")
            acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
            ref  = s3.checkbox("Dossier refusé", value=bool(row.get("Dossier refusé")), key=f"mod_ref_{SID}")
            ref_d  = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=f"mod_refd_{SID}")
            ann  = s4.checkbox("Dossier annulé", value=bool(row.get("Dossier annulé")), key=f"mod_ann_{SID}")
            ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
            rfe  = s5.checkbox("RFE", value=bool(row.get("RFE")), key=f"mod_rfe_{SID}")
            if rfe and not any([sent, acc, ref, ann]):
                st.warning("⚠️ RFE ne peut être coché qu’avec un autre statut.")

            note = st.text_area("Notes", value=_safe_str(row.get("Notes","")), key=f"mod_note_{SID}")

            if st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}"):
                if not nom or not sel_cat or not sel_sub:
                    st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                    st.stop()

                total = float(honor) + float(other)
                paye  = float(_to_float(row.get("Payé", 0.0)))
                reste = max(0.0, total - paye)

                live.at[idx, "Nom"] = nom
                live.at[idx, "Date"] = dt
                live.at[idx, "Mois"] = _safe_str(mois)
                live.at[idx, "Categorie"] = sel_cat
                live.at[idx, "Sous-categorie"] = sel_sub
                live.at[idx, "Visa"] = visa_final or sel_sub
                live.at[idx, HONO] = float(honor)
                live.at[idx, AUTRE] = float(other)
                live.at[idx, TOTAL] = float(total)
                live.at[idx, "Reste"] = float(reste)
                live.at[idx, "Options"] = {"options": opts_sel, "exclusive": None}
                live.at[idx, "Notes"] = note
                live.at[idx, "Dossier envoyé"] = 1 if sent else 0
                live.at[idx, "Date d'envoi"] = (dt if sent else None) if not sent_d else sent_d
                live.at[idx, "Dossier accepté"] = 1 if acc else 0
                live.at[idx, "Date d'acceptation"] = acc_d if acc else None
                live.at[idx, "Dossier refusé"] = 1 if ref else 0
                live.at[idx, "Date de refus"] = ref_d if ref else None
                live.at[idx, "Dossier annulé"] = 1 if ann else 0
                live.at[idx, "Date d'annulation"] = ann_d if ann else None
                live.at[idx, "RFE"] = 1 if rfe else 0

                ok, err = write_clients_maybe_bin(live)
                if ok:
                    st.success("Modifications enregistrées.")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(f"Erreur d’écriture : {err}")

            # --- Paiements ---
            st.markdown("#### 💵 Paiements")
            reste_actu = float(_to_float(live.loc[idx, "Reste"]))
            st.info(f"Reste actuel : {_fmt_money(reste_actu)}")
            paycol1, paycol2, paycol3 = st.columns(3)
            if reste_actu > 0:
                pay_amt  = paycol1.number_input("Montant à encaisser", min_value=0.0, step=10.0, format="%.2f", key=f"p_add_{SID}")
                pay_date = paycol2.date_input("Date paiement", value=date.today(), key=f"p_date_{SID}")
                mode     = paycol3.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"p_mode_{SID}")
                if st.button("Ajouter le paiement", key=f"p_btn_{SID}"):
                    if pay_amt <= 0:
                        st.warning("Montant > 0 requis.")
                        st.stop()
                    pays = _json_loads_or(live.loc[idx, "Paiements"], [])
                    pays.append({"date": str(pay_date), "montant": float(pay_amt), "mode": mode})
                    paye_new  = float(_to_float(live.loc[idx, "Payé"])) + float(pay_amt)
                    reste_new = max(0.0, float(_to_float(live.loc[idx, TOTAL])) - paye_new)
                    live.at[idx, "Paiements"] = pays
                    live.at[idx, "Payé"] = paye_new
                    live.at[idx, "Reste"] = reste_new
                    ok, err = write_clients_maybe_bin(live)
                    if ok:
                        st.success("Paiement ajouté.")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(f"Erreur écriture : {err}")

            hist = _json_loads_or(live.loc[idx, "Paiements"], [])
            if hist:
                st.write("Historique des paiements :")
                st.table(pd.DataFrame(hist))

    # ------- Supprimer -------
    elif op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if live.empty:
            st.info("Aucun client.")
        else:
            names = sorted(live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(live["ID_Client"].dropna().astype(str).unique().tolist())
            c1, c2 = st.columns(2)
            sel_name = c1.selectbox("Nom", [""]+names, index=0, key=f"del_n_{SID}")
            sel_id   = c2.selectbox("ID_Client", [""]+ids, index=0, key=f"del_i_{SID}")

            mask = None
            if sel_id:
                mask = (live["ID_Client"].astype(str) == sel_id)
            elif sel_name:
                mask = (live["Nom"].astype(str) == sel_name)

            if mask is not None and mask.any():
                row = live[mask].iloc[0]
                st.write({"Dossier N": row.get("Dossier N",""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                    out = live[~mask].copy()
                    ok, err = write_clients_maybe_bin(out)
                    if ok:
                        st.success("Client supprimé.")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(f"Erreur écriture : {err}")

# ==============================================
# 📄 Visa (aperçu)
# ==============================================
with tabs[4]:
    st.subheader("📄 Visa — aperçu structure & test")
    if df_visa.empty:
        st.info("Feuille Visa vide ou introuvable.")
    else:
        st.caption("La 1ère ligne contient les intitulés d’options. Chaque ligne: Categorie, Sous-categorie, puis la valeur **1** dans les colonnes d’options disponibles.")
        st.dataframe(df_visa, use_container_width=True, height=320)

        st.markdown("#### 🎯 Test interactif")
        cats = sorted(list(visa_map.keys()))
        tcat = st.selectbox("Catégorie", [""] + cats, index=0, key=f"v_cat_{SID}")
        tsub = ""
        if tcat:
            subs = sorted(list(visa_map.get(tcat, {}).keys()))
            tsub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=f"v_sub_{SID}")
        options_available = visa_map[tcat][tsub]["options"] if tcat and tsub and tcat in visa_map and tsub in visa_map[tcat] else []
        chosen = render_option_checkboxes(options_available, keyprefix=f"v_opts_{SID}")
        st.write("**Visa final :**", compute_visa_string(tsub, chosen) if tsub else "(—)")

# ==============================================
# 💾 Export global ZIP
# ==============================================
st.markdown("---")
st.subheader("💾 Export global (Clients + Visa)")
colz1, colz2 = st.columns([1,3])
with colz1:
    if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID}"):
        try:
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients
                with BytesIO() as xbuf:
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                        df_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                    zf.writestr("Clients.xlsx", xbuf.getvalue())
                # Visa
                if st.session_state.get("visa_bin"):
                    zf.writestr("Visa.xlsx", st.session_state["visa_bin"])
                else:
                    try:
                        with open(visa_path, "rb") as f:
                            zf.writestr("Visa.xlsx", f.read())
                    except Exception:
                        pass
            st.session_state[f"zip_export_{SID}"] = buf.getvalue()
            st.success("Archive prête.")
        except Exception as e:
            st.error(f"Erreur de préparation : {_safe_str(e)}")

with colz2:
    if st.session_state.get(f"zip_export_{SID}"):
        st.download_button(
            "⬇️ Télécharger l’export (ZIP)",
            st.session_state[f"zip_export_{SID}"],
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
            key=f"zip_dl_{SID}",
        )