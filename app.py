# =========================
# PARTIE 1/4 — BOOTSTRAP
# =========================
from __future__ import annotations

import json, os, re, io, zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Dict, Any, List, Tuple, Optional

import streamlit as st
import pandas as pd

st.set_page_config(page_title="Visa Manager", layout="wide")

# ---------- Constantes colonnes (SANS ACCENTS) ----------
COLS = [
    "ID_Client","Dossier N","Nom","Date","Categories","Sous-categorie","Visa",
    "Montant honoraires (US $)","Payé","Solde","Acompte 1","Acompte 2","RFE",
    "Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé",
    "Commentaires","Autres frais (US $)"
]

SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

STATE_FILE = ".vm_state.json"  # mémoire des derniers chemins

# ---------- Outils de clés Streamlit ----------
def skey(*parts: str) -> str:
    return "vm_" + "_".join([str(p) for p in parts])

# ---------- Helpers sûrs ----------
def _safe_str(x: Any) -> str:
    try:
        if x is None:
            return ""
        if isinstance(x, float) and pd.isna(x):
            return ""
        return str(x)
    except Exception:
        return ""

def _to_num(x: Any) -> float:
    if x is None:
        return 0.0
    try:
        if isinstance(x, (int, float)):
            return float(x)
        s = str(x)
        s = re.sub(r"[^\d\.\-]", "", s)
        return float(s) if s else 0.0
    except Exception:
        return 0.0

def _fmt_money(v: float) -> str:
    try:
        return f"${v:,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(val: Any) -> date:
    """Retourne une date utilisable par st.date_input (jamais NaT)."""
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return date.today()
        return d.date()
    except Exception:
        return date.today()

def _make_client_id(nom: str, d: date) -> str:
    base = re.sub(r"[^a-z0-9\-]+", "-", nom.lower().strip())
    base = re.sub(r"-{2,}", "-", base).strip("-") or "client"
    return f"{base}-{d.strftime('%Y%m%d')}"

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    try:
        existing = pd.to_numeric(df.get("Dossier N", pd.Series(dtype=float)), errors="coerce").dropna()
        mx = int(existing.max()) if not existing.empty else start - 1
        return mx + 1 if mx >= start else start
    except Exception:
        return start

# ---------- Mémoire des chemins ----------
def load_state() -> Dict[str, Any]:
    try:
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def save_state(d: Dict[str, Any]) -> None:
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ---------- Lecture fichiers (xlsx/csv) ----------
@st.cache_data(show_spinner=False)
def read_any_table(path_or_bytes: Any, sheet: Optional[str] = None) -> pd.DataFrame:
    """Lit CSV ou Excel. Si sheet est fourni, tente l’onglet; sinon la première feuille."""
    if path_or_bytes is None:
        return pd.DataFrame()
    # BytesIO (upload) ou chemin
    if isinstance(path_or_bytes, (bytes, bytearray)):
        bio = BytesIO(path_or_bytes)
        try:
            return pd.read_csv(bio)
        except Exception:
            bio.seek(0)
            return pd.read_excel(bio, sheet_name=sheet)
    if hasattr(path_or_bytes, "read"):  # UploadedFile
        content = path_or_bytes.read()
        bio = BytesIO(content)
        try:
            return pd.read_csv(bio)
        except Exception:
            bio.seek(0)
            return pd.read_excel(bio, sheet_name=sheet)
    # Chemin
    ext = str(path_or_bytes).lower()
    if ext.endswith(".csv"):
        return pd.read_csv(path_or_bytes)
    # Excel
    try:
        return pd.read_excel(path_or_bytes, sheet_name=sheet)
    except Exception:
        # si pas de sheet, lit première feuille
        return pd.read_excel(path_or_bytes)

# ---------- Normalisation Clients ----------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS)
    # Renommer colonnes proches vers nos libellés
    ren = {}
    for c in df.columns:
        nc = c.strip()
        # tolérance accents/variantes communes
        low = nc.lower()
        m = {
            "id_client": "ID_Client",
            "dossier n": "Dossier N",
            "nom": "Nom",
            "date": "Date",
            "categories": "Categories",
            "catégorie": "Categories",
            "categorie": "Categories",
            "sous-categorie": "Sous-categorie",
            "sous-catégorie": "Sous-categorie",
            "sous categorie": "Sous-categorie",
            "visa": "Visa",
            "montant honoraires (us $)": "Montant honoraires (US $)",
            "paye": "Payé",
            "payé": "Payé",
            "solde": "Solde",
            "acompte 1": "Acompte 1",
            "acompte 2": "Acompte 2",
            "rfe": "RFE",
            "dossiers envoyé": "Dossiers envoyé",
            "dossier envoyé": "Dossiers envoyé",
            "dossier approuvé": "Dossier approuvé",
            "dossier refuse": "Dossier refusé",
            "dossier refusé": "Dossier refusé",
            "dossier annulé": "Dossier Annulé",
            "dossier annule": "Dossier Annulé",
            "commentaires": "Commentaires",
            "autres frais (us $)": "Autres frais (US $)",
        }
        ren[nc] = m.get(low, nc)
    df = df.rename(columns=ren)

    # Ajouter colonnes manquantes
    for c in COLS:
        if c not in df.columns:
            df[c] = "" if c in ("ID_Client","Nom","Categories","Sous-categorie","Visa","Commentaires") else 0

    # Types
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for c in ["Montant honoraires (US $)","Payé","Solde","Acompte 1","Acompte 2","Autres frais (US $)"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    for c in ["RFE","Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    # ID_Client si manquant
    df["ID_Client"] = df.apply(
        lambda r: (r["ID_Client"] if _safe_str(r["ID_Client"]) else _make_client_id(_safe_str(r["Nom"]), _date_for_widget(r["Date"]))),
        axis=1
    )

    # Recalcul Solde si 0 alors qu’on a des montants
    zero_solde = (df["Solde"] == 0) & ((df["Montant honoraires (US $)"] + df["Autres frais (US $)"]) > 0)
    df.loc[zero_solde, "Solde"] = (df.loc[zero_solde, "Montant honoraires (US $)"] +
                                   df.loc[zero_solde, "Autres frais (US $)"] -
                                   df.loc[zero_solde, "Payé"]).clip(lower=0)

    # réordonner
    return df[COLS]

# ---------- Construction de la map Visa (Cat -> Sous-cat -> options) ----------
def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, List[str]]]:
    if df_visa is None or df_visa.empty:
        return {}
    # détecter colonnes cat/sous-cat
    cat_col = next((c for c in df_visa.columns if c.lower().startswith("cat")), "Categories")
    sub_col = next((c for c in df_visa.columns if "sous" in c.lower()), "Sous-categorie")
    cols = [c for c in df_visa.columns if c not in (cat_col, sub_col, "Visa")]

    vm: Dict[str, Dict[str, List[str]]] = {}
    for _, row in df_visa.iterrows():
        cat = _safe_str(row.get(cat_col, "")).strip()
        sub = _safe_str(row.get(sub_col, "")).strip()
        if not cat or not sub:
            continue
        opts = []
        for oc in cols:
            v = row.get(oc, "")
            if str(v).strip().lower() in ("1", "x", "yes", "true", "oui"):
                opts.append(oc)
        vm.setdefault(cat, {}).setdefault(sub, [])
        for o in opts:
            if o not in vm[cat][sub]:
                vm[cat][sub].append(o)
    return vm

# ============ Barre latérale : chargement & mémoire ============
st.sidebar.header("📂 Fichiers")
mode = st.sidebar.radio("Mode de chargement", ["Un fichier (Clients)", "Deux fichiers (Clients + Visa)"], horizontal=False, key=skey("mode"))

# Fichiers upload
uf_clients = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx","csv"], key=skey("up_clients"))
uf_visa    = None
if mode == "Deux fichiers (Clients + Visa)":
    uf_visa = st.sidebar.file_uploader("Visa (xlsx/csv)", type=["xlsx","csv"], key=skey("up_visa"))

# Chemins mémorisés
st.sidebar.caption("Derniers chemins mémorisés :")
state = load_state()
last_clients = state.get("clients_path", "")
last_visa    = state.get("visa_path", "")

# Sélecteurs de chemin (optionnels) pour sauvegarde directe
st.sidebar.markdown("**Chemin de sauvegarde** (sur ton PC / Drive / OneDrive) :")
save_clients_path = st.sidebar.text_input("Sauvegarder Clients vers…", value=_safe_str(last_clients), key=skey("save_clients"))
save_visa_path    = st.sidebar.text_input("Sauvegarder Visa vers…", value=_safe_str(last_visa), key=skey("save_visa"))

# Détermination sources
clients_src = uf_clients if uf_clients is not None else (last_clients if last_clients else None)
visa_src    = (uf_visa if uf_visa is not None else (last_visa if last_visa else clients_src))  # par défaut Visa = même fichier

# Chargement
df_clients_raw = normalize_clients(read_any_table(clients_src))
df_visa_raw    = read_any_table(visa_src, sheet=SHEET_VISA) if visa_src and mode=="Deux fichiers (Clients + Visa)" else read_any_table(visa_src)

# Sauvegarde de l’état si on a des sources valides
if clients_src is not None:
    state["clients_path"] = (clients_src.name if hasattr(clients_src, "name") else clients_src)
if visa_src is not None:
    state["visa_path"] = (visa_src.name if hasattr(visa_src, "name") else visa_src)
save_state(state)

# Visa map
visa_map = build_visa_map(df_visa_raw)

st.title("🛂 Visa Manager")

# Tabs
tabs = st.tabs(["📄 Fichiers chargés","📊 Dashboard","📈 Analyses","🏦 Escrow","👤 Compte client","🧾 Gestion","📄 Visa (aperçu)","💾 Export"])

with tabs[0]:
    st.subheader("📄 Fichiers chargés")
    st.write("**Clients** :", "`" + _safe_str(state.get("clients_path","")) + "`" if state.get("clients_path") else "_non défini_")
    st.write("**Visa** :", "`" + _safe_str(state.get("visa_path","")) + "`" if state.get("visa_path") else "_non défini_")
    st.dataframe(df_clients_raw.head(20), use_container_width=True)
    st.dataframe(df_visa_raw.head(20), use_container_width=True)


# =========================
# PARTIE 2/4 — DASHBOARD & ANALYSES
# =========================

# --------- Préparation des données communes ---------
df_all = df_clients_raw.copy()

# Champs dérivés year / month (à partir de "Date")
if "Date" in df_all.columns:
    dts = pd.to_datetime(df_all["Date"], errors="coerce")
    df_all["_Année_"]  = dts.dt.year
    df_all["_MoisNum_"] = dts.dt.month
    df_all["Mois"]     = dts.dt.month.map(lambda m: f"{int(m):02d}" if pd.notna(m) else "")
else:
    df_all["_Année_"] = pd.NA
    df_all["_MoisNum_"] = pd.NA
    df_all["Mois"] = ""

# Sécurisation numériques
for c in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde","Acompte 1","Acompte 2"]:
    if c in df_all.columns:
        df_all[c] = pd.to_numeric(df_all[c], errors="coerce").fillna(0.0)

# --------- 📊 DASHBOARD ---------
with tabs[1]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # Filtres
        years  = sorted([int(x) for x in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1,13)]
        cats   = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        st.markdown("#### 🎛️ Filtres")
        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", years, default=[], key=skey("dash","years"))
        fm = a2.multiselect("Mois (MM)", months, default=[], key=skey("dash","months"))
        fc = a3.multiselect("Catégories", cats, default=[], key=skey("dash","cats"))
        fs = a4.multiselect("Sous-catégories", subs, default=[], key=skey("dash","subs"))
        fv = a5.multiselect("Visa", visas, default=[], key=skey("dash","visas"))

        view = df_all.copy()
        if fy: view = view[view["_Année_"].isin(fy)]
        if fm: view = view[view["Mois"].astype(str).isin(fm)]
        if fc: view = view[view["Categories"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        # KPI (format réduit)
        k1,k2,k3,k4,k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(view)}")
        total_usd = float(view.get("Montant honoraires (US $)", pd.Series(dtype=float)).sum() +
                          view.get("Autres frais (US $)", pd.Series(dtype=float)).sum())
        k2.metric("Honoraires+Frais", _fmt_money(total_usd))
        k3.metric("Payé", _fmt_money(float(view.get("Payé", pd.Series(dtype=float)).sum())))
        k4.metric("Solde", _fmt_money(float(view.get("Solde", pd.Series(dtype=float)).sum())))
        # % envoyés
        sent_cnt = int(pd.to_numeric(view.get("Dossiers envoyé", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
        pct_sent = int(round(100.0 * sent_cnt / len(view), 0)) if len(view) else 0
        k5.metric("Envoyés (%)", f"{pct_sent}%")

        # Répartition par catégorie (valeur = nombre de dossiers)
        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if "Categories" in view.columns and not view.empty:
            vc = view["Categories"].value_counts().reset_index()
            vc.columns = ["Categories","Nombre"]
            st.bar_chart(vc.set_index("Categories"))

        # Flux par mois (Honoraires, Autres frais, Payé, Solde)
        st.markdown("#### 💵 Flux par mois")
        if not view.empty:
            g = view.copy()
            g["MoisLabel"] = g.apply(
                lambda r: f"{int(r['_Année_']):04d}-{int(r['_MoisNum_']):02d}" if pd.notna(r["_Année_"]) and pd.notna(r["_MoisNum_"]) else "NaT",
                axis=1
            )
            grp = g.groupby("MoisLabel", as_index=False)[
                ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]
            ].sum().sort_values("MoisLabel")
            st.line_chart(grp.set_index("MoisLabel"))

        # Détails
        st.markdown("#### 📋 Détails (après filtres)")
        det = view.copy()
        if "Date" in det.columns:
            det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
        # jolis montants
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]:
            if c in det.columns:
                det[c] = pd.to_numeric(det[c], errors="coerce").fillna(0.0).map(_fmt_money)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Date","Mois","Categories","Sous-categorie","Visa",
            "Montant honoraires (US $)","Autres frais (US $)","Payé","Solde",
            "Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé","RFE","Commentaires"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_","_MoisNum_","Categories","Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det
        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=skey("dash","table"))

# --------- 📈 ANALYSES ---------
with tabs[2]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        years  = sorted([int(x) for x in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1,13)]
        cats   = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        st.markdown("#### 🎛️ Filtres Analyses")
        b1,b2,b3,b4,b5 = st.columns(5)
        ay = b1.multiselect("Année", years, default=[], key=skey("anal","years"))
        am = b2.multiselect("Mois (MM)", months, default=[], key=skey("anal","months"))
        ac = b3.multiselect("Catégories", cats, default=[], key=skey("anal","cats"))
        asub = b4.multiselect("Sous-catégories", subs, default=[], key=skey("anal","subs"))
        av = b5.multiselect("Visa", visas, default=[], key=skey("anal","visas"))

        A = df_all.copy()
        if ay:   A = A[A["_Année_"].isin(ay)]
        if am:   A = A[A["Mois"].astype(str).isin(am)]
        if ac:   A = A[A["Categories"].astype(str).isin(ac)]
        if asub: A = A[A["Sous-categorie"].astype(str).isin(asub)]
        if av:   A = A[A["Visa"].astype(str).isin(av)]

        # KPI compacts
        k1,k2,k3,k4 = st.columns([1,1,1,1])
        k1.metric("Dossiers", f"{len(A)}")
        k2.metric("Honoraires", _fmt_money(float(A.get("Montant honoraires (US $)", pd.Series(dtype=float)).sum())))
        k3.metric("Payé", _fmt_money(float(A.get("Payé", pd.Series(dtype=float)).sum())))
        k4.metric("Solde", _fmt_money(float(A.get("Solde", pd.Series(dtype=float)).sum())))

        # % par catégorie
        st.markdown("#### 📌 Répartition % par catégorie")
        if not A.empty and "Categories" in A.columns:
            tot = len(A)
            pct = (A["Categories"].value_counts(normalize=True)*100).round(1).reset_index()
            pct.columns = ["Categories","%"]
            st.dataframe(pct, use_container_width=True, height=240, key=skey("anal","pct_cat"))

        # % par sous-catégorie
        st.markdown("#### 📌 Répartition % par sous-catégorie")
        if not A.empty and "Sous-categorie" in A.columns:
            tot = len(A)
            pct2 = (A["Sous-categorie"].value_counts(normalize=True)*100).round(1).reset_index()
            pct2.columns = ["Sous-categorie","%"]
            st.dataframe(pct2, use_container_width=True, height=240, key=skey("anal","pct_sub"))

        # Comparaison période A vs B (Années ou Mois)
        st.markdown("#### 🔁 Comparaison de périodes")
        ca, cb = st.columns(2)
        # Période A
        ca.subheader("Période A")
        pa_years  = ca.multiselect("Année (A)", years, default=[], key=skey("cmp","ya"))
        pa_months = ca.multiselect("Mois (A)", months, default=[], key=skey("cmp","ma"))
        # Période B
        cb.subheader("Période B")
        pb_years  = cb.multiselect("Année (B)", years, default=[], key=skey("cmp","yb"))
        pb_months = cb.multiselect("Mois (B)", months, default=[], key=skey("cmp","mb"))

        def _slice_period(base: pd.DataFrame, ys: List[int], ms: List[str]) -> pd.DataFrame:
            S = base.copy()
            if ys: S = S[S["_Année_"].isin(ys)]
            if ms: S = S[S["Mois"].astype(str).isin(ms)]
            return S

        A1 = _slice_period(df_all, pa_years, pa_months)
        A2 = _slice_period(df_all, pb_years, pb_months)

        c1, c2 = st.columns(2)
        # Synthèse A
        c1.markdown("**Synthèse A**")
        a_tot = float(A1.get("Montant honoraires (US $)", pd.Series(dtype=float)).sum() + A1.get("Autres frais (US $)", pd.Series(dtype=float)).sum())
        a_pay = float(A1.get("Payé", pd.Series(dtype=float)).sum())
        a_sol = float(A1.get("Solde", pd.Series(dtype=float)).sum())
        c1.write({
            "Dossiers": len(A1),
            "Honoraires+Frais": _fmt_money(a_tot),
            "Payé": _fmt_money(a_pay),
            "Solde": _fmt_money(a_sol),
        })

        # Synthèse B
        c2.markdown("**Synthèse B**")
        b_tot = float(A2.get("Montant honoraires (US $)", pd.Series(dtype=float)).sum() + A2.get("Autres frais (US $)", pd.Series(dtype=float)).sum())
        b_pay = float(A2.get("Payé", pd.Series(dtype=float)).sum())
        b_sol = float(A2.get("Solde", pd.Series(dtype=float)).sum())
        c2.write({
            "Dossiers": len(A2),
            "Honoraires+Frais": _fmt_money(b_tot),
            "Payé": _fmt_money(b_pay),
            "Solde": _fmt_money(b_sol),
        })



# =========================
# PARTIE 3/4 — ESCROW & COMPTE CLIENT
# =========================

# ---- petites aides paiements ----
PAY_COL = "Paiements"  # champ texte (JSON) où l'on stocke l'historique
PAY_MODES = ["Chèque", "CB", "Cash", "Virement", "Venmo"]

def _parse_pay_list(val) -> list:
    """Retourne une liste d'objets {date, mode, montant} à partir du texte JSON ou d'une liste existante."""
    if isinstance(val, list):
        return val
    s = str(val or "").strip()
    if not s or s in ["nan", "None", "NaN"]:
        return []
    try:
        obj = json.loads(s)
        return obj if isinstance(obj, list) else []
    except Exception:
        return []

def _append_payment(row: pd.Series, pay_date: date, mode: str, amount: float) -> list:
    lst = _parse_pay_list(row.get(PAY_COL))
    lst.append({
        "date": pay_date.strftime("%Y-%m-%d"),
        "mode": str(mode),
        "montant": float(amount),
    })
    return lst

def _recompute_finance(row: pd.Series) -> Tuple[float, float]:
    """Recalcule Payé & Solde à partir des montants + paiements historisés."""
    honor = float(pd.to_numeric(row.get("Montant honoraires (US $)"), errors="coerce") or 0.0)
    other = float(pd.to_numeric(row.get("Autres frais (US $)"), errors="coerce") or 0.0)
    total = honor + other
    pays  = _parse_pay_list(row.get(PAY_COL))
    paid  = float(pd.to_numeric(row.get("Payé"), errors="coerce") or 0.0)
    # Sécurise: "Payé" = paiements historisés + acomptes (si présents)
    acc1  = float(pd.to_numeric(row.get("Acompte 1"), errors="coerce") or 0.0)
    acc2  = float(pd.to_numeric(row.get("Acompte 2"), errors="coerce") or 0.0)
    paid_hist = sum(float(pd.to_numeric(p.get("montant"), errors="coerce") or 0.0) for p in pays)
    paid_new  = acc1 + acc2 + paid_hist
    reste = max(0.0, total - paid_new)
    return float(paid_new), float(reste)

def _bool01(x) -> int:
    try:
        v = int(pd.to_numeric(x, errors="coerce") or 0)
        return 1 if v == 1 else 0
    except Exception:
        return 0

def _date_or_str(val) -> str:
    if isinstance(val, (date, datetime)):
        return val.strftime("%Y-%m-%d")
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.notna(d):
            return d.date().strftime("%Y-%m-%d")
    except Exception:
        pass
    return ""

# ---------------------------------------------------
# 🏦 ESCROW — synthèse
# ---------------------------------------------------
with tabs[3]:
    st.subheader("🏦 Escrow")

    if df_clients_raw.empty:
        st.info("Aucun client chargé.")
    else:
        # Recharge la version disque courante (pour ne pas travailler sur un cache périmé)
        dfE = read_clients_file(clients_path_curr).copy()

        # champs numériques sûrs
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde","Acompte 1","Acompte 2"]:
            if c in dfE.columns:
                dfE[c] = pd.to_numeric(dfE[c], errors="coerce").fillna(0.0)

        # Recalcule Payé / Solde à partir des historiques pour cohérence
        if not dfE.empty:
            paid_list = []
            reste_list = []
            for _, r in dfE.iterrows():
                p, rr = _recompute_finance(r)
                paid_list.append(p); reste_list.append(rr)
            dfE["Payé"] = paid_list
            dfE["Solde"] = reste_list

        # KPI compacts
        t1, t2, t3 = st.columns([1,1,1])
        tot = float(dfE["Montant honoraires (US $)"].sum() + dfE["Autres frais (US $)"].sum()) if not dfE.empty else 0.0
        pay = float(dfE["Payé"].sum()) if not dfE.empty else 0.0
        rst = float(dfE["Solde"].sum()) if not dfE.empty else 0.0
        t1.metric("Total (US $)", _fmt_money(tot))
        t2.metric("Payé", _fmt_money(pay))
        t3.metric("Solde", _fmt_money(rst))

        # Liste des dossiers "envoyés" pour signaler si l'escrow doit être transféré
        st.markdown("#### 📬 Dossiers envoyés — à vérifier (transfert ESCROW)")
        sent_mask = dfE["Dossiers envoyé"].apply(_bool01) == 1 if "Dossiers envoyé" in dfE.columns else pd.Series([], dtype=bool)
        to_show = dfE[sent_mask].copy() if not dfE.empty else dfE.head(0)
        if to_show.empty:
            st.info("Aucun dossier marqué 'Dossiers envoyé'.")
        else:
            # Montant hypotétique à transférer = honoraires encaissés (acomptes + paiements) plafonnés aux honoraires
            to_show["_Encaisse_honoraires"] = (
                to_show["Acompte 1"] + to_show["Acompte 2"] +
                to_show[PAY_COL].apply(lambda v: sum(float(pd.to_numeric(p.get("montant"), errors="coerce") or 0.0) for p in _parse_pay_list(v)))
            ).clip(upper=to_show["Montant honoraires (US $)"])
            st.dataframe(
                to_show[[
                    "Dossier N","ID_Client","Nom","Visa","Montant honoraires (US $)","_Encaisse_honoraires","Payé","Solde"
                ]].rename(columns={"_Encaisse_honoraires":"ESCROW à transférer"}),
                use_container_width=True,
                key=skey("escrow","table")
            )
        st.caption("NB : L’ESCROW correspond ici aux honoraires encaissés avant/après envoi. "
                   "Tu peux ensuite opérer les transferts sur ton compte ordinaire.")

# ---------------------------------------------------
# 👤 COMPTE CLIENT — détail + encaissements
# ---------------------------------------------------
with tabs[4]:
    st.subheader("👤 Compte client")

    if df_clients_raw.empty:
        st.info("Charge d’abord tes fichiers Clients dans la barre latérale.")
    else:
        dfC = read_clients_file(clients_path_curr).copy()

        # Sélecteur client (Nom + ID)
        c1, c2 = st.columns([2,2])
        all_names = sorted(dfC["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in dfC.columns else []
        all_ids   = sorted(dfC["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in dfC.columns else []
        sel_name  = c1.selectbox("Nom", [""] + all_names, index=0, key=skey("acct","name"))
        sel_id    = c2.selectbox("ID_Client", [""] + all_ids, index=0, key=skey("acct","id"))

        # Trouver la ligne
        mask = None
        if sel_id:
            mask = (dfC["ID_Client"].astype(str) == sel_id)
        elif sel_name:
            mask = (dfC["Nom"].astype(str) == sel_name)

        if mask is None or not mask.any():
            st.stop()

        idx = dfC[mask].index[0]
        row = dfC.loc[idx].copy()

        # Recalcule Payé / Solde
        paid_new, reste_new = _recompute_finance(row)
        # Met à jour la vue (sans encore écrire disque)
        row["Payé"] = paid_new
        row["Solde"] = reste_new

        # En-tête synthèse
        st.markdown(f"**Dossier N :** {row.get('Dossier N','')} &nbsp;&nbsp; | &nbsp;&nbsp; **ID_Client :** {row.get('ID_Client','')}")
        st.markdown(f"**Nom :** {row.get('Nom','')}  &nbsp;&nbsp; | &nbsp;&nbsp; **Visa :** {row.get('Visa','')}")
        s1,s2,s3,s4 = st.columns(4)
        s1.metric("Honoraires", _fmt_money(float(pd.to_numeric(row.get("Montant honoraires (US $)"), errors="coerce") or 0.0)))
        s2.metric("Autres frais", _fmt_money(float(pd.to_numeric(row.get("Autres frais (US $)"), errors="coerce") or 0.0)))
        s3.metric("Payé", _fmt_money(float(paid_new)))
        s4.metric("Solde", _fmt_money(float(reste_new)))

        # Historique règlements
        st.markdown("### 💳 Historique des règlements")
        pays = _parse_pay_list(row.get(PAY_COL))
        # Ajoute acomptes 1/2 en tête si présents (pour traçabilité)
        acc1 = float(pd.to_numeric(row.get("Acompte 1"), errors="coerce") or 0.0)
        acc2 = float(pd.to_numeric(row.get("Acompte 2"), errors="coerce") or 0.0)
        extra_rows = []
        if acc1 > 0:
            extra_rows.append({"date": _date_or_str(row.get("Date")), "mode": "Acompte 1", "montant": acc1})
        if acc2 > 0:
            extra_rows.append({"date": _date_or_str(row.get("Date")), "mode": "Acompte 2", "montant": acc2})
        hist = extra_rows + pays

        if not hist:
            st.info("Aucun règlement historisé.")
        else:
            hdf = pd.DataFrame(hist)
            if "date" in hdf.columns:
                hdf["date"] = pd.to_datetime(hdf["date"], errors="coerce").dt.date.astype(str)
            st.dataframe(
                hdf.rename(columns={"date":"Date","mode":"Mode","montant":"Montant (US $)"}),
                use_container_width=True, height=240, key=skey("acct","hist")
            )

        st.markdown("---")
        st.markdown("### ➕ Ajouter un règlement (si dossier non soldé)")
        if float(reste_new) <= 0.0:
            st.success("Dossier soldé — aucun règlement supplémentaire requis.")
        else:
            p1, p2, p3 = st.columns([1.3,1.2,1.2])
            # Date par défaut = aujourd'hui
            pay_d = p1.date_input("Date", value=date.today(), key=skey("acct","pdate"))
            pay_m = p2.selectbox("Mode", PAY_MODES, index=1, key=skey("acct","pmode"))
            pay_a = p3.number_input("Montant (US $)", min_value=0.0, step=10.0, value=0.0, format="%.2f", key=skey("acct","pamt"))

            if st.button("💾 Enregistrer le règlement", key=skey("acct","psave")):
                if float(pay_a) <= 0.0:
                    st.warning("Le montant doit être > 0.")
                    st.stop()

                # Append paiement + recalcul puis écriture disque
                new_list = _append_payment(row, pay_d, pay_m, float(pay_a))
                dfC.at[idx, PAY_COL] = json.dumps(new_list, ensure_ascii=False)

                # Recalcule global pour la ligne
                paid_final, reste_final = _recompute_finance(dfC.loc[idx])
                dfC.at[idx, "Payé"] = paid_final
                dfC.at[idx, "Solde"] = reste_final

                # Écrit et rafraîchit
                write_clients_file(clients_path_curr, dfC)
                st.success("Règlement ajouté.")
                st.cache_data.clear()
                st.rerun()

        st.markdown("---")
        st.markdown("### 📌 Statut du dossier")
        s1, s2 = st.columns(2)

        col_status_left  = s1
        col_status_right = s2

        # Booleens (01)
        env   = col_status_left.checkbox("Dossiers envoyé", value=_bool01(row.get("Dossiers envoyé")), key=skey("acct","sent"))
        appr  = col_status_left.checkbox("Dossier approuvé", value=_bool01(row.get("Dossier approuvé")), key=skey("acct","acc"))
        refus = col_status_left.checkbox("Dossier refusé",   value=_bool01(row.get("Dossier refusé")),   key=skey("acct","ref"))
        ann   = col_status_left.checkbox("Dossier Annulé",   value=_bool01(row.get("Dossier Annulé")),   key=skey("acct","ann"))
        rfe   = col_status_left.checkbox("RFE",              value=_bool01(row.get("RFE")),              key=skey("acct","rfe"))

        # Dates associées
        d_env   = col_status_right.date_input("Date d'envoi",        value=date.today(), key=skey("acct","d_sent"))
        d_acc   = col_status_right.date_input("Date d'acceptation",  value=date.today(), key=skey("acct","d_acc"))
        d_ref   = col_status_right.date_input("Date de refus",       value=date.today(), key=skey("acct","d_ref"))
        d_ann   = col_status_right.date_input("Date d'annulation",   value=date.today(), key=skey("acct","d_ann"))

        if rfe and not any([env, appr, refus, ann]):
            st.warning("RFE ne peut être coché qu’avec un autre statut (envoyé, approuvé, refusé ou annulé).")

        if st.button("💾 Enregistrer le statut", key=skey("acct","save_stat")):
            dfC.at[idx, "Dossiers envoyé"]   = 1 if env else 0
            dfC.at[idx, "Dossier approuvé"]  = 1 if appr else 0
            dfC.at[idx, "Dossier refusé"]    = 1 if refus else 0
            dfC.at[idx, "Dossier Annulé"]    = 1 if ann else 0
            dfC.at[idx, "RFE"]               = 1 if rfe else 0
            dfC.at[idx, "Date d'envoi"]      = d_env.strftime("%Y-%m-%d") if env else ""
            dfC.at[idx, "Date d'acceptation"]= d_acc.strftime("%Y-%m-%d") if appr else ""
            dfC.at[idx, "Date de refus"]     = d_ref.strftime("%Y-%m-%d") if refus else ""
            dfC.at[idx, "Date d'annulation"] = d_ann.strftime("%Y-%m-%d") if ann else ""
            write_clients_file(clients_path_curr, dfC)
            st.success("Statut mis à jour.")
            st.cache_data.clear()
            st.rerun()



# =========================
# PARTIE 4/4 — GESTION (CRUD) • VISA (APERÇU) • EXPORT
# =========================

# --- petits helpers robustes (au cas où ils n'existent pas déjà dans les parties précédentes)
def _coerce_int(x, default=0):
    try:
        v = int(pd.to_numeric(x, errors="coerce"))
        return v
    except Exception:
        return default

def _next_dossier_number(df: pd.DataFrame, start_at: int = 13057) -> int:
    if "Dossier N" not in df.columns or df.empty:
        return start_at
    vals = pd.to_numeric(df["Dossier N"], errors="coerce").dropna().astype(int)
    return (vals.max() + 1) if len(vals) else start_at

def _make_id_client(name: str, dt_val) -> str:
    # essaie d'utiliser un helper déjà défini (si présent)
    try:
        return _make_client_id(name, dt_val)  # type: ignore[name-defined]
    except Exception:
        pass
    # fallback local: NOM-YYYYMMDD-XX (uniques dans le fichier)
    base = str(name).strip().lower().replace(" ", "-")
    try:
        d = pd.to_datetime(dt_val, errors="coerce")
        stamp = d.strftime("%Y%m%d") if pd.notna(d) else datetime.now().strftime("%Y%m%d")
    except Exception:
        stamp = datetime.now().strftime("%Y%m%d")
    base = f"{base}-{stamp}"
    # garantie unicité en suffixant -0,-1,...
    df_curr = read_clients_file(clients_path_curr)
    exist = set(df_curr["ID_Client"].astype(str)) if "ID_Client" in df_curr.columns else set()
    if base not in exist:
        return base
    i = 0
    while f"{base}-{i}" in exist:
        i += 1
    return f"{base}-{i}"

def _safe_date_for_widget(val):
    if isinstance(val, (date, datetime)):
        return val if isinstance(val, date) else val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.notna(d):
            return d.date()
    except Exception:
        pass
    return date.today()

# champs obligatoires utilisés partout
REQ_COLS = [
    "ID_Client", "Dossier N", "Nom", "Date", "Categorie", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde",
    "Acompte 1", "Acompte 2", "Commentaires",
    "RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé",
    "Date d'envoi", "Date d'acceptation", "Date de refus", "Date d'annulation",
    PAY_COL  # "Paiements" défini en partie 3/4
]

def _ensure_client_columns(df: pd.DataFrame) -> pd.DataFrame:
    for c in REQ_COLS:
        if c not in df.columns:
            df[c] = "" if c not in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde","Acompte 1","Acompte 2"] else 0.0
    # types numériques sûrs
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde","Acompte 1","Acompte 2"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    # bool 0/1
    for c in ["RFE","Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé"]:
        df[c] = df[c].apply(lambda x: 1 if _coerce_int(x,0)==1 else 0)
    return df

# ------------------------------------------
# 🧾 GESTION : Ajouter / Modifier / Supprimer
# ------------------------------------------
with tabs[5]:
    st.subheader("🧾 Gestion (Ajouter / Modifier / Supprimer)")

    if df_clients_raw.empty:
        st.info("Charge d’abord ton fichier Clients dans la barre latérale.")
    else:
        df_live = read_clients_file(clients_path_curr).copy()
        df_live = _ensure_client_columns(df_live)

        action = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=skey("crud","op"))

        # ---------- AJOUT ----------
        if action == "Ajouter":
            st.markdown("### ➕ Ajouter un client")

            a1, a2, a3 = st.columns([2,1,1])
            nom    = a1.text_input("Nom", "", key=skey("add","nom"))
            dval   = a2.date_input("Date de création", value=date.today(), key=skey("add","date"))
            # Catégorie / Sous-catégorie / Visa guidés par df_visa_raw (si présent)
            cats = sorted(df_visa_raw["Categorie"].dropna().astype(str).unique().tolist()) if not df_visa_raw.empty and "Categorie" in df_visa_raw.columns else []
            cat  = a3.selectbox("Catégorie", [""]+cats, index=0, key=skey("add","cat"))

            sub = ""
            visa = ""
            if cat:
                sub_opts = sorted(df_visa_raw.loc[df_visa_raw["Categorie"].astype(str)==cat, "Sous-categorie"].dropna().astype(str).unique()) \
                           if "Sous-categorie" in df_visa_raw.columns else []
                sub = st.selectbox("Sous-catégorie", [""]+list(sub_opts), index=0, key=skey("add","sub"))

                if sub:
                    # options (cases à cocher) si présentes en colonnes (valeur 1 sur la ligne concernée)
                    row_v = df_visa_raw[(df_visa_raw["Categorie"].astype(str)==cat) &
                                        (df_visa_raw["Sous-categorie"].astype(str)==sub)]
                    opt_cols = []
                    if not row_v.empty:
                        r0 = row_v.iloc[0]
                        opt_cols = [c for c in df_visa_raw.columns if c not in ["Categorie","Sous-categorie"] and _coerce_int(r0.get(c),0)==1]
                    chosen = []
                    if opt_cols:
                        st.caption("Options disponibles :")
                        for copt in opt_cols:
                            if st.checkbox(copt, key=skey("add","opt",copt)):
                                chosen.append(copt)
                    # Règle d’assemblage du label Visa = "Sous-categorie (+ options cochées jointes par un espace)"
                    visa = f"{sub}" + ("" if not chosen else " " + " ".join(chosen))

            b1, b2 = st.columns(2)
            honor  = b1.number_input("Montant honoraires (US $)", min_value=0.0, step=50.0, value=0.0, format="%.2f", key=skey("add","honor"))
            other  = b2.number_input("Autres frais (US $)",       min_value=0.0, step=25.0, value=0.0, format="%.2f", key=skey("add","other"))
            c1, c2 = st.columns(2)
            acc1   = c1.number_input("Acompte 1", min_value=0.0, step=10.0, value=0.0, format="%.2f", key=skey("add","acc1"))
            acc2   = c2.number_input("Acompte 2", min_value=0.0, step=10.0, value=0.0, format="%.2f", key=skey("add","acc2"))
            com    = st.text_area("Commentaires", "", key=skey("add","com"))

            st.markdown("#### 📌 Statuts initiaux")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent  = s1.checkbox("Dossiers envoyé", key=skey("add","sent"))
            acc   = s2.checkbox("Dossier approuvé", key=skey("add","acc"))
            ref   = s3.checkbox("Dossier refusé",   key=skey("add","ref"))
            ann   = s4.checkbox("Dossier Annulé",   key=skey("add","ann"))
            rfe   = s5.checkbox("RFE",              key=skey("add","rfe"))
            d_sent = s1.date_input("Date d'envoi", value=date.today(), key=skey("add","dsent"))
            d_acc  = s2.date_input("Date d'acceptation", value=date.today(), key=skey("add","dacc"))
            d_ref  = s3.date_input("Date de refus", value=date.today(), key=skey("add","dref"))
            d_ann  = s4.date_input("Date d'annulation", value=date.today(), key=skey("add","dann"))

            if rfe and not any([sent, acc, ref, ann]):
                st.warning("RFE ne peut être coché qu’avec un autre statut (envoyé, approuvé, refusé ou annulé).")

            if st.button("💾 Enregistrer le client", key=skey("add","save")):
                if not nom:
                    st.warning("Le nom est obligatoire.")
                    st.stop()
                if not cat or not sub:
                    st.warning("Choisir Catégorie et Sous-catégorie.")
                    st.stop()

                df_curr = read_clients_file(clients_path_curr).copy()
                df_curr = _ensure_client_columns(df_curr)

                did  = _make_id_client(nom, dval)
                dnum = _next_dossier_number(df_curr, start_at=13057)
                total = float(honor) + float(other)
                # paye et solde selon règles (acomptes inclus)
                paye = float(acc1) + float(acc2)
                reste = max(0.0, total - paye)

                new_row = {
                    "ID_Client": did,
                    "Dossier N": dnum,
                    "Nom": nom,
                    "Date": dval,
                    "Categorie": cat,
                    "Sous-categorie": sub,
                    "Visa": visa if visa else sub,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Payé": float(paye),
                    "Solde": float(reste),
                    "Acompte 1": float(acc1),
                    "Acompte 2": float(acc2),
                    "Commentaires": com,
                    "RFE": 1 if rfe else 0,
                    "Dossiers envoyé": 1 if sent else 0,
                    "Dossier approuvé": 1 if acc else 0,
                    "Dossier refusé": 1 if ref else 0,
                    "Dossier Annulé": 1 if ann else 0,
                    "Date d'envoi": d_sent.strftime("%Y-%m-%d") if sent else "",
                    "Date d'acceptation": d_acc.strftime("%Y-%m-%d") if acc else "",
                    "Date de refus": d_ref.strftime("%Y-%m-%d") if ref else "",
                    "Date d'annulation": d_ann.strftime("%Y-%m-%d") if ann else "",
                    PAY_COL: json.dumps([], ensure_ascii=False),
                }
                df_new = pd.concat([df_curr, pd.DataFrame([new_row])], ignore_index=True)
                write_clients_file(clients_path_curr, df_new)
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()

        # ---------- MODIFICATION ----------
        elif action == "Modifier":
            st.markdown("### ✏️ Modifier un client")
            if df_live.empty:
                st.info("Aucun client.")
            else:
                names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
                ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
                m1, m2 = st.columns(2)
                sel_name = m1.selectbox("Nom", [""]+names, index=0, key=skey("mod","name"))
                sel_id   = m2.selectbox("ID_Client", [""]+ids, index=0, key=skey("mod","id"))

                mask = None
                if sel_id:
                    mask = (df_live["ID_Client"].astype(str)==sel_id)
                elif sel_name:
                    mask = (df_live["Nom"].astype(str)==sel_name)

                if mask is None or not mask.any():
                    st.stop()

                idx = df_live[mask].index[0]
                row = df_live.loc[idx].copy()

                d1, d2, d3 = st.columns([2,1,1])
                nom  = d1.text_input("Nom", str(row.get("Nom","")), key=skey("mod","nomv"))
                dval = _safe_date_for_widget(row.get("Date"))
                dt   = d2.date_input("Date de création", value=dval, key=skey("mod","date"))

                # Catégorie / Sous-catégorie / Visa
                cats = sorted(df_visa_raw["Categorie"].dropna().astype(str).unique().tolist()) if not df_visa_raw.empty and "Categorie" in df_visa_raw.columns else []
                preset_cat = str(row.get("Categorie",""))
                cat_idx = (cats.index(preset_cat)+1) if preset_cat in cats else 0
                cat  = d3.selectbox("Catégorie", [""]+cats, index=cat_idx, key=skey("mod","cat"))

                sub = str(row.get("Sous-categorie",""))
                visa = str(row.get("Visa",""))
                if cat:
                    sub_opts = sorted(df_visa_raw.loc[df_visa_raw["Categorie"].astype(str)==cat, "Sous-categorie"].dropna().astype(str).unique()) \
                               if "Sous-categorie" in df_visa_raw.columns else []
                    sub_idx = (list(sub_opts).index(sub)+1) if sub in sub_opts else 0
                    sub = st.selectbox("Sous-catégorie", [""]+list(sub_opts), index=sub_idx, key=skey("mod","sub"))
                    # proposer de régénérer le visa depuis les cases à cocher si besoin
                    visa_old = visa
                    row_v = df_visa_raw[(df_visa_raw["Categorie"].astype(str)==cat) &
                                        (df_visa_raw["Sous-categorie"].astype(str)==sub)]
                    chosen = []
                    if not row_v.empty:
                        r0 = row_v.iloc[0]
                        opt_cols = [c for c in df_visa_raw.columns if c not in ["Categorie","Sous-categorie"] and _coerce_int(r0.get(c),0)==1]
                        if opt_cols:
                            st.caption("Options (si tu coches, le champ 'Visa' sera recalculé) :")
                            for copt in opt_cols:
                                if st.checkbox(copt, key=skey("mod","opt",copt)):
                                    chosen.append(copt)
                    if chosen:
                        visa = f"{sub}" + " " + " ".join(chosen)
                    else:
                        # laisser la valeur existante par défaut
                        visa = st.text_input("Visa", visa_old, key=skey("mod","visatxt"))
                else:
                    visa = st.text_input("Visa", visa, key=skey("mod","visatxt0"))

                f1, f2 = st.columns(2)
                honor = f1.number_input("Montant honoraires (US $)", min_value=0.0,
                                        value=float(pd.to_numeric(row.get("Montant honoraires (US $)"), errors="coerce") or 0.0),
                                        step=50.0, format="%.2f", key=skey("mod","honor"))
                other = f2.number_input("Autres frais (US $)", min_value=0.0,
                                        value=float(pd.to_numeric(row.get("Autres frais (US $)"), errors="coerce") or 0.0),
                                        step=25.0, format="%.2f", key=skey("mod","other"))
                g1, g2 = st.columns(2)
                acc1 = g1.number_input("Acompte 1", min_value=0.0,
                                       value=float(pd.to_numeric(row.get("Acompte 1"), errors="coerce") or 0.0),
                                       step=10.0, format="%.2f", key=skey("mod","acc1"))
                acc2 = g2.number_input("Acompte 2", min_value=0.0,
                                       value=float(pd.to_numeric(row.get("Acompte 2"), errors="coerce") or 0.0),
                                       step=10.0, format="%.2f", key=skey("mod","acc2"))
                com  = st.text_area("Commentaires", str(row.get("Commentaires","")), key=skey("mod","com"))

                st.markdown("#### 📌 Statuts")
                s1, s2, s3, s4, s5 = st.columns(5)
                sent = s1.checkbox("Dossiers envoyé", value=_coerce_int(row.get("Dossiers envoyé"),0)==1, key=skey("mod","sent"))
                appr = s2.checkbox("Dossier approuvé", value=_coerce_int(row.get("Dossier approuvé"),0)==1, key=skey("mod","appr"))
                refus= s3.checkbox("Dossier refusé",   value=_coerce_int(row.get("Dossier refusé"),0)==1, key=skey("mod","refus"))
                ann  = s4.checkbox("Dossier Annulé",   value=_coerce_int(row.get("Dossier Annulé"),0)==1, key=skey("mod","ann"))
                rfe  = s5.checkbox("RFE",              value=_coerce_int(row.get("RFE"),0)==1, key=skey("mod","rfe"))
                dsent = s1.date_input("Date d'envoi",        value=_safe_date_for_widget(row.get("Date d'envoi")),        key=skey("mod","dsent"))
                dacc  = s2.date_input("Date d'acceptation",  value=_safe_date_for_widget(row.get("Date d'acceptation")),  key=skey("mod","dacc"))
                dref  = s3.date_input("Date de refus",       value=_safe_date_for_widget(row.get("Date de refus")),       key=skey("mod","dref"))
                dann  = s4.date_input("Date d'annulation",   value=_safe_date_for_widget(row.get("Date d'annulation")),   key=skey("mod","dann"))

                if rfe and not any([sent, appr, refus, ann]):
                    st.warning("RFE ne peut être coché qu’avec un autre statut (envoyé, approuvé, refusé ou annulé).")

                if st.button("💾 Enregistrer les modifications", key=skey("mod","save")):
                    df_live.at[idx, "Nom"] = nom
                    df_live.at[idx, "Date"] = dt
                    df_live.at[idx, "Categorie"] = cat
                    df_live.at[idx, "Sous-categorie"] = sub
                    df_live.at[idx, "Visa"] = visa if visa else sub
                    df_live.at[idx, "Montant honoraires (US $)"] = float(honor)
                    df_live.at[idx, "Autres frais (US $)"] = float(other)
                    # recalcul simple (les historiques détaillés restent dans l’onglet Compte client)
                    total = float(honor) + float(other)
                    paye  = float(pd.to_numeric(df_live.at[idx, "Payé"], errors="coerce") or 0.0)
                    # si paye < acomptes, réaligner
                    acc1v = float(acc1); acc2v = float(acc2)
                    min_paid = acc1v + acc2v
                    paye = max(paye, min_paid)
                    df_live.at[idx, "Payé"] = paye
                    df_live.at[idx, "Solde"] = max(0.0, total - paye)
                    df_live.at[idx, "Acompte 1"] = acc1v
                    df_live.at[idx, "Acompte 2"] = acc2v
                    df_live.at[idx, "Commentaires"] = com
                    df_live.at[idx, "RFE"] = 1 if rfe else 0
                    df_live.at[idx, "Dossiers envoyé"]  = 1 if sent else 0
                    df_live.at[idx, "Dossier approuvé"] = 1 if appr else 0
                    df_live.at[idx, "Dossier refusé"]   = 1 if refus else 0
                    df_live.at[idx, "Dossier Annulé"]   = 1 if ann else 0
                    df_live.at[idx, "Date d'envoi"]       = dsent.strftime("%Y-%m-%d") if sent else ""
                    df_live.at[idx, "Date d'acceptation"] = dacc.strftime("%Y-%m-%d")  if appr else ""
                    df_live.at[idx, "Date de refus"]      = dref.strftime("%Y-%m-%d")  if refus else ""
                    df_live.at[idx, "Date d'annulation"]  = dann.strftime("%Y-%m-%d")  if ann else ""

                    write_clients_file(clients_path_curr, df_live)
                    st.success("Modifications enregistrées.")
                    st.cache_data.clear()
                    st.rerun()

        # ---------- SUPPRESSION ----------
        elif action == "Supprimer":
            st.markdown("### 🗑️ Supprimer un client")
            if df_live.empty:
                st.info("Aucun client.")
            else:
                names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
                ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
                d1, d2 = st.columns(2)
                sel_name = d1.selectbox("Nom", [""]+names, index=0, key=skey("del","name"))
                sel_id   = d2.selectbox("ID_Client", [""]+ids, index=0, key=skey("del","id"))
                mask = None
                if sel_id:
                    mask = (df_live["ID_Client"].astype(str)==sel_id)
                elif sel_name:
                    mask = (df_live["Nom"].astype(str)==sel_name)
                if mask is not None and mask.any():
                    r0 = df_live[mask].iloc[0]
                    st.write({"Dossier N": r0.get("Dossier N",""), "Nom": r0.get("Nom",""), "Visa": r0.get("Visa","")})
                    if st.button("❗ Confirmer la suppression", key=skey("del","go")):
                        df_new = df_live[~mask].copy()
                        write_clients_file(clients_path_curr, df_new)
                        st.success("Client supprimé.")
                        st.cache_data.clear()
                        st.rerun()

# -------------------------
# 📄 Visa (aperçu simple)
# -------------------------
with tabs[6]:
    st.subheader("📄 Visa (aperçu)")
    if df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True, height=380, key=skey("visa","view"))

# -------------------------
# 💾 Export (Clients & Visa)
# -------------------------
with tabs[7]:
    st.subheader("💾 Export")
    c1, c2 = st.columns(2)

    # Export Clients.xlsx (depuis l'état disque courant)
    df_export = read_clients_file(clients_path_curr).copy() if clients_path_curr else pd.DataFrame()
    xbuf = None
    if not df_export.empty:
        xbuf = BytesIO()
        with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
            df_export.to_excel(wr, sheet_name="Clients", index=False)
    if xbuf:
        c1.download_button(
            label="⬇️ Télécharger Clients.xlsx",
            data=xbuf.getvalue(),
            file_name="Clients.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=skey("dl","clients")
        )
    else:
        c1.info("Aucun client à exporter.")

    # Export Visa.xlsx (si chargé)
    if visa_path_curr:
        try:
            c2.download_button(
                label="⬇️ Télécharger Visa.xlsx",
                data=open(visa_path_curr, "rb").read(),
                file_name="Visa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=skey("dl","visa")
            )
        except Exception:
            # fallback: réécrire depuis df_visa_raw
            if not df_visa_raw.empty:
                vb = BytesIO()
                with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                    df_visa_raw.to_excel(wr, sheet_name="Visa", index=False)
                c2.download_button(
                    label="⬇️ Télécharger Visa.xlsx (reconstruit)",
                    data=vb.getvalue(),
                    file_name="Visa.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=skey("dl","visa2")
                )
            else:
                c2.info("Aucun visa à exporter.")