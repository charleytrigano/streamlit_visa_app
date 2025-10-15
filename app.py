# ==============================
# 🛂 VISA MANAGER — PARTIE 1/4
# ==============================
import os, json, re, zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple, Optional

import pandas as pd
import numpy as np
import streamlit as st

# ------------------
# Constantes colonnes
# ------------------
COLS_EXPECTED = [
    "ID_Client","Dossier N","Nom","Date","Categories","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Payé","Solde",
    "Acompte 1","Acompte 2","RFE","Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé",
    "Commentaires"
]

# Fichier JSON pour mémoriser les derniers chemins
LAST_JSON = ".visa_manager_last.json"

# ------------------
# Helpers format & num
# ------------------
def _safe_str(x):
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)

def _to_num(x, default=0.0) -> float:
    try:
        if isinstance(x, (int, float, np.number)):
            return float(x)
        s = _safe_str(x)
        if s.strip() == "":
            return float(default)
        s = re.sub(r"[^\d\.\-]", "", s)
        return float(s) if s not in ("", "-", ".", "-.") else float(default)
    except Exception:
        return float(default)

def _fmt_money(v: float) -> str:
    try:
        return f"${float(v):,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(v):
    """Retourne une date Python (ou date.today()) pour date_input."""
    if isinstance(v, date):
        return v
    if isinstance(v, datetime):
        return v.date()
    try:
        d = pd.to_datetime(v, errors="coerce")
        if pd.notna(d):
            return d.date()
    except Exception:
        pass
    return date.today()

def _month_str_from_date(v) -> str:
    try:
        d = pd.to_datetime(v, errors="coerce")
        if pd.notna(d):
            return f"{int(d.month):02d}"
    except Exception:
        pass
    return ""

# ------------------
# Persistance chemins
# ------------------
def load_last_paths() -> Dict[str, str]:
    try:
        if os.path.exists(LAST_JSON):
            with open(LAST_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {"clients": "", "visa": ""}

def save_last_paths(clients_path: str, visa_path: str):
    try:
        data = {"clients": clients_path or "", "visa": visa_path or ""}
        with open(LAST_JSON, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ------------------
# Lecture table générique
# ------------------
def read_any_table(path_or_buffer) -> Optional[pd.DataFrame]:
    if path_or_buffer is None:
        return None
    try:
        if isinstance(path_or_buffer, (bytes, bytearray, BytesIO)):
            # Tentative Excel puis CSV
            try:
                return pd.read_excel(path_or_buffer)
            except Exception:
                path_or_buffer.seek(0)
                return pd.read_csv(path_or_buffer, sep=None, engine="python")
        if isinstance(path_or_buffer, str):
            if path_or_buffer.lower().endswith((".xlsx", ".xlsm", ".xls")):
                return pd.read_excel(path_or_buffer)
            else:
                return pd.read_csv(path_or_buffer, sep=None, engine="python")
        # fallback
        return pd.read_excel(path_or_buffer)
    except Exception:
        try:
            # Dernière chance CSV
            return pd.read_csv(path_or_buffer, sep=None, engine="python")
        except Exception:
            return None

# ------------------
# Normalisation CLIENTS
# ------------------
def normalize_clients(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS_EXPECTED)

    # Renommer colonnes proches / tolérer variations
    rename_map = {
        "Categorie": "Categories",
        "Catégorie": "Categories",
        "Sous-categories": "Sous-categorie",
        "Sous-catégorie": "Sous-categorie",
        "Montant honoraires": "Montant honoraires (US $)",
        "Autres frais": "Autres frais (US $)",
        "Accompte 1": "Acompte 1",
        "Accompte 2": "Acompte 2",
        "Dossier envoyé": "Dossiers envoyé",  # pour harmoniser au pluriel donné
        "Dossier approuvé": "Dossier approuvé",
        "Dossier refusé": "Dossier refusé",
        "Dossier annulé": "Dossier Annulé",
        "Commentaires/Autres frais": "Commentaires",
        "Solde (US $)": "Solde",
    }
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(columns=rename_map)

    # Ajouter colonnes manquantes
    for c in COLS_EXPECTED:
        if c not in df.columns:
            df[c] = np.nan

    # Types
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde","Acompte 1","Acompte 2"]:
        df[c] = df[c].apply(_to_num)

    for c in ["RFE","Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé"]:
        df[c] = df[c].apply(lambda x: 1 if _to_num(x) == 1 else 0)

    # ID auto si manquant
    def _make_id(row):
        base = re.sub(r"[^A-Za-z0-9]+", "", _safe_str(row.get("Nom",""))).upper() or "CLIENT"
        d = row.get("Date")
        dstr = ""
        try:
            if isinstance(d, (date, datetime)):
                dstr = d.strftime("%Y%m%d")
            else:
                dt = pd.to_datetime(d, errors="coerce")
                dstr = dt.strftime("%Y%m%d") if pd.notna(dt) else date.today().strftime("%Y%m%d")
        except Exception:
            dstr = date.today().strftime("%Y%m%d")
        return f"{base}-{dstr}"

    df["ID_Client"] = df["ID_Client"].fillna("").astype(str)
    df.loc[df["ID_Client"].str.strip() == "", "ID_Client"] = df[df["ID_Client"].str.strip() == ""].apply(_make_id, axis=1)

    # Dossier N auto-incrément si manquant (à partir 13057)
    def _next_dossier_int(existing):
        used = [int(_to_num(x)) for x in existing if _to_num(x) > 0]
        start = 13057
        c = start
        used_set = set(used)
        while c in used_set:
            c += 1
        return c

    mask_dn = df["Dossier N"].isna() | (df["Dossier N"].astype(str).str.strip() == "")
    if mask_dn.any():
        next_val = _next_dossier_int(df["Dossier N"].tolist())
        idxs = df[mask_dn].index.tolist()
        for i in idxs:
            df.at[i, "Dossier N"] = next_val
            next_val += 1

    # Recalcul Total / Payé / Solde si possible
    if "Total (US $)" not in df.columns:
        df["Total (US $)"] = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]

    # Si Payé absent, calcul via Acomptes
    if df["Payé"].isna().any() or (df["Payé"] == 0).all():
        df["Payé"] = df["Acompte 1"].apply(_to_num) + df["Acompte 2"].apply(_to_num) + df["Payé"].apply(_to_num)

    # Solde
    df["Solde"] = df["Total (US $)"].apply(_to_num) - df["Payé"].apply(_to_num)

    # Mois & Année (cachés)
    df["_Année_"] = df["Date"].dt.year
    df["Mois"]    = df["Date"].dt.month.apply(lambda m: f"{int(m):02d}" if pd.notna(m) else "")
    df["_MoisNum_"]= df["Date"].dt.month

    # Nettoyage strings
    for c in ["Nom","Categories","Sous-categorie","Visa","Commentaires"]:
        df[c] = df[c].astype(str).fillna("").str.strip()

    return df[COLS_EXPECTED + ["Total (US $)","_Année_","Mois","_MoisNum_"]]

# ------------------
# Normalisation VISA
# ------------------
def normalize_visa(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Categories","Sous-categorie","Visa"])
    df = df.copy()
    # Harmoniser noms colonnes si besoin
    rename_map = {
        "Categorie":"Categories",
        "Catégorie":"Categories",
        "Sous-categories":"Sous-categorie",
        "Sous-catégorie":"Sous-categorie",
    }
    df.columns = [c.strip() for c in df.columns]
    df = df.rename(columns=rename_map)
    # Garder colonnes clés si existent
    keep = [c for c in ["Categories","Sous-categorie","Visa"] if c in df.columns]
    if not keep:
        # structure minimale
        return pd.DataFrame(columns=["Categories","Sous-categorie","Visa"])
    # Nettoyage
    for c in keep:
        df[c] = df[c].astype(str).fillna("").str.strip()
    # Drop lignes vides
    df = df[(df["Categories"]!="") | (("Sous-categorie" in df.columns) & (df["Sous-categorie"]!="")) | (("Visa" in df.columns) & (df["Visa"]!=""))]
    return df[keep]

# ------------------
# Build visa_map: {cat: {sub: [visa options]}}
# ------------------
def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, List[str]]]:
    visa_map: Dict[str, Dict[str, List[str]]] = {}
    if df_visa is None or df_visa.empty:
        return visa_map
    cats = df_visa["Categories"].dropna().astype(str).unique().tolist() if "Categories" in df_visa.columns else []
    for c in cats:
        submap: Dict[str, List[str]] = {}
        sdf = df_visa[df_visa["Categories"].astype(str) == str(c)]
        subs = sdf["Sous-categorie"].dropna().astype(str).unique().tolist() if "Sous-categorie" in sdf.columns else []
        if not subs:
            subs = [""]
        for s in subs:
            v = sdf[sdf["Sous-categorie"].astype(str) == str(s)]["Visa"].dropna().astype(str).unique().tolist() if "Sous-categorie" in sdf.columns and "Visa" in sdf.columns else []
            submap[s] = sorted([x for x in v if x != ""])
        visa_map[str(c)] = submap
    return visa_map

# ------------------
# Mémo streamlit cache
# ------------------
@st.cache_data(show_spinner=False)
def read_clients_file(path) -> pd.DataFrame:
    df = read_any_table(path)
    return normalize_clients(df)

@st.cache_data(show_spinner=False)
def read_visa_file(path) -> pd.DataFrame:
    df = read_any_table(path)
    return normalize_visa(df)

# =========================
# UI — BARRE LATERALE : chargement
# =========================
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

last_paths = load_last_paths()
if "clients_path" not in st.session_state:
    st.session_state.clients_path = last_paths.get("clients","")
if "visa_path" not in st.session_state:
    st.session_state.visa_path = last_paths.get("visa","")

with st.sidebar:
    st.header("📂 Fichiers")
    mode = st.radio("Mode de chargement", ["Un fichier (Clients)","Deux fichiers (Clients + Visa)"], index=0)
    up_clients = st.file_uploader("Clients (xlsx/csv)", type=["xlsx","xls","csv"], key="upl_clients")
    up_visa = None
    if mode == "Deux fichiers (Clients + Visa)":
        up_visa = st.file_uploader("Visa (xlsx/csv)", type=["xlsx","xls","csv"], key="upl_visa")

    colp1, colp2 = st.columns(2)
    with colp1:
        if st.button("Utiliser ces fichiers", use_container_width=True):
            # Sauver en fichiers temporaires locaux (pour garder un chemin)
            if up_clients is not None:
                tmpc = f"./upload_{up_clients.name}"
                with open(tmpc, "wb") as f:
                    f.write(up_clients.getbuffer())
                st.session_state.clients_path = tmpc
            if up_visa is not None:
                tmpv = f"./upload_{up_visa.name}"
                with open(tmpv, "wb") as f:
                    f.write(up_visa.getbuffer())
                st.session_state.visa_path = tmpv
            if mode == "Un fichier (Clients)" and up_clients is not None:
                # Si un seul fichier, on suppose que Visa est dans le même fichier/onglet = fallback
                st.session_state.visa_path = st.session_state.clients_path
            save_last_paths(st.session_state.clients_path, st.session_state.visa_path)
            st.success("Fichiers mémorisés.")
            st.cache_data.clear()
            st.experimental_rerun()

    with colp2:
        if st.button("Oublier fichiers", use_container_width=True):
            st.session_state.clients_path = ""
            st.session_state.visa_path = ""
            save_last_paths("", "")
            st.cache_data.clear()
            st.experimental_rerun()

clients_path_curr = st.session_state.get("clients_path","")
visa_path_curr    = st.session_state.get("visa_path","")

st.markdown("### 📄 Fichiers chargés")
st.write("**Clients** :", f"`{clients_path_curr}`" if clients_path_curr else "_(aucun)_")
st.write("**Visa**    :", f"`{visa_path_curr}`" if visa_path_curr else "_(aucun)_")

# Charger DataFrames
df_clients_raw = read_clients_file(clients_path_curr) if clients_path_curr else pd.DataFrame(columns=COLS_EXPECTED)
df_visa_raw    = read_visa_file(visa_path_curr) if visa_path_curr else pd.DataFrame(columns=["Categories","Sous-categorie","Visa"])
visa_map       = build_visa_map(df_visa_raw)

# Préparer DF global (df_all)
df_all = df_clients_raw.copy()
if not df_all.empty:
    if "Total (US $)" not in df_all.columns:
        df_all["Total (US $)"] = df_all["Montant honoraires (US $)"] + df_all["Autres frais (US $)"]



# =========================
# Création des onglets (nommés)
# =========================
tab_dash, tab_analyses, tab_escrow, tab_compte, tab_gestion, tab_visa, tab_export = st.tabs([
    "📊 Dashboard",
    "📈 Analyses",
    "🏦 Escrow",
    "👤 Compte client",
    "🧾 Gestion",
    "📄 Visa (aperçu)",
    "💾 Export"
])

# =========================
# 📊 DASHBOARD
# =========================
with tab_dash:
    st.subheader("📊 Dashboard")
    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # KPIs (réduits visuellement via colonnes)
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(df_all)}")
        k2.metric("Honoraires+Frais", _fmt_money(df_all["Total (US $)"].sum()))
        k3.metric("Payé", _fmt_money(df_all["Payé"].sum()))
        k4.metric("Solde", _fmt_money(df_all["Solde"].sum()))
        nb_env = int((df_all["Dossiers envoyé"] == 1).sum())
        pct_env = int(round((nb_env / max(1, len(df_all))) * 100, 0))
        k5.metric("Envoyés (%)", f"{pct_env}%")

        # Filtres
        st.markdown("#### 🎛️ Filtres")
        cats  = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subs  = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3 = st.columns(3)
        fc = a1.multiselect("Catégories", cats, default=[])
        fs = a2.multiselect("Sous-catégories", subs, default=[])
        fv = a3.multiselect("Visa", visas, default=[])

        view = df_all.copy()
        if fc: view = view[view["Categories"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        # Graph 1 : Nombre de dossiers par catégorie
        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if not view.empty and "Categories" in view.columns:
            g1 = view["Categories"].value_counts().reset_index()
            g1.columns = ["Catégorie","Nombre"]
            st.bar_chart(g1.set_index("Catégorie"))
        else:
            st.write("0")

        # Graph 2 : Flux mensuels (honoraires, frais, payé, solde)
        st.markdown("#### 💵 Flux par mois")
        if not view.empty:
            tmp = view.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)[["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]].sum().sort_values("Mois")
            gm = gm.set_index("Mois")
            st.line_chart(gm)
        else:
            st.write("Aucune donnée après filtres.")

        # Détails
        st.markdown("#### 📋 Détails (après filtres)")
        show = ["Dossier N","ID_Client","Nom","Date","Mois","Categories","Sous-categorie","Visa",
                "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
                "Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé","RFE"]
        show = [c for c in show if c in view.columns]
        view2 = view.copy()
        # Format monnaie
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde"]:
            if c in view2.columns:
                view2[c] = view2[c].apply(_fmt_money)
        # Date propre
        if "Date" in view2.columns:
            try:
                view2["Date"] = pd.to_datetime(view2["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                view2["Date"] = view2["Date"].astype(str)

        st.dataframe(view2[show].reset_index(drop=True), use_container_width=True, height=420)

# =========================
# 📈 ANALYSES
# =========================
with tab_analyses:
    st.subheader("📈 Analyses")
    if df_all.empty:
        st.info("Aucune donnée.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(df_all["Categories"].dropna().astype(str).unique().tolist())
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist())
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist())

        b1, b2, b3, b4, b5 = st.columns(5)
        fy = b1.multiselect("Année", yearsA, default=[])
        fm = b2.multiselect("Mois (MM)", monthsA, default=[])
        fc = b3.multiselect("Catégorie", catsA, default=[])
        fs = b4.multiselect("Sous-catégorie", subsA, default=[])
        fv = b5.multiselect("Visa", visasA, default=[])

        A = df_all.copy()
        if fy: A = A[A["_Année_"].isin(fy)]
        if fm: A = A[A["Mois"].astype(str).isin(fm)]
        if fc: A = A[A["Categories"].astype(str).isin(fc)]
        if fs: A = A[A["Sous-categorie"].astype(str).isin(fs)]
        if fv: A = A[A["Visa"].astype(str).isin(fv)]

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(A)}")
        k2.metric("Honoraires", _fmt_money(A["Montant honoraires (US $)"].sum()))
        k3.metric("Payé", _fmt_money(A["Payé"].sum()))
        k4.metric("Solde", _fmt_money(A["Solde"].sum()))

        # % par Catégorie / Sous-catégorie
        st.markdown("#### % par catégorie")
        if not A.empty:
            totalA = max(1.0, float(A["Total (US $)"].sum()))
            part = (A.groupby("Categories", as_index=False)["Total (US $)"].sum()
                      .assign(Part=lambda df: (df["Total (US $)"]/totalA*100.0).round(1)))
            st.dataframe(part, use_container_width=True)
        st.markdown("#### % par sous-catégorie")
        if not A.empty and "Sous-categorie" in A.columns:
            totalA = max(1.0, float(A["Total (US $)"].sum()))
            part2 = (A.groupby("Sous-categorie", as_index=False)["Total (US $)"].sum()
                        .assign(Part=lambda df: (df["Total (US $)"]/totalA*100.0).round(1)))
            st.dataframe(part2, use_container_width=True)

        # Comparaison simple période (Années A vs B)
        st.markdown("#### Comparaison par année")
        c1, c2 = st.columns(2)
        ya = c1.multiselect("Années A", yearsA, default=yearsA[:1])
        yb = c2.multiselect("Années B", yearsA, default=yearsA[-1:])

        def agg_years(sel):
            if not sel:
                return pd.DataFrame(columns=["_Année_","Total (US $)","Payé","Solde","Dossiers"])
            X = df_all[df_all["_Année_"].isin(sel)]
            return (X.groupby("_Année_", as_index=False)
                     .agg({"Total (US $)":"sum","Payé":"sum","Solde":"sum","ID_Client":"count"})
                     .rename(columns={"ID_Client":"Dossiers"}))

        ga = agg_years(ya)
        gb = agg_years(yb)
        st.write("**Années A**")
        st.dataframe(ga, use_container_width=True)
        st.write("**Années B**")
        st.dataframe(gb, use_container_width=True)

        # Détails
        st.markdown("#### Détails")
        det = A.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde"]:
            if c in det.columns:
                det[c] = det[c].apply(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)
        showA = ["Dossier N","ID_Client","Nom","Date","Mois","Categories","Sous-categorie","Visa",
                 "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
                 "Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé","RFE"]
        showA = [c for c in showA if c in det.columns]
        st.dataframe(det[showA].reset_index(drop=True), use_container_width=True, height=400)



# =========================
# 🏦 ESCROW (synthèse simple)
# =========================
with tab_escrow:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        c1, c2, c3 = st.columns([1,1,1])
        c1.metric("Total", _fmt_money(df_all["Total (US $)"].sum()))
        c2.metric("Payé", _fmt_money(df_all["Payé"].sum()))
        c3.metric("Solde", _fmt_money(df_all["Solde"].sum()))

        agg = (df_all.groupby("Categories", as_index=False)[["Total (US $)","Payé","Solde"]].sum())
        agg["% Payé"] = ((agg["Payé"] / agg["Total (US $)"]).replace([np.inf, -np.inf, np.nan], 0)*100).round(1)
        st.dataframe(agg, use_container_width=True)

        st.caption("NB: Si tu veux un escrow strict, on peut tracer les honoraires perçus avant 'Dossiers envoyé' puis signaler les transferts à faire une fois envoyé.")

# =========================
# 👤 COMPTE CLIENT (timeline & paiements)
# =========================
with tab_compte:
    st.subheader("👤 Compte client")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        ids = sorted(df_all["ID_Client"].dropna().astype(str).unique().tolist())
        sel_id = st.selectbox("Choisir un ID_Client", [""]+ids, index=0)
        if sel_id:
            row = df_all[df_all["ID_Client"].astype(str) == sel_id].iloc[0].to_dict()
            # En-tête
            st.markdown(f"**Nom** : {_safe_str(row.get('Nom',''))} — **Dossier N** : {_safe_str(row.get('Dossier N',''))}")
            st.markdown(f"**Visa** : {_safe_str(row.get('Categories',''))} / {_safe_str(row.get('Sous-categorie',''))} / {_safe_str(row.get('Visa',''))}")

            # Financier
            h1,h2,h3,h4 = st.columns(4)
            total = _to_num(row.get("Montant honoraires (US $)",0)) + _to_num(row.get("Autres frais (US $)",0))
            paye  = _to_num(row.get("Payé",0))
            solde = total - paye
            h1.metric("Honoraires", _fmt_money(row.get("Montant honoraires (US $)",0)))
            h2.metric("Autres frais", _fmt_money(row.get("Autres frais (US $)",0)))
            h3.metric("Payé", _fmt_money(paye))
            h4.metric("Solde", _fmt_money(solde))

            # Timeline statuts + dates (lecture simple: dates stockées dans Commentaires si besoin)
            st.markdown("#### 🧾 Statuts du dossier")
            s1, s2 = st.columns(2)
            s1.write(f"- Dossiers envoyé : {int(_to_num(row.get('Dossiers envoyé',0)))}")
            s1.write(f"- Dossier approuvé : {int(_to_num(row.get('Dossier approuvé',0)))}")
            s1.write(f"- Dossier refusé : {int(_to_num(row.get('Dossier refusé',0)))}")
            s1.write(f"- Dossier Annulé : {int(_to_num(row.get('Dossier Annulé',0)))}")
            s2.write(f"- RFE : {int(_to_num(row.get('RFE',0)))}")
            s2.write(f"- Commentaires : {_safe_str(row.get('Commentaires',''))}")

            # Paiements rapides (Acompte 1 / Acompte 2 + Extra)
            st.markdown("#### 💵 Paiements")
            p1, p2, p3 = st.columns([1,1,1])
            new_pay = p1.number_input("Nouveau paiement (US $)", min_value=0.0, step=10.0, format="%.2f")
            pay_mode = p2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], index=0)
            pay_date = p3.date_input("Date paiement", value=date.today())
            if st.button("➕ Ajouter paiement"):
                # Intégrer au 'Payé' (simple)
                curr = _to_num(row.get("Payé",0))
                new_total = curr + float(new_pay)
                df_all.loc[df_all["ID_Client"].astype(str)==sel_id, "Payé"] = new_total
                df_all.loc[df_all["ID_Client"].astype(str)==sel_id, "Solde"] = (df_all["Total (US $)"] - df_all["Payé"])
                # Sauvegarde fichier Clients si connu
                if clients_path_curr:
                    try:
                        # réécrire tout (écrase)
                        with pd.ExcelWriter(clients_path_curr, engine="openpyxl") as wr:
                            df_all[COLS_EXPECTED + ["Total (US $)","_Année_","Mois","_MoisNum_"]].to_excel(wr, index=False, sheet_name="Clients")
                    except Exception:
                        try:
                            df_all.to_csv(clients_path_curr, index=False)
                        except Exception:
                            pass
                st.success("Paiement ajouté.")
                st.cache_data.clear()
                st.experimental_rerun()



# =========================
# 🧾 GESTION (Ajouter / Modifier / Supprimer)
# =========================
with tab_gestion:
    st.subheader("🧾 Gestion des clients")

    op = st.radio("Action", ["Ajouter","Modifier","Supprimer"], horizontal=True)
    live = df_all.copy()

    # ---------- Ajouter ----------
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")
        c1,c2,c3 = st.columns(3)
        nom  = c1.text_input("Nom", "")
        dval = c2.date_input("Date de création", value=date.today())
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=int(date.today().month)-1)

        # Cascade Visa
        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégories", [""]+cats, index=0)
        sel_sub = ""
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0)
        visa_final = ""
        if sel_cat and sel_sub:
            options = visa_map.get(sel_cat, {}).get(sel_sub, [])
            if options:
                visa_final = st.selectbox("Visa", [""]+options, index=0)
            else:
                visa_final = sel_sub

        f1,f2 = st.columns(2)
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f")
        other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f")
        com   = st.text_area("Commentaires", "")

        s1,s2,s3,s4,s5 = st.columns(5)
        sent  = s1.checkbox("Dossiers envoyé", value=False)
        appr  = s2.checkbox("Dossier approuvé", value=False)
        refus = s3.checkbox("Dossier refusé", value=False)
        ann   = s4.checkbox("Dossier Annulé", value=False)
        rfe   = s5.checkbox("RFE", value=False)

        if st.button("💾 Enregistrer le client"):
            if not nom:
                st.warning("Nom requis.")
                st.stop()
            total = float(honor) + float(other)
            # ID unique
            base = re.sub(r"[^A-Za-z0-9]+","", nom).upper() or "CLIENT"
            did  = f"{base}-{_date_for_widget(dval).strftime('%Y%m%d')}"
            # Dossier N
            used = [int(_to_num(x)) for x in live["Dossier N"].tolist() if _to_num(x) > 0]
            start = 13057
            nxt = start
            us = set(used)
            while nxt in us:
                nxt += 1
            new_row = {
                "ID_Client": did, "Dossier N": nxt, "Nom": nom, "Date": dval, "Categories": sel_cat,
                "Sous-categorie": sel_sub, "Visa": (visa_final or sel_sub),
                "Montant honoraires (US $)": float(honor), "Autres frais (US $)": float(other),
                "Payé": 0.0, "Solde": total, "Acompte 1": 0.0, "Acompte 2": 0.0,
                "RFE": 1 if rfe else 0, "Dossiers envoyé": 1 if sent else 0, "Dossier approuvé": 1 if appr else 0,
                "Dossier refusé": 1 if refus else 0, "Dossier Annulé": 1 if ann else 0,
                "Commentaires": com
            }
            live = pd.concat([live, pd.DataFrame([new_row])], ignore_index=True)
            # recalc annexes
            live["Total (US $)"] = live["Montant honoraires (US $)"] + live["Autres frais (US $)"]
            live["Solde"] = live["Total (US $)"] - live["Payé"]

            # sauvegarde
            if clients_path_curr:
                try:
                    with pd.ExcelWriter(clients_path_curr, engine="openpyxl") as wr:
                        live[COLS_EXPECTED + ["Total (US $)","_Année_","Mois","_MoisNum_"]].to_excel(wr, index=False, sheet_name="Clients")
                except Exception:
                    try:
                        live.to_csv(clients_path_curr, index=False)
                    except Exception:
                        pass
            st.success("Client ajouté.")
            save_last_paths(clients_path_curr, visa_path_curr)
            st.cache_data.clear()
            st.experimental_rerun()

    # ---------- Modifier ----------
    if op == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if live.empty:
            st.info("Aucun client.")
        else:
            names = sorted(live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(live["ID_Client"].dropna().astype(str).unique().tolist())
            m1, m2 = st.columns(2)
            tname = m1.selectbox("Nom", [""]+names, index=0)
            tid   = m2.selectbox("ID_Client", [""]+ids, index=0)
            mask = None
            if tid: mask = (live["ID_Client"].astype(str) == tid)
            elif tname: mask = (live["Nom"].astype(str) == tname)
            if mask is None or not mask.any():
                st.stop()
            idx = live[mask].index[0]
            row = live.loc[idx].to_dict()

            d1,d2,d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")))
            dval = d2.date_input("Date de création", value=_date_for_widget(row.get("Date")))
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=(int(_safe_str(row.get("Mois","01")) or "1")-1))

            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            curr_cat = _safe_str(row.get("Categories",""))
            sel_cat = st.selectbox("Catégories", [""]+cats, index=(cats.index(curr_cat)+1 if curr_cat in cats else 0))
            curr_sub = _safe_str(row.get("Sous-categorie",""))
            sel_sub = ""
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=(subs.index(curr_sub)+1 if curr_sub in subs else 0))
            curr_visa = _safe_str(row.get("Visa",""))
            visa_final = curr_visa
            if sel_cat and sel_sub:
                options = visa_map.get(sel_cat, {}).get(sel_sub, [])
                if options:
                    visa_final = st.selectbox("Visa", [""]+options,
                                              index=(options.index(curr_visa)+1 if curr_visa in options else 0))
                else:
                    visa_final = sel_sub

            f1,f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0,
                                    value=float(_to_num(row.get("Montant honoraires (US $)",0))), step=50.0, format="%.2f")
            other = f2.number_input("Autres frais (US $)", min_value=0.0,
                                    value=float(_to_num(row.get("Autres frais (US $)",0))), step=20.0, format="%.2f")
            com = st.text_area("Commentaires", _safe_str(row.get("Commentaires","")))

            s1,s2,s3,s4,s5 = st.columns(5)
            sent  = s1.checkbox("Dossiers envoyé", value=bool(_to_num(row.get("Dossiers envoyé",0))==1))
            appr  = s2.checkbox("Dossier approuvé", value=bool(_to_num(row.get("Dossier approuvé",0))==1))
            refus = s3.checkbox("Dossier refusé", value=bool(_to_num(row.get("Dossier refusé",0))==1))
            ann   = s4.checkbox("Dossier Annulé", value=bool(_to_num(row.get("Dossier Annulé",0))==1))
            rfe   = s5.checkbox("RFE", value=bool(_to_num(row.get("RFE",0))==1))

            if st.button("💾 Enregistrer modifications"):
                live.at[idx,"Nom"] = nom
                live.at[idx,"Date"] = dval
                live.at[idx,"Categories"] = sel_cat
                live.at[idx,"Sous-categorie"] = sel_sub
                live.at[idx,"Visa"] = visa_final
                live.at[idx,"Montant honoraires (US $)"] = float(honor)
                live.at[idx,"Autres frais (US $)"] = float(other)
                live.at[idx,"Total (US $)"] = float(honor)+float(other)
                live.at[idx,"Solde"] = live.at[idx,"Total (US $)"] - _to_num(live.at[idx,"Payé"])
                live.at[idx,"Commentaires"] = com
                live.at[idx,"Dossiers envoyé"] = 1 if sent else 0
                live.at[idx,"Dossier approuvé"] = 1 if appr else 0
                live.at[idx,"Dossier refusé"] = 1 if refus else 0
                live.at[idx,"Dossier Annulé"] = 1 if ann else 0
                live.at[idx,"RFE"] = 1 if rfe else 0

                if clients_path_curr:
                    try:
                        with pd.ExcelWriter(clients_path_curr, engine="openpyxl") as wr:
                            live[COLS_EXPECTED + ["Total (US $)","_Année_","Mois","_MoisNum_"]].to_excel(wr, index=False, sheet_name="Clients")
                    except Exception:
                        try:
                            live.to_csv(clients_path_curr, index=False)
                        except Exception:
                            pass
                st.success("Modifications enregistrées.")
                save_last_paths(clients_path_curr, visa_path_curr)
                st.cache_data.clear()
                st.experimental_rerun()

    # ---------- Supprimer ----------
    if op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if live.empty:
            st.info("Aucun client.")
        else:
            ids = sorted(live["ID_Client"].dropna().astype(str).unique().tolist())
            tid = st.selectbox("ID_Client", [""]+ids, index=0)
            if tid:
                r = live[live["ID_Client"].astype(str)==tid].iloc[0].to_dict()
                st.write({"Dossier N": r.get("Dossier N",""), "Nom": r.get("Nom",""), "Visa": r.get("Visa","")})
                if st.button("❗ Confirmer la suppression"):
                    newdf = live[live["ID_Client"].astype(str)!=tid].copy()
                    if clients_path_curr:
                        try:
                            with pd.ExcelWriter(clients_path_curr, engine="openpyxl") as wr:
                                newdf[COLS_EXPECTED + ["Total (US $)","_Année_","Mois","_MoisNum_"]].to_excel(wr, index=False, sheet_name="Clients")
                        except Exception:
                            try:
                                newdf.to_csv(clients_path_curr, index=False)
                            except Exception:
                                pass
                    st.success("Client supprimé.")
                    save_last_paths(clients_path_curr, visa_path_curr)
                    st.cache_data.clear()
                    st.experimental_rerun()

# =========================
# 📄 VISA (aperçu)
# =========================
with tab_visa:
    st.subheader("📄 Visa — aperçu")
    if df_visa_raw.empty:
        st.info("Aucun fichier Visa.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True)

# =========================
# 💾 EXPORT
# =========================
with tab_export:
    st.subheader("💾 Export")
    c1, c2 = st.columns(2)
    if c1.button("Exporter Clients (xlsx)"):
        if df_all.empty:
            st.warning("Aucun client.")
        else:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as wr:
                df_all[COLS_EXPECTED + ["Total (US $)","_Année_","Mois","_MoisNum_"]].to_excel(wr, index=False, sheet_name="Clients")
            st.download_button("⬇️ Télécharger Clients.xlsx", data=buf.getvalue(), file_name="Clients_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if c2.button("Exporter Visa (xlsx)"):
        if df_visa_raw.empty:
            st.warning("Aucun Visa.")
        else:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as wr:
                df_visa_raw.to_excel(wr, index=False, sheet_name="Visa")
            st.download_button("⬇️ Télécharger Visa.xlsx", data=buf.getvalue(), file_name="Visa_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")