# app.py — Visa Manager (Partie 1/4)

from __future__ import annotations
import os, json, re, zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Dict, Any, List, Tuple, Optional

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------------------
# CONSTANTES — noms de colonnes conformes à ton fichier (ne pas changer)
# --------------------------------------------------------------------------------------
COL_ID          = "ID_Client"
COL_DOSSIER     = "Dossier N"
COL_NOM         = "Nom"
COL_DATE        = "Date"
COL_CAT         = "Categories"
COL_SUB         = "Sous-categorie"
COL_VISA        = "Visa"
COL_HONO        = "Montant honoraires (US $)"
COL_AUTRE       = "Autres frais (US $)"
COL_PAYE        = "Payé"
COL_SOLDE       = "Solde"
COL_ACPT1       = "Acompte 1"
COL_ACPT2       = "Acompte 2"
COL_RFE         = "RFE"
COL_SENT        = "Dossiers envoyé"       # respecter ton en-tête exact
COL_OK          = "Dossier approuvé"
COL_REFUS       = "Dossier refusé"
COL_ANNUL       = "Dossier Annulé"
COL_COMM        = "Commentaires"

REQUIRED_COLS = [
    COL_ID, COL_DOSSIER, COL_NOM, COL_DATE, COL_CAT, COL_SUB, COL_VISA,
    COL_HONO, COL_PAYE, COL_SOLDE, COL_ACPT1, COL_ACPT2, COL_RFE,
    COL_SENT, COL_OK, COL_REFUS, COL_ANNUL, COL_COMM, COL_AUTRE
]

APP_TITLE   = "🛂 Visa Manager"
SESSION_FILE = "last_session.json"  # persistance des derniers fichiers

# --------------------------------------------------------------------------------------
# PETITS OUTILS
# --------------------------------------------------------------------------------------
def _fmt_money(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def _to_float(x) -> float:
    try:
        if pd.isna(x):
            return 0.0
        if isinstance(x, str):
            s = x.replace(" ", "").replace("\u202f", "").replace(",", ".")
            s = re.sub(r"[^\d\.\-]", "", s)
            return float(s) if s else 0.0
        return float(x)
    except Exception:
        return 0.0

def _safe_bool(x) -> int:
    try:
        if isinstance(x, (bool, int)):
            return int(bool(x))
        if isinstance(x, str):
            sx = x.strip().lower()
            return 1 if sx in {"1","true","vrai","oui","yes","y"} else 0
        return 0
    except Exception:
        return 0

def _safe_date(x) -> Optional[date]:
    if isinstance(x, date) and not isinstance(x, datetime):
        return x
    try:
        d = pd.to_datetime(x, errors="coerce")
        if pd.isna(d):
            return None
        return d.date()
    except Exception:
        return None

def _date_for_widget(x) -> date:
    d = _safe_date(x)
    return d if d else date.today()

def _norm_name(s: str) -> str:
    s = str(s or "").strip()
    s = s.lower()
    s = re.sub(r"[^a-z0-9\- ]+", " ", s)
    s = re.sub(r"\s+", "-", s).strip("-")
    return s or "client"

def _make_id_if_needed(row: pd.Series) -> str:
    cid = str(row.get(COL_ID, "")).strip()
    if cid:
        return cid
    base = _norm_name(row.get(COL_NOM, "client"))
    d    = _safe_date(row.get(COL_DATE)) or date.today()
    return f"{base}-{d.strftime('%Y%m%d')}"

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    try:
        cur = pd.to_numeric(df.get(COL_DOSSIER, pd.Series([], dtype=float)), errors="coerce")
        mx  = int(cur.dropna().max()) if cur.notna().any() else (start - 1)
        return max(start, mx + 1)
    except Exception:
        return start

def _ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    for c in REQUIRED_COLS:
        if c not in df.columns:
            df[c] = pd.NA
    return df[REQUIRED_COLS].copy()

# --------------------------------------------------------------------------------------
# PERSISTANCE DES CHEMINS (mêmes après redémarrage)
# --------------------------------------------------------------------------------------
def load_last_paths() -> Dict[str, str]:
    try:
        if os.path.exists(SESSION_FILE):
            with open(SESSION_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return {k: str(v) for k,v in data.items()}
    except Exception:
        pass
    return {}

def save_last_paths(d: Dict[str, str]) -> None:
    try:
        with open(SESSION_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def sid(key: str) -> str:
    # espace de nom pour éviter les collisions de clés Streamlit
    return f"vmgr_{key}"

# --------------------------------------------------------------------------------------
# LECTURE / ECRITURE FICHIERS
# --------------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def read_any(path: str, sheet: Optional[str] = None) -> pd.DataFrame:
    if not path:
        return pd.DataFrame()
    ext = os.path.splitext(path)[1].lower()
    if ext in {".xlsx", ".xls"}:
        # feuille "Clients" par défaut si présente
        try:
            if sheet is None:
                xl = pd.ExcelFile(path)
                target = "Clients" if "Clients" in xl.sheet_names else xl.sheet_names[0]
                df = pd.read_excel(path, sheet_name=target)
            else:
                df = pd.read_excel(path, sheet_name=sheet)
        except Exception:
            df = pd.read_excel(path)  # fallback: première feuille
    else:
        # CSV: tentative ; ou point-virgule
        try:
            df = pd.read_csv(path)
        except Exception:
            df = pd.read_csv(path, sep=";")
    return df

def write_any(df: pd.DataFrame, path: str, sheet: Optional[str] = None) -> None:
    ext = os.path.splitext(path)[1].lower()
    if ext in {".xlsx", ".xls"}:
        with pd.ExcelWriter(path, engine="openpyxl") as wr:
            target = sheet or "Clients"
            df.to_excel(wr, sheet_name=target, index=False)
    else:
        df.to_csv(path, index=False)

# --------------------------------------------------------------------------------------
# CHARGEMENT INITIAL (barre latérale + mémoire persistante)
# --------------------------------------------------------------------------------------
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title(APP_TITLE)

paths = load_last_paths()
if "paths" not in st.session_state:
    st.session_state["paths"] = paths

with st.sidebar:
    st.header("📂 Fichiers")
    mode = st.radio("Mode de chargement", ["Un fichier (Clients)", "Deux fichiers (Clients + Visa)"], key=sid("mode"))

    # Sélecteurs / upload
    up_clients = st.file_uploader("Clients (xlsx/csv)", type=["xlsx","xls","csv"], key=sid("up_clients"))
    up_visa    = None
    if mode == "Deux fichiers (Clients + Visa)":
        up_visa = st.file_uploader("Visa (xlsx/csv)", type=["xlsx","xls","csv"], key=sid("up_visa"))

    # Chemins mémorisés
    st.caption("Derniers chemins mémorisés :")
    last_clients_path = st.session_state["paths"].get("clients_path","")
    last_visa_path    = st.session_state["paths"].get("visa_path","")
    st.text_input("Dernier Clients", value=last_clients_path, key=sid("last_clients"), disabled=True)
    st.text_input("Dernier Visa", value=last_visa_path, key=sid("last_visa"), disabled=True)

    # Choix d’un chemin de sauvegarde
    st.markdown("**Chemin de sauvegarde** (sur ton PC / Drive / OneDrive) :")
    save_clients_path = st.text_input("Sauvegarder Clients vers…", value=last_clients_path or "clients.xlsx", key=sid("save_clients"))
    save_visa_path    = st.text_input("Sauvegarder Visa vers…", value=last_visa_path or "visa.xlsx", key=sid("save_visa"))

    # Boutons d’action sur chemins
    if st.button("💾 Mémoriser ces chemins", key=sid("mem_paths")):
        st.session_state["paths"]["clients_path"] = save_clients_path
        if mode == "Deux fichiers (Clients + Visa)":
            st.session_state["paths"]["visa_path"] = save_visa_path
        save_last_paths(st.session_state["paths"])
        st.success("Chemins mémorisés.")

# Charger les DataFrames selon priorité : uploader > chemins mémorisés > vide
def _df_from_uploader_or_path(uploader, path, sheet=None) -> Tuple[pd.DataFrame,str]:
    if uploader is not None:
        # contenu uploadé
        try:
            ext = os.path.splitext(uploader.name)[1].lower()
            if ext in {".xlsx",".xls"}:
                xl = pd.ExcelFile(uploader)
                target = "Clients" if sheet is None and "Clients" in xl.sheet_names else (sheet or xl.sheet_names[0])
                df = pd.read_excel(uploader, sheet_name=target)
            else:
                try:
                    df = pd.read_csv(uploader)
                except Exception:
                    df = pd.read_csv(uploader, sep=";")
            tmp_path = os.path.join(".", f"upload_{uploader.name}")
            write_any(df, tmp_path, sheet="Clients")
            return df, tmp_path
        except Exception:
            pass
    # sinon, chemin mémorisé si présent
    if path and os.path.exists(path):
        return read_any(path, sheet=sheet), path
    # vide
    return pd.DataFrame(), ""

# Clients
df_clients_raw, clients_path = _df_from_uploader_or_path(up_clients, st.session_state["paths"].get("clients_path",""), sheet="Clients")
if not df_clients_raw.empty:
    st.session_state["paths"]["clients_path"] = clients_path
    save_last_paths(st.session_state["paths"])

# Visa
df_visa_raw, visa_path = pd.DataFrame(), ""
if mode == "Deux fichiers (Clients + Visa)":
    df_visa_raw, visa_path = _df_from_uploader_or_path(up_visa, st.session_state["paths"].get("visa_path",""), sheet=None)
    if not df_visa_raw.empty:
        st.session_state["paths"]["visa_path"] = visa_path
        save_last_paths(st.session_state["paths"])

# Normalisation initiale Clients
def normalize_clients(df0: pd.DataFrame) -> pd.DataFrame:
    if df0.empty:
        return pd.DataFrame(columns=REQUIRED_COLS)
    df = df0.copy()
    df = _ensure_columns(df)

    # types
    df[COL_DATE]  = df[COL_DATE].apply(_safe_date)
    for c in [COL_HONO, COL_AUTRE, COL_PAYE, COL_SOLDE, COL_ACPT1, COL_ACPT2]:
        df[c] = df[c].apply(_to_float)

    for c in [COL_RFE, COL_SENT, COL_OK, COL_REFUS, COL_ANNUL]:
        df[c] = df[c].apply(_safe_bool)

    # ID et Dossier
    df[COL_ID]      = df.apply(_make_id_if_needed, axis=1)
    # Dossier N si manquant → incrémental
    mask_empty_dossier = ~pd.to_numeric(df[COL_DOSSIER], errors="coerce").notna()
    if mask_empty_dossier.any():
        base = _next_dossier(df)
        idxs = df[mask_empty_dossier].index.tolist()
        for k, i in enumerate(idxs):
            df.at[i, COL_DOSSIER] = base + k

    # recalcul Solde si incohérences
    total_calc = df[COL_HONO] + df[COL_AUTRE]
    paid_calc  = df[COL_PAYE].fillna(0.0)
    df[COL_SOLDE] = (total_calc - paid_calc).clip(lower=0.0)

    return df

df_all = normalize_clients(df_clients_raw)


# ================================
# 📊 DASHBOARD & 📈 ANALYSES
# ================================

st.markdown("### 📄 Fichiers chargés")
c1, c2 = st.columns(2)
c1.write(f"**Clients** : `{clients_path or '—'}`")
c2.write(f"**Visa** : `{visa_path or '—'}`")

tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Compte client", "🧾 Gestion", "📄 Visa (aperçu)", "💾 Export"])

# --------------------------------
# 📊 DASHBOARD (vue synthétique)
# --------------------------------
with tabs[0]:
    st.subheader("📊 Dashboard")
    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # KPI compacts
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        total_files = len(df_all)
        honor_total = float((df_all[COL_HONO] + df_all[COL_AUTRE]).sum())
        paye_total  = float(df_all[COL_PAYE].sum())
        solde_total = float(df_all[COL_SOLDE].sum())
        sent_ratio  = (df_all[COL_SENT].sum() / max(1, total_files)) * 100.0

        k1.metric("Dossiers", f"{total_files}")
        k2.metric("Honoraires+Frais", _fmt_money(honor_total))
        k3.metric("Payé", _fmt_money(paye_total))
        k4.metric("Solde", _fmt_money(solde_total))
        k5.metric("Envoyés (%)", f"{sent_ratio:.0f}%")

        # Filtres rapides
        st.markdown("#### 🎛️ Filtres")
        f1, f2, f3 = st.columns(3)
        cats = sorted(df_all[COL_CAT].dropna().astype(str).unique().tolist())
        subs = sorted(df_all[COL_SUB].dropna().astype(str).unique().tolist())
        visas = sorted(df_all[COL_VISA].dropna().astype(str).unique().tolist())
        fc = f1.multiselect("Catégories", cats, default=[])
        fs = f2.multiselect("Sous-catégories", subs, default=[])
        fv = f3.multiselect("Visa", visas, default=[])

        view = df_all.copy()
        if fc: view = view[view[COL_CAT].astype(str).isin(fc)]
        if fs: view = view[view[COL_SUB].astype(str).isin(fs)]
        if fv: view = view[view[COL_VISA].astype(str).isin(fv)]

        # Graphique barre — nombre dossiers par catégorie
        if not view.empty:
            st.markdown("#### 📦 Nombre de dossiers par catégorie")
            vc = view[COL_CAT].value_counts().reset_index()
            vc.columns = ["Catégorie", "Nombre"]
            st.bar_chart(vc.set_index("Catégorie"))

            # Graphique lignes — honoraires+frais par mois (si Date exploitable)
            if COL_DATE in view.columns:
                dfm = view.copy()
                dfm["_Mois_"] = pd.to_datetime(dfm[COL_DATE], errors="coerce").dt.to_period("M").astype(str)
                gm = dfm.groupby("_Mois_", as_index=False)[[COL_HONO, COL_AUTRE, COL_PAYE, COL_SOLDE]].sum()
                gm = gm.sort_values("_Mois_")
                st.markdown("#### 💵 Flux par mois")
                st.line_chart(gm.set_index("_Mois_")[[COL_HONO, COL_AUTRE, COL_PAYE, COL_SOLDE]])

        # Tableau
        st.markdown("#### 📋 Détails (après filtres)")
        show_cols = [c for c in [
            COL_DOSSIER, COL_ID, COL_NOM, COL_DATE, COL_CAT, COL_SUB, COL_VISA,
            COL_HONO, COL_AUTRE, COL_PAYE, COL_SOLDE, COL_SENT, COL_OK, COL_REFUS, COL_ANNUL, COL_RFE, COL_COMM
        ] if c in view.columns]
        st.dataframe(view[show_cols].reset_index(drop=True), use_container_width=True)

# --------------------------------
# 📈 ANALYSES (filtres + comparaisons)
# --------------------------------
with tabs[1]:
    st.subheader("📈 Analyses")
    if df_all.empty:
        st.info("Aucun client chargé.")
    else:
        # Préparation (année/mois)
        tmp = df_all.copy()
        tmp["_Date_"] = pd.to_datetime(tmp[COL_DATE], errors="coerce")
        tmp["_Année_"] = tmp["_Date_"].dt.year
        tmp["_Mois_"]  = tmp["_Date_"].dt.month

        a1, a2, a3, a4, a5 = st.columns(5)
        years  = sorted([int(x) for x in tmp["_Année_"].dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1,13)]
        cats   = sorted(tmp[COL_CAT].dropna().astype(str).unique().tolist())
        subs   = sorted(tmp[COL_SUB].dropna().astype(str).unique().tolist())
        visas  = sorted(tmp[COL_VISA].dropna().astype(str).unique().tolist())

        fy = a1.multiselect("Année", years, default=[])
        fm = a2.multiselect("Mois (MM)", months, default=[])
        fc = a3.multiselect("Catégorie", cats, default=[])
        fs = a4.multiselect("Sous-catégorie", subs, default=[])
        fv = a5.multiselect("Visa", visas, default=[])

        ff = tmp.copy()
        if fy: ff = ff[ff["_Année_"].isin(fy)]
        if fm: ff = ff[ff["_Mois_"].astype(str).str.zfill(2).isin(fm)]
        if fc: ff = ff[ff[COL_CAT].astype(str).isin(fc)]
        if fs: ff = ff[ff[COL_SUB].astype(str).isin(fs)]
        if fv: ff = ff[ff[COL_VISA].astype(str).isin(fv)]

        # KPI
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(ff)}")
        k2.metric("Honoraires", _fmt_money(float(ff[COL_HONO].sum())))
        k3.metric("Autres frais", _fmt_money(float(ff[COL_AUTRE].sum())))
        k4.metric("Payé", _fmt_money(float(ff[COL_PAYE].sum())))
        k5.metric("Solde", _fmt_money(float(ff[COL_SOLDE].sum())))

        # % par catégorie / sous-catégorie
        st.markdown("#### % par catégorie")
        gcat = ff.groupby(COL_CAT, as_index=False)[[COL_HONO, COL_AUTRE, COL_PAYE, COL_SOLDE]].sum()
        tot = max(1.0, float(gcat[COL_HONO].sum() + gcat[COL_AUTRE].sum()))
        gcat["% CA"] = 100.0 * (gcat[COL_HONO] + gcat[COL_AUTRE]) / tot
        st.dataframe(gcat.sort_values("% CA", ascending=False), use_container_width=True)

        st.markdown("#### % par sous-catégorie")
        gsub = ff.groupby(COL_SUB, as_index=False)[[COL_HONO, COL_AUTRE, COL_PAYE, COL_SOLDE]].sum()
        tot2 = max(1.0, float(gsub[COL_HONO].sum() + gsub[COL_AUTRE].sum()))
        gsub["% CA"] = 100.0 * (gsub[COL_HONO] + gsub[COL_AUTRE]) / tot2
        st.dataframe(gsub.sort_values("% CA", ascending=False), use_container_width=True)

        # Comparaison période A vs B
        st.markdown("### 🔀 Comparaison de périodes")
        ca, cb = st.columns(2)
        with ca:
            st.caption("Période A")
            ya = st.multiselect("Année (A)", years, default=[])
            ma = st.multiselect("Mois (A)", months, default=[])
        with cb:
            st.caption("Période B")
            yb = st.multiselect("Année (B)", years, default=[])
            mb = st.multiselect("Mois (B)", months, default=[])

        def _subset(y, m):
            d = tmp.copy()
            if y: d = d[d["_Année_"].isin(y)]
            if m: d = d[d["_Mois_"].astype(str).str.zfill(2).isin(m)]
            return d

        A = _subset(ya, ma)
        B = _subset(yb, mb)

        c1, c2, c3 = st.columns(3)
        c1.metric("CA A", _fmt_money(float((A[COL_HONO]+A[COL_AUTRE]).sum())))
        c2.metric("CA B", _fmt_money(float((B[COL_HONO]+B[COL_AUTRE]).sum())))
        delta = float((B[COL_HONO]+B[COL_AUTRE]).sum() - (A[COL_HONO]+A[COL_AUTRE]).sum())
        c3.metric("Δ B - A", _fmt_money(delta))

        # Détails
        st.markdown("### 🧾 Détails filtrés")
        detail = ff.copy()
        detail_show = [c for c in [
            COL_DOSSIER, COL_ID, COL_NOM, COL_DATE, COL_CAT, COL_SUB, COL_VISA,
            COL_HONO, COL_AUTRE, COL_PAYE, COL_SOLDE, COL_SENT, COL_OK, COL_REFUS, COL_ANNUL, COL_RFE, COL_COMM
        ] if c in detail.columns]
        st.dataframe(detail[detail_show].reset_index(drop=True), use_container_width=True)



# ================================
# 🏦 ESCROW (synthèse & alertes)
# ================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse & alertes")
    if df_all.empty:
        st.info("Aucun client chargé.")
    else:
        work = df_all.copy()
        work["Total (US $)"] = work[COL_HONO].fillna(0) + work[COL_AUTRE].fillna(0)

        # Règle simple :
        # - Dossiers NON envoyés  -> Payé reste en ESCROW
        # - Dossiers envoyés      -> Payé est "à transférer" (alerte)
        escrow_non_envoye = float(work.loc[work[COL_SENT] == 0, COL_PAYE].sum())
        a_transferer      = float(work.loc[work[COL_SENT] == 1, COL_PAYE].sum())

        # KPI compacts
        c1, c2, c3 = st.columns([1,1,1])
        c1.metric("En ESCROW (non envoyés)", _fmt_money(escrow_non_envoye))
        c2.metric("À transférer (envoyés)", _fmt_money(a_transferer))
        c3.metric("Solde total dossiers", _fmt_money(float(work[COL_SOLDE].sum())))

        st.markdown("#### 🔔 Dossiers envoyés — à transférer")
        to_claim = work[(work[COL_SENT] == 1) & (work[COL_PAYE] > 0)]
        if to_claim.empty:
            st.success("Aucun dossier envoyé avec encaissement en attente de transfert.")
        else:
            show_cols = [c for c in [COL_DOSSIER, COL_ID, COL_NOM, COL_CAT, COL_SUB, COL_VISA, COL_PAYE, COL_DATE] if c in to_claim.columns]
            st.dataframe(to_claim[show_cols].reset_index(drop=True), use_container_width=True)

        st.caption("Règle ESCROW appliquée : les encaissements restent en attente tant que le dossier n’est pas « envoyé ». \
Quand le dossier passe à « envoyé », le total payé apparaît comme À transférer.")


# ==========================================
# 👤 COMPTE CLIENT — détail et règlements
# ==========================================
with tabs[3]:
    st.subheader("👤 Compte client — détail & règlements")
    if df_all.empty:
        st.info("Aucun client chargé.")
    else:
        # Sélection du client
        ids  = df_all[COL_ID].dropna().astype(str).unique().tolist()
        noms = df_all[COL_NOM].dropna().astype(str).unique().tolist()
        s1, s2 = st.columns(2)
        sel_id  = s1.selectbox("ID_Client", [""] + sorted(ids), index=0, key=sid("acct_id"))
        sel_nom = s2.selectbox("Nom", [""] + sorted(noms), index=0, key=sid("acct_nom"))

        mask = None
        if sel_id:
            mask = (df_all[COL_ID].astype(str) == sel_id)
        elif sel_nom:
            mask = (df_all[COL_NOM].astype(str) == sel_nom)

        if mask is None or not mask.any():
            st.stop()

        row = df_all[mask].iloc[0].copy()

        # Récap
        b1, b2, b3, b4 = st.columns(4)
        total_calc = float(row.get(COL_HONO, 0) + row.get(COL_AUTRE, 0))
        b1.metric("Honoraires", _fmt_money(float(row.get(COL_HONO, 0))))
        b2.metric("Autres frais", _fmt_money(float(row.get(COL_AUTRE, 0))))
        b3.metric("Payé", _fmt_money(float(row.get(COL_PAYE, 0))))
        b4.metric("Solde", _fmt_money(float(row.get(COL_SOLDE, max(0.0, total_calc - float(row.get(COL_PAYE, 0)))))))

        st.markdown("#### 📄 Informations dossier")
        ci1, ci2, ci3 = st.columns(3)
        ci1.write(f"**Dossier N** : {row.get(COL_DOSSIER,'')}")
        ci2.write(f"**Catégorie** : {row.get(COL_CAT,'')}")
        ci3.write(f"**Sous-catégorie** : {row.get(COL_SUB,'')}")
        st.write(f"**Visa** : {row.get(COL_VISA,'')}")
        st.write(f"**Date** : {(_safe_date(row.get(COL_DATE)) or '')}")

        st.markdown("#### 📌 Statuts")
        st.write(f"- Dossiers envoyé : {'✅' if int(row.get(COL_SENT,0) or 0)==1 else '❌'}")
        st.write(f"- Dossier approuvé : {'✅' if int(row.get(COL_OK,0) or 0)==1 else '❌'}")
        st.write(f"- Dossier refusé : {'✅' if int(row.get(COL_REFUS,0) or 0)==1 else '❌'}")
        st.write(f"- Dossier annulé : {'✅' if int(row.get(COL_ANNUL,0) or 0)==1 else '❌'}")
        st.write(f"- RFE : {'✅' if int(row.get(COL_RFE,0) or 0)==1 else '❌'}")

        # --- Historique paiements (à partir du champ Commentaires) ---
        st.markdown("### 🧾 Historique des paiements")
        def parse_payments_from_comments(txt: str) -> List[Dict[str, Any]]:
            out: List[Dict[str,Any]] = []
            if not isinstance(txt, str) or not txt.strip():
                return out
            # Format attendu par l'app quand on ajoute : [YYYY-MM-DD] Paiement MODE $AMOUNT
            for line in txt.splitlines():
                line = line.strip()
                m = re.match(r"\[(\d{4}-\d{2}-\d{2})\]\s*Paiement\s+([A-Za-z]+)\s+\$?([\d\.,]+)", line)
                if m:
                    d, mode, amt = m.groups()
                    out.append({"date": d, "mode": mode.upper(), "amount": _to_float(amt)})
            return out

        pay_hist = parse_payments_from_comments(row.get(COL_COMM, ""))
        if pay_hist:
            st.table(pd.DataFrame(pay_hist))
        else:
            st.info("Aucun paiement historisé dans les commentaires.")

        # --- Ajouter un règlement complémentaire ---
        st.markdown("### ➕ Ajouter un règlement")
        p1, p2, p3, p4 = st.columns([1,1,1,2])
        pay_date = p1.date_input("Date", value=date.today(), key=sid("pay_date"))
        pay_mode = p2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=sid("pay_mode"))
        pay_amt  = p3.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=sid("pay_amt"))
        add_btn  = p4.button("💾 Enregistrer le paiement", key=sid("pay_add_btn"))

        if add_btn:
            if float(pay_amt) <= 0:
                st.warning("Le montant doit être > 0.")
                st.stop()

            # Recharger le fichier source pour éviter les décalages
            live = read_any(clients_path, sheet="Clients")
            if live.empty:
                st.error("Impossible de relire le fichier Clients pour mise à jour.")
                st.stop()

            # Normaliser les colonnes minimales au cas où
            live = _ensure_columns(live)
            # Localiser la ligne
            mask2 = (live[COL_ID].astype(str) == str(row.get(COL_ID)))
            if not mask2.any():
                st.error("Ligne introuvable dans le fichier source.")
                st.stop()

            idx = live[mask2].index[0]

            # Mettre à jour Payé & Solde
            paid0 = _to_float(live.at[idx, COL_PAYE])
            sold0 = _to_float(live.at[idx, COL_SOLDE])
            hono0 = _to_float(live.at[idx, COL_HONO])
            autr0 = _to_float(live.at[idx, COL_AUTRE])
            total = hono0 + autr0

            new_paid = paid0 + float(pay_amt)
            new_paid = min(new_paid, total)  # on ne dépasse pas le total
            new_sold = max(0.0, total - new_paid)

            live.at[idx, COL_PAYE]  = new_paid
            live.at[idx, COL_SOLDE] = new_sold

            # Journaliser dans Commentaires
            comm0 = str(live.at[idx, COL_COMM]) if pd.notna(live.at[idx, COL_COMM]) else ""
            line = f"[{pay_date.strftime('%Y-%m-%d')}] Paiement {pay_mode} {pay_amt:.2f} USD"
            comm1 = (comm0 + "\n" + line).strip() if comm0 else line
            live.at[idx, COL_COMM] = comm1

            # Écrire sur disque
            write_any(live, clients_path, sheet="Clients")
            st.success("Paiement ajouté et fichier clients mis à jour.")
            st.cache_data.clear()
            st.rerun()

        # --- Commentaires libres / note interne ---
        st.markdown("### 📝 Commentaires")
        new_comm = st.text_area("Commentaires (journal libre)", value=str(row.get(COL_COMM,"")), height=140, key=sid("comm_text"))
        if st.button("💾 Sauvegarder les commentaires", key=sid("comm_save_btn")):
            live = read_any(clients_path, sheet="Clients")
            if live.empty:
                st.error("Impossible de relire le fichier Clients pour mise à jour.")
                st.stop()
            live = _ensure_columns(live)
            mask2 = (live[COL_ID].astype(str) == str(row.get(COL_ID)))
            if not mask2.any():
                st.error("Ligne introuvable dans le fichier source.")
                st.stop()
            idx = live[mask2].index[0]
            live.at[idx, COL_COMM] = new_comm
            write_any(live, clients_path, sheet="Clients")
            st.success("Commentaires sauvegardés.")
            st.cache_data.clear()
            st.rerun()



# =====================================================
# 🧾 GESTION (Ajouter / Modifier / Supprimer clients)
# =====================================================
with tabs[4]:
    st.subheader("🧾 Gestion des clients")
    if df_all.empty:
        st.info("Charge d’abord tes fichiers (barre latérale).")
    else:
        # On relit toujours le fichier "source" pour éviter les décalages
        df_live = read_any(clients_path, sheet="Clients")
        df_live = _ensure_columns(df_live)

        op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=sid("crud_op"))

        # -------------------
        # ➕ AJOUTER
        # -------------------
        if op == "Ajouter":
            st.markdown("### ➕ Ajouter un client")

            c1, c2, c3 = st.columns(3)
            nom  = c1.text_input("Nom", key=sid("add_nom"))
            dval = date.today()
            dt   = c2.date_input("Date de création", value=dval, key=sid("add_date"))
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=dval.month-1, key=sid("add_mois"))

            st.markdown("#### 🎯 Choix Visa (cascade)")
            cats = sorted(list(visa_map.keys()))
            sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=sid("add_cat"))

            sel_sub = ""
            options_picked: List[str] = []
            visa_final = ""

            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=sid("add_sub"))

                if sel_sub:
                    # Options de la sous-catégorie
                    vm = visa_map.get(sel_cat, {}).get(sel_sub, {})
                    opts = vm.get("options", [])
                    exclusive = bool(vm.get("exclusive", False))
                    if opts:
                        st.caption("Options disponibles")
                        if exclusive:
                            opt = st.radio("Choisir une option", [""] + opts, horizontal=True, key=sid("add_opt_radio"))
                            options_picked = [opt] if opt else []
                        else:
                            options_picked = []
                            cols = st.columns(min(4, len(opts)) or 1)
                            for i, o in enumerate(opts):
                                if cols[i % len(cols)].checkbox(o, key=sid(f"add_opt_{i}")):
                                    options_picked.append(o)

                    # Construction du libellé Visa (règle simple : Sous-catégorie + première option si exclusive/unique)
                    if options_picked:
                        visa_final = f"{sel_sub} {options_picked[0]}"
                    else:
                        visa_final = sel_sub

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, step=50.0, format="%.2f", key=sid("add_hono"))
            autres = f2.number_input("Autres frais (US $)", min_value=0.0, step=20.0, format="%.2f", key=sid("add_autre"))

            st.markdown("#### 📌 Statuts initiaux (avec dates)")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent   = s1.checkbox("Dossier envoyé", key=sid("add_sent"))
            sent_d = s1.date_input("Date d'envoi", value=None, key=sid("add_sent_d"))
            ok     = s2.checkbox("Dossier approuvé", key=sid("add_ok"))
            ok_d   = s2.date_input("Date d'acceptation", value=None, key=sid("add_ok_d"))
            refus  = s3.checkbox("Dossier refusé", key=sid("add_refus"))
            refus_d= s3.date_input("Date de refus", value=None, key=sid("add_refus_d"))
            annul  = s4.checkbox("Dossier annulé", key=sid("add_annul"))
            annul_d= s4.date_input("Date d'annulation", value=None, key=sid("add_annul_d"))
            rfe    = s5.checkbox("RFE", key=sid("add_rfe"))

            if rfe and not any([sent, ok, refus, annul]):
                st.warning("RFE doit être accompagné d’au moins un autre statut (envoyé, approuvé, refusé, annulé).")

            comm = st.text_area("Commentaires (journal / notes)", key=sid("add_comm"), height=120)

            if st.button("💾 Enregistrer le client", key=sid("add_save")):
                if not nom:
                    st.warning("Le nom est requis."); st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisissez Catégorie et Sous-catégorie."); st.stop()

                total = float(honor) + float(autres)
                paye  = 0.0
                reste = max(0.0, total - paye)

                # ID & dossier
                did = _make_client_id(nom, _date_for_widget(dt))
                dossier_n = _next_dossier(df_live, start=13057)

                # ligne
                new = {
                    COL_ID: did,
                    COL_DOSSIER: dossier_n,
                    COL_NOM: nom,
                    COL_DATE: _date_for_widget(dt),
                    COL_CAT: sel_cat,
                    COL_SUB: sel_sub,
                    COL_VISA: (visa_final or sel_sub),
                    COL_HONO: float(honor),
                    COL_PAYE: float(paye),
                    COL_SOLDE: float(reste),
                    COL_COMM: comm,
                    COL_AUTRE: float(autres),
                    COL_SENT: int(1 if sent else 0),
                    "Date d'envoi": _date_for_widget(sent_d) if sent_d else ("1970-01-01" if sent else ""),
                    COL_OK: int(1 if ok else 0),
                    "Date d'acceptation": _date_for_widget(ok_d) if ok_d else "",
                    COL_REFUS: int(1 if refus else 0),
                    "Date de refus": _date_for_widget(refus_d) if refus_d else "",
                    COL_ANNUL: int(1 if annul else 0),
                    "Date d'annulation": _date_for_widget(annul_d) if annul_d else "",
                    COL_RFE: int(1 if rfe else 0),
                }
                df_new = pd.concat([df_live, pd.DataFrame([new])], ignore_index=True)
                write_any(df_new, clients_path, sheet="Clients")
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()

        # -------------------
        # ✏️ MODIFIER
        # -------------------
        elif op == "Modifier":
            st.markdown("### ✏️ Modifier un client")
            if df_live.empty:
                st.info("Aucun client.")
                st.stop()

            ids  = sorted(df_live[COL_ID].dropna().astype(str).unique().tolist())
            noms = sorted(df_live[COL_NOM].dropna().astype(str).unique().tolist())
            s1, s2 = st.columns(2)
            sel_id  = s1.selectbox("ID_Client", [""] + ids, index=0, key=sid("mod_id"))
            sel_nom = s2.selectbox("Nom", [""] + noms, index=0, key=sid("mod_nom"))

            mask = None
            if sel_id:
                mask = (df_live[COL_ID].astype(str) == sel_id)
            elif sel_nom:
                mask = (df_live[COL_NOM].astype(str) == sel_nom)

            if mask is None or not mask.any():
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()

            d1, d2, d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get(COL_NOM,"")), key=sid("mod_nom_val"))
            # date safe
            pdate = _safe_date(row.get(COL_DATE)) or date.today()
            dt   = d2.date_input("Date de création", value=pdate, key=sid("mod_date"))
            # catégorie / sous-cat
            cats = sorted(list(visa_map.keys()))
            current_cat = _safe_str(row.get(COL_CAT,""))
            sel_cat = d3.selectbox("Catégorie", [""] + cats,
                                   index=(cats.index(current_cat)+1 if current_cat in cats else 0),
                                   key=sid("mod_cat"))

            col_sc1, col_sc2 = st.columns(2)
            sel_sub = _safe_str(row.get(COL_SUB,""))
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = col_sc1.selectbox("Sous-catégorie", [""] + subs,
                                            index=(subs.index(sel_sub)+1 if sel_sub in subs else 0),
                                            key=sid("mod_sub"))
            else:
                sel_sub = col_sc1.selectbox("Sous-catégorie", [sel_sub], index=0, key=sid("mod_sub_disabled"))

            # Options & visa
            visa_final = _safe_str(row.get(COL_VISA, sel_sub))
            options_picked: List[str] = []
            if sel_cat and sel_sub:
                vm = visa_map.get(sel_cat, {}).get(sel_sub, {})
                opts = vm.get("options", [])
                exclusive = bool(vm.get("exclusive", False))
                if opts:
                    col_sc2.caption("Options disponibles")
                    if exclusive:
                        # Préselection sur base du visa existant
                        preset = ""
                        for o in opts:
                            if visa_final.strip().endswith(o):
                                preset = o
                                break
                        opt = col_sc2.radio("Choisir une option", [""] + opts, horizontal=True,
                                            index=([""]+opts).index(preset) if preset in opts else 0,
                                            key=sid("mod_opt_radio"))
                        options_picked = [opt] if opt else []
                    else:
                        options_picked = []
                        cols = col_sc2.columns(min(4, len(opts)) or 1)
                        for i, o in enumerate(opts):
                            checked = False
                            if o and o in visa_final:
                                checked = True
                            if cols[i % len(cols)].checkbox(o, value=checked, key=sid(f"mod_opt_{i}")):
                                options_picked.append(o)

                # Calcul visa final
                if options_picked:
                    visa_final = f"{sel_sub} {options_picked[0]}"
                else:
                    visa_final = sel_sub

            m1, m2, m3 = st.columns(3)
            honor = m1.number_input("Montant honoraires (US $)", min_value=0.0,
                                    value=float(_to_float(row.get(COL_HONO,0))), step=50.0, format="%.2f",
                                    key=sid("mod_hono"))
            autres = m2.number_input("Autres frais (US $)", min_value=0.0,
                                     value=float(_to_float(row.get(COL_AUTRE,0))), step=20.0, format="%.2f",
                                     key=sid("mod_autre"))
            paye = m3.number_input("Payé (US $)", min_value=0.0,
                                   value=float(_to_float(row.get(COL_PAYE,0))), step=10.0, format="%.2f",
                                   key=sid("mod_paye"))
            total = float(honor) + float(autres)
            reste = max(0.0, total - float(paye))

            st.markdown("#### 📌 Statuts (avec dates)")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent   = s1.checkbox("Dossier envoyé", value=bool(int(row.get(COL_SENT,0) or 0)==1), key=sid("mod_sent"))
            sent_d = s1.date_input("Date d'envoi",
                                   value=_date_for_widget(row.get("Date d'envoi")),
                                   key=sid("mod_sent_d"))
            ok     = s2.checkbox("Dossier approuvé", value=bool(int(row.get(COL_OK,0) or 0)==1), key=sid("mod_ok"))
            ok_d   = s2.date_input("Date d'acceptation",
                                   value=_date_for_widget(row.get("Date d'acceptation")),
                                   key=sid("mod_ok_d"))
            refus  = s3.checkbox("Dossier refusé", value=bool(int(row.get(COL_REFUS,0) or 0)==1), key=sid("mod_refus"))
            refus_d= s3.date_input("Date de refus",
                                   value=_date_for_widget(row.get("Date de refus")),
                                   key=sid("mod_refus_d"))
            annul  = s4.checkbox("Dossier annulé", value=bool(int(row.get(COL_ANNUL,0) or 0)==1), key=sid("mod_annul"))
            annul_d= s4.date_input("Date d'annulation",
                                   value=_date_for_widget(row.get("Date d'annulation")),
                                   key=sid("mod_annul_d"))
            rfe    = s5.checkbox("RFE", value=bool(int(row.get(COL_RFE,0) or 0)==1), key=sid("mod_rfe"))

            comm = st.text_area("Commentaires (journal / notes)",
                                value=str(row.get(COL_COMM,"")), height=120, key=sid("mod_comm"))

            if st.button("💾 Sauvegarder", key=sid("mod_save")):
                if not nom:
                    st.warning("Le nom est requis."); st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisissez Catégorie et Sous-catégorie."); st.stop()

                df_live.at[idx, COL_NOM]   = nom
                df_live.at[idx, COL_DATE]  = _date_for_widget(dt)
                df_live.at[idx, COL_CAT]   = sel_cat
                df_live.at[idx, COL_SUB]   = sel_sub
                df_live.at[idx, COL_VISA]  = visa_final
                df_live.at[idx, COL_HONO]  = float(honor)
                df_live.at[idx, COL_PAYE]  = float(paye)
                df_live.at[idx, COL_SOLDE] = float(reste)
                df_live.at[idx, COL_AUTRE] = float(autres)
                df_live.at[idx, COL_COMM]  = comm

                df_live.at[idx, COL_SENT] = int(1 if sent else 0)
                df_live.at[idx, "Date d'envoi"] = _date_for_widget(sent_d) if sent_d else ("1970-01-01" if sent else "")

                df_live.at[idx, COL_OK] = int(1 if ok else 0)
                df_live.at[idx, "Date d'acceptation"] = _date_for_widget(ok_d) if ok_d else ""

                df_live.at[idx, COL_REFUS] = int(1 if refus else 0)
                df_live.at[idx, "Date de refus"] = _date_for_widget(refus_d) if refus_d else ""

                df_live.at[idx, COL_ANNUL] = int(1 if annul else 0)
                df_live.at[idx, "Date d'annulation"] = _date_for_widget(annul_d) if annul_d else ""

                df_live.at[idx, COL_RFE] = int(1 if rfe else 0)

                write_any(df_live, clients_path, sheet="Clients")
                st.success("Modifications enregistrées.")
                st.cache_data.clear()
                st.rerun()

        # -------------------
        # 🗑️ SUPPRIMER
        # -------------------
        elif op == "Supprimer":
            st.markdown("### 🗑️ Supprimer un client")
            if df_live.empty:
                st.info("Aucun client.")
                st.stop()

            ids  = sorted(df_live[COL_ID].dropna().astype(str).unique().tolist())
            noms = sorted(df_live[COL_NOM].dropna().astype(str).unique().tolist())
            s1, s2 = st.columns(2)
            sel_id  = s1.selectbox("ID_Client", [""] + ids, index=0, key=sid("del_id"))
            sel_nom = s2.selectbox("Nom", [""] + noms, index=0, key=sid("del_nom"))

            mask = None
            if sel_id:
                mask = (df_live[COL_ID].astype(str) == sel_id)
            elif sel_nom:
                mask = (df_live[COL_NOM].astype(str) == sel_nom)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({COL_DOSSIER: row.get(COL_DOSSIER,""), COL_ID: row.get(COL_ID,""), COL_NOM: row.get(COL_NOM,""), COL_VISA: row.get(COL_VISA,"")})
                if st.button("❗ Confirmer la suppression", key=sid("del_btn")):
                    df_new = df_live[~mask].copy()
                    write_any(df_new, clients_path, sheet="Clients")
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()


# ================================
# 📄 VISA — aperçu et filtres
# ================================
with tabs[5]:
    st.subheader("📄 Visa (aperçu)")
    if df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        # Filtres simples
        cats = sorted(df_visa_raw["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_visa_raw.columns else []
        subs = sorted(df_visa_raw["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_visa_raw.columns else []

        f1, f2 = st.columns(2)
        fc = f1.multiselect("Catégories", cats, default=[], key=sid("visa_fc"))
        fs = f2.multiselect("Sous-catégories", subs, default=[], key=sid("visa_fs"))

        view = df_visa_raw.copy()
        if fc and "Categorie" in view.columns:
            view = view[view["Categorie"].astype(str).isin(fc)]
        if fs and "Sous-categorie" in view.columns:
            view = view[view["Sous-categorie"].astype(str).isin(fs)]

        st.dataframe(view.reset_index(drop=True), use_container_width=True)


# ================================
# 💾 EXPORT (Clients + Visa)
# ================================
with tabs[6]:
    st.subheader("💾 Export")
    st.caption("Exporte les données Clients et Visa en un seul fichier Excel (2 onglets).")

    if st.button("Préparer l’Excel", key=sid("exp_prep")):
        try:
            export_buf = BytesIO()
            with pd.ExcelWriter(export_buf, engine="openpyxl") as wr:
                # Clients : on relit le fichier pour être sûr d’avoir l'état disque
                dfC = read_any(clients_path, sheet="Clients")
                dfC = _ensure_columns(dfC)
                dfC.to_excel(wr, sheet_name="Clients", index=False)

                # Visa : on exporte la table brute
                if not df_visa_raw.empty:
                    df_visa_raw.to_excel(wr, sheet_name="Visa", index=False)
            st.session_state["export_xlsx"] = export_buf.getvalue()
            st.success("Export prêt.")
        except Exception as e:
            st.error("Erreur à la préparation de l’export : " + _safe_str(e))

    if st.session_state.get("export_xlsx"):
        st.download_button(
            "⬇️ Télécharger Export.xlsx",
            data=st.session_state["export_xlsx"],
            file_name="Export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=sid("exp_dl"),
        )