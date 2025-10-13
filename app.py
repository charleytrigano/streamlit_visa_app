# ==============================================
# 🛂 Visa Manager — PARTIE 1 / 4
#  - Imports & config
#  - Constantes & helpers (dates robustes, formats, num)
#  - IO Excel (Clients & Visa)
#  - Chargement fichiers (2 fichiers / 1 classeur)
#  - Construction visa_map à partir d'une feuille "Visa" (cases=1)
# ==============================================

from __future__ import annotations

import os, io, json, zipfile, re
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple, Any

import pandas as pd
import streamlit as st

# ---------------- Page & thème ----------------
st.set_page_config(page_title="Visa Manager", layout="wide")

# ---------------- Constantes ------------------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

WORK_DIR = os.path.join(os.getcwd(), "data_vm")
os.makedirs(WORK_DIR, exist_ok=True)

DOSSIER_START = 13057  # démarrage du compteur "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"
DOSSIER_COL = "Dossier N"

# ---------------- Helpers génériques ----------
def skey(*parts: Any) -> str:
    """Clé unique et stable pour widgets Streamlit."""
    return "k_" + "_".join(str(p) for p in parts)

def _safe_str(x: Any) -> str:
    try:
        s = "" if x is None else str(x)
    except Exception:
        s = ""
    return s

def _fmt_money_us(x: float | int | str) -> str:
    try:
        v = float(x)
    except Exception:
        v = 0.0
    # format US simple avec séparateur milliers
    return f"${v:,.2f}"

def _safe_num_series(df_or_s: Any, col_or_idx: Any) -> pd.Series:
    """
    Retourne une série numérique sûre:
    - supprime symboles
    - virgule ou point -> point
    """
    if isinstance(df_or_s, pd.Series):
        s = df_or_s
    else:
        s = pd.Series(df_or_s[col_or_idx]) if isinstance(df_or_s, pd.DataFrame) else pd.Series([], dtype=float)
    s = s.astype(str).fillna("")
    s = s.str.replace(r"[^\d,\.\-]", "", regex=True).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _date_for_widget(value, fallback: date | None = None) -> date:
    """
    Retourne toujours un objet date valide pour st.date_input.
    Évite les plantages (None/NaT/texte).
    """
    if fallback is None:
        fallback = date.today()
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    try:
        ts = pd.to_datetime(value, errors="coerce")
        if pd.isna(ts):
            return fallback
        if isinstance(ts, pd.Timestamp):
            return ts.to_pydatetime().date()
    except Exception:
        pass
    return fallback

def _norm(s: str) -> str:
    """Normalise clé (pour clés internes), sans accents spéciaux."""
    s = _safe_str(s).strip().lower()
    # autoriser a-z, 0-9, + / _ - et espace (échappements sécurisés)
    return re.sub(r"[^a-z0-9+\-_/ ]+", " ", s)

# ---------- Persistance des derniers chemins ----------
def _save_last_paths(clients_path: str | None, visa_path: str | None) -> None:
    st.session_state["last_clients_path"] = clients_path
    st.session_state["last_visa_path"] = visa_path

def _load_last_paths() -> Tuple[str | None, str | None]:
    return (
        st.session_state.get("last_clients_path"),
        st.session_state.get("last_visa_path"),
    )

# ---------- Fabrication ID & Dossier ----------
def _make_client_id(nom: str, d: date) -> str:
    base = _safe_str(nom).strip()
    base = re.sub(r"\s+", "_", base)
    return f"{base}-{d:%Y%m%d}"

def _next_dossier(df_clients: pd.DataFrame, start: int = DOSSIER_START) -> int:
    if DOSSIER_COL in df_clients.columns:
        vals = pd.to_numeric(df_clients[DOSSIER_COL], errors="coerce").dropna().astype(int)
        return (vals.max() + 1) if len(vals) else start
    return start

# ---------- IO Excel : Clients ----------
def _read_clients(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        # si pas de feuille nommée, tenter la première
        try:
            xl = pd.ExcelFile(path)
            df = xl.parse(xl.sheet_names[0])
        except Exception:
            return pd.DataFrame()
    # Colonnes minimales
    for c in [DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois", "Categorie", "Sous-categorie", "Visa",
              HONO, AUTRE, TOTAL, "Payé", "Reste", "Paiements", "Options",
              "Dossier envoyé", "Date d'envoi",
              "Dossier accepté", "Date d'acceptation",
              "Dossier refusé", "Date de refus",
              "Dossier annulé", "Date d'annulation",
              "RFE", "Commentaires autres"]:
        if c not in df.columns:
            df[c] = None
    # types sûrs
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = _safe_num_series(df, c)
    return df

def _write_clients(df: pd.DataFrame, path: str) -> None:
    # recalcul Total si besoin
    if HONO in df.columns and AUTRE in df.columns:
        df[TOTAL] = _safe_num_series(df, HONO) + _safe_num_series(df, AUTRE)
    with pd.ExcelWriter(path, engine="openpyxl") as wr:
        df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)

# ---------- IO Excel : Visa (structure cases=1) ----------
def _read_visa_table(path: str) -> pd.DataFrame:
    """
    Lit la feuille Visa. Format attendu:
    - Colonnes obligatoires: 'Categorie', 'Sous-categorie'
    - Colonnes d'options: en-têtes en ligne 1; dans les lignes, 1 => option disponible.
    """
    try:
        df = pd.read_excel(path, sheet_name=SHEET_VISA)
    except Exception:
        try:
            xl = pd.ExcelFile(path)
            # prioriser une feuille nommée Visa; sinon 1ère
            if SHEET_VISA in xl.sheet_names:
                df = xl.parse(SHEET_VISA)
            else:
                df = xl.parse(xl.sheet_names[0])
        except Exception:
            return pd.DataFrame()

    # nettoyer
    # Harmoniser noms de deux colonnes clés
    rename_map = {}
    for col in df.columns:
        cn = _safe_str(col)
        if _norm(cn) in ("categorie",):
            rename_map[col] = "Categorie"
        elif _norm(cn) in ("sous-categorie", "sous_categorie", "sous-categories", "sous categories"):
            rename_map[col] = "Sous-categorie"
    df = df.rename(columns=rename_map)

    # S'assurer présence des colonnes clés
    if "Categorie" not in df.columns or "Sous-categorie" not in df.columns:
        # tenter les 2 premières colonnes
        cols = df.columns.tolist()
        if len(cols) >= 2:
            df = df.rename(columns={cols[0]: "Categorie", cols[1]: "Sous-categorie"})
        else:
            return pd.DataFrame()

    # Remplir NaN par vide
    df["Categorie"] = df["Categorie"].fillna("").astype(str).str.strip()
    df["Sous-categorie"] = df["Sous-categorie"].fillna("").astype(str).str.strip()

    # Retirer lignes vides
    df = df[~((df["Categorie"] == "") & (df["Sous-categorie"] == ""))].copy()
    return df

def build_visa_map(df_visa_raw: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, List[str]]]]:
    """
    Construit une structure:
    { "Categorie": {
         "Sous-categorie": {
             "options": ["COS","EOS", "Autre en-tête", ...]
         }
      }, ... }
    Une option est retenue si la cellule de cette colonne == 1 (ou '1', True).
    """
    if df_visa_raw.empty:
        return {}

    # colonnes d'options = toutes sauf Categorie & Sous-categorie
    opt_cols = [c for c in df_visa_raw.columns if c not in ("Categorie", "Sous-categorie")]
    m: Dict[str, Dict[str, Dict[str, List[str]]]] = {}
    for _, r in df_visa_raw.iterrows():
        cat = _safe_str(r.get("Categorie","")).strip()
        sub = _safe_str(r.get("Sous-categorie","")).strip()
        if not cat or not sub:
            continue
        if cat not in m:
            m[cat] = {}
        if sub not in m[cat]:
            m[cat][sub] = {"options": []}
        # options actives
        opts = []
        for oc in opt_cols:
            val = r.get(oc, None)
            active = False
            if isinstance(val, (int, float)) and float(val) == 1.0:
                active = True
            elif _safe_str(val).strip() == "1":
                active = True
            elif isinstance(val, bool) and val:
                active = True
            if active:
                opts.append(_safe_str(oc).strip())
        m[cat][sub]["options"] = sorted(list(dict.fromkeys(opts)))
    return m

# ==============================================
# 🎛️ CHARGEMENT FICHIERS (UI) + MÉMO des derniers chemins
# ==============================================
st.markdown("# 🛂 Visa Manager")
st.markdown("## 📂 Fichiers")

mode = st.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"],
                horizontal=True, key=skey("load","mode"))

clients_path: str | None = None
visa_path: str | None = None

if mode == "Deux fichiers (Clients & Visa)":
    up_c = st.file_uploader("Clients (xlsx)", type=["xlsx"], key=skey("load","clients"))
    up_v = st.file_uploader("Visa (xlsx)", type=["xlsx"], key=skey("load","visa"))
    if up_c:
        clients_path = os.path.join(WORK_DIR, "clients_loaded.xlsx")
        with open(clients_path, "wb") as f:
            f.write(up_c.getvalue())
    if up_v:
        visa_path = os.path.join(WORK_DIR, "visa_loaded.xlsx")
        with open(visa_path, "wb") as f:
            f.write(up_v.getvalue())

else:
    up_one = st.file_uploader("Classeur unique (2 onglets: Clients & Visa)", type=["xlsx"], key=skey("load","one"))
    if up_one:
        one_path = os.path.join(WORK_DIR, "workbook_loaded.xlsx")
        with open(one_path, "wb") as f:
            f.write(up_one.getvalue())
        # décomposer en deux chemins "logiques" (on gardera le même fichier sous-jacent)
        clients_path = one_path
        visa_path    = one_path

# Récupérer derniers chemins si rien de chargé maintenant
if not clients_path or not visa_path:
    last_c, last_v = _load_last_paths()
    if not clients_path and last_c and os.path.exists(last_c):
        clients_path = last_c
    if not visa_path and last_v and os.path.exists(last_v):
        visa_path = last_v

# Si toujours rien, proposer un exemple minimal en mémoire
if not clients_path or not os.path.exists(clients_path):
    # créer un fichier Clients vide conforme
    clients_path = os.path.join(WORK_DIR, "modele_clients.xlsx")
    if not os.path.exists(clients_path):
        df_empty = pd.DataFrame(columns=[
            DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
            "Categorie", "Sous-categorie", "Visa",
            HONO, AUTRE, TOTAL, "Commentaires autres",
            "Payé", "Reste", "Paiements", "Options",
            "Dossier envoyé", "Date d'envoi",
            "Dossier accepté", "Date d'acceptation",
            "Dossier refusé", "Date de refus",
            "Dossier annulé", "Date d'annulation",
            "RFE",
        ])
        _write_clients(df_empty, clients_path)

if not visa_path or not os.path.exists(visa_path):
    # créer un fichier Visa modèle simple si manquant
    visa_path = os.path.join(WORK_DIR, "modele_visa.xlsx")
    if not os.path.exists(visa_path):
        dfv = pd.DataFrame({
            "Categorie": ["Affaires/Tourisme","Affaires/Tourisme","Etudiants","Etudiants"],
            "Sous-categorie": ["B-1","B-2","F-1","F-2"],
            "COS": [1,1,1,1],
            "EOS": [1,1,1,1],
        })
        with pd.ExcelWriter(visa_path, engine="openpyxl") as wr:
            dfv.to_excel(wr, sheet_name=SHEET_VISA, index=False)

# mémoriser
_save_last_paths(clients_path, visa_path)

# Charger données
df_clients = _read_clients(clients_path)
df_visa_raw = _read_visa_table(visa_path)

# Construire la carte Visa
visa_map = build_visa_map(df_visa_raw.copy()) if not df_visa_raw.empty else {}

# Dataframe combiné pour Dashboard/Analyses
df_all = df_clients.copy()
# Enrichissements date -> année/mois numériques pour tri/graph
if not df_all.empty:
    # Date en string -> datetime
    try:
        dt_series = pd.to_datetime(df_all["Date"], errors="coerce")
    except Exception:
        dt_series = pd.Series([pd.NaT]*len(df_all))
    df_all["_Année_"]  = dt_series.dt.year.fillna(0).astype(int)
    # mois comme "MM" déjà en colonne "Mois", on prépare aussi un mois numérique pour tri
    df_all["_MoisNum_"] = pd.to_numeric(df_all.get("Mois", ""), errors="coerce").fillna(0).astype(int)




# ==============================================
# 🧭 PARTIE 2 / 4 — Tabs + Dashboard + Aperçu Visa
#  - Création des onglets
#  - DASHBOARD : filtres, KPI, graphes, liste
#  - VISA (aperçu) : cascade Catégorie → Sous-catégorie → options (cases)
# ==============================================

# ---------- Création des onglets ----------
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "🧾 Gestion", "📄 Visa (aperçu)"])

# ===========
# 📊 DASHBOARD
# ===========
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucune donnée client chargée.")
    else:
        # --- Filtres ---
        years  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1, 13)]
        cats   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        c1, c2, c3, c4, c5 = st.columns([1,1,1,1,1])
        fy = c1.multiselect("Année", years, default=[], key=skey("dash","years"))
        fm = c2.multiselect("Mois (MM)", months, default=[], key=skey("dash","months"))
        fc = c3.multiselect("Catégorie", cats, default=[], key=skey("dash","cats"))
        fs = c4.multiselect("Sous-catégorie", subs, default=[], key=skey("dash","subs"))
        fv = c5.multiselect("Visa", visas, default=[], key=skey("dash","visas"))

        ff = df_all.copy()
        if fy: ff = ff[ff["_Année_"].isin(fy)]
        if fm: ff = ff[ff["Mois"].astype(str).isin(fm)]
        if fc: ff = ff[ff["Categorie"].astype(str).isin(fc)]
        if fs: ff = ff[ff["Sous-categorie"].astype(str).isin(fs)]
        if fv: ff = ff[ff["Visa"].astype(str).isin(fv)]

        # --- KPI compacts ---
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(ff)}")
        k2.metric("Honoraires", _fmt_money_us(_safe_num_series(ff, HONO).sum()))
        k3.metric("Autres frais", _fmt_money_us(_safe_num_series(ff, AUTRE).sum()))
        k4.metric("Payé", _fmt_money_us(_safe_num_series(ff, "Payé").sum()))
        k5.metric("Reste", _fmt_money_us(_safe_num_series(ff, "Reste").sum()))

        st.caption("💡 Astuce : utilisez les filtres ci-dessus pour segmenter instantanément le tableau et les graphes.")

        # --- Graphes simples ---
        gcol1, gcol2 = st.columns(2)

        with gcol1:
            st.markdown("#### Dossiers par catégorie")
            if not ff.empty and "Categorie" in ff.columns:
                vc = ff["Categorie"].value_counts().reset_index()
                vc.columns = ["Categorie", "Nombre"]
                st.bar_chart(vc.set_index("Categorie"))
            else:
                st.info("Pas de catégories à afficher.")

        with gcol2:
            st.markdown("#### Honoraires par mois")
            if not ff.empty and "Mois" in ff.columns:
                tmp = ff.copy()
                tmp["Mois"] = tmp["Mois"].astype(str)
                gm = tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois")
                st.line_chart(gm.set_index("Mois"))
            else:
                st.info("Pas de données mensuelles.")

        # --- Détails des dossiers filtrés ---
        st.markdown("#### 📋 Détails")
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
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa", "Date", "Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"
        ] if c in view.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

        st.dataframe(
            view_sorted[show_cols].reset_index(drop=True),
            use_container_width=True,
            key=skey("dash","detail")
        )

# ==================
# 📄 VISA (APERÇU)
# ==================
with tabs[5]:
    st.subheader("📄 Visa — aperçu & structure")

    if not visa_map:
        st.info("Aucune structure Visa chargée (feuille 'Visa').")
    else:
        # Sélection cascade
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=skey("visa","cat"))
        sel_sub = ""
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=skey("visa","sub"))

        if sel_cat and sel_sub:
            opts = visa_map.get(sel_cat, {}).get(sel_sub, {}).get("options", [])
            st.markdown("##### Options disponibles")
            if not opts:
                st.info("Aucune option pour cette sous-catégorie.")
            else:
                # Affichage des options (lecture seule)
                cols = st.columns(4)
                for i, opt in enumerate(opts):
                    cols[i % 4].checkbox(opt, value=True, key=skey("visa","opt",sel_cat,sel_sub,opt), disabled=True)

        # Rappel de la structure brute
        st.markdown("---")
        st.markdown("#### Structure brute (extrait)")
        try:
            preview_cols = ["Categorie", "Sous-categorie"] + [
                c for c in df_visa_raw.columns if c not in ("Categorie", "Sous-categorie")
            ][:8]  # limiter l'aperçu
            st.dataframe(df_visa_raw[preview_cols].head(50), use_container_width=True, height=300, key=skey("visa","raw"))
        except Exception:
            st.dataframe(df_visa_raw.head(50), use_container_width=True, height=300, key=skey("visa","raw2"))




# ==============================================
# 📈 PARTIE 3 / 4 — Analyses & Escrow
#  - Analyses : filtres, KPI compacts, % par catégorie/sous-cat,
#               comparaisons période A vs B (année/mois),
#               graphes volumes & montants
#  - Escrow : synthèse, KPI compacts, par catégorie & alertes basiques
# ==============================================

# =================
# 📈 ONGLET ANALYSES
# =================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        # --- Filtres globaux pour analyses ---
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3, a4, a5 = st.columns([1,1,1,1,1])
        fy = a1.multiselect("Année", yearsA, default=[], key=skey("ana","y"))
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=skey("ana","m"))
        fc = a3.multiselect("Catégorie", catsA, default=[], key=skey("ana","c"))
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=skey("ana","s"))
        fv = a5.multiselect("Visa", visasA, default=[], key=skey("ana","v"))

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # --- KPI compacts ---
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money_us(_safe_num_series(dfA, HONO).sum()))
        k3.metric("Autres frais", _fmt_money_us(_safe_num_series(dfA, AUTRE).sum()))
        k4.metric("Payé", _fmt_money_us(_safe_num_series(dfA, "Payé").sum()))
        k5.metric("Reste", _fmt_money_us(_safe_num_series(dfA, "Reste").sum()))

        st.markdown("---")

        # --- % par Catégorie / Sous-catégorie (volumes) ---
        g1, g2 = st.columns(2)

        with g1:
            st.markdown("#### % dossiers par catégorie")
            if not dfA.empty and "Categorie" in dfA.columns:
                vc = dfA["Categorie"].value_counts(dropna=True).reset_index()
                vc.columns = ["Categorie", "Nombre"]
                total = vc["Nombre"].sum()
                if total > 0:
                    vc["%"] = (vc["Nombre"] / total * 100).round(1)
                st.dataframe(vc, use_container_width=True, height=240, key=skey("ana","pcat"))
            else:
                st.info("Pas de catégories.")

        with g2:
            st.markdown("#### % dossiers par sous-catégorie")
            if not dfA.empty and "Sous-categorie" in dfA.columns:
                vs = dfA["Sous-categorie"].value_counts(dropna=True).reset_index()
                vs.columns = ["Sous-categorie", "Nombre"]
                total = vs["Nombre"].sum()
                if total > 0:
                    vs["%"] = (vs["Nombre"] / total * 100).round(1)
                st.dataframe(vs, use_container_width=True, height=240, key=skey("ana","psub"))
            else:
                st.info("Pas de sous-catégories.")

        st.markdown("---")

        # --- Graphes volumes & montants par mois ---
        gg1, gg2 = st.columns(2)

        with gg1:
            st.markdown("#### Volumes : dossiers par mois")
            if not dfA.empty:
                tmp = dfA.copy()
                tmp["Mois"] = tmp["Mois"].astype(str)
                vc_m = tmp["Mois"].value_counts().rename_axis("Mois").reset_index(name="Dossiers").sort_values("Mois")
                if not vc_m.empty:
                    st.bar_chart(vc_m.set_index("Mois"))
                else:
                    st.info("Aucun volume mensuel.")
            else:
                st.info("Pas de données.")

        with gg2:
            st.markdown("#### Montants : honoraires par mois")
            if not dfA.empty:
                tmp = dfA.copy()
                tmp["Mois"] = tmp["Mois"].astype(str)
                gm = tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois")
                if not gm.empty:
                    st.line_chart(gm.set_index("Mois"))
                else:
                    st.info("Aucun montant mensuel.")
            else:
                st.info("Pas de données.")

        st.markdown("---")

        # --- Comparaison Période A vs B (Année & Mois) ---
        st.markdown("### Comparaison période A vs B")

        ca1, ca2, cb1, cb2 = st.columns(4)
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=skey("cmp","ya"))
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=skey("cmp","ma"))
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=skey("cmp","yb"))
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=skey("cmp","mb"))

        def _apply_period_filter(df: pd.DataFrame, ys: List[int], ms: List[str]) -> pd.DataFrame:
            d = df.copy()
            if ys: d = d[d["_Année_"].isin(ys)]
            if ms: d = d[d["Mois"].astype(str).isin(ms)]
            return d

        dA = _apply_period_filter(df_all, pa_years, pa_month)
        dB = _apply_period_filter(df_all, pb_years, pb_month)

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### Période A — KPI")
            a1, a2, a3 = st.columns(3)
            a1.metric("Dossiers", f"{len(dA)}")
            a2.metric("Honoraires", _fmt_money_us(_safe_num_series(dA, HONO).sum()))
            a3.metric("Payé", _fmt_money_us(_safe_num_series(dA, "Payé").sum()))
        with c2:
            st.markdown("##### Période B — KPI")
            b1, b2, b3 = st.columns(3)
            b1.metric("Dossiers", f"{len(dB)}")
            b2.metric("Honoraires", _fmt_money_us(_safe_num_series(dB, HONO).sum()))
            b3.metric("Payé", _fmt_money_us(_safe_num_series(dB, "Payé").sum()))

        # Détail comparatif par catégorie (honoraires)
        st.markdown("#### Comparaison par catégorie (Honoraires)")
        def _agg_cat(df: pd.DataFrame) -> pd.DataFrame:
            if df.empty or "Categorie" not in df.columns:
                return pd.DataFrame(columns=["Categorie", "Honoraires"])
            g = df.groupby("Categorie", as_index=False)[HONO].sum()
            g = g.rename(columns={HONO: "Honoraires"})
            return g

        agA = _agg_cat(dA)
        agB = _agg_cat(dB)
        comp = pd.merge(agA, agB, on="Categorie", how="outer", suffixes=(" A", " B")).fillna(0.0)
        if not comp.empty:
            comp["Δ"] = comp["Honoraires A"] - comp["Honoraires B"]
            st.dataframe(comp.sort_values("Categorie"), use_container_width=True, key=skey("cmp","table"))
        else:
            st.info("Comparaison indisponible (pas de données).")

        st.markdown("---")

        # --- Détails des dossiers filtrés (pour analyses) ---
        st.markdown("### 🧾 Détails des dossiers filtrés (Analyses)")
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
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa", "Date", "Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det

        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=skey("ana","detail"))

# ==============
# 🏦 ONGLET ESCROW
# ==============
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")

    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Payé"]  = _safe_num_series(dfE, "Payé")
        dfE["Reste"] = _safe_num_series(dfE, "Reste")
        dfE[TOTAL]   = _safe_num_series(dfE, TOTAL)

        # KPI compacts
        e1, e2, e3 = st.columns([1,1,1])
        e1.metric("Total (US $)", _fmt_money_us(float(dfE[TOTAL].sum())))
        e2.metric("Payé", _fmt_money_us(float(dfE["Payé"].sum())))
        e3.metric("Reste", _fmt_money_us(float(dfE["Reste"].sum())))

        st.markdown("#### Par catégorie")
        agg = dfE.groupby("Categorie", as_index=False)[[TOTAL, "Payé", "Reste"]].sum()
        if not agg.empty:
            agg["% Payé"] = (agg["Payé"] / agg[TOTAL]).replace([pd.NA, pd.NaT], 0).fillna(0.0) * 100
            st.dataframe(agg.sort_values("Categorie"), use_container_width=True, key=skey("esc","agg"))
        else:
            st.info("Pas de regroupement possible.")

        st.markdown("---")
        st.markdown("#### Alerte : dossiers envoyés mais non soldés")
        # Un dossier "envoyé" avec reste > 0 => alerte (escrow à réclamer / transfert)
        mask_alert = (dfE.get("Dossier envoyé", 0).fillna(0).astype(int) == 1) & (dfE["Reste"] > 0.0)
        alert_df = dfE[mask_alert].copy()
        if not alert_df.empty:
            show_cols = [c for c in [
                DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
                HONO, AUTRE, TOTAL, "Payé", "Reste", "Date d'envoi"
            ] if c in alert_df.columns]
            # afficher les montants formatés
            for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                if c in alert_df.columns:
                    alert_df[c] = alert_df[c].map(_fmt_money_us)
            st.dataframe(alert_df[show_cols].reset_index(drop=True), use_container_width=True, key=skey("esc","alert"))
            st.caption("💡 Ces dossiers sont « envoyés » mais présentent encore un reste à encaisser/transférer.")
        else:
            st.success("Aucune alerte : tous les dossiers envoyés sont soldés.")



# =====================================================
# 🧾 ONGLET : GESTION (CRUD) — dans l’onglet Clients
# =====================================================
with tabs[4]:
    st.markdown("---")
    st.subheader("🧾 Gestion des clients (Ajouter / Modifier / Supprimer)")

    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_op_{SID}")

    df_live = _read_clients(clients_path)

    # ---------- AJOUT ----------
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")
        c1, c2, c3 = st.columns(3)
        nom  = c1.text_input("Nom", "", key=f"add_nom_{SID}")
        dt   = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=int(date.today().month)-1, key=f"add_mois_{SID}")

        # Sélection visa
        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=f"add_cat_{SID}")
        sel_sub = ""
        visa_final = ""
        opts_dict = {"exclusive": None, "options": []}
        info_msg = ""
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=f"add_sub_{SID}")
            if sel_sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=f"add_opts_{SID}", preselected={}
                )
        if info_msg:
            st.info(info_msg)

        f1, f2 = st.columns(2)
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
        other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")
        commentaire = st.text_area("Commentaires (autres frais / remarques)", "", key=f"add_com_{SID}")

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
            st.warning("⚠️ La case RFE ne peut être cochée qu’avec un autre statut (envoyé, accepté, refusé ou annulé).")

        if st.button("💾 Enregistrer le client", key=f"btn_add_{SID}"):
            if not nom:
                st.warning("Veuillez saisir le nom.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Veuillez choisir la catégorie et la sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            dossier_n = _next_dossier(df_live, start=13057)
            did = _make_client_id(nom, dt)
            new_row = {
                "Dossier N": dossier_n,
                "ID_Client": did,
                "Nom": nom,
                "Date": dt,
                "Mois": mois,
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": visa_final if visa_final else sel_sub,
                "Montant honoraires (US $)": honor,
                "Autres frais (US $)": other,
                "Commentaires": commentaire,
                "Total (US $)": total,
                "Payé": 0.0,
                "Reste": total,
                "Options": opts_dict,
                "Dossier envoyé": 1 if sent else 0,
                "Date d'envoi": sent_d,
                "Dossier accepté": 1 if acc else 0,
                "Date d'acceptation": acc_d,
                "Dossier refusé": 1 if ref else 0,
                "Date de refus": ref_d,
                "Dossier annulé": 1 if ann else 0,
                "Date d'annulation": ann_d,
                "RFE": 1 if rfe else 0
            }
            df_live = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            _write_clients(df_live, clients_path)
            st.success("Client ajouté.")
            st.cache_data.clear()
            st.rerun()


    # ---------- MODIFIER ----------
    elif op == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if df_live.empty:
            st.info("Aucun client à modifier.")
        else:
            names = df_live["Nom"].dropna().unique().tolist()
            target = st.selectbox("Sélectionner un client", [""]+names, key=f"mod_target_{SID}")
            if target:
                row = df_live[df_live["Nom"] == target].iloc[0].copy()

                def _to_int_bool(v):
                    try:
                        if isinstance(v, bool): return v
                        if isinstance(v, (int, float)) and not pd.isna(v): return int(v) == 1
                        if isinstance(v, str): return v.strip().lower() in ["1","true","oui","yes","x"]
                    except Exception:
                        pass
                    return False

                c1, c2, c3 = st.columns(3)
                nom  = c1.text_input("Nom", row.get("Nom",""), key=f"mod_nom_{SID}")
                dt   = c2.date_input("Date de création", value=_date_for_widget(row.get("Date")), key=f"mod_date_{SID}")
                mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                    index=int(str(row.get("Mois","01")))-1, key=f"mod_mois_{SID}")

                st.markdown("#### 🎯 Choix Visa")
                cats = sorted(list(visa_map.keys()))
                preset_cat = _safe_str(row.get("Categorie",""))
                sel_cat = st.selectbox("Catégorie", [""]+cats,
                    index=(cats.index(preset_cat)+1 if preset_cat in cats else 0), key=f"mod_cat_{SID}")
                subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
                preset_sub = _safe_str(row.get("Sous-categorie",""))
                sel_sub = st.selectbox("Sous-catégorie", [""]+subs,
                    index=(subs.index(preset_sub)+1 if preset_sub in subs else 0), key=f"mod_sub_{SID}")

                f1, f2 = st.columns(2)
                honor = f1.number_input("Montant honoraires (US $)", min_value=0.0,
                        value=float(row.get("Montant honoraires (US $)",0)), step=50.0, format="%.2f", key=f"mod_h_{SID}")
                other = f2.number_input("Autres frais (US $)", min_value=0.0,
                        value=float(row.get("Autres frais (US $)",0)), step=20.0, format="%.2f", key=f"mod_o_{SID}")
                commentaire = st.text_area("Commentaires", _safe_str(row.get("Commentaires","")), key=f"mod_com_{SID}")

                st.markdown("#### 📌 Statuts")
                s1, s2, s3, s4, s5 = st.columns(5)
                envoye = _to_int_bool(row.get("Dossier envoyé"))
                accepte = _to_int_bool(row.get("Dossier accepté"))
                refuse = _to_int_bool(row.get("Dossier refusé"))
                annule = _to_int_bool(row.get("Dossier annulé"))
                rfe = _to_int_bool(row.get("RFE"))
                sent   = s1.checkbox("Dossier envoyé", value=envoye, key=f"mod_sent_{SID}")
                sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
                acc    = s2.checkbox("Dossier accepté", value=accepte, key=f"mod_acc_{SID}")
                acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
                ref    = s3.checkbox("Dossier refusé", value=refuse, key=f"mod_ref_{SID}")
                ref_d  = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=f"mod_refd_{SID}")
                ann    = s4.checkbox("Dossier annulé", value=annule, key=f"mod_ann_{SID}")
                ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
                rfe    = s5.checkbox("RFE", value=rfe, key=f"mod_rfe_{SID}")

                if st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}"):
                    idx = df_live[df_live["Nom"] == target].index[0]
                    df_live.at[idx, "Nom"] = nom
                    df_live.at[idx, "Date"] = dt
                    df_live.at[idx, "Mois"] = mois
                    df_live.at[idx, "Categorie"] = sel_cat
                    df_live.at[idx, "Sous-categorie"] = sel_sub
                    df_live.at[idx, "Montant honoraires (US $)"] = honor
                    df_live.at[idx, "Autres frais (US $)"] = other
                    df_live.at[idx, "Commentaires"] = commentaire
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
                    st.success("Client modifié.")
                    st.cache_data.clear()
                    st.rerun()

    # ---------- SUPPRIMER ----------
    elif op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client à supprimer.")
        else:
            noms = df_live["Nom"].dropna().unique().tolist()
            target = st.selectbox("Sélectionner un client à supprimer", [""]+noms, key=f"del_target_{SID}")
            if target and st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                df_live = df_live[df_live["Nom"] != target]
                _write_clients(df_live, clients_path)
                st.success("Client supprimé.")
                st.cache_data.clear()
                st.rerun()