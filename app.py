# ================================
# PARTIE 1/6 — Imports & Utilitaires
# ================================
from __future__ import annotations

import json, re, os, zipfile, uuid
from io import BytesIO
from datetime import datetime, date
from typing import Dict, List, Tuple, Any

import pandas as pd
import streamlit as st

# ----------------
# CLÉS & CONSTANTES
# ----------------
APP_TITLE = "🛂 Visa Manager"
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

# Mémorisation du dernier fichier (dans le répertoire courant)
LAST_MEMO_FILE = "last_paths.json"

def _load_last_paths() -> Tuple[str|None, str|None, str|None]:
    """Charge les derniers chemins (mode, clients_path, visa_path)."""
    try:
        with open(LAST_MEMO_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d.get("mode"), d.get("clients_path"), d.get("visa_path")
    except Exception:
        return None, None, None

def _save_last_paths(mode: str|None, clients_path: str|None, visa_path: str|None) -> None:
    try:
        with open(LAST_MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump({"mode": mode, "clients_path": clients_path, "visa_path": visa_path}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ----------------
# GÉNÉRATEURS DE CLÉS & FORMATAGE
# ----------------
_SID_PREFIX = st.session_state.get("_sid_prefix") or str(uuid.uuid4())[:8]
st.session_state["_sid_prefix"] = _SID_PREFIX
def sid(*parts) -> str:
    return _SID_PREFIX + "_" + "_".join(str(p) for p in parts)

def _fmt_money(x) -> str:
    try:
        v = float(x)
    except Exception:
        v = 0.0
    return f"${v:,.2f}"

def _safe_str(x: Any) -> str:
    try:
        if x is None:
            return ""
        return str(x)
    except Exception:
        return ""

def _to_date(val) -> date|None:
    """Coerce en date (ou None)."""
    if isinstance(val, date):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        return d.date() if pd.notna(d) else None
    except Exception:
        return None

def _nnum(s: pd.Series|Any, default: float = 0.0) -> pd.Series|float:
    """Numérise de manière robuste."""
    if isinstance(s, pd.Series):
        return pd.to_numeric(s, errors="coerce").fillna(default)
    try:
        return float(s)
    except Exception:
        return default

# ----------------
# LECTURE/ÉCRITURE FICHIERS
# ----------------
@st.cache_data(show_spinner=False)
def read_any_table(path_or_buf, sheet: str|None=None) -> pd.DataFrame:
    """Lit un CSV/XLSX. Si sheet est fourni, lit cet onglet."""
    if hasattr(path_or_buf, "read"):  # uploaded file
        try:
            if sheet:
                return pd.read_excel(path_or_buf, sheet_name=sheet)
            # Si un seul tableau dans le fichier
            try:
                return pd.read_excel(path_or_buf)
            except Exception:
                path_or_buf.seek(0)
                return pd.read_csv(path_or_buf)
        except Exception:
            path_or_buf.seek(0)
            return pd.read_csv(path_or_buf)
    # chemin disque
    p = str(path_or_buf)
    ext = os.path.splitext(p)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(p, sheet_name=sheet) if sheet else pd.read_excel(p)
    return pd.read_csv(p)

@st.cache_data(show_spinner=False)
def read_clients_file(path_or_buf) -> pd.DataFrame:
    df = read_any_table(path_or_buf)
    return df

@st.cache_data(show_spinner=False)
def read_visa_file(path_or_buf) -> pd.DataFrame:
    df = read_any_table(path_or_buf)
    return df

def write_clients_file(df: pd.DataFrame, path: str) -> None:
    """Écrit le fichier Clients exactement à l’endroit choisi par l’utilisateur."""
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        with pd.ExcelWriter(path, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
    else:
        df.to_csv(path, index=False, encoding="utf-8-sig")

# ----------------
# NORMALISATION COLONNES CLIENTS
# ----------------
RENAMES = {
    "Categories": "Categorie",
    "Sous-catégorie": "Sous-categorie",   # sans accents dans l’app
    "Sous-categorie": "Sous-categorie",
    "Dossiers envoyé": "Dossier envoyé",
    "Dossier Envoyé": "Dossier envoyé",
    "Dossier approuve": "Dossier approuvé",
    "Dossier Approuvé": "Dossier approuvé",
    "Dossier refuse": "Dossier refusé",
    "Dossier Refusé": "Dossier refusé",
    "Dossier Annulé": "Dossier annulé",
    "Solde": "Reste",
}

REQ_CLIENT_COLS = [
    "ID_Client","Dossier N","Nom","Date","Mois",
    "Categorie","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Payé","Reste",
    "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé",
    "Commentaires"
]

def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=REQ_CLIENT_COLS + ["_Année_","_MoisNum_"])
    # renames
    cols = {c: RENAMES.get(c, c) for c in df.columns}
    df = df.rename(columns=cols)
    # create missing
    for c in REQ_CLIENT_COLS:
        if c not in df.columns:
            df[c] = pd.NA
    # Date → Année/Mois
    d = pd.to_datetime(df["Date"], errors="coerce")
    df["_Année_"]   = d.dt.year.fillna(0).astype(int)
    df["_MoisNum_"] = d.dt.month.fillna(0).astype(int)
    # Mois (MM)
    if "Mois" in df.columns:
        # si déjà string MM, on garde — sinon on recalcule
        m = df["Mois"].astype(str).str.zfill(2)
        df["Mois"] = m.where(m.str.match(r"^\d{2}$"), df["_MoisNum_"].apply(lambda x: f"{int(x):02d}" if x>0 else ""))
    else:
        df["Mois"] = df["_MoisNum_"].apply(lambda x: f"{int(x):02d}" if x>0 else "")
    # Numériques
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Reste"]:
        df[c] = _nnum(df[c])
    # Statuts binaires
    for c in ["RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    # Listes/json
    if "Paiements" not in df.columns:
        df["Paiements"] = [[] for _ in range(len(df))]
    # Options (pour enregistrement des choix visa)
    if "Options" not in df.columns:
        df["Options"] = [{} for _ in range(len(df))]
    return df

# ----------------
# VISA MAP (structure à partir du fichier Visa)
# ----------------
def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    Construit :
    {
      "Affaires/Tourisme": {
          "B-1": {"options": ["COS","EOS"], "exclusive": True},
          "B-2": {"options": ["COS","EOS"], "exclusive": True},
      },
      "Etudiants": {
          "F-1": {"options": ["COS","EOS"], "exclusive": True},
          ...
      }
    }
    Règle : pour chaque ligne, toutes les colonnes (hors Catégorie & Sous-categorie)
    dont la valeur == 1 sont des cases à cocher disponibles.
    """
    if df_visa is None or df_visa.empty:
        return {}
    # normalise nom des 2 colonnes pivot
    rcols = {c: c for c in df_visa.columns}
    for c in ["Categorie","Catégorie"]:
        if c in df_visa.columns:
            rcols[c] = "Categorie"
    for c in ["Sous-categorie","Sous-catégorie"]:
        if c in df_visa.columns:
            rcols[c] = "Sous-categorie"
    df = df_visa.rename(columns=rcols).copy()

    # colonnes d’options = toutes sauf Categorie/Sous-categorie
    option_cols = [c for c in df.columns if c not in ["Categorie","Sous-categorie"]]

    visa_map: Dict[str, Dict[str, Dict[str, Any]]] = {}
    for _, row in df.iterrows():
        cat = _safe_str(row.get("Categorie","")).strip()
        sub = _safe_str(row.get("Sous-categorie","")).strip()
        if not cat or not sub:
            continue
        opts = []
        for c in option_cols:
            val = row.get(c)
            try:
                v = float(val)
            except Exception:
                v = None
            if v == 1:   # coche disponible
                opts.append(c)
        if cat not in visa_map:
            visa_map[cat] = {}
        visa_map[cat][sub] = {"options": opts, "exclusive": True}  # par défaut exclusif (COS/EOS)
    return visa_map


# ===========================================
# PARTIE 2/6 — Barre latérale & Chargement
# ===========================================
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title(APP_TITLE)

# --- Barre latérale — chargement fichiers
st.sidebar.header("📂 Fichiers")
mode_choice = st.sidebar.radio(
    "Mode de chargement",
    ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"],
    key=sid("mode")
)

last_mode, last_cli, last_visa = _load_last_paths()
clients_file = None
visa_file    = None

if mode_choice == "Deux fichiers (Clients & Visa)":
    clients_file = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx","xls","csv"], key=sid("up_clients"))
    visa_file    = st.sidebar.file_uploader("Visa (xlsx/csv)",     type=["xlsx","xls","csv"], key=sid("up_visa"))
else:
    uni_file = st.sidebar.file_uploader("Fichier unique (2 onglets 'Clients' et 'Visa')", type=["xlsx","xls"], key=sid("up_uni"))
    if uni_file:
        try:
            df_clients_raw = read_any_table(uni_file, sheet=SHEET_CLIENTS)
            df_visa_raw    = read_any_table(uni_file, sheet=SHEET_VISA)
            st.session_state["df_clients_raw"] = df_clients_raw
            st.session_state["df_visa_raw"]    = df_visa_raw
            _save_last_paths("uni", "./upload_"+SHEET_CLIENTS+".xlsx", "./upload_"+SHEET_VISA+".xlsx")
        except Exception as e:
            st.sidebar.error(f"Echec de lecture des onglets: {e}")

# --- Si fichiers séparés, charge & mémorise
if mode_choice == "Deux fichiers (Clients & Visa)":
    if clients_file is not None:
        st.session_state["df_clients_raw"] = read_clients_file(clients_file)
        _save_last_paths("split", "./upload_clients.xlsx", (last_visa or "./upload_visa.xlsx"))
    elif last_mode == "split" and last_cli and os.path.exists(last_cli):
        st.session_state["df_clients_raw"] = read_clients_file(last_cli)
    if visa_file is not None:
        st.session_state["df_visa_raw"] = read_visa_file(visa_file)
        _save_last_paths("split", (last_cli or "./upload_clients.xlsx"), "./upload_visa.xlsx")
    elif last_mode == "split" and last_visa and os.path.exists(last_visa):
        st.session_state["df_visa_raw"] = read_visa_file(last_visa)

# --- Récupère en mémoire
df_clients_raw: pd.DataFrame = st.session_state.get("df_clients_raw", pd.DataFrame())
df_visa_raw:    pd.DataFrame = st.session_state.get("df_visa_raw",    pd.DataFrame())

# Normalise
df_all = normalize_clients(df_clients_raw.copy())
visa_map = build_visa_map(df_visa_raw.copy())

# Affiche info fichiers
with st.expander("📄 Fichiers chargés", expanded=True):
    st.write("**Clients** :", "chargé" if not df_all.empty else "—")
    st.write("**Visa** :", "chargé" if not df_visa_raw.empty else "—")
    if not df_visa_raw.empty:
        st.caption(f"Catégories Visa : {', '.join(sorted(visa_map.keys()))}")

# Création onglets principaux
tabs = st.tabs([
    "📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Compte client",
    "🧾 Gestion", "📄 Visa (aperçu)", "💾 Export"
])


# ========================
# PARTIE 3/6 — Dashboard
# ========================
with tabs[0]:
    st.subheader("📊 Dashboard")

    # ---- Helpers locaux pour éviter les collisions de clés & options vides
    _SID = st.session_state.get("_sid_prefix", "sid")
    def _k(*parts):  # clé unique et stable
        return _SID + "_dash_" + "_".join(str(p) for p in parts)

    def _clean_list(series_like):
        try:
            s = pd.Series(series_like).dropna().astype(str).str.strip()
            s = s[s != ""]
            return sorted(s.unique().tolist())
        except Exception:
            return []

    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # Listes filtres (toujours calculées sur df_all, pas sur un sous-ensemble)
        years  = sorted([int(x) for x in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1,13)]
        cats   = _clean_list(df_all.get("Categorie", []))
        subs   = _clean_list(df_all.get("Sous-categorie", []))
        visas  = _clean_list(df_all.get("Visa", []))

        # Widgets — toujours avec des clés uniques
        a1, a2, a3, a4, a5 = st.columns([1,1,1,1,2])
        fy = a1.multiselect("Année", years, default=[], key=_k("y"))
        fm = a2.multiselect("Mois (MM)", months, default=[], key=_k("m"))
        fc = a3.multiselect("Catégories", cats, default=[], key=_k("c"))
        fs = a4.multiselect("Sous-catégories", subs, default=[], key=_k("s"))
        fv = a5.multiselect("Visa", visas, default=[], key=_k("v"))

        # Application des filtres
        f = df_all.copy()
        if fy: f = f[f["_Année_"].isin(fy)]
        if fm: f = f[f["Mois"].astype(str).isin(fm)]
        if fc: f = f[f["Categorie"].astype(str).isin(fc)]
        if fs: f = f[f["Sous-categorie"].astype(str).isin(fs)]
        if fv: f = f[f["Visa"].astype(str).isin(fv)]

        # KPIs (taille compacte)
        k1,k2,k3,k4,k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(f)}")
        total = (pd.to_numeric(f.get("Montant honoraires (US $)", 0), errors="coerce").fillna(0) +
                 pd.to_numeric(f.get("Autres frais (US $)", 0), errors="coerce").fillna(0)).sum()
        paye  = pd.to_numeric(f.get("Payé", 0), errors="coerce").fillna(0).sum()
        reste = pd.to_numeric(f.get("Reste", 0), errors="coerce").fillna(0).sum()
        k2.metric("Honoraires+Frais", _fmt_money(total))
        k3.metric("Payé", _fmt_money(paye))
        k4.metric("Solde", _fmt_money(reste))
        pct_env = 0
        if "Dossier envoyé" in f.columns and len(f) > 0:
            pct_env = int((pd.to_numeric(f["Dossier envoyé"], errors="coerce").fillna(0) > 0).mean() * 100)
        k5.metric("Envoyés (%)", f"{pct_env}%")

        # Graphique : Dossiers par catégorie
        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if "Categorie" in f.columns and not f.empty:
            vc = f["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie","Nombre"]
            st.bar_chart(vc.set_index("Categorie"))
        else:
            st.info("Pas de données catégories avec les filtres actuels.")

        # Graphique : flux par mois
        st.markdown("#### 💵 Flux par mois")
        if not f.empty:
            gdf = f.copy()
            gdf["Montant honoraires (US $)"] = pd.to_numeric(gdf.get("Montant honoraires (US $)", 0), errors="coerce").fillna(0)
            gdf["Autres frais (US $)"]       = pd.to_numeric(gdf.get("Autres frais (US $)", 0), errors="coerce").fillna(0)
            gdf["Payé"]                      = pd.to_numeric(gdf.get("Payé", 0), errors="coerce").fillna(0)
            gdf["Reste"]                     = pd.to_numeric(gdf.get("Reste", 0), errors="coerce").fillna(0)
            gdf["Mois"]                      = gdf["Mois"].astype(str)

            g = (gdf.groupby("Mois", as_index=False)[
                    ["Montant honoraires (US $)","Autres frais (US $)","Payé","Reste"]
                ].sum()
                 .sort_values("Mois")
                 .set_index("Mois"))
            st.line_chart(g)
        else:
            st.info("Aucune donnée pour tracer les flux avec les filtres actuels.")

        # Détails
        st.markdown("#### 📋 Détails (après filtres)")
        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Payé","Reste",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in f.columns]

        view = f.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Reste"]:
            if c in view.columns:
                view[c] = pd.to_numeric(view[c], errors="coerce").fillna(0).map(_fmt_money)
        # Tri robuste (seulement colonnes existantes)
        sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
        view = view.sort_values(by=sort_keys) if sort_keys else view
        st.dataframe(view[show_cols].reset_index(drop=True), use_container_width=True, key=_k("table"))

# ========================
# PARTIE 3/6 — Analyses
# ========================
with tabs[1]:
    st.subheader("📈 Analyses")

    # Helpers locaux
    _SID2 = st.session_state.get("_sid_prefix", "sid")
    def _kA(*parts):
        return _SID2 + "_ana_" + "_".join(str(p) for p in parts)

    def _clean_list(series_like):
        try:
            s = pd.Series(series_like).dropna().astype(str).str.strip()
            s = s[s != ""]
            return sorted(s.unique().tolist())
        except Exception:
            return []

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = _clean_list(df_all.get("Categorie", []))
        subsA   = _clean_list(df_all.get("Sous-categorie", []))
        visasA  = _clean_list(df_all.get("Visa", []))

        c1,c2,c3,c4,c5 = st.columns(5)
        fy = c1.multiselect("Année", yearsA, default=[], key=_kA("y"))
        fm = c2.multiselect("Mois (MM)", monthsA, default=[], key=_kA("m"))
        fc = c3.multiselect("Catégorie", catsA, default=[], key=_kA("c"))
        fs = c4.multiselect("Sous-catégorie", subsA, default=[], key=_kA("s"))
        fv = c5.multiselect("Visa", visasA, default=[], key=_kA("v"))

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        k1,k2,k3,k4 = st.columns(4, gap="small")
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(pd.to_numeric(dfA.get("Montant honoraires (US $)",0), errors="coerce").fillna(0).sum()))
        k3.metric("Payé", _fmt_money(pd.to_numeric(dfA.get("Payé",0), errors="coerce").fillna(0).sum()))
        k4.metric("Reste", _fmt_money(pd.to_numeric(dfA.get("Reste",0), errors="coerce").fillna(0).sum()))

        # % par catégorie
        st.markdown("#### Répartition par catégorie (%)")
        if not dfA.empty and "Categorie" in dfA.columns:
            base = dfA["Categorie"].astype(str).str.strip()
            vc = (base.value_counts(normalize=True)*100).round(1).astype(str) + "%"
            st.dataframe(vc.to_frame("Part"), use_container_width=True, key=_kA("cat_pct"))
        else:
            st.info("Pas de catégorie sur l’échantillon.")

        # Comparaison période A vs B
        st.markdown("#### Comparaison A vs B")
        ca1, ca2, cb1, cb2 = st.columns(4)
        ya = ca1.multiselect("Année (A)", yearsA, default=[], key=_kA("ya"))
        ma = ca2.multiselect("Mois (A)", monthsA, default=[], key=_kA("ma"))
        yb = cb1.multiselect("Année (B)", yearsA, default=[], key=_kA("yb"))
        mb = cb2.multiselect("Mois (B)", monthsA, default=[], key=_kA("mb"))

        def _slice(ylist, mlist):
            d = df_all.copy()
            if ylist: d = d[d["_Année_"].isin(ylist)]
            if mlist: d = d[d["Mois"].astype(str).isin(mlist)]
            return d

        A = _slice(ya, ma)
        B = _slice(yb, mb)
        colA, colB = st.columns(2)
        def _kpi_block(col, lab, dset):
            col.metric(f"Dossiers ({lab})", f"{len(dset)}")
            col.metric("Honoraires", _fmt_money(pd.to_numeric(dset.get("Montant honoraires (US $)",0), errors="coerce").fillna(0).sum()))
            col.metric("Payé", _fmt_money(pd.to_numeric(dset.get("Payé",0), errors="coerce").fillna(0).sum()))
            col.metric("Reste", _fmt_money(pd.to_numeric(dset.get("Reste",0), errors="coerce").fillna(0).sum()))
        with colA:
            _kpi_block(st, "A", A)
        with colB:
            _kpi_block(st, "B", B)


# ======================
# PARTIE 4/6 — Escrow
# ======================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        d = df_all.copy()
        d["Total (US $)"] = _nnum(d["Montant honoraires (US $)"]) + _nnum(d["Autres frais (US $)"])
        agg = d.groupby("Categorie", as_index=False)[["Total (US $)","Payé","Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg["Total (US $)"]).replace([pd.NA, pd.NaT], 0).fillna(0.0) * 100
        st.dataframe(agg, use_container_width=True, key=sid("esc_agg"))
        t1,t2,t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money(float(agg["Total (US $)"].sum())))
        t2.metric("Payé", _fmt_money(float(agg["Payé"].sum())))
        t3.metric("Reste", _fmt_money(float(agg["Reste"].sum())))
        st.caption("NB : si tu veux un escrow 'strict', on peut suivre ce qui est perçu avant 'Dossier envoyé', puis marquer le transfert.")

# ===================================
# PARTIE 4/6 — 👤 Compte client
# ===================================
with tabs[3]:
    st.subheader("👤 Compte client")
    if df_all.empty:
        st.info("Charge d’abord des clients.")
    else:
        names = sorted(df_all["Nom"].dropna().astype(str).unique().tolist())
        sel = st.selectbox("Choisir un client", [""]+names, key=sid("acct_sel"))
        if sel:
            row = df_all[df_all["Nom"].astype(str)==sel].iloc[0].to_dict()

            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Honoraires", _fmt_money(_nnum(row.get("Montant honoraires (US $)",0))))
            c2.metric("Autres frais", _fmt_money(_nnum(row.get("Autres frais (US $)",0))))
            c3.metric("Payé", _fmt_money(_nnum(row.get("Payé",0))))
            c4.metric("Reste", _fmt_money(_nnum(row.get("Reste",0))))

            st.markdown("#### Statuts & Dates")
            s1,s2,s3,s4,s5 = st.columns(5)
            env = int(row.get("Dossier envoyé",0) or 0) == 1
            acc = int(row.get("Dossier approuvé",0) or 0) == 1
            ref = int(row.get("Dossier refusé",0) or 0) == 1
            ann = int(row.get("Dossier annulé",0) or 0) == 1
            rfe = int(row.get("RFE",0) or 0) == 1
            s1.write(f"Envoyé : {'✅' if env else '—'}  | Date : {_safe_str(row.get('Date d’envoi',''))}")
            s2.write(f"Accepté : {'✅' if acc else '—'} | Date : {_safe_str(row.get('Date d’acceptation',''))}")
            s3.write(f"Refusé : {'✅' if ref else '—'}  | Date : {_safe_str(row.get('Date de refus',''))}")
            s4.write(f"Annulé : {'✅' if ann else '—'}  | Date : {_safe_str(row.get('Date d’annulation',''))}")
            s5.write(f"RFE : {'✅' if rfe else '—'}")

            # Paiements (timeline simple + ajout)
            st.markdown("#### Règlements")
            pay_series = row.get("Paiements", [])
            # Affichage
            if isinstance(pay_series, list) and pay_series:
                dfp = pd.DataFrame(pay_series)
                st.dataframe(dfp, use_container_width=True, key=sid("acct_pay"))
            else:
                st.info("Aucun règlement.")

            st.markdown("##### ➕ Ajouter un règlement")
            pcol1,pcol2,pcol3,pcol4 = st.columns(4)
            pdate = pcol1.date_input("Date", value=date.today(), key=sid("pay_date"))
            pamt  = pcol2.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=sid("pay_amt"))
            pmode = pcol3.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=sid("pay_mode"))
            padd  = pcol4.text_input("Note", "", key=sid("pay_note"))
            if st.button("Enregistrer le règlement", key=sid("pay_save")):
                df_live = df_all.copy()
                idx = df_live[df_live["Nom"].astype(str)==sel].index[0]
                # append
                pays = df_live.at[idx, "Paiements"]
                if not isinstance(pays, list):
                    pays = []
                pays.append({
                    "date": (pdate if isinstance(pdate,(date,datetime)) else date.today()).strftime("%Y-%m-%d"),
                    "montant": float(pamt or 0.0),
                    "mode": pmode,
                    "note": padd
                })
                df_live.at[idx, "Paiements"] = pays
                # recalc Payé/Reste
                total = float(df_live.at[idx, "Montant honoraires (US $)"]) + float(df_live.at[idx, "Autres frais (US $)"])
                paye  = sum([float(x.get("montant",0.0) or 0.0) for x in pays])
                df_live.at[idx, "Payé"] = paye
                df_live.at[idx, "Reste"] = max(0.0, total - paye)
                # sauver en mémoire (si l’utilisateur a chargé depuis un fichier disque, tu peux lui proposer d’écrire)
                st.session_state["df_clients_raw"] = df_live.copy()
                st.success("Règlement ajouté. (Pense à exporter / sauvegarder ton fichier)")
                st.cache_data.clear()
                st.rerun()


# =====================================
# PARTIE 5/6 — 🧾 Gestion (CRUD)
# =====================================
with tabs[4]:
    st.subheader("🧾 Gestion des clients")
    if df_all.empty:
        st.info("Charge d’abord des clients.")
    else:
        op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=sid("crud_op"))

        def _next_dossier(df: pd.DataFrame, start=13057) -> int:
            try:
                mx = pd.to_numeric(df.get("Dossier N", pd.Series([start-1])), errors="coerce").fillna(start-1).max()
                return int(mx)+1
            except Exception:
                return start

        def _make_client_id(nom: str, d: date) -> str:
            base = re.sub(r"[^A-Za-z0-9]+","", (nom or "").strip())[:16] or "CL"
            return f"{base}-{d:%Y%m%d}"

        # Sélecteurs Visa (cascade + options depuis visa_map)
        def visa_cascade_editor(prefix_key: str, preset_cat: str="", preset_sub: str="", preset_opts: dict|None=None) -> Tuple[str,str,dict]:
            cats = sorted(list(visa_map.keys()))
            sel_cat = st.selectbox("Catégorie", [""]+cats, index=(cats.index(preset_cat)+1 if preset_cat in cats else 0), key=sid(prefix_key,"cat"))
            sel_sub = ""
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=(subs.index(preset_sub)+1 if preset_sub in subs else 0), key=sid(prefix_key,"sub"))
            opts_dict = {"exclusive": True, "options": []}
            if sel_cat and sel_sub:
                data = visa_map.get(sel_cat, {}).get(sel_sub, {"options":[], "exclusive":True})
                opts = data.get("options", [])
                exclusive = bool(data.get("exclusive", True))
                chosen: List[str] = []
                st.markdown("**Options disponibles :**")
                if exclusive:
                    opt = st.radio("Choix exclusif", [""]+opts, index=0 if not preset_opts else ([""]+opts).index(preset_opts.get("exclusive","")) if preset_opts.get("exclusive","") in opts else 0, key=sid(prefix_key,"opt1"))
                    if opt:
                        chosen = [opt]
                    opts_dict = {"exclusive": opt or None, "options": chosen}
                else:
                    for o in opts:
                        if st.checkbox(o, value=(o in (preset_opts or {}).get("options", [])), key=sid(prefix_key,"opt",o)):
                            chosen.append(o)
                    opts_dict = {"exclusive": None, "options": chosen}
            return sel_cat, sel_sub, opts_dict

        if op == "Ajouter":
            c1,c2,c3 = st.columns(3)
            nom = c1.text_input("Nom", "", key=sid("add_nom"))
            dt  = c2.date_input("Date de création", value=date.today(), key=sid("add_date"))
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=date.today().month-1, key=sid("add_mois"))

            st.markdown("#### 🎯 Choix Visa")
            cat, sub, opts = visa_cascade_editor("add")

            f1,f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, step=50.0, format="%.2f", key=sid("add_h"))
            autre = f2.number_input("Autres frais (US $)", min_value=0.0, step=20.0, format="%.2f", key=sid("add_o"))
            comment = st.text_area("Commentaires (Autres frais)", "", key=sid("add_comm"))

            st.markdown("#### 📌 Statuts initiaux")
            s1,s2,s3,s4,s5 = st.columns(5)
            sent = s1.checkbox("Dossier envoyé", key=sid("add_sent"))
            acc  = s2.checkbox("Dossier approuvé", key=sid("add_acc"))
            ref  = s3.checkbox("Dossier refusé", key=sid("add_ref"))
            ann  = s4.checkbox("Dossier annulé", key=sid("add_ann"))
            rfe  = s5.checkbox("RFE", key=sid("add_rfe"))
            if rfe and not any([sent,acc,ref,ann]):
                st.warning("RFE ne peut être coché qu’avec un autre statut.")

            if st.button("💾 Enregistrer le client", key=sid("btn_add")):
                if not nom or not cat or not sub:
                    st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                    st.stop()
                df_live = df_all.copy()
                did = _make_client_id(nom, dt)
                dossier_n = _next_dossier(df_live, start=13057)
                total = float(honor)+float(autre)
                new_row = {
                    "Dossier N": dossier_n,
                    "ID_Client": did,
                    "Nom": nom,
                    "Date": dt,
                    "Mois": f"{int(mois):02d}",
                    "Categorie": cat,
                    "Sous-categorie": sub,
                    "Visa": (opts.get("exclusive") or sub) if opts else sub,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(autre),
                    "Payé": 0.0,
                    "Reste": total,
                    "RFE": 1 if rfe else 0,
                    "Dossier envoyé": 1 if sent else 0,
                    "Dossier approuvé": 1 if acc else 0,
                    "Dossier refusé": 1 if ref else 0,
                    "Dossier annulé": 1 if ann else 0,
                    "Commentaires": comment,
                    "Paiements": [],
                    "Options": opts or {}
                }
                df_live = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                st.session_state["df_clients_raw"] = df_live.copy()
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()

        elif op == "Modifier":
            names = sorted(df_all["Nom"].dropna().astype(str).unique().tolist())
            target = st.selectbox("Nom", [""]+names, key=sid("mod_sel"))
            if target:
                df_live = df_all.copy()
                idx = df_live[df_live["Nom"].astype(str)==target].index[0]
                row = df_live.loc[idx].copy()

                d1,d2,d3 = st.columns(3)
                nom = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=sid("mod_nom"))
                dval = _to_date(row.get("Date"))
                if dval is None: dval = date.today()
                dt   = d2.date_input("Date de création", value=dval, key=sid("mod_date"))
                mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=(int(_safe_str(row.get("Mois","01")))-1), key=sid("mod_mois"))

                st.markdown("#### 🎯 Choix Visa")
                preset_cat = _safe_str(row.get("Categorie",""))
                preset_sub = _safe_str(row.get("Sous-categorie",""))
                preset_opts = row.get("Options", {})
                if not isinstance(preset_opts, dict):
                    try:
                        preset_opts = json.loads(_safe_str(preset_opts) or "{}")
                    except Exception:
                        preset_opts = {}
                cat, sub, opts = visa_cascade_editor("mod", preset_cat, preset_sub, preset_opts)

                f1,f2 = st.columns(2)
                honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=float(_nnum(row.get("Montant honoraires (US $)",0))), step=50.0, format="%.2f", key=sid("mod_h"))
                autre = f2.number_input("Autres frais (US $)", min_value=0.0, value=float(_nnum(row.get("Autres frais (US $)",0))), step=20.0, format="%.2f", key=sid("mod_o"))
                comment = st.text_area("Commentaires (Autres frais)", _safe_str(row.get("Commentaires","")), key=sid("mod_comm"))

                st.markdown("#### 📌 Statuts")
                s1,s2,s3,s4,s5 = st.columns(5)
                sent = s1.checkbox("Dossier envoyé", value=bool(int(row.get("Dossier envoyé",0) or 0)), key=sid("mod_sent"))
                acc  = s2.checkbox("Dossier approuvé", value=bool(int(row.get("Dossier approuvé",0) or 0)), key=sid("mod_acc"))
                ref  = s3.checkbox("Dossier refusé", value=bool(int(row.get("Dossier refusé",0) or 0)), key=sid("mod_ref"))
                ann  = s4.checkbox("Dossier annulé", value=bool(int(row.get("Dossier annulé",0) or 0)), key=sid("mod_ann"))
                rfe  = s5.checkbox("RFE", value=bool(int(row.get("RFE",0) or 0)), key=sid("mod_rfe"))

                if st.button("💾 Enregistrer les modifications", key=sid("btn_mod")):
                    if not nom or not cat or not sub:
                        st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                        st.stop()
                    df_live.at[idx, "Nom"] = nom
                    df_live.at[idx, "Date"] = dt
                    df_live.at[idx, "Mois"] = f"{int(mois):02d}"
                    df_live.at[idx, "Categorie"] = cat
                    df_live.at[idx, "Sous-categorie"] = sub
                    df_live.at[idx, "Visa"] = (opts.get("exclusive") or sub) if opts else sub
                    df_live.at[idx, "Montant honoraires (US $)"] = float(honor)
                    df_live.at[idx, "Autres frais (US $)"] = float(autre)
                    total = float(honor)+float(autre)
                    paye  = float(_nnum(df_live.at[idx, "Payé"]))
                    df_live.at[idx, "Reste"] = max(0.0, total - paye)
                    df_live.at[idx, "Commentaires"] = comment
                    df_live.at[idx, "Options"] = opts or {}
                    df_live.at[idx, "Dossier envoyé"] = 1 if sent else 0
                    df_live.at[idx, "Dossier approuvé"] = 1 if acc else 0
                    df_live.at[idx, "Dossier refusé"] = 1 if ref else 0
                    df_live.at[idx, "Dossier annulé"] = 1 if ann else 0
                    df_live.at[idx, "RFE"] = 1 if rfe else 0

                    st.session_state["df_clients_raw"] = df_live.copy()
                    st.success("Modifications enregistrées.")
                    st.cache_data.clear()
                    st.rerun()

        elif op == "Supprimer":
            names = sorted(df_all["Nom"].dropna().astype(str).unique().tolist())
            target = st.selectbox("Nom", [""]+names, key=sid("del_sel"))
            if target and st.button("❗ Confirmer la suppression", key=sid("btn_del")):
                df_live = df_all.copy()
                df_live = df_live[df_live["Nom"].astype(str) != target].copy()
                st.session_state["df_clients_raw"] = df_live.copy()
                st.success("Client supprimé.")
                st.cache_data.clear()
                st.rerun()


# ==================================
# PARTIE 6/6 — 📄 Visa (aperçu)
# ==================================
with tabs[5]:
    st.subheader("📄 Visa (aperçu)")
    if df_visa_raw.empty:
        st.info("Charge un fichier Visa.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True, key=sid("visa_preview"))

# =========================
# PARTIE 6/6 — 💾 Export
# =========================
with tabs[6]:
    st.subheader("💾 Export")
    st.caption("Export du fichier Clients uniquement (tu restes décisionnaire du chemin).")

    # Préparer un fichier à télécharger tel quel depuis l’app
    export_buf = BytesIO()
    df_export = st.session_state.get("df_clients_raw", df_all.copy())
    if df_export is None: df_export = pd.DataFrame(columns=REQ_CLIENT_COLS)

    with pd.ExcelWriter(export_buf, engine="openpyxl") as wr:
        df_export.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)

    st.download_button(
        "⬇️ Télécharger Clients.xlsx",
        data=export_buf.getvalue(),
        file_name="Clients.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=sid("dl_clients"),
    )

    st.markdown("---")
    st.caption("💡 Pour **écrire directement sur ton disque** (PC / Drive), utilise la sauvegarde externe du navigateur (ex. choisir un dossier) ou re-charge ce fichier comme source lors de ta prochaine session — l’app retient le dernier mode/état dans `last_paths.json` si tu travailles avec des fichiers locaux.")