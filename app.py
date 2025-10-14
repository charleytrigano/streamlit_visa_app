# ===============================
# 🛂 VISA MANAGER — PARTIE 1/4
# (imports, constantes, helpers,
#  chargement fichiers & onglets)
# ===============================

from __future__ import annotations

import json, re, os, zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Any, Dict, Tuple, List, Optional

import pandas as pd
import streamlit as st

# ---------- Réglages généraux ----------
st.set_page_config(
    page_title="Visa Manager",
    page_icon="🛂",
    layout="wide"
)

# Un identifiant de session pour éviter les collisions de clés Streamlit
SID = "vm"

def skey(*parts: str) -> str:
    """Construit une clé Streamlit unique & stable."""
    return f"{SID}_" + "_".join(str(p) for p in parts if p is not None)

# ---------- Dossiers / mémorisation derniers fichiers ----------
APP_DIR = os.getcwd()
STATE_DIR = os.path.join(APP_DIR, ".visamanager")
os.makedirs(STATE_DIR, exist_ok=True)
LAST_PATHS_JSON = os.path.join(STATE_DIR, "last_paths.json")

def _save_last_paths(clients_path: str | None, visa_path: str | None) -> None:
    try:
        obj = {"clients": clients_path or "", "visa": visa_path or ""}
        with open(LAST_PATHS_JSON, "w", encoding="utf-8") as f:
            json.dump(obj, f)
    except Exception:
        pass

def _load_last_paths() -> Tuple[str | None, str | None]:
    try:
        if os.path.exists(LAST_PATHS_JSON):
            with open(LAST_PATHS_JSON, "r", encoding="utf-8") as f:
                obj = json.load(f)
                return (obj.get("clients") or None, obj.get("visa") or None)
    except Exception:
        pass
    return (None, None)

# ---------- Constantes colonnes ----------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

DOSSIER_COL = "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"

# ---------- Helpers sûrs ----------
def _safe_str(x: Any) -> str:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x)
    except Exception:
        return ""

def _fmt_money(v: float | int | str) -> str:
    try:
        f = float(v)
    except Exception:
        f = 0.0
    return f"${f:,.2f}"

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index, dtype="float64")
    s = df[col]
    try:
        # Nettoyage éventuel (séparateurs, symboles, etc.)
        return (
            s.astype(str)
             .str.replace(r"[^\d\-,.]", "", regex=True)
             .str.replace(",", ".", regex=False)
             .replace("", "0")
             .astype(float)
        )
    except Exception:
        return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _date_for_widget(val: Any) -> Optional[date]:
    """Convertit val -> date ou None pour les widgets Streamlit."""
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.notna(d):
            return d.date()
        return None
    except Exception:
        return None

def _norm(s: str) -> str:
    """Normalise une étiquette pour clés (catégories, options)."""
    s = _safe_str(s).strip().lower()
    # Regex sûre : on garde lettres/chiffres + quelques symboles
    s = re.sub(r"[^a-z0-9+\-/_ ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

# ---------- Aide : retrouver une colonne (avec / sans accents) ----------
def _find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = list(df.columns)
    norm = {i: _norm(c) for i, c in enumerate(cols)}
    norm_cands = [_norm(c) for c in candidates]
    for i, nc in norm.items():
        if nc in norm_cands:
            return cols[i]
    return None

# ---------- Lecture fichiers ----------
@st.cache_data(show_spinner=False)
def read_excel(path: str, sheet: str | None = None) -> pd.DataFrame | Dict[str, pd.DataFrame]:
    if sheet is None:
        # retourne toutes les feuilles
        xls = pd.ExcelFile(path)
        return {sn: xls.parse(sn) for sn in xls.sheet_names}
    else:
        return pd.read_excel(path, sheet_name=sheet)

def _coerce_clients_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Renomme/ajoute les colonnes nécessaires côté Clients."""
    df = df.copy()

    # Harmonisation noms
    col_map = {}
    # Categorie / Sous-categorie / Visa
    c_cat = _find_col(df, ["Categorie", "Catégorie"])
    c_sub = _find_col(df, ["Sous-categorie", "Sous-catégorie", "Sous-categories", "Sous-categories 1"])
    c_vis = _find_col(df, ["Visa"])
    if c_cat and c_cat != "Categorie": col_map[c_cat] = "Categorie"
    if c_sub and c_sub != "Sous-categorie": col_map[c_sub] = "Sous-categorie"
    if c_vis and c_vis != "Visa": col_map[c_vis] = "Visa"

    # Date / Mois
    c_date = _find_col(df, ["Date"])
    c_mois = _find_col(df, ["Mois"])
    if c_date and c_date != "Date": col_map[c_date] = "Date"
    if c_mois and c_mois != "Mois": col_map[c_mois] = "Mois"

    # Identifiants / nom
    c_nom = _find_col(df, ["Nom"])
    c_id  = _find_col(df, ["ID_Client", "ID Client"])
    c_dos = _find_col(df, [DOSSIER_COL, "Dossier", "DossierN"])
    if c_nom and c_nom != "Nom": col_map[c_nom] = "Nom"
    if c_id  and c_id  != "ID_Client": col_map[c_id] = "ID_Client"
    if c_dos and c_dos != DOSSIER_COL: col_map[c_dos] = DOSSIER_COL

    # Montants
    c_h = _find_col(df, [HONO, "Honoraires", "Montant honoraires", "Honoraires (US $)"])
    c_o = _find_col(df, [AUTRE, "Autres frais", "Autres (US $)"])
    c_t = _find_col(df, [TOTAL, "Total", "Total US $"])
    if c_h and c_h != HONO: col_map[c_h] = HONO
    if c_o and c_o != AUTRE: col_map[c_o] = AUTRE
    if c_t and c_t != TOTAL: col_map[c_t] = TOTAL

    # Paiements & statut
    c_pay = _find_col(df, ["Payé", "Paye", "Paye (US $)"])
    if c_pay and c_pay != "Payé": col_map[c_pay] = "Payé"

    for lab in ["Dossier envoyé","Date d'envoi","Dossier accepté","Date d'acceptation",
                "Dossier refusé","Date de refus","Dossier annulé","Date d'annulation","RFE"]:
        c = _find_col(df, [lab])
        if c and c != lab:
            col_map[c] = lab

    # Commentaire
    c_com = _find_col(df, ["Commentaire", "Commentaires", "Notes"])
    if c_com and c_com != "Commentaire":
        col_map[c_com] = "Commentaire"

    if col_map:
        df = df.rename(columns=col_map)

    # Ajouts manquants
    for needed in ["Nom","ID_Client","Categorie","Sous-categorie","Visa","Date","Mois",
                   HONO,AUTRE,TOTAL,"Payé","Reste","Paiements",
                   "Dossier envoyé","Date d'envoi","Dossier accepté","Date d'acceptation",
                   "Dossier refusé","Date de refus","Dossier annulé","Date d'annulation","RFE",
                   DOSSIER_COL,"Commentaire","Options"]:
        if needed not in df.columns:
            df[needed] = "" if needed in ["Nom","ID_Client","Categorie","Sous-categorie","Visa","Date","Mois","Paiements","Commentaire","Options"] else 0

    return df

def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    """Calcule Total/Payé/Reste, fabrique _Année_/_MoisNum_ pour les analyses."""
    df = _coerce_clients_columns(df).copy()

    # Dates & mois
    d = pd.to_datetime(df["Date"], errors="coerce")
    df["_Année_"] = d.dt.year.fillna(0).astype(int)
    # Mois au format MM (texte)
    df["Mois"] = df["Mois"].apply(lambda x: f"{int(_safe_str(x) or 0):02d}" if _safe_str(x).isdigit() else _safe_str(x))
    try:
        df["_MoisNum_"] = pd.to_numeric(df["Mois"], errors="coerce").fillna(d.dt.month).fillna(0).astype(int)
    except Exception:
        df["_MoisNum_"] = d.dt.month.fillna(0).astype(int)

    # Numériques
    h = _safe_num_series(df, HONO)
    o = _safe_num_series(df, AUTRE)
    p = _safe_num_series(df, "Payé")

    if TOTAL not in df.columns or (TOTAL in df.columns and _safe_num_series(df, TOTAL).sum() == 0):
        df[TOTAL] = (h + o).round(2)
    if "Reste" not in df.columns:
        df["Reste"] = (df[TOTAL] - p).clip(lower=0).round(2)
    else:
        df["Reste"] = (df[TOTAL] - p).clip(lower=0).round(2)

    # Statuts => int 0/1
    for lab in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        if lab in df.columns:
            df[lab] = pd.to_numeric(df[lab], errors="coerce").fillna(0).astype(int)

    return df

@st.cache_data(show_spinner=False)
def read_clients_file(path: str, sheet: str = SHEET_CLIENTS) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        # tenter une lecture simple (une seule feuille)
        df = pd.read_excel(path)
    return normalize_clients(df)

# ---------- Référentiel Visa : construction d’une map ----------
@st.cache_data(show_spinner=False)
def read_visa_file(path: str, sheet: str = SHEET_VISA) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        df = pd.read_excel(path)
    return df.copy()

def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    Construit :
    { "Categorie": {
         "Sous-categorie": {
             "exclusive": [liste radio] ou None,
             "options": [cases à cocher]
         }, ...
      }, ... }
    Les colonnes 'COS'/'EOS' deviennent une option exclusive si la valeur == 1 sur la ligne.
    Les autres colonnes avec 1 deviennent des cases à cocher.
    """
    if df_visa is None or df_visa.empty:
        return {}

    # Colonnes de base
    c_cat = _find_col(df_visa, ["Categorie","Catégorie"])
    c_sub = _find_col(df_visa, ["Sous-categorie","Sous-catégorie","Sous-categories","Sous-categories 1","Sous-categories 2"])
    if not c_cat or not c_sub:
        return {}

    cat_values = df_visa[c_cat].astype(str).fillna("")
    sub_values = df_visa[c_sub].astype(str).fillna("")

    # En-têtes des options (ligne 1 = entête, mais on parcourt par ligne)
    all_cols = list(df_visa.columns)
    # Colonnes d’options = toutes les colonnes hors Cat/Sous-cat
    opt_cols = [c for c in all_cols if c not in (c_cat, c_sub)]

    visa_map: Dict[str, Dict[str, Dict[str, Any]]] = {}

    for i, row in df_visa.iterrows():
        cat = _safe_str(row.get(c_cat, "")).strip()
        sub = _safe_str(row.get(c_sub, "")).strip()
        if not cat or not sub:
            continue

        excl: List[str] = []
        others: List[str] = []

        for oc in opt_cols:
            lab = _safe_str(df_visa.columns[df_visa.columns.get_loc(oc)])
            val = row.get(oc, "")
            # on considère "1", 1, True comme actif
            active = False
            try:
                if isinstance(val, str) and val.strip() == "1":
                    active = True
                elif isinstance(val, (int, float)) and float(val) == 1.0:
                    active = True
                elif isinstance(val, bool) and val:
                    active = True
            except Exception:
                active = False

            if active:
                # COS/EOS -> exclusif
                if _norm(lab) in ("cos","eos"):
                    if lab not in excl:
                        excl.append(lab)
                else:
                    if lab not in others:
                        others.append(lab)

        visa_map.setdefault(cat, {})
        visa_map[cat].setdefault(sub, {"exclusive": None, "options": []})
        visa_map[cat][sub]["exclusive"] = excl if excl else None
        visa_map[cat][sub]["options"]   = others

    return visa_map

# ---------- Écriture Clients ----------
def write_df_to_bytes(df: pd.DataFrame, sheet_name: str = SHEET_CLIENTS) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        df.to_excel(wr, index=False, sheet_name=sheet_name)
    return buf.getvalue()

def _write_clients(df: pd.DataFrame, path: str, sheet: str = SHEET_CLIENTS) -> None:
    """Écrit la feuille Clients uniquement."""
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as wr:
            df.to_excel(wr, index=False, sheet_name=sheet)
        _save_last_paths(path, st.session_state.get(skey("files","visa_path")))
    except Exception as e:
        st.error("Erreur à l’écriture du fichier Clients : " + _safe_str(e))

# ---------- Barre latérale : chargement fichiers ----------
st.sidebar.header("📂 Fichiers")

# Récup dernier chemins
last_clients, last_visa = _load_last_paths()

mode = st.sidebar.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"],
                        index=0, key=skey("files","mode"))

df_clients_raw = pd.DataFrame()
df_visa_raw    = pd.DataFrame()
clients_path   = None
visa_path      = None

if mode == "Deux fichiers (Clients & Visa)":
    up_c = st.sidebar.file_uploader("Clients (xlsx)", type=["xlsx"], key=skey("files","upc"))
    up_v = st.sidebar.file_uploader("Visa (xlsx)",    type=["xlsx"], key=skey("files","upv"))

    if up_c is not None:
        clients_bytes = up_c.read()
        tmpc = os.path.join(STATE_DIR, "clients_uploaded.xlsx")
        with open(tmpc, "wb") as f: f.write(clients_bytes)
        clients_path = tmpc

    if up_v is not None:
        visa_bytes = up_v.read()
        tmpv = os.path.join(STATE_DIR, "visa_uploaded.xlsx")
        with open(tmpv, "wb") as f: f.write(visa_bytes)
        visa_path = tmpv

    # Si non upload, tenter la dernière mémorisation
    if not clients_path and last_clients and os.path.exists(last_clients):
        clients_path = last_clients
    if not visa_path and last_visa and os.path.exists(last_visa):
        visa_path = last_visa

    # Lecture
    if clients_path:
        try:
            df_clients_raw = read_clients_file(clients_path)
        except Exception as e:
            st.sidebar.error("Lecture Clients : " + _safe_str(e))
    if visa_path:
        try:
            df_visa_raw = read_visa_file(visa_path)
        except Exception as e:
            st.sidebar.error("Lecture Visa : " + _safe_str(e))

    # Sauvegarder derniers chemins si valides
    if clients_path or visa_path:
        _save_last_paths(clients_path, visa_path)

else:  # un seul fichier avec 2 onglets
    up_one = st.sidebar.file_uploader("Fichier unique (2 onglets 'Clients' & 'Visa')", type=["xlsx"], key=skey("files","upa"))
    if up_one is not None:
        one_bytes = up_one.read()
        tmpp = os.path.join(STATE_DIR, "all_in_one.xlsx")
        with open(tmpp, "wb") as f: f.write(one_bytes)
        clients_path = tmpp
        visa_path    = tmpp
    else:
        # tentative reprise derniers chemins si un seul fichier précédemment
        if last_clients and os.path.exists(last_clients):
            clients_path = last_clients
        if last_visa and os.path.exists(last_visa):
            visa_path = last_visa

    # Lecture
    if clients_path:
        try:
            df_clients_raw = read_clients_file(clients_path, sheet=SHEET_CLIENTS)
        except Exception as e:
            st.sidebar.error("Lecture Clients : " + _safe_str(e))
    if visa_path:
        try:
            df_visa_raw = read_visa_file(visa_path, sheet=SHEET_VISA)
        except Exception as e:
            st.sidebar.error("Lecture Visa : " + _safe_str(e))

    # Sauvegarder derniers chemins
    if clients_path or visa_path:
        _save_last_paths(clients_path, visa_path)

# ---------- Construire la carte Visa (options) ----------
visa_map: Dict[str, Dict[str, Dict[str, Any]]] = {}
try:
    if not df_visa_raw.empty:
        visa_map = build_visa_map(df_visa_raw.copy())
except Exception as e:
    st.error("Erreur construction référentiel Visa : " + _safe_str(e))
    visa_map = {}

# ---------- En-tête & info fichiers ----------
st.title("🛂 Visa Manager")

with st.expander("📄 Fichiers chargés", expanded=False):
    st.write({
        "Mode": mode,
        "Clients": clients_path or "(aucun)",
        "Visa": visa_path or "(aucun)",
        "Mémoire": {"derniers": {"clients": last_clients, "visa": last_visa}}
    })

# ---------- Création des onglets (contenu dans les parties suivantes) ----------
tabs = st.tabs([
    "📊 Dashboard",      # (Partie 2/4)
    "🏦 Escrow",         # (Partie 3/4)
    "👤 Compte client",  # (Partie 3/4)
    "🧾 Gestion",        # (Partie 4/4)
    "📄 Visa (aperçu)",  # (Partie 4/4)
    "💾 Export",         # (Partie 4/4)
    "📈 Analyses"        # (Partie 3/4)
])



# ===============================
# 🧩 PARTIE 2/4 — 📊 DASHBOARD
# ===============================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_clients_raw.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        df_all = normalize_clients(df_clients_raw).copy()

        # --- Filtres ---
        yearsA  = sorted([int(x) for x in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        f1, f2, f3, f4, f5 = st.columns([1,1,1,1,1])
        fy = f1.multiselect("Année", yearsA, default=[], key=skey("dash","years"))
        fm = f2.multiselect("Mois (MM)", monthsA, default=[], key=skey("dash","months"))
        fc = f3.multiselect("Catégorie", catsA, default=[], key=skey("dash","cats"))
        fs = f4.multiselect("Sous-catégorie", subsA, default=[], key=skey("dash","subs"))
        fv = f5.multiselect("Visa", visasA, default=[], key=skey("dash","visas"))

        ff = df_all.copy()
        if fy: ff = ff[ff["_Année_"].isin(fy)]
        if fm: ff = ff[ff["Mois"].astype(str).isin(fm)]
        if fc: ff = ff[ff["Categorie"].astype(str).isin(fc)]
        if fs: ff = ff[ff["Sous-categorie"].astype(str).isin(fs)]
        if fv: ff = ff[ff["Visa"].astype(str).isin(fv)]

        # --- KPI (compacts) ---
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(ff)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(ff, HONO).sum()))
        k3.metric("Autres frais", _fmt_money(_safe_num_series(ff, AUTRE).sum()))
        k4.metric("Payé", _fmt_money(_safe_num_series(ff, "Payé").sum()))
        k5.metric("Reste", _fmt_money(_safe_num_series(ff, "Reste").sum()))

        st.caption("Astuce : utilise les filtres ci-dessus pour affiner les KPI.")

        # --- % par Catégorie / Sous-catégorie ---
        st.markdown("### Répartition (%)")
        cA, cB = st.columns(2)

        with cA:
            if not ff.empty and "Categorie" in ff.columns:
                grp = ff.groupby("Categorie", as_index=False)[TOTAL].sum()
                tot = float(grp[TOTAL].sum())
                if tot > 0:
                    grp["% Total"] = (grp[TOTAL] / tot * 100).round(1)
                    st.dataframe(grp[["Categorie","% Total"]].sort_values("% Total", ascending=False),
                                 use_container_width=True, hide_index=True, key=skey("dash","pct_cat"))
                else:
                    st.info("Aucune valeur de Total pour calculer la répartition par catégorie.")

        with cB:
            if not ff.empty and "Sous-categorie" in ff.columns:
                grp2 = ff.groupby("Sous-categorie", as_index=False)[TOTAL].sum()
                tot2 = float(grp2[TOTAL].sum())
                if tot2 > 0:
                    grp2["% Total"] = (grp2[TOTAL] / tot2 * 100).round(1)
                    st.dataframe(grp2[["Sous-categorie","% Total"]].sort_values("% Total", ascending=False),
                                 use_container_width=True, hide_index=True, key=skey("dash","pct_sub"))
                else:
                    st.info("Aucune valeur de Total pour calculer la répartition par sous-catégorie.")

        # --- Graphiques ---
        st.markdown("### Graphiques")

        # Dossiers par catégorie
        if not ff.empty and "Categorie" in ff.columns:
            vc = ff["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie", "Nombre"]
            try:
                import plotly.express as px
                fig1 = px.bar(vc, x="Categorie", y="Nombre", title="Dossiers par catégorie")
                st.plotly_chart(fig1, use_container_width=True, key=skey("dash","g_cat"))
            except Exception:
                st.bar_chart(vc.set_index("Categorie"), use_container_width=True)

        # Honoraires par mois (MM)
        if not ff.empty and "Mois" in ff.columns:
            tmp = ff.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois")
            if not gm.empty:
                try:
                    import plotly.express as px
                    fig2 = px.line(gm, x="Mois", y=HONO, markers=True, title="Honoraires par mois")
                    st.plotly_chart(fig2, use_container_width=True, key=skey("dash","g_hono_mois"))
                except Exception:
                    st.line_chart(gm.set_index("Mois"), use_container_width=True)

        st.markdown("---")

        # --- Détails (table) ---
        st.markdown("### 📋 Détails des dossiers filtrés")
        view = ff.copy()

        # formatage montants & dates
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in view.columns:
                view[c] = _safe_num_series(view, c).map(_fmt_money)
        if "Date" in view.columns:
            try:
                view["Date"] = pd.to_datetime(view["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                view["Date"] = view["Date"].astype(str)

        # colonnes à afficher (sans doublons)
        base_cols = [
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
            "Date", "Mois", HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"
        ]
        show_cols = [c for c in base_cols if c in view.columns]
        # tri
        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

        st.dataframe(
            view_sorted[show_cols].reset_index(drop=True),
            use_container_width=True,
            key=skey("dash","detail")
        )

        # --- Rappel des filtres actifs ---
        with st.expander("🧾 Filtres actifs", expanded=False):
            st.write({
                "Année": fy or "—",
                "Mois": fm or "—",
                "Catégorie": fc or "—",
                "Sous-catégorie": fs or "—",
                "Visa": fv or "—",
            })



# ===============================
# 🧩 PARTIE 3/4 — Escrow + Compte client + Analyses
# ===============================

# Utilitaire local (sécurise la valeur passée aux date_input)
def _date_for_widget_safe(val: Any, fallback: Optional[date] = None) -> Optional[date]:
    d = _date_for_widget(val)
    if d is None:
        return fallback
    return d

# --------------------------------------------------
# 🏦 ONGLET 2 : Escrow (synthèse simple)
# --------------------------------------------------
with tabs[1]:
    st.subheader("🏦 Escrow — synthèse")
    if df_clients_raw.empty:
        st.info("Aucun client chargé.")
    else:
        dfE = normalize_clients(df_clients_raw).copy()
        dfE["Payé"]  = _safe_num_series(dfE, "Payé")
        dfE["Reste"] = _safe_num_series(dfE, "Reste")
        dfE[TOTAL]   = _safe_num_series(dfE, TOTAL)

        # KPI compacts
        k1, k2, k3 = st.columns([1,1,1])
        k1.metric("Total (US $)", _fmt_money(float(dfE[TOTAL].sum())))
        k2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        k3.metric("Reste", _fmt_money(float(dfE["Reste"].sum())))

        # Agrégat par catégorie
        st.markdown("### Par catégorie")
        agg = dfE.groupby("Categorie", as_index=False)[[TOTAL, "Payé", "Reste"]].sum()
        if not agg.empty:
            agg["% Payé"] = (agg["Payé"] / agg[TOTAL] * 100).fillna(0).round(1)
            st.dataframe(agg, use_container_width=True, hide_index=True, key=skey("escrow","agg"))
        else:
            st.info("Aucune donnée agrégée.")

        st.caption(
            "NB : on considère ici l’« escrow » comme la partie honoraires à encaisser. "
            "Tu peux filtrer dans le Dashboard pour isoler des sous-ensembles."
        )

# --------------------------------------------------
# 👤 ONGLET 3 : Compte client (suivi dossier + paiements)
# --------------------------------------------------
with tabs[2]:
    st.subheader("👤 Compte client")

    if df_clients_raw.empty:
        st.info("Aucun client chargé.")
    else:
        live = normalize_clients(df_clients_raw).copy()

        # Sélection du client
        c1, c2 = st.columns([2,1])
        names = sorted(live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in live.columns else []
        sel_name = c1.selectbox("Nom", [""] + names, index=0, key=skey("acct","name"))
        subset = live[live["Nom"].astype(str) == sel_name].copy() if sel_name else pd.DataFrame()

        if subset.empty:
            st.info("Choisis un client pour afficher le compte.")
        else:
            # Si plusieurs lignes (plusieurs dossiers homonymes), choisir par ID_Client
            ids = subset["ID_Client"].astype(str).tolist() if "ID_Client" in subset.columns else []
            sel_id = c2.selectbox("ID_Client", [""] + ids, index=1 if ids else 0, key=skey("acct","id"))
            row = subset[subset["ID_Client"].astype(str) == sel_id].iloc[0] if sel_id else subset.iloc[0]

            # En-tête compte
            st.markdown("### Dossier")
            h1, h2, h3, h4 = st.columns([1,1,1,1])
            h1.write(f"**Dossier N** : {_safe_str(row.get(DOSSIER_COL,''))}")
            h2.write(f"**Catégorie** : {_safe_str(row.get('Categorie',''))}")
            h3.write(f"**Sous-catégorie** : {_safe_str(row.get('Sous-categorie',''))}")
            h4.write(f"**Visa** : {_safe_str(row.get('Visa',''))}")

            # Montants
            st.markdown("### Montants")
            m1, m2, m3, m4 = st.columns([1,1,1,1])
            m1.metric("Honoraires", _fmt_money(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0]))
            m2.metric("Autres frais", _fmt_money(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0]))
            m3.metric("Payé", _fmt_money(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0]))
            reste_val = float(_safe_num_series(pd.DataFrame([row]), "Reste").iloc[0])
            m4.metric("Reste", _fmt_money(reste_val))

            # Historique paiements (stocké dans colonne Paiements sous forme liste JSON ou liste)
            st.markdown("### Paiements")
            pay_list = []
            raw_pay = row.get("Paiements", [])
            if isinstance(raw_pay, str):
                try:
                    pay_list = json.loads(raw_pay) if raw_pay.strip() else []
                except Exception:
                    pay_list = []
            elif isinstance(raw_pay, list):
                pay_list = raw_pay
            else:
                pay_list = []

            # Tableau historique (si vide -> info)
            if pay_list:
                jdf = pd.DataFrame(pay_list)
                # normaliser colonnes
                if "amount" in jdf.columns:
                    jdf["amount"] = pd.to_numeric(jdf["amount"], errors="coerce").fillna(0.0)
                    jdf["amount_fmt"] = jdf["amount"].map(_fmt_money)
                if "date" in jdf.columns:
                    try:
                        jdf["date"] = pd.to_datetime(jdf["date"], errors="coerce").dt.date.astype(str)
                    except Exception:
                        jdf["date"] = jdf["date"].astype(str)
                showp = [c for c in ["date","method","amount_fmt","note"] if c in jdf.columns]
                if showp:
                    st.dataframe(jdf[showp], use_container_width=True, hide_index=True, key=skey("acct","payhist"))
                else:
                    st.write(pay_list)
            else:
                st.info("Aucun paiement enregistré.")

            # Ajouter un paiement si reste > 0
            if reste_val > 0:
                st.markdown("#### ➕ Ajouter un paiement")
                p1, p2, p3, p4 = st.columns([1,1,1,2])
                p_date = p1.date_input("Date", value=date.today(), key=skey("acct","pdate"))
                p_method = p2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], index=0, key=skey("acct","pmethod"))
                p_amount = p3.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=skey("acct","pamount"))
                p_note = p4.text_input("Note (optionnel)", "", key=skey("acct","pnote"))

                if st.button("💾 Enregistrer le paiement", key=skey("acct","pay_save")):
                    add = float(p_amount or 0.0)
                    if add <= 0:
                        st.warning("Le montant doit être > 0.")
                        st.stop()

                    # recharger fichier
                    live2 = read_clients_file(st.session_state.get(skey("files","clients_path")) or st.session_state.get(skey("files","all_path"), ""))
                    if live2.empty:
                        live2 = live

                    mask = (live2["ID_Client"].astype(str) == _safe_str(row.get("ID_Client","")))
                    if not mask.any():
                        st.error("Ligne introuvable pour ce client.")
                        st.stop()

                    # récupérer / normaliser paiements
                    raw = live2.loc[mask, "Paiements"].iloc[0]
                    if isinstance(raw, str):
                        try:
                            pay_arr = json.loads(raw) if raw.strip() else []
                        except Exception:
                            pay_arr = []
                    elif isinstance(raw, list):
                        pay_arr = raw
                    else:
                        pay_arr = []

                    pay_arr.append({
                        "date": (p_date if isinstance(p_date, (date, datetime)) else date.today()).strftime("%Y-%m-%d"),
                        "method": p_method,
                        "amount": float(add),
                        "note": _safe_str(p_note),
                    })

                    # recalc payé / reste
                    honor = float(_safe_num_series(live2.loc[mask].copy(), HONO).iloc[0])
                    other = float(_safe_num_series(live2.loc[mask].copy(), AUTRE).iloc[0])
                    total = honor + other
                    paye  = sum(float(x.get("amount",0.0) or 0.0) for x in pay_arr)
                    reste = max(0.0, total - paye)

                    # écrire
                    live2.loc[mask, "Paiements"] = [pay_arr]
                    live2.loc[mask, "Payé"] = paye
                    live2.loc[mask, "Reste"] = reste

                    # persister
                    _write_clients(live2, st.session_state.get(skey("files","clients_path")) or st.session_state.get(skey("files","all_path"), ""))
                    st.success("Paiement ajouté.")
                    st.cache_data.clear()
                    st.rerun()

            # Statuts dossier
            st.markdown("### Statuts du dossier")
            s1, s2, s3, s4, s5 = st.columns(5)
            envoye   = int(row.get("Dossier envoyé", 0) or 0) == 1
            accepte  = int(row.get("Dossier accepté", 0) or 0) == 1
            refuse   = int(row.get("Dossier refusé", 0) or 0) == 1
            annule   = int(row.get("Dossier annulé", 0) or 0) == 1
            rfe      = int(row.get("RFE", 0) or 0) == 1

            sent   = s1.checkbox("Dossier envoyé", value=envoye, key=skey("acct","sent"))
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget_safe(row.get("Date d'envoi")), key=skey("acct","sentd"))
            acc    = s2.checkbox("Dossier accepté", value=accepte, key=skey("acct","acc"))
            acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget_safe(row.get("Date d'acceptation")), key=skey("acct","accd"))
            ref    = s3.checkbox("Dossier refusé", value=refuse, key=skey("acct","ref"))
            ref_d  = s3.date_input("Date de refus", value=_date_for_widget_safe(row.get("Date de refus")), key=skey("acct","refd"))
            ann    = s4.checkbox("Dossier annulé", value=annule, key=skey("acct","ann"))
            ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget_safe(row.get("Date d'annulation")), key=skey("acct","annd"))
            rfe_v  = s5.checkbox("RFE", value=rfe, key=skey("acct","rfe"))

            if rfe_v and not any([sent, acc, ref, ann]):
                st.warning("⚠️ La case RFE ne peut être cochée qu’avec un autre statut (envoyé, accepté, refusé ou annulé).")

            if st.button("💾 Enregistrer les statuts", key=skey("acct","savestatus")):
                live3 = read_clients_file(st.session_state.get(skey("files","clients_path")) or st.session_state.get(skey("files","all_path"), ""))
                if live3.empty:
                    live3 = live
                mask = (live3["ID_Client"].astype(str) == _safe_str(row.get("ID_Client","")))
                if not mask.any():
                    st.error("Ligne introuvable.")
                    st.stop()
                live3.loc[mask, "Dossier envoyé"]      = 1 if sent else 0
                live3.loc[mask, "Date d'envoi"]        = sent_d
                live3.loc[mask, "Dossier accepté"]     = 1 if acc else 0
                live3.loc[mask, "Date d'acceptation"]  = acc_d
                live3.loc[mask, "Dossier refusé"]      = 1 if ref else 0
                live3.loc[mask, "Date de refus"]       = ref_d
                live3.loc[mask, "Dossier annulé"]      = 1 if ann else 0
                live3.loc[mask, "Date d'annulation"]   = ann_d
                live3.loc[mask, "RFE"]                 = 1 if rfe_v else 0

                _write_clients(live3, st.session_state.get(skey("files","clients_path")) or st.session_state.get(skey("files","all_path"), ""))
                st.success("Statuts enregistrés.")
                st.cache_data.clear()
                st.rerun()

# --------------------------------------------------
# 📈 ONGLET 7 : Analyses (KPI + graph + comparaison)
# --------------------------------------------------
with tabs[6]:
    st.subheader("📈 Analyses")

    if df_clients_raw.empty:
        st.info("Aucun client chargé.")
    else:
        dfA = normalize_clients(df_clients_raw).copy()

        yearsA  = sorted([int(x) for x in pd.to_numeric(dfA["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(dfA["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in dfA.columns else []
        subsA   = sorted(dfA["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in dfA.columns else []
        visasA  = sorted(dfA["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in dfA.columns else []

        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=skey("an","y"))
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=skey("an","m"))
        fc = a3.multiselect("Catégorie", catsA, default=[], key=skey("an","c"))
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=skey("an","s"))
        fv = a5.multiselect("Visa", visasA, default=[], key=skey("an","v"))

        fa = dfA.copy()
        if fy: fa = fa[fa["_Année_"].isin(fy)]
        if fm: fa = fa[fa["Mois"].astype(str).isin(fm)]
        if fc: fa = fa[fa["Categorie"].astype(str).isin(fc)]
        if fs: fa = fa[fa["Sous-categorie"].astype(str).isin(fs)]
        if fv: fa = fa[fa["Visa"].astype(str).isin(fv)]

        # KPI compacts
        k1, k2, k3, k4 = st.columns([1,1,1,1])
        k1.metric("Dossiers", f"{len(fa)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(fa, HONO).sum()))
        k3.metric("Payé", _fmt_money(_safe_num_series(fa, "Payé").sum()))
        k4.metric("Reste", _fmt_money(_safe_num_series(fa, "Reste").sum()))

        st.markdown("### Graphiques")
        # Dossiers par catégorie
        if not fa.empty and "Categorie" in fa.columns:
            vc = fa["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie", "Nombre"]
            try:
                import plotly.express as px
                figA = px.bar(vc, x="Categorie", y="Nombre", title="Dossiers par catégorie")
                st.plotly_chart(figA, use_container_width=True, key=skey("an","g_cat"))
            except Exception:
                st.bar_chart(vc.set_index("Categorie"), use_container_width=True)

        # Honoraires par mois
        if not fa.empty and "Mois" in fa.columns:
            tmp = fa.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)[HONO].sum().sort_values("Mois")
            if not gm.empty:
                try:
                    import plotly.express as px
                    figB = px.line(gm, x="Mois", y=HONO, markers=True, title="Honoraires par mois")
                    st.plotly_chart(figB, use_container_width=True, key=skey("an","g_hono"))
                except Exception:
                    st.line_chart(gm.set_index("Mois"), use_container_width=True)

        # Comparaison A vs B (par année/mois)
        st.markdown("### Comparaison période A vs B")

        ca, cb = st.columns(2)
        with ca:
            pa_years = ca.multiselect("Année (A)", yearsA, default=[], key=skey("cmp","ya"))
            pa_month = ca.multiselect("Mois (A)", monthsA, default=[], key=skey("cmp","ma"))
        with cb:
            pb_years = cb.multiselect("Année (B)", yearsA, default=[], key=skey("cmp","yb"))
            pb_month = cb.multiselect("Mois (B)", monthsA, default=[], key=skey("cmp","mb"))

        def filt(df: pd.DataFrame, yrs: List[int], mos: List[str]) -> pd.DataFrame:
            x = df.copy()
            if yrs: x = x[x["_Année_"].isin(yrs)]
            if mos: x = x[x["Mois"].astype(str).isin(mos)]
            return x

        A = filt(dfA, pa_years, pa_month)
        B = filt(dfA, pb_years, pb_month)

        g1, g2, g3 = st.columns(3)
        g1.metric("Dossiers A", f"{len(A)}")
        g2.metric("Dossiers B", f"{len(B)}")
        delta = len(A) - len(B)
        g3.metric("Δ A vs B", f"{delta:+d}")

        # Tableau détaillé filtré
        st.markdown("### 🧾 Détails des dossiers filtrés")
        det = fa.copy()
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).map(_fmt_money)
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

        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=skey("an","detail"))



# ===============================
# 🧩 PARTIE 4/4 — GESTION (CRUD) + VISA (aperçu) + EXPORT
# ===============================

# --------------------------------------------------
# 🧾 ONGLET 4 : Clients (liste simple, lecture seule)
# --------------------------------------------------
with tabs[3]:
    st.subheader("👥 Clients — liste")
    if df_clients_raw.empty:
        st.info("Aucun client chargé.")
    else:
        view = normalize_clients(df_clients_raw).copy()

        # petite zone de filtres
        yearsA  = sorted([int(x) for x in pd.to_numeric(view["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(view["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in view.columns else []
        subsA   = sorted(view["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in view.columns else []
        visasA  = sorted(view["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in view.columns else []

        c1, c2, c3, c4, c5 = st.columns(5)
        fy = c1.multiselect("Année", yearsA, default=[], key=skey("cl","y"))
        fm = c2.multiselect("Mois (MM)", monthsA, default=[], key=skey("cl","m"))
        fc = c3.multiselect("Catégorie", catsA, default=[], key=skey("cl","c"))
        fs = c4.multiselect("Sous-catégorie", subsA, default=[], key=skey("cl","s"))
        fv = c5.multiselect("Visa", visasA, default=[], key=skey("cl","v"))

        if fy: view = view[view["_Année_"].isin(fy)]
        if fm: view = view[view["Mois"].astype(str).isin(fm)]
        if fc: view = view[view["Categorie"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        # formatage
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in view.columns:
                view[c] = _safe_num_series(view, c).map(_fmt_money)
        if "Date" in view.columns:
            try:
                view["Date"] = pd.to_datetime(view["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                view["Date"] = view["Date"].astype(str)

        show_cols = [c for c in [
            DOSSIER_COL,"ID_Client","Nom","Categorie","Sous-categorie","Visa",
            "Date","Mois", HONO, AUTRE, TOTAL, "Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in view.columns]

        sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

        st.dataframe(view_sorted[show_cols].reset_index(drop=True),
                     use_container_width=True, key=skey("cl","table"))


# --------------------------------------------------
# 🧾 ONGLET 5 : Gestion (CRUD complet)
# --------------------------------------------------
with tabs[4]:
    st.subheader("🧾 Gestion — Ajouter / Modifier / Supprimer")

    # état courant depuis le chemin mémorisé
    clients_path_curr = st.session_state.get(skey("files","clients_path")) or st.session_state.get(skey("files","all_path"), "")
    df_live = read_clients_file(clients_path_curr)
    if df_live.empty and not df_clients_raw.empty:
        df_live = normalize_clients(df_clients_raw).copy()

    op = st.radio("Action", ["Ajouter","Modifier","Supprimer"], horizontal=True, key=skey("crud","op"))

    # ---------- AJOUT ----------
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")

        d1, d2, d3 = st.columns(3)
        nom  = d1.text_input("Nom", "", key=skey("add","nom"))
        dval = date.today()
        dt   = d2.date_input("Date de création", value=dval, key=skey("add","date"))
        mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=int(date.today().month)-1, key=skey("add","mois"))

        # Cascade Visa
        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""]+cats, index=0, key=skey("add","cat"))
        sel_sub = ""
        visa_final, opts_dict, info_msg = "", {"exclusive": None, "options": []}, ""
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=skey("add","sub"))
            if sel_sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=skey("add","opts"), preselected={}
                )
        if info_msg:
            st.info(info_msg)

        f1, f2 = st.columns(2)
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=skey("add","hono"))
        other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=skey("add","autre"))
        comment_autre = st.text_area("Commentaire (Autres frais)", "", key=skey("add","comment_autre"))

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
            st.warning("⚠️ RFE ne peut être coché qu’avec un autre statut (envoyé/accepté/refusé/annulé).")

        if st.button("💾 Enregistrer le client", key=skey("add","save")):
            if not nom:
                st.warning("Le nom est requis.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Choisis la catégorie et la sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            paye  = 0.0
            reste = max(0.0, total - paye)
            did   = _make_client_id(nom, dt)
            dossier_n = _next_dossier(df_live, start=13057)

            new_row = {
                DOSSIER_COL: dossier_n,
                "ID_Client": did,
                "Nom": nom,
                "Date": dt,
                "Mois": f"{int(mois):02d}" if isinstance(mois,(int,str)) else _safe_str(mois),
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": (visa_final if visa_final else sel_sub),
                HONO: float(honor),
                AUTRE: float(other),
                "Commentaire frais": comment_autre,
                TOTAL: total,
                "Payé": paye,
                "Reste": reste,
                "Paiements": [],
                "Options": opts_dict,
                "Dossier envoyé": 1 if sent else 0,
                "Date d'envoi": dt if sent and not sent_d else sent_d,
                "Dossier accepté": 1 if acc else 0,
                "Date d'acceptation": acc_d,
                "Dossier refusé": 1 if ref else 0,
                "Date de refus": ref_d,
                "Dossier annulé": 1 if ann else 0,
                "Date d'annulation": ann_d,
                "RFE": 1 if rfe else 0,
            }
            df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            _write_clients(df_new, clients_path_curr)
            st.success("Client ajouté.")
            st.cache_data.clear()
            st.rerun()

    # ---------- MODIFICATION ----------
    elif op == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if df_live.empty:
            st.info("Aucun client.")
        else:
            ids = df_live["ID_Client"].dropna().astype(str).tolist() if "ID_Client" in df_live.columns else []
            sel_id = st.selectbox("ID_Client", [""]+sorted(ids), index=0, key=skey("mod","id"))
            if not sel_id:
                st.stop()

            mask = (df_live["ID_Client"].astype(str) == sel_id)
            if not mask.any():
                st.warning("Ligne introuvable.")
                st.stop()
            row = df_live[mask].iloc[0].copy()

            d1, d2, d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=skey("mod","nom"))
            dval = _date_for_widget_safe(row.get("Date"), fallback=date.today())
            dt   = d2.date_input("Date de création", value=dval, key=skey("mod","date"))
            mois_def = _safe_str(row.get("Mois","01"))
            mois_idx = max(0, min(11, int(mois_def or "1")-1)) if mois_def.isdigit() else 0
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=mois_idx, key=skey("mod","mois"))

            # Cascade Visa (avec présélection)
            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""]+cats,
                                   index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                                   key=skey("mod","cat"))
            sel_sub = _safe_str(row.get("Sous-categorie",""))
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""]+subs,
                                       index=(subs.index(sel_sub)+1 if sel_sub in subs else 0),
                                       key=skey("mod","sub"))

            # options existantes
            preset_opts = row.get("Options", {})
            if not isinstance(preset_opts, dict):
                try:
                    preset_opts = json.loads(_safe_str(preset_opts) or "{}")
                    if not isinstance(preset_opts, dict):
                        preset_opts = {}
                except Exception:
                    preset_opts = {}

            visa_final, opts_dict, info_msg = "", {"exclusive": None, "options": []}, ""
            if sel_cat and sel_sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=skey("mod","opts"), preselected=preset_opts
                )
            if info_msg:
                st.info(info_msg)

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0]),
                                    step=50.0, format="%.2f", key=skey("mod","hono"))
            other = f2.number_input("Autres frais (US $)", min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0]),
                                    step=20.0, format="%.2f", key=skey("mod","autre"))
            comment_autre = st.text_area("Commentaire (Autres frais)",
                                         _safe_str(row.get("Commentaire frais","")),
                                         key=skey("mod","comment_autre"))

            st.markdown("#### 📌 Statuts")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent   = s1.checkbox("Dossier envoyé", value=bool(int(row.get("Dossier envoyé",0) or 0)), key=skey("mod","sent"))
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget_safe(row.get("Date d'envoi")),
                                   key=skey("mod","sentd"))
            acc    = s2.checkbox("Dossier accepté", value=bool(int(row.get("Dossier accepté",0) or 0)), key=skey("mod","acc"))
            acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget_safe(row.get("Date d'acceptation")),
                                   key=skey("mod","accd"))
            ref    = s3.checkbox("Dossier refusé", value=bool(int(row.get("Dossier refusé",0) or 0)), key=skey("mod","ref"))
            ref_d  = s3.date_input("Date de refus", value=_date_for_widget_safe(row.get("Date de refus")),
                                   key=skey("mod","refd"))
            ann    = s4.checkbox("Dossier annulé", value=bool(int(row.get("Dossier annulé",0) or 0)), key=skey("mod","ann"))
            ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget_safe(row.get("Date d'annulation")),
                                   key=skey("mod","annd"))
            rfe_v  = s5.checkbox("RFE", value=bool(int(row.get("RFE",0) or 0)), key=skey("mod","rfe"))

            if rfe_v and not any([sent, acc, ref, ann]):
                st.warning("⚠️ RFE ne peut être coché qu’avec un autre statut.")

            if st.button("💾 Enregistrer les modifications", key=skey("mod","save")):
                if not nom:
                    st.warning("Le nom est requis.")
                    st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisis la catégorie et la sous-catégorie.")
                    st.stop()

                total = float(honor) + float(other)
                paye  = float(_safe_num_series(df_live[mask].copy(), "Payé").iloc[0]) if "Payé" in df_live.columns else 0.0
                reste = max(0.0, total - paye)

                df_live.loc[mask, "Nom"] = nom
                df_live.loc[mask, "Date"] = dt
                df_live.loc[mask, "Mois"] = f"{int(mois):02d}" if isinstance(mois,(int,str)) else _safe_str(mois)
                df_live.loc[mask, "Categorie"] = sel_cat
                df_live.loc[mask, "Sous-categorie"] = sel_sub
                df_live.loc[mask, "Visa"] = (visa_final if visa_final else sel_sub)
                df_live.loc[mask, HONO] = float(honor)
                df_live.loc[mask, AUTRE] = float(other)
                df_live.loc[mask, "Commentaire frais"] = comment_autre
                df_live.loc[mask, TOTAL] = total
                df_live.loc[mask, "Reste"] = reste
                df_live.loc[mask, "Options"] = [opts_dict]

                df_live.loc[mask, "Dossier envoyé"] = 1 if sent else 0
                df_live.loc[mask, "Date d'envoi"] = sent_d
                df_live.loc[mask, "Dossier accepté"] = 1 if acc else 0
                df_live.loc[mask, "Date d'acceptation"] = acc_d
                df_live.loc[mask, "Dossier refusé"] = 1 if ref else 0
                df_live.loc[mask, "Date de refus"] = ref_d
                df_live.loc[mask, "Dossier annulé"] = 1 if ann else 0
                df_live.loc[mask, "Date d'annulation"] = ann_d
                df_live.loc[mask, "RFE"] = 1 if rfe_v else 0

                _write_clients(df_live, clients_path_curr)
                st.success("Modifications enregistrées.")
                st.cache_data.clear()
                st.rerun()

    # ---------- SUPPRESSION ----------
    elif op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client.")
        else:
            ids = df_live["ID_Client"].dropna().astype(str).tolist() if "ID_Client" in df_live.columns else []
            sel = st.selectbox("ID_Client", [""]+sorted(ids), index=0, key=skey("del","id"))
            if sel:
                row = df_live[df_live["ID_Client"].astype(str)==sel].iloc[0]
                st.write({
                    "Dossier N": row.get(DOSSIER_COL,""),
                    "Nom": row.get("Nom",""),
                    "Visa": row.get("Visa",""),
                })
                if st.button("❗ Confirmer la suppression", key=skey("del","go")):
                    df_new = df_live[df_live["ID_Client"].astype(str) != sel].copy()
                    _write_clients(df_new, clients_path_curr)
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()


# --------------------------------------------------
# 📄 ONGLET 6 : Visa (aperçu)
# --------------------------------------------------
with tabs[5]:
    st.subheader("📄 Visa — aperçu du référentiel")
    if df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        st.markdown("#### Catégories et sous-catégories")
        cats = sorted(df_visa_raw["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_visa_raw.columns else []
        sel_cat = st.selectbox("Catégorie", [""]+cats, index=0, key=skey("visa","cat"))
        if sel_cat:
            subs = sorted(df_visa_raw.loc[df_visa_raw["Categorie"].astype(str)==sel_cat, "Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_visa_raw.columns else []
            sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=skey("visa","sub"))
        else:
            sel_sub = ""

        st.markdown("#### Options disponibles")
        if sel_cat and sel_sub:
            # montrer les options (selon visa_map)
            visa_final, opts, info = build_visa_option_selector(
                visa_map, sel_cat, sel_sub, keyprefix=skey("visa","opts"), preselected={}
            )
            if info:
                st.info(info)
            st.caption(f"Visa construit : **{visa_final or sel_sub}**")
        else:
            st.info("Choisis une catégorie puis une sous-catégorie pour voir les options.")


# --------------------------------------------------
# 💾 Export global (Clients + Visa) + Sauvegardes rapides
# --------------------------------------------------
st.markdown("---")
st.subheader("💾 Export / Sauvegardes")

colz1, colz2, colz3 = st.columns([1,1,2])

# 1) Export ZIP (Clients normalisés + Visa tel quel si possible)
with colz1:
    if st.button("Préparer l’export ZIP", key=skey("exp","zipbtn")):
        try:
            from io import BytesIO
            import zipfile

            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients (toujours normalisés)
                if not df_clients_raw.empty:
                    df_export = normalize_clients(df_clients_raw).copy()
                    with BytesIO() as xbuf:
                        with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                            df_export.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                        zf.writestr("Clients.xlsx", xbuf.getvalue())
                # Visa (fichier source si dispo, sinon feuille unique)
                vpath = st.session_state.get(skey("files","visa_path")) or st.session_state.get(skey("files","all_path"), "")
                if vpath:
                    try:
                        zf.write(vpath, "Visa.xlsx")
                    except Exception:
                        try:
                            with BytesIO() as vb:
                                with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                    df_visa_raw.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                                zf.writestr("Visa.xlsx", vb.getvalue())
                        except Exception:
                            pass

            st.session_state[skey("exp","zipdata")] = buf.getvalue()
            st.success("Export prêt.")
        except Exception as e:
            st.error("Erreur export : " + _safe_str(e))

with colz2:
    if st.session_state.get(skey("exp","zipdata")):
        st.download_button(
            "⬇️ Télécharger l’export (ZIP)",
            data=st.session_state[skey("exp","zipdata")],
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
            key=skey("exp","zipdl"),
        )

# 2) Sauvegarde rapide du classeur Clients courant (au chemin d’origine)
with colz3:
    st.markdown("**Sauvegarde rapide**")
    path_clients = st.session_state.get(skey("files","clients_path")) or st.session_state.get(skey("files","all_path"), "")
    if not path_clients:
        st.info("Aucun chemin de sauvegarde connu (charge un fichier via la barre latérale).")
    else:
        if st.button("💾 Écraser le fichier Clients", key=skey("exp","save_clients")):
            try:
                # Toujours ré-écrire la version normalisée
                live_write = normalize_clients(read_clients_file(path_clients))
                with pd.ExcelWriter(path_clients, engine="openpyxl") as wr:
                    live_write.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                st.success(f"Clients sauvegardés dans : {path_clients}")
            except Exception as e:
                st.error("Échec sauvegarde Clients : " + _safe_str(e))

        # Si on a aussi un visa_path et que l’utilisateur veut tout ré-emballer dans un seul classeur
        visa_path = st.session_state.get(skey("files","visa_path"))
        if visa_path and st.button("💾 Créer un fichier unique (2 onglets)", key=skey("exp","save_both")):
            try:
                out = os.path.join(os.path.dirname(path_clients), "export_clients_visa.xlsx")
                dfC = normalize_clients(read_clients_file(path_clients))
                dfV = df_visa_raw.copy()
                with pd.ExcelWriter(out, engine="openpyxl") as wr:
                    dfC.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                    dfV.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                st.success(f"Fichier unique créé : {out}")
            except Exception as e:
                st.error("Échec création du classeur 2 onglets : " + _safe_str(e))