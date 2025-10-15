# ========== BLOC 1/10 — Setup, chargement fichiers & création des onglets ==========

import io
import os
import json
import uuid
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import numpy as np
import streamlit as st

# ---------- Config de page ----------
st.set_page_config(page_title="Visa Manager", layout="wide")

# ---------- Clé de session stable pour les widgets ----------
SID = st.session_state.get("_sid", None)
if not SID:
    SID = uuid.uuid4().hex[:6]
    st.session_state["_sid"] = SID

def skey(*parts: str) -> str:
    """Fabrique une clé unique et stable par widget."""
    return f"{'_'.join(parts)}_{SID}"

# ---------- Dossier de travail local (pour mémoriser les derniers chemins) ----------
WORK_DIR = Path(".")
STATE_FILE = WORK_DIR / ".visa_manager_state.json"

def _safe_str(x) -> str:
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x) if x is not None else ""

def _to_num(x, default=0.0) -> float:
    """Convertit en float de façon robuste."""
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = _safe_str(x)
    if not s:
        return float(default)
    # Nettoyage des symboles ($, espaces, etc.)
    s = s.replace("$", "").replace("€", "").replace(" ", "").replace("\xa0", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return float(default)

def _ensure_cols(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    """Ajoute les colonnes manquantes comme vides/NaN."""
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    return df

# ---------- Colonnes attendues côté Clients (selon ta trame) ----------
CLIENT_COLS = [
    "ID_Client", "Dossier N", "Nom", "Date", "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde",
    "Acompte 1", "Acompte 2",
    "RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé",
    "Commentaires"
]

# ---------- Lecture fichier générique (xlsx/csv) ----------
def read_any_table(file_or_path) -> pd.DataFrame | None:
    if file_or_path is None:
        return None
    try:
        if hasattr(file_or_path, "read"):  # UploadedFile streamlit
            name = getattr(file_or_path, "name", "uploaded")
            data = file_or_path.read()
            bio = io.BytesIO(data)
            if name.lower().endswith(".csv"):
                df = pd.read_csv(io.BytesIO(data), encoding="utf-8", sep=",")
                return df
            else:
                # essaie de lire 1er onglet par défaut
                return pd.read_excel(bio)
        else:
            p = str(file_or_path)
            if p.lower().endswith(".csv"):
                return pd.read_csv(p, encoding="utf-8", sep=",")
            else:
                return pd.read_excel(p)
    except Exception as e:
        st.error(f"Erreur lecture: {e}")
        return None

def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    """Force la trame attendue et types essentiels."""
    if df is None or df.empty:
        return pd.DataFrame(columns=CLIENT_COLS)

    # Aligne les noms de colonnes possibles -> cibles
    rename_map = {
        "Categorie": "Categories",
        "Catégorie": "Categories",
        "Sous-catégorie": "Sous-categorie",
        "Montant honoraires": "Montant honoraires (US $)",
        "Autres frais": "Autres frais (US $)",
        "Payee": "Payé",
        "Solde (US $)": "Solde",
        "Dossier envoyé": "Dossiers envoyé",
        "Dossier accepté": "Dossier approuvé",
        "Dossier refuse": "Dossier refusé",
        "Dossier annule": "Dossier Annulé",
    }
    df = df.rename(columns=rename_map)
    df = _ensure_cols(df, CLIENT_COLS)

    # Types pratiques
    # Date -> YYYY-MM-DD (string) + dérivés Année/Mois
    try:
        d = pd.to_datetime(df["Date"], errors="coerce")
    except Exception:
        d = pd.to_datetime(pd.Series([], dtype=object), errors="coerce")
    df["Date"] = d.dt.date.astype(str)
    df["_Année_"] = d.dt.year
    df["_MoisNum_"] = d.dt.month
    df["Mois"] = d.dt.month.fillna(0).astype(int).astype(str).str.zfill(2)

    # Numériques
    for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Acompte 1", "Acompte 2"]:
        df[c] = df[c].apply(_to_num)

    # Bools/0-1 pour statuts (si present)
    for c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
        if c in df.columns:
            df[c] = df[c].apply(lambda v: 1 if str(v).strip() in ("1","True","true","Oui","oui","OUI","yes","Yes") else 0)

    # ID_Client si absent
    def _make_id(nom, sdate):
        base = _safe_str(nom).strip().replace(" ", "_")
        if not base:
            base = "client"
        try:
            d = pd.to_datetime(sdate, errors="coerce")
            if pd.notna(d):
                return f"{base}-{d.strftime('%Y%m%d')}"
        except Exception:
            pass
        return f"{base}-{uuid.uuid4().hex[:6]}"

    if df["ID_Client"].isna().all():
        df["ID_Client"] = df.apply(lambda r: _make_id(r.get("Nom",""), r.get("Date","")), axis=1)

    return df

# ---------- Visa : on accepte soit un 2e fichier, soit on déduit depuis Clients ----------
def normalize_visa(df_visa: pd.DataFrame | None, df_cli: pd.DataFrame) -> pd.DataFrame:
    """
    Visa minimal : on doit au moins avoir Categories / Sous-categorie / Visa.
    Si pas de fichier Visa fourni, on reconstruit depuis Clients.
    """
    if df_visa is None or df_visa.empty:
        # reconstruit à partir des colonnes clients
        base = df_cli[["Categories", "Sous-categorie", "Visa"]].dropna(how="all").copy()
        base = base.replace({np.nan: ""})
        base = base.drop_duplicates().reset_index(drop=True)
        return base

    # Renomme si besoin
    rmap = {"Categorie": "Categories", "Catégorie": "Categories", "Sous-catégorie": "Sous-categorie"}
    df_visa = df_visa.rename(columns=rmap)
    need_cols = ["Categories", "Sous-categorie", "Visa"]
    df_visa = _ensure_cols(df_visa, need_cols)
    df_visa = df_visa[need_cols].fillna("")
    df_visa = df_visa.drop_duplicates().reset_index(drop=True)
    return df_visa

# ---------- Mémorisation des derniers chemins ----------
def _load_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def _save_state(d: dict):
    try:
        STATE_FILE.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

state = _load_state()

# ---------- Barre latérale : chargement fichiers ----------
st.sidebar.header("📂 Fichiers")

mode = st.sidebar.radio("Mode de chargement", ["Un fichier (Clients)", "Deux fichiers (Clients + Visa)"],
                        horizontal=False, key=skey("load","mode"))

up_clients = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx","xls","csv"], key=skey("up","clients"))
up_visa    = None
if mode == "Deux fichiers (Clients + Visa)":
    up_visa = st.sidebar.file_uploader("Visa (xlsx/csv)", type=["xlsx","xls","csv"], key=skey("up","visa"))

# Bouton pour recharger depuis les derniers chemins mémorisés
st.sidebar.caption("Derniers chemins mémorisés :")
st.sidebar.code(
    f"Dernier Clients : {state.get('clients_path','')}\nDernier Visa    : {state.get('visa_path','')}",
    language="text"
)

reload_last = st.sidebar.button("🔁 Recharger les derniers fichiers", key=skey("btn","reload_last"))

# Option de sauvegarde (chemin choisi par l’utilisateur)
st.sidebar.caption("**Chemin de sauvegarde** (sur ton PC / Drive / OneDrive) :")
save_clients_to = st.sidebar.text_input("Sauvegarder Clients vers…", value=state.get("clients_save_to",""),
                                        key=skey("save","clients"))
save_visa_to    = st.sidebar.text_input("Sauvegarder Visa vers…", value=state.get("visa_save_to",""),
                                        key=skey("save","visa"))

# ---------- Déterminer les sources à lire ----------
clients_src = None
visa_src = None

if reload_last:
    # recharge depuis les chemins mémorisés si dispo
    cpath = state.get("clients_path", "")
    vpath = state.get("visa_path", "")
    clients_src = cpath if cpath else None
    if mode == "Deux fichiers (Clients + Visa)":
        visa_src = vpath if vpath else None
else:
    # fichiers uploadés en session
    clients_src = up_clients if up_clients is not None else state.get("clients_path", None)
    if mode == "Deux fichiers (Clients + Visa)":
        visa_src = up_visa if up_visa is not None else state.get("visa_path", None)

# ---------- Lecture & normalisation ----------
df_clients_raw = normalize_clients(read_any_table(clients_src))
df_visa_raw    = normalize_visa(read_any_table(visa_src), df_clients_raw)

# Mémorise les chemins si on a uploadé des fichiers
def _persist_uploaded(uploaded, default_name):
    """Écrit l’upload en local pour avoir un chemin persistant et pouvoir « recharger derniers fichiers »."""
    if uploaded is None:
        return None
    data = uploaded.read() if hasattr(uploaded, "read") else None
    if data is None:
        return None
    out = WORK_DIR / f"upload_{default_name}"
    with open(out, "wb") as f:
        f.write(data)
    return str(out)

changed = False
if isinstance(up_clients, st.runtime.uploaded_file_manager.UploadedFile):
    loc = _persist_uploaded(up_clients, up_clients.name)
    if loc:
        state["clients_path"] = loc
        changed = True

if mode == "Deux fichiers (Clients + Visa)" and isinstance(up_visa, st.runtime.uploaded_file_manager.UploadedFile):
    loc = _persist_uploaded(up_visa, up_visa.name)
    if loc:
        state["visa_path"] = loc
        changed = True

# Mémorise aussi les chemins de sauvegarde si changés
if save_clients_to != state.get("clients_save_to",""):
    state["clients_save_to"] = save_clients_to
    changed = True
if save_visa_to != state.get("visa_save_to",""):
    state["visa_save_to"] = save_visa_to
    changed = True

if changed:
    _save_state(state)

# ---------- En-tête & info fichiers ----------
st.title("🛂 Visa Manager")

st.subheader("📄 Fichiers chargés")
st.write(f"**Clients** : `{state.get('clients_path', '—')}`")
if mode == "Deux fichiers (Clients + Visa)":
    st.write(f"**Visa** : `{state.get('visa_path', '—')}`")
else:
    st.write("**Visa** : (déduit depuis le fichier Clients)")

# ---------- Création des onglets (doit être AVANT tout `with tabs[i]:`) ----------
tabs = st.tabs([
    "📄 Fichiers chargés",
    "📊 Dashboard",
    "📈 Analyses",
    "🏦 Escrow",
    "👤 Compte client",
    "🧾 Gestion",
    "📄 Visa (aperçu)",
    "💾 Export"
])
# ========== FIN BLOC 1/10 ==========



# # ==============================================
# BLOC 2/10 — 📊 Dashboard (construction df_all + KPI + filtres + graphiques + détails)
# ==============================================
import pandas as pd
from datetime import datetime, date
import streamlit as st
from io import BytesIO

# ---------- Helpers locaux (autonomes pour ce bloc) ----------
def _safe_str(x):
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _to_num_col(df, col):
    if col not in df.columns:
        df[col] = 0.0
    s = df[col]
    if not pd.api.types.is_numeric_dtype(s):
        s = (
            s.astype(str)
             .str.replace(r"[^\d,.\-]", "", regex=True)
             .str.replace(",", ".", regex=False)
        )
    s = pd.to_numeric(s, errors="coerce").fillna(0.0)
    df[col] = s
    return df

def _parse_date_col(df, col):
    if col not in df.columns:
        df[col] = pd.NaT
    df[col] = pd.to_datetime(df[col], errors="coerce")
    return df

def _fmt_money(v):
    try:
        return f"${float(v):,.2f}"
    except Exception:
        return "$0.00"

def _standardize_client_columns(df):
    """Aligne les noms de colonnes ‘clients’ sur un schéma commun."""
    if df is None or df.empty:
        return pd.DataFrame()
    # Renommages tolérants (accents / variations)
    ren = {
        "Categorie":"Categorie", "Catégorie":"Categorie",
        "Sous-categorie":"Sous-categorie","Sous-catégorie":"Sous-categorie",
        "Montant honoraires (US $)":"Montant honoraires (US $)",
        "Autres frais (US $)":"Autres frais (US $)",
        "Payé":"Payé","Paye":"Payé",
        "Solde":"Solde","Reste":"Solde",
        "Date":"Date","Mois":"Mois","Visa":"Visa","Nom":"Nom",
        "Dossier N":"Dossier N","ID_Client":"ID_Client",
        "Commentaires":"Commentaires",
        "Dossier envoyé":"Dossier envoyé",
        "Dossier accepté":"Dossier accepté",
        "Dossier refusé":"Dossier refusé",
        "Dossier annulé":"Dossier annulé",
        "RFE":"RFE"
    }
    # essaie de mapper en insensible à la casse/espace
    cols_map = {}
    for c in df.columns:
        c_clean = c.strip()
        if c_clean in ren:
            cols_map[c] = ren[c_clean]
        else:
            # quelques alias probables
            low = c_clean.lower()
            if low == "categorie": cols_map[c] = "Categorie"
            elif low in ("sous-categorie","sous-catégorie","sous categorie","sous-catégorie"):
                cols_map[c] = "Sous-categorie"
            elif "honoraire" in low: cols_map[c] = "Montant honoraires (US $)"
            elif "autres frais" in low: cols_map[c] = "Autres frais (US $)"
            elif low in ("paye","payé"): cols_map[c] = "Payé"
            elif low in ("solde","reste"): cols_map[c] = "Solde"
            elif low == "visa": cols_map[c] = "Visa"
            elif low == "nom": cols_map[c] = "Nom"
            elif "dossier" in low and "n" in low: cols_map[c] = "Dossier N"
            elif "id_client" in low: cols_map[c] = "ID_Client"
            elif "comment" in low: cols_map[c] = "Commentaires"
            elif "envoy" in low: cols_map[c] = "Dossier envoyé"
            elif "accept" in low: cols_map[c] = "Dossier accepté"
            elif "refus" in low: cols_map[c] = "Dossier refusé"
            elif "annul" in low: cols_map[c] = "Dossier annulé"
            elif low == "rfe": cols_map[c] = "RFE"
            elif low == "mois": cols_map[c] = "Mois"
            elif low == "date": cols_map[c] = "Date"
            else:
                cols_map[c] = c  # conserve
    df = df.rename(columns=cols_map)

    # Colonnes minimales
    for c in ["Nom","Categorie","Sous-categorie","Visa","Date","Mois",
              "Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]:
        if c not in df.columns:
            df[c] = pd.NA

    # Typages
    df = _to_num_col(df, "Montant honoraires (US $)")
    df = _to_num_col(df, "Autres frais (US $)")
    df = _to_num_col(df, "Payé")
    # si 'Solde' absent ou incohérent, recalcule à partir du total
    total_calc = df["Montant honoraires (US $)"].fillna(0) + df["Autres frais (US $)"].fillna(0)
    if "Solde" not in df.columns or df["Solde"].isna().all():
        df["Solde"] = (total_calc - df["Payé"]).clip(lower=0)
    else:
        df = _to_num_col(df, "Solde")

    # Dates / Année / Mois
    df = _parse_date_col(df, "Date")
    # Complète Mois si absent
    if "Mois" in df.columns:
        df["Mois"] = df["Mois"].astype(str).str.extract(r"(\d{1,2})", expand=False).fillna("")
        df["Mois"] = df["Mois"].apply(lambda x: f"{int(x):02d}" if x and x.isdigit() else "")
    else:
        df["Mois"] = ""

    df["_Année_"] = df["Date"].dt.year.fillna(pd.NA)
    df["_MoisNum_"] = df["Date"].dt.month.fillna(pd.NA)
    return df

def _load_df_from_session_or_path(session_key_df, session_key_path):
    """Essaye d'abord df en session, sinon lit le chemin en session si présent."""
    df = st.session_state.get(session_key_df)
    if df is None:
        p = st.session_state.get(session_key_path)
        if p:
            try:
                if str(p).lower().endswith(".csv"):
                    df = pd.read_csv(p)
                else:
                    df = pd.read_excel(p)
            except Exception:
                df = None
    return df if isinstance(df, pd.DataFrame) else pd.DataFrame()

# ---------- Récupération sources ----------
# attendus (posés par le bloc Fichiers) :
# - st.session_state["clients_df_raw"] ou st.session_state["clients_path_curr"]
# - st.session_state["visa_df_raw"]    ou st.session_state["visa_path_curr"]

clients_raw = st.session_state.get("clients_df_raw")
if clients_raw is None or not isinstance(clients_raw, pd.DataFrame) or clients_raw.empty:
    clients_raw = _load_df_from_session_or_path("clients_df_raw", "clients_path_curr")

visa_raw = st.session_state.get("visa_df_raw")
if visa_raw is None or not isinstance(visa_raw, pd.DataFrame) or visa_raw.empty:
    # si l’app utilise le même fichier pour tout, le chemin peut être sur clients_path_curr
    visa_raw = _load_df_from_session_or_path("visa_df_raw", "visa_path_curr")
    if visa_raw.empty:
        # fallback : tente aussi le chemin clients si visa manquant
        visa_raw = _load_df_from_session_or_path("visa_df_raw", "clients_path_curr")

# Normalisation clients
df_clients = _standardize_client_columns(clients_raw.copy()) if not clients_raw.empty else pd.DataFrame()

# Pour le Dashboard on n’a pas besoin de pivot « Visa structure » ; on garde juste les colonnes visa du client.
df_all = df_clients.copy()

# ---------- Interface Dashboard ----------
st.markdown("### 📊 Dashboard")

if df_all.empty:
    st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
else:
    # Listes filtres
    cats  = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
    subs  = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
    visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

    # KPI
    tot_dossiers = len(df_all)
    total_hono   = float(_to_num_col(df_all.copy(), "Montant honoraires (US $)")["Montant honoraires (US $)"].sum())
    total_autre  = float(_to_num_col(df_all.copy(), "Autres frais (US $)")["Autres frais (US $)"].sum())
    total_all    = total_hono + total_autre
    total_paye   = float(_to_num_col(df_all.copy(), "Payé")["Payé"].sum())
    total_solde  = float(_to_num_col(df_all.copy(), "Solde")["Solde"].sum())
    # % envoyés (si colonne existe)
    if "Dossier envoyé" in df_all.columns:
        try:
            sent_ratio = (pd.to_numeric(df_all["Dossier envoyé"], errors="coerce").fillna(0) > 0).mean() * 100
        except Exception:
            sent_ratio = 0.0
    else:
        sent_ratio = 0.0

    k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
    with k1: st.metric("Dossiers", f"{tot_dossiers}")
    with k2: st.metric("Honoraires+Frais", _fmt_money(total_all))
    with k3: st.metric("Payé", _fmt_money(total_paye))
    with k4: st.metric("Solde", _fmt_money(total_solde))
    with k5: st.metric("Envoyés (%)", f"{sent_ratio:.0f}%")

    st.markdown("#### 🎛️ Filtres")
    a1, a2, a3 = st.columns(3)
    fc = a1.multiselect("Catégories", cats, default=[], key="dash_f_cats")
    fs = a2.multiselect("Sous-catégories", subs, default=[], key="dash_f_subs")
    fv = a3.multiselect("Visa", visas, default=[], key="dash_f_visas")

    view = df_all.copy()
    if fc: view = view[view["Categorie"].astype(str).isin(fc)]
    if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
    if fv: view = view[view["Visa"].astype(str).isin(fv)]

    # Graphique : dossiers par catégorie
    st.markdown("#### 📦 Nombre de dossiers par catégorie")
    if not view.empty and "Categorie" in view.columns:
        vc = view["Categorie"].astype(str).value_counts().reset_index()
        vc.columns = ["Categorie", "Nombre"]
        st.bar_chart(vc.set_index("Categorie"))
    else:
        st.caption("Aucune donnée pour ce graphique.")

    # Graphique : flux par mois (honoraires / autres / payé / solde)
    st.markdown("#### 💵 Flux par mois")
    g = view.copy()
    g = _parse_date_col(g, "Date")
    g["MoisLbl"] = g["Date"].dt.to_period("M").astype(str)
    g = _to_num_col(g, "Montant honoraires (US $)")
    g = _to_num_col(g, "Autres frais (US $)")
    g = _to_num_col(g, "Payé")
    g = _to_num_col(g, "Solde")
    if not g.empty:
        gb = (g.groupby("MoisLbl", as_index=False)
                [["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]]
                .sum()
                .sort_values("MoisLbl"))
        st.line_chart(gb.set_index("MoisLbl"))
    else:
        st.caption("Aucune donnée pour ce graphique.")

    # Tableau détaillé
    st.markdown("#### 📋 Détails (après filtres)")
    det = view.copy()
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]:
        if c in det.columns:
            det[c] = pd.to_numeric(det[c], errors="coerce").fillna(0.0).apply(_fmt_money)
    if "Date" in det.columns:
        try:
            det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
        except Exception:
            det["Date"] = det["Date"].astype(str)

    show_cols = [c for c in [
        "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa",
        "Date","Mois","Montant honoraires (US $)","Autres frais (US $)","Payé","Solde",
        "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE","Commentaires"
    ] if c in det.columns]

    sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
    det_sorted = det.sort_values(by=sort_keys) if sort_keys else det
    st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key="dash_detail_table")




# ================================
# PARTIE 3/6 — 📊 Dashboard
# ================================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # KPIs
        left, right = st.columns([1.2, 2.8])
        with left:
            k1, k2 = st.columns([1,1])
            k3, k4 = st.columns([1,1])
            k1.metric("Dossiers", f"{len(df_all)}")
            k2.metric("Honoraires+Frais", _fmt_money((_series_num(df_all,"Montant honoraires (US $)") + _series_num(df_all,"Autres frais (US $)")).sum()))
            k3.metric("Payé", _fmt_money(_series_num(df_all, "Payé").sum()))
            k4.metric("Solde", _fmt_money(_series_num(df_all, "Reste").sum()))
            # % envoyés
            pct_env = 0.0
            if len(df_all) > 0:
                pct_env = 100.0 * (_series_num(df_all, "Dossier envoyé")>0).sum() / len(df_all)
            st.metric("Envoyés (%)", f"{pct_env:0.0f}%")

        with right:
            st.markdown("#### 🎛️ Filtres")
            cats  = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
            subs  = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
            visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

            a1, a2, a3 = st.columns(3)
            fc = a1.multiselect("Catégories", cats, default=[], key=f"dash_c_{SID}")
            fs = a2.multiselect("Sous-catégories", subs, default=[], key=f"dash_s_{SID}")
            fv = a3.multiselect("Visa", visas, default=[], key=f"dash_v_{SID}")

        view = df_all.copy()
        if fc: view = view[view["Categorie"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if not view.empty and "Categorie" in view.columns:
            vc = view["Categorie"].value_counts().sort_index()
            st.bar_chart(vc)

        st.markdown("#### 💵 Flux par mois")
        flux = pd.DataFrame({
            "Mois": view["Mois"].astype(str),
            "Montant honoraires (US $)": _series_num(view, "Montant honoraires (US $)"),
            "Autres frais (US $)": _series_num(view, "Autres frais (US $)"),
            "Payé": _series_num(view, "Payé"),
            "Solde": _series_num(view, "Reste")
        })
        flux = flux.groupby("Mois", as_index=False).sum().sort_values("Mois")
        st.line_chart(flux.set_index("Mois"))

        st.markdown("#### 📋 Détails (après filtres)")
        det = view.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in det.columns:
                det[c] = _series_num(det, c).map(_fmt_money)
        if "Date" in det.columns:
            det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]
        sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
        det = det.sort_values(by=sort_keys) if sort_keys else det
        st.dataframe(det[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_table_{SID}")




# ================================
# PARTIE 4/6 — 📈 Analyses / 🏦 Escrow / 📄 Visa (aperçu)
# ================================

# -------- Analyses --------
with tabs[1]:
    st.subheader("📈 Analyses")
    if df_all.empty:
        st.info("Aucune donnée.")
    else:
        yearsA  = sorted(df_all["_Année_"].dropna().astype(int).unique().tolist())
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist())
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist())
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist())

        a1,a2,a3,a4,a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_y_{SID}")
        fm = a2.multiselect("Mois", monthsA, default=[], key=f"a_m_{SID}")
        fc = a3.multiselect("Catégories", catsA, default=[], key=f"a_c_{SID}")
        fs = a4.multiselect("Sous-catégories", subsA, default=[], key=f"a_s_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_v_{SID}")

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(_series_num(dfA, "Montant honoraires (US $)").sum()))
        k3.metric("Payé", _fmt_money(_series_num(dfA, "Payé").sum()))
        k4.metric("Reste", _fmt_money(_series_num(dfA, "Reste").sum()))

        st.markdown("#### 📦 Répartition par catégorie (en %)")
        if not dfA.empty:
            vc = dfA["Categorie"].value_counts(dropna=True)
            pct = (vc / vc.sum() * 100.0).round(1)
            st.bar_chart(pct.sort_index())

        st.markdown("#### 🧩 Répartition par sous-catégorie (en %)")
        if not dfA.empty:
            vs = dfA["Sous-categorie"].value_counts(dropna=True)
            pct2 = (vs / vs.sum() * 100.0).round(1)
            st.bar_chart(pct2.sort_index())

        st.markdown("#### 📈 Honoraires par mois")
        tmp = dfA.copy()
        g = tmp.groupby("Mois", as_index=False)["Montant honoraires (US $)"].sum().sort_values("Mois")
        st.line_chart(g.set_index("Mois"))

        st.markdown("#### 🧾 Détails")
        det = dfA.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in det.columns:
                det[c] = _series_num(det, c).map(_fmt_money)
        if "Date" in det.columns:
            det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)

        cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]
        st.dataframe(det[cols].reset_index(drop=True), use_container_width=True, key=f"a_table_{SID}")

# -------- Escrow --------
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Total (US $)"] = _series_num(dfE, "Total (US $)")
        dfE["Payé"] = _series_num(dfE, "Payé")
        dfE["Reste"] = _series_num(dfE, "Reste")

        t1,t2,t3 = st.columns(3)
        t1.metric("Total", _fmt_money(dfE["Total (US $)"].sum()))
        t2.metric("Payé", _fmt_money(dfE["Payé"].sum()))
        t3.metric("Reste", _fmt_money(dfE["Reste"].sum()))

        st.markdown("#### Par catégorie")
        agg = dfE.groupby("Categorie", as_index=False)[["Total (US $)","Payé","Reste"]].sum()
        st.dataframe(agg, use_container_width=True, key=f"esc_agg_{SID}")

        st.caption("NB : pour un suivi ESCROW strict, on peut isoler les honoraires pré-envoi et déclencher un transfert quand le statut passe à « Envoyé ».")

# -------- Visa (aperçu) --------
with tabs[5]:
    st.subheader("📄 Visa — aperçu & filtres")
    if df_visa_raw.empty:
        st.info("Aucun fichier Visa.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True, key=f"visa_raw_{SID}")
        st.markdown("#### Carte Catégorie → Sous-catégorie → Options disponibles")
        cats = sorted(list(visa_map.keys()))
        c1, c2 = st.columns(2)
        selc = c1.selectbox("Catégorie", [""]+cats, index=0, key=f"v_cat_{SID}")
        if selc:
            subs = sorted(list(visa_map.get(selc, {}).keys()))
            sels = c2.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"v_sub_{SID}")
            if sels:
                opts = visa_map.get(selc,{}).get(sels,{}).get("options",[])
                st.write("**Options** :", ", ".join(opts) if opts else "Aucune (visa direct)")




# ================================
# PARTIE 5/6 — 👤 Compte client (timeline + paiements)
# ================================
with tabs[3]:
    st.subheader("👤 Compte client")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        names = sorted(df_all["Nom"].dropna().astype(str).unique().tolist())
        ids   = sorted(df_all["ID_Client"].dropna().astype(str).unique().tolist())
        c1,c2 = st.columns(2)
        pick_name = c1.selectbox("Nom", [""]+names, index=0, key=f"acc_nom_{SID}")
        pick_id   = c2.selectbox("ID_Client", [""]+ids, index=0, key=f"acc_id_{SID}")

        mask = None
        if pick_id:
            mask = (df_all["ID_Client"].astype(str) == pick_id)
        elif pick_name:
            mask = (df_all["Nom"].astype(str) == pick_name)

        if mask is not None and mask.any():
            row = df_all[mask].iloc[0].copy()

            st.markdown("#### 📌 Dossier")
            s1,s2,s3,s4 = st.columns(4)
            s1.write(f"Dossier N : {_safe_str(row.get('Dossier N',''))}")
            s2.write(f"Nom : {_safe_str(row.get('Nom',''))}")
            s3.write(f"Visa : {_safe_str(row.get('Visa',''))}")
            s4.write(f"Catégorie : {_safe_str(row.get('Categorie',''))} / {_safe_str(row.get('Sous-categorie',''))}")

            st.markdown("#### ✅ Statut & dates")
            env = int(_to_num(row.get("Dossier envoyé",0)))==1
            acc = int(_to_num(row.get("Dossier accepté",0)))==1
            ref = int(_to_num(row.get("Dossier refusé",0)))==1
            ann = int(_to_num(row.get("Dossier annulé",0)))==1
            rfe = int(_to_num(row.get("RFE",0)))==1

            colA, colB = st.columns(2)
            with colA:
                st.write("- Dossier envoyé :", "Oui" if env else "Non",
                         "| Date :", _safe_str(row.get("Date d'envoi","")))
                st.write("- Dossier accepté :", "Oui" if acc else "Non",
                         "| Date :", _safe_str(row.get("Date d'acceptation","")))
                st.write("- Dossier refusé :", "Oui" if ref else "Non",
                         "| Date :", _safe_str(row.get("Date de refus","")))
                st.write("- Dossier annulé :", "Oui" if ann else "Non",
                         "| Date :", _safe_str(row.get("Date d'annulation","")))
            with colB:
                st.write("- RFE :", "Oui" if rfe else "Non")

            st.markdown("#### 💳 Paiements")
            # Paiements stockés en JSON ou liste
            rawp = row.get("Paiements","")
            pay_list = []
            if isinstance(rawp, list):
                pay_list = rawp
            else:
                try:
                    pay_list = json.loads(_safe_str(rawp) or "[]")
                    if not isinstance(pay_list, list): pay_list = []
                except Exception:
                    pay_list = []

            if pay_list:
                dfp = pd.DataFrame(pay_list)
                if "date" in dfp.columns:
                    try:
                        dfp["date"] = pd.to_datetime(dfp["date"], errors="coerce").dt.date.astype(str)
                    except Exception:
                        pass
                st.dataframe(dfp, use_container_width=True, key=f"pay_hist_{SID}")
            else:
                st.info("Aucun paiement saisi.")

            st.markdown("##### ➕ Ajouter un paiement (tant que le dossier n’est pas soldé)")
            reste = float(_to_num(row.get("Reste", 0.0)))
            if reste <= 0:
                st.success("Ce dossier est soldé.")
            else:
                pcol1,pcol2,pcol3,pcol4 = st.columns([1,1,1,2])
                pdate = pcol1.date_input("Date", value=date.today(), key=f"pay_date_{SID}")
                pamt  = pcol2.number_input("Montant", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=f"pay_amt_{SID}")
                pmode = pcol3.selectbox("Mode", ["Chèque","CB","Cash","Virement","Venmo"], index=1, key=f"pay_mode_{SID}")
                pok   = pcol4.button("💾 Enregistrer le paiement", key=f"pay_save_{SID}")

                if pok:
                    add = float(pamt or 0.0)
                    if add <= 0:
                        st.warning("Montant > 0 requis.")
                    else:
                        # MAJ paiements + Payé + Reste
                        pay_list.append({
                            "date": pdate.strftime("%Y-%m-%d"),
                            "montant": add,
                            "mode": pmode
                        })
                        # Recalcule
                        paye_new = float(_to_num(row.get("Payé", 0.0))) + add
                        total    = float(_to_num(row.get("Total (US $)", 0.0)))
                        reste_new= max(0.0, total - paye_new)

                        # Persister dans df_all puis fichier source
                        idx_global = df_all[mask].index[0]
                        df_all.at[idx_global, "Paiements"] = json.dumps(pay_list, ensure_ascii=False)
                        df_all.at[idx_global, "Payé"] = paye_new
                        df_all.at[idx_global, "Reste"] = reste_new

                        # Écrire dans fichier clients
                        try:
                            write_clients_file(df_all, clients_src if isinstance(clients_src,str) else save_clients_to or "clients_sauvegarde.xlsx")
                            st.success("Paiement ajouté.")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur sauvegarde : {_safe_str(e)}")




# ================================
# PARTIE 6/6 — 🧾 Gestion (CRUD) + 💾 Export
# ================================
with tabs[4]:
    st.subheader("🧾 Gestion (Ajouter / Modifier / Supprimer)")

    # Helpers statut
    def _status_to_flags(status: str):
        s = (status or "").strip().lower()
        return {
            "Dossier envoyé":  1 if s=="envoyé" else 0,
            "Dossier accepté": 1 if s=="accepté" else 0,
            "Dossier refusé":  1 if s=="refusé" else 0,
            "Dossier annulé":  1 if s=="annulé" else 0,
        }
    def _flags_to_status(row):
        if int(_to_num(row.get("Dossier accepté",0)))==1: return "Accepté"
        if int(_to_num(row.get("Dossier refusé",0)))==1:  return "Refusé"
        if int(_to_num(row.get("Dossier annulé",0)))==1:  return "Annulé"
        if int(_to_num(row.get("Dossier envoyé",0)))==1:  return "Envoyé"
        return "Aucun"
    def _status_date_key(statut):
        lut = {"Envoyé":"Date d'envoi","Accepté":"Date d'acceptation","Refusé":"Date de refus","Annulé":"Date d'annulation"}
        return lut.get(statut, None)

    df_live = df_all.copy()

    op = st.radio("Action", ["Ajouter","Modifier","Supprimer"], horizontal=True, key=f"crud_op_{SID}")

    # ------- AJOUT -------
    if op == "Ajouter":
        c1,c2,c3 = st.columns(3)
        nom = c1.text_input("Nom", "", key=f"add_nom_{SID}")
        dt  = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois= c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=(date.today().month-1), key=f"add_mois_{SID}")

        st.markdown("#### 🎯 Visa")
        cats = sorted(list(visa_map.keys()))
        cat  = st.selectbox("Catégorie", [""]+cats, index=0, key=f"add_cat_{SID}")
        sub  = ""
        visa_final, opts_dict = "", {"exclusive": None, "options":[]}
        if cat:
            subs = sorted(list(visa_map.get(cat,{}).keys()))
            sub  = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"add_sub_{SID}")
            if sub:
                opts = visa_map.get(cat,{}).get(sub,{}).get("options",[])
                st.caption("Options (issues du fichier Visa) : " + (", ".join(opts) if opts else "aucune"))

        f1,f2 = st.columns(2)
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
        other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")
        comment = st.text_area("Commentaires (Autres frais / notes)", "", key=f"add_comm_{SID}")

        st.markdown("#### 📌 Statut & RFE")
        st_choices = ["Aucun","Envoyé","Accepté","Refusé","Annulé"]
        statut = st.selectbox("Statut", st_choices, index=0, key=f"add_stat_{SID}")
        rfe_on = st.toggle("RFE", value=False, key=f"add_rfe_{SID}")
        if rfe_on and statut=="Aucun":
            st.warning("RFE nécessite un statut sélectionné.")

        dkey = _status_date_key(statut)
        stat_date = None
        if statut!="Aucun":
            stat_date = st.date_input(f"Date pour « {statut} »", value=date.today(), key=f"add_statd_{SID}")

        if st.button("💾 Enregistrer le client", key=f"btn_add_{SID}"):
            if not nom or not cat or not sub:
                st.warning("Nom, Catégorie, Sous-catégorie requis.")
            else:
                did = f"{_safe_str(nom).strip()}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                dossier_n = int(df_live["Dossier N"].max())+1 if "Dossier N" in df_live.columns and not df_live.empty else 13057
                total = float(honor)+float(other)
                row = {
                    "Dossier N": dossier_n, "ID_Client": did, "Nom": nom,
                    "Date": dt, "Mois": mois,
                    "Categorie": cat, "Sous-categorie": sub, "Visa": sub,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Total (US $)": total,
                    "Payé": 0.0, "Reste": total,
                    "Paiements": json.dumps([], ensure_ascii=False),
                    "Commentaires": comment,
                    "Dossier envoyé":0, "Dossier accepté":0, "Dossier refusé":0, "Dossier annulé":0,
                    "Date d'envoi": None, "Date d'acceptation": None, "Date de refus": None, "Date d'annulation": None,
                    "RFE": 1 if (rfe_on and statut!="Aucun") else 0
                }
                flags = _status_to_flags(statut)
                for k,v in flags.items(): row[k]=v
                if dkey: row[dkey]=stat_date
                df_live = pd.concat([df_live, pd.DataFrame([row])], ignore_index=True)
                write_clients_file(df_live, clients_src if isinstance(clients_src,str) else (save_clients_to or "clients_sauvegarde.xlsx"))
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()

    # ------- MODIFIER -------
    elif op == "Modifier":
        if df_live.empty:
            st.info("Aucun client à modifier.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
            m1,m2 = st.columns(2)
            tgt_name = m1.selectbox("Nom", [""]+names, index=0, key=f"mod_nom_{SID}")
            tgt_id   = m2.selectbox("ID_Client", [""]+ids, index=0, key=f"mod_id_{SID}")

            mask=None
            if tgt_id: mask = (df_live["ID_Client"].astype(str)==tgt_id)
            elif tgt_name: mask = (df_live["Nom"].astype(str)==tgt_name)

            if mask is not None and mask.any():
                idx = df_live[mask].index[0]
                row = df_live.loc[idx].copy()

                d1,d2,d3 = st.columns(3)
                nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nv_{SID}")
                dt   = d2.date_input("Date de création", value=_date_for_widget(row.get("Date")), key=f"mod_dt_{SID}")
                mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                    index=max(0, int(_safe_str(row.get("Mois","01"))[:2]) - 1), key=f"mod_m_{SID}")

                # Visa cascade
                st.markdown("#### 🎯 Visa")
                cats = sorted(list(visa_map.keys()))
                preset_cat = _safe_str(row.get("Categorie",""))
                cat  = st.selectbox("Catégorie", [""]+cats, index=(cats.index(preset_cat)+1 if preset_cat in cats else 0), key=f"mod_cat_{SID}")
                sub  = _safe_str(row.get("Sous-categorie",""))
                if cat:
                    subs = sorted(list(visa_map.get(cat,{}).keys()))
                    sub  = st.selectbox("Sous-catégorie", [""]+subs, index=(subs.index(sub)+1 if sub in subs else 0), key=f"mod_sub_{SID}")

                f1,f2 = st.columns(2)
                honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=float(_to_num(row.get("Montant honoraires (US $)",0.0))), step=50.0, format="%.2f", key=f"mod_h_{SID}")
                other = f2.number_input("Autres frais (US $)", min_value=0.0, value=float(_to_num(row.get("Autres frais (US $)",0.0))), step=20.0, format="%.2f", key=f"mod_o_{SID}")
                comment = st.text_area("Commentaires", _safe_str(row.get("Commentaires","")), key=f"mod_com_{SID}")

                st.markdown("#### 📌 Statut & RFE")
                st_choices = ["Aucun","Envoyé","Accepté","Refusé","Annulé"]
                current = "Aucun"
                if int(_to_num(row.get("Dossier accepté",0)))==1: current="Accepté"
                elif int(_to_num(row.get("Dossier refusé",0)))==1: current="Refusé"
                elif int(_to_num(row.get("Dossier annulé",0)))==1: current="Annulé"
                elif int(_to_num(row.get("Dossier envoyé",0)))==1: current="Envoyé"
                statut = st.selectbox("Statut", st_choices, index=st_choices.index(current), key=f"mod_stat_{SID}")
                rfe_on = st.toggle("RFE", value=(int(_to_num(row.get("RFE",0)))==1), key=f"mod_rfe_{SID}")

                dkey = _status_date_key(statut)
                stat_date = _date_for_widget(row.get(dkey)) if dkey else date.today()
                if statut!="Aucun" and dkey:
                    stat_date = st.date_input(f"Date pour « {statut} »", value=_date_for_widget(row.get(dkey)), key=f"mod_statd_{SID}")

                if st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}"):
                    if not nom or not cat or not sub:
                        st.warning("Nom, Catégorie, Sous-catégorie requis.")
                    else:
                        total = float(honor)+float(other)
                        paye  = float(_to_num(row.get("Payé",0.0)))
                        reste = max(0.0, total - paye)

                        df_live.at[idx,"Nom"]=nom
                        df_live.at[idx,"Date"]=dt
                        df_live.at[idx,"Mois"]=f"{int(mois):02d}"
                        df_live.at[idx,"Categorie"]=cat
                        df_live.at[idx,"Sous-categorie"]=sub
                        df_live.at[idx,"Visa"]=sub
                        df_live.at[idx,"Montant honoraires (US $)"]=float(honor)
                        df_live.at[idx,"Autres frais (US $)"]=float(other)
                        df_live.at[idx,"Total (US $)"]=total
                        df_live.at[idx,"Reste"]=reste
                        df_live.at[idx,"Commentaires"]=comment

                        # reset statuts + dates
                        for k in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé"]:
                            df_live.at[idx,k]=0
                        for k in ["Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation"]:
                            df_live.at[idx,k]=None
                        flags=_status_to_flags(statut)
                        for k,v in flags.items(): df_live.at[idx,k]=v
                        if statut!="Aucun" and dkey:
                            df_live.at[idx,dkey]=stat_date
                        df_live.at[idx,"RFE"]=1 if (rfe_on and statut!="Aucun") else 0

                        write_clients_file(df_live, clients_src if isinstance(clients_src,str) else (save_clients_to or "clients_sauvegarde.xlsx"))
                        st.success("Modifications enregistrées.")
                        st.cache_data.clear()
                        st.rerun()

    # ------- SUPPRIMER -------
    elif op == "Supprimer":
        if df_live.empty:
            st.info("Aucun client.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
            s1,s2=st.columns(2)
            tgt_name = s1.selectbox("Nom", [""]+names, index=0, key=f"del_nom_{SID}")
            tgt_id   = s2.selectbox("ID_Client", [""]+ids, index=0, key=f"del_id_{SID}")

            mask=None
            if tgt_id: mask=(df_live["ID_Client"].astype(str)==tgt_id)
            elif tgt_name: mask=(df_live["Nom"].astype(str)==tgt_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({"Dossier N":row.get("Dossier N",""), "Nom":row.get("Nom",""), "Visa":row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                    df_live = df_live[~mask].copy()
                    write_clients_file(df_live, clients_src if isinstance(clients_src,str) else (save_clients_to or "clients_sauvegarde.xlsx"))
                    st.success("Supprimé.")
                    st.cache_data.clear()
                    st.rerun()

# -------- Export --------
with tabs[6]:
    st.subheader("💾 Export")
    colz1, colz2 = st.columns([1,3])
    with colz1:
        if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID}"):
            try:
                buf = BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    # Clients
                    with BytesIO() as xbuf:
                        with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                            df_all.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                        zf.writestr("Clients.xlsx", xbuf.getvalue())
                    # Visa si présent
                    if visa_src:
                        try:
                            if isinstance(visa_src, str) and os.path.exists(visa_src):
                                zf.write(visa_src, "Visa.xlsx")
                            else:
                                # upload → re-écrit depuis df_visa_raw
                                with BytesIO() as vb:
                                    with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                        df_visa_raw.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                                    zf.writestr("Visa.xlsx", vb.getvalue())
                        except Exception:
                            pass
                st.session_state[f"zip_export_{SID}"] = buf.getvalue()
                st.success("Archive prête.")
            except Exception as e:
                st.error(f"Erreur : {_safe_str(e)}")
    with colz2:
        if st.session_state.get(f"zip_export_{SID}"):
            st.download_button(
                "⬇️ Télécharger l’export (ZIP)",
                data=st.session_state[f"zip_export_{SID}"],
                file_name="Export_Visa_Manager.zip",
                mime="application/zip",
                key=f"zip_dl_{SID}",
            )




# ==============================
# BLOC 2/10 — Sidebar fichiers, lecture & onglets
# ==============================

st.set_page_config(page_title="Visa Manager", layout="wide")

# --- Mémoire des derniers chemins
lp_clients, lp_visa, lp_save = load_last_paths()
st.session_state.setdefault(f"last_clients_{SID}", lp_clients)
st.session_state.setdefault(f"last_visa_{SID}", lp_visa)
st.session_state.setdefault(f"last_save_dir_{SID}", lp_save)

# ============ BARRE LATÉRALE ============ #
with st.sidebar:
    st.header("## 📂 Fichiers")

    mode = st.radio(
        "Mode de chargement",
        ["Un fichier (Clients)", "Deux fichiers (Clients + Visa)"],
        index=1,
        key=f"mode_{SID}",
    )

    up_clients = st.file_uploader("Clients (xlsx/csv)", type=["xlsx","xls","csv"], key=f"up_clients_{SID}")
    up_visa    = None
    if mode == "Deux fichiers (Clients + Visa)":
        up_visa = st.file_uploader("Visa (xlsx/csv)", type=["xlsx","xls","csv"], key=f"up_visa_{SID}")

    st.caption("—")

    st.markdown("**Derniers chemins mémorisés :**")
    st.write("Dernier Clients :", f"`{st.session_state[f'last_clients_{SID}'] or ''}`")
    st.write("Dernier Visa :", f"`{st.session_state[f'last_visa_{SID}'] or ''}`")

    st.caption("**Chemin de sauvegarde** (sur ton PC / Drive / OneDrive) :")
    save_dir = st.text_input(
        "Dossier par défaut pour sauvegarder",
        value=st.session_state[f"last_save_dir_{SID}"],
        key=f"save_dir_{SID}",
        placeholder="ex: C:/Users/…/Documents/VisaManager",
    )

    st.caption("—")

    # Champs chemins manuels (optionnels)
    path_clients = st.text_input(
        "Chemin Clients (xlsx/csv)",
        value=st.session_state[f"last_clients_{SID}"],
        key=f"path_clients_{SID}",
        placeholder="ex: C:/…/donnees_visa_clients1.xlsx",
    )
    if mode == "Deux fichiers (Clients + Visa)":
        path_visa = st.text_input(
            "Chemin Visa (xlsx/csv)",
            value=st.session_state[f"last_visa_{SID}"],
            key=f"path_visa_{SID}",
            placeholder="ex: C:/…/donnees_visa_clients1.xlsx",
        )
    else:
        path_visa = ""

    col_sb1, col_sb2 = st.columns(2)
    with col_sb1:
        if st.button("💾 Mémoriser ces chemins", key=f"btn_mem_{SID}"):
            st.session_state[f"last_clients_{SID}"]  = path_clients
            st.session_state[f"last_visa_{SID}"]     = path_visa
            st.session_state[f"last_save_dir_{SID}"] = save_dir
            save_last_paths(path_clients, path_visa, save_dir)
            st.success("Chemins mémorisés.")
    with col_sb2:
        if st.button("↩️ Restaurer derniers choix", key=f"btn_res_{SID}"):
            lp_clients, lp_visa, lp_save = load_last_paths()
            st.session_state[f"last_clients_{SID}"]  = lp_clients
            st.session_state[f"last_visa_{SID}"]     = lp_visa
            st.session_state[f"last_save_dir_{SID}"] = lp_save
            st.experimental_rerun()

# --- Sources : uploader prioritaire, sinon chemin saisi, sinon dernier chemin
clients_src = up_clients if up_clients is not None else (
    st.session_state[f"path_clients_{SID}"] or st.session_state[f"last_clients_{SID}"]
)
if mode == "Deux fichiers (Clients + Visa)":
    visa_src = up_visa if up_visa is not None else (
        st.session_state[f"path_visa_{SID}"] or st.session_state[f"last_visa_{SID}"]
    )
else:
    # un seul fichier = on utilise le même pour Visa si la feuille existe (ou autre structure)
    visa_src = clients_src

# --- Lecture & normalisation
df_clients_raw = normalize_clients(read_any_table(clients_src))
df_visa_raw    = normalize_visa(read_any_table(visa_src))

# --- Map des visas -> {Categorie: {Sous-categorie: {"exclusive": None, "options":[…]}}}
visa_map = build_visa_map(df_visa_raw)

# --- Affichage résumé des fichiers chargés
st.title("🛂 Visa Manager")

tabs = st.tabs([
    "📄 Fichiers chargés",
    "📊 Dashboard",
    "🏦 Escrow",
    "👤 Compte client",
    "🧾 Gestion",
    "📄 Visa (aperçu)",
    "💾 Export",
    "📈 Analyses",
])

with tabs[0]:
    st.markdown("### 📄 Fichiers chargés")
    cli_label = getattr(up_clients, "name", None) or (clients_src if isinstance(clients_src, str) else "")
    vis_label = getattr(up_visa, "name", None) or (visa_src if isinstance(visa_src, str) else "")
    st.write("**Clients** :", f"`{cli_label or '(aucun)'}`")
    st.write("**Visa** :",    f"`{vis_label or '(aucun)'}`")

    st.caption("Astuces :")
    st.markdown("- Utilise l’uploader **ou** saisis un chemin absolu vers le fichier.")
    st.markdown("- Le bouton **Mémoriser ces chemins** enregistre tes derniers choix pour la prochaine session.")
    st.markdown("- Le mode *Un fichier* utilise le même fichier pour Clients et Visa (si deux onglets).")

# (Fin du BLOC 2/10)




# ==============================
# BLOC 3/10 — Dashboard
# ==============================

with tabs[1]:
    st.markdown("### 📊 Dashboard")

    df_all = df_clients_raw.copy()
    if df_all is None or df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # --- Colonnes attendues & conversions sûres
        num_cols = [
            "Montant honoraires (US $)",
            "Autres frais (US $)",
            "Payé",  # somme des règlements
            "Solde",
        ]
        for c in num_cols:
            if c not in df_all.columns:
                df_all[c] = 0.0
            df_all[c] = _to_num(df_all[c])

        # Total (Honoraires + Autres frais)
        if "Total (US $)" not in df_all.columns:
            df_all["Total (US $)"] = df_all["Montant honoraires (US $)"] + df_all["Autres frais (US $)"]
        else:
            df_all["Total (US $)"] = _to_num(df_all["Total (US $)"])

        # Solde si absent / incohérent
        df_all["Solde"] = df_all["Solde"]
        need_s = df_all["Solde"].isna() | (df_all["Solde"] < 0)
        df_all.loc[need_s, "Solde"] = (df_all["Total (US $)"] - df_all["Payé"]).clip(lower=0)

        # Dérivés Date -> Année / Mois (MM)
        if "Date" in df_all.columns:
            dd = pd.to_datetime(df_all["Date"], errors="coerce")
            df_all["_Année_"]   = dd.dt.year.fillna(0).astype(int)
            df_all["_MoisNum_"] = dd.dt.month.fillna(0).astype(int)
            df_all["Mois"]      = dd.dt.month.fillna(0).astype(int).map(lambda m: f"{m:02d}" if m else "")
        else:
            df_all["_Année_"] = 0
            df_all["_MoisNum_"] = 0
            df_all["Mois"] = ""

        # Normalisation noms colonnes clés (compat avec anciens fichiers)
        if "Categorie" not in df_all.columns and "Categories" in df_all.columns:
            df_all = df_all.rename(columns={"Categories": "Categorie"})
        if "Sous-categorie" not in df_all.columns and "Sous-categories" in df_all.columns:
            df_all = df_all.rename(columns={"Sous-categories": "Sous-categorie"})

        # ------- KPI compacts
        cK1, cK2, cK3, cK4, cK5 = st.columns([1,1,1,1,1])
        cK1.metric("Dossiers", f"{len(df_all)}")
        cK2.metric("Honoraires+Frais", _fmt_money(df_all["Total (US $)"].sum()))
        cK3.metric("Payé", _fmt_money(df_all["Payé"].sum()))
        cK4.metric("Solde", _fmt_money(df_all["Solde"].sum()))

        # % envoyés
        sent_col = None
        for cand in ["Dossier envoyé", "Dossiers envoyé", "Dossiers envoyés"]:
            if cand in df_all.columns:
                sent_col = cand
                break
        if sent_col:
            sent_rate = ( (_to_num(df_all[sent_col]) > 0).sum() / max(1, len(df_all)) ) * 100
        else:
            sent_rate = 0.0
        cK5.metric("Envoyés (%)", f"{sent_rate:.0f}%")

        st.divider()

        # ------- Filtres
        years   = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist() if y > 0])
        months  = [f"{m:02d}" for m in range(1,13)]
        cats    = sorted(df_all.get("Categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        subs    = sorted(df_all.get("Sous-categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        visas   = sorted(df_all.get("Visa", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", years, default=[], key=f"dash_years_{SID}")
        fm = a2.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{SID}")
        fc = a3.multiselect("Catégories", cats, default=[], key=f"dash_cats_{SID}")
        fs = a4.multiselect("Sous-catégories", subs, default=[], key=f"dash_subs_{SID}")
        fv = a5.multiselect("Visa", visas, default=[], key=f"dash_visas_{SID}")

        # Appliquer filtres
        view = df_all.copy()
        if fy: view = view[view["_Année_"].isin(fy)]
        if fm: view = view[view["Mois"].astype(str).isin(fm)]
        if fc: view = view[view["Categorie"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        st.caption(f"**Résultats après filtres : {len(view)} dossier(s)**")

        # ------- Graphique 1 : Nombre par catégorie
        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if not view.empty and "Categorie" in view.columns:
            vc = view["Categorie"].value_counts().sort_index()
            st.bar_chart(vc, key=f"bar_cat_{SID}")
        else:
            st.info("Pas de données de catégories disponibles.")

        # ------- Graphique 2 : Flux par mois (Honoraires, Frais, Payé, Solde)
        st.markdown("#### 💵 Flux par mois")
        if not view.empty and "Mois" in view.columns and view["Mois"].ne("").any():
            g = view.copy()
            g["Mois"] = g["Mois"].replace("", np.nan)
            g = g.dropna(subset=["Mois"])
            grp = (
                g.groupby("Mois", as_index=True)[
                    ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"]
                ].sum()
                .reindex([f"{m:02d}" for m in range(1,13)])  # ordre 01..12
                .fillna(0.0)
            )
            st.line_chart(grp, key=f"line_flux_{SID}")
        else:
            st.info("Pas assez de données mensuelles pour tracer les flux.")

        # ------- Détails tableau
        st.markdown("#### 📋 Détails (après filtres)")
        detail = view.copy()

        # Formattage affichage monétaire
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Total (US $)", "Payé", "Solde"]:
            if c in detail.columns:
                detail[c] = _to_num(detail[c]).map(_fmt_money)

        # Date -> str propre
        if "Date" in detail.columns:
            try:
                detail["Date"] = pd.to_datetime(detail["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                detail["Date"] = detail["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Date","Mois","Categorie","Sous-categorie","Visa",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé","RFE","Commentaires"
        ] if c in detail.columns]

        # Tri lisible
        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in detail.columns]
        detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail

        # Eviter l’erreur pyarrow si noms dupliqués
        detail_sorted = detail_sorted.loc[:, ~detail_sorted.columns.duplicated()].copy()

        st.dataframe(detail_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_tbl_{SID}")




# ==============================
# BLOC 4/10 — Analyses
# ==============================

with tabs[7]:
    st.markdown("### 📈 Analyses")

    dfA = df_clients_raw.copy()
    if dfA.empty:
        st.warning("Aucune donnée client chargée.")
    else:
        # Sélection des périodes A et B (Année/Mois)
        dfA["_Année_"]   = pd.to_numeric(dfA.get("_Année_", np.nan), errors="coerce")
        dfA["_MoisNum_"] = pd.to_numeric(dfA.get("_MoisNum_", np.nan), errors="coerce")
        dfA["Mois"]      = dfA["_MoisNum_"].fillna(0).astype(int).map(lambda m: f"{m:02d}" if m else "")

        yearsA  = sorted(dfA["_Année_"].dropna().astype(int).unique().tolist())
        monthsA = [f"{m:02d}" for m in range(1, 13)]

        ca1, ca2 = st.columns(2)
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=f"cmp_ya_{SID}")
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=f"cmp_ma_{SID}")

        cb1, cb2 = st.columns(2)
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=f"cmp_yb_{SID}")
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=f"cmp_mb_{SID}")

        dfA["Total"] = _to_num(dfA["Montant honoraires (US $)"]) + _to_num(dfA["Autres frais (US $)"])
        dfA["Categorie"] = dfA.get("Categories", dfA.get("Categorie", ""))
        dfA["Sous-categorie"] = dfA.get("Sous-categorie", "")

        # Filtres Périodes
        A = dfA.copy()
        B = dfA.copy()
        if pa_years: A = A[A["_Année_"].isin(pa_years)]
        if pa_month: A = A[A["Mois"].astype(str).isin(pa_month)]
        if pb_years: B = B[B["_Année_"].isin(pb_years)]
        if pb_month: B = B[B["Mois"].astype(str).isin(pb_month)]

        st.divider()

        cA, cB = st.columns(2)
        cA.metric("Total A", _fmt_money(A["Total"].sum()), help="Somme des honoraires + frais pour période A")
        cB.metric("Total B", _fmt_money(B["Total"].sum()), help="Somme des honoraires + frais pour période B")

        diff_val = B["Total"].sum() - A["Total"].sum()
        diff_pct = (diff_val / A["Total"].sum() * 100) if A["Total"].sum() > 0 else 0
        st.metric("Évolution", f"{diff_pct:+.1f}%", _fmt_money(diff_val))

        st.divider()

        # ------- % par catégorie (période A)
        st.markdown("#### 📊 Répartition par catégorie (période A)")
        if not A.empty and "Categorie" in A.columns:
            grpA = A.groupby("Categorie")["Total"].sum().sort_values(ascending=False)
            totA = grpA.sum()
            df_pctA = pd.DataFrame({
                "Montant": grpA,
                "Part (%)": (grpA / totA * 100).round(1)
            })
            st.dataframe(df_pctA, use_container_width=True)
            st.bar_chart(df_pctA["Part (%)"], key=f"barA_{SID}")
        else:
            st.info("Pas de données pour la période A.")

        st.markdown("#### 📊 Répartition par sous-catégorie (période A)")
        if not A.empty and "Sous-categorie" in A.columns:
            grpA2 = A.groupby("Sous-categorie")["Total"].sum().sort_values(ascending=False)
            totA2 = grpA2.sum()
            df_pctA2 = pd.DataFrame({
                "Montant": grpA2,
                "Part (%)": (grpA2 / totA2 * 100).round(1)
            })
            st.dataframe(df_pctA2, use_container_width=True)
            st.bar_chart(df_pctA2["Part (%)"], key=f"barA2_{SID}")
        else:
            st.info("Pas de sous-catégories pour la période A.")

        st.divider()

        # ------- % par catégorie (période B)
        st.markdown("#### 📊 Répartition par catégorie (période B)")
        if not B.empty and "Categorie" in B.columns:
            grpB = B.groupby("Categorie")["Total"].sum().sort_values(ascending=False)
            totB = grpB.sum()
            df_pctB = pd.DataFrame({
                "Montant": grpB,
                "Part (%)": (grpB / totB * 100).round(1)
            })
            st.dataframe(df_pctB, use_container_width=True)
            st.bar_chart(df_pctB["Part (%)"], key=f"barB_{SID}")
        else:
            st.info("Pas de données pour la période B.")

        st.markdown("#### 📊 Répartition par sous-catégorie (période B)")
        if not B.empty and "Sous-categorie" in B.columns:
            grpB2 = B.groupby("Sous-categorie")["Total"].sum().sort_values(ascending=False)
            totB2 = grpB2.sum()
            df_pctB2 = pd.DataFrame({
                "Montant": grpB2,
                "Part (%)": (grpB2 / totB2 * 100).round(1)
            })
            st.dataframe(df_pctB2, use_container_width=True)
            st.bar_chart(df_pctB2["Part (%)"], key=f"barB2_{SID}")
        else:
            st.info("Pas de sous-catégories pour la période B.")

        st.divider()

        # ------- Évolution par catégorie
        st.markdown("#### 📈 Évolution par catégorie (A → B)")
        if not A.empty and not B.empty:
            comb = pd.DataFrame({
                "A": A.groupby("Categorie")["Total"].sum(),
                "B": B.groupby("Categorie")["Total"].sum()
            }).fillna(0)
            comb["Diff"] = comb["B"] - comb["A"]
            comb["Diff %"] = np.where(comb["A"] > 0, comb["Diff"] / comb["A"] * 100, np.nan).round(1)
            st.dataframe(comb, use_container_width=True)
            st.bar_chart(comb["Diff %"], key=f"barDiff_{SID}")
        else:
            st.info("Pas de données suffisantes pour la comparaison A→B.")




# ==============================
# BLOC 5/10 — Escrow
# ==============================

with tabs[2]:
    st.markdown("### 🏦 Escrow")

    dfE = df_clients_raw.copy()
    if dfE.empty:
        st.info("Aucun client chargé.")
    else:
        # Conversion monétaire sûre
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"]:
            if c not in dfE.columns:
                dfE[c] = 0.0
            dfE[c] = _to_num(dfE[c])

        dfE["Total"] = dfE["Montant honoraires (US $)"] + dfE["Autres frais (US $)"]
        dfE["Solde"] = dfE["Total"] - dfE["Payé"]

        # --- KPI compacts
        c1, c2, c3, c4 = st.columns([1,1,1,1])
        c1.metric("Total dossiers", f"{len(dfE)}")
        c2.metric("Montant total", _fmt_money(dfE["Total"].sum()))
        c3.metric("Payé", _fmt_money(dfE["Payé"].sum()))
        c4.metric("Solde restant", _fmt_money(dfE["Solde"].sum()))

        st.divider()

        # --- Filtrage par catégorie/sous-cat
        cats = sorted(dfE.get("Categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        subs = sorted(dfE.get("Sous-categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

        f1, f2 = st.columns(2)
        fcat = f1.multiselect("Catégories", cats, default=[], key=f"escrow_cats_{SID}")
        fsub = f2.multiselect("Sous-catégories", subs, default=[], key=f"escrow_subs_{SID}")

        dfV = dfE.copy()
        if fcat: dfV = dfV[dfV["Categorie"].astype(str).isin(fcat)]
        if fsub: dfV = dfV[dfV["Sous-categorie"].astype(str).isin(fsub)]

        st.caption(f"**{len(dfV)} dossiers affichés après filtres.**")

        # --- Tableau synthèse
        df_sum = dfV.groupby("Categorie")[["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"]].sum().round(2)
        df_sum["% Payé"] = np.where(df_sum["Total (US $)"] > 0, (df_sum["Payé"] / (df_sum["Montant honoraires (US $)"] + df_sum["Autres frais (US $)"])) * 100, 0)
        st.dataframe(df_sum, use_container_width=True)

        # --- Graphique par catégorie
        st.markdown("#### 💰 Répartition des montants par catégorie")
        if not df_sum.empty:
            st.bar_chart(df_sum[["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"]])
        else:
            st.info("Aucune donnée à afficher.")

        st.divider()

        # --- Liste détaillée
        st.markdown("#### 📋 Détails des paiements individuels")
        show_cols = [c for c in [
            "ID_Client","Nom","Categorie","Sous-categorie","Visa","Montant honoraires (US $)",
            "Autres frais (US $)","Payé","Solde","Commentaires"
        ] if c in dfV.columns]

        # Format monétaire
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"]:
            if c in dfV.columns:
                dfV[c] = _to_num(dfV[c]).map(_fmt_money)

        st.dataframe(dfV[show_cols].reset_index(drop=True), use_container_width=True)



# ==============================
# BLOC 6/10 — Compte client & Gestion
# ==============================

# --- petites utilitaires locales (défensives ; réutilisent les helpers globaux si déjà définis) ---
def _safe_date(val):
    try:
        if isinstance(val, (date, datetime)):
            return val.date() if isinstance(val, datetime) else val
        d = pd.to_datetime(val, errors="coerce")
        return d.date() if pd.notna(d) else None
    except Exception:
        return None

def _ensure_col(df, col, default):
    if col not in df.columns:
        df[col] = default
    return df

def _make_id(df, name, d):
    base = re.sub(r"[^a-z0-9]+", "-", str(name).strip().lower())
    if not base:
        base = "client"
    stamp = (d if isinstance(d, date) else date.today()).strftime("%Y%m%d")
    candidate = f"{base}-{stamp}"
    i = 0
    while (df.get("ID_Client", pd.Series(dtype=str)).astype(str) == candidate).any():
        i += 1
        candidate = f"{base}-{stamp}-{i}"
    return candidate

def _next_dossier_num(df, start=13057):
    try:
        exist = pd.to_numeric(df.get("Dossier N", pd.Series(dtype=str)), errors="coerce").dropna()
        return int(exist.max()) + 1 if not exist.empty else int(start)
    except Exception:
        return int(start)

def _write_clients_df(df, path_str):
    p = Path(path_str)
    p.parent.mkdir(parents=True, exist_ok=True)
    if p.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
        with pd.ExcelWriter(p, engine="openpyxl") as wr:
            df.to_excel(wr, index=False)
    else:
        df.to_csv(p, index=False, encoding="utf-8")

# S’assure que certaines colonnes existent pour la suite
for col, default in [
    ("Montant honoraires (US $)", 0.0),
    ("Autres frais (US $)", 0.0),
    ("Payé", 0.0),
    ("Solde", 0.0),
    ("Commentaires", ""),
    ("Paiements", ""),  # JSON historique facultatif
]:
    df_clients_raw = _ensure_col(df_clients_raw, col, default)

# ==============================================
# 👤 ONGLET : Compte client (vue dossier + historique paiements)
# ==============================================
with tabs[3]:
    st.markdown("### 👤 Compte client")
    if df_clients_raw.empty:
        st.info("Aucun client chargé.")
    else:
        # Sélecteurs
        left, right = st.columns([2,3])
        nom_list = sorted(df_clients_raw["Nom"].dropna().astype(str).unique().tolist())
        id_list  = sorted(df_clients_raw["ID_Client"].dropna().astype(str).unique().tolist())
        sel_nom  = left.selectbox("Nom", [""] + nom_list, index=0, key=f"acct_nom_{SID}")
        sel_id   = right.selectbox("ID_Client", [""] + id_list, index=0, key=f"acct_id_{SID}")

        mask = None
        if sel_id:
            mask = (df_clients_raw["ID_Client"].astype(str) == sel_id)
        elif sel_nom:
            mask = (df_clients_raw["Nom"].astype(str) == sel_nom)

        if mask is None or not mask.any():
            st.stop()

        row = df_clients_raw[mask].iloc[0].copy()

        # Carte récap
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Honoraires", _fmt_money(_to_num(row.get("Montant honoraires (US $)", 0.0))))
        c2.metric("Autres frais", _fmt_money(_to_num(row.get("Autres frais (US $)", 0.0))))
        c3.metric("Payé", _fmt_money(_to_num(row.get("Payé", 0.0))))
        # recalcul solde robuste
        total_row = _to_num(row.get("Montant honoraires (US $)", 0.0)) + _to_num(row.get("Autres frais (US $)", 0.0))
        reste_row = max(0.0, total_row - _to_num(row.get("Payé", 0.0)))
        c4.metric("Solde", _fmt_money(reste_row))

        st.divider()

        # Historique des paiements (JSON list [{date, mode, montant, note}])
        pay_hist = []
        try:
            rawp = row.get("Paiements", "")
            if isinstance(rawp, str) and rawp.strip():
                pay_hist = json.loads(rawp)
            elif isinstance(rawp, list):
                pay_hist = rawp
        except Exception:
            pay_hist = []

        st.markdown("#### 💳 Historique des paiements")
        if pay_hist:
            ph = pd.DataFrame(pay_hist)
            if "montant" in ph.columns:
                ph["montant"] = pd.to_numeric(ph["montant"], errors="coerce").fillna(0.0).map(_fmt_money)
            st.dataframe(ph, use_container_width=True, hide_index=True)
        else:
            st.info("Aucun paiement saisi pour ce client.")

        # Ajouter un paiement (si solde > 0)
        if reste_row > 0:
            st.markdown("#### ➕ Ajouter un paiement")
            p1, p2, p3, p4 = st.columns([1,1,1,2])
            pay_date = p1.date_input("Date", value=date.today(), key=f"acct_paydate_{SID}")
            pay_mode = p2.selectbox("Mode", ["CB", "Chèque", "Cash", "Virement", "Venmo"], key=f"acct_paymode_{SID}")
            pay_amt  = p3.number_input("Montant (US $)", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=f"acct_payamt_{SID}")
            pay_note = p4.text_input("Note (optionnel)", "", key=f"acct_paynote_{SID}")
            if st.button("💾 Enregistrer le paiement", key=f"acct_paybtn_{SID}"):
                if float(pay_amt) <= 0:
                    st.warning("Le montant doit être > 0.")
                    st.stop()
                # Mettre à jour df
                idx = df_clients_raw[mask].index[0]
                # Historique
                new_item = {
                    "date": str(pay_date),
                    "mode": str(pay_mode),
                    "montant": float(pay_amt),
                    "note": str(pay_note or "")
                }
                try:
                    base_list = []
                    rawp = df_clients_raw.at[idx, "Paiements"]
                    if isinstance(rawp, str) and rawp.strip():
                        base_list = json.loads(rawp)
                    elif isinstance(rawp, list):
                        base_list = rawp
                    base_list.append(new_item)
                    df_clients_raw.at[idx, "Paiements"] = json.dumps(base_list, ensure_ascii=False)
                except Exception:
                    df_clients_raw.at[idx, "Paiements"] = json.dumps([new_item], ensure_ascii=False)

                # Montants
                new_paye = _to_num(df_clients_raw.at[idx, "Payé"]) + float(pay_amt)
                df_clients_raw.at[idx, "Payé"] = new_paye
                # recalcul solde
                h = _to_num(df_clients_raw.at[idx, "Montant honoraires (US $)"])
                o = _to_num(df_clients_raw.at[idx, "Autres frais (US $)"])
                df_clients_raw.at[idx, "Solde"] = max(0.0, h + o - new_paye)

                # Sauvegarde immédiate sur le même fichier
                _write_clients_df(df_clients_raw, clients_src)
                st.success("Paiement ajouté et sauvegardé.")
                st.cache_data.clear()
                st.rerun()
        else:
            st.success("Ce dossier est soldé.")

        st.divider()

        # Statut du dossier (sélecteur + date + RFE)
        st.markdown("#### 🗂️ Statut du dossier")
        s1, s2, s3 = st.columns([2,2,1])
        statut_opts = ["", "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé"]
        # déterminer statut courant (priorité selon colonnes présentes ; non-booleans désormais)
        current_statut = ""
        for label in statut_opts[1:]:
            if int(_to_num(row.get(label, 0))) == 1:
                current_statut = label
                break
        statut = s1.selectbox("Statut", statut_opts, index=(statut_opts.index(current_statut) if current_statut in statut_opts else 0), key=f"acct_statut_{SID}")
        # date associée
        date_map = {
            "Dossier envoyé": "Date d'envoi",
            "Dossier accepté": "Date d'acceptation",
            "Dossier refusé": "Date de refus",
            "Dossier annulé": "Date d'annulation",
        }
        dkey = date_map.get(statut, None)
        dval = _safe_date(row.get(dkey)) if dkey else None
        sd = s2.date_input("Date statut", value=(dval if dval else date.today()) if statut else None, key=f"acct_statdate_{SID}", disabled=(not statut))
        rfe_val = (int(_to_num(row.get("RFE", 0))) == 1)
        rfe = s3.checkbox("RFE", value=rfe_val, key=f"acct_rfe_{SID}", disabled=(not statut))

        if st.button("💾 Enregistrer le statut", key=f"acct_statbtn_{SID}"):
            idx = df_clients_raw[mask].index[0]
            # remettre à zéro
            for label in date_map.keys():
                if label in df_clients_raw.columns:
                    df_clients_raw.at[idx, label] = 0
            for label in date_map.values():
                if label in df_clients_raw.columns:
                    df_clients_raw.at[idx, label] = None
            # poser le nouveau
            if statut:
                df_clients_raw.at[idx, statut] = 1
                if dkey:
                    df_clients_raw.at[idx, dkey] = sd
            # RFE seulement si un statut actif
            df_clients_raw.at[idx, "RFE"] = 1 if (statut and rfe) else 0

            _write_clients_df(df_clients_raw, clients_src)
            st.success("Statut mis à jour.")
            st.cache_data.clear()
            st.rerun()

# ==============================================
# 🧾 ONGLET : Gestion (Ajouter / Modifier / Supprimer)
# ==============================================
with tabs[4]:
    st.markdown("### 🧾 Gestion des clients")
    if df_clients_raw.empty:
        st.info("Aucun client chargé.")
    else:
        op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_op_{SID}")

        # ---------- AJOUT ----------
        if op == "Ajouter":
            st.markdown("#### ➕ Ajouter un client")
            a1, a2, a3 = st.columns([2,1,1])
            nom  = a1.text_input("Nom", "", key=f"add_nom_{SID}")
            dcrt = a2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
            mois = a3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=date.today().month-1, key=f"add_mois_{SID}")

            b1, b2, b3 = st.columns(3)
            cat  = b1.selectbox("Catégorie", [""] + sorted(df_clients_raw["Categorie"].dropna().astype(str).unique().tolist()), index=0, key=f"add_cat_{SID}")
            sub  = b2.selectbox("Sous-catégorie", [""] + sorted(df_clients_raw["Sous-categorie"].dropna().astype(str).unique().tolist()), index=0, key=f"add_sub_{SID}")
            visa = b3.text_input("Visa", "", key=f"add_visa_{SID}")

            c1, c2 = st.columns(2)
            honor = c1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
            other = c2.number_input("Autres frais (US $)",       min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")

            com = st.text_area("Commentaires (optionnel)", "", key=f"add_com_{SID}")

            if st.button("💾 Enregistrer le client", key=f"add_btn_{SID}"):
                if not nom:
                    st.warning("Le nom est requis.")
                    st.stop()
                did = _make_id(df_clients_raw, nom, dcrt)
                dossier_n = _next_dossier_num(df_clients_raw, start=13057)
                total = float(honor) + float(other)

                new_row = {
                    "ID_Client": did,
                    "Dossier N": dossier_n,
                    "Nom": nom,
                    "Date": dcrt,
                    "Mois": f"{int(mois):02d}",
                    "Categorie": cat,
                    "Sous-categorie": sub,
                    "Visa": visa or sub,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Payé": 0.0,
                    "Solde": total,
                    "Commentaires": com,
                    "Paiements": "[]",
                    "Dossier envoyé": 0,
                    "Date d'envoi": None,
                    "Dossier accepté": 0,
                    "Date d'acceptation": None,
                    "Dossier refusé": 0,
                    "Date de refus": None,
                    "Dossier annulé": 0,
                    "Date d'annulation": None,
                    "RFE": 0,
                }
                df_new = pd.concat([df_clients_raw, pd.DataFrame([new_row])], ignore_index=True)
                _write_clients_df(df_new, clients_src)
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()

        # ---------- MODIFICATION ----------
        elif op == "Modifier":
            st.markdown("#### ✏️ Modifier un client")
            names = sorted(df_clients_raw["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(df_clients_raw["ID_Client"].dropna().astype(str).unique().tolist())
            m1, m2 = st.columns(2)
            target_name = m1.selectbox("Nom", [""]+names, index=0, key=f"mod_nm_{SID}")
            target_id   = m2.selectbox("ID_Client", [""]+ids, index=0, key=f"mod_id_{SID}")

            mask = None
            if target_id:
                mask = (df_clients_raw["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_clients_raw["Nom"].astype(str) == target_name)

            if mask is None or not mask.any():
                st.stop()

            idx = df_clients_raw[mask].index[0]
            row = df_clients_raw.loc[idx].copy()

            d1, d2, d3 = st.columns([2,1,1])
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nom_{SID}")
            dval = _safe_date(row.get("Date")) or date.today()
            dcrt = d2.date_input("Date de création", value=dval, key=f"mod_date_{SID}")
            try:
                mval = int(str(row.get("Mois","01"))[:2])
                mval = 1 if mval < 1 or mval > 12 else mval
            except Exception:
                mval = 1
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=mval-1, key=f"mod_mois_{SID}")

            e1, e2, e3 = st.columns(3)
            cat  = e1.selectbox("Catégorie", [""] + sorted(df_clients_raw["Categorie"].dropna().astype(str).unique().tolist()),
                                index= (1 + sorted(df_clients_raw["Categorie"].dropna().astype(str).unique().tolist()).index(_safe_str(row.get("Categorie","")))) if _safe_str(row.get("Categorie","")) in sorted(df_clients_raw["Categorie"].dropna().astype(str).unique().tolist()) else 0,
                                key=f"mod_cat_{SID}")
            sub  = e2.selectbox("Sous-catégorie", [""] + sorted(df_clients_raw["Sous-categorie"].dropna().astype(str).unique().tolist()),
                                index= (1 + sorted(df_clients_raw["Sous-categorie"].dropna().astype(str).unique().tolist()).index(_safe_str(row.get("Sous-categorie","")))) if _safe_str(row.get("Sous-categorie","")) in sorted(df_clients_raw["Sous-categorie"].dropna().astype(str).unique().tolist()) else 0,
                                key=f"mod_sub_{SID}")
            visa = e3.text_input("Visa", _safe_str(row.get("Visa","")), key=f"mod_visa_{SID}")

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=float(_to_num(row.get("Montant honoraires (US $)",0.0))), step=50.0, format="%.2f", key=f"mod_h_{SID}")
            other = f2.number_input("Autres frais (US $)",       min_value=0.0, value=float(_to_num(row.get("Autres frais (US $)",0.0))),       step=20.0, format="%.2f", key=f"mod_o_{SID}")

            com = st.text_area("Commentaires", _safe_str(row.get("Commentaires","")), key=f"mod_com_{SID}")

            if st.button("💾 Enregistrer les modifications", key=f"mod_btn_{SID}"):
                df_clients_raw.at[idx, "Nom"] = nom
                df_clients_raw.at[idx, "Date"] = dcrt
                df_clients_raw.at[idx, "Mois"] = f"{int(mois):02d}"
                df_clients_raw.at[idx, "Categorie"] = cat
                df_clients_raw.at[idx, "Sous-categorie"] = sub
                df_clients_raw.at[idx, "Visa"] = visa or sub
                df_clients_raw.at[idx, "Montant honoraires (US $)"] = float(honor)
                df_clients_raw.at[idx, "Autres frais (US $)"] = float(other)
                total = float(honor) + float(other)
                paye  = float(_to_num(df_clients_raw.at[idx, "Payé"]))
                df_clients_raw.at[idx, "Solde"] = max(0.0, total - paye)
                df_clients_raw.at[idx, "Commentaires"] = com

                _write_clients_df(df_clients_raw, clients_src)
                st.success("Client modifié.")
                st.cache_data.clear()
                st.rerun()

        # ---------- SUPPRESSION ----------
        else:
            st.markdown("#### 🗑️ Supprimer un client")
            names = sorted(df_clients_raw["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(df_clients_raw["ID_Client"].dropna().astype(str).unique().tolist())
            s1, s2 = st.columns(2)
            target_name = s1.selectbox("Nom", [""]+names, index=0, key=f"del_nm_{SID}")
            target_id   = s2.selectbox("ID_Client", [""]+ids, index=0, key=f"del_id_{SID}")

            mask = None
            if target_id:
                mask = (df_clients_raw["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_clients_raw["Nom"].astype(str) == target_name)

            if mask is not None and mask.any():
                r = df_clients_raw[mask].iloc[0]
                st.write({"Dossier N": r.get("Dossier N",""), "Nom": r.get("Nom",""), "Visa": r.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"del_btn_{SID}"):
                    df_new = df_clients_raw[~mask].copy()
                    _write_clients_df(df_new, clients_src)
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()




# ==============================
# BLOC 7/10 — Visa (aperçu & cohérence)
# ==============================

with tabs[5]:
    st.markdown("### 📄 Visa — aperçu & contrôle de cohérence")

    if df_visa_raw is None or df_visa_raw.empty:
        st.info("Aucun tableau Visa chargé.")
    else:
        # --- Colonnes de base attendues (tolérance accents / variantes) ---
        def _col(df, *cands):
            for c in cands:
                if c in df.columns:
                    return c
            return None

        COL_CAT = _col(df_visa_raw, "Categorie", "Catégorie", "Category")
        COL_SUB = _col(df_visa_raw, "Sous-categorie", "Sous-catégorie", "Subcategory", "Sous-categories", "Sous-categories 1")
        base_cols = [c for c in [COL_CAT, COL_SUB] if c]

        # Colonnes "options" (cases à cocher) = tout le reste
        opt_cols = [c for c in df_visa_raw.columns if c not in base_cols]

        # --- Filtres latéraux (propres à l'onglet) ---
        cats = sorted(df_visa_raw[COL_CAT].dropna().astype(str).unique().tolist()) if COL_CAT else []
        subs = sorted(df_visa_raw[COL_SUB].dropna().astype(str).unique().tolist()) if COL_SUB else []

        f1, f2 = st.columns(2)
        fc = f1.multiselect("Catégories", cats, default=[], key=f"visa_cat_{SID}")
        fs = f2.multiselect("Sous-catégories", subs, default=[], key=f"visa_sub_{SID}")

        vf = df_visa_raw.copy()
        if fc and COL_CAT:
            vf = vf[vf[COL_CAT].astype(str).isin(fc)]
        if fs and COL_SUB:
            vf = vf[vf[COL_SUB].astype(str).isin(fs)]

        # Affichage du tableau filtré
        st.markdown("#### Tableau Visa (filtré)")
        st.dataframe(vf.reset_index(drop=True), use_container_width=True, hide_index=True, key=f"visa_table_{SID}")

        # --- Construction d’une carte Catégorie → Sous-catégorie → Options disponibles ---
        def build_map(df):
            m = {}
            if not (COL_CAT and COL_SUB):
                return m
            # options disponibles = colonnes opt où la valeur == 1 sur au moins une ligne du couple (cat, sub)
            for (cat, sub), grp in df.groupby([COL_CAT, COL_SUB], dropna=True):
                if pd.isna(cat) or pd.isna(sub):
                    continue
                cat_s = str(cat).strip()
                sub_s = str(sub).strip()
                m.setdefault(cat_s, {})
                opts = []
                for oc in opt_cols:
                    try:
                        v = pd.to_numeric(grp[oc], errors="coerce").fillna(0).astype(float)
                        if (v == 1).any():
                            opts.append(str(oc))
                    except Exception:
                        pass
                m[cat_s][sub_s] = sorted(opts)
            return m

        visa_tree = build_map(df_visa_raw)

        st.markdown("#### Arborescence (catégorie → sous-catégorie → options)")
        if not visa_tree:
            st.info("Impossible de construire l’arborescence (colonnes Catégorie/Sous-catégorie manquantes?).")
        else:
            # rendu compact
            for cat in sorted(visa_tree.keys()):
                with st.expander(f"📁 {cat}", expanded=False):
                    for sub in sorted(visa_tree[cat].keys()):
                        opts = visa_tree[cat][sub]
                        st.write(f"- **{sub}** : {(', '.join(opts)) if opts else '—'}")

        st.divider()

        # --- Contrôle de cohérence : clients vs visa_tree ---
        st.markdown("#### 🔎 Contrôle de cohérence Client ⇄ Visa")
        if df_clients_raw is None or df_clients_raw.empty:
            st.info("Aucun client chargé — impossible de vérifier la cohérence.")
        else:
            # Vérifier que (Categorie, Sous-categorie) existe dans la carte
            cc = df_clients_raw.copy()
            ccat = "Categorie" if "Categorie" in cc.columns else ("Catégorie" if "Catégorie" in cc.columns else None)
            csub = "Sous-categorie" if "Sous-categorie" in cc.columns else ("Sous-catégorie" if "Sous-catégorie" in cc.columns else None)

            bad_rows = []
            if ccat and csub:
                for i, r in cc.iterrows():
                    cat = str(r.get(ccat, "")).strip()
                    sub = str(r.get(csub, "")).strip()
                    if cat and sub:
                        if cat not in visa_tree or sub not in visa_tree.get(cat, {}):
                            bad_rows.append({
                                "ID_Client": r.get("ID_Client",""),
                                "Nom": r.get("Nom",""),
                                "Categorie": cat,
                                "Sous-categorie": sub,
                                "Visa": r.get("Visa",""),
                            })
            if bad_rows:
                st.warning("Des clients ont une catégorie/sous-catégorie qui n’existe pas dans la table Visa.")
                st.dataframe(pd.DataFrame(bad_rows), use_container_width=True, hide_index=True, key=f"visa_inco_{SID}")
            else:
                st.success("Cohérence OK : toutes les catégories/sous-catégories client existent dans le tableau Visa.")




# ==============================
# BLOC 8/10 — 👤 Compte client (détails financiers + chronologie)
# ==============================

with tabs[3]:
    st.markdown("### 👤 Compte client — Détails et historique")

    df_acc = read_clients_file(clients_src)
    if df_acc is None or df_acc.empty:
        st.info("Aucun client chargé.")
        st.stop()

    c1, c2 = st.columns([2, 2])
    noms = sorted(df_acc["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_acc.columns else []
    ids = sorted(df_acc["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_acc.columns else []

    sel_nom = c1.selectbox("Nom", [""] + noms, index=0, key=f"acct_nom_{SID}")
    sel_id = c2.selectbox("ID_Client", [""] + ids, index=0, key=f"acct_id_{SID}")

    mask = None
    if sel_id:
        mask = (df_acc["ID_Client"].astype(str) == sel_id)
    elif sel_nom:
        mask = (df_acc["Nom"].astype(str) == sel_nom)

    if mask is None or not mask.any():
        st.stop()

    row = df_acc[mask].iloc[0]

    st.markdown(f"#### 🧾 Dossier N° {row.get('Dossier N', '?')} — {_safe_str(row.get('Nom', ''))}")

    # --- Section financière ---
    honor = float(_ensure_num(row.get("Montant honoraires (US $)", 0)))
    other = float(_ensure_num(row.get("Autres frais (US $)", 0)))
    paye = float(_ensure_num(row.get("Payé", 0)))
    solde = float(_ensure_num(row.get("Solde", honor + other - paye)))

    f1, f2, f3, f4 = st.columns(4)
    f1.metric("Honoraires", f"${honor:,.2f}")
    f2.metric("Autres frais", f"${other:,.2f}")
    f3.metric("Payé", f"${paye:,.2f}")
    f4.metric("Solde", f"${solde:,.2f}")

    st.divider()

    # --- Chronologie & Statuts ---
    s_env = int(_ensure_num(row.get("Dossier envoyé", 0))) == 1
    s_acc = int(_ensure_num(row.get("Dossier approuvé", 0))) == 1
    s_ref = int(_ensure_num(row.get("Dossier refusé", 0))) == 1
    s_ann = int(_ensure_num(row.get("Dossier annulé", 0))) == 1
    s_rfe = int(_ensure_num(row.get("RFE", 0))) == 1

    def sdate(col):
        val = _safe_str(row.get(col, ""))
        return val if val else "—"

    st.markdown("##### 📅 Statuts du dossier")

    f1, f2, f3 = st.columns(3)
    f1.write("**RFE :** " + ("Oui" if s_rfe else "Non"))
    f2.write("**Catégorie :** " + _safe_str(row.get("Categorie", "")))
    f3.write("**Sous-catégorie :** " + _safe_str(row.get("Sous-categorie", "")))

    st.divider()

    f1, f2, f3, f4 = st.columns(4)
    f1.write("**Dossier envoyé :** " + ("Oui" if s_env else "Non") + " — " + sdate("Date d'envoi"))
    f2.write("**Dossier approuvé :** " + ("Oui" if s_acc else "Non") + " — " + sdate("Date d'acceptation"))
    f3.write("**Dossier refusé :** " + ("Oui" if s_ref else "Non") + " — " + sdate("Date de refus"))
    f4.write("**Dossier annulé :** " + ("Oui" if s_ann else "Non") + " — " + sdate("Date d'annulation"))

    st.divider()

    st.markdown("##### 💬 Commentaires")
    st.info(_safe_str(row.get("Commentaires", "(aucun)")))

    # --- Paiements (liste et historique) ---
    st.markdown("##### 💵 Historique des paiements")
    try:
        payments = json.loads(row.get("Paiements", "[]")) if isinstance(row.get("Paiements"), str) else []
    except Exception:
        payments = []

    if payments:
        ptable = pd.DataFrame(payments)
        st.dataframe(ptable, use_container_width=True)
    else:
        st.write("Aucun paiement enregistré.")

    st.divider()

    # --- Ajout d’un nouveau paiement ---
    st.markdown("##### ➕ Ajouter un paiement")
    pay_col1, pay_col2, pay_col3 = st.columns([1, 1, 2])
    new_date = pay_col1.date_input("Date", value=date.today(), key=f"pay_date_{SID}")
    new_amount = pay_col2.number_input("Montant (US $)", min_value=0.0, step=50.0, format="%.2f", key=f"pay_amt_{SID}")
    new_note = pay_col3.text_input("Note", "", key=f"pay_note_{SID}")

    if st.button("💾 Ajouter ce paiement", key=f"add_payment_{SID}"):
        payments.append({
            "date": str(new_date),
            "montant": float(new_amount),
            "note": new_note,
        })
        total_paye = sum(p.get("montant", 0) for p in payments)
        total_due = honor + other
        solde = max(0.0, total_due - total_paye)

        row["Paiements"] = json.dumps(payments, ensure_ascii=False)
        row["Payé"] = total_paye
        row["Solde"] = solde

        df_acc.loc[mask, "Paiements"] = row["Paiements"]
        df_acc.loc[mask, "Payé"] = total_paye
        df_acc.loc[mask, "Solde"] = solde

        write_clients_file(df_acc, clients_save_path)
        st.success("Paiement ajouté avec succès ✅")
        st.cache_data.clear()
        st.rerun()



# ==============================
# BLOC 9/10 — 🧾 Gestion (CRUD : Ajouter / Modifier / Supprimer)
# (Dépend de : read_clients_file, write_clients_file, visa_df, visa_map, SID, _fmt_money, _ensure_num)
# ==============================

with tabs[4]:
    st.markdown("### 🧾 Gestion des clients (Ajouter • Modifier • Supprimer)")

    # Fabrique de clés UI
    def skey(*parts): 
        return f"{SID}_" + "_".join(str(p) for p in parts)

    # Helpers locaux
    def _month_mm(d):
        try:
            d = pd.to_datetime(d, errors="coerce")
            if pd.isna(d): 
                d = date.today()
            return f"{int(d.month):02d}"
        except Exception:
            return f"{date.today().month:02d}"

    def _parse_date_widget(v):
        """Accepte date/datetime/str -> date (ou None)."""
        if isinstance(v, date) and not isinstance(v, datetime):
            return v
        try:
            d = pd.to_datetime(v, errors="coerce")
            if pd.isna(d): 
                return None
            return d.date()
        except Exception:
            return None

    def _safe_str(v):
        return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)

    def _make_client_id(nom, d):
        base = (
            _safe_str(nom)
            .strip()
            .lower()
            .replace(" ", "-")
            .replace("'", "")
            .replace('"', "")
        )
        try:
            d = pd.to_datetime(d, errors="coerce")
            if pd.isna(d):
                d = pd.Timestamp.today()
        except Exception:
            d = pd.Timestamp.today()
        return f"{base}-{d:%Y%m%d}"

    def available_visa_options(cat, sub, visa_df):
        """
        Lit la ligne correspondant à (cat, sub) dans visa_df.
        Toute colonne (hors Catégorie/Sous-catégorie…) avec valeur == 1 est proposée comme option.
        """
        if visa_df is None or visa_df.empty or not cat or not sub:
            return []
        dfv = visa_df.copy()
        # colonnes d’axes
        axes_cols = [c for c in dfv.columns if c.lower().startswith("categorie") or c.lower().startswith("sous")]
        mask = (dfv[axes_cols[0]].astype(str) == str(cat)) & (dfv[axes_cols[1]].astype(str) == str(sub))
        if not mask.any():
            return []
        row = dfv[mask].iloc[0]
        opts = []
        for c in dfv.columns:
            if c in axes_cols:
                continue
            val = row.get(c, 0)
            try:
                val = int(float(val))
            except Exception:
                val = 0
            if val == 1:
                opts.append(str(c))
        return opts

    # Charger état courant
    df_live = read_clients_file(clients_src)
    if df_live is None:
        df_live = pd.DataFrame()
    if "Paiements" not in df_live.columns:
        df_live["Paiements"] = ""

    # Sélecteur d’opération
    op = st.radio(
        "Action",
        options=["Ajouter", "Modifier", "Supprimer"],
        horizontal=True,
        key=skey("crud","op")
    )

    # ---------- AJOUT ----------
    if op == "Ajouter":
        st.markdown("#### ➕ Ajouter un client")

        a1, a2, a3 = st.columns([2,1,1])
        nom = a1.text_input("Nom", "", key=skey("add","nom"))
        dval = a2.date_input("Date de création", value=date.today(), key=skey("add","date"))
        mois = a3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=date.today().month-1, key=skey("add","mois"))

        # Cascade Visa depuis visa_df
        cats = []
        subs = []
        if (visa_df is not None) and (not visa_df.empty):
            ccol = [c for c in visa_df.columns if c.lower().startswith("categorie")]
            scol = [c for c in visa_df.columns if c.lower().startswith("sous")]
            if ccol and scol:
                cats = sorted(visa_df[ccol[0]].dropna().astype(str).unique().tolist())
        sel_cat = st.selectbox("Catégorie", [""]+cats, index=0, key=skey("add","cat"))
        if sel_cat and (visa_df is not None) and (not visa_df.empty):
            ccol = [c for c in visa_df.columns if c.lower().startswith("categorie")]
            scol = [c for c in visa_df.columns if c.lower().startswith("sous")]
            if ccol and scol:
                subs = sorted(visa_df.loc[visa_df[ccol[0]].astype(str)==sel_cat, scol[0]].dropna().astype(str).unique().tolist())
        sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=skey("add","sub"))

        # Options selon la grille (ex: COS/EOS…)
        picked_opts = []
        visa_label  = ""
        if sel_cat and sel_sub:
            disp_opts = available_visa_options(sel_cat, sel_sub, visa_df)
            if disp_opts:
                st.caption("Options disponibles (coche une ou plusieurs si nécessaire) :")
                cols = st.columns(max(1, min(4, len(disp_opts))))
                chosen = []
                for i, opt in enumerate(disp_opts):
                    if cols[i % len(cols)].checkbox(opt, key=skey("add","opt",i)):
                        chosen.append(opt)
                picked_opts = chosen
                if len(chosen) == 1:
                    visa_label = f"{sel_sub} {chosen[0]}"
                elif len(chosen) > 1:
                    visa_label = f"{sel_sub} " + ",".join(chosen)
                else:
                    visa_label = sel_sub
            else:
                visa_label = sel_sub

        f1, f2 = st.columns(2)
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=skey("add","honor"))
        other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=skey("add","other"))
        comments = st.text_area("Commentaires (autres frais, notes…)", "", key=skey("add","comm"))

        st.markdown("##### Statut (toggle) + dates (texte libre)")
        s1, s2, s3, s4, s5 = st.columns(5)
        env = s1.toggle("Envoyé", key=skey("add","env"))
        acc = s2.toggle("Approuvé", key=skey("add","acc"))
        ref = s3.toggle("Refusé", key=skey("add","ref"))
        ann = s4.toggle("Annulé", key=skey("add","ann"))
        rfe = s5.toggle("RFE", key=skey("add","rfe"))

        d1, d2, d3, d4 = st.columns(4)
        de = d1.text_input("Date d'envoi", "", key=skey("add","de"))
        da = d2.text_input("Date d'acceptation", "", key=skey("add","da"))
        dr = d3.text_input("Date de refus", "", key=skey("add","dr"))
        dn = d4.text_input("Date d'annulation", "", key=skey("add","dn"))

        if st.button("💾 Enregistrer le client", key=skey("add","save")):
            if not nom:
                st.warning("Le nom est requis.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Choisis une Catégorie et une Sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            paye  = float(_ensure_num(0))
            solde = max(0.0, total - paye)

            # Dossier N (auto) -> max + 1, sinon 13057
            start_no = 13057
            if "Dossier N" in df_live.columns and not df_live.empty:
                try:
                    start_no = max(start_no, int(pd.to_numeric(df_live["Dossier N"], errors="coerce").max() or start_no) + 1)
                except Exception:
                    start_no = start_no + 1

            did = _make_client_id(nom, dval)

            new_row = {
                "ID_Client": did,
                "Dossier N": start_no,
                "Nom": nom,
                "Date": dval,
                "Mois": str(mois),
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": visa_label or sel_sub,
                "Montant honoraires (US $)": float(honor),
                "Autres frais (US $)": float(other),
                "Payé": float(paye),
                "Solde": float(solde),
                "Paiements": json.dumps([], ensure_ascii=False),
                "Commentaires": comments,
                "Dossier envoyé": 1 if env else 0,
                "Dossier approuvé": 1 if acc else 0,
                "Dossier refusé": 1 if ref else 0,
                "Dossier annulé": 1 if ann else 0,
                "RFE": 1 if rfe else 0,
                "Date d'envoi": de,
                "Date d'acceptation": da,
                "Date de refus": dr,
                "Date d'annulation": dn,
                "Options": json.dumps(picked_opts, ensure_ascii=False) if picked_opts else json.dumps([], ensure_ascii=False),
            }
            df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            write_clients_file(df_new, clients_save_path)
            st.success("Client ajouté.")
            st.cache_data.clear()
            st.rerun()

    # ---------- MODIFIER ----------
    if op == "Modifier":
        st.markdown("#### ✏️ Modifier un client")
        if df_live.empty:
            st.info("Aucun client à modifier.")
            st.stop()

        m1, m2 = st.columns([2,2])
        noms = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
        ids  = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
        sel_nom = m1.selectbox("Nom", [""]+noms, index=0, key=skey("mod","nom"))
        sel_id  = m2.selectbox("ID_Client", [""]+ids, index=0, key=skey("mod","id"))

        mask = None
        if sel_id:
            mask = (df_live["ID_Client"].astype(str) == sel_id)
        elif sel_nom:
            mask = (df_live["Nom"].astype(str) == sel_nom)
        if mask is None or not mask.any():
            st.stop()

        idx = df_live[mask].index[0]
        row = df_live.loc[idx].copy()

        # Champs
        d1, d2, d3 = st.columns([2,1,1])
        nom = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=skey("mod","nomv"))
        dval = row.get("Date")
        try:
            dparsed = pd.to_datetime(dval, errors="coerce")
            ddefault = dparsed.date() if pd.notna(dparsed) else date.today()
        except Exception:
            ddefault = date.today()
        dte  = d2.date_input("Date de création", value=ddefault, key=skey("mod","date"))
        mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=(int(_safe_str(row.get("Mois","01"))) - 1 if _safe_str(row.get("Mois","01")).isdigit() else date.today().month-1),
                            key=skey("mod","mois"))

        # Visa cascade
        cats = []
        subs = []
        if (visa_df is not None) and (not visa_df.empty):
            ccol = [c for c in visa_df.columns if c.lower().startswith("categorie")]
            scol = [c for c in visa_df.columns if c.lower().startswith("sous")]
            if ccol and scol:
                cats = sorted(visa_df[ccol[0]].dropna().astype(str).unique().tolist())

        preset_cat = _safe_str(row.get("Categorie",""))
        sel_cat = st.selectbox("Catégorie", [""]+cats,
                               index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                               key=skey("mod","cat"))

        if sel_cat and (visa_df is not None) and (not visa_df.empty):
            ccol = [c for c in visa_df.columns if c.lower().startswith("categorie")]
            scol = [c for c in visa_df.columns if c.lower().startswith("sous")]
            if ccol and scol:
                subs = sorted(visa_df.loc[visa_df[ccol[0]].astype(str)==sel_cat, scol[0]].dropna().astype(str).unique().tolist())

        preset_sub = _safe_str(row.get("Sous-categorie",""))
        sel_sub = st.selectbox("Sous-catégorie", [""]+subs,
                               index=(subs.index(preset_sub)+1 if preset_sub in subs else 0),
                               key=skey("mod","sub"))

        picked_opts = []
        visa_label  = _safe_str(row.get("Visa",""))
        if sel_cat and sel_sub:
            disp_opts = available_visa_options(sel_cat, sel_sub, visa_df)
            if disp_opts:
                st.caption("Options disponibles :")
                # Pré-sélection à partir de "Options" s'il existe
                preset = row.get("Options", "[]")
                try:
                    preset_list = json.loads(preset) if isinstance(preset, str) else (preset if isinstance(preset, list) else [])
                except Exception:
                    preset_list = []
                cols = st.columns(max(1, min(4, len(disp_opts))))
                chosen = []
                for i, opt in enumerate(disp_opts):
                    checked = opt in preset_list
                    if cols[i % len(cols)].checkbox(opt, value=checked, key=skey("mod","opt",i)):
                        chosen.append(opt)
                picked_opts = chosen
                if len(chosen) == 1:
                    visa_label = f"{sel_sub} {chosen[0]}"
                elif len(chosen) > 1:
                    visa_label = f"{sel_sub} " + ",".join(chosen)
                else:
                    visa_label = sel_sub
            else:
                visa_label = sel_sub

        f1, f2 = st.columns(2)
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0,
                                value=float(_ensure_num(row.get("Montant honoraires (US $)",0))),
                                step=50.0, format="%.2f", key=skey("mod","honor"))
        other = f2.number_input("Autres frais (US $)", min_value=0.0,
                                value=float(_ensure_num(row.get("Autres frais (US $)",0))),
                                step=20.0, format="%.2f", key=skey("mod","other"))
        comments = st.text_area("Commentaires", _safe_str(row.get("Commentaires","")), key=skey("mod","comm"))

        st.markdown("##### Statut + dates")
        s1, s2, s3, s4, s5 = st.columns(5)
        env = s1.toggle("Envoyé", value=bool(int(_ensure_num(row.get("Dossier envoyé",0)))), key=skey("mod","env"))
        acc = s2.toggle("Approuvé", value=bool(int(_ensure_num(row.get("Dossier approuvé",0)))), key=skey("mod","acc"))
        ref = s3.toggle("Refusé",  value=bool(int(_ensure_num(row.get("Dossier refusé",0)))),  key=skey("mod","ref"))
        ann = s4.toggle("Annulé",  value=bool(int(_ensure_num(row.get("Dossier annulé",0)))),  key=skey("mod","ann"))
        rfe = s5.toggle("RFE",     value=bool(int(_ensure_num(row.get("RFE",0)))),             key=skey("mod","rfe"))

        t1, t2, t3, t4 = st.columns(4)
        de = t1.text_input("Date d'envoi", _safe_str(row.get("Date d'envoi","")), key=skey("mod","de"))
        da = t2.text_input("Date d'acceptation", _safe_str(row.get("Date d'acceptation","")), key=skey("mod","da"))
        dr = t3.text_input("Date de refus", _safe_str(row.get("Date de refus","")), key=skey("mod","dr"))
        dn = t4.text_input("Date d'annulation", _safe_str(row.get("Date d'annulation","")), key=skey("mod","dn"))

        if st.button("💾 Enregistrer les modifications", key=skey("mod","save")):
            if not nom:
                st.warning("Le nom est requis.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Choisis une Catégorie et une Sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            paye  = float(_ensure_num(row.get("Payé",0)))
            solde = max(0.0, total - paye)

            df_live.at[idx, "Nom"] = nom
            df_live.at[idx, "Date"] = dte
            df_live.at[idx, "Mois"] = str(mois)
            df_live.at[idx, "Categorie"] = sel_cat
            df_live.at[idx, "Sous-categorie"] = sel_sub
            df_live.at[idx, "Visa"] = visa_label or sel_sub
            df_live.at[idx, "Montant honoraires (US $)"] = float(honor)
            df_live.at[idx, "Autres frais (US $)"] = float(other)
            df_live.at[idx, "Payé"] = float(paye)
            df_live.at[idx, "Solde"] = float(solde)
            df_live.at[idx, "Commentaires"] = comments
            df_live.at[idx, "Options"] = json.dumps(picked_opts, ensure_ascii=False) if picked_opts else json.dumps([], ensure_ascii=False)
            df_live.at[idx, "Dossier envoyé"] = 1 if env else 0
            df_live.at[idx, "Dossier approuvé"] = 1 if acc else 0
            df_live.at[idx, "Dossier refusé"] = 1 if ref else 0
            df_live.at[idx, "Dossier annulé"] = 1 if ann else 0
            df_live.at[idx, "RFE"] = 1 if rfe else 0
            df_live.at[idx, "Date d'envoi"] = de
            df_live.at[idx, "Date d'acceptation"] = da
            df_live.at[idx, "Date de refus"] = dr
            df_live.at[idx, "Date d'annulation"] = dn

            write_clients_file(df_live, clients_save_path)
            st.success("Modifications enregistrées.")
            st.cache_data.clear()
            st.rerun()

    # ---------- SUPPRIMER ----------
    if op == "Supprimer":
        st.markdown("#### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client à supprimer.")
            st.stop()

        s1, s2 = st.columns([2,2])
        noms = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
        ids  = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
        sel_nom = s1.selectbox("Nom", [""]+noms, index=0, key=skey("del","nom"))
        sel_id  = s2.selectbox("ID_Client", [""]+ids, index=0, key=skey("del","id"))

        mask = None
        if sel_id:
            mask = (df_live["ID_Client"].astype(str) == sel_id)
        elif sel_nom:
            mask = (df_live["Nom"].astype(str) == sel_nom)

        if mask is not None and mask.any():
            row = df_live[mask].iloc[0]
            st.write({
                "Dossier N": row.get("Dossier N",""),
                "Nom": row.get("Nom",""),
                "Visa": row.get("Visa",""),
            })
            if st.button("❗ Confirmer la suppression", key=skey("del","go")):
                df_new = df_live[~mask].copy()
                write_clients_file(df_new, clients_save_path)
                st.success("Client supprimé.")
                st.cache_data.clear()
                st.rerun()



# ==============================
# BLOC 10/10 — 📄 Visa (aperçu) & 💾 Export
# (Dépend de : visa_df, clients_src, visa_src, read_clients_file, _safe_str, SID)
# ==============================

# -------- 📄 Visa (aperçu) --------
with tabs[5]:
    st.markdown("### 📄 Visa (aperçu)")
    if visa_df is None or visa_df.empty:
        st.info("Aucun tableau Visa chargé.")
    else:
        vdf = visa_df.copy()

        # Colonnes d'axes (Catégorie + 1ère Sous-catégorie)
        cat_cols = [c for c in vdf.columns if c.lower().startswith("categorie")]
        sub_cols = [c for c in vdf.columns if c.lower().startswith("sous")]
        opt_cols = [c for c in vdf.columns if c not in (cat_cols[:1] + sub_cols[:1])]

        # Filtres
        c1, c2 = st.columns(2)
        cats = sorted(vdf[cat_cols[0]].dropna().astype(str).unique().tolist()) if cat_cols else []
        subs = sorted(vdf[sub_cols[0]].dropna().astype(str).unique().tolist()) if sub_cols else []

        fcat = c1.multiselect("Catégories", cats, default=[], key=f"visa_cat_{SID}")
        fsub = c2.multiselect("Sous-catégories", subs, default=[], key=f"visa_sub_{SID}")

        vf = vdf.copy()
        if fcat and cat_cols:
            vf = vf[vf[cat_cols[0]].astype(str).isin(fcat)]
        if fsub and sub_cols:
            vf = vf[vf[sub_cols[0]].astype(str).isin(fsub)]

        # Conversion options 1/0 -> ✓/vide (pour lecture plus claire)
        for c in opt_cols:
            try:
                vf[c] = (pd.to_numeric(vf[c], errors="coerce").fillna(0).astype(int) == 1).map({True:"✓", False:""})
            except Exception:
                pass

        # Affichage
        show_cols = (cat_cols[:1] + sub_cols[:1] + opt_cols) if (cat_cols and sub_cols) else vf.columns.tolist()
        st.dataframe(vf[show_cols].reset_index(drop=True), use_container_width=True, key=f"visa_preview_{SID}")

        st.caption("Astuce : la colonne Visa du client est générée par « Sous-catégorie + option(s) cochée(s) ».")

# -------- 💾 Export (ZIP) --------
with tabs[6]:
    st.markdown("### 💾 Export (Clients + Visa)")
    z1, z2 = st.columns([1,3])

    if st.button("Préparer l’archive ZIP", key=f"zip_make_{SID}"):
        try:
            from io import BytesIO
            import zipfile

            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients (reprend le fichier tel quel si possible)
                try:
                    # si l'utilisateur a chargé un fichier source, on exporte ce fichier original
                    if isinstance(clients_src, str) and len(clients_src) and os.path.exists(clients_src):
                        zf.write(clients_src, arcname=os.path.basename(_safe_str(clients_src)) or "Clients.xlsx")
                    else:
                        # sinon, on reconstruit depuis df
                        dfC = read_clients_file(clients_src)
                        with BytesIO() as xb:
                            with pd.ExcelWriter(xb, engine="openpyxl") as wr:
                                dfC.to_excel(wr, sheet_name="Clients", index=False)
                            zf.writestr("Clients.xlsx", xb.getvalue())
                except Exception:
                    pass

                # Visa (idem)
                try:
                    if isinstance(visa_src, str) and len(visa_src) and os.path.exists(visa_src):
                        zf.write(visa_src, arcname=os.path.basename(_safe_str(visa_src)) or "Visa.xlsx")
                    else:
                        dfV = visa_df if visa_df is not None else pd.DataFrame()
                        with BytesIO() as vb:
                            with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                dfV.to_excel(wr, sheet_name="Visa", index=False)
                            zf.writestr("Visa.xlsx", vb.getvalue())
                except Exception:
                    pass

            st.session_state[f"zip_{SID}"] = buf.getvalue()
            st.success("Archive ZIP prête.")
        except Exception as e:
            st.error(f"Erreur lors de la préparation : {_safe_str(e)}")

    if st.session_state.get(f"zip_{SID}"):
        st.download_button(
            "⬇️ Télécharger l’archive",
            data=st.session_state[f"zip_{SID}"],
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
            key=f"zip_dl_{SID}",
        )

# -------- 🧭 Pied de page --------
st.markdown("---")
st.caption("Visa Manager — version latérale simple (sans Plotly) • Filtres persistants et derniers chemins mémorisés.")