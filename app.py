
# ================================
# 🛂 Visa Manager — PARTIE 1/4
# Imports, constantes, helpers, I/O, lecture fichiers, carte Visa
# ================================

from __future__ import annotations

import json
import re
import zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Dict, List, Tuple, Any

import pandas as pd
import streamlit as st

# ---------------- Page config ----------------
st.set_page_config(page_title="Visa Manager", layout="wide", page_icon="🛂")

# --- Identifiant de session unique pour les clés de widgets ---
if "SID" not in st.session_state:
    st.session_state["SID"] = "main"
SID = st.session_state["SID"]

def skey(*parts) -> str:
    """Génère une clé stable pour widgets (évite collisions)."""
    return "_".join([*map(lambda x: str(x).replace(" ", "_"), parts), st.session_state.get("SID", "main")])

# ---------------- Constantes ----------------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"
DOSSIER_COL = "Dossier N"

# ---------------- Utils "sûrs" ----------------
def _safe_str(x: Any) -> str:
    try:
        if x is None:
            return ""
        return str(x)
    except Exception:
        return ""

def _fmt_money(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def _safe_num_series(df_or_ser: Any, col_or_idx: Any):
    """Retourne une Series numérique propre (remplace NaN/non-num par 0)."""
    try:
        if isinstance(df_or_ser, pd.DataFrame):
            s = df_or_ser.get(col_or_idx, pd.Series(dtype=float))
        else:
            s = df_or_ser
        s = pd.to_numeric(s, errors="coerce").fillna(0.0)
        return s
    except Exception:
        return pd.Series(dtype=float)

def _to_int_bool(v: Any) -> bool:
    try:
        if isinstance(v, bool):
            return v
        if isinstance(v, (int, float)) and not pd.isna(v):
            return int(v) == 1
        if isinstance(v, str):
            return v.strip().lower() in ["1", "true", "oui", "yes", "x"]
    except Exception:
        pass
    return False

def _date_for_widget(val: Any):
    """Convertit val en date (widget Streamlit), sinon None."""
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return None
        return d.date()
    except Exception:
        return None

def _make_client_id(nom: str, d: date) -> str:
    """ID_CLIENT basé sur nom + AAAAMMJJ, gère doublons suffixés."""
    base = re.sub(r"[^A-Za-z0-9_-]+", "-", _safe_str(nom)).strip("-").lower() or "client"
    return f"{base}-{d:%Y%m%d}"

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    """Renvoie le prochain numéro de dossier à partir de 'start'."""
    if DOSSIER_COL in df.columns:
        nums = pd.to_numeric(df[DOSSIER_COL], errors="coerce").dropna()
        if not nums.empty:
            return int(max(int(nums.max()) + 1, start))
    return int(start)

# ---------------- Persistance des chemins / buffers ----------------
# On mémorise dans la session le dernier fichier utilisé (clients & visa)
if "LAST_FILES" not in st.session_state:
    st.session_state["LAST_FILES"] = {
        "clients_name": "",
        "clients_bytes": None,
        "visa_name": "",
        "visa_bytes": None,
        "mode": "two",  # "two" (2 fichiers) ou "single" (un xlsx avec 2 onglets)
        "single_name": "",
        "single_bytes": None,
    }

LAST = st.session_state["LAST_FILES"]

# ---------------- Fonctions I/O ----------------
@st.cache_data(show_spinner=False)
def read_sheet(xlsx_path_or_bytes, sheet_name: str) -> pd.DataFrame:
    """Lit une feuille excel depuis chemin (str) ou bytes."""
    try:
        return pd.read_excel(xlsx_path_or_bytes, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def _read_clients(source) -> pd.DataFrame:
    df = read_sheet(source, SHEET_CLIENTS)
    # Normalisation colonnes attendues
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c not in df.columns:
            df[c] = 0.0
    if "Mois" in df.columns:
        df["Mois"] = df["Mois"].astype(str).str.zfill(2)
    # colonnes statut
    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        if c not in df.columns:
            df[c] = 0
    # dates statut
    for c in ["Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation"]:
        if c not in df.columns:
            df[c] = None
    if "_Année_" not in df.columns:
        # derive année/mois numériques pour analyses/tri
        try:
            d = pd.to_datetime(df.get("Date", pd.NaT), errors="coerce")
            df["_Année_"] = d.dt.year
            df["_MoisNum_"] = d.dt.month
        except Exception:
            df["_Année_"] = pd.NA
            df["_MoisNum_"] = pd.NA
    # Paiements/Options: forcer un format JSON/obj
    if "Paiements" not in df.columns:
        df["Paiements"] = [[] for _ in range(len(df))]
    if "Options" not in df.columns:
        df["Options"] = [{} for _ in range(len(df))]
    return df

def _write_clients(df: pd.DataFrame, target_buffer: BytesIO | None) -> None:
    """
    Écrit la feuille Clients dans le dernier container (2 modes possibles) :
    - mode 'two' -> on tient un buffer 'clients_bytes'
    - mode 'single' -> on met à jour la feuille 'Clients' au sein du fichier unique
    """
    if LAST["mode"] == "single":
        # on réécrit le fichier unique (2 onglets)
        base = BytesIO(LAST["single_bytes"]) if LAST["single_bytes"] else BytesIO()
        # Si vide, on crée un nouveau xlsx avec 2 feuilles vides par défaut
        out = BytesIO()
        # Lire existant si possible
        try:
            whole_clients = df.copy()
            visa_df = read_sheet(base if base.getbuffer().nbytes else LAST["single_bytes"], SHEET_VISA)
        except Exception:
            whole_clients = df.copy()
            visa_df = pd.DataFrame()
        with pd.ExcelWriter(out, engine="openpyxl") as wr:
            whole_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
            if not visa_df.empty:
                visa_df.to_excel(wr, sheet_name=SHEET_VISA, index=False)
            else:
                pd.DataFrame(columns=["Categorie","Sous-categorie"]).to_excel(
                    wr, sheet_name=SHEET_VISA, index=False
                )
        LAST["single_bytes"] = out.getvalue()
        st.cache_data.clear()
        return
    else:
        # mode "two" : on met à jour uniquement le buffer clients
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        LAST["clients_bytes"] = out.getvalue()
        st.cache_data.clear()
        return

@st.cache_data(show_spinner=False)
def read_visa_raw(xlsx_path_or_bytes) -> pd.DataFrame:
    """
    Lit l'onglet Visa tel quel. On s'attend à :
    - Colonnes au minimum : Categorie, Sous-categorie
    - Colonnes options (ex: COS, EOS, ...), avec '1' pour indiquer que l'option existe.
    """
    try:
        return pd.read_excel(xlsx_path_or_bytes, sheet_name=SHEET_VISA)
    except Exception:
        try:
            # Si le fichier visa est seul sans onglet "Visa", on lit la première feuille
            return pd.read_excel(xlsx_path_or_bytes)
        except Exception:
            return pd.DataFrame()

# ---------------- Chargement fichiers (UI) ----------------
st.sidebar.header("📂 Fichiers")
mode = st.sidebar.radio(
    "Mode de chargement",
    ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"],
    index=(0 if LAST["mode"] == "two" else 1),
    key=skey("mode"),
)
LAST["mode"] = "two" if mode.startswith("Deux") else "single"

if LAST["mode"] == "two":
    up_c = st.sidebar.file_uploader("Clients (xlsx)", type=["xlsx"], key=skey("up_clients"))
    up_v = st.sidebar.file_uploader("Visa (xlsx)", type=["xlsx"], key=skey("up_visa"))

    if up_c is not None:
        LAST["clients_name"] = up_c.name
        LAST["clients_bytes"] = up_c.read()
    if up_v is not None:
        LAST["visa_name"] = up_v.name
        LAST["visa_bytes"] = up_v.read()

    if st.sidebar.button("💾 Conserver comme derniers fichiers", key=skey("save_last_two")):
        st.success("Derniers fichiers mémorisés (Clients & Visa).")

else:
    up_s = st.sidebar.file_uploader("Fichier unique (2 onglets : Clients & Visa)", type=["xlsx"], key=skey("up_single"))
    if up_s is not None:
        LAST["single_name"] = up_s.name
        LAST["single_bytes"] = up_s.read()
    if st.sidebar.button("💾 Conserver comme dernier fichier unique", key=skey("save_last_single")):
        st.success("Dernier fichier unique mémorisé.")

# ---------------- Sélection de la source à lire ----------------
if LAST["mode"] == "single":
    clients_source = BytesIO(LAST["single_bytes"]) if LAST["single_bytes"] else None
    visa_source    = BytesIO(LAST["single_bytes"]) if LAST["single_bytes"] else None
else:
    clients_source = BytesIO(LAST["clients_bytes"]) if LAST["clients_bytes"] else None
    visa_source    = BytesIO(LAST["visa_bytes"]) if LAST["visa_bytes"] else None

# UI info fichiers
st.markdown("# 🛂 Visa Manager")
with st.expander("📄 Fichiers en mémoire", expanded=False):
    st.write({
        "mode": LAST["mode"],
        "clients": LAST["clients_name"] if LAST["clients_bytes"] else "(aucun)",
        "visa": LAST["visa_name"] if LAST["visa_bytes"] else "(aucun)",
        "single": LAST["single_name"] if LAST["single_bytes"] else "(aucun)",
    })

# ---------------- Lecture DataFrames ----------------
df_clients = _read_clients(clients_source) if clients_source else pd.DataFrame()
df_visa_raw = read_visa_raw(visa_source) if visa_source else pd.DataFrame()

# ---------------- Construction de la carte Visa (catégorie -> sous-catégorie -> options) ----------------
def build_visa_map(dfv: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    dfv : colonnes
      - 'Categorie' (sans accent)
      - 'Sous-categorie' (sans accent)
      - d'autres colonnes optionnelles (ex: COS, EOS, ...); une valeur '1' indique que l'option est disponible.

    Retour:
    {
      "Affaires/Tourisme": {
         "B-1": {"options": ["COS","EOS"], "exclusive": ["COS","EOS"]},
         "B-2": {"options": ["COS","EOS"], "exclusive": ["COS","EOS"]},
      },
      "Etudiants": {
         "F-1": {"options": ["COS","EOS"], "exclusive": ["COS","EOS"]},
         ...
      }
    }
    """
    out: Dict[str, Dict[str, Dict[str, Any]]] = {}

    if dfv.empty:
        return out

    # Normaliser noms (pas d’accents, espaces conservés)
    def clean_header(s: str) -> str:
        s = _safe_str(s)
        # on supprime *seulement* les caractères manifestement non alpha-num, on garde +/_- et espace
        s = re.sub(r"[^A-Za-z0-9+/_\- ]+", " ", s)
        return re.sub(r"\s+", " ", s).strip()

    dfv = dfv.rename(columns={c: clean_header(c) for c in dfv.columns})

    if "Categorie" not in dfv.columns or "Sous-categorie" not in dfv.columns:
        return out

    # Colonnes options = toutes les colonnes autres que Catégorie & Sous-catégorie
    option_cols = [c for c in dfv.columns if c not in ["Categorie", "Sous-categorie"]]

    for _, row in dfv.iterrows():
        cat = _safe_str(row.get("Categorie", "")).strip()
        sub = _safe_str(row.get("Sous-categorie", "")).strip()
        if not cat or not sub:
            continue
        if cat not in out:
            out[cat] = {}
        if sub not in out[cat]:
            out[cat][sub] = {"options": [], "exclusive": []}

        # Options disponibles (valeur == 1)
        opts = []
        for oc in option_cols:
            val = row.get(oc, "")
            try:
                is_one = (str(val).strip() == "1") or (float(val) == 1.0)
            except Exception:
                is_one = False
            if is_one:
                label = oc.strip()
                if label and label not in opts:
                    opts.append(label)

        # Détecter le couple exclusif COS/EOS si présent
        exclusive = []
        if "COS" in opts and "EOS" in opts:
            exclusive = ["COS", "EOS"]

        out[cat][sub]["options"] = opts
        out[cat][sub]["exclusive"] = exclusive

    return out

visa_map: Dict[str, Dict[str, Dict[str, Any]]] = build_visa_map(df_visa_raw.copy()) if not df_visa_raw.empty else {}

# ---------------- UI helper : sélection d’options par sous-catégorie ----------------
def build_visa_option_selector(
    vmap: Dict[str, Dict[str, Dict[str, Any]]],
    cat: str,
    sub: str,
    keyprefix: str,
    preselected: Dict[str, Any] | None = None
) -> Tuple[str, Dict[str, Any], str]:
    """
    Affiche dynamiquement les options liées à (cat, sub).
    Retourne:
      - visa_final (ex: "B-1 COS")
      - opts_dict ({"exclusive": "COS"/"EOS"/None, "options": [autres checkés]})
      - info_msg (texte complémentaire si utile)
    """
    preselected = preselected or {}
    info_msg = ""
    if cat not in vmap or sub not in vmap[cat]:
        return sub, {"exclusive": None, "options": []}, info_msg

    block = vmap[cat][sub]
    opts = block.get("options", [])
    excl = block.get("exclusive", [])

    chosen_excl = None
    others: List[str] = []

    # Exclusif COS/EOS => radio
    if len(excl) == 2 and all(x in opts for x in excl):
        preset = preselected.get("exclusive")
        if preset not in excl:
            preset = excl[0]
        chosen_excl = st.radio("Choix exclusif", excl, index=excl.index(preset), key=skey(keyprefix, "excl"))
    else:
        chosen_excl = None

    # Autres options -> cases à cocher
    non_excl = [o for o in opts if o not in (excl or [])]
    if non_excl:
        st.caption("Options complémentaires :")
        cols = st.columns(min(3, len(non_excl)))
        for i, opt in enumerate(non_excl):
            preset = opt in preselected.get("options", [])
            checked = cols[i % len(cols)].checkbox(opt, value=preset, key=skey(keyprefix, "opt", opt))
            if checked:
                others.append(opt)

    # Libellé final du visa
    visa_final = sub
    if chosen_excl:
        visa_final = f"{sub} {chosen_excl}".strip()

    return visa_final, {"exclusive": chosen_excl, "options": others}, info_msg




# ================================
# 🛂 Visa Manager — PARTIE 2/4
# Tabs, Dashboard, Visa (aperçu)
# ================================

# ---------- Création des onglets principaux ----------
tabs = st.tabs([
    "📊 Dashboard",     # tabs[0]
    "📈 Analyses",      # tabs[1] (rendu en partie 3)
    "🏦 Escrow",        # tabs[2] (rendu en partie 3)
    "👤 Clients",       # tabs[3] (aperçu simple en partie 2, détails en partie 4)
    "🧾 Gestion",       # tabs[4] (CRUD complet en partie 4)
    "📄 Visa (aperçu)", # tabs[5]
])

# ==========================
# 📊 ONGLET : Dashboard
# ==========================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_clients.empty:
        st.info("Aucune donnée client chargée. Charge un fichier dans la barre latérale.")
    else:
        # Préparer listes pour filtres
        # Années, mois, catégories, sous-catégories, visas
        try:
            years = sorted([int(y) for y in pd.to_numeric(df_clients["_Année_"], errors="coerce").dropna().unique().tolist()])
        except Exception:
            years = []
        months = [f"{m:02d}" for m in range(1, 13)]
        cats   = sorted(df_clients.get("Categorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
        subs   = sorted(df_clients.get("Sous-categorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
        visas  = sorted(df_clients.get("Visa", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())

        # Filtres en 5 colonnes
        c1, c2, c3, c4, c5 = st.columns([1,1,1,1,1])
        f_years = c1.multiselect("Année", years, default=[], key=skey("dash","years"))
        f_month = c2.multiselect("Mois (MM)", months, default=[], key=skey("dash","months"))
        f_cat   = c3.multiselect("Catégorie", cats, default=[], key=skey("dash","cats"))
        f_sub   = c4.multiselect("Sous-catégorie", subs, default=[], key=skey("dash","subs"))
        f_visa  = c5.multiselect("Visa", visas, default=[], key=skey("dash","visas"))

        # Application des filtres
        ff = df_clients.copy()
        if f_years: ff = ff[ff["_Année_"].isin(f_years)]
        if f_month: ff = ff[ff["Mois"].astype(str).isin(f_month)]
        if f_cat:   ff = ff[ff["Categorie"].astype(str).isin(f_cat)]
        if f_sub:   ff = ff[ff["Sous-categorie"].astype(str).isin(f_sub)]
        if f_visa:  ff = ff[ff["Visa"].astype(str).isin(f_visa)]

        # KPI (version compacte)
        # Normaliser numériques
        ff["Payé"]  = _safe_num_series(ff, "Payé")
        ff["Reste"] = _safe_num_series(ff, "Reste")
        ff[TOTAL]   = _safe_num_series(ff, TOTAL)
        ff[HONO]    = _safe_num_series(ff, HONO)

        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(ff)}")
        k2.metric("Honoraires", _fmt_money(float(ff[HONO].sum())))
        k3.metric("Total (US $)", _fmt_money(float(ff[TOTAL].sum())))
        k4.metric("Payé", _fmt_money(float(ff["Payé"].sum())))
        k5.metric("Reste", _fmt_money(float(ff["Reste"].sum())))

        # Graphiques
        gcol1, gcol2 = st.columns([1,1])

        # Dossiers par mois (barres)
        with gcol1:
            st.markdown("#### 📦 Dossiers par mois")
            tmp = ff.copy()
            if "Mois" in tmp.columns:
                vc = tmp["Mois"].astype(str).value_counts().reindex(months, fill_value=0)
                st.bar_chart(vc)

        # Honoraires par mois (ligne)
        with gcol2:
            st.markdown("#### 💵 Honoraires par mois")
            if "Mois" in ff.columns:
                gm = ff.groupby("Mois", as_index=False)[HONO].sum()
                # s'assurer d'avoir les 12 mois dans l'ordre
                allm = pd.DataFrame({"Mois": months}).merge(gm, on="Mois", how="left").fillna(0.0)
                st.line_chart(allm.set_index("Mois"))

        # Tableau détaillé
        st.markdown("#### 🧾 Détails")
        view = ff.copy()

        # Formattage affichage
        for col in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if col in view.columns:
                view[col] = _safe_num_series(view, col).map(_fmt_money)

        if "Date" in view.columns:
            try:
                view["Date"] = pd.to_datetime(view["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                view["Date"] = view["Date"].astype(str)

        show_cols = [c for c in [
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
            "Date", "Mois", HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"
        ] if c in view.columns]

        # Tri sûr si colonnes présentes
        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

        st.dataframe(
            view_sorted[show_cols].reset_index(drop=True),
            use_container_width=True,
            key=skey("dash","table")
        )

# ==============================
# 👤 ONGLET : Clients (aperçu)
# ==============================
with tabs[3]:
    st.subheader("👤 Clients — aperçu rapide")
    if df_clients.empty:
        st.info("Aucun client chargé.")
    else:
        # Petit aperçu synthétique
        base_cols = [c for c in [
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
            "Date", "Mois", TOTAL, "Payé", "Reste"
        ] if c in df_clients.columns]

        dfv = df_clients.copy()
        dfv["Payé"]  = _safe_num_series(dfv, "Payé")
        dfv["Reste"] = _safe_num_series(dfv, "Reste")
        dfv[TOTAL]   = _safe_num_series(dfv, TOTAL)

        st.dataframe(
            dfv[base_cols].reset_index(drop=True),
            use_container_width=True,
            key=skey("clients","preview")
        )

# ==============================
# 📄 ONGLET : Visa (aperçu)
# ==============================
with tabs[5]:
    st.subheader("📄 Visa — aperçu structure")

    if not visa_map:
        st.info("Aucune structure Visa chargée.")
    else:
        # Sélecteurs cascade
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=skey("visa","cat"))
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=skey("visa","sub"))
        else:
            sel_sub = ""

        # Affichage des options détectées + sélecteur dynamique (preview)
        if sel_cat and sel_sub:
            block = visa_map[sel_cat][sel_sub]
            opts = block.get("options", [])
            excl = block.get("exclusive", [])

            st.caption("Options disponibles pour cette sous-catégorie :")
            st.write({"options": opts, "exclusive": excl})

            st.markdown("##### 🔧 Aperçu du sélecteur dynamique")
            visa_final, picked, _ = build_visa_option_selector(
                visa_map, sel_cat, sel_sub, keyprefix="visa_prev", preselected={}
            )
            st.info(f"Visa final prévisualisé : **{visa_final or sel_sub}**")




# ================================
# 🛂 Visa Manager — PARTIE 3/4
# Analyses + Escrow
# ================================

# ==============================================
# 📈 ONGLET : Analyses (filtres, KPI, graphiques, détails)
# ==============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_clients.empty:
        st.info("Aucune donnée client.")
    else:
        # Jeux de valeurs pour filtres
        try:
            yearsA = sorted([int(y) for y in pd.to_numeric(df_clients["_Année_"], errors="coerce").dropna().unique().tolist()])
        except Exception:
            yearsA = []
        monthsA = [f"{m:02d}" for m in range(1, 12 + 1)]
        catsA   = sorted(df_clients.get("Categorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
        subsA   = sorted(df_clients.get("Sous-categorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
        visasA  = sorted(df_clients.get("Visa", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())

        a1, a2, a3, a4, a5 = st.columns([1,1,1,1,1])
        fy = a1.multiselect("Année", yearsA, default=[], key=skey("ana","years"))
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=skey("ana","months"))
        fc = a3.multiselect("Catégorie", catsA, default=[], key=skey("ana","cats"))
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=skey("ana","subs"))
        fv = a5.multiselect("Visa", visasA, default=[], key=skey("ana","visas"))

        dfA = df_clients.copy()
        # normaliser numériques pour KPI/graph
        dfA["Payé"]  = _safe_num_series(dfA, "Payé")
        dfA["Reste"] = _safe_num_series(dfA, "Reste")
        dfA[TOTAL]   = _safe_num_series(dfA, TOTAL)
        dfA[HONO]    = _safe_num_series(dfA, HONO)

        # Application filtres
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPI (compacts)
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(float(dfA[HONO].sum())))
        k3.metric("Total (US $)", _fmt_money(float(dfA[TOTAL].sum())))
        k4.metric("Payé", _fmt_money(float(dfA["Payé"].sum())))
        k5.metric("Reste", _fmt_money(float(dfA["Reste"].sum())))

        # --- Graphiques
        g1, g2 = st.columns([1,1])

        with g1:
            st.markdown("#### 📦 Dossiers par catégorie")
            if not dfA.empty and "Categorie" in dfA.columns:
                vc = dfA["Categorie"].value_counts().sort_index()
                st.bar_chart(vc)

            st.markdown("#### 🧮 % par catégorie (sur dossiers filtrés)")
            if not dfA.empty and "Categorie" in dfA.columns:
                total_n = max(1, len(dfA))
                pc = (dfA["Categorie"].value_counts().sort_index() / total_n * 100).round(1)
                st.dataframe(pc.rename("%").to_frame(), use_container_width=True, key=skey("ana","pcat"))

        with g2:
            st.markdown("#### 💵 Honoraires par mois")
            if not dfA.empty and "Mois" in dfA.columns:
                gm = dfA.groupby("Mois", as_index=False)[HONO].sum()
                # garantir l’ordre 01..12
                gm = pd.DataFrame({"Mois": monthsA}).merge(gm, on="Mois", how="left").fillna(0.0)
                st.line_chart(gm.set_index("Mois")[HONO])

            st.markdown("#### 🧮 % par sous-catégorie")
            if not dfA.empty and "Sous-categorie" in dfA.columns:
                total_n = max(1, len(dfA))
                ps = (dfA["Sous-categorie"].value_counts().sort_index() / total_n * 100).round(1)
                st.dataframe(ps.rename("%").to_frame(), use_container_width=True, key=skey("ana","psub"))

        # --- Détails des dossiers filtrés
        st.markdown("#### 🧾 Détails des dossiers filtrés")
        det = dfA.copy()

        # format monétaires
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).apply(_fmt_money)
        # format date
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
            "Date", "Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Date d'envoi",
            "Dossier accepté", "Date d'acceptation",
            "Dossier refusé", "Date de refus",
            "Dossier annulé", "Date d'annulation",
            "RFE"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det

        st.dataframe(det_sorted[show_cols].reset_index(drop=True),
                     use_container_width=True,
                     key=skey("ana","detail"))

        # --- Comparaison période A vs B (Années/Mois) — volumétries & honoraires
        st.markdown("---")
        st.markdown("### 🔁 Comparaison de périodes (A vs B)")
        ca1, ca2, cb1, cb2 = st.columns(4)
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=skey("cmp","ya"))
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=skey("cmp","ma"))
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=skey("cmp","yb"))
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=skey("cmp","mb"))

        def _slice(df, years_sel, months_sel):
            x = df.copy()
            if years_sel:  x = x[x["_Année_"].isin(years_sel)]
            if months_sel: x = x[x["Mois"].astype(str).isin(months_sel)]
            return x

        A = _slice(df_clients, pa_years, pa_month)
        B = _slice(df_clients, pb_years, pb_month)
        for d in (A, B):
            d["Payé"]  = _safe_num_series(d, "Payé")
            d["Reste"] = _safe_num_series(d, "Reste")
            d[TOTAL]   = _safe_num_series(d, TOTAL)
            d[HONO]    = _safe_num_series(d, HONO)

        cA, cB = st.columns(2)
        with cA:
            st.markdown("#### Période A")
            st.write({
                "Dossiers": len(A),
                "Honoraires": _fmt_money(float(A[HONO].sum())),
                "Total": _fmt_money(float(A[TOTAL].sum())),
                "Payé": _fmt_money(float(A["Payé"].sum())),
                "Reste": _fmt_money(float(A["Reste"].sum())),
            })
        with cB:
            st.markdown("#### Période B")
            st.write({
                "Dossiers": len(B),
                "Honoraires": _fmt_money(float(B[HONO].sum())),
                "Total": _fmt_money(float(B[TOTAL].sum())),
                "Payé": _fmt_money(float(B["Payé"].sum())),
                "Reste": _fmt_money(float(B["Reste"].sum())),
            })

        st.markdown("#### Différence (A - B)")
        st.write({
            "Δ Dossiers": len(A) - len(B),
            "Δ Honoraires": _fmt_money(float(A[HONO].sum() - B[HONO].sum())),
            "Δ Total": _fmt_money(float(A[TOTAL].sum() - B[TOTAL].sum())),
            "Δ Payé": _fmt_money(float(A["Payé"].sum() - B["Payé"].sum())),
            "Δ Reste": _fmt_money(float(A["Reste"].sum() - B["Reste"].sum())),
        })


# ==============================================
# 🏦 ONGLET : Escrow — synthèse simple
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
        dfE[HONO]    = _safe_num_series(dfE, HONO)

        # KPI réduits
        t1, t2, t3 = st.columns([1,1,1])
        t1.metric("Total (US $)", _fmt_money(float(dfE[TOTAL].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Reste", _fmt_money(float(dfE["Reste"].sum())))

        st.markdown("#### Par catégorie")
        agg = dfE.groupby("Categorie", as_index=False)[[TOTAL, "Payé", "Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg[TOTAL]).replace([pd.NA, pd.NaT], 0).fillna(0) * 100
        st.dataframe(agg, use_container_width=True, key=skey("esc","agg_cat"))

        st.markdown("#### Détail clients")
        show_cols = [c for c in [
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Date d'envoi"
        ] if c in dfE.columns]

        st.dataframe(dfE[show_cols].reset_index(drop=True),
                     use_container_width=True,
                     key=skey("esc","detail"))




# ================================
# 🛂 Visa Manager — PARTIE 4/4
# Clients (détaillé) + Gestion CRUD (Ajouter / Modifier / Supprimer)
# ================================

# -------------------------------------------------------
# 👤 ONGLET : Clients — aperçu étendu + recherche simple
# -------------------------------------------------------
with tabs[3]:
    st.subheader("👤 Clients — aperçu détaillé")

    if df_clients.empty:
        st.info("Aucun client chargé.")
    else:
        cc1, cc2, cc3 = st.columns([1,1,2])
        q_name = cc1.text_input("Recherche par nom", "", key=skey("cli","qname"))
        q_id   = cc2.text_input("Recherche par ID_Client", "", key=skey("cli","qid"))

        dfv = df_clients.copy()
        if q_name:
            dfv = dfv[dfv["Nom"].astype(str).str.contains(q_name, case=False, na=False)]
        if q_id:
            dfv = dfv[dfv["ID_Client"].astype(str).str.contains(q_id, case=False, na=False)]

        # Monétaires en affichage formaté
        for col in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if col in dfv.columns:
                dfv[col] = _safe_num_series(dfv, col).map(_fmt_money)

        if "Date" in dfv.columns:
            try:
                dfv["Date"] = pd.to_datetime(dfv["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                dfv["Date"] = dfv["Date"].astype(str)

        show_cols = [c for c in [
            DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
            "Date", "Mois",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Dossier envoyé", "Date d'envoi",
            "Dossier accepté", "Date d'acceptation",
            "Dossier refusé", "Date de refus",
            "Dossier annulé", "Date d'annulation",
            "RFE",
            "Commentaires"
        ] if c in dfv.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in dfv.columns]
        dfv_sorted = dfv.sort_values(by=sort_keys) if sort_keys else dfv

        st.dataframe(
            dfv_sorted[show_cols].reset_index(drop=True),
            use_container_width=True,
            key=skey("cli","table")
        )


# -------------------------------------------------------
# 🧾 GESTION — Ajouter / Modifier / Supprimer
# -------------------------------------------------------
with tabs[4]:
    st.subheader("🧾 Gestion des clients")
    df_live = _read_clients(clients_source) if clients_source else pd.DataFrame()

    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=skey("crud","op"))

    # ---------- utilitaires communs ----------
    def _recompute_finance(row_like: dict) -> Tuple[float, float, float]:
        """Retourne (honoraires, total, payé, reste) à partir du dict, en prenant en compte Paiements."""
        h = float(row_like.get(HONO, 0.0) or 0.0)
        o = float(row_like.get(AUTRE, 0.0) or 0.0)
        t = h + o
        # Paiements : liste de dicts [{"date": "...", "mode": "...", "montant": float}, ...]
        pay_list = row_like.get("Paiements", [])
        paid = 0.0
        if isinstance(pay_list, list):
            for p in pay_list:
                try:
                    paid += float(p.get("montant", 0.0) or 0.0)
                except Exception:
                    pass
        r = max(0.0, t - paid)
        return h, t, paid, r

    def _parse_paiements(raw) -> list:
        if isinstance(raw, list):
            return raw
        if isinstance(raw, str) and raw.strip():
            try:
                v = json.loads(raw)
                return v if isinstance(v, list) else []
            except Exception:
                return []
        return []

    # =====================
    # ➕ AJOUTER UN CLIENT
    # =====================
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")

        a1, a2, a3 = st.columns([1,1,1])
        nom  = a1.text_input("Nom", "", key=skey("add","nom"))
        dval = _date_for_widget(date.today())
        dt   = a2.date_input("Date de création", value=dval, key=skey("add","date"))
        mois = a3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=(dt.month-1 if isinstance(dt, date) else 0), key=skey("add","mois"))

        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=skey("add","cat"))
        sel_sub = ""
        visa_final = ""
        opts_dict = {"exclusive": None, "options": []}
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=skey("add","sub"))
            if sel_sub:
                visa_final, opts_dict, _ = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=skey("add","opts"), preselected={}
                )

        f1, f2 = st.columns(2)
        honor = f1.number_input(HONO, min_value=0.0, value=0.0, step=50.0, format="%.2f", key=skey("add","h"))
        other = f2.number_input(AUTRE, min_value=0.0, value=0.0, step=20.0, format="%.2f", key=skey("add","o"))
        coms  = st.text_area("Commentaires (autres frais / notes)", "", key=skey("add","coms"))

        st.markdown("#### 📌 Statuts initiaux")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent = s1.checkbox("Dossier envoyé", key=skey("add","sent"))
        sent_d = s1.date_input("Date d'envoi", value=None, key=skey("add","sentd"))
        acc = s2.checkbox("Dossier accepté", key=skey("add","acc"))
        acc_d = s2.date_input("Date d'acceptation", value=None, key=skey("add","accd"))
        ref = s3.checkbox("Dossier refusé", key=skey("add","ref"))
        ref_d = s3.date_input("Date de refus", value=None, key=skey("add","refd"))
        ann = s4.checkbox("Dossier annulé", key=skey("add","ann"))
        ann_d = s4.date_input("Date d'annulation", value=None, key=skey("add","annd"))
        rfe = s5.checkbox("RFE", key=skey("add","rfe"))
        if rfe and not any([sent, acc, ref, ann]):
            st.warning("⚠️ La case RFE doit être associée à un autre statut (envoyé/accepté/refusé/annulé).")

        if st.button("💾 Enregistrer le client", key=skey("add","save")):
            if not nom:
                st.warning("Veuillez saisir le nom.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Veuillez choisir la catégorie et la sous-catégorie.")
                st.stop()

            did = _make_client_id(nom, dt if isinstance(dt, date) else date.today())
            dossier_n = _next_dossier(df_live, start=13057)

            new_row = {
                DOSSIER_COL: dossier_n,
                "ID_Client": did,
                "Nom": nom,
                "Date": dt,
                "Mois": f"{int(mois):02d}" if isinstance(mois, (int,str)) else "",
                "Categorie": sel_cat,
                "Sous-categorie": sel_sub,
                "Visa": visa_final or sel_sub,
                HONO: float(honor),
                AUTRE: float(other),
                TOTAL: float(honor) + float(other),
                "Payé": 0.0,
                "Reste": float(honor) + float(other),
                "Paiements": [],
                "Options": opts_dict,
                "Dossier envoyé": 1 if sent else 0,
                "Date d'envoi": sent_d if sent_d else (dt if sent else None),
                "Dossier accepté": 1 if acc else 0,
                "Date d'acceptation": acc_d,
                "Dossier refusé": 1 if ref else 0,
                "Date de refus": ref_d,
                "Dossier annulé": 1 if ann else 0,
                "Date d'annulation": ann_d,
                "RFE": 1 if rfe else 0,
                "Commentaires": coms,
            }
            df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
            _write_clients(df_new, clients_source)
            st.success("Client ajouté.")
            st.cache_data.clear()
            st.rerun()

    # =====================
    # ✏️ MODIFIER UN CLIENT
    # =====================
    elif op == "Modifier":
        st.markdown("### ✏️ Modifier un client")

        if df_live.empty:
            st.info("Aucun client à modifier.")
        else:
            names = sorted(df_live.get("Nom", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
            ids   = sorted(df_live.get("ID_Client", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())

            m1, m2 = st.columns(2)
            target_name = m1.selectbox("Nom", [""] + names, index=0, key=skey("mod","selname"))
            target_id   = m2.selectbox("ID_Client", [""] + ids, index=0, key=skey("mod","selid"))

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is None or not mask.any():
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()

            d1, d2, d3 = st.columns([1,1,1])
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=skey("mod","nom"))
            dval = _date_for_widget(row.get("Date"))
            dt   = d2.date_input("Date de création", value=dval, key=skey("mod","date"))
            mois_default = _safe_str(row.get("Mois","01"))
            try:
                mois_idx = max(0, min(11, int(mois_default) - 1))
            except Exception:
                mois_idx = 0
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=mois_idx, key=skey("mod","mois"))

            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""] + cats, index=(cats.index(preset_cat)+1 if preset_cat in cats else 0), key=skey("mod","cat"))

            subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
            preset_sub = _safe_str(row.get("Sous-categorie",""))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=(subs.index(preset_sub)+1 if preset_sub in subs else 0), key=skey("mod","sub"))

            # options déjà enregistrées
            preset_opts = row.get("Options", {})
            if not isinstance(preset_opts, dict):
                try:
                    preset_opts = json.loads(_safe_str(preset_opts) or "{}")
                    if not isinstance(preset_opts, dict):
                        preset_opts = {}
                except Exception:
                    preset_opts = {}

            visa_final, opts_dict, _ = "", {"exclusive": None, "options": []}, ""
            if sel_cat and sel_sub:
                visa_final, opts_dict, _ = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=skey("mod","opts"), preselected=preset_opts
                )

            f1, f2 = st.columns(2)
            honor = f1.number_input(HONO, min_value=0.0, value=float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0]), step=50.0, format="%.2f", key=skey("mod","h"))
            other = f2.number_input(AUTRE, min_value=0.0, value=float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0]), step=20.0, format="%.2f", key=skey("mod","o"))
            coms  = st.text_area("Commentaires (autres frais / notes)", _safe_str(row.get("Commentaires","")), key=skey("mod","coms"))

            st.markdown("#### 📌 Statuts")
            s1, s2, s3, s4, s5 = st.columns(5)
            envoye  = s1.checkbox("Dossier envoyé", value=_to_int_bool(row.get("Dossier envoyé",0)), key=skey("mod","sent"))
            sent_d  = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=skey("mod","sentd"))
            accepte = s2.checkbox("Dossier accepté", value=_to_int_bool(row.get("Dossier accepté",0)), key=skey("mod","acc"))
            acc_d   = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=skey("mod","accd"))
            refuse  = s3.checkbox("Dossier refusé", value=_to_int_bool(row.get("Dossier refusé",0)), key=skey("mod","ref"))
            ref_d   = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=skey("mod","refd"))
            annule  = s4.checkbox("Dossier annulé", value=_to_int_bool(row.get("Dossier annulé",0)), key=skey("mod","ann"))
            ann_d   = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=skey("mod","annd"))
            rfe     = s5.checkbox("RFE", value=_to_int_bool(row.get("RFE",0)), key=skey("mod","rfe"))
            if rfe and not any([envoye, accepte, refuse, annule]):
                st.warning("⚠️ RFE doit être associé à un statut (envoyé/accepté/refusé/annulé).")

            # Bloc paiements (consultation + ajout si non soldé)
            st.markdown("#### 💳 Paiements")
            pay_list = _parse_paiements(row.get("Paiements"))

            # tableau paiements
            if pay_list:
                dfp = pd.DataFrame(pay_list)
                # normaliser colonnes
                if "date" in dfp.columns:
                    try:
                        dfp["date"] = pd.to_datetime(dfp["date"], errors="coerce").dt.date.astype(str)
                    except Exception:
                        dfp["date"] = dfp["date"].astype(str)
                if "montant" in dfp.columns:
                    dfp["montant"] = pd.to_numeric(dfp["montant"], errors="coerce").fillna(0.0)
                st.dataframe(dfp.rename(columns={"date":"Date", "mode":"Mode", "montant":"Montant"}), use_container_width=True, key=skey("mod","pays"))

            # Ajout règlement si reste > 0
            h, t, paid, rest = _recompute_finance({
                HONO: honor, AUTRE: other, "Paiements": pay_list
            })
            st.caption(f"Total: {_fmt_money(t)} — Payé: {_fmt_money(paid)} — Reste: {_fmt_money(rest)}")

            if rest > 0:
                r1, r2, r3 = st.columns([1,1,1])
                pay_date = r1.date_input("Date paiement", value=_date_for_widget(date.today()), key=skey("mod","paydate"))
                pay_mode = r2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=skey("mod","paymode"))
                pay_amt  = r3.number_input("Montant", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=skey("mod","payamt"))
                if st.button("➕ Ajouter ce paiement", key=skey("mod","paybtn")):
                    if pay_amt <= 0:
                        st.warning("Montant de paiement invalide.")
                        st.stop()
                    pay_list.append({"date": str(pay_date) if isinstance(pay_date, date) else _safe_str(pay_date),
                                     "mode": pay_mode, "montant": float(pay_amt)})
                    # recalcule
                    _, t2, paid2, rest2 = _recompute_finance({HONO: honor, AUTRE: other, "Paiements": pay_list})
                    # injecte dans la ligne avant sauvegarde finale
                    row["Paiements"] = pay_list
                    row["Payé"] = paid2
                    row["Reste"] = rest2
                    # reflète dans df_live (sans encore tout enregistrer)
                    df_live.at[idx, "Paiements"] = json.dumps(pay_list, ensure_ascii=False)
                    df_live.at[idx, "Payé"] = paid2
                    df_live.at[idx, "Reste"] = rest2
                    _write_clients(df_live, clients_source)
                    st.success("Paiement ajouté.")
                    st.cache_data.clear()
                    st.rerun()

            if st.button("💾 Enregistrer les modifications", key=skey("mod","save")):
                if not nom:
                    st.warning("Le nom est requis.")
                    st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisissez la catégorie et la sous-catégorie.")
                    st.stop()

                # recalcul finance avec paiements existants
                h, t, paid, rest = _recompute_finance({HONO: honor, AUTRE: other, "Paiements": _parse_paiements(row.get("Paiements"))})

                df_live.at[idx, "Nom"] = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(mois):02d}" if isinstance(mois,(int,str)) else ""
                df_live.at[idx, "Categorie"] = sel_cat
                df_live.at[idx, "Sous-categorie"] = sel_sub
                df_live.at[idx, "Visa"] = (visa_final or sel_sub)
                df_live.at[idx, HONO] = float(honor)
                df_live.at[idx, AUTRE] = float(other)
                df_live.at[idx, TOTAL] = float(t)
                df_live.at[idx, "Payé"] = float(paid)
                df_live.at[idx, "Reste"] = float(rest)
                df_live.at[idx, "Options"] = opts_dict
                df_live.at[idx, "Commentaires"] = coms

                df_live.at[idx, "Dossier envoyé"] = 1 if envoye else 0
                df_live.at[idx, "Date d'envoi"] = sent_d
                df_live.at[idx, "Dossier accepté"] = 1 if accepte else 0
                df_live.at[idx, "Date d'acceptation"] = acc_d
                df_live.at[idx, "Dossier refusé"] = 1 if refuse else 0
                df_live.at[idx, "Date de refus"] = ref_d
                df_live.at[idx, "Dossier annulé"] = 1 if annule else 0
                df_live.at[idx, "Date d'annulation"] = ann_d
                df_live.at[idx, "RFE"] = 1 if rfe else 0

                _write_clients(df_live, clients_source)
                st.success("Modifications enregistrées.")
                st.cache_data.clear()
                st.rerun()

    # =====================
    # 🗑️ SUPPRIMER UN CLIENT
    # =====================
    elif op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client à supprimer.")
        else:
            names = sorted(df_live.get("Nom", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
            ids   = sorted(df_live.get("ID_Client", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())

            d1, d2 = st.columns(2)
            target_name = d1.selectbox("Nom", [""] + names, index=0, key=skey("del","name"))
            target_id   = d2.selectbox("ID_Client", [""] + ids, index=0, key=skey("del","id"))

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({
                    "Dossier N": row.get(DOSSIER_COL, ""),
                    "Nom": row.get("Nom", ""),
                    "Visa": row.get("Visa", ""),
                    "Montant": _fmt_money(float(_safe_num_series(pd.DataFrame([row]), TOTAL).iloc[0] if TOTAL in row else 0.0))
                })
                if st.button("❗ Confirmer la suppression", key=skey("del","confirm")):
                    df_new = df_live[~mask].copy()
                    _write_clients(df_new, clients_source)
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()