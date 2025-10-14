# ================================
# 🛂 Visa Manager — PARTIE 1/4
# Imports, constantes, helpers, I/O, parsing Visa
# ================================
from __future__ import annotations

import json, re, zipfile, uuid
from io import BytesIO
from datetime import date, datetime
from typing import Dict, Any, List, Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st

# ----------------
# Config de page
# ----------------
st.set_page_config(page_title="Visa Manager", page_icon="🛂", layout="wide")

# ----------------
# Constantes
# ----------------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"
DOSSIER_COL = "Dossier N"

# Clés session (générateur)
SID = st.session_state.get("_sid") or str(uuid.uuid4())[:8]
st.session_state["_sid"] = SID
def skey(*parts: str) -> str:
    return "k_" + SID + "_" + "_".join(parts)

# ----------------
# Helpers généraux
# ----------------
def _safe_str(x: Any) -> str:
    try:
        if x is None:
            return ""
        return str(x)
    except Exception:
        return ""

def _norm(s: str) -> str:
    """Normalise pour matching (sans accents supposés) en minuscules."""
    s = _safe_str(s).strip().lower()
    # attention: pas de classes invalides -> échappements simples
    s = re.sub(r"[^a-z0-9+/_\- ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _to_num(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series([], dtype=float)
    ss = s.astype(str)
    # Retire les $ espaces et remplace virgules françaises si besoin
    ss = ss.str.replace(r"[^\d,.\-]", "", regex=True).str.replace(",", ".", regex=False)
    return pd.to_numeric(ss, errors="coerce").fillna(0.0)

def _safe_num_series(df_or_series: Any, col: str) -> pd.Series:
    """Retourne une Series numérique sûre (0.0 si col absente)."""
    try:
        if isinstance(df_or_series, pd.DataFrame):
            if col not in df_or_series.columns:
                return pd.Series([0.0] * len(df_or_series), index=df_or_series.index, dtype=float)
            return _to_num(df_or_series[col])
        elif isinstance(df_or_series, pd.Series):
            return _to_num(df_or_series)
        else:
            return pd.Series([], dtype=float)
    except Exception:
        return pd.Series([], dtype=float)

def _fmt_money(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(val):
    """Date sûre pour st.date_input (ou None)."""
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        t = pd.to_datetime(val, errors="coerce")
        return t.date() if pd.notna(t) else None
    except Exception:
        return None

def _to_iso_date(v):
    """Convertit v en 'YYYY-MM-DD' (fallback = today)."""
    if isinstance(v, date) and not isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime):
        return v.date().strftime("%Y-%m-%d")
    try:
        t = pd.to_datetime(v, errors="coerce")
        if pd.notna(t):
            return t.date().strftime("%Y-%m-%d")
    except Exception:
        pass
    return date.today().strftime("%Y-%m-%d")

# ----------------
# Mémoire dernier chargement (session)
# ----------------
def _save_last(name: str, data: bytes):
    st.session_state[f"last_{name}_bytes"] = data

def _load_last(name: str) -> Optional[bytes]:
    return st.session_state.get(f"last_{name}_bytes")

# ----------------
# Lecture / Écriture fichiers
# ----------------
@st.cache_data(show_spinner=False)
def read_excel_bytes(xls_bytes: bytes, sheet: Optional[str] = None) -> Dict[str, pd.DataFrame]:
    """Lit un XLSX depuis bytes. Retourne dict sheet_name -> DataFrame.
    Si 'sheet' est fourni, on ne renvoie que cette feuille sous ce nom.
    """
    bio = BytesIO(xls_bytes)
    xls = pd.ExcelFile(bio)
    if sheet:
        return {sheet: pd.read_excel(xls, sheet_name=sheet)}
    return {sn: pd.read_excel(xls, sheet_name=sn) for sn in xls.sheet_names}

def write_clients_to_bytes(df: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
    return bio.getvalue()

def write_two_sheets_to_bytes(df_clients: pd.DataFrame, df_visa: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        df_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        df_visa.to_excel(wr, sheet_name=SHEET_VISA, index=False)
    return bio.getvalue()

# ----------------
# Normalisation Clients
# ----------------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
            "Categorie", "Sous-categorie", "Visa",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Paiements", "Options",
            "Dossier envoyé", "Date d'envoi",
            "Dossier accepté", "Date d'acceptation",
            "Dossier refusé", "Date de refus",
            "Dossier annulé", "Date d'annulation",
            "RFE", "Commentaire"
        ])

    # Colonnes minimales
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c not in df.columns:
            df[c] = 0.0

    # Numérise et recalcule
    df[HONO]  = _safe_num_series(df, HONO)
    df[AUTRE] = _safe_num_series(df, AUTRE)
    df["Payé"] = _safe_num_series(df, "Payé")

    # Total
    if TOTAL not in df.columns or df[TOTAL].fillna(0).sum() == 0:
        df[TOTAL] = df[HONO] + df[AUTRE]

    # Reste
    if "Reste" not in df.columns or df["Reste"].isna().any():
        df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # Date/Mois/Année
    if "Date" in df.columns:
        dts = pd.to_datetime(df["Date"], errors="coerce")
    else:
        dts = pd.Series([pd.NaT] * len(df))
        df["Date"] = None

    df["Mois"] = df.get("Mois", pd.Series([None]*len(df))).astype(str)
    # si mois vide, tirer du champ Date
    for i, v in enumerate(df["Mois"]):
        if not _safe_str(v):
            try:
                mm = dts.iloc[i].month if pd.notna(dts.iloc[i]) else None
                df.at[i, "Mois"] = f"{mm:02d}" if mm else ""
            except Exception:
                df.at[i, "Mois"] = ""

    df["_Année_"]   = dts.dt.year.astype("Int64")
    df["_MoisNum_"] = dts.dt.month.astype("Int64")

    # Garanti les colonnes statut / options
    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        if c not in df.columns:
            df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    for c in ["Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation"]:
        if c not in df.columns:
            df[c] = None

    if "Paiements" not in df.columns:
        df["Paiements"] = [[] for _ in range(len(df))]
    if "Options" not in df.columns:
        df["Options"] = [{} for _ in range(len(df))]
    if "Commentaire" not in df.columns:
        df["Commentaire"] = ""

    # ID/Dossier
    if DOSSIER_COL not in df.columns:
        df[DOSSIER_COL] = None
    if "ID_Client" not in df.columns:
        df["ID_Client"] = None

    # Ordre conseillé (sans casser si manques)
    ordered = [c for c in [
        DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
        "Categorie", "Sous-categorie", "Visa",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Paiements", "Options", "Commentaire",
        "Dossier envoyé", "Date d'envoi",
        "Dossier accepté", "Date d'acceptation",
        "Dossier refusé", "Date de refus",
        "Dossier annulé", "Date d'annulation",
        "RFE", "_Année_", "_MoisNum_"
    ] if c in df.columns] + [c for c in df.columns if c not in {
        DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
        "Categorie", "Sous-categorie", "Visa",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Paiements", "Options", "Commentaire",
        "Dossier envoyé", "Date d'envoi",
        "Dossier accepté", "Date d'acceptation",
        "Dossier refusé", "Date de refus",
        "Dossier annulé", "Date d'annulation",
        "RFE", "_Année_", "_MoisNum_"
    }]

    return df[ordered]

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if df is None or df.empty or DOSSIER_COL not in df.columns:
        return int(start)
    vals = pd.to_numeric(df[DOSSIER_COL], errors="coerce").dropna()
    return int((vals.max() if len(vals) else (start - 1)) + 1)

def _make_client_id(nom: str, dt: date | datetime | None) -> str:
    base = _norm(nom).replace(" ", "")
    if not base:
        base = "client"
    if isinstance(dt, datetime):
        d = dt.date()
    elif isinstance(dt, date):
        d = dt
    else:
        d = date.today()
    return f"{base}-{d:%Y%m%d}"

# ----------------
# Parsing structure Visa (df_visa_raw)
# ----------------
def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """Construit un mapping à partir de l’onglet Visa :
       {
         "Categorie" : {
            "Sous-categorie": {
               "visa_label": "...",
               "options": {
                  "exclusive": "NomSuite" ou None,
                  "options": [{"title":"COS","key":"cos"}, {"title":"EOS","key":"eos"}, ...]
               }
            }
         }
       }
       Règles :
       - On lit les colonnes 'Catégorie' et 'Sous-categories 1' (ou 'Sous-categorie' si déjà présent)
       - Les colonnes supplémentaires (en-tête ligne 1) définissent les options :
           * si l’en-tête commence par '(...)' => groupe exclusif (radio)
           * si l’en-tête contient '[X]' => cases à cocher (checkbox)
           * autres colonnes ignorées
       - Une cellule == 1 (ou '1' / True) => l’option est proposée pour la ligne (Sous-catégorie) concernée.
    """
    df = df_visa.copy()
    # Harmonisation nom colonnes
    rename_map = {}
    for c in df.columns:
        cn = _norm(c)
        if cn == "categorie":
            rename_map[c] = "Categorie"
        elif cn in ("sous-categories1", "sous-categorie", "sous-categories_1"):
            rename_map[c] = "Sous-categorie"
    if rename_map:
        df = df.rename(columns=rename_map)

    if "Categorie" not in df.columns or "Sous-categorie" not in df.columns:
        # tente sur "Sous-categories 1"
        cand = [c for c in df.columns if _norm(c) == "sous-categories1"]
        if cand:
            df = df.rename(columns={cand[0]: "Sous-categorie"})
    if "Categorie" not in df.columns or "Sous-categorie" not in df.columns:
        return {}

    # Détecte colonnes d’options (ligne d’entête)
    cols = [c for c in df.columns if c not in ["Categorie", "Sous-categorie"]]
    option_cols = []
    for c in cols:
        header = _safe_str(c).strip()
        if header.startswith("(") and header.endswith(")"):
            # groupe exclusif
            option_cols.append(("exclusive", header.strip("()")))
        elif "[X]" in header or header.startswith("[") and header.endswith("]"):
            # checkbox
            title = header.replace("[X]", "").replace("[x]", "").replace("[", "").replace("]", "").strip()
            option_cols.append(("checkbox", title))
        else:
            # colonne lambda (ignorer)
            pass

    # Construit map
    out: Dict[str, Dict[str, Dict[str, Any]]] = {}
    for _, row in df.iterrows():
        cat = _safe_str(row.get("Categorie", "")).strip()
        sub = _safe_str(row.get("Sous-categorie", "")).strip()
        if not cat or not sub:
            continue

        cat_map = out.setdefault(cat, {})
        if sub not in cat_map:
            cat_map[sub] = {"visa_label": sub, "options": {"exclusive": None, "options": []}}

        # Parcours des colonnes options : une valeur == 1 => active l’option
        for kind, title in option_cols:
            if title not in df.columns:
                # (cas improbable si titre = entête lui-même)
                continue
            val = row.get(title, None)
            active = False
            if isinstance(val, (int, float)) and not pd.isna(val):
                active = (int(val) == 1)
            elif _safe_str(val) == "1":
                active = True
            if active:
                if kind == "exclusive":
                    cat_map[sub]["options"]["exclusive"] = title
                elif kind == "checkbox":
                    cat_map[sub]["options"]["options"].append({
                        "title": title,
                        "key": _norm(title).replace(" ", "_")
                    })

    return out



# ================================
# 🛂 Visa Manager — PARTIE 2/4
# Fichiers (upload, mémoire), visa_map, onglets, sélecteur d'options
# ================================

st.markdown("## 📂 Fichiers")

# ----------------
# Zone de chargement
# ----------------
mode = st.radio(
    "Mode de chargement",
    ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"],
    horizontal=True,
    key=skey("load", "mode"),
)

cA, cB = st.columns(2)

clients_bytes = st.session_state.get("clients_bytes")
visa_bytes    = st.session_state.get("visa_bytes")

if mode == "Deux fichiers (Clients & Visa)":
    with cA:
        up_clients = st.file_uploader("Clients (xlsx)", type=["xlsx"], key=skey("up", "clients"))
        if up_clients is not None:
            clients_bytes = up_clients.read()
            _save_last("clients", clients_bytes)
            st.session_state["clients_bytes"] = clients_bytes
            st.success("Feuille Clients chargée.")

    with cB:
        up_visa = st.file_uploader("Visa (xlsx)", type=["xlsx"], key=skey("up", "visa"))
        if up_visa is not None:
            visa_bytes = up_visa.read()
            _save_last("visa", visa_bytes)
            st.session_state["visa_bytes"] = visa_bytes
            st.success("Feuille Visa chargée.")

else:
    up_both = st.file_uploader("Un seul fichier (2 onglets: Clients & Visa)", type=["xlsx"], key=skey("up", "both"))
    if up_both is not None:
        both_bytes = up_both.read()
        try:
            sheets = read_excel_bytes(both_bytes)
            # Clients
            if SHEET_CLIENTS in sheets:
                clients_df_tmp = sheets[SHEET_CLIENTS]
            else:
                # premier onglet par défaut
                first_sn = list(sheets.keys())[0] if sheets else SHEET_CLIENTS
                clients_df_tmp = sheets.get(first_sn, pd.DataFrame())
            # Visa
            if SHEET_VISA in sheets:
                visa_df_tmp = sheets[SHEET_VISA]
            else:
                # tente un onglet "Visa" sinon vide
                visa_df_tmp = sheets.get(SHEET_VISA, pd.DataFrame())

            # Stocke mémoire bytes séparés
            st.session_state["clients_bytes"] = write_clients_to_bytes(clients_df_tmp)
            st.session_state["visa_bytes"]    = write_two_sheets_to_bytes(pd.DataFrame(), visa_df_tmp)  # on encapsule VISA seul
            _save_last("clients", st.session_state["clients_bytes"])
            _save_last("visa",    st.session_state["visa_bytes"])
            clients_bytes = st.session_state["clients_bytes"]
            visa_bytes    = st.session_state["visa_bytes"]
            st.success("Fichier multi-onglets chargé (Clients & Visa).")
        except Exception as e:
            st.error("Lecture impossible : " + _safe_str(e))

# Boutons derniers fichiers mémorisés
cL1, cL2, cL3 = st.columns([1,1,2])
with cL1:
    if st.button("↩️ Recharger dernier Clients", key=skey("btn", "lastC")):
        last = _load_last("clients")
        if last:
            st.session_state["clients_bytes"] = last
            clients_bytes = last
            st.success("Dernier Clients rechargé.")
        else:
            st.info("Aucun Clients mémorisé.")
with cL2:
    if st.button("↩️ Recharger dernier Visa", key=skey("btn", "lastV")):
        last = _load_last("visa")
        if last:
            st.session_state["visa_bytes"] = last
            visa_bytes = last
            st.success("Dernier Visa rechargé.")
        else:
            st.info("Aucun Visa mémorisé.")

st.markdown("---")

# ----------------
# Lecture des DataFrames depuis les bytes en mémoire
# ----------------
def _read_clients_from_state() -> pd.DataFrame:
    b = st.session_state.get("clients_bytes")
    if not b:
        return pd.DataFrame()
    try:
        d = read_excel_bytes(b)
        # si le fichier ne contient qu'une feuille, on prend la première
        if SHEET_CLIENTS in d:
            df = d[SHEET_CLIENTS]
        else:
            first = list(d.values())[0]
            df = first
        return normalize_clients(df.copy())
    except Exception:
        return pd.DataFrame()

def _read_visa_from_state() -> pd.DataFrame:
    b = st.session_state.get("visa_bytes")
    if not b:
        return pd.DataFrame()
    try:
        d = read_excel_bytes(b)
        if SHEET_VISA in d:
            return d[SHEET_VISA].copy()
        else:
            # si c'était stocké seul, il peut être dans la première feuille
            first = list(d.values())[0]
            return first.copy()
    except Exception:
        return pd.DataFrame()

def _write_clients_to_state(df_clients: pd.DataFrame):
    bytes_out = write_clients_to_bytes(df_clients)
    st.session_state["clients_bytes"] = bytes_out
    _save_last("clients", bytes_out)

def _write_two_to_state(df_clients: pd.DataFrame, df_visa: pd.DataFrame):
    bytes_out = write_two_sheets_to_bytes(df_clients, df_visa)
    # on duplique pour rester compatible
    st.session_state["clients_bytes"] = write_clients_to_bytes(df_clients)
    st.session_state["visa_bytes"]    = write_two_sheets_to_bytes(pd.DataFrame(), df_visa)
    _save_last("clients", st.session_state["clients_bytes"])
    _save_last("visa",    st.session_state["visa_bytes"])

# Expose alias cohérents pour les autres parties
_read_clients  = _read_clients_from_state
_write_clients = _write_clients_to_state

# ----------------
# Construit les DataFrames de travail
# ----------------
df_clients_raw = _read_clients_from_state()
df_visa_raw    = _read_visa_from_state()

df_all = normalize_clients(df_clients_raw.copy()) if not df_clients_raw.empty else pd.DataFrame(columns=[
    DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
    "Categorie", "Sous-categorie", "Visa",
    HONO, AUTRE, TOTAL, "Payé", "Reste",
    "Paiements", "Options", "Commentaire",
    "Dossier envoyé", "Date d'envoi",
    "Dossier accepté", "Date d'acceptation",
    "Dossier refusé", "Date de refus",
    "Dossier annulé", "Date d'annulation",
    "RFE", "_Année_", "_MoisNum_"
])

# Construire la structure des visas
visa_map = build_visa_map(df_visa_raw.copy()) if not df_visa_raw.empty else {}

# ----------------
# Aperçu Visa (facultatif rapide)
# ----------------
with st.expander("📄 Aperçu rapide de la structure Visa (débogage)", expanded=False):
    if visa_map:
        st.json(visa_map)
    else:
        st.caption("Aucune structure Visa chargée.")

# ----------------
# Création des onglets (ordre figé)
# ----------------
tabs = st.tabs([
    "📊 Dashboard",   # tabs[0]
    "📈 Analyses",    # tabs[1]
    "🏦 Escrow",      # tabs[2]
    "👤 Compte client", # tabs[3]
    "🧾 Gestion",     # tabs[4]
    "📄 Visa (aperçu)"# tabs[5]
])

# ----------------
# UI sélecteur d’options (COS/EOS etc.) en fonction de visa_map
# ----------------
def build_visa_option_selector(visa_map: Dict[str, Any], cat: str, sub: str,
                               keyprefix: str, preselected: Dict[str, Any] | None = None
                               ) -> Tuple[str, Dict[str, Any], str]:
    """Affiche les options (checkbox/radio) en fonction de la catégorie & sous-catégorie.
       Retourne (visa_label, options_dict, info_msg).
       options_dict = {"exclusive": <val ou None>, "options":[{"title":..., "key":..., "checked": bool}, ...]}
       Les titres des options viennent de la ligne d'entête du fichier Visa (cases cochables).
    """
    info = ""
    if not visa_map or cat not in visa_map or sub not in visa_map[cat]:
        return sub, {"exclusive": None, "options": []}, info

    node = visa_map[cat][sub]
    visa_label = _safe_str(node.get("visa_label", sub)) or sub
    opt = node.get("options", {"exclusive": None, "options": []})
    ex_label = opt.get("exclusive")  # nom du groupe exclusif, si présent
    checks   = opt.get("options", []) or []

    # Applique pré-sélection éventuelle
    preselected = preselected or {}
    sel_ex = preselected.get("exclusive", None)

    # Rend l'UI
    st.markdown(f"**Visa :** {visa_label}")

    # Groupe exclusif (radio) si défini
    if ex_label:
        choices = [ex_label]  # le libellé encadrant
        # Si on veut des sous-choix spécifiques on peut lister ici, mais dans le modèle actuel
        # l'intitulé «(XXX)» indique un groupe, donc on affiche un switch binaire.
        # On le traduit en un toggle Oui/Non sous forme de checkbox :
        use_ex = st.checkbox(ex_label, value=bool(sel_ex), key=f"{keyprefix}_ex")
        sel_ex = ex_label if use_ex else None
    else:
        use_ex = None

    selected_options = []
    if checks:
        st.caption("Options disponibles :")
        for i, opti in enumerate(checks):
            title = _safe_str(opti.get("title", ""))
            keyk  = _safe_str(opti.get("key", f"opt{i}"))
            default_checked = False
            if isinstance(preselected.get("options"), list):
                # si dicts complets
                for d in preselected["options"]:
                    if _safe_str(d.get("key")) == keyk and bool(d.get("checked", False)):
                        default_checked = True
                        break
            elif isinstance(preselected.get("options"), dict):
                # ancien format dict key->bool
                default_checked = bool(preselected["options"].get(keyk, False))
            # checkbox
            chk = st.checkbox(title, value=default_checked, key=f"{keyprefix}_chk_{keyk}")
            selected_options.append({"title": title, "key": keyk, "checked": bool(chk)})

    opts_dict = {"exclusive": sel_ex, "options": selected_options}

    # Construit un libellé final utile (ex: "B-1 COS", "F-1 EOS", etc.)
    suffixes = []
    if sel_ex:
        suffixes.append(sel_ex)
    suffixes.extend([d["title"] for d in selected_options if d.get("checked")])
    if suffixes:
        visa_label = f"{visa_label} " + " ".join(suffixes)

    return visa_label, opts_dict, info




# ================================
# 🛂 Visa Manager — PARTIE 3/4
# Dashboard, Analyses, Escrow, Gestion
# ================================

# ----------------
# 📊 DASHBOARD
# ----------------
with tabs[0]:
    st.header("📊 Dashboard général")

    if df_all.empty:
        st.info("Aucun client chargé pour le moment.")
    else:
        nb_clients = len(df_all)
        total_hono = _safe_num_series(df_all, HONO).sum()
        total_autre = _safe_num_series(df_all, AUTRE).sum()
        total_total = _safe_num_series(df_all, TOTAL).sum()
        total_paye = _safe_num_series(df_all, "Payé").sum()
        total_reste = _safe_num_series(df_all, "Reste").sum()

        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        kpi1.metric("Clients", nb_clients)
        kpi2.metric("Honoraires", _fmt_money(total_hono))
        kpi3.metric("Autres frais", _fmt_money(total_autre))
        kpi4.metric("Payé", _fmt_money(total_paye))
        kpi5.metric("Reste", _fmt_money(total_reste))

        # pourcentage par catégorie
        if "Categorie" in df_all.columns:
            st.subheader("Répartition par catégorie")
            cat_counts = df_all["Categorie"].value_counts(dropna=True)
            cat_percent = (cat_counts / cat_counts.sum()) * 100
            df_cat = pd.DataFrame({"Catégorie": cat_counts.index, "Nombre": cat_counts.values, "%": cat_percent.values})
            st.dataframe(df_cat, use_container_width=True)

        # par sous-catégorie
        if "Sous-categorie" in df_all.columns:
            st.subheader("Répartition par sous-catégorie")
            sub_counts = df_all["Sous-categorie"].value_counts(dropna=True)
            sub_percent = (sub_counts / sub_counts.sum()) * 100
            df_sub = pd.DataFrame({"Sous-catégorie": sub_counts.index, "Nombre": sub_counts.values, "%": sub_percent.values})
            st.dataframe(df_sub, use_container_width=True)

# ----------------
# 📈 ANALYSES
# ----------------
with tabs[1]:
    st.header("📈 Analyses")

    if df_all.empty:
        st.info("Charge un fichier Clients pour voir les analyses.")
    else:
        years = sorted(df_all["_Année_"].dropna().unique().tolist())
        months = [f"{i:02d}" for i in range(1, 13)]
        col1, col2 = st.columns(2)
        fy = col1.multiselect("Année", years, default=years)
        fm = col2.multiselect("Mois", months, default=[])

        df_filt = df_all.copy()
        if fy:
            df_filt = df_filt[df_filt["_Année_"].isin(fy)]
        if fm:
            df_filt = df_filt[df_filt["Mois"].astype(str).isin(fm)]

        st.markdown("### Totaux par période sélectionnée")
        total_hono = _safe_num_series(df_filt, HONO).sum()
        total_autre = _safe_num_series(df_filt, AUTRE).sum()
        total_total = _safe_num_series(df_filt, TOTAL).sum()
        total_paye = _safe_num_series(df_filt, "Payé").sum()
        total_reste = _safe_num_series(df_filt, "Reste").sum()

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Honoraires", _fmt_money(total_hono))
        c2.metric("Autres frais", _fmt_money(total_autre))
        c3.metric("Total", _fmt_money(total_total))
        c4.metric("Payé", _fmt_money(total_paye))
        c5.metric("Reste", _fmt_money(total_reste))

        # graphique comparatif par année ou mois
        st.markdown("### Évolution des montants (Honoraires + Autres frais)")
        df_grp = df_filt.groupby(["_Année_", "Mois"], dropna=True)[[HONO, AUTRE, TOTAL]].sum().reset_index()
        if not df_grp.empty:
            import plotly.express as px
            df_grp["Période"] = df_grp["_Année_"].astype(str) + "-" + df_grp["Mois"].astype(str)
            fig = px.bar(df_grp, x="Période", y="TOTAL", title="Montant total par période")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.caption("Aucune donnée pour cette période.")

# ----------------
# 🏦 ESCROW
# ----------------
with tabs[2]:
    st.header("🏦 Escrow")

    if df_all.empty:
        st.info("Charge un fichier Clients pour voir les totaux.")
    else:
        total_hono = _safe_num_series(df_all, HONO).sum()
        total_paye = _safe_num_series(df_all, "Payé").sum()
        total_reste = _safe_num_series(df_all, "Reste").sum()

        k1, k2, k3 = st.columns(3)
        k1.metric("Total honoraires", _fmt_money(total_hono))
        k2.metric("Total payé", _fmt_money(total_paye))
        k3.metric("Reste dû", _fmt_money(total_reste))

        st.markdown("### Détails Escrow")
        df_display = df_all[[DOSSIER_COL, "Nom", HONO, AUTRE, "Payé", "Reste"]].copy()
        st.dataframe(df_display, use_container_width=True)

# ----------------
# 🧾 GESTION
# ----------------
with tabs[4]:
    st.header("🧾 Gestion des données")
    st.markdown("Télécharger ou sauvegarder les fichiers consolidés.")

    if not df_all.empty:
        if st.button("💾 Télécharger Clients (xlsx)", key=skey("dl", "clients")):
            bytes_out = write_clients_to_bytes(df_all)
            st.download_button(
                label="⬇️ Télécharger Clients.xlsx",
                data=bytes_out,
                file_name="Clients_modifié.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if not df_visa_raw.empty:
            if st.button("💾 Télécharger Clients + Visa (xlsx)", key=skey("dl", "both")):
                bytes_out = write_two_sheets_to_bytes(df_all, df_visa_raw)
                st.download_button(
                    label="⬇️ Télécharger Fichier complet.xlsx",
                    data=bytes_out,
                    file_name="VisaManager_Complet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info("Charge les fichiers avant d'exporter.")




# ================================
# 🛂 Visa Manager — PARTIE 4/4
# Compte client (fiche + paiements + statuts + visa) & Visa (aperçu)
# ================================

# --------- utilitaire paiements ---------
def _parse_payments_cell(val) -> List[Dict[str, Any]]:
    """Retourne une liste de paiements :
       [{"date":"YYYY-MM-DD","amount":float,"method":"...","note":"..."}]
    """
    if isinstance(val, list):
        return val
    s = _safe_str(val)
    if not s:
        return []
    try:
        data = json.loads(s)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []

def _payments_to_str(lst: List[Dict[str, Any]]) -> str:
    try:
        return json.dumps(lst, ensure_ascii=False)
    except Exception:
        return "[]"

# ================================
# 👤 COMPTE CLIENT — fiche & paiements
# ================================
with tabs[3]:
    st.header("👤 Compte client")

    df_live = _read_clients()  # état courant
    if df_live.empty:
        st.info("Aucun client chargé.")
    else:
        # Sélection client (par Nom ou ID)
        csel1, csel2 = st.columns([2,2])
        names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
        ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
        sel_name = csel1.selectbox("Nom", [""]+names, index=0, key=skey("acct","name"))
        sel_id   = csel2.selectbox("ID_Client", [""]+ids, index=0, key=skey("acct","id"))

        mask = None
        if sel_id:
            mask = (df_live["ID_Client"].astype(str) == sel_id)
        elif sel_name:
            mask = (df_live["Nom"].astype(str) == sel_name)

        if mask is None or not mask.any():
            st.stop()

        idx = df_live[mask].index[0]
        row = df_live.loc[idx].copy()

        # ---------- Affichage synthèse ----------
        st.subheader("Résumé du dossier")
        r1, r2, r3, r4, r5, r6 = st.columns(6)
        r1.metric("Dossier N", _safe_str(row.get(DOSSIER_COL,"")))
        r2.metric("Honoraires", _fmt_money(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0]))
        r3.metric("Autres frais", _fmt_money(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0]))
        r4.metric("Total", _fmt_money(_safe_num_series(pd.DataFrame([row]), TOTAL).iloc[0]))
        r5.metric("Payé", _fmt_money(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0]))
        r6.metric("Reste", _fmt_money(_safe_num_series(pd.DataFrame([row]), "Reste").iloc[0]))

        st.markdown("---")

        # ---------- Édition fiche ----------
        st.subheader("✏️ Éditer la fiche")
        d1, d2, d3 = st.columns([2,1,1])

        nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=skey("acct","nom"))
        dval = _date_for_widget(row.get("Date")) or date.today()
        dt   = d2.date_input("Date de création", value=dval, key=skey("acct","date"))
        mois_def = _safe_str(row.get("Mois",""))
        try:
            mois_index = (int(mois_def) - 1) if mois_def.isdigit() else (dt.month-1)
        except Exception:
            mois_index = dt.month - 1
        mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=max(0, min(11, mois_index)), key=skey("acct","mois"))

        # Visa cascade (catégorie → sous-catégorie → options)
        st.markdown("#### 🎯 Visa")
        cats = sorted(list(visa_map.keys()))
        preset_cat = _safe_str(row.get("Categorie",""))
        sel_cat = st.selectbox("Catégorie", [""]+cats, index=(cats.index(preset_cat)+1 if preset_cat in cats else 0), key=skey("acct","cat"))

        subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
        preset_sub = _safe_str(row.get("Sous-categorie",""))
        sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=(subs.index(preset_sub)+1 if preset_sub in subs else 0), key=skey("acct","sub"))

        # options pré-sélectionnées
        preset_opts = row.get("Options", {})
        if not isinstance(preset_opts, dict):
            try:
                preset_opts = json.loads(_safe_str(preset_opts) or "{}")
                if not isinstance(preset_opts, dict):
                    preset_opts = {}
            except Exception:
                preset_opts = {}

        visa_final, opts_dict, _ = ("", {"exclusive": None, "options": []}, "")
        if sel_cat and sel_sub:
            visa_final, opts_dict, _ = build_visa_option_selector(
                visa_map, sel_cat, sel_sub, keyprefix=skey("acct","opts"), preselected=preset_opts
            )

        # Montants + commentaire
        st.markdown("#### 💵 Montants")
        f1, f2, f3 = st.columns([1,1,2])
        honor = f1.number_input(HONO, min_value=0.0,
                                value=float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0]),
                                step=50.0, format="%.2f", key=skey("acct","hono"))
        other = f2.number_input(AUTRE, min_value=0.0,
                                value=float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0]),
                                step=20.0, format="%.2f", key=skey("acct","autre"))
        comm  = f3.text_area("Commentaire (Autres frais / notes)", _safe_str(row.get("Commentaire","")), key=skey("acct","comm"))

        # Statuts + dates (RFE seulement si un autre statut coché)
        st.markdown("#### 📌 Statuts du dossier")
        s1, s2, s3, s4, s5 = st.columns(5)
        envoye = s1.checkbox("Dossier envoyé", value=bool(int(row.get("Dossier envoyé",0) or 0)), key=skey("acct","sent"))
        sent_d_val = _date_for_widget(row.get("Date d'envoi")) or (dt if envoye else date.today())
        sent_d = s1.date_input("Date d'envoi", value=sent_d_val, key=skey("acct","sentd"))

        accepte = s2.checkbox("Dossier accepté", value=bool(int(row.get("Dossier accepté",0) or 0)), key=skey("acct","acc"))
        acc_d_val = _date_for_widget(row.get("Date d'acceptation")) or date.today()
        acc_d = s2.date_input("Date d'acceptation", value=acc_d_val, key=skey("acct","accd"))

        refuse = s3.checkbox("Dossier refusé", value=bool(int(row.get("Dossier refusé",0) or 0)), key=skey("acct","ref"))
        ref_d_val = _date_for_widget(row.get("Date de refus")) or date.today()
        ref_d = s3.date_input("Date de refus", value=ref_d_val, key=skey("acct","refd"))

        annule = s4.checkbox("Dossier annulé", value=bool(int(row.get("Dossier annulé",0) or 0)), key=skey("acct","ann"))
        ann_d_val = _date_for_widget(row.get("Date d'annulation")) or date.today()
        ann_d = s4.date_input("Date d'annulation", value=ann_d_val, key=skey("acct","annd"))

        rfe = s5.checkbox("RFE", value=bool(int(row.get("RFE",0) or 0)), key=skey("acct","rfe"))
        if rfe and not any([envoye, accepte, refuse, annule]):
            st.warning("⚠️ RFE ne peut être coché qu’avec Envoyé / Accepté / Refusé / Annulé.")

        # ---------- Paiements ----------
        st.markdown("#### 💳 Paiements")
        pay_list = _parse_payments_cell(row.get("Paiements"))

        if len(pay_list):
            df_pay = pd.DataFrame(pay_list)
            # normalisation affichage
            if "date" in df_pay.columns:
                try:
                    df_pay["date"] = pd.to_datetime(df_pay["date"], errors="coerce").dt.date.astype(str)
                except Exception:
                    df_pay["date"] = df_pay["date"].astype(str)
            if "amount" in df_pay.columns:
                df_pay["amount"] = pd.to_numeric(df_pay["amount"], errors="coerce").fillna(0.0)
            st.dataframe(
                df_pay.rename(columns={"date":"Date","amount":"Montant","method":"Mode","note":"Note"}),
                use_container_width=True,
                height=220,
                key=skey("acct","paytbl")
            )
        else:
            st.caption("Aucun règlement enregistré.")

        reste = float(_safe_num_series(pd.DataFrame([row]), "Reste").iloc[0])
        can_add = (reste > 0.0)

        addc1, addc2, addc3, addc4 = st.columns([1,1,1,2])
        pay_amt = addc1.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", value=0.0, key=skey("acct","payamt"))
        pay_date = addc2.date_input("Date règlement", value=date.today(), key=skey("acct","paydate"))
        pay_method = addc3.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], index=0, key=skey("acct","paymethod"))
        pay_note = addc4.text_input("Note", "", key=skey("acct","paynote"))

        add_btn = st.button("➕ Ajouter règlement", disabled=not can_add, key=skey("acct","addpaybtn"))
        if add_btn:
            if pay_amt <= 0.0:
                st.warning("Le montant doit être > 0.")
            else:
                # recalcul
                total = float(honor) + float(other)
                already = float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0])
                new_paid = already + float(pay_amt)
                if new_paid > total:
                    st.warning("Le paiement dépasse le total. Ajustez le montant.")
                else:
                    pay_list.append({
                        "date": _to_iso_date(pay_date),
                        "amount": float(pay_amt),
                        "method": _safe_str(pay_method),
                        "note": _safe_str(pay_note)
                    })
                    df_live.at[idx, "Paiements"] = _payments_to_str(pay_list)
                    df_live.at[idx, "Payé"] = new_paid
                    df_live.at[idx, "Reste"] = max(0.0, total - new_paid)
                    _write_clients(df_live)
                    st.success("Règlement ajouté.")
                    st.cache_data.clear()
                    st.rerun()

        # ---------- Enregistrement fiche ----------
        st.markdown("---")
        if st.button("💾 Enregistrer la fiche", key=skey("acct","save")):
            if not nom:
                st.warning("Le nom est requis.")
                st.stop()

            total = float(honor) + float(other)
            paye  = float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0])
            reste = max(0.0, total - paye)

            # champs de base
            df_live.at[idx, "Nom"]   = nom
            df_live.at[idx, "Date"]  = dt
            df_live.at[idx, "Mois"]  = f"{int(mois):02d}" if _safe_str(mois).isdigit() else _safe_str(mois)
            df_live.at[idx, "Categorie"] = sel_cat
            df_live.at[idx, "Sous-categorie"] = sel_sub
            df_live.at[idx, "Visa"] = (visa_final if (sel_cat and sel_sub) else _safe_str(row.get("Visa","")))
            df_live.at[idx, HONO] = float(honor)
            df_live.at[idx, AUTRE] = float(other)
            df_live.at[idx, TOTAL] = total
            df_live.at[idx, "Reste"] = reste
            df_live.at[idx, "Options"] = opts_dict
            df_live.at[idx, "Commentaire"] = comm

            # statuts
            df_live.at[idx, "Dossier envoyé"]  = 1 if envoye else 0
            df_live.at[idx, "Date d'envoi"]    = sent_d
            df_live.at[idx, "Dossier accepté"] = 1 if accepte else 0
            df_live.at[idx, "Date d'acceptation"] = acc_d
            df_live.at[idx, "Dossier refusé"]  = 1 if refuse else 0
            df_live.at[idx, "Date de refus"]   = ref_d
            df_live.at[idx, "Dossier annulé"]  = 1 if annule else 0
            df_live.at[idx, "Date d'annulation"] = ann_d
            df_live.at[idx, "RFE"] = 1 if (rfe and any([envoye, accepte, refuse, annule])) else 0

            _write_clients(df_live)
            st.success("Fiche enregistrée.")
            st.cache_data.clear()
            st.rerun()

# ================================
# 📄 VISA (aperçu)
# ================================
with tabs[5]:
    st.header("📄 Visa — aperçu")
    if df_visa_raw.empty:
        st.info("Aucune feuille Visa chargée.")
    else:
        colf1, colf2 = st.columns([1,2])
        # filtres simples
        cats = sorted(df_visa_raw["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_visa_raw.columns else []
        catF = colf1.selectbox("Catégorie", [""]+cats, index=0, key=skey("viz","cat"))
        if catF and "Sous-categorie" in df_visa_raw.columns:
            subs = sorted(df_visa_raw.loc[df_visa_raw["Categorie"].astype(str)==catF, "Sous-categorie"].dropna().astype(str).unique().tolist())
        else:
            subs = sorted(df_visa_raw.get("Sous-categorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
        subF = colf1.selectbox("Sous-catégorie", [""]+subs, index=0, key=skey("viz","sub"))

        view = df_visa_raw.copy()
        if catF:
            view = view[view["Categorie"].astype(str)==catF]
        if subF:
            view = view[view["Sous-categorie"].astype(str)==subF]

        st.dataframe(view.reset_index(drop=True), use_container_width=True, height=350, key=skey("viz","tbl"))
