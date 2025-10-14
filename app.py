# ================================
# 🛂 Visa Manager — PARTIE 1/4
# Setup, Helpers, Persistance, Chargement fichiers, Visa map
# ================================

from __future__ import annotations

import json
import re
import zipfile
from io import BytesIO
from pathlib import Path
from datetime import date, datetime
from typing import Dict, List, Tuple, Any

import pandas as pd
import streamlit as st

# -------------------------------------------------------
# Configuration de la page
# -------------------------------------------------------
st.set_page_config(
    page_title="Visa Manager",
    layout="wide",
)

# -------------------------------------------------------
# Constantes colonnes (noms attendus dans l’onglet Clients)
# -------------------------------------------------------
DOSSIER_COL = "Dossier N"
HONO = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"

SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"

# -------------------------------------------------------
# Helpers génériques
# -------------------------------------------------------
def _safe_str(x: Any) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _to_int_bool(v: Any) -> bool:
    try:
        return int(v or 0) == 1
    except Exception:
        return False

def _safe_num_series(df: pd.DataFrame | pd.Series, col: str) -> pd.Series:
    """Retourne une Series numérique sûre (remplace NaN par 0.0)."""
    if isinstance(df, pd.Series):
        s = df
    else:
        s = df.get(col, pd.Series([], dtype=float))
    try:
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    except Exception:
        # Si conversion impossible
        try:
            s2 = s.astype(str).str.replace(r"[^\d\-,.]", "", regex=True)
            return pd.to_numeric(s2, errors="coerce").fillna(0.0)
        except Exception:
            return pd.Series([0.0] * len(s), index=s.index)

def _fmt_money(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(val: Any, default_to_today: bool = True) -> date | None:
    """
    Retourne un objet date (ou None) acceptable par st.date_input.
    Gère date/datetime/str/NaT/None.
    """
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        ts = pd.to_datetime(val, errors="coerce")
        if pd.isna(ts):
            return date.today() if default_to_today else None
        return ts.date()
    except Exception:
        return date.today() if default_to_today else None

# Générateur de clés uniques (pour éviter les collisions streamlit)
SID = "vm"  # suffixe stable pour cette session
def skey(*parts: str) -> str:
    return "k_" + "_".join([SID] + [p.replace(" ", "_") for p in parts])

# -------------------------------------------------------
# Persistance des derniers chemins utilisés
# -------------------------------------------------------
APP_DIR = Path(".")
STATE_FILE = APP_DIR / ".visa_manager_state.json"

def _save_last_paths(clients_path: str | None, visa_path: str | None, mode_dual: bool):
    try:
        data = {
            "mode_dual": bool(mode_dual),
            "clients_path": str(clients_path) if clients_path else "",
            "visa_path": str(visa_path) if visa_path else "",
        }
        STATE_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass

def _load_last_paths() -> tuple[bool, str | None, str | None]:
    try:
        if not STATE_FILE.exists():
            return True, None, None  # par défaut “Deux fichiers”
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        mode_dual = bool(data.get("mode_dual", True))
        clients_path = data.get("clients_path") or None
        visa_path = data.get("visa_path") or None
        return mode_dual, clients_path, visa_path
    except Exception:
        return True, None, None

# -------------------------------------------------------
# Lecture Excel (UploadedFile ou Chemin)
# -------------------------------------------------------
def _read_excel_any(source: Any, sheet_name: str | None = None) -> pd.DataFrame:
    """
    Accepte :
      - UploadedFile streamlit (avec .read disponible)
      - chemin str/Path vers un fichier xlsx
    """
    if source is None:
        return pd.DataFrame()

    # UploadedFile -> BytesIO
    if hasattr(source, "read"):
        data = source.read()
        bio = BytesIO(data)
        return pd.read_excel(bio, sheet_name=sheet_name)

    # Chemin
    p = Path(str(source))
    if not p.exists():
        return pd.DataFrame()
    return pd.read_excel(p, sheet_name=sheet_name)

def _write_clients(df: pd.DataFrame, source: Any):
    """
    Écrit l’onglet Clients :
      - si source est UploadedFile initial, on demande un chemin de sauvegarde
      - si source est un chemin .xlsx, on réécrit dedans (onglet Clients uniquement)
    """
    # Cas “fichier unique 2 onglets” : on sauvegardera dans le même fichier (géré plus loin)
    # Ici on gère seulement le cas “deux fichiers” Clients séparé.
    try:
        # si source est un chemin
        if isinstance(source, (str, Path)):
            p = Path(str(source))
            # Si le fichier existe déjà, on le met à jour juste pour l'onglet Clients
            if p.exists():
                with pd.ExcelWriter(p, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
                    df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
            else:
                with pd.ExcelWriter(p, engine="openpyxl") as wr:
                    df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        else:
            # UploadedFile : demander un chemin de sauvegarde en session
            st.warning("Source Clients chargée via upload. Utilisez l’export global en bas de page pour récupérer le fichier.")
    except Exception as e:
        st.error("Erreur d’écriture Clients : " + _safe_str(e))

def _write_two_tabs(df_clients: pd.DataFrame, df_visa: pd.DataFrame, path: str | Path):
    """Écrit un seul fichier avec les deux onglets 'Clients' et 'Visa'."""
    try:
        p = Path(str(path))
        with pd.ExcelWriter(p, engine="openpyxl") as wr:
            df_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
            df_visa.to_excel(wr, sheet_name=SHEET_VISA, index=False)
    except Exception as e:
        st.error("Erreur d’écriture fichier 2 onglets : " + _safe_str(e))

# -------------------------------------------------------
# Construction de la table “Clients” normalisée minimale
# -------------------------------------------------------
def _ensure_clients_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Crée les colonnes manquantes avec des valeurs par défaut."""
    need_cols = [
        DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
        "Categorie", "Sous-categorie", "Visa",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Paiements", "Options",
        "Dossier envoyé", "Date d'envoi",
        "Dossier accepté", "Date d'acceptation",
        "Dossier refusé", "Date de refus",
        "Dossier annulé", "Date d'annulation",
        "RFE",
        "Commentaires",
    ]
    for c in need_cols:
        if c not in df.columns:
            if c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                df[c] = 0.0
            elif c in ["Paiements", "Options"]:
                df[c] = ""
            elif c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
                df[c] = 0
            else:
                df[c] = ""
    # Normalisation Date/Mois
    if "Date" in df.columns:
        try:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        except Exception:
            pass
    if "Mois" in df.columns:
        df["Mois"] = df["Mois"].astype(str).str.zfill(2)
        # colonnes techniques pour tri/analyses
        try:
            years = pd.to_datetime(df["Date"], errors="coerce").dt.year
            months = pd.to_datetime(df["Date"], errors="coerce").dt.month
            df["_Année_"] = years.fillna(0).astype(int)
            df["_MoisNum_"] = months.fillna(0).astype(int)
        except Exception:
            df["_Année_"] = 0
            df["_MoisNum_"] = 0
    return df

# -------------------------------------------------------
# Construction du visa_map à partir de l’onglet Visa
# - Colonnes attendues : "Categorie", "Sous-categorie", puis colonnes d’options (ex: COS, EOS, …) contenant 1
# - Si “Visa” existe, on l’utilise comme libellé final. Sinon : “Sous-categorie + option cochée”.
# -------------------------------------------------------
def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    if df_visa.empty:
        return {}

    # uniformiser noms de colonnes (sans accents dans les valeurs, mais colonnes en FR OK)
    cols = {c: c.strip() for c in df_visa.columns}
    dfv = df_visa.rename(columns=cols).copy()

    base_cols = {"Categorie", "Sous-categorie", "Visa"}
    all_cols = list(dfv.columns)
    option_cols = [c for c in all_cols if c not in base_cols]

    # Remplacer NaN par vide/0
    for c in ["Categorie","Sous-categorie","Visa"]:
        if c in dfv.columns:
            dfv[c] = dfv[c].fillna("").astype(str)

    # Pour les colonnes d’options, considérer “1” comme coché
    for oc in option_cols:
        dfv[oc] = pd.to_numeric(dfv[oc], errors="coerce").fillna(0).astype(int)

    vm: Dict[str, Dict[str, Dict[str, Any]]] = {}
    for _, row in dfv.iterrows():
        cat = _safe_str(row.get("Categorie","")).strip()
        sub = _safe_str(row.get("Sous-categorie","")).strip()
        visa_label = _safe_str(row.get("Visa","")).strip()

        if not cat or not sub:
            continue

        checked = []
        for oc in option_cols:
            try:
                if int(row.get(oc, 0)) == 1:
                    checked.append(oc)
            except Exception:
                pass

        # Si pas de “Visa” explicite : si options cochées, on produit (“sub + option”)
        # Sinon, on garde sub comme Visa par défaut.
        options_def = {
            "exclusive": None,   # None = pas d’exclusivité
            "options": checked   # liste d’options dispos pour l’affichage (cases à cocher)
        }

        # On ne “fige” pas le Visa ici ; le libellé final se construira côté UI
        if cat not in vm:
            vm[cat] = {}
        vm[cat][sub] = {
            "options": options_def,
            "has_label": bool(visa_label),
            "label": visa_label,  # peut être vide
        }

    return vm

# -------------------------------------------------------
# UI — Chargement des fichiers (Upload ou Chemin)
# -------------------------------------------------------
st.title("🛂 Visa Manager")

st.markdown("## 📂 Fichiers")

last_mode_dual, last_clients_path, last_visa_path = _load_last_paths()

mode = st.radio(
    "Mode de chargement",
    ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"],
    index=(0 if last_mode_dual else 1),
    horizontal=True,
    key=skey("load","mode"),
)

clients_source = None
visa_source = None
two_tabs_path_text = None

if mode == "Deux fichiers (Clients & Visa)":
    c1, c2 = st.columns(2)
    with c1:
        st.caption("**Clients (xlsx)** — *Upload ou chemin local*")
        up_c = st.file_uploader("Upload Clients", type=["xlsx"], key=skey("up","clients"))
        path_c = st.text_input(
            "Chemin Clients (facultatif si upload)",
            value=(last_clients_path or ""),
            placeholder="ex: /home/user/clients.xlsx",
            key=skey("path","clients"),
        )
        if up_c is not None:
            clients_source = up_c
        elif path_c.strip():
            clients_source = path_c.strip()
        elif last_clients_path and Path(last_clients_path).exists():
            clients_source = last_clients_path

    with c2:
        st.caption("**Visa (xlsx)** — *Upload ou chemin local*")
        up_v = st.file_uploader("Upload Visa", type=["xlsx"], key=skey("up","visa"))
        path_v = st.text_input(
            "Chemin Visa (facultatif si upload)",
            value=(last_visa_path or ""),
            placeholder="ex: /home/user/visa.xlsx",
            key=skey("path","visa"),
        )
        if up_v is not None:
            visa_source = up_v
        elif path_v.strip():
            visa_source = path_v.strip()
        elif last_visa_path and Path(last_visa_path).exists():
            visa_source = last_visa_path

    # Sauver l'état
    _save_last_paths(
        clients_path=(getattr(clients_source, "name", str(clients_source)) if clients_source else None),
        visa_path=(getattr(visa_source, "name", str(visa_source)) if visa_source else None),
        mode_dual=True,
    )

else:
    st.caption("**Fichier unique (2 onglets : Clients & Visa)** — *Upload ou chemin local*")
    up_one = st.file_uploader("Upload fichier unique", type=["xlsx"], key=skey("up","one"))
    path_one = st.text_input(
        "Chemin (facultatif si upload)",
        value=(last_clients_path or ""),
        placeholder="ex: /home/user/donnees.xlsx",
        key=skey("path","one"),
    )
    if up_one is not None:
        clients_source = up_one
        visa_source = up_one
        two_tabs_path_text = None
    elif path_one.strip():
        clients_source = path_one.strip()
        visa_source = path_one.strip()
        two_tabs_path_text = path_one.strip()
    elif last_clients_path and Path(last_clients_path).exists():
        clients_source = last_clients_path
        visa_source = last_clients_path
        two_tabs_path_text = last_clients_path

    _save_last_paths(
        clients_path=(getattr(clients_source, "name", str(clients_source)) if clients_source else None),
        visa_path=(getattr(visa_source, "name", str(visa_source)) if visa_source else None),
        mode_dual=False,
    )

# -------------------------------------------------------
# Lecture des DataFrames
# -------------------------------------------------------
if mode == "Deux fichiers (Clients & Visa)":
    df_clients = _read_excel_any(clients_source, sheet_name=None)
    if isinstance(df_clients, dict):
        # si l’utilisateur a donné un fichier avec plusieurs onglets, on cherche "Clients"
        df_clients = df_clients.get(SHEET_CLIENTS, pd.DataFrame())
    df_visa_raw = _read_excel_any(visa_source, sheet_name=None)
    if isinstance(df_visa_raw, dict):
        df_visa_raw = df_visa_raw.get(SHEET_VISA, pd.DataFrame())
else:
    # Fichier unique
    df_all_sheets = _read_excel_any(clients_source, sheet_name=None)
    if isinstance(df_all_sheets, dict):
        df_clients = df_all_sheets.get(SHEET_CLIENTS, pd.DataFrame())
        df_visa_raw = df_all_sheets.get(SHEET_VISA, pd.DataFrame())
    else:
        # fichier sans onglets attendus
        df_clients = pd.DataFrame()
        df_visa_raw = pd.DataFrame()

# Normaliser Clients
df_clients = df_clients.copy()
if not df_clients.empty:
    # S'assurer que toutes les colonnes attendues sont là
    df_clients = _ensure_clients_columns(df_clients)
else:
    # squelette vide prêt à l’emploi
    df_clients = _ensure_clients_columns(pd.DataFrame())

# Construire le visa_map
df_visa_raw = df_visa_raw.copy()
visa_map: Dict[str, Dict[str, Dict[str, Any]]] = {}
if not df_visa_raw.empty:
    # On s'assure que les colonnes de base existent (au minimum cat/sub)
    # Si ce n’est pas le cas, on tente des variantes de nommage.
    rename_candidates = {}
    for c in df_visa_raw.columns:
        cn = _safe_str(c).strip()
        if cn.lower() in {"categorie","catégorie"}:
            rename_candidates[c] = "Categorie"
        elif cn.lower() in {"sous-categorie","sous-catégorie","sous categorie","sous catégorie"}:
            rename_candidates[c] = "Sous-categorie"
        elif cn.lower() == "visa":
            rename_candidates[c] = "Visa"
    if rename_candidates:
        df_visa_raw = df_visa_raw.rename(columns=rename_candidates)

    if "Categorie" in df_visa_raw.columns and "Sous-categorie" in df_visa_raw.columns:
        visa_map = build_visa_map(df_visa_raw)
    else:
        st.warning("L’onglet Visa doit contenir au moins les colonnes ‘Categorie’ et ‘Sous-categorie’.")

# -------------------------------------------------------
# Barres d’onglets de l’app (Dashboard, Analyses, Escrow, Clients, Gestion, Visa)
# (Le contenu des onglets est fourni dans les parties suivantes)
# -------------------------------------------------------
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "🧾 Gestion", "📄 Visa (aperçu)"])

# Les parties 2/4, 3/4 et 4/4 viendront compléter l’UI à partir de `tabs[...]`.
# Ne rien mettre ici pour éviter les références avant définition.




# ================================
# 🛂 Visa Manager — PARTIE 2/4
# Onglet 📊 Dashboard : filtres, KPI, graphiques, tableau
# ================================

with tabs[0]:
    st.subheader("📊 Dashboard")

    df_all = df_clients.copy()

    # ---------- Jeu de colonnes sûr ----------
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c not in df_all.columns:
            df_all[c] = 0.0

    if "Date" not in df_all.columns:
        df_all["Date"] = pd.NaT
    if "Mois" not in df_all.columns:
        df_all["Mois"] = ""
    if "_Année_" not in df_all.columns or "_MoisNum_" not in df_all.columns:
        try:
            years = pd.to_datetime(df_all["Date"], errors="coerce").dt.year
            months = pd.to_datetime(df_all["Date"], errors="coerce").dt.month
            df_all["_Année_"] = years.fillna(0).astype(int)
            df_all["_MoisNum_"] = months.fillna(0).astype(int)
        except Exception:
            df_all["_Année_"] = 0
            df_all["_MoisNum_"] = 0

    # ---------- Listes de filtres ----------
    years = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist() if int(y) > 0])
    monthsA = [f"{m:02d}" for m in range(1, 13)]
    cats  = sorted(df_all.get("Categorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
    subs  = sorted(df_all.get("Sous-categorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
    visas = sorted(df_all.get("Visa", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())

    st.markdown("#### 🔎 Filtres")
    f1, f2, f3, f4, f5, f6 = st.columns([1.1, 1.1, 1.3, 1.3, 1.4, 1.2])

    fy = f1.multiselect("Année", years, default=[], key=skey("dash", "years"))
    fm = f2.multiselect("Mois (MM)", monthsA, default=[], key=skey("dash", "months"))
    fc = f3.multiselect("Catégorie", cats, default=[], key=skey("dash", "cats"))
    fs = f4.multiselect("Sous-catégorie", subs, default=[], key=skey("dash", "subs"))
    fv = f5.multiselect("Visa", visas, default=[], key=skey("dash", "visas"))

    txt = f6.text_input("🔍 Nom contient", value="", key=skey("dash", "search"))

    # ---------- Application des filtres ----------
    ff = df_all.copy()
    if fy: ff = ff[ff["_Année_"].isin(fy)]
    if fm: ff = ff[ff["Mois"].astype(str).isin(fm)]
    if fc: ff = ff[ff["Categorie"].astype(str).isin(fc)]
    if fs: ff = ff[ff["Sous-categorie"].astype(str).isin(fs)]
    if fv: ff = ff[ff["Visa"].astype(str).isin(fv)]
    if txt:
        pat = str(txt).strip().lower()
        ff = ff[ff["Nom"].astype(str).str.lower().str.contains(pat, na=False)]

    # ---------- KPI compacts ----------
    st.markdown("#### ✅ Indicateurs")
    k1, k2, k3, k4, k5, k6 = st.columns([0.9, 1.0, 1.0, 1.0, 0.9, 0.9])

    total_dossiers = int(len(ff))
    sum_hono  = float(_safe_num_series(ff, HONO).sum())
    sum_other = float(_safe_num_series(ff, AUTRE).sum())
    sum_total = float(_safe_num_series(ff, TOTAL).sum())
    sum_paye  = float(_safe_num_series(ff, "Payé").sum())
    sum_reste = float(_safe_num_series(ff, "Reste").sum())

    k1.metric("Dossiers", f"{total_dossiers}")
    k2.metric("Honoraires", _fmt_money(sum_hono))
    k3.metric("Autres frais", _fmt_money(sum_other))
    k4.metric("Total", _fmt_money(sum_total))
    k5.metric("Payé", _fmt_money(sum_paye))
    k6.metric("Reste", _fmt_money(sum_reste))

    # ---------- % par Catégorie / Sous-catégorie ----------
    st.markdown("#### % Répartition")
    pc1, pc2 = st.columns(2)
    if not ff.empty:
        # % par Catégorie
        if "Categorie" in ff.columns:
            dist_c = ff["Categorie"].value_counts(dropna=True)
            pct_c = (dist_c / max(1, dist_c.sum()) * 100.0).round(1)
            pc1.dataframe(
                pct_c.rename("Pourcentage (%)").to_frame(),
                use_container_width=True,
                height=220
            )
        # % par Sous-catégorie
        if "Sous-categorie" in ff.columns:
            dist_s = ff["Sous-categorie"].value_counts(dropna=True)
            pct_s = (dist_s / max(1, dist_s.sum()) * 100.0).round(1)
            pc2.dataframe(
                pct_s.rename("Pourcentage (%)").to_frame(),
                use_container_width=True,
                height=220
            )

    # ---------- Graphiques ----------
    st.markdown("#### 📈 Graphiques")

    g1, g2 = st.columns(2)

    # Dossiers par mois (barres)
    with g1:
        if not ff.empty:
            tmp = ff.copy()
            tmp["Mois"] = tmp["Mois"].astype(str).str.zfill(2)
            grp = tmp.groupby("Mois", as_index=False).size().rename(columns={"size": "Dossiers"})
            if not grp.empty:
                st.bar_chart(grp.set_index("Mois"), use_container_width=True)
            else:
                st.info("Aucune donnée pour le graphique Dossiers par mois.")
        else:
            st.info("Aucune donnée pour le graphique Dossiers par mois.")

    # Honoraires par mois (ligne)
    with g2:
        if not ff.empty:
            tmp = ff.copy()
            tmp["Mois"] = tmp["Mois"].astype(str).str.zfill(2)
            grp_h = tmp.groupby("Mois", as_index=False)[HONO].sum().rename(columns={HONO: "Honoraires"})
            if not grp_h.empty:
                st.line_chart(grp_h.set_index("Mois"), use_container_width=True)
            else:
                st.info("Aucune donnée pour le graphique Honoraires par mois.")
        else:
            st.info("Aucune donnée pour le graphique Honoraires par mois.")

    # ---------- Détails des dossiers filtrés ----------
    st.markdown("#### 📋 Détails des dossiers filtrés")

    view = ff.copy()
    # Formatage monétaire pour affichage
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c in view.columns:
            view[c] = _safe_num_series(ff, c).apply(_fmt_money)
    # Date en texte pour lisibilité
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

    # tri stable
    sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
    view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

    st.dataframe(
        view_sorted[show_cols].reset_index(drop=True),
        use_container_width=True,
        key=skey("dash", "table")
    )




# ================================
# 🛂 Visa Manager — PARTIE 3/4
# Onglet 📈 Analyses : filtres, KPI, graphiques, comparaisons A vs B
# ================================

with tabs[1]:
    st.subheader("📈 Analyses")

    df_all = df_clients.copy()

    # Colonnes sûres
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c not in df_all.columns:
            df_all[c] = 0.0
    if "Mois" not in df_all.columns:
        df_all["Mois"] = ""
    if "_Année_" not in df_all.columns or "_MoisNum_" not in df_all.columns:
        try:
            years = pd.to_datetime(df_all.get("Date"), errors="coerce").dt.year
            months = pd.to_datetime(df_all.get("Date"), errors="coerce").dt.month
            df_all["_Année_"] = years.fillna(0).astype(int)
            df_all["_MoisNum_"] = months.fillna(0).astype(int)
        except Exception:
            df_all["_Année_"] = 0
            df_all["_MoisNum_"] = 0

    # ---------- Filtres globaux ----------
    yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist() if int(y) > 0])
    monthsA = [f"{m:02d}" for m in range(1, 13)]
    catsA   = sorted(df_all.get("Categorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
    subsA   = sorted(df_all.get("Sous-categorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
    visasA  = sorted(df_all.get("Visa", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())

    st.markdown("### 🔎 Filtres (vue globale)")
    a1, a2, a3, a4, a5 = st.columns([1.1, 1.1, 1.3, 1.3, 1.2])

    fy = a1.multiselect("Année", yearsA, default=[], key=skey("ana","years"))
    fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=skey("ana","months"))
    fc = a3.multiselect("Catégorie", catsA, default=[], key=skey("ana","cats"))
    fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=skey("ana","subs"))
    fv = a5.multiselect("Visa", visasA, default=[], key=skey("ana","visas"))

    dfA = df_all.copy()
    if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
    if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
    if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
    if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
    if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

    # ---------- KPI ----------
    st.markdown("#### ✅ Indicateurs (vue filtrée)")
    k1, k2, k3, k4 = st.columns([0.9, 1.0, 1.0, 1.0])

    k1.metric("Dossiers", f"{len(dfA)}")
    k2.metric("Honoraires", _fmt_money(float(_safe_num_series(dfA, HONO).sum())))
    k3.metric("Payé", _fmt_money(float(_safe_num_series(dfA, "Payé").sum())))
    k4.metric("Reste", _fmt_money(float(_safe_num_series(dfA, "Reste").sum())))

    # ---------- Graphiques synthétiques ----------
    st.markdown("#### 📊 Graphiques synthétiques")

    g1, g2 = st.columns(2)

    # Dossiers par catégorie
    with g1:
        if not dfA.empty and "Categorie" in dfA.columns:
            vc = dfA["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie", "Nombre"]
            if not vc.empty:
                st.bar_chart(vc.set_index("Categorie"), use_container_width=True, key=skey("ana","bar_cat"))
            else:
                st.info("Aucune donnée pour Catégorie.")
        else:
            st.info("Aucune donnée pour Catégorie.")

    # Honoraires par mois
    with g2:
        if not dfA.empty and "Mois" in dfA.columns:
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str).str.zfill(2)
            gm = tmp.groupby("Mois", as_index=False)[HONO].sum().rename(columns={HONO:"Honoraires"}).sort_values("Mois")
            if not gm.empty:
                st.line_chart(gm.set_index("Mois"), use_container_width=True, key=skey("ana","line_hono"))
            else:
                st.info("Aucune donnée pour Honoraires par mois.")
        else:
            st.info("Aucune donnée pour Honoraires par mois.")

    # ---------- Détails (table) ----------
    st.markdown("#### 🧾 Détails (vue filtrée)")
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
        DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa", "Date", "Mois",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
    ] if c in det.columns]

    sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
    det_sorted = det.sort_values(by=sort_keys) if sort_keys else det

    st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=skey("ana","detail"))

    # ======================================================
    # 🔁 Comparaison Période A vs Période B (mois/années/catégories/sous-catégories/visas)
    # ======================================================
    st.markdown("---")
    st.markdown("### 🔁 Comparaison Période A vs Période B")

    if df_all.empty:
        st.info("Aucune donnée pour la comparaison.")
    else:
        # Sélecteurs A
        st.markdown("#### A) Période A")
        ca1, ca2, ca3, ca4, ca5 = st.columns([1.1, 1.1, 1.3, 1.3, 1.2])
        pa_years = ca1.multiselect("Année (A)", yearsA, default=[], key=skey("cmp","ya"))
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=skey("cmp","ma"))
        pa_cat   = ca3.multiselect("Catégorie (A)", catsA, default=[], key=skey("cmp","ca"))
        pa_sub   = ca4.multiselect("Sous-catégorie (A)", subsA, default=[], key=skey("cmp","sa"))
        pa_visa  = ca5.multiselect("Visa (A)", visasA, default=[], key=skey("cmp","va"))

        A = df_all.copy()
        if pa_years: A = A[A["_Année_"].isin(pa_years)]
        if pa_month: A = A[A["Mois"].astype(str).isin(pa_month)]
        if pa_cat:   A = A[A["Categorie"].astype(str).isin(pa_cat)]
        if pa_sub:   A = A[A["Sous-categorie"].astype(str).isin(pa_sub)]
        if pa_visa:  A = A[A["Visa"].astype(str).isin(pa_visa)]

        # Sélecteurs B
        st.markdown("#### B) Période B")
        cb1, cb2, cb3, cb4, cb5 = st.columns([1.1, 1.1, 1.3, 1.3, 1.2])
        pb_years = cb1.multiselect("Année (B)", yearsA, default=[], key=skey("cmp","yb"))
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=skey("cmp","mb"))
        pb_cat   = cb3.multiselect("Catégorie (B)", catsA, default=[], key=skey("cmp","cb"))
        pb_sub   = cb4.multiselect("Sous-catégorie (B)", subsA, default=[], key=skey("cmp","sb"))
        pb_visa  = cb5.multiselect("Visa (B)", visasA, default=[], key=skey("cmp","vb"))

        B = df_all.copy()
        if pb_years: B = B[B["_Année_"].isin(pb_years)]
        if pb_month: B = B[B["Mois"].astype(str).isin(pb_month)]
        if pb_cat:   B = B[B["Categorie"].astype(str).isin(pb_cat)]
        if pb_sub:   B = B[B["Sous-categorie"].astype(str).isin(pb_sub)]
        if pb_visa:  B = B[B["Visa"].astype(str).isin(pb_visa)]

        # KPI comparatifs
        st.markdown("#### 📌 KPI comparatifs")
        ck1, ck2, ck3, ck4, ck5, ck6 = st.columns([0.9, 1.0, 1.0, 1.0, 1.0, 1.0])

        def _kpis(df: pd.DataFrame) -> dict:
            return {
                "n": len(df),
                "h": float(_safe_num_series(df, HONO).sum()),
                "p": float(_safe_num_series(df, "Payé").sum()),
                "r": float(_safe_num_series(df, "Reste").sum()),
                "t": float(_safe_num_series(df, TOTAL).sum()),
            }

        KA = _kpis(A)
        KB = _kpis(B)

        ck1.metric("Dossiers A", f"{KA['n']}", delta=(KA['n'] - KB['n']))
        ck2.metric("Honoraires A", _fmt_money(KA["h"]), delta=_fmt_money(KA["h"] - KB["h"]))
        ck3.metric("Payé A", _fmt_money(KA["p"]), delta=_fmt_money(KA["p"] - KB["p"]))
        ck4.metric("Reste A", _fmt_money(KA["r"]), delta=_fmt_money(KA["r"] - KB["r"]))
        ck5.metric("Total A", _fmt_money(KA["t"]), delta=_fmt_money(KA["t"] - KB["t"]))
        # taux payé
        rateA = (KA["p"] / KA["t"] * 100.0) if KA["t"] > 0 else 0.0
        rateB = (KB["p"] / KB["t"] * 100.0) if KB["t"] > 0 else 0.0
        ck6.metric("% Payé A", f"{rateA:.1f}%", delta=f"{(rateA-rateB):.1f}%")

        # Graphiques comparaison (barres par mois)
        st.markdown("#### 📊 Comparaison par mois (A vs B)")

        cg1, cg2 = st.columns(2)

        with cg1:
            if not A.empty:
                tA = A.copy()
                tA["Mois"] = tA["Mois"].astype(str).str.zfill(2)
                gA = tA.groupby("Mois", as_index=False)[HONO].sum().rename(columns={HONO: "Honoraires A"})
                if not gA.empty:
                    st.bar_chart(gA.set_index("Mois"), use_container_width=True, key=skey("cmp","barA"))
                else:
                    st.info("Aucune donnée pour A (par mois).")
            else:
                st.info("Aucune donnée pour A (par mois).")

        with cg2:
            if not B.empty:
                tB = B.copy()
                tB["Mois"] = tB["Mois"].astype(str).str.zfill(2)
                gB = tB.groupby("Mois", as_index=False)[HONO].sum().rename(columns={HONO: "Honoraires B"})
                if not gB.empty:
                    st.bar_chart(gB.set_index("Mois"), use_container_width=True, key=skey("cmp","barB"))
                else:
                    st.info("Aucune donnée pour B (par mois).")
            else:
                st.info("Aucune donnée pour B (par mois).")

        # Détail côté A & B
        st.markdown("#### 🧾 Détails A & B")
        dA, dB = st.columns(2)

        def _fmt_table(df: pd.DataFrame) -> pd.DataFrame:
            res = df.copy()
            for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                if c in res.columns:
                    res[c] = _safe_num_series(res, c).apply(_fmt_money)
            if "Date" in res.columns:
                try:
                    res["Date"] = pd.to_datetime(res["Date"], errors="coerce").dt.date.astype(str)
                except Exception:
                    res["Date"] = res["Date"].astype(str)
            keep = [c for c in [
                DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
                "Date", "Mois", HONO, AUTRE, TOTAL, "Payé", "Reste"
            ] if c in res.columns]
            keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in res.columns]
            return (res.sort_values(by=keys) if keys else res)[keep].reset_index(drop=True)

        with dA:
            st.caption("**Période A**")
            st.dataframe(_fmt_table(A), use_container_width=True, key=skey("cmp","tableA"))
        with dB:
            st.caption("**Période B**")
            st.dataframe(_fmt_table(B), use_container_width=True, key=skey("cmp","tableB"))




# ================================
# 🛂 Visa Manager — PARTIE 4/4
# Onglets : 🏦 Escrow • 👤 Compte client • 🧾 Gestion • 📄 Visa • 💾 Export
# ================================

# Petite utilitaire locale sûre (au cas où)
def _date_for_widget(val):
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, date) else val.date()
    try:
        d2 = pd.to_datetime(val, errors="coerce")
        return d2.date() if pd.notna(d2) else None
    except Exception:
        return None

def _ensure_paylist(x):
    """Retourne une liste de paiements [{'date': 'YYYY-MM-DD', 'mode': 'CB', 'montant': 100.0}, ...]."""
    if isinstance(x, list):
        return x
    if pd.isna(x) or x is None:
        return []
    s = str(x).strip()
    if not s:
        return []
    try:
        obj = json.loads(s)
        return obj if isinstance(obj, list) else []
    except Exception:
        return []

# ================
# 🏦  ONGLET ESCROW
# ================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse & transferts potentiels")

    df_all = df_clients.copy()
    if df_all.empty:
        st.info("Aucun client.")
    else:
        for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
            if c not in df_all.columns:
                df_all[c] = 0.0

        # Montant honoraires qui peuvent être « logés » en escrow (avant envoi)
        # Hypothèse simple : ESCROW disponible = min(Payé, Honoraires) si dossier non encore envoyé.
        df_all["Payé_num"]  = _safe_num_series(df_all, "Payé")
        df_all["Hono_num"]  = _safe_num_series(df_all, HONO)
        df_all["Escrow dispo"] = 0.0

        sent_flag = df_all.get("Dossier envoyé", pd.Series([0]*len(df_all))).fillna(0)
        sent_flag = sent_flag.astype(int)

        not_sent_mask = (sent_flag != 1)
        df_all.loc[not_sent_mask, "Escrow dispo"] = np.minimum(df_all.loc[not_sent_mask, "Payé_num"],
                                                               df_all.loc[not_sent_mask, "Hono_num"])

        # KPI compacts
        e1, e2, e3 = st.columns([1,1,1])
        e1.metric("Honoraires totaux", _fmt_money(float(df_all["Hono_num"].sum())))
        e2.metric("Payé total", _fmt_money(float(df_all["Payé_num"].sum())))
        e3.metric("ESCROW disponible (non envoyés)", _fmt_money(float(df_all["Escrow dispo"].sum())))

        st.markdown("#### 📋 Dossiers non encore envoyés — honoraires payés à loger en ESCROW")
        cols_keep = [c for c in [DOSSIER_COL,"ID_Client","Nom","Categorie","Sous-categorie","Visa",
                                 HONO,"Payé","Escrow dispo"] if c in df_all.columns]
        view = df_all.loc[not_sent_mask, cols_keep].copy()
        if not view.empty:
            # format
            if HONO in view.columns:
                view[HONO] = _safe_num_series(view, HONO).apply(_fmt_money)
            if "Payé" in view.columns:
                view["Payé"] = _safe_num_series(view, "Payé").apply(_fmt_money)
            view["Escrow dispo"] = view["Escrow dispo"].apply(_fmt_money)
            st.dataframe(view.reset_index(drop=True), use_container_width=True, height=240, key=skey("esc","table"))
        else:
            st.info("Aucun dossier en attente d’envoi avec honoraires payés.")

        st.caption("ℹ️ Rappel : lorsqu’un « Dossier envoyé » est coché, l’ESCROW de ce dossier doit être transféré vers le compte d’encaissement.")

# ===========================
# 👤  ONGLET  COMPTE CLIENT
# ===========================
with tabs[3]:
    st.subheader("👤 Compte client — solde, paiements & statuts")

    df_all = df_clients.copy()
    if df_all.empty:
        st.info("Aucun client.")
    else:
        names = sorted(df_all.get("Nom", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
        ids   = sorted(df_all.get("ID_Client", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())

        s1, s2 = st.columns([1.5,1.0])
        sel_name = s1.selectbox("Nom", [""]+names, index=0, key=skey("acct","name"))
        sel_id   = s2.selectbox("ID_Client", [""]+ids, index=0, key=skey("acct","id"))

        mask = None
        if sel_id:
            mask = (df_all["ID_Client"].astype(str) == sel_id)
        elif sel_name:
            mask = (df_all["Nom"].astype(str) == sel_name)

        if mask is None or not mask.any():
            st.info("Sélectionnez un client.")
        else:
            row = df_all[mask].iloc[0].copy()

            # Montants numériques
            honor = float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0] if HONO in row else 0.0)
            other = float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0] if AUTRE in row else 0.0)
            total = float(_safe_num_series(pd.DataFrame([row]), TOTAL).iloc[0] if TOTAL in row else honor+other)
            paye  = float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0] if "Payé" in row else 0.0)
            reste = float(_safe_num_series(pd.DataFrame([row]), "Reste").iloc[0] if "Reste" in row else max(0.0, total-paye))

            k1,k2,k3,k4 = st.columns([1,1,1,1])
            k1.metric("Honoraires", _fmt_money(honor))
            k2.metric("Autres frais", _fmt_money(other))
            k3.metric("Payé", _fmt_money(paye))
            k4.metric("Reste", _fmt_money(reste))

            # Paiements
            st.markdown("#### 💵 Paiements")
            pay_list = _ensure_paylist(row.get("Paiements"))
            if len(pay_list):
                disp = pd.DataFrame(pay_list)
                # propreté
                if "montant" in disp.columns:
                    disp["montant"] = pd.to_numeric(disp["montant"], errors="coerce").fillna(0.0)
                st.dataframe(disp, use_container_width=True, height=220, key=skey("acct","pays"))
            else:
                st.info("Aucun paiement enregistré pour ce client.")

            if reste > 0:
                st.markdown("##### ➕ Ajouter un règlement")
                c1,c2,c3 = st.columns([1.2,1.2,1.0])
                pdate = c1.date_input("Date règlement", value=date.today(), key=skey("acct","pdate"))
                pmode = c2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=skey("acct","pmode"))
                pamt  = c3.number_input("Montant (US $)", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=skey("acct","pamt"))
                if st.button("💾 Ajouter ce paiement", key=skey("acct","addpay")):
                    add = float(pamt or 0.0)
                    if add <= 0:
                        st.warning("Montant invalide.")
                    else:
                        live = _read_clients(clients_path)
                        m2 = (live["ID_Client"].astype(str) == str(row.get("ID_Client","")))
                        if not m2.any():
                            st.error("Client introuvable.")
                        else:
                            i = live[m2].index[0]
                            pl = _ensure_paylist(live.at[i, "Paiements"] if "Paiements" in live.columns else [])
                            pl.append({
                                "date": (pdate if isinstance(pdate, (date, datetime)) else date.today()).strftime("%Y-%m-%d"),
                                "mode": pmode,
                                "montant": add
                            })
                            live.at[i, "Paiements"] = pl
                            # Recalcule Payé/Reste
                            live.at[i, "Payé"] = float(_safe_num_series(pd.DataFrame(pl), "montant").sum())
                            live.at[i, "Total (US $)"] = float(_safe_num_series(live.loc[[i]], HONO).iloc[0] +
                                                               _safe_num_series(live.loc[[i]], AUTRE).iloc[0])
                            live.at[i, "Reste"] = max(0.0, float(live.at[i, "Total (US $)"]) - float(live.at[i, "Payé"]))
                            _write_clients(live, clients_path)
                            st.success("Paiement ajouté.")
                            st.cache_data.clear()
                            st.rerun()

            st.markdown("#### 📌 Statuts du dossier")
            s1,s2,s3,s4,s5 = st.columns(5)
            envoye = bool(int(row.get("Dossier envoyé",0) or 0) == 1)
            accepte = bool(int(row.get("Dossier accepté",0) or 0) == 1)
            refuse  = bool(int(row.get("Dossier refusé",0) or 0) == 1)
            annule  = bool(int(row.get("Dossier annulé",0) or 0) == 1)
            rfe     = bool(int(row.get("RFE",0) or 0) == 1)

            s1.write(f"Date : {_safe_str(row.get('Date d'envoi',''))}")
            s2.write(f"Date : {_safe_str(row.get('Date d'acceptation',''))}")
            s3.write(f"Date : {_safe_str(row.get('Date de refus',''))}")
            s4.write(f"Date : {_safe_str(row.get('Date d'annulation',''))}")
            s5.write(f"RFE : {'Oui' if rfe else 'Non'}")

# ===========================
# 🧾  ONGLET  GESTION (CRUD)
# ===========================
with tabs[4]:
    st.subheader("🧾 Gestion des clients (Ajouter / Modifier / Supprimer)")

    df_live = _read_clients(clients_path)
    op = st.radio("Action", ["Ajouter","Modifier","Supprimer"], horizontal=True, key=skey("crud","op"))

    # -------- AJOUT --------
    if op == "Ajouter":
        c1,c2,c3 = st.columns(3)
        nom  = c1.text_input("Nom", "", key=skey("add","nom"))
        dval = date.today()
        dt   = c2.date_input("Date de création", value=dval, key=skey("add","date"))
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=dval.month-1, key=skey("add","mois"))

        # Visa cascade
        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        cat  = st.selectbox("Catégorie", [""]+cats, index=0, key=skey("add","cat"))
        sub  = ""
        visa_final = ""
        opts_dict = {"exclusive": None, "options": []}
        info_msg = ""
        if cat:
            subs = sorted(list(visa_map.get(cat, {}).keys()))
            sub  = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=skey("add","sub"))
            if sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, cat, sub, keyprefix=skey("add","opts"), preselected={}
                )
        if info_msg:
            st.info(info_msg)

        f1,f2 = st.columns(2)
        honor = f1.number_input(HONO, min_value=0.0, value=0.0, step=50.0, format="%.2f", key=skey("add","h"))
        other = f2.number_input(AUTRE, min_value=0.0, value=0.0, step=20.0, format="%.2f", key=skey("add","o"))
        st.text_area("Commentaire (autres frais)", "", key=skey("add","comm"), height=80)

        st.markdown("#### 📌 Statuts initiaux")
        s1,s2,s3,s4,s5 = st.columns(5)
        sent   = s1.checkbox("Dossier envoyé", key=skey("add","sent"))
        sent_d = s1.date_input("Date d'envoi", value=None, key=skey("add","sentd"))
        acc    = s2.checkbox("Dossier accepté", key=skey("add","acc"))
        acc_d  = s2.date_input("Date d'acceptation", value=None, key=skey("add","accd"))
        ref    = s3.checkbox("Dossier refusé", key=skey("add","ref"))
        ref_d  = s3.date_input("Date de refus", value=None, key=skey("add","refd"))
        ann    = s4.checkbox("Dossier annulé", key=skey("add","ann"))
        ann_d  = s4.date_input("Date d'annulation", value=None, key=skey("add","annd"))
        rfe    = s5.checkbox("RFE", key=skey("add","rfe"))
        if rfe and not any([sent,acc,ref,ann]):
            st.warning("RFE doit être associé à un autre statut (envoyé/accepté/refusé/annulé).")

        if st.button("💾 Enregistrer le client", key=skey("add","save")):
            if not nom:
                st.warning("Nom requis.")
            elif not cat or not sub:
                st.warning("Choisir Catégorie et Sous-catégorie.")
            else:
                total = float(honor)+float(other)
                dossier_n = _next_dossier(df_live, start=13057)
                did = _make_client_id(nom, dt)
                new_row = {
                    DOSSIER_COL: dossier_n, "ID_Client": did, "Nom": nom,
                    "Date": dt, "Mois": f"{int(mois):02d}",
                    "Categorie": cat, "Sous-categorie": sub,
                    "Visa": visa_final if visa_final else sub,
                    HONO: float(honor), AUTRE: float(other), TOTAL: total,
                    "Payé": 0.0, "Reste": total,
                    "Paiements": [], "Options": opts_dict,
                    "Dossier envoyé": 1 if sent else 0, "Date d'envoi": sent_d,
                    "Dossier accepté": 1 if acc else 0, "Date d'acceptation": acc_d,
                    "Dossier refusé": 1 if ref else 0, "Date de refus": ref_d,
                    "Dossier annulé": 1 if ann else 0, "Date d'annulation": ann_d,
                    "RFE": 1 if rfe else 0,
                }
                df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                _write_clients(df_new, clients_path)
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
            s1,s2 = st.columns(2)
            tname = s1.selectbox("Nom", [""]+names, index=0, key=skey("mod","selname"))
            tid   = s2.selectbox("ID_Client", [""]+ids, index=0, key=skey("mod","selid"))

            mask = None
            if tid:
                mask = (df_live["ID_Client"].astype(str) == tid)
            elif tname:
                mask = (df_live["Nom"].astype(str) == tname)

            if mask is None or not mask.any():
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()

            d1,d2,d3 = st.columns(3)
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=skey("mod","nom"))
            dval = _date_for_widget(row.get("Date")) or date.today()
            dt   = d2.date_input("Date de création", value=dval, key=skey("mod","date"))
            mois_default = _safe_str(row.get("Mois","01"))
            try:
                mois_index = max(0, int(mois_default) - 1)
            except Exception:
                mois_index = 0
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=mois_index, key=skey("mod","mois"))

            # Visa cascade (lecture valeurs)
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            cat  = st.selectbox("Catégorie", [""]+cats, index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                                key=skey("mod","cat"))

            sub = _safe_str(row.get("Sous-categorie",""))
            if cat:
                subs = sorted(list(visa_map.get(cat, {}).keys()))
                sub  = st.selectbox("Sous-catégorie", [""]+subs, index=(subs.index(sub)+1 if sub in subs else 0),
                                    key=skey("mod","sub"))

            preset_opts = row.get("Options", {})
            if not isinstance(preset_opts, dict):
                try:
                    preset_opts = json.loads(_safe_str(preset_opts) or "{}")
                    if not isinstance(preset_opts, dict):
                        preset_opts = {}
                except Exception:
                    preset_opts = {}
            visa_final, opts_dict, info_msg = "", {"exclusive": None, "options": []}, ""
            if cat and sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, cat, sub, keyprefix=skey("mod","opts"), preselected=preset_opts
                )
            if info_msg:
                st.info(info_msg)

            f1,f2 = st.columns(2)
            honor = f1.number_input(HONO, min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0]),
                                    step=50.0, format="%.2f", key=skey("mod","h"))
            other = f2.number_input(AUTRE, min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0]),
                                    step=20.0, format="%.2f", key=skey("mod","o"))
            st.text_area("Commentaire (autres frais)", _safe_str(row.get("Commentaire","")), key=skey("mod","comm"), height=80)

            st.markdown("#### 📌 Statuts")
            s1,s2,s3,s4,s5 = st.columns(5)
            sent   = s1.checkbox("Dossier envoyé", value=bool(int(row.get("Dossier envoyé",0) or 0)==1), key=skey("mod","sent"))
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=skey("mod","sentd"))
            acc    = s2.checkbox("Dossier accepté", value=bool(int(row.get("Dossier accepté",0) or 0)==1), key=skey("mod","acc"))
            acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=skey("mod","accd"))
            ref    = s3.checkbox("Dossier refusé", value=bool(int(row.get("Dossier refusé",0) or 0)==1), key=skey("mod","ref"))
            ref_d  = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=skey("mod","refd"))
            ann    = s4.checkbox("Dossier annulé", value=bool(int(row.get("Dossier annulé",0) or 0)==1), key=skey("mod","ann"))
            ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=skey("mod","annd"))
            rfe    = s5.checkbox("RFE", value=bool(int(row.get("RFE",0) or 0)==1), key=skey("mod","rfe"))

            if st.button("💾 Enregistrer les modifications", key=skey("mod","save")):
                if not nom:
                    st.warning("Nom requis.")
                    st.stop()
                if not cat or not sub:
                    st.warning("Choisir Catégorie et Sous-catégorie.")
                    st.stop()

                df_live.at[idx, "Nom"]  = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(mois):02d}"
                df_live.at[idx, "Categorie"] = cat
                df_live.at[idx, "Sous-categorie"] = sub
                df_live.at[idx, "Visa"] = visa_final if visa_final else sub
                df_live.at[idx, HONO] = float(honor)
                df_live.at[idx, AUTRE] = float(other)
                df_live.at[idx, TOTAL] = float(honor) + float(other)
                # recalcul reste en conservant Payé existant
                pay_list = _ensure_paylist(df_live.at[idx, "Paiements"] if "Paiements" in df_live.columns else [])
                pay_sum = float(_safe_num_series(pd.DataFrame(pay_list), "montant").sum()) if len(pay_list) else float(_safe_num_series(df_live.loc[[idx]], "Payé").iloc[0])
                df_live.at[idx, "Payé"] = pay_sum
                df_live.at[idx, "Reste"] = max(0.0, float(df_live.at[idx, TOTAL]) - pay_sum)
                df_live.at[idx, "Options"] = opts_dict
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
                st.success("Modifications enregistrées.")
                st.cache_data.clear()
                st.rerun()

    # ------ SUPPRIMER ------
    elif op == "Supprimer":
        if df_live.empty:
            st.info("Aucun client à supprimer.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
            s1,s2 = st.columns(2)
            tname = s1.selectbox("Nom", [""]+names, index=0, key=skey("del","name"))
            tid   = s2.selectbox("ID_Client", [""]+ids, index=0, key=skey("del","id"))

            mask = None
            if tid:
                mask = (df_live["ID_Client"].astype(str) == tid)
            elif tname:
                mask = (df_live["Nom"].astype[str] == tname)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({"Dossier N": row.get(DOSSIER_COL,""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=skey("del","ok")):
                    df_new = df_live[~mask].copy()
                    _write_clients(df_new, clients_path)
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()

# ======================
# 📄  ONGLET  VISA
# ======================
with tabs[5]:
    st.subheader("📄 Visa — aperçu structure")
    if df_visa_raw is None or df_visa_raw.empty:
        st.info("Aucune feuille Visa chargée.")
    else:
        # Aperçu Catégorie/Sous-catégorie + colonnes d’options disponibles (ligne 1)
        cols = df_visa_raw.columns.tolist()
        st.write("Colonnes :", ", ".join(cols))
        st.dataframe(df_visa_raw, use_container_width=True, height=300, key=skey("visa","raw"))

        st.markdown("#### 🎯 Options détectées (par sous-catégorie)")
        st.write("Les options (cases à cocher / exclusives) proviennent des en-têtes marquées **[X]** ou **(…)** en ligne 1.")
        st.json(visa_map, expanded=False)

# ======================
# 💾  EXPORT GLOBAL
# ======================
st.markdown("---")
st.subheader("💾 Export global (Clients + Visa)")
colz1, colz2 = st.columns([1,3])

with colz1:
    if st.button("Préparer ZIP", key=skey("zip","prep")):
        try:
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients « propres »
                df_export = _read_clients(clients_path)
                with BytesIO() as xbuf:
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                        df_export.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                    zf.writestr("Clients.xlsx", xbuf.getvalue())
                # Visa (fichier tel quel si possible)
                try:
                    zf.write(visa_path, "Visa.xlsx")
                except Exception:
                    try:
                        dfv0 = pd.read_excel(visa_path, sheet_name=SHEET_VISA)
                        with BytesIO() as vb:
                            with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                dfv0.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                            zf.writestr("Visa.xlsx", vb.getvalue())
                    except Exception:
                        pass
            st.session_state[skey("zip","data")] = buf.getvalue()
            st.success("Archive prête.")
        except Exception as e:
            st.error("Erreur de préparation : " + _safe_str(e))

with colz2:
    z = st.session_state.get(skey("zip","data"))
    if z:
        st.download_button(
            "⬇️ Télécharger l’archive (ZIP)",
            data=z,
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
            key=skey("zip","dl")
        )