# ==============================================
# 🛂 VISA MANAGER — Application Streamlit (full)
# ==============================================

from __future__ import annotations

import streamlit as st
import pandas as pd
import json
from datetime import date, datetime
from pathlib import Path
from io import BytesIO
import unicodedata
from uuid import uuid4
import zipfile

# -------- Page & titre
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

# Espace de noms unique pour éviter les collisions de clés Streamlit
SID = st.session_state.setdefault("WIDGET_NS", str(uuid4()))

# ==============================================
# Mémoire des derniers chemins de fichiers
# ==============================================
LAST_PATHS_FILE = ".visa_manager_last.json"

def _load_last_paths() -> dict:
    try:
        p = Path(LAST_PATHS_FILE)
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def _save_last_paths(clients_path: str, visa_path: str) -> None:
    try:
        payload = {"clients": clients_path, "visa": visa_path}
        Path(LAST_PATHS_FILE).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

# ==============================================
# Constantes & colonnes attendues
# ==============================================
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

CLIENTS_COLS = [
    "Dossier N","ID_Client","Nom","Date","Mois",
    "Categorie","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
    "Payé","Reste","Paiements","Options",
    "Dossier envoyé","Date d'envoi",
    "Dossier accepté","Date d'acceptation",
    "Dossier refusé","Date de refus",
    "Dossier annulé","Date d'annulation",
    "RFE"
]

# Valeurs par défaut + surcharge depuis la mémoire
CLIENTS_FILE_FALLBACK = "donnees_visa_clients1_adapte.xlsx"
VISA_FILE_FALLBACK    = "donnees_visa_clients1.xlsx"
_last = _load_last_paths()
CLIENTS_FILE_DEFAULT = _last.get("clients", CLIENTS_FILE_FALLBACK)
VISA_FILE_DEFAULT    = _last.get("visa",    VISA_FILE_FALLBACK)

# ==============================================
# Utilitaires généraux
# ==============================================
def _safe_str(x) -> str:
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)

def _to_num(s: pd.Series) -> pd.Series:
    s = s.astype(str)
    s = s.str.replace(r"[^\d,.\-]", "", regex=True).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df), index=df.index, dtype=float)
    v = df[col]
    if pd.api.types.is_numeric_dtype(v):
        return v.fillna(0.0).astype(float)
    return _to_num(v)

def _fmt_money(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _uniquify_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(map(str, df.columns))
    seen, out = {}, []
    for c in cols:
        if c not in seen:
            seen[c] = 1
            out.append(c)
        else:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
    df2 = df.copy()
    df2.columns = out
    return df2

def ensure_file(path: str, sheet_name: str, cols: list[str]) -> None:
    p = Path(path)
    if not p.exists():
        df = pd.DataFrame(columns=cols)
        with pd.ExcelWriter(p, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=sheet_name, index=False)

def _norm(s: str) -> str:
    s2 = unicodedata.normalize("NFKD", s)
    s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
    s2 = s2.strip().lower().replace("\u00a0", " ")
    s2 = s2.replace("-", " ").replace("_", " ")
    return " ".join(s2.split())

# Dates sûres pour widgets / enregistrement
def _date_for_widget(val):
    """Toujours une date valide pour st.date_input (fallback = aujourd'hui)."""
    try:
        if isinstance(val, date) and not isinstance(val, datetime):
            return val
        if isinstance(val, datetime):
            return val.date()
        d = pd.to_datetime(val, errors="coerce")
        return d.date() if pd.notna(d) else date.today()
    except Exception:
        return date.today()

def _date_or_none(val):
    """Retourne date (date) ou None pour stocker dans fichier."""
    try:
        if isinstance(val, date) and not isinstance(val, datetime):
            return val
        if isinstance(val, datetime):
            return val.date()
        d = pd.to_datetime(val, errors="coerce")
        return d.date() if pd.notna(d) else None
    except Exception:
        return None

# Crée des fichiers vides si absents
ensure_file(CLIENTS_FILE_DEFAULT, SHEET_CLIENTS, CLIENTS_COLS)
ensure_file(VISA_FILE_DEFAULT, SHEET_VISA, ["Categorie","Sous-categorie","COS","EOS"])

# ==============================================
# Parsing de la feuille Visa → visa_map {cat:{sub:[options]}}
# ==============================================
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    """
    Construit un mapping: {Categorie: {Sous-categorie: [options...]}}
    - Chaque option correspond à une colonne cochée (=1, x, oui, true...) sur la ligne de la sous-catégorie.
    - Si aucune option cochée, on conserve la sous-catégorie seule comme option.
    - Injection automatique de la catégorie '2-Etudiants' si absente, avec F-1/F-2 COS/EOS.
    """
    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        if isinstance(v, (int, float)):
            return float(v) == 1.0
        s = str(v).strip().lower()
        return s in {"1","x","true","vrai","oui","yes"}

    out: dict[str, dict[str, list[str]]] = {}
    found_students = False

    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        xls = None

    sheets = [sheet_name] if (sheet_name and xls) else (xls.sheet_names if xls else [])
    if sheets:
        for sn in sheets:
            try:
                dfv = pd.read_excel(xlsx_path, sheet_name=sn)
            except Exception:
                continue
            if dfv.empty:
                continue

            dfv = _uniquify_columns(dfv)
            dfv.columns = dfv.columns.map(str).str.strip()

            cmap = {_norm(c): c for c in dfv.columns}
            cat_col = next((cmap[k] for k in cmap if "categorie" in k), None)
            sub_col = next((cmap[k] for k in cmap if k.startswith("sous")), None)
            if not cat_col:
                continue
            if not sub_col:
                dfv["_Sous_"] = ""
                sub_col = "_Sous_"

            check_cols = [c for c in dfv.columns if c not in {cat_col, sub_col}]

            cats_in_sheet = dfv[cat_col].dropna().astype(str).map(str.strip)
            if any("etudiant" in _norm(c) for c in cats_in_sheet):
                found_students = True

            for _, row in dfv.iterrows():
                cat = _safe_str(row.get(cat_col, "")).strip()
                sub = _safe_str(row.get(sub_col, "")).strip()
                if not cat:
                    continue
                opts = []
                for cc in check_cols:
                    if _is_checked(row.get(cc)):
                        # libellé d'option = "Sous-categorie + nom de colonne"
                        lab = (f"{sub} {cc}".strip() if sub else str(cc).strip())
                        opts.append(lab)
                if not opts and sub:
                    opts = [sub]
                if opts:
                    out.setdefault(cat, {})
                    out[cat].setdefault(sub, [])
                    out[cat][sub].extend(opts)

    # Injection pour catégories contenant 'etudiant'
    for cat_name in list(out.keys()):
        if "etudiant" in _norm(cat_name):
            submap = out.setdefault(cat_name, {})
            for sub in ("F-1","F-2"):
                arr = submap.setdefault(sub, [])
                for w in (f"{sub} COS", f"{sub} EOS"):
                    if w not in arr:
                        arr.append(w)
                submap[sub] = sorted(set(arr))

    # Si aucune catégorie étudiants trouvée, on ajoute '2-Etudiants'
    if not found_students:
        out.setdefault("2-Etudiants", {})
        out["2-Etudiants"].setdefault("F-1", ["F-1 COS", "F-1 EOS"])
        out["2-Etudiants"].setdefault("F-2", ["F-2 COS", "F-2 EOS"])

    # Nettoyage doublons
    for cat, subs in out.items():
        for sub, arr in subs.items():
            subs[sub] = sorted(set(arr))
    return out

# ==============================================
# I/O & normalisation des Clients
# ==============================================
def _normalize_options_json(x) -> dict:
    try:
        d = json.loads(_safe_str(x) or "{}")
        if not isinstance(d, dict):
            return {}
        excl = d.get("exclusive", None)
        opts = d.get("options", [])
        if not isinstance(opts, list):
            opts = []
        return {"exclusive": excl, "options": [str(o) for o in opts]}
    except Exception:
        return {"exclusive": None, "options": []}

def _normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # colonnes manquantes
    for c in CLIENTS_COLS:
        if c not in df.columns:
            df[c] = None

    # Dates
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    for c in ["Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.date

    # Mois
    df["Mois"] = df.apply(
        lambda r: f"{pd.to_datetime(r['Date']).month:02d}" if pd.notna(r["Date"]) else (_safe_str(r.get("Mois",""))[:2] or None),
        axis=1
    )

    # Montants
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        df[c] = _safe_num_series(df, c)

    # Paiements (JSON list)
    def _parse_p(x):
        try:
            j = json.loads(_safe_str(x) or "[]")
            return j if isinstance(j, list) else []
        except Exception:
            return []
    df["Paiements"] = df["Paiements"].apply(_parse_p)

    def _sum_json(lst):
        try:
            return float(sum(float(it.get("amount",0.0) or 0.0) for it in (lst or [])))
        except Exception:
            return 0.0
    paid_json = df["Paiements"].apply(_sum_json)
    df["Payé"] = pd.concat([df["Payé"].fillna(0.0).astype(float), paid_json], axis=1).max(axis=1)

    df["Total (US $)"] = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]
    df["Reste"] = (df["Total (US $)"] - df["Payé"]).clip(lower=0.0)

    # Options (dict JSON)
    df["Options"] = df["Options"].apply(_normalize_options_json)

    # Statuts -> bool
    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        df[c] = df[c].apply(lambda v: bool(str(v).strip().lower() in {"1","true","vrai","oui","yes","x"}))

    # Index temporels auxiliaires
    df["_Année_"]   = df["Date"].apply(lambda d: d.year if pd.notna(d) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda d: d.month if pd.notna(d) else pd.NA)
    return _uniquify_columns(df)

@st.cache_data(show_spinner=False)
def _read_clients(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        df = pd.DataFrame(columns=CLIENTS_COLS)
    return _normalize_clients(df)

def _write_clients(df: pd.DataFrame, path: str) -> None:
    """Ecrit le fichier Clients et pousse l'état précédent dans la pile UNDO."""
    st.session_state.setdefault("undo_stack", [])
    try:
        prev = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        prev = pd.DataFrame(columns=CLIENTS_COLS)
    st.session_state["undo_stack"].append(prev.copy())

    df2 = df.copy()
    df2["Options"] = df2["Options"].apply(lambda d: json.dumps(_normalize_options_json(d), ensure_ascii=False))
    df2["Paiements"] = df2["Paiements"].apply(lambda l: json.dumps(l, ensure_ascii=False))
    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        df2[c] = df2[c].apply(lambda b: 1 if bool(b) else 0)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        _uniquify_columns(df2).to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)

def undo_last_write(path: str):
    stack = st.session_state.get("undo_stack", [])
    if not stack:
        st.warning("Aucune opération à annuler.")
        return
    prev_df = stack.pop()
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        _uniquify_columns(prev_df).to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
    st.success("Dernière action annulée.")

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if "Dossier N" in df.columns:
        s = pd.to_numeric(df["Dossier N"], errors="coerce")
        if s.notna().any():
            return int(s.max()) + 1
    return int(start)

# --- Identifiant client robuste (prend str/date/datetime/None) ---
def _make_client_id(nom: str, d) -> str:
    base = _safe_str(nom).strip()
    base = unicodedata.normalize("NFKD", base)
    base = "".join(ch for ch in base if not unicodedata.combining(ch))
    base = base.replace(" ", "_").replace("/", "-").replace("\\", "-")
    base = base.strip("_-") or "Client"
    try:
        if not isinstance(d, (date, datetime)):
            d = pd.to_datetime(d, errors="coerce")
        if pd.isna(d):
            d = datetime.today()
        if isinstance(d, date) and not isinstance(d, datetime):
            d = datetime(d.year, d.month, d.day)
        stamp = d.strftime("%Y%m%d")
    except Exception:
        stamp = datetime.today().strftime("%Y%m%d")
    return f"{base}-{stamp}"

# ==============================================
# Helper UI — sélecteur d'options VISA (radio + cases)
# ==============================================
def build_visa_option_selector(visa_map: dict, cat: str, sub: str, keyprefix: str, preselected: dict | None = None):
    """
    UI des options Visa pour (cat, sub) :
      - Radio 'exclusive' si options de la forme 'sub XXX' (ex: 'F-1 COS', 'F-1 EOS')
      - Cases à cocher pour les autres options
    Retourne (visa_final, opts_dict, info_msg).
    """
    arr = visa_map.get(cat, {}).get(sub, [])
    prefix = f"{sub} "
    suffixes = sorted({o[len(prefix):] for o in arr if o.startswith(prefix) and len(o) > len(prefix)})
    others = sorted([o for o in arr if not (o.startswith(prefix) and len(o) > len(prefix))])

    preselected = preselected or {}
    pre_excl = preselected.get("exclusive")
    pre_opts = preselected.get("options", []) if isinstance(preselected.get("options", []), list) else []

    chosen_excl = None
    if suffixes:
        radio_opts = [""] + suffixes
        default_idx = radio_opts.index(pre_excl) if pre_excl in radio_opts else 0
        chosen_excl = st.radio(
            f"Option exclusive — {sub}",
            options=radio_opts,
            index=default_idx,
            key=f"{keyprefix}_excl"
        )

    chosen_multi = []
    for i, lab in enumerate(others):
        default = lab in pre_opts
        if st.checkbox(lab, value=default, key=f"{keyprefix}_chk_{i}"):
            chosen_multi.append(lab)

    visa_final = sub
    if chosen_excl:
        visa_final = f"{sub} {chosen_excl}".strip()

    info_msg = "" if arr else "Aucune option cochée pour cette sous-catégorie dans la feuille Visa."
    return visa_final, {"exclusive": (chosen_excl or None), "options": chosen_multi}, info_msg

# ==============================================
# Barre latérale (fichiers, uploads, mémorisation, UNDO)
# ==============================================
with st.sidebar:
    st.header("🧭 Navigation")

    # Chemins mémorisés (modifiables)
    clients_path = st.text_input("Fichier Clients", value=CLIENTS_FILE_DEFAULT, key=f"sb_clients_path_{SID}")
    visa_path    = st.text_input("Fichier Visa",    value=VISA_FILE_DEFAULT,    key=f"sb_visa_path_{SID}")

    # Uploads optionnels (écrasent le fichier indiqué et mémorisent)
    up_c = st.file_uploader("Charger Excel Clients (remplace le fichier indiqué)", type=["xlsx"], key=f"up_c_{SID}")
    if up_c is not None:
        try:
            Path(clients_path).write_bytes(up_c.getvalue())
            _save_last_paths(clients_path, visa_path)
            st.success("Fichier Clients importé et mémorisé.")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"Import Clients impossible : {e}")

    up_v = st.file_uploader("Charger Excel Visa (remplace le fichier indiqué)", type=["xlsx"], key=f"up_v_{SID}")
    if up_v is not None:
        try:
            Path(visa_path).write_bytes(up_v.getvalue())
            _save_last_paths(clients_path, visa_path)
            st.success("Fichier Visa importé et mémorisé.")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"Import Visa impossible : {e}")

    # Bouton pour mémoriser explicitement les chemins saisis
    if st.button("💾 Mémoriser ces chemins par défaut", key=f"save_paths_{SID}"):
        _save_last_paths(clients_path, visa_path)
        st.success("Chemins mémorisés.")
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.subheader("🧾 Gestion")
    if st.button("↩️ Annuler dernière action (UNDO)", key=f"undo_{SID}"):
        undo_last_write(clients_path)
        st.cache_data.clear()
        st.rerun()

# Sauvegarde auto si l’utilisateur modifie les champs (qualité de vie)
if clients_path != CLIENTS_FILE_DEFAULT or visa_path != VISA_FILE_DEFAULT:
    _save_last_paths(clients_path, visa_path)

# ==============================================
# Chargement des données
# ==============================================
visa_map = parse_visa_sheet(visa_path)
df_all   = _read_clients(clients_path)

# ==============================================
# Création des onglets
# ==============================================
tabs = st.tabs([
    "📊 Dashboard", "📈 Analyses", "🏦 Escrow",
    "🧾 Visa (gestion)", "📄 Visa (aperçu)", "👤 Clients"
])

# ==============================================
# 📊 DASHBOARD
# ==============================================
with tabs[0]:
    st.subheader("📊 Dashboard — tous les clients")

    # Filtres dashboard (barre latérale)
    with st.sidebar:
        st.subheader("🔎 Filtres Dashboard")
        if df_all.empty:
            years = []; months = []; cats = []; subs = []; visas = []
        else:
            years  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
            months = [f"{m:02d}" for m in range(1,13)]
            cats   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist())
            subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist())
            visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist())

        dash_years  = st.multiselect("Année", years, default=[], key=f"dash_years_{SID}")
        dash_months = st.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{SID}")
        dash_cats   = st.multiselect("Catégories", cats, default=[], key=f"dash_cats_{SID}")
        dash_subs   = st.multiselect("Sous-catégories", subs, default=[], key=f"dash_subs_{SID}")
        dash_visas  = st.multiselect("Visa", visas, default=[], key=f"dash_visas_{SID}")

    df = df_all.copy()
    if dash_years:  df = df[df["_Année_"].isin(dash_years)]
    if dash_months: df = df[df["Mois"].isin(dash_months)]
    if dash_cats:   df = df[df["Categorie"].astype(str).isin(dash_cats)]
    if dash_subs:   df = df[df["Sous-categorie"].astype(str).isin(dash_subs)]
    if dash_visas:  df = df[df["Visa"].astype(str).isin(dash_visas)]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(df)}")
    k2.metric("Honoraires", _fmt_money(_safe_num_series(df,"Montant honoraires (US $)").sum()))
    k3.metric("Payé", _fmt_money(_safe_num_series(df,"Payé").sum()))
    k4.metric("Solde", _fmt_money(_safe_num_series(df,"Reste").sum()))

    # Détails (tri sûr)
    view = df.copy()
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        if c in view.columns:
            view[c] = _safe_num_series(view, c).map(_fmt_money)
    if "Date" in view.columns:
        try:
            view["Date"] = pd.to_datetime(view["Date"], errors="coerce").dt.date.astype(str)
        except Exception:
            view["Date"] = view["Date"].astype(str)

    show_cols = [c for c in [
        "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
        "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
        "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
    ] if c in view.columns]

    sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
    view_sorted = view.sort_values(by=sort_keys) if sort_keys else view
    st.dataframe(view_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_tbl_{SID}")

# ==============================================
# 📈 ANALYSES
# ==============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_years_{SID}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"a_months_{SID}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"a_cats_{SID}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"a_subs_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_visas_{SID}")

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPI
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(dfA, "Montant honoraires (US $)").sum()))
        k3.metric("Payé", _fmt_money(_safe_num_series(dfA, "Payé").sum()))
        k4.metric("Reste", _fmt_money(_safe_num_series(dfA, "Reste").sum()))

        # Graphiques
        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("### 📊 Dossiers par catégorie")
            vc = dfA["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie", "Nombre"]
            st.bar_chart(vc.set_index("Categorie"))

        if not dfA.empty and "Mois" in dfA.columns:
            st.markdown("### 📈 Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)["Montant honoraires (US $)"].sum().sort_values("Mois")
            st.line_chart(gm.set_index("Mois"))

        # Tableau détaillé
        st.markdown("### 🧾 Détails des dossiers filtrés")
        det = dfA.copy()
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Total (US $)", "Payé", "Reste"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).apply(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det

        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"a_detail_{SID}")

# ==============================================
# 🏦 ESCROW — synthèse
# ==============================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Payé"]  = _safe_num_series(dfE, "Payé")
        dfE["Reste"] = _safe_num_series(dfE, "Reste")
        dfE["Total (US $)"] = _safe_num_series(dfE, "Total (US $)")

        agg = dfE.groupby("Categorie", as_index=False)[["Total (US $)", "Payé", "Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg["Total (US $)"]).replace([pd.NA, pd.NaT], 0).fillna(0.0) * 100
        st.dataframe(agg, use_container_width=True, key=f"esc_agg_{SID}")

        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money(float(dfE["Total (US $)"].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Reste", _fmt_money(float(dfE["Reste"].sum())))

        st.caption("NB : pour un escrow « strict », on peut isoler les honoraires perçus avant envoi, puis signaler les transferts à effectuer une fois « Dossier envoyé » coché.")

# ==============================================
# 🧾 VISA (GESTION) — CRUD sur la feuille Visa
# ==============================================
def _read_visa_raw(path: str) -> pd.DataFrame:
    try:
        dfv = pd.read_excel(path, sheet_name=SHEET_VISA)
    except Exception:
        dfv = pd.DataFrame(columns=["Categorie","Sous-categorie","COS","EOS"])
    # normaliser nom de colonne sous-catégorie
    cols = [("Sous-categorie" if _norm(c).startswith("sous") else c) for c in dfv.columns]
    dfv.columns = cols
    return dfv

def _write_visa_raw(dfv: pd.DataFrame, path: str):
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        _uniquify_columns(dfv).to_excel(wr, sheet_name=SHEET_VISA, index=False)

with tabs[3]:
    st.subheader("🧾 Visa — Gestion (ajouter / modifier / supprimer)")
    dfv = _read_visa_raw(visa_path).copy()

    st.markdown("#### Aperçu actuel")
    st.dataframe(dfv, use_container_width=True, key=f"visa_view_{SID}")

    st.markdown("### ➕ Ajouter une ligne")
    with st.form(key=f"visa_add_{SID}", clear_on_submit=True):
        c1, c2 = st.columns(2)
        v_cat = c1.text_input("Catégorie", "")
        v_sub = c2.text_input("Sous-catégorie", "")
        # options existantes comme colonnes
        opt_cols = [c for c in dfv.columns if c not in {"Categorie","Sous-categorie"}]
        st.caption("Cochez les options à 1 si applicable (les colonnes non listées peuvent être ajoutées plus bas).")
        chosen = {}
        oc1, oc2, oc3 = st.columns(3)
        buckets = [oc1, oc2, oc3]
        for i, c in enumerate(opt_cols):
            chosen[c] = buckets[i % 3].checkbox(c, value=False, key=f"visa_add_opt_{SID}_{i}")

        submitted = st.form_submit_button("💾 Ajouter")
        if submitted:
            if not v_cat.strip():
                st.warning("Catégorie obligatoire.")
            else:
                new_row = {"Categorie": v_cat.strip(), "Sous-categorie": v_sub.strip()}
                for c in opt_cols:
                    new_row[c] = 1 if chosen.get(c) else None
                dfv_new = pd.concat([dfv, pd.DataFrame([new_row])], ignore_index=True)
                _write_visa_raw(dfv_new, visa_path)
                st.success("Ligne ajoutée au Visa.")
                st.cache_data.clear()
                st.experimental_rerun()

    st.markdown("---")
    st.markdown("### ✏️ Modifier / 🗑️ Supprimer une ligne existante")
    if dfv.empty:
        st.info("Aucune ligne Visa à modifier.")
    else:
        # Sélection de la ligne
        idxs = dfv.index.tolist()
        labels = [f"{i} — {dfv.loc[i, 'Categorie']} / {dfv.loc[i, 'Sous-categorie']}" for i in idxs]
        sel_idx_label = st.selectbox("Sélectionnez une ligne :", labels, key=f"visa_sel_{SID}")
        sel_idx = int(sel_idx_label.split(" — ")[0]) if sel_idx_label else idxs[0]

        # Formulaire de modification
        row = dfv.loc[sel_idx].copy()
        m1, m2 = st.columns(2)
        new_cat = m1.text_input("Catégorie", _safe_str(row.get("Categorie","")), key=f"visa_mod_cat_{SID}")
        new_sub = m2.text_input("Sous-catégorie", _safe_str(row.get("Sous-categorie","")), key=f"visa_mod_sub_{SID}")

        opt_cols = [c for c in dfv.columns if c not in {"Categorie","Sous-categorie"}]
        st.caption("Options (1 = coché)")
        mc1, mc2, mc3 = st.columns(3)
        buckets = [mc1, mc2, mc3]
        opts_values = {}
        for i, c in enumerate(opt_cols):
            current = str(row.get(c,"")).strip().lower() in {"1","x","true","oui","yes"}
            opts_values[c] = buckets[i % 3].checkbox(c, value=current, key=f"visa_mod_opt_{SID}_{i}")

        col_btn1, col_btn2, col_btn3 = st.columns(3)
        if col_btn1.button("💾 Enregistrer les modifications", key=f"visa_save_{SID}"):
            dfv.at[sel_idx, "Categorie"] = new_cat.strip()
            dfv.at[sel_idx, "Sous-categorie"] = new_sub.strip()
            for c, v in opts_values.items():
                dfv.at[sel_idx, c] = 1 if v else None
            _write_visa_raw(dfv, visa_path)
            st.success("Ligne mise à jour.")
            st.cache_data.clear()
            st.experimental_rerun()

        if col_btn2.button("🗑️ Supprimer cette ligne", key=f"visa_del_{SID}"):
            dfv2 = dfv.drop(index=sel_idx).reset_index(drop=True)
            _write_visa_raw(dfv2, visa_path)
            st.success("Ligne supprimée.")
            st.cache_data.clear()
            st.experimental_rerun()

        # Gestion des colonnes d'options
        st.markdown("---")
        st.markdown("### ⚙️ Colonnes d’options (ajouter / supprimer)")
        nc = st.text_input("Ajouter une colonne d’option (ex: Premium, COS, EOS…)", key=f"visa_addcol_{SID}")
        if st.button("➕ Ajouter la colonne", key=f"visa_addcol_btn_{SID}"):
            if nc.strip() and nc.strip() not in dfv.columns:
                dfv[nc.strip()] = None
                _write_visa_raw(dfv, visa_path)
                st.success(f"Colonne '{nc.strip()}' ajoutée.")
                st.cache_data.clear()
                st.experimental_rerun()
            else:
                st.warning("Nom invalide ou déjà existant.")

        if opt_cols:
            sc = st.selectbox("Supprimer une colonne d’option :", [""] + opt_cols, key=f"visa_delcol_{SID}")
            if sc and st.button("🗑️ Supprimer la colonne", key=f"visa_delcol_btn_{SID}"):
                dfv2 = dfv.drop(columns=[sc])
                _write_visa_raw(dfv2, visa_path)
                st.success(f"Colonne '{sc}' supprimée.")
                st.cache_data.clear()
                st.experimental_rerun()

# ==============================================
# 📄 VISA (APERÇU)
# ==============================================
with tabs[4]:
    st.subheader("📄 Visa — aperçu et filtre")
    visa_map = parse_visa_sheet(visa_path)  # reparse en cas de modif

    cats = sorted(list(visa_map.keys()))
    sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=f"vprev_cat_{SID}")
    subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
    sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=f"vprev_sub_{SID}")

    if sel_cat and sel_sub:
        # Montrer options dynamiques
        st.markdown("#### Options disponibles")
        visa_final, opts_dict, info_msg = build_visa_option_selector(
            visa_map, sel_cat, sel_sub, keyprefix=f"vprev_opts_{SID}", preselected={}
        )
        st.info(f"Visa sélectionné : **{visa_final}**")
        if info_msg:
            st.caption(info_msg)
    elif sel_cat:
        st.info("Choisissez une sous-catégorie.")

# ==============================================
# 👤 CLIENTS — Gestion (CRUD complet + statuts + paiements)
# ==============================================
with tabs[5]:
    st.subheader("👤 Clients — Ajouter / Modifier / Supprimer")

    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_op_{SID}")
    df_live = _read_clients(clients_path)

    # ---------- AJOUT ----------
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")

        c1, c2, c3 = st.columns(3)
        nom = c1.text_input("Nom", "", key=f"add_nom_{SID}")
        dt  = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=int(date.today().month)-1, key=f"add_mois_{SID}")

        # Cascade Visa
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
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f",
                                key=f"add_h_{SID}")
        other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f",
                                key=f"add_o_{SID}")

        st.markdown("#### 📌 Statuts initiaux")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent    = s1.checkbox("Dossier envoyé", key=f"add_sent_{SID}")
        sent_d  = s1.date_input("Date d'envoi", value=_date_for_widget(None), key=f"add_sentd_{SID}")
        acc     = s2.checkbox("Dossier accepté", key=f"add_acc_{SID}")
        acc_d   = s2.date_input("Date d'acceptation", value=_date_for_widget(None), key=f"add_accd_{SID}")
        ref     = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
        ref_d   = s3.date_input("Date de refus", value=_date_for_widget(None), key=f"add_refd_{SID}")
        ann     = s4.checkbox("Dossier annulé", key=f"add_ann_{SID}")
        ann_d   = s4.date_input("Date d'annulation", value=_date_for_widget(None), key=f"add_annd_{SID}")
        rfe     = s5.checkbox("RFE", key=f"add_rfe_{SID}")

        if rfe and not any([sent, acc, ref, ann]):
            st.warning("⚠️ La case RFE ne peut être cochée qu’avec un autre statut (envoyé, accepté, refusé ou annulé).")

        save_add = st.button("💾 Enregistrer le client", key=f"btn_add_{SID}")
        if save_add:
            if not nom:
                st.warning("Veuillez saisir le nom.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Veuillez choisir la catégorie et la sous-catégorie.")
            else:
                total = float(honor) + float(other)
                paye  = 0.0
                reste = max(0.0, total - paye)
                did = _make_client_id(nom, dt)
                dossier_n = _next_dossier(df_live, start=13057)

                new_row = {
                    "Dossier N": dossier_n,
                    "ID_Client": did,
                    "Nom": nom,
                    "Date": dt,
                    "Mois": f"{int(mois):02d}" if isinstance(mois, (int,str)) else _safe_str(mois),
                    "Categorie": sel_cat,
                    "Sous-categorie": sel_sub,
                    "Visa": visa_final if visa_final else sel_sub,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Total (US $)": total,
                    "Payé": paye,
                    "Reste": reste,
                    "Paiements": [],
                    "Options": opts_dict,
                    "Dossier envoyé": bool(sent),
                    "Date d'envoi": _date_or_none(sent_d) if sent else None,
                    "Dossier accepté": bool(acc),
                    "Date d'acceptation": _date_or_none(acc_d) if acc else None,
                    "Dossier refusé": bool(ref),
                    "Date de refus": _date_or_none(ref_d) if ref else None,
                    "Dossier annulé": bool(ann),
                    "Date d'annulation": _date_or_none(ann_d) if ann else None,
                    "RFE": bool(rfe),
                }
                df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                _write_clients(df_new, clients_path)
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()

    # ---------- MODIFICATION ----------
    elif op == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if df_live.empty:
            st.info("Aucun client à modifier.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            sel1, sel2 = st.columns(2)
            target_name = sel1.selectbox("Nom", [""]+names, index=0, key=f"mod_nom_{SID}")
            target_id   = sel2.selectbox("ID_Client", [""]+ids, index=0, key=f"mod_id_{SID}")

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is None or not mask.any():
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()

            c1, c2, c3 = st.columns(3)
            nom = c1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nomv_{SID}")
            dval = row.get("Date")
            dt  = c2.date_input("Date de création", value=_date_for_widget(dval), key=f"mod_date_{SID}")
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=max(0, int(_safe_str(row.get("Mois","01")))-1), key=f"mod_mois_{SID}")

            # options déjà enregistrées
            preset_opts = row.get("Options", {})
            if not isinstance(preset_opts, dict):
                try:
                    preset_opts = json.loads(_safe_str(preset_opts) or "{}")
                    if not isinstance(preset_opts, dict):
                        preset_opts = {}
                except Exception:
                    preset_opts = {}

            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""] + cats,
                                   index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                                   key=f"mod_cat_{SID}")

            sel_sub = _safe_str(row.get("Sous-categorie",""))
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""] + subs,
                                       index=(subs.index(sel_sub)+1 if sel_sub in subs else 0),
                                       key=f"mod_sub_{SID}")

            visa_final, opts_dict, info_msg = "", {"exclusive": None, "options": []}, ""
            if sel_cat and sel_sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=f"mod_opts_{SID}", preselected=preset_opts
                )
            if info_msg:
                st.info(info_msg)

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), "Montant honoraires (US $)").iloc[0]),
                                    step=50.0, format="%.2f", key=f"mod_h_{SID}")
            other = f2.number_input("Autres frais (US $)", min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), "Autres frais (US $)").iloc[0]),
                                    step=20.0, format="%.2f", key=f"mod_o_{SID}")

            st.markdown("#### 📌 Statuts (modifiables partout)")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent   = s1.checkbox("Dossier envoyé", value=bool(row.get("Dossier envoyé")), key=f"mod_sent_{SID}")
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
            acc    = s2.checkbox("Dossier accepté", value=bool(row.get("Dossier accepté")), key=f"mod_acc_{SID}")
            acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
            ref    = s3.checkbox("Dossier refusé", value=bool(row.get("Dossier refusé")), key=f"mod_ref_{SID}")
            ref_d  = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=f"mod_refd_{SID}")
            ann    = s4.checkbox("Dossier annulé", value=bool(row.get("Dossier annulé")), key=f"mod_ann_{SID}")
            ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
            rfe    = s5.checkbox("RFE", value=bool(row.get("RFE")), key=f"mod_rfe_{SID}")

            if rfe and not any([sent, acc, ref, ann]):
                st.warning("⚠️ RFE doit être couplé avec un autre statut (envoyé/accepté/refusé/annulé).")

            # Paiements (ajout)
            st.markdown("#### 💵 Paiements (ajouter un acompte)")
            pay_c1, pay_c2, pay_c3 = st.columns(3)
            pay_amt = pay_c1.number_input("Montant (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"pay_amt_{SID}")
            pay_dt  = pay_c2.date_input("Date du paiement", value=date.today(), key=f"pay_date_{SID}")
            pay_mode = pay_c3.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"pay_mode_{SID}")
            if st.button("➕ Ajouter le paiement", key=f"pay_add_{SID}"):
                # récupérer liste actuelle
                try:
                    curr = row.get("Paiements", [])
                    if not isinstance(curr, list):
                        curr = json.loads(_safe_str(curr) or "[]")
                        if not isinstance(curr, list):
                            curr = []
                except Exception:
                    curr = []
                curr.append({"amount": float(pay_amt), "date": str(_date_or_none(pay_dt)), "mode": pay_mode})
                df_live.at[idx, "Paiements"] = curr
                # recalcul payé/reste
                tmp = pd.DataFrame([df_live.loc[idx]])
                paid = float(_safe_num_series(tmp, "Payé").iloc[0])
                tot  = float(_safe_num_series(tmp, "Total (US $)").iloc[0])
                # Payé max entre total paiements et ancienne valeur
                def _sum_json(lst):
                    try:
                        return float(sum(float(it.get("amount",0.0) or 0.0) for it in (lst or [])))
                    except Exception:
                        return 0.0
                paid_json = _sum_json(curr)
                df_live.at[idx, "Payé"] = max(paid, paid_json)
                df_live.at[idx, "Reste"] = max(0.0, tot - df_live.at[idx, "Payé"])
                _write_clients(df_live, clients_path)
                st.success("Paiement ajouté.")
                st.cache_data.clear()
                st.rerun()

            # Historique paiements
            st.markdown("##### Historique des paiements")
            hist = row.get("Paiements", [])
            if not isinstance(hist, list):
                try:
                    hist = json.loads(_safe_str(hist) or "[]")
                    if not isinstance(hist, list):
                        hist = []
                except Exception:
                    hist = []
            if hist:
                st.table(pd.DataFrame(hist))
            else:
                st.caption("Aucun paiement enregistré.")

            save_mod = st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}")
            if save_mod:
                if not nom:
                    st.warning("Le nom est requis.")
                    st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisissez Catégorie et Sous-catégorie.")
                    st.stop()

                total = float(honor) + float(other)
                # recalcul payé/restes à partir de paiements déjà stockés
                try:
                    curr = df_live.at[idx, "Paiements"]
                    if not isinstance(curr, list):
                        curr = json.loads(_safe_str(curr) or "[]")
                        if not isinstance(curr, list):
                            curr = []
                except Exception:
                    curr = []
                def _sum_json(lst):
                    try:
                        return float(sum(float(it.get("amount",0.0) or 0.0) for it in (lst or [])))
                    except Exception:
                        return 0.0
                paid_json = _sum_json(curr)
                paye  = max(float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0]), paid_json)
                reste = max(0.0, total - paye)

                df_live.at[idx, "Nom"] = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(mois):02d}" if isinstance(mois,(int,str)) else _safe_str(mois)
                df_live.at[idx, "Categorie"] = sel_cat
                df_live.at[idx, "Sous-categorie"] = sel_sub
                df_live.at[idx, "Visa"] = (visa_final if visa_final else sel_sub)
                df_live.at[idx, "Montant honoraires (US $)"] = float(honor)
                df_live.at[idx, "Autres frais (US $)"] = float(other)
                df_live.at[idx, "Total (US $)"] = total
                df_live.at[idx, "Payé"] = paye
                df_live.at[idx, "Reste"] = reste
                df_live.at[idx, "Options"] = opts_dict
                df_live.at[idx, "Dossier envoyé"] = bool(sent)
                df_live.at[idx, "Date d'envoi"] = _date_or_none(sent_d) if sent else None
                df_live.at[idx, "Dossier accepté"] = bool(acc)
                df_live.at[idx, "Date d'acceptation"] = _date_or_none(acc_d) if acc else None
                df_live.at[idx, "Dossier refusé"] = bool(ref)
                df_live.at[idx, "Date de refus"] = _date_or_none(ref_d) if ref else None
                df_live.at[idx, "Dossier annulé"] = bool(ann)
                df_live.at[idx, "Date d'annulation"] = _date_or_none(ann_d) if ann else None
                df_live.at[idx, "RFE"] = bool(rfe)

                _write_clients(df_live, clients_path)
                st.success("Modifications enregistrées.")
                st.cache_data.clear()
                st.rerun()

    # ---------- SUPPRESSION ----------
    elif op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client à supprimer.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            sel1, sel2 = st.columns(2)
            target_name = sel1.selectbox("Nom", [""]+names, index=0, key=f"del_nom_{SID}")
            target_id   = sel2.selectbox("ID_Client", [""]+ids, index=0, key=f"del_id_{SID}")

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({"Dossier N": row.get("Dossier N",""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                    df_new = df_live[~mask].copy()
                    _write_clients(df_new, clients_path)
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()

# ==============================================
# 💾 Export global (Clients + Visa) — optionnel
# ==============================================
st.markdown("---")
st.markdown("### 💾 Export global (Clients + Visa)")
colz1, colz2 = st.columns([1,3])
with colz1:
    if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID}"):
        try:
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients propre
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
            st.session_state[f"zip_export_{SID}"] = buf.getvalue()
            st.success("Archive prête.")
        except Exception as e:
            st.error("Erreur de préparation : " + _safe_str(e))

with colz2:
    if st.session_state.get(f"zip_export_{SID}"):
        st.download_button(
            label="⬇️ Télécharger l’export (ZIP)",
            data=st.session_state[f"zip_export_{SID}"],
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
            key=f"zip_dl_{SID}",
        )