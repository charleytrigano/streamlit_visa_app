from __future__ import annotations

# =========================
# VISA MANAGER — APP (Part 1/3)
# =========================

import json
import zipfile
from io import BytesIO
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import streamlit as st
import openpyxl  # writer Excel

# --------- Config ----------
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")
st.caption("Catégorie → Sous-catégorie → cases (Visa), Clients, Paiements, ESCROW, Analyses — Excel source")

# --------- Constantes colonnes ----------
DEFAULT_CLIENTS = "donnees_visa_clients1_adapte.xlsx"
DEFAULT_VISA    = "donnees_visa_clients1.xlsx"

DOSSIER_COL = "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"
PAY_JSON = "Paiements"  # JSON list: [{date, mode, amount}]

CLIENTS_COLUMNS = [
    DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
    "Categorie", "Sous-categorie", "Visa",
    HONO, AUTRE, TOTAL, "Payé", "Reste",
    PAY_JSON, "ESCROW transféré (US $)", "Journal ESCROW",
    "Dossier envoyé","Date envoyé",
    "Dossier approuvé","Date approuvé",
    "RFE","Date RFE",
    "Dossier refusé","Date refusé",
    "Dossier annulé","Date annulé",
]

STATE_LAST = "last_excel_paths"  # (clients_path, visa_path)

# --------- Petits helpers ----------
def _safe_str(x) -> str:
    try:
        if pd.isna(x): return ""
    except Exception:
        pass
    return str(x)

def _to_num(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.replace(r"[^\d,.\-]", "", regex=True).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df), index=df.index, dtype=float)
    v = df[col]
    if pd.api.types.is_numeric_dtype(v):
        return v.fillna(0.0).astype(float)
    return _to_num(v)

def _fmt_money_us(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _uniquify_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(map(str, df.columns))
    seen = {}
    new_cols = []
    for c in cols:
        if c not in seen:
            seen[c] = 1; new_cols.append(c)
        else:
            seen[c] += 1; new_cols.append(f"{c}_{seen[c]}")
    out = df.copy()
    out.columns = new_cols
    return out

# ------- Mémoire chemins -------
def save_last_paths(clients_path: str, visa_path: str) -> None:
    st.session_state[STATE_LAST] = (clients_path, visa_path)

def load_last_paths(default_clients: str, default_visa: str) -> tuple[str, str]:
    return st.session_state.get(STATE_LAST, (default_clients, default_visa))

# --------- Création fichiers modèles ----------
def _create_clients_template(path: str|Path, sheet_name: str="Clients") -> None:
    df = pd.DataFrame(columns=CLIENTS_COLUMNS)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        df.to_excel(wr, sheet_name=sheet_name, index=False)

def _create_visa_template(path: str|Path, sheet_name: str="Visa") -> None:
    # Modèle minimal : Categorie, Sous-categorie 1, cases (COS/EOS)
    rows = [
        {"Categorie":"B-1","Sous-categorie 1":"Affaires","COS":"x","EOS":""},
        {"Categorie":"B-2","Sous-categorie 1":"Tourisme","COS":"x","EOS":"x"},
        {"Categorie":"F-1","Sous-categorie 1":"Etudiant","COS":"x","EOS":"x"},
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        df.to_excel(wr, sheet_name=sheet_name, index=False)

def ensure_files_exist(clients_path: str|Path, visa_path: str|Path) -> None:
    cp = Path(clients_path); vp = Path(visa_path)
    if not cp.exists(): _create_clients_template(cp)
    if not vp.exists(): _create_visa_template(vp)

def safe_excel_first_sheet(path: str|Path, preferred: str|None=None) -> str:
    with pd.ExcelFile(path) as xls:
        sheets = xls.sheet_names
    if preferred and preferred in sheets: return preferred
    return sheets[0] if sheets else "Clients"

# --------- Export bytes ----------
def excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        for name, df in sheets.items():
            df2 = _uniquify_columns(df.copy())
            df2.to_excel(wr, sheet_name=name, index=False)
    bio.seek(0)
    return bio.getvalue()

def make_zip_clients_visa(clients_df: pd.DataFrame, visa_df: pd.DataFrame) -> bytes:
    clients_bytes = excel_bytes({"Clients": clients_df})
    visa_bytes    = excel_bytes({"Visa": visa_df})
    bio = BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Clients.xlsx", clients_bytes)
        zf.writestr("Visa.xlsx",    visa_bytes)
    bio.seek(0)
    return bio.getvalue()

# --------- Normalisation en-têtes Visa ----------
import unicodedata
def _norm_header(s: str) -> str:
    """Normalise un en-tête: supprime accents, minuscule, compacte espaces/tirets/underscores."""
    if s is None: return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.strip().lower()
    for ch in ["\u00A0", "\t"]:
        s = s.replace(ch, " ")
    s = s.replace("-", " ").replace("_", " ")
    s = " ".join(s.split())
    return s

def detect_visa_sheet_name(xlsx_path: str | Path) -> str:
    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        return "Visa"
    for sn in xls.sheet_names:
        try:
            tmp = pd.read_excel(xlsx_path, sheet_name=sn, nrows=5)
            tmp = _uniquify_columns(tmp)
            tmp.columns = tmp.columns.map(str).str.strip()
            # on accepte Catégorie ou Categorie
            norm = { _norm_header(c): c for c in tmp.columns }
            if any(k in norm for k in ("categorie","category")):
                return sn
        except Exception:
            continue
    return xls.sheet_names[0] if xls.sheet_names else "Visa"

# --------- Parsing Visa imbriqué (très tolérant) ---------
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    """
    Sortie:
    { "Categorie": { "Sous-categorie X": ["X COS", "X EOS", ...] (ou ["X"] si aucune case cochée) }, ... }
    - Tolère accents / pluriels / tirets / underscores / espaces.
    - Considère 'coché' toute valeur non vide sauf {0,false,non,no,off,null,none}.
    - Ignore les colonnes de cases totalement vides.
    """
    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)): return False
        if isinstance(v, (int, float)): return float(v) != 0.0
        s = str(v).strip().lower()
        if s == "": return False
        return s not in {"0","false","faux","non","no","off","null","none"}

    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        return {}

    sheets_to_try = [sheet_name] if sheet_name else xls.sheet_names
    for sn in sheets_to_try:
        if sn is None: continue
        try:
            dfv = pd.read_excel(xlsx_path, sheet_name=sn)
        except Exception:
            continue
        if dfv.empty:
            continue

        dfv = _uniquify_columns(dfv)
        dfv.columns = dfv.columns.map(str).str.strip()

        colmap = { _norm_header(c): c for c in dfv.columns }
        cat_col = None
        for key in ("categorie", "category"):
            if key in colmap:
                cat_col = colmap[key]; break
        if not cat_col:
            # variantes sans espace
            for nk, orig in colmap.items():
                if nk.replace(" ", "") in {"categorie","category"}:
                    cat_col = orig; break
        if not cat_col:
            continue

        sub_col = None
        for key in ("sous categorie 1","sous categorie","sous categories 1","sous categories","sous cat 1","sous cat"):
            if key in colmap:
                sub_col = colmap[key]; break
        if not sub_col:
            for nk, orig in colmap.items():
                if nk.replace(" ", "") in {"souscategorie1","souscategorie"}:
                    sub_col = orig; break
        if not sub_col:
            dfv["_Sous_"] = ""
            sub_col = "_Sous_"

        checkbox_cols = [c for c in dfv.columns if c not in {cat_col, sub_col}]

        # garde uniquement les colonnes 'cases' qui ont au moins une valeur cochée
        usable_cols = []
        for c in checkbox_cols:
            col = dfv[c]
            if any(_is_checked(v) for v in col):
                usable_cols.append(c)

        out: dict[str, dict[str, list[str]]] = {}
        for _, row in dfv.iterrows():
            cat  = _safe_str(row.get(cat_col, "")).strip()
            sous = _safe_str(row.get(sub_col, "")).strip()
            if not cat:
                continue

            labels_for_sub: list[str] = []
            for cc in usable_cols:
                if _is_checked(row.get(cc, None)):
                    labels_for_sub.append(f"{sous} {cc}".strip())

            if not labels_for_sub and sous:
                labels_for_sub = [sous]

            if labels_for_sub:
                out.setdefault(cat, {})
                out[cat].setdefault(sous, [])
                out[cat][sous].extend(labels_for_sub)

        if out:
            for cat, submap in out.items():
                for sous, arr in submap.items():
                    submap[sous] = sorted(set(arr))
            return out

    return {}

def _visa_all_categories(visa_map_nested: dict[str, dict[str, list[str]]]) -> list[str]:
    return sorted(list(visa_map_nested.keys()))

def _visa_subcats_for(cat: str, visa_map_nested: dict[str, dict[str, list[str]]]) -> list[str]:
    return sorted(list(visa_map_nested.get(cat, {}).keys()))

def _visa_options_for(cat: str, subcat: str, visa_map_nested: dict[str, dict[str, list[str]]]) -> list[str]:
    return sorted(list(visa_map_nested.get(cat, {}).get(subcat, [])))

# --------- Clients I/O + normalisation ----------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in CLIENTS_COLUMNS:
        if c not in df.columns: df[c] = None

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df["Mois"] = df.apply(
        lambda r: f"{pd.to_datetime(r['Date']).month:02d}" if pd.notna(r["Date"]) else (_safe_str(r.get("Mois",""))[:2] or None),
        axis=1
    )

    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste", "ESCROW transféré (US $)"]:
        df[c] = _safe_num_series(df, c)
    df[TOTAL] = df[HONO] + df[AUTRE]

    def _norm_p(x):
        try:
            j = json.loads(_safe_str(x) or "[]")
            return j if isinstance(j, list) else []
        except Exception:
            return []
    df[PAY_JSON] = df[PAY_JSON].apply(_norm_p)

    def _sum_json(lst):
        try:
            return float(sum(float(it.get("amount",0.0) or 0.0) for it in (lst or [])))
        except Exception:
            return 0.0
    json_paid = df[PAY_JSON].apply(_sum_json)
    df["Payé"] = pd.concat([df["Payé"].fillna(0.0).astype(float), json_paid], axis=1).max(axis=1)
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    df["_Année_"]   = df["Date"].apply(lambda d: d.year if pd.notna(d) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda d: d.month if pd.notna(d) else pd.NA)
    df["_Mois_"]    = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) else pd.NA)

    for c in ["Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"]:
        df[c] = df[c].apply(lambda v: bool(v) if pd.notna(v) else False)

    def _norm_j(x):
        try:
            j = json.loads(_safe_str(x) or "[]")
            return j if isinstance(j, list) else []
        except Exception:
            return []
    df["Journal ESCROW"] = df["Journal ESCROW"].apply(_norm_j)

    return _uniquify_columns(df)

@st.cache_data(show_spinner=False)
def read_sheet(xlsx_path: str | Path, sheet_name: str) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    return normalize_clients(_uniquify_columns(df))

def write_sheet_inplace(xlsx_path: str | Path, sheet_name: str, df: pd.DataFrame) -> None:
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="w") as wr:
            df.to_excel(wr, sheet_name=sheet_name, index=False)
        return
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as wr:
        try:
            book = wr.book
            if sheet_name in book.sheetnames:
                sh = book[sheet_name]; book.remove(sh); book.create_sheet(sheet_name)
        except Exception:
            pass
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
        df.to_excel(wr, sheet_name=sheet_name, index=False)

# --------- ID & Dossier ----------
def _make_client_id_from_row(r: dict) -> str:
    nom = _safe_str(r.get("Nom","")).strip().replace(" ", "_")
    dt = r.get("Date", date.today())
    if isinstance(dt, str):
        try: dt = pd.to_datetime(dt).date()
        except Exception: dt = date.today()
    return f"{nom}-{dt:%Y%m%d}"

def next_dossier_number(df: pd.DataFrame, start: int = 13057) -> int:
    if DOSSIER_COL in df.columns and pd.to_numeric(df[DOSSIER_COL], errors="coerce").notna().any():
        return int(pd.to_numeric(df[DOSSIER_COL], errors="coerce").max()) + 1
    return int(start)

# --------- Barre latérale : chemins / import / export ----------
st.sidebar.header("📂 Fichiers (mémoire)")
clients_default, visa_default = load_last_paths(DEFAULT_CLIENTS, DEFAULT_VISA)

clients_path = st.sidebar.text_input("Fichier Clients (.xlsx)", value=clients_default, key="cli_path")
visa_path    = st.sidebar.text_input("Fichier Visa (.xlsx)",    value=visa_default,   key="visa_path")

ensure_files_exist(clients_path, visa_path)
save_last_paths(clients_path, visa_path)

# Chargements
try:
    sheet_choice = safe_excel_first_sheet(clients_path, preferred="Clients")
except Exception:
    _create_clients_template(clients_path); sheet_choice = "Clients"

try:
    df_clients = read_sheet(clients_path, sheet_name=sheet_choice)
except Exception:
    _create_clients_template(clients_path); df_clients = read_sheet(clients_path, "Clients")

try:
    visa_map = parse_visa_sheet(visa_path, sheet_name=None)
except Exception:
    _create_visa_template(visa_path); visa_map = parse_visa_sheet(visa_path, sheet_name=None)

st.sidebar.markdown("---")
st.sidebar.subheader("⬆️ Import / ⬇️ Export")

upl_cli = st.sidebar.file_uploader("Remplacer le fichier Clients", type=["xlsx"], key="upl_clients")
if upl_cli is not None:
    Path(clients_path).write_bytes(upl_cli.read())
    st.sidebar.success("Clients remplacé.")
    df_clients = read_sheet(clients_path, sheet_choice)

upl_visa = st.sidebar.file_uploader("Remplacer le fichier Visa", type=["xlsx"], key="upl_visa")
if upl_visa is not None:
    Path(visa_path).write_bytes(upl_visa.read())
    st.sidebar.success("Visa remplacé.")
    visa_map = parse_visa_sheet(visa_path, sheet_name=None)

if st.sidebar.button("🔄 Recharger Visa"):
    visa_map = parse_visa_sheet(visa_path, sheet_name=None)
    st.sidebar.success("Référentiel Visa rechargé.")

# Export ZIP (Clients.xlsx + Visa.xlsx)
try:
    visa_sheet = detect_visa_sheet_name(visa_path)
    df_visa_export = pd.read_excel(visa_path, sheet_name=visa_sheet)
    zip_bytes = make_zip_clients_visa(df_clients.copy(), df_visa_export)
    st.sidebar.download_button(
        label="📦 Télécharger ZIP (Clients + Visa)",
        data=zip_bytes,
        file_name="Visa_Manager_Export.zip",
        mime="application/zip",
        key="dl_zip_clients_visa",
    )
except Exception as e:
    st.sidebar.caption(f"Export ZIP impossible : {e}")

# Aperçu et debug Visa
with st.sidebar.expander("🔎 Aperçu Visa détecté", expanded=False):
    if visa_map:
        for cat, submap in visa_map.items():
            lines = []
            for sous, opts in submap.items():
                if opts:
                    lines.append(f"- {sous}: {', '.join(opts)}")
                else:
                    lines.append(f"- {sous}")
            st.markdown(f"**{cat}**  \n" + "\n".join(lines))
    else:
        st.caption("Aucune donnée Visa.")

with st.sidebar.expander("🧰 Debug Visa (lecture brute + colonnes cases)", expanded=False):
    try:
        vsn = detect_visa_sheet_name(visa_path)
        raw = pd.read_excel(visa_path, sheet_name=vsn)
        st.write("Feuille lue:", vsn)
        st.write("Colonnes:", list(raw.columns))
        if not raw.empty:
            cols = list(_uniquify_columns(raw).columns)
            cols = [c.strip() for c in cols]
            colmap = { _norm_header(c): c for c in cols }
            cat_col = next((colmap[k] for k in ("categorie","category") if k in colmap), None)
            sub_col = None
            for k in ("sous categorie 1","sous categorie","sous categories 1","sous categories","sous cat 1","sous cat"):
                if k in colmap: sub_col = colmap[k]; break
            if not sub_col:
                for nk, orig in colmap.items():
                    if nk.replace(" ","") in {"souscategorie1","souscategorie"}:
                        sub_col = orig; break
            if not sub_col:
                raw["_Sous_"] = ""; sub_col = "_Sous_"
            cand = [c for c in cols if c not in {cat_col, sub_col}]
            def _is_checked(v):
                if v is None or (isinstance(v, float) and pd.isna(v)): return False
                if isinstance(v, (int,float)): return float(v) != 0.0
                s = str(v).strip().lower()
                if s == "": return False
                return s not in {"0","false","faux","non","no","off","null","none"}
            usable = [c for c in cand if any(_is_checked(v) for v in raw[c])]
            st.write("Colonnes 'cases' retenues:", usable)
            st.dataframe(raw.head(5), use_container_width=True)
    except Exception as e:
        st.caption(f"Debug indisponible: {e}")


# ================== TABS ==================
tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs(
    ["Dashboard", "Clients", "Analyses", "ESCROW"]
)

# ================ DASHBOARD ================
with tab_dash:
    st.subheader("📊 Tableau de bord")

    cats_all  = sorted(df_clients["Categorie"].dropna().astype(str).unique().tolist())
    visas_all = sorted(df_clients["Visa"].dropna().astype(str).unique().tolist())

    c1,c2,c3,c4 = st.columns([1,1,1,2])
    with c1:
        sel_cats = st.multiselect("Catégories", cats_all, default=[], key="dash_cats")
    with c2:
        sel_visas = st.multiselect("Visa", visas_all, default=[], key="dash_visas")
    with c3:
        sel_solde = st.selectbox("Solde", ["Tous","Soldé (Reste = 0)","Non soldé (Reste > 0)"], index=0, key="dash_solde")
    with c4:
        q = st.text_input("Recherche (Nom / ID / Dossier / Visa)", "", key="dash_q")

    ff = df_clients.copy()
    if sel_cats:
        ff = ff[ff["Categorie"].astype(str).isin(sel_cats)]
    if sel_visas:
        ff = ff[ff["Visa"].astype(str).isin(sel_visas)]
    if sel_solde == "Soldé (Reste = 0)":
        ff = ff[_safe_num_series(ff,"Reste") <= 1e-9]
    elif sel_solde == "Non soldé (Reste > 0)":
        ff = ff[_safe_num_series(ff,"Reste") > 1e-9]
    if q:
        qn = q.lower().strip()
        def _m(r):
            hay = " ".join([
                _safe_str(r.get("Nom","")),
                _safe_str(r.get("ID_Client","")),
                _safe_str(r.get("Categorie","")),
                _safe_str(r.get("Visa","")),
                str(r.get(DOSSIER_COL,"")),
            ]).lower()
            return qn in hay
        ff = ff[ff.apply(_m, axis=1)]

    st.info(f"**{len(ff)} dossiers** avec les filtres.")
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(ff)}")
    k2.metric("Honoraires", _fmt_money_us(_safe_num_series(ff,HONO).sum()))
    k3.metric("Payé", _fmt_money_us(_safe_num_series(ff,"Payé").sum()))
    k4.metric("Reste", _fmt_money_us(_safe_num_series(ff,"Reste").sum()))

    st.markdown("---")
    view = ff.copy()
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c in view.columns:
            view[c] = _safe_num_series(view,c).map(_fmt_money_us)
    if "Date" in view.columns:
        view["Date"] = view["Date"].astype(str)

    show_cols = [c for c in [
        DOSSIER_COL,"ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"
    ] if c in view.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
    view_sorted = view.sort_values(by=sort_cols) if sort_cols else view

    st.dataframe(_uniquify_columns(view_sorted[show_cols].reset_index(drop=True)), use_container_width=True)

# ================ CLIENTS (CRUD + Paiements + Cases dépendantes) ================
with tab_clients:
    st.subheader("👥 Clients — Créer / Modifier / Supprimer / Paiements")

    live = df_clients.copy()

    cL, cR = st.columns([1,1])

    # --- Sélection client
    with cL:
        st.markdown("### 🔎 Sélection")
        if live.empty:
            st.caption("Aucun client.")
            sel_idx = None
            sel_row = None
        else:
            labels = (live.get("Nom","").astype(str) + " — " + live.get("ID_Client","").astype(str)).fillna("")
            sel_idx = st.selectbox("Client", options=list(live.index),
                                   format_func=lambda i: labels.iloc[i],
                                   key=f"cli_sel")
            sel_row = live.loc[sel_idx] if sel_idx is not None else None

    # --- Création client
    with cR:
        st.markdown("### ➕ Nouveau client")
        new_name = st.text_input("Nom", key="new_nom")
        new_date = st.date_input("Date de création", value=date.today(), key="new_date")

        # Catégorie
        cats = _visa_all_categories(visa_map)
        new_cat = st.selectbox("Categorie", options=[""]+cats, index=0, key="new_cat_sel")

        # Sous-catégorie dépendante
        sub_opts = _visa_subcats_for(new_cat, visa_map) if new_cat else []
        new_sub  = st.selectbox("Sous-categorie", options=[""]+sub_opts, index=0, key="new_sub_sel")

        # Cases (par sous-catégorie) -> choix unique via cases à cocher contrôlées
        opt_list = _visa_options_for(new_cat, new_sub, visa_map) if (new_cat and new_sub) else []
        st.caption("Choix des cases (une seule). Si aucune case disponible, le Visa = Sous-catégorie seule.")
        checked_keys = []
        chosen_case = ""
        for i, lab in enumerate(opt_list):
            # lab est déjà "Sous-categorie <Case>", on extrait juste la partie 'Case' pour l’affichage
            suf = lab[len(new_sub):].strip() if new_sub and lab.startswith(new_sub) else lab
            k = f"new_case_{i}"
            val = st.checkbox(f"{suf or new_sub}", key=k, value=False, disabled=False)
            if val:
                checked_keys.append(k); chosen_case = lab
        # Enforce un seul choix
        if len(checked_keys) > 1:
            st.warning("Une seule case doit être cochée. Décochez les autres.")

        # Valeur Visa finale
        if opt_list:
            new_visa = chosen_case if chosen_case else ""
        else:
            new_visa = new_sub if new_sub else ""

        new_hono = st.number_input(HONO, min_value=0.0, step=10.0, format="%.2f", key="new_hono")
        new_autr = st.number_input(AUTRE, min_value=0.0, step=10.0, format="%.2f", key="new_autre")

        if st.button("💾 Créer", key="btn_create"):
            if not new_name: st.warning("Nom obligatoire."); st.stop()
            if not new_cat: st.warning("Categorie obligatoire."); st.stop()
            if not new_sub: st.warning("Sous-categorie obligatoire."); st.stop()
            # si options existent, on veut une et une seule case cochée
            if opt_list and not chosen_case:
                st.warning("Cochez une case pour définir le Visa."); st.stop()

            base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
            base_norm = normalize_clients(base_raw.copy())

            dossier = next_dossier_number(base_norm)
            client_id = _make_client_id_from_row({"Nom": new_name, "Date": new_date})
            origin = client_id; i = 0
            while (base_norm["ID_Client"].astype(str) == client_id).any():
                i += 1; client_id = f"{origin}-{i}"

            total = float(new_hono) + float(new_autr)
            new_row = {
                DOSSIER_COL: dossier,
                "ID_Client": client_id,
                "Nom": new_name,
                "Date": pd.to_datetime(new_date).date(),
                "Mois": f"{new_date.month:02d}",
                "Categorie": new_cat,
                "Sous-categorie": new_sub,
                "Visa": new_visa,
                HONO: float(new_hono),
                AUTRE: float(new_autr),
                TOTAL: total,
                "Payé": 0.0,
                "Reste": total,
                PAY_JSON: "[]",
                "ESCROW transféré (US $)": 0.0,
                "Journal ESCROW": "[]",
                "Dossier envoyé": False, "Date envoyé": "",
                "Dossier approuvé": False, "Date approuvé": "",
                "RFE": False, "Date RFE": "",
                "Dossier refusé": False, "Date refusé": "",
                "Dossier annulé": False, "Date annulé": "",
            }
            base_raw = pd.concat([base_raw, pd.DataFrame([new_row])], ignore_index=True)
            base_norm = normalize_clients(base_raw)
            write_sheet_inplace(clients_path, sheet_choice, base_norm)
            st.success("✅ Client créé.")
            st.rerun()

    st.markdown("---")

    if sel_row is None:
        st.stop()

    # --- Edition & Paiements
    idx = sel_idx
    ed = sel_row.to_dict()

    e1,e2,e3 = st.columns(3)
    with e1:
        ed_nom  = st.text_input("Nom", value=_safe_str(ed.get("Nom","")), key=f"ed_nom_{idx}")
        ed_date = st.date_input("Date de création",
                                value=(pd.to_datetime(ed.get("Date")).date() if pd.notna(ed.get("Date")) else date.today()),
                                key=f"ed_date_{idx}")

    # Bloc Catégorie → Sous-catégorie → cases
    with e2:
        cats = _visa_all_categories(visa_map)
        curr_cat = _safe_str(ed.get("Categorie",""))
        ed_cat = st.selectbox("Categorie",
                              options=[""]+cats,
                              index=(cats.index(curr_cat)+1 if curr_cat in cats else 0),
                              key=f"ed_cat_{idx}")

        sub_opts = _visa_subcats_for(ed_cat, visa_map) if ed_cat else []
        curr_sub = _safe_str(ed.get("Sous-categorie",""))
        ed_sub = st.selectbox("Sous-categorie",
                              options=[""]+sub_opts,
                              index=(sub_opts.index(curr_sub)+1 if curr_sub in sub_opts else 0),
                              key=f"ed_sub_{idx}")

    with e3:
        # options de cases pour la sous-catégorie
        opt_list = _visa_options_for(ed_cat, ed_sub, visa_map) if (ed_cat and ed_sub) else []
        curr_visa = _safe_str(ed.get("Visa",""))

        st.caption("Choix des cases (une seule). Si aucune case disponible, le Visa = Sous-catégorie seule.")
        chosen_case = ""
        checked_count = 0
        # initialise: coche la case correspondant au Visa actuel si possible
        for i, lab in enumerate(opt_list):
            suf = lab[len(ed_sub):].strip() if ed_sub and lab.startswith(ed_sub) else lab
            default_checked = (curr_visa == lab)
            val = st.checkbox(f"{suf or ed_sub}", key=f"ed_case_{idx}_{i}", value=default_checked)
            if val:
                checked_count += 1
                chosen_case = lab

        if opt_list and checked_count == 0:
            st.info("Cochez la case du Visa souhaité.")
        if checked_count > 1:
            st.warning("Une seule case doit être cochée.")

        ed_visa_final = chosen_case if (opt_list and chosen_case) else (ed_sub if ed_sub else "")

        ed_hono = st.number_input(HONO, min_value=0.0, value=float(ed.get(HONO,0.0)), step=10.0, format="%.2f", key=f"ed_hono_{idx}")
        ed_autr = st.number_input(AUTRE, min_value=0.0, value=float(ed.get(AUTRE,0.0)), step=10.0, format="%.2f", key=f"ed_autre_{idx}")

    st.markdown("#### 🧾 Statuts du dossier")
    s1,s2,s3 = st.columns(3)
    with s1:
        ed_env = st.checkbox("Dossier envoyé", value=bool(ed.get("Dossier envoyé",False)), key=f"ed_env_{idx}")
        ed_app = st.checkbox("Dossier approuvé", value=bool(ed.get("Dossier approuvé",False)), key=f"ed_app_{idx}")
    with s2:
        ed_rfe = st.checkbox("RFE", value=bool(ed.get("RFE",False)), key=f"ed_rfe_{idx}")
        ed_ref = st.checkbox("Dossier refusé", value=bool(ed.get("Dossier refusé",False)), key=f"ed_ref_{idx}")
    with s3:
        ed_ann = st.checkbox("Dossier annulé", value=bool(ed.get("Dossier annulé",False)), key=f"ed_ann_{idx}")

    st.caption("💡 RFE n’a de sens que si le dossier a un statut (Envoyé / Approuvé / Refusé / Annulé).")

    st.markdown("### 💳 Paiements (acomptes)")
    p1,p2,p3,p4 = st.columns([1,1,1,2])
    with p1:
        p_date = st.date_input("Date paiement", value=date.today(), key=f"p_date_{idx}")
    with p2:
        p_mode = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"p_mode_{idx}")
    with p3:
        p_amt = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"p_amt_{idx}")
    with p4:
        if st.button("➕ Ajouter paiement", key=f"p_add_{idx}"):
            base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
            idc = _safe_str(ed.get("ID_Client",""))
            if idc and "ID_Client" in base_raw.columns:
                idxs = base_raw.index[base_raw["ID_Client"].astype(str)==idc].tolist()
                real_idx = idxs[0] if idxs else idx
            else:
                real_idx = idx
            if float(p_amt) <= 0:
                st.warning("Montant > 0 requis.")
            else:
                row = base_raw.loc[real_idx].to_dict()
                try:
                    plist = json.loads(_safe_str(row.get(PAY_JSON,"[]")) or "[]")
                    if not isinstance(plist,list): plist=[]
                except Exception:
                    plist=[]
                plist.append({"date": str(p_date), "mode": p_mode, "amount": float(p_amt)})
                row[PAY_JSON] = json.dumps(plist, ensure_ascii=False)
                base_raw.loc[real_idx] = row
                base_norm = normalize_clients(base_raw.copy())
                write_sheet_inplace(clients_path, sheet_choice, base_norm)
                st.success("Paiement ajouté.")
                st.rerun()

    # Historique paiements
    try:
        hist = json.loads(_safe_str(sel_row.get(PAY_JSON,"[]")) or "[]")
        if not isinstance(hist,list): hist=[]
    except Exception:
        hist=[]
    st.write("**Historique paiements**")
    if hist:
        h = pd.DataFrame(hist)
        if "amount" in h.columns: h["amount"] = h["amount"].astype(float).map(_fmt_money_us)
        st.dataframe(h, use_container_width=True)
    else:
        st.caption("Aucun paiement saisi.")

    st.markdown("---")
    a1,a2 = st.columns([1,1])
    if a1.button("💾 Sauvegarder les modifications", key=f"save_{idx}"):
        base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
        idc = _safe_str(ed.get("ID_Client",""))
        if idc and "ID_Client" in base_raw.columns:
            idxs = base_raw.index[base_raw["ID_Client"].astype(str)==idc].tolist()
            real_idx = idxs[0] if idxs else idx
        else:
            real_idx = idx
        if real_idx is None or not (0 <= real_idx < len(base_raw)):
            st.error("Ligne introuvable."); st.stop()

        row = base_raw.loc[real_idx].to_dict()
        row["Nom"] = ed_nom
        row["Date"] = pd.to_datetime(ed_date).date()
        row["Mois"] = f"{ed_date.month:02d}"
        row["Categorie"] = ed_cat
        row["Sous-categorie"] = ed_sub
        row["Visa"] = ed_visa_final
        row[HONO] = float(ed_hono)
        row[AUTRE] = float(ed_autr)
        row[TOTAL] = float(ed_hono) + float(ed_autr)
        row["Dossier envoyé"] = bool(ed_env)
        row["Dossier approuvé"] = bool(ed_app)
        row["RFE"] = bool(ed_rfe)
        row["Dossier refusé"] = bool(ed_ref)
        row["Dossier annulé"] = bool(ed_ann)

        base_raw.loc[real_idx] = row
        base_norm = normalize_clients(base_raw.copy())
        write_sheet_inplace(clients_path, sheet_choice, base_norm)
        st.success("✅ Modifications sauvegardées.")
        st.rerun()

    if a2.button("🗑️ Supprimer ce client", key=f"del_{idx}"):
        base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
        idc = _safe_str(ed.get("ID_Client",""))
        if idc and "ID_Client" in base_raw.columns:
            base_raw = base_raw.loc[base_raw["ID_Client"].astype(str)!=idc].reset_index(drop=True)
        else:
            base_raw = base_raw.drop(index=idx).reset_index(drop=True)
        base_norm = normalize_clients(base_raw.copy())
        write_sheet_inplace(clients_path, sheet_choice, base_norm)
        st.success("🗑️ Client supprimé.")
        st.rerun()


# ================== ANALYSES ==================
with tab_analyses:
    st.subheader("📈 Analyses")

    base = df_clients.copy()
    # Filtres simples (année / mois / cat / visa / solde)
    yearsA  = sorted([int(y) for y in pd.to_numeric(base["_Année_"], errors="coerce").dropna().unique().tolist()])
    monthsA = [f"{m:02d}" for m in range(1,13)]
    catsA   = sorted(base["Categorie"].dropna().astype(str).unique().tolist())
    visasA  = sorted(base["Visa"].dropna().astype(str).unique().tolist())

    cR1, cR2, cR3, cR4, cR5 = st.columns([1,1,1,1,1])
    sel_years  = cR1.multiselect("Année", yearsA, default=[], key="an_years")
    sel_months = cR2.multiselect("Mois (MM)", monthsA, default=[], key="an_months")
    sel_cats   = cR3.multiselect("Catégories", catsA, default=[], key="an_cats")
    sel_visas  = cR4.multiselect("Visa", visasA, default=[], key="an_visas")
    sel_solde  = cR5.selectbox("Solde", ["Tous","Soldé","Non soldé"], index=0, key="an_solde")

    f = base.copy()
    if sel_years:
        f = f[f["_Année_"].isin(sel_years)]
    if sel_months:
        f = f[f["_Mois_"].isin(sel_months)]
    if sel_cats:
        f = f[f["Categorie"].astype(str).isin(sel_cats)]
    if sel_visas:
        f = f[f["Visa"].astype(str).isin(sel_visas)]
    if sel_solde == "Soldé":
        f = f[_safe_num_series(f,"Reste") <= 1e-9]
    elif sel_solde == "Non soldé":
        f = f[_safe_num_series(f,"Reste") > 1e-9]

    # KPIs
    st.markdown("### KPI")
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Honoraires", _fmt_money_us(_safe_num_series(f,HONO).sum()))
    k3.metric("Encaissements (Payé)", _fmt_money_us(_safe_num_series(f,"Payé").sum()))
    k4.metric("Solde à encaisser", _fmt_money_us(_safe_num_series(f,"Reste").sum()))

    st.markdown("---")
    # Volumes par mois (année/mois)
    vol = f.copy()
    if "Date" in vol.columns:
        vol = vol[pd.notna(vol["Date"])]
    vol_group = vol.groupby(["_Année_","_Mois_"], dropna=True).size().reset_index(name="Dossiers")
    if not vol_group.empty:
        st.markdown("#### Volumes par Année/Mois")
        st.dataframe(vol_group.sort_values(["_Année_","_Mois_"]).reset_index(drop=True), use_container_width=True)
    else:
        st.caption("Aucun volume à afficher (filtre trop restrictif ?)")

    # CA par mois
    ca = f.groupby(["_Année_","_Mois_"], dropna=True)[HONO].sum().reset_index(name="Honoraires")
    if not ca.empty:
        st.markdown("#### Honoraires par Année/Mois")
        ca["Honoraires"] = ca["Honoraires"].map(_fmt_money_us)
        st.dataframe(ca.sort_values(["_Année_","_Mois_"]).reset_index(drop=True), use_container_width=True)

    # Détails (clients listés)
    st.markdown("### 📋 Détails (clients filtrés)")
    detail = f.copy()
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c in detail.columns:
            detail[c] = _safe_num_series(detail, c).map(_fmt_money_us)
    if "Date" in detail.columns:
        detail["Date"] = detail["Date"].astype(str)

    show_cols = [c for c in [
        DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa", "Date", "Mois",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"
    ] if c in detail.columns]
    sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in detail.columns]
    detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail
    st.dataframe(_uniquify_columns(detail_sorted[show_cols].reset_index(drop=True)), use_container_width=True)

# ================== ESCROW ==================
with tab_escrow:
    st.subheader("🏦 ESCROW — suivi des honoraires encaissés")

    g = df_clients.copy()
    g["Payé_num"] = _safe_num_series(g, "Payé")
    g["Hon_num"]  = _safe_num_series(g, HONO)
    g["Esc_trans"] = _safe_num_series(g, "ESCROW transféré (US $)")

    # Solde ESCROW dispo = min(Payé, Honoraires) - déjà transféré (pas négatif)
    g["ESCROW dispo"] = (g[["Payé_num","Hon_num"]].min(axis=1) - g["Esc_trans"]).clip(lower=0.0)

    # Filtres rapides
    cc1, cc2 = st.columns([1,1])
    with cc1:
        only_dispo = st.checkbox("Afficher uniquement dossiers avec ESCROW disponible", value=True, key="esc_only_pos")
    with cc2:
        q = st.text_input("Recherche (Nom / ID / Dossier / Visa)", "", key="esc_q")

    if only_dispo:
        g = g[g["ESCROW dispo"] > 1e-9]
    if q:
        qn = q.lower().strip()
        def _m2(r):
            hay = " ".join([
                _safe_str(r.get("Nom","")),
                _safe_str(r.get("ID_Client","")),
                _safe_str(r.get("Visa","")),
                str(r.get(DOSSIER_COL,"")),
            ]).lower()
            return qn in hay
        g = g[g.apply(_m2, axis=1)]

    st.markdown("### Transferts à marquer")
    if g.empty:
        st.caption("Aucun dossier avec ESCROW disponible selon les filtres.")
    else:
        for _, r in g.sort_values(["_Année_","_MoisNum_","Nom"], na_position="last").iterrows():
            with st.expander(f"{r.get('Nom','')} — {r.get('ID_Client','')}  |  Visa: {r.get('Visa','')}  |  Dispo: {_fmt_money_us(r['ESCROW dispo'])}", expanded=False):
                amt = st.number_input(
                    "Montant à marquer comme transféré (US $)",
                    min_value=0.0,
                    value=float(r["ESCROW dispo"]),
                    step=10.0,
                    format="%.2f",
                    key=f"esc_amt_{_safe_str(r.get('ID_Client',''))}"
                )
                dte = st.date_input("Date transfert", value=date.today(), key=f"esc_dt_{_safe_str(r.get('ID_Client',''))}")
                if st.button("✅ Marquer transféré", key=f"esc_mark_{_safe_str(r.get('ID_Client',''))}"):
                    if amt <= 0:
                        st.warning("Montant > 0 requis."); st.stop()
                    base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
                    # Localiser la ligne par ID_Client
                    idc = _safe_str(r.get("ID_Client",""))
                    if idc and "ID_Client" in base_raw.columns:
                        idxs = base_raw.index[base_raw["ID_Client"].astype(str)==idc].tolist()
                        if not idxs:
                            st.error("Ligne introuvable."); st.stop()
                        ridx = idxs[0]
                    else:
                        st.error("ID_Client manquant."); st.stop()

                    row = base_raw.loc[ridx].to_dict()
                    # maj montant transféré
                    prev = 0.0
                    try: prev = float(row.get("ESCROW transféré (US $)", 0.0) or 0.0)
                    except Exception: prev = 0.0
                    row["ESCROW transféré (US $)"] = prev + float(amt)

                    # journal
                    try:
                        j = json.loads(_safe_str(row.get("Journal ESCROW","[]")) or "[]")
                        if not isinstance(j,list): j=[]
                    except Exception:
                        j=[]
                    j.append({"ts": str(dte), "amount": float(amt), "note": "Transfert ESCROW -> Courant"})
                    row["Journal ESCROW"] = json.dumps(j, ensure_ascii=False)

                    base_raw.loc[ridx] = row
                    base_norm = normalize_clients(base_raw.copy())
                    write_sheet_inplace(clients_path, sheet_choice, base_norm)
                    st.success("Transfert enregistré.")
                    st.rerun()

    st.markdown("---")
    st.markdown("### Historique ESCROW (par client)")

    # Sélection client pour l'historique
    if not df_clients.empty:
        labels = (df_clients.get("Nom","").astype(str) + " — " + df_clients.get("ID_Client","").astype(str)).fillna("")
        pick_idx = st.selectbox("Client", options=list(df_clients.index),
                                format_func=lambda i: labels.iloc[i],
                                key="esc_hist_sel")
        row = df_clients.loc[pick_idx]
        st.write(f"**Honoraires**: {_fmt_money_us(float(row.get(HONO,0.0)))}  |  **Payé**: {_fmt_money_us(float(row.get('Payé',0.0)))}  |  **Transféré**: {_fmt_money_us(float(row.get('ESCROW transféré (US $)',0.0)))}")
        try:
            journal = json.loads(_safe_str(row.get("Journal ESCROW","[]")) or "[]")
            if not isinstance(journal, list): journal=[]
        except Exception:
            journal=[]
        if journal:
            jdf = pd.DataFrame(journal)
            if "amount" in jdf.columns:
                jdf["amount"] = jdf["amount"].astype(float).map(_fmt_money_us)
            if "ts" in jdf.columns:
                jdf["ts"] = jdf["ts"].astype(str)
            st.dataframe(jdf.sort_values(by=[c for c in ["ts"] if c in jdf.columns]).reset_index(drop=True),
                         use_container_width=True)
        else:
            st.caption("Aucune écriture ESCROW.")