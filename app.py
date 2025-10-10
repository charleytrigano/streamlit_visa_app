from __future__ import annotations

# =========================
# VISA MANAGER — APP (Part 1/2)
# =========================

import json
import zipfile
from io import BytesIO
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import streamlit as st
import openpyxl  # pour ExcelWriter

# --------- Config ----------
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")
st.caption("Gestion Clients, Visas, Paiements, ESCROW et Analyses — Excel en source unique")

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
    "Categorie", "Visa",
    HONO, AUTRE, TOTAL, "Payé", "Reste",
    PAY_JSON, "ESCROW transféré (US $)", "Journal ESCROW",
    "Dossier envoyé","Date envoyé",
    "Dossier approuvé","Date approuvé",
    "RFE","Date RFE",
    "Dossier refusé","Date refusé",
    "Dossier annulé","Date annulé",
]

# ------- Mémoire chemins + exports -------
STATE_LAST = "last_excel_paths"  # (clients_path, visa_path)

def save_last_paths(clients_path: str, visa_path: str) -> None:
    st.session_state[STATE_LAST] = (clients_path, visa_path)

def load_last_paths(default_clients: str, default_visa: str) -> tuple[str, str]:
    return st.session_state.get(STATE_LAST, (default_clients, default_visa))

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

def excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Construit un .xlsx en mémoire contenant plusieurs feuilles."""
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        for name, df in sheets.items():
            df2 = _uniquify_columns(df.copy())
            df2.to_excel(wr, sheet_name=name, index=False)
    bio.seek(0)
    return bio.getvalue()

def detect_visa_sheet_name(xlsx_path: str | Path) -> str:
    """Retourne la 1re feuille contenant 'Categorie' (sinon 1re feuille)."""
    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        return "Visa"
    for sn in xls.sheet_names:
        try:
            tmp = pd.read_excel(xlsx_path, sheet_name=sn, nrows=5)
            tmp = _uniquify_columns(tmp)
            tmp.columns = tmp.columns.map(str).str.strip()
            if "Categorie" in tmp.columns:
                return sn
        except Exception:
            continue
    return xls.sheet_names[0] if xls.sheet_names else "Visa"

def make_zip_clients_visa(clients_df: pd.DataFrame, visa_df: pd.DataFrame) -> bytes:
    """Crée un ZIP avec Clients.xlsx et Visa.xlsx."""
    clients_bytes = excel_bytes({"Clients": clients_df})
    visa_bytes    = excel_bytes({"Visa": visa_df})
    bio = BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Clients.xlsx", clients_bytes)
        zf.writestr("Visa.xlsx",    visa_bytes)
    bio.seek(0)
    return bio.getvalue()

# --------- Helpers généraux ----------
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

# --------- Templates & existence ----------
def _create_clients_template(path: str|Path, sheet_name: str="Clients") -> None:
    df = pd.DataFrame(columns=CLIENTS_COLUMNS)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        df.to_excel(wr, sheet_name=sheet_name, index=False)

def _create_visa_template(path: str|Path, sheet_name: str="Visa") -> None:
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

# --------- Visa parsing robuste ----------
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, list[str]]:
    """
    Scan toutes les feuilles si nécessaire ; génère {Categorie: [ 'Sous-categorie Case', ... ]}.
    Interprète les cases cochées: 1, x, ✓, true, oui, yes, ...
    """
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
        dfv = _uniquify_columns(dfv)
        dfv.columns = dfv.columns.map(str).str.strip()
        if "Categorie" not in dfv.columns:
            continue
        sub_col = None
        for c in ["Sous-categorie 1","Sous-categorie","Sous-categories 1","Sous-categories"]:
            if c in dfv.columns:
                sub_col = c; break
        if not sub_col:
            dfv["_Sous_"] = ""
            sub_col = "_Sous_"
        checkbox_cols = [c for c in dfv.columns if c not in {"Categorie", sub_col}]
        def _is_checked(v) -> bool:
            if pd.isna(v): return False
            if isinstance(v,(int,float)): return float(v) != 0.0
            s = str(v).strip().lower()
            return s in {"1","x","✓","true","vrai","oui","yes","y","o"}
        out: dict[str, list[str]] = {}
        for _, row in dfv.iterrows():
            cat = _safe_str(row.get("Categorie","")).strip()
            if not cat: continue
            sous = _safe_str(row.get(sub_col,"")).strip()
            for col in checkbox_cols:
                if _is_checked(row.get(col, None)):
                    lbl = f"{sous} {col}".strip()
                    out.setdefault(cat, []).append(lbl)
        if out:
            for k,v in out.items():
                out[k] = sorted(set(v))
            return out
    return {}

# --------- Clients I/O ----------
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

    df["_Année_"] = df["Date"].apply(lambda d: d.year if pd.notna(d) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda d: d.month if pd.notna(d) else pd.NA)
    df["_Mois_"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) else pd.NA)

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
    # supprimer puis réécrire la feuille
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

# --------- Barre latérale : chemins + mémoire + import/export ZIP ---------
st.sidebar.header("📂 Fichiers (mémoire)")

clients_path_default, visa_path_default = load_last_paths(DEFAULT_CLIENTS, DEFAULT_VISA)
clients_path = st.sidebar.text_input("Fichier Clients (.xlsx)", value=clients_path_default, key="cli_path")
visa_path    = st.sidebar.text_input("Fichier Visa (.xlsx)",    value=visa_path_default, key="visa_path")

ensure_files_exist(clients_path, visa_path)
save_last_paths(clients_path, visa_path)

# Feuille Clients valide
try:
    sheet_choice = safe_excel_first_sheet(clients_path, preferred="Clients")
except Exception:
    _create_clients_template(clients_path); sheet_choice = "Clients"

# Charger données
try:
    df_clients = read_sheet(clients_path, sheet_name=sheet_choice)
except Exception:
    _create_clients_template(clients_path); df_clients = read_sheet(clients_path, "Clients")

try:
    visa_choices_by_category = parse_visa_sheet(visa_path, sheet_name=None)  # auto-detect
except Exception:
    _create_visa_template(visa_path); visa_choices_by_category = parse_visa_sheet(visa_path, sheet_name=None)

# Import / Export UI
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
    visa_choices_by_category = parse_visa_sheet(visa_path, sheet_name=None)

if st.sidebar.button("🔄 Recharger Visa"):
    visa_choices_by_category = parse_visa_sheet(visa_path, sheet_name=None)
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

with st.sidebar.expander("🔎 Aperçu Visa détecté", expanded=False):
    if visa_choices_by_category:
        for cat, opts in visa_choices_by_category.items():
            st.write(f"**{cat}** → {', '.join(opts)}")
    else:
        st.caption("Aucune option Visa détectée.")


# ================== TABS ==================
tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs(
    ["Dashboard", "Clients", "Analyses", "ESCROW"]
)

# ================ DASHBOARD ================
with tab_dash:
    st.subheader("📊 Tableau de bord — Synthèse")

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
        DOSSIER_COL,"ID_Client","Nom","Categorie","Visa","Date","Mois",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"
    ] if c in view.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
    view_sorted = view.sort_values(by=sort_cols) if sort_cols else view

    st.dataframe(_uniquify_columns(view_sorted[show_cols].reset_index(drop=True)), use_container_width=True)

# ================ CLIENTS (CRUD + paiements) ================
with tab_clients:
    st.subheader("👥 Clients — Créer / Modifier / Supprimer / Paiements")

    live = df_clients.copy()

    cL, cR = st.columns([1,1])

    # --- Sélection client à droite
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
        cats = sorted(visa_choices_by_category.keys())
        new_cat = st.selectbox("Categorie", options=[""]+cats, index=0, key="new_cat_sel")
        visa_opts = visa_choices_by_category.get(new_cat, [])
        if new_cat and not visa_opts:
            st.warning("Aucune option Visa détectée pour cette catégorie (référentiel Visa).")
        new_visa = st.selectbox("Visa (auto depuis onglet Visa)", options=[""]+visa_opts, index=0, key="new_visa_sel")

        new_hono = st.number_input(HONO, min_value=0.0, step=10.0, format="%.2f", key="new_hono")
        new_autr = st.number_input(AUTRE, min_value=0.0, step=10.0, format="%.2f", key="new_autre")

        if st.button("💾 Créer", key="btn_create"):
            if not new_name: st.warning("Nom obligatoire."); st.stop()
            if not new_cat: st.warning("Categorie obligatoire."); st.stop()
            if not new_visa: st.warning("Visa obligatoire."); st.stop()

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

    # --- Edition & Paiements pour le client sélectionné
    idx = sel_idx
    ed = sel_row.to_dict()

    e1,e2,e3 = st.columns(3)
    with e1:
        ed_nom  = st.text_input("Nom", value=_safe_str(ed.get("Nom","")), key=f"ed_nom_{idx}")
        ed_date = st.date_input("Date de création",
                                value=(pd.to_datetime(ed.get("Date")).date() if pd.notna(ed.get("Date")) else date.today()),
                                key=f"ed_date_{idx}")
    with e2:
        cats = sorted(visa_choices_by_category.keys())
        curr_cat = _safe_str(ed.get("Categorie",""))
        ed_cat = st.selectbox("Categorie",
                              options=[""]+cats,
                              index=(cats.index(curr_cat)+1 if curr_cat in cats else 0),
                              key=f"ed_cat_{idx}")
        visa_opts = visa_choices_by_category.get(ed_cat, [])
        if ed_cat and not visa_opts:
            st.warning("Aucune option Visa détectée pour cette catégorie (référentiel Visa).")
    with e3:
        curr_visa = _safe_str(ed.get("Visa",""))
        ed_visa = st.selectbox("Visa (auto depuis onglet Visa)",
                               options=[""]+visa_opts,
                               index=(visa_opts.index(curr_visa)+1 if curr_visa in visa_opts else 0),
                               key=f"ed_visa_{idx}")
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
        row["Visa"] = ed_visa
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

# ================ ANALYSES =================
with tab_analyses:
    st.subheader("📈 Analyses — volumes & financier")

    yearsA  = sorted([int(y) for y in df_clients["_Année_"].dropna().unique()]) if not df_clients.empty else []
    monthsA = [f"{m:02d}" for m in sorted([int(m) for m in df_clients["_MoisNum_"].dropna().unique()])] if not df_clients.empty else []
    catsA   = sorted(df_clients["Categorie"].dropna().astype(str).unique().tolist())
    visasA  = sorted(df_clients["Visa"].dropna().astype(str).unique().tolist())

    c1,c2,c3,c4 = st.columns([1,1,1,2])
    with c1: sel_years  = st.multiselect("Année", yearsA, default=[], key="a_years")
    with c2: sel_months = st.multiselect("Mois (MM)", monthsA, default=[], key="a_months")
    with c3: sel_cats   = st.multiselect("Catégories", catsA, default=[], key="a_cats")
    with c4: sel_visas  = st.multiselect("Visa", visasA, default=[], key="a_visas")

    ff = df_clients.copy()
    if sel_years:  ff = ff[ff["_Année_"].isin(sel_years)]
    if sel_months: ff = ff[ff["_Mois_"].astype(str).isin(sel_months)]
    if sel_cats:   ff = ff[ff["Categorie"].astype(str).isin(sel_cats)]
    if sel_visas:  ff = ff[ff["Visa"].astype(str).isin(sel_visas)]

    st.info(f"{len(ff)} dossiers dans le périmètre.")
    sk1,sk2,sk3,sk4 = st.columns(4)
    sk1.metric("Dossiers", f"{len(ff)}")
    sk2.metric("Honoraires", _fmt_money_us(_safe_num_series(ff,HONO).sum()))
    sk3.metric("Payé", _fmt_money_us(_safe_num_series(ff,"Payé").sum()))
    sk4.metric("Reste", _fmt_money_us(_safe_num_series(ff,"Reste").sum()))

    st.markdown("---")
    st.markdown("#### 📆 Par année")
    if not ff.empty and ff["_Année_"].notna().any():
        def _sum(df_loc, col): return _safe_num_series(df_loc,col).sum()
        grpY = ff.groupby("_Année_", dropna=True).apply(
            lambda g: pd.Series({
                "Dossiers": int(g.shape[0]),
                "Honoraires": _sum(g,HONO),
                "Payé": _sum(g,"Payé"),
                "Reste": _sum(g,"Reste")
            })
        ).reset_index().rename(columns={"_Année_":"Année"}).sort_values("Année")
        st.dataframe(_uniquify_columns(grpY), use_container_width=True)
    else:
        st.caption("Aucune année exploitable.")

    st.markdown("#### 🗓️ Par mois (toutes années)")
    if not ff.empty and ff["_Mois_"].notna().any():
        def _sum(df_loc, col): return _safe_num_series(df_loc,col).sum()
        grpM = ff.groupby("_Mois_", dropna=True).apply(
            lambda g: pd.Series({
                "Dossiers": int(g.shape[0]),
                "Honoraires": _sum(g,HONO),
                "Payé": _sum(g,"Payé"),
                "Reste": _sum(g,"Reste")
            })
        ).reset_index().rename(columns={"_Mois_":"Mois"}).sort_values("Mois")
        st.dataframe(_uniquify_columns(grpM), use_container_width=True)
    else:
        st.caption("Aucun mois exploitable.")

    st.markdown("#### 📋 Détails")
    det = ff.copy()
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c in det.columns: det[c] = _safe_num_series(det,c).map(_fmt_money_us)
    if "Date" in det.columns: det["Date"] = det["Date"].astype(str)
    show_cols = [c for c in [
        DOSSIER_COL,"ID_Client","Nom","Categorie","Visa","Date","Mois",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"
    ] if c in det.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
    det_sorted = det.sort_values(by=sort_cols) if sort_cols else det
    st.dataframe(_uniquify_columns(det_sorted[show_cols].reset_index(drop=True)), use_container_width=True)

# ================ ESCROW =================
with tab_escrow:
    st.subheader("🏦 ESCROW — dépôts, transferts & alertes")

    rows = []
    for i, r in df_clients.iterrows():
        hono = float(r.get(HONO,0.0) or 0.0)
        try:
            plist = json.loads(_safe_str(r.get(PAY_JSON,"[]")) or "[]")
            if not isinstance(plist,list): plist=[]
        except Exception:
            plist=[]
        paid_json = sum(float(it.get("amount",0.0) or 0.0) for it in plist)
        pay = max(float(r.get("Payé",0.0) or 0.0), paid_json)
        transf = float(r.get("ESCROW transféré (US $)",0.0) or 0.0)
        dispo = max(min(pay, hono) - transf, 0.0)
        rows.append({
            "idx": i,
            DOSSIER_COL: r.get(DOSSIER_COL,""),
            "ID_Client": r.get("ID_Client",""),
            "Nom": r.get("Nom",""),
            "Categorie": r.get("Categorie",""),
            "Visa": r.get("Visa",""),
            "Dossier envoyé": bool(r.get("Dossier envoyé",False)),
            HONO: hono,
            "Payé_calc": pay,
            "ESCROW transféré (US $)": transf,
            "ESCROW dispo": dispo,
            "Journal ESCROW": _safe_str(r.get("Journal ESCROW","[]")),
        })
    jdf = pd.DataFrame(rows)

    c1,c2,c3,c4 = st.columns([1,1,1,2])
    with c1: only_dispo = st.toggle("Uniquement ESCROW disponible", value=True, key="esc_only_dispo")
    with c2: only_sent  = st.toggle("Uniquement dossiers envoyés", value=False, key="esc_only_sent")
    with c3: order_dispo = st.toggle("Trier par dispo", value=True, key="esc_order")
    with c4: q = st.text_input("Recherche (Nom/ID/Dossier/Visa)", "", key="esc_q")

    if only_dispo: jdf = jdf[jdf["ESCROW dispo"] > 0.0]
    if only_sent:  jdf = jdf[jdf["Dossier envoyé"]==True]
    if q:
        qn = q.lower().strip()
        def _m(row):
            hay = " ".join([
                _safe_str(row.get("Nom","")), _safe_str(row.get("ID_Client","")),
                str(row.get(DOSSIER_COL,"")), _safe_str(row.get("Visa","")), _safe_str(row.get("Categorie","")),
            ]).lower()
            return qn in hay
        jdf = jdf[jdf.apply(_m, axis=1)]
    if order_dispo and not jdf.empty:
        jdf = jdf.sort_values("ESCROW dispo", ascending=False)

    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.1rem}
    .small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(jdf)}")
    k2.metric("Honoraires périmètre", _fmt_money_us(float(jdf[HONO].sum()) if not jdf.empty else 0.0))
    k3.metric("ESCROW transféré", _fmt_money_us(float(jdf["ESCROW transféré (US $)"].sum()) if not jdf.empty else 0.0))
    k4.metric("ESCROW dispo", _fmt_money_us(float(jdf["ESCROW dispo"].sum()) if not jdf.empty else 0.0))
    st.markdown('</div>', unsafe_allow_html=True)

    if jdf.empty:
        st.info("Aucun dossier avec les filtres.")
        st.stop()

    show = jdf.copy()
    for c in [HONO, "Payé_calc", "ESCROW transféré (US $)", "ESCROW dispo"]:
        show[c] = show[c].map(_fmt_money_us)
    st.dataframe(
        show[[DOSSIER_COL,"ID_Client","Nom","Categorie","Visa",HONO,"Payé_calc","ESCROW transféré (US $)","ESCROW dispo","Dossier envoyé"]]
        .reset_index(drop=True),
        use_container_width=True
    )

    st.markdown("### ↗️ Enregistrer un transfert ESCROW")
    st.caption("Disponible = min(Payé, Honoraires) − Transféré.")

    for _, r in jdf.iterrows():
        st.markdown(f"**{r['Nom']} — {r['ID_Client']} — Dossier {r[DOSSIER_COL]}**")
        cA,cB,cC,cD = st.columns([1,1,1,2])
        with cA: st.write("Disponible :", _fmt_money_us(float(r["ESCROW dispo"])))
        with cB: t_date = st.date_input("Date transfert", value=date.today(), key=f"esc_dt_{r['ID_Client']}")
        with cC: amt = st.number_input("Montant (US $)", min_value=0.0, value=float(r["ESCROW dispo"]), step=10.0, format="%.2f", key=f"esc_amt_{r['ID_Client']}")
        with cD: note = st.text_input("Note (optionnel)", value="", key=f"esc_note_{r['ID_Client']}")

        ok = float(r["ESCROW dispo"]) > 0 and float(amt) > 0 and float(amt) <= float(r["ESCROW dispo"]) + 1e-9
        if st.button("💸 Enregistrer transfert", key=f"esc_btn_{r['ID_Client']}"):
            if not ok:
                st.warning("Montant invalide."); st.stop()

            base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
            idc = _safe_str(r.get("ID_Client",""))
            if idc and "ID_Client" in base_raw.columns:
                idxs = base_raw.index[base_raw["ID_Client"].astype(str) == idc].tolist()
                real_idx = idxs[0] if idxs else int(r["idx"])
            else:
                real_idx = int(r["idx"])

            if real_idx is None or not (0 <= real_idx < len(base_raw)):
                st.error("Ligne introuvable."); st.stop()

            row = base_raw.loc[real_idx].to_dict()
            curr_tr = float(row.get("ESCROW transféré (US $)",0.0) or 0.0)
            row["ESCROW transféré (US $)"] = curr_tr + float(amt)

            try:
                jlog = json.loads(_safe_str(row.get("Journal ESCROW","[]")) or "[]")
                if not isinstance(jlog,list): jlog=[]
            except Exception:
                jlog=[]
            jlog.append({"ts": datetime.now().isoformat(timespec="seconds"), "date": str(t_date), "amount": float(amt), "note": note})
            row["Journal ESCROW"] = json.dumps(jlog, ensure_ascii=False)

            base_raw.loc[real_idx] = row
            base_norm = normalize_clients(base_raw.copy())
            write_sheet_inplace(clients_path, sheet_choice, base_norm)
            st.success("✅ Transfert enregistré.")
            st.rerun()