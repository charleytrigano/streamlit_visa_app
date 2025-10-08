=========================
# VISA APP — PARTIE 1/5
# =========================

from __future__ import annotations
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
import json, unicodedata

# --- Constantes principales ---
DOSSIER_COL = "Dossier N"
HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"
S_ENVOYE, D_ENVOYE   = "Dossier envoyé", "Date envoyé"
S_APPROUVE, D_APPROUVE = "Dossier approuvé", "Date approuvé"
S_RFE, D_RFE         = "RFE", "Date RFE"
S_REFUSE, D_REFUSE   = "Dossier refusé", "Date refusé"
S_ANNULE, D_ANNULE   = "Dossier annulé", "Date annulé"

# --- Fonctions utilitaires ---
def _fmt_money_us(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return "$0.00"

def _to_num(s):
    if s is None: return 0.0
    s = pd.Series(s)
    try:
        s = s.astype(str).replace("", np.nan)
        s = s.str.replace(r"[^\d,.\-]", "", regex=True)
        s = s.str.replace(",", ".", regex=False)
        s = pd.to_numeric(s, errors="coerce").fillna(0.0)
    except Exception:
        s = pd.Series([0.0]*len(s))
    return s

def _safe_str(x) -> str:
    if pd.isna(x): return ""
    return str(x)

def _norm_txt(x: str) -> str:
    """Normalise une chaîne : sans accents, minuscules, espaces uniformisés."""
    s = "" if x is None else str(x)
    s = s.strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = " ".join(s.lower().split())
    return s

def _find_col(df: pd.DataFrame, targets: list[str]) -> str | None:
    """Cherche une colonne existante dont le nom correspond aux cibles (tolérance casse/accents)."""
    if df is None or df.empty:
        return None
    norm_map = { _norm_txt(c): str(c) for c in df.columns.astype(str) }
    for t in targets:
        nt = _norm_txt(t)
        if nt in norm_map:
            return norm_map[nt]
    # fallback : recherche partielle
    for t in targets:
        nt = _norm_txt(t)
        for k, orig in norm_map.items():
            if nt in k:
                return orig
    return None

def filter_by_selection(df: pd.DataFrame, sel: dict, visas_aut: list[str] | None = None) -> pd.DataFrame:
    """Filtre df selon Catégorie / Visa / Sous-type, insensible aux accents/espaces/casse."""
    if df is None or df.empty:
        return df
    f = df.copy()

    # Colonnes à trouver
    col_cat  = _find_col(f, ["Catégorie","Categorie","Category"])
    col_visa = _find_col(f, ["Visa"])
    col_sub  = _find_col(f, ["Sous-type","Soustype","Sous type","Type","Subtype"])

    # Colonnes normalisées temporaires
    if col_cat:  f["__norm_cat"]  = f[col_cat].astype(str).map(_norm_txt)
    else:        f["__norm_cat"]  = ""
    if col_visa: f["__norm_visa"] = f[col_visa].astype(str).map(_norm_txt)
    else:        f["__norm_visa"] = ""
    if col_sub:  f["__norm_sub"]  = f[col_sub].astype(str).map(_norm_txt)
    else:        f["__norm_sub"]  = ""

    sel_cat  = _norm_txt(sel.get("Catégorie",""))
    sel_visa = _norm_txt(sel.get("Visa",""))
    sel_sub  = _norm_txt(sel.get("Sous-type",""))

    if sel_cat:
        f = f[f["__norm_cat"] == sel_cat]
    if sel_visa:
        f = f[f["__norm_visa"] == sel_visa]
    if sel_sub:
        f = f[f["__norm_sub"] == sel_sub]

    # Si une liste de visas autorisés est fournie
    if visas_aut:
        visas_norm = {_norm_txt(v) for v in visas_aut}
        f = f[f["__norm_visa"].isin(visas_norm)]

    return f.drop(columns=[c for c in f.columns if c.startswith("__norm_")], errors="ignore")

# --- Lecture du fichier et normalisation ---

def normalize_dataframe(df: pd.DataFrame, visa_ref=None) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    rename = {}
    for c in df.columns:
        lc = str(c).lower().strip()
        if lc in ("montant honoraires", "montant honoraires (us $)", "honoraires", "montant"):
            rename[c] = HONO
        elif lc in ("autres frais", "autres frais (us $)", "autres"):
            rename[c] = AUTRE
        elif lc in ("total", "total (us $)"):
            rename[c] = TOTAL
        elif lc == "dossier n" or lc == "dossier":
            rename[c] = DOSSIER_COL
        elif lc in ("reste (us $)", "solde (us $)", "solde"):
            rename[c] = "Reste"
        elif lc in ("paye (us $)","payé (us $)","paye","payé"):
            rename[c] = "Payé"
        elif lc == "sous-type" or lc == "sous type":
            rename[c] = "Sous-type"
        elif lc in ("categorie","catégorie"):              # ✅ AJOUT FONDAMENTAL
            rename[c] = "Catégorie"

    df = df.rename(columns=rename)

    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c in df.columns:
            df[c] = _to_num(df[c])

    if HONO in df.columns and AUTRE in df.columns:
        df[TOTAL] = df[HONO] + df[AUTRE]
    elif TOTAL not in df.columns:
        df[TOTAL] = 0.0

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Mois"] = df["Date"].dt.month.astype("Int64")

    # Initialiser colonnes statut si absentes
    for s in [S_ENVOYE,S_APPROUVE,S_RFE,S_REFUSE,S_ANNULE]:
        if s not in df.columns:
            df[s] = False
    for d in [D_ENVOYE,D_APPROUVE,D_RFE,D_REFUSE,D_ANNULE]:
        if d not in df.columns:
            df[d] = pd.NaT

    return df

# --- Fichier Excel ---
def list_sheets(path: Path) -> list[str]:
    try:
        return pd.ExcelFile(path).sheet_names
    except Exception:
        return []

def read_sheet(path: Path, sheet_name: str, normalize=True, visa_ref=None) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()
    return normalize_dataframe(df, visa_ref=visa_ref) if normalize else df

# --- Gestion du workspace ---
def workspace_file() -> Path:
    w = Path("workspace.json")
    return w

def save_workspace_path(path: Path):
    try:
        json.dump({"path": str(path)}, open(workspace_file(), "w"))
    except Exception:
        pass

def _load_last_path() -> Path | None:
    f = workspace_file()
    if not f.exists(): return None
    try:
        data = json.load(open(f))
        p = Path(data.get("path",""))
        return p if p.exists() else None
    except Exception:
        return None

def current_file_path() -> Path | None:
    p = st.session_state.get("current_path")
    if p: return Path(p)
    p2 = _load_last_path()
    return p2

def set_current_file_from_upload(uploaded) -> Path | None:
    """Sauvegarde le fichier téléchargé dans le répertoire courant."""
    try:
        target = Path("uploaded.xlsx")
        with open(target, "wb") as f:
            f.write(uploaded.read())
        save_workspace_path(target)
        st.session_state["current_path"] = str(target)
        return target
    except Exception:
        return None



# =========================
# VISA APP — PARTIE 2/5
# =========================

st.set_page_config(page_title="Visa Manager", layout="wide")

# ---------- Barre latérale : gestion du fichier ----------
st.sidebar.header("📂 Fichier Excel")
uploaded = st.sidebar.file_uploader("Charger/Remplacer fichier (.xlsx)", type=["xlsx"], key="uploader")
if uploaded is not None:
    p = set_current_file_from_upload(uploaded)
    if p:
        st.sidebar.success(f"Fichier chargé: {p.name}")

path_text = st.sidebar.text_input("Ou saisir le chemin d’un fichier existant", value=st.session_state.get("current_path", ""))
colB1, colB2 = st.sidebar.columns(2)
if colB1.button("📄 Ouvrir ce fichier", key="open_manual"):
    p = Path(path_text)
    if p.exists():
        save_workspace_path(p)
        st.sidebar.success(f"Ouvert: {p.name}")
        st.rerun()
    else:
        st.sidebar.error("Chemin invalide.")
if colB2.button("♻️ Reprendre le dernier fichier", key="open_last"):
    p = _load_last_path()
    if p:
        save_workspace_path(p)
        st.sidebar.success(f"Repris: {p.name}")
        st.rerun()
    else:
        st.sidebar.info("Aucun fichier précédemment enregistré.")

current_path = current_file_path()
if current_path is None:
    st.warning("Aucun fichier sélectionné. Charge un .xlsx ou choisis un chemin valide.")
    st.stop()

# ---------- Feuilles disponibles ----------
sheets = list_sheets(current_path)
if not sheets:
    st.error("Impossible de lire le classeur. Assure-toi que le fichier est un .xlsx valide.")
    st.stop()

st.sidebar.markdown("---")
st.sidebar.write("**Feuilles détectées :**")
for i, sn in enumerate(sheets):
    st.sidebar.write(f"- {i+1}. {sn}")

# Détection d’une feuille « clients »
client_target_sheet = None
for sn in sheets:
    df_try = read_sheet(current_path, sn, normalize=False)
    if {"Nom", "Visa"}.issubset(set(df_try.columns.astype(str))):
        client_target_sheet = sn
        break

sheet_choice = st.sidebar.selectbox(
    "Feuille à afficher sur le Dashboard :",
    sheets,
    index=max(0, sheets.index(client_target_sheet) if client_target_sheet in sheets else 0),
    key="sheet_choice_select"
)

# ---------- Titre & onglets ----------
st.title("🛂 Visa Manager — US $")

tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs(
    ["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"]
)

# ---------- Référentiels Visa ----------
visa_ref_tree   = read_visa_reference_tree(current_path)  # Catégorie / Visa / Sous-type
visa_ref_simple = read_visa_reference(current_path)       # Mapping simple Catégorie/Visa

# ================= DASHBOARD =================
with tab_dash:
    df_raw = read_sheet(current_path, sheet_choice, normalize=False)

    # Si la feuille active est le référentiel Visa, on l’affiche simplement
    if looks_like_reference(df_raw) and sheet_choice == "Visa":
        st.subheader("📄 Référentiel — Catégories / Visa / Sous-type")
        st.dataframe(visa_ref_tree, use_container_width=True)
        st.stop()

    # Données normalisées pour Dashboard
    df = read_sheet(current_path, sheet_choice, normalize=True, visa_ref=visa_ref_simple)

    # --- Filtres (clés uniques dash_*) ---
    st.markdown("### 🔎 Filtres (Catégorie → Visa → Sous-type)")
    with st.container():
        cTopL, cTopR = st.columns([1,2])
        show_all = cTopL.checkbox("Afficher tous les dossiers", value=False, key=f"dash_show_all_{sheet_choice}")
        cTopL.caption("Sélection hiérarchique")

        with cTopL:
            sel_path_dash = cascading_visa_picker_tree(visa_ref_tree, key_prefix=f"dash_tree_{sheet_choice}")
        visas_aut = visas_autorises_from_tree(visa_ref_tree, sel_path_dash)

        cR1, cR2, cR3 = cTopR.columns(3)
        years  = sorted({d.year for d in df["Date"] if pd.notna(d)}) if "Date" in df.columns else []
        months = sorted([m for m in df["Mois"].dropna().unique()]) if "Mois" in df.columns else []
        sel_years  = cR1.multiselect("Année", years, default=[], key=f"dash_years_{sheet_choice}")
        sel_months = cR2.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{sheet_choice}")
        include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True, key=f"dash_na_{sheet_choice}")

    # Application des filtres
    f = df.copy()
    if not show_all:
        for col in ["Catégorie","Visa","Sous-type"]:
            if col in f.columns:
                val = str(sel_path_dash.get(col, "")).strip()
                if val:
                    f = f[f[col].astype(str) == val]
        if "Visa" in f.columns and visas_aut:
            f = f[f["Visa"].astype(str).isin(visas_aut)]

    if "Date" in f.columns and sel_years:
        mask = f["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask |= f["Date"].isna()
        f = f[mask]
    if "Mois" in f.columns and sel_months:
        mask = f["Mois"].isin(sel_months)
        if include_na_dates: mask |= f["Mois"].isna()
        f = f[mask]

    hidden = len(df) - len(f)
    if hidden > 0:
        st.caption(f"🔎 {hidden} ligne(s) masquée(s) par les filtres.")

    # KPI compacts
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Total (US $)", _fmt_money_us(float(f.get(TOTAL, pd.Series(dtype=float)).sum())))
    k3.metric("Payé (US $)", _fmt_money_us(float(f.get("Payé", pd.Series(dtype=float)).sum())))
    k4.metric("Solde (US $)", _fmt_money_us(float(f.get("Reste", pd.Series(dtype=float)).sum())))
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    st.subheader("📋 Données (aperçu)")
    cols_show = [c for c in [
        DOSSIER_COL,"ID_Client","Nom","Date","Mois",
        "Catégorie","Visa","Sous-type",
        HONO, AUTRE, TOTAL, "Payé","Reste",
        S_ENVOYE, D_ENVOYE, S_APPROUVE, D_APPROUVE, S_RFE, D_RFE, S_REFUSE, D_REFUSE, S_ANNULE, D_ANNULE
    ] if c in f.columns]
    view = f.copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in view.columns: view[col] = pd.to_numeric(view[col], errors="coerce").fillna(0.0).map(_fmt_money_us)
    if "Date" in view.columns: view["Date"] = view["Date"].astype(str)
    st.dataframe(view[cols_show], use_container_width=True)


# =========================
# VISA APP — PARTIE 3/5
# =========================

with tab_clients:
    st.subheader("👤 Clients — Créer / Modifier / Supprimer (écriture directe)")
    if client_target_sheet is None:
        st.warning("Aucune feuille *Clients* valide (Nom & Visa)."); st.stop()
    if st.button("🔄 Recharger le fichier", key="reload_btn_crud"):
        st.rerun()

    live_raw = read_sheet(current_path, client_target_sheet, normalize=False).copy()
    live_raw = ensure_dossier_numbers(live_raw)
    live_raw["_RowID"] = range(len(live_raw))

    action = st.radio("Action", ["Créer", "Modifier", "Supprimer"], horizontal=True, key="crud_action")

    # --- CREER ---
    if action == "Créer":
        st.markdown("### ➕ Nouveau client")
        # Colonnes minimales
        for must in [DOSSIER_COL,"ID_Client","Nom","Date","Mois",
                     "Catégorie","Visa","Sous-type",
                     HONO, AUTRE, TOTAL, "Payé","Reste", ESC_TR, ESC_JR] + STATUS_COLS + STATUS_DATES + ["Paiements"]:
            if must not in live_raw.columns:
                if must in {HONO, AUTRE, TOTAL, "Payé","Reste", ESC_TR}: live_raw[must]=0.0
                elif must in {"Paiements", ESC_JR, "Nom","Date","Mois","Catégorie","Visa","Sous-type"}:
                    live_raw[must]=""
                elif must in STATUS_DATES: live_raw[must]=""
                elif must in STATUS_COLS: live_raw[must]=False
                elif must==DOSSIER_COL: live_raw[must]=0
                else: live_raw[must]=""

        next_num = next_dossier_number(live_raw)
        with st.form("create_form", clear_on_submit=False):
            c0, c1, c2 = st.columns([1,1,1])
            c0.metric("Prochain Dossier N", f"{next_num}")
            nom_in = c1.text_input("Nom")
            d = c2.date_input("Date", value=date.today())

            st.caption("Sélection hiérarchique (Catégorie → Visa → Sous-type)")
            sel_path = cascading_visa_picker_tree(visa_ref_tree, key_prefix="create_tree")
            cat  = sel_path.get("Catégorie","")
            visa = sel_path.get("Visa","")
            stype= sel_path.get("Sous-type","")

            c5,c6 = st.columns(2)
            honoraires = c5.number_input(HONO, value=0.0, step=10.0, format="%.2f")
            autres     = c6.number_input(AUTRE, value=0.0, step=10.0, format="%.2f")

            st.markdown("#### État du dossier")
            r1c1, r1c2 = st.columns(2)
            v_env   = r1c1.checkbox(S_ENVOYE,  value=False)
            dt_env  = r1c2.date_input(D_ENVOYE, value=date.today(), disabled=not v_env, key="dt_env_cre")
            r2c1, r2c2 = st.columns(2)
            v_app   = r2c1.checkbox(S_APPROUVE, value=False)
            dt_app  = r2c2.date_input(D_APPROUVE, value=date.today(), disabled=not v_app, key="dt_app_cre")
            r3c1, r3c2 = st.columns(2)
            v_rfe   = r3c1.checkbox(S_RFE,      value=False)
            dt_rfe  = r3c2.date_input(D_RFE,    value=date.today(), disabled=not v_rfe, key="dt_rfe_cre")
            r4c1, r4c2 = st.columns(2)
            v_ref   = r4c1.checkbox(S_REFUSE,   value=False)
            dt_ref  = r4c2.date_input(D_REFUSE, value=date.today(), disabled=not v_ref, key="dt_ref_cre")
            r5c1, r5c2 = st.columns(2)
            v_ann   = r5c1.checkbox(S_ANNULE,   value=False)
            dt_ann  = r5c2.date_input(D_ANNULE, value=date.today(), disabled=not v_ann, key="dt_ann_cre")

            ok = st.form_submit_button("💾 Sauvegarder (dans le fichier)", type="primary")

        if ok:
            if v_rfe and not (v_env or v_ref or v_ann):
                st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché."); st.stop()
            existing_names = set(live_raw["Nom"].dropna().astype(str))
            base_name = _safe_str(nom_in); use_name = base_name
            if base_name in existing_names:
                k = 0
                while f"{base_name}-{k}" in existing_names: k += 1
                use_name = f"{base_name}-{k}"
            gen_id = _make_client_id_from_row({"Nom": use_name, "Date": d})
            existing_ids = set(live_raw["ID_Client"].astype(str)) if "ID_Client" in live_raw.columns else set()
            new_id = gen_id; n=1
            while new_id in existing_ids: n+=1; new_id=f"{gen_id}-{n:02d}"
            total = float((honoraires or 0.0)+(autres or 0.0))
            new_row = {
                DOSSIER_COL: int(next_num), "ID_Client": new_id, "Nom": use_name,
                "Date": str(d), "Mois": f"{d.month:02d}",
                "Catégorie": _safe_str(cat), "Visa": _safe_str(visa), "Sous-type": _safe_str(stype),
                HONO: float(honoraires or 0.0), AUTRE: float(autres or 0.0),
                TOTAL: total, "Payé": 0.0, "Reste": max(total, 0.0),
                ESC_TR: 0.0, ESC_JR: "", "Paiements": "",
                S_ENVOYE: bool(v_env),   D_ENVOYE:   (str(dt_env) if v_env else ""),
                S_APPROUVE: bool(v_app), D_APPROUVE: (str(dt_app) if v_app else ""),
                S_RFE: bool(v_rfe),      D_RFE:      (str(dt_rfe) if v_rfe else ""),
                S_REFUSE: bool(v_ref),   D_REFUSE:   (str(dt_ref) if v_ref else ""),
                S_ANNULE: bool(v_ann),   D_ANNULE:   (str(dt_ann) if v_ann else "")
            }
            live_after = pd.concat([live_raw.drop(columns=["_RowID"]), pd.DataFrame([new_row])], ignore_index=True)
            live_after = ensure_dossier_numbers(live_after)
            write_sheet_inplace(current_path, client_target_sheet, live_after); save_workspace_path(current_path)
            st.success(f"Client créé **dans le fichier** (Dossier N {next_num}). ✅"); st.rerun()

    # --- MODIFIER ---
    if action == "Modifier":
        st.markdown("### ✏️ Modifier un client (fiche + paiements + dates)")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            options = [(int(r["_RowID"]),
                        int(r.name),
                        f'{int(r.get(DOSSIER_COL,0))} — { _safe_str(r.get("ID_Client")) } — { _safe_str(r.get("Nom")) }')
                       for _,r in live_raw.iterrows()]
            labels  = [lab for _,__,lab in options]
            sel_lab = st.selectbox("Sélection", labels, key="edit_sel_label")
            sel_rowid, orig_pos, _ = [t for t in options if t[2]==sel_lab][0]
            idx = live_raw.index[live_raw["_RowID"]==sel_rowid][0]
            init = live_raw.loc[idx].to_dict()

            with st.form(f"edit_form_{sel_rowid}", clear_on_submit=False):
                c0, c1, c2 = st.columns([1,1,1])
                c0.metric("Dossier N", f'{int(init.get(DOSSIER_COL,0))}')
                nom = c1.text_input("Nom", value=_safe_str(init.get("Nom")), key=f"edit_nom_{sel_rowid}")
                try:
                    d_init = pd.to_datetime(init.get("Date")).date() if _safe_str(init.get("Date")) else date.today()
                except Exception:
                    d_init = date.today()
                d = c2.date_input("Date", value=d_init, key=f"edit_date_{sel_rowid}")

                st.caption("Sélection hiérarchique (Catégorie → Visa → Sous-type)")
                init_path = {
                    "Catégorie": _safe_str(init.get("Catégorie")),
                    "Visa": _safe_str(init.get("Visa")),
                    "Sous-type": _safe_str(init.get("Sous-type")).upper(),
                }
                sel_path = cascading_visa_picker_tree(visa_ref_tree, key_prefix=f"edit_tree_{sel_rowid}", init=init_path)
                cat  = sel_path.get("Catégorie","")
                visa = sel_path.get("Visa","")
                stype= sel_path.get("Sous-type","")

                def _f(v, alt=0.0):
                    try: return float(v)
                    except Exception: return float(alt)
                hono0  = _f(init.get(HONO, init.get("Montant", 0.0)))
                autre0 = _f(init.get(AUTRE, 0.0))
                paye0  = _f(init.get("Payé", 0.0))
                c5,c6 = st.columns(2)
                honoraires = c5.number_input(HONO, value=hono0, step=10.0, format="%.2f", key=f"edit_hono_{sel_rowid}")
                autres     = c6.number_input(AUTRE, value=autre0, step=10.0, format="%.2f", key=f"edit_autre_{sel_rowid}")
                c7,c8 = st.columns(2)
                total_preview = float(honoraires + autres); c7.metric("Total (US $)", _fmt_money_us(total_preview))
                st.caption(f"Payé actuel : {_fmt_money_us(paye0)} — Solde après sauvegarde : {_fmt_money_us(max(total_preview - paye0, 0.0))}")

                st.markdown("#### État du dossier (avec dates)")
                def _get_dt(key):
                    v = _safe_str(init.get(key))
                    try: return pd.to_datetime(v).date() if v else date.today()
                    except Exception: return date.today()

                r1c1, r1c2 = st.columns(2)
                v_env = r1c1.checkbox(S_ENVOYE, value=bool(init.get(S_ENVOYE)), key=f"edit_env_{sel_rowid}")
                dt_env = r1c2.date_input(D_ENVOYE, value=_get_dt(D_ENVOYE), disabled=not v_env, key=f"edit_dt_env_{sel_rowid}")

                r2c1, r2c2 = st.columns(2)
                v_app = r2c1.checkbox(S_APPROUVE, value=bool(init.get(S_APPROUVE)), key=f"edit_app_{sel_rowid}")
                dt_app = r2c2.date_input(D_APPROUVE, value=_get_dt(D_APPROUVE), disabled=not v_app, key=f"edit_dt_app_{sel_rowid}")

                r3c1, r3c2 = st.columns(2)
                v_rfe = r3c1.checkbox(S_RFE, value=bool(init.get(S_RFE)), key=f"edit_rfe_{sel_rowid}")
                dt_rfe = r3c2.date_input(D_RFE, value=_get_dt(D_RFE), disabled=not v_rfe, key=f"edit_dt_rfe_{sel_rowid}")

                r4c1, r4c2 = st.columns(2)
                v_ref = r4c1.checkbox(S_REFUSE, value=bool(init.get(S_REFUSE)), key=f"edit_ref_{sel_rowid}")
                dt_ref = r4c2.date_input(D_REFUSE, value=_get_dt(D_REFUSE), disabled=not v_ref, key=f"edit_dt_ref_{sel_rowid}")

                r5c1, r5c2 = st.columns(2)
                v_ann = r5c1.checkbox(S_ANNULE, value=bool(init.get(S_ANNULE)), key=f"edit_ann_{sel_rowid}")
                dt_ann = r5c2.date_input(D_ANNULE, value=_get_dt(D_ANNULE), disabled=not v_ann, key=f"edit_dt_ann_{sel_rowid}")

                ok_fiche = st.form_submit_button("💾 Enregistrer la fiche (dans le fichier)", type="primary")

            if ok_fiche:
                if v_rfe and not (v_env or v_ref or v_ann):
                    st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché."); st.stop()

                live = read_sheet(current_path, client_target_sheet, normalize=False).copy()

                # Re-trouve la ligne à mettre à jour
                t_idx = None
                key_id = _safe_str(init.get("ID_Client"))
                if key_id and "ID_Client" in live.columns:
                    hits = live.index[live["ID_Client"].astype(str) == key_id]
                    if len(hits)>0: t_idx = hits[0]
                if t_idx is None and (DOSSIER_COL in live.columns) and (init.get(DOSSIER_COL) not in [None, ""]):
                    try:
                        num = int(_to_int(pd.Series([init.get(DOSSIER_COL)])).iloc[0])
                        hits = live.index[_to_int(live[DOSSIER_COL]) == num]
                        if len(hits)>0: t_idx = hits[0]
                    except Exception:
                        pass
                if t_idx is None and (orig_pos is not None) and 0 <= int(orig_pos) < len(live):
                    t_idx = int(orig_pos)
                if t_idx is None:
                    st.error("Ligne introuvable."); st.stop()

                total = float((honoraires or 0.0)+(autres or 0.0))
                for c in [HONO, AUTRE, TOTAL, "Payé","Reste","Paiements", ESC_TR, ESC_JR,
                          "Nom","Date","Mois","Catégorie","Visa","Sous-type"] + STATUS_COLS + STATUS_DATES + [DOSSIER_COL]:
                    if c not in live.columns:
                        live[c] = 0.0 if c in [HONO,AUTRE,TOTAL,"Payé","Reste",ESC_TR] else ""
                for b in STATUS_COLS:
                    if b not in live.columns: live[b] = False

                live.at[t_idx,"Nom"]=_safe_str(nom)
                live.at[t_idx,"Date"]=str(d)
                live.at[t_idx,"Mois"]=f"{d.month:02d}"
                live.at[t_idx,"Catégorie"]= _safe_str(cat)
                live.at[t_idx,"Visa"]= _safe_str(visa)
                live.at[t_idx,"Sous-type"]= _safe_str(stype)
                live.at[t_idx, HONO]=float(honoraires or 0.0)
                live.at[t_idx, AUTRE]=float(autres or 0.0)

                # Statuts + dates
                live.at[t_idx, S_ENVOYE]   = bool(v_env);  live.at[t_idx, D_ENVOYE]   = (str(dt_env) if v_env else "")
                live.at[t_idx, S_APPROUVE] = bool(v_app);  live.at[t_idx, D_APPROUVE] = (str(dt_app) if v_app else "")
                live.at[t_idx, S_RFE]      = bool(v_rfe);  live.at[t_idx, D_RFE]      = (str(dt_rfe) if v_rfe else "")
                live.at[t_idx, S_REFUSE]   = bool(v_ref);  live.at[t_idx, D_REFUSE]   = (str(dt_ref) if v_ref else "")
                live.at[t_idx, S_ANNULE]   = bool(v_ann);  live.at[t_idx, D_ANNULE]   = (str(dt_ann) if v_ann else "")

                # Recalc payé/reste
                pay_json = live.at[t_idx,"Paiements"] if "Paiements" in live.columns else ""
                paid = _sum_payments(_parse_json_list(pay_json))
                live.at[t_idx, "Payé"]  = float(paid)
                live.at[t_idx, TOTAL]   = total
                live.at[t_idx, "Reste"] = max(total - float(paid), 0.0)

                live = ensure_dossier_numbers(live)
                write_sheet_inplace(current_path, client_target_sheet, live); save_workspace_path(current_path)
                st.success("Fiche enregistrée **dans le fichier**. ✅"); st.rerun()

            # Historique & gestion des paiements
            live_now = read_sheet(current_path, client_target_sheet, normalize=False)
            ixs = live_now.index[live_now.get("ID_Client","").astype(str)==_safe_str(init.get("ID_Client"))]
            st.markdown("#### 💳 Historique & gestion des règlements")
            if len(ixs)==0:
                st.info("Ligne introuvable pour les paiements.")
            else:
                i = ixs[0]
                if "Paiements" not in live_now.columns: live_now["Paiements"] = ""
                plist = _parse_json_list(live_now.at[i,"Paiements"])
                if plist:
                    dfp = pd.DataFrame(plist)
                    if "date" in dfp.columns: dfp["date"] = pd.to_datetime(dfp["date"], errors="coerce").dt.date.astype(str)
                    if "amount" in dfp.columns: dfp["Montant ($)"] = dfp["amount"].apply(lambda x: _fmt_money_us(float(x) if pd.notna(x) else 0.0))
                    for col in ["mode","note"]:
                        if col not in dfp.columns: dfp[col] = ""
                    show = dfp[["date","mode","Montant ($)","note"]] if set(["date","mode","note"]).issubset(dfp.columns) else dfp
                    with st.expander("Historique des règlements (cliquer pour ouvrir)", expanded=True):
                        st.table(show.rename(columns={"date":"Date","mode":"Mode","note":"Note"}))
                        if len(plist)>0:
                            del_idx = st.number_input("Supprimer la ligne n° (1..n)", min_value=1, max_value=len(plist), value=1, step=1, key=f"del_pay_idx_{i}")
                            if st.button("🗑️ Supprimer cette ligne", key=f"del_pay_btn_{i}"):
                                try:
                                    del plist[int(del_idx)-1]
                                    live_now.at[i,"Paiements"] = json.dumps(plist, ensure_ascii=False)
                                    total_paid = _sum_payments(plist)
                                    hono = _to_num(pd.Series([live_now.at[i, HONO] if HONO in live_now.columns else 0.0])).iloc[0]
                                    autr = _to_num(pd.Series([live_now.at[i, AUTRE] if AUTRE in live_now.columns else 0.0])).iloc[0]
                                    total = float(hono + autr)
                                    live_now.at[i,"Payé"]  = float(total_paid)
                                    live_now.at[i,"Reste"] = max(total - float(total_paid), 0.0)
                                    live_now.at[i,TOTAL]   = total
                                    write_sheet_inplace(current_path, client_target_sheet, live_now)
                                    st.success("Ligne supprimée et soldes recalculés. ✅"); st.rerun()
                                except Exception as e:
                                    st.error(f"Erreur suppression : {e}")
                else:
                    st.caption("Aucun paiement enregistré pour ce client.")
                cA, cB, cC, cD = st.columns([1,1,1,2])
                pay_date = cA.date_input("Date", value=date.today(), key=f"pay_date_{i}")
                pay_mode = cB.selectbox("Mode", ["CB","Chèque","Espèces","Virement","Venmo","Autre"], key=f"pay_mode_{i}")
                pay_amt  = cC.number_input("Montant ($)", min_value=0.0, step=10.0, format="%.2f", key=f"pay_amt_{i}")
                pay_note = cD.text_input("Note", "", key=f"pay_note_{i}")
                if st.button("💾 Enregistrer ce règlement (dans le fichier)", key=f"pay_add_btn_{i}"):
                    try:
                        add = float(pay_amt or 0.0)
                        if add <= 0: st.warning("Le montant doit être > 0."); st.stop()
                        norm = normalize_dataframe(live_now.copy(), visa_ref=read_visa_reference(current_path))
                        mask_id = norm["ID_Client"].astype(str) == _safe_str(init.get("ID_Client"))
                        reste_curr = float(norm.loc[mask_id, "Reste"].sum()) if mask_id.any() else 0.0
                        if add > reste_curr + 1e-9: add = reste_curr
                        plist.append({"date": str(pay_date), "amount": float(add), "mode": pay_mode, "note": pay_note})
                        live_now.at[i,"Paiements"] = json.dumps(plist, ensure_ascii=False)
                        total_paid = _sum_payments(plist)
                        hono = _to_num(pd.Series([live_now.at[i, HONO] if HONO in live_now.columns else 0.0])).iloc[0]
                        autr = _to_num(pd.Series([live_now.at[i, AUTRE] if AUTRE in live_now.columns else 0.0])).iloc[0]
                        total = float(hono + autr)
                        live_now.at[i,"Payé"]  = float(total_paid)
                        live_now.at[i,"Reste"] = max(total - float(total_paid), 0.0)
                        live_now.at[i,TOTAL]   = total
                        write_sheet_inplace(current_path, client_target_sheet, live_now)
                        st.success("Règlement ajouté. ✅"); st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")

    # --- SUPPRIMER ---
    if action == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client (écrit directement)")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            options = [(int(r["_RowID"]),
                        f'{int(r.get(DOSSIER_COL,0))} — { _safe_str(r.get("ID_Client")) } — { _safe_str(r.get("Nom")) }')
                       for _,r in live_raw.iterrows()]
            labels  = [lab for _,lab in options]
            sel_lab = st.selectbox("Sélection", labels, key="del_select")
            sel_rowid = [rid for rid,lab in options if lab==sel_lab][0]
            idx = live_raw.index[live_raw["_RowID"]==sel_rowid][0]
            st.error("⚠️ Action irréversible.")
            if st.button("Supprimer (dans le fichier)", key="del_btn"):
                live = live_raw.drop(columns=["_RowID"]).copy()
                key = _safe_str(live_raw.at[idx, "ID_Client"])
                if key and "ID_Client" in live.columns:
                    live = live[live["ID_Client"].astype(str)!=key].reset_index(drop=True)
                else:
                    nom = _safe_str(live_raw.at[idx,"Nom"]); dat = _safe_str(live_raw.at[idx,"Date"])
                    live = live[~((live.get("Nom","").astype(str)==nom)&(live.get("Date","").astype(str)==dat))].reset_index(drop=True)
                live = ensure_dossier_numbers(live)
                write_sheet_inplace(current_path, client_target_sheet, live); save_workspace_path(current_path)
                st.success("Client supprimé **dans le fichier**. ✅"); st.rerun()



# =========================
# VISA APP — PARTIE 4/5
# =========================

with tab_analyses:
    st.subheader("📊 Analyses — Volumes, Financier & Comparaisons")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa)."); st.stop()

    visa_ref_simple = read_visa_reference(current_path)
    dfA_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    dfA = normalize_dataframe(dfA_raw, visa_ref=visa_ref_simple).copy()
    if dfA.empty: st.info("Aucune donnée pour analyser."); st.stop()

    # Filtres (clés uniques anal_*)
    with st.container():
        cL, cR = st.columns([1,2])
        show_all_A = cL.checkbox("Afficher tous les dossiers", value=False, key="anal_show_all")
        cL.caption("Sélection hiérarchique (Catégorie → Visa → Sous-type)")
        with cL:
            sel_path_anal = cascading_visa_picker_tree(read_visa_reference_tree(current_path), key_prefix="anal_tree")
        visas_aut_A = visas_autorises_from_tree(read_visa_reference_tree(current_path), sel_path_anal)

        cR1, cR2, cR3 = cR.columns(3)
        yearsA  = sorted({d.year for d in dfA["Date"] if pd.notna(d)}) if "Date" in dfA.columns else []
        monthsA = [f"{m:02d}" for m in range(1,13)]
        sel_years  = cR1.multiselect("Année", yearsA, default=[], key="anal_years")
        sel_months = cR2.multiselect("Mois (MM)", monthsA, default=[], key="anal_months")
        include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True, key="anal_na")

    fA = dfA.copy()
    if not show_all_A:
        for col in ["Catégorie","Visa","Sous-type"]:
            if col in fA.columns:
                val = _safe_str(sel_path_anal.get(col, ""))
                if val:
                    fA = fA[fA[col].astype(str) == val]
        if "Visa" in fA.columns and visas_aut_A:
            fA = fA[fA["Visa"].astype(str).isin(visas_aut_A)]

    if "Date" in fA.columns and sel_years:
        mask_year = fA["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask_year |= fA["Date"].isna()
        fA = fA[mask_year]
    if "Mois" in fA.columns and sel_months:
        mask_month = fA["Mois"].isin(sel_months)
        if include_na_dates: mask_month |= fA["Mois"].isna()
        fA = fA[mask_month]

    fA["Année"] = fA["Date"].apply(lambda x: x.year if pd.notna(x) else pd.NA)
    fA["MoisNum"] = fA["Date"].apply(lambda x: int(x.month) if pd.notna(x) else pd.NA)
    fA["Periode"] = fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")

    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in fA.columns: fA[col] = pd.to_numeric(fA[col], errors="coerce").fillna(0.0)

    def derive_statut(row) -> str:
        if bool(row.get(S_APPROUVE, False)): return "Approuvé"
        if bool(row.get(S_REFUSE, False)):   return "Refusé"
        if bool(row.get(S_ANNULE, False)):   return "Annulé"
        return "En attente"
    fA["Statut"] = fA.apply(derive_statut, axis=1)

    # KPI
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(fA)}")
    k2.metric("Total (US $)", _fmt_money_us(float(fA.get(TOTAL, pd.Series(dtype=float)).sum())) )
    k3.metric("Payé (US $)", _fmt_money_us(float(fA.get("Payé", pd.Series(dtype=float)).sum())) )
    k4.metric("Solde (US $)", _fmt_money_us(float(fA.get("Reste", pd.Series(dtype=float)).sum())) )
    st.markdown('</div>', unsafe_allow_html=True)

    # Volumes créations
    st.markdown("### 📈 Volumes de créations")
    vol_crees = fA.groupby("Periode").size().reset_index(name="Créés")
    df_vol = vol_crees.rename(columns={"Créés":"Volume"}).assign(Indic="Créés")
    if not df_vol.empty:
        try:
            st.altair_chart(
                alt.Chart(df_vol).mark_line(point=True).encode(
                    x=alt.X("Periode:N", sort=None, title="Période"),
                    y=alt.Y("Volume:Q"),
                    color=alt.Color("Indic:N", legend=alt.Legend(title="")),
                    tooltip=["Periode","Indic","Volume"]
                ).properties(height=260), use_container_width=True
            )
        except Exception:
            st.dataframe(df_vol, use_container_width=True)

    st.divider()

    # Comparaisons YoY & MoM
    st.markdown("## 🔁 Comparaisons (YoY & MoM)")

    by_year = fA.dropna(subset=["Année"]).groupby("Année").agg(
        Dossiers=("Nom","count"),
        Honoraires=(HONO,"sum"),
        Autres=(AUTRE,"sum"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index().sort_values("Année")

    c1, c2 = st.columns(2)
    if not by_year.empty:
        try:
            c1.altair_chart(
                alt.Chart(by_year.melt("Année", ["Dossiers"])).mark_bar().encode(
                    x=alt.X("Année:N"), y=alt.Y("value:Q", title="Volume"),
                    color=alt.Color("variable:N", legend=None),
                    tooltip=["Année","value"]
                ).properties(title="Nombre de dossiers", height=260), use_container_width=True
            )
        except Exception:
            c1.dataframe(by_year[["Année","Dossiers"]], use_container_width=True)

        try:
            metric_vars = ["Honoraires","Autres","Total","Payé","Reste"]
            yo = by_year.melt("Année", metric_vars, var_name="Indicateur", value_name="Montant")
            c2.altair_chart(
                alt.Chart(yo).mark_bar().encode(
                    x=alt.X("Année:N"),
                    y=alt.Y("Montant:Q"),
                    color=alt.Color("Indicateur:N"),
                    tooltip=["Année","Indicateur", alt.Tooltip("Montant:Q", format="$.2f")]
                ).properties(title="Montants par année", height=260), use_container_width=True
            )
        except Exception:
            c2.dataframe(by_year.drop(columns=["Dossiers"]), use_container_width=True)

    st.markdown("### 📅 Mois (1..12) — Année sur année")
    by_year_month = fA.dropna(subset=["Année","MoisNum"]).groupby(["Année","MoisNum"]).agg(
        Dossiers=("Nom","count"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index()

    c3, c4 = st.columns(2)
    if not by_year_month.empty:
        try:
            c3.altair_chart(
                alt.Chart(by_year_month).mark_line(point=True).encode(
                    x=alt.X("MoisNum:O", title="Mois"),
                    y=alt.Y("Dossiers:Q"),
                    color=alt.Color("Année:N"),
                    tooltip=["Année","MoisNum","Dossiers"]
                ).properties(title="Dossiers par mois (YoY)", height=260), use_container_width=True
            )
        except Exception:
            c3.dataframe(by_year_month.pivot(index="MoisNum", columns="Année", values="Dossiers"), use_container_width=True)

        try:
            c4.altair_chart(
                alt.Chart(by_year_month.melt(["Année","MoisNum"], ["Total","Payé","Reste"],
                                             var_name="Indicateur", value_name="Montant")
                ).mark_line(point=True).encode(
                    x=alt.X("MoisNum:O", title="Mois"),
                    y=alt.Y("Montant:Q"),
                    color=alt.Color("Année:N"),
                    tooltip=["Année","MoisNum","Indicateur", alt.Tooltip("Montant:Q", format="$.2f")]
                ).properties(title="Montants par mois (YoY)", height=260),
                use_container_width=True
            )
        except Exception:
            c4.dataframe(by_year_month.pivot_table(index="MoisNum", columns="Année", values="Total"), use_container_width=True)

    st.markdown("### 🛂 Par type de visa — Année sur année")
    topN = st.slider("Top N visas (par Total)", 3, 20, 10, 1, key="cmp_topn")
    metric_cmp = st.selectbox("Indicateur", ["Dossiers","Total","Payé","Reste","Honoraires","Autres"], index=1, key="cmp_metric")

    by_year_visa = fA.dropna(subset=["Année"]).groupby(["Année","Visa"]).agg(
        Dossiers=("Nom","count"),
        Honoraires=(HONO,"sum"),
        Autres=(AUTRE,"sum"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index()

    top_visas = (by_year_visa.groupby("Visa")["Total"].sum()
                 .sort_values(ascending=False).head(topN).index.tolist())
    by_year_visa_top = by_year_visa[by_year_visa["Visa"].isin(top_visas)].copy()

    if not by_year_visa_top.empty:
        try:
            st.altair_chart(
                alt.Chart(by_year_visa_top).mark_bar().encode(
                    x=alt.X("Visa:N", sort=top_visas),
                    y=alt.Y(f"{metric_cmp}:Q"),
                    color=alt.Color("Année:N"),
                    tooltip=["Visa","Année", alt.Tooltip(f"{metric_cmp}:Q", format="$.2f" if metric_cmp!="Dossiers" else "")],
                ).properties(height=300), use_container_width=True
            )
        except Exception:
            st.dataframe(by_year_visa_top.pivot_table(index="Visa", columns="Année", values=metric_cmp, aggfunc="sum"),
                         use_container_width=True)

    st.divider()
    st.markdown("### 🔎 Détails (clients)")
    details_cols = [c for c in ["Periode",DOSSIER_COL,"ID_Client","Nom",
                                "Catégorie","Visa","Sous-type",
                                "Date", HONO, AUTRE, TOTAL, "Payé","Reste","Statut","Année","MoisNum"] if c in fA.columns]
    details = fA[details_cols].copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in details.columns: details[col] = details[col].apply(lambda x: _fmt_money_us(x) if pd.notna(x) else "")
    st.dataframe(details.sort_values(["Année","MoisNum","Catégorie","Nom"]), use_container_width=True)




# =========================
# VISA APP — PARTIE 5/5
# =========================

with tab_escrow:
    st.subheader("🏦 ESCROW — dépôts sur honoraires & transferts")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa)."); st.stop()
    live_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    live = normalize_dataframe(live_raw, visa_ref=read_visa_reference(current_path)).copy()
    if ESC_TR not in live.columns: live[ESC_TR] = 0.0
    else: live[ESC_TR] = pd.to_numeric(live[ESC_TR], errors="coerce").fillna(0.0)
    live["ESCROW dispo"] = live.apply(lambda r: float(max(min(float(r.get("Payé",0.0)), float(r.get(HONO,0.0))) - float(r.get(ESC_TR,0.0)), 0.0)), axis=1)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Dossiers", f"{len(live)}")
    c2.metric("ESCROW total dispo", _fmt_money_us(float(live["ESCROW dispo"].sum())))
    envoyes = live[(live[S_ENVOYE]==True)]
    a_transferer = envoyes[envoyes["ESCROW dispo"]>0.004].reset_index(drop=True)
    c3.metric("Dossiers envoyés (à réclamer)", f"{len(a_transferer)}")
    c4.metric("Montant à réclamer", _fmt_money_us(float(a_transferer["ESCROW dispo"].sum())))

    st.divider()
    st.markdown("### 📌 À transférer (dossiers **envoyés**)")
    if a_transferer.empty:
        st.success("Aucun transfert en attente pour des dossiers envoyés.")
    else:
        for i, r in a_transferer.sort_values("Date").reset_index(drop=True).iterrows():
            rid = f"{_safe_str(r.get('ID_Client'))}_{int(r.get(DOSSIER_COL,0))}_{i}"
            with st.expander(f'🔔 {r.get(DOSSIER_COL,"")} — {r.get("ID_Client","")} — {r.get("Nom","")} — {r.get("Catégorie","")} / {r.get("Visa","")} — ESCROW dispo: {_fmt_money_us(r["ESCROW dispo"])}'):
                cA, cB, cC = st.columns(3)
                cA.metric("Honoraires", _fmt_money_us(float(r.get(HONO,0.0))))
                cB.metric("Déjà transféré", _fmt_money_us(float(r.get(ESC_TR,0.0))))
                cC.metric("Payé", _fmt_money_us(float(r.get("Payé",0.0))))
                amt = st.number_input("Montant à marquer comme transféré (US $)",
                                      min_value=0.0,
                                      value=float(r["ESCROW dispo"]),
                                      step=10.0, format="%.2f",
                                      key=f"esc_amt_{rid}")
                note = st.text_input("Note (facultatif)", "", key=f"esc_note_{rid}")
                if st.button("✅ Marquer transféré (écrit dans le fichier)", key=f"esc_btn_{rid}"):
                    try:
                        live_w = read_sheet(current_path, client_target_sheet, normalize=False).copy()
                        for c in [ESC_TR, ESC_JR]:
                            if c not in live_w.columns: live_w[c] = 0.0 if c==ESC_TR else ""
                        # retrouver la ligne par ID_Client (fallback par Dossier N)
                        idxs = live_w.index[live_w.get("ID_Client","").astype(str)==str(r.get("ID_Client",""))]
                        if len(idxs)==0 and DOSSIER_COL in live_w.columns:
                            idxs = live_w.index[_to_int(live_w[DOSSIER_COL]) == int(_to_int(pd.Series([r.get(DOSSIER_COL,0)])).iloc[0])]
                        if len(idxs)==0: st.error("Ligne introuvable."); st.stop()
                        row_i = idxs[0]

                        # dispo recalculée
                        tmp = normalize_dataframe(live_w.copy(), visa_ref=read_visa_reference(current_path))
                        disp = float(tmp.loc[tmp["ID_Client"].astype(str)==str(r.get("ID_Client","")), :].apply(
                            lambda rr: float(max(min(float(rr.get("Payé",0.0)), float(rr.get(HONO,0.0))) - float(rr.get(ESC_TR,0.0)), 0.0)),
                            axis=1).iloc[0])
                        add = float(min(max(amt,0.0), disp))
                        live_w.at[row_i, ESC_TR] = float(pd.to_numeric(pd.Series([live_w.at[row_i, ESC_TR]]), errors="coerce").fillna(0.0).iloc[0] + add)
                        # Journal
                        lst = _parse_json_list(live_w.at[row_i, ESC_JR])
                        lst.append({"ts": datetime.now().isoformat(timespec="seconds"), "amount": float(add), "note": _safe_str(note)})
                        live_w.at[row_i, ESC_JR] = json.dumps(lst, ensure_ascii=False)
                        live_w = ensure_dossier_numbers(live_w)
                        write_sheet_inplace(current_path, client_target_sheet, live_w)
                        st.success("Transfert ESCROW enregistré **dans le fichier**. ✅"); st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")

    st.divider()
    st.markdown("### 📥 En cours d’alimentation (dossiers **non envoyés**)")
    non_env = live[(live[S_ENVOYE]!=True) & (live["ESCROW dispo"]>0.004)].copy()
    if non_env.empty:
        st.info("Rien en attente côté dossiers non envoyés.")
    else:
        show = non_env[[DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Date",HONO,"Payé",ESC_TR,"ESCROW dispo"]].copy()
        for col in [HONO,"Payé",ESC_TR,"ESCROW dispo"]:
            show[col] = show[col].map(_fmt_money_us)
        st.dataframe(show, use_container_width=True)

    st.divider()
    st.markdown("### 🧾 Historique des transferts (journal)")
    has_journal = live[live[ESC_JR].astype(str).str.len() > 0]
    if has_journal.empty:
        st.caption("Aucun journal de transfert pour le moment.")
    else:
        rows = []
        for _, rr in has_journal.iterrows():
            entries = _parse_json_list(rr.get(ESC_JR, ""))
            for e in entries:
                rows.append({
                    DOSSIER_COL: rr.get(DOSSIER_COL, ""),
                    "ID_Client": rr.get("ID_Client", ""),
                    "Nom": rr.get("Nom", ""),
                    "Visa": rr.get("Visa", ""),
                    "Date": rr.get("Date", ""),
                    "Horodatage": e.get("ts", ""),
                    "Montant (US $)": float(e.get("amount", 0.0)),
                    "Note": e.get("note", "")
                })
        jdf = pd.DataFrame(rows)
        if not jdf.empty:
            jdf["Montant (US $)"] = jdf["Montant (US $)"].apply(lambda x: _fmt_money_us(float(x) if pd.notna(x) else 0.0))
            if "Horodatage" in jdf.columns:
                try:
                    jdf = jdf.sort_values("Horodatage")
                except Exception:
                    pass
        st.dataframe(jdf, use_container_width=True)
