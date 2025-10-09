# =============================================
# 🧭 PARTIE 1/5 — INITIALISATION & CHARGEMENT
# =============================================
from __future__ import annotations
import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime, date
import json, re, unicodedata

st.set_page_config(
    page_title="📑 Gestion des Visas — Villa Tobias",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================
# 📁 FONCTIONS DE BASE
# =============================================

def _safe_str(x):
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip()
    except Exception:
        return ""

def _norm_txt(x: str) -> str:
    s = _safe_str(x)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s*[/\-]\s*", " ", s)
    s = re.sub(r"[^a-zA-Z0-9\s]+", " ", s)
    return " ".join(s.lower().split())

def _to_num(s):
    if isinstance(s, pd.Series):
        s = s.astype(str).str.replace(r"[^\d,.\-]", "", regex=True)
        def _one(x):
            if x == "" or x == "-":
                return 0.0
            if x.count(",") == 1 and x.count(".") == 0:
                x = x.replace(",", ".")
            if x.count(".") == 1 and x.count(",") >= 1:
                x = x.replace(",", "")
            try:
                return float(x)
            except:
                return 0.0
        return s.map(_one)
    return 0.0

# =============================================
# ⚙️ CHARGEMENT DES FICHIERS
# =============================================

@st.cache_data(show_spinner=False)
def load_excel(path: str | Path, sheet_candidates=("Clients", "Clients_normalises")) -> pd.DataFrame:
    """Lecture d’un fichier Excel de clients."""
    path = Path(path)
    if not path.exists():
        return pd.DataFrame()
    xls = pd.ExcelFile(path)
    for sn in sheet_candidates:
        if sn in xls.sheet_names:
            df = pd.read_excel(path, sheet_name=sn)
            return df
    # fallback: premier onglet
    return pd.read_excel(path, sheet_name=xls.sheet_names[0])


@st.cache_data(show_spinner=False)
def load_visa_structure(path: str | Path, sheet_candidates=("Visa", "Visa_normalise")) -> pd.DataFrame:
    """Lecture d’un référentiel Visa (Catégorie + Sous-catégories)."""
    path = Path(path)
    if not path.exists():
        return pd.DataFrame()
    xls = pd.ExcelFile(path)
    for sn in sheet_candidates:
        if sn in xls.sheet_names:
            df = pd.read_excel(path, sheet_name=sn)
            return df
    return pd.read_excel(path, sheet_name=xls.sheet_names[0])


# =============================================
# 🎛️ INTERFACE DE CHOIX DE FICHIERS
# =============================================
st.sidebar.header("📂 Fichiers")

# --- Chargement du fichier principal Clients ---
clients_file = st.sidebar.file_uploader("Fichier Clients", type=["xlsx"], key="up_clients")
default_clients_path = st.session_state.get("clients_last_path", "")
clients_path_text = st.sidebar.text_input("Chemin Clients.xlsx", value=default_clients_path)

clients_path: Path | None = None
if clients_file is not None:
    clients_path = Path(clients_file.name).resolve()
    clients_path.write_bytes(clients_file.getvalue())
    st.session_state["clients_last_path"] = str(clients_path)
elif clients_path_text:
    p = Path(clients_path_text)
    clients_path = p if p.exists() else None

# --- Chargement du référentiel Visa ---
visa_file = st.sidebar.file_uploader("Référentiel Visa", type=["xlsx"], key="up_visa")
default_visa_path = st.session_state.get("visa_last_path", "")
visa_path_text = st.sidebar.text_input("Chemin Visa.xlsx", value=default_visa_path)

visa_path: Path | None = None
if visa_file is not None:
    visa_path = Path(visa_file.name).resolve()
    visa_path.write_bytes(visa_file.getvalue())
    st.session_state["visa_last_path"] = str(visa_path)
elif visa_path_text:
    p = Path(visa_path_text)
    visa_path = p if p.exists() else None

# --- Si aucun fichier n'est encore sélectionné ---
if clients_path is None or not clients_path.exists():
    st.warning("🟡 Veuillez d’abord choisir ou indiquer le fichier **Clients.xlsx**.")
    st.stop()

if visa_path is None or not visa_path.exists():
    st.warning("🟡 Veuillez d’abord choisir ou indiquer le fichier **Visa.xlsx**.")
    st.stop()

# =============================================
# 🧾 CHARGEMENT EFFECTIF DES DONNÉES
# =============================================
df_clients = load_excel(clients_path)
df_visa = load_visa_structure(visa_path)

if df_clients.empty:
    st.error("Le fichier Clients est vide ou illisible.")
    st.stop()

if df_visa.empty:
    st.error("Le fichier Visa est vide ou illisible.")
    st.stop()

st.sidebar.success(f"✅ {len(df_clients)} clients chargés.")
st.sidebar.success(f"✅ {len(df_visa)} catégories Visa chargées.")

# =============================================
# 🧮 NORMALISATION DE BASE
# =============================================
for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Reste"]:
    if c in df_clients.columns:
        df_clients[c] = _to_num(df_clients[c])
if "Total (US $)" in df_clients.columns:
    df_clients["Total (US $)"] = df_clients["Montant honoraires (US $)"] + df_clients["Autres frais (US $)"]
else:
    df_clients["Total (US $)"] = df_clients["Montant honoraires (US $)"] + df_clients["Autres frais (US $)"]

df_clients["Reste"] = (df_clients["Total (US $)"] - df_clients["Payé"]).clip(lower=0.0)

# =============================================
# ✅ RÉSULTAT DE CHARGEMENT
# =============================================
st.success(f"Fichier **{clients_path.name}** et référentiel **{visa_path.name}** chargés avec succès.")
st.write("Aperçu des 5 premiers clients :")
st.dataframe(df_clients.head())


# =========================
# VISA APP — PARTIE 2/5
# =========================

st.set_page_config(page_title="Visa Manager — US $", layout="wide")
st.title("🛂 Visa Manager — US $")

# --- Barre latérale : chargement des fichiers ---
st.sidebar.header("📁 Fichiers")
last_clients, last_visa = _load_last_paths()

# Clients
up_clients = st.sidebar.file_uploader("Classeur Clients (.xlsx)", type=["xlsx"], key="up_clients")
if up_clients is not None:
    buf = up_clients.getvalue()
    cpath = Path(up_clients.name).resolve()
    cpath.write_bytes(buf)
    _save_last_paths(clients=cpath)

clients_path = st.sidebar.text_input("Chemin Clients", value=str(last_clients) if last_clients else "")
clients_path = Path(clients_path) if clients_path else None

# Visa.xlsx
up_visa = st.sidebar.file_uploader("Référentiel Visa.xlsx (onglet 'Visa')", type=["xlsx"], key="up_visa")
if up_visa is not None:
    buf = up_visa.getvalue()
    vpath = Path(up_visa.name).resolve()
    vpath.write_bytes(buf)
    _save_last_paths(visa=vpath)

visa_path = st.sidebar.text_input("Chemin Visa.xlsx", value=str(last_visa) if last_visa else "")
visa_path = Path(visa_path) if visa_path else None

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Recharger", use_container_width=True):
    st.rerun()

# --- Contrôles ---
if not clients_path or not clients_path.exists():
    st.warning("Charge un **classeur Clients** (.xlsx)."); st.stop()
if not visa_path or not visa_path.exists():
    st.warning("Charge le **référentiel Visa.xlsx** (onglet 'Visa')."); st.stop()

# --- Feuille Clients à utiliser ---
sheets = list_sheets(clients_path)
if not sheets:
    st.error("Impossible de lire le classeur Clients."); st.stop()

# Détection d'une feuille "clients"
cand = None
for sn in sheets:
    df0 = read_sheet(clients_path, sn)
    if {"Nom","Visa"}.issubset(set(df0.columns.astype(str))):
        cand = sn; break

sheet_choice = st.sidebar.selectbox("Feuille Clients :", sheets, index=(sheets.index(cand) if cand in sheets else 0), key="sheet_choice")

# --- Lecture données ---
df_clients_raw = read_sheet(clients_path, sheet_choice, normalize=False)
df_clients     = read_sheet(clients_path, sheet_choice, normalize=True)

df_visa = read_visa_matrix(visa_path)
if df_visa.empty:
    st.error("Onglet 'Visa' introuvable ou vide dans Visa.xlsx."); st.stop()

# --- Onglets ---
tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"])

# ================= DASHBOARD =================
with tab_dash:
    st.subheader("📊 Dashboard")
    st.caption("Filtres contextuels basés sur Visa.xlsx : coche des catégories, puis leurs sous-catégories s’affichent.")

    # Filtres contextuels (cases ; passe as_toggle=True si tu veux des bascules)
    sel = build_checkbox_filters_grouped(df_visa, keyprefix=f"flt_dash_{sheet_choice}", as_toggle=False)

    # Filtrage
    f = filter_clients_by_ref(df_clients, sel)

    # Filtres date simples en plus (Année/Mois)
    cR1, cR2, cR3 = st.columns(3)
    years  = sorted({d.year for d in f["Date"] if pd.notna(d)}) if "Date" in f.columns else []
    months = sorted([m for m in f["Mois"].dropna().unique()]) if "Mois" in f.columns else []
    sel_years  = cR1.multiselect("Année", years, default=[], key=f"dash_years_{sheet_choice}")
    sel_months = cR2.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{sheet_choice}")
    include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True, key=f"dash_na_{sheet_choice}")

    if "Date" in f.columns and sel_years:
        mask = f["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask |= f["Date"].isna()
        f = f[mask]
    if "Mois" in f.columns and sel_months:
        mask = f["Mois"].isin(sel_months)
        if include_na_dates: mask |= f["Mois"].isna()
        f = f[mask]

    hidden = len(df_clients) - len(f)
    if hidden > 0:
        st.caption(f"🔎 {hidden} ligne(s) masquée(s) par les filtres.")

    # KPI compacts
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Honoraires", _fmt_money_us(float(f.get(HONO, pd.Series(dtype=float)).sum())))
    k3.metric("Payé", _fmt_money_us(float(f.get("Payé", pd.Series(dtype=float)).sum())))
    k4.metric("Solde", _fmt_money_us(float(f.get("Reste", pd.Series(dtype=float)).sum())))
    st.markdown('</div>', unsafe_allow_html=True)

    # Tableau
    st.divider()
    st.subheader("📋 Dossiers filtrés")
    cols_show = [c for c in [DOSSIER_COL,"ID_Client","Nom","Date","Mois","Catégorie","Visa",HONO,AUTRE,TOTAL,"Payé","Reste",
                             S_ENVOYE,D_ENVOYE,S_APPROUVE,D_APPROUVE,S_RFE,D_RFE,S_REFUSE,D_REFUSE,S_ANNULE,D_ANNULE] if c in f.columns]
    view = f.copy()
    for col in [HONO,AUTRE,TOTAL,"Payé","Reste"]:
        if col in view.columns: view[col] = pd.to_numeric(view[col], errors="coerce").fillna(0.0).map(_fmt_money_us)
    if "Date" in view.columns: view["Date"] = view["Date"].astype(str)
    st.dataframe(view[cols_show], use_container_width=True)

# =========================
# VISA APP — PARTIE 3/5
# =========================

# --- helpers locaux si absents ---
if 'next_dossier_number' not in globals():
    def next_dossier_number(df: pd.DataFrame) -> int:
        if df is None or df.empty or DOSSIER_COL not in df.columns:
            return 13057
        try:
            nums = pd.to_numeric(df[DOSSIER_COL], errors="coerce")
            m = int(nums.max()) if nums.notna().any() else 13056
        except Exception:
            m = 13056
        return max(m, 13056) + 1

if '_make_client_id_from_row' not in globals():
    def _make_client_id_from_row(row: dict) -> str:
        nom = _safe_str(row.get("Nom"))
        d = row.get("Date")
        try:
            d = pd.to_datetime(d).date()
        except Exception:
            d = date.today()
        base = f"{nom}-{d.strftime('%Y%m%d')}"
        base = re.sub(r"[^A-Za-z0-9\-]+", "", base.replace(" ", "-"))
        return base.lower()

# --- onglets si non créés ---
if 'tab_clients' not in globals():
    tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"])

with tab_clients:
    st.subheader("👥 Clients — créer / modifier / supprimer / paiements")

    # sécurité: chemins posés en PARTIE 2
    if 'clients_path' not in globals() or clients_path is None or not Path(clients_path).exists():
        st.info("Charge d’abord le **classeur Clients** (barre latérale).")
        st.stop()

    # feuille
    if 'sheet_choice' not in globals() or not sheet_choice:
        sheets = list_sheets(clients_path)
        sheet_choice = sheets[0] if sheets else None
    if sheet_choice is None:
        st.error("Aucune feuille valide dans le classeur."); st.stop()

    live_raw = read_sheet(clients_path, sheet_name=sheet_choice) if 'read_sheet' in globals() else pd.read_excel(clients_path, sheet_name=sheet_choice)
    live = normalize_clients(live_raw)

    # --- Sélecteur client existant ---
    cL, cR = st.columns([1,1])
    with cL:
        st.markdown("### 🔎 Sélection")
        if live.empty:
            st.caption("Aucun client pour le moment.")
            sel_idx = None
            sel_row = None
        else:
            labels = (live["Nom"].fillna("").astype(str) + " — " + live.get("ID_Client","").astype(str))
            sel_idx = st.selectbox("Client", options=list(live.index), format_func=lambda i: labels.iloc[i], key=f"cli_sel_{sheet_choice}")
            sel_row = live.loc[sel_idx] if sel_idx is not None else None

    # --- Création nouveau client ---
    with cR:
        st.markdown("### ➕ Nouveau client")
        new_name = st.text_input("Nom", key=f"new_nom_{sheet_choice}")
        new_date = st.date_input("Date création", value=date.today(), key=f"new_date_{sheet_choice}")

        # Visa via code (catégories du référentiel)
        if 'df_visa' in globals() and not df_visa.empty:
            codes = sorted(df_visa["VisaCode"].dropna().unique().tolist())
        else:
            codes = sorted(live["Visa"].dropna().astype(str).unique().tolist())
        new_visa = st.selectbox("Visa (code)", options=[""]+codes, index=0, key=f"new_visa_{sheet_choice}")

        new_hono = st.number_input(HONO, min_value=0.0, step=10.0, format="%.2f", key=f"new_hono_{sheet_choice}")
        new_autr = st.number_input(AUTRE, min_value=0.0, step=10.0, format="%.2f", key=f"new_autr_{sheet_choice}")

        if st.button("💾 Créer", key=f"btn_new_{sheet_choice}"):
            if not new_name:
                st.warning("Renseigne le **Nom**.")
            elif not new_visa:
                st.warning("Choisis un **Visa**.")
            else:
                base_raw = read_sheet(clients_path, sheet_choice).copy()
                base_norm = normalize_clients(base_raw)

                dossier = next_dossier_number(base_norm)
                client_id = _make_client_id_from_row({"Nom": new_name, "Date": new_date})
                # éviter collision ID_Client
                origin = client_id; i = 0
                while "ID_Client" in base_norm.columns and (base_norm["ID_Client"].astype(str) == client_id).any():
                    i += 1; client_id = f"{origin}-{i}"

                new_row = {
                    DOSSIER_COL: dossier,
                    "ID_Client": client_id,
                    "Nom": new_name,
                    "Date": pd.to_datetime(new_date).date(),
                    "Mois": f"{new_date.month:02d}",
                    "Catégorie": new_visa,  # si tu veux Catégorie distincte, remplace ici
                    "Visa": _visa_code_only(new_visa),
                    HONO: float(new_hono),
                    AUTRE: float(new_autr),
                    TOTAL: float(new_hono) + float(new_autr),
                    "Payé": 0.0,
                    "Reste": float(new_hono) + float(new_autr),
                    PAY_JSON: "[]"
                }

                # append et écrire
                base_raw = pd.concat([base_raw, pd.DataFrame([new_row])], ignore_index=True)
                base_raw = normalize_clients(base_raw)
                write_sheet_inplace(clients_path, sheet_choice, base_raw)
                st.success("Client créé.")
                st.rerun()

    st.markdown("---")

    if sel_row is None:
        st.info("Sélectionne un client à gauche, ou crée un nouveau client.")
        st.stop()

    # --- Formulaire édition ---
    idx = sel_idx
    ed = sel_row.to_dict()

    e1,e2,e3 = st.columns(3)
    with e1:
        ed_nom = st.text_input("Nom", value=_safe_str(ed.get("Nom","")), key=f"ed_nom_{idx}_{sheet_choice}")
        ed_date = st.date_input("Date création", value=(pd.to_datetime(ed.get("Date")).date() if pd.notna(ed.get("Date")) else date.today()),
                                key=f"ed_date_{idx}_{sheet_choice}")
    with e2:
        # choisir visa code
        codes_all = sorted(df_visa["VisaCode"].dropna().unique().tolist()) if 'df_visa' in globals() and not df_visa.empty else sorted(live["Visa"].dropna().astype(str).unique().tolist())
        current_code = _visa_code_only(ed.get("Visa",""))
        ed_visa = st.selectbox("Visa (code)", options=[""]+codes_all, index=(codes_all.index(current_code)+1 if current_code in codes_all else 0),
                               key=f"ed_visa_{idx}_{sheet_choice}")
    with e3:
        ed_hono = st.number_input(HONO, min_value=0.0, value=float(ed.get(HONO,0.0)), step=10.0, format="%.2f", key=f"ed_hono_{idx}_{sheet_choice}")
        ed_autr = st.number_input(AUTRE, min_value=0.0, value=float(ed.get(AUTRE,0.0)), step=10.0, format="%.2f", key=f"ed_autr_{idx}_{sheet_choice}")

    # --- Paiements ---
    st.markdown("#### 💳 Paiements (acomptes multiples)")
    p1,p2,p3,p4 = st.columns([1,1,1,2])
    with p1:
        p_date = st.date_input("Date paiement", value=date.today(), key=f"p_date_{idx}_{sheet_choice}")
    with p2:
        p_mode = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"p_mode_{idx}_{sheet_choice}")
    with p3:
        p_amt  = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"p_amt_{idx}_{sheet_choice}")
    with p4:
        if st.button("➕ Ajouter paiement", key=f"btn_addpay_{idx}_{sheet_choice}"):
            base_raw = read_sheet(clients_path, sheet_choice).copy()
            base_norm = normalize_clients(base_raw)
            reste_curr = float(base_norm.loc[idx, "Reste"])
            if float(p_amt) <= 0:
                st.warning("Le montant doit être > 0.")
            elif reste_curr <= 0:
                st.info("Dossier déjà soldé.")
            else:
                row = base_raw.loc[idx].to_dict()
                try:
                    plist = json.loads(row.get(PAY_JSON,"[]"))
                    if not isinstance(plist, list): plist=[]
                except Exception:
                    plist = []
                plist.append({"date": str(p_date), "mode": p_mode, "amount": float(p_amt)})
                row[PAY_JSON] = json.dumps(plist, ensure_ascii=False)
                base_raw.loc[idx] = row
                base_raw = normalize_clients(base_raw)
                write_sheet_inplace(clients_path, sheet_choice, base_raw)
                st.success("Paiement ajouté.")
                st.rerun()

    # Historique paiements
    try:
        plist = json.loads(live_raw.loc[idx].get(PAY_JSON,"[]"))
        if not isinstance(plist, list): plist=[]
    except Exception:
        plist = []
    st.write("**Historique des paiements**")
    if plist:
        h = pd.DataFrame(plist)
        if "amount" in h.columns: h["amount"] = h["amount"].map(_fmt_money_us)
        st.dataframe(h, use_container_width=True)
    else:
        st.caption("Aucun paiement saisi.")

    # --- Boutons actions ---
    a1,a2 = st.columns([1,1])
    if a1.button("💾 Sauvegarder les modifications", key=f"btn_save_{idx}_{sheet_choice}"):
        base_raw = read_sheet(clients_path, sheet_choice).copy()
        if idx >= len(base_raw):
            st.error("Ligne introuvable."); st.stop()
        row = base_raw.loc[idx].to_dict()
        row["Nom"]  = ed_nom
        row["Date"] = pd.to_datetime(ed_date).date()
        row["Mois"] = f"{ed_date.month:02d}"
        if ed_visa: row["Visa"] = _visa_code_only(ed_visa)
        row[HONO] = float(ed_hono)
        row[AUTRE]= float(ed_autr)
        row[TOTAL]= float(ed_hono) + float(ed_autr)
        base_raw.loc[idx] = row
        base_raw = normalize_clients(base_raw)
        write_sheet_inplace(clients_path, sheet_choice, base_raw)
        st.success("Modifications sauvegardées.")
        st.rerun()

    if a2.button("🗑️ Supprimer ce client", key=f"btn_del_{idx}_{sheet_choice}"):
        base_raw = read_sheet(clients_path, sheet_choice).copy()
        if 0 <= idx < len(base_raw):
            base_raw = base_raw.drop(index=idx).reset_index(drop=True)
            base_raw = normalize_clients(base_raw)
            write_sheet_inplace(clients_path, sheet_choice, base_raw)
            st.success("Client supprimé.")
            st.rerun()
        else:
            st.error("Ligne introuvable.")


# =========================
# VISA APP — PARTIE 4/5
# =========================
try:
    import altair as alt
except Exception:
    alt = None

if 'tab_analyses' not in globals():
    tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"])

with tab_analyses:
    st.subheader("📊 Analyses — Volumes & Financier")
    if 'clients_path' not in globals() or clients_path is None or not Path(clients_path).exists():
        st.info("Charge d’abord le **classeur Clients**."); st.stop()

    dfA = normalize_clients(read_sheet(clients_path, sheet_choice))
    if dfA.empty:
        st.info("Aucune donnée à analyser."); st.stop()

    # Filtres contextuels identiques au Dashboard (cases / bascules par catégorie)
    if 'df_visa' in globals() and not df_visa.empty:
        selA = build_checkbox_filters_grouped(df_visa, keyprefix=f"anal_{sheet_choice}", as_toggle=False)
        fA = filter_clients_by_ref(dfA, selA)
    else:
        selA = {"__whitelist_visa__": []}
        fA = dfA.copy()

    # Enrichissements
    fA["Année"] = fA["Date"].apply(lambda x: x.year if pd.notna(x) else pd.NA)
    fA["MoisNum"] = fA["Date"].apply(lambda x: int(x.month) if pd.notna(x) else pd.NA)
    fA["Periode"] = fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")

    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        if c in fA.columns: fA[c] = pd.to_numeric(fA[c], errors="coerce").fillna(0.0)

    # KPI compacts
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(fA)}")
    k2.metric("Honoraires", _fmt_money_us(float(fA[HONO].sum())))
    k3.metric("Payé", _fmt_money_us(float(fA["Payé"].sum())))
    k4.metric("Solde", _fmt_money_us(float(fA["Reste"].sum())))
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    st.markdown("### 📈 Volumes de créations (par période)")
    vol = fA.groupby("Periode").size().reset_index(name="Créés")
    if alt is not None and not vol.empty:
        try:
            ch = alt.Chart(vol).mark_line(point=True).encode(
                x=alt.X("Periode:N", sort=None), y="Créés:Q", tooltip=["Periode","Créés"]
            ).properties(height=260)
            st.altair_chart(ch, use_container_width=True)
        except Exception:
            st.dataframe(vol, use_container_width=True)
    else:
        st.dataframe(vol, use_container_width=True)

    st.divider()
    st.markdown("### 🔁 Comparaisons année / mois")
    by_year = fA.dropna(subset=["Année"]).groupby("Année").agg(
        Dossiers=("Nom","count"),
        Honoraires=(HONO,"sum"),
        Autres=(AUTRE,"sum"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index().sort_values("Année")
    c1,c2 = st.columns(2)
    c1.dataframe(by_year, use_container_width=True)

    by_month = fA.dropna(subset=["MoisNum"]).groupby("MoisNum").agg(
        Dossiers=("Nom","count"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index().sort_values("MoisNum")
    c2.dataframe(by_month, use_container_width=True)

    st.divider()
    st.markdown("### 🔎 Détails (clients)")
    show_cols = [c for c in ["Periode",DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Date",HONO,AUTRE,TOTAL,"Payé","Reste"] if c in fA.columns]
    vf = fA.copy()
    for c in [HONO,AUTRE,TOTAL,"Payé","Reste"]:
        if c in vf.columns: vf[c] = vf[c].apply(lambda x: _fmt_money_us(x) if pd.notna(x) else "")
    if "Date" in vf.columns: vf["Date"] = vf["Date"].astype(str)
    st.dataframe(vf[show_cols].sort_values(["Année","MoisNum","Catégorie","Nom"]), use_container_width=True)


# =========================
# VISA APP — PARTIE 5/5
# =========================

# constantes si absentes
if 'ESC_TR' not in globals(): ESC_TR = "ESCROW transféré (US $)"
if 'ESC_JR' not in globals(): ESC_JR = "Journal ESCROW"
for _c in [ESC_TR, ESC_JR]:
    if _c not in normalize_clients(pd.DataFrame()).columns:
        pass  # normalize_clients ajoute par défaut si besoin, sinon on gèrera dynamiquement

if 'tab_escrow' not in globals():
    tab_dash, tab_clients, tab_analyses, tab_escrow = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"])

with tab_escrow:
    st.subheader("🏦 ESCROW — suivi & transferts")

    if 'clients_path' not in globals() or clients_path is None or not Path(clients_path).exists():
        st.info("Charge d’abord le **classeur Clients**."); st.stop()

    base_raw = read_sheet(clients_path, sheet_choice)
    dfE = normalize_clients(base_raw.copy())
    if dfE.empty:
        st.info("Aucun dossier."); st.stop()

    # Ajoute colonnes manquantes côté RAW si besoin
    for col in [ESC_TR, ESC_JR]:
        if col not in base_raw.columns:
            base_raw[col] = "" if col==ESC_JR else 0.0

    # disponible ESCROW = min(Payé, Honoraires) - déjà transféré
    tr_vals = pd.to_numeric(dfE.get(ESC_TR, 0.0), errors="coerce").fillna(0.0)
    dfE["Dispo ESCROW"] = (dfE["Payé"].clip(upper=dfE[HONO]) - tr_vals).clip(lower=0.0)

    # Alerte : dossiers "envoyés" => ici on se base juste sur Dispo>0 pour simplifier
    to_claim = dfE[dfE["Dispo ESCROW"] > 0.0]
    if len(to_claim):
        tmp = to_claim[[c for c in [DOSSIER_COL,"ID_Client","Nom","Visa",HONO,"Payé","Dispo ESCROW"] if c in to_claim.columns]].copy()
        for col in [HONO,"Payé","Dispo ESCROW"]:
            if col in tmp.columns: tmp[col] = pd.to_numeric(tmp[col], errors="coerce").fillna(0.0).map(_fmt_money_us)
        st.warning(f"⚠️ {len(to_claim)} dossier(s) ont de l’ESCROW disponible.")
        st.dataframe(tmp, use_container_width=True)

    st.divider()
    st.markdown("### 🔁 Marquer un transfert d’ESCROW → Compte ordinaire")

    df_with_dispo = dfE[dfE["Dispo ESCROW"] > 0.0].reset_index(drop=True)
    if df_with_dispo.empty:
        st.caption("Aucun dossier avec ESCROW disponible.")
    else:
        for i, r in df_with_dispo.iterrows():
            dispo = float(r["Dispo ESCROW"])
            header = f"{r.get(DOSSIER_COL,'')} — {r.get('Nom','')} — Visa {r.get('Visa','')} — Dispo: {_fmt_money_us(dispo)}"
            with st.expander(header, expanded=False):
                amt = st.number_input("Montant à transférer (US $)", min_value=0.0, value=float(dispo), step=10.0, format="%.2f",
                                      key=f"esc_amt_{sheet_choice}_{i}")
                note = st.text_input("Note (optionnelle)", key=f"esc_note_{sheet_choice}_{i}")
                if st.button("💾 Enregistrer le transfert", key=f"esc_save_{sheet_choice}_{i}"):
                    # on identifie la ligne par ID_Client si possible
                    idc = _safe_str(r.get("ID_Client",""))
                    if idc and "ID_Client" in base_raw.columns:
                        try:
                            real_idx = base_raw.index[base_raw["ID_Client"].astype(str) == idc][0]
                        except Exception:
                            real_idx = None
                    else:
                        real_idx = int(r.name) if isinstance(r.name, (int, np.integer)) else None

                    if real_idx is None or real_idx >= len(base_raw):
                        st.error("Ligne introuvable.")
                    else:
                        row = base_raw.loc[real_idx].to_dict()
                        # journal
                        try:
                            jr = json.loads(row.get(ESC_JR, "[]"))
                            if not isinstance(jr, list): jr=[]
                        except Exception:
                            jr = []
                        jr.append({"ts": pd.Timestamp.now().isoformat(timespec="seconds"), "amount": float(amt), "note": _safe_str(note)})
                        row[ESC_JR] = json.dumps(jr, ensure_ascii=False)
                        # cumule transféré
                        try:
                            curr_tr = float(row.get(ESC_TR, 0.0) or 0.0)
                        except Exception:
                            curr_tr = 0.0
                        row[ESC_TR] = curr_tr + float(amt)
                        base_raw.loc[real_idx] = row

                        # normalise & écrit
                        base_norm = normalize_clients(base_raw.copy())
                        write_sheet_inplace(clients_path, sheet_choice, base_norm)
                        st.success("Transfert enregistré.")
                        st.rerun()

    st.divider()
    st.markdown("### 📒 Journal ESCROW (tous dossiers)")
    rows = []
    for j, r in base_raw.iterrows():
        try:
            jr = json.loads(r.get(ESC_JR, "[]"))
            if not isinstance(jr, list): jr=[]
        except Exception:
            jr = []
        for ent in jr:
            rows.append({
                "Horodatage": ent.get("ts",""),
                DOSSIER_COL: r.get(DOSSIER_COL,""),
                "ID_Client": r.get("ID_Client",""),
                "Nom": r.get("Nom",""),
                "Visa": r.get("Visa",""),
                "Montant": float(ent.get("amount",0.0)),
                "Note": ent.get("note","")
            })
    if rows:
        jdf = pd.DataFrame(rows)
        # tri temporel si possible
        try:
            jdf["Horodatage_dt"] = pd.to_datetime(jdf["Horodatage"], errors="coerce")
            jdf = jdf.sort_values("Horodatage_dt").drop(columns=["Horodatage_dt"])
        except Exception:
            jdf = jdf.sort_values("Horodatage")
        jdf["Montant"] = jdf["Montant"].map(_fmt_money_us)
        st.dataframe(jdf, use_container_width=True)
    else:
        st.caption("Aucun transfert journalisé.")