# app.py
import io
import json
import hashlib
from datetime import date
from pathlib import Path

import streamlit as st
import pandas as pd

st.set_page_config(page_title="📊 Visas — Simplifié", layout="wide")
st.title("📊 Visas — Tableau simplifié")

# --- KPI compacts (CSS) ---
st.markdown("""
<style>
.small-kpi [data-testid="stMetricValue"] { font-size: 1.15rem; line-height: 1.1; }
.small-kpi [data-testid="stMetricLabel"] { font-size: 0.80rem; opacity: 0.8; }
</style>
""", unsafe_allow_html=True)

# =========================
# Helpers
# =========================
def _first_col(df: pd.DataFrame, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def _safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def _to_num(s: pd.Series) -> pd.Series:
    cleaned = (
        s.astype(str)
         .str.replace("\u00a0", "", regex=False)
         .str.replace("\u202f", "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.replace("$", "", regex=False)
         .str.replace(",", "", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)

def _to_date(s: pd.Series) -> pd.Series:
    d = pd.to_datetime(s, errors="coerce")
    try:
        d = d.dt.tz_localize(None)
    except Exception:
        pass
    return d.dt.normalize().dt.date  # YYYY-MM-DD

def _fmt_money_us(v: float) -> str:
    try:
        return f"${float(v):,.2f}"
    except Exception:
        return "$0.00"

def _make_client_id_from_row(row) -> str:
    base = "|".join([
        _safe_str(row.get("Nom")),
        _safe_str(row.get("Telephone")),
        _safe_str(row.get("Date")),
    ])
    h = hashlib.sha1(base.encode("utf-8")).hexdigest()[:8].upper()
    return f"CL-{h}"

def _dedupe_ids(series: pd.Series) -> pd.Series:
    s = series.copy()
    counts = {}
    for i, val in enumerate(s):
        val = _safe_str(val)
        if not val:
            continue
        counts[val] = counts.get(val, 0) + 1
        if counts[val] > 1:
            s.iloc[i] = f"{val}-{counts[val]:02d}"
    return s

def looks_like_reference(df: pd.DataFrame) -> bool:
    cols = set(map(str.lower, df.columns.astype(str)))
    has_ref = {"categories", "visa"} <= cols
    no_money = not ({"montant", "honoraires", "acomptes", "payé", "reste", "solde"} & cols)
    return has_ref and no_money

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Uniformise Date/Visa/Montant/Payé/Reste, génère ID_Client, calcule Mois=MM (interne)."""
    df = df.copy()

    # Date (sans heure)
    if "Date" in df.columns:
        df["Date"] = _to_date(df["Date"])
    else:
        df["Date"] = pd.NaT

    # Mois (MM) interne (non affiché, mais utile pour les graphes)
    df["Mois"] = df["Date"].apply(lambda x: f"{x.month:02d}" if pd.notna(x) else pd.NA)

    # Visa / Categories
    visa_col = _first_col(df, ["Visa", "Categories", "Catégorie", "TypeVisa"])
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"

    # Montant / Payé
    if "Montant" in df.columns:
        df["Montant"] = _to_num(df["Montant"])
    else:
        src_montant = _first_col(df, ["Honoraires", "Total", "Amount"])
        df["Montant"] = _to_num(df[src_montant]) if src_montant else 0.0

    if "Payé" in df.columns:
        df["Payé"] = _to_num(df["Payé"])
    else:
        src_paye = _first_col(df, ["Acomptes", "Paye", "Paid"])
        df["Payé"] = _to_num(df[src_paye]) if src_paye else 0.0

    # Reste (toujours calculé)
    df["Reste"] = (df["Montant"] - df["Payé"]).fillna(0.0)

    # ID client auto
    if "ID_Client" not in df.columns:
        df["ID_Client"] = ""
    need_id = df["ID_Client"].astype(str).str.strip().eq("") | df["ID_Client"].isna()
    if need_id.any():
        generated = df.loc[need_id].apply(_make_client_id_from_row, axis=1)
        df.loc[need_id, "ID_Client"] = generated
    df["ID_Client"] = _dedupe_ids(df["ID_Client"].astype(str).str.strip())

    return df

def write_updated_excel_bytes(original_bytes: bytes, sheet_to_replace: str, new_df: pd.DataFrame) -> bytes:
    """Remplace (ou crée) la feuille `sheet_to_replace` et conserve les autres."""
    xls = pd.ExcelFile(io.BytesIO(original_bytes))
    out = io.BytesIO()
    target_written = False
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name in xls.sheet_names:
            if name == sheet_to_replace:
                dfw = new_df.copy()
                for c in dfw.columns:
                    if dfw[c].dtype == "object":
                        dfw[c] = dfw[c].astype(str).fillna("")
                dfw.to_excel(writer, sheet_name=name, index=False)
                target_written = True
            else:
                pd.read_excel(xls, sheet_name=name).to_excel(writer, sheet_name=name, index=False)
        if not target_written:
            dfw = new_df.copy()
            for c in dfw.columns:
                if dfw[c].dtype == "object":
                    dfw[c] = dfw[c].astype(str).fillna("")
            dfw.to_excel(writer, sheet_name=sheet_to_replace, index=False)
    out.seek(0)
    return out.read()

# =========================
# Chargement source (SANS CACHE, ID basé sur le contenu)
# =========================
def load_excel_bytes(xlsx_input):
    """
    Charge l'Excel et retourne (sheet_names, data_bytes, source_id, kind, name, path) sans cache.
    Identifiant de source basé sur le CONTENU (hash + taille) pour éviter toute confusion.
    """
    if hasattr(xlsx_input, "read"):  # upload
        data = xlsx_input.read()
        name = getattr(xlsx_input, "name", "uploaded.xlsx")
        kind = "upload"; path = None
    else:  # chemin disque
        p = Path(xlsx_input)
        data = p.read_bytes()
        name = p.name
        kind = "path"; path = str(p)

    h = hashlib.sha1(data).hexdigest()[:10]
    src_id = f"{kind}:{name}:{len(data)}:{h}"
    xls = pd.ExcelFile(io.BytesIO(data))
    return xls.sheet_names, data, src_id, kind, name, path

def read_sheet_from_bytes(data_bytes: bytes, sheet_name: str, normalize: bool) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(data_bytes))
    if sheet_name not in xls.sheet_names:
        base = pd.DataFrame(columns=["ID_Client","Nom","Telephone","Email","Date","Visa","Montant","Payé","Reste",
                                     "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"])
        return normalize_dataframe(base) if normalize else base
    df = pd.read_excel(xls, sheet_name=sheet_name)
    if normalize and not looks_like_reference(df):
        df = normalize_dataframe(df)
    return df

DEFAULT_CANDIDATES = [
    "/mnt/data/Visa_Clients_20251001-114844.xlsx",
    "/mnt/data/visa_analytics_datecol.xlsx",
]

# === Sidebar: Source & Sauvegarde
st.sidebar.header("Données")
source_mode = st.sidebar.radio("Source", ["Fichier par défaut", "Importer un Excel"])

if source_mode == "Fichier par défaut":
    path = next((p for p in DEFAULT_CANDIDATES if Path(p).exists()), None)
    if not path:
        st.sidebar.error("Aucun fichier par défaut trouvé. Importez un fichier.")
        st.stop()
    sheet_names0, data_bytes, source_id, source_kind, source_name, source_path = load_excel_bytes(path)
else:
    up = st.sidebar.file_uploader("Dépose un Excel (.xlsx, .xls)", type=["xlsx", "xls"])
    if not up:
        st.info("Importe un fichier pour commencer.")
        st.stop()
    sheet_names0, data_bytes, source_id, source_kind, source_name, source_path = load_excel_bytes(up)

if "excel_bytes_current" not in st.session_state or st.session_state.get("excel_source_id") != source_id:
    st.session_state["excel_bytes_current"] = data_bytes
    st.session_state["excel_source_id"] = source_id
    st.session_state["excel_source_kind"] = source_kind
    st.session_state["excel_source_name"] = source_name
    st.session_state["excel_source_path"] = source_path

current_bytes = st.session_state["excel_bytes_current"]
sheet_names_current = pd.ExcelFile(io.BytesIO(current_bytes)).sheet_names

st.sidebar.caption(f"Fichier courant : **{st.session_state['excel_source_name']}**")
if st.sidebar.button("🔄 Rafraîchir l’affichage"):
    st.rerun()

# Écriture sur disque si fichier local côté serveur
if st.session_state["excel_source_kind"] == "path":
    if st.sidebar.button("💾 Écrire sur disque (même fichier)"):
        try:
            Path(st.session_state["excel_source_path"]).write_bytes(st.session_state["excel_bytes_current"])
            st.sidebar.success(f"Écrit dans {st.session_state['excel_source_name']}")
        except Exception as e:
            st.sidebar.error(f"Échec écriture : {e}")

# Téléchargement même nom
st.sidebar.download_button(
    "💾 Sauvegarder (même nom)",
    data=st.session_state["excel_bytes_current"],
    file_name=st.session_state["excel_source_name"],
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Télécharge le fichier actuel en mémoire en conservant exactement le même nom."
)

# Redirection feuille (post-write)
if "pending_sheet_choice" in st.session_state:
    pending = st.session_state.pop("pending_sheet_choice")
    if pending in sheet_names_current:
        st.session_state["sheet_choice"] = pending

# Sélection des feuilles
preferred_order = ["Clients", "Visa", "Données normalisées"]
default_sheet = next((s for s in preferred_order if s in sheet_names_current), sheet_names_current[0])
sheet_choice_default = st.session_state.get("sheet_choice", default_sheet)
if sheet_choice_default not in sheet_names_current:
    sheet_choice_default = default_sheet

sheet_choice = st.sidebar.selectbox(
    "Feuille Excel (vue Dashboard)",
    sheet_names_current,
    index=sheet_names_current.index(sheet_choice_default),
    key="sheet_choice"
)

CLIENT_SHEET_DEFAULT = "Clients"
client_sheet_exists = CLIENT_SHEET_DEFAULT in sheet_names_current
client_target_sheet = st.sidebar.selectbox(
    "Feuille *Clients* (cible CRUD)",
    sheet_names_current,
    index=sheet_names_current.index(CLIENT_SHEET_DEFAULT) if client_sheet_exists else sheet_names_current.index(sheet_choice),
    help="Toutes les opérations Créer/Modifier/Supprimer s’appliquent à cette feuille."
)

# Référentiel Visa ?
sample_df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=False).head(5)
is_ref = looks_like_reference(sample_df)

# =========================
# TABS
# =========================
tabs = st.tabs(["Dashboard", "Clients (CRUD)"])

# =========================
# TAB Dashboard
# =========================
with tabs[0]:
    df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=True)

    # Filtres (par défaut rien n'est coché) + inclure lignes sans date
    with st.container():
        c1, c2, c3 = st.columns(3)
        years = sorted({d.year for d in df["Date"] if pd.notna(d)}) if "Date" in df.columns else []
        months_present = sorted(df["Mois"].dropna().unique()) if "Mois" in df.columns else []
        visas = sorted(df["Visa"].dropna().astype(str).unique()) if "Visa" in df.columns else []

        year_sel  = c1.multiselect("Année", years, default=[])
        month_sel = c2.multiselect("Mois (MM)", months_present, default=[])
        visa_sel  = c3.multiselect("Type de visa", visas, default=[])

        c4, c5, c6 = st.columns([1,1,1])
        include_na_dates = c6.checkbox("Inclure lignes sans date", value=True)

        def make_range_slider(df_src: pd.DataFrame, col: str, label: str, container, fmt=lambda x: f"{x:,.2f}"):
            if col not in df_src.columns or df_src[col].dropna().empty:
                container.caption(f"{label} : aucune donnée")
                return None
            vmin = float(df_src[col].min())
            vmax = float(df_src[col].max())
            if not (vmin < vmax):
                container.caption(f"{label} : valeur unique = {fmt(vmin)}")
                return (vmin, vmax)
            span = vmax - vmin
            step = 1.0 if span > 1000 else 0.1 if span > 10 else 0.01
            return container.slider(label, min_value=vmin, max_value=vmax, value=(vmin, vmax), step=step)

        pay_range   = make_range_slider(df, "Payé", "Payé (min-max)", c4, fmt=lambda x: _fmt_money_us(x))
        solde_range = make_range_slider(df, "Reste", "Solde / Reste (min-max)", c5, fmt=lambda x: _fmt_money_us(x))

    # Application des filtres
    f = df.copy()
    if "Date" in f.columns and year_sel:
        mask_year = f["Date"].apply(lambda x: (pd.notna(x) and x.year in year_sel))
        if include_na_dates: mask_year = mask_year | f["Date"].isna()
        f = f[mask_year]
    if "Mois" in f.columns and month_sel:
        mask_month = f["Mois"].isin(month_sel)
        if include_na_dates: mask_month = mask_month | f["Mois"].isna()
        f = f[mask_month]
    if "Visa" in f.columns and visa_sel:
        f = f[f["Visa"].astype(str).isin(visa_sel)]
    if "Payé" in f.columns and pay_range is not None:
        f = f[(f["Payé"] >= pay_range[0]) & (f["Payé"] <= pay_range[1])]
    if "Reste" in f.columns and solde_range is not None:
        f = f[(f["Reste"] >= solde_range[0]) & (f["Reste"] <= solde_range[1])]

    hidden = len(df) - len(f)
    if hidden > 0:
        st.caption(f"🔎 {hidden} ligne(s) masquée(s) par les filtres. Décoche tout pour tout voir.")

    # KPI
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Montant total", _fmt_money_us(float(f["Montant"].sum())) if "Montant" in f.columns else "—")
    k3.metric("Payé",         _fmt_money_us(float(f["Payé"].sum()))     if "Payé" in f.columns else "—")
    k4.metric("Reste",        _fmt_money_us(float(f["Reste"].sum()))    if "Reste" in f.columns else "—")
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    st.subheader("📈 Nombre de dossiers par mois (MM)")
    if "Mois" in f.columns:
        counts = (
            f.dropna(subset=["Mois"])
             .groupby("Mois")
             .size()
             .rename("Dossiers")
             .reset_index()
             .sort_values("Mois")
        )
        st.bar_chart(counts.set_index("Mois"))
    else:
        st.info("Aucune colonne 'Mois' exploitable.")

    # Tableau principal
    st.subheader("📋 Données")
    cols_show = [c for c in ["ID_Client","Nom","Telephone","Email","Date","Visa","Montant","Payé","Reste",
                             "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"] if c in f.columns]
    table = f.copy()
    for col in ["Montant","Payé","Reste"]:
        if col in table.columns:
            table[col] = table[col].map(_fmt_money_us)
    if "Date" in table.columns:
        table["Date"] = table["Date"].astype(str)
    st.dataframe(
        table[cols_show].sort_values(by=[c for c in ["Date","Visa"] if c in table.columns], na_position="last"),
        use_container_width=True
    )

# =========================
# TAB Clients (CRUD)
# =========================
with tabs[1]:
    st.subheader("👤 Clients — Créer / Modifier / Supprimer")
    st.caption(f"Feuille cible : **{client_target_sheet}**")

    if st.button("🔄 Rafraîchir"):
        st.rerun()

    # Charger la feuille cible (brute)
    orig = read_sheet_from_bytes(current_bytes, client_target_sheet, normalize=False).copy()
    orig = orig.copy()
    orig["_RowID"] = range(len(orig))

    # Colonnes logiques utilisées pour la contrainte RFE
    BOOL_RFE = [c for c in orig.columns if c.lower() == "rfe"]
    BOOL_ENVOYE = [c for c in orig.columns if c.lower() == "dossier envoyé"]
    BOOL_REFUSE = [c for c in orig.columns if c.lower() == "dossier refusé"]
    BOOL_ANNULE = [c for c in orig.columns if c.lower() == "dossier annulé"]
    BOOL_APPROUVE = [c for c in orig.columns if c.lower() == "dossier approuvé"]

    # Liste de visas de référence si la feuille "Visa" existe
    try:
        visa_ref = read_sheet_from_bytes(current_bytes, "Visa", normalize=False)
        visa_options = sorted(visa_ref["Visa"].dropna().astype(str).unique()) if "Visa" in visa_ref.columns else []
    except Exception:
        visa_options = []

    action = st.radio("Action", ["Créer", "Modifier", "Supprimer"], horizontal=True)

    # ---------- CRÉER ----------
    if action == "Créer":
        st.markdown("### ➕ Nouveau client")

        # Squelette si feuille vide
        if len(orig.columns) == 1 and "_RowID" in orig.columns:
            base_cols = ["ID_Client","Nom","Telephone","Email","Date","Visa","Montant","Payé","Reste",
                         "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"]
            orig = pd.DataFrame(columns=base_cols + ["_RowID"])

        # --- Formulaire (sans Catégorie, sans Statut, sans Paiements) ---
        with st.form("create_form", clear_on_submit=False):
            c1, c2 = st.columns(2)
            nom = c1.text_input("Nom")
            tel = c2.text_input("Telephone")
            c3, c4 = st.columns(2)
            email = c3.text_input("Email")
            d = c4.date_input("Date", value=date.today())
            # Visa seul (si liste dispo -> selectbox, sinon champ texte)
            if visa_options:
                visa = st.selectbox("Visa", visa_options, index=0 if visa_options else None)
            else:
                visa = st.text_input("Visa")
            c5, c6 = st.columns(2)
            montant = c5.number_input("Montant (US $)", value=0.0, step=10.0, format="%.2f")
            paye    = c6.number_input("Payé (US $)", value=0.0, step=10.0, format="%.2f")

            # Booléens (optionnels, affichés si colonnes existent dans l'Excel)
            if BOOL_ENVOYE or BOOL_REFUSE or BOOL_ANNULE or BOOL_APPROUVE or BOOL_RFE:
                st.markdown("#### État du dossier")
            val_envoye  = st.checkbox("Dossier envoyé", value=False)  if BOOL_ENVOYE else False
            val_refuse  = st.checkbox("Dossier refusé", value=False)  if BOOL_REFUSE else False
            val_annule  = st.checkbox("Dossier annulé", value=False)  if BOOL_ANNULE else False
            val_approuve= st.checkbox("Dossier approuvé", value=False)if BOOL_APPROUVE else False
            val_rfe     = st.checkbox("RFE", value=False)             if BOOL_RFE else False

            submitted = st.form_submit_button("💾 Sauvegarder", type="primary")

        if submitted:
            # Contrainte RFE : ne peut être True que si envoyé/refusé/annulé au moins un True
            if val_rfe and not (val_envoye or val_refuse or val_annule):
                st.error("RFE ne peut être activé que si **Dossier envoyé** OU **Dossier refusé** OU **Dossier annulé** est coché.")
                st.stop()

            live_before = read_sheet_from_bytes(
                st.session_state["excel_bytes_current"],
                client_target_sheet,
                normalize=False
            ).copy()
            if live_before.empty and client_target_sheet not in pd.ExcelFile(io.BytesIO(st.session_state["excel_bytes_current"])).sheet_names:
                live_before = pd.DataFrame(columns=["ID_Client","Nom","Telephone","Email","Date","Visa","Montant","Payé","Reste",
                                                    "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"])

            # Colonnes minimales
            for must in ["ID_Client","Nom","Telephone","Email","Date","Visa","Montant","Payé","Reste",
                         "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"]:
                if must not in live_before.columns:
                    live_before[must] = "" if must not in {"Montant","Payé","Reste",
                                                           "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"} else 0.0

            # ID auto
            gen_id = _make_client_id_from_row({"Nom": nom, "Telephone": tel, "Date": d})
            existing_ids = set(live_before["ID_Client"].astype(str)) if "ID_Client" in live_before.columns else set()
            tmp = gen_id; n=1
            while tmp in existing_ids:
                n += 1
                tmp = f"{gen_id}-{n:02d}"
            id_client = tmp

            # Nouvelle ligne
            new_row = {}
            for c in live_before.columns:
                cl = c.lower()
                if c == "ID_Client": new_row[c] = id_client
                elif c == "Nom": new_row[c] = _safe_str(nom)
                elif c == "Telephone": new_row[c] = _safe_str(tel)
                elif c == "Email": new_row[c] = _safe_str(email)
                elif c == "Date": new_row[c] = str(d) if d else ""
                elif c == "Visa": new_row[c] = _safe_str(visa)
                elif c == "Montant": new_row[c] = float(montant or 0)
                elif c == "Payé": new_row[c] = float(paye or 0)
                elif c == "Reste": new_row[c] = 0.0  # recalculé en dessous
                elif cl == "rfe": new_row[c] = bool(val_rfe)
                elif cl == "dossier envoyé": new_row[c] = bool(val_envoye)
                elif cl == "dossier approuvé": new_row[c] = bool(val_approuve)
                elif cl == "dossier refusé": new_row[c] = bool(val_refuse)
                elif cl == "dossier annulé": new_row[c] = bool(val_annule)
                else:
                    new_row[c] = _safe_str("")  # champs annexes ignorés en création

            # Recalcul Reste
            new_row["Reste"] = float(new_row.get("Montant", 0)) - float(new_row.get("Payé", 0))

            live_after = pd.concat([live_before, pd.DataFrame([new_row])], ignore_index=True)
            try:
                updated_bytes = write_updated_excel_bytes(
                    st.session_state["excel_bytes_current"],
                    client_target_sheet,
                    live_after
                )
                st.session_state["excel_bytes_current"] = updated_bytes
                st.session_state["pending_sheet_choice"] = client_target_sheet
                st.success(f"✅ Client sauvegardé (ID: {id_client}). Utilise **Sauvegarder (même nom)** (sidebar).")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur à l’écriture : {e}")

    # ---------- MODIFIER ----------
    if action == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if orig.drop(columns=["_RowID"]).empty:
            st.info("Aucun client à modifier.")
        else:
            options = [(int(r["_RowID"]), f"{_safe_str(r.get('ID_Client'))} — {_safe_str(r.get('Nom'))}") for _, r in orig.iterrows()]
            sel_label = st.selectbox("Sélection", [lab for _, lab in options])
            sel_rowid = [rid for rid, lab in options if lab == sel_label][0]
            sel_idx = orig.index[orig["_RowID"] == sel_rowid][0]

            init = orig.loc[sel_idx].to_dict()

            with st.form("edit_form", clear_on_submit=False):
                c1, c2 = st.columns(2)
                nom = c1.text_input("Nom", value=_safe_str(init.get("Nom")))
                tel = c2.text_input("Telephone", value=_safe_str(init.get("Telephone")))
                c3, c4 = st.columns(2)
                email = c3.text_input("Email", value=_safe_str(init.get("Email")))
                try:
                    d_init = pd.to_datetime(init.get("Date")).date() if _safe_str(init.get("Date")) else date.today()
                except Exception:
                    d_init = date.today()
                d = c4.date_input("Date", value=d_init)
                # Visa
                if visa_options:
                    try:
                        idx = visa_options.index(_safe_str(init.get("Visa")))
                    except Exception:
                        idx = 0 if visa_options else None
                    visa = st.selectbox("Visa", visa_options, index=idx)
                else:
                    visa = st.text_input("Visa", value=_safe_str(init.get("Visa")))
                c5, c6 = st.columns(2)
                try: montant = float(init.get("Montant", 0))
                except Exception: montant = 0.0
                try: paye = float(init.get("Payé", 0))
                except Exception: paye = 0.0
                montant = c5.number_input("Montant (US $)", value=montant, step=10.0, format="%.2f")
                paye    = c6.number_input("Payé (US $)", value=paye,    step=10.0, format="%.2f")

                # Booléens (affichés si colonnes existent)
                if BOOL_ENVOYE or BOOL_REFUSE or BOOL_ANNULE or BOOL_APPROUVE or BOOL_RFE:
                    st.markdown("#### État du dossier")
                val_envoye   = st.checkbox("Dossier envoyé",  value=bool(init.get("Dossier envoyé")))  if BOOL_ENVOYE else False
                val_refuse   = st.checkbox("Dossier refusé",  value=bool(init.get("Dossier refusé")))  if BOOL_REFUSE else False
                val_annule   = st.checkbox("Dossier annulé",  value=bool(init.get("Dossier annulé")))  if BOOL_ANNULE else False
                val_approuve = st.checkbox("Dossier approuvé",value=bool(init.get("Dossier approuvé")))if BOOL_APPROUVE else False
                val_rfe      = st.checkbox("RFE",             value=bool(init.get("RFE")))             if BOOL_RFE else False

                submitted = st.form_submit_button("💾 Enregistrer", type="primary")

            if submitted:
                if val_rfe and not (val_envoye or val_refuse or val_annule):
                    st.error("RFE ne peut être activé que si **Dossier envoyé** OU **Dossier refusé** OU **Dossier annulé** est coché.")
                    st.stop()

                live = read_sheet_from_bytes(st.session_state["excel_bytes_current"], client_target_sheet, normalize=False).copy()
                if live.empty:
                    st.error("Feuille cible introuvable.")
                else:
                    # Trouver la ligne par ID si possible
                    target_idx = None
                    if "ID_Client" in live.columns and _safe_str(init.get("ID_Client")):
                        hits = live.index[live["ID_Client"].astype(str) == _safe_str(init.get("ID_Client"))]
                        if len(hits) > 0: target_idx = hits[0]
                    if target_idx is None:
                        mask = (live.get("Nom","").astype(str) == _safe_str(init.get("Nom"))) & \
                               (live.get("Telephone","").astype(str) == _safe_str(init.get("Telephone")))
                        hit2 = live.index[mask]
                        target_idx = hit2[0] if len(hit2) > 0 else None

                    if target_idx is None:
                        st.error("Ligne cible introuvable.")
                    else:
                        live.at[target_idx, "Nom"] = _safe_str(nom)
                        live.at[target_idx, "Telephone"] = _safe_str(tel)
                        live.at[target_idx, "Email"] = _safe_str(email)
                        live.at[target_idx, "Date"] = str(d) if d else ""
                        live.at[target_idx, "Visa"] = _safe_str(visa)
                        live.at[target_idx, "Montant"] = float(montant or 0)
                        live.at[target_idx, "Payé"] = float(paye or 0)
                        # Recalcule Reste
                        live.at[target_idx, "Reste"] = float(live.at[target_idx, "Montant"]) - float(live.at[target_idx, "Payé"])

                        # Booléens si présents
                        if BOOL_ENVOYE:  live.at[target_idx, BOOL_ENVOYE[0]]  = bool(val_envoye)
                        if BOOL_REFUSE:  live.at[target_idx, BOOL_REFUSE[0]]  = bool(val_refuse)
                        if BOOL_ANNULE:  live.at[target_idx, BOOL_ANNULE[0]]  = bool(val_annule)
                        if BOOL_APPROUVE:live.at[target_idx, BOOL_APPROUVE[0]]= bool(val_approuve)
                        if BOOL_RFE:     live.at[target_idx, BOOL_RFE[0]]     = bool(val_rfe)

                        try:
                            updated_bytes = write_updated_excel_bytes(st.session_state["excel_bytes_current"], client_target_sheet, live)
                            st.session_state["excel_bytes_current"] = updated_bytes
                            st.success("✅ Modifications enregistrées. Utilise **Sauvegarder (même nom)** (sidebar).")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur à l’écriture : {e}")

    # ---------- SUPPRIMER ----------
    if action == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if orig.drop(columns=["_RowID"]).empty:
            st.info("Aucun client à supprimer.")
        else:
            options = [(int(r["_RowID"]), f"{_safe_str(r.get('ID_Client'))} — {_safe_str(r.get('Nom'))}") for _, r in orig.iterrows()]
            sel_label = st.selectbox("Sélection", [lab for _, lab in options])
            sel_rowid = [rid for rid, lab in options if lab == sel_label][0]
            sel_idx = orig.index[orig["_RowID"] == sel_rowid][0]

            st.error("⚠️ Cette action est irréversible.")
            confirm = st.checkbox("Je confirme la suppression définitive de ce client.")
            if st.button("Supprimer", type="primary", disabled=not confirm):
                try:
                    live = read_sheet_from_bytes(st.session_state["excel_bytes_current"], client_target_sheet, normalize=False).copy()
                    if "ID_Client" in live.columns and _safe_str(orig.at[sel_idx, "ID_Client"]):
                        live = live[live["ID_Client"].astype(str) != _safe_str(orig.at[sel_idx, "ID_Client"])].reset_index(drop=True)
                    else:
                        nom = _safe_str(orig.at[sel_idx, "Nom"]); tel = _safe_str(orig.at[sel_idx, "Telephone"])
                        live = live[~((live.get("Nom","").astype(str)==nom) & (live.get("Telephone","").astype(str)==tel))].reset_index(drop=True)

                    updated_bytes = write_updated_excel_bytes(st.session_state["excel_bytes_current"], client_target_sheet, live)
                    st.session_state["excel_bytes_current"] = updated_bytes
                    st.success("✅ Client supprimé. Utilise **Sauvegarder (même nom)** (sidebar).")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur à l’écriture : {e}")

    # ----- Debug -----
    with st.expander("🛠️ Debug"):
        st.write("Source:", st.session_state.get("excel_source_id"))
        try:
            live_sheets = pd.ExcelFile(io.BytesIO(st.session_state["excel_bytes_current"])).sheet_names
        except Exception:
            live_sheets = []
        st.write("Feuilles (mémoire):", live_sheets)
        if "ID_Client" in orig.columns:
            st.write("Nb lignes (feuille cible, avant event):", len(orig.drop(columns=["_RowID"], errors="ignore")))
        st.caption("ℹ️ Tout est en **mémoire**. Utilise **Sauvegarder (même nom)** pour récupérer le fichier sous le nom d’origine.")
