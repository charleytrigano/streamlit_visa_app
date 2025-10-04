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
    # Parse US: "$1,234.56" -> 1234.56
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

def _parse_paiements(x):
    if isinstance(x, list):
        return x
    if pd.isna(x):
        return []
    try:
        v = json.loads(x)
        return v if isinstance(v, list) else []
    except Exception:
        return []

def _sum_payments(pay_list) -> float:
    total = 0.0
    for p in (pay_list or []):
        try:
            amt = float(p.get("amount", 0) or 0) if isinstance(p, dict) else float(p)
        except Exception:
            amt = 0.0
        total += amt
    return total

def _make_client_id_from_row(row) -> str:
    # ID stable depuis Nom + Telephone + Date
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
    """Détecte un onglet de référence (ex: 'Visa' avec Categories/Visa/Definition)."""
    cols = set(map(str.lower, df.columns.astype(str)))
    has_ref = {"categories", "visa"} <= cols
    no_money = not ({"montant", "honoraires", "acomptes", "payé", "reste", "solde"} & cols)
    return has_ref and no_money

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Uniformise Date/Visa/Statut/Montant/Payé/Reste, génère ID_Client, calcule Mois=MM (interne)."""
    df = df.copy()

    # Date (sans heure)
    if "Date" in df.columns:
        df["Date"] = _to_date(df["Date"])
    else:
        df["Date"] = pd.NaT

    # Mois (MM) interne
    df["Mois"] = df["Date"].apply(lambda x: f"{x.month:02d}" if pd.notna(x) else pd.NA)

    # Visa / Categories
    visa_col = _first_col(df, ["Visa", "Categories", "Catégorie", "TypeVisa"])
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"

    # Statut
    if "__Statut règlement__" in df.columns and "Statut" not in df.columns:
        df = df.rename(columns={"__Statut règlement__": "Statut"})
    if "Statut" not in df.columns:
        df["Statut"] = "Inconnu"
    else:
        df["Statut"] = df["Statut"].astype(str).fillna("Inconnu")

    # Montant
    if "Montant" in df.columns:
        df["Montant"] = _to_num(df["Montant"])
    else:
        src_montant = _first_col(df, ["Honoraires", "Total", "Amount"])
        df["Montant"] = _to_num(df[src_montant]) if src_montant else 0.0

    # Paiements JSON -> TotalAcomptes
    if "Paiements" in df.columns:
        parsed = df["Paiements"].apply(_parse_paiements)
        df["TotalAcomptes"] = parsed.apply(_sum_payments)

    # Payé
    if "Payé" in df.columns:
        df["Payé"] = _to_num(df["Payé"])
    else:
        src_paye = _first_col(df, ["TotalAcomptes", "Acomptes", "Paye", "Paid"])
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
        # si la feuille n'existait pas, on la crée
        if not target_written:
            dfw = new_df.copy()
            for c in dfw.columns:
                if dfw[c].dtype == "object":
                    dfw[c] = dfw[c].astype(str).fillna("")
            dfw.to_excel(writer, sheet_name=sheet_to_replace, index=False)
    out.seek(0)
    return out.read()

def build_categories_to_visa_map(data_bytes: bytes, visa_sheet_name: str = "Visa") -> dict:
    """Construit un mapping {categories_normalisées: visa} depuis l’onglet 'Visa'."""
    try:
        xls = pd.ExcelFile(io.BytesIO(data_bytes))
        if visa_sheet_name not in xls.sheet_names:
            return {}
        ref = pd.read_excel(xls, sheet_name=visa_sheet_name)
        if "Categories" not in ref.columns or "Visa" not in ref.columns:
            return {}
        m = {}
        for _, r in ref.iterrows():
            cat = _safe_str(r.get("Categories")).lower()
            vis = _safe_str(r.get("Visa"))
            if cat:
                m[cat] = vis
        return m
    except Exception:
        return {}

def enrich_visa_from_categories(df: pd.DataFrame, cat2visa: dict):
    """Si Visa est vide mais Categories correspond dans le mapping, remplit Visa."""
    if not cat2visa or "Categories" not in df.columns or "Visa" not in df.columns:
        return df, 0
    df = df.copy()
    mask_empty_visa = df["Visa"].astype(str).str.strip().eq("") | df["Visa"].isna() | (df["Visa"] == "Inconnu")
    to_fill = 0
    def _fill(row):
        nonlocal to_fill
        if mask_empty_visa.loc[row.name]:
            key = _safe_str(row.get("Categories")).lower()
            if key in cat2visa and _safe_str(cat2visa[key]):
                to_fill += 1
                return cat2visa[key]
        return row["Visa"]
    if mask_empty_visa.any():
        df.loc[mask_empty_visa, "Visa"] = df[mask_empty_visa].apply(_fill, axis=1)
    return df, to_fill

# =========================
# Cache sérialisable
# =========================
@st.cache_data
def load_excel_bytes(xlsx_input):
    """Retourne (sheet_names, data_bytes, source_id) pour un chemin ou un fichier uploadé."""
    if hasattr(xlsx_input, "read"):
        data = xlsx_input.read()
        src_id = f"upload:{getattr(xlsx_input, 'name', 'uploaded')}"
    else:
        data = Path(xlsx_input).read_bytes()
        src_id = f"path:{xlsx_input}"
    xls = pd.ExcelFile(io.BytesIO(data))
    return xls.sheet_names, data, src_id

def read_sheet_from_bytes(data_bytes: bytes, sheet_name: str, normalize: bool) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(data_bytes))
    if sheet_name not in xls.sheet_names:
        # crée feuille vide avec colonnes de base
        base = pd.DataFrame(columns=["ID_Client","Nom","Telephone","Email","Date","Visa","Statut","Montant","Payé","Reste","Paiements"])
        return normalize_dataframe(base) if normalize else base
    df = pd.read_excel(xls, sheet_name=sheet_name)
    if normalize and not looks_like_reference(df):
        df = normalize_dataframe(df)
    return df

# =========================
# Source & Sélection feuille
# =========================
DEFAULT_CANDIDATES = [
    "/mnt/data/Visa_Clients_20251001-114844.xlsx",
    "/mnt/data/visa_analytics_datecol.xlsx",
]

st.sidebar.header("Données")
source_mode = st.sidebar.radio("Source", ["Fichier par défaut", "Importer un Excel"])

if source_mode == "Fichier par défaut":
    path = next((p for p in DEFAULT_CANDIDATES if Path(p).exists()), None)
    if not path:
        st.sidebar.error("Aucun fichier par défaut trouvé. Importez un fichier.")
        st.stop()
    st.sidebar.success(f"Fichier: {path}")
    sheet_names, data_bytes, source_id = load_excel_bytes(path)
else:
    up = st.sidebar.file_uploader("Dépose un Excel (.xlsx, .xls)", type=["xlsx", "xls"])
    if not up:
        st.info("Importe un fichier pour commencer.")
        st.stop()
    sheet_names, data_bytes, source_id = load_excel_bytes(up)

# état courant
if "excel_bytes_current" not in st.session_state or st.session_state.get("excel_source_id") != source_id:
    st.session_state["excel_bytes_current"] = data_bytes
    st.session_state["excel_source_id"] = source_id
current_bytes = st.session_state["excel_bytes_current"]

# --- Redirection feuille (post-écriture) AVANT le selectbox ---
if "pending_sheet_choice" in st.session_state:
    pending = st.session_state.pop("pending_sheet_choice")
    if pending in sheet_names:
        st.session_state["sheet_choice"] = pending

# bouton refresh
if st.sidebar.button("🔄 Rafraîchir"):
    st.rerun()

# choix onglet (vue Dashboard)
preferred_order = ["Données normalisées", "Clients", "Visa"]
default_sheet = next((s for s in preferred_order if s in sheet_names), sheet_names[0])
sheet_choice_default = st.session_state.get("sheet_choice", default_sheet)
if sheet_choice_default not in sheet_names:
    sheet_choice_default = default_sheet
sheet_choice = st.sidebar.selectbox(
    "Feuille Excel (vue Dashboard)",
    sheet_names,
    index=sheet_names.index(sheet_choice_default),
    key="sheet_choice"
)

# feuille cible CRUD Clients
CLIENT_SHEET_DEFAULT = "Clients"
client_sheet_exists = CLIENT_SHEET_DEFAULT in sheet_names
client_target_sheet = st.sidebar.selectbox(
    "Feuille *Clients* (cible CRUD)",
    sheet_names,
    index=sheet_names.index(CLIENT_SHEET_DEFAULT) if client_sheet_exists else sheet_names.index(sheet_choice),
    help="Toutes les opérations Créer/Modifier/Supprimer s’appliquent à cette feuille."
)

# détection référentiel
sample_df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=False).head(5)
is_ref = looks_like_reference(sample_df)

# =========================
# TABS UI
# =========================
if is_ref and sheet_choice.lower() == "visa":
    tabs = st.tabs(["Référentiel Visa (CRUD)"])
else:
    tabs = st.tabs(["Dashboard", "Clients (CRUD)"])

# =========================
# TAB 1 : Référentiel Visa (CRUD)
# =========================
if is_ref and sheet_choice.lower() == "visa":
    with tabs[0]:
        st.subheader("📚 Référentiel — Visa (éditable)")
        full_ref_df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=False).copy()
        default_cols = ["Categories", "Visa", "Definition"]
        for c in default_cols:
            if c not in full_ref_df.columns:
                full_ref_df[c] = ""
        ordered_cols = [c for c in default_cols if c in full_ref_df.columns] + [c for c in full_ref_df.columns if c not in default_cols]
        full_ref_df = full_ref_df[ordered_cols].copy()
        for c in full_ref_df.columns:
            if full_ref_df[c].dtype != "object":
                full_ref_df[c] = full_ref_df[c].astype(str)
            full_ref_df[c] = full_ref_df[c].fillna("")

        edited_df = st.data_editor(
            full_ref_df, num_rows="dynamic", use_container_width=True, hide_index=True, key="visa_editor",
        )
        col_save, col_dl, col_reset = st.columns([1,1,1])
        if col_save.button("💾 Enregistrer (remplace la feuille 'Visa')", type="primary"):
            try:
                updated_bytes = write_updated_excel_bytes(current_bytes, sheet_choice, edited_df)
                st.session_state["excel_bytes_current"] = updated_bytes
                if source_mode == "Fichier par défaut" and source_id.startswith("path:"):
                    original_path = source_id.split("path:", 1)[1]
                    try:
                        Path(original_path).write_bytes(updated_bytes)
                        st.success(f"Fichier mis à jour : {original_path}")
                    except Exception as e:
                        st.info(f"Écriture disque impossible. Téléchargez le fichier. Détail: {e}")
                st.success("Modifications enregistrées.")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur à l’enregistrement : {e}")
        col_dl.download_button(
            "⬇️ Télécharger l’Excel mis à jour",
            data=st.session_state["excel_bytes_current"],
            file_name="visa_mis_a_jour.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        if col_reset.button("↩️ Réinitialiser"):
            st.session_state.pop("excel_bytes_current", None)
            st.session_state["excel_bytes_current"] = data_bytes
            st.success("Réinitialisé.")
            st.rerun()

# =========================
# MODE Dossiers — TAB Dashboard
# =========================
if not (is_ref and sheet_choice.lower() == "visa"):
    df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=True)
    cat2visa = build_categories_to_visa_map(current_bytes, visa_sheet_name="Visa")
    df, nb_filled = enrich_visa_from_categories(df, cat2visa)

    with tabs[0]:
        if sheet_choice != client_target_sheet:
            st.info(f"ℹ️ Le CRUD Clients cible la feuille **{client_target_sheet}**.")

        # --- helper sliders robustes ---
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

        # Filtres
        with st.container():
            c1, c2, c3 = st.columns(3)
            years = sorted({d.year for d in df["Date"] if pd.notna(d)}) if "Date" in df.columns else []
            months_present = sorted(df["Mois"].dropna().unique()) if "Mois" in df.columns else []
            visas = sorted(df["Visa"].dropna().astype(str).unique()) if "Visa" in df.columns else []

            # reset des filtres si on vient d'écrire
            if st.session_state.get("reset_filters_after_write"):
                st.session_state.pop("reset_filters_after_write", None)

            year_sel = c1.multiselect("Année", years, default=years or None)
            month_sel = c2.multiselect("Mois (MM)", months_present, default=months_present or None)
            visa_sel  = c3.multiselect("Type de visa", visas, default=visas or None)

            c4, c5 = st.columns(2)
            pay_range   = make_range_slider(df, "Payé", "Payé (min-max)", c4, fmt=lambda x: _fmt_money_us(x))
            solde_range = make_range_slider(df, "Reste", "Solde / Reste (min-max)", c5, fmt=lambda x: _fmt_money_us(x))

        # Appliquer filtres
        f = df.copy()
        if "Date" in f.columns and year_sel:
            f = f[f["Date"].apply(lambda x: pd.notna(x) and x.year in year_sel)]
        if "Mois" in f.columns and month_sel:
            f = f[f["Mois"].isin(month_sel)]
        if "Visa" in f.columns and visa_sel:
            f = f[f["Visa"].astype(str).isin(visa_sel)]
        if "Payé" in f.columns and pay_range is not None:
            f = f[(f["Payé"] >= pay_range[0]) & (f["Payé"] <= pay_range[1])]
        if "Reste" in f.columns and solde_range is not None:
            f = f[(f["Reste"] >= solde_range[0]) & (f["Reste"] <= solde_range[1])]

        # KPI
        st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(f)}")
        k2.metric("Montant total", _fmt_money_us(float(f["Montant"].sum())) if "Montant" in f.columns else "—")
        k3.metric("Payé",         _fmt_money_us(float(f["Payé"].sum()))     if "Payé" in f.columns else "—")
        k4.metric("Reste",        _fmt_money_us(float(f["Reste"].sum()))    if "Reste" in f.columns else "—")
        st.markdown('</div>', unsafe_allow_html=True)

        st.divider()

        # Graph par Mois (MM)
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

        # Ajout d'acompte (ID fiable)
        st.subheader("➕ Ajouter un acompte (US $)")
        pending = df[df["Reste"] > 0.0005].copy() if "Reste" in df.columns else pd.DataFrame()
        if pending.empty:
            st.success("Tous les dossiers sont soldés ✅")
        else:
            pending["_label"] = pending.apply(
                lambda r: f'{r.get("ID_Client","")} — {r.get("Nom","")} — Reste {_fmt_money_us(float(r.get("Reste",0)))}',
                axis=1
            )
            label_to_id = pending.set_index("_label")["ID_Client"].to_dict() if "ID_Client" in pending.columns else {}

            csel, camt, cdate, cmode = st.columns([2,1,1,1])
            selected_label = csel.selectbox("Dossier à créditer", pending["_label"].tolist())
            amount = camt.number_input("Montant ($)", min_value=0.0, step=10.0, format="%.2f")
            pay_date = cdate.date_input("Date", value=date.today())
            mode = cmode.selectbox("Mode", ["CB", "Virement", "Espèces", "Chèque", "Autre"])
            note = st.text_input("Note (facultatif)", "")
            if st.button("💾 Ajouter l’acompte"):
                if amount <= 0:
                    st.warning("Montant invalide.")
                else:
                    try:
                        original_df_dash = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=False)
                        if "Paiements" not in original_df_dash.columns:
                            original_df_dash["Paiements"] = ""

                        target_row_idx = None
                        if "ID_Client" in original_df_dash.columns and label_to_id:
                            target_id = label_to_id.get(selected_label, "")
                            if target_id:
                                hits = original_df_dash.index[original_df_dash["ID_Client"].astype(str) == str(target_id)]
                                if len(hits) > 0:
                                    target_row_idx = hits[0]
                        if target_row_idx is None:
                            target_row_idx = pending.loc[pending["_label"] == selected_label].index[0]

                        raw = original_df_dash.at[target_row_idx, "Paiements"] if target_row_idx in original_df_dash.index else ""
                        try:
                            pay_list = json.loads(raw) if isinstance(raw, str) and raw.strip() else []
                            if not isinstance(pay_list, list):
                                pay_list = []
                        except Exception:
                            pay_list = []
                        pay_list.append({"date": str(pay_date), "amount": float(amount), "mode": mode, "note": note})

                        if target_row_idx not in original_df_dash.index:
                            raise RuntimeError("Ligne cible introuvable pour l’ajout d’acompte.")

                        original_df_dash.at[target_row_idx, "Paiements"] = json.dumps(pay_list, ensure_ascii=False)
                        updated_bytes = write_updated_excel_bytes(current_bytes, sheet_choice, original_df_dash)
                        st.session_state["excel_bytes_current"] = updated_bytes
                        st.session_state["reset_filters_after_write"] = True
                        if source_mode == "Fichier par défaut" and source_id.startswith("path:"):
                            original_path = source_id.split("path:", 1)[1]
                            try:
                                Path(original_path).write_bytes(updated_bytes)
                                st.success(f"Acompte ajouté et écrit dans : {original_path}")
                            except Exception as e:
                                st.info(f"Écriture disque impossible, fichier mémoire à jour. Détail: {e}")
                        st.success(f"Acompte {_fmt_money_us(amount)} ajouté.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur lors de l’ajout : {e}")

        # Tableau
        st.subheader("📋 Données")
        cols_show = [c for c in ["ID_Client","Nom","Telephone","Email","Date","Visa","Statut","Montant","Payé","Reste"] if c in f.columns]
        table = f.copy()
        for col in ["Montant","Payé","Reste"]:
            if col in table.columns:
                table[col] = table[col].map(_fmt_money_us)
        if "Date" in table.columns:
            table["Date"] = table["Date"].astype(str)
        st.dataframe(
            table[cols_show].sort_values(by=[c for c in ["Date","Visa","Statut"] if c in table.columns], na_position="last"),
            use_container_width=True
        )

# =========================
# TAB Clients (CRUD)
# =========================
if not (is_ref and sheet_choice.lower() == "visa"):
    with tabs[1]:
        st.subheader("👤 Clients — Créer / Modifier / Supprimer")
        st.caption(f"Feuille cible : **{client_target_sheet}** — toutes les colonnes existantes sont proposées.")
        if st.button("🔄 Rafraîchir"):
            st.rerun()

        # Lire la feuille cible; si absente, on crée un squelette
        orig = read_sheet_from_bytes(current_bytes, client_target_sheet, normalize=False).copy()
        if orig.empty and client_target_sheet not in pd.ExcelFile(io.BytesIO(current_bytes)).sheet_names:
            orig = pd.DataFrame(columns=["ID_Client","Nom","Telephone","Email","Date","Visa","Statut","Montant","Payé","Reste","Paiements"])

        # Colonne technique _RowID pour cibler exactement la bonne ligne
        orig = orig.copy()
        orig["_RowID"] = range(len(orig))

        # Colonnes connues
        bool_cols = [c for c in orig.columns if c.lower() in {"rfe","dossier envoyé","dossier approuvé","dossier refusé","dossier annulé"}]
        date_cols = [c for c in orig.columns if c.lower() in {"date"}]
        money_cols = [c for c in orig.columns if c.lower() in {"honoraires","acomptes","solde","montant","payé","reste"}]
        json_cols  = [c for c in orig.columns if c.lower() in {"paiements"}]

        # Actions
        action = st.radio("Action", ["Créer", "Modifier", "Supprimer"], horizontal=True)

        id_col = "ID_Client" if "ID_Client" in orig.columns else None
        name_col = "Nom" if "Nom" in orig.columns else None
        def _row_label(row):
            parts = []
            if id_col: parts.append(_safe_str(row.get(id_col)))
            if name_col: parts.append(_safe_str(row.get(name_col)))
            if "Telephone" in orig.columns: parts.append(_safe_str(row.get("Telephone")))
            return " — ".join([p for p in parts if p]) or f"Ligne {row.get('_RowID')}"

        # ---------- CRÉER ----------
        if action == "Créer":
            st.markdown("### ➕ Nouveau client")

            if len(orig.columns) == 0:
                orig = pd.DataFrame(columns=["ID_Client","Nom","Telephone","Email","Date","Visa","Statut","Montant","Payé","Reste","Paiements","_RowID"])

            form_values = {}
            with st.form("create_form", clear_on_submit=False):
                for c in [col for col in orig.columns if col != "_RowID"]:
                    label = c
                    if c in bool_cols:
                        form_values[c] = st.checkbox(label, value=False)
                    elif c in date_cols or c == "Date":
                        form_values[c] = st.date_input(label, value=date.today())
                    elif c in money_cols or c in {"Montant","Payé","Reste"}:
                        form_values[c] = st.number_input(label, value=0.0, step=10.0, format="%.2f")
                    elif c in json_cols or c == "Paiements":
                        form_values[c] = st.text_area(label, value="", placeholder='[{"date":"2025-10-03","amount":100}]')
                    else:
                        default_val = "" if c != "Statut" else "Inconnu"
                        form_values[c] = st.text_input(label, value=default_val)
                # ▼▼▼ libellé changé ici ▼▼▼
                submitted = st.form_submit_button("💾 Sauvegarder", type="primary")

            if submitted:
                # Colonnes essentielles si absentes
                essential = ["ID_Client","Nom","Telephone","Date","Montant","Payé","Reste","Paiements"]
                for must in essential:
                    if must not in orig.columns:
                        orig[must] = "" if must not in {"Montant","Payé","Reste"} else 0.0

                # ID si vide
                id_val = _safe_str(form_values.get("ID_Client", ""))
                if not id_val:
                    base_row = {
                        "Nom": form_values.get("Nom",""),
                        "Telephone": form_values.get("Telephone",""),
                        "Date": form_values.get("Date", date.today())
                    }
                    gen_id = _make_client_id_from_row(base_row)
                    existing_ids = set(orig["ID_Client"].astype(str)) if "ID_Client" in orig.columns else set()
                    tmp = gen_id; n=1
                    while tmp in existing_ids:
                        n += 1
                        tmp = f"{gen_id}-{n:02d}"
                    form_values["ID_Client"] = tmp

                # Normaliser types
                new_row = {}
                for c in [col for col in orig.columns if col != "_RowID"]:
                    v = form_values.get(c, "")
                    if c in date_cols or c == "Date":
                        new_row[c] = str(v) if v else str(date.today())
                    elif c in money_cols or c in {"Montant","Payé","Reste"}:
                        new_row[c] = float(v or 0)
                    elif c in bool_cols:
                        new_row[c] = bool(v)
                    elif c in json_cols or c == "Paiements":
                        try:
                            if _safe_str(v):
                                parsed = json.loads(v)
                                if not isinstance(parsed, list): parsed = []
                                new_row[c] = json.dumps(parsed, ensure_ascii=False)
                            else:
                                new_row[c] = ""
                        except Exception:
                            new_row[c] = ""
                    else:
                        new_row[c] = _safe_str(v)

                # Calcul Reste
                if "Montant" in orig.columns and "Payé" in orig.columns:
                    try:
                        m = float(new_row.get("Montant", 0))
                        p = float(new_row.get("Payé", 0))
                        new_row["Reste"] = m - p
                    except Exception:
                        new_row["Reste"] = 0.0

                # append + nouveau _RowID
                new_row["_RowID"] = (orig["_RowID"].max() + 1) if not orig["_RowID"].empty else 0
                orig = pd.concat([orig, pd.DataFrame([new_row])], ignore_index=True)

                try:
                    # Écriture sans la colonne technique
                    to_write = orig.drop(columns=["_RowID"], errors="ignore")
                    updated_bytes = write_updated_excel_bytes(current_bytes, client_target_sheet, to_write)
                    st.session_state["excel_bytes_current"] = updated_bytes
                    st.session_state["pending_sheet_choice"] = client_target_sheet
                    st.session_state["reset_filters_after_write"] = True
                    if source_mode == "Fichier par défaut" and source_id.startswith("path:"):
                        original_path = source_id.split("path:", 1)[1]
                        try:
                            Path(original_path).write_bytes(updated_bytes)
                            st.success(f"Client créé et écrit dans : {original_path}")
                        except Exception as e:
                            st.info(f"Écriture disque impossible. Téléchargez le fichier. Détail: {e}")
                    st.success("✅ Client sauvegardé.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur à l’écriture : {e}")

        # ---------- MODIFIER ----------
        if action == "Modifier":
            st.markdown("### ✏️ Modifier un client")
            if orig.drop(columns=["_RowID"]).empty:
                st.info("Aucun client à modifier.")
            else:
                options = [
                    (int(r["_RowID"]), _row_label(r))
                    for _, r in orig.iterrows()
                ]
                sel_label = st.selectbox("Sélection", [lab for _, lab in options])
                sel_rowid = [rid for rid, lab in options if lab == sel_label][0]
                sel_idx = orig.index[orig["_RowID"] == sel_rowid][0]

                init = orig.loc[sel_idx].to_dict()
                form_values = {}
                with st.form("edit_form", clear_on_submit=False):
                    for c in [col for col in orig.columns if col != "_RowID"]:
                        v = init.get(c, "")
                        if c in bool_cols:
                            form_values[c] = st.checkbox(c, value=bool(v))
                        elif c in date_cols or c == "Date":
                            try:
                                d = pd.to_datetime(v).date() if _safe_str(v) else date.today()
                            except Exception:
                                d = date.today()
                            form_values[c] = st.date_input(c, value=d)
                        elif c in money_cols or c in {"Montant","Payé","Reste"}:
                            try:
                                fv = float(v) if _safe_str(v) else 0.0
                            except Exception:
                                fv = 0.0
                            form_values[c] = st.number_input(c, value=fv, step=10.0, format="%.2f")
                        elif c in json_cols or c == "Paiements":
                            form_values[c] = st.text_area(c, value=_safe_str(v), height=120)
                        else:
                            form_values[c] = st.text_input(c, value=_safe_str(v))
                    submitted = st.form_submit_button("💾 Enregistrer", type="primary")

                if submitted:
                    if "ID_Client" in orig.columns and not _safe_str(form_values.get("ID_Client")):
                        base_row = {"Nom": form_values.get("Nom",""), "Telephone": form_values.get("Telephone",""), "Date": form_values.get("Date", date.today())}
                        form_values["ID_Client"] = _make_client_id_from_row(base_row)

                    for c in [col for col in orig.columns if col != "_RowID"]:
                        v = form_values.get(c, "")
                        if c in date_cols or c == "Date":
                            orig.at[sel_idx, c] = str(v) if v else str(date.today())
                        elif c in money_cols or c in {"Montant","Payé","Reste"}:
                            try: orig.at[sel_idx, c] = float(v or 0)
                            except Exception: orig.at[sel_idx, c] = 0.0
                        elif c in bool_cols:
                            orig.at[sel_idx, c] = bool(v)
                        elif c in json_cols or c == "Paiements":
                            try:
                                if _safe_str(v):
                                    parsed = json.loads(v)
                                    if not isinstance(parsed, list): parsed = []
                                    orig.at[sel_idx, c] = json.dumps(parsed, ensure_ascii=False)
                                else:
                                    orig.at[sel_idx, c] = ""
                            except Exception:
                                orig.at[sel_idx, c] = ""
                        else:
                            orig.at[sel_idx, c] = _safe_str(v)

                    # recalcul Reste
                    if {"Montant","Payé"}.issubset(orig.columns):
                        try:
                            orig.at[sel_idx, "Reste"] = float(orig.at[sel_idx, "Montant"]) - float(orig.at[sel_idx, "Payé"])
                        except Exception:
                            orig.at[sel_idx, "Reste"] = 0.0

                    try:
                        to_write = orig.drop(columns=["_RowID"], errors="ignore")
                        updated_bytes = write_updated_excel_bytes(current_bytes, client_target_sheet, to_write)
                        st.session_state["excel_bytes_current"] = updated_bytes
                        st.session_state["pending_sheet_choice"] = client_target_sheet
                        st.session_state["reset_filters_after_write"] = True
                        if source_mode == "Fichier par défaut" and source_id.startswith("path:"):
                            original_path = source_id.split("path:", 1)[1]
                            try:
                                Path(original_path).write_bytes(updated_bytes)
                                st.success(f"Client modifié et écrit dans : {original_path}")
                            except Exception as e:
                                st.info(f"Écriture disque impossible. Téléchargez le fichier. Détail: {e}")
                        st.success("✅ Modifications enregistrées.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur à l’écriture : {e}")

        # ---------- SUPPRIMER ----------
        if action == "Supprimer":
            st.markdown("### 🗑️ Supprimer un client")
            if orig.drop(columns=["_RowID"]).empty:
                st.info("Aucun client à supprimer.")
            else:
                options = [
                    (int(r["_RowID"]), _row_label(r))
                    for _, r in orig.iterrows()
                ]
                sel_label = st.selectbox("Sélection", [lab for _, lab in options])
                sel_rowid = [rid for rid, lab in options if lab == sel_label][0]
                sel_idx = orig.index[orig["_RowID"] == sel_rowid][0]

                st.error("⚠️ Cette action est irréversible.")
                confirm = st.checkbox("Je confirme la suppression définitive de ce client.")
                if st.button("Supprimer", type="primary", disabled=not confirm):
                    try:
                        orig = orig.drop(index=sel_idx).reset_index(drop=True)
                        to_write = orig.drop(columns=["_RowID"], errors="ignore")
                        updated_bytes = write_updated_excel_bytes(current_bytes, client_target_sheet, to_write)
                        st.session_state["excel_bytes_current"] = updated_bytes
                        st.session_state["pending_sheet_choice"] = client_target_sheet
                        st.session_state["reset_filters_after_write"] = True
                        if source_mode == "Fichier par défaut" and source_id.startswith("path:"):
                            original_path = source_id.split("path:", 1)[1]
                            try:
                                Path(original_path).write_bytes(updated_bytes)
                                st.success(f"Client supprimé et écrit dans : {original_path}")
                            except Exception as e:
                                st.info(f"Écriture disque impossible. Téléchargez le fichier. Détail: {e}")
                        st.success("✅ Client supprimé.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur à l’écriture : {e}")

