import io
import json
import hashlib
from pathlib import Path
import streamlit as st
import pandas as pd

st.set_page_config(page_title="📊 Visas — Simplifié", layout="wide")
st.title("📊 Visas — Tableau simplifié")

# ---------------- Utils ----------------
def _first_col(df: pd.DataFrame, candidates) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def _to_num(s: pd.Series) -> pd.Series:
    # Nettoie "1 234,56 €" → 1234.56
    cleaned = (
        s.astype(str)
         .str.replace("\u00a0", "", regex=False)   # espace insécable
         .str.replace("\u202f", "", regex=False)   # fine space
         .str.replace(" ", "", regex=False)
         .str.replace("€", "", regex=False)
         .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)

def _to_date(s: pd.Series) -> pd.Series:
    d = pd.to_datetime(s, errors="coerce")
    # enlève les fuseaux si présents et supprime l'heure
    try:
        d = d.dt.tz_localize(None)
    except Exception:
        pass
    # tronque à la date (00:00), puis renvoie type date
    return d.dt.normalize().dt.date

def _safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def _parse_paiements(x):
    if isinstance(x, list):
        return x
    if pd.isna(x):
        return []
    try:
        return json.loads(x)
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
    """
    ID stable depuis Nom + Telephone + Date (Email ignoré).
    Exemple: CL-2F9A3C81
    """
    base = "|".join([
        _safe_str(row.get("Nom")),
        _safe_str(row.get("Telephone")),
        _safe_str(row.get("Date")),
    ])
    h = hashlib.sha1(base.encode("utf-8")).hexdigest()[:8].upper()
    return f"CL-{h}"

def _dedupe_ids(series: pd.Series) -> pd.Series:
    """
    Si des IDs identiques existent (même hash pour lignes similaires),
    ajoute un suffixe -01, -02, ...
    """
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

    # --- Date (sans heure) ---
    if "Date" in df.columns:
        df["Date"] = _to_date(df["Date"])
    else:
        df["Date"] = pd.NaT

    # --- Mois (MM) pour les regroupements internes (non affiché) ---
    df["Mois"] = df["Date"].apply(lambda x: f"{x.month:02d}" if pd.notna(x) else pd.NA)

    # --- Visa / Categories
    visa_col = _first_col(df, ["Visa", "Categories", "Catégorie", "TypeVisa"])
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"

    # --- Statut
    if "__Statut règlement__" in df.columns and "Statut" not in df.columns:
        df = df.rename(columns={"__Statut règlement__": "Statut"})
    if "Statut" not in df.columns:
        df["Statut"] = "Inconnu"
    else:
        df["Statut"] = df["Statut"].astype(str).fillna("Inconnu")

    # --- Montant
    if "Montant" in df.columns:
        df["Montant"] = _to_num(df["Montant"])
    else:
        src_montant = _first_col(df, ["Honoraires", "Total", "Amount"])
        df["Montant"] = _to_num(df[src_montant]) if src_montant else 0.0

    # --- Paiements (JSON) -> TotalAcomptes
    if "Paiements" in df.columns:
        parsed = df["Paiements"].apply(_parse_paiements)
        df["TotalAcomptes"] = parsed.apply(_sum_payments)

    # --- Payé
    if "Payé" in df.columns:
        df["Payé"] = _to_num(df["Payé"])
    else:
        src_paye = _first_col(df, ["TotalAcomptes", "Acomptes", "Paye", "Paid"])
        df["Payé"] = _to_num(df[src_paye]) if src_paye else 0.0

    # --- Reste (toujours calculé)
    df["Reste"] = (df["Montant"] - df["Payé"]).fillna(0.0)

    # --- ID client auto (Nom + Telephone + Date)
    if "ID_Client" not in df.columns:
        df["ID_Client"] = ""
    need_id = df["ID_Client"].astype(str).str.strip().eq("") | df["ID_Client"].isna()
    if need_id.any():
        generated = df.loc[need_id].apply(_make_client_id_from_row, axis=1)
        df.loc[need_id, "ID_Client"] = generated
    df["ID_Client"] = _dedupe_ids(df["ID_Client"].astype(str).str.strip())

    return df

def write_updated_excel_bytes(original_bytes: bytes, sheet_to_replace: str, new_df: pd.DataFrame) -> bytes:
    """Recharge toutes les feuilles, remplace sheet_to_replace par new_df, renvoie les nouveaux octets Excel."""
    xls = pd.ExcelFile(io.BytesIO(original_bytes))
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name in xls.sheet_names:
            if name == sheet_to_replace:
                dfw = new_df.copy()
                for c in dfw.columns:
                    if dfw[c].dtype == "object":
                        dfw[c] = dfw[c].astype(str).fillna("")
                dfw.to_excel(writer, sheet_name=name, index=False)
            else:
                pd.read_excel(xls, sheet_name=name).to_excel(writer, sheet_name=name, index=False)
    out.seek(0)
    return out.read()

# ---- Référence: map Categories -> Visa
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

def enrich_visa_from_categories(df: pd.DataFrame, cat2visa: dict) -> tuple[pd.DataFrame, int]:
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

# ---------------- Cache sérialisable ----------------
@st.cache_data
def load_excel_bytes(xlsx_input):
    """Retourne (sheet_names, data_bytes, source_id) pour un chemin ou un fichier uploadé."""
    if hasattr(xlsx_input, "read"):  # UploadedFile
        data = xlsx_input.read()
        src_id = f"upload:{getattr(xlsx_input, 'name', 'uploaded')}"
    else:  # chemin
        data = Path(xlsx_input).read_bytes()
        src_id = f"path:{xlsx_input}"
    xls = pd.ExcelFile(io.BytesIO(data))
    return xls.sheet_names, data, src_id

def read_sheet_from_bytes(data_bytes: bytes, sheet_name: str, normalize: bool) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(data_bytes))
    df = pd.read_excel(xls, sheet_name=sheet_name)
    if normalize and not looks_like_reference(df):
        df = normalize_dataframe(df)
    return df

# ---------------- Source & Sélection feuille ----------------
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

# Initialise/actualise l'état courant des octets selon la source
if "excel_bytes_current" not in st.session_state or st.session_state.get("excel_source_id") != source_id:
    st.session_state["excel_bytes_current"] = data_bytes
    st.session_state["excel_source_id"] = source_id

current_bytes = st.session_state["excel_bytes_current"]

# Choix explicite de la feuille (inclut 'Visa')
preferred_order = ["Données normalisées", "Clients", "Visa"]
default_sheet = next((s for s in preferred_order if s in sheet_names), sheet_names[0])
sheet_choice = st.sidebar.selectbox("Feuille", sheet_names, index=sheet_names.index(default_sheet), key="sheet_choice")

# Lecture de la feuille (échantillon pour la détection)
sample_df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=False).head(5)
is_ref = looks_like_reference(sample_df)

# ---------------- MODE RÉFÉRENCE — ÉDITION VISA ----------------
if is_ref and sheet_choice.lower() == "visa":
    st.subheader("📚 Table de référence — Visa")
    st.caption("Ajoute / modifie / supprime directement ci-dessous. Les lignes et cellules sont éditables. "
               "Utilise le bouton **+ Ajouter une ligne** en bas du tableau, ou le menu ⋮ pour supprimer.")

    # Feuille complète depuis les octets courants
    full_ref_df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=False).copy()

    # Colonnes par défaut
    default_cols = ["Categories", "Visa", "Definition"]
    for c in default_cols:
        if c not in full_ref_df.columns:
            full_ref_df[c] = ""

    # Ordre colonnes
    ordered_cols = [c for c in default_cols if c in full_ref_df.columns] + [c for c in full_ref_df.columns if c not in default_cols]
    full_ref_df = full_ref_df[ordered_cols]

    # Forcer string
    for c in full_ref_df.columns:
        if full_ref_df[c].dtype != "object":
            full_ref_df[c] = full_ref_df[c].astype(str)
        full_ref_df[c] = full_ref_df[c].fillna("")

    # Éditeur
    edited_df = st.data_editor(
        full_ref_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="visa_editor",
    )

    st.divider()
    col_save, col_dl, col_reset = st.columns([1,1,1])

    if col_save.button("💾 Enregistrer (remplace la feuille 'Visa')", type="primary"):
        try:
            updated_bytes = write_updated_excel_bytes(current_bytes, sheet_choice, edited_df)
            st.session_state["excel_bytes_current"] = updated_bytes
            if source_mode == "Fichier par défaut" and source_id.startswith("path:"):
                original_path = source_id.split("path:", 1)[1]
                try:
                    Path(original_path).write_bytes(updated_bytes)
                    st.success(f"Fichier mis à jour sur le disque : {original_path}")
                except Exception as e:
                    st.info(f"Impossible d’écrire sur le disque. Télécharge le fichier ci-dessous. Détail: {e}")
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

    if col_reset.button("↩️ Réinitialiser (annuler les modifs non enregistrées)"):
        st.session_state.pop("excel_bytes_current", None)
        st.session_state["excel_bytes_current"] = data_bytes
        st.success("Réinitialisé.")
        st.rerun()

    st.stop()

# ---------------- MODE DOSSIERS ----------------
# 1) lecture normalisée
df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=True)

# 2) jointure auto Categories -> Visa (depuis l'onglet Visa si présent)
cat2visa = build_categories_to_visa_map(current_bytes, visa_sheet_name="Visa")
df_enriched, nb_filled = enrich_visa_from_categories(df, cat2visa)

# si on a enrichi, propose d'écrire dans l'Excel
if nb_filled > 0 and "Visa" in df_enriched.columns and sheet_choice.lower() in ["clients", "données normalisées", "donnees normalisees"]:
    st.info(f"🔁 {nb_filled} valeur(s) 'Visa' complétée(s) depuis 'Categories' grâce à l'onglet de référence **Visa**.")
    cols_top = st.columns([1,1])
    if cols_top[0].button("💾 Écrire les 'Visa' complétés dans l’Excel", type="primary"):
        try:
            original_df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=False)
            if "Visa" not in original_df.columns:
                original_df["Visa"] = ""
            # on écrit sur les mêmes indices
            original_df.loc[df_enriched.index, "Visa"] = df_enriched["Visa"].values
            updated_bytes = write_updated_excel_bytes(current_bytes, sheet_choice, original_df)
            st.session_state["excel_bytes_current"] = updated_bytes
            if source_mode == "Fichier par défaut" and source_id.startswith("path:"):
                original_path = source_id.split("path:", 1)[1]
                try:
                    Path(original_path).write_bytes(updated_bytes)
                    st.success(f"✅ Écrit dans le fichier : {original_path}")
                except Exception as e:
                    st.info(f"Impossible d’écrire sur le disque. Télécharge le fichier mis à jour ci-dessous. Détail: {e}")
            st.success("✅ 'Visa' complétés enregistrés.")
            st.rerun()
        except Exception as e:
            st.error(f"Erreur à l’écriture : {e}")

    cols_top[1].download_button(
        "⬇️ Télécharger l’Excel avec 'Visa' complétés",
        data=st.session_state.get("excel_bytes_current", current_bytes),
        file_name="clients_visa_completes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# 3) Remplace df par la version enrichie pour le tableau de bord
df = df_enriched

# --- Détection des ID ajoutés et écriture dans l’Excel (Clients / Données normalisées) ---
if sheet_choice.lower() in ["clients", "données normalisées", "donnees normalisees"]:
    original_df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=False)
    had_id_col = "ID_Client" in original_df.columns
    orig_ids = original_df["ID_Client"].astype(str).str.strip() if had_id_col else pd.Series([""] * len(original_df))

    if len(orig_ids) == len(df):
        missing_before = orig_ids.eq("") | orig_ids.isna()
        new_ids = df["ID_Client"].astype(str).str.strip()
        newly_filled = (missing_before) & new_ids.ne("")

        if newly_filled.any():
            st.info(f"🆔 {newly_filled.sum()} ID_Client ont été générés automatiquement.")
            cols_id = st.columns([1,1])
            if cols_id[0].button("💾 Écrire les ID_Client dans l’Excel", type="primary"):
                try:
                    original_upd = original_df.copy()
                    if "ID_Client" not in original_upd.columns:
                        original_upd["ID_Client"] = ""
                    original_upd.loc[newly_filled.index, "ID_Client"] = df.loc[newly_filled.index, "ID_Client"].values
                    updated_bytes = write_updated_excel_bytes(current_bytes, sheet_choice, original_upd)
                    st.session_state["excel_bytes_current"] = updated_bytes
                    if source_mode == "Fichier par défaut" and source_id.startswith("path:"):
                        original_path = source_id.split("path:", 1)[1]
                        try:
                            Path(original_path).write_bytes(updated_bytes)
                            st.success(f"✅ Écrit dans le fichier : {original_path}")
                        except Exception as e:
                            st.info(f"Impossible d’écrire sur le disque. Télécharge le fichier mis à jour ci-dessous. Détail: {e}")
                    st.success("✅ ID_Client enregistrés dans l’Excel.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur à l’écriture : {e}")

            cols_id[1].download_button(
                "⬇️ Télécharger l’Excel avec ID_Client",
                data=st.session_state.get("excel_bytes_current", current_bytes),
                file_name="clients_avec_id.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# --- Filtres & KPIs & affichages dossiers ---
with st.container():
    c1, c2, c3 = st.columns(3)
    # Année/Mois n'apparaissent plus à l'écran
    years = sorted([int(y) for y in []])  # placeholder pour garder la structure si besoin plus tard
    visas = sorted(df["Visa"].dropna().astype(str).unique()) if "Visa" in df.columns else []
    statuses = sorted(df["Statut"].dropna().astype(str).unique()) if "Statut" in df.columns else []

    # Filtres disponibles : Visa / Statut seulement (Date reste visible mais sans heure)
    visa_sel = c1.multiselect("Type de visa", visas, default=visas or None)
    stat_sel = c2.multiselect("Statut", statuses, default=statuses or None)
    # c3 laissé vide pour équilibre visuel

f = df.copy()
if "Visa" in f.columns and visa_sel:
    f = f[f["Visa"].astype(str).isin(visa_sel)]
if "Statut" in f.columns and stat_sel:
    f = f[f["Statut"].astype(str).isin(stat_sel)]

k1, k2, k3, k4 = st.columns(4)
k1.metric("Dossiers", f"{len(f)}")
k2.metric("Montant total", f"{f['Montant'].sum():,.2f} €" if "Montant" in f.columns else "—")
k3.metric("Payé", f"{f['Payé'].sum():,.2f} €" if "Payé" in f.columns else "—")
k4.metric("Reste", f"{f['Reste'].sum():,.2f} €" if "Reste" in f.columns else "—")

st.divider()

# Graphique simple par Mois (MM). ATTENTION: agrège tous les mois de l'année ensemble (janv de toutes années).
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

st.subheader("📋 Données")
# Année/Mois ne sont pas affichées
cols_show = [c for c in ["ID_Client","Nom","Telephone","Email","Date","Visa","Statut","Montant","Payé","Reste"] if c in f.columns]
table = f.copy()

# Formatage propre montants + date str
def _fmt_money_col(df, name):
    if name in df.columns:
        df[name] = df[name].map(lambda v: f"{v:,.2f} €".replace(",", " ").replace(".", ","))
_fmt_money_col(table, "Montant")
_fmt_money_col(table, "Payé")
_fmt_money_col(table, "Reste")
if "Date" in table.columns:
    table["Date"] = table["Date"].astype(str)  # YYYY-MM-DD sans heure

st.dataframe(
    table[cols_show].sort_values(by=[c for c in ["Date","Visa","Statut"] if c in table.columns], na_position="last"),
    use_container_width=True
)

st.caption("• Mois = MM (non affiché), Date sans heure • Reste = Montant − Payé • Onglet Visa éditable • ID_Client auto (Nom + Telephone + Date) • Jointure Categories→Visa.")
