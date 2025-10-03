import io
import json
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
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _to_date(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

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

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Uniformise Date/Année/Mois/Visa/Statut/Montant/Payé/Reste si c'est un tableau 'dossiers'."""
    df = df.copy()

    if "Date" in df.columns:
        df["Date"] = _to_date(df["Date"])
    else:
        df["Date"] = pd.NaT

    if "Année" not in df.columns:
        df["Année"] = df["Date"].dt.year
    if "Mois" not in df.columns:
        df["Mois"] = df["Date"].dt.to_period("M").astype(str)

    visa_col = _first_col(df, ["Visa", "Categories", "Catégorie", "TypeVisa"])
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"

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

    # Paiements (JSON) -> TotalAcomptes
    if "Paiements" in df.columns:
        parsed = df["Paiements"].apply(_parse_paiements)
        df["TotalAcomptes"] = parsed.apply(_sum_payments)

    # Payé
    if "Payé" in df.columns:
        df["Payé"] = _to_num(df["Payé"])
    else:
        src_paye = _first_col(df, ["TotalAcomptes", "Acomptes", "Paye", "Paid"])
        df["Payé"] = _to_num(df[src_paye]) if src_paye else 0.0

    # Reste
    if "Reste" in df.columns:
        df["Reste"] = _to_num(df["Reste"])
    else:
        src_reste = _first_col(df, ["Solde", "SoldeCalc"])
        if src_reste:
            df["Reste"] = _to_num(df[src_reste])
        else:
            df["Reste"] = (df["Montant"] - df["Payé"]).fillna(0.0)

    return df

def looks_like_reference(df: pd.DataFrame) -> bool:
    """Détecte un onglet de référence (ex: 'Visa' avec Categories/Visa/Definition)."""
    cols = set(map(str.lower, df.columns.astype(str)))
    has_ref = {"categories", "visa"} <= cols
    no_money = not ({"montant", "honoraires", "acomptes", "payé", "reste", "solde"} & cols)
    return has_ref and no_money

def write_updated_excel_bytes(original_bytes: bytes, sheet_to_replace: str, new_df: pd.DataFrame) -> bytes:
    """Recharge toutes les feuilles, remplace sheet_to_replace par new_df, renvoie les nouveaux octets Excel."""
    xls = pd.ExcelFile(io.BytesIO(original_bytes))
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name in xls.sheet_names:
            if name == sheet_to_replace:
                # cast en str pour éviter des types 'object' non éditables
                dfw = new_df.copy()
                for c in dfw.columns:
                    if dfw[c].dtype == "object":
                        dfw[c] = dfw[c].astype(str).fillna("")
                dfw.to_excel(writer, sheet_name=name, index=False)
            else:
                pd.read_excel(xls, sheet_name=name).to_excel(writer, sheet_name=name, index=False)
    out.seek(0)
    return out.read()

# ---------------- Cache sérialisable ----------------
@st.cache_data
def load_excel_bytes(xlsx_input):
    """
    Retourne (sheet_names, data_bytes, source_id) pour un chemin ou un fichier uploadé.
    """
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

    # charge la feuille complète depuis les OCTETS COURANTS
    full_ref_df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=False).copy()

    # Colonnes par défaut
    default_cols = ["Categories", "Visa", "Definition"]
    for c in default_cols:
        if c not in full_ref_df.columns:
            full_ref_df[c] = ""

    # Ordonne : d'abord les colonnes par défaut, puis les autres
    ordered_cols = [c for c in default_cols if c in full_ref_df.columns] + [c for c in full_ref_df.columns if c not in default_cols]
    full_ref_df = full_ref_df[ordered_cols]

    # Forcer string pour rendre l'édition plus fluide
    for c in full_ref_df.columns:
        if full_ref_df[c].dtype != "object":
            full_ref_df[c] = full_ref_df[c].astype(str)
        full_ref_df[c] = full_ref_df[c].fillna("")

    # ÉDITEUR : clé dédiée + lignes dynamiques
    edited_df = st.data_editor(
        full_ref_df,
        num_rows="dynamic",       # permet d'ajouter/supprimer des lignes
        use_container_width=True,
        hide_index=True,
        key="visa_editor",        # clé nécessaire pour bien gérer l'état d'édition
    )

    st.divider()
    col_save, col_dl, col_reset = st.columns([1,1,1])

    # Sauvegarde → met à jour les octets courants + (si possible) le fichier d'origine
    if col_save.button("💾 Enregistrer (remplace la feuille 'Visa')", type="primary"):
        try:
            updated_bytes = write_updated_excel_bytes(current_bytes, sheet_choice, edited_df)
            st.session_state["excel_bytes_current"] = updated_bytes  # ← met à jour la SOURCE courante
            # Si la source est un fichier par défaut sur disque, on tente d'écrire dessus
            if source_mode == "Fichier par défaut" and source_id.startswith("path:"):
                original_path = source_id.split("path:", 1)[1]
                try:
                    Path(original_path).write_bytes(updated_bytes)
                    st.success(f"Fichier mis à jour sur le disque : {original_path}")
                except Exception as e:
                    st.info(f"Impossible d’écrire sur le disque. Télécharge le fichier ci-dessous. Détail: {e}")
            st.success("Modifications enregistrées.")
            st.rerun()  # ← très important : on relit la version à jour
        except Exception as e:
            st.error(f"Erreur à l’enregistrement : {e}")

    # Télécharger l’Excel mis à jour (depuis l'état courant)
    col_dl.download_button(
        "⬇️ Télécharger l’Excel mis à jour",
        data=st.session_state["excel_bytes_current"],
        file_name="visa_mis_a_jour.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Réinitialiser (revient aux octets de la lecture initiale de la session)
    if col_reset.button("↩️ Réinitialiser (annuler les modifs non enregistrées)"):
        st.session_state.pop("excel_bytes_current", None)  # on efface l'état courant
        # recharge depuis la source initiale
        st.session_state["excel_bytes_current"] = data_bytes
        st.success("Réinitialisé.")
        st.rerun()

    st.stop()

# ---------------- MODE DOSSIERS ----------------
# si ce n'est pas une table de référence 'Visa', on normalise et on affiche KPIs + graph + table
df = read_sheet_from_bytes(current_bytes, sheet_choice, normalize=True)

with st.container():
    c1, c2, c3 = st.columns(3)
    years = sorted([int(y) for y in df["Année"].dropna().unique()]) if "Année" in df else []
    visas = sorted(df["Visa"].dropna().astype(str).unique())
    statuses = sorted(df["Statut"].dropna().astype(str).unique())

    year_sel = c1.multiselect("Années", years, default=years or None)
    visa_sel = c2.multiselect("Type de visa", visas, default=visas or None)
    stat_sel = c3.multiselect("Statut", statuses, default=statuses or None)

f = df.copy()
if year_sel:
    f = f[f["Année"].isin(year_sel)]
if visa_sel:
    f = f[f["Visa"].astype(str).isin(visa_sel)]
if stat_sel:
    f = f[f["Statut"].astype(str).isin(stat_sel)]

k1, k2, k3, k4 = st.columns(4)
k1.metric("Dossiers", f"{len(f)}")
k2.metric("Montant total", f"{f['Montant'].sum():,.2f} €")
k3.metric("Payé", f"{f['Payé'].sum():,.2f} €")
k4.metric("Reste", f"{f['Reste'].sum():,.2f} €")

st.divider()

st.subheader("📈 Nombre de dossiers par mois")
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
cols_show = [c for c in ["Date","Année","Mois","Visa","Statut","Montant","Payé","Reste"] if c in f.columns]
st.dataframe(
    f[cols_show].sort_values(by=[c for c in ["Date","Visa","Statut"] if c in f.columns], na_position="last"),
    use_container_width=True
)

st.caption("Dans l’onglet **Visa**, tu peux ajouter/modifier/supprimer des lignes, puis **Enregistrer**. "
           "Le fichier courant est mis à jour et tu peux le **Télécharger**.")





