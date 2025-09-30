import streamlit as st
import pandas as pd
import io
from datetime import datetime

# =============================
# Clear cache via URL param ?clear=1 (optionnel)
# =============================
try:
    params = st.query_params
    if params.get("clear", ["0"])[0] == "1":
        st.cache_data.clear()
        try:
            st.cache_resource.clear()
        except Exception:
            pass
        # Nettoie les params pour éviter les boucles
        st.experimental_set_query_params()
        st.rerun()
except Exception:
    pass

# =============================
# Helpers
# =============================

def _find_col(possible_names, columns):
    """Return the first matching column name (case-insensitive, accent-insensitive-ish)."""
    # simple normalization: lower + remove accents-like chars
    import unicodedata
    def norm(s: str) -> str:
        s = str(s)
        s = ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
        return s.lower().strip()
    cols_norm = {norm(c): c for c in columns}
    for name in possible_names:
        key = norm(name)
        if key in cols_norm:
            return cols_norm[key]
    return None


def _as_bool_series(s):
    """Coerce a pandas Series with checkbox-like values to boolean (True/False)."""
    import numpy as np
    if s is None:
        return pd.Series([], dtype=bool)
    vals = s.astype(str).str.strip().str.lower()
    truthy = {"1", "true", "vrai", "yes", "oui", "y", "o", "x", "✓", "checked"}
    falsy = {"0", "false", "faux", "no", "non", "n", "", "none", "nan"}
    out = vals.apply(lambda v: True if v in truthy else (False if v in falsy else np.nan))
    return out.fillna(False)


def _to_numeric(s):
    return pd.to_numeric(s, errors="coerce")

def _pick_sheet(sheet_names, preferred_order=("Visa", "Clients")):
    """Return the first sheet that exists from preferred_order, else the first available."""
    for name in preferred_order:
        if name in sheet_names:
            return name
    return sheet_names[0]

@st.cache_data(show_spinner=False)
def load_data(xlsx_input, preferred_sheet_order=("Visa", "Clients")):
    """Load a sheet from an Excel *path or file-like* with robust sheet selection.

    xlsx_input can be a str path (e.g., "C:/.../file.xlsx") or an uploaded file.
    Returns (df, used_sheet_name, all_sheet_names)
    """
    xls = pd.ExcelFile(xlsx_input)
    sheet_names = xls.sheet_names
    used_sheet = _pick_sheet(sheet_names, preferred_sheet_order)
    df = pd.read_excel(xls, sheet_name=used_sheet)
    # Normalize column names: strip spaces/newlines
    df.columns = [str(c).strip() for c in df.columns]
    return df, used_sheet, sheet_names

@st.cache_data(show_spinner=False)
def load_all_sheets(xlsx_file):
    """Load all sheets into a dict {sheet_name: DataFrame} with normalized columns."""
    xls = pd.ExcelFile(xlsx_file)
    out = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        df.columns = [str(c).strip() for c in df.columns]
        out[name] = df
    return out, xls.sheet_names

@st.cache_data(show_spinner=False)
def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Feuille1") -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()

# =============================
# UI
# =============================
st.set_page_config(
    page_title="Visa App",
    page_icon="🛂",
    layout="wide",
)

st.title("🛂 Visa App — Excel → analyse & export")
st.caption("Sélection automatique de l'onglet **Visa** ou **Clients** si disponible.")

with st.sidebar:
    st.header("Importer votre Excel")
    up = st.file_uploader("Fichier .xlsx", type=["xlsx"], help="Choisissez le classeur contenant les onglets 'Visa' et/ou 'Clients'.")
    prefer_visa = st.toggle("Préférer l'onglet 'Visa' s'il existe", value=True)

    st.divider()
    if st.button("♻️ Vider le cache et recharger", use_container_width=True):
        st.cache_data.clear()
        try:
            st.cache_resource.clear()
        except Exception:
            pass
        st.success("Cache vidé. Rechargement…")
        st.rerun()

    st.markdown("**Astuce** : vous pouvez aussi ajouter `?clear=1` à l'URL pour vider le cache au chargement.")

if not up:
    st.info("Chargez un fichier Excel (.xlsx) dans la barre latérale pour commencer.")
    st.stop()

# Chargement robuste
# Autoriser un chemin local en plus de l'upload
col_path, col_hint = st.columns([3,2])
with col_path:
    data_path = st.text_input("Ou saisissez un chemin local vers le .xlsx (optionnel)", value="", help="Exemple: C:/Users/charl/Desktop/visa_app/data.xlsx")
with col_hint:
    st.caption("Si un chemin est fourni, il sera prioritaire sur l'upload.")

preferred = ("Visa", "Clients") if prefer_visa else ("Clients", "Visa")

src = data_path if data_path.strip() else up
if not src:
    st.info("Chargez un fichier ou renseignez un chemin local pour continuer.")
    st.stop()

try:
    df, used_sheet, sheet_names = load_data(src, preferred)
    all_sheets, _ = load_all_sheets(src)
except ValueError as e:
    st.error(f"Erreur lors de la lecture du classeur : {e}")
    st.stop()

st.success(f"✅ Onglet utilisé : **{used_sheet}** · Onglets trouvés : {', '.join(sheet_names)}")

# =============================
# Dashboard — Clients (si l'onglet existe)
# =============================
if "Clients" in all_sheets:
    st.subheader("📊 Dashboard — Clients")
    clients_df = all_sheets["Clients"].copy()

    # Colonnes cibles (avec variantes possibles)
    col_type = _find_col(["Type Visa", "Type", "Visa"], clients_df.columns) or ""
    col_hon = _find_col(["Honoraires", "Frais", "Montant"], clients_df.columns)
    col_solde = _find_col(["Solde"], clients_df.columns)
    col_envoye = _find_col(["Dossier envoyé", "Dossier envoye", "Envoye", "Envoyé"], clients_df.columns)
    col_refuse = _find_col(["Dossier refusé", "Dossier refuse", "Refuse", "Refusé"], clients_df.columns)
    col_approuve = _find_col(["Dossier approuvé", "Dossier approuve", "Approuve", "Approuvé"], clients_df.columns)

    sent = _as_bool_series(clients_df[col_envoye]) if col_envoye else pd.Series([False]*len(clients_df))
    refused = _as_bool_series(clients_df[col_refuse]) if col_refuse else pd.Series([False]*len(clients_df))
    approved = _as_bool_series(clients_df[col_approuve]) if col_approuve else pd.Series([False]*len(clients_df))

    hon_total = _to_numeric(clients_df[col_hon]).sum() if col_hon else 0.0
    solde_total = _to_numeric(clients_df[col_solde]).sum() if col_solde else 0.0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Dossiers", f"{len(clients_df):,}")
    c2.metric("Envoyés", f"{int(sent.sum()):,}")
    c3.metric("Approuvés", f"{int(approved.sum()):,}")
    c4.metric("Refusés", f"{int(refused.sum()):,}")
    c5.metric("Honoraires (Σ)", f"{hon_total:,.2f}")

    if col_solde:
        st.caption(f"Solde (Σ): {solde_total:,.2f}")

    # Graphique de répartition par Type Visa (top 15)
    if col_type:
        st.markdown("**Répartition par type de visa**")
        counts = clients_df[col_type].astype(str).str.strip().value_counts().head(15)
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots()
        ax.bar(range(len(counts)), counts.values)
        ax.set_xticks(range(len(counts)))
        ax.set_xticklabels(list(counts.index), rotation=45, ha="right")
        ax.set_xlabel(col_type)
        ax.set_ylabel("Occurrences")
        ax.set_title("Top valeurs — Type de visa")
        fig.tight_layout()
        st.pyplot(fig)
    else:
        st.info("Colonne 'Type Visa' non trouvée : le graphique de répartition est masqué.")

# =============================
# Tableau + Filtres
# =============================

st.subheader("Aperçu & filtres")

# Recherche plein-texte simple
col_search, col_rows = st.columns([3,1])
with col_search:
    q = st.text_input("Recherche (plein-texte sur toutes les colonnes)", placeholder="Tapez un mot-clé…")
with col_rows:
    max_rows = st.number_input("Lignes à afficher", min_value=5, max_value=5000, value=100, step=5)

filtered = df.copy()
if q:
    mask = pd.Series(False, index=filtered.index)
    for c in filtered.columns:
        try:
            mask = mask | filtered[c].astype(str).str.contains(q, case=False, na=False)
        except Exception:
            pass
    filtered = filtered[mask]

# Filtres par colonne (catégorielles seulement)
with st.expander("Filtres par colonne (catégories)"):
    for col in filtered.select_dtypes(include=["object", "category"]).columns:
        unique_vals = sorted([v for v in filtered[col].dropna().unique() if v != ""], key=lambda x: str(x).lower())
        if 1 < len(unique_vals) <= 1000:
            sel = st.multiselect(f"{col}", unique_vals, default=None)
            if sel:
                filtered = filtered[filtered[col].isin(sel)]

# Stats rapides
st.markdown(
    f"**{len(filtered):,}** lignes affichées (sur **{len(df):,}**), **{len(df.columns)}** colonnes."
)

st.dataframe(filtered.head(int(max_rows)), use_container_width=True)

# =============================
# Exports
# =============================

st.subheader("Exports")
col1, col2 = st.columns(2)
filename_stamp = datetime.now().strftime("%Y%m%d-%H%M%S")

with col1:
    csv_bytes = filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Télécharger CSV (filtré)",
        data=csv_bytes,
        file_name=f"{used_sheet}_filtre_{filename_stamp}.csv",
        mime="text/csv",
        use_container_width=True,
    )
with col2:
    xls_bytes = to_excel_bytes(filtered, sheet_name=used_sheet)
    st.download_button(
        "⬇️ Télécharger Excel (filtré)",
        data=xls_bytes,
        file_name=f"{used_sheet}_filtre_{filename_stamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# =============================
# Onglet secondaire : Clients / Visa si présents
# =============================

st.divider()
other_tabs = [name for name in ("Visa", "Clients") if name in all_sheets and name != used_sheet]
if other_tabs:
    st.subheader("Autre onglet disponible")
    sel = st.selectbox("Afficher l'autre onglet :", options=other_tabs)
    df_other = all_sheets[sel]
    st.caption(f"Aperçu de l'onglet **{sel}**")
    st.dataframe(df_other.head(200), use_container_width=True)
    other_csv = df_other.to_csv(index=False).encode("utf-8")
    st.download_button(
        f"⬇️ Télécharger CSV — {sel}",
        data=other_csv,
        file_name=f"{sel}_{filename_stamp}.csv",
        mime="text/csv",
    )
else:
    st.caption("Aucun autre onglet à afficher.")

# =============================
# Notes
# =============================
with st.expander("Aide / Dépannage"):
    st.markdown(
        """
        - Si vous voyez une erreur de type **Worksheet not found**, renommez l'onglet ou laissez l'app choisir automatiquement l'onglet disponible.
        - Cette app normalise simplement les noms de colonnes (suppression d'espaces inutiles). Aucun mapping de colonnes n'est imposé.
        - Les filtres par colonnes s'appliquent aux colonnes de type texte/catégorie.
        - L'export Excel utilise **openpyxl**; si besoin, installez `pip install openpyxl`.
        """
    )
