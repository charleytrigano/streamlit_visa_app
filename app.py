import streamlit as st
import pandas as pd
import io
from datetime import datetime

# =============================
# Helpers
# =============================

def _pick_sheet(sheet_names, preferred_order=("Visa", "Clients")):
    """Return the first sheet that exists from preferred_order, else the first available."""
    for name in preferred_order:
        if name in sheet_names:
            return name
    return sheet_names[0]

@st.cache_data(show_spinner=False)
def load_data(xlsx_file, preferred_sheet_order=("Visa", "Clients")):
    """Load a sheet from an Excel file-like object with robust sheet selection.

    Returns (df, used_sheet_name, all_sheet_names)
    """
    xls = pd.ExcelFile(xlsx_file)
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
    st.markdown("**Astuce** : si le nom de l'onglet change, l'application basculera automatiquement sur le premier onglet disponible.")

if not up:
    st.info("Chargez un fichier Excel (.xlsx) dans la barre latérale pour commencer.")
    st.stop()

# Chargement robuste
preferred = ("Visa", "Clients") if prefer_visa else ("Clients", "Visa")
df, used_sheet, sheet_names = load_data(up, preferred)
all_sheets, _ = load_all_sheets(up)

st.success(f"✅ Onglet utilisé : **{used_sheet}** · Onglets trouvés : {', '.join(sheet_names)}")

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

