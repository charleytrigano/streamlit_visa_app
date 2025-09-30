# app.py
import io
from datetime import datetime
from typing import Tuple, List, Dict

import pandas as pd
import streamlit as st


# =============================
# Clear cache via URL param ?clear=1 (optionnel)
# =============================
try:
    params = st.experimental_get_query_params()
    if params.get("clear", ["0"])[0] == "1":
        st.cache_data.clear()
        try:
            st.cache_resource.clear()
        except Exception:
            pass
        # Nettoie le paramètre pour éviter les boucles de rerun
        st.experimental_set_query_params()
        st.rerun()
except Exception:
    pass


# =============================
# Helpers
# =============================
def _pick_sheet(sheet_names: List[str], preferred_order=("Visa", "Clients")) -> str:
    """Retourne la première feuille existante parmi preferred_order, sinon la première feuille."""
    for name in preferred_order:
        if name in sheet_names:
            return name
    return sheet_names[0]


@st.cache_data(show_spinner=False)
def load_data(xlsx_input, preferred_sheet_order=("Visa", "Clients")) -> Tuple[pd.DataFrame, str, List[str]]:
    """
    Charge un onglet depuis un classeur Excel (chemin local OU fichier uploadé Streamlit).
    Retourne: (df, used_sheet, sheet_names)
    """
    xls = pd.ExcelFile(xlsx_input)
    sheet_names = xls.sheet_names
    used_sheet = _pick_sheet(sheet_names, preferred_sheet_order)
    df = pd.read_excel(xls, sheet_name=used_sheet)
    # Normalisation légère des noms de colonnes
    df.columns = [str(c).strip() for c in df.columns]
    return df, used_sheet, sheet_names


@st.cache_data(show_spinner=False)
def load_all_sheets(xlsx_input) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    """Charge toutes les feuilles dans un dict {nom: DataFrame} avec colonnes normalisées."""
    xls = pd.ExcelFile(xlsx_input)
    out = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        df.columns = [str(c).strip() for c in df.columns]
        out[name] = df
    return out, xls.sheet_names


@st.cache_data(show_spinner=False)
def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Feuille1") -> bytes:
    """Convertit un DataFrame en bytes Excel (XLSX) pour download_button."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()


# =============================
# UI de la page
# =============================
st.set_page_config(page_title="Visa App", page_icon="🛂", layout="wide")
st.title("🛂 Visa App — Excel → analyse & export")
st.caption("Sélection automatique de l’onglet **Visa** (sinon **Clients**, sinon premier onglet disponible).")


# =============================
# Barre latérale
# =============================
with st.sidebar:
    st.header("Importer votre Excel")

    # Upload
    up = st.file_uploader(
        "Fichier .xlsx",
        type=["xlsx"],
        help="Choisissez le classeur contenant les onglets 'Visa' et/ou 'Clients'.",
    )

    # Chemin local optionnel
    data_path = st.text_input(
        "Ou saisissez un chemin local vers le .xlsx (optionnel)",
        value="",
        help="Exemple: C:/Users/charl/Desktop/visa_app/data.xlsx",
    )

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

    st.markdown("**Astuce** : ajoutez `?clear=1` à l’URL pour vider le cache au chargement.")


# =============================
# Sélection de la source (upload ou chemin)
# =============================
src = data_path if data_path.strip() else up
if not src:
    st.info("Chargez un fichier Excel (.xlsx) ou renseignez un chemin local dans la barre latérale pour commencer.")
    st.stop()

preferred = ("Visa", "Clients") if prefer_visa else ("Clients", "Visa")

# =============================
# Chargement robuste avec gestion d'erreurs
# =============================
try:
    df, used_sheet, sheet_names = load_data(src, preferred)
    all_sheets, _ = load_all_sheets(src)
except ValueError as e:
    # Message clair si la feuille demandée n'existe pas
    st.error(f"Erreur lors de la lecture du classeur : {e}")
    st.stop()
except Exception as e:
    st.exception(e)
    st.stop()

st.success(f"✅ Onglet utilisé : **{used_sheet}** · Onglets trouvés : {', '.join(sheet_names)}")


# =============================
# Recherche & Filtres
# =============================
st.subheader("Aperçu & filtres")

col_search, col_rows = st.columns([3, 1])
with col_search:
    q = st.text_input("Recherche (plein-texte sur toutes les colonnes)", placeholder="Tapez un mot-clé…")
with col_rows:
    max_rows = st.number_input("Lignes à afficher", min_value=5, max_value=5000, value=100, step=5)

filtered = df.copy()
if q:
    # Filtre plein-texte simple sur toutes colonnes
    mask = pd.Series(False, index=filtered.index)
    for c in filtered.columns:
        try:
            mask = mask | filtered[c].astype(str).str.contains(q, case=False, na=False)
        except Exception:
            pass
    filtered = filtered[mask]

with st.expander("Filtres par colonne (catégories)"):
    # Filtres multiselect pour colonnes texte/catégorie
    for col in filtered.select_dtypes(include=["object", "category"]).columns:
        unique_vals = sorted([v for v in filtered[col].dropna().unique() if str(v) != ""], key=lambda x: str(x).lower())
        if 1 < len(unique_vals) <= 1000:
            sel = st.multiselect(f"{col}", unique_vals, default=None)
            if sel:
                filtered = filtered[filtered[col].isin(sel)]

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
# Aperçu de l'autre onglet (si présent)
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
# Aide / Dépannage
# =============================
with st.expander("Aide / Dépannage"):
    st.markdown(
        """
        - Si vous voyiez **Worksheet not found**, assurez-vous que l’onglet existe ou laissez l’app choisir automatiquement l’onglet disponible.
        - Les noms de colonnes sont légèrement normalisés (suppression d'espaces).
        - Les filtres par colonnes s’appliquent aux colonnes de type texte/catégorie.
        - L’export Excel utilise **openpyxl** (`pip install openpyxl` si nécessaire).
        - Pour forcer un rafraîchissement complet: bouton **♻️** en sidebar ou ajoutez `?clear=1` à l’URL.
        """
    )


