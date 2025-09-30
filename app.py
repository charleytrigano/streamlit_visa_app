# app.py — Visa App avec navigation latérale (Visa / Clients) et CRUD Clients
import io
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


# =============================
# Clear cache via URL param ?clear=1 (API moderne)
# =============================
try:
    params = st.query_params  # MutableMapping
    clear_val = params.get("clear", "0")
    if isinstance(clear_val, list):
        clear_val = clear_val[0]
    if clear_val == "1":
        st.cache_data.clear()
        try:
            st.cache_resource.clear()
        except Exception:
            pass
        # Nettoie les query params et relance
        st.query_params.clear()
        st.rerun()
except Exception:
    pass


# =============================
# Helpers
# =============================

def _find_col(possible_names: List[str], columns: List[str]):
    """Retourne la 1re colonne correspondante (insensible aux accents/majuscules)."""
    import unicodedata

    def norm(s: str) -> str:
        s = str(s)
        s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
        return s.lower().strip()

    cols_norm = {norm(c): c for c in columns}
    for name in possible_names:
        key = norm(name)
        if key in cols_norm:
            return cols_norm[key]
    return None


def _as_bool_series(s: pd.Series) -> pd.Series:
    """Convertit des valeurs 'case à cocher' en booléen (gère 1/0, oui/non, x, ✓...)."""
    import numpy as np

    if s is None:
        return pd.Series([], dtype=bool)
    vals = s.astype(str).str.strip().str.lower()
    truthy = {"1", "true", "vrai", "yes", "oui", "y", "o", "x", "✓", "checked"}
    falsy = {"0", "false", "faux", "no", "non", "n", "", "none", "nan"}
    out = vals.apply(lambda v: True if v in truthy else (False if v in falsy else np.nan))
    return out.fillna(False)


def _to_numeric(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def _pick_sheet(sheet_names: List[str], preferred_order=("Visa", "Clients")) -> str:
    """Retourne la 1re feuille trouvée parmi preferred_order, sinon la 1re disponible."""
    for name in preferred_order:
        if name in sheet_names:
            return name
    return sheet_names[0]


@st.cache_data(show_spinner=False)
def load_all_sheets(xlsx_input) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    """Charge toutes les feuilles dans un dict {nom: DataFrame} avec colonnes normalisées."""
    xls = pd.ExcelFile(xlsx_input)
    out = {}
    for name in xls.sheet_names:
        _df = pd.read_excel(xls, sheet_name=name)
        _df.columns = [str(c).strip() for c in _df.columns]
        out[name] = _df
    return out, xls.sheet_names


@st.cache_data(show_spinner=False)
def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Feuille1") -> bytes:
    """Convertit un DataFrame en bytes Excel (XLSX)."""
    import openpyxl  # noqa: F401
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buffer.getvalue()


@st.cache_data(show_spinner=False)
def to_excel_bytes_multi(sheets: Dict[str, pd.DataFrame]) -> bytes:
    """Crée un classeur XLSX avec plusieurs onglets à partir d'un dict {nom: df}."""
    import openpyxl  # noqa: F401
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, _df in sheets.items():
            _df.to_excel(writer, index=False, sheet_name=name)
    return buffer.getvalue()


# =============================
# UI
# =============================
st.set_page_config(page_title="Visa App", page_icon="🛂", layout="wide")
st.title("🛂 Visa App — Excel → analyse & export")
st.caption("Navigation latérale : **Visa** et **Clients** (CRUD Clients inclus).")


# =============================
# Sidebar — source de données & navigation
# =============================
with st.sidebar:
    st.header("Importer votre Excel")
    up = st.file_uploader(
        "Fichier .xlsx",
        type=["xlsx"],
        help="Classeur contenant les onglets 'Visa' et/ou 'Clients'.",
    )
    data_path = st.text_input(
        "Ou saisissez un chemin local vers le .xlsx (optionnel)",
        value="",
        help="Exemple: C:/Users/charl/Desktop/visa_app/data.xlsx",
    )

    st.divider()
    page = st.radio("Sections", ["Visa", "Clients"], index=0)

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
# Sélection de la source (upload OU chemin)
# =============================
src = data_path if data_path.strip() else up
if not src:
    st.info("Chargez un fichier Excel (.xlsx) **ou** renseignez un chemin local dans la barre latérale pour commencer.")
    st.stop()

# =============================
# Chargement de toutes les feuilles
# =============================
try:
    all_sheets, sheet_names = load_all_sheets(src)
except ValueError as e:
    st.error(f"Erreur lors de la lecture du classeur : {e}")
    st.stop()

st.success(f"✅ Onglets trouvés : {', '.join(sheet_names)}")

visa_df = all_sheets.get("Visa")
clients_df_loaded = all_sheets.get("Clients")

# Met en mémoire de session une copie éditable des Clients (pour CRUD)
if "clients_df" not in st.session_state:
    st.session_state.clients_df = clients_df_loaded.copy() if clients_df_loaded is not None else pd.DataFrame()


# =============================
# PAGE: VISA
# =============================
if page == "Visa":
    st.subheader("🛂 Visa — tableau & filtres")

    if visa_df is None:
        st.warning("L’onglet **Visa** est introuvable dans le classeur.")
    else:
        df = visa_df.copy()
        col_search, col_rows = st.columns([3, 1])
        with col_search:
            q = st.text_input("Recherche (plein-texte)", placeholder="Tapez un mot-clé…")
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

        with st.expander("Filtres par colonne (catégories)"):
            for col in filtered.select_dtypes(include=["object", "category"]).columns:
                unique_vals = sorted([v for v in filtered[col].dropna().unique() if str(v) != ""], key=lambda x: str(x).lower())
                if 1 < len(unique_vals) <= 1000:
                    sel = st.multiselect(f"{col}", unique_vals, default=None)
                    if sel:
                        filtered = filtered[filtered[col].isin(sel)]

        st.markdown(f"**{len(filtered):,}** lignes affichées (sur **{len(df):,}**), **{len(df.columns)}** colonnes.")
        st.dataframe(filtered.head(int(max_rows)), use_container_width=True)

        st.subheader("Exports — Visa")
        c1, c2 = st.columns(2)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        with c1:
            csv_bytes = filtered.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Télécharger CSV — Visa",
                data=csv_bytes,
                file_name=f"Visa_{stamp}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with c2:
            xls_bytes = to_excel_bytes(filtered, sheet_name="Visa")
            st.download_button(
                "⬇️ Télécharger Excel — Visa",
                data=xls_bytes,
                file_name=f"Visa_{stamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


# =============================
# PAGE: CLIENTS (CRUD)
# =============================
if page == "Clients":
    st.subheader("👥 Clients — ajouter / modifier / supprimer")

    if st.session_state.clients_df is None or st.session_state.clients_df.empty:
        st.warning("L’onglet **Clients** est introuvable ou vide dans le classeur.")
        # Option pour créer un squelette vide
        if st.button("Créer l’onglet Clients vide"):
            st.session_state.clients_df = pd.DataFrame([
                {"Dossier": "", "Date": "", "Nom": "", "Type Visa": "", "Téléphone": "", "Email": "",
                 "Date facture": "", "Honoraires": "", "Date acompte 1": "", "Acompte 1": "",
                 "Date acompte 2": "", "Acompte 2": "", "Date acompte 3": "", "Acompte 3": "",
                 "Solde": "", "Date envoi": "", "Dossier envoyé": "", "Date retour": "",
                 "Dossier refusé": "", "Dossier approuvé": "", "RFE": ""}
