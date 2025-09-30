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
                unique_vals = sorted(
                    [v for v in filtered[col].dropna().unique() if str(v) != ""],
                    key=lambda x: str(x).lower(),
                )
                if 1 < len(unique_vals) <= 1000:
                    sel = st.multiselect(f"{col}", unique_vals, default=None)
                    if sel:
                        filtered = filtered[filtered[col].isin(sel)]

        st.markdown(f"**{len(filtered):,}** lignes affichées (sur **{len[df]:,}**), **{len(df.columns)}** colonnes.")
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
                {
                    "Dossier": "",
                    "Date": "",
                    "Nom": "",
                    "Type Visa": "",
                    "Téléphone": "",
                    "Email": "",
                    "Date facture": "",
                    "Honoraires": "",
                    "Date acompte 1": "",
                    "Acompte 1": "",
                    "Date acompte 2": "",
                    "Acompte 2": "",
                    "Date acompte 3": "",
                    "Acompte 3": "",
                    "Solde": "",
                    "Date envoi": "",
                    "Dossier envoyé": "",
                    "Date retour": "",
                    "Dossier refusé": "",
                    "Dossier approuvé": "",
                    "RFE": "",
                }
            ])
            st.rerun()
    else:
        clients_df = st.session_state.clients_df

        tabs = st.tabs(["Ajouter", "Modifier / Supprimer", "Tableau & exports"])

        # --- Ajouter ---
        with tabs[0]:
            st.caption("Ajouter un nouveau client (les champs sont libres — adaptez à vos colonnes)")
            cols = list(clients_df.columns)
            # champs principaux suggérés
            d1, d2, d3 = st.columns(3)
            with d1:
                v_dossier = st.text_input("Dossier", value="")
                v_nom = st.text_input("Nom", value="")
                v_type = st.text_input("Type Visa", value="")
            with d2:
                v_tel = st.text_input("Téléphone", value="")
                v_email = st.text_input("Email", value="")
                v_hon = st.text_input("Honoraires", value="")
            with d3:
                v_envoye = st.checkbox("Dossier envoyé")
                v_refuse = st.checkbox("Dossier refusé")
                v_approuve = st.checkbox("Dossier approuvé")
                v_rfe = st.checkbox("RFE (doit être combiné avec un des 3 statuts)")

            if st.button("➕ Ajouter ce client", type="primary"):
                new_row = {c: "" for c in cols}
                # injecte les valeurs communes si elles existent dans les colonnes
                for k, val in {
                    "Dossier": v_dossier,
                    "Nom": v_nom,
                    "Type Visa": v_type,
                    "Téléphone": v_tel,
                    "Email": v_email,
                    "Honoraires": v_hon,
                    "Dossier envoyé": v_envoye,
                    "Dossier refusé": v_refuse,
                    "Dossier approuvé": v_approuve,
                    "RFE": v_rfe,
                }.items():
                    if k in new_row:
                        new_row[k] = val
                st.session_state.clients_df = pd.concat([clients_df, pd.DataFrame([new_row])], ignore_index=True)
                st.success("Client ajouté.")
                st.rerun()

        # --- Modifier / Supprimer ---
        with tabs[1]:
            st.caption("Modifiez directement dans le tableau. Cochez des lignes à supprimer puis cliquez sur Supprimer.")
            editable = st.data_editor(
                clients_df,
                use_container_width=True,
                num_rows="dynamic",
                key="clients_editor",
            )
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("💾 Enregistrer les modifications"):
                    st.session_state.clients_df = editable
                    st.success("Modifications enregistrées en mémoire.")
            with c2:
                to_delete = st.multiselect("Sélectionner les index à supprimer", options=list(editable.index))
                if st.button("🗑️ Supprimer les lignes sélectionnées") and to_delete:
                    st.session_state.clients_df = editable.drop(index=to_delete).reset_index(drop=True)
                    st.success(f"Supprimé : {len(to_delete)} ligne(s).")
                    st.rerun()
            with c3:
                if st.button("↩️ Réinitialiser depuis le fichier chargé"):
                    st.session_state.clients_df = clients_df_loaded.copy() if clients_df_loaded is not None else pd.DataFrame()
                    st.success("Réinitialisé.")
                    st.rerun()

        # --- Tableau & exports ---
        with tabs[2]:
            st.dataframe(st.session_state.clients_df, use_container_width=True)
            e1, e2 = st.columns(2)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            with e1:
                clients_csv = st.session_state.clients_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    "⬇️ Télécharger CSV — Clients (modifié)",
                    data=clients_csv,
                    file_name=f"Clients_mod_{stamp}.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            with e2:
                clients_xlsx = to_excel_bytes(st.session_state.clients_df, sheet_name="Clients")
                st.download_button(
                    "⬇️ Télécharger Excel — Clients (modifié)",
                    data=clients_xlsx,
                    file_name=f"Clients_mod_{stamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            # Export du classeur complet (Visa + Clients)
            st.markdown("**Exporter le classeur complet (Visa + Clients)**")
            sheets_out = {}
            if visa_df is not None:
                sheets_out["Visa"] = visa_df
            if not st.session_state.clients_df.empty:
                sheets_out["Clients"] = st.session_state.clients_df
            if sheets_out:
                full_xlsx = to_excel_bytes_multi(sheets_out)
                st.download_button(
                    "⬇️ Télécharger Excel — Classeur complet",
                    data=full_xlsx,
                    file_name=f"Visa_Clients_{stamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


# =============================
# Aide / Dépannage
# =============================
with st.expander("Aide / Dépannage"):
    st.markdown(
        """
        - Le disque de Streamlit Cloud est éphémère : les ajouts/modifs Clients sont conservés en **mémoire de session**
          et disponibles au téléchargement (CSV/XLSX). Pour persister côté serveur, stockez dans un bucket (S3/GCS)
          ou téléchargez le classeur complet puis remplacez votre fichier source.
        - `RFE` est détecté comme colonne et peut être cochée en combinaison avec *Envoyé/Refusé/Approuvé*.
        - Pour forcer un rafraîchissement : bouton **♻️** en sidebar ou ajoutez `?clear=1` à l’URL.
        """
    )
