# app.py — Visa App (Clients = source de vérité)
# Features added: DateAnnulation, DossierAnnule, dynamic payments, RFE validation,
# save options: Download XLSX, save to local path (when running locally), st.secrets-driven Google Drive/OneDrive upload stubs.

import io
import json
from datetime import datetime, date
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st


# =============================
# Config & helpers
# =============================
st.set_page_config(page_title="Visa App", page_icon="🛂", layout="wide")

# Clear cache via URL param ?clear=1
try:
    params = st.query_params
    clear_val = params.get("clear", "0")
    if isinstance(clear_val, list):
        clear_val = clear_val[0]
    if clear_val == "1":
        st.cache_data.clear()
        try:
            st.cache_resource.clear()
        except Exception:
            pass
        st.query_params.clear()
        st.rerun()
except Exception:
    pass


def _norm_cols(cols: List[str]) -> List[str]:
    return [str(c).strip() for c in cols]


def _find_col(possible_names: List[str], columns: List[str]):
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
    import numpy as np
    if s is None:
        return pd.Series([], dtype=bool)
    vals = s.astype(str).str.strip().str.lower()
    truthy = {"1", "true", "vrai", "yes", "oui", "y", "o", "x", "✓", "checked"}
    falsy = {"0", "false", "faux", "no", "non", "n", "", "none", "nan"}
    out = vals.apply(lambda v: True if v in truthy else (False if v in falsy else pd.NA))
    return out.fillna(False)


@st.cache_data(show_spinner=False)
def load_all_sheets(xlsx_input) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    xls = pd.ExcelFile(xlsx_input)
    out = {}
    for name in xls.sheet_names:
        _df = pd.read_excel(xls, sheet_name=name)
        _df.columns = _norm_cols(_df.columns)
        out[name] = _df
    return out, xls.sheet_names


@st.cache_data(show_spinner=False)
def to_excel_bytes_multi(sheets: Dict[str, pd.DataFrame]) -> bytes:
    import openpyxl  # noqa: F401
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, _df in sheets.items():
            _df.to_excel(writer, index=False, sheet_name=name[:31])
    return buffer.getvalue()


# Finance helpers: dynamic payments stored as JSON in column 'Paiements' (list of {date,amount})
def compute_finances(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # ensure numeric honoraires
    if "Honoraires" not in df.columns:
        df["Honoraires"] = 0
    df["Honoraires"] = pd.to_numeric(df["Honoraires"], errors="coerce").fillna(0)

    # payments column
    if "Paiements" not in df.columns:
        df["Paiements"] = "[]"

    def sum_payments(cell):
        try:
            lst = cell if isinstance(cell, list) else json.loads(cell) if cell and pd.notna(cell) else []
        except Exception:
            lst = []
        s = 0
        for p in lst:
            try:
                s += float(p.get("amount", 0))
            except Exception:
                pass
        return s

    df["TotalAcomptes"] = df["Paiements"].apply(sum_payments)
    df["SoldeCalc"] = (df["Honoraires"] - df["TotalAcomptes"]).round(2)
    return df


def validate_rfe_row(row: pd.Series) -> Tuple[bool, str]:
    rfe = bool(row.get("RFE", False))
    sent = bool(row.get("Dossier envoyé", False) or row.get("Dossier envoye", False))
    refused = bool(row.get("Dossier refusé", False) or row.get("Dossier refuse", False))
    approved = bool(row.get("Dossier approuvé", False) or row.get("Dossier approuve", False))
    canceled = bool(row.get("DossierAnnule", False) or row.get("Dossier Annule", False) or row.get("Dossier annulé", False))
    # RFE cannot be true if none of sent/refused/approved and not canceled
    if rfe and not (sent or refused or approved):
        return False, "RFE doit être combinée avec Envoyé / Refusé / Approuvé"
    # Can't be both approved and refused
    if approved and refused:
        return False, "Un dossier ne peut pas être à la fois Approuvé et Refusé"
    # If canceled, clear other statuses
    if canceled and (sent or refused or approved):
        return False, "Un dossier annulé ne peut pas être marqué Envoyé/Refusé/Approuvé"
    return True, ""


# =============================
# Sidebar: source, save options
# =============================
with st.sidebar:
    st.header("Fichier source & sauvegarde")
    up = st.file_uploader("Fichier .xlsx", type=["xlsx"], help="Classeur contenant 'Visa' et 'Clients'.")
    data_path = st.text_input("Ou chemin local vers le .xlsx (optionnel)")

    st.markdown("---")
    st.subheader("Sauvegarde")
    save_mode = st.selectbox("Mode de sauvegarde", ["Download (toujours disponible)", "Save to local path (serveur/PC)", "Google Drive (secrets req.)", "OneDrive (secrets req.)"]) 
    save_path = st.text_input("Chemin local pour sauvegarde (si Save to local path)")
    st.markdown("Les sauvegardes vers Google Drive/OneDrive requièrent des identifiants dans `st.secrets`.")

    st.markdown("---")
    st.info("Navigation : utilisez le menu en bas pour basculer entre Visa et Clients")

src = data_path if data_path.strip() else up
if not src:
    st.info("Chargez un fichier ou renseignez un chemin local pour commencer.")
    st.stop()

# load
try:
    all_sheets, sheet_names = load_all_sheets(src)
except Exception as e:
    st.error(f"Erreur lecture fichier: {e}")
    st.stop()

st.success(f"Onglets trouvés: {', '.join(sheet_names)}")

visa_df = all_sheets.get("Visa")
clients_df_loaded = all_sheets.get("Clients")

# ensure clients has canonical columns including new ones
base_cols = [
    "DossierID", "DateCreation", "Nom", "TypeVisa", "Telephone", "Email",
    "DateFacture", "Honoraires", "Solde", "DateEnvoi", "Dossier envoyé",
    "DateRetour", "Dossier refusé", "Dossier approuvé", "RFE",
    "DateAnnulation", "DossierAnnule", "Notes", "Paiements"
]

if clients_df_loaded is None:
    clients_df_loaded = pd.DataFrame(columns=base_cols)
else:
    # normalize column names that users may have
    clients_df_loaded.columns = _norm_cols(clients_df_loaded.columns)
    for c in base_cols:
        if c not in clients_df_loaded.columns:
            clients_df_loaded[c] = "" if c != "Honoraires" else 0

# session copy for edits
if "clients_df" not in st.session_state:
    st.session_state.clients_df = clients_df_loaded.copy()

def compute_finances(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcule TotalAcomptes et SoldeCalc à partir de la colonne 'Paiements' (JSON/list)
    et de 'Honoraires'. Retourne une copie du dataframe avec colonnes ajoutées/normalisées.
    """
    df = df.copy()

    # Ensure Honoraires exists and is numeric
    if "Honoraires" not in df.columns:
        df["Honoraires"] = 0
    df["Honoraires"] = pd.to_numeric(df["Honoraires"], errors="coerce").fillna(0.0)

    # Ensure Paiements exists (store as JSON-string or list)
    if "Paiements" not in df.columns:
        df["Paiements"] = "[]"

    def sum_payments(cell):
        # parse various possible formats safely -> return float sum
        try:
            if isinstance(cell, list):
                lst = cell
            elif isinstance(cell, str):
                # empty or JSON list
                cell_strip = cell.strip()
                if cell_strip == "":
                    lst = []
                else:
                    try:
                        lst = json.loads(cell_strip)
                    except Exception:
                        # maybe it is a representation like "[{'date':...}]" — fallback: empty
                        lst = []
            elif pd.isna(cell):
                lst = []
            else:
                lst = []
        except Exception:
            lst = []

        total = 0.0
        for p in lst:
            try:
                if isinstance(p, dict):
                    amt = float(p.get("amount", 0) or 0)
                elif isinstance(p, (int, float)):
                    amt = float(p)
                else:
                    amt = float(str(p))
            except Exception:
                amt = 0.0
            total += amt
        return total

    # Compute TotalAcomptes (as numeric)
    df["TotalAcomptes"] = df["Paiements"].apply(sum_payments)
    df["TotalAcomptes"] = pd.to_numeric(df["TotalAcomptes"], errors="coerce").fillna(0.0)

    # Compute SoldeCalc as numeric and round
    df["SoldeCalc"] = (df["Honoraires"].astype(float) - df["TotalAcomptes"].astype(float)).round(2)

    return df

# Navigation
page = st.selectbox("Page", ["Visa", "Clients"], index=0)

# =============================
# Page Visa
# =============================
if page == "Visa":
    st.header("🛂 Visa")
    if visa_df is None:
        st.warning("Onglet Visa introuvable")
    else:
        st.dataframe(visa_df.head(500), use_container_width=True)

# =============================
# Page Clients
# =============================
if page == "Clients":
    st.header("👥 Clients — gestion & suivi")

    df = st.session_state.clients_df

    # Top KPIs
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total dossiers", f"{len(df):,}")
    c2.metric("Total encaissé", f"{df['TotalAcomptes'].sum():,.2f}")
    c3.metric("Total honoraires", f"{df['Honoraires'].sum():,.2f}")
    c4.metric("Solde total", f"{df['SoldeCalc'].sum():,.2f}")

    # filter / select
    with st.expander("Filtrer / Rechercher"):
        q = st.text_input("Recherche (nom / dossier / email)")
        status_filter = st.selectbox("Filtrer par statut", ["Tous", "Envoyé", "Approuvé", "Refusé", "Annulé", "RFE"])

    filtered = df.copy()
    if q:
        mask = pd.Series(False, index=filtered.index)
        for c in ["DossierID", "Nom", "Email", "TypeVisa"]:
            if c in filtered.columns:
                mask = mask | filtered[c].astype(str).str.contains(q, case=False, na=False)
        filtered = filtered[mask]
    if status_filter != "Tous":
        if status_filter == "Envoyé":
            filtered = filtered[filtered["Dossier envoyé"] == True]
        elif status_filter == "Approuvé":
            filtered = filtered[filtered["Dossier approuvé"] == True]
        elif status_filter == "Refusé":
            filtered = filtered[filtered["Dossier refusé"] == True]
        elif status_filter == "Annulé":
            filtered = filtered[filtered["DossierAnnule"] == True]
        elif status_filter == "RFE":
            filtered = filtered[filtered["RFE"] == True]

    st.dataframe(filtered.reset_index(drop=True), use_container_width=True)

    # Select a client to open detail / edit
    sel_idx = st.number_input("Ouvrir dossier (index affiché)", min_value=0, max_value=max(0, len(filtered)-1), value=0)
    if len(filtered) == 0:
        st.info("Aucun dossier à afficher")
    else:
        sel_row = filtered.reset_index(drop=True).loc[int(sel_idx)]
        st.subheader(f"Dossier: {sel_row.get('DossierID','(sans id)')} — {sel_row.get('Nom','')}")

        # detail form
        with st.form("client_form"):
            cols1, cols2 = st.columns(2)
            with cols1:
                dossier_id = st.text_input("DossierID", value=sel_row.get("DossierID", ""))
                nom = st.text_input("Nom", value=sel_row.get("Nom", ""))
                typevisa = st.text_input("TypeVisa", value=sel_row.get("TypeVisa", ""))
                email = st.text_input("Email", value=sel_row.get("Email", ""))
            with cols2:
                telephone = st.text_input("Telephone", value=sel_row.get("Telephone", ""))
                honoraires = st.number_input("Honoraires", value=float(sel_row.get("Honoraires", 0)), format="%.2f")
                notes = st.text_area("Notes", value=sel_row.get("Notes", ""))

            st.markdown("---")
            st.write("Statuts / dates")
            st_col1, st_col2, st_col3 = st.columns(3)
            with st_col1:
                dossier_envoye = st.checkbox("Dossier envoyé", value=bool(sel_row.get("Dossier envoyé", False)))
                dossier_refuse = st.checkbox("Dossier refusé", value=bool(sel_row.get("Dossier refusé", False)))
            with st_col2:
                dossier_approuve = st.checkbox("Dossier approuvé", value=bool(sel_row.get("Dossier approuvé", False)))
                dossier_annule = st.checkbox("DossierAnnule (annulé)", value=bool(sel_row.get("DossierAnnule", False)))
            with st_col3:
                rfe = st.checkbox("RFE (doit être combiné)", value=bool(sel_row.get("RFE", False)))
                date_envoi = st.date_input("DateEnvoi", value=sel_row.get("DateEnvoi") if pd.notna(sel_row.get("DateEnvoi","")) and sel_row.get("DateEnvoi")!="" else date.today())

            st.markdown("---")
            st.write("Paiements")
            # show existing payments
            payments = sel_row.get("Paiements", "[]")
            try:
                payments_list = payments if isinstance(payments, list) else json.loads(payments) if payments and pd.notna(payments) else []
            except Exception:
                payments_list = []
            for i, p in enumerate(payments_list):
                st.write(f"{i+1}. {p.get('date','')} — {p.get('amount','')}")

            # add payment
            new_pay_date = st.date_input("Date paiement (nouveau)", value=date.today())
            new_pay_amount = st.number_input("Montant (nouveau)", value=0.0, format="%.2f")

            submitted = st.form_submit_button("Enregistrer les modifications")
            if submitted:
                # build updated row
                updated = sel_row.copy()
                updated["DossierID"] = dossier_id
                updated["Nom"] = nom
                updated["TypeVisa"] = typevisa
                updated["Email"] = email
                updated["Telephone"] = telephone
                updated["Honoraires"] = float(honoraires)
                updated["Notes"] = notes
                updated["Dossier envoyé"] = bool(dossier_envoye)
                updated["Dossier refusé"] = bool(dossier_refuse)
                updated["Dossier approuvé"] = bool(dossier_approuve)
                updated["DossierAnnule"] = bool(dossier_annule)
                updated["RFE"] = bool(rfe)
                updated["DateEnvoi"] = str(date_envoi)
                # append new payment if >0
                if new_pay_amount and float(new_pay_amount) > 0:
                    payments_list.append({"date": str(new_pay_date), "amount": float(new_pay_amount)})
                updated["Paiements"] = json.dumps(payments_list)

                # validate
                ok, msg = validate_rfe_row(updated)
                if not ok:
                    st.error(msg)
                else:
                    # locate index in session_state df and update
                    # find by DossierID first, otherwise by index matching
                    idxs = st.session_state.clients_df.index[st.session_state.clients_df.get("DossierID") == sel_row.get("DossierID")].tolist()
                    if not idxs:
                        # fallback: match by exact row equality (not ideal) — append as new
                        st.session_state.clients_df = pd.concat([st.session_state.clients_df, pd.DataFrame([updated])], ignore_index=True)
                    else:
                        st.session_state.clients_df.loc[idxs[0], :] = pd.Series(updated)
                    # recompute finances
                    st.session_state.clients_df = compute_finances(st.session_state.clients_df)
                    st.success("Modifications sauvegardées en session.")

    # quick export / save actions
    st.markdown("---")
    exp_col1, exp_col2, exp_col3 = st.columns(3)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    with exp_col1:
        csv_bytes = st.session_state.clients_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Télécharger CSV — Clients", data=csv_bytes, file_name=f"Clients_{stamp}.csv", mime="text/csv")
    with exp_col2:
        xls_bytes = to_excel_bytes_multi({"Clients": st.session_state.clients_df, **({"Visa": visa_df} if visa_df is not None else {})})
        st.download_button("⬇️ Télécharger XLSX — Classeur (Visa+Clients)", data=xls_bytes, file_name=f"Visa_Clients_{stamp}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with exp_col3:
        if save_mode == "Save to local path (serveur/PC)":
            if save_path:
                try:
                    with open(save_path, "wb") as f:
                        f.write(to_excel_bytes_multi({"Clients": st.session_state.clients_df, **({"Visa": visa_df} if visa_df is not None else {})}))
                    st.success(f"Fichier écrit: {save_path}")
                except Exception as e:
                    st.error(f"Erreur écriture locale: {e}")
            else:
                st.warning("Renseignez un chemin local dans la sidebar.")
        elif save_mode == "Google Drive (secrets req.)":
            # minimal attempt: expects st.secrets['gdrive'] with credentials JSON or token info
            try:
                creds = st.secrets.get("gdrive")
                if not creds:
                    st.error("Aucun secret gdrive trouvé. Ajoutez vos identifiants dans st.secrets['gdrive']")
                else:
                    st.info("Upload Google Drive non-implémenté automatiquement. Voir README pour config.")
            except Exception as e:
                st.error(f"Google Drive error: {e}")
        elif save_mode == "OneDrive (secrets req.)":
            st.info("OneDrive upload non-implémenté automatiquement. Voir README pour config OAuth.")

# End of app


