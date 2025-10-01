# app.py — version finale et complète
import json
from datetime import datetime, date
import pandas as pd
import streamlit as st
import numpy as np # Ajout pour le to_datetime de pandas/numpy dans la gestion des dates

# Importer les utilitaires depuis le nouveau fichier
from utils import (
    load_all_sheets,
    to_excel_bytes_multi,
    compute_finances,
    validate_rfe_row,
    harmonize_clients_df,
    _norm_cols 
)

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

# --- DÉBUT DE LA SECTION SIDEBAR (où data_path est défini) ---
# Sidebar / source / save options
with st.sidebar:
    st.header("Fichier source & sauvegarde")
    up = st.file_uploader("Fichier .xlsx", type=["xlsx"], help="Classeur contenant 'Visa' et 'Clients'.")
    data_path = st.text_input("Ou chemin local vers le .xlsx (optionnel)") # <-- DÉFINITION DE data_path
    st.markdown("---")
    st.subheader("Sauvegarde")
    save_mode = st.selectbox("Mode de sauvegarde", ["Download (toujours disponible)", "Save to local path (serveur/PC)", "Google Drive (secrets req.)", "OneDrive (secrets req.)"])
    save_path = st.text_input("Chemin local pour sauvegarde (si Save to local path)")
    st.markdown("---")
    st.info("Navigation : utilisez le menu en bas pour basculer entre Visa et Clients")

# --- FIN DE LA SECTION SIDEBAR ---

# Cette ligne peut maintenant s'exécuter car data_path et up sont définis
src = data_path if data_path.strip() else up
if not src:
    st.info("Chargez un fichier ou renseignez un chemin local pour commencer.")
    st.stop()

# load sheets
try:
    with st.spinner("Chargement et nettoyage des données..."):
        all_sheets, sheet_names = load_all_sheets(src)
except Exception as e:
    st.error(f"Erreur lecture fichier: {e}")
    st.stop()

st.success(f"Onglets trouvés: {', '.join(sheet_names)}")

visa_df = all_sheets.get("Visa")
clients_df_loaded = all_sheets.get("Clients")

# Normalize and ensure base columns
base_cols = [
    "DossierID", "DateCreation", "Nom", "TypeVisa", "Telephone", "Email",
    "DateFacture", "Honoraires", "Solde", "DateEnvoi", "Dossier envoyé",
    "DateRetour", "Dossier refusé", "Dossier approuvé", "RFE",
    "DateAnnulation", "DossierAnnule", "Notes", "Paiements" 
]

if clients_df_loaded is None:
    clients_df_loaded = pd.DataFrame(columns=base_cols)
else:
    # --- ÉTAPE CRUCIALE : HARMONISATION DES DONNÉES ---
    clients_df_loaded = harmonize_clients_df(clients_df_loaded) 
    
    # S'assurer que les colonnes de base existent après l'harmonisation
    for c in base_cols:
        if c not in clients_df_loaded.columns:
            # Initialisation avec pd.NaT pour les dates, 0.0 pour les honoraires, et vide pour les autres strings
            clients_df_loaded[c] = "" if c not in ["Honoraires", "DateCreation", "DateFacture", "DateEnvoi", "DateRetour", "DateAnnulation"] else (0.0 if c == "Honoraires" else pd.NaT)

# Initialisation de la session
if "clients_df" not in st.session_state:
    date_cols = ["DateCreation", "DateFacture", "DateEnvoi", "DateRetour", "DateAnnulation"]
    for col in date_cols:
        if col in clients_df_loaded.columns:
            # Conversion robuste des colonnes de dates
            clients_df_loaded[col] = pd.to_datetime(clients_df_loaded[col], errors='coerce')
            
    st.session_state.clients_df = clients_df_loaded.copy()

# ensure finances computed (this guarantees TotalAcomptes exists)
st.session_state.clients_df = compute_finances(st.session_state.clients_df)

# Navigation
page = st.selectbox("Page", ["Visa", "Clients"], index=0)

# Page Visa
if page == "Visa":
    st.header("🛂 Visa")
    if visa_df is None:
        st.warning("Onglet Visa introuvable")
    else:
        st.dataframe(visa_df.head(500), use_container_width=True)

# Page Clients
if page == "Clients":
    st.header("👥 Clients — gestion & suivi")

    df = st.session_state.clients_df

    # KPIs (use .get to be safe)
    total_dossiers = len(df) if df is not None else 0
    total_encaissé = float(df.get("TotalAcomptes", pd.Series([0])).sum()) if df is not None else 0.0
    total_honoraires = float(df.get("Honoraires", pd.Series([0])).sum()) if df is not None else 0.0
    total_solde = float(df.get("SoldeCalc", pd.Series([0])).sum()) if df is not None else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total dossiers", f"{total_dossiers:,}")
    c2.metric("Total encaissé", f"{total_encaissé:,.2f}")
    c3.metric("Total honoraires", f"{total_honoraires:,.2f}")
    c4.metric("Solde total", f"{total_solde:,.2f}")

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
            filtered = filtered[filtered.get("Dossier envoyé", False) == True]
        elif status_filter == "Approuvé":
            filtered = filtered[filtered.get("Dossier approuvé", False) == True]
        elif status_filter == "Refusé":
            filtered = filtered[filtered.get("Dossier refusé", False) == True]
        elif status_filter == "Annulé":
            filtered = filtered[filtered.get("DossierAnnule", False) == True]
        elif status_filter == "RFE":
            filtered = filtered[filtered.get("RFE", False) == True]

    st.dataframe(filtered.reset_index(drop=True), use_container_width=True)

    # select a client by index in filtered
    if len(filtered) == 0:
        st.info("Aucun dossier à afficher")
    else:
        max_idx = max(0, len(filtered)-1)
        sel_idx = st.number_input("Ouvrir dossier (index affiché)", min_value=0, max_value=max_idx, value=min(0, max_idx))
        sel_row = filtered.reset_index(drop=True).loc[int(sel_idx)] 
        
        st.subheader(f"Dossier: {sel_row.get('DossierID','(sans id)')} — {sel_row.get('Nom','')}")

        def get_date_for_input(col_name, row):
            dt = row.get(col_name)
            if pd.notna(dt) and isinstance(dt, (datetime, date, pd.Timestamp)):
                return dt
            if isinstance(dt, str):
                 try:
                     return pd.to_datetime(dt).date()
                 except:
                     pass
            return date.today()

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
                honoraires = st.number_input("Honoraires", value=float(sel_row.get("Honoraires", 0.0)), format="%.2f")
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
                date_envoi = st.date_input("DateEnvoi", value=get_date_for_input("DateEnvoi", sel_row))

            st.markdown("---")
            st.write("Paiements (Total encaissé: " + f"{sel_row.get('TotalAcomptes', 0.0):.2f}" + ")")
            
            payments_list = sel_row.get("Paiements", [])
            
            if isinstance(payments_list, str):
                try:
                    payments_list = json.loads(payments_list) if payments_list and pd.notna(payments_list) else []
                except Exception:
                    payments_list = []
            elif not isinstance(payments_list, list):
                 payments_list = []
            
            for i, p in enumerate(payments_list):
                p_date = p.get('date', 'N/A')
                p_amount = p.get('amount', 0)
                st.markdown(f"**{i+1}. {p_date}** — {p_amount:.2f}")

            st.markdown("---")
            st.write("Ajouter un nouveau paiement")
            col_pay1, col_pay2 = st.columns(2)
            with col_pay1:
                new_pay_date = st.date_input("Date du paiement", value=date.today())
            with col_pay2:
                new_pay_amount = st.number_input("Montant", value=0.0, min_value=0.0, format="%.2f")

            submitted = st.form_submit_button("Enregistrer les modifications")
            if submitted:
                # Récupérer l'index original dans le DataFrame de session
                original_index = filtered.index[int(sel_idx)]

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
                updated["DateEnvoi"] = date_envoi
                
                current_payments_list = payments_list.copy()
                if new_pay_amount and float(new_pay_amount) > 0:
                    current_payments_list.append({"date": str(new_pay_date), "amount": float(new_pay_amount)})
                
                updated["Paiements"] = current_payments_list

                ok, msg = validate_rfe_row(updated)
                if not ok:
                    st.error(msg)
                else:
                    dossier_id_original = sel_row.get("DossierID")
                    
                    # Logique de mise à jour (recherche par ID ou par index)
                    if not dossier_id_original:
                        update_row_idx = original_index
                    else:
                         idxs = st.session_state.clients_df.index[st.session_state.clients_df.get("DossierID") == dossier_id_original].tolist()
                         update_row_idx = idxs[0] if idxs else None

                    if update_row_idx is not None:
                         st.session_state.clients_df.loc[update_row_idx] = updated
                         st.success("Modifications sauvegardées en session.")
                    else:
                        st.session_state.clients_df = pd.concat([st.session_state.clients_df, pd.DataFrame([updated])], ignore_index=True)
                        st.success("Nouveau dossier ajouté en session.")

                    st.session_state.clients_df = compute_finances(st.session_state.clients_df)
                    st.rerun()

    # quick export / save actions
    st.markdown("---")
    exp_col1, exp_col2, exp_col3 = st.columns(3)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    with exp_col1:
        # Assurer que 'Paiements' est sauvé en JSON string pour l'export CSV
        csv_df = st.session_state.clients_df.copy()
        csv_df["Paiements"] = csv_df["Paiements"].apply(json.dumps)
        csv_bytes = csv_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Télécharger CSV — Clients", data=csv_bytes, file_name=f"Clients_{stamp}.csv", mime="text/csv")
    with exp_col2:
        # Assurer que 'Paiements' est sauvé en JSON string pour l'export XLSX
        xls_df = st.session_state.clients_df.copy()
        xls_df["Paiements"] = xls_df["Paiements"].apply(json.dumps)
        xls_bytes = to_excel_bytes_multi({"Clients": xls_df, **({"Visa": visa_df} if visa_df is not None else {})})
        st.download_button("⬇️ Télécharger XLSX — Classeur (Visa+Clients)", data=xls_bytes, file_name=f"Visa_Clients_{stamp}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with exp_col3:
        if save_mode == "Save to local path (serveur/PC)":
            if save_path:
                try:
                    save_df = st.session_state.clients_df.copy()
                    save_df["Paiements"] = save_df["Paiements"].apply(json.dumps)
                    xls_bytes = to_excel_bytes_multi({"Clients": save_df, **({"Visa": visa_df} if visa_df is not None else {})})
                    with open(save_path, "wb") as f:
                        f.write(xls_bytes)
                    st.success(f"Fichier écrit: {save_path}")
                except Exception as e:
                    st.error(f"Erreur écriture locale: {e}")
            else:
                st.warning("Renseignez un chemin local dans la sidebar.")
        elif save_mode == "Google Drive (secrets req.)":
            creds = st.secrets.get("gdrive")
            if not creds:
                st.error("Aucun secret gdrive trouvé. Ajoutez vos identifiants dans st.secrets['gdrive']")
            else:
                st.info("Upload Google Drive non-implémenté automatiquement. Voir README pour config.")
        elif save_mode == "OneDrive (secrets req.)":
            st.info("OneDrive upload non-implémenté automatiquement. Voir README pour config OAuth.")



