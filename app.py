# app.py — Version finale avec hyper-unicité des clés de widget (Corrigé 22)
import json
from datetime import datetime, date
import pandas as pd
import streamlit as st
import numpy as np 

# Importer les utilitaires (assurez-vous que utils.py est à jour)
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

# --- 1. FONCTIONS DE GESTION DES DONNÉES (INITIALISATION ET UTILITAIRES) ---

def initialize_session_state(all_sheets):
    """Charge les DataFrames initiaux et les stocke dans st.session_state."""
    
    # ------------------- DONNÉES CLIENTS -------------------
    clients_df_loaded = all_sheets.get("Clients")
    
    if clients_df_loaded is None:
        # Recherche robuste pour "Clients"
        clients_key = next((k for k in all_sheets.keys() if "client" in str(k).lower()), None)
        if clients_key:
            clients_df_loaded = all_sheets.get(clients_key)
            if clients_df_loaded is not None:
                st.info(f"Onglet 'Clients' non trouvé. Utilisation de l'onglet '{clients_key}'.")

    # Colonnes de base pour les clients
    base_cols_clients = [
        "DossierID", "DateCreation", "Nom", "TypeVisa", "Telephone", "Email",
        "DateFacture", "Honoraires", "Solde", "DateEnvoi", "Dossier envoyé",
        "DateRetour", "Dossier refusé", "Dossier approuvé", "RFE",
        "DateAnnulation", "DossierAnnule", "Notes", "Paiements" 
    ]
    
    if clients_df_loaded is None or clients_df_loaded.empty:
        if clients_df_loaded is None:
            st.warning("Onglet Clients introuvable ou illisible. Création d'un DataFrame Clients vide.")
        clients_df_loaded = pd.DataFrame(columns=base_cols_clients)
    else:
        clients_df_loaded = harmonize_clients_df(clients_df_loaded) 
        for c in base_cols_clients:
            if c not in clients_df_loaded.columns:
                 # Initialisation par défaut basée sur le type attendu
                default_val = pd.NaT if c.startswith("Date") else (0.0 if c == "Honoraires" else ([] if c == "Paiements" else ""))
                clients_df_loaded[c] = default_val

    # Finalisation des types de colonnes (Dates)
    date_cols = [c for c in base_cols_clients if c.startswith("Date")]
    for col in date_cols:
        clients_df_loaded[col] = pd.to_datetime(clients_df_loaded.get(col), errors='coerce')
            
    st.session_state.clients_df = compute_finances(clients_df_loaded.copy())
    st.session_state.clients_df = st.session_state.clients_df.astype({"Paiements": object})

    # ------------------- DONNÉES VISA -------------------
    visa_df_loaded = all_sheets.get("Visa")

    base_cols_visa = ["Categories", "Visa", "Definition"]

    if visa_df_loaded is None or visa_df_loaded.empty:
        if visa_df_loaded is None:
            st.warning("Onglet Visa introuvable ou illisible. Création d'un DataFrame Visa vide.")
        st.session_state.visa_df = pd.DataFrame(columns=base_cols_visa)
    else:
        # Nettoyage de colonnes pour Visa (norm_cols est déjà dans load_all_sheets)
        visa_df_loaded = visa_df_loaded.rename(columns={c:c for c in visa_df_loaded.columns if c in base_cols_visa})
        # S'assurer que les colonnes existent
        for col in base_cols_visa:
            if col not in visa_df_loaded.columns:
                visa_df_loaded[col] = ""
        st.session_state.visa_df = visa_df_loaded[base_cols_visa].copy()

    # Initialisation de l'index pour le formulaire (utiliser un index propre)
    if not st.session_state.visa_df.index.name:
        st.session_state.visa_df = st.session_state.visa_df.reset_index(drop=True)


def get_date_for_input(col_name, row):
    """Fonction utilitaire pour formatter les dates pour les date_input de Streamlit."""
    dt = row.get(col_name)
    if pd.notna(dt) and isinstance(dt, (datetime, date, pd.Timestamp)):
        return dt.date()
    return date.today()

# --- 2. LOGIQUE PRINCIPALE DE L'APPLICATION ---

# Sidebar / source / save options
with st.sidebar:
    st.header("Fichier source & sauvegarde")
    up = st.file_uploader("Fichier .xlsx", type=["xlsx"], help="Classeur contenant 'Visa' et 'Clients'.")
    data_path = st.text_input("Ou chemin local vers le .xlsx (optionnel)") 
    st.markdown("---")
    st.subheader("Sauvegarde")
    save_mode = st.selectbox("Mode de sauvegarde", ["Download (toujours disponible)", "Save to local path (serveur/PC)", "Google Drive (secrets req.)", "OneDrive (secrets req.)"])
    save_path = st.text_input("Chemin local pour sauvegarde (si Save to local path)", value="data_sauvegardee.xlsx")
    st.markdown("---")
    st.info("Navigation : utilisez le menu en bas pour basculer entre Clients et Visa")

src = data_path if data_path.strip() else up
if not src:
    st.info("Chargez un fichier ou renseignez un chemin local pour commencer.")
    st.stop()

# load sheets
if "clients_df" not in st.session_state:
    try:
        with st.spinner("Chargement et nettoyage des données..."):
            all_sheets, _ = load_all_sheets(src)
            initialize_session_state(all_sheets)
    except Exception as e:
        st.error(f"Erreur lecture fichier: {e}")
        st.stop()
    st.rerun()

# Navigation principale
page = st.selectbox("Page", ["Clients", "Visa"], index=0)

# --- 3. RENDU DES PAGES ---

if page == "Clients":
    
    st.header("👥 Clients — gestion & suivi")
    df = st.session_state.clients_df
    
    # Navigation CRUD dans la page Clients
    crud_mode = st.radio("Action Clients", ["Lister/Modifier/Supprimer", "Ajouter un nouveau dossier"], index=0, horizontal=True)

    if crud_mode == "Ajouter un nouveau dossier":
        st.subheader("Ajouter un nouveau dossier client")
        
        # Créer une ligne vide pour l'ajout
        empty_row = pd.Series("", index=df.columns)
        empty_row["Paiements"] = [] 
        
        # Ligne 162 (ajout)
        render_client_form(df, empty_row, action="add")

    elif crud_mode == "Lister/Modifier/Supprimer":
        
        # KPIs (omitted for brevity)
        total_dossiers = len(df)
        total_encaissé = df["TotalAcomptes"].sum()
        total_honoraires = df["Honoraires"].sum()
        total_solde = df["SoldeCalc"].sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total dossiers", f"{total_dossiers:,}")
        c2.metric("Total encaissé", f"{total_encaissé:,.2f} €")
        c3.metric("Total honoraires", f"{total_honoraires:,.2f} €")
        c4.metric("Solde total", f"{total_solde:,.2f} €")

        # Filtrage
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
            col_map = {"Envoyé": "Dossier envoyé", "Approuvé": "Dossier approuvé", "Refusé": "Dossier refusé", "Annulé": "DossierAnnule", "RFE": "RFE"}
            col_name = col_map.get(status_filter)
            if col_name:
                 filtered = filtered[filtered.get(col_name, False) == True]


        st.dataframe(filtered.reset_index(drop=True), use_container_width=True)

        # Sélection et modification
        if len(filtered) > 0: 
            
            # --- ZONE CRITIQUE DE SÉLECTION D'INDEX STABILISÉE ---
            
            max_idx = len(filtered) - 1
            
            # 1. INITIALISATION ET CONTRÔLE D'INDEX CRITIQUE
            if 'client_sel_idx' not in st.session_state:
                st.session_state.client_sel_idx = 0
            
            current_index = st.session_state.client_sel_idx
            
            # Fix index if out of bounds (protection primaire)
            if current_index > max_idx or current_index < 0:
                st.session_state.client_sel_idx = min(max_idx, max(0, current_index)) 
                st.warning("Index Client réinitialisé (valeur hors limites).")
                st.rerun() 
            
            final_safe_index = st.session_state.client_sel_idx

            # 2. L'utilisateur choisit l'index affiché (Clé Statique)
            sel_idx_float = st.number_input(
                "Ouvrir dossier (index affiché)", 
                min_value=0, 
                max_value=max_idx, 
                value=final_safe_index, 
                key="client_idx_input_static" 
            )
            
            sel_idx = int(sel_idx_float) 
            
            # 3. Mettre à jour la session state
            if sel_idx != final_safe_index:
                st.session_state.client_sel_idx = sel_idx
                st.rerun() 
            
            # --- DÉFENSE ULTIME CONTRE IndexError (try/except) ---
            try:
                # Accès aux données garanti
                sel_row_filtered = filtered.iloc[st.session_state.client_sel_idx] 
                original_session_index = sel_row_filtered.name 

                st.subheader(f"Modifier Dossier: {sel_row_filtered.get('DossierID','(sans id)')} — {sel_row_filtered.get('Nom','')}")
                
                # Ligne 246 (mise à jour)
                render_client_form(df, sel_row_filtered, action="update", original_index=original_session_index)

            except IndexError as e:
                # Si l'index est désynchronisé (après une suppression rapide), on réinitialise et on relance
                st.session_state.client_sel_idx = 0
                st.error("Erreur d'index détectée après modification. Redémarrage automatique.")
                st.rerun()
                st.stop()
            
        else:
            st.info("Aucun dossier client ne correspond aux filtres.")


elif page == "Visa":
    st.header("🛂 Visa — Gestion des types")
    df = st.session_state.visa_df
    
    # Navigation CRUD dans la page Visa
    crud_mode = st.radio("Action Visa", ["Lister/Modifier/Supprimer", "Ajouter un nouveau type"], index=0, horizontal=True)

    if crud_mode == "Ajouter un nouveau type":
        st.subheader("Ajouter un nouveau type de visa")
        empty_row = pd.Series("", index=df.columns)
        render_visa_form(df, empty_row, action="add")
        
    elif crud_mode == "Lister/Modifier/Supprimer":
        
        st.dataframe(df, use_container_width=True)
        
        if len(df) > 0: 
            max_idx = len(df) - 1
            
            # --- ZONE CRITIQUE DE SÉLECTION D'INDEX STABILISÉE ---
            if 'visa_sel_idx' not in st.session_state:
                st.session_state.visa_sel_idx = 0
            
            current_index = st.session_state.visa_sel_idx
            
            # 1. CONTRÔLE D'INDEX ET CORRECTION CRITIQUE
            if current_index > max_idx or current_index < 0:
                 st.session_state.visa_sel_idx = min(max_idx, max(0, current_index))
                 st.warning("Index Visa réinitialisé (valeur hors limites).")
                 st.rerun()
                 
            # 2. L'index est garanti d'être valide ici.
            final_safe_index = st.session_state.visa_sel_idx
            
            # Clé statique pour éviter les problèmes de recréation de widget
            sel_idx_float = st.number_input(
                "Ouvrir visa (index affiché)", 
                min_value=0, 
                max_value=max_idx, 
                value=final_safe_index,
                key="visa_idx_input_static" 
            )
            
            sel_idx = int(sel_idx_float)
            
            # 3. Mettre à jour la session state (si l'utilisateur a changé la valeur)
            if sel_idx != final_safe_index:
                st.session_state.visa_sel_idx = sel_idx
                st.rerun()

            # --- DÉFENSE ULTIME CONTRE IndexError (try/except) ---
            try:
                # Accès aux données garanti
                sel_row = df.iloc[final_safe_index]
                
                st.subheader(f"Modifier Visa: {sel_row.get('Visa', 'N/A')}")
                
                # Ligne 317 (mise à jour)
                render_visa_form(df, sel_row, action="update", original_index=final_safe_index) 
            
            except IndexError as e:
                # Si l'index est désynchronisé (après une suppression rapide), on réinitialise et on relance
                st.session_state.visa_sel_idx = 0
                st.error("Erreur d'index détectée après modification. Redémarrage automatique.")
                st.rerun()
                st.stop()
            

        else:
            st.info("Aucun type de visa à gérer.")
        
# --- 4. DEFINITION DES FORMULAIRES (CRUD) ---

def render_client_form(df, sel_row, action, original_index=None):
    """Rendu du formulaire d'ajout/modification/suppression pour un client."""
    
    is_add = (action == "add")
    button_label = "Ajouter le dossier" if is_add else "Enregistrer les modifications"
    
    # 1. CLÉ DU FORMULAIRE : Unique par action et index
    unique_form_key = f"{action}_{original_index}" if action == 'update' and original_index is not None else f"{action}_new"

    with st.form(f"client_form_{unique_form_key}"):
        
        # Corps du formulaire CLIENTS
        cols1, cols2 = st.columns(2)
        with cols1:
            # Pour l'ajout, l'ID doit être modifiable ; pour la modification, il est figé (car il sert d'identifiant)
            dossier_id = st.text_input("DossierID", value=sel_row.get("DossierID", ""), disabled=not is_add)
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
        
        # Récupération des valeurs booléennes
        envoye = bool(sel_row.get("Dossier envoyé", False))
        refuse = bool(sel_row.get("Dossier refusé", False))
        approuve = bool(sel_row.get("Dossier approuvé", False))
        annule = bool(sel_row.get("DossierAnnule", False))
        rfe_val = bool(sel_row.get("RFE", False))
        
        with st_col1:
            dossier_envoye = st.checkbox("Dossier envoyé", value=envoye)
            dossier_refuse = st.checkbox("Dossier refusé", value=refuse)
        with st_col2:
            dossier_approuve = st.checkbox("Dossier approuvé", value=approuve)
            dossier_annule = st.checkbox("DossierAnnule (annulé)", value=annule)
        with st_col3:
            rfe = st.checkbox("RFE (doit être combiné)", value=rfe_val)
            date_envoi = st.date_input("DateEnvoi", value=get_date_for_input("DateEnvoi", sel_row))

        st.markdown("---")
        
        payments_list = sel_row.get("Paiements", [])
        # S'assurer que payments_list est une liste propre (au cas où ce serait un string/NaN)
        if isinstance(payments_list, str):
            try:
                payments_list = json.loads(payments_list) if payments_list and pd.notna(payments_list) else []
            except Exception:
                payments_list = []
        elif not isinstance(payments_list, list):
             payments_list = []

        st.write("Paiements (Total encaissé: " + f"{sel_row.get('TotalAcomptes', 0.0):.2f} €" + ")")
        
        for i, p in enumerate(payments_list):
            p_date = p.get('date', 'N/A')
            p_amount = p.get('amount', 0)
            st.markdown(f"**{i+1}. {p_date}** — {p_amount:.2f} €")

        st.markdown("---")
        st.write("Ajouter un nouveau paiement")
        col_pay1, col_pay2 = st.columns(2)
        
        # 2. CLÉS DE WIDGETS DE PAIEMENT : Utilisation de la clé unique du formulaire
        pay_date_key = f"pay_date_{unique_form_key}"
        pay_amount_key = f"pay_amount_{unique_form_key}"
        
        with col_pay1:
            new_pay_date = st.date_input("Date du paiement", value=date.today(), key=pay_date_key)
        with col_pay2:
            new_pay_amount = st.number_input("Montant", value=0.0, min_value=0.0, format="%.2f", key=pay_amount_key)


        # Boutons d'action
        col_buttons = st.columns(3)
        submitted = col_buttons[0].form_submit_button(button_label)
        
        delete_button = None
        if not is_add:
            delete_button = col_buttons[1].form_submit_button("❌ Supprimer le dossier")

        if submitted:
            # 3. RÉCUPÉRATION DES VALEURS DE WIDGETS : Utilisation des nouvelles clés uniques
            final_pay_amount = st.session_state.get(pay_amount_key, 0.0)
            final_pay_date = st.session_state.get(pay_date_key, date.today())
            
            if not dossier_id and is_add:
                 st.error("Veuillez entrer un DossierID pour l'ajout.")
            else:
                update_client_data(df, sel_row, original_index, {
                    "DossierID": dossier_id, "Nom": nom, "TypeVisa": typevisa, "Email": email,
                    "Telephone": telephone, "Honoraires": float(honoraires), "Notes": notes,
                    "Dossier envoyé": dossier_envoye, "Dossier refusé": dossier_refuse,
                    "Dossier approuvé": dossier_approuve, "DossierAnnule": dossier_annule,
                    "RFE": rfe, "DateEnvoi": date_envoi, 
                    "Paiements_New_Amount": float(final_pay_amount), 
                    "Paiements_New_Date": final_pay_date
                }, action)
        
        if not is_add and delete_button:
            update_client_data(df, sel_row, original_index, {}, "delete")


def update_client_data(df, sel_row, original_index, form_data, action):
    """Logique de mise à jour/ajout/suppression pour les clients."""
    
    if action == "delete":
        st.session_state.clients_df = st.session_state.clients_df.drop(original_index, axis=0)
        st.session_state.clients_df = compute_finances(st.session_state.clients_df)
        # Réinitialiser l'index pour se placer sur le premier élément (0) si la liste n'est pas vide
        if 'client_sel_idx' in st.session_state:
             st.session_state.client_sel_idx = 0 
        st.success("Dossier client supprimé.")
        st.rerun()
        return 

    # Préparation des données mises à jour (inchangé)
    updated = sel_row.copy()
    
    for key, value in form_data.items():
        if not key.startswith("Paiements_New"):
            updated[key] = value

    # Gestion des paiements (inchangé)
    current_payments_list = updated.get("Paiements", [])
    if isinstance(current_payments_list, str): 
        current_payments_list = []
        
    new_pay_amount = form_data.get("Paiements_New_Amount", 0.0)
    new_pay_date = form_data.get("Paiements_New_Date", date.today())
    
    if new_pay_amount and float(new_pay_amount) > 0:
        current_payments_list.append({"date": str(new_pay_date), "amount": float(new_pay_amount)})

    updated["Paiements"] = current_payments_list.copy()
    
    # Validation (inchangé)
    ok, msg = validate_rfe_row(updated)
    if not ok:
        st.error(msg)
        return

    # Enregistrement (inchangé)
    if action == "update":
        st.session_state.clients_df.loc[original_index, updated.index] = updated.astype(object)
        st.success("Modifications client enregistrées.")
    elif action == "add":
        new_row_df = pd.DataFrame([updated])
        st.session_state.clients_df = pd.concat([st.session_state.clients_df, new_row_df], ignore_index=True)
        st.success("Nouveau dossier client ajouté.")

    # Recalculer les finances et relancer (inchangé)
    st.session_state.clients_df = compute_finances(st.session_state.clients_df)
    st.rerun()


def render_visa_form(df, sel_row, action, original_index=None):
    """Rendu du formulaire d'ajout/modification/suppression pour un type de visa."""
    
    is_add = (action == "add")
    button_label = "Ajouter le type" if is_add else "Enregistrer les modifications"

    # 1. CLÉ DU FORMULAIRE : Unique par action et index
    unique_form_key = f"{action}_{original_index}" if action == 'update' and original_index is not None else f"{action}_new"

    with st.form(f"visa_form_{unique_form_key}"):
        
        # Corps du formulaire VISAS
        visa_code = st.text_input("Code Visa", value=sel_row.get("Visa", ""), disabled=not is_add)
        category = st.text_input("Catégorie", value=sel_row.get("Categories", ""))
        definition = st.text_area("Définition", value=sel_row.get("Definition", ""))

        # Boutons d'action
        col_buttons = st.columns(3)
        submitted = col_buttons[0].form_submit_button(button_label)
        
        delete_button = None
        if not is_add:
            delete_button = col_buttons[1].form_submit_button("❌ Supprimer le type")

        if submitted:
            if not visa_code:
                 st.error("Veuillez entrer un Code Visa.")
                 return

            if action == "add" and visa_code in st.session_state.visa_df['Visa'].values:
                 st.error(f"Le code Visa '{visa_code}' existe déjà. Veuillez modifier l'entrée existante.")
                 return

            # Préparation des données
            updated = sel_row.copy()
            updated["Visa"] = visa_code
            updated["Categories"] = category
            updated["Definition"] = definition
            
            # Enregistrement
            if action == "update":
                st.session_state.visa_df.loc[original_index, :] = updated
                st.success("Type de visa modifié.")
            elif action == "add":
                new_row_df = pd.DataFrame([updated])
                st.session_state.visa_df = pd.concat([st.session_state.visa_df, new_row_df], ignore_index=True)
                st.success("Nouveau type de visa ajouté.")
            
            st.rerun()
        
        if not is_add and delete_button:
            st.session_state.visa_df = st.session_state.visa_df.drop(original_index, axis=0)
            st.session_state.visa_df = st.session_state.visa_df.reset_index(drop=True)
            # Réinitialiser l'index
            if 'visa_sel_idx' in st.session_state:
                 st.session_state.visa_sel_idx = 0 
            st.success("Type de visa supprimé.")
            st.rerun()
            return 

# --- 5. LOGIQUE DE SAUVEGARDE GLOBALE (inchangé) ---

if src and (page == "Clients" or page == "Visa"):
    
    st.markdown("---")
    st.subheader("Exporter et Sauvegarder les Données")
    
    exp_col1, exp_col2, exp_col3 = st.columns(3)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    
    # Préparation du DataFrame pour l'export (Paiements en JSON string)
    clients_df_export = st.session_state.clients_df.copy()
    clients_df_export["Paiements"] = clients_df_export["Paiements"].apply(json.dumps)
    
    # Création du dictionnaire d'onglets pour l'export XLSX
    all_sheets_export = {
        "Clients": clients_df_export,
        "Visa": st.session_state.visa_df
    }

    with exp_col1:
        # Téléchargement CSV Clients
        csv_bytes = clients_df_export.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Télécharger CSV — Clients", data=csv_bytes, file_name=f"Clients_{stamp}.csv", mime="text/csv")
    
    with exp_col2:
        # Téléchargement XLSX Classeur
        xls_bytes = to_excel_bytes_multi(all_sheets_export)
        st.download_button("⬇️ Télécharger XLSX — Classeur", data=xls_bytes, file_name=f"Visa_Clients_{stamp}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    with exp_col3:
        if save_mode == "Save to local path (serveur/PC)":
            if save_path:
                try:
                    xls_bytes = to_excel_bytes_multi(all_sheets_export)
                    with open(save_path, "wb") as f:
                        f.write(xls_bytes)
                    st.success(f"Fichier écrit: {save_path}")
                except Exception as e:
                    st.error(f"Erreur écriture locale: {e}")
            else:
                st.warning("Renseignez un chemin local dans la sidebar.")
        elif save_mode in ["Google Drive (secrets req.)", "OneDrive (secrets req.)"]:
            st.info("Les modes de sauvegarde avancés nécessitent une configuration spécifique des secrets/API.")
