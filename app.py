# app.py — Version finale avec sauvegarde locale explicite (Corrigé 32)
import json
import os # Ajout pour les opérations de fichier
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
        clients_key = next((k for k in all_sheets.keys() if "client" in str(k).lower()), None)
        if clients_key:
            clients_df_loaded = all_sheets.get(clients_key)
            if clients_df_loaded is not None:
                st.info(f"Onglet 'Clients' non trouvé. Utilisation de l'onglet '{clients_key}'.")

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
                default_val = pd.NaT if c.startswith("Date") else (0.0 if c == "Honoraires" else ([] if c == "Paiements" else ""))
                clients_df_loaded[c] = default_val

    # FORCER TOUTES LES COLONNES DE DATE EN DATETIME
    date_cols = [c for c in base_cols_clients if c.startswith("Date")]
    for col in date_cols:
        clients_df_loaded[col] = pd.to_datetime(clients_df_loaded.get(col), errors='coerce')
        
    # S'assurer que 'DateCreation' est une date valide pour les filtres
    clients_df_loaded['DateCreation'] = clients_df_loaded['DateCreation'].fillna(pd.Timestamp(date.today()))
        
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
        visa_df_loaded = visa_df_loaded.rename(columns={c:c for c in visa_df_loaded.columns if c in base_cols_visa})
        for col in base_cols_visa:
            if col not in visa_df_loaded.columns:
                visa_df_loaded[col] = ""
        st.session_state.visa_df = visa_df_loaded[base_cols_visa].copy()

    if not st.session_state.visa_df.index.name:
        st.session_state.visa_df = st.session_state.visa_df.reset_index(drop=True)


def get_date_for_input(col_name, row):
    """Fonction utilitaire pour formatter les dates pour les date_input de Streamlit."""
    dt = row.get(col_name)
    if pd.notna(dt) and isinstance(dt, (datetime, date, pd.Timestamp)):
        return dt.date()
    return date.today()
    
def get_export_data():
    """Prépare les DataFrames pour l'exportation et la sauvegarde."""
    clients_df_export = st.session_state.clients_df.copy()
    # Sérialiser la colonne 'Paiements' en JSON pour l'enregistrement Excel
    clients_df_export["Paiements"] = clients_df_export["Paiements"].apply(lambda x: json.dumps(x) if isinstance(x, list) else x)
    
    all_sheets_export = {
        "Clients": clients_df_export,
        "Visa": st.session_state.visa_df
    }
    return all_sheets_export

def save_to_local_path(all_sheets_export, path):
    """Écrit le fichier XLSX à l'emplacement local spécifié."""
    if not path:
        return False, "Le chemin de sauvegarde local n'est pas spécifié."
        
    try:
        # Assurez-vous que le répertoire existe
        dir_name = os.path.dirname(path)
        if dir_name and not os.path.exists(dir_name):
            os.makedirs(dir_name)
            
        xls_bytes = to_excel_bytes_multi(all_sheets_export)
        with open(path, "wb") as f:
            f.write(xls_bytes)
        
        return True, f"Fichier sauvegardé avec succès à : **{path}**"
    except Exception as e:
        return False, f"Erreur d'écriture locale. Vérifiez le chemin et les permissions : {e}"

# --- 2. LOGIQUE PRINCIPALE DE L'APPLICATION ---

# Sidebar / source / save options
with st.sidebar:
    st.header("Fichier source & sauvegarde")
    up = st.file_uploader("Fichier .xlsx", type=["xlsx"], help="Classeur contenant 'Visa' et 'Clients'.")
    data_path = st.text_input("Ou chemin local vers le .xlsx (optionnel)") 
    st.markdown("---")
    st.subheader("Sauvegarde")
    save_mode = st.selectbox("Mode de sauvegarde", ["Download (toujours disponible)", "Save to local path (serveur/PC)"])
    
    # 💥 CHEMIN LOCAL PAR DÉFAUT + INFO
    default_save_path = os.path.join(os.getcwd(), "data_sauvegardee.xlsx")
    save_path = st.text_input("Chemin local (ex: C:\\Users\\...\\data.xlsx)", value=default_save_path)
    
    # 💥 BOUTON DE SAUVEGARDE LOCALE EXPLICITE
    if save_mode == "Save to local path (serveur/PC)":
        if st.button("💾 SAUVEGARDER MAINTENANT (Local)"):
            all_sheets_export = get_export_data()
            success, message = save_to_local_path(all_sheets_export, save_path)
            
            if success:
                st.success(message)
            else:
                st.error(message)

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

# --- GESTION DE LA VUE PAR ÉTAT ---
if "current_view" not in st.session_state:
     st.session_state.current_view = "clients_list"

def set_view(view_name):
    """Callback pour changer la vue et relancer l'application."""
    st.session_state.current_view = view_name
    if view_name in ["clients_list", "visa_list"]:
        st.session_state.pop("client_sel_idx", None)
        st.session_state.pop("visa_sel_idx", None)


# Navigation principale (mise à jour pour gérer le changement de vue)
page = st.selectbox("Page", ["Clients", "Visa"], index=0, 
                    key="main_page_select", 
                    on_change=lambda: set_view("clients_list" if st.session_state.main_page_select == "Clients" else "visa_list"))

# ----------------------------------------------------------------------
# --- 3. DEFINITION DES FORMULAIRES (CRUD) ---
# ----------------------------------------------------------------------

def render_client_form(df, sel_row, action, original_index=None):
    """Rendu du formulaire d'ajout/modification/suppression pour un client."""
    
    is_add = (action == "add")
    button_label = "Ajouter le dossier" if is_add else "Enregistrer les modifications"
    
    unique_form_key = f"{action}_{original_index}" if action == 'update' and original_index is not None else f"{action}_new"

    with st.form(f"client_form_{unique_form_key}"):
        
        honoraires_val = sel_row.get("Honoraires", 0.0)
        if honoraires_val == "" or pd.isna(honoraires_val):
            honoraires_default = 0.0
        else:
            honoraires_default = float(honoraires_val)

        cols1, cols2 = st.columns(2)
        with cols1:
            dossier_id = st.text_input("DossierID", value=sel_row.get("DossierID", ""), disabled=not is_add, key=f"dossier_id_{unique_form_key}")
            nom = st.text_input("Nom", value=sel_row.get("Nom", ""), key=f"nom_{unique_form_key}")
            typevisa = st.text_input("TypeVisa", value=sel_row.get("TypeVisa", ""), key=f"typevisa_{unique_form_key}")
            email = st.text_input("Email", value=sel_row.get("Email", ""), key=f"email_{unique_form_key}")
        with cols2:
            telephone = st.text_input("Telephone", value=sel_row.get("Telephone", ""), key=f"telephone_{unique_form_key}")
            honoraires = st.number_input("Honoraires", value=honoraires_default, format="%.2f", key=f"honoraires_{unique_form_key}")
            notes = st.text_area("Notes", value=sel_row.get("Notes", ""), key=f"notes_{unique_form_key}")
        
        st.markdown("---")
        st.write("Statuts / dates")
        st_col1, st_col2, st_col3 = st.columns(3)
        
        envoye = bool(sel_row.get("Dossier envoyé", False))
        refuse = bool(sel_row.get("Dossier refusé", False))
        approuve = bool(sel_row.get("Dossier approuvé", False))
        annule = bool(sel_row.get("DossierAnnule", False))
        rfe_val = bool(sel_row.get("RFE", False))
        
        with st_col1:
            dossier_envoye = st.checkbox("Dossier envoyé", value=envoye, key=f"envoye_{unique_form_key}")
            dossier_refuse = st.checkbox("Dossier refusé", value=refuse, key=f"refuse_{unique_form_key}")
        with st_col2:
            dossier_approuve = st.checkbox("Dossier approuvé", value=approuve, key=f"approuve_{unique_form_key}")
            dossier_annule = st.checkbox("DossierAnnule (annulé)", value=annule, key=f"annule_{unique_form_key}")
        with st_col3:
            rfe = st.checkbox("RFE (doit être combiné)", value=rfe_val, key=f"rfe_{unique_form_key}")
            date_envoi = st.date_input("DateEnvoi", value=get_date_for_input("DateEnvoi", sel_row), key=f"date_envoi_{unique_form_key}")

        st.markdown("---")
        
        payments_list = sel_row.get("Paiements", [])
        if isinstance(payments_list, str):
            try:
                payments_list = json.loads(payments_list) if payments_list and pd.notna(payments_list) else []
            except Exception:
                payments_list = []
        elif not isinstance(payments_list, list):
             payments_list = []

        total_payed_val = sel_row.get('TotalAcomptes', 0.0)
        try:
            if pd.isna(total_payed_val) or total_payed_val == "":
                total_payed_safe = 0.0
            else:
                total_payed_safe = float(total_payed_val)
        except (TypeError, ValueError):
            total_payed_safe = 0.0
            
        st.write("Paiements (Total encaissé: " + f"{total_payed_safe:.2f} €" + ")")
        
        for i, p in enumerate(payments_list):
            p_date = p.get('date', 'N/A')
            p_amount = p.get('amount', 0)
            st.markdown(f"**{i+1}. {p_date}** — {p_amount:.2f} €")

        st.markdown("---")
        st.write("Ajouter un nouveau paiement")
        col_pay1, col_pay2 = st.columns(2)
        
        pay_date_key = f"pay_date_{unique_form_key}"
        pay_amount_key = f"pay_amount_{unique_form_key}"
        
        with col_pay1:
            new_pay_date = st.date_input("Date du paiement", value=date.today(), key=pay_date_key)
        with col_pay2:
            new_pay_amount = st.number_input("Montant", value=0.0, min_value=0.0, format="%.2f", key=pay_amount_key)


        col_buttons = st.columns(3)
        submitted = col_buttons[0].form_submit_button(button_label) 
        
        delete_button = None
        if not is_add:
            delete_button = col_buttons[1].form_submit_button("❌ Supprimer le dossier")

        if submitted:
            
            final_dossier_id = st.session_state.get(f"dossier_id_{unique_form_key}")
            final_nom = st.session_state.get(f"nom_{unique_form_key}")
            final_typevisa = st.session_state.get(f"typevisa_{unique_form_key}")
            final_email = st.session_state.get(f"email_{unique_form_key}")
            final_telephone = st.session_state.get(f"telephone_{unique_form_key}")
            final_honoraires = st.session_state.get(f"honoraires_{unique_form_key}")
            final_notes = st.session_state.get(f"notes_{unique_form_key}")
            
            final_envoye = st.session_state.get(f"envoye_{unique_form_key}")
            final_refuse = st.session_state.get(f"refuse_{unique_form_key}")
            final_approuve = st.session_state.get(f"approuve_{unique_form_key}")
            final_annule = st.session_state.get(f"annule_{unique_form_key}")
            final_rfe = st.session_state.get(f"rfe_{unique_form_key}")
            final_date_envoi = st.session_state.get(f"date_envoi_{unique_form_key}")
            
            final_pay_amount = st.session_state.get(pay_amount_key, 0.0)
            final_pay_date = st.session_state.get(pay_date_key, date.today())
            
            final_dossier_id = final_dossier_id if final_dossier_id is not None else sel_row.get("DossierID", "")

            if not final_dossier_id and is_add:
                 st.error("Veuillez entrer un DossierID pour l'ajout.")
            else:
                update_client_data(df, sel_row, original_index, {
                    "DossierID": final_dossier_id, "Nom": final_nom, "TypeVisa": final_typevisa, "Email": final_email,
                    "Telephone": final_telephone, "Honoraires": float(final_honoraires), "Notes": final_notes,
                    "Dossier envoyé": final_envoye, "Dossier refusé": final_refuse,
                    "Dossier approuvé": final_approuve, "DossierAnnule": final_annule,
                    "RFE": final_rfe, "DateEnvoi": final_date_envoi, 
                    "Paiements_New_Amount": float(final_pay_amount), 
                    "Paiements_New_Date": final_pay_date
                }, action)
        
        if not is_add and delete_button:
            update_client_data(df, sel_row, original_index, {}, "delete")


def update_client_data(df, sel_row, original_index, form_data, action):
    """Logique de mise à jour/ajout/suppression pour les clients."""
    
    date_cols_to_convert = ["DateCreation", "DateFacture", "DateEnvoi", "DateRetour", "DateAnnulation"]
    
    if action == "delete":
        st.session_state.clients_df = st.session_state.clients_df.drop(original_index, axis=0)
        st.session_state.clients_df = compute_finances(st.session_state.clients_df)
        st.session_state.client_sel_idx = 0 
        set_view("clients_list") 
        st.success("Dossier client supprimé.")
        st.rerun()
        return 

    updated = sel_row.copy()
    
    for key, value in form_data.items():
        if not key.startswith("Paiements_New"):
            updated[key] = value

    if action == "add" and "DateCreation" not in updated or pd.isna(updated.get("DateCreation")):
        updated["DateCreation"] = date.today()

    current_payments_list = updated.get("Paiements", [])
    if isinstance(current_payments_list, str): 
        current_payments_list = []
        
    new_pay_amount = form_data.get("Paiements_New_Amount", 0.0)
    new_pay_date = form_data.get("Paiements_New_Date", date.today())
    
    if new_pay_amount and float(new_pay_amount) > 0:
        current_payments_list.append({"date": str(new_pay_date), "amount": float(new_pay_amount)})

    updated["Paiements"] = current_payments_list.copy()
    
    ok, msg = validate_rfe_row(updated)
    if not ok:
        st.error(msg)
        return

    if action == "update":
        st.session_state.clients_df.loc[original_index, updated.index] = updated.astype(object)
        st.success("Modifications client enregistrées.")
    elif action == "add":
        new_row_df = pd.DataFrame([updated])
        for col in date_cols_to_convert:
            if col in new_row_df.columns:
                 new_row_df[col] = pd.to_datetime(new_row_df[col], errors='coerce')
        
        st.session_state.clients_df = pd.concat([st.session_state.clients_df, new_row_df], ignore_index=True)
        st.success("Nouveau dossier client ajouté.")

    for col in date_cols_to_convert:
        if col in st.session_state.clients_df.columns:
            st.session_state.clients_df[col] = pd.to_datetime(st.session_state.clients_df[col], errors='coerce')

    st.session_state.clients_df = compute_finances(st.session_state.clients_df)
    set_view("clients_list") 
    st.rerun()


def render_visa_form(df, sel_row, action, original_index=None):
    """Rendu du formulaire d'ajout/modification/suppression pour un type de visa."""
    
    is_add = (action == "add")
    button_label = "Ajouter le type" if is_add else "Enregistrer les modifications"

    unique_form_key = f"{action}_{original_index}" if action == 'update' and original_index is not None else f"{action}_new"

    with st.form(f"visa_form_{unique_form_key}"):
        
        visa_code = st.text_input("Code Visa", value=sel_row.get("Visa", ""), disabled=not is_add, key=f"visa_code_{unique_form_key}")
        category = st.text_input("Catégorie", value=sel_row.get("Categories", ""), key=f"category_{unique_form_key}")
        definition = st.text_area("Définition", value=sel_row.get("Definition", ""), key=f"definition_{unique_form_key}")

        col_buttons = st.columns(3)
        submitted = col_buttons[0].form_submit_button(button_label)
        
        delete_button = None
        if not is_add:
            delete_button = col_buttons[1].form_submit_button("❌ Supprimer le type")

        if submitted:
            
            final_visa_code = st.session_state.get(f"visa_code_{unique_form_key}")
            final_category = st.session_state.get(f"category_{unique_form_key}")
            final_definition = st.session_state.get(f"definition_{unique_form_key}")
            
            final_visa_code = final_visa_code if final_visa_code is not None else sel_row.get("Visa", "")

            if not final_visa_code:
                 st.error("Veuillez entrer un Code Visa.")
                 return

            if action == "add" and final_visa_code in st.session_state.visa_df['Visa'].values:
                 st.error(f"Le code Visa '{final_visa_code}' existe déjà. Veuillez modifier l'entrée existante.")
                 return

            updated = sel_row.copy()
            updated["Visa"] = final_visa_code
            updated["Categories"] = final_category
            updated["Definition"] = final_definition
            
            if action == "update":
                st.session_state.visa_df.loc[original_index, :] = updated
                st.success("Type de visa modifié.")
            elif action == "add":
                new_row_df = pd.DataFrame([updated])
                st.session_state.visa_df = pd.concat([st.session_state.visa_df, new_row_df], ignore_index=True)
                st.success("Nouveau type de visa ajouté.")
            
            set_view("visa_list") 
            st.rerun()
        
        if not is_add and delete_button:
            st.session_state.visa_df = st.session_state.visa_df.drop(original_index, axis=0)
            st.session_state.visa_df = st.session_state.visa_df.reset_index(drop=True)
            st.session_state.visa_sel_idx = 0 
            set_view("visa_list") 
            st.success("Type de visa supprimé.")
            st.rerun()
            return 
            
# ----------------------------------------------------------------------
# --- 4. RENDU DES PAGES (APPEL DES FONCTIONS) ---
# ----------------------------------------------------------------------

if page == "Clients":
    
    st.header("👥 Clients — gestion & suivi")
    df = st.session_state.clients_df
    current_view = st.session_state.current_view

    # --- CLIENTS : VUE AJOUT (add) ---
    if current_view == "clients_add":
        st.subheader("Ajouter un nouveau dossier client")
        empty_row = pd.Series("", index=df.columns)
        empty_row["Paiements"] = [] 
        
        render_client_form(df, empty_row, action="add")
        st.markdown("---")
        st.button("↩️ Retour à la liste des clients", on_click=lambda: set_view("clients_list"))


    # --- CLIENTS : VUE MODIFICATION (edit) ---
    elif current_view == "clients_edit":
        
        filtered = st.session_state.get("clients_filtered_df", df.copy())
        max_idx = len(filtered) - 1

        final_safe_index_filtered = st.session_state.get("client_sel_idx", 0)
        
        if final_safe_index_filtered < 0 or final_safe_index_filtered > max_idx:
             set_view("clients_list")
             st.rerun()

        try:
            sel_row_filtered = filtered.iloc[final_safe_index_filtered] 
            original_session_index = sel_row_filtered.name 

            st.subheader(f"Modifier Dossier: {sel_row_filtered.get('DossierID','(sans id)')} — {sel_row_filtered.get('Nom','')}")
            
            render_client_form(df, sel_row_filtered, action="update", original_index=original_session_index)
            st.markdown("---")
            st.button("↩️ Retour à la liste des clients", on_click=lambda: set_view("clients_list"))

        except IndexError as e:
            st.error("Dossier introuvable (IndexError). Retour à la liste.")
            set_view("clients_list")
            st.rerun()

    
    # --- CLIENTS : VUE LISTE (list) ---
    else: # current_view == "clients_list"
        
        total_dossiers = len(df)
        total_encaissé = df["TotalAcomptes"].sum()
        total_honoraires = df["Honoraires"].sum()
        total_solde = df["SoldeCalc"].sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total dossiers", f"{total_dossiers:,}")
        c2.metric("Total encaissé", f"{total_encaissé:,.2f} €")
        c3.metric("Total honoraires", f"{total_honoraires:,.2f} €")
        c4.metric("Solde total", f"{total_solde:,.2f} €")
        
        st.markdown("---")
        
        col_buttons_list = st.columns([1, 4])
        with col_buttons_list[0]:
             st.button("➕ Ajouter un nouveau dossier", on_click=lambda: set_view("clients_add"))

        st.markdown("---")

        # FILTRES MIS À JOUR
        with st.expander("Filtrer / Rechercher"):
            q = st.text_input("Recherche (nom / dossier / email)")
            
            col_date, col_visa = st.columns(2)
            
            df_temp = df.copy() 
            df_temp['Year'] = df_temp['DateCreation'].dt.year.fillna(0).astype(int)
            df_temp['Month'] = df_temp['DateCreation'].dt.month.fillna(0).astype(int)
            
            years = sorted(df_temp['Year'].unique().tolist(), reverse=True)
            months = {1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin",
                      7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"}
            
            with col_date:
                selected_year = st.selectbox("Filtrer par Année de création", ["Toutes"] + [y for y in years if y > 0])
                selected_month = st.selectbox("Filtrer par Mois de création", ["Tous"] + list(months.values()))
            
            with col_visa:
                visa_types = sorted(df_temp['TypeVisa'].dropna().unique().tolist())
                selected_visa = st.multiselect("Filtrer par Type de Visa", visa_types)
                
            status_filter = st.selectbox("Filtrer par Statut", ["Tous", "Envoyé", "Approuvé", "Refusé", "Annulé", "RFE"])

            st.markdown("---")
            st.subheader("Filtres financiers (€)")
            
            min_h = int(df['Honoraires'].min()) if not df['Honoraires'].empty and df['Honoraires'].min() is not np.nan else 0
            max_h = int(df['Honoraires'].max()) if not df['Honoraires'].empty and df['Honoraires'].max() is not np.nan else 0
            min_p = int(df['TotalAcomptes'].min()) if not df['TotalAcomptes'].empty and df['TotalAcomptes'].min() is not np.nan else 0
            max_p = int(df['TotalAcomptes'].max()) if not df['TotalAcomptes'].empty and df['TotalAcomptes'].max() is not np.nan else 0
            min_d = int(df['SoldeCalc'].min()) if not df['SoldeCalc'].empty and df['SoldeCalc'].min() is not np.nan else 0
            max_d = int(df['SoldeCalc'].max()) if not df['SoldeCalc'].empty and df['SoldeCalc'].max() is not np.nan else 0
            
            if min_h == max_h: max_h = 1 if max_h == 0 else (max_h + 1)
            if min_p == max_p: max_p = 1 if max_p == 0 else (max_p + 1)
            if min_d == max_d: max_d = 1 if max_d == 0 else (max_d + 1)
            
            h_range_default = (min_h, max_h)
            p_range_default = (min_p, max_p)
            d_range_default = (min_d, max_d)
            
            if min_h > max_h: h_range_default = (0, 0)
            if min_p > max_p: p_range_default = (0, 0)
            if min_d > max_d: d_range_default = (0, 0)
            
            col_h, col_p, col_d = st.columns(3)

            with col_h: honoraires_range = st.slider("Honoraires (Total)", min_h, max_h, h_range_default)
            with col_p: paye_range = st.slider("Montant Payé", min_p, max_p, p_range_default)
            with col_d: du_range = st.slider("Montant Dû (Solde)", min_d, max_d, d_range_default)


        # --- LOGIQUE D'APPLICATION DES FILTRES ---
        filtered = df_temp.copy() 
        
        if q:
            mask = pd.Series(False, index=filtered.index)
            for c in ["DossierID", "Nom", "Email", "TypeVisa"]:
                if c in filtered.columns:
                    mask = mask | filtered[c].astype(str).str.contains(q, case=False, na=False)
            filtered = filtered[mask]
        
        if selected_year != "Toutes": filtered = filtered[filtered['Year'] == int(selected_year)]
            
        if selected_month != "Tous":
            month_num = [k for k, v in months.items() if v == selected_month][0]
            filtered = filtered[filtered['Month'] == month_num]
            
        if selected_visa: filtered = filtered[filtered['TypeVisa'].isin(selected_visa)]

        if status_filter != "Tous":
            col_map = {"Envoyé": "Dossier envoyé", "Approuvé": "Dossier approuvé", "Refusé": "Dossier refusé", "Annulé": "DossierAnnule", "RFE": "RFE"}
            col_name = col_map.get(status_filter)
            if col_name: filtered = filtered[filtered.get(col_name, False) == True]
                 
        filtered = filtered[
            (filtered['Honoraires'] >= honoraires_range[0]) & (filtered['Honoraires'] <= honoraires_range[1])
        ]
        filtered = filtered[
            (filtered['TotalAcomptes'] >= paye_range[0]) & (filtered['TotalAcomptes'] <= paye_range[1])
        ]
        filtered = filtered[
            (filtered['SoldeCalc'] >= du_range[0]) & (filtered['SoldeCalc'] <= du_range[1])
        ]
        
        filtered = filtered.drop(columns=['Year', 'Month'], errors='ignore')

        st.dataframe(filtered.reset_index(drop=True).drop(columns=['Paiements', 'TotalAcomptes', 'SoldeCalc'], errors='ignore'), use_container_width=True)
        st.session_state.clients_filtered_df = filtered.copy() 

        if len(filtered) > 0: 
            max_idx = len(filtered) - 1
            current_index = st.session_state.get('client_sel_idx', 0)
            
            if current_index > max_idx or current_index < 0: current_index = 0
            final_safe_index = current_index

            sel_idx_float = st.number_input(
                "Index du dossier à modifier", 
                min_value=0, 
                max_value=max_idx, 
                value=final_safe_index, 
                key="client_idx_input_static" 
            )
            
            sel_idx = int(sel_idx_float) 
            
            if sel_idx != final_safe_index:
                st.session_state.client_sel_idx = sel_idx
                st.rerun() 
            else:
                 st.session_state.client_sel_idx = final_safe_index

            st.button("✏️ Modifier le dossier sélectionné", on_click=lambda: set_view("clients_edit"))
            
        else:
            st.info("Aucun dossier client ne correspond aux filtres.")


elif page == "Visa":
    st.header("🛂 Visa — Gestion des types")
    df = st.session_state.visa_df
    current_view = st.session_state.current_view
    
    if current_view == "visa_add":
        st.subheader("Ajouter un nouveau type de visa")
        empty_row = pd.Series("", index=df.columns)
        render_visa_form(df, empty_row, action="add")
        st.markdown("---")
        st.button("↩️ Retour à la liste des visas", on_click=lambda: set_view("visa_list"))
        
    elif current_view == "visa_edit":
        
        max_idx = len(df) - 1
        final_safe_index = st.session_state.get("visa_sel_idx", 0)

        if final_safe_index < 0 or final_safe_index > max_idx:
             set_view("visa_list")
             st.rerun()
             
        try:
            sel_row = df.iloc[final_safe_index]
            
            st.subheader(f"Modifier Visa: {sel_row.get('Visa', 'N/A')}")
            
            render_visa_form(df, sel_row, action="update", original_index=final_safe_index) 
            st.markdown("---")
            st.button("↩️ Retour à la liste des visas", on_click=lambda: set_view("visa_list"))
        
        except IndexError as e:
            st.error("Type de visa introuvable (IndexError). Retour à la liste.")
            set_view("visa_list")
            st.rerun()
            

    else: # current_view == "visa_list"
        
        st.dataframe(df, use_container_width=True)
        st.markdown("---")
        st.button("➕ Ajouter un nouveau type", on_click=lambda: set_view("visa_add"))
        st.markdown("---")
        
        if len(df) > 0: 
            max_idx = len(df) - 1
            current_index = st.session_state.get('visa_sel_idx', 0)
            
            if current_index > max_idx or current_index < 0: current_index = 0
                 
            final_safe_index = current_index
            
            sel_idx_float = st.number_input(
                "Index du visa à modifier", 
                min_value=0, 
                max_value=max_idx, 
                value=final_safe_index,
                key="visa_idx_input_static" 
            )
            
            sel_idx = int(sel_idx_float)
            
            if sel_idx != final_safe_index:
                st.session_state.visa_sel_idx = sel_idx
                st.rerun()
            else:
                 st.session_state.visa_sel_idx = final_safe_index

            st.button("✏️ Modifier le type de visa sélectionné", on_click=lambda: set_view("visa_edit"))

        else:
            st.info("Aucun type de visa à gérer.")
        
# --- 5. LOGIQUE DE SAUVEGARDE GLOBALE (Téléchargement) ---

if src and (page == "Clients" or page == "Visa"):
    
    st.markdown("---")
    st.subheader("Exporter et Télécharger les Données")
    
    exp_col1, exp_col2 = st.columns(2)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    
    all_sheets_export = get_export_data()
    xls_bytes = to_excel_bytes_multi(all_sheets_export)
    
    with exp_col1:
        # Téléchargement XLSX Classeur
        st.download_button("⬇️ Télécharger XLSX — Classeur", 
                           data=xls_bytes, 
                           file_name=f"Visa_Clients_Sauvegarde_{stamp}.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           help="Télécharge l'intégralité du classeur mis à jour.")
    
    with exp_col2:
        # Téléchargement CSV Clients
        clients_df_export = all_sheets_export.get("Clients").copy()
        csv_bytes = clients_df_export.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Télécharger CSV — Clients", 
                           data=csv_bytes, 
                           file_name=f"Clients_{stamp}.csv", 
                           mime="text/csv",
                           help="Télécharge seulement la liste des clients mise à jour au format CSV.")
