# Début de utils.py
import io
import json
import unicodedata
from typing import Dict, List, Tuple
from datetime import date 

import pandas as pd
import streamlit as st # <-- L'importation de Streamlit est correcte ici
# ... (le reste du fichier)

def _norm_cols(cols: List[str]) -> List[str]:
    """Nettoie les noms de colonnes (enlève espaces)"""
    return [str(c).strip() for c in cols]

def _find_col(possible_names: List[str], columns: List[str]):
    """Recherche un nom de colonne en ignorant la casse et les accents."""
    def norm(s: str) -> str:
        s = str(s)
        # Normalisation pour enlever les accents (ex: "é" -> "e")
        s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
        return s.lower().strip()
    cols_norm = {norm(c): c for c in columns}
    for name in possible_names:
        key = norm(name)
        if key in cols_norm:
            return cols_norm[key]
    return None

# ... (Autres fonctions _as_bool_series, load_all_sheets, to_excel_bytes_multi inchangées) ...

def _parse_payments_to_list(cell):
    """Analyse une cellule 'Paiements' (string JSON ou liste) en une liste de dicts."""
    try:
        if isinstance(cell, list):
            return cell
        if isinstance(cell, str):
            s = cell.strip()
            if not s:
                return []
            try:
                parsed = json.loads(s)
                return parsed if isinstance(parsed, list) else []
            except json.JSONDecodeError:
                return []
        if pd.isna(cell):
            return []
    except Exception:
        return []
    return []

def harmonize_clients_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    1. Standardise les noms de colonnes (ex: Dossier -> DossierID).
    2. Migre les anciens paiements (Date Acompte X / Acompte X) vers 'Paiements'.
    3. Supprime les colonnes dupliquées/anciennes.
    """
    df = df.copy()
    columns = list(df.columns)
    
    # --- 1. Standardisation des noms de colonnes ---
    col_std_mapping = {
        "DossierID": ["Dossier"],
        "DateCreation": ["Date"],
        "TypeVisa": ["Type Visa"],
        "DateFacture": ["Date facture"],
        "DateEnvoi": ["Date envoi"],
        "DateRetour": ["Date retour"],
        "Dossier envoyé": ["Dossier envoye"],
        "Dossier refusé": ["Dossier refuse"],
        "Dossier approuvé": ["Dossier approuve"],
        "DossierAnnule": ["Dossier Annule", "Dossier annulé"],
    }
    
    cols_to_drop = []

    for standard_col, possible_names in col_std_mapping.items():
        if standard_col not in df.columns:
            found_col = _find_col(possible_names, columns)
            if found_col:
                df = df.rename(columns={found_col: standard_col})
        else:
            for name in possible_names:
                found_col = _find_col([name], columns)
                if found_col and found_col != standard_col:
                    cols_to_drop.append(found_col)

    columns = list(df.columns) 

    # --- 2. Migration des anciens paiements ---
    
    # S'assurer que 'Paiements' est présent et formaté comme liste de dicts
    if "Paiements" not in df.columns:
         # Initialisation avec dtype=object pour garantir qu'elle peut contenir des listes de dicts.
         df["Paiements"] = pd.Series([[] for _ in range(len(df))], index=df.index, dtype=object) 
    
    # Assurer que Paiements existants sont lus comme une liste pour l'état initial
    df["Paiements"] = df["Paiements"].apply(_parse_payments_to_list)
    
    # ********** CORRECTIF DU VALUER ERROR **********
    # Forcer le type 'object'. Cela garantit que pandas autorise l'affectation de listes de dicts 
    # à des cellules individuelles via .loc.
    df["Paiements"] = df["Paiements"].astype(object) 
    # **********************************************
    
    # Identifier les lignes où 'Paiements' est vide (ce sont les lignes à migrer)
    no_payments_mask = df["Paiements"].apply(lambda x: len(x) == 0)

    legacy_payments = {idx: [] for idx in df.index[no_payments_mask]}
    legacy_pay_cols = []
    
    for i in range(1, 6): 
        date_col_name = _find_col([f"Date Acompte {i}"], columns)
        amount_col_name = _find_col([f"Acompte {i}"], columns)
        
        if date_col_name and amount_col_name:
            legacy_pay_cols.extend([c for c in [date_col_name, amount_col_name] if c not in legacy_pay_cols])
            
            for idx in legacy_payments.keys():
                date_val = df.loc[idx, date_col_name]
                amount_val = df.loc[idx, amount_col_name]
                
                try:
                    pay_amount = float(amount_val) if pd.notna(amount_val) else 0.0
                    
                    if pay_amount > 0:
                        dt_obj = pd.to_datetime(date_val, errors='coerce')
                        pay_date = str(dt_obj.date()) if pd.notna(dt_obj) else str(date.today()) 
                        
                        payment = {"date": pay_date, "amount": pay_amount}
                        legacy_payments[idx].append(payment)
                except Exception:
                    pass 

    # Mettre à jour la colonne 'Paiements' avec les données migrées (liste de dicts)
    for idx, payments_list in legacy_payments.items():
        if payments_list:
            df.loc[idx, "Paiements"] = payments_list

    # --- 3. Nettoyage final des colonnes ---
    all_cols_to_drop = list(set(cols_to_drop + legacy_pay_cols))
    df = df.drop(columns=[c for c in all_cols_to_drop if c in df.columns], errors='ignore')
    
    return df

def compute_finances(df: pd.DataFrame) -> pd.DataFrame:
    # ... (Reste de la fonction inchangée) ...
    # ...
    return df

def validate_rfe_row(row: pd.Series) -> Tuple[bool, str]:
    # ... (Reste de la fonction inchangée) ...
    # ...
    return True, ""

