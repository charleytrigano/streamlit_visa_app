# utils.py
import io
import json
import unicodedata
from typing import Dict, List, Tuple
from datetime import date # Import nécessaire pour la migration des paiements

import pandas as pd
import streamlit as st

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

def _as_bool_series(s: pd.Series) -> pd.Series:
    """Convertit une colonne en booléens de manière robuste."""
    if s is None:
        return pd.Series([], dtype=bool)
    vals = s.astype(str).str.strip().str.lower()
    truthy = {"1", "true", "vrai", "yes", "oui", "y", "o", "x", "✓", "checked"}
    falsy = {"0", "false", "faux", "no", "non", "n", "", "none", "nan"}
    out = vals.apply(lambda v: True if v in truthy else (False if v in falsy else pd.NA))
    return out.fillna(False)

@st.cache_data(show_spinner=False)
def load_all_sheets(xlsx_input) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    """Charge toutes les feuilles d'un fichier Excel."""
    xls = pd.ExcelFile(xlsx_input)
    out = {}
    for name in xls.sheet_names:
        _df = pd.read_excel(xls, sheet_name=name)
        _df.columns = _norm_cols(_df.columns)
        out[name] = _df
    return out, xls.sheet_names

@st.cache_data(show_spinner=False)
def to_excel_bytes_multi(sheets: Dict[str, pd.DataFrame]) -> bytes:
    """Convertit un dictionnaire de DataFrames en un fichier Excel binaire."""
    import openpyxl  # noqa: F401
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, _df in sheets.items():
            # Limite le nom de l'onglet à 31 caractères
            _df.to_excel(writer, index=False, sheet_name=name[:31])
    return buffer.getvalue()

def _parse_payments_to_list(cell):
    """Analyse une cellule 'Paiements' (string JSON ou liste) en une liste de dicts."""
    try:
        if isinstance(cell, list):
            return cell
        if isinstance(cell, str):
            s = cell.strip()
            if not s:
                return []
            # Tente de charger JSON
            try:
                parsed = json.loads(s)
                return parsed if isinstance(parsed, list) else []
            except json.JSONDecodeError:
                # Si non-JSON valide, retourne vide
                return []
        if pd.isna(cell):
            return []
    except Exception:
        # En cas d'erreur inattendue
        return []
    # Autres types -> vide
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
        "Dossier envoyé": ["Dossier envoye"],
        "Dossier refusé": ["Dossier refuse"],
        "Dossier approuvé": ["Dossier approuve"],
        "DossierAnnule": ["Dossier Annule", "Dossier annulé"],
    }
    
    cols_to_drop = []

    for standard_col, possible_names in col_std_mapping.items():
        if standard_col not in df.columns:
            # Chercher un nom alternatif à renommer
            found_col = _find_col(possible_names, columns)
            if found_col:
                df = df.rename(columns={found_col: standard_col})
        else:
            # La colonne standard existe, chercher et marquer les doublons pour suppression
            for name in possible_names:
                found_col = _find_col([name], columns)
                if found_col and found_col != standard_col:
                    cols_to_drop.append(found_col)

    # Re-obtenir la liste des colonnes après le renommage
    columns = list(df.columns) 

    # --- 2. Migration des anciens paiements ---
    
    # S'assurer que 'Paiements' est présent et formaté comme liste de dicts
    if "Paiements" not in df.columns:
         df["Paiements"] = df.apply(lambda x: [], axis=1) # Initialiser avec liste vide
    
    # Assurer que Paiements existants sont lus comme une liste pour l'état initial
    df["Paiements"] = df["Paiements"].apply(_parse_payments_to_list)
    
    # Identifier les lignes où 'Paiements' est vide (à migrer)
    no_payments_mask = df["Paiements"].apply(lambda x: len(x) == 0)

    legacy_payments = {idx: [] for idx in df.index[no_payments_mask]}
    legacy_pay_cols = []
    
    # Colonnes d'acomptes
    for i in range(1, 4): # Acompte 1, 2, 3
        date_col_name = _find_col([f"Date Acompte {i}"], columns)
        amount_col_name = _find_col([f"Acompte {i}"], columns)
        
        if date_col_name and amount_col_name:
            # Marquer pour suppression, si elles ne sont pas déjà dans la liste
            legacy_pay_cols.extend([c for c in [date_col_name, amount_col_name] if c not in legacy_pay_cols])
            
            for idx in legacy_payments.keys():
                date_val = df.loc[idx, date_col_name]
                amount_val = df.loc[idx, amount_col_name]
                
                try:
                    pay_amount = float(amount_val) if pd.notna(amount_val) else 0.0
                    
                    if pay_amount > 0:
                        # Conversion de date robuste
                        dt_obj = pd.to_datetime(date_val, errors='coerce')
                        pay_date = str(dt_obj.date()) if pd.notna(dt_obj) else str(date.today())
                        
                        payment = {"date": pay_date, "amount": pay_amount}
                        legacy_payments[idx].append(payment)
                except Exception:
                    pass # Ignore les lignes avec données illisibles

    # Mettre à jour la colonne 'Paiements' avec les données migrées (liste de dicts)
    for idx, payments_list in legacy_payments.items():
        if payments_list:
            df.loc[idx, "Paiements"] = payments_list

    # --- 3. Nettoyage final des colonnes ---
    all_cols_to_drop = list(set(cols_to_drop + legacy_pay_cols))
    df = df.drop(columns=[c for c in all_cols_to_drop if c in df.columns], errors='ignore')
    
    return df

# ... (Le reste de utils.py: compute_finances, validate_rfe_row restent inchangés) ...

def compute_finances(df: pd.DataFrame) -> pd.DataFrame:
# ... (Fonction inchangée) ...
# ... (La fonction originale est conservée, elle est maintenant plus robuste car harmonize_clients_df nettoie les paiements avant)
