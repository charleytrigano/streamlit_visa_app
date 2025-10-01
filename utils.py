# utils.py
import io
import json
import unicodedata
from typing import Dict, List, Tuple

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
    # Autres types (int, float, etc.) -> vide
    return []

def compute_finances(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcule 'TotalAcomptes' et 'SoldeCalc' à partir de 'Honoraires' et 'Paiements'.
    Assure l'existence et le format numérique des colonnes clés.
    """
    df = df.copy()

    # 1. Honoraires: S'assurer qu'elle existe et est numérique
    if "Honoraires" not in df.columns:
        df["Honoraires"] = 0.0
    df["Honoraires"] = pd.to_numeric(df["Honoraires"], errors="coerce").fillna(0.0)

    # 2. Paiements: S'assurer qu'elle existe
    if "Paiements" not in df.columns:
        df["Paiements"] = "[]"

    # 3. Calcul de TotalAcomptes
    def sum_payments(payments_list):
        total = 0.0
        for p in payments_list:
            try:
                # Récupère le montant, gère les clés manquantes ou les valeurs non numériques
                amt = float(p.get("amount", 0) or 0) if isinstance(p, dict) else float(p)
            except Exception:
                amt = 0.0
            total += amt
        return total

    # Applique d'abord le parsing, puis la somme
    df["_ParsedPaiements"] = df["Paiements"].apply(_parse_payments_to_list)
    df["TotalAcomptes"] = df["_ParsedPaiements"].apply(sum_payments)
    
    # Assurer que TotalAcomptes est propre
    df["TotalAcomptes"] = pd.to_numeric(df["TotalAcomptes"], errors="coerce").fillna(0.0)
    
    # 4. SoldeCalc
    df["SoldeCalc"] = (df["Honoraires"] - df["TotalAcomptes"]).round(2)
    
    # 5. Mettre à jour la colonne 'Paiements' avec la version liste pour éviter les re-parsings inutiles dans la session
    df["Paiements"] = df["_ParsedPaiements"]
    df = df.drop(columns=["_ParsedPaiements"])

    return df

def validate_rfe_row(row: pd.Series) -> Tuple[bool, str]:
    """Valide la cohérence des statuts d'un dossier (RFE, Envoyé, Refusé, Approuvé, Annulé)."""
    # Utiliser .get() avec _as_bool_series implicitement sur les colonnes pour une robustesse maximale
    rfe = bool(row.get("RFE", False))
    sent = bool(row.get("Dossier envoyé", False) or row.get("Dossier envoye", False))
    refused = bool(row.get("Dossier refusé", False) or row.get("Dossier refuse", False))
    approved = bool(row.get("Dossier approuvé", False) or row.get("Dossier approuve", False))
    canceled = bool(row.get("DossierAnnule", False) or row.get("Dossier Annule", False) or row.get("Dossier annulé", False))

    if rfe and not (sent or refused or approved):
        return False, "RFE doit être combinée avec Envoyé / Refusé / Approuvé"
    if approved and refused:
        return False, "Un dossier ne peut pas être à la fois Approuvé et Refusé"
    if canceled and (sent or refused or approved):
        return False, "Un dossier annulé ne peut pas être marqué Envoyé/Refusé/Approuvé"
    
    return True, ""