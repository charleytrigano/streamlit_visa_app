import streamlit as st
import pandas as pd
import altair as alt
import io

st.set_page_config(page_title="📊 Visas & Règlements", layout="wide")
st.title("📊 Tableau de bord — Visas & Règlements")

# ========= Helpers =========
def _find_col(df: pd.DataFrame, candidates):
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        for low, orig in cols.items():
            if cand in low:
                return orig
    return None

def _coerce_money(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series([], dtype="float64")
    s = (
        s.astype(str)
         .replace({',':'.', '€':'', 'EUR':'', 'euros':'', 'Euros':'', ' ':''}, regex=True)
    )
    return pd.to_numeric(s, errors="coerce")

def normalize_any_excel(xls_file) -> pd.DataFrame:
    """
    Accepte : chemin, buffer BytesIO (Streamlit uploader), ou file-like.
    1) Si l’onglet 'Données normalisées' existe => le lit tel quel.
    2) Sinon : prend la feuille la plus grande, détecte les colonnes et normalise.
    """
    xfile = pd.ExcelFile(xls_file)
    if "Données normalisées" in xfile.sheet_names:
        df = pd.read_excel(xls_file, sheet_name="Données normalisées")
    else:
        # Choisir la plus grosse feuille
        sheets = {sh: pd.read_excel(xls_file, sheet_name=sh) for sh in xfile.sheet_names}
        main_sheet = max(sheets, key=lambda k: len(sheets[k]))
        df = sheets[main_sheet].copy()

        # Nettoyage
        df = df.dropna(how="all")
        df.columns = [str(c).strip() for c in df.columns]

        # Colonnes clés
        date_col = "Date" if "Date" in df.columns else _find_col(df, ["date", "délivr", "delivr", "émission", "emission"])
        visa_col = "Visa" if "Visa" in df.columns else _find_col(df, ["visa", "type de visa", "categorie", "catégorie"])
        statut_col = "Statut" if "Statut" in df.columns else _find_col(df, ["règl", "regl", "paiement", "payment", "statut", "status"])
        amount_col = _find_col(df, ["montant", "total", "prix", "tarif"])
        paid_col   = _find_col(df, ["payé", "paye", "versé", "acompte", "reçu", "encaisse"])
        due_col    = _find_col(df, ["reste", "solde", "à payer", "a payer", "du"])

        # Date
        if date_col:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        else:
            # si pas de vraie date, crée une colonne vide pour éviter les crashs
            date_col = "Date"
            df[date_col] = pd.NaT

        # Type visa
        if not visa_col:
            visa_col = "Visa"
            df[visa_col] = "Inconnu"

        # Statut règlement
        if not statut_col:
            statut_col = "Statut"
            df[statut_col] = "Inconnu"

        # Montants
        df["Montant"] = _coerce_money(df[amount_col]) if amount_col else pd.NA
        df["Payé"]    = _coerce_money(df[paid_col])   if paid_col   else pd.NA
        if due_col:
            df["Reste"] = _coerce_money(df[due_col])
        else:
            df["Reste"] = df["Montant"] - df["Payé"] if ("Montant" in df and "Payé" in df) else pd.NA

        # Année / Mois
        df["Année"] = pd.to_datetime(df[date_col], errors="coerce").dt.year
        df["Mois"]  = pd.to_datetime(df[date_col], errors="coerce").dt.to_period("M").astype(str)

        # Harmonisation des noms finaux
        if visa_col != "Visa":
            df = df.rename(columns={visa_col: "Visa"})
        if statut_col != "Statut":
            df = df.rename(columns={
