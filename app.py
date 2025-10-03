import json
from pathlib import Path
import streamlit as st
import pandas as pd

st.set_page_config(page_title="📊 Visas — Simplifié", layout="wide")
st.title("📊 Visas — Tableau simplifié")

# ---------------- Utils ----------------
def _first_col(df: pd.DataFrame, candidates) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None

def _to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _to_date(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def _parse_paiements(x):
    if isinstance(x, list):
        return x
    if pd.isna(x):
        return []
    try:
        return json.loads(x)
    except Exception:
        return []

def _sum_payments(pay_list) -> float:
    total = 0.0
    for p in (pay_list or []):
        try:
            amt = float(p.get("amount", 0) or 0) if isinstance(p, dict) else float(p)
        except Exception:
            amt = 0.0
        total += amt
    return total

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Uniformise Date/Année/Mois/Visa/Statut/Montant/Payé/Reste si c'est un tableau 'dossiers'."""
    df = df.copy()

    if "Date" in df.columns:
        df["Date"] = _to_date(df["Date"])
    else:
        df["Date"] = pd.NaT

    if "Année" not in df.columns:
        df["Année"] = df["Date"].dt.year
    if "Mois" not in df.columns:
        df["Mois"] = df["Date"].dt.to_period("M").astype(str)

    visa_col = _first_col(df, ["Visa", "Categories", "Catégorie", "TypeVisa"])
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"

    if "__Statut règlement__" in df.columns and "Statut" not in df.columns:
        df = df.rename(columns={"__Statut règlement__": "Statut"})
    if "Statut" not in df.columns:
        df["Statut"] = "Inconnu"
    else:
        df["Statut"] = df["Statut"].astype(str).fillna("Inconnu")

    if "Montant" in df.columns:
        df["Montant"] = _to_num(df["Montant"])
    else:
        src_montant = _first_col(df, ["Honoraires", "Total", "Amount"])
        df["Montant"] = _to_num(df[src_montant]) if src_montant else 0.0

    if "Paiements" in df.columns:
        parsed = df["Paiements"].apply(_parse_paiements)
        df["TotalAcomptes"] = parsed.apply(_sum_payments)

    if "Payé" in df.columns:
        df["Payé"] = _to_num(df["Payé"])
    else:
        src_paye = _first_col(df, ["TotalAcomptes", "Acomptes", "Paye", "Paid"])
        df["Payé"] = _to_num(df[src_paye]) if src_paye else 0.0

    if "Reste" in df.columns:
        df["Reste"] = _to_num(df["Reste"])
    else:
        src_reste = _first_col(df, ["Solde", "SoldeCalc"])
        if src_reste:
            df["Reste"] = _to_num(df[src_reste])
        else:
            df["Reste"] = (df["Montant"] - df["Payé"]).fillna(0.0)

    return df

def looks_like_reference(df: pd.DataFrame) -> bool:
    """
    Détecte un onglet de référence (ex: feuille 'Visa' avec colonnes
    'Categories', 'Visa', 'Definition') qui n'est pas une liste de dossiers.
    """
    cols = set(map(str.lower, df.columns.astype(str)))
    # Présence d'un trio typique de dictionnaire et absence de colonnes montants
    has_ref = {"categories", "visa"} <= cols
    no_money = not ({"montant", "honoraires", "acomptes", "payé", "reste", "solde"} & cols)
    return has_ref and no_money

@st.cache_data
def read_excel(xlsx_obj):
    xls = pd.ExcelFile(xlsx_obj)
    return xls.sheet_names, xls  # retourne aussi l'ExcelFile pour charger la feuille choisie

@st.cache_data
def load_sheet(xls: pd.ExcelFile, sheet_name: str, normalize: bool):
    df = pd.read_excel(xls, sheet_name=sheet_name)
    if normalize and not looks_like_reference(df):
        df = normalize_dataframe(df)
    return df

# ---------------- Source & Sélection feuille ----------------
DEFAULT_CANDIDATES = [
    "/mnt/data/Visa_Clients_20251001-114844.xlsx",
    "/mnt/data/visa_analytics_datecol.xlsx",
]

st.sidebar.header("Données")
source_mode = st.sidebar.radio("Source", ["Fichier par défaut", "Importer un Excel"])

if source_mode == "Fichier par défaut":
    path = next((p for p in DEFAULT_CANDIDATES if Path(p).exists()), None)
    if not path:
        st.sidebar.error("Aucun fichier par défaut trouvé. Importez un fichier.")
        st.stop()
    st.sidebar.success(f"Fichier: {path}")
    sheet_names, xls = read_excel(path)
else:
    up = st.sidebar.file_uploader("Dépose un Excel (.xlsx, .xls)", type=["xlsx", "xls"])
    if not up:
        st.info("Importe un fichier pour commencer.")
        st.stop()
    sheet_names, xls = read_excel(up)

# Choix explicite de la feuille (oui, y compris 'Visa')
preferred_order = ["Données normalisées", "Clients", "Visa"]
default_sheet = next((s for s in preferred_order if s in sheet_names), sheet_names[0])
sheet_choice = st.sidebar.selectbox("Feuille", sheet_names, index=sheet_names.index(default_sheet))

# Charger. Si la feuille ressemble à un dictionnaire de références, on **normalize=False**
df_raw = load_sheet(xls, sheet_choice, normalize=not looks_like_reference(pd.read_excel(xls, sheet_name=sheet_choice, nrows=5)))
is_ref = looks_like_reference(df_raw)

if is_ref:
    st.info("ℹ️ Cette feuille ressemble à une **table de référence** (ex: correspondance Catégories ↔ Visa). "
            "Elle s'affiche telle quelle. Pour analyser des dossiers, sélectionnez l'onglet **Clients** ou "
            "**Données normalisées** dans la sidebar.")
    st.dataframe(df_raw, use_container_width=True)
    st.stop()

df = df_raw  # déjà normalisé

# ---------------- Filtres ----------------
with st.container():
    c1, c2, c3 = st.columns(3)
    years = sorted([int(y) for y in df["Année"].dropna().unique()]) if "Année" in df else []
    visas = sorted(df["Visa"].dropna().astype(str).unique())
    statuses = sorted(df["Statut"].dropna().astype(str).unique())

    year_sel = c1.multiselect("Années", years, default=years or None)
    visa_sel = c2.multiselect("Type de visa", visas, default=visas or None)
    stat_sel = c3.multiselect("Statut", statuses, default=statuses or None)

f = df.copy()
if year_sel:
    f = f[f["Année"].isin(year_sel)]
if visa_sel:
    f = f[f["Visa"].astype(str).isin(visa_sel)]
if stat_sel:
    f = f[f["Statut"].astype(str).isin(stat_sel)]

# ---------------- KPIs ----------------
k1, k2, k3, k4 = st.columns(4)
k1.metric("Dossiers", f"{len(f)}")
k2.metric("Montant total", f"{f['Montant'].sum():,.2f} €")
k3.metric("Payé", f"{f['Payé'].sum():,.2f} €")
k4.metric("Reste", f"{f['Reste'].sum():,.2f} €")

st.divider()

# ---------------- Graphique ----------------
st.subheader("📈 Nombre de dossiers par mois")
if "Mois" in f.columns:
    counts = (
        f.dropna(subset=["Mois"])
         .groupby("Mois")
         .size()
         .rename("Dossiers")
         .reset_index()
         .sort_values("Mois")
    )
    st.bar_chart(counts.set_index("Mois"))
else:
    st.info("Aucune colonne 'Mois' exploitable.")

# ---------------- Tableau ----------------
st.subheader("📋 Données")
cols_show = [c for c in ["Date","Année","Mois","Visa","Statut","Montant","Payé","Reste"] if c in f.columns]
st.dataframe(
    f[cols_show].sort_values(by=[c for c in ["Date","Visa","Statut"] if c in f.columns], na_position="last"),
    use_container_width=True
)

st.caption("Astuce : choisissez l’onglet dans la sidebar. Si vous sélectionnez 'Visa' (table de référence), le programme l’affiche tel quel.")

