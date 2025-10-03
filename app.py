import json
from pathlib import Path
import streamlit as st
import pandas as pd

st.set_page_config(page_title="📊 Visas — Simplifié", layout="wide")
st.title("📊 Visas — Tableau simplifié")

# ---------- Utilitaires internes (sans module externe) ----------

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
    """Accepte liste ou chaîne JSON; renvoie liste de dicts ou nombres."""
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
    """Uniformise les colonnes clés: Date, Année, Mois, Visa, Statut, Montant, Payé, Reste."""
    df = df.copy()

    # --- Date / Année / Mois
    if "Date" in df.columns:
        df["Date"] = _to_date(df["Date"])
    else:
        df["Date"] = pd.NaT

    if "Année" not in df.columns:
        df["Année"] = df["Date"].dt.year

    if "Mois" not in df.columns:
        # format AAAA-MM pour un groupby simple
        df["Mois"] = df["Date"].dt.to_period("M").astype(str)

    # --- Type de visa
    visa_col = _first_col(df, ["Visa", "Categories", "Catégorie", "TypeVisa"])
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"

    # --- Statut (optionnel)
    if "__Statut règlement__" in df.columns and "Statut" not in df.columns:
        df = df.rename(columns={"__Statut règlement__": "Statut"})
    if "Statut" not in df.columns:
        df["Statut"] = "Inconnu"
    else:
        df["Statut"] = df["Statut"].astype(str).fillna("Inconnu")

    # --- Montant / Payé / Reste
    # Montant
    if "Montant" in df.columns:
        df["Montant"] = _to_num(df["Montant"])
    else:
        src_montant = _first_col(df, ["Honoraires", "Total", "Amount"])
        df["Montant"] = _to_num(df[src_montant]) if src_montant else 0.0

    # Paiements (JSON ou liste) -> TotalAcomptes
    if "Paiements" in df.columns:
        parsed = df["Paiements"].apply(_parse_paiements)
        df["TotalAcomptes"] = parsed.apply(_sum_payments)

    # Payé
    if "Payé" in df.columns:
        df["Payé"] = _to_num(df["Payé"])
    else:
        src_paye = _first_col(df, ["TotalAcomptes", "Acomptes", "Paye", "Paid"])
        df["Payé"] = _to_num(df[src_paye]) if src_paye else 0.0

    # Reste
    if "Reste" in df.columns:
        df["Reste"] = _to_num(df["Reste"])
    else:
        src_reste = _first_col(df, ["Solde", "SoldeCalc"])
        if src_reste:
            df["Reste"] = _to_num(df[src_reste])
        else:
            df["Reste"] = (df["Montant"] - df["Payé"]).fillna(0.0)

    return df

@st.cache_data
def load_excel_smart(xlsx_obj) -> pd.DataFrame:
    """Choisit automatiquement la feuille: 'Données normalisées' > 'Clients' > 1ère."""
    xls = pd.ExcelFile(xlsx_obj)
    names = xls.sheet_names
    target = "Données normalisées" if "Données normalisées" in names else (
        "Clients" if "Clients" in names else names[0]
    )
    base = pd.read_excel(xls, sheet_name=target)
    return normalize_dataframe(base)

# ---------- Source de données (très simple) ----------

DEFAULT_CANDIDATES = [
    "/mnt/data/Visa_Clients_20251001-114844.xlsx",  # fourni
    "/mnt/data/visa_analytics_datecol.xlsx",        # fourni
]

st.sidebar.header("Données")
mode = st.sidebar.radio("Source", ["Fichier par défaut", "Importer un Excel"])

if mode == "Fichier par défaut":
    path = next((p for p in DEFAULT_CANDIDATES if Path(p).exists()), None)
    if not path:
        st.sidebar.error("Aucun fichier par défaut trouvé. Merci d'en importer un.")
        st.stop()
    st.sidebar.success(f"Fichier: {path}")
    df = load_excel_smart(path)
else:
    up = st.sidebar.file_uploader("Dépose un fichier Excel (.xlsx, .xls)", type=["xlsx", "xls"])
    if not up:
        st.info("Importe un fichier pour commencer.")
        st.stop()
    df = load_excel_smart(up)

# ---------- Filtres légers ----------
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

# ---------- KPIs minimales ----------
k1, k2, k3, k4 = st.columns(4)
k1.metric("Dossiers", f"{len(f)}")
k2.metric("Montant total", f"{f['Montant'].sum():,.2f} €")
k3.metric("Payé", f"{f['Payé'].sum():,.2f} €")
k4.metric("Reste", f"{f['Reste'].sum():,.2f} €")

st.divider()

# ---------- Graphe simple : nombre par mois ----------
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

# ---------- Tableau détaillé (toggle) ----------
st.subheader("📋 Données")
st.dataframe(
    f[["Date","Année","Mois","Visa","Statut","Montant","Payé","Reste"]]
      .sort_values(by=["Date","Visa","Statut"], na_position="last"),
    use_container_width=True
)

st.caption("Astuce : le programme détecte automatiquement les colonnes (Honoraires/Acomptes/Solde ou Montant/Payé/Reste) et lit aussi les Paiements JSON si présents.")
