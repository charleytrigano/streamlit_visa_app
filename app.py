import streamlit as st
import pandas as pd
import altair as alt

from utils import compute_finances  # ← on exploite l’utilitaire

st.set_page_config(page_title="📊 Visas & Règlements", layout="wide")
st.title("📊 Tableau de bord — Visas & Règlements")

# ---------- Helpers ----------
def _coerce_datetime(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def _coerce_numeric(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")

def _first_existing(df: pd.DataFrame, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def _ensure_money_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aligne les colonnes sur Montant/Payé/Reste en s'appuyant sur :
    - Honoraires / Acomptes / Solde (ton Excel)
    - Paiements (JSON) via compute_finances si présent
    """
    df = df.copy()

    # Si Paiements existe, calcule TotalAcomptes/ SoldeCalc
    if "Paiements" in df.columns or "Honoraires" in df.columns:
        df = compute_finances(df)  # crée TotalAcomptes & SoldeCalc à partir d'Honoraires/Paiements

    # Montant
    if "Montant" not in df.columns:
        src_montant = _first_existing(df, ["Honoraires", "Total", "Amount"])
        df["Montant"] = _coerce_numeric(df[src_montant]) if src_montant else 0.0
    else:
        df["Montant"] = _coerce_numeric(df["Montant"]).fillna(0.0)

    # Payé
    if "Payé" not in df.columns:
        # priorité aux champs calculés, puis Acomptes
        src_paye = _first_existing(df, ["TotalAcomptes", "Acomptes", "Paye", "Paid"])
        df["Payé"] = _coerce_numeric(df[src_paye]) if src_paye else 0.0
    else:
        df["Payé"] = _coerce_numeric(df["Payé"]).fillna(0.0)

    # Reste
    if "Reste" not in df.columns:
        # priorité Solde (si déjà fourni), sinon Montant - Payé, sinon SoldeCalc
        src_reste = _first_existing(df, ["Solde", "SoldeCalc", "Reste"])
        if src_reste:
            df["Reste"] = _coerce_numeric(df[src_reste])
        else:
            df["Reste"] = (df["Montant"] - df["Payé"])
    else:
        df["Reste"] = _coerce_numeric(df["Reste"])

    df["Montant"] = df["Montant"].fillna(0.0)
    df["Payé"] = df["Payé"].fillna(0.0)
    df["Reste"] = df["Reste"].fillna(df["Montant"] - df["Payé"])
    return df

def _ensure_date_fields(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # Date
    if "Date" in df.columns:
        df["Date"] = _coerce_datetime(df["Date"])
    else:
        # si pas de Date, on laisse vide; le dashboard fonctionnera quand même
        df["Date"] = pd.NaT

    # Année / Mois
    if "Année" not in df.columns:
        df["Année"] = df["Date"].dt.year

    if "Mois" not in df.columns:
        # chaîne AAAA-MM
        df["Mois"] = df["Date"].dt.to_period("M").astype(str)

    return df

def _ensure_visa_and_status(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    # Type de visa
    visa_col = _first_existing(df, ["Visa", "Categories", "Catégorie", "TypeVisa"])
    if visa_col is None:
        df["Visa"] = "Inconnu"
    else:
        df["Visa"] = df[visa_col].astype(str).fillna("Inconnu")

    # Statut de règlement
    if "__Statut règlement__" in df.columns and "Statut" not in df.columns:
        df = df.rename(columns={"__Statut règlement__": "Statut"})
    if "Statut" not in df.columns:
        df["Statut"] = "Inconnu"
    else:
        df["Statut"] = df["Statut"].astype(str).fillna("Inconnu")

    return df

@st.cache_data
def load_any_sheet(xlsx_obj) -> pd.DataFrame:
    """
    Charge intelligemment le fichier :
    - si 'Données normalisées' existe, on la prend
    - sinon on privilégie 'Clients', puis la 1ère feuille
    - normalisation colonnes pour le dashboard
    """
    xls = pd.ExcelFile(xlsx_obj)
    sheet_names = xls.sheet_names
    target = None
    for candidate in ["Données normalisées", "Clients"]:
        if candidate in sheet_names:
            target = candidate
            break
    if target is None:
        target = sheet_names[0]

    df = pd.read_excel(xls, sheet_name=target)

    df = _ensure_money_columns(df)
    df = _ensure_date_fields(df)
    df = _ensure_visa_and_status(df)

    return df

# ---------- Données ----------
DEFAULT_DATA_PATH = "/mnt/data/Visa_Clients_20251001-114844.xlsx"  # ton fichier actuel par défaut

st.sidebar.header("Données")
mode = st.sidebar.radio("Source", ["Fichier par défaut", "Importer un autre fichier Excel"])
if mode == "Fichier par défaut":
    st.sidebar.info(f"Lecture : {DEFAULT_DATA_PATH}")
    df = load_any_sheet(DEFAULT_DATA_PATH)
else:
    up = st.sidebar.file_uploader("Dépose ton Excel (Clients / Visa ou Données normalisées)", type=["xlsx", "xls"])
    if up is not None:
        df = load_any_sheet(up)
    else:
        st.stop()

# ---------- Filtres ----------
with st.container():
    c1, c2, c3, c4 = st.columns([1,1,1,1])
    years = sorted([int(y) for y in df["Année"].dropna().unique()]) if "Année" in df else []
    types = sorted(df["Visa"].dropna().astype(str).unique())
    statuses = sorted(df["Statut"].dropna().astype(str).unique())

    year_sel = c1.multiselect("Années", years, default=years or None)
    type_sel = c2.multiselect("Types de visa", types, default=types or None)
    status_sel = c3.multiselect("Statuts de règlement", statuses, default=statuses or None)
    show_table = c4.toggle("Afficher le tableau détaillé", value=False)

f = df.copy()
if year_sel:
    f = f[f["Année"].isin(year_sel)]
if type_sel:
    f = f[f["Visa"].astype(str).isin(type_sel)]
if status_sel:
    f = f[f["Statut"].astype(str).isin(status_sel)]

# ---------- KPIs ----------
k1, k2, k3, k4 = st.columns(4)
total_visas = len(f.dropna(subset=["Date"])) if "Date" in f.columns else len(f)
k1.metric("Visas (sélection)", f"{total_visas}")
k2.metric("Montant total", f"{f['Montant'].sum(skipna=True):,.2f} €" if "Montant" in f else "—")
k3.metric("Payé", f"{f['Payé'].sum(skipna=True):,.2f} €" if "Payé" in f else "—")
k4.metric("Reste", f"{f['Reste'].sum(skipna=True):,.2f} €" if "Reste" in f else "—")

st.divider()

# ---------- Graphiques ----------
colA, colB = st.columns(2)

if "Mois" in f.columns and "Visa" in f.columns:
    monthly_counts = (
        f.dropna(subset=["Mois"])
         .groupby(["Mois","Visa"]).size().reset_index(name="Nombre de visas")
    )
    with colA:
        st.subheader("Visas par mois (par type)")
        if not monthly_counts.empty:
            chart_month = alt.Chart(monthly_counts).mark_bar().encode(
                x=alt.X("Mois:O", sort="ascending"),
                y="Nombre de visas:Q",
                color="Visa:N",
                tooltip=["Mois","Visa","Nombre de visas"]
            ).properties(height=360)
            st.altair_chart(chart_month, use_container_width=True)
        else:
            st.info("Pas de données mensuelles à afficher.")

if "Année" in f.columns and "Visa" in f.columns:
    yearly_counts = (
        f.dropna(subset=["Année"])
         .groupby(["Année","Visa"]).size().reset_index(name="Nombre de visas")
    )
    with colB:
        st.subheader("Visas par année (par type)")
        if not yearly_counts.empty:
            chart_year = alt.Chart(yearly_counts).mark_bar().encode(
                x=alt.X("Année:O", sort="ascending"),
                y="Nombre de visas:Q",
                color="Visa:N",
                tooltip=["Année","Visa","Nombre de visas"]
            ).properties(height=360)
            st.altair_chart(chart_year, use_container_width=True)
        else:
            st.info("Pas de données annuelles à afficher.")

st.subheader("Suivi des règlements")
if set(["Montant","Payé","Reste","Statut"]).issubset(f.columns):
    pay = (
        f.groupby(["Statut"])[["Montant","Payé","Reste"]]
         .sum(min_count=1).reset_index()
         .sort_values("Reste", ascending=False)
    )
    st.dataframe(pay, use_container_width=True)
    if not pay.empty:
        chart_reste = alt.Chart(pay).mark_bar().encode(
            x=alt.X("Statut:N", sort="-y"),
            y="Reste:Q",
            tooltip=["Statut","Montant","Payé","Reste"]
        ).properties(height=300)
        st.altair_chart(chart_reste, use_container_width=True)
else:
    st.info("Colonnes de montants manquantes pour le suivi des règlements.")

# ---------- Tableau détaillé (optionnel) ----------
if show_table:
    st.subheader("Données (après filtres)")
    order_cols = ["Date","Année","Mois"] if "Date" in f.columns else f.columns.tolist()
    st.dataframe(
        f.sort_values(by=order_cols),
        use_container_width=True
    )

st.caption("💡 Astuce : le type de visa est détecté depuis 'Visa' ou 'Categories'. Les montants sont harmonisés à partir d'Honoraires/Acomptes/Solde ou directement via Paiements.")
