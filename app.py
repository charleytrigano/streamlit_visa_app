import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(page_title="📊 Visas & Règlements", layout="wide")

st.title("📊 Tableau de bord — Visas & Règlements")

# ---------- Chargement des données ----------
@st.cache_data
def load_data(xlsx_path: str) -> pd.DataFrame:
    # Le fichier est celui généré : "Données normalisées" avec colonnes Date / Année / Mois
    df = pd.read_excel(xlsx_path, sheet_name="Données normalisées")
    # Conversions sûres
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for num_col in ["Montant", "Payé", "Reste"]:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce")
    # Valeurs par défaut si colonnes manquantes
    if "Année" not in df.columns and "Date" in df.columns:
        df["Année"] = df["Date"].dt.year
    if "Mois" not in df.columns and "Date" in df.columns:
        df["Mois"] = df["Date"].dt.to_period("M").astype(str)
    # Colonnes clés potentielles
    if "Visa" not in df.columns:
        # si votre colonne s'appelle autrement (ex: "Categories"), renommez-la ici :
        # df = df.rename(columns={"Categories": "Visa"})
        df["Visa"] = "Inconnu"
    if "__Statut règlement__" in df.columns and "Statut" not in df.columns:
        df = df.rename(columns={"__Statut règlement__": "Statut"})
    if "Statut" not in df.columns:
        df["Statut"] = "Inconnu"
    return df

# Chemin par défaut (même dossier que app.py)
DEFAULT_DATA_PATH = "visa_analytics_datecol.xlsx"

st.sidebar.header("Données")
mode = st.sidebar.radio("Source", ["Fichier normalisé (recommandé)", "Importer un autre fichier Excel"])
if mode == "Fichier normalisé (recommandé)":
    data_path = DEFAULT_DATA_PATH
    st.sidebar.info(f"Lecture : {data_path}")
    df = load_data(data_path)
else:
    up = st.sidebar.file_uploader("Déposez ici votre Excel (mêmes colonnes)", type=["xlsx","xls"])
    if up is not None:
        df = load_data(up)
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
    st.dataframe(
        f.sort_values(by=["Date"] if "Date" in f.columns else f.columns.tolist()),
        use_container_width=True
    )

st.caption("💡 Astuce : si le nom de votre colonne de type de visa n'est pas 'Visa', renommez-la dans la source ou adaptez la ligne de renommage dans le code.")
