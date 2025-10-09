import streamlit as st
import pandas as pd

st.set_page_config(page_title="Visa — Filtres dynamiques", layout="wide")

# === Chargement du fichier Excel ===
@st.cache_data
def load_visa_structure(path):
    df = pd.read_excel(path)
    df = df.fillna("")  # éviter les NaN
    return df

visa_df = load_visa_structure("Visa.xlsx")

# === Construction d'une structure hiérarchique ===
def build_hierarchy(df):
    hierarchy = {}
    for _, row in df.iterrows():
        cat = row["Catégorie"]
        if not cat:
            continue
        if cat not in hierarchy:
            hierarchy[cat] = {}

        # construire les sous-niveaux dynamiquement
        sub = hierarchy[cat]
        for i in range(1, 9):
            col = f"Sous-categories {i}"
            val = row[col].strip() if col in row and isinstance(row[col], str) else ""
            if val:
                if val not in sub:
                    sub[val] = {}
                sub = sub[val]
    return hierarchy

visa_hierarchy = build_hierarchy(visa_df)

st.sidebar.success("✅ Structure hiérarchique chargée")


# === Sélecteurs hiérarchiques ===
st.header("🎯 Filtres hiérarchiques Visa")

selected_filters = {}
level = 1
current_level = visa_hierarchy

while current_level and level <= 8:
    options = list(current_level.keys())
    if not options:
        break

    selected = st.multiselect(
        f"Niveau {level} — Sélection :",
        options,
        key=f"lvl_{level}"
    )

    selected_filters[f"niveau_{level}"] = selected

    # Si un seul choix est fait, on descend dans la hiérarchie
    if len(selected) == 1:
        current_level = current_level[selected[0]]
        level += 1
    else:
        break

# === Affichage du chemin sélectionné ===
path = " > ".join([val[0] for val in selected_filters.values() if val])
if path:
    st.info(f"🧭 Chemin sélectionné : **{path}**")
else:
    st.warning("Aucune sélection effectuée pour le moment.")


# === Simulation d'un tableau de dossiers (à remplacer par tes données réelles) ===
data = {
    "Nom du client": ["Dupont", "Smith", "Garcia", "Lee"],
    "Catégorie": ["E-2", "H-1B", "E-2", "EB-5"],
    "Sous-catégorie": ["E-2 Inv. Ren.", "Extension", "E-2 Inv.", "I-526"]
}
df_dossiers = pd.DataFrame(data)

# === Filtrage dynamique ===
mask = pd.Series(True, index=df_dossiers.index)
for level_key, values in selected_filters.items():
    if values:
        # On filtre sur la catégorie ou sous-catégorie correspondante
        col_candidates = [c for c in df_dossiers.columns if "Catégorie" in c or "Sous" in c]
        for col in col_candidates:
            mask &= df_dossiers[col].isin(values)

filtered = df_dossiers[mask]

st.subheader("📋 Résultats filtrés")
st.dataframe(filtered, use_container_width=True)

st.caption(f"{len(filtered)} dossier(s) correspondant(s) à la sélection.")