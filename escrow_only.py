
# -*- coding: utf-8 -*-
"""
escrow_only.py — Module Streamlit autonome pour la gestion Escrow
- Utilise le même classeur Excel: "Clients BL.xlsx"
- Crée/maintient la feuille "Escrow" dans ce classeur
- Synchronise automatiquement : tout dossier (Clients/Dossiers) avec Escrow=1 et Dossier envoyé=1
  -> ajoute une ligne "À réclamer" dans la feuille Escrow (si absente)
- Affiche un badge rouge (dans le titre) quand il y a des escrows à réclamer
- Permet de marquer comme "Réclamé" et export Excel
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from pathlib import Path

EXCEL_FILE = "Clients BL.xlsx"   # même fichier que ton app principale
MAIN_SHEETS_CANDIDATES = ["Clients", "Dossiers"]
ESCROW_SHEET = "Escrow"

st.set_page_config(page_title="Escrow", page_icon="🛡️", layout="wide")

# ---------- Utils ----------
def to_dt(x):
    try:
        return pd.to_datetime(x, errors="coerce")
    except Exception:
        return pd.NaT

def _ensure_workbook_and_sheets(path: Path):
    """Crée le classeur minimal s'il n'existe pas, avec feuilles Clients + Escrow."""
    if not path.exists():
        df_clients = pd.DataFrame(columns=[
            "Dossier N","Nom","Date","Montant total","Acompte 1",
            "Date Acompte 1","Dossier envoyé","Date envoi","Escrow"
        ])
        df_escrow = pd.DataFrame(columns=[
            "Dossier N","Nom","Montant","Date envoi","État","Date réclamation"
        ])
        with pd.ExcelWriter(path, engine="openpyxl") as w:
            df_clients.to_excel(w, index=False, sheet_name="Clients")
            df_escrow.to_excel(w, index=False, sheet_name=ESCROW_SHEET)

def _read_excel(path: Path):
    """Lit le fichier, détecte la feuille principale (Clients/Dossiers) et Escrow."""
    _ensure_workbook_and_sheets(path)
    xls = pd.ExcelFile(path)
    # détecter feuille principale
    main_sheet = None
    for cand in MAIN_SHEETS_CANDIDATES:
        for s in xls.sheet_names:
            if s.strip().lower() == cand.lower():
                main_sheet = s
                break
        if main_sheet:
            break
    if not main_sheet:
        raise ValueError(f"Aucune feuille principale trouvée parmi {MAIN_SHEETS_CANDIDATES}. Feuilles: {xls.sheet_names}")
    # si Escrow n'existe pas, la créer à la volée
    if ESCROW_SHEET not in xls.sheet_names:
        df_esc = pd.DataFrame(columns=["Dossier N","Nom","Montant","Date envoi","État","Date réclamation"])
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as w:
            df_esc.to_excel(w, index=False, sheet_name=ESCROW_SHEET)
        xls = pd.ExcelFile(path)  # recharger

    df_main = pd.read_excel(xls, main_sheet)
    df_escrow = pd.read_excel(xls, ESCROW_SHEET)
    return main_sheet, df_main, df_escrow

def _save_excel(path: Path, df_main: pd.DataFrame, main_sheet: str, df_escrow: pd.DataFrame):
    """Écrit les deux feuilles en conservant le même fichier."""
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df_main.to_excel(w, index=False, sheet_name=main_sheet)
        df_escrow.to_excel(w, index=False, sheet_name=ESCROW_SHEET)

def _sync_from_main(df_main: pd.DataFrame, df_escrow: pd.DataFrame):
    """Ajoute dans df_escrow les dossiers à réclamer manquants (Escrow=1 et Envoyé=1)."""
    # normaliser colonnes attendues
    needed_cols = ["Dossier N","Nom","Acompte 1","Dossier envoyé","Date envoi","Escrow"]
    for col in needed_cols:
        if col not in df_main.columns:
            # si certaines colonnes n'existent pas, les créer vides
            df_main[col] = pd.NA

    # construire set des dossiers déjà en escrow
    existing = set(df_escrow["Dossier N"].astype(str)) if not df_escrow.empty and "Dossier N" in df_escrow.columns else set()

    added = 0
    rows = []
    for _, r in df_main.iterrows():
        try:
            esc = int(r.get("Escrow", 0)) if pd.notna(r.get("Escrow", 0)) else 0
            sent = int(r.get("Dossier envoyé", 0)) if pd.notna(r.get("Dossier envoyé", 0)) else 0
        except Exception:
            esc = 0; sent = 0
        if esc == 1 and sent == 1:
            num = str(r.get("Dossier N", "")).strip()
            if num and num not in existing:
                rows.append({
                    "Dossier N": num,
                    "Nom": r.get("Nom", ""),
                    "Montant": r.get("Acompte 1", 0),
                    "Date envoi": r.get("Date envoi", ""),
                    "État": "À réclamer",
                    "Date réclamation": ""
                })
                existing.add(num)
                added += 1
    if rows:
        df_escrow = pd.concat([df_escrow, pd.DataFrame(rows)], ignore_index=True)
    return df_escrow, added

def a_reclamer(df_escrow: pd.DataFrame):
    if "État" not in df_escrow.columns:
        return df_escrow.iloc[0:0]
    return df_escrow[df_escrow["État"].fillna("") == "À réclamer"]

def reclames(df_escrow: pd.DataFrame):
    if "État" not in df_escrow.columns:
        return df_escrow.iloc[0:0]
    return df_escrow[df_escrow["État"].fillna("") == "Réclamé"]

# ============ Chargement ============
excel_path = Path(EXCEL_FILE)
main_sheet, df_main, df_esc = _read_excel(excel_path)

# sync automatique
df_esc, added = _sync_from_main(df_main, df_esc)
if added > 0:
    _save_excel(excel_path, df_main, main_sheet, df_esc)

pending = a_reclamer(df_esc)
n_pending = len(pending)

title = "🛡️ Escrow 🔴" if n_pending > 0 else "🛡️ Escrow"
st.title(title)

st.caption(f"Classeur : {EXCEL_FILE} — Feuille principale détectée : **{main_sheet}**")
if added > 0:
    st.success(f"Synchronisation automatique : {added} dossier(s) ajoutés à la feuille Escrow.")

# ============ Vue principale ============
st.subheader("À réclamer")
cols_show = ["Dossier N","Nom","Date envoi","Montant"]
view = pending.copy()
for c in cols_show:
    if c not in view.columns:
        view[c] = ""
st.dataframe(view[cols_show], use_container_width=True, height=260)

st.subheader("Réclamés")
r = reclames(df_esc)
if not r.empty:
    r_view = r.copy()
    for c in cols_show:
        if c not in r_view.columns:
            r_view[c] = ""
    st.dataframe(r_view[cols_show], use_container_width=True, height=260)
else:
    st.dataframe(r, use_container_width=True, height=100)

st.markdown("---")
st.subheader("Actions")
col1, col2, col3 = st.columns([2,2,1])

with col1:
    num_to_mark = st.text_input("Numéro de dossier à marquer comme réclamé")
    if st.button("✅ Marquer comme réclamé"):
        if num_to_mark:
            idx = df_esc.index[df_esc["Dossier N"].astype(str) == str(num_to_mark)]
            if len(idx):
                j = idx[0]
                df_esc.at[j, "État"] = "Réclamé"
                df_esc.at[j, "Date réclamation"] = datetime.now().strftime("%Y-%m-%d")
                _save_excel(excel_path, df_main, main_sheet, df_esc)
                st.success(f"Dossier {num_to_mark} marqué comme réclamé et sauvegardé.")
            else:
                st.warning("Numéro de dossier introuvable dans Escrow.")

with col2:
    # Export Excel (vue Escrow complète)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_esc.to_excel(w, index=False, sheet_name="Escrow")
    st.download_button("📥 Exporter Escrow (xlsx)", data=buf.getvalue(),
                       file_name="Escrow_export.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with col3:
    if st.button("🔄 Resynchroniser"):
        # relire + resynchroniser
        _, df_main2, df_esc2 = _read_excel(excel_path)
        df_esc2, added2 = _sync_from_main(df_main2, df_esc2)
        if added2 > 0:
            _save_excel(excel_path, df_main2, main_sheet, df_esc2)
            st.success(f"Synchronisation : {added2} nouveau(x) dossier(s) ajouté(s).")
        else:
            st.info("Aucun nouveau dossier à ajouter.")
