
# -*- coding: utf-8 -*-
"""
app.py — Visa Manager (version finale intégrée)
- Onglets complets : 📄 Fichiers | 📊 Dashboard | 📈 Analyses | ➕ Ajouter | ✏️ Gestion | 💳 Compta Client | 💾 Export | 🛡️ Escrow
- Escrow intégré (badge rouge si "À réclamer")
- Un seul classeur Excel : "Clients BL.xlsx"
  • Feuille principale : "Clients" OU "Dossiers" (détection automatique)
  • Feuille "Escrow" pour le suivi des acomptes à réclamer
- Sauvegarde automatique dans "Clients BL.xlsx"
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import date, datetime
from pathlib import Path

# ====================
#   CONFIG & HEADER
# ====================
APP_TITLE = "Visa Manager"
EXCEL_FILE = "Clients BL.xlsx"
MAIN_SHEETS_CANDIDATES = ["Clients", "Dossiers"]
ESCROW_SHEET = "Escrow"

st.set_page_config(page_title=APP_TITLE, page_icon="🛂", layout="wide")
st.title("🛂 " + APP_TITLE)

# ========= Utils Excel =========
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
    # si Escrow n'existe pas, la créer
    if ESCROW_SHEET not in xls.sheet_names:
        df_esc = pd.DataFrame(columns=["Dossier N","Nom","Montant","Date envoi","État","Date réclamation"])
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as w:
            df_esc.to_excel(w, index=False, sheet_name=ESCROW_SHEET)
        xls = pd.ExcelFile(path)  # recharger

    df_main = pd.read_excel(xls, main_sheet)
    df_escrow = pd.read_excel(xls, ESCROW_SHEET)
    return main_sheet, df_main, df_escrow

def _save_excel(path: Path, df_main: pd.DataFrame, main_sheet: str, df_escrow: pd.DataFrame):
    """Écrit les deux feuilles dans le même fichier."""
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df_main.to_excel(w, index=False, sheet_name=main_sheet)
        df_escrow.to_excel(w, index=False, sheet_name=ESCROW_SHEET)

# ========= Helpers données =========
def _to_num(x):
    try:
        if pd.isna(x): return 0.0
        s = str(x).strip().replace('\u202f','').replace('\xa0','')
        s = s.replace("€","").replace(" ", "").replace(",", ".")
        return float(s)
    except Exception:
        try: return float(x)
        except Exception: return 0.0

def recalc(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return df
    df = df.copy()
    # Dates
    for col in df.columns:
        if "Date" in col and df[col].dtype == object:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    # Montants
    if "Montant total" in df.columns:
        df["Montant total"] = df["Montant total"].apply(_to_num)
    if "Acompte 1" in df.columns:
        df["Acompte 1"] = df["Acompte 1"].apply(_to_num)
    if "Montant total" in df.columns and "Acompte 1" in df.columns:
        df["Solde"] = (df["Montant total"] - df["Acompte 1"]).fillna(0)
    return df

def download_excel(df: pd.DataFrame, filename: str, label: str):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    st.download_button(label, buf.getvalue(), file_name=filename,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ========= Escrow logique =========
def escrow_sync_from_main(df_main: pd.DataFrame, df_escrow: pd.DataFrame):
    """Ajoute dans df_escrow les dossiers à réclamer manquants (Escrow=1 et Envoyé=1)."""
    # Colonnes attendues
    needed_cols = ["Dossier N","Nom","Acompte 1","Dossier envoyé","Date envoi","Escrow"]
    for col in needed_cols:
        if col not in df_main.columns:
            df_main[col] = pd.NA

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

def escrow_a_reclamer(df_escrow: pd.DataFrame):
    if df_escrow is None or df_escrow.empty: return df_escrow
    if "État" not in df_escrow.columns:
        return df_escrow.iloc[0:0]
    return df_escrow[df_escrow["État"].fillna("") == "À réclamer"]

def escrow_reclames(df_escrow: pd.DataFrame):
    if df_escrow is None or df_escrow.empty: return df_escrow
    if "État" not in df_escrow.columns:
        return df_escrow.iloc[0:0]
    return df_escrow[df_escrow["État"].fillna("") == "Réclamé"]

def escrow_mark_reclaimed(df_escrow: pd.DataFrame, dossier_num: str):
    if df_escrow is None or df_escrow.empty: return df_escrow, False
    idx = df_escrow.index[df_escrow["Dossier N"].astype(str) == str(dossier_num)]
    if len(idx):
        j = idx[0]
        df_escrow.at[j, "État"] = "Réclamé"
        df_escrow.at[j, "Date réclamation"] = datetime.now().strftime("%Y-%m-%d")
        return df_escrow, True
    return df_escrow, False

# ========= Chargement initial =========
excel_path = Path(EXCEL_FILE)
main_sheet, df_main, df_escrow = _read_excel(excel_path)

# Sync automatique (au chargement)
df_escrow, _added_auto = escrow_sync_from_main(df_main, df_escrow)
if _added_auto > 0:
    _save_excel(excel_path, df_main, main_sheet, df_escrow)

# Mémorisation session
if "df_main" not in st.session_state or "df_escrow" not in st.session_state or "main_sheet" not in st.session_state:
    st.session_state.main_sheet = main_sheet
    st.session_state.df_main = df_main
    st.session_state.df_escrow = df_escrow
else:
    main_sheet = st.session_state.main_sheet
    df_main = st.session_state.df_main
    df_escrow = st.session_state.df_escrow

# Escrow badge
try:
    escrow_pending = escrow_a_reclamer(st.session_state.df_escrow)
except Exception:
    escrow_pending = pd.DataFrame(columns=["Dossier N","Nom","Montant","Date envoi","État","Date réclamation"])
n_escrow = len(escrow_pending) if escrow_pending is not None else 0
escrow_label = "🛡️ Escrow 🔴" if n_escrow > 0 else "🛡️ Escrow"

# ========= Tabs =========
tabs = st.tabs([
    "📄 Fichiers",
    "📊 Dashboard",
    "📈 Analyses",
    "➕ Ajouter",
    "✏️ Gestion",
    "💳 Compta Client",
    "💾 Export",
    escrow_label
])

# ====================
#   📄 FICHIERS
# ====================
with tabs[0]:
    st.header("📄 Fichiers")
    st.caption(f"Classeur Excel : **{EXCEL_FILE}** — Feuille principale détectée : **{main_sheet}**")
    c1, c2 = st.columns([2,1])
    with c1:
        st.subheader(f"Aperçu ({main_sheet})")
        st.dataframe(recalc(st.session_state.df_main), use_container_width=True, height=360)
    with c2:
        st.subheader("Aperçu (Escrow)")
        st.dataframe(st.session_state.df_escrow, use_container_width=True, height=360)

    st.markdown("---")
    st.subheader("Remplacer les données (import d'un fichier Excel)")  # remplace totalement
    up = st.file_uploader("Sélectionnez un fichier .xlsx (doit contenir 'Clients' ou 'Dossiers' et 'Escrow')", type=["xlsx"])
    if up is not None:
        try:
            xls = pd.ExcelFile(up)
            # Cherche Clients ou Dossiers
            target_main = None
            for name in MAIN_SHEETS_CANDIDATES:
                if name in xls.sheet_names: target_main = name; break
            if not target_main:
                st.error(f"Feuille 'Clients' ou 'Dossiers' introuvable. Feuilles: {xls.sheet_names}")
            elif ESCROW_SHEET not in xls.sheet_names:
                st.error("Feuille 'Escrow' introuvable.")
            else:
                new_main = pd.read_excel(xls, target_main)
                new_esc = pd.read_excel(xls, ESCROW_SHEET)
                # sync auto (ajoute 'À réclamer' manquants)
                new_esc, added = escrow_sync_from_main(new_main, new_esc)
                _save_excel(excel_path, new_main, target_main, new_esc)
                st.session_state.main_sheet = target_main
                st.session_state.df_main = new_main
                st.session_state.df_escrow = new_esc
                st.success(f"✅ Données remplacées et sauvegardées ({added} ajout(s) dans Escrow).")
        except Exception as e:
            st.error(f"Erreur de lecture : {e}")

    if st.button("💾 Enregistrer (sauvegarde Excel)"):
        _save_excel(excel_path, st.session_state.df_main, st.session_state.main_sheet, st.session_state.df_escrow)
        st.success(f"Fichier sauvegardé : {excel_path}")

# ====================
#   📊 DASHBOARD
# ====================
with tabs[1]:
    st.header("📊 Tableau de bord")
    df = recalc(st.session_state.df_main)
    if df is None or df.empty:
        st.info("Aucune donnée. Ajoutez un dossier ou chargez votre fichier Excel.")
    else:
        total_dossiers = len(df)
        total_montant = df["Montant total"].sum() if "Montant total" in df.columns else 0
        total_acompte = df["Acompte 1"].sum() if "Acompte 1" in df.columns else 0
        total_solde = df["Solde"].sum() if "Solde" in df.columns else (total_montant - total_acompte)

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Nombre de dossiers", int(total_dossiers))
        c2.metric("Montant total", f"{total_montant:,.2f} €".replace(",", " ").replace(".", ","))
        c3.metric("Total encaissé (Acompte 1)", f"{total_acompte:,.2f} €".replace(",", " ").replace(".", ","))
        c4.metric("Solde restant", f"{total_solde:,.2f} €".replace(",", " ").replace(".", ","))

        if "Date" in df.columns and not df["Date"].isna().all():
            dft = df.dropna(subset=["Date"]).copy()
            dft["Mois"] = dft["Date"].dt.to_period("M").astype(str)
            monthly = dft.groupby("Mois")[["Montant total"]].sum().reset_index() if "Montant total" in dft.columns else None
            if monthly is not None and not monthly.empty:
                fig = px.bar(monthly, x="Mois", y="Montant total", title="Montant total par mois")
                st.plotly_chart(fig, use_container_width=True)

        st.subheader("Aperçu récents")
        if "Date" in df.columns:
            st.dataframe(df.sort_values("Date", ascending=False).head(20), use_container_width=True, height=380)
        else:
            st.dataframe(df.head(20), use_container_width=True, height=380)

# ====================
#   📈 ANALYSES
# ====================
with tabs[2]:
    st.header("📈 Analyses")
    df = recalc(st.session_state.df_main)
    if df is None or df.empty:
        st.info("Aucune donnée.")
    else:
        if "Solde" in df.columns:
            st.subheader("Distribution des soldes")
            st.plotly_chart(px.histogram(df, x="Solde", nbins=20), use_container_width=True)
        st.subheader("Vue complète (lecture seule)")
        st.dataframe(df, use_container_width=True, height=420)

# ====================
#   ➕ AJOUTER
# ====================
with tabs[3]:
    st.header("➕ Ajouter un dossier")
    with st.form("form_add"):
        col1,col2,col3 = st.columns(3)
        dossier_num = col1.text_input("Dossier N")
        nom_client = col2.text_input("Nom")
        date_dossier = col3.date_input("Date", date.today())

        col4,col5,col6 = st.columns(3)
        montant_total = col4.text_input("Montant total (€)", value="0")
        acompte1 = col5.text_input("Acompte 1 (€)", value="0")
        date_acompte1 = col6.date_input("Date Acompte 1", date.today())

        col7,col8 = st.columns(2)
        dossier_envoye = col7.checkbox("Dossier envoyé ?")
        date_envoi = col8.date_input("Date envoi", date.today())

        escrow_flag = st.checkbox("Escrow ?")
        ok = st.form_submit_button("Ajouter")

    if ok:
        new_row = {
            "Dossier N": dossier_num,
            "Nom": nom_client,
            "Date": pd.to_datetime(date_dossier),
            "Montant total": montant_total,
            "Acompte 1": acompte1,
            "Date Acompte 1": pd.to_datetime(date_acompte1),
            "Dossier envoyé": 1 if dossier_envoye else 0,
            "Date envoi": pd.to_datetime(date_envoi) if dossier_envoye else "",
            "Escrow": 1 if escrow_flag else 0
        }
        # Ajoute au df principal
        st.session_state.df_main = pd.concat([st.session_state.df_main, pd.DataFrame([new_row])], ignore_index=True)
        # Sync immédiate Escrow si besoin
        st.session_state.df_escrow, added = escrow_sync_from_main(st.session_state.df_main, st.session_state.df_escrow)
        _save_excel(excel_path, st.session_state.df_main, st.session_state.main_sheet, st.session_state.df_escrow)
        st.success("✅ Dossier ajouté et sauvegardé.")
        if added > 0:
            st.info(f"Escrow : {added} ligne(s) ajoutée(s).")

# ====================
#   ✏️ GESTION
# ====================
with tabs[4]:
    st.header("✏️ Gestion des dossiers")
    df = recalc(st.session_state.df_main)
    st.dataframe(df, use_container_width=True, height=360)

    st.markdown("---")
    st.subheader("Modifier un dossier (envoi / escrow)")
    colA, colB = st.columns([1,2])
    dossier_to_edit = colA.text_input("Dossier N à modifier")
    new_envoye = colB.checkbox("Dossier envoyé ?")
    date_envoi_new = colB.date_input("Date d'envoi", date.today())
    new_escrow_flag = colB.checkbox("Mettre Escrow ?")

    if st.button("Enregistrer la modification"):
        idx = st.session_state.df_main.index[st.session_state.df_main["Dossier N"].astype(str) == str(dossier_to_edit)]
        if len(idx) == 0:
            st.warning("Dossier non trouvé.")
        else:
            i = idx[0]
            st.session_state.df_main.at[i, "Dossier envoyé"] = 1 if new_envoye else 0
            st.session_state.df_main.at[i, "Date envoi"] = pd.to_datetime(date_envoi_new) if new_envoye else ""
            st.session_state.df_main.at[i, "Escrow"] = 1 if new_escrow_flag else 0
            # Sync Escrow
            st.session_state.df_escrow, added = escrow_sync_from_main(st.session_state.df_main, st.session_state.df_escrow)
            _save_excel(excel_path, st.session_state.df_main, st.session_state.main_sheet, st.session_state.df_escrow)
            st.success("✅ Dossier modifié et sauvegardé.")
            if added > 0:
                st.info(f"Escrow : {added} ligne(s) ajoutée(s).")

# ====================
#   💳 COMPTA CLIENT
# ====================
with tabs[5]:
    st.header("💳 Compta Client")
    df = recalc(st.session_state.df_main)
    if df is None or df.empty:
        st.info("Aucune donnée.")
    else:
        st.subheader("Totaux par client (Montant / Acompte / Solde)")
        cols = [c for c in ["Montant total","Acompte 1","Solde"] if c in df.columns]
        if not cols:
            st.info("Colonnes montants absentes.")
        else:
            grp = df.groupby("Nom")[cols].sum().reset_index()
            st.dataframe(grp.sort_values(cols[0], ascending=False), use_container_width=True, height=420)

# ====================
#   💾 EXPORT
# ====================
with tabs[6]:
    st.header("💾 Export")
    st.write("Téléchargez vos données au format Excel :")
    c1,c2,c3 = st.columns(3)
    with c1:
        download_excel(st.session_state.df_main, "dossiers_export.xlsx", "📥 Exporter Dossiers/Clients")
    with c2:
        download_excel(st.session_state.df_escrow, "escrow_export.xlsx", "📥 Exporter Escrow")
    with c3:
        if st.button("💾 Sauvegarder le classeur (overwrite)"):
            _save_excel(excel_path, st.session_state.df_main, st.session_state.main_sheet, st.session_state.df_escrow)
            st.success(f"Classeur sauvegardé : {excel_path}")

# ====================
#   🛡️ ESCROW
# ====================
with tabs[7]:
    st.header("🛡️ Escrow")
    # Point rouge uniquement si à réclamer
    if n_escrow > 0:
        st.error(f"⚠️ {n_escrow} dossier(s) Escrow à réclamer !")
    else:
        st.success("✅ Aucun Escrow à réclamer.")

    col1, col2 = st.columns(2)
    col1.subheader("À réclamer")
    cols_show = ["Dossier N","Nom","Date envoi","Montant"]
    view = escrow_pending.copy() if escrow_pending is not None else pd.DataFrame(columns=cols_show)
    for c in cols_show:
        if c not in view.columns:
            view[c] = ""
    col1.dataframe(view[cols_show], use_container_width=True, height=300)

    col2.subheader("Réclamés")
    r = escrow_reclames(st.session_state.df_escrow)
    if r is not None and not r.empty:
        r_view = r.copy()
        for c in cols_show:
            if c not in r_view.columns:
                r_view[c] = ""
        col2.dataframe(r_view[cols_show], use_container_width=True, height=300)
    else:
        col2.info("Aucun dossier réclamé.")

    st.markdown("---")
    st.subheader("✅ Marquer un Escrow comme réclamé")
    num_rec = st.text_input("Numéro de dossier")
    if st.button("Marquer comme réclamé"):
        st.session_state.df_escrow, ok = escrow_mark_reclaimed(st.session_state.df_escrow, num_rec)
        if ok:
            _save_excel(excel_path, st.session_state.df_main, st.session_state.main_sheet, st.session_state.df_escrow)
            st.success(f"Dossier {num_rec} marqué comme réclamé et sauvegardé.")
        else:
            st.warning("Numéro de dossier introuvable dans Escrow.")
