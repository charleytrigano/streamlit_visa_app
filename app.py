# =========================
# VISA MANAGER — APP COMPLETE
# =========================

from __future__ import annotations

import streamlit as st
import pandas as pd
import json
from datetime import date, datetime
from pathlib import Path
from io import BytesIO
import unicodedata
from uuid import uuid4
import zipfile
import altair as alt

# ---- Page
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

# ---- Espace de noms unique pour éviter collisions de widgets
SID = st.session_state.setdefault("WIDGET_NS", str(uuid4()))

# =========================
# Constantes
# =========================
CLIENTS_FILE_DEFAULT = "donnees_visa_clients1_adapte.xlsx"
VISA_FILE_DEFAULT    = "donnees_visa_clients1.xlsx"

SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

CLIENTS_COLS = [
    "Dossier N","ID_Client","Nom","Date","Mois",
    "Categorie","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
    "Payé","Reste","Paiements","Options",
    "Dossier envoyé","Date d'envoi",
    "Dossier accepté","Date d'acceptation",
    "Dossier refusé","Date de refus",
    "Dossier annulé","Date d'annulation",
    "RFE"
]

# =========================
# Utilitaires
# =========================
def _safe_str(x) -> str:
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)

def _to_num(s: pd.Series) -> pd.Series:
    s = s.astype(str)
    s = s.str.replace(r"[^\d,.\-]", "", regex=True).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0]*len(df), index=df.index, dtype=float)
    v = df[col]
    if pd.api.types.is_numeric_dtype(v):
        return v.fillna(0.0).astype(float)
    return _to_num(v)

def _fmt_money(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _uniquify_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = list(map(str, df.columns))
    seen, out = {}, []
    for c in cols:
        if c not in seen:
            seen[c] = 1
            out.append(c)
        else:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
    df2 = df.copy()
    df2.columns = out
    return df2

def ensure_file(path: str, sheet_name: str, cols: list[str]) -> None:
    p = Path(path)
    if not p.exists():
        df = pd.DataFrame(columns=cols)
        with pd.ExcelWriter(p, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=sheet_name, index=False)

# Crée les fichiers vides au besoin (pour éviter les erreurs au premier démarrage)
ensure_file(CLIENTS_FILE_DEFAULT, SHEET_CLIENTS, CLIENTS_COLS)
ensure_file(VISA_FILE_DEFAULT, SHEET_VISA, ["Categorie","Sous-categorie 1","COS","EOS"])

def _norm(s: str) -> str:
    s2 = unicodedata.normalize("NFKD", s)
    s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
    s2 = s2.strip().lower().replace("\u00a0", " ")
    s2 = s2.replace("-", " ").replace("_", " ")
    return " ".join(s2.split())

# =========================
# Parsing de la feuille Visa
# =========================
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    """
    Retourne: {Categorie: {Sous-categorie: [options...]}}
    - Les options proviennent des colonnes dont la cellule = 1 (ou 'x' / 'oui'…)
    - Injection auto "2-Etudiants" -> F-1/F-2 COS/EOS si aucune catégorie 'etudiant' détectée
    """
    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        if isinstance(v, (int, float)):
            return float(v) == 1.0
        s = str(v).strip().lower()
        return s in {"1","x","true","vrai","oui","yes"}

    out: dict[str, dict[str, list[str]]] = {}
    found_students = False

    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        xls = None

    sheets = [sheet_name] if (sheet_name and xls) else (xls.sheet_names if xls else [])
    if sheets:
        for sn in sheets:
            try:
                dfv = pd.read_excel(xlsx_path, sheet_name=sn)
            except Exception:
                continue
            if dfv.empty:
                continue

            dfv = _uniquify_columns(dfv)
            dfv.columns = dfv.columns.map(str).str.strip()

            cmap = {_norm(c): c for c in dfv.columns}
            cat_col = next((cmap[k] for k in cmap if "categorie" in k), None)
            sub_col = next((cmap[k] for k in cmap if k.startswith("sous")), None)
            if not cat_col:
                continue
            if not sub_col:
                dfv["_Sous_"] = ""
                sub_col = "_Sous_"

            check_cols = [c for c in dfv.columns if c not in {cat_col, sub_col}]

            cats_in_sheet = dfv[cat_col].dropna().astype(str).map(str.strip)
            if any("etudiant" in _norm(c) for c in cats_in_sheet):
                found_students = True

            for _, row in dfv.iterrows():
                cat = _safe_str(row.get(cat_col, "")).strip()
                sub = _safe_str(row.get(sub_col, "")).strip()
                if not cat:
                    continue
                opts = []
                for cc in check_cols:
                    if _is_checked(row.get(cc)):
                        opts.append(f"{sub} {cc}".strip())
                if not opts and sub:
                    opts = [sub]
                if opts:
                    out.setdefault(cat, {})
                    out[cat].setdefault(sub, [])
                    out[cat][sub].extend(opts)

    # Injection pour catégories contenant 'etudiant'
    for cat_name in list(out.keys()):
        if "etudiant" in _norm(cat_name):
            submap = out.setdefault(cat_name, {})
            for sub in ("F-1","F-2"):
                arr = submap.setdefault(sub, [])
                for w in (f"{sub} COS", f"{sub} EOS"):
                    if w not in arr:
                        arr.append(w)
                submap[sub] = sorted(set(arr))

    # Si aucune catégorie étudiants trouvée, on ajoute "2-Etudiants"
    if not found_students:
        out.setdefault("2-Etudiants", {})
        out["2-Etudiants"].setdefault("F-1", ["F-1 COS", "F-1 EOS"])
        out["2-Etudiants"].setdefault("F-2", ["F-2 COS", "F-2 EOS"])

    # Nettoyage final
    for cat, subs in out.items():
        for sub, arr in subs.items():
            subs[sub] = sorted(set(arr))
    return out

# =========================
# I/O & normalisation Clients
# =========================
def _normalize_options_json(x) -> dict:
    try:
        d = json.loads(_safe_str(x) or "{}")
        if not isinstance(d, dict):
            return {}
        excl = d.get("exclusive", None)
        opts = d.get("options", [])
        if not isinstance(opts, list):
            opts = []
        return {"exclusive": excl, "options": [str(o) for o in opts]}
    except Exception:
        return {"exclusive": None, "options": []}

def _normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in CLIENTS_COLS:
        if c not in df.columns:
            df[c] = None

    # Dates
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    for c in ["Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.date

    # Mois
    df["Mois"] = df.apply(
        lambda r: f"{pd.to_datetime(r['Date']).month:02d}" if pd.notna(r["Date"]) else (_safe_str(r.get("Mois",""))[:2] or None),
        axis=1
    )

    # Montants
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        df[c] = _safe_num_series(df, c)

    # Paiements
    def _parse_p(x):
        try:
            j = json.loads(_safe_str(x) or "[]")
            return j if isinstance(j, list) else []
        except Exception:
            return []
    df["Paiements"] = df["Paiements"].apply(_parse_p)

    def _sum_json(lst):
        try:
            return float(sum(float(it.get("amount",0.0) or 0.0) for it in (lst or [])))
        except Exception:
            return 0.0
    paid_json = df["Paiements"].apply(_sum_json)
    df["Payé"] = pd.concat([df["Payé"].fillna(0.0).astype(float), paid_json], axis=1).max(axis=1)

    df["Total (US $)"] = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]
    df["Reste"] = (df["Total (US $)"] - df["Payé"]).clip(lower=0.0)

    # Options (dict JSON)
    df["Options"] = df["Options"].apply(_normalize_options_json)

    # Statuts -> bool
    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        df[c] = df[c].apply(lambda v: bool(str(v).strip().lower() in {"1","true","vrai","oui","yes","x"}))

    # Index temporels auxiliaires
    df["_Année_"]   = df["Date"].apply(lambda d: d.year if pd.notna(d) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda d: d.month if pd.notna(d) else pd.NA)
    return _uniquify_columns(df)

@st.cache_data(show_spinner=False)
def _read_clients(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        df = pd.DataFrame(columns=CLIENTS_COLS)
    return _normalize_clients(df)

def _write_clients(df: pd.DataFrame, path: str) -> None:
    """Ecrit le fichier Clients et pousse l'état précédent dans la pile UNDO."""
    st.session_state.setdefault("undo_stack", [])
    try:
        prev = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        prev = pd.DataFrame(columns=CLIENTS_COLS)
    st.session_state["undo_stack"].append(prev.copy())

    df2 = df.copy()
    df2["Options"] = df2["Options"].apply(lambda d: json.dumps(_normalize_options_json(d), ensure_ascii=False))
    df2["Paiements"] = df2["Paiements"].apply(lambda l: json.dumps(l, ensure_ascii=False))
    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        df2[c] = df2[c].apply(lambda b: 1 if bool(b) else 0)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        _uniquify_columns(df2).to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)

def undo_last_write(path: str):
    """Annule la dernière écriture du fichier Clients, si possible."""
    stack = st.session_state.get("undo_stack", [])
    if not stack:
        st.warning("Aucune opération à annuler.")
        return
    prev_df = stack.pop()
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        _uniquify_columns(prev_df).to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if "Dossier N" in df.columns:
        s = pd.to_numeric(df["Dossier N"], errors="coerce")
        if s.notna().any():
            return int(s.max()) + 1
    return int(start)

def _make_client_id(nom: str, d: date) -> str:
    base = _safe_str(nom).strip().replace(" ","_")
    return f"{base}-{d:%Y%m%d}"


# --- 🧾 ONGLET : Clients ---
with tabs["👤 Clients"]:
    st.markdown("## 👤 Suivi et gestion des clients")

    if "df_clients" not in locals():
        st.error("❌ Fichier client non trouvé.")
    else:
        df_clients_display = df_clients.copy()
        # Ajout des totaux calculés
        if {"Montant honoraires (US $)", "Autres frais (US $)"}.issubset(df_clients_display.columns):
            df_clients_display["Total (US $)"] = (
                df_clients_display["Montant honoraires (US $)"].fillna(0)
                + df_clients_display["Autres frais (US $)"].fillna(0)
            )

        st.markdown("### 📋 Liste des clients")
        st.dataframe(df_clients_display, use_container_width=True)

        # --- Sélection d’un client pour le suivi ---
        st.markdown("### 🔍 Détails d’un client")
        client_names = df_clients_display["Nom"].dropna().unique().tolist()
        selected_client = st.selectbox("Sélectionner un client", [""] + client_names)

        if selected_client:
            cdata = df_clients_display[df_clients_display["Nom"] == selected_client].iloc[0]
            st.markdown(f"#### 👤 {selected_client}")
            cols = st.columns(2)
            cols[0].metric("Total dû", f"{cdata.get('Total (US $)', 0):,.2f} $")
            cols[1].metric("Payé", f"{cdata.get('Payé (US $)', 0):,.2f} $")

            st.markdown("##### 📦 Informations principales")
            st.write({
                "Catégorie": cdata.get("Catégorie", ""),
                "Visa": cdata.get("Visa", ""),
                "Sous-catégorie": cdata.get("Sous-catégorie", ""),
                "Pays": cdata.get("Pays", ""),
                "Date de création": cdata.get("Date", ""),
            })

            # --- Section de suivi du dossier ---
            st.markdown("##### 🧾 Statuts du dossier")
            s1, s2, s3 = st.columns(3)
            s1.checkbox("Dossier envoyé", key="sent_chk")
            s1.date_input("Date d’envoi", key="sent_date")
            s2.checkbox("Dossier accepté", key="approved_chk")
            s2.date_input("Date d’acceptation", key="approved_date")
            s3.checkbox("Dossier refusé", key="refused_chk")
            s3.date_input("Date de refus", key="refused_date")

            st.markdown("##### ⚠️ Autres statuts")
            c1, c2 = st.columns(2)
            c1.checkbox("RFE reçu", key="rfe_chk")
            c1.date_input("Date RFE", key="rfe_date")
            c2.checkbox("Dossier annulé", key="cancel_chk")
            c2.date_input("Date d’annulation", key="cancel_date")

            st.markdown("##### 💵 Paiements")
            pay_history = st.data_editor(
                pd.DataFrame(columns=["Date", "Montant (US $)", "Méthode"]),
                num_rows="dynamic",
                key="pay_hist",
            )
            st.success("✅ Suivi mis à jour (simulation).")

# --- 🧾 ONGLET : Dashboard général ---
with tabs["📊 Dashboard"]:
    st.markdown("## 📊 Tableau de bord des visas")

    if "df_clients" not in locals():
        st.error("❌ Données clients non disponibles.")
    else:
        st.markdown("### 🎯 Indicateurs globaux")

        total_clients = len(df_clients)
        total_visa = df_clients["Visa"].nunique()
        total_amount = df_clients["Montant honoraires (US $)"].sum()

        c1, c2, c3 = st.columns(3)
        c1.metric("Total clients", total_clients)
        c2.metric("Types de visa", total_visa)
        c3.metric("Total honoraires", f"{total_amount:,.2f} $")

        # Filtres latéraux
        st.sidebar.markdown("### 🎚️ Filtres Dashboard")
        fy = st.sidebar.multiselect("Année", sorted(df_clients["Année"].dropna().unique()))
        fv = st.sidebar.multiselect("Visa", sorted(df_clients["Visa"].dropna().unique()))
        fc = st.sidebar.multiselect("Catégorie", sorted(df_clients["Catégorie"].dropna().unique()))

        df_dash = df_clients.copy()
        if fy:
            df_dash = df_dash[df_dash["Année"].isin(fy)]
        if fv:
            df_dash = df_dash[df_dash["Visa"].isin(fv)]
        if fc:
            df_dash = df_dash[df_dash["Catégorie"].isin(fc)]

        st.markdown("### 📈 Répartition par visa")
        visa_chart = df_dash["Visa"].value_counts().reset_index()
        visa_chart.columns = ["Visa", "Nombre"]
        st.bar_chart(visa_chart.set_index("Visa"))

        st.markdown("### 💰 Somme totale par catégorie")
        cat_chart = (
            df_dash.groupby("Catégorie")["Montant honoraires (US $)"].sum().sort_values(ascending=False)
        )
        st.bar_chart(cat_chart)


# =========================
# 📈 ONGLET : Analyses
# =========================
with tabs["📈 Analyses"]:
    st.markdown("## 📈 Analyses")
    if "df_clients" not in locals() or df_clients.empty:
        st.info("Aucune donnée client disponible pour l’analyse.")
    else:
        # Filtres simples d'analyse
        yearsA  = sorted(df_clients["Année"].dropna().unique()) if "Année" in df_clients.columns else []
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_clients["Catégorie"].dropna().astype(str).unique()) if "Catégorie" in df_clients.columns else []
        visasA  = sorted(df_clients["Visa"].dropna().astype(str).unique()) if "Visa" in df_clients.columns else []

        a1, a2, a3, a4 = st.columns(4)
        fy = a1.multiselect("Année", yearsA, default=[])
        fm = a2.multiselect("Mois (MM)", monthsA, default=[])
        fc = a3.multiselect("Catégorie", catsA, default=[])
        fv = a4.multiselect("Visa", visasA, default=[])

        dfA = df_clients.copy()
        if fy: dfA = dfA[dfA["Année"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Catégorie"].astype(str).isin(fc)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPID
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", len(dfA))
        k2.metric("Honoraires", f"${float(dfA['Montant honoraires (US $)'].fillna(0).sum()):,.2f}")
        paid = float(dfA.get("Payé", 0).fillna(0).sum()) if "Payé" in dfA.columns else 0.0
        rest = float(dfA.get("Reste", 0).fillna(0).sum()) if "Reste" in dfA.columns else 0.0
        k3.metric("Payé", f"${paid:,.2f}")
        k4.metric("Reste", f"${rest:,.2f}")

        st.markdown("### 📊 Dossiers par catégorie")
        if not dfA.empty:
            grp_cat = dfA.groupby("Catégorie", as_index=False).size().rename(columns={"size":"Nombre"})
            st.bar_chart(grp_cat.set_index("Catégorie"))

        st.markdown("### 📈 Honoraires par mois")
        if not dfA.empty:
            dfA["Mois"] = dfA["Mois"].astype(str)
            grp_m = dfA.groupby("Mois", as_index=False)["Montant honoraires (US $)"].sum()
            grp_m = grp_m.sort_values("Mois")
            st.line_chart(grp_m.set_index("Mois"))

        st.markdown("### 🧾 Détails (après filtres)")
        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Catégorie","Sous-catégorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in dfA.columns]
        st.dataframe(dfA[show_cols].reset_index(drop=True), use_container_width=True)


# =========================
# 🏦 ONGLET : Escrow (synthèse)
# =========================
with tabs["🏦 Escrow"]:
    st.markdown("## 🏦 Escrow — synthèse")
    if "df_clients" not in locals() or df_clients.empty:
        st.info("Aucun client.")
    else:
        dfE = df_clients.copy()
        # Calculs simples : on assimile Payé comme fonds reçus, Reste à encaisser, etc.
        dfE["Payé"] = pd.to_numeric(dfE.get("Payé", 0), errors="coerce").fillna(0.0)
        dfE["Reste"] = pd.to_numeric(dfE.get("Reste", 0), errors="coerce").fillna(0.0)
        dfE["Total (US $)"] = pd.to_numeric(dfE.get("Total (US $)", 0), errors="coerce").fillna(0.0)
        agg = dfE.groupby("Catégorie", as_index=False)[["Total (US $)","Payé","Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg["Total (US $)"]).replace([pd.NA, pd.NaT, float("inf")], 0).fillna(0.0) * 100
        st.dataframe(agg, use_container_width=True)

        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", f"${float(dfE['Total (US $)'].sum()):,.2f}")
        t2.metric("Payé", f"${float(dfE['Payé'].sum()):,.2f}")
        t3.metric("Reste", f"${float(dfE['Reste'].sum()):,.2f}")

        st.caption("NB : Pour un vrai compte escrow, on peut isoler les honoraires perçus avant envoi, puis déclencher un transfert lors du statut 'Dossier envoyé'.")


# =========================
# 📄 ONGLET : Visa (aperçu brut)
# =========================
with tabs["📄 Visa (aperçu)"]:
    st.markdown("## 📄 Aperçu du fichier Visa")
    try:
        visa_preview = pd.read_excel(visa_file_path, sheet_name=visa_sheet or 0)
        st.dataframe(visa_preview, use_container_width=True)
    except Exception as e:
        st.error(f"Impossible de lire la feuille Visa : {e}")


# =========================
# 💾 Export global (Clients + Visa)
# =========================
st.markdown("---")
st.markdown("### 💾 Export global")
colz1, colz2 = st.columns([1,3])
with colz1:
    if st.button("Préparer l’archive ZIP"):
        try:
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Réécriture Clients “propres”
                df_export = df_clients.copy()
                with BytesIO() as xbuf:
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                        df_export.to_excel(wr, sheet_name=clients_sheet, index=False)
                    zf.writestr("Clients.xlsx", xbuf.getvalue())
                # Visa original
                zf.write(visa_file_path, "Visa.xlsx")
            st.session_state["zip_export"] = buf.getvalue()
            st.success("Archive prête. Cliquez pour télécharger.")
        except Exception as e:
            st.error(f"Erreur lors de la préparation : {e}")

with colz2:
    if st.session_state.get("zip_export"):
        st.download_button(
            label="⬇️ Télécharger l’export (ZIP)",
            data=st.session_state["zip_export"],
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
        )


# =========================
# 👤 Vue Compte Client — Statuts & dates (bloc corrigé)
# =========================
def render_client_status_block(row: pd.Series):
    st.markdown("### 📌 Statuts & dates")
    s1, s2, s3, s4, s5 = st.columns(5)

    # Envoyé
    s1.write("**Envoyé** : " + ("✅" if bool(row.get("Dossier envoyé")) else "❌"))
    s1.write("• Date : " + _safe_str(row.get("Date d'envoi", "")))

    # Accepté
    s2.write("**Accepté** : " + ("✅" if bool(row.get("Dossier accepté")) else "❌"))
    s2.write("• Date : " + _safe_str(row.get("Date d'acceptation", "")))

    # Refusé
    s3.write("**Refusé** : " + ("✅" if bool(row.get("Dossier refusé")) else "❌"))
    s3.write("• Date : " + _safe_str(row.get("Date de refus", "")))

    # Annulé
    s4.write("**Annulé** : " + ("✅" if bool(row.get("Dossier annulé")) else "❌"))
    s4.write("• Date : " + _safe_str(row.get("Date d'annulation", "")))

    # RFE
    s5.write("**RFE** : " + ("✅" if bool(row.get("RFE")) else "❌"))