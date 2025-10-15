# ================================
# PARTIE 1/6 — Imports, constantes, helpers, I/O (corrigée)
# ================================
from io import BytesIO
from datetime import date, datetime
from uuid import uuid4
import json, os, zipfile
import pandas as pd
import streamlit as st

# ---- Config de page & SID unique pour les clés Streamlit
st.set_page_config(page_title="Visa Manager", layout="wide")
if "SID" not in st.session_state:
    st.session_state["SID"] = str(uuid4())
SID = st.session_state["SID"]

# ---- Canon de colonnes côté Clients (tolérant)
COLS_CANON = {
    "ID_Client": ["ID_Client", "ID Client", "Id_Client", "Id Client"],
    "Dossier N": ["Dossier N", "Dossier No", "N° Dossier", "Dossier", "Dossier Numero"],
    "Nom": ["Nom", "Client", "Name"],
    "Date": ["Date", "Date creation", "Date de creation", "Date de création"],
    "Mois": ["Mois", "Month"],
    "Categorie": ["Categorie", "Catégorie", "Categories", "Catégories"],
    "Sous-categorie": ["Sous-categorie", "Sous-catégorie", "Sous categorie", "Sous categories"],
    "Visa": ["Visa", "Type Visa"],
    "Montant honoraires (US $)": ["Montant honoraires (US $)", "Honoraires (US $)", "Montant honoraires"],
    "Autres frais (US $)": ["Autres frais (US $)", "Autres frais"],
    "Total (US $)": ["Total (US $)", "Total"],
    "Payé": ["Payé", "Paye", "Encaisse"],
    "Reste": ["Reste", "Solde"],
    "Acompte 1": ["Acompte 1", "Acompte1"],
    "Acompte 2": ["Acompte 2", "Acompte2"],
    "RFE": ["RFE"],
    "Dossier envoyé": ["Dossier envoyé", "Envoye", "Dossiers envoyé"],
    "Dossier accepté": ["Dossier accepté", "Accepte"],
    "Dossier refusé": ["Dossier refusé", "Refuse"],
    "Dossier annulé": ["Dossier annulé", "Annule"],
    "Date d'envoi": ["Date d'envoi", "Date envoi"],
    "Date d'acceptation": ["Date d'acceptation", "Date acceptation"],
    "Date de refus": ["Date de refus", "Date refus"],
    "Date d'annulation": ["Date d'annulation", "Date annulation"],
    "Commentaires": ["Commentaires", "Commentaire", "Notes"],
    "Paiements": ["Paiements", "Reglements", "Règlements"],
}

SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"
STATE_FILE    = ".visa_state.json"

# ---- Helpers sûrs
def _safe_str(v):
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return "" if v is None else str(v)

def _to_num(v, default=0.0):
    try:
        return float(pd.to_numeric(v, errors="coerce"))
    except Exception:
        return default

def _series_num(df, col):
    if col not in df.columns:
        return pd.Series([0.0]*len(df), index=df.index, dtype=float)
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0).astype(float)

def _fmt_money(x):
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _date_for_widget(v, fallback=None):
    fb = fallback if isinstance(fallback, (date, datetime)) else date.today()
    if isinstance(v, date): return v
    if isinstance(v, datetime): return v.date()
    try:
        d2 = pd.to_datetime(v, errors="coerce")
        if pd.notna(d2): return d2.date()
    except Exception:
        pass
    return fb if isinstance(fb, date) else date.today()

# ---- Persistance des derniers chemins
def save_last_paths(clients_path, visa_path):
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump({"clients": _safe_str(clients_path), "visa": _safe_str(visa_path)}, f)
    except Exception:
        pass

def load_last_paths():
    try:
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                obj = json.load(f)
                return (obj.get("clients",""), obj.get("visa",""))
    except Exception:
        pass
    return ("","")

# ---- Lecture générique (xlsx/csv, chemin ou upload)
def read_any_table(source, sheet_name=None) -> pd.DataFrame|None:
    if source is None:
        return None
    # Upload (BytesIO or UploadedFile)
    if hasattr(source, "read"):
        try:
            source.seek(0)
            return pd.read_excel(source, sheet_name=sheet_name) if sheet_name else pd.read_excel(source)
        except Exception:
            try:
                source.seek(0)
                return pd.read_csv(source)
            except Exception:
                return None
    # Chemin
    p = str(source)
    if not os.path.exists(p):
        return None
    ext = os.path.splitext(p)[1].lower()
    try:
        if ext in [".xlsx", ".xls"]:
            return pd.read_excel(p, sheet_name=sheet_name) if sheet_name else pd.read_excel(p)
        if ext == ".csv":
            return pd.read_csv(p)
    except Exception:
        return None
    return None

# ---- Normalisation Clients
def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols_map = {}
    for canon, variants in COLS_CANON.items():
        for v in variants:
            for c in df.columns:
                if _safe_str(c).strip().lower() == _safe_str(v).strip().lower():
                    cols_map[c] = canon
    df = df.rename(columns=cols_map)

    # Colonnes minimales
    for need in ["ID_Client","Dossier N","Nom","Date","Mois",
                 "Categorie","Sous-categorie","Visa",
                 "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
                 "Payé","Reste","Paiements","Commentaires",
                 "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE",
                 "Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation"]:
        if need not in df.columns:
            df[need] = "" if need in ["ID_Client","Nom","Categorie","Sous-categorie","Visa","Paiements","Commentaires"] else 0

    # Types numériques
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        df[c] = _series_num(df, c)

    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    # Dates
    for c in ["Date","Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Mois / Année techniques
    if "Mois" not in df.columns:
        df["Mois"] = df["Date"].dt.month.fillna(1).astype(int).map(lambda m: f"{int(m):02d}")
    else:
        df["Mois"] = df["Mois"].astype(str).str.zfill(2)

    df["_Année_"]   = pd.to_numeric(df["Date"].dt.year, errors="coerce").fillna(1900).astype(int)
    df["_MoisNum_"] = pd.to_numeric(df["Mois"], errors="coerce").fillna(1).astype(int)

    # Total / Reste
    mask_total_zero = (df["Total (US $)"]<=0) | df["Total (US $)"].isna()
    df.loc[mask_total_zero, "Total (US $)"] = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]
    df["Reste"] = df["Total (US $)"] - df["Payé"]
    return df

def read_clients_file(path_or_io) -> pd.DataFrame:
    df = read_any_table(path_or_io) or read_any_table(path_or_io, sheet_name=SHEET_CLIENTS)
    return normalize_columns(df) if df is not None else pd.DataFrame()

def write_clients_file(df: pd.DataFrame, dest_path: str|BytesIO):
    if hasattr(dest_path, "write") and not isinstance(dest_path, str):
        with pd.ExcelWriter(dest_path, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        return
    with pd.ExcelWriter(dest_path, engine="openpyxl") as wr:
        df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)



# ================================
# PARTIE 2/6 — Lecture & normalisation VISA
# ================================
def _normalize_visa_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Renomme Catégorie / Sous-catégorie, enlève colonnes vides, nettoie en-têtes."""
    if df is None or df.empty:
        return pd.DataFrame()

    # Supprimer entièrement les colonnes vides (100% NA)
    df = df.loc[:, ~df.isna().all(axis=0)]

    # Renommer Catégorie/Sous-catégorie
    ren = {}
    for c in df.columns:
        cl = _safe_str(c).strip().lower()
        if cl in ["categorie", "catégorie", "categories", "catégories"]:
            ren[c] = "Categorie"
        elif cl in ["sous-categorie", "sous-catégorie", "sous categorie", "sous-categories", "sous categories"]:
            ren[c] = "Sous-categorie"
    if ren:
        df = df.rename(columns=ren)

    # S’assurer des colonnes de base
    if "Categorie" not in df.columns: df["Categorie"] = ""
    if "Sous-categorie" not in df.columns: df["Sous-categorie"] = ""

    # Enlever lignes totalement vides
    df = df.dropna(how="all").reset_index(drop=True)

    # Nettoyer espaces/NaN sur Catégorie/Sous-catégorie
    df["Categorie"] = df["Categorie"].apply(lambda x: _safe_str(x).strip())
    df["Sous-categorie"] = df["Sous-categorie"].apply(lambda x: _safe_str(x).strip())

    # Garder seulement les lignes qui ont au moins Catégorie & Sous-catégorie
    df = df[(df["Categorie"]!="") & (df["Sous-categorie"]!="")].copy()

    # Transformer les “1” en bool/int pour les colonnes d’options
    for c in df.columns:
        if c not in ["Categorie","Sous-categorie"]:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    return df

def build_visa_map(visa_df: pd.DataFrame) -> dict:
    """
    Construit la hiérarchie Catégorie -> Sous-catégorie -> options disponibles.
    Les options sont toutes les colonnes (hors Catégorie / Sous-catégorie) dont la valeur == 1 sur la ligne.
    """
    d = {}
    if visa_df is None or visa_df.empty:
        return d

    opt_cols = [c for c in visa_df.columns if c not in ["Categorie","Sous-categorie"]]
    for _, r in visa_df.iterrows():
        cat = _safe_str(r.get("Categorie")).strip()
        sub = _safe_str(r.get("Sous-categorie")).strip()
        if not cat or not sub:
            continue
        d.setdefault(cat, {})
        d[cat].setdefault(sub, {"exclusive": None, "options": []})

        opts = []
        for oc in opt_cols:
            try:
                if int(r.get(oc, 0)) == 1:
                    opts.append(_safe_str(oc).strip())
            except Exception:
                pass
        d[cat][sub]["options"] = sorted(list(set(opts)))
    return d

@st.cache_data(show_spinner=False)
def read_visa_file(path_or_io) -> tuple[pd.DataFrame, dict]:
    """
    Lit le fichier Visa (xlsx/csv), normalise, et renvoie (visa_df_norm, visa_map).
    - Si le fichier a plusieurs onglets, essaie 'Visa' en priorité.
    """
    # 1) lecture brute
    df = read_any_table(path_or_io, sheet_name=SHEET_VISA)
    if df is None:
        df = read_any_table(path_or_io)  # peut être déjà la bonne feuille

    if df is None or df.empty:
        return pd.DataFrame(), {}

    # 2) normalisation des colonnes
    df_norm = _normalize_visa_columns(df)

    # 3) carte des visas
    vmap = build_visa_map(df_norm)
    return df_norm, vmap




# ================================
# PARTIE 3/6 — 📊 Dashboard
# ================================
with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # KPIs
        left, right = st.columns([1.2, 2.8])
        with left:
            k1, k2 = st.columns([1,1])
            k3, k4 = st.columns([1,1])
            k1.metric("Dossiers", f"{len(df_all)}")
            k2.metric("Honoraires+Frais", _fmt_money((_series_num(df_all,"Montant honoraires (US $)") + _series_num(df_all,"Autres frais (US $)")).sum()))
            k3.metric("Payé", _fmt_money(_series_num(df_all, "Payé").sum()))
            k4.metric("Solde", _fmt_money(_series_num(df_all, "Reste").sum()))
            # % envoyés
            pct_env = 0.0
            if len(df_all) > 0:
                pct_env = 100.0 * (_series_num(df_all, "Dossier envoyé")>0).sum() / len(df_all)
            st.metric("Envoyés (%)", f"{pct_env:0.0f}%")

        with right:
            st.markdown("#### 🎛️ Filtres")
            cats  = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
            subs  = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
            visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

            a1, a2, a3 = st.columns(3)
            fc = a1.multiselect("Catégories", cats, default=[], key=f"dash_c_{SID}")
            fs = a2.multiselect("Sous-catégories", subs, default=[], key=f"dash_s_{SID}")
            fv = a3.multiselect("Visa", visas, default=[], key=f"dash_v_{SID}")

        view = df_all.copy()
        if fc: view = view[view["Categorie"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if not view.empty and "Categorie" in view.columns:
            vc = view["Categorie"].value_counts().sort_index()
            st.bar_chart(vc)

        st.markdown("#### 💵 Flux par mois")
        flux = pd.DataFrame({
            "Mois": view["Mois"].astype(str),
            "Montant honoraires (US $)": _series_num(view, "Montant honoraires (US $)"),
            "Autres frais (US $)": _series_num(view, "Autres frais (US $)"),
            "Payé": _series_num(view, "Payé"),
            "Solde": _series_num(view, "Reste")
        })
        flux = flux.groupby("Mois", as_index=False).sum().sort_values("Mois")
        st.line_chart(flux.set_index("Mois"))

        st.markdown("#### 📋 Détails (après filtres)")
        det = view.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in det.columns:
                det[c] = _series_num(det, c).map(_fmt_money)
        if "Date" in det.columns:
            det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]
        sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
        det = det.sort_values(by=sort_keys) if sort_keys else det
        st.dataframe(det[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_table_{SID}")




# ================================
# PARTIE 4/6 — 📈 Analyses / 🏦 Escrow / 📄 Visa (aperçu)
# ================================

# -------- Analyses --------
with tabs[1]:
    st.subheader("📈 Analyses")
    if df_all.empty:
        st.info("Aucune donnée.")
    else:
        yearsA  = sorted(df_all["_Année_"].dropna().astype(int).unique().tolist())
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist())
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist())
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist())

        a1,a2,a3,a4,a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_y_{SID}")
        fm = a2.multiselect("Mois", monthsA, default=[], key=f"a_m_{SID}")
        fc = a3.multiselect("Catégories", catsA, default=[], key=f"a_c_{SID}")
        fs = a4.multiselect("Sous-catégories", subsA, default=[], key=f"a_s_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_v_{SID}")

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(_series_num(dfA, "Montant honoraires (US $)").sum()))
        k3.metric("Payé", _fmt_money(_series_num(dfA, "Payé").sum()))
        k4.metric("Reste", _fmt_money(_series_num(dfA, "Reste").sum()))

        st.markdown("#### 📦 Répartition par catégorie (en %)")
        if not dfA.empty:
            vc = dfA["Categorie"].value_counts(dropna=True)
            pct = (vc / vc.sum() * 100.0).round(1)
            st.bar_chart(pct.sort_index())

        st.markdown("#### 🧩 Répartition par sous-catégorie (en %)")
        if not dfA.empty:
            vs = dfA["Sous-categorie"].value_counts(dropna=True)
            pct2 = (vs / vs.sum() * 100.0).round(1)
            st.bar_chart(pct2.sort_index())

        st.markdown("#### 📈 Honoraires par mois")
        tmp = dfA.copy()
        g = tmp.groupby("Mois", as_index=False)["Montant honoraires (US $)"].sum().sort_values("Mois")
        st.line_chart(g.set_index("Mois"))

        st.markdown("#### 🧾 Détails")
        det = dfA.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in det.columns:
                det[c] = _series_num(det, c).map(_fmt_money)
        if "Date" in det.columns:
            det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)

        cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]
        st.dataframe(det[cols].reset_index(drop=True), use_container_width=True, key=f"a_table_{SID}")

# -------- Escrow --------
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Total (US $)"] = _series_num(dfE, "Total (US $)")
        dfE["Payé"] = _series_num(dfE, "Payé")
        dfE["Reste"] = _series_num(dfE, "Reste")

        t1,t2,t3 = st.columns(3)
        t1.metric("Total", _fmt_money(dfE["Total (US $)"].sum()))
        t2.metric("Payé", _fmt_money(dfE["Payé"].sum()))
        t3.metric("Reste", _fmt_money(dfE["Reste"].sum()))

        st.markdown("#### Par catégorie")
        agg = dfE.groupby("Categorie", as_index=False)[["Total (US $)","Payé","Reste"]].sum()
        st.dataframe(agg, use_container_width=True, key=f"esc_agg_{SID}")

        st.caption("NB : pour un suivi ESCROW strict, on peut isoler les honoraires pré-envoi et déclencher un transfert quand le statut passe à « Envoyé ».")

# -------- Visa (aperçu) --------
with tabs[5]:
    st.subheader("📄 Visa — aperçu & filtres")
    if df_visa_raw.empty:
        st.info("Aucun fichier Visa.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True, key=f"visa_raw_{SID}")
        st.markdown("#### Carte Catégorie → Sous-catégorie → Options disponibles")
        cats = sorted(list(visa_map.keys()))
        c1, c2 = st.columns(2)
        selc = c1.selectbox("Catégorie", [""]+cats, index=0, key=f"v_cat_{SID}")
        if selc:
            subs = sorted(list(visa_map.get(selc, {}).keys()))
            sels = c2.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"v_sub_{SID}")
            if sels:
                opts = visa_map.get(selc,{}).get(sels,{}).get("options",[])
                st.write("**Options** :", ", ".join(opts) if opts else "Aucune (visa direct)")




# ================================
# PARTIE 5/6 — 👤 Compte client (timeline + paiements)
# ================================
with tabs[3]:
    st.subheader("👤 Compte client")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        names = sorted(df_all["Nom"].dropna().astype(str).unique().tolist())
        ids   = sorted(df_all["ID_Client"].dropna().astype(str).unique().tolist())
        c1,c2 = st.columns(2)
        pick_name = c1.selectbox("Nom", [""]+names, index=0, key=f"acc_nom_{SID}")
        pick_id   = c2.selectbox("ID_Client", [""]+ids, index=0, key=f"acc_id_{SID}")

        mask = None
        if pick_id:
            mask = (df_all["ID_Client"].astype(str) == pick_id)
        elif pick_name:
            mask = (df_all["Nom"].astype(str) == pick_name)

        if mask is not None and mask.any():
            row = df_all[mask].iloc[0].copy()

            st.markdown("#### 📌 Dossier")
            s1,s2,s3,s4 = st.columns(4)
            s1.write(f"Dossier N : {_safe_str(row.get('Dossier N',''))}")
            s2.write(f"Nom : {_safe_str(row.get('Nom',''))}")
            s3.write(f"Visa : {_safe_str(row.get('Visa',''))}")
            s4.write(f"Catégorie : {_safe_str(row.get('Categorie',''))} / {_safe_str(row.get('Sous-categorie',''))}")

            st.markdown("#### ✅ Statut & dates")
            env = int(_to_num(row.get("Dossier envoyé",0)))==1
            acc = int(_to_num(row.get("Dossier accepté",0)))==1
            ref = int(_to_num(row.get("Dossier refusé",0)))==1
            ann = int(_to_num(row.get("Dossier annulé",0)))==1
            rfe = int(_to_num(row.get("RFE",0)))==1

            colA, colB = st.columns(2)
            with colA:
                st.write("- Dossier envoyé :", "Oui" if env else "Non",
                         "| Date :", _safe_str(row.get("Date d'envoi","")))
                st.write("- Dossier accepté :", "Oui" if acc else "Non",
                         "| Date :", _safe_str(row.get("Date d'acceptation","")))
                st.write("- Dossier refusé :", "Oui" if ref else "Non",
                         "| Date :", _safe_str(row.get("Date de refus","")))
                st.write("- Dossier annulé :", "Oui" if ann else "Non",
                         "| Date :", _safe_str(row.get("Date d'annulation","")))
            with colB:
                st.write("- RFE :", "Oui" if rfe else "Non")

            st.markdown("#### 💳 Paiements")
            # Paiements stockés en JSON ou liste
            rawp = row.get("Paiements","")
            pay_list = []
            if isinstance(rawp, list):
                pay_list = rawp
            else:
                try:
                    pay_list = json.loads(_safe_str(rawp) or "[]")
                    if not isinstance(pay_list, list): pay_list = []
                except Exception:
                    pay_list = []

            if pay_list:
                dfp = pd.DataFrame(pay_list)
                if "date" in dfp.columns:
                    try:
                        dfp["date"] = pd.to_datetime(dfp["date"], errors="coerce").dt.date.astype(str)
                    except Exception:
                        pass
                st.dataframe(dfp, use_container_width=True, key=f"pay_hist_{SID}")
            else:
                st.info("Aucun paiement saisi.")

            st.markdown("##### ➕ Ajouter un paiement (tant que le dossier n’est pas soldé)")
            reste = float(_to_num(row.get("Reste", 0.0)))
            if reste <= 0:
                st.success("Ce dossier est soldé.")
            else:
                pcol1,pcol2,pcol3,pcol4 = st.columns([1,1,1,2])
                pdate = pcol1.date_input("Date", value=date.today(), key=f"pay_date_{SID}")
                pamt  = pcol2.number_input("Montant", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=f"pay_amt_{SID}")
                pmode = pcol3.selectbox("Mode", ["Chèque","CB","Cash","Virement","Venmo"], index=1, key=f"pay_mode_{SID}")
                pok   = pcol4.button("💾 Enregistrer le paiement", key=f"pay_save_{SID}")

                if pok:
                    add = float(pamt or 0.0)
                    if add <= 0:
                        st.warning("Montant > 0 requis.")
                    else:
                        # MAJ paiements + Payé + Reste
                        pay_list.append({
                            "date": pdate.strftime("%Y-%m-%d"),
                            "montant": add,
                            "mode": pmode
                        })
                        # Recalcule
                        paye_new = float(_to_num(row.get("Payé", 0.0))) + add
                        total    = float(_to_num(row.get("Total (US $)", 0.0)))
                        reste_new= max(0.0, total - paye_new)

                        # Persister dans df_all puis fichier source
                        idx_global = df_all[mask].index[0]
                        df_all.at[idx_global, "Paiements"] = json.dumps(pay_list, ensure_ascii=False)
                        df_all.at[idx_global, "Payé"] = paye_new
                        df_all.at[idx_global, "Reste"] = reste_new

                        # Écrire dans fichier clients
                        try:
                            write_clients_file(df_all, clients_src if isinstance(clients_src,str) else save_clients_to or "clients_sauvegarde.xlsx")
                            st.success("Paiement ajouté.")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur sauvegarde : {_safe_str(e)}")




# ================================
# PARTIE 6/6 — 🧾 Gestion (CRUD) + 💾 Export
# ================================
with tabs[4]:
    st.subheader("🧾 Gestion (Ajouter / Modifier / Supprimer)")

    # Helpers statut
    def _status_to_flags(status: str):
        s = (status or "").strip().lower()
        return {
            "Dossier envoyé":  1 if s=="envoyé" else 0,
            "Dossier accepté": 1 if s=="accepté" else 0,
            "Dossier refusé":  1 if s=="refusé" else 0,
            "Dossier annulé":  1 if s=="annulé" else 0,
        }
    def _flags_to_status(row):
        if int(_to_num(row.get("Dossier accepté",0)))==1: return "Accepté"
        if int(_to_num(row.get("Dossier refusé",0)))==1:  return "Refusé"
        if int(_to_num(row.get("Dossier annulé",0)))==1:  return "Annulé"
        if int(_to_num(row.get("Dossier envoyé",0)))==1:  return "Envoyé"
        return "Aucun"
    def _status_date_key(statut):
        lut = {"Envoyé":"Date d'envoi","Accepté":"Date d'acceptation","Refusé":"Date de refus","Annulé":"Date d'annulation"}
        return lut.get(statut, None)

    df_live = df_all.copy()

    op = st.radio("Action", ["Ajouter","Modifier","Supprimer"], horizontal=True, key=f"crud_op_{SID}")

    # ------- AJOUT -------
    if op == "Ajouter":
        c1,c2,c3 = st.columns(3)
        nom = c1.text_input("Nom", "", key=f"add_nom_{SID}")
        dt  = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois= c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=(date.today().month-1), key=f"add_mois_{SID}")

        st.markdown("#### 🎯 Visa")
        cats = sorted(list(visa_map.keys()))
        cat  = st.selectbox("Catégorie", [""]+cats, index=0, key=f"add_cat_{SID}")
        sub  = ""
        visa_final, opts_dict = "", {"exclusive": None, "options":[]}
        if cat:
            subs = sorted(list(visa_map.get(cat,{}).keys()))
            sub  = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"add_sub_{SID}")
            if sub:
                opts = visa_map.get(cat,{}).get(sub,{}).get("options",[])
                st.caption("Options (issues du fichier Visa) : " + (", ".join(opts) if opts else "aucune"))

        f1,f2 = st.columns(2)
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
        other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=f"add_o_{SID}")
        comment = st.text_area("Commentaires (Autres frais / notes)", "", key=f"add_comm_{SID}")

        st.markdown("#### 📌 Statut & RFE")
        st_choices = ["Aucun","Envoyé","Accepté","Refusé","Annulé"]
        statut = st.selectbox("Statut", st_choices, index=0, key=f"add_stat_{SID}")
        rfe_on = st.toggle("RFE", value=False, key=f"add_rfe_{SID}")
        if rfe_on and statut=="Aucun":
            st.warning("RFE nécessite un statut sélectionné.")

        dkey = _status_date_key(statut)
        stat_date = None
        if statut!="Aucun":
            stat_date = st.date_input(f"Date pour « {statut} »", value=date.today(), key=f"add_statd_{SID}")

        if st.button("💾 Enregistrer le client", key=f"btn_add_{SID}"):
            if not nom or not cat or not sub:
                st.warning("Nom, Catégorie, Sous-catégorie requis.")
            else:
                did = f"{_safe_str(nom).strip()}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
                dossier_n = int(df_live["Dossier N"].max())+1 if "Dossier N" in df_live.columns and not df_live.empty else 13057
                total = float(honor)+float(other)
                row = {
                    "Dossier N": dossier_n, "ID_Client": did, "Nom": nom,
                    "Date": dt, "Mois": mois,
                    "Categorie": cat, "Sous-categorie": sub, "Visa": sub,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Total (US $)": total,
                    "Payé": 0.0, "Reste": total,
                    "Paiements": json.dumps([], ensure_ascii=False),
                    "Commentaires": comment,
                    "Dossier envoyé":0, "Dossier accepté":0, "Dossier refusé":0, "Dossier annulé":0,
                    "Date d'envoi": None, "Date d'acceptation": None, "Date de refus": None, "Date d'annulation": None,
                    "RFE": 1 if (rfe_on and statut!="Aucun") else 0
                }
                flags = _status_to_flags(statut)
                for k,v in flags.items(): row[k]=v
                if dkey: row[dkey]=stat_date
                df_live = pd.concat([df_live, pd.DataFrame([row])], ignore_index=True)
                write_clients_file(df_live, clients_src if isinstance(clients_src,str) else (save_clients_to or "clients_sauvegarde.xlsx"))
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()

    # ------- MODIFIER -------
    elif op == "Modifier":
        if df_live.empty:
            st.info("Aucun client à modifier.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
            m1,m2 = st.columns(2)
            tgt_name = m1.selectbox("Nom", [""]+names, index=0, key=f"mod_nom_{SID}")
            tgt_id   = m2.selectbox("ID_Client", [""]+ids, index=0, key=f"mod_id_{SID}")

            mask=None
            if tgt_id: mask = (df_live["ID_Client"].astype(str)==tgt_id)
            elif tgt_name: mask = (df_live["Nom"].astype(str)==tgt_name)

            if mask is not None and mask.any():
                idx = df_live[mask].index[0]
                row = df_live.loc[idx].copy()

                d1,d2,d3 = st.columns(3)
                nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nv_{SID}")
                dt   = d2.date_input("Date de création", value=_date_for_widget(row.get("Date")), key=f"mod_dt_{SID}")
                mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                    index=max(0, int(_safe_str(row.get("Mois","01"))[:2]) - 1), key=f"mod_m_{SID}")

                # Visa cascade
                st.markdown("#### 🎯 Visa")
                cats = sorted(list(visa_map.keys()))
                preset_cat = _safe_str(row.get("Categorie",""))
                cat  = st.selectbox("Catégorie", [""]+cats, index=(cats.index(preset_cat)+1 if preset_cat in cats else 0), key=f"mod_cat_{SID}")
                sub  = _safe_str(row.get("Sous-categorie",""))
                if cat:
                    subs = sorted(list(visa_map.get(cat,{}).keys()))
                    sub  = st.selectbox("Sous-catégorie", [""]+subs, index=(subs.index(sub)+1 if sub in subs else 0), key=f"mod_sub_{SID}")

                f1,f2 = st.columns(2)
                honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=float(_to_num(row.get("Montant honoraires (US $)",0.0))), step=50.0, format="%.2f", key=f"mod_h_{SID}")
                other = f2.number_input("Autres frais (US $)", min_value=0.0, value=float(_to_num(row.get("Autres frais (US $)",0.0))), step=20.0, format="%.2f", key=f"mod_o_{SID}")
                comment = st.text_area("Commentaires", _safe_str(row.get("Commentaires","")), key=f"mod_com_{SID}")

                st.markdown("#### 📌 Statut & RFE")
                st_choices = ["Aucun","Envoyé","Accepté","Refusé","Annulé"]
                current = "Aucun"
                if int(_to_num(row.get("Dossier accepté",0)))==1: current="Accepté"
                elif int(_to_num(row.get("Dossier refusé",0)))==1: current="Refusé"
                elif int(_to_num(row.get("Dossier annulé",0)))==1: current="Annulé"
                elif int(_to_num(row.get("Dossier envoyé",0)))==1: current="Envoyé"
                statut = st.selectbox("Statut", st_choices, index=st_choices.index(current), key=f"mod_stat_{SID}")
                rfe_on = st.toggle("RFE", value=(int(_to_num(row.get("RFE",0)))==1), key=f"mod_rfe_{SID}")

                dkey = _status_date_key(statut)
                stat_date = _date_for_widget(row.get(dkey)) if dkey else date.today()
                if statut!="Aucun" and dkey:
                    stat_date = st.date_input(f"Date pour « {statut} »", value=_date_for_widget(row.get(dkey)), key=f"mod_statd_{SID}")

                if st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}"):
                    if not nom or not cat or not sub:
                        st.warning("Nom, Catégorie, Sous-catégorie requis.")
                    else:
                        total = float(honor)+float(other)
                        paye  = float(_to_num(row.get("Payé",0.0)))
                        reste = max(0.0, total - paye)

                        df_live.at[idx,"Nom"]=nom
                        df_live.at[idx,"Date"]=dt
                        df_live.at[idx,"Mois"]=f"{int(mois):02d}"
                        df_live.at[idx,"Categorie"]=cat
                        df_live.at[idx,"Sous-categorie"]=sub
                        df_live.at[idx,"Visa"]=sub
                        df_live.at[idx,"Montant honoraires (US $)"]=float(honor)
                        df_live.at[idx,"Autres frais (US $)"]=float(other)
                        df_live.at[idx,"Total (US $)"]=total
                        df_live.at[idx,"Reste"]=reste
                        df_live.at[idx,"Commentaires"]=comment

                        # reset statuts + dates
                        for k in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé"]:
                            df_live.at[idx,k]=0
                        for k in ["Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation"]:
                            df_live.at[idx,k]=None
                        flags=_status_to_flags(statut)
                        for k,v in flags.items(): df_live.at[idx,k]=v
                        if statut!="Aucun" and dkey:
                            df_live.at[idx,dkey]=stat_date
                        df_live.at[idx,"RFE"]=1 if (rfe_on and statut!="Aucun") else 0

                        write_clients_file(df_live, clients_src if isinstance(clients_src,str) else (save_clients_to or "clients_sauvegarde.xlsx"))
                        st.success("Modifications enregistrées.")
                        st.cache_data.clear()
                        st.rerun()

    # ------- SUPPRIMER -------
    elif op == "Supprimer":
        if df_live.empty:
            st.info("Aucun client.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
            s1,s2=st.columns(2)
            tgt_name = s1.selectbox("Nom", [""]+names, index=0, key=f"del_nom_{SID}")
            tgt_id   = s2.selectbox("ID_Client", [""]+ids, index=0, key=f"del_id_{SID}")

            mask=None
            if tgt_id: mask=(df_live["ID_Client"].astype(str)==tgt_id)
            elif tgt_name: mask=(df_live["Nom"].astype(str)==tgt_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({"Dossier N":row.get("Dossier N",""), "Nom":row.get("Nom",""), "Visa":row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                    df_live = df_live[~mask].copy()
                    write_clients_file(df_live, clients_src if isinstance(clients_src,str) else (save_clients_to or "clients_sauvegarde.xlsx"))
                    st.success("Supprimé.")
                    st.cache_data.clear()
                    st.rerun()

# -------- Export --------
with tabs[6]:
    st.subheader("💾 Export")
    colz1, colz2 = st.columns([1,3])
    with colz1:
        if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID}"):
            try:
                buf = BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    # Clients
                    with BytesIO() as xbuf:
                        with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                            df_all.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                        zf.writestr("Clients.xlsx", xbuf.getvalue())
                    # Visa si présent
                    if visa_src:
                        try:
                            if isinstance(visa_src, str) and os.path.exists(visa_src):
                                zf.write(visa_src, "Visa.xlsx")
                            else:
                                # upload → re-écrit depuis df_visa_raw
                                with BytesIO() as vb:
                                    with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                        df_visa_raw.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                                    zf.writestr("Visa.xlsx", vb.getvalue())
                        except Exception:
                            pass
                st.session_state[f"zip_export_{SID}"] = buf.getvalue()
                st.success("Archive prête.")
            except Exception as e:
                st.error(f"Erreur : {_safe_str(e)}")
    with colz2:
        if st.session_state.get(f"zip_export_{SID}"):
            st.download_button(
                "⬇️ Télécharger l’export (ZIP)",
                data=st.session_state[f"zip_export_{SID}"],
                file_name="Export_Visa_Manager.zip",
                mime="application/zip",
                key=f"zip_dl_{SID}",
            )