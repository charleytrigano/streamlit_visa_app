# =========================
# PARTIE 1/6
# =========================
import json, re, os, zipfile
from io import BytesIO
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import numpy as np
import streamlit as st

# ---------- Constantes colonnes (canonique) ----------
COLS = [
    "ID_Client","Dossier N","Nom","Date","Categorie","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
    "Acompte 1","Acompte 2","RFE","Dossier envoyé","Dossier approuvé","Dossier refusé",
    "Dossier annulé","Commentaires","_Année_","_MoisNum_","Mois"
]

# mapping pour normaliser les noms du fichier utilisateur
NMAP = {
    "Categories":"Categorie",
    "Dossiers envoyé":"Dossier envoyé",
    "Dossier Annulé":"Dossier annulé",
    "Montant honoraires (US$)":"Montant honoraires (US $)",
    "Autres frais (US$)":"Autres frais (US $)",
    "Solde du -":"Solde",
}

# ---------- Petit décor & SID ----------
st.set_page_config(page_title="Visa Manager", layout="wide")
if "_sid" not in st.session_state:
    st.session_state["_sid"] = str(int(datetime.now().timestamp()))
SID = st.session_state["_sid"]

# ---------- Mémoire de chemins ----------
WORK_DIR = Path(".")
LAST_JSON = WORK_DIR / "last_paths.json"

def save_last_paths(clients:str|None, visa:str|None, save_dir:str|None):
    data = {"clients":clients or "", "visa":visa or "", "save_dir":save_dir or ""}
    try:
        LAST_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2))
    except Exception:
        pass

def load_last_paths():
    if LAST_JSON.exists():
        try:
            d = json.loads(LAST_JSON.read_text())
            return d.get("clients",""), d.get("visa",""), d.get("save_dir","")
        except Exception:
            return "","",""
    return "","",""

# ---------- Aides types & formats ----------
def _safe_str(x):
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x)

def _to_num(x):
    try:
        if isinstance(x, (int,float)):
            return float(x)
        s = _safe_str(x)
        s = re.sub(r"[^\d\.,\-]", "", s)
        s = s.replace(",", ".")
        return float(s) if s else 0.0
    except Exception:
        return 0.0

def _safe_num_series(df, col):
    if col not in df.columns:
        return pd.Series([0.0]*len(df), index=df.index, dtype=float)
    return df[col].apply(_to_num).astype(float)

def _fmt_money(v):
    try:
        return f"${float(v):,.2f}".replace(",", " ").replace(".00", ".00")
    except Exception:
        return "$0.00"

def _date_for_widget(val):
    # retourne un objet date « propre » pour st.date_input
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.notna(d):
            return d.date()
    except Exception:
        pass
    return date.today()

def _ensure_time_cols(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    # Mois
    if "Mois" not in out.columns or out["Mois"].isna().all():
        if "Date" in out.columns:
            m = pd.to_datetime(out["Date"], errors="coerce").dt.month.fillna(1).astype(int)
            out["Mois"] = m.apply(lambda x: f"{int(x):02d}")
        else:
            out["Mois"] = "01"
    # _MoisNum_
    out["_MoisNum_"] = pd.to_numeric(out["Mois"], errors="coerce").fillna(1).astype(int)
    # _Année_
    if "_Année_" not in out.columns or out["_Année_"].isna().all():
        if "Date" in out.columns:
            y = pd.to_datetime(out["Date"], errors="coerce").dt.year
            y = y.fillna(date.today().year).astype(int)
            out["_Année_"] = y
        else:
            out["_Année_"] = date.today().year
    return out

def _next_dossier(df, start=13057):
    try:
        existing = pd.to_numeric(df.get("Dossier N", pd.Series([], dtype=str)), errors="coerce").dropna()
        if existing.empty:
            return start
        return int(existing.max()) + 1
    except Exception:
        return start

def _make_client_id(nom, dte:date):
    base = re.sub(r"[^a-z0-9\-]+", "", _safe_str(nom).strip().lower().replace(" ", "-"))
    if not base:
        base = "client"
    return f"{base}-{dte.strftime('%Y%m%d')}"

# ---------- Lecture / écriture ----------
@st.cache_data(show_spinner=False)
def read_any(path: str|Path) -> pd.DataFrame:
    p = str(path)
    if p.lower().endswith(".csv"):
        return pd.read_csv(p)
    return pd.read_excel(p)

@st.cache_data(show_spinner=False)
def read_clients_file(path:str|Path) -> pd.DataFrame:
    df = read_any(path).copy()
    # normalise titres
    cols = [NMAP.get(c, c) for c in df.columns]
    df.columns = cols
    # complète colonnes manquantes
    for c in COLS:
        if c not in df.columns:
            df[c] = pd.NA
    # calcule Total & Solde si absents
    if df["Total (US $)"].isna().all():
        df["Total (US $)"] = _safe_num_series(df, "Montant honoraires (US $)") + _safe_num_series(df, "Autres frais (US $)")
    if df["Solde"].isna().all():
        df["Solde"] = _safe_num_series(df, "Total (US $)") - _safe_num_series(df, "Payé")
        df["Solde"] = df["Solde"].clip(lower=0)
    df = _ensure_time_cols(df)
    return df

def write_clients_file(df:pd.DataFrame, out_path:str|Path):
    if str(out_path).lower().endswith(".csv"):
        df.to_csv(out_path, index=False)
    else:
        with pd.ExcelWriter(out_path, engine="openpyxl") as wr:
            df.to_excel(wr, index=False)

# ---------- État global mémoire chemins ----------
last_clients, last_visa, last_save_dir = load_last_paths()
if "clients_path" not in st.session_state:
    st.session_state["clients_path"] = last_clients
if "visa_path" not in st.session_state:
    st.session_state["visa_path"] = last_visa
if "save_dir" not in st.session_state:
    st.session_state["save_dir"] = last_save_dir

clients_path_curr = st.session_state["clients_path"]
visa_path_curr = st.session_state["visa_path"]
save_dir_curr = st.session_state["save_dir"]

# =========================
# PARTIE 2/6
# =========================
st.sidebar.header("📂 Fichiers")
mode = st.sidebar.radio("Mode de chargement", ["Un fichier (Clients)","Deux fichiers (Clients + Visa)"], key=f"mode_{SID}")

up_clients = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx","csv"], key=f"up_c_{SID}")
up_visa    = st.sidebar.file_uploader("Visa (xlsx/csv)", type=["xlsx","csv"], key=f"up_v_{SID}") if mode=="Deux fichiers (Clients + Visa)" else None

# Dépôt sur disque pour relecture automatique
UPLOAD_DIR = WORK_DIR / "upload_cache"
UPLOAD_DIR.mkdir(exist_ok=True)

def _persist_upload(uploader, default_name):
    if uploader is None:
        return ""
    try:
        path = UPLOAD_DIR / default_name
        with open(path, "wb") as f:
            f.write(uploader.read())
        return str(path)
    except Exception:
        return ""

if up_clients is not None:
    clients_path_curr = _persist_upload(up_clients, "upload_clients.xlsx")
    st.session_state["clients_path"] = clients_path_curr

if mode=="Deux fichiers (Clients + Visa)":
    if up_visa is not None:
        visa_path_curr = _persist_upload(up_visa, "upload_visa.xlsx")
        st.session_state["visa_path"] = visa_path_curr
else:
    # un seul fichier, utilisé comme "clients" et "visa" (si tu veux aussi un aperçu)
    visa_path_curr = st.session_state["visa_path"] or st.session_state["clients_path"]
    st.session_state["visa_path"] = visa_path_curr

# Chemin de sauvegarde (facultatif)
st.sidebar.caption("**Chemin de sauvegarde** (sur ton PC / Drive / OneDrive) :")
save_dir_curr = st.sidebar.text_input("Dossier de sauvegarde", value=save_dir_curr, key=f"save_dir_{SID}")
st.session_state["save_dir"] = save_dir_curr
if st.sidebar.button("💾 Mémoriser ces chemins", key=f"btn_mem_{SID}"):
    save_last_paths(st.session_state["clients_path"], st.session_state["visa_path"], st.session_state["save_dir"])
    st.sidebar.success("Chemins mémorisés.")

# Charger données clients
try:
    df_all = read_clients_file(clients_path_curr) if clients_path_curr else pd.DataFrame()
except Exception as e:
    st.warning("Aucun client chargé. (Charge un fichier Clients)")
    df_all = pd.DataFrame()

# Visa (simple aperçu : on lit juste pour voir catégories/sous-catégories/visa si inclus dans le même fichier)
try:
    df_visa_raw = read_any(visa_path_curr) if visa_path_curr else pd.DataFrame()
except Exception:
    df_visa_raw = pd.DataFrame()

# KPI bandeau si data
st.title("🛂 Visa Manager")
if not df_all.empty:
    k1,k2,k3,k4,k5 = st.columns([1,1,1,1,1])
    total = _safe_num_series(df_all,"Total (US $)").sum()
    paye  = _safe_num_series(df_all,"Payé").sum()
    reste = _safe_num_series(df_all,"Solde").sum() if "Solde" in df_all.columns else max(0.0, total-paye)
    k1.metric("Dossiers", f"{len(df_all)}")
    k2.metric("Honoraires+Frais", _fmt_money(total))
    k3.metric("Payé", _fmt_money(paye))
    k4.metric("Solde", _fmt_money(reste))
    env = int((_safe_num_series(df_all, "Dossier envoyé")>0).sum())
    tot = len(df_all) if len(df_all)>0 else 1
    k5.metric("Envoyés (%)", f"{(env/tot*100):.0f}%")
else:
    st.caption("Charge tes fichiers Clients / Visa dans la barre latérale.")

# Tabs
tabs = st.tabs(["📄 Fichiers chargés","📊 Dashboard","🏦 Escrow","👤 Compte client","🧾 Gestion","📄 Visa (aperçu)","💾 Export","📈 Analyses"])

# =========================
# PARTIE 3/6
# =========================

# --------- Onglet 1 : Fichiers chargés ----------
with tabs[0]:
    st.subheader("📄 Fichiers chargés")
    st.write("**Clients** :", f"`{clients_path_curr}`" if clients_path_curr else "_aucun_")
    st.write("**Visa** :", f"`{visa_path_curr}`" if visa_path_curr else "_aucun_")

# --------- Onglet 2 : Dashboard ----------
with tabs[1]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucun client chargé. Charge un fichier Clients.")
    else:
        dfD = df_all.copy()
        cats = sorted(dfD.get("Categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        subs = sorted(dfD.get("Sous-categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        vis  = sorted(dfD.get("Visa", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

        st.markdown("#### 🎛️ Filtres")
        a1,a2,a3 = st.columns(3)
        fc = a1.multiselect("Catégories", cats, default=[])
        fs = a2.multiselect("Sous-catégories", subs, default=[])
        fv = a3.multiselect("Visa", vis, default=[])

        view = dfD.copy()
        if fc: view = view[view["Categorie"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        # Dossiers par catégorie
        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if "Categorie" in view.columns and not view.empty:
            vc = view["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie","Dossiers"]
            st.bar_chart(vc.set_index("Categorie"))
        else:
            st.info("Aucune catégorie à afficher.")

        # Flux par mois (charts natifs)
        st.markdown("#### 💵 Flux par mois")
        tmp = view.copy()
        tmp["Mois"] = tmp["Mois"].astype(str)
        agg = (tmp.groupby("Mois", as_index=False)
                  .agg({"Montant honoraires (US $)":"sum","Autres frais (US $)":"sum","Payé":"sum"}))
        agg = agg.sort_values("Mois")
        agg["Solde"] = (agg["Montant honoraires (US $)"] + agg["Autres frais (US $)"] - agg["Payé"]).clip(lower=0)
        st.bar_chart(agg.set_index("Mois")[["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]])

        # Détails
        st.markdown("#### 📋 Détails (après filtres)")
        det = view.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).apply(_fmt_money)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in det.columns]
        view_sorted = det.sort_values(by=sort_keys) if sort_keys else det
        st.dataframe(view_sorted[show_cols].reset_index(drop=True), use_container_width=True)

# =========================
# PARTIE 4/6
# =========================

# --------- Onglet 3 : Escrow ----------
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Total (US $)"] = _safe_num_series(dfE,"Total (US $)")
        dfE["Payé"] = _safe_num_series(dfE,"Payé")
        dfE["Solde"] = _safe_num_series(dfE,"Solde")
        t1,t2,t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money(dfE["Total (US $)"].sum()))
        t2.metric("Payé", _fmt_money(dfE["Payé"].sum()))
        t3.metric("Solde", _fmt_money(dfE["Solde"].sum()))
        st.caption("Escrow informatif : honoraires versés avant « Dossier envoyé », à transférer ensuite.")

# --------- Onglet 4 : Compte client ----------
with tabs[3]:
    st.subheader("👤 Compte client (timeline & règlements)")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        names = sorted(df_all["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_all.columns else []
        ids   = sorted(df_all["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_all.columns else []
        c1,c2 = st.columns(2)
        sname = c1.selectbox("Nom", [""]+names, index=0, key=f"acct_nom_{SID}")
        sid   = c2.selectbox("ID_Client", [""]+ids, index=0, key=f"acct_id_{SID}")

        mask = None
        if sid:
            mask = (df_all["ID_Client"].astype(str) == sid)
        elif sname:
            mask = (df_all["Nom"].astype(str) == sname)

        if mask is None or not mask.any():
            st.stop()

        row = df_all[mask].iloc[0].copy()

        # Bandeau montants
        b1,b2,b3,b4 = st.columns(4)
        tot = _safe_num_series(pd.DataFrame([row]),"Total (US $)").iloc[0]
        pay = _safe_num_series(pd.DataFrame([row]),"Payé").iloc[0]
        sol = _safe_num_series(pd.DataFrame([row]),"Solde").iloc[0]
        b1.metric("Total", _fmt_money(tot))
        b2.metric("Payé", _fmt_money(pay))
        b3.metric("Solde", _fmt_money(sol))
        b4.metric("Honoraires", _fmt_money(_safe_num_series(pd.DataFrame([row]),"Montant honoraires (US $)").iloc[0]))

        # Statuts + dates
        st.markdown("#### 🗂️ Statuts")
        s1,s2 = st.columns(2)
        with s1:
            st.write(f"- Dossier envoyé : {int(_to_num(row.get('Dossier envoyé',0)) or 0)}"
                     f" | Date : {_safe_str(row.get('Date d\\'envoi',''))}")
            st.write(f"- Dossier accepté : {int(_to_num(row.get('Dossier approuvé',0)) or 0)}"
                     f" | Date : {_safe_str(row.get('Date d\\'acceptation',''))}")
            st.write(f"- RFE : {int(_to_num(row.get('RFE',0)) or 0)}")
        with s2:
            st.write(f"- Dossier refusé : {int(_to_num(row.get('Dossier refusé',0)) or 0)}"
                     f" | Date : {_safe_str(row.get('Date de refus',''))}")
            st.write(f"- Dossier annulé : {int(_to_num(row.get('Dossier annulé',0)) or 0)}"
                     f" | Date : {_safe_str(row.get('Date d\\'annulation',''))}")
            st.write(f"- Commentaires : {_safe_str(row.get('Commentaires',''))}")

        # Règlement additionnel
        st.markdown("#### ➕ Ajouter un règlement")
        r1,r2 = st.columns(2)
        add_amt = r1.number_input("Montant (US $)", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=f"addpay_{SID}")
        add_dt  = r2.date_input("Date règlement", value=date.today(), key=f"addpdt_{SID}")
        if st.button("Enregistrer le paiement", key=f"btn_addpay_{SID}"):
            # maj df_all et disque
            idx = df_all[mask].index[0]
            p_old = _safe_num_series(pd.DataFrame([row]),"Payé").iloc[0]
            s_old = _safe_num_series(pd.DataFrame([row]),"Solde").iloc[0]
            p_new = float(p_old) + float(add_amt)
            t_val = _safe_num_series(pd.DataFrame([row]),"Total (US $)").iloc[0]
            s_new = max(0.0, float(t_val) - p_new)
            df_all.at[idx,"Payé"] = p_new
            df_all.at[idx,"Solde"] = s_new
            # sauvegarde : réécrit le fichier source si choisi
            if clients_path_curr:
                try:
                    write_clients_file(df_all, clients_path_curr)
                    st.success("Règlement enregistré et fichier mis à jour.")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error("Erreur sauvegarde : " + _safe_str(e))
            else:
                st.warning("Aucun chemin Clients pour sauvegarder.")

# ================================================
# PARTIE 5/6 — 📄 Visa (aperçu)
# ================================================

SID5 = st.session_state.get("_sid", "p5")

def _p5_truthy(v):
    s = str(v).strip().lower()
    return s in {"1", "true", "yes", "y", "x", "oui"}

def _p5_safe_str(x):
    try:
        if x is None:
            return ""
        if isinstance(x, float) and pd.isna(x):
            return ""
        return str(x)
    except Exception:
        return ""

with tabs[5]:
    st.subheader("📄 Visa (aperçu)")

    if df_visa_raw is None or df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        # Détection des colonnes Catégorie / Sous-catégorie
        cat_col = next((c for c in ["Categorie", "Catégorie", "Category"] if c in df_visa_raw.columns), None)
        sub_col = next((c for c in ["Sous-categorie", "Sous-catégorie", "Sous catégorie"] if c in df_visa_raw.columns), None)

        if not cat_col or not sub_col:
            st.warning("Le fichier Visa doit contenir les colonnes 'Catégorie' et 'Sous-catégorie'.")
            st.dataframe(df_visa_raw, use_container_width=True)
            st.stop()

        cats = sorted(df_visa_raw[cat_col].dropna().astype(str).unique().tolist())
        c1, c2 = st.columns(2)
        sel_cat = c1.selectbox("Catégorie", [""] + cats, index=0, key=f"visa_cat_{SID5}")

        subs = []
        if sel_cat:
            subs = sorted(
                df_visa_raw[df_visa_raw[cat_col].astype(str) == sel_cat][sub_col]
                .dropna().astype(str).unique().tolist()
            )

        sel_sub = c2.selectbox("Sous-catégorie", [""] + subs, index=0, key=f"visa_sub_{SID5}")

        # Colonnes d'options
        option_cols = [c for c in df_visa_raw.columns if c not in {cat_col, sub_col, "Visa"}]

        # Filtrage
        flt = df_visa_raw.copy()
        if sel_cat:
            flt = flt[flt[cat_col].astype(str) == sel_cat]
        if sel_sub:
            flt = flt[flt[sub_col].astype(str) == sel_sub]

        st.markdown("#### Tableau Visa filtré")
        st.dataframe(flt.reset_index(drop=True), use_container_width=True, height=300, key=f"visa_filtered_{SID5}")

        # Détection des options cochées
        st.markdown("#### Options disponibles pour cette sous-catégorie")
        if not sel_cat or not sel_sub:
            st.caption("Choisis une catégorie et une sous-catégorie pour voir les options.")
        else:
            available = set()
            for _, row in flt.iterrows():
                for col in option_cols:
                    if _p5_truthy(row.get(col, "")):
                        available.add(col)

            if not available:
                st.info("Aucune option cochée détectée.")
            else:
                cols = st.columns(min(len(available), 4) or 1)
                chosen = []
                for i, opt in enumerate(sorted(available)):
                    with cols[i % len(cols)]:
                        ck = st.checkbox(opt, key=f"visa_opt_{SID5}_{i}")
                        if ck:
                            chosen.append(opt)

                visa_label = sel_sub
                if chosen:
                    visa_label += " - " + ", ".join(chosen)
                st.success(f"Visa proposé : {visa_label}")

        # Statut du dossier
        st.markdown("#### Statut du dossier")
        s1, s2 = st.columns(2)
        with s1:
            try:
                row0 = flt.iloc[0] if not flt.empty else {}
                envoye = int(row0.get("Dossier envoyé", 0) or 0)
                approuve = int(row0.get("Dossier approuvé", 0) or 0)
                refuse = int(row0.get("Dossier refusé", 0) or 0)
                annule = int(row0.get("Dossier annulé", 0) or 0)
                rfe = int(row0.get("RFE", 0) or 0)
                date_envoi = _p5_safe_str(row0.get("Date d'envoi", ""))

                txt = (
                    f"Dossier envoyé : {envoye} | "
                    f"Approuvé : {approuve} | "
                    f"Refusé : {refuse} | "
                    f"Annulé : {annule} | "
                    f"RFE : {rfe}"
                )
                if date_envoi:
                    txt += f" | Date d'envoi : {date_envoi}"

                s1.write(txt)
            except Exception as e:
                s1.error(f"Erreur lecture statut : {_p5_safe_str(e)}")

# =======================================================
# PARTIE 6/6 - Analyses (sans Plotly) & Export
# =======================================================
SID6 = st.session_state.get("_sid", "p6")

def _ensure_time_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Ajoute _Année_, _MoisNum_ et Mois si absents (à partir de Date/Mois)."""
    out = df.copy()
    if "Mois" not in out.columns:
        if "Date" in out.columns:
            try:
                m = pd.to_datetime(out["Date"], errors="coerce").dt.month
                out["Mois"] = m.fillna(1).astype(int).apply(lambda x: f"{int(x):02d}")
            except Exception:
                out["Mois"] = "01"
        else:
            out["Mois"] = "01"
    try:
        out["_MoisNum_"] = pd.to_numeric(out["Mois"], errors="coerce").fillna(1).astype(int)
    except Exception:
        out["_MoisNum_"] = 1
    if "_Année_" not in out.columns:
        if "Date" in out.columns:
            try:
                out["_Année_"] = pd.to_datetime(out["Date"], errors="coerce").dt.year
                if out["_Année_"].isna().all():
                    out["_Année_"] = date.today().year
                else:
                    mode_y = out["_Année_"].mode()
                    out["_Année_"] = out["_Année_"].fillna(mode_y.iloc[0] if not mode_y.empty else date.today().year).astype(int)
            except Exception:
                out["_Année_"] = date.today().year
        else:
            out["_Année_"] = date.today().year
    return out

def _pct(a, b):
    a = float(a or 0); b = float(b or 0)
    return (a / b * 100.0) if b > 0 else 0.0


# -----------------------------
# Onglet Analyses (tabs[7])
# -----------------------------
with tabs[7]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucun client chargé.")
    else:
        dfA0 = _ensure_time_cols(df_all)

        yearsA  = sorted(pd.to_numeric(dfA0["_Année_"], errors="coerce").dropna().astype(int).unique().tolist())
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(dfA0.get("Categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        subsA   = sorted(dfA0.get("Sous-categorie", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        visasA  = sorted(dfA0.get("Visa", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

        st.markdown("#### 🎛️ Filtres (ensemble global)")
        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_years_{SID6}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"a_months_{SID6}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"a_cats_{SID6}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"a_subs_{SID6}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_visas_{SID6}")

        dfA = dfA0.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        hono = _safe_num_series(dfA, "Montant honoraires (US $)")
        autre = _safe_num_series(dfA, "Autres frais (US $)")
        total = (_safe_num_series(dfA, "Total (US $)") if "Total (US $)" in dfA.columns else (hono + autre))
        paye  = _safe_num_series(dfA, "Payé")
        reste = (_safe_num_series(dfA, "Solde") if "Solde" in dfA.columns else (total - paye)).clip(lower=0)

        # KPI compacts
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires+Frais", _fmt_money(float((hono+autre).sum())))
        k3.metric("Payé", _fmt_money(float(paye.sum())))
        k4.metric("Solde", _fmt_money(float(reste.sum())))
        pct_env = _pct(dfA.get("Dossier envoyé", 0).sum(), len(dfA))
        k5.metric("Envoyés (%)", f"{pct_env:.0f}%")

        # Répartition par Catégorie / Sous-catégorie
        st.markdown("#### 📊 Répartition (nombre de dossiers)")
        c11, c12 = st.columns(2)
        if not dfA.empty:
            df_cnt_cat = (dfA.groupby("Categorie", as_index=False)
                            .size().rename(columns={"size":"Dossiers"})).sort_values("Dossiers", ascending=False)
            df_cnt_cat["%"] = (df_cnt_cat["Dossiers"] / max(1, df_cnt_cat["Dossiers"].sum()) * 100).round(1)
            c11.dataframe(df_cnt_cat, use_container_width=True, height=240, key=f"a_cnt_cat_{SID6}")

            if "Sous-categorie" in dfA.columns:
                df_cnt_sub = (dfA.groupby("Sous-categorie", as_index=False)
                                .size().rename(columns={"size":"Dossiers"})).sort_values("Dossiers", ascending=False)
                df_cnt_sub["%"] = (df_cnt_sub["Dossiers"] / max(1, df_cnt_sub["Dossiers"].sum()) * 100).round(1)
                c12.dataframe(df_cnt_sub, use_container_width=True, height=240, key=f"a_cnt_sub_{SID6}")
            else:
                c12.info("Aucune sous-catégorie dans les données.")

        # Flux par mois (bar_chart natif Streamlit)
        st.markdown("#### 💵 Flux par mois")
        tmp = dfA.copy()
        tmp["Mois"] = tmp["Mois"].astype(str)
        flux = (tmp.groupby("Mois", as_index=False)
                    .agg({
                        "Montant honoraires (US $)": "sum",
                        "Autres frais (US $)": "sum",
                        "Payé": "sum"
                    }))
        flux = flux.sort_values("Mois")
        flux["Solde"] = (flux["Montant honoraires (US $)"] + flux["Autres frais (US $)"] - flux["Payé"]).clip(lower=0)
        st.bar_chart(flux.set_index("Mois")[["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]])

        # Comparaison A vs B
        st.markdown("#### ⚖️ Comparaison A vs B (périodes / filtres)")
        ca1, ca2, ca3 = st.columns(3)
        ya = ca1.multiselect("Année (A)", yearsA, default=[], key=f"cmp_ya_{SID6}")
        ma = ca2.multiselect("Mois (A)", monthsA, default=[], key=f"cmp_ma_{SID6}")
        ca = ca3.multiselect("Catégories (A)", catsA, default=[], key=f"cmp_ca_{SID6}")

        cb1, cb2, cb3 = st.columns(3)
        yb = cb1.multiselect("Année (B)", yearsA, default=[], key=f"cmp_yb_{SID6}")
        mb = cb2.multiselect("Mois (B)", monthsA, default=[], key=f"cmp_mb_{SID6}")
        cb = cb3.multiselect("Catégories (B)", catsA, default=[], key=f"cmp_cb_{SID6}")

        def _apply_filters(df, yy, mm, cc):
            d = df.copy()
            if yy: d = d[d["_Année_"].isin(yy)]
            if mm: d = d[d["Mois"].astype(str).isin(mm)]
            if cc: d = d[d["Categorie"].astype(str).isin(cc)]
            return d

        A = _apply_filters(dfA0, ya, ma, ca)
        B = _apply_filters(dfA0, yb, mb, cb)

        def _kpis(df):
            h = _safe_num_series(df, "Montant honoraires (US $)")
            a = _safe_num_series(df, "Autres frais (US $)")
            t = (h + a)
            p = _safe_num_series(df, "Payé")
            r = (t - p).clip(lower=0)
            return {
                "Dossiers": len(df),
                "Honoraires+Frais": float(t.sum()),
                "Payé": float(p.sum()),
                "Solde": float(r.sum())
            }

        kA = _kpis(A); kB = _kpis(B)

        cA, cB = st.columns(2)
        with cA:
            st.markdown("**Période A**")
            st.metric("Dossiers", f"{kA['Dossiers']}")
            st.metric("Honoraires+Frais", _fmt_money(kA["Honoraires+Frais"]))
            st.metric("Payé", _fmt_money(kA["Payé"]))
            st.metric("Solde", _fmt_money(kA["Solde"]))
        with cB:
            st.markdown("**Période B**")
            st.metric("Dossiers", f"{kB['Dossiers']}")
            st.metric("Honoraires+Frais", _fmt_money(kB["Honoraires+Frais"]))
            st.metric("Payé", _fmt_money(kB["Payé"]))
            st.metric("Solde", _fmt_money(kB["Solde"]))

        # Détails
        st.markdown("#### 📋 Détails (après filtres globaux)")
        det = dfA.copy()
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Total (US $)", "Payé", "Solde"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).apply(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","RFE","Commentaires"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det

        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"a_detail_{SID6}")


# -----------------------------
# Onglet Export (tabs[6])
# -----------------------------
with tabs[6]:
    st.subheader("💾 Export")

    colz1, colz2 = st.columns([1,3])
    with colz1:
        if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID6}"):
            try:
                buf = BytesIO()
                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    # Export Clients propre
                    try:
                        df_export = read_clients_file(clients_path_curr)
                        with BytesIO() as xbuf:
                            with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                                df_export.to_excel(wr, sheet_name="Clients", index=False)
                            zf.writestr("Clients.xlsx", xbuf.getvalue())
                    except Exception as e:
                        st.warning(f"Clients : export partiel ({_safe_str(e)})")

                    # Export Visa (tel quel si possible)
                    try:
                        zf.write(visa_path_curr, "Visa.xlsx")
                    except Exception:
                        try:
                            dfv0 = pd.read_excel(visa_path_curr)
                            with BytesIO() as vb:
                                with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                    dfv0.to_excel(wr, sheet_name="Visa", index=False)
                                zf.writestr("Visa.xlsx", vb.getvalue())
                        except Exception as e2:
                            st.warning(f"Visa : export partiel ({_safe_str(e2)})")

                st.session_state[f"zip_export_{SID6}"] = buf.getvalue()
                st.success("Archive prête.")
            except Exception as e:
                st.error("Erreur de préparation : " + _safe_str(e))

    with colz2:
        if st.session_state.get(f"zip_export_{SID6}"):
            st.download_button(
                label="⬇️ Télécharger l’export (ZIP)",
                data=st.session_state[f"zip_export_{SID6}"],
                file_name="Export_Visa_Manager.zip",
                mime="application/zip",
                key=f"zip_dl_{SID6}",
            )
        else:
            st.caption("Clique sur « Préparer l’archive ZIP » pour générer un export complet.")