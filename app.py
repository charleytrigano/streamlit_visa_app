# =========================
# Visa Manager — Version Complète et Corrigée
# =========================

import os
import json
import re
import io
import zipfile
import uuid
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional
from pathlib import Path

import pandas as pd
import streamlit as st

from datetime import date, datetime
import pandas as pd

def _date_for_widget(val):
    """Convertit proprement une valeur Excel/pandas en date utilisable dans Streamlit."""
    if val is None or pd.isna(val):
        return date.today()
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return date.today()
        return d.date()
    except Exception:
        return date.today()

# =========================
# Constantes et configuration
# =========================
APP_TITLE = "🛂 Visa Manager"

COLS_CLIENTS = [
    "ID_Client", "Dossier N", "Nom", "Date",
    "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)",
    "Payé", "Solde", "Acompte 1", "Acompte 2",
    "RFE", "Dossiers envoyé", "Dossier approuvé",
    "Dossier refusé", "Dossier Annulé", "Commentaires"
]

MEMO_FILE = "_vmemory.json"
SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"
SID = "vmgr"

# =========================
# Fonctions utilitaires
# =========================

def _safe_str(x: Any) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _to_num(x: Any) -> float:
    if isinstance(x, (int, float)):
        return float(x)
    s = _safe_str(x)
    if not s:
        return 0.0
    s = re.sub(r"[^\d,.-]", "", s)
    s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0

def _fmt_money(v: float) -> str:
    try:
        return "${:,.2f}".format(float(v))
    except Exception:
        return "$0.00"

# Vérifier si row est défini et si "Date" est bien dans row
if row is not None and "Date" in row:
    dval = _date_for_widget(row.get("Date"))
else:
    dval = date.today()  # Définit une valeur par défaut (aujourd'hui)

if isinstance(dval, datetime):
    dval = dval.date()

dt = d2.date_input("Date de création", value=dval, key=skey("mod", "date"))

if dval is None or isinstance(dval, pd._libs.tslibs.nattype.NaTType):
    dval = date.today()

dt = d2.date_input("Date de création", value=dval, key=skey("mod", "date"))


dval = _date_for_widget(row.get("Date"))
if dval is None:
    dval = date.today()

if isinstance(dval, datetime):
    dval = dval.date()

dt = d2.date_input("Date de création", value=dval, key=skey("mod", "date"))


def _normalize_clients_numeric(df: pd.DataFrame) -> pd.DataFrame:
    num_cols = ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Acompte 1", "Acompte 2"]
    for c in num_cols:
        if c in df.columns:
            df[c] = df[c].apply(_to_num)
    if "Montant honoraires (US $)" in df.columns and "Autres frais (US $)" in df.columns:
        total = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]
        paye = df["Payé"] if "Payé" in df.columns else 0.0
        df["Solde"] = (total - paye).clip(lower=0.0)
    return df

def _normalize_status(df: pd.DataFrame) -> pd.DataFrame:
    for c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: 1 if str(x).strip() in ["1", "True", "true", "OUI", "Oui", "oui", "X", "x"] else 0)
        else:
            df[c] = 0
    return df

def normalize_clients(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS_CLIENTS)
    df = df.copy()
    ren = {
        "Categorie": "Categories",
        "Catégorie": "Categories",
        "Sous-categorie": "Sous-categorie",
        "Sous-catégorie": "Sous-categorie",
        "Payee": "Payé",
        "Payé (US $)": "Payé",
        "Montant honoraires": "Montant honoraires (US $)",
        "Autres frais": "Autres frais (US $)",
        "Dossier envoye": "Dossiers envoyé",
        "Dossier envoyé": "Dossiers envoyé",
    }
    df.rename(columns={k: v for k, v in ren.items() if k in df.columns}, inplace=True)
    df = _ensure_columns(df, COLS_CLIENTS)
    if "Date" in df.columns:
        try:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        except Exception:
            pass
    df = _normalize_clients_numeric(df)
    df = _normalize_status(df)
    df["Nom"] = df["Nom"].astype(str)
    df["Categories"] = df["Categories"].astype(str)
    df["Sous-categorie"] = df["Sous-categorie"].astype(str)
    df["Visa"] = df["Visa"].astype(str)
    df["Commentaires"] = df["Commentaires"].astype(str)
    try:
        dser = pd.to_datetime(df["Date"], errors="coerce")
        df["_Annee_"] = dser.dt.year.fillna(0).astype(int)
        df["_MoisNum_"] = dser.dt.month.fillna(0).astype(int)
        df["Mois"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if m and m == m else "")
    except Exception:
        df["_Annee_"] = 0
        df["_MoisNum_"] = 0
        df["Mois"] = ""
    return df

def read_any_table(src: Any, sheet: Optional[str] = None) -> Optional[pd.DataFrame]:
    if src is None:
        return None
    if hasattr(src, "read") and hasattr(src, "name"):
        name = src.name.lower()
        data = src.read()
        bio = BytesIO(data)
        if name.endswith(".csv"):
            return pd.read_csv(bio)
        return pd.read_excel(bio, sheet_name=(sheet if sheet else 0))
    if isinstance(src, (str, os.PathLike)):
        p = str(src)
        if not os.path.exists(p):
            return None
        if p.lower().endswith(".csv"):
            return pd.read_csv(p)
        return pd.read_excel(p, sheet_name=(sheet if sheet else 0))
    if isinstance(src, (io.BytesIO, BytesIO)):
        try:
            bio2 = BytesIO(src.getvalue())
            return pd.read_excel(bio2, sheet_name=(sheet if sheet else 0))
        except Exception:
            src.seek(0)
            return pd.read_csv(src)
    return None

def read_sheet_from_path(path: str, sheet_name: str) -> Optional[pd.DataFrame]:
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return None

def load_last_paths() -> Tuple[str, str, str]:
    if not os.path.exists(MEMO_FILE):
        return "", "", ""
    try:
        with open(MEMO_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("clients", ""), data.get("visa", ""), data.get("save_dir", "")
    except Exception:
        return "", "", ""

def save_last_paths(clients_path: str, visa_path: str, save_dir: str) -> None:
    data = {"clients": clients_path or "", "visa": visa_path or "", "save_dir": save_dir or ""}
    try:
        with open(MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def skey(*parts: str) -> str:
    return f"{SID}_" + "_".join([p for p in parts if p])

def _norm(s: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]", "_", s.strip().lower())

def make_client_id(nom: str, dval: date) -> str:
    return f"{_norm(nom)}_{int(datetime.now().timestamp())}"

def next_dossier(df: pd.DataFrame) -> int:
    max_dossier = df.get("Dossier N", pd.Series([13056])).astype(str).str.extract(r"(\d+)").fillna(13056).astype(int).max()
    return max_dossier + 1

def _to_float(x: Any) -> float:
    return _to_num(x)

def _ensure_dir(pdir: Path) -> None:
    pdir.mkdir(parents=True, exist_ok=True)

def build_visa_map(dfv: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    vm: Dict[str, Dict[str, Dict[str, Any]]] = {}
    if dfv is None or dfv.empty:
        return vm
    cols = [c for c in dfv.columns if _safe_str(c)]
    if "Categories" not in cols and "Catégorie" in cols:
        dfv = dfv.rename(columns={"Catégorie": "Categories"})
    if "Sous-categorie" not in cols and "Sous-catégorie" in cols:
        dfv = dfv.rename(columns={"Sous-catégorie": "Sous-categorie"})
    if "Categories" not in dfv.columns or "Sous-categorie" not in dfv.columns:
        return vm
    fixed = ["Categories", "Sous-categorie"]
    option_cols = [c for c in dfv.columns if c not in fixed]
    for _, row in dfv.iterrows():
        cat = _safe_str(row.get("Categories", "")).strip()
        sub = _safe_str(row.get("Sous-categorie", "")).strip()
        if not cat or not sub:
            continue
        vm.setdefault(cat, {})
        vm[cat].setdefault(sub, {"exclusive": None, "options": []})
        opts = []
        for oc in option_cols:
            val = _safe_str(row.get(oc, "")).strip()
            if val in ["1", "x", "X", "oui", "Oui", "OUI", "True", "true"]:
                opts.append(oc)
        exclusive = None
        if set([o.upper() for o in opts]) == set(["COS", "EOS"]):
            exclusive = "radio_group"
        vm[cat][sub] = {"exclusive": exclusive, "options": opts}
    return vm

def visa_option_selector(vm: Dict[str, Any], cat: str, sub: str, keybase: str) -> str:
    if cat not in vm or sub not in vm[cat]:
        return sub
    meta = vm[cat][sub]
    opts = meta.get("options", [])
    if not opts:
        return sub
    if meta.get("exclusive") == "radio_group" and set([o.upper() for o in opts]) == set(["COS", "EOS"]):
        pick = st.radio("Options", ["COS", "EOS"], horizontal=True, key=skey(keybase, "opt"))
        return f"{sub} {pick}"
    else:
        picks = st.multiselect("Options", opts, default=[], key=skey(keybase, "opts"))
        if not picks:
            return sub
        return f"{sub} {picks[0]}"

def _ensure_time_features(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()
    if "Date" in df.columns:
        try:
            dd = pd.to_datetime(df["Date"], errors="coerce")
        except Exception:
            dd = pd.to_datetime(pd.Series([], dtype="datetime64[ns]"))
        df["_Année_"] = dd.dt.year
        df["_MoisNum_"] = dd.dt.month
        df["Mois"] = dd.dt.month.apply(lambda m: f"{int(m):02d}" if pd.notna(m) else "")
    else:
        if "_Année_" not in df.columns:
            df["_Année_"] = pd.NA
        if "_MoisNum_" not in df.columns:
            df["_MoisNum_"] = pd.NA
        if "Mois" not in df.columns:
            df["Mois"] = ""
    return df

# =========================
# Interface Streamlit
# =========================

st.set_page_config(page_title="Visa Manager", layout="wide")
st.title(APP_TITLE)

# Barre latérale
st.sidebar.header("📂 Fichiers")
last_clients, last_visa, last_save_dir = load_last_paths()

mode = st.sidebar.radio(
    "Mode de chargement",
    ["Un fichier (Clients)", "Deux fichiers (Clients & Visa)"],
    index=0,
    key=skey("mode")
)

up_clients = st.sidebar.file_uploader(
    "Clients (xlsx/csv)", type=["xlsx", "xls", "csv"], key=skey("up_clients")
)
up_visa = None
if mode == "Deux fichiers (Clients & Visa)":
    up_visa = st.sidebar.file_uploader(
        "Visa (xlsx/csv)", type=["xlsx", "xls", "csv"], key=skey("up_visa")
    )

clients_path_in = st.sidebar.text_input("ou chemin local Clients", value=last_clients, key=skey("cli_path"))
visa_path_in = st.sidebar.text_input("ou chemin local Visa", value=(last_visa if mode != "Un fichier (Clients)" else ""), key=skey("vis_path"))
save_dir_in = st.sidebar.text_input("Dossier de sauvegarde", value=last_save_dir, key=skey("save_dir"))

if st.sidebar.button("📥 Charger", key=skey("btn_load")):
    save_last_paths(clients_path_in, visa_path_in, save_dir_in)
    st.success("Chemins mémorisés. Re-lancement pour prise en compte.")
    st.rerun()

# Lecture des fichiers
clients_src = up_clients if up_clients is not None else (clients_path_in if clients_path_in else last_clients)
df_clients_raw = normalize_clients(read_any_table(clients_src))

if mode == "Deux fichiers (Clients & Visa)":
    visa_src = up_visa if up_visa is not None else (visa_path_in if visa_path_in else last_visa)
else:
    visa_src = up_clients if up_clients is not None else (clients_path_in if clients_path_in else last_clients)

# Correction de la lecture du fichier Visa
df_visa_raw = read_any_table(visa_src, sheet=SHEET_VISA)
if df_visa_raw is None:
    df_visa_raw = read_any_table(visa_src)
if df_visa_raw is None:
    df_visa_raw = pd.DataFrame()

# Affichage des fichiers chargés
with st.expander("📄 Fichiers chargés", expanded=True):
    st.write("**Clients** :", ("(aucun)" if (df_clients_raw is None or df_clients_raw.empty) else (getattr(clients_src, 'name', str(clients_src)))))
    st.write("**Visa** :", ("(aucun)" if (df_visa_raw is None or df_visa_raw.empty) else (getattr(visa_src, 'name', str(visa_src)))))

# Construction de la carte Visa
visa_map = build_visa_map(df_visa_raw)

df_all = _ensure_time_features(df_clients_raw)

# Création des onglets
tabs = st.tabs([
    "📄 Fichiers",
    "📊 Dashboard",
    "📈 Analyses",
    "🏦 Escrow",
    "👤 Compte client",
    "🧾 Gestion",
    "📄 Visa (aperçu)",
    "💾 Export",
])

# Onglet Dashboard
with tabs[1]:
    st.subheader("📊 Dashboard")
    if df_all is None or df_all.empty:
        st.info("Aucun client chargé. Chargez les fichiers dans la barre latérale.")
    else:
        cats = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subs = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []
        years = sorted(pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().astype(int).unique().tolist())

        a1, a2, a3, a4 = st.columns([1, 1, 1, 1])
        fc = a1.multiselect("Catégories", cats, default=[], key=skey("dash", "cats"))
        fs = a2.multiselect("Sous-catégories", subs, default=[], key=skey("dash", "subs"))
        fv = a3.multiselect("Visa", visas, default=[], key=skey("dash", "visas"))
        fy = a4.multiselect("Année", years, default=[], key=skey("dash", "years"))

        view = df_all.copy()
        if fc:
            view = view[view["Categories"].astype(str).isin(fc)]
        if fs:
            view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv:
            view = view[view["Visa"].astype(str).isin(fv)]
        if fy:
            view = view[view["_Année_"].isin(fy)]

        k1, k2, k3, k4, k5 = st.columns([1, 1, 1, 1, 1])
        k1.metric("Dossiers", f"{len(view)}")
        total = (view["Montant honoraires (US $)"].apply(_to_num) + view["Autres frais (US $)"].apply(_to_num)).sum()
        paye = view["Payé"].apply(_to_num).sum()
        solde = view["Solde"].apply(_to_num).sum()



        env_pct = 0
        if "Dossiers envoyé" in view.columns and len(view) > 0:
            env_pct = int(100 * (view["Dossiers envoyé"].apply(_to_num).clip(lower=0, upper=1).sum() / len(view)))
        k2.metric("Honoraires+Frais", _fmt_money(total))
        k3.metric("Payé", _fmt_money(paye))
        k4.metric("Solde", _fmt_money(solde))
        k5.metric("Envoyés (%)", f"{env_pct}%")

        if not view.empty and "Categories" in view.columns:
            vc = view["Categories"].value_counts().rename_axis("Categorie").reset_index(name="Nombre")
            st.bar_chart(vc.set_index("Categorie"))

        if not view.empty and "Mois" in view.columns:
            tmp = view.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            g = tmp.groupby("Mois", as_index=False).agg({
                "Montant honoraires (US $)": "sum",
                "Autres frais (US $)": "sum",
                "Payé": "sum",
                "Solde": "sum",
            }).sort_values("Mois")
            g = g.fillna(0)
            g = g.set_index("Mois")
            st.bar_chart(g)

        show_cols = [c for c in [
            "Dossier N", "ID_Client", "Nom", "Date", "Mois", "Categories", "Sous-categorie", "Visa",
            "Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Commentaires",
            "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé", "RFE"
        ] if c in view.columns]

        detail = view.copy()
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde"]:
            if c in detail.columns:
                detail[c] = detail[c].apply(_to_num).map(_fmt_money)
        if "Date" in detail.columns:
            try:
                detail["Date"] = pd.to_datetime(detail["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                detail["Date"] = detail["Date"].astype(str)

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categories", "Nom"] if c in detail.columns]
        detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail
        st.dataframe(detail_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=skey("dash", "table"))

# Onglet Analyses
with tabs[2]:
    st.subheader("📈 Analyses")
    if df_all is None or df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA = sorted(pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().astype(int).unique().tolist())
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subsA = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        b1, b2, b3, b4, b5 = st.columns(5)
        fy = b1.multiselect("Année", yearsA, default=[], key=skey("an", "years"))
        fm = b2.multiselect("Mois (MM)", monthsA, default=[], key=skey("an", "months"))
        fc = b3.multiselect("Catégories", catsA, default=[], key=skey("an", "cats"))
        fs = b4.multiselect("Sous-catégories", subsA, default=[], key=skey("an", "subs"))
        fv = b5.multiselect("Visa", visasA, default=[], key=skey("an", "visas"))

        dfA = df_all.copy()
        if fy:
            dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm:
            dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc:
            dfA = dfA[dfA["Categories"].astype(str).isin(fc)]
        if fs:
            dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv:
            dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Dossiers", f"{len(dfA)}")
        c2.metric("Honoraires", _fmt_money(dfA["Montant honoraires (US $)"].apply(_to_num).sum()))
        c3.metric("Payé", _fmt_money(dfA["Payé"].apply(_to_num).sum()))
        c4.metric("Solde", _fmt_money(dfA["Solde"].apply(_to_num).sum()))


        if not dfA.empty and "Categories" in dfA.columns:
            total_cnt = max(1, len(dfA))
            rep = dfA["Categories"].value_counts().rename_axis("Categorie").reset_index(name="Nbr")
            rep["%"] = (rep["Nbr"] / total_cnt * 100).round(1)
            st.dataframe(rep, use_container_width=True, hide_index=True, key=skey("an", "rep_cat"))

        if not dfA.empty and "Sous-categorie" in dfA.columns:
            total_cnt = max(1, len(dfA))
            rep2 = dfA["Sous-categorie"].value_counts().rename_axis("Sous-categorie").reset_index(name="Nbr")
            rep2["%"] = (rep2["Nbr"] / total_cnt * 100).round(1)
            st.dataframe(rep2, use_container_width=True, hide_index=True, key=skey("an", "rep_sub"))

        ca1, ca2, ca3 = st.columns(3)
        pa_years = ca1.multiselect("Années (A)", yearsA, default=[], key=skey("cmp", "ya"))
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=skey("cmp", "ma"))
        pa_cat = ca3.multiselect("Catégories (A)", catsA, default=[], key=skey("cmp", "ca"))

        cb1, cb2, cb3 = st.columns(3)
        pb_years = cb1.multiselect("Années (B)", yearsA, default=[], key=skey("cmp", "yb"))
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=skey("cmp", "mb"))
        pb_cat = cb3.multiselect("Catégories (B)", catsA, default=[], key=skey("cmp", "cb"))

        def _slice(df, ys, ms, cs):
            s = df.copy()
            if ys:
                s = s[s["_Année_"].isin(ys)]
            if ms:
                s = s[s["Mois"].astype(str).isin(ms)]
            if cs:
                s = s[s["Categories"].astype(str).isin(cs)]
            return s

        A = _slice(df_all, pa_years, pa_month, pa_cat)
        B = _slice(df_all, pb_years, pb_month, pb_cat)

        def _kpis(df):
            return {
                "Dossiers": len(df),
                "Honoraires": df["Montant honoraires (US $)"].apply(_to_num).sum(),
                "Payé": df["Payé"].apply(_to_num).sum(),
                "Solde": df["Solde"].apply(_to_num).sum(),
            }


        kA, kB = _kpis(A), _kpis(B)
        dcmp = pd.DataFrame({
            "KPI": ["Dossiers", "Honoraires", "Payé", "Solde"],
            "A": [kA["Dossiers"], kA["Honoraires"], kA["Payé"], kA["Solde"]],
            "B": [kB["Dossiers"], kB["Honoraires"], kB["Payé"], kB["Solde"]],
            "Δ (B - A)": [kB["Dossiers"] - kA["Dossiers"], kB["Honoraires"] - kA["Honoraires"], kB["Payé"] - kA["Payé"], kB["Solde"] - kA["Solde"]],
        })
        for c in ["A", "B", "Δ (B - A)"]:
            dcmp.loc[1:3, c] = dcmp.loc[1:3, c].astype(float).map(_fmt_money)
        st.dataframe(dcmp, use_container_width=True, hide_index=True, key=skey("an", "cmp_table"))

# Onglet Escrow
with tabs[3]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all is None or df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Total (US $)"] = _to_num(dfE.get("Montant honoraires (US $)", 0)) + _to_num(dfE.get("Autres frais (US $)", 0))
        dfE["Payé"] = _to_num(dfE.get("Payé", 0))
        dfE["Solde"] = _to_num(dfE.get("Solde", 0))

        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money(float(dfE["Total (US $)"].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Solde", _fmt_money(float(dfE["Solde"].sum())))

        agg = dfE.groupby("Categories", as_index=False)[["Total (US $)", "Payé", "Solde"]].sum().sort_values("Total (US $)", ascending=False)
        st.dataframe(agg, use_container_width=True, key=skey("esc", "agg"))

# Onglet Compte client
with tabs[4]:
    st.subheader("👤 Compte client")
    if df_all is None or df_all.empty:
        st.info("Aucun client chargé.")
    else:
        left, right = st.columns(2)
        ids = sorted(df_all["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_all.columns else []
        noms = sorted(df_all["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_all.columns else []

        sel_id = left.selectbox("ID_Client", [""] + ids, index=0, key=skey("acct", "id"))
        sel_nom = right.selectbox("Nom", [""] + noms, index=0, key=skey("acct", "nm"))

        subset = df_all.copy()
        if sel_id:
            subset = subset[subset["ID_Client"].astype(str) == sel_id]
        elif sel_nom:
            subset = subset[subset["Nom"].astype(str) == sel_nom]

        if subset.empty:
            st.warning("Sélectionnez un client pour afficher le compte.")
        else:
            row = subset.iloc[0].to_dict()

            r1, r2, r3, r4 = st.columns(4)
            r1.metric("Dossier N", _safe_str(row.get("Dossier N", "")))
            total = float(_to_num(row.get("Montant honoraires (US $)", 0)) + _to_num(row.get("Autres frais (US $)", 0)))
            r2.metric("Total", _fmt_money(total))
            r3.metric("Payé", _fmt_money(float(_to_num(row.get("Payé", 0)))))
            r4.metric("Solde", _fmt_money(float(_to_num(row.get("Solde", 0)))))

            d1, d2, d3 = st.columns(3)
            d1.write(f"**Catégorie :** {_safe_str(row.get('Categories', ''))}")
            d1.write(f"**Sous-catégorie :** {_safe_str(row.get('Sous-categorie', ''))}")
            d1.write(f"**Visa :** {_safe_str(row.get('Visa', ''))}")
            d2.write(f"**Date :** {_safe_str(row.get('Date', ''))}")
            d2.write(f"**Mois (MM) :** {_safe_str(row.get('Mois', ''))}")
            d3.write(f"**Commentaires :** {_safe_str(row.get('Commentaires', ''))}")

            s1, s2 = st.columns(2)

            def sdate(label):
                val = row.get(label, "")
                if isinstance(val, (date, datetime)):
                    return val.strftime("%Y-%m-%d")
                try:
                    d = pd.to_datetime(val, errors="coerce")
                    return d.date().strftime("%Y-%m-%d") if pd.notna(d) else ""
                except Exception:
                    return _safe_str(val)

            s1.write(f"- **Dossier envoyé** : {'Oui' if sdate('Date denvoi') else 'Non'} | Date : {sdate('Date denvoi')}")
            s1.write(f"- **Dossier approuvé** : {'Oui' if sdate('Date dacceptation') else 'Non'} | Date : {sdate('Date dacceptation')}")
            s2.write(f"- **Dossier refusé** : {'Oui' if sdate('Date de refus') else 'Non'} | Date : {sdate('Date de refus')}")
            s2.write(f"- **Dossier annulé** : {'Oui' if sdate('Date dannulation') else 'Non'} | Date : {sdate('Date dannulation')}")
            rfeflag = int(_to_num(row.get("RFE", 0)) or 0)
            st.write(f"- **RFE** : {'Oui' if rfeflag else 'Non'}")

            mvts = []
            if "Acompte 1" in row and _to_num(row["Acompte 1"]) > 0:
                mvts.append({"Libellé": "Acompte 1", "Montant": float(_to_num(row["Acompte 1"]))})
            if "Acompte 2" in row and _to_num(row["Acompte 2"]) > 0:
                mvts.append({"Libellé": "Acompte 2", "Montant": float(_to_num(row["Acompte 2"]))})
            if mvts:
                dfm = pd.DataFrame(mvts)
                dfm["Montant"] = dfm["Montant"].map(_fmt_money)
                st.dataframe(dfm, use_container_width=True, hide_index=True, key=skey("acct", "mvts"))
            else:
                st.caption("Aucun acompte enregistré dans le fichier (colonnes « Acompte 1 » / « Acompte 2 »).")

# Onglet Gestion
with tabs[5]:
    st.subheader("🧾 Gestion (Ajouter / Modifier / Supprimer)")
    df_live = df_all.copy() if df_all is not None else pd.DataFrame()

    if df_live.empty:
        st.info("Aucun client à gérer (chargez un fichier Clients).")
    else:
        op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=skey("crud", "op"))

        cats = sorted(df_visa_raw["Categories"].dropna().astype(str).unique().tolist()) if ("Categories" in df_visa_raw.columns and not df_visa_raw.empty) else sorted(df_live["Categories"].dropna().astype(str).unique().tolist())

        def subs_for(cat):
            if "Categories" in df_visa_raw.columns and "Sous-categorie" in df_visa_raw.columns:
                return sorted(df_visa_raw[df_visa_raw["Categories"].astype(str) == cat]["Sous-categorie"].dropna().astype(str).unique().tolist())
            return sorted(df_live[df_live["Categories"].astype(str) == cat]["Sous-categorie"].dropna().astype(str).unique().tolist())

        if op == "Ajouter":
            st.markdown("### ➕ Ajouter")
            c1, c2, c3 = st.columns(3)
            nom = c1.text_input("Nom", "", key=skey("add", "nom"))
            dval = _date_for_widget(date.today())
            dt = c2.date_input("Date de création", value=dval, key=skey("add", "date"))
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1, 13)], index=int(dval.month) - 1, key=skey("add", "mois"))

            v1, v2, v3 = st.columns(3)
            cat = v1.selectbox("Catégorie", [""] + cats, index=0, key=skey("add", "cat"))
            sub = ""
            if cat:
                subs = subs_for(cat)
                sub = v2.selectbox("Sous-catégorie", [""] + subs, index=0, key=skey("add", "sub"))
            visa_val = v3.text_input("Visa (libre ou dérivé)", sub if sub else "", key=skey("add", "visa"))

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=skey("add", "h"))
            other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=skey("add", "o"))
            acomp1 = st.number_input("Acompte 1", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=skey("add", "a1"))
            acomp2 = st.number_input("Acompte 2", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=skey("add", "a2"))
            comm = st.text_area("Commentaires", "", key=skey("add", "com"))

            s1, s2 = st.columns(2)
            sent_d = s1.date_input("Date denvoi", value=None, key=skey("add", "sentd"))
            acc_d = s1.date_input("Date dacceptation", value=None, key=skey("add", "accd"))
            ref_d = s2.date_input("Date de refus", value=None, key=skey("add", "refd"))
            ann_d = s2.date_input("Date dannulation", value=None, key=skey("add", "annd"))
            rfe = st.checkbox("RFE", value=False, key=skey("add", "rfe"))

            if st.button("💾 Enregistrer", key=skey("add", "save")):
                if not nom or not cat or not sub:
                    st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                    st.stop()
                total = float(honor) + float(other)
                paye = float(acomp1) + float(acomp2)
                solde = max(0.0, total - paye)
                new_id = f"{_norm(nom)}-{int(datetime.now().timestamp())}"
                new_dossier = next_dossier(df_live)

                new_row = {
                    "ID_Client": new_id,
                    "Dossier N": new_dossier,
                    "Nom": nom,
                    "Date": dt,
                    "Mois": f"{int(mois):02d}",
                    "Categories": cat,
                    "Sous-categorie": sub,
                    "Visa": visa_val,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Payé": paye,
                    "Solde": solde,
                    "Acompte 1": float(acomp1),
                    "Acompte 2": float(acomp2),
                    "Commentaires": comm,
                    "Date denvoi": sent_d,
                    "Date dacceptation": acc_d,
                    "Date de refus": ref_d,
                    "Date dannulation": ann_d,
                    "RFE": 1 if rfe else 0,
                }
                df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                st.success("Client ajouté (en mémoire). Utilisez l’onglet Export pour sauvegarder.")
                st.cache_data.clear()
                st.rerun()

        elif op == "Modifier":
            st.markdown("### ✏️ Modifier")
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            m1, m2 = st.columns(2)
            target_name = m1.selectbox("Nom", [""] + names, index=0, key=skey("mod", "nom"))
            target_id = m2.selectbox("ID_Client", [""] + ids, index=0, key=skey("mod", "id"))

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if not (mask is not None and mask.any()):
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()

            d1, d2, d3 = st.columns(3)
            nom = d1.text_input("Nom", _safe_str(row.get("Nom", "")), key=skey("mod", "nomv"))
            dval = _date_for_widget(row.get("Date"))
            dt = d2.date_input("Date de création", value=dval, key=skey("mod", "date"))
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1, 13)], index=max(0, int(_safe_str(row.get("Mois", "01"))) - 1), key=skey("mod", "mois"))

            v1, v2, v3 = st.columns(3)
            preset_cat = _safe_str(row.get("Categories", ""))
            cat = v1.selectbox("Catégorie", [""] + cats, index=(cats.index(preset_cat) + 1 if preset_cat in cats else 0), key=skey("mod", "cat"))
            sub = _safe_str(row.get("Sous-categorie", ""))
            if cat:
                subs = subs_for(cat)
                sub = v2.selectbox("Sous-catégorie", [""] + subs, index=(subs.index(sub) + 1 if sub in subs else 0), key=skey("mod", "sub"))
            visa_val = v3.text_input("Visa (libre ou dérivé)", _safe_str(row.get("Visa", "")), key=skey("mod", "visa"))

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=float(_to_num(row.get("Montant honoraires (US $)", 0))), step=50.0, format="%.2f", key=skey("mod", "h"))
            other = f2.number_input("Autres frais (US $)", min_value=0.0, value=float(_to_num(row.get("Autres frais (US $)", 0))), step=20.0, format="%.2f", key=skey("mod", "o"))
            acomp1 = st.number_input("Acompte 1", min_value=0.0, value=float(_to_num(row.get("Acompte 1", 0))), step=10.0, format="%.2f", key=skey("mod", "a1"))
            acomp2 = st.number_input("Acompte 2", min_value=0.0, value=float(_to_num(row.get("Acompte 2", 0))), step=10.0, format="%.2f", key=skey("mod", "a2"))
            comm = st.text_area("Commentaires", _safe_str(row.get("Commentaires", "")), key=skey("mod", "com"))

            s1, s2 = st.columns(2)
            sent_d = s1.date_input("Date denvoi", value=_date_for_widget(row.get("Date denvoi")), key=skey("mod", "sentd"))
            acc_d = s1.date_input("Date dacceptation", value=_date_for_widget(row.get("Date dacceptation")), key=skey("mod", "accd"))
            ref_d = s2.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=skey("mod", "refd"))
            ann_d = s2.date_input("Date dannulation", value=_date_for_widget(row.get("Date dannulation")), key=skey("mod", "annd"))
            rfe = st.checkbox("RFE", value=bool(int(_to_num(row.get("RFE", 0)) or 0)), key=skey("mod", "rfe"))

            if st.button("💾 Enregistrer les modifications", key=skey("mod", "save")):
                if not nom or not cat or not sub:
                    st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                    st.stop()
                total = float(honor) + float(other)
                paye = float(acomp1) + float(acomp2)
                solde = max(0.0, total - paye)

                df_live.at[idx, "Nom"] = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(mois):02d}"
                df_live.at[idx, "Categories"] = cat
                df_live.at[idx, "Sous-categorie"] = sub
                df_live.at[idx, "Visa"] = visa_val
                df_live.at[idx, "Montant honoraires (US $)"] = float(honor)
                df_live.at[idx, "Autres frais (US $)"] = float(other)
                df_live.at[idx, "Acompte 1"] = float(acomp1)
                df_live.at[idx, "Acompte 2"] = float(acomp2)
                df_live.at[idx, "Payé"] = float(paye)
                df_live.at[idx, "Solde"] = float(solde)
                df_live.at[idx, "Commentaires"] = comm
                df_live.at[idx, "Date denvoi"] = sent_d
                df_live.at[idx, "Date dacceptation"] = acc_d
                df_live.at[idx, "Date de refus"] = ref_d
                df_live.at[idx, "Date dannulation"] = ann_d
                df_live.at[idx, "RFE"] = 1 if rfe else 0

                st.success("Modifié (en mémoire). Utilisez Export pour sauvegarder.")
                st.cache_data.clear()
                st.rerun()

        elif op == "Supprimer":
            st.markdown("### 🗑️ Supprimer")
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            d1, d2 = st.columns(2)
            target_name = d1.selectbox("Nom", [""] + names, index=0, key=skey("del", "nom"))
            target_id = d2.selectbox("ID_Client", [""] + ids, index=0, key=skey("del", "id"))

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({"Dossier N": row.get("Dossier N", ""), "Nom": row.get("Nom", ""), "Visa": row.get("Visa", "")})
                if st.button("❗ Confirmer la suppression", key=skey("del", "btn")):
                    df_new = df_live[~mask].copy()
                    st.success("Client supprimé (en mémoire). Utilisez Export pour sauvegarder.")
                    st.cache_data.clear()
                    st.rerun()

# Onglet Visa
with tabs[6]:
    st.subheader("📄 Visa — aperçu")
    if df_visa_raw is None or df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True, key=skey("visa", "view"))

# Onglet Export
with tabs[7]:
    st.subheader("💾 Export")
    colx, coly = st.columns(2)

    with colx:
        if df_all is None or df_all.empty:
            st.info("Pas de Clients à exporter.")
        else:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                df_all.to_excel(w, index=False, sheet_name="Clients")
            st.download_button(
                "⬇️ Exporter Clients.xlsx",
                data=buf.getvalue(),
                file_name="Clients_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=skey("exp", "clients"),
            )

    with coly:
        if df_visa_raw is None or df_visa_raw.empty:
            st.info("Pas de Visa à exporter.")
        else:
            bufv = BytesIO()
            with pd.ExcelWriter(bufv, engine="openpyxl") as w:
                df_visa_raw.to_excel(w, index=False, sheet_name="Visa")
            st.download_button(
                "⬇️ Exporter Visa.xlsx",
                data=bufv.getvalue(),
                file_name="Visa_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=skey("exp", "visa"),
            )
