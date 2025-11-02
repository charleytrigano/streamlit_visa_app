import os
import json
import re
import io
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional
from pathlib import Path

import pandas as pd
import streamlit as st

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
    "Dossier refusé", "Dossier Annulé", "Commentaires",
    "Escrow"
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

def _date_for_widget(val: Any) -> date:
    if isinstance(val, date):
        return val
    if isinstance(val, datetime):
        return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d):
            return date.today()
        return d.date()
    except Exception:
        return date.today()

def _ensure_columns(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            if c in ["Payé", "Solde", "Montant honoraires (US $)", "Autres frais (US $)", "Acompte 1", "Acompte 2"]:
                out[c] = 0.0
            elif c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
                out[c] = 0
            elif c == "Escrow":
                out[c] = 0
            else:
                out[c] = ""
    return out[cols]

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
    if "Escrow" in df.columns:
        df["Escrow"] = df["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ["1", "true", "t", "yes", "oui", "y", "x"] else 0)
    else:
        df["Escrow"] = 0
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
        yearsA = sorted(pd.to_numeric(df_all["_Annee_"], errors="coerce").dropna().astype(int).unique().tolist())
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
            dfA = dfA[dfA["_Annee_"].isin(fy)]
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
                s = s[s["_Annee_"].isin(ys)]
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
        if "Escrow" not in dfE.columns:
            dfE["Escrow"] = 0
        dfE["Escrow"] = dfE["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ["1", "true", "t", "yes", "oui", "y", "x"] else 0)
        escrow_view = dfE[dfE["Escrow"] == 1].copy()

        if escrow_view.empty:
            st.info("Aucun dossier en Escrow.")
        else:
            escrow_view["Montant Escrow"] = escrow_view["Acompte 1"].apply(_to_num)
            escrow_view["Etat"] = escrow_view.apply(lambda r: "Réclamé" if (pd.notna(r.get("Date denvoi")) and r.get("Date denvoi")) else "À réclamer", axis=1)
            total_escrow = float(escrow_view["Montant Escrow"].sum())

            st.markdown(f"**Nombre dossiers Escrow : {len(escrow_view)}**")
            st.markdown(f"**Total montants Escrow : {_fmt_money(total_escrow)}**")
            st.dataframe(escrow_view[["Nom","Dossier N","Date","Date denvoi","Montant Escrow","Etat"]].reset_index(drop=True), use_container_width=True, height=320)
            st.markdown("#### Historique Escrow")
            st.dataframe(escrow_view[["Nom","Dossier N","Date","Montant Escrow","Date denvoi","Etat"]].sort_values("Date").reset_index(drop=True), use_container_width=True, height=220)
            # Export XLSX
            if st.button("Exporter les dossiers escrow en XLSX"):
                buf = BytesIO()
                export_df = escrow_view[["Nom","Dossier N","Date","Date denvoi","Montant Escrow","Etat"]]
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    export_df.to_excel(writer, index=False, sheet_name="Escrow")
                buf.seek(0)
                st.download_button("Télécharger XLSX", data=buf.getvalue(), file_name="escrow_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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

# Onglet Gestion (ajouter/modifier : alerte escrow lors de saisie/modification)
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
            escrow_val = st.checkbox("Escrow", value=False, key=skey("add", "escrow"))

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
                    "Escrow": 1 if escrow_val else 0
                }
                # --- Alerte escrow lors de saisie ---
                if new_row.get("Escrow",0) == 1 and pd.notna(new_row.get("Date denvoi")) and new_row.get("Date denvoi"):
                    montant_escrow = _to_num(new_row.get("Acompte 1",0))
                    st.info(f"⚠️ Escrow activé : Dossier {new_row.get('Dossier N','')} / Client {new_row.get('Nom','')} — Montant à réclamer : {_fmt_money(montant_escrow)}")
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
            escrow_val = st.checkbox("Escrow", value=bool(row.get("Escrow",0)), key=skey("mod", "escrow"))

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
                df_live.at[idx, "Escrow"] = 1 if escrow_val else 0
                # --- Alerte escrow lors de modification ---
                if df_live.at[idx, "Escrow"] == 1 and pd.notna(df_live.at[idx, "Date denvoi"]) and df_live.at[idx, "Date denvoi"]:
                    montant_escrow = _to_num(df_live.at[idx, "Acompte 1"])
                    st.info(f"⚠️ Escrow activé : Dossier {df_live.at[idx,'Dossier N']} / Client {df_live.at[idx,'Nom']} — Montant à réclamer : {_fmt_money(montant_escrow)}")

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

# === Début Partie 5 ===

# Fonctions supplémentaires pour validation, alertes et historique Escrow
def get_escrow_alert(row):
    """Retourne une alerte Escrow si dossier Escrow avec date d'envoi."""
    if row.get("Escrow", 0) == 1 and pd.notna(row.get("Date denvoi")) and row.get("Date denvoi"):
        montant_escrow = _to_num(row.get("Acompte 1", 0))
        return f"⚠️ Escrow activé : Dossier {row.get('Dossier N', '')} / Client {row.get('Nom', '')} — Montant à réclamer : {_fmt_money(montant_escrow)}"
    return None

def escrow_history(df):
    """Retourne le DataFrame historique des dossiers escrow."""
    escrow_rows = df[df["Escrow"] == 1].copy()
    escrow_rows["Montant Escrow"] = escrow_rows["Acompte 1"].apply(_to_num)
    escrow_rows["Etat"] = escrow_rows.apply(lambda r: "Réclamé" if (pd.notna(r.get("Date denvoi")) and r.get("Date denvoi")) else "À réclamer", axis=1)
    return escrow_rows[["Nom", "Dossier N", "Date", "Date denvoi", "Montant Escrow", "Etat"]].sort_values("Date")

# Ajout d’une section pour visualiser l’historique complet Escrow (optionnel, onglet dédié)
if "Escrow" in df_all.columns:
    with st.expander("📜 Historique complet Escrow", expanded=False):
        escrow_hist = escrow_history(df_all)
        st.dataframe(escrow_hist, use_container_width=True)
        if not escrow_hist.empty:
            st.markdown(f"**Total dossiers Escrow : {len(escrow_hist)}**")
            st.markdown(f"**Montant total Escrow : {_fmt_money(escrow_hist['Montant Escrow'].sum())}**")

# Ajout d’une section pour validation automatique des dossiers Escrow lors des exports
def export_escrow_xlsx(df, file_name="escrow_export.xlsx"):
    buf = BytesIO()
    export_df = df[df["Escrow"] == 1].copy()
    export_df["Montant Escrow"] = export_df["Acompte 1"].apply(_to_num)
    export_df["Etat"] = export_df.apply(lambda r: "Réclamé" if (pd.notna(r.get("Date denvoi")) and r.get("Date denvoi")) else "À réclamer", axis=1)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df[["Nom","Dossier N","Date","Date denvoi","Montant Escrow","Etat"]].to_excel(writer, index=False, sheet_name="Escrow")
    buf.seek(0)
    st.download_button("Télécharger export XLSX complet Escrow", data=buf.getvalue(), file_name=file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Possibilité d’ajouter d’autres fonctions annexes ici (validation, reporting, etc.)

# === Fin Partie 5 ===
# === Début Partie 6 ===

# Exemple de fonction pour filtrer les dossiers escrow à partir de df_live (utile en gestion)
def filter_escrow(df):
    """Filtre le DataFrame pour ne garder que les dossiers Escrow."""
    if "Escrow" not in df.columns:
        return pd.DataFrame()
    return df[df["Escrow"] == 1].copy()

# Option de visualisation rapide des dossiers escrow dans l’onglet Gestion
with st.expander("🔍 Voir les dossiers Escrow (Gestion)", expanded=False):
    escrow_gest = filter_escrow(df_live)
    if escrow_gest.empty:
        st.info("Aucun dossier Escrow dans la gestion en cours.")
    else:
        escrow_gest["Montant Escrow"] = escrow_gest["Acompte 1"].apply(_to_num)
        escrow_gest["Etat"] = escrow_gest.apply(lambda r: "Réclamé" if (pd.notna(r.get("Date denvoi")) and r.get("Date denvoi")) else "À réclamer", axis=1)
        st.dataframe(escrow_gest[["Nom", "Dossier N", "Date", "Date denvoi", "Montant Escrow", "Etat"]], use_container_width=True)
        st.markdown(f"**Nombre dossiers Escrow : {len(escrow_gest)}**")
        st.markdown(f"**Montant total Escrow : {_fmt_money(escrow_gest['Montant Escrow'].sum())}**")

# Ajout d’une option pour exporter tous les dossiers Escrow depuis la gestion
if st.button("Exporter tous les dossiers Escrow (Gestion)", key=skey("exp", "escrow_gest")):
    buf = BytesIO()
    export_df = escrow_gest[["Nom", "Dossier N", "Date", "Date denvoi", "Montant Escrow", "Etat"]]
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Escrow_Gestion")
    buf.seek(0)
    st.download_button("Télécharger XLSX (Escrow Gestion)", data=buf.getvalue(), file_name="escrow_gestion_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option d’affichage des alertes et validation Escrow pour chaque dossier dans la gestion
with st.expander("⚠️ Alertes Escrow (Gestion)", expanded=False):
    for idx, row in escrow_gest.iterrows():
        alert = get_escrow_alert(row)
        if alert:
            st.warning(alert)

# === Fin Partie 6 ===

# === Début Partie 7 ===

# Ajout d’une fonction d’audit Escrow pour voir les écarts (exemple audit)
def escrow_audit(df):
    """Audit des dossiers Escrow : signale les dossiers où le montant d'acompte est incohérent."""
    audit_rows = []
    for _, row in df[df["Escrow"] == 1].iterrows():
        acompte = _to_num(row.get("Acompte 1", 0))
        solde = _to_num(row.get("Solde", 0))
        if acompte < 100 or solde < 0:
            audit_rows.append({
                "Nom": row.get("Nom", ""),
                "Dossier N": row.get("Dossier N", ""),
                "Acompte 1": acompte,
                "Solde": solde,
                "Alerte": "Montant acompte trop faible ou solde négatif"
            })
    return pd.DataFrame(audit_rows)

# Affichage audit Escrow dans un onglet dédié ou dans Escrow
with st.expander("🔎 Audit Escrow (contrôles)", expanded=False):
    audit_df = escrow_audit(df_all)
    if audit_df.empty:
        st.success("Aucun écart détecté dans les dossiers Escrow.")
    else:
        st.warning("Des anomalies ont été détectées !")
        st.dataframe(audit_df, use_container_width=True)
        for _, row in audit_df.iterrows():
            st.error(f"Dossier {row['Dossier N']} ({row['Nom']}) — {row['Alerte']}")

# Ajout d’une option pour récapitulatif global Escrow (pour reporting)
with st.expander("📊 Synthèse Escrow (reporting)", expanded=False):
    escrow_rows = df_all[df_all["Escrow"] == 1].copy()
    total_escrow = escrow_rows["Acompte 1"].apply(_to_num).sum()
    nb_escrow = len(escrow_rows)
    st.markdown(f"**Nombre total dossiers Escrow : {nb_escrow}**")
    st.markdown(f"**Montant total Escrow : {_fmt_money(total_escrow)}**")
    # Répartition par catégorie
    if "Categories" in escrow_rows.columns:
        rep_cat = escrow_rows["Categories"].value_counts().rename_axis("Catégorie").reset_index(name="Nombre")
        st.dataframe(rep_cat, use_container_width=True)

# === Fin Partie 7 ===

# === Début Partie 8 ===

# Ajout d’une fonction de synthèse graphique Escrow (exemple avec matplotlib si souhaité)
import matplotlib.pyplot as plt

def plot_escrow_by_month(df):
    escrow_df = df[df["Escrow"] == 1].copy()
    escrow_df["Date"] = pd.to_datetime(escrow_df["Date"], errors="coerce")
    escrow_df["month"] = escrow_df["Date"].dt.to_period("M")
    monthly = escrow_df.groupby("month")["Acompte 1"].apply(lambda x: x.apply(_to_num).sum()).reset_index()
    fig, ax = plt.subplots()
    ax.bar(monthly["month"].astype(str), monthly["Acompte 1"], color="#0071bd")
    ax.set_xlabel("Mois")
    ax.set_ylabel("Montant Escrow")
    ax.set_title("Montant Escrow par mois")
    plt.xticks(rotation=45)
    st.pyplot(fig)

with st.expander("📉 Graphique Escrow par mois", expanded=False):
    plot_escrow_by_month(df_all)

# Option : ajout de synthèse Escrow par visa/catégorie
with st.expander("📑 Répartition Escrow par type de Visa", expanded=False):
    escrow = df_all[df_all["Escrow"] == 1].copy()
    if not escrow.empty and "Visa" in escrow.columns:
        rep_visa = escrow["Visa"].value_counts().rename_axis("Visa").reset_index(name="Nombre dossiers Escrow")
        st.dataframe(rep_visa, use_container_width=True)
        st.bar_chart(rep_visa.set_index("Visa"))

# Option : affichage de l’évolution du nombre de dossiers Escrow
with st.expander("📈 Évolution du nombre de dossiers Escrow", expanded=False):
    if not escrow.empty and "Date" in escrow.columns:
        escrow["Date"] = pd.to_datetime(escrow["Date"], errors="coerce")
        escrow["Année"] = escrow["Date"].dt.year
        yearly_count = escrow.groupby("Année").size().reset_index(name="Nombre dossiers")
        st.line_chart(yearly_count.set_index("Année"))

# === Fin Partie 8 ===

# === Début Partie 9 ===

# Option avancée : synthèse/export Escrow par sous-catégorie ou autre critère
with st.expander("📂 Export Escrow par sous-catégorie", expanded=False):
    if "Sous-categorie" in df_all.columns and "Escrow" in df_all.columns:
        escrow_subcat = df_all[df_all["Escrow"] == 1].copy()
        rep_subcat = escrow_subcat["Sous-categorie"].value_counts().rename_axis("Sous-catégorie").reset_index(name="Nombre dossiers Escrow")
        st.dataframe(rep_subcat, use_container_width=True)
        # Export XLSX
        if st.button("Exporter Escrow par sous-catégorie", key=skey("exp", "escrow_subcat")):
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                escrow_subcat.to_excel(writer, index=False, sheet_name="Escrow_SousCat")
            buf.seek(0)
            st.download_button("Télécharger XLSX (Escrow sous-catégorie)", data=buf.getvalue(), file_name="escrow_souscategorie_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : export global de tous les tableaux Escrow en un seul fichier
with st.expander("🗃️ Export global Escrow (multi-tableaux)", expanded=False):
    if "Escrow" in df_all.columns:
        escrow_all = df_all[df_all["Escrow"] == 1].copy()
        escrow_all["Montant Escrow"] = escrow_all["Acompte 1"].apply(_to_num)
        escrow_all["Etat"] = escrow_all.apply(lambda r: "Réclamé" if (pd.notna(r.get("Date denvoi")) and r.get("Date denvoi")) else "À réclamer", axis=1)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            escrow_all.to_excel(writer, index=False, sheet_name="Escrow")
            # Ajout synthèse par catégorie
            if "Categories" in escrow_all.columns:
                cat_tab = escrow_all["Categories"].value_counts().rename_axis("Catégorie").reset_index(name="Nombre")
                cat_tab.to_excel(writer, index=False, sheet_name="Synthese_Categorie")
            # Ajout synthèse par sous-catégorie
            if "Sous-categorie" in escrow_all.columns:
                subcat_tab = escrow_all["Sous-categorie"].value_counts().rename_axis("Sous-catégorie").reset_index(name="Nombre")
                subcat_tab.to_excel(writer, index=False, sheet_name="Synthese_SousCat")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Escrow global)", data=buf.getvalue(), file_name="escrow_global_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 9 ===

# === Début Partie 10 ===

# Option : affichage et export des dossiers Escrow par année
with st.expander("📅 Export Escrow par année", expanded=False):
    if "Date" in df_all.columns and "Escrow" in df_all.columns:
        escrow_year = df_all[df_all["Escrow"] == 1].copy()
        escrow_year["Date"] = pd.to_datetime(escrow_year["Date"], errors="coerce")
        escrow_year["Année"] = escrow_year["Date"].dt.year
        rep_year = escrow_year.groupby("Année").size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(rep_year, use_container_width=True)
        # Export XLSX
        if st.button("Exporter Escrow par année", key=skey("exp", "escrow_year")):
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                escrow_year.to_excel(writer, index=False, sheet_name="Escrow_Année")
            buf.seek(0)
            st.download_button("Télécharger XLSX (Escrow année)", data=buf.getvalue(), file_name="escrow_annee_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : synthèse finale et résumé Escrow
with st.expander("🧾 Synthèse finale Escrow", expanded=False):
    if "Escrow" in df_all.columns:
        escrow_final = df_all[df_all["Escrow"] == 1].copy()
        total_final = escrow_final["Acompte 1"].apply(_to_num).sum()
        st.markdown(f"**Nombre total dossiers Escrow : {len(escrow_final)}**")
        st.markdown(f"**Montant total Escrow : {_fmt_money(total_final)}**")
        # Synthèse par état
        escrow_final["Etat"] = escrow_final.apply(lambda r: "Réclamé" if (pd.notna(r.get("Date denvoi")) and r.get("Date denvoi")) else "À réclamer", axis=1)
        rep_etat = escrow_final["Etat"].value_counts().rename_axis("Etat").reset_index(name="Nombre")
        st.dataframe(rep_etat, use_container_width=True)

# === Fin Partie 10 ===

# Onglet Analyses
with tabs[2]:
    st.subheader("📈 Analyses")
    if df_all is None or df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA = sorted(pd.to_numeric(df_all["_Annee_"], errors="coerce").dropna().astype(int).unique().tolist())
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
            dfA = dfA[dfA["_Annee_"].isin(fy)]
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
                s = s[s["_Annee_"].isin(ys)]
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
        if "Escrow" not in dfE.columns:
            dfE["Escrow"] = 0
        dfE["Escrow"] = dfE["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ["1", "true", "t", "yes", "oui", "y", "x"] else 0)
        escrow_view = dfE[dfE["Escrow"] == 1].copy()

        if escrow_view.empty:
            st.info("Aucun dossier en Escrow.")
        else:
            escrow_view["Montant Escrow"] = escrow_view["Acompte 1"].apply(_to_num)
            escrow_view["Etat"] = escrow_view.apply(lambda r: "Réclamé" if (pd.notna(r.get("Date denvoi")) and r.get("Date denvoi")) else "À réclamer", axis=1)
            total_escrow = float(escrow_view["Montant Escrow"].sum())

            st.markdown(f"**Nombre dossiers Escrow : {len(escrow_view)}**")
            st.markdown(f"**Total montants Escrow : {_fmt_money(total_escrow)}**")
            st.dataframe(escrow_view[["Nom","Dossier N","Date","Date denvoi","Montant Escrow","Etat"]].reset_index(drop=True), use_container_width=True, height=320)
            st.markdown("#### Historique Escrow")
            st.dataframe(escrow_view[["Nom","Dossier N","Date","Montant Escrow","Date denvoi","Etat"]].sort_values("Date").reset_index(drop=True), use_container_width=True, height=220)
            # Export XLSX
            if st.button("Exporter les dossiers escrow en XLSX"):
                buf = BytesIO()
                export_df = escrow_view[["Nom","Dossier N","Date","Date denvoi","Montant Escrow","Etat"]]
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    export_df.to_excel(writer, index=False, sheet_name="Escrow")
                buf.seek(0)
                st.download_button("Télécharger XLSX", data=buf.getvalue(), file_name="escrow_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Début Partie 12 ===

# Option : affichage synthèse Escrow par type de visa et export associé
with st.expander("📒 Synthèse Escrow par type de Visa", expanded=False):
    if "Escrow" in df_all.columns and "Visa" in df_all.columns:
        escrow_visa = df_all[df_all["Escrow"] == 1].copy()
        synth_visa = escrow_visa.groupby("Visa").size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_visa, use_container_width=True)
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            synth_visa.to_excel(writer, index=False, sheet_name="Synthese_Escrow_Visa")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow Visa)", data=buf.getvalue(), file_name="escrow_synthese_visa.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : affichage synthèse Escrow par mois et export associé
with st.expander("📆 Synthèse Escrow par mois", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns:
        escrow_month = df_all[df_all["Escrow"] == 1].copy()
        escrow_month["Date"] = pd.to_datetime(escrow_month["Date"], errors="coerce")
        escrow_month["Mois"] = escrow_month["Date"].dt.month
        synth_month = escrow_month.groupby("Mois").size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_month, use_container_width=True)
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            synth_month.to_excel(writer, index=False, sheet_name="Synthese_Escrow_Mois")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow Mois)", data=buf.getvalue(), file_name="escrow_synthese_mois.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : export complet Escrow (archive finale)
with st.expander("🗄️ Export complet Escrow (archive finale)", expanded=False):
    if "Escrow" in df_all.columns:
        escrow_complete = df_all[df_all["Escrow"] == 1].copy()
        escrow_complete["Montant Escrow"] = escrow_complete["Acompte 1"].apply(_to_num)
        escrow_complete["Etat"] = escrow_complete.apply(lambda r: "Réclamé" if (pd.notna(r.get("Date denvoi")) and r.get("Date denvoi")) else "À réclamer", axis=1)
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            escrow_complete.to_excel(writer, index=False, sheet_name="Escrow_Complet")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Escrow complet)", data=buf.getvalue(), file_name="escrow_complet_final.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 12 ===

# === Début Partie 13 ===

# Option : affichage synthèse Escrow par trimestre et export associé
with st.expander("🗓️ Synthèse Escrow par trimestre", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns:
        escrow_trim = df_all[df_all["Escrow"] == 1].copy()
        escrow_trim["Date"] = pd.to_datetime(escrow_trim["Date"], errors="coerce")
        escrow_trim["Trimestre"] = escrow_trim["Date"].dt.quarter
        synth_trim = escrow_trim.groupby("Trimestre").size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_trim, use_container_width=True)
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            synth_trim.to_excel(writer, index=False, sheet_name="Synthese_Escrow_Trimestre")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow Trimestre)", data=buf.getvalue(), file_name="escrow_synthese_trimestre.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : affichage synthèse Escrow par jour et export associé
with st.expander("📅 Synthèse Escrow par jour", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns:
        escrow_day = df_all[df_all["Escrow"] == 1].copy()
        escrow_day["Date"] = pd.to_datetime(escrow_day["Date"], errors="coerce")
        escrow_day["Jour"] = escrow_day["Date"].dt.day
        synth_day = escrow_day.groupby("Jour").size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_day, use_container_width=True)
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            synth_day.to_excel(writer, index=False, sheet_name="Synthese_Escrow_Jour")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow Jour)", data=buf.getvalue(), file_name="escrow_synthese_jour.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : synthèse multi-axes Escrow (année, mois, catégorie)
with st.expander("🔀 Synthèse Escrow année/mois/catégorie", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Categories" in df_all.columns:
        escrow_multi = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi["Date"] = pd.to_datetime(escrow_multi["Date"], errors="coerce")
        escrow_multi["Année"] = escrow_multi["Date"].dt.year
        escrow_multi["Mois"] = escrow_multi["Date"].dt.month
        synth_multi = escrow_multi.groupby(["Année", "Mois", "Categories"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi, use_container_width=True)
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            synth_multi.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisCat")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow An/Mois/Catégorie)", data=buf.getvalue(), file_name="escrow_synthese_anneemoiscategorie.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 13 ===

# === Début Partie 14 ===

# Option : synthèse Escrow par montant (tranches) et export associé
with st.expander("💰 Synthèse Escrow par tranche de montant", expanded=False):
    if "Escrow" in df_all.columns and "Acompte 1" in df_all.columns:
        escrow_tranche = df_all[df_all["Escrow"] == 1].copy()
        escrow_tranche["Montant Escrow"] = escrow_tranche["Acompte 1"].apply(_to_num)
        bins = [0, 500, 1000, 2000, 5000, 10000, float('inf')]
        labels = ["0-500", "501-1000", "1001-2000", "2001-5000", "5001-10000", "10001+"]
        escrow_tranche["Tranche"] = pd.cut(escrow_tranche["Montant Escrow"], bins=bins, labels=labels, right=False)
        synth_tranche = escrow_tranche.groupby("Tranche").size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_tranche, use_container_width=True)
        st.bar_chart(synth_tranche.set_index("Tranche"))
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            synth_tranche.to_excel(writer, index=False, sheet_name="Synthese_Escrow_Tranche")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow Tranche)", data=buf.getvalue(), file_name="escrow_synthese_tranche.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : synthèse Escrow par paiement (payé/non payé) et export associé
with st.expander("💸 Synthèse Escrow par statut paiement", expanded=False):
    if "Escrow" in df_all.columns and "Payé" in df_all.columns:
        escrow_pay = df_all[df_all["Escrow"] == 1].copy()
        escrow_pay["PayéStatut"] = escrow_pay["Payé"].apply(lambda x: "Payé" if _to_num(x) > 0 else "Non payé")
        synth_pay = escrow_pay.groupby("PayéStatut").size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_pay, use_container_width=True)
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            synth_pay.to_excel(writer, index=False, sheet_name="Synthese_Escrow_Paiement")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow Paiement)", data=buf.getvalue(), file_name="escrow_synthese_paiement.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : synthèse Escrow par année et tranche de montant
with st.expander("📆💰 Synthèse Escrow par année & tranche", expanded=False):
    if "Escrow" in df_all.columns and "Acompte 1" in df_all.columns and "Date" in df_all.columns:
        escrow_year_tranche = df_all[df_all["Escrow"] == 1].copy()
        escrow_year_tranche["Montant Escrow"] = escrow_year_tranche["Acompte 1"].apply(_to_num)
        escrow_year_tranche["Date"] = pd.to_datetime(escrow_year_tranche["Date"], errors="coerce")
        escrow_year_tranche["Année"] = escrow_year_tranche["Date"].dt.year
        bins = [0, 500, 1000, 2000, 5000, 10000, float('inf')]
        labels = ["0-500", "501-1000", "1001-2000", "2001-5000", "5001-10000", "10001+"]
        escrow_year_tranche["Tranche"] = pd.cut(escrow_year_tranche["Montant Escrow"], bins=bins, labels=labels, right=False)
        synth_year_tranche = escrow_year_tranche.groupby(["Année", "Tranche"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_year_tranche, use_container_width=True)
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            synth_year_tranche.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeTranche")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow Année&Tranche)", data=buf.getvalue(), file_name="escrow_synthese_annee_tranche.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 14 ===

# === Début Partie 15 ===

# Option : synthèse Escrow par statut dossier (envoyé, approuvé, refusé, annulé) et export associé
with st.expander("📋 Synthèse Escrow par statut dossier", expanded=False):
    if "Escrow" in df_all.columns:
        escrow_statut = df_all[df_all["Escrow"] == 1].copy()
        # Création du statut
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                return "Envoyé"
            elif row.get("Dossier approuvé", 0) == 1:
                return "Approuvé"
            elif row.get("Dossier refusé", 0) == 1:
                return "Refusé"
            elif row.get("Dossier Annulé", 0) == 1:
                return "Annulé"
            else:
                return "En attente"
        escrow_statut["Statut dossier"] = escrow_statut.apply(statut_dossier, axis=1)
        synth_statut = escrow_statut.groupby("Statut dossier").size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_statut, use_container_width=True)
        st.bar_chart(synth_statut.set_index("Statut dossier"))
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            synth_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_Statut")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow Statut dossier)", data=buf.getvalue(), file_name="escrow_synthese_statut.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : synthèse Escrow dossiers avec RFE
with st.expander("📝 Synthèse Escrow dossiers avec RFE", expanded=False):
    if "Escrow" in df_all.columns and "RFE" in df_all.columns:
        escrow_rfe = df_all[(df_all["Escrow"] == 1) & (df_all["RFE"] == 1)].copy()
        st.markdown(f"**Nombre dossiers Escrow avec RFE : {len(escrow_rfe)}**")
        st.dataframe(escrow_rfe, use_container_width=True)
        # Export XLSX
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            escrow_rfe.to_excel(writer, index=False, sheet_name="Escrow_RFE")
        buf.seek(0)
        st.download_button("Télécharger XLSX (Escrow RFE)", data=buf.getvalue(), file_name="escrow_rfe_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 15 ===

# === Début Partie 16 ===

# Option : synthèse Escrow dossiers annulés/refusés et export associé
with st.expander("🚫 Synthèse Escrow dossiers annulés/refusés", expanded=False):
    if "Escrow" in df_all.columns and "Dossier Annulé" in df_all.columns and "Dossier refusé" in df_all.columns:
        escrow_annule = df_all[(df_all["Escrow"] == 1) & (df_all["Dossier Annulé"] == 1)].copy()
        escrow_refuse = df_all[(df_all["Escrow"] == 1) & (df_all["Dossier refusé"] == 1)].copy()
        st.markdown(f"**Nombre dossiers Escrow annulés : {len(escrow_annule)}**")
        st.dataframe(escrow_annule, use_container_width=True)
        st.markdown(f"**Nombre dossiers Escrow refusés : {len(escrow_refuse)}**")
        st.dataframe(escrow_refuse, use_container_width=True)
        # Export XLSX annulés
        buf_annule = BytesIO()
        with pd.ExcelWriter(buf_annule, engine="openpyxl") as writer:
            escrow_annule.to_excel(writer, index=False, sheet_name="Escrow_Annule")
        buf_annule.seek(0)
        st.download_button("Télécharger XLSX (Escrow annulés)", data=buf_annule.getvalue(), file_name="escrow_annules_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # Export XLSX refusés
        buf_refuse = BytesIO()
        with pd.ExcelWriter(buf_refuse, engine="openpyxl") as writer:
            escrow_refuse.to_excel(writer, index=False, sheet_name="Escrow_Refuse")
        buf_refuse.seek(0)
        st.download_button("Télécharger XLSX (Escrow refusés)", data=buf_refuse.getvalue(), file_name="escrow_refuses_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : synthèse Escrow dossiers envoyés non encore approuvés/refusés/annulés
with st.expander("🔔 Synthèse Escrow envoyés en attente", expanded=False):
    if "Escrow" in df_all.columns and "Dossiers envoyé" in df_all.columns:
        escrow_envoye = df_all[
            (df_all["Escrow"] == 1) &
            (df_all["Dossiers envoyé"] == 1) &
            (df_all.get("Dossier approuvé", 0) == 0) &
            (df_all.get("Dossier refusé", 0) == 0) &
            (df_all.get("Dossier Annulé", 0) == 0)
        ].copy()
        st.markdown(f"**Nombre dossiers Escrow envoyés en attente de décision : {len(escrow_envoye)}**")
        st.dataframe(escrow_envoye, use_container_width=True)
        # Export XLSX
        buf_envoye = BytesIO()
        with pd.ExcelWriter(buf_envoye, engine="openpyxl") as writer:
            escrow_envoye.to_excel(writer, index=False, sheet_name="Escrow_Envoye_Attente")
        buf_envoye.seek(0)
        st.download_button("Télécharger XLSX (Escrow envoyés en attente)", data=buf_envoye.getvalue(), file_name="escrow_envoyes_attente_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 16 ===

# === Début Partie 17 ===

# Option : synthèse Escrow dossiers approuvés et export associé
with st.expander("✅ Synthèse Escrow dossiers approuvés", expanded=False):
    if "Escrow" in df_all.columns and "Dossier approuvé" in df_all.columns:
        escrow_approved = df_all[(df_all["Escrow"] == 1) & (df_all["Dossier approuvé"] == 1)].copy()
        st.markdown(f"**Nombre dossiers Escrow approuvés : {len(escrow_approved)}**")
        st.dataframe(escrow_approved, use_container_width=True)
        # Export XLSX
        buf_approved = BytesIO()
        with pd.ExcelWriter(buf_approved, engine="openpyxl") as writer:
            escrow_approved.to_excel(writer, index=False, sheet_name="Escrow_Approuve")
        buf_approved.seek(0)
        st.download_button("Télécharger XLSX (Escrow approuvés)", data=buf_approved.getvalue(), file_name="escrow_approuves_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : synthèse Escrow dossiers en attente d’envoi
with st.expander("🕒 Synthèse Escrow en attente d’envoi", expanded=False):
    if "Escrow" in df_all.columns and "Dossiers envoyé" in df_all.columns:
        escrow_waiting = df_all[(df_all["Escrow"] == 1) & (df_all["Dossiers envoyé"] != 1)].copy()
        st.markdown(f"**Nombre dossiers Escrow en attente d’envoi : {len(escrow_waiting)}**")
        st.dataframe(escrow_waiting, use_container_width=True)
        # Export XLSX
        buf_waiting = BytesIO()
        with pd.ExcelWriter(buf_waiting, engine="openpyxl") as writer:
            escrow_waiting.to_excel(writer, index=False, sheet_name="Escrow_Attente_Envoi")
        buf_waiting.seek(0)
        st.download_button("Télécharger XLSX (Escrow attente envoi)", data=buf_waiting.getvalue(), file_name="escrow_attente_envoi_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : synthèse finale Escrow tous statuts et export
with st.expander("📊 Synthèse finale Escrow tous statuts", expanded=False):
    if "Escrow" in df_all.columns:
        escrow_final_all = df_all[df_all["Escrow"] == 1].copy()
        # Statut dossier
        def statut_final(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_final_all["Statut final"] = escrow_final_all.apply(statut_final, axis=1)
        synth_final = escrow_final_all.groupby("Statut final").size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_final, use_container_width=True)
        st.bar_chart(synth_final.set_index("Statut final"))
        # Export XLSX
        buf_final = BytesIO()
        with pd.ExcelWriter(buf_final, engine="openpyxl") as writer:
            synth_final.to_excel(writer, index=False, sheet_name="Synthese_Escrow_Final")
        buf_final.seek(0)
        st.download_button("Télécharger XLSX (Synthèse Escrow tous statuts)", data=buf_final.getvalue(), file_name="escrow_synthese_final_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 17 ===

# === Début Partie 18 ===

# Option : synthèse Escrow avec alertes sur montants et dates incohérents
with st.expander("⚠️ Alertes Escrow montants et dates incohérents", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Acompte 1" in df_all.columns:
        escrow_alerts = []
        for _, row in df_all[df_all["Escrow"] == 1].iterrows():
            montant = _to_num(row.get("Acompte 1", 0))
            date_envoi = row.get("Date denvoi", "")
            date_creation = row.get("Date", "")
            if montant <= 0:
                escrow_alerts.append({
                    "Nom": row.get("Nom", ""),
                    "Dossier N": row.get("Dossier N", ""),
                    "Alerte": "Montant Escrow nul ou négatif"
                })
            if pd.notna(date_envoi) and pd.notna(date_creation):
                try:
                    d_envoi = pd.to_datetime(date_envoi, errors="coerce")
                    d_creation = pd.to_datetime(date_creation, errors="coerce")
                    if d_envoi < d_creation:
                        escrow_alerts.append({
                            "Nom": row.get("Nom", ""),
                            "Dossier N": row.get("Dossier N", ""),
                            "Alerte": "Date d'envoi antérieure à la date de création"
                        })
                except Exception:
                    pass
        if escrow_alerts:
            alert_df = pd.DataFrame(escrow_alerts)
            st.warning("Des alertes ont été détectées dans les dossiers Escrow !")
            st.dataframe(alert_df, use_container_width=True)
            # Export XLSX
            buf_alert = BytesIO()
            with pd.ExcelWriter(buf_alert, engine="openpyxl") as writer:
                alert_df.to_excel(writer, index=False, sheet_name="Escrow_Alertes")
            buf_alert.seek(0)
            st.download_button("Télécharger XLSX (Alertes Escrow)", data=buf_alert.getvalue(), file_name="escrow_alertes_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("Aucune alerte détectée dans les dossiers Escrow.")

# Option : synthèse Escrow avec commentaires spécifiques
with st.expander("💬 Synthèse Escrow avec commentaires", expanded=False):
    if "Escrow" in df_all.columns and "Commentaires" in df_all.columns:
        escrow_comments = df_all[(df_all["Escrow"] == 1) & (df_all["Commentaires"].astype(str).str.strip() != "")].copy()
        st.markdown(f"**Nombre dossiers Escrow avec commentaires : {len(escrow_comments)}**")
        st.dataframe(escrow_comments[["Nom", "Dossier N", "Commentaires"]], use_container_width=True)
        # Export XLSX
        buf_comments = BytesIO()
        with pd.ExcelWriter(buf_comments, engine="openpyxl") as writer:
            escrow_comments.to_excel(writer, index=False, sheet_name="Escrow_Commentaires")
        buf_comments.seek(0)
        st.download_button("Télécharger XLSX (Escrow commentaires)", data=buf_comments.getvalue(), file_name="escrow_commentaires_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 18 ===

# === Début Partie 19 ===

# Option : Synthèse Escrow dossiers dont le solde n’est pas à zéro (alerte sur paiement partiel)
with st.expander("⚠️ Escrow dossiers solde non nul (paiement partiel)", expanded=False):
    if "Escrow" in df_all.columns and "Solde" in df_all.columns:
        escrow_solde = df_all[(df_all["Escrow"] == 1) & (_to_num(df_all["Solde"]) > 0)].copy()
        st.markdown(f"**Nombre dossiers Escrow solde non nul : {len(escrow_solde)}**")
        st.dataframe(escrow_solde, use_container_width=True)
        # Export XLSX
        buf_solde = BytesIO()
        with pd.ExcelWriter(buf_solde, engine="openpyxl") as writer:
            escrow_solde.to_excel(writer, index=False, sheet_name="Escrow_Solde_Non_Nul")
        buf_solde.seek(0)
        st.download_button("Télécharger XLSX (Escrow solde non nul)", data=buf_solde.getvalue(), file_name="escrow_solde_non_nul_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : Synthèse Escrow dossiers dont le solde est négatif (erreur ou anomalie)
with st.expander("🚨 Escrow dossiers solde négatif (anomalie)", expanded=False):
    if "Escrow" in df_all.columns and "Solde" in df_all.columns:
        escrow_neg = df_all[(df_all["Escrow"] == 1) & (_to_num(df_all["Solde"]) < 0)].copy()
        st.markdown(f"**Nombre dossiers Escrow solde négatif : {len(escrow_neg)}**")
        st.dataframe(escrow_neg, use_container_width=True)
        # Export XLSX
        buf_neg = BytesIO()
        with pd.ExcelWriter(buf_neg, engine="openpyxl") as writer:
            escrow_neg.to_excel(writer, index=False, sheet_name="Escrow_Solde_Negatif")
        buf_neg.seek(0)
        st.download_button("Télécharger XLSX (Escrow solde négatif)", data=buf_neg.getvalue(), file_name="escrow_solde_negatif_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : Synthèse Escrow dossiers avec acompte 2 (multi-acompte)
with st.expander("💵 Escrow dossiers avec acompte 2", expanded=False):
    if "Escrow" in df_all.columns and "Acompte 2" in df_all.columns:
        escrow_acomp2 = df_all[(df_all["Escrow"] == 1) & (_to_num(df_all["Acompte 2"]) > 0)].copy()
        st.markdown(f"**Nombre dossiers Escrow avec acompte 2 : {len(escrow_acomp2)}**")
        st.dataframe(escrow_acomp2, use_container_width=True)
        # Export XLSX
        buf_acomp2 = BytesIO()
        with pd.ExcelWriter(buf_acomp2, engine="openpyxl") as writer:
            escrow_acomp2.to_excel(writer, index=False, sheet_name="Escrow_Acompte_2")
        buf_acomp2.seek(0)
        st.download_button("Télécharger XLSX (Escrow acompte 2)", data=buf_acomp2.getvalue(), file_name="escrow_acompte2_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 19 ===

# === Début Partie 20 ===

# Option : Synthèse Escrow par catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par catégorie & statut dossier", expanded=False):
    if "Escrow" in df_all.columns and "Categories" in df_all.columns:
        escrow_multi_statut = df_all[df_all["Escrow"] == 1].copy()
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_statut["Statut dossier"] = escrow_multi_statut.apply(statut_dossier, axis=1)
        synth_multi_statut = escrow_multi_statut.groupby(["Categories", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_statut, use_container_width=True)
        # Export XLSX
        buf_multi_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_statut, engine="openpyxl") as writer:
            synth_multi_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_CategorieStatut")
        buf_multi_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Catégorie & Statut)", data=buf_multi_statut.getvalue(), file_name="escrow_categorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : Synthèse Escrow par type de visa et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par visa & statut dossier", expanded=False):
    if "Escrow" in df_all.columns and "Visa" in df_all.columns:
        escrow_multi_visa_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_visa_statut["Statut dossier"] = escrow_multi_visa_statut.apply(statut_dossier, axis=1)
        synth_multi_visa_statut = escrow_multi_visa_statut.groupby(["Visa", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_visa_statut, use_container_width=True)
        # Export XLSX
        buf_multi_visa_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_visa_statut, engine="openpyxl") as writer:
            synth_multi_visa_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_VisaStatut")
        buf_multi_visa_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Visa & Statut)", data=buf_multi_visa_statut.getvalue(), file_name="escrow_visa_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 20 ===

# === Début Partie 21 ===

# Option : Synthèse Escrow par sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par sous-catégorie & statut dossier", expanded=False):
    if "Escrow" in df_all.columns and "Sous-categorie" in df_all.columns:
        escrow_multi_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_subcat_statut["Statut dossier"] = escrow_multi_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_subcat_statut = escrow_multi_subcat_statut.groupby(["Sous-categorie", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_subcat_statut, engine="openpyxl") as writer:
            synth_multi_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_SousCatStatut")
        buf_multi_subcat_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Sous-catégorie & Statut)", data=buf_multi_subcat_statut.getvalue(), file_name="escrow_souscategorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : Synthèse Escrow par année, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Sous-categorie" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_annee_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_subcat_statut["Date"] = pd.to_datetime(escrow_multi_annee_subcat_statut["Date"], errors="coerce")
        escrow_multi_annee_subcat_statut["Année"] = escrow_multi_annee_subcat_statut["Date"].dt.year
        escrow_multi_annee_subcat_statut["Statut dossier"] = escrow_multi_annee_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_subcat_statut = escrow_multi_annee_subcat_statut.groupby(["Année", "Sous-categorie", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_subcat_statut, engine="openpyxl") as writer:
            synth_multi_annee_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeSousCatStatut")
        buf_multi_annee_subcat_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Année/Sous-catégorie/Statut)", data=buf_multi_annee_subcat_statut.getvalue(), file_name="escrow_annee_souscategorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 21 ===

# === Début Partie 22 ===

# Option : Synthèse Escrow par mois, catégorie, et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par mois/catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Categories" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_month_cat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_month_cat_statut["Date"] = pd.to_datetime(escrow_multi_month_cat_statut["Date"], errors="coerce")
        escrow_multi_month_cat_statut["Mois"] = escrow_multi_month_cat_statut["Date"].dt.month
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_month_cat_statut["Statut dossier"] = escrow_multi_month_cat_statut.apply(statut_dossier, axis=1)
        synth_multi_month_cat_statut = escrow_multi_month_cat_statut.groupby(["Mois", "Categories", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_month_cat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_month_cat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_month_cat_statut, engine="openpyxl") as writer:
            synth_multi_month_cat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_MoisCatStatut")
        buf_multi_month_cat_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Mois/Catégorie/Statut)", data=buf_multi_month_cat_statut.getvalue(), file_name="escrow_mois_categorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : Synthèse Escrow par mois, sous-catégorie, et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par mois/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Sous-categorie" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_month_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_month_subcat_statut["Date"] = pd.to_datetime(escrow_multi_month_subcat_statut["Date"], errors="coerce")
        escrow_multi_month_subcat_statut["Mois"] = escrow_multi_month_subcat_statut["Date"].dt.month
        escrow_multi_month_subcat_statut["Statut dossier"] = escrow_multi_month_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_month_subcat_statut = escrow_multi_month_subcat_statut.groupby(["Mois", "Sous-categorie", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_month_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_month_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_month_subcat_statut, engine="openpyxl") as writer:
            synth_multi_month_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_MoisSousCatStatut")
        buf_multi_month_subcat_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Mois/Sous-catégorie/Statut)", data=buf_multi_month_subcat_statut.getvalue(), file_name="escrow_mois_souscategorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 22 ===

# === Début Partie 23 ===

# Option : Synthèse Escrow par trimestre, catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par trimestre/catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Categories" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_trim_cat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_trim_cat_statut["Date"] = pd.to_datetime(escrow_multi_trim_cat_statut["Date"], errors="coerce")
        escrow_multi_trim_cat_statut["Trimestre"] = escrow_multi_trim_cat_statut["Date"].dt.quarter
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_trim_cat_statut["Statut dossier"] = escrow_multi_trim_cat_statut.apply(statut_dossier, axis=1)
        synth_multi_trim_cat_statut = escrow_multi_trim_cat_statut.groupby(["Trimestre", "Categories", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_trim_cat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_trim_cat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_trim_cat_statut, engine="openpyxl") as writer:
            synth_multi_trim_cat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_TrimCatStatut")
        buf_multi_trim_cat_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Trimestre/Catégorie/Statut)", data=buf_multi_trim_cat_statut.getvalue(), file_name="escrow_trimestre_categorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : Synthèse Escrow par trimestre, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par trimestre/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Sous-categorie" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_trim_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_trim_subcat_statut["Date"] = pd.to_datetime(escrow_multi_trim_subcat_statut["Date"], errors="coerce")
        escrow_multi_trim_subcat_statut["Trimestre"] = escrow_multi_trim_subcat_statut["Date"].dt.quarter
        escrow_multi_trim_subcat_statut["Statut dossier"] = escrow_multi_trim_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_trim_subcat_statut = escrow_multi_trim_subcat_statut.groupby(["Trimestre", "Sous-categorie", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_trim_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_trim_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_trim_subcat_statut, engine="openpyxl") as writer:
            synth_multi_trim_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_TrimSousCatStatut")
        buf_multi_trim_subcat_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Trimestre/Sous-catégorie/Statut)", data=buf_multi_trim_subcat_statut.getvalue(), file_name="escrow_trimestre_souscategorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 23 ===

# === Début Partie 24 ===

# Option : Synthèse Escrow par année, catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Categories" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_annee_cat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_cat_statut["Date"] = pd.to_datetime(escrow_multi_annee_cat_statut["Date"], errors="coerce")
        escrow_multi_annee_cat_statut["Année"] = escrow_multi_annee_cat_statut["Date"].dt.year
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_cat_statut["Statut dossier"] = escrow_multi_annee_cat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_cat_statut = escrow_multi_annee_cat_statut.groupby(["Année", "Categories", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_cat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_cat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_cat_statut, engine="openpyxl") as writer:
            synth_multi_annee_cat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeCatStatut")
        buf_multi_annee_cat_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Année/Catégorie/Statut)", data=buf_multi_annee_cat_statut.getvalue(), file_name="escrow_annee_categorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : Synthèse Escrow par année, visa et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/visa/statut", expanded=False):
    if "Escrow" in df_all.columns and "Visa" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_annee_visa_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_visa_statut["Date"] = pd.to_datetime(escrow_multi_annee_visa_statut["Date"], errors="coerce")
        escrow_multi_annee_visa_statut["Année"] = escrow_multi_annee_visa_statut["Date"].dt.year
        escrow_multi_annee_visa_statut["Statut dossier"] = escrow_multi_annee_visa_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_visa_statut = escrow_multi_annee_visa_statut.groupby(["Année", "Visa", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_visa_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_visa_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_visa_statut, engine="openpyxl") as writer:
            synth_multi_annee_visa_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeVisaStatut")
        buf_multi_annee_visa_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Année/Visa/Statut)", data=buf_multi_annee_visa_statut.getvalue(), file_name="escrow_annee_visa_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 24 ===

# === Début Partie 25 ===

# Option : Synthèse Escrow par année, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Sous-categorie" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_annee_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_subcat_statut["Date"] = pd.to_datetime(escrow_multi_annee_subcat_statut["Date"], errors="coerce")
        escrow_multi_annee_subcat_statut["Année"] = escrow_multi_annee_subcat_statut["Date"].dt.year
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_subcat_statut["Statut dossier"] = escrow_multi_annee_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_subcat_statut = escrow_multi_annee_subcat_statut.groupby(["Année", "Sous-categorie", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_subcat_statut, engine="openpyxl") as writer:
            synth_multi_annee_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeSousCatStatut")
        buf_multi_annee_subcat_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Année/Sous-catégorie/Statut)", data=buf_multi_annee_subcat_statut.getvalue(), file_name="escrow_annee_souscategorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : Synthèse Escrow par année, type de visa et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/visa/statut", expanded=False):
    if "Escrow" in df_all.columns and "Visa" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_annee_visa_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_visa_statut["Date"] = pd.to_datetime(escrow_multi_annee_visa_statut["Date"], errors="coerce")
        escrow_multi_annee_visa_statut["Année"] = escrow_multi_annee_visa_statut["Date"].dt.year
        escrow_multi_annee_visa_statut["Statut dossier"] = escrow_multi_annee_visa_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_visa_statut = escrow_multi_annee_visa_statut.groupby(["Année", "Visa", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_visa_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_visa_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_visa_statut, engine="openpyxl") as writer:
            synth_multi_annee_visa_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeVisaStatut")
        buf_multi_annee_visa_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Année/Visa/Statut)", data=buf_multi_annee_visa_statut.getvalue(), file_name="escrow_annee_visa_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 25 ===

# === Début Partie 26 ===

# Option : Synthèse Escrow par année, trimestre et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/trimestre/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns:
        escrow_multi_annee_trim_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_trim_statut["Date"] = pd.to_datetime(escrow_multi_annee_trim_statut["Date"], errors="coerce")
        escrow_multi_annee_trim_statut["Année"] = escrow_multi_annee_trim_statut["Date"].dt.year
        escrow_multi_annee_trim_statut["Trimestre"] = escrow_multi_annee_trim_statut["Date"].dt.quarter
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_trim_statut["Statut dossier"] = escrow_multi_annee_trim_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_trim_statut = escrow_multi_annee_trim_statut.groupby(["Année", "Trimestre", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_trim_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_trim_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_trim_statut, engine="openpyxl") as writer:
            synth_multi_annee_trim_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeTrimStatut")
        buf_multi_annee_trim_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Année/Trimestre/Statut)", data=buf_multi_annee_trim_statut.getvalue(), file_name="escrow_annee_trimestre_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Option : Synthèse Escrow par année, trimestre, catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/trimestre/catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Categories" in df_all.columns:
        escrow_multi_annee_trim_cat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_trim_cat_statut["Date"] = pd.to_datetime(escrow_multi_annee_trim_cat_statut["Date"], errors="coerce")
        escrow_multi_annee_trim_cat_statut["Année"] = escrow_multi_annee_trim_cat_statut["Date"].dt.year
        escrow_multi_annee_trim_cat_statut["Trimestre"] = escrow_multi_annee_trim_cat_statut["Date"].dt.quarter
        escrow_multi_annee_trim_cat_statut["Statut dossier"] = escrow_multi_annee_trim_cat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_trim_cat_statut = escrow_multi_annee_trim_cat_statut.groupby(["Année", "Trimestre", "Categories", "Statut dossier"]).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_trim_cat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_trim_cat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_trim_cat_statut, engine="openpyxl") as writer:
            synth_multi_annee_trim_cat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeTrimCatStatut")
        buf_multi_annee_trim_cat_statut.seek(0)
        st.download_button("Télécharger XLSX (Escrow Année/Trimestre/Catégorie/Statut)", data=buf_multi_annee_trim_cat_statut.getvalue(), file_name="escrow_annee_trimestre_categorie_statut_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === Fin Partie 26 ===

# === Début Partie 27 ===

# Option : Synthèse Escrow par année, trimestre, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/trimestre/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Sous-categorie" in df_all.columns:
        escrow_multi_annee_trim_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_trim_subcat_statut["Date"] = pd.to_datetime(escrow_multi_annee_trim_subcat_statut["Date"], errors="coerce")
        escrow_multi_annee_trim_subcat_statut["Année"] = escrow_multi_annee_trim_subcat_statut["Date"].dt.year
        escrow_multi_annee_trim_subcat_statut["Trimestre"] = escrow_multi_annee_trim_subcat_statut["Date"].dt.quarter
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_trim_subcat_statut["Statut dossier"] = escrow_multi_annee_trim_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_trim_subcat_statut = escrow_multi_annee_trim_subcat_statut.groupby(
            ["Année", "Trimestre", "Sous-categorie", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_trim_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_trim_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_trim_subcat_statut, engine="openpyxl") as writer:
            synth_multi_annee_trim_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeTrimSousCatStatut")
        buf_multi_annee_trim_subcat_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Trimestre/Sous-catégorie/Statut)",
            data=buf_multi_annee_trim_subcat_statut.getvalue(),
            file_name="escrow_annee_trimestre_souscategorie_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Option : Synthèse Escrow par année, trimestre, type de visa et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/trimestre/visa/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Visa" in df_all.columns:
        escrow_multi_annee_trim_visa_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_trim_visa_statut["Date"] = pd.to_datetime(escrow_multi_annee_trim_visa_statut["Date"], errors="coerce")
        escrow_multi_annee_trim_visa_statut["Année"] = escrow_multi_annee_trim_visa_statut["Date"].dt.year
        escrow_multi_annee_trim_visa_statut["Trimestre"] = escrow_multi_annee_trim_visa_statut["Date"].dt.quarter
        escrow_multi_annee_trim_visa_statut["Statut dossier"] = escrow_multi_annee_trim_visa_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_trim_visa_statut = escrow_multi_annee_trim_visa_statut.groupby(
            ["Année", "Trimestre", "Visa", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_trim_visa_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_trim_visa_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_trim_visa_statut, engine="openpyxl") as writer:
            synth_multi_annee_trim_visa_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeTrimVisaStatut")
        buf_multi_annee_trim_visa_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Trimestre/Visa/Statut)",
            data=buf_multi_annee_trim_visa_statut.getvalue(),
            file_name="escrow_annee_trimestre_visa_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === Fin Partie 27 ===

# === Début Partie 28 ===

# Option : Synthèse Escrow par année, trimestre, type de visa et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/trimestre/visa/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Visa" in df_all.columns:
        escrow_multi_annee_trim_visa_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_trim_visa_statut["Date"] = pd.to_datetime(escrow_multi_annee_trim_visa_statut["Date"], errors="coerce")
        escrow_multi_annee_trim_visa_statut["Année"] = escrow_multi_annee_trim_visa_statut["Date"].dt.year
        escrow_multi_annee_trim_visa_statut["Trimestre"] = escrow_multi_annee_trim_visa_statut["Date"].dt.quarter
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_trim_visa_statut["Statut dossier"] = escrow_multi_annee_trim_visa_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_trim_visa_statut = escrow_multi_annee_trim_visa_statut.groupby(
            ["Année", "Trimestre", "Visa", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_trim_visa_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_trim_visa_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_trim_visa_statut, engine="openpyxl") as writer:
            synth_multi_annee_trim_visa_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeTrimVisaStatut")
        buf_multi_annee_trim_visa_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Trimestre/Visa/Statut)",
            data=buf_multi_annee_trim_visa_statut.getvalue(),
            file_name="escrow_annee_trimestre_visa_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Option : Synthèse Escrow par année, trimestre, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/trimestre/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Sous-categorie" in df_all.columns:
        escrow_multi_annee_trim_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_trim_subcat_statut["Date"] = pd.to_datetime(escrow_multi_annee_trim_subcat_statut["Date"], errors="coerce")
        escrow_multi_annee_trim_subcat_statut["Année"] = escrow_multi_annee_trim_subcat_statut["Date"].dt.year
        escrow_multi_annee_trim_subcat_statut["Trimestre"] = escrow_multi_annee_trim_subcat_statut["Date"].dt.quarter
        escrow_multi_annee_trim_subcat_statut["Statut dossier"] = escrow_multi_annee_trim_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_trim_subcat_statut = escrow_multi_annee_trim_subcat_statut.groupby(
            ["Année", "Trimestre", "Sous-categorie", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_trim_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_trim_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_trim_subcat_statut, engine="openpyxl") as writer:
            synth_multi_annee_trim_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeTrimSousCatStatut")
        buf_multi_annee_trim_subcat_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Trimestre/Sous-catégorie/Statut)",
            data=buf_multi_annee_trim_subcat_statut.getvalue(),
            file_name="escrow_annee_trimestre_souscategorie_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === Fin Partie 28 ===

# === Début Partie 29 ===

# Option : Synthèse Escrow par année, mois, catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Categories" in df_all.columns:
        escrow_multi_annee_mois_cat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_cat_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_cat_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_cat_statut["Année"] = escrow_multi_annee_mois_cat_statut["Date"].dt.year
        escrow_multi_annee_mois_cat_statut["Mois"] = escrow_multi_annee_mois_cat_statut["Date"].dt.month
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_mois_cat_statut["Statut dossier"] = escrow_multi_annee_mois_cat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_cat_statut = escrow_multi_annee_mois_cat_statut.groupby(
            ["Année", "Mois", "Categories", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_cat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_cat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_cat_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_cat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisCatStatut")
        buf_multi_annee_mois_cat_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Catégorie/Statut)",
            data=buf_multi_annee_mois_cat_statut.getvalue(),
            file_name="escrow_annee_mois_categorie_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Option : Synthèse Escrow par année, mois, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Sous-categorie" in df_all.columns:
        escrow_multi_annee_mois_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_subcat_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_subcat_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_subcat_statut["Année"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.year
        escrow_multi_annee_mois_subcat_statut["Mois"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.month
        escrow_multi_annee_mois_subcat_statut["Statut dossier"] = escrow_multi_annee_mois_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_subcat_statut = escrow_multi_annee_mois_subcat_statut.groupby(
            ["Année", "Mois", "Sous-categorie", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_subcat_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisSousCatStatut")
        buf_multi_annee_mois_subcat_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Sous-catégorie/Statut)",
            data=buf_multi_annee_mois_subcat_statut.getvalue(),
            file_name="escrow_annee_mois_souscategorie_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === Fin Partie 29 ===

# === Début Partie 30 ===

# Option : Synthèse Escrow par année, mois, type de visa et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/visa/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Visa" in df_all.columns:
        escrow_multi_annee_mois_visa_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_visa_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_visa_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_visa_statut["Année"] = escrow_multi_annee_mois_visa_statut["Date"].dt.year
        escrow_multi_annee_mois_visa_statut["Mois"] = escrow_multi_annee_mois_visa_statut["Date"].dt.month
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_mois_visa_statut["Statut dossier"] = escrow_multi_annee_mois_visa_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_visa_statut = escrow_multi_annee_mois_visa_statut.groupby(
            ["Année", "Mois", "Visa", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_visa_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_visa_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_visa_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_visa_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisVisaStatut")
        buf_multi_annee_mois_visa_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Visa/Statut)",
            data=buf_multi_annee_mois_visa_statut.getvalue(),
            file_name="escrow_annee_mois_visa_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Option : Synthèse Escrow par année, mois, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Sous-categorie" in df_all.columns:
        escrow_multi_annee_mois_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_subcat_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_subcat_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_subcat_statut["Année"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.year
        escrow_multi_annee_mois_subcat_statut["Mois"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.month
        escrow_multi_annee_mois_subcat_statut["Statut dossier"] = escrow_multi_annee_mois_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_subcat_statut = escrow_multi_annee_mois_subcat_statut.groupby(
            ["Année", "Mois", "Sous-categorie", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_subcat_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisSousCatStatut")
        buf_multi_annee_mois_subcat_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Sous-catégorie/Statut)",
            data=buf_multi_annee_mois_subcat_statut.getvalue(),
            file_name="escrow_annee_mois_souscategorie_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === Fin Partie 30 ===

# === Début Partie 31 ===

# Option : Synthèse Escrow par année, mois, visa et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/visa/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Visa" in df_all.columns:
        escrow_multi_annee_mois_visa_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_visa_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_visa_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_visa_statut["Année"] = escrow_multi_annee_mois_visa_statut["Date"].dt.year
        escrow_multi_annee_mois_visa_statut["Mois"] = escrow_multi_annee_mois_visa_statut["Date"].dt.month
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_mois_visa_statut["Statut dossier"] = escrow_multi_annee_mois_visa_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_visa_statut = escrow_multi_annee_mois_visa_statut.groupby(
            ["Année", "Mois", "Visa", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_visa_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_visa_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_visa_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_visa_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisVisaStatut")
        buf_multi_annee_mois_visa_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Visa/Statut)",
            data=buf_multi_annee_mois_visa_statut.getvalue(),
            file_name="escrow_annee_mois_visa_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Option : Synthèse Escrow par année, mois, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Sous-categorie" in df_all.columns:
        escrow_multi_annee_mois_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_subcat_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_subcat_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_subcat_statut["Année"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.year
        escrow_multi_annee_mois_subcat_statut["Mois"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.month
        escrow_multi_annee_mois_subcat_statut["Statut dossier"] = escrow_multi_annee_mois_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_subcat_statut = escrow_multi_annee_mois_subcat_statut.groupby(
            ["Année", "Mois", "Sous-categorie", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_subcat_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisSousCatStatut")
        buf_multi_annee_mois_subcat_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Sous-catégorie/Statut)",
            data=buf_multi_annee_mois_subcat_statut.getvalue(),
            file_name="escrow_annee_mois_souscategorie_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === Fin Partie 31 ===

# === Début Partie 32 ===

# Option : Synthèse Escrow par année, mois, catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Categories" in df_all.columns:
        escrow_multi_annee_mois_cat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_cat_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_cat_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_cat_statut["Année"] = escrow_multi_annee_mois_cat_statut["Date"].dt.year
        escrow_multi_annee_mois_cat_statut["Mois"] = escrow_multi_annee_mois_cat_statut["Date"].dt.month
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_mois_cat_statut["Statut dossier"] = escrow_multi_annee_mois_cat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_cat_statut = escrow_multi_annee_mois_cat_statut.groupby(
            ["Année", "Mois", "Categories", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_cat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_cat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_cat_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_cat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisCatStatut")
        buf_multi_annee_mois_cat_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Catégorie/Statut)",
            data=buf_multi_annee_mois_cat_statut.getvalue(),
            file_name="escrow_annee_mois_categorie_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Option : Synthèse Escrow par année, mois, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Sous-categorie" in df_all.columns:
        escrow_multi_annee_mois_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_subcat_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_subcat_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_subcat_statut["Année"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.year
        escrow_multi_annee_mois_subcat_statut["Mois"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.month
        escrow_multi_annee_mois_subcat_statut["Statut dossier"] = escrow_multi_annee_mois_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_subcat_statut = escrow_multi_annee_mois_subcat_statut.groupby(
            ["Année", "Mois", "Sous-categorie", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_subcat_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisSousCatStatut")
        buf_multi_annee_mois_subcat_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Sous-catégorie/Statut)",
            data=buf_multi_annee_mois_subcat_statut.getvalue(),
            file_name="escrow_annee_mois_souscategorie_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === Fin Partie 32 ===

# === Début Partie 33 ===

# Option : Synthèse Escrow par année, mois, type de visa et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/visa/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Visa" in df_all.columns:
        escrow_multi_annee_mois_visa_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_visa_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_visa_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_visa_statut["Année"] = escrow_multi_annee_mois_visa_statut["Date"].dt.year
        escrow_multi_annee_mois_visa_statut["Mois"] = escrow_multi_annee_mois_visa_statut["Date"].dt.month
        def statut_dossier(row):
            if row.get("Dossiers envoyé", 0) == 1:
                if row.get("Dossier approuvé", 0) == 1:
                    return "Approuvé"
                elif row.get("Dossier refusé", 0) == 1:
                    return "Refusé"
                elif row.get("Dossier Annulé", 0) == 1:
                    return "Annulé"
                else:
                    return "Envoyé en attente"
            else:
                return "Non envoyé"
        escrow_multi_annee_mois_visa_statut["Statut dossier"] = escrow_multi_annee_mois_visa_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_visa_statut = escrow_multi_annee_mois_visa_statut.groupby(
            ["Année", "Mois", "Visa", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_visa_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_visa_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_visa_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_visa_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisVisaStatut")
        buf_multi_annee_mois_visa_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Visa/Statut)",
            data=buf_multi_annee_mois_visa_statut.getvalue(),
            file_name="escrow_annee_mois_visa_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Option : Synthèse Escrow par année, mois, sous-catégorie et statut dossier (multi-axes) et export associé
with st.expander("🔢 Synthèse Escrow par année/mois/sous-catégorie/statut", expanded=False):
    if "Escrow" in df_all.columns and "Date" in df_all.columns and "Sous-categorie" in df_all.columns:
        escrow_multi_annee_mois_subcat_statut = df_all[df_all["Escrow"] == 1].copy()
        escrow_multi_annee_mois_subcat_statut["Date"] = pd.to_datetime(escrow_multi_annee_mois_subcat_statut["Date"], errors="coerce")
        escrow_multi_annee_mois_subcat_statut["Année"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.year
        escrow_multi_annee_mois_subcat_statut["Mois"] = escrow_multi_annee_mois_subcat_statut["Date"].dt.month
        escrow_multi_annee_mois_subcat_statut["Statut dossier"] = escrow_multi_annee_mois_subcat_statut.apply(statut_dossier, axis=1)
        synth_multi_annee_mois_subcat_statut = escrow_multi_annee_mois_subcat_statut.groupby(
            ["Année", "Mois", "Sous-categorie", "Statut dossier"]
        ).size().reset_index(name="Nombre dossiers Escrow")
        st.dataframe(synth_multi_annee_mois_subcat_statut, use_container_width=True)
        # Export XLSX
        buf_multi_annee_mois_subcat_statut = BytesIO()
        with pd.ExcelWriter(buf_multi_annee_mois_subcat_statut, engine="openpyxl") as writer:
            synth_multi_annee_mois_subcat_statut.to_excel(writer, index=False, sheet_name="Synthese_Escrow_AnneeMoisSousCatStatut")
        buf_multi_annee_mois_subcat_statut.seek(0)
        st.download_button(
            "Télécharger XLSX (Escrow Année/Mois/Sous-catégorie/Statut)",
            data=buf_multi_annee_mois_subcat_statut.getvalue(),
            file_name="escrow_annee_mois_souscategorie_statut_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# === Fin du script Escrow ===
