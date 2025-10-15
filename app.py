# =========================
# Visa Manager — PARTIE 1/4
# Imports, constantes, utilitaires, mémoire des chemins
# =========================

import os, json, re, io, zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Tuple, Dict, Any, List, Optional

import pandas as pd
import streamlit as st

APP_TITLE = "🛂 Visa Manager"

# ---------- Colonnes attendues côté Clients ----------
COLS_CLIENTS = [
    "ID_Client", "Dossier N", "Nom", "Date",
    "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)",
    "Payé", "Solde", "Acompte 1", "Acompte 2",
    "RFE", "Dossiers envoyé", "Dossier approuvé",
    "Dossier refusé", "Dossier Annulé", "Commentaires"
]

# ---------- Constantes diverses ----------
MEMO_FILE = "_vmemory.json"
SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"

# Sécurité clés Streamlit
SID = "vmgr"

st.set_page_config(page_title="Visa Manager", layout="wide")
st.title(APP_TITLE)

# ============== Utilitaires ==============

def _safe_str(x: Any) -> str:
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def _to_num(x: Any) -> float:
    """Convertit de manière robuste en float (0.0 si vide)."""
    if isinstance(x, (int, float)):
        return float(x)
    s = _safe_str(x)
    if not s:
        return 0.0
    s = re.sub(r"[^\d,.\-]", "", s)
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
    """Retourne un objet date sûr pour st.date_input."""
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
    """Ajoute les colonnes manquantes avec valeurs par défaut."""
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            if c in ["Payé", "Solde", "Montant honoraires (US $)", "Autres frais (US $)", "Acompte 1", "Acompte 2"]:
                out[c] = 0.0
            elif c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
                out[c] = 0
            else:
                out[c] = ""
    return out[cols]

def _normalize_clients_numeric(df: pd.DataFrame) -> pd.DataFrame:
    num_cols = ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Acompte 1", "Acompte 2"]
    for c in num_cols:
        if c in df.columns:
            df[c] = df[c].apply(_to_num)
    # recalculs de securité
    if "Montant honoraires (US $)" in df.columns and "Autres frais (US $)" in df.columns:
        total = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]
        paye = df["Payé"] if "Payé" in df.columns else 0.0
        df["Solde"] = (total - paye).clip(lower=0.0)
    return df

def _normalize_status(df: pd.DataFrame) -> pd.DataFrame:
    """Statuts en 0/1 (checkbox)."""
    for c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: 1 if str(x).strip() in ["1", "True", "true", "OUI", "Oui", "oui", "X", "x"] else 0)
        else:
            df[c] = 0
    return df

def normalize_clients(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS_CLIENTS)
    # harmonise colonnes
    df = df.copy()
    # Rename possibles (tolérance minimale)
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
        "Dossier envoyé": "Dossiers envoyé",  # on standardise sur pluriel
    }
    df.rename(columns={k:v for k,v in ren.items() if k in df.columns}, inplace=True)

    # colonnes obligatoires
    df = _ensure_columns(df, COLS_CLIENTS)

    # date (string -> date/str)
    if "Date" in df.columns:
        try:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        except Exception:
            pass

    # numérique + statuts
    df = _normalize_clients_numeric(df)
    df = _normalize_status(df)

    # nettoie champs texte essentiels
    df["Nom"] = df["Nom"].astype(str)
    df["Categories"] = df["Categories"].astype(str)
    df["Sous-categorie"] = df["Sous-categorie"].astype(str)
    df["Visa"] = df["Visa"].astype(str)
    df["Commentaires"] = df["Commentaires"].astype(str)

    # Mois (MM) et Annee cachees pour filtres
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

# ============== Lecture de fichiers (xlsx/csv) ==============

def read_any_table(src: Any, sheet: Optional[str] = None) -> Optional[pd.DataFrame]:
    """
    src peut être : chemin str, BytesIO, UploadedFile streamlit.
    sheet : None pour CSV, ou nom d'onglet pour Excel.
    """
    if src is None:
        return None

    # UploadedFile de Streamlit
    if hasattr(src, "read") and hasattr(src, "name"):
        name = src.name.lower()
        data = src.read()
        bio = BytesIO(data)
        if name.endswith(".csv"):
            return pd.read_csv(bio)
        return pd.read_excel(bio, sheet_name=(sheet if sheet else 0))
    # Chemin local
    if isinstance(src, (str, os.PathLike)):
        p = str(src)
        if not os.path.exists(p):
            return None
        if p.lower().endswith(".csv"):
            return pd.read_csv(p)
        return pd.read_excel(p, sheet_name=(sheet if sheet else 0))
    # BytesIO
    if isinstance(src, (io.BytesIO, BytesIO)):
        # pas de nom : on tente excel puis csv
        try:
            bio2 = BytesIO(src.getvalue())
            return pd.read_excel(bio2, sheet_name=(sheet if sheet else 0))
        except Exception:
            src.seek(0)
            return pd.read_csv(src)
    return None

# ============== Mémoire des derniers chemins ==============

def load_last_paths() -> Tuple[str, str, str]:
    """Retourne (last_clients, last_visa, last_save_dir)."""
    if not os.path.exists(MEMO_FILE):
        return "", "", ""
    try:
        with open(MEMO_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("clients",""), data.get("visa",""), data.get("save_dir","")
    except Exception:
        return "", "", ""

def save_last_paths(clients_path: str, visa_path: str, save_dir: str) -> None:
    data = {"clients": clients_path or "", "visa": visa_path or "", "save_dir": save_dir or ""}
    try:
        with open(MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# petit helper pour des clés uniques
def skey(*parts: str) -> str:
    return f"{SID}_" + "_".join([p for p in parts if p])

# =========================
# FIN PARTIE 1/4
# =========================



# =========================
# Visa Manager — PARTIE 2/4
# Barre latérale : chargement & mémoire
# Carte Visa (catégories / sous-catégories / options)
# Création des onglets
# =========================

st.sidebar.header("📂 Fichiers")

# Derniers chemins mémorisés
last_clients, last_visa, last_save_dir = load_last_paths()

mode = st.sidebar.radio(
    "Mode de chargement",
    ["Un fichier (Clients)", "Deux fichiers (Clients & Visa)"],
    index=0,
    key=skey("mode")
)

# Uploaders
up_clients = st.sidebar.file_uploader(
    "Clients (xlsx/csv)", type=["xlsx","xls","csv"], key=skey("up_clients")
)
up_visa = None
if mode == "Deux fichiers (Clients & Visa)":
    up_visa = st.sidebar.file_uploader(
        "Visa (xlsx/csv)", type=["xlsx","xls","csv"], key=skey("up_visa")
    )

st.sidebar.markdown("—")
# Saisies de chemins locaux (optionnel) + boutons "charger"
clients_path_in = st.sidebar.text_input("ou chemin local Clients", value=last_clients, key=skey("cli_path"))
visa_path_in    = st.sidebar.text_input("ou chemin local Visa", value=(last_visa if mode!="Un fichier (Clients)" else ""), key=skey("vis_path"))
save_dir_in     = st.sidebar.text_input("Dossier de sauvegarde", value=last_save_dir, key=skey("save_dir"))

if st.sidebar.button("📥 Charger", key=skey("btn_load")):
    # on mémorise ce que l'utilisateur a indiqué
    save_last_paths(clients_path_in, visa_path_in, save_dir_in)
    st.success("Chemins mémorisés. Re-lancement pour prise en compte.")
    st.rerun()

# ----- Lecture Clients -----
clients_src = up_clients if up_clients is not None else (clients_path_in if clients_path_in else last_clients)
df_clients_raw = normalize_clients(read_any_table(clients_src))

# ----- Lecture Visa -------
if mode == "Deux fichiers (Clients & Visa)":
    visa_src = up_visa if up_visa is not None else (visa_path_in if visa_path_in else last_visa)
else:
    # si un seul fichier : on autorise un onglet "Visa" dans le même fichier
    visa_src = up_clients if up_clients is not None else (clients_path_in if clients_path_in else last_clients)

df_visa_raw = read_any_table(visa_src, sheet=SHEET_VISA) or read_any_table(visa_src)  # tolérant
if df_visa_raw is None:
    df_visa_raw = pd.DataFrame()

# Affichage résumé chargement
with st.expander("📄 Fichiers chargés", expanded=True):
    st.write("**Clients** :", ("(aucun)" if (df_clients_raw is None or df_clients_raw.empty) else (getattr(clients_src, 'name', str(clients_src)))))
    st.write("**Visa** :",    ("(aucun)" if (df_visa_raw is None or df_visa_raw.empty) else (getattr(visa_src, 'name', str(visa_src)))))

# ---------- Construction carte Visa ----------
# Format attendu du fichier Visa :
# - Colonnes : "Categories", "Sous-categorie", et une série d'options en en-tête (ex: COS, EOS, ...).
# - Chaque ligne : 1 pour l'option disponible, vide sinon.
def build_visa_map(dfv: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    vm: Dict[str, Dict[str, Dict[str, Any]]] = {}
    if dfv is None or dfv.empty:
        return vm

    cols = [c for c in dfv.columns if _safe_str(c)]
    # colonnes minimales
    if "Categories" not in cols and "Catégorie" in cols:
        dfv = dfv.rename(columns={"Catégorie": "Categories"})
    if "Sous-categorie" not in cols and "Sous-catégorie" in cols:
        dfv = dfv.rename(columns={"Sous-catégorie": "Sous-categorie"})

    if "Categories" not in dfv.columns or "Sous-categorie" not in dfv.columns:
        return vm

    fixed = ["Categories","Sous-categorie"]
    option_cols = [c for c in dfv.columns if c not in fixed]

    for _, row in dfv.iterrows():
        cat = _safe_str(row.get("Categories","")).strip()
        sub = _safe_str(row.get("Sous-categorie","")).strip()
        if not cat or not sub:
            continue
        vm.setdefault(cat, {})
        vm[cat].setdefault(sub, {"exclusive": None, "options": []})

        opts = []
        for oc in option_cols:
            val = _safe_str(row.get(oc,"")).strip()
            # on considère 1, X, Oui = disponible
            if val in ["1","x","X","oui","Oui","OUI","True","true"]:
                opts.append(oc)
        # stratégie : si 2 options exactement "COS" et "EOS" -> exclusives
        exclusive = None
        if set([o.upper() for o in opts]) == set(["COS","EOS"]):
            exclusive = "radio_group"

        vm[cat][sub] = {"exclusive": exclusive, "options": opts}
    return vm

visa_map = build_visa_map(df_visa_raw)

# Helper d’UI : propose les options d’un (cat, sub)
def visa_option_selector(vm: Dict[str, Any], cat: str, sub: str, keybase: str) -> str:
    """
    Retourne le Visa final choisi, formé de :
    - si options exclusives (COS/EOS): "sub + ' ' + choix"
    - sinon: "sub" (ou concat multi si besoin)
    """
    if cat not in vm or sub not in vm[cat]:
        return sub
    meta = vm[cat][sub]
    opts = meta.get("options", [])
    if not opts:
        return sub

    if meta.get("exclusive") == "radio_group" and set([o.upper() for o in opts]) == set(["COS","EOS"]):
        pick = st.radio("Options", ["COS", "EOS"], horizontal=True, key=skey(keybase,"opt"))
        return f"{sub} {pick}"
    else:
        picks = st.multiselect("Options", opts, default=[], key=skey(keybase,"opts"))
        if not picks:
            return sub
        # par simplicité on ne concatène pas ; on prend le premier pick
        return f"{sub} {picks[0]}"

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



# =========================
# Visa Manager — PARTIE 3/4
# 📊 Dashboard • 📈 Analyses • 🏦 Escrow
# =========================

def _ensure_time_features(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()
    if "Date" in df.columns:
        try:
            dd = pd.to_datetime(df["Date"], errors="coerce")
        except Exception:
            dd = pd.to_datetime(pd.Series([], dtype="datetime64[ns]"))
        df["_Année_"]   = dd.dt.year
        df["_MoisNum_"] = dd.dt.month
        df["Mois"]      = dd.dt.month.apply(lambda m: f"{int(m):02d}" if pd.notna(m) else "")
    else:
        if "_Année_" not in df.columns:   df["_Année_"] = pd.NA
        if "_MoisNum_" not in df.columns: df["_MoisNum_"] = pd.NA
        if "Mois" not in df.columns:      df["Mois"] = ""
    return df

df_all = _ensure_time_features(df_clients_raw)

# ======= ONGLET : Dashboard =======
with tabs[1]:
    st.subheader("📊 Dashboard")
    if df_all is None or df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # Filtres haut de page
        cats  = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subs  = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []
        years = sorted(pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().astype(int).unique().tolist())

        a1, a2, a3, a4 = st.columns([1,1,1,1])
        fc = a1.multiselect("Catégories", cats, default=[], key=skey("dash","cats"))
        fs = a2.multiselect("Sous-catégories", subs, default=[], key=skey("dash","subs"))
        fv = a3.multiselect("Visa", visas, default=[], key=skey("dash","visas"))
        fy = a4.multiselect("Année", years, default=[], key=skey("dash","years"))

        view = df_all.copy()
        if fc: view = view[view["Categories"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]
        if fy: view = view[view["_Année_"].isin(fy)]

        # KPI (format compact)
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(view)}")
        total = (_to_num(view.get("Montant honoraires (US $)", 0)) + _to_num(view.get("Autres frais (US $)", 0))).sum()
        paye  = _to_num(view.get("Payé", 0)).sum()
        solde = _to_num(view.get("Solde", 0)).sum()
        env_pct = 0
        if "Dossiers envoyé" in view.columns and len(view)>0:
            env_pct = int(100 * (_to_num(view["Dossiers envoyé"]).clip(lower=0, upper=1).sum() / len(view)))
        k2.metric("Honoraires+Frais", _fmt_money(total))
        k3.metric("Payé", _fmt_money(paye))
        k4.metric("Solde", _fmt_money(solde))
        k5.metric("Envoyés (%)", f"{env_pct}%")

        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if not view.empty and "Categories" in view.columns:
            vc = view["Categories"].value_counts().rename_axis("Categorie").reset_index(name="Nombre")
            st.bar_chart(vc.set_index("Categorie"))

        st.markdown("#### 💵 Flux par mois")
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

        st.markdown("#### 📋 Détails (après filtres)")
        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Date","Mois","Categories","Sous-categorie","Visa",
            "Montant honoraires (US $)","Autres frais (US $)","Payé","Solde","Commentaires",
            "Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé","RFE"
        ] if c in view.columns]

        # Format monétaire lisible
        detail = view.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]:
            if c in detail.columns:
                detail[c] = _to_num(detail[c]).map(_fmt_money)
        if "Date" in detail.columns:
            try:
                detail["Date"] = pd.to_datetime(detail["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                detail["Date"] = detail["Date"].astype(str)

        sort_keys = [c for c in ["_Année_","_MoisNum_","Categories","Nom"] if c in detail.columns]
        detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail
        st.dataframe(detail_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=skey("dash","table"))


# ======= ONGLET : Analyses =======
with tabs[2]:
    st.subheader("📈 Analyses")
    if df_all is None or df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted(pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().astype(int).unique().tolist())
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(df_all["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_all.columns else []
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        b1, b2, b3, b4, b5 = st.columns(5)
        fy = b1.multiselect("Année", yearsA, default=[], key=skey("an","years"))
        fm = b2.multiselect("Mois (MM)", monthsA, default=[], key=skey("an","months"))
        fc = b3.multiselect("Catégories", catsA, default=[], key=skey("an","cats"))
        fs = b4.multiselect("Sous-catégories", subsA, default=[], key=skey("an","subs"))
        fv = b5.multiselect("Visa", visasA, default=[], key=skey("an","visas"))

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categories"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPI compacts
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Dossiers", f"{len(dfA)}")
        c2.metric("Honoraires", _fmt_money(_to_num(dfA.get("Montant honoraires (US $)",0)).sum()))
        c3.metric("Payé", _fmt_money(_to_num(dfA.get("Payé",0)).sum()))
        c4.metric("Solde", _fmt_money(_to_num(dfA.get("Solde",0)).sum()))

        # % par catégorie
        st.markdown("#### 📊 Répartition par catégorie (%)")
        if not dfA.empty and "Categories" in dfA.columns:
            total_cnt = max(1, len(dfA))
            rep = dfA["Categories"].value_counts().rename_axis("Categorie").reset_index(name="Nbr")
            rep["%"] = (rep["Nbr"] / total_cnt * 100).round(1)
            st.dataframe(rep, use_container_width=True, hide_index=True, key=skey("an","rep_cat"))

        # % par sous-catégorie
        st.markdown("#### 📊 Répartition par sous-catégorie (%)")
        if not dfA.empty and "Sous-categorie" in dfA.columns:
            total_cnt = max(1, len(dfA))
            rep2 = dfA["Sous-categorie"].value_counts().rename_axis("Sous-categorie").reset_index(name="Nbr")
            rep2["%"] = (rep2["Nbr"] / total_cnt * 100).round(1)
            st.dataframe(rep2, use_container_width=True, hide_index=True, key=skey("an","rep_sub"))

        # Comparaison période A vs période B (simple)
        st.markdown("#### 🔁 Comparaison deux périodes (A vs B)")
        ca1, ca2, ca3 = st.columns(3)
        pa_years = ca1.multiselect("Années (A)", yearsA, default=[], key=skey("cmp","ya"))
        pa_month = ca2.multiselect("Mois (A)", monthsA, default=[], key=skey("cmp","ma"))
        pa_cat   = ca3.multiselect("Catégories (A)", catsA, default=[], key=skey("cmp","ca"))

        cb1, cb2, cb3 = st.columns(3)
        pb_years = cb1.multiselect("Années (B)", yearsA, default=[], key=skey("cmp","yb"))
        pb_month = cb2.multiselect("Mois (B)", monthsA, default=[], key=skey("cmp","mb"))
        pb_cat   = cb3.multiselect("Catégories (B)", catsA, default=[], key=skey("cmp","cb"))

        def _slice(df, ys, ms, cs):
            s = df.copy()
            if ys: s = s[s["_Année_"].isin(ys)]
            if ms: s = s[s["Mois"].astype(str).isin(ms)]
            if cs: s = s[s["Categories"].astype(str).isin(cs)]
            return s

        A = _slice(df_all, pa_years, pa_month, pa_cat)
        B = _slice(df_all, pb_years, pb_month, pb_cat)

        def _kpis(df):
            return {
                "Dossiers": len(df),
                "Honoraires": _to_num(df.get("Montant honoraires (US $)",0)).sum(),
                "Payé": _to_num(df.get("Payé",0)).sum(),
                "Solde": _to_num(df.get("Solde",0)).sum(),
            }

        kA, kB = _kpis(A), _kpis(B)
        dcmp = pd.DataFrame({
            "KPI": ["Dossiers","Honoraires","Payé","Solde"],
            "A":   [kA["Dossiers"], kA["Honoraires"], kA["Payé"], kA["Solde"]],
            "B":   [kB["Dossiers"], kB["Honoraires"], kB["Payé"], kB["Solde"]],
            "Δ (B - A)": [kB["Dossiers"]-kA["Dossiers"], kB["Honoraires"]-kA["Honoraires"], kB["Payé"]-kA["Payé"], kB["Solde"]-kA["Solde"]],
        })
        # format monétaire
        for c in ["A","B","Δ (B - A)"]:
            dcmp.loc[1:3, c] = dcmp.loc[1:3, c].astype(float).map(_fmt_money)
        st.dataframe(dcmp, use_container_width=True, hide_index=True, key=skey("an","cmp_table"))


# ======= ONGLET : Escrow =======
with tabs[3]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all is None or df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Total (US $)"] = _to_num(dfE.get("Montant honoraires (US $)",0)) + _to_num(dfE.get("Autres frais (US $)",0))
        dfE["Payé"]  = _to_num(dfE.get("Payé",0))
        dfE["Solde"] = _to_num(dfE.get("Solde",0))

        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money(float(dfE["Total (US $)"].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Solde", _fmt_money(float(dfE["Solde"].sum())))

        st.markdown("#### Détails par catégorie")
        agg = dfE.groupby("Categories", as_index=False)[["Total (US $)","Payé","Solde"]].sum().sort_values("Total (US $)", ascending=False)
        st.dataframe(agg, use_container_width=True, key=skey("esc","agg"))

        st.caption("NB : si vous souhaitez un escrow strict, on peut isoler les honoraires encaissés avant « Dossiers envoyé » puis marquer les transferts une fois l’envoi effectué.")



# =========================
# Visa Manager — PARTIE 4/4
# 👤 Compte client • 🧾 Gestion • 📄 Visa (aperçu) • 💾 Export
# =========================

# --- petits helpers de robustesse (si non définis plus haut) ---
try:
    _ = _date_for_widget  # noqa
except NameError:
    def _date_for_widget(val):
        """Retourne un objet date sûr pour date_input."""
        if isinstance(val, (datetime, date)):
            return val if isinstance(val, date) else val.date()
        try:
            d = pd.to_datetime(val, errors="coerce")
            return d.date() if pd.notna(d) else date.today()
        except Exception:
            return date.today()

try:
    _ = skey  # noqa
except NameError:
    def skey(scope: str, name: str) -> str:
        rid = st.session_state.get("_sid", str(uuid.uuid4())[:8])
        st.session_state["_sid"] = rid
        return f"{scope}_{name}_{rid}"

# =========================
# 👤 ONGLET : Compte client
# =========================
with tabs[4]:
    st.subheader("👤 Compte client")

    if df_all is None or df_all.empty:
        st.info("Aucun client chargé.")
    else:
        # Sélection client par ID ou Nom
        left, right = st.columns(2)
        ids  = sorted(df_all["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_all.columns else []
        noms = sorted(df_all["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_all.columns else []

        sel_id  = left.selectbox("ID_Client", [""] + ids, index=0, key=skey("acct","id"))
        sel_nom = right.selectbox("Nom", [""] + noms, index=0, key=skey("acct","nm"))

        subset = df_all.copy()
        if sel_id:
            subset = subset[subset["ID_Client"].astype(str) == sel_id]
        elif sel_nom:
            subset = subset[subset["Nom"].astype(str) == sel_nom]

        if subset.empty:
            st.warning("Sélectionne un client pour afficher le compte.")
        else:
            row = subset.iloc[0].to_dict()

            # En-tête résumé
            st.markdown("#### 📌 Résumé")
            r1, r2, r3, r4 = st.columns(4)
            r1.metric("Dossier N", _safe_str(row.get("Dossier N","")))
            total = float(_to_num(row.get("Montant honoraires (US $)", 0)) + _to_num(row.get("Autres frais (US $)", 0)))
            r2.metric("Total", _fmt_money(total))
            r3.metric("Payé", _fmt_money(float(_to_num(row.get("Payé", 0)))))
            r4.metric("Solde", _fmt_money(float(_to_num(row.get("Solde", 0)))))

            # Détails dossier
            st.markdown("#### 🧾 Détails du dossier")
            d1, d2, d3 = st.columns(3)
            d1.write(f"**Catégorie :** {_safe_str(row.get('Categories',''))}")
            d1.write(f"**Sous-catégorie :** {_safe_str(row.get('Sous-categorie',''))}")
            d1.write(f"**Visa :** {_safe_str(row.get('Visa',''))}")
            d2.write(f"**Date :** {_safe_str(row.get('Date',''))}")
            d2.write(f"**Mois (MM) :** {_safe_str(row.get('Mois',''))}")
            d3.write(f"**Commentaires :** {_safe_str(row.get('Commentaires',''))}")

            # Statuts (les colonnes ne sont plus des cases à cocher mais des dates éventuelles)
            st.markdown("#### 🗂️ Statuts")
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

            s1.write(f"- **Dossier envoyé** : { 'Oui' if sdate(\"Date d'envoi\") else 'Non'} | Date : {sdate(\"Date d'envoi\")}")
            s1.write(f"- **Dossier approuvé** : { 'Oui' if sdate(\"Date d'acceptation\") else 'Non'} | Date : {sdate(\"Date d'acceptation\")}")
            s2.write(f"- **Dossier refusé** : { 'Oui' if sdate(\"Date de refus\") else 'Non'} | Date : {sdate(\"Date de refus\")}")
            s2.write(f"- **Dossier annulé** : { 'Oui' if sdate(\"Date d'annulation\") else 'Non'} | Date : {sdate(\"Date d'annulation\")}")
            rfeflag = int(_to_num(row.get("RFE", 0)) or 0)
            st.write(f"- **RFE** : {'Oui' if rfeflag else 'Non'}")

            # Mouvements financiers (simple : Acompte 1 / 2 si présents)
            st.markdown("#### 💳 Mouvements financiers")
            mvts = []
            if "Acompte 1" in row and _to_num(row["Acompte 1"]) > 0:
                mvts.append({"Libellé":"Acompte 1","Montant": float(_to_num(row["Acompte 1"]))})
            if "Acompte 2" in row and _to_num(row["Acompte 2"]) > 0:
                mvts.append({"Libellé":"Acompte 2","Montant": float(_to_num(row["Acompte 2"]))})
            if mvts:
                dfm = pd.DataFrame(mvts)
                dfm["Montant"] = dfm["Montant"].map(_fmt_money)
                st.dataframe(dfm, use_container_width=True, hide_index=True, key=skey("acct","mvts"))
            else:
                st.caption("Aucun acompte enregistré dans le fichier (colonnes « Acompte 1 » / « Acompte 2 »).")

# =========================
# 🧾 ONGLET : Gestion (CRUD)
# =========================
with tabs[5]:
    st.subheader("🧾 Gestion (Ajouter / Modifier / Supprimer)")
    df_live = df_all.copy() if df_all is not None else pd.DataFrame()

    if df_live.empty:
        st.info("Aucun client à gérer (charge un fichier Clients).")
    else:
        op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=skey("crud","op"))

        # listes de référence pour la cascade
        cats = sorted(df_visa_raw["Categories"].dropna().astype(str).unique().tolist()) if ("Categories" in df_visa_raw.columns and not df_visa_raw.empty) else sorted(df_live["Categories"].dropna().astype(str).unique().tolist())
        def subs_for(cat):
            if "Categories" in df_visa_raw.columns and "Sous-categorie" in df_visa_raw.columns:
                return sorted(df_visa_raw[df_visa_raw["Categories"].astype(str)==cat]["Sous-categorie"].dropna().astype(str).unique().tolist())
            return sorted(df_live[df_live["Categories"].astype(str)==cat]["Sous-categorie"].dropna().astype(str).unique().tolist())

        # -------- AJOUT --------
        if op == "Ajouter":
            st.markdown("### ➕ Ajouter")
            c1, c2, c3 = st.columns(3)
            nom  = c1.text_input("Nom", "", key=skey("add","nom"))
            dval = _date_for_widget(date.today())
            dt   = c2.date_input("Date de création", value=dval, key=skey("add","date"))
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=int(dval.month)-1, key=skey("add","mois"))

            st.markdown("#### 🎯 Visa")
            v1, v2, v3 = st.columns(3)
            cat = v1.selectbox("Catégorie", [""]+cats, index=0, key=skey("add","cat"))
            sub = ""
            if cat:
                subs = subs_for(cat)
                sub = v2.selectbox("Sous-catégorie", [""]+subs, index=0, key=skey("add","sub"))
            visa_val = v3.text_input("Visa (libre ou dérivé)", sub if sub else "", key=skey("add","visa"))

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=skey("add","h"))
            other = f2.number_input("Autres frais (US $)",      min_value=0.0, value=0.0, step=20.0, format="%.2f", key=skey("add","o"))
            acomp1 = st.number_input("Acompte 1", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=skey("add","a1"))
            acomp2 = st.number_input("Acompte 2", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=skey("add","a2"))
            comm   = st.text_area("Commentaires", "", key=skey("add","com"))

            st.markdown("#### 🗂️ Statuts (dates)")
            s1, s2 = st.columns(2)
            sent_d = s1.date_input("Date d'envoi", value=None, key=skey("add","sentd"))
            acc_d  = s1.date_input("Date d'acceptation", value=None, key=skey("add","accd"))
            ref_d  = s2.date_input("Date de refus", value=None, key=skey("add","refd"))
            ann_d  = s2.date_input("Date d'annulation", value=None, key=skey("add","annd"))
            rfe    = st.checkbox("RFE", value=False, key=skey("add","rfe"))

            if st.button("💾 Enregistrer", key=skey("add","save")):
                if not nom or not cat or not sub:
                    st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                    st.stop()
                total = float(honor) + float(other)
                paye  = float(acomp1) + float(acomp2)
                solde = max(0.0, total - paye)
                # ID_Client & Dossier N (simples)
                new_id = f"{_norm(nom)}-{int(datetime.now().timestamp())}"
                new_dossier = int(df_live.get("Dossier N", pd.Series([13056])).astype(str).str.extract(r"(\d+)").fillna(13056).astype(int).max()) + 1

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
                    "Date d'envoi": sent_d,
                    "Date d'acceptation": acc_d,
                    "Date de refus": ref_d,
                    "Date d'annulation": ann_d,
                    "RFE": 1 if rfe else 0,
                }
                df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                # sauvegarde côté appli (fichier « Clients courant »)
                if clients_src and clients_src.get("type") == "file":
                    try:
                        with pd.ExcelWriter(clients_src["path"], engine="openpyxl") as w:
                            df_new.to_excel(w, index=False)
                        st.success("Client ajouté et fichier mis à jour.")
                    except Exception as e:
                        st.error(f"Impossible d'écrire le fichier : {e}")
                st.cache_data.clear()
                st.rerun()

        # -------- MODIFIER --------
        if op == "Modifier":
            st.markdown("### ✏️ Modifier")
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            m1, m2 = st.columns(2)
            target_name = m1.selectbox("Nom", [""]+names, index=0, key=skey("mod","nom"))
            target_id   = m2.selectbox("ID_Client", [""]+ids, index=0, key=skey("mod","id"))

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
            nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=skey("mod","nomv"))
            dval = _date_for_widget(row.get("Date"))
            dt   = d2.date_input("Date de création", value=dval, key=skey("mod","date"))
            mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=max(0, int(_safe_str(row.get("Mois","01"))) - 1),
                                key=skey("mod","mois"))

            st.markdown("#### 🎯 Visa")
            v1, v2, v3 = st.columns(3)
            preset_cat = _safe_str(row.get("Categories",""))
            cat = v1.selectbox("Catégorie", [""]+cats,
                               index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                               key=skey("mod","cat"))
            sub = _safe_str(row.get("Sous-categorie",""))
            if cat:
                subs = subs_for(cat)
                sub = v2.selectbox("Sous-catégorie", [""]+subs,
                                   index=(subs.index(sub)+1 if sub in subs else 0),
                                   key=skey("mod","sub"))
            visa_val = v3.text_input("Visa (libre ou dérivé)", _safe_str(row.get("Visa","")), key=skey("mod","visa"))

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0,
                                    value=float(_to_num(row.get("Montant honoraires (US $)", 0))),
                                    step=50.0, format="%.2f", key=skey("mod","h"))
            other = f2.number_input("Autres frais (US $)", min_value=0.0,
                                    value=float(_to_num(row.get("Autres frais (US $)", 0))),
                                    step=20.0, format="%.2f", key=skey("mod","o"))
            acomp1 = st.number_input("Acompte 1", min_value=0.0,
                                     value=float(_to_num(row.get("Acompte 1", 0))),
                                     step=10.0, format="%.2f", key=skey("mod","a1"))
            acomp2 = st.number_input("Acompte 2", min_value=0.0,
                                     value=float(_to_num(row.get("Acompte 2", 0))),
                                     step=10.0, format="%.2f", key=skey("mod","a2"))
            comm   = st.text_area("Commentaires", _safe_str(row.get("Commentaires","")), key=skey("mod","com"))

            st.markdown("#### 🗂️ Statuts (dates)")
            s1, s2 = st.columns(2)
            sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=skey("mod","sentd"))
            acc_d  = s1.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=skey("mod","accd"))
            ref_d  = s2.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=skey("mod","refd"))
            ann_d  = s2.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=skey("mod","annd"))
            rfe    = st.checkbox("RFE", value=bool(int(_to_num(row.get("RFE", 0)) or 0)), key=skey("mod","rfe"))

            if st.button("💾 Enregistrer les modifications", key=skey("mod","save")):
                if not nom or not cat or not sub:
                    st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                    st.stop()
                total = float(honor) + float(other)
                paye  = float(acomp1) + float(acomp2)
                solde = max(0.0, total - paye)

                df_live.at[idx, "Nom"]  = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(mois):02d}"
                df_live.at[idx, "Categories"] = cat
                df_live.at[idx, "Sous-categorie"] = sub
                df_live.at[idx, "Visa"] = visa_val
                df_live.at[idx, "Montant honoraires (US $)"] = float(honor)
                df_live.at[idx, "Autres frais (US $)"]       = float(other)
                df_live.at[idx, "Acompte 1"] = float(acomp1)
                df_live.at[idx, "Acompte 2"] = float(acomp2)
                df_live.at[idx, "Payé"]      = float(paye)
                df_live.at[idx, "Solde"]     = float(solde)
                df_live.at[idx, "Commentaires"] = comm
                df_live.at[idx, "Date d'envoi"]       = sent_d
                df_live.at[idx, "Date d'acceptation"] = acc_d
                df_live.at[idx, "Date de refus"]      = ref_d
                df_live.at[idx, "Date d'annulation"]  = ann_d
                df_live.at[idx, "RFE"] = 1 if rfe else 0

                if clients_src and clients_src.get("type") == "file":
                    try:
                        with pd.ExcelWriter(clients_src["path"], engine="openpyxl") as w:
                            df_live.to_excel(w, index=False)
                        st.success("Modifications enregistrées.")
                    except Exception as e:
                        st.error(f"Écriture fichier impossible : {e}")
                st.cache_data.clear()
                st.rerun()

        # -------- SUPPRIMER --------
        if op == "Supprimer":
            st.markdown("### 🗑️ Supprimer")
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            d1, d2 = st.columns(2)
            target_name = d1.selectbox("Nom", [""]+names, index=0, key=skey("del","nom"))
            target_id   = d2.selectbox("ID_Client", [""]+ids, index=0, key=skey("del","id"))

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({"Dossier N": row.get("Dossier N",""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=skey("del","btn")):
                    df_new = df_live[~mask].copy()
                    if clients_src and clients_src.get("type") == "file":
                        try:
                            with pd.ExcelWriter(clients_src["path"], engine="openpyxl") as w:
                                df_new.to_excel(w, index=False)
                            st.success("Client supprimé.")
                        except Exception as e:
                            st.error(f"Écriture fichier impossible : {e}")
                    st.cache_data.clear()
                    st.rerun()

# =========================
# 📄 ONGLET : Visa (aperçu)
# =========================
with tabs[6]:
    st.subheader("📄 Visa — aperçu")
    if df_visa_raw is None or df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True, key=skey("visa","view"))

# =========================
# 💾 ONGLET : Export
# =========================
with tabs[7]:
    st.subheader("💾 Export")
    st.caption("Exporte les jeux de données courants tels que chargés/édités dans l'application.")

    colx, coly = st.columns(2)

    # Export Clients
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
                key=skey("exp","clients"),
            )

    # Export Visa
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
                key=skey("exp","visa"),
            )




# ==========================
# 🛂 Visa Manager – PARTIE 2
# ==========================

# --------- Barre latérale : chargement / mémorisation chemins ----------
st.sidebar.header("📂 Fichiers")
mode = st.sidebar.radio("Mode de chargement", ["Un fichier (Clients)", "Deux fichiers (Clients & Visa)"], index=1, key=f"mode_{SID}")

last_clients, last_visa, last_save = load_last_paths()

up_clients = st.sidebar.file_uploader("Clients (xlsx/csv)", type=["xlsx","xls","csv"], key=f"up_clients_{SID}")
up_visa    = None
if mode == "Deux fichiers (Clients & Visa)":
    up_visa = st.sidebar.file_uploader("Visa (xlsx/csv)", type=["xlsx","xls","csv"], key=f"up_visa_{SID}")

st.sidebar.caption("**Derniers chemins mémorisés :**")
st.sidebar.code(f"Dernier Clients : {last_clients or '(non défini)'}\nDernier Visa    : {last_visa or '(non défini)'}", language="bash")

save_dir = st.sidebar.text_input("Chemin de sauvegarde (PC / Drive / OneDrive)", value=last_save or "", key=f"save_dir_{SID}")
if st.sidebar.button("💾 Mémoriser ces chemins", key=f"btn_mem_{SID}"):
    cpath = last_clients
    vpath = last_visa
    # si on vient d’uploader, on stocke les noms upload_*
    if up_clients is not None:
        cpath = f"upload_{up_clients.name}"
    if up_visa is not None:
        vpath = f"upload_{up_visa.name}"
    save_last_paths(cpath, vpath, save_dir)
    st.sidebar.success("Chemins mémorisés.")

st.title("🛂 Visa Manager")

# --------- Lecture données selon mode ----------
def resolve_clients_source():
    # priorité à l’upload courant
    if up_clients is not None:
        return up_clients
    # sinon réutilise dernier chemin si dispo
    if last_clients and Path(last_clients).exists():
        return last_clients
    return None

def resolve_visa_source():
    if mode == "Deux fichiers (Clients & Visa)":
        if up_visa is not None:
            return up_visa
        if last_visa and Path(last_visa).exists():
            return last_visa
        return None
    else:
        # un seul fichier : on lit l’onglet Visa du même fichier si possible
        src = resolve_clients_source()
        if src is None or hasattr(src, "read"):  # UploadedFile sans multi-sheet direct
            return None
        # on utilisera read_sheet_from_path
        return src

clients_src = resolve_clients_source()
visa_src    = resolve_visa_source()

df_clients_raw = read_any_table(clients_src)
if df_clients_raw is None and isinstance(clients_src, (str, Path)):
    # Essaye onglet Clients si xlsx
    df_clients_raw = read_sheet_from_path(clients_src, SHEET_CLIENTS)

df_visa_raw = read_any_table(visa_src)
if df_visa_raw is None and isinstance(visa_src, (str, Path)):
    df_visa_raw = read_sheet_from_path(visa_src, SHEET_VISA)

# Normalisation clients
if df_clients_raw is not None and not df_clients_raw.empty:
    df_clients = _normalize_columns(df_clients_raw.copy())
else:
    df_clients = pd.DataFrame(columns=REQUIRED_CLIENT_COLS)

# Carte Visa
visa_map = build_visa_map(df_visa_raw.copy()) if df_visa_raw is not None and not df_visa_raw.empty else {}

# Calculs de base
df_all = df_clients.copy()
df_all["Total (US $)"] = df_all["Montant honoraires (US $)"] + df_all["Autres frais (US $)"]
# si "Solde" vide, on recalcule
mask_na_solde = pd.to_numeric(df_all["Solde"], errors="coerce").isna()
df_all.loc[mask_na_solde, "Solde"] = (df_all["Total (US $)"] - df_all["Payé"]).clip(lower=0.0)

# --------- Tabs ----------
tabs = st.tabs([
    "📄 Fichiers chargés",
    "📊 Dashboard",
    "📈 Analyses",
    "🏦 Escrow",
    "👤 Compte client",
    "🧾 Gestion",
    "📄 Visa (aperçu)",
    "💾 Export",
])

# ============= ONGLET 0 : Fichiers chargés =============
with tabs[0]:
    st.subheader("📄 Fichiers chargés")
    st.write("**Clients** :", f"`{getattr(up_clients, 'name', '') or last_clients or '(aucun)'}`")
    if mode == "Deux fichiers (Clients & Visa)":
        st.write("**Visa** :", f"`{getattr(up_visa, 'name', '') or last_visa or '(aucun)'}`")
    else:
        st.write("**Visa** : (déduit depuis le fichier Clients)")

    st.write(f"Nb clients : **{len(df_all)}**")
    if df_visa_raw is not None:
        st.write(f"Nb lignes Visa : **{len(df_visa_raw)}**")

# ============= ONGLET 1 : Dashboard =============
with tabs[1]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Aucun client chargé. Charge les fichiers dans la barre latérale.")
    else:
        # KPI (compacts)
        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        k1.metric("Dossiers", f"{len(df_all)}")
        k2.metric("Honoraires+Frais", _fmt_money(float((df_all["Total (US $)"]).sum())))
        k3.metric("Payé", _fmt_money(float((df_all["Payé"]).sum())))
        k4.metric("Solde", _fmt_money(float((df_all["Solde"]).sum())))
        pct_envoye = 0.0
        try:
            num = int(pd.to_numeric(df_all["Dossiers envoyé"], errors="coerce").fillna(0).astype(int).sum())
            pct_envoye = (num / max(1,len(df_all))) * 100.0
        except Exception:
            pass
        k5.metric("Envoyés (%)", f"{pct_envoye:.0f}%")

        # Filtres
        st.markdown("#### 🎛️ Filtres")
        cats  = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist())
        subs  = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist())
        visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist())

        a1,a2,a3 = st.columns(3)
        fc = a1.multiselect("Catégories", cats, default=[])
        fs = a2.multiselect("Sous-catégories", subs, default=[])
        fv = a3.multiselect("Visa", visas, default=[])

        view = df_all.copy()
        if fc: view = view[view["Categorie"].astype(str).isin(fc)]
        if fs: view = view[view["Sous-categorie"].astype(str).isin(fs)]
        if fv: view = view[view["Visa"].astype(str).isin(fv)]

        # Graph 1 : Nombre dossiers par catégorie
        st.markdown("#### 📦 Nombre de dossiers par catégorie")
        if not view.empty and "Categorie" in view.columns:
            vc = view["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie","Nombre"]
            st.bar_chart(vc.set_index("Categorie"))
        else:
            st.info("Aucune donnée avec ces filtres.")

        # Graph 2 : flux mensuels
        st.markdown("#### 💵 Flux par mois")
        g = view.copy()
        if not g.empty:
            g["Mois"] = g["Mois"].astype(str)
            gm = g.groupby("Mois", as_index=False)[["Montant honoraires (US $)","Autres frais (US $)","Payé","Solde"]].sum().sort_values("Mois")
            st.line_chart(gm.set_index("Mois"))
        else:
            st.info("Aucune donnée pour la courbe mensuelle.")

        # Détails
        st.markdown("#### 📋 Détails (après filtres)")
        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé","RFE",
            "Date d'envoi","Date d'acceptation","Date de refus","Date d'annulation",
        ] if c in view.columns]
        sort_keys = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
        view_sorted = view.sort_values(by=sort_keys) if sort_keys else view
        st.dataframe(view_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"dash_detail_{SID}")



# ==========================
# 🛂 Visa Manager – PARTIE 3
# ==========================

# ============= ONGLET 2 : Analyses =============
with tabs[2]:
    st.subheader("📈 Analyses")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        yearsA  = sorted(pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().astype(int).unique().tolist())
        monthsA = [f"{m:02d}" for m in range(1,13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist())
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist())
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist())

        b1,b2,b3,b4,b5 = st.columns(5)
        fy = b1.multiselect("Année", yearsA, default=[])
        fm = b2.multiselect("Mois (MM)", monthsA, default=[])
        fc = b3.multiselect("Catégorie", catsA, default=[])
        fs = b4.multiselect("Sous-catégorie", subsA, default=[])
        fv = b5.multiselect("Visa", visasA, default=[])

        A = df_all.copy()
        if fy: A = A[A["_Année_"].isin(fy)]
        if fm: A = A[A["Mois"].astype(str).isin(fm)]
        if fc: A = A[A["Categorie"].astype(str).isin(fc)]
        if fs: A = A[A["Sous-categorie"].astype(str).isin(fs)]
        if fv: A = A[A["Visa"].astype(str).isin(fv)]

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Dossiers", f"{len(A)}")
        c2.metric("Honoraires", _fmt_money(float(A["Montant honoraires (US $)"].sum())))
        c3.metric("Payé", _fmt_money(float(A["Payé"].sum())))
        c4.metric("Solde", _fmt_money(float(A["Solde"].sum())))

        st.markdown("#### 📦 Dossiers par catégorie")
        if not A.empty:
            vc = A["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie","Nombre"]
            st.bar_chart(vc.set_index("Categorie"))
        else:
            st.info("Aucune donnée.")

        st.markdown("#### 🧾 Détail filtré")
        show = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Solde",
            "Dossiers envoyé","Dossier approuvé","Dossier refusé","Dossier Annulé","RFE",
        ] if c in A.columns]
        st.dataframe(A[show].reset_index(drop=True), use_container_width=True, key=f"ana_detail_{SID}")

# ============= ONGLET 3 : Escrow (synthèse) =============
with tabs[3]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Total (US $)"] = dfE["Montant honoraires (US $)"] + dfE["Autres frais (US $)"]
        agg = dfE.groupby("Categorie", as_index=False)[["Total (US $)","Payé","Solde"]].sum()
        st.dataframe(agg, use_container_width=True)
        t1,t2,t3 = st.columns(3)
        t1.metric("Total", _fmt_money(float(dfE["Total (US $)"].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Solde", _fmt_money(float(dfE["Solde"].sum())))
        st.caption("NB : si tu veux un escrow strict (honoraires avant envoi → transfert après « Dossier envoyé »), on pourra isoler ces flux.")

# ============= ONGLET 4 : Compte client (consultation) =============
with tabs[4]:
    st.subheader("👤 Compte client")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        names = sorted(df_all["Nom"].dropna().astype(str).unique().tolist())
        sel_name = st.selectbox("Sélectionne un client", [""]+names, index=0)
        if sel_name:
            sub = df_all[df_all["Nom"].astype(str)==sel_name].copy()
            st.write(f"**Dossiers trouvés :** {len(sub)}")
            for _, row in sub.iterrows():
                st.markdown("---")
                s1,s2,s3 = st.columns([2,2,2])
                s1.write(f"**Dossier N** : {_safe_str(row.get('Dossier N',''))}")
                s1.write(f"**ID_Client** : {_safe_str(row.get('ID_Client',''))}")
                s1.write(f"**Nom** : {_safe_str(row.get('Nom',''))}")
                s2.write(f"**Catégorie** : {_safe_str(row.get('Categorie',''))}")
                s2.write(f"**Sous-catégorie** : {_safe_str(row.get('Sous-categorie',''))}")
                s2.write(f"**Visa** : {_safe_str(row.get('Visa',''))}")
                s3.write(f"**Date** : {_safe_str(row.get('Date',''))}")
                s3.write(f"**Mois** : {_safe_str(row.get('Mois',''))}")

                h = _to_float(row.get("Montant honoraires (US $)",0))
                o = _to_float(row.get("Autres frais (US $)",0))
                p = _to_float(row.get("Payé",0))
                s = _to_float(row.get("Solde",0))

                k1,k2,k3,k4 = st.columns(4)
                k1.metric("Honoraires", _fmt_money(h))
                k2.metric("Autres frais", _fmt_money(o))
                k3.metric("Payé", _fmt_money(p))
                k4.metric("Solde", _fmt_money(s))

                # Statuts (lecture)
                st.markdown("**Statuts**")
                def sdate(col):
                    return _safe_str(row.get(col, ""))

                c1,c2 = st.columns(2)
                with c1:
                    st.write(f"- **Dossier envoyé** : {'Oui' if int(_to_float(row.get('Dossiers envoyé',0)))==1 else 'Non'}"
                             f" | Date : {sdate(\"Date d'envoi\")}")
                    st.write(f"- **Dossier approuvé** : {'Oui' if int(_to_float(row.get('Dossier approuvé',0)))==1 else 'Non'}"
                             f" | Date : {sdate('Date d\\'acceptation')}")
                with c2:
                    st.write(f"- **Dossier refusé** : {'Oui' if int(_to_float(row.get('Dossier refusé',0)))==1 else 'Non'}"
                             f" | Date : {sdate('Date de refus')}")
                    st.write(f"- **Dossier annulé** : {'Oui' if int(_to_float(row.get('Dossier Annulé',0)))==1 else 'Non'}"
                             f" | Date : {sdate('Date d\\'annulation')}")
                st.write(f"- **RFE** : {'Oui' if int(_to_float(row.get('RFE',0)))==1 else 'Non'}")

# ============= ONGLET 6 : Visa (aperçu) =============
with tabs[6]:
    st.subheader("📄 Visa (aperçu)")
    if df_visa_raw is None or df_visa_raw.empty:
        st.info("Aucune table Visa chargée.")
    else:
        st.dataframe(df_visa_raw, use_container_width=True)
        # mini navigateur Catégorie → Sous-catégorie → options
        st.markdown("#### 🎯 Sélecteur Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""]+cats, index=0, key=f"v_cat_{SID}")
        opts = []
        if sel_cat:
            subs = sorted(list(visa_map[sel_cat].keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"v_sub_{SID}")
            if sel_sub:
                info = visa_map[sel_cat][sel_sub]
                opts = info.get("options", [])
        if opts:
            st.write("Options disponibles :", ", ".join(opts))
        else:
            if sel_cat:
                st.info("Pas d'options pour cette sous-catégorie.")



# ==========================
# 🛂 Visa Manager – PARTIE 4
# ==========================

# ============= ONGLET 5 : Gestion (CRUD) =============
with tabs[5]:
    st.subheader("🧾 Gestion (Ajouter / Modifier / Supprimer)")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        op = st.radio("Action", ["Ajouter","Modifier","Supprimer"], horizontal=True, key=f"crud_{SID}")

        if op == "Ajouter":
            st.markdown("### ➕ Ajouter")
            c1,c2,c3 = st.columns(3)
            nom  = c1.text_input("Nom", "", key=f"add_nom_{SID}")
            dval = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=date.today().month-1, key=f"add_mois_{SID}")

            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            sel_cat = st.selectbox("Catégorie", [""]+cats, index=0, key=f"add_cat_{SID}")
            sel_sub = ""
            selected_opts = []
            if sel_cat:
                subs = sorted(list(visa_map[sel_cat].keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""]+subs, index=0, key=f"add_sub_{SID}")
                if sel_sub:
                    opts = visa_map[sel_cat][sel_sub].get("options", [])
                    selected_opts = st.multiselect("Options", opts, default=[], key=f"add_opts_{SID}")

            f1,f2,f3 = st.columns(3)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"add_h_{SID}")
            other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=f"add_o_{SID}")
            comment = f3.text_input("Commentaires (autres frais)", "", key=f"add_c_{SID}")

            st.markdown("#### 📌 Statuts initiaux")
            s1,s2,s3,s4,s5 = st.columns(5)
            sent = s1.checkbox("Dossier envoyé", key=f"add_sent_{SID}")
            sent_d = s1.date_input("Date d'envoi", value=None, key=f"add_sentd_{SID}")
            acc = s2.checkbox("Dossier approuvé", key=f"add_acc_{SID}")
            acc_d = s2.date_input("Date d'acceptation", value=None, key=f"add_accd_{SID}")
            ref = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
            ref_d = s3.date_input("Date de refus", value=None, key=f"add_refd_{SID}")
            ann = s4.checkbox("Dossier Annulé", key=f"add_ann_{SID}")
            ann_d = s4.date_input("Date d'annulation", value=None, key=f"add_annd_{SID}")
            rfe = s5.checkbox("RFE", key=f"add_rfe_{SID}")
            if rfe and not any([sent,acc,ref,ann]):
                st.warning("RFE doit être associé à un des autres statuts.")

            if st.button("💾 Enregistrer le client", key=f"add_btn_{SID}"):
                if not nom or not sel_cat or not sel_sub:
                    st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                    st.stop()
                total = float(honor) + float(other)
                paye  = 0.0
                solde = max(0.0, total - paye)
                did   = make_client_id(nom, dval)
                dossier_no = next_dossier(df_all)

                new_row = {
                    "Dossier N": dossier_no,
                    "ID_Client": did,
                    "Nom": nom,
                    "Date": dval,
                    "Mois": f"{int(date(dval.year,dval.month,1).month):02d}",
                    "Categorie": sel_cat,
                    "Sous-categorie": sel_sub,
                    "Visa": (sel_sub + (" " + " ".join(selected_opts) if selected_opts else "")),
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Payé": 0.0,
                    "Solde": solde,
                    "Acompte 1": 0.0,
                    "Acompte 2": 0.0,
                    "Commentaires": comment,
                    "Dossiers envoyé": 1 if sent else 0,
                    "Dossier approuvé": 1 if acc else 0,
                    "Dossier refusé": 1 if ref else 0,
                    "Dossier Annulé": 1 if ann else 0,
                    "RFE": 1 if rfe else 0,
                    "Date d'envoi": sent_d,
                    "Date d'acceptation": acc_d,
                    "Date de refus": ref_d,
                    "Date d'annulation": ann_d,
                }
                df_all = pd.concat([df_all, pd.DataFrame([new_row])], ignore_index=True)
                st.success("Client ajouté (en mémoire). Utilise l’onglet Export pour sauver.")

        elif op == "Modifier":
            st.markdown("### ✏️ Modifier")
            if df_all.empty:
                st.info("Aucun client.")
            else:
                ids = df_all["ID_Client"].dropna().astype(str).unique().tolist()
                tid = st.selectbox("Choisir ID_Client", [""]+sorted(ids), index=0, key=f"mod_id_{SID}")
                if tid:
                    idx = df_all[df_all["ID_Client"].astype(str)==tid].index
                    if len(idx)==0:
                        st.warning("ID introuvable.")
                        st.stop()
                    i = idx[0]
                    row = df_all.loc[i].copy()

                    d1,d2,d3 = st.columns(3)
                    nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nom_{SID}")
                    dval = _date_for_widget(row.get("Date"))
                    dt   = d2.date_input("Date de création", value=dval or date.today(), key=f"mod_date_{SID}")
                    mdef = _safe_str(row.get("Mois","01"))
                    try:
                        im = max(1, min(12, int(mdef)))
                    except Exception:
                        im = 1
                    mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=im-1, key=f"mod_mois_{SID}")

                    st.markdown("#### 🎯 Choix Visa")
                    cats = sorted(list(visa_map.keys()))
                    preset_cat = _safe_str(row.get("Categorie",""))
                    cat = st.selectbox("Catégorie", [""]+cats, index=(cats.index(preset_cat)+1 if preset_cat in cats else 0), key=f"mod_cat_{SID}")

                    preset_sub = _safe_str(row.get("Sous-categorie",""))
                    ss = ""
                    if cat:
                        subs = sorted(list(visa_map[cat].keys()))
                        ss = st.selectbox("Sous-catégorie", [""]+subs, index=(subs.index(preset_sub)+1 if preset_sub in subs else 0), key=f"mod_sub_{SID}")

                    selected_opts = []
                    if cat and ss:
                        opts = visa_map[cat][ss].get("options", [])
                        # essaie de précocher ce qui figure déjà dans la chaîne Visa
                        already = _safe_str(row.get("Visa",""))
                        pre = [o for o in opts if o in already]
                        selected_opts = st.multiselect("Options", opts, default=pre, key=f"mod_opts_{SID}")

                    f1,f2,f3 = st.columns(3)
                    honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=float(_to_float(row.get("Montant honoraires (US $)",0))), step=50.0, format="%.2f", key=f"mod_h_{SID}")
                    other = f2.number_input("Autres frais (US $)", min_value=0.0, value=float(_to_float(row.get("Autres frais (US $)",0))), step=10.0, format="%.2f", key=f"mod_o_{SID}")
                    comment = f3.text_input("Commentaires (autres frais)", _safe_str(row.get("Commentaires","")), key=f"mod_c_{SID}")

                    st.markdown("#### 📌 Statuts")
                    s1,s2,s3,s4,s5 = st.columns(5)
                    sent = s1.checkbox("Dossier envoyé", value=bool(int(_to_float(row.get("Dossiers envoyé",0)))==1), key=f"mod_sent_{SID}")
                    sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
                    acc = s2.checkbox("Dossier approuvé", value=bool(int(_to_float(row.get("Dossier approuvé",0)))==1), key=f"mod_acc_{SID}")
                    acc_d = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
                    ref = s3.checkbox("Dossier refusé", value=bool(int(_to_float(row.get("Dossier refusé",0)))==1), key=f"mod_ref_{SID}")
                    ref_d = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=f"mod_refd_{SID}")
                    ann = s4.checkbox("Dossier Annulé", value=bool(int(_to_float(row.get("Dossier Annulé",0)))==1), key=f"mod_ann_{SID}")
                    ann_d = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
                    rfe = s5.checkbox("RFE", value=bool(int(_to_float(row.get("RFE",0)))==1), key=f"mod_rfe_{SID}")

                    if st.button("💾 Enregistrer", key=f"mod_save_{SID}"):
                        df_all.at[i,"Nom"] = nom
                        df_all.at[i,"Date"] = dt
                        df_all.at[i,"Mois"] = mois
                        df_all.at[i,"Categorie"] = cat
                        df_all.at[i,"Sous-categorie"] = ss
                        df_all.at[i,"Visa"] = (ss + (" " + " ".join(selected_opts) if selected_opts else "")) if ss else _safe_str(row.get("Visa",""))
                        df_all.at[i,"Montant honoraires (US $)"] = float(honor)
                        df_all.at[i,"Autres frais (US $)"] = float(other)
                        df_all.at[i,"Total (US $)"] = float(honor)+float(other)
                        # recalcul solde (en gardant le payé existant)
                        paye = float(_to_float(row.get("Payé",0)))
                        df_all.at[i,"Payé"] = paye
                        df_all.at[i,"Solde"] = max(0.0, float(honor)+float(other)-paye)
                        df_all.at[i,"Commentaires"] = comment

                        df_all.at[i,"Dossiers envoyé"] = 1 if sent else 0
                        df_all.at[i,"Date d'envoi"] = sent_d
                        df_all.at[i,"Dossier approuvé"] = 1 if acc else 0
                        df_all.at[i,"Date d'acceptation"] = acc_d
                        df_all.at[i,"Dossier refusé"] = 1 if ref else 0
                        df_all.at[i,"Date de refus"] = ref_d
                        df_all.at[i,"Dossier Annulé"] = 1 if ann else 0
                        df_all.at[i,"Date d'annulation"] = ann_d
                        df_all.at[i,"RFE"] = 1 if rfe else 0

                        st.success("Modifié (en mémoire). Utilise Export pour sauver.")

        elif op == "Supprimer":
            st.markdown("### 🗑️ Supprimer")
            ids = df_all["ID_Client"].dropna().astype(str).unique().tolist()
            tid = st.selectbox("Choisir ID_Client", [""]+sorted(ids), index=0, key=f"del_id_{SID}")
            if tid:
                mask = (df_all["ID_Client"].astype(str)==tid)
                if mask.any():
                    row = df_all[mask].iloc[0]
                    st.write({"Dossier N": row.get("Dossier N",""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                    if st.button("❗ Confirmer la suppression", key=f"del_btn_{SID}"):
                        df_all = df_all[~mask].copy()
                        st.success("Supprimé (en mémoire). Utilise Export pour sauver.")

# ============= ONGLET 7 : Export =============
with tabs[7]:
    st.subheader("💾 Export")
    st.caption("Export en **xlsx** (1 onglet Clients + 1 onglet Visa si dispo).")
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        df_all.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        if df_visa_raw is not None and not df_visa_raw.empty:
            df_visa_raw.to_excel(wr, sheet_name=SHEET_VISA, index=False)
    xlsx_data = buf.getvalue()

    st.download_button(
        "⬇️ Télécharger Excel",
        data=xlsx_data,
        file_name=f"VisaManager_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_all_{SID}",
        use_container_width=True
    )

    # Sauvegarde locale si chemin défini
    if save_dir:
        pdir = Path(save_dir)
        _ensure_dir(pdir)
        if st.button("💽 Sauver sur le dossier choisi", key=f"btn_save_local_{SID}"):
            try:
                outp = pdir / f"VisaManager_{date.today().isoformat()}.xlsx"
                with open(outp, "wb") as f:
                    f.write(xlsx_data)
                st.success(f"Fichier sauvegardé : {outp}")
                # mémorise ces chemins
                cpath = getattr(up_clients, "name", "") and f"upload_{up_clients.name}" or last_clients
                vpath = (getattr(up_visa, "name", "") and f"upload_{up_visa.name}") if (mode=="Deux fichiers (Clients & Visa)") else last_visa
                save_last_paths(cpath, vpath, str(pdir))
            except Exception as e:
                st.error(f"Erreur de sauvegarde : {e}"