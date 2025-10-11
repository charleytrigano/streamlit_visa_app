from __future__ import annotations

import streamlit as st
import pandas as pd
import json
from datetime import date, datetime
from pathlib import Path
import unicodedata
import openpyxl
import altair as alt

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

# ------------------------------------------------------------
# CONSTANTES
# ------------------------------------------------------------
CLIENTS_FILE_DEFAULT = "donnees_visa_clients1_adapte.xlsx"
VISA_FILE_DEFAULT    = "donnees_visa_clients1.xlsx"
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

HONO  = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"

CLIENTS_COLS = [
    "Dossier N","ID_Client","Nom","Date","Mois",
    "Categorie","Sous-categorie","Visa",
    HONO, AUTRE, TOTAL,
    "Payé","Reste","Paiements","Options",
    "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"
]

TOGGLE_COLUMNS = {
    "AOS","CP","USCIS","I-130","I-140","I-140 & AOS","I-829","I-407",
    "Work Permit","Re-entry Permit","Consultation","Analysis","Referral",
    "Derivatives","Travel Permit","USC","LPR","Perm"
}

# ------------------------------------------------------------
# FICHIERS & UTILITAIRES
# ------------------------------------------------------------
def ensure_file(path: str, sheet_name: str, cols: list[str]) -> None:
    p = Path(path)
    if not p.exists():
        df = pd.DataFrame(columns=cols)
        with pd.ExcelWriter(p, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=sheet_name, index=False)

ensure_file(CLIENTS_FILE_DEFAULT, SHEET_CLIENTS, CLIENTS_COLS)
ensure_file(VISA_FILE_DEFAULT, SHEET_VISA, ["Categorie","Sous-categorie 1","COS","EOS"])

def _safe_str(x) -> str:
    try:
        if pd.isna(x): return ""
    except Exception:
        pass
    return str(x)

def _to_num(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.replace(r"[^\d,.\-]", "", regex=True).str.replace(",", ".", regex=False)
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
    seen, new_cols = {}, []
    for c in cols:
        if c not in seen:
            seen[c] = 1; new_cols.append(c)
        else:
            seen[c] += 1; new_cols.append(f"{c}_{seen[c]}")
    out = df.copy(); out.columns = new_cols
    return out

# ------------------------------------------------------------
# VISA — parsing (cellule = 1) + injection Étudiants F-1/F-2 COS/EOS
# ------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_raw_visa_df(xlsx_path: str | Path, sheet_name: str = SHEET_VISA) -> pd.DataFrame:
    try:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()
    df = _uniquify_columns(df)
    df.columns = df.columns.map(str).str.strip()
    return df

def _find_cat_sub_columns(df: pd.DataFrame) -> tuple[str|None,str|None]:
    def _norm(s: str) -> str:
        s2 = unicodedata.normalize("NFKD", s)
        s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
        s2 = s2.lower().strip().replace("\u00a0"," ")
        s2 = s2.replace("-", " ").replace("_"," ")
        return " ".join(s2.split())
    cmap = {_norm(c): c for c in df.columns}
    cat_col = next((cmap[k] for k in cmap if "categorie" in k), None)
    sub_col = None
    for k in cmap:
        if k.startswith("sous"):
            sub_col = cmap[k]; break
    return cat_col, sub_col

@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)): return False
        if isinstance(v, (int,float)): return float(v) == 1.0
        s = str(v).strip().lower()
        return s in {"1","true","vrai","oui","yes","x"}

    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        return {}

    sheets = [sheet_name] if sheet_name else xls.sheet_names
    for sn in sheets:
        try:
            dfv = pd.read_excel(xlsx_path, sheet_name=sn)
        except Exception:
            continue
        if dfv.empty: continue

        dfv = _uniquify_columns(dfv)
        dfv.columns = dfv.columns.map(str).str.strip()

        # cols
        cat_col, sub_col = _find_cat_sub_columns(dfv)
        if not cat_col: continue
        if not sub_col:
            dfv["_Sous_"] = ""; sub_col = "_Sous_"
        check_cols = [c for c in dfv.columns if c not in {cat_col, sub_col}]

        out: dict[str, dict[str, list[str]]] = {}
        for _, row in dfv.iterrows():
            cat = _safe_str(row.get(cat_col,"")).strip()
            sub = _safe_str(row.get(sub_col,"")).strip()
            if not cat: continue
            opts = []
            for cc in check_cols:
                if _is_checked(row.get(cc)):
                    opts.append(f"{sub} {cc}".strip())
            if not opts and sub:
                opts = [sub]
            if opts:
                out.setdefault(cat,{})
                out[cat].setdefault(sub,[])
                out[cat][sub].extend(opts)

        # injection Étudiants
        def _inject_students(d: dict[str, dict[str, list[str]]]) -> None:
            keys = [k for k in d.keys() if k.strip().lower() in {"etudiants","etudiant","students","student"}]
            if not keys: return
            for k in keys:
                subs = d.setdefault(k,{})
                for sub in ("F-1","F-2"):
                    arr = subs.setdefault(sub,[])
                    for w in (f"{sub} COS", f"{sub} EOS"):
                        if w not in arr: arr.append(w)
                    subs[sub] = sorted(set(arr))
        if out:
            for k,v in out.items():
                for s, arr in v.items():
                    v[s] = sorted(set(arr))
            _inject_students(out)
            return out
    return {}

def _normalize_options_json(x) -> dict:
    try:
        d = json.loads(_safe_str(x) or "{}")
        if not isinstance(d, dict): return {}
        excl = d.get("exclusive", None)
        opts = d.get("options", [])
        if not isinstance(opts, list): opts=[]
        return {"exclusive": excl, "options": [str(o) for o in opts]}
    except Exception:
        return {"exclusive": None, "options": []}

def render_dynamic_steps(cat: str, sub: str, keyprefix: str, visa_file: str, preselected: dict|None=None) -> tuple[str,str,dict]:
    if not (cat and sub):
        return "", "Choisir Catégorie et Sous-catégorie.", {"exclusive": None, "options": []}

    vdf = load_raw_visa_df(visa_file, SHEET_VISA)
    if vdf.empty:
        return "", "Feuille Visa vide ou introuvable.", {"exclusive": None, "options": []}

    cat_col, sub_col = _find_cat_sub_columns(vdf)
    if not cat_col:
        return "", "Colonne 'Catégorie' absente.", {"exclusive": None, "options": []}

    if sub_col:
        row = vdf[(vdf[cat_col].astype(str).str.strip()==cat) & (vdf[sub_col].astype(str).str.strip()==sub)]
    else:
        row = vdf[vdf[cat_col].astype(str).str.strip()==cat]
    if row.empty:
        return "", "Combinaison non trouvée.", {"exclusive": None, "options": []}
    row = row.iloc[0]
    option_cols = [c for c in vdf.columns if c not in {cat_col, sub_col}]

    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)): return False
        if isinstance(v, (int,float)): return float(v) == 1.0
        s = str(v).strip().lower()
        return s in {"1","true","vrai","oui","yes","x"}

    possibles = [c for c in option_cols if _is_checked(row.get(c))]
    exclusive = None
    if {"COS","EOS"}.issubset(set(possibles)):
        exclusive = ("COS","EOS")
    elif {"USCIS","CP"}.issubset(set(possibles)):
        exclusive = ("USCIS","CP")

    pre = _normalize_options_json(preselected or {})
    visa_final, info_msg = "", ""
    selected_opts: list[str] = []
    selected_excl: str|None = None

    if exclusive:
        st.caption("Choix exclusif")
        default_index = 0
        if pre["exclusive"] in exclusive:
            default_index = list(exclusive).index(pre["exclusive"])
        choice = st.radio("Sélectionner", list(exclusive), index=default_index, horizontal=True, key=f"{keyprefix}_exclusive")
        selected_excl = choice
        visa_final = f"{sub} {choice}".strip()

    others = [c for c in possibles if not (exclusive and c in exclusive)]
    if others: st.caption("Options complémentaires")
    for i, col in enumerate(others):
        default_val = col in pre["options"]
        if col in TOGGLE_COLUMNS:
            val = st.toggle(col, value=default_val, key=f"{keyprefix}_tg_{i}")
        else:
            val = st.checkbox(col, value=default_val, key=f"{keyprefix}_cb_{i}")
        if val: selected_opts.append(col)

    if not visa_final:
        if len(selected_opts)==0:
            info_msg = "Coche une option (ou utilise le choix exclusif)."
        elif len(selected_opts)>1:
            info_msg = "Une seule option possible."
        else:
            visa_final = f"{sub} {selected_opts[0]}".strip()

    if not visa_final and not possibles:
        visa_final = sub

    return visa_final, info_msg, {"exclusive": selected_excl, "options": selected_opts}

# ------------------------------------------------------------
# CLIENTS — I/O & normalisation
# ------------------------------------------------------------
def _normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in CLIENTS_COLS:
        if c not in df.columns:
            df[c] = None

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df["Mois"] = df.apply(
        lambda r: f"{pd.to_datetime(r['Date']).month:02d}" if pd.notna(r["Date"]) else (_safe_str(r.get("Mois",""))[:2] or None),
        axis=1
    )

    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = _safe_num_series(df, c)

    # Paiements JSON -> liste
    def _parse_p(x):
        try:
            j = json.loads(_safe_str(x) or "[]");  return j if isinstance(j, list) else []
        except Exception:
            return []
    df["Paiements"] = df["Paiements"].apply(_parse_p)

    def _sum_json(lst):
        try:    return float(sum(float(it.get("amount",0.0) or 0.0) for it in (lst or [])))
        except: return 0.0
    paid_json = df["Paiements"].apply(_sum_json)
    df["Payé"] = pd.concat([df["Payé"].fillna(0.0).astype(float), paid_json], axis=1).max(axis=1)

    df[TOTAL] = df[HONO] + df[AUTRE]
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # Options JSON
    df["Options"] = df["Options"].apply(_normalize_options_json)

    # Statuts -> bool
    for s in ["Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"]:
        df[s] = df[s].apply(lambda v: bool(v) if isinstance(v, (bool,int,float)) else str(v).strip().lower() in {"1","true","vrai","oui","yes","x"})

    df["_Année_"]   = df["Date"].apply(lambda d: d.year if pd.notna(d) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda d: d.month if pd.notna(d) else pd.NA)
    return _uniquify_columns(df)

def _read_clients(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    return _normalize_clients(df)

def _write_clients(df: pd.DataFrame, path: str) -> None:
    df = df.copy()
    df["Options"] = df["Options"].apply(lambda d: json.dumps(_normalize_options_json(d), ensure_ascii=False))
    df["Paiements"] = df["Paiements"].apply(lambda l: json.dumps(l, ensure_ascii=False))
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        _uniquify_columns(df).to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if "Dossier N" in df.columns:
        s = pd.to_numeric(df["Dossier N"], errors="coerce")
        if s.notna().any():
            return int(s.max()) + 1
    return int(start)

def _normalize_name_for_id(nom: str) -> str:
    s = unicodedata.normalize("NFKD", nom)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace(" ", "_")
    return "".join(ch for ch in s if ch.isalnum() or ch in {"_", "-"})

def _next_client_id(base_df: pd.DataFrame, nom: str) -> str:
    base = _normalize_name_for_id(_safe_str(nom))
    # formattage suffixe 2 chiffres : -01, -02, ...
    mask = base_df["ID_Client"].astype(str).str.startswith(base + "-")
    existing = base_df.loc[mask, "ID_Client"].astype(str).tolist()
    if not existing:
        return f"{base}-01"
    # extraire suffixes
    maxn = 1
    for cid in existing:
        parts = cid.rsplit("-", 1)
        if len(parts) == 2 and parts[0] == base:
            try:
                n = int(parts[1])
                if n > maxn: maxn = n
            except:
                pass
    return f"{base}-{maxn+1:02d}"

# ------------------------------------------------------------
# BARRE LATÉRALE — fichiers
# ------------------------------------------------------------
with st.sidebar:
    st.markdown("### 📁 Fichiers")
    clients_path = st.text_input("Fichier Clients", value=CLIENTS_FILE_DEFAULT, key="sb_clients_path")
    visa_path    = st.text_input("Fichier Visa",    value=VISA_FILE_DEFAULT,    key="sb_visa_path")
    st.caption("Astuce : ces chemins sont mémorisés dans la session.")

# Charger mapping Visa
visa_map = parse_visa_sheet(visa_path)

# ------------------------------------------------------------
# ONGLETs
# ------------------------------------------------------------
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "👤 Clients", "🧾 Gestion", "📄 Visa (aperçu)"])

# ---------------------- DASHBOARD ----------------------------
with tabs[0]:
    st.subheader("📊 Dashboard")
    df = _read_clients(clients_path)

    # Filtres (dans l’onglet)
    f1,f2,f3,f4 = st.columns([1,1,2,2])
    years  = sorted([int(y) for y in pd.to_numeric(df["_Année_"], errors="coerce").dropna().unique().tolist()]) if not df.empty else []
    months = [f"{m:02d}" for m in range(1,13)]
    cats   = sorted(df["Categorie"].dropna().astype(str).unique().tolist()) if not df.empty else []
    visas  = sorted(df["Visa"].dropna().astype(str).unique().tolist()) if not df.empty else []

    fy = f1.multiselect("Année", years, default=[])
    fm = f2.multiselect("Mois (MM)", months, default=[])
    fc = f3.multiselect("Catégories", cats, default=[])
    fv = f4.multiselect("Visa", visas, default=[])

    f = df.copy()
    if fy: f = f[f["_Année_"].isin(fy)]
    if fm: f = f[f["Mois"].isin(fm)]
    if fc: f = f[f["Categorie"].astype(str).isin(fc)]
    if fv: f = f[f["Visa"].astype(str).isin(fv)]

    # KPI compacts
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Honoraires", _fmt_money(_safe_num_series(f,HONO).sum()))
    k3.metric("Autres", _fmt_money(_safe_num_series(f,AUTRE).sum()))
    k4.metric("Payé", _fmt_money(_safe_num_series(f,"Payé").sum()))
    k5.metric("Solde", _fmt_money(_safe_num_series(f,"Reste").sum()))

    view = f.copy()
    for c in [HONO,AUTRE,TOTAL,"Payé","Reste"]:
        if c in view.columns: view[c] = _safe_num_series(view,c).map(_fmt_money)
    if "Date" in view.columns: view["Date"] = view["Date"].astype(str)
    show_cols = [c for c in [
        "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
        HONO,AUTRE,TOTAL,"Payé","Reste",
        "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"
    ] if c in view.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
    view = view.sort_values(by=sort_cols) if sort_cols else view
    st.dataframe(_uniquify_columns(view[show_cols].reset_index(drop=True)), use_container_width=True)

# ---------------------- ANALYSES ----------------------------
with tabs[1]:
    st.subheader("📈 Analyses")
    df = _read_clients(clients_path)
    if df.empty:
        st.info("Pas de données.")
    else:
        a1,a2,a3,a4 = st.columns(4)
        years  = sorted([int(y) for y in pd.to_numeric(df["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1,13)]
        cats   = sorted(df["Categorie"].dropna().astype(str).unique().tolist())
        visas  = sorted(df["Visa"].dropna().astype(str).unique().tolist())

        fy = a1.multiselect("Année", years, default=[])
        fm = a2.multiselect("Mois (MM)", months, default=[])
        fc = a3.multiselect("Catégories", cats, default=[])
        fv = a4.multiselect("Visa", visas, default=[])

        f = df.copy()
        if fy: f = f[f["_Année_"].isin(fy)]
        if fm: f = f[f["Mois"].isin(fm)]
        if fc: f = f[f["Categorie"].astype(str).isin(fc)]
        if fv: f = f[f["Visa"].astype(str).isin(fv)]

        # KPI
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(f)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(f,HONO).sum()))
        k3.metric("Encaissements", _fmt_money(_safe_num_series(f,"Payé").sum()))
        k4.metric("Solde", _fmt_money(_safe_num_series(f,"Reste").sum()))

        # Graphes
        st.markdown("#### Volumes par mois (barres)")
        if not f.empty:
            vol_mois = f.groupby(["_Année_","Mois"], dropna=True).size().reset_index(name="Dossiers")
            chart = alt.Chart(vol_mois).mark_bar().encode(
                x=alt.X("Mois:N", sort=months),
                y="Dossiers:Q",
                color="__year__:N"
            ).transform_calculate(__year__="toString(datum._Année_)").properties(height=250)
            st.altair_chart(chart, use_container_width=True)

        st.markdown("#### Volumes par catégorie")
        vol_cat = f.groupby(["Categorie"], dropna=True).size().reset_index(name="Dossiers")
        st.dataframe(vol_cat.sort_values("Dossiers", ascending=False), use_container_width=True)

        st.markdown("#### CA par catégorie (honoraires)")
        ca_cat = f.groupby(["Categorie"], dropna=True)[HONO].sum().reset_index().rename(columns={HONO:"Honoraires"})
        ca_cat["Honoraires"] = ca_cat["Honoraires"].astype(float).map(_fmt_money)
        st.dataframe(ca_cat.sort_values("Honoraires", ascending=False), use_container_width=True)

# ---------------------- ESCROW ------------------------------
with tabs[2]:
    st.subheader("🏦 Escrow — dossiers envoyés")
    df = _read_clients(clients_path)
    if df.empty:
        st.info("Pas de données.")
    else:
        # Seuls les dossiers envoyés sont éligibles au transfert escrow (simplifié)
        envoyes = df[df["Dossier envoyé"] == True].copy()
        k1,k2,k3 = st.columns(3)
        k1.metric("Dossiers envoyés", f"{len(envoyes)}")
        k2.metric("Payé (envoyés)", _fmt_money(_safe_num_series(envoyes,"Payé").sum()))
        k3.metric("Solde total (envoyés)", _fmt_money(_safe_num_series(envoyes,"Reste").sum()))

        if envoyes.empty:
            st.success("Aucun dossier ‘envoyé’.")
        else:
            show = envoyes[[
                "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa",
                HONO,"Payé","Reste","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"
            ]].copy()
            for c in [HONO,"Payé","Reste"]:
                show[c] = _safe_num_series(show,c).map(_fmt_money)
            st.dataframe(show.reset_index(drop=True), use_container_width=True)
            st.caption("Note : logique de transfert détaillée possible (journal des transferts, etc.).")

# ---------------------- CLIENTS (fiche comptable) ----------
with tabs[3]:
    st.subheader("👤 Fiche client")
    df = _read_clients(clients_path)
    if df.empty:
        st.info("Aucun client.")
    else:
        # Sélection client
        by = st.radio("Sélection par :", ["Nom","ID_Client"], horizontal=True)
        if by == "Nom":
            names = df["Nom"].astype(str).tolist()
            sel = st.selectbox("Nom", names)
            row = df[df["Nom"].astype(str)==sel].iloc[0]
            idx = df.index[df["Nom"].astype(str)==sel][0]
        else:
            ids = df["ID_Client"].astype(str).tolist()
            sel = st.selectbox("ID_Client", ids)
            row = df[df["ID_Client"].astype(str)==sel].iloc[0]
            idx = df.index[df["ID_Client"].astype(str)==sel][0]

        # KPI
        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Dossier N", str(row["Dossier N"]))
        k2.metric("Total", _fmt_money(float(row[TOTAL])))
        k3.metric("Payé", _fmt_money(float(row["Payé"])))
        k4.metric("Reste", _fmt_money(float(row["Reste"])))

        # Statuts
        st.markdown("#### Statut du dossier")
        s1,s2,s3,s4,s5 = st.columns(5)
        sent    = s1.checkbox("Dossier envoyé", value=bool(row["Dossier envoyé"]), key=f"cli_sent_{idx}")
        appr    = s2.checkbox("Dossier approuvé", value=bool(row["Dossier approuvé"]), key=f"cli_appr_{idx}")
        rfe     = s3.checkbox("RFE", value=bool(row["RFE"]), key=f"cli_rfe_{idx}")
        refus   = s4.checkbox("Dossier refusé", value=bool(row["Dossier refusé"]), key=f"cli_refus_{idx}")
        annul   = s5.checkbox("Dossier annulé", value=bool(row["Dossier annulé"]), key=f"cli_annul_{idx}")
        if st.button("💾 Mettre à jour les statuts", key=f"cli_stat_save_{idx}"):
            base = _read_clients(clients_path)
            base.loc[idx,["Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"]] = [sent,appr,rfe,refus,annul]
            _write_clients(_normalize_clients(base), clients_path)
            st.success("Statuts mis à jour."); st.rerun()

        # Paiements (timeline)
        st.markdown("#### Paiements")
        hist = row["Paiements"] if isinstance(row["Paiements"], list) else []
        if hist:
            h = pd.DataFrame(hist)
            if "amount" in h.columns:
                h["amount"] = h["amount"].astype(float)
            if "date" in h.columns:
                h["date"] = pd.to_datetime(h["date"], errors="coerce")
            st.dataframe(h, use_container_width=True)
            if "date" in h.columns and "amount" in h.columns:
                chart = alt.Chart(h.dropna(subset=["date"])).mark_bar().encode(
                    x="date:T", y="amount:Q", tooltip=["date:T","amount:Q","mode:N"]
                ).properties(height=200)
                st.altair_chart(chart, use_container_width=True)
        else:
            st.caption("Aucun paiement.")

# ---------------------- GESTION (Ajouter/Modifier/Supprimer)
with tabs[4]:
    st.subheader("🧾 Gestion")
    df = _read_clients(clients_path)

    mode = st.radio("Action", ["Ajouter","Modifier","Supprimer"], horizontal=True)

    if mode == "Ajouter":
        st.markdown("#### ➕ Ajouter un client")
        c1,c2,c3 = st.columns(3)
        with c1:
            nom = st.text_input("Nom", key="add_nom")
            dcr = st.date_input("Date de création", value=date.today(), key="add_date")
        with c2:
            cats = sorted(list(visa_map.keys()))
            cat = st.selectbox("Catégorie", options=[""]+cats, index=0, key="add_cat")
            subs = sorted(list(visa_map.get(cat, {}).keys())) if cat else []
            sub = st.selectbox("Sous-catégorie", options=[""]+subs, index=0, key="add_sub")
        with c3:
            visa_final, info_msg, opts = render_dynamic_steps(cat, sub, "add_steps", visa_file=visa_path, preselected=None)
            if info_msg: st.info(info_msg)
            hono = st.number_input(HONO, min_value=0.0, step=10.0, format="%.2f", key="add_hono")
            autre= st.number_input(AUTRE, min_value=0.0, step=10.0, format="%.2f", key="add_autre")

        if st.button("💾 Créer", key="btn_add_create"):
            if not nom or not cat or not sub:
                st.warning("Nom, Catégorie et Sous-catégorie sont requis."); st.stop()
            base = _read_clients(clients_path)
            dossier = _next_dossier(base)
            cid = _next_client_id(base, nom)  # <-- format Nom-01, Nom-02, ...
            visa_val = visa_final if visa_final else sub
            total = float(hono)+float(autre)
            row = {
                "Dossier N": dossier,"ID_Client": cid,"Nom": nom,
                "Date": pd.to_datetime(dcr).date(),"Mois": f"{dcr.month:02d}",
                "Categorie": cat,"Sous-categorie": sub,"Visa": visa_val,
                HONO: float(hono), AUTRE: float(autre),
                TOTAL: total,"Payé": 0.0,"Reste": total,
                "Paiements": json.dumps([], ensure_ascii=False),
                "Options": json.dumps(_normalize_options_json(opts), ensure_ascii=False),
                "Dossier envoyé": False, "Dossier approuvé": False, "RFE": False,
                "Dossier refusé": False, "Dossier annulé": False
            }
            base = pd.concat([base, pd.DataFrame([row])], ignore_index=True)
            _write_clients(_normalize_clients(base), clients_path)
            st.success(f"✅ Client créé (Dossier {dossier}, ID {cid})."); st.rerun()

    elif mode == "Modifier":
        st.markdown("#### 🛠️ Modifier un client")
        if df.empty:
            st.info("Aucun client.")
        else:
            idx = st.selectbox("Sélectionne la ligne", options=list(df.index),
                               format_func=lambda i: f"{df.loc[i,'Nom']} — {df.loc[i,'ID_Client']}")
            row = df.loc[idx]
            c1,c2,c3 = st.columns(3)
            with c1:
                nom = st.text_input("Nom", value=_safe_str(row["Nom"]), key=f"mod_nom_{idx}")
                dcr = st.date_input("Date de création", value=(pd.to_datetime(row["Date"]).date() if pd.notna(row["Date"]) else date.today()), key=f"mod_date_{idx}")
            with c2:
                cats = sorted(list(visa_map.keys()))
                cat  = st.selectbox("Catégorie", options=[""]+cats,
                                    index=(cats.index(_safe_str(row["Categorie"]))+1 if _safe_str(row["Categorie"]) in cats else 0),
                                    key=f"mod_cat_{idx}")
                subs = sorted(list(visa_map.get(cat, {}).keys())) if cat else []
                sub  = st.selectbox("Sous-catégorie", options=[""]+subs,
                                    index=(subs.index(_safe_str(row["Sous-categorie"]))+1 if _safe_str(row["Sous-categorie"]) in subs else 0),
                                    key=f"mod_sub_{idx}")
            with c3:
                cur_opts = _normalize_options_json(row.get("Options", {}))
                visa_final, info_msg, opts = render_dynamic_steps(cat, sub, f"mod_steps_{idx}", visa_file=visa_path, preselected=cur_opts)
                if info_msg: st.info(info_msg)
                hono = st.number_input(HONO,  min_value=0.0, value=float(row[HONO]),  step=10.0, format="%.2f", key=f"mod_hono_{idx}")
                autre= st.number_input(AUTRE, min_value=0.0, value=float(row[AUTRE]), step=10.0, format="%.2f", key=f"mod_autre_{idx}")

            # paiements
            st.markdown("##### 💳 Paiements")
            p1,p2,p3,p4 = st.columns([1,1,1,2])
            with p1: pdt = st.date_input("Date paiement", value=date.today(), key=f"mod_paydt_{idx}")
            with p2: pmd = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"mod_mode_{idx}")
            with p3: pmt = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"mod_amt_{idx}")
            with p4:
                if st.button("➕ Ajouter paiement", key=f"mod_addpay_{idx}"):
                    base = _read_clients(clients_path)
                    plist = base.loc[idx,"Paiements"]; 
                    if not isinstance(plist, list): plist=[]
                    plist.append({"date": str(pdt), "mode": pmd, "amount": float(pmt)})
                    base.loc[idx,"Paiements"] = plist
                    base = _normalize_clients(base)
                    _write_clients(base, clients_path)
                    st.success("Paiement ajouté."); st.rerun()

            # statuts
            st.markdown("##### 🔖 Statuts")
            s1,s2,s3,s4,s5 = st.columns(5)
            sent  = s1.checkbox("Dossier envoyé", value=bool(row["Dossier envoyé"]), key=f"mod_sent_{idx}")
            appr  = s2.checkbox("Dossier approuvé", value=bool(row["Dossier approuvé"]), key=f"mod_appr_{idx}")
            rfe   = s3.checkbox("RFE", value=bool(row["RFE"]), key=f"mod_rfe_{idx}")
            refus = s4.checkbox("Dossier refusé", value=bool(row["Dossier refusé"]), key=f"mod_refus_{idx}")
            annul = s5.checkbox("Dossier annulé", value=bool(row["Dossier annulé"]), key=f"mod_annul_{idx}")

            if st.button("💾 Sauvegarder", key=f"mod_save_{idx}"):
                base = _read_clients(clients_path)
                base.loc[idx,"Nom"] = nom
                base.loc[idx,"Date"] = pd.to_datetime(dcr).date()
                base.loc[idx,"Mois"] = f"{dcr.month:02d}"
                base.loc[idx,"Categorie"] = cat
                base.loc[idx,"Sous-categorie"] = sub
                base.loc[idx,"Visa"] = visa_final if visa_final else (sub or "")
                base.loc[idx,HONO] = float(hono)
                base.loc[idx,AUTRE] = float(autre)
                base.loc[idx,TOTAL] = float(hono)+float(autre)
                base.loc[idx,"Options"] = opts
                base.loc[idx,["Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"]] = [sent,appr,rfe,refus,annul]
                _write_clients(_normalize_clients(base), clients_path)
                st.success("✅ Modifications enregistrées."); st.rerun()

    else:  # Supprimer
        st.markdown("#### 🗑️ Supprimer")
        if df.empty:
            st.info("Aucun client.")
        else:
            idx = st.selectbox("Sélectionne la ligne à supprimer", options=list(df.index),
                               format_func=lambda i: f"{df.loc[i,'Nom']} — {df.loc[i,'ID_Client']}")
            if st.button("Confirmer la suppression", type="primary", key="btn_confirm_del"):
                base = _read_clients(clients_path)
                base = base.drop(index=idx).reset_index(drop=True)
                _write_clients(_normalize_clients(base), clients_path)
                st.success("Client supprimé."); st.rerun()

# ---------------------- VISA (aperçu) -----------------------
with tabs[5]:
    st.subheader("📄 Référentiel Visa (lecture Excel : cellule = 1 → option)")
    if not visa_map:
        st.warning("Aucune donnée Visa trouvée. Vérifie le fichier.")
    else:
        c1,c2 = st.columns(2)
        with c1:
            cat = st.selectbox("Catégorie", [""]+sorted(list(visa_map.keys())), key="vz_cat")
        with c2:
            subs = sorted(list(visa_map.get(cat, {}).keys())) if cat else []
            sub  = st.selectbox("Sous-catégorie", [""]+subs, key="vz_sub")
        if cat and sub:
            st.write("**Options** :", ", ".join(sorted(visa_map.get(cat, {}).get(sub, []))) or "(Aucune → visa = sous-catégorie)")
        with st.expander("Aperçu complet"):
            for k, submap in visa_map.items():
                st.write(f"**{k}**")
                for s, arr in submap.items():
                    st.caption(f"- {s} → {', '.join(arr)}")