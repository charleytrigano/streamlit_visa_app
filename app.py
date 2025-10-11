from __future__ import annotations

import streamlit as st
import pandas as pd
import json
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import unicodedata
import zipfile
import openpyxl

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

CLIENTS_COLS = [
    "Dossier N","ID_Client","Nom","Date","Mois",
    "Categorie","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
    "Payé","Reste","Paiements","Options"
]

TOGGLE_COLUMNS = {
    "AOS","CP","USCIS","I-130","I-140","I-140 & AOS","I-829","I-407",
    "Work Permit","Re-entry Permit","Consultation","Analysis","Referral",
    "Derivatives","Travel Permit","USC","LPR","Perm"
}

# ------------------------------------------------------------
# UTILITAIRES
# ------------------------------------------------------------
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

def ensure_file(path: str, sheet_name: str, cols: list[str]) -> None:
    p = Path(path)
    if not p.exists():
        df = pd.DataFrame(columns=cols)
        with pd.ExcelWriter(p, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=sheet_name, index=False)

# créer si absents
ensure_file(CLIENTS_FILE_DEFAULT, SHEET_CLIENTS, CLIENTS_COLS)
ensure_file(VISA_FILE_DEFAULT, SHEET_VISA, ["Categorie","Sous-categorie 1"])

# ------------------------------------------------------------
# VISA – parser (cellule=1 => option active) + injection F-1/F-2 COS/EOS
# ------------------------------------------------------------
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)): 
            return False
        if isinstance(v, (int, float)): 
            return float(v) == 1.0
        s = str(v).strip().lower()
        return s in {"1","true","vrai","oui","yes","x"}

    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        return {}

    sheets_to_try = [sheet_name] if sheet_name else xls.sheet_names
    for sn in sheets_to_try:
        try:
            dfv = pd.read_excel(xlsx_path, sheet_name=sn)
        except Exception:
            continue
        if dfv.empty:
            continue

        dfv = _uniquify_columns(dfv)
        dfv.columns = dfv.columns.map(str).str.strip()

        def _norm(s: str) -> str:
            s2 = unicodedata.normalize("NFKD", s)
            s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
            s2 = s2.strip().lower().replace("\u00a0", " ")
            s2 = s2.replace("-", " ").replace("_", " ")
            return " ".join(s2.split())

        cmap = {_norm(c): c for c in dfv.columns}
        cat_col = next((cmap[k] for k in cmap if "categorie" in k), None)
        sub_col = next((cmap[k] for k in cmap if "sous" in k), None)
        if not cat_col:
            continue
        if not sub_col:
            dfv["_Sous_"] = ""; sub_col = "_Sous_"

        check_cols = [c for c in dfv.columns if c not in {cat_col, sub_col}]
        out: dict[str, dict[str, list[str]]] = {}

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

        # injecter F-1 / F-2 COS/EOS si catégorie étudiants présente
        def _inject_students(d: dict[str, dict[str, list[str]]]) -> None:
            keys = [k for k in d.keys() if k.strip().lower() in {"etudiants","etudiant","students","student"}]
            if not keys: return
            for k in keys:
                subs = d.setdefault(k, {})
                for sub in ("F-1","F-2"):
                    arr = subs.setdefault(sub, [])
                    for w in (f"{sub} COS", f"{sub} EOS"):
                        if w not in arr: arr.append(w)
                    subs[sub] = sorted(set(arr))

        if out:
            for cat, subs in out.items():
                for sub, arr in subs.items():
                    subs[sub] = sorted(set(arr))
            _inject_students(out)
            return out
    return {}

@st.cache_data(show_spinner=False)
def load_raw_visa_df(xlsx_path: str | Path, sheet_name: str = SHEET_VISA) -> pd.DataFrame:
    try:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()
    return _uniquify_columns(df)

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

def _normalize_options_json(x) -> dict:
    try:
        d = json.loads(_safe_str(x) or "{}")
        if not isinstance(d, dict): return {}
        excl = d.get("exclusive", None)
        opts = d.get("options", []); 
        if not isinstance(opts, list): opts = []
        return {"exclusive": excl, "options": [str(o) for o in opts]}
    except Exception:
        return {"exclusive": None, "options": []}

def render_dynamic_steps(cat: str, sub: str, keyprefix: str, visa_file: str, preselected: dict | None = None) -> tuple[str,str,dict]:
    if not (cat and sub):
        return "", "Choisir d'abord Catégorie et Sous-catégorie.", {"exclusive": None, "options": []}

    vdf = load_raw_visa_df(visa_file, SHEET_VISA)
    if vdf.empty:
        return "", "Feuille Visa introuvable.", {"exclusive": None, "options": []}

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
    selected_opts, selected_excl = [], None
    visa_final, info_msg = "", ""

    # exclusif
    if exclusive:
        st.caption("Choix exclusif")
        def_index = 0
        if pre["exclusive"] in exclusive:
            def_index = list(exclusive).index(pre["exclusive"])
        choice = st.radio("Sélectionner", options=list(exclusive), index=def_index,
                          horizontal=True, key=f"{keyprefix}_x")
        selected_excl = choice
        visa_final = f"{sub} {choice}".strip()

    # autres
    others = [c for c in possibles if not (exclusive and c in exclusive)]
    if others:
        st.caption("Options complémentaires")
    for i, col in enumerate(others):
        default_val = col in pre["options"]
        if col in TOGGLE_COLUMNS:
            val = st.toggle(col, value=default_val, key=f"{keyprefix}_t{i}")
        else:
            val = st.checkbox(col, value=default_val, key=f"{keyprefix}_c{i}")
        if val:
            selected_opts.append(col)

    if not visa_final:
        if len(selected_opts)==0:
            info_msg = "Coche une option (une seule) ou utilise l’exclusif."
        elif len(selected_opts)>1:
            info_msg = "Une seule option possible dans ce cas."
        else:
            visa_final = f"{sub} {selected_opts[0]}".strip()

    if not visa_final and not possibles:
        visa_final = sub

    return visa_final, info_msg, {"exclusive": selected_excl, "options": selected_opts}

# ------------------------------------------------------------
# CLIENTS I/O + normalisation
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
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        df[c] = _safe_num_series(df, c)

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

    df["Total (US $)"] = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]
    df["Reste"] = (df["Total (US $)"] - df["Payé"]).clip(lower=0.0)

    df["Options"] = df["Options"].apply(_normalize_options_json)

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

def _make_client_id(nom: str, d: date) -> str:
    return f"{_safe_str(nom).strip().replace(' ','_')}-{d:%Y%m%d}"

# ------------------------------------------------------------
# BARRE LATÉRALE — étapes
# ------------------------------------------------------------
with st.sidebar:
    st.markdown("## 🧭 Étapes")
    st.caption("Choisis l’action et travaille toujours sur le même fichier.")
    # fichiers
    clients_path = st.text_input("Fichier Clients", value=CLIENTS_FILE_DEFAULT, key="sb_clients_path")
    visa_path    = st.text_input("Fichier Visa",    value=VISA_FILE_DEFAULT, key="sb_visa_path")

    # action
    action = st.radio("Action clients", options=["Ajouter","Modifier","Supprimer"], horizontal=False, key="sb_action")

    # filtres rapides (Dashboard)
    st.markdown("---")
    st.markdown("### 🔎 Filtres Dashboard")
    sb_year  = st.multiselect("Année", [], key="sb_years")  # remplis plus bas
    sb_month = st.multiselect("Mois (MM)", [f"{m:02d}" for m in range(1,13)], key="sb_months")
    sb_cat   = st.multiselect("Catégories", [], key="sb_cats")
    sb_visa  = st.multiselect("Visa", [], key="sb_visas")

# charger mapping visa
visa_map = parse_visa_sheet(visa_path)

# ------------------------------------------------------------
# ONGLETs
# ------------------------------------------------------------
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "📄 Visa (aperçu)"])

# ---------------------- DASHBOARD ----------------------------
with tabs[0]:
    st.subheader("📊 Dashboard — tous les clients")
    df = _read_clients(clients_path)

    # mettre valeurs dans la sidebar
    with st.sidebar:
        if not df.empty:
            years = sorted([int(y) for y in pd.to_numeric(df["_Année_"], errors="coerce").dropna().unique().tolist()])
            st.session_state.sb_years = st.session_state.sb_years or years
            st.session_state.sb_cats  = st.session_state.sb_cats or sorted(df["Categorie"].dropna().astype(str).unique().tolist())
            st.session_state.sb_visas = st.session_state.sb_visas or sorted(df["Visa"].dropna().astype(str).unique().tolist())

    # appliquer filtres latéraux
    f = df.copy()
    if st.session_state.sb_years:
        f = f[f["_Année_"].isin(st.session_state.sb_years)]
    if st.session_state.sb_months:
        f = f[f["Mois"].isin(st.session_state.sb_months)]
    if st.session_state.sb_cats:
        f = f[f["Categorie"].astype(str).isin(st.session_state.sb_cats)]
    if st.session_state.sb_visas:
        f = f[f["Visa"].astype(str).isin(st.session_state.sb_visas)]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Honoraires", _fmt_money(_safe_num_series(f,"Montant honoraires (US $)").sum()))
    k3.metric("Payé", _fmt_money(_safe_num_series(f,"Payé").sum()))
    k4.metric("Solde", _fmt_money(_safe_num_series(f,"Reste").sum()))

    # tableau
    view = f.copy()
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        if c in view.columns:
            view[c] = _safe_num_series(view,c).map(_fmt_money)
    if "Date" in view.columns: view["Date"] = view["Date"].astype(str)
    view["Options (résumé)"] = view["Options"].apply(lambda d: f"[{(d or {}).get('exclusive')}] + {', '.join((d or {}).get('options', []))}" if isinstance(d, dict) else "")
    show_cols = [c for c in [
        "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
        "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste","Options (résumé)"
    ] if c in view.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
    view = view.sort_values(by=sort_cols) if sort_cols else view
    st.dataframe(_uniquify_columns(view[show_cols].reset_index(drop=True)), use_container_width=True)

    st.markdown("---")
    st.markdown("### ✏️ Gestion (suivant l’action choisie dans la barre latérale)")
    # panneau central selon action
    if action == "Ajouter":
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
            hono = st.number_input("Montant honoraires (US $)", min_value=0.0, step=10.0, format="%.2f", key="add_hono")
            autre= st.number_input("Autres frais (US $)",     min_value=0.0, step=10.0, format="%.2f", key="add_autre")

        if st.button("💾 Créer", key="btn_add_create"):
            if not nom or not cat or not sub:
                st.warning("Nom, Catégorie et Sous-catégorie sont requis."); st.stop()
            base = _read_clients(clients_path)
            dossier = _next_dossier(base)
            cid_base = _make_client_id(nom, dcr); cid = cid_base; i=0
            while (base["ID_Client"].astype(str)==cid).any():
                i+=1; cid=f"{cid_base}-{i}"
            visa_val = visa_final if visa_final else sub
            total = float(hono)+float(autre)
            row = {
                "Dossier N": dossier,"ID_Client": cid,"Nom": nom,
                "Date": pd.to_datetime(dcr).date(),"Mois": f"{dcr.month:02d}",
                "Categorie": cat,"Sous-categorie": sub,"Visa": visa_val,
                "Montant honoraires (US $)": float(hono),"Autres frais (US $)": float(autre),
                "Total (US $)": total,"Payé": 0.0,"Reste": total,
                "Paiements": json.dumps([], ensure_ascii=False),
                "Options": json.dumps(_normalize_options_json(opts), ensure_ascii=False),
            }
            base = pd.concat([base, pd.DataFrame([row])], ignore_index=True)
            _write_clients(_normalize_clients(base), clients_path)
            st.success("✅ Client créé."); st.rerun()

    elif action == "Modifier":
        st.markdown("#### 🛠️ Modifier un client")
        if df.empty:
            st.info("Aucun client."); 
        else:
            idx = st.selectbox("Sélectionne la ligne à modifier", options=list(df.index),
                               format_func=lambda i: f"{df.loc[i,'Nom']} — {df.loc[i,'ID_Client']}", key="mod_idx")
            row = df.loc[idx]
            c1,c2,c3 = st.columns(3)
            with c1:
                nom = st.text_input("Nom", value=_safe_str(row["Nom"]), key=f"mod_nom_{idx}")
                dcr = st.date_input("Date de création", value=(pd.to_datetime(row["Date"]).date() if pd.notna(row["Date"]) else date.today()), key=f"mod_date_{idx}")
            with c2:
                cats = sorted(list(visa_map.keys()))
                cat  = st.selectbox("Catégorie", options=[""]+cats,
                                    index=(cats.index(row["Categorie"])+1 if _safe_str(row["Categorie"]) in cats else 0),
                                    key=f"mod_cat_{idx}")
                subs = sorted(list(visa_map.get(cat, {}).keys())) if cat else []
                sub  = st.selectbox("Sous-catégorie", options=[""]+subs,
                                    index=(subs.index(row["Sous-categorie"])+1 if _safe_str(row["Sous-categorie"]) in subs else 0),
                                    key=f"mod_sub_{idx}")
            with c3:
                cur_opts = _normalize_options_json(row.get("Options", {}))
                visa_final, info_msg, opts = render_dynamic_steps(cat, sub, f"mod_steps_{idx}", visa_file=visa_path, preselected=cur_opts)
                if info_msg: st.info(info_msg)
                hono = st.number_input("Montant honoraires (US $)", min_value=0.0, value=float(row["Montant honoraires (US $)"]), step=10.0, format="%.2f", key=f"mod_hono_{idx}")
                autre= st.number_input("Autres frais (US $)",     min_value=0.0, value=float(row["Autres frais (US $)"]),     step=10.0, format="%.2f", key=f"mod_autre_{idx}")

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

            # historique
            hist = row["Paiements"] if isinstance(row["Paiements"], list) else []
            if hist:
                h = pd.DataFrame(hist); 
                if "amount" in h.columns: h["amount"] = h["amount"].astype(float).map(_fmt_money)
                st.dataframe(h, use_container_width=True)
            else:
                st.caption("Aucun paiement.")

            if st.button("💾 Sauvegarder", key=f"mod_save_{idx}"):
                base = _read_clients(clients_path)
                base.loc[idx,"Nom"] = nom
                base.loc[idx,"Date"] = pd.to_datetime(dcr).date()
                base.loc[idx,"Mois"] = f"{dcr.month:02d}"
                base.loc[idx,"Categorie"] = cat
                base.loc[idx,"Sous-categorie"] = sub
                base.loc[idx,"Visa"] = visa_final if visa_final else (sub or "")
                base.loc[idx,"Montant honoraires (US $)"] = float(hono)
                base.loc[idx,"Autres frais (US $)"] = float(autre)
                base.loc[idx,"Total (US $)"] = float(hono)+float(autre)
                base.loc[idx,"Options"] = opts
                _write_clients(_normalize_clients(base), clients_path)
                st.success("✅ Modifications enregistrées."); st.rerun()

    elif action == "Supprimer":
        st.markdown("#### 🗑️ Supprimer un client")
        if df.empty:
            st.info("Aucun client.")
        else:
            idx = st.selectbox("Sélectionne la ligne à supprimer", options=list(df.index),
                               format_func=lambda i: f"{df.loc[i,'Nom']} — {df.loc[i,'ID_Client']}", key="del_idx")
            if st.button("Confirmer la suppression", type="primary", key="btn_confirm_del"):
                base = _read_clients(clients_path)
                base = base.drop(index=idx).reset_index(drop=True)
                _write_clients(_normalize_clients(base), clients_path)
                st.success("Client supprimé."); st.rerun()

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

        fy = a1.multiselect("Année", years, default=[], key="ana_y")
        fm = a2.multiselect("Mois (MM)", months, default=[], key="ana_m")
        fc = a3.multiselect("Catégories", cats, default=[], key="ana_c")
        fv = a4.multiselect("Visa", visas, default=[], key="ana_v")

        f = df.copy()
        if fy: f = f[f["_Année_"].isin(fy)]
        if fm: f = f[f["Mois"].isin(fm)]
        if fc: f = f[f["Categorie"].astype(str).isin(fc)]
        if fv: f = f[f["Visa"].astype(str).isin(fv)]

        k1,k2,k3,k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(f)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(f,"Montant honoraires (US $)").sum()))
        k3.metric("Encaissements", _fmt_money(_safe_num_series(f,"Payé").sum()))
        k4.metric("Solde à encaisser", _fmt_money(_safe_num_series(f,"Reste").sum()))

        st.markdown("#### Volumes par catégorie")
        vol_cat = f.groupby(["Categorie"], dropna=True).size().reset_index(name="Dossiers")
        st.dataframe(vol_cat.sort_values("Dossiers", ascending=False), use_container_width=True)

        st.markdown("#### Volumes par sous-catégorie")
        vol_sub = f.groupby(["Sous-categorie"], dropna=True).size().reset_index(name="Dossiers")
        st.dataframe(vol_sub.sort_values("Dossiers", ascending=False), use_container_width=True)

        st.markdown("#### Volumes par Visa")
        vol_visa = f.groupby(["Visa"], dropna=True).size().reset_index(name="Dossiers")
        st.dataframe(vol_visa.sort_values("Dossiers", ascending=False), use_container_width=True)

        st.markdown("#### Détails (clients filtrés)")
        detail = f.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in detail.columns: detail[c] = _safe_num_series(detail,c).map(_fmt_money)
        if "Date" in detail.columns: detail["Date"] = detail["Date"].astype(str)
        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"
        ] if c in detail.columns]
        sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in detail.columns]
        st.dataframe(_uniquify_columns(detail.sort_values(by=sort_cols)[show_cols].reset_index(drop=True)),
                     use_container_width=True)

# ---------------------- ESCROW ------------------------------
with tabs[2]:
    st.subheader("🏦 Escrow — suivi des honoraires en attente de transfert")
    df = _read_clients(clients_path)
    if df.empty:
        st.info("Pas de données.")
    else:
        # On considère l'Escrow comme la partie "Montant honoraires" encaissée (Payé) mais non encore transférée.
        # Pour rester simple ici, on affiche juste les dossiers non soldés + bouton pour marquer un transfert (note).
        en_cours = df[df["Reste"] > 0].copy()
        if en_cours.empty:
            st.success("Tous les dossiers sont soldés ✅")
        else:
            en_cours["Honoraires"] = _safe_num_series(en_cours,"Montant honoraires (US $)")
            en_cours["Payé"]      = _safe_num_series(en_cours,"Payé")
            en_cours["À transférer (indicatif)"] = (en_cours["Payé"]).clip(lower=0.0)
            show = en_cours[["Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa",
                             "Honoraires","Payé","Reste","À transférer (indicatif)"]].copy()
            for c in ["Honoraires","Payé","Reste","À transférer (indicatif)"]:
                show[c] = show[c].astype(float).map(_fmt_money)
            st.dataframe(show.reset_index(drop=True), use_container_width=True)
            st.caption("Astuce : utilise l’onglet *Modifier* dans la barre latérale pour ajouter des paiements jusqu’au solde.")

# ---------------------- VISA (aperçu) -----------------------
with tabs[3]:
    st.subheader("📄 Référentiel Visa (lecture Excel : cellule = 1 → option)")
    visa_map = parse_visa_sheet(visa_path)
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