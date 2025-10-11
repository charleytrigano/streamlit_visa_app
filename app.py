from __future__ import annotations

# ========= Imports =========
import streamlit as st
import pandas as pd
import json
from datetime import date, datetime
from pathlib import Path
from io import BytesIO
import openpyxl
import zipfile
import unicodedata
from uuid import uuid4
import altair as alt

# ========= Page & Namespace anti-collisions =========
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

SID = st.session_state.setdefault("WIDGET_NS", str(uuid4()))

# ========= Constantes =========
CLIENTS_FILE_DEFAULT = "donnees_visa_clients1_adapte.xlsx"
VISA_FILE_DEFAULT    = "donnees_visa_clients1.xlsx"

SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

CLIENTS_COLS = [
    "Dossier N","ID_Client","Nom","Date","Mois",
    "Categorie","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
    "Payé","Reste","Paiements","Options",
    # colonnes d’état si présentes :
    "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"
]

TOGGLE_COLUMNS = {
    "AOS","CP","USCIS","I-130","I-140","I-140 & AOS","I-829","I-407",
    "Work Permit","Re-entry Permit","Consultation","Analysis","Referral",
    "Derivatives","Travel Permit","USC","LPR","Perm"
}

# ========= Utilitaires =========
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

# Crée les fichiers min si absents
ensure_file(CLIENTS_FILE_DEFAULT, SHEET_CLIENTS, CLIENTS_COLS)
ensure_file(VISA_FILE_DEFAULT, SHEET_VISA, ["Categorie","Sous-categorie 1"])

def load_raw_visa_df(xlsx_path: str | Path, sheet_name: str = SHEET_VISA) -> pd.DataFrame:
    try:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()
    return _uniquify_columns(df)

def _find_cat_sub_columns(df: pd.DataFrame) -> tuple[str|None, str|None]:
    def _norm(s: str) -> str:
        s2 = unicodedata.normalize("NFKD", s)
        s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
        s2 = s2.lower().strip().replace("\u00a0"," ")
        s2 = s2.replace("-", " ").replace("_"," ")
        return " ".join(s2.split())
    cmap = {_norm(c): c for c in df.columns}
    cat_col = next((cmap[k] for k in cmap if "categorie" in k), None)
    sub_col = next((cmap[k] for k in cmap if k.startswith("sous")), None)
    return cat_col, sub_col

# ========= Parse Visa (cellule = 1 => option) + F-1/F-2 Etudiants/2-Etudiants =========
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    """
    Construit:
    {
      "Categorie": {
         "Sous-categorie": ["Sous-categorie COS","Sous-categorie EOS", ...]  # cellule = 1
      }
    }
    Injection F-1/F-2 (COS/EOS) si nom catégorie contient 'etudiant' (Etudiants, 2-Etudiants, etc.)
    """
    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        if isinstance(v, (int, float)):
            return float(v) == 1.0
        s = str(v).strip().lower()
        return s in {"1","x","true","vrai","oui","yes"}

    def _norm(s: str) -> str:
        s2 = unicodedata.normalize("NFKD", s)
        s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
        s2 = s2.strip().lower().replace("\u00a0", " ")
        s2 = s2.replace("-", " ").replace("_", " ")
        return " ".join(s2.split())

    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        return {}

    sheets_to_try = [sheet_name] if sheet_name else []
    if not sheets_to_try:
        try:
            with pd.ExcelFile(xlsx_path) as x:
                sheets_to_try = x.sheet_names
        except Exception:
            return {}

    for sn in sheets_to_try:
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

        # Étudiants / 2-Etudiants → injecte F-1/F-2 COS/EOS
        cats_in_sheet = sorted(set(dfv[cat_col].dropna().astype(str).str.strip()))
        student_cats = [c for c in cats_in_sheet if "etudiant" in _norm(c)]
        for cat_name in student_cats:
            subs = out.setdefault(cat_name, {})
            for sub in ("F-1","F-2"):
                arr = subs.setdefault(sub, [])
                for w in (f"{sub} COS", f"{sub} EOS"):
                    if w not in arr:
                        arr.append(w)
                subs[sub] = sorted(set(arr))

        if out:
            for cat, subs in out.items():
                for sub, arr in subs.items():
                    subs[sub] = sorted(set(arr))
            return out

    return {}

# ========= Clients I/O & Normalisation =========
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

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df["Mois"] = df.apply(
        lambda r: f"{pd.to_datetime(r['Date']).month:02d}" if pd.notna(r["Date"]) else (_safe_str(r.get("Mois",""))[:2] or None),
        axis=1
    )
    # numériques
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        df[c] = _safe_num_series(df, c)

    # paiements JSON list
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

    # options JSON dict
    df["Options"] = df["Options"].apply(_normalize_options_json)

    df["_Année_"]   = df["Date"].apply(lambda d: d.year if pd.notna(d) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda d: d.month if pd.notna(d) else pd.NA)
    return _uniquify_columns(df)

def _read_clients(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        df = pd.DataFrame(columns=CLIENTS_COLS)
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
    base = _safe_str(nom).strip().replace(" ","_")
    return f"{base}-{d:%Y%m%d}"

# ========= Barre latérale : chemins & action =========
with st.sidebar:
    st.header("🧭 Navigation")
    clients_path = st.text_input("Fichier Clients", value=CLIENTS_FILE_DEFAULT, key=f"sb_clients_path_{SID}")
    visa_path    = st.text_input("Fichier Visa",    value=VISA_FILE_DEFAULT,    key=f"sb_visa_path_{SID}")
    st.markdown("---")
    st.subheader("👤 Gestion Clients")
    action = st.radio("Action", options=["Ajouter","Modifier","Supprimer"], key=f"sb_action_{SID}")

# Charge données
visa_map = parse_visa_sheet(visa_path)
df_all   = _read_clients(clients_path)

# ========= Onglets =========
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "📄 Visa (aperçu)"])

# ========= Dashboard =========
with tabs[0]:
    st.subheader("📊 Dashboard — tous les clients")

    # Filtres dans la sidebar (namespace)
    with st.sidebar:
        st.markdown("---")
        st.subheader("🔎 Filtres Dashboard")
        years  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()]) if not df_all.empty else []
        months = [f"{m:02d}" for m in range(1,13)]
        cats   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if not df_all.empty else []
        subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if not df_all.empty else []
        visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist())      if not df_all.empty else []

        dash_years  = st.multiselect("Année", years, default=[], key=f"dash_years_{SID}")
        dash_months = st.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{SID}")
        dash_cats   = st.multiselect("Catégories", cats, default=[], key=f"dash_cats_{SID}")
        dash_subs   = st.multiselect("Sous-catégories", subs, default=[], key=f"dash_subs_{SID}")
        dash_visas  = st.multiselect("Visa", visas, default=[], key=f"dash_visas_{SID}")

    df = df_all.copy()
    if dash_years:  df = df[df["_Année_"].isin(dash_years)]
    if dash_months: df = df[df["Mois"].isin(dash_months)]
    if dash_cats:   df = df[df["Categorie"].astype(str).isin(dash_cats)]
    if dash_subs:   df = df[df["Sous-categorie"].astype(str).isin(dash_subs)]
    if dash_visas:  df = df[df["Visa"].astype(str).isin(dash_visas)]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(df)}")
    k2.metric("Honoraires", _fmt_money(_safe_num_series(df,"Montant honoraires (US $)").sum()))
    k3.metric("Payé", _fmt_money(_safe_num_series(df,"Payé").sum()))
    k4.metric("Solde", _fmt_money(_safe_num_series(df,"Reste").sum()))

    # Tableau
    view = df.copy()
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        if c in view.columns:
            view[c] = _safe_num_series(view,c).map(_fmt_money)
    if "Date" in view.columns:
        view["Date"] = view["Date"].astype(str)
    view["Options (résumé)"] = view["Options"].apply(
        lambda d: f"[{(d or {}).get('exclusive')}] + {', '.join((d or {}).get('options', []))}" if isinstance(d, dict) else ""
    )
    show_cols = [c for c in [
        "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
        "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste","Options (résumé)"
    ] if c in view.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Sous-categorie","Nom"] if c in view.columns]
    view = view.sort_values(by=sort_cols) if sort_cols else view
    st.dataframe(_uniquify_columns(view[show_cols].reset_index(drop=True)), use_container_width=True)

    st.markdown("---")
    st.markdown("### ✏️ Gestion (selon l’action de la barre latérale)")

    # ----- Ajouter -----
    if action == "Ajouter":
        st.markdown("#### ➕ Ajouter un client")
        c1,c2,c3 = st.columns(3)
        with c1:
            nom = st.text_input("Nom", key=f"add_nom_{SID}")
            dcr = st.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        with c2:
            cats = sorted(list(visa_map.keys()))
            cat = st.selectbox("Catégorie", options=[""]+cats, index=0, key=f"add_cat_{SID}")
            subs = sorted(list(visa_map.get(cat, {}).keys())) if cat else []
            sub  = st.selectbox("Sous-catégorie", options=[""]+subs, index=0, key=f"add_sub_{SID}")
        with c3:
            visa_final, info_msg, opts = ("","",{"exclusive": None, "options": []})
            if cat and sub:
                visa_final, info_msg, opts = st.container()._parent.render(
                    # hack avoided; call helper directly:
                )
            # Appel normal :
            visa_final, info_msg, opts = (render_dynamic_steps(cat, sub, f"add_steps_{SID}", visa_file=visa_path, preselected=None)
                                          if (cat and sub) else ("","",{"exclusive": None,"options": []}))
            if info_msg: st.info(info_msg)
            hono = st.number_input("Montant honoraires (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"add_hono_{SID}")
            autre= st.number_input("Autres frais (US $)",     min_value=0.0, step=10.0, format="%.2f", key=f"add_autre_{SID}")

        if st.button("💾 Créer", key=f"btn_add_create_{SID}"):
            if not nom or not cat or not sub:
                st.warning("Nom, Catégorie et Sous-catégorie sont requis."); st.stop()
            base = _read_clients(clients_path)
            dossier = _next_dossier(base)
            cid_base = _make_client_id(nom, dcr); cid = cid_base; i = 0
            while (base["ID_Client"].astype(str) == cid).any():
                i += 1; cid = f"{cid_base}-{i}"
            visa_val = visa_final if visa_final else sub
            total = float(hono) + float(autre)
            row = {
                "Dossier N": dossier,
                "ID_Client": cid,
                "Nom": nom,
                "Date": pd.to_datetime(dcr).date(),
                "Mois": f"{dcr.month:02d}",
                "Categorie": cat,
                "Sous-categorie": sub,
                "Visa": visa_val,
                "Montant honoraires (US $)": float(hono),
                "Autres frais (US $)": float(autre),
                "Total (US $)": total,
                "Payé": 0.0,
                "Reste": total,
                "Paiements": json.dumps([], ensure_ascii=False),
                "Options": json.dumps(_normalize_options_json(opts), ensure_ascii=False),
            }
            base = pd.concat([base, pd.DataFrame([row])], ignore_index=True)
            _write_clients(_normalize_clients(base), clients_path)
            st.success("✅ Client créé."); st.rerun()

    # ----- Modifier -----
    elif action == "Modifier":
        st.markdown("#### 🛠️ Modifier un client")
        if df_all.empty:
            st.info("Aucun client.")
        else:
            idx = st.selectbox("Sélectionne la ligne à modifier", options=list(df_all.index),
                               format_func=lambda i: f"{df_all.loc[i,'Nom']} — {df_all.loc[i,'ID_Client']}",
                               key=f"mod_idx_{SID}")
            row = df_all.loc[idx]

            c1,c2,c3 = st.columns(3)
            with c1:
                nom = st.text_input("Nom", value=_safe_str(row["Nom"]), key=f"mod_nom_{idx}_{SID}")
                dcr = st.date_input("Date de création",
                                    value=(pd.to_datetime(row["Date"]).date() if pd.notna(row["Date"]) else date.today()),
                                    key=f"mod_date_{idx}_{SID}")
            with c2:
                cats = sorted(list(visa_map.keys()))
                cur_cat = _safe_str(row["Categorie"])
                cat = st.selectbox("Catégorie", options=[""]+cats,
                                   index=(cats.index(cur_cat)+1 if cur_cat in cats else 0),
                                   key=f"mod_cat_{idx}_{SID}")
                subs = sorted(list(visa_map.get(cat, {}).keys())) if cat else []
                cur_sub = _safe_str(row["Sous-categorie"])
                sub = st.selectbox("Sous-catégorie", options=[""]+subs,
                                   index=(subs.index(cur_sub)+1 if cur_sub in subs else 0),
                                   key=f"mod_sub_{idx}_{SID}")
            with c3:
                cur_opts = _normalize_options_json(row.get("Options", {}))
                visa_final, info_msg, opts = render_dynamic_steps(cat, sub, f"mod_steps_{idx}_{SID}",
                                                                  visa_file=visa_path, preselected=cur_opts)
                if info_msg: st.info(info_msg)
                hono = st.number_input("Montant honoraires (US $)", min_value=0.0,
                                       value=float(row["Montant honoraires (US $)"]), step=10.0, format="%.2f",
                                       key=f"mod_hono_{idx}_{SID}")
                autre= st.number_input("Autres frais (US $)", min_value=0.0,
                                       value=float(row["Autres frais (US $)"]), step=10.0, format="%.2f",
                                       key=f"mod_autre_{idx}_{SID}")

            st.markdown("##### 💳 Paiements")
            p1,p2,p3,p4 = st.columns([1,1,1,2])
            with p1: pdt = st.date_input("Date paiement", value=date.today(), key=f"mod_paydt_{idx}_{SID}")
            with p2: pmd = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"mod_mode_{idx}_{SID}")
            with p3: pmt = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"mod_amt_{idx}_{SID}")
            with p4:
                if st.button("➕ Ajouter paiement", key=f"mod_addpay_{idx}_{SID}"):
                    base = _read_clients(clients_path)
                    plist = base.loc[idx,"Paiements"]
                    if not isinstance(plist, list):
                        try:
                            plist = json.loads(_safe_str(plist) or "[]")
                        except Exception:
                            plist = []
                    if float(pmt) > 0:
                        plist.append({"date": str(pdt), "mode": pmd, "amount": float(pmt)})
                        base.loc[idx,"Paiements"] = plist
                        base = _normalize_clients(base)
                        _write_clients(base, clients_path)
                        st.success("Paiement ajouté."); st.rerun()
                    else:
                        st.warning("Montant > 0 requis.")

            hist = row["Paiements"] if isinstance(row["Paiements"], list) else []
            if hist:
                h = pd.DataFrame(hist)
                if "amount" in h.columns:
                    h["amount"] = h["amount"].astype(float).map(_fmt_money)
                st.dataframe(h, use_container_width=True)
            else:
                st.caption("Aucun paiement.")

            if st.button("💾 Sauvegarder", key=f"mod_save_{idx}_{SID}"):
                base = _read_clients(clients_path)
                base.loc[idx,"Nom"] = nom
                base.loc[idx,"Date"] = pd.to_datetime(dcr).date()
                base.loc[idx,"Mois"] = f"{dcr.month:02d}"
                base.loc[idx,"Categorie"] = cat
                base.loc[idx,"Sous-categorie"] = sub
                base.loc[idx,"Visa"] = visa_final if visa_final else (sub or "")
                base.loc[idx,"Montant honoraires (US $)"] = float(hono)
                base.loc[idx,"Autres frais (US $)"] = float(autre)
                base.loc[idx,"Total (US $)"] = float(hono) + float(autre)
                base.loc[idx,"Options"] = opts
                _write_clients(_normalize_clients(base), clients_path)
                st.success("✅ Modifications enregistrées."); st.rerun()

    # ----- Supprimer -----
    elif action == "Supprimer":
        st.markdown("#### 🗑️ Supprimer un client")
        if df_all.empty:
            st.info("Aucun client.")
        else:
            idx = st.selectbox("Sélectionne la ligne à supprimer", options=list(df_all.index),
                               format_func=lambda i: f"{df_all.loc[i,'Nom']} — {df_all.loc[i,'ID_Client']}",
                               key=f"del_idx_{SID}")
            if st.button("Confirmer la suppression", type="primary", key=f"btn_confirm_del_{SID}"):
                base = _read_clients(clients_path)
                base = base.drop(index=idx).reset_index(drop=True)
                _write_clients(_normalize_clients(base), clients_path)
                st.success("Client supprimé."); st.rerun()

# ========= Analyses + Graphiques =========
with tabs[1]:
    st.subheader("📈 Analyses")
    df = df_all.copy()
    if df.empty:
        st.info("Pas de données.")
    else:
        a1, a2, a3, a4 = st.columns(4)
        years  = sorted([int(y) for y in pd.to_numeric(df["_Année_"], errors="coerce").dropna().unique().tolist()])
        months = [f"{m:02d}" for m in range(1,13)]
        cats   = sorted(df["Categorie"].dropna().astype(str).unique().tolist())
        visas  = sorted(df["Visa"].dropna().astype(str).unique().tolist())

        fy = a1.multiselect("Année", years, default=[], key=f"ana_years_{SID}")
        fm = a2.multiselect("Mois (MM)", months, default=[], key=f"ana_months_{SID}")
        fc = a3.multiselect("Catégories", cats, default=[], key=f"ana_cats_{SID}")
        fv = a4.multiselect("Visa", visas, default=[], key=f"ana_visas_{SID}")

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

        # ---------- Graphiques ----------
        # 1) Volumes par mois (barres)
        if not f.empty:
            gd = f.copy()
            gd["_YYYYMM_"] = gd.apply(lambda r: f"{int(r['_Année_']):04d}-{int(r['_MoisNum_']):02d}" if pd.notna(r["_Année_"]) and pd.notna(r["_MoisNum_"]) else None, axis=1)
            g_vol = gd.groupby("_YYYYMM_", dropna=True).size().reset_index(name="Dossiers")
            g_vol = g_vol[g_vol["_YYYYMM_"].notna()]
            if not g_vol.empty:
                chart_vol = alt.Chart(g_vol).mark_bar().encode(
                    x=alt.X("_YYYYMM_", sort=None, title="Année-Mois"),
                    y=alt.Y("Dossiers", title="Nombre de dossiers"),
                    tooltip=["_YYYYMM_","Dossiers"]
                ).properties(height=280)
                st.altair_chart(chart_vol, use_container_width=True)

        # 2) Honoraires & Encaissements par mois (lignes)
        if not f.empty:
            gg = f.copy()
            gg["_YYYYMM_"] = gg.apply(lambda r: f"{int(r['_Année_']):04d}-{int(r['_MoisNum_']):02d}" if pd.notna(r["_Année_"]) and pd.notna(r["_MoisNum_"]) else None, axis=1)
            gg["Honoraires"]   = _safe_num_series(gg, "Montant honoraires (US $)")
            gg["Encaissements"]= _safe_num_series(gg, "Payé")
            g2 = gg.groupby("_YYYYMM_", dropna=True)[["Honoraires","Encaissements"]].sum().reset_index()
            g2 = g2[g2["_YYYYMM_"].notna()]
            if not g2.empty:
                g2_m = g2.melt("_YYYYMM_", var_name="Type", value_name="Montant")
                chart_ca = alt.Chart(g2_m).mark_line(point=True).encode(
                    x=alt.X("_YYYYMM_", sort=None, title="Année-Mois"),
                    y=alt.Y("Montant", title="US $"),
                    color="Type",
                    tooltip=["_YYYYMM_","Type","Montant"]
                ).properties(height=280)
                st.altair_chart(chart_ca, use_container_width=True)

        st.markdown("#### Détails (clients filtrés)")
        detail = f.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in detail.columns:
                detail[c] = _safe_num_series(detail,c).map(_fmt_money)
        if "Date" in detail.columns:
            detail["Date"] = detail["Date"].astype(str)
        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"
        ] if c in detail.columns]
        sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Sous-categorie","Nom"] if c in detail.columns]
        detail = detail.sort_values(by=sort_cols) if sort_cols else detail
        st.dataframe(_uniquify_columns(detail[show_cols].reset_index(drop=True)), use_container_width=True)

# ========= Escrow =========
with tabs[2]:
    st.subheader("🏦 Escrow — honoraires en attente de transfert")
    df = df_all.copy()
    if df.empty:
        st.info("Pas de données.")
    else:
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
            st.caption("Astuce : ajoute des acomptes via **Modifier** jusqu’au solde.")

# ========= Visa (aperçu) =========
with tabs[3]:
    st.subheader("📄 Référentiel Visa (cellule = 1 → option)")
    vmap = parse_visa_sheet(visa_path)
    if not vmap:
        st.warning("Aucune donnée Visa trouvée. Vérifie le fichier et l’onglet.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            cat = st.selectbox("Catégorie", [""]+sorted(list(vmap.keys())), key=f"vz_cat_{SID}")
        with c2:
            subs = sorted(list(vmap.get(cat, {}).keys())) if cat else []
            sub  = st.selectbox("Sous-catégorie", [""]+subs, key=f"vz_sub_{SID}")
        if cat and sub:
            opts = sorted(list(vmap.get(cat, {}).get(sub, [])))
            st.write("**Options détectées** :", ", ".join(opts) if opts else "(Aucune → le visa = sous-catégorie)")
        with st.expander("Aperçu complet"):
            for k, submap in vmap.items():
                st.write(f"**{k}**")
                for s, arr in submap.items():
                    st.caption(f"- {s} → {', '.join(arr)}")

# ========= Exports =========
st.markdown("---")
st.subheader("📤 Exports rapides")

def _excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        _uniquify_columns(df).to_excel(wr, sheet_name=sheet_name, index=False)
    bio.seek(0)
    return bio.getvalue()

colE1, colE2, colE3 = st.columns(3)
with colE1:
    try:
        st.download_button(
            "⬇️ Télécharger Clients.xlsx",
            data=_excel_bytes(df_all, SHEET_CLIENTS),
            file_name="Clients.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_clients_xlsx_{SID}",
        )
    except Exception as e:
        st.caption(f"Export Clients.xlsx indisponible : {e}")

with colE2:
    try:
        st.download_button(
            "⬇️ Télécharger Visa.xlsx",
            data=Path(visa_path).read_bytes(),
            file_name="Visa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_visa_xlsx_{SID}",
        )
    except Exception as e:
        st.caption(f"Export Visa.xlsx indisponible : {e}")

with colE3:
    try:
        bio = BytesIO()
        with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("Clients.xlsx", _excel_bytes(df_all, SHEET_CLIENTS))
            try:
                zf.write(visa_path, arcname="Visa.xlsx")
            except Exception:
                vdf = load_raw_visa_df(visa_path, SHEET_VISA)
                if not vdf.empty:
                    zf.writestr("Visa.xlsx", _excel_bytes(vdf, SHEET_VISA))
        bio.seek(0)
        st.download_button(
            "📦 ZIP : Clients + Visa",
            data=bio.getvalue(),
            file_name="Visa_Manager_Export.zip",
            mime="application/zip",
            key=f"dl_zip_all_{SID}",
        )
    except Exception as e:
        st.caption(f"Export ZIP indisponible : {e}")