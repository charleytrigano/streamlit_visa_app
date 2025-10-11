# =========================
# APP VISA MANAGER — PARTIE 1/3
# (imports, constantes, utilitaires, parsing Visa, I/O clients)
# =========================

from __future__ import annotations

# ---- Imports
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

# Espace de noms unique pour éviter les collisions de widgets
SID = st.session_state.setdefault("WIDGET_NS", str(uuid4()))

# =========================
# Constantes
# =========================
CLIENTS_FILE_DEFAULT = "donnees_visa_clients1_adapte.xlsx"
VISA_FILE_DEFAULT    = "donnees_visa_clients1.xlsx"

SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

# Colonnes attendues côté Clients (sans accents, et avec nouveaux statuts + dates)
CLIENTS_COLS = [
    "Dossier N","ID_Client","Nom","Date","Mois",
    "Categorie","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
    "Payé","Reste","Paiements","Options",
    # ---- Statuts + dates (nouveau)
    "Dossier envoyé","Date d'envoi",
    "Dossier accepté","Date d'acceptation",
    "Dossier refusé","Date de refus",
    "Dossier annulé","Date d'annulation",
    # RFE (ne peut être VRAI que si au moins un autre statut est coché)
    "RFE"
]

# =========================
# Utilitaires génériques
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

# Crée des fichiers vides si absents (structure OK)
ensure_file(CLIENTS_FILE_DEFAULT, SHEET_CLIENTS, CLIENTS_COLS)
# Pour Visa on met juste 2 colonnes minimales au cas où
ensure_file(VISA_FILE_DEFAULT, SHEET_VISA, ["Categorie","Sous-categorie 1"])

# =========================
# Normalisation de texte (sans accents)
# =========================
def _norm(s: str) -> str:
    s2 = unicodedata.normalize("NFKD", s)
    s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
    s2 = s2.strip().lower().replace("\u00a0", " ")
    s2 = s2.replace("-", " ").replace("_", " ")
    return " ".join(s2.split())

# =========================
# Parsing de la feuille Visa
# - Lit l’onglet Visa (cases = 1)
# - Produit un mapping: {Categorie: {Sous-categorie: [options...]}}
# - Injection automatique "2-Etudiants" → F-1/F-2 COS/EOS si rien de 'etudiant' trouvé
# =========================
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    """
    Construit un mapping:
    {
      "Categorie": {
         "Sous-categorie": ["Sous-categorie COS","Sous-categorie EOS", ...]  # colonnes dont la cellule = 1
      }
    }
    Injection automatique F-1/F-2 (COS/EOS) pour toute catégorie dont le nom contient 'etudiant'
    (ex: 'Etudiants', '2-Etudiants', etc.). Si aucune catégorie 'étudiants' n'est trouvée
    dans le fichier, on AJOUTE quand même '2-Etudiants' avec F-1/F-2 COS/EOS par défaut.
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

    # Lecture du classeur
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

            # Détection présence d’une catégorie 'étudiants'
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
                    # aucune case cochée → on retient la sous-catégorie brute
                    opts = [sub]
                if opts:
                    out.setdefault(cat, {})
                    out[cat].setdefault(sub, [])
                    out[cat][sub].extend(opts)

    # Injection automatique pour catégories existantes contenant 'etudiant'
    for cat_name in list(out.keys()):
        if "etudiant" in _norm(cat_name):
            submap = out.setdefault(cat_name, {})
            for sub in ("F-1","F-2"):
                arr = submap.setdefault(sub, [])
                for w in (f"{sub} COS", f"{sub} EOS"):
                    if w not in arr:
                        arr.append(w)
                submap[sub] = sorted(set(arr))

    # Si AUCUNE catégorie 'étudiants' trouvée, on ajoute quand même 2-Etudiants
    if not found_students:
        out.setdefault("2-Etudiants", {})
        out["2-Etudiants"].setdefault("F-1", ["F-1 COS", "F-1 EOS"])
        out["2-Etudiants"].setdefault("F-2", ["F-2 COS", "F-2 EOS"])

    # Nettoyage doublons + tri
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
    # colonnes manquantes
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

    # Paiements (liste JSON)
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

    # Statuts -> booléens propres
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
    df = df.copy()
    df["Options"] = df["Options"].apply(lambda d: json.dumps(_normalize_options_json(d), ensure_ascii=False))
    df["Paiements"] = df["Paiements"].apply(lambda l: json.dumps(l, ensure_ascii=False))
    # booleans -> 1/0 pour stockage
    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        df[c] = df[c].apply(lambda b: 1 if bool(b) else 0)
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


# =========================
# APP VISA MANAGER — PARTIE 2/3
# (UI, Dashboard, Ajout / Modif / Suppression, statuts + dates + règle RFE)
# =========================

# ---------- Rendu dynamique des options Visa depuis la feuille Visa ----------
def render_dynamic_steps(cat: str, sub: str, keyprefix: str, visa_file: str, preselected: dict|None=None):
    """
    Affiche les options dynamiques issues de la feuille Visa pour (cat, sub).
    - Si les options sont de la forme "<sub> XXX" (ex: "F-1 COS", "F-1 EOS"),
      on propose un choix exclusif (radio) sur les suffixes (COS/EOS).
    - Les autres options (non-prefixées par "sub ") sont rendues en cases à cocher.
    Retourne (visa_final, info, options_dict)
    """
    vmap = parse_visa_sheet(visa_file)
    submap = vmap.get(cat, {}) if vmap else {}
    all_opts = submap.get(sub, [])

    info = ""
    pre = preselected or {}
    # normalise preselected
    try:
        if isinstance(pre, str):
            pre = json.loads(pre or "{}")
    except Exception:
        pre = {}
    excl_pre = pre.get("exclusive")
    opts_pre = set(pre.get("options", []))

    # sépare "<sub> SFX" vs "autres"
    prefix = f"{sub} "
    suffixes, others = [], []
    for o in all_opts:
        if o.startswith(prefix) and len(o) > len(prefix):
            suffixes.append(o[len(prefix):])   # garde SFX
        else:
            others.append(o)

    # Exclusif
    chosen_excl = None
    if suffixes:
        sup_opts = ["" ] + sorted(set(suffixes))
        chosen_excl = st.radio(
            f"Option exclusive — {sub}",
            options=sup_opts,
            index=(sup_opts.index(excl_pre) if excl_pre in sup_opts else 0),
            key=f"{keyprefix}_excl_{SID}"
        )
        chosen_excl = chosen_excl or None

    # Multiples
    chosen_multi = []
    for i, lab in enumerate(sorted(set(others))):
        default = (lab in opts_pre)
        ok = st.checkbox(lab, value=default, key=f"{keyprefix}_chk_{i}_{SID}")
        if ok:
            chosen_multi.append(lab)

    options_dict = {"exclusive": chosen_excl, "options": chosen_multi}

    # Construction visa final
    if chosen_excl:
        visa_final = f"{sub} {chosen_excl}"
    else:
        visa_final = sub

    if not all_opts:
        info = "Aucune option cochée dans la feuille Visa pour cette sous-catégorie — le Visa sera la sous-catégorie elle-même."

    return visa_final, info, options_dict

# ---------- Barre latérale (fichiers + action) ----------
with st.sidebar:
    st.header("🧭 Navigation")
    clients_path = st.text_input("Fichier Clients", value=CLIENTS_FILE_DEFAULT, key=f"sb_clients_path_{SID}")
    visa_path    = st.text_input("Fichier Visa",    value=VISA_FILE_DEFAULT,    key=f"sb_visa_path_{SID}")
    st.markdown("---")
    st.subheader("👤 Gestion")
    action = st.radio("Action", options=["Ajouter","Modifier","Supprimer"], key=f"sb_action_{SID}")

# Chargement des données
visa_map = parse_visa_sheet(visa_path)
df_all   = _read_clients(clients_path)

# ---------- Onglets (partie 2 remplit uniquement le Dashboard) ----------
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "📄 Visa (aperçu)"])

# =========================
# 📊 DASHBOARD
# =========================
with tabs[0]:
    st.subheader("📊 Dashboard — tous les clients")

    # Filtres (dans la sidebar pour rester discrets)
    with st.sidebar:
        st.markdown("---")
        st.subheader("🔎 Filtres Dashboard")
        if df_all.empty:
            years = []; months = []; cats = []; subs = []; visas = []
        else:
            years  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
            months = [f"{m:02d}" for m in range(1,13)]
            cats   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist())
            subs   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist())
            visas  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist())

        dash_years  = st.multiselect("Année", years, default=[], key=f"dash_years_{SID}")
        dash_months = st.multiselect("Mois (MM)", months, default=[], key=f"dash_months_{SID}")
        dash_cats   = st.multiselect("Catégories", cats, default=[], key=f"dash_cats_{SID}")
        dash_subs   = st.multiselect("Sous-catégories", subs, default=[], key=f"dash_subs_{SID}")
        dash_visas  = st.multiselect("Visa", visas, default=[], key=f"dash_visas_{SID}")

    # applique filtres
    df = df_all.copy()
    if dash_years:  df = df[df["_Année_"].isin(dash_years)]
    if dash_months: df = df[df["Mois"].isin(dash_months)]
    if dash_cats:   df = df[df["Categorie"].astype(str).isin(dash_cats)]
    if dash_subs:   df = df[df["Sous-categorie"].astype(str).isin(dash_subs)]
    if dash_visas:  df = df[df["Visa"].astype(str).isin(dash_visas)]

    # KPIs
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
        "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
        "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE",
        "Options (résumé)"
    ] if c in view.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Sous-categorie","Nom"] if c in view.columns]
    view = view.sort_values(by=sort_cols) if sort_cols else view
    st.dataframe(_uniquify_columns(view[show_cols].reset_index(drop=True)), use_container_width=True)

    st.markdown("---")
    st.markdown("### ✏️ Gestion (voir la barre latérale)")

    # ---------- AJOUTER ----------
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
            if cat and sub:
                visa_final, info_msg, opts = render_dynamic_steps(cat, sub, f"add_steps_{SID}", visa_file=visa_path, preselected=None)
                if info_msg: st.info(info_msg)
            else:
                visa_final, opts = "", {"exclusive": None, "options": []}
            hono = st.number_input("Montant honoraires (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"add_hono_{SID}")
            autre= st.number_input("Autres frais (US $)",     min_value=0.0, step=10.0, format="%.2f", key=f"add_autre_{SID}")

        st.markdown("##### 📌 Statuts du dossier")
        s1,s2,s3,s4,s5 = st.columns(5)
        with s1:
            st_envoye = st.checkbox("Dossier envoyé", key=f"add_env_{SID}")
            dt_envoye = st.date_input("Date d’envoi", value=date.today(), key=f"add_dt_env_{SID}") if st_envoye else None
        with s2:
            st_accepte = st.checkbox("Dossier accepté", key=f"add_acc_{SID}")
            dt_accepte = st.date_input("Date d’acceptation", value=date.today(), key=f"add_dt_acc_{SID}") if st_accepte else None
        with s3:
            st_refuse = st.checkbox("Dossier refusé", key=f"add_ref_{SID}")
            dt_refuse = st.date_input("Date de refus", value=date.today(), key=f"add_dt_ref_{SID}") if st_refuse else None
        with s4:
            st_annule = st.checkbox("Dossier annulé", key=f"add_ann_{SID}")
            dt_annule = st.date_input("Date d’annulation", value=date.today(), key=f"add_dt_ann_{SID}") if st_annule else None
        with s5:
            st_rfe = st.checkbox("RFE", key=f"add_rfe_{SID}")

        if st.button("💾 Créer", key=f"btn_add_create_{SID}"):
            if not nom or not cat or not sub:
                st.warning("Nom, Catégorie et Sous-catégorie sont requis."); st.stop()

            # Règle RFE : doit être couplée à au moins un autre statut
            if st_rfe and not (st_envoye or st_accepte or st_refuse or st_annule):
                st.error("RFE ne peut être cochée que si au moins un autre statut est coché."); st.stop()

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
                "Options": json.dumps(opts, ensure_ascii=False),
                "Dossier envoyé": 1 if st_envoye else 0,
                "Date d'envoi": dt_envoye if st_envoye else None,
                "Dossier accepté": 1 if st_accepte else 0,
                "Date d'acceptation": dt_accepte if st_accepte else None,
                "Dossier refusé": 1 if st_refuse else 0,
                "Date de refus": dt_refuse if st_refuse else None,
                "Dossier annulé": 1 if st_annule else 0,
                "Date d'annulation": dt_annule if st_annule else None,
                "RFE": 1 if st_rfe else 0,
            }
            base = pd.concat([base, pd.DataFrame([row])], ignore_index=True)
            _write_clients(_normalize_clients(base), clients_path)
            st.success("✅ Client créé."); st.rerun()

    # ---------- MODIFIER ----------
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
                cur_opts = row.get("Options", {})
                visa_final, info_msg, opts = render_dynamic_steps(cat, sub, f"mod_steps_{idx}_{SID}",
                                                                  visa_file=visa_path, preselected=cur_opts)
                if info_msg: st.info(info_msg)
                hono = st.number_input("Montant honoraires (US $)", min_value=0.0,
                                       value=float(row["Montant honoraires (US $)"]), step=10.0, format="%.2f",
                                       key=f"mod_hono_{idx}_{SID}")
                autre= st.number_input("Autres frais (US $)", min_value=0.0,
                                       value=float(row["Autres frais (US $)"]), step=10.0, format="%.2f",
                                       key=f"mod_autre_{idx}_{SID}")

            st.markdown("##### 📌 Statuts du dossier")
            s1,s2,s3,s4,s5 = st.columns(5)
            with s1:
                st_envoye = st.checkbox("Dossier envoyé", value=bool(row["Dossier envoyé"]), key=f"mod_env_{idx}_{SID}")
                dt_envoye = st.date_input("Date d’envoi",
                                          value=(pd.to_datetime(row["Date d'envoi"]).date() if pd.notna(row["Date d'envoi"]) else date.today()),
                                          key=f"mod_dt_env_{idx}_{SID}") if st_envoye else None
            with s2:
                st_accepte = st.checkbox("Dossier accepté", value=bool(row["Dossier accepté"]), key=f"mod_acc_{idx}_{SID}")
                dt_accepte = st.date_input("Date d’acceptation",
                                           value=(pd.to_datetime(row["Date d'acceptation"]).date() if pd.notna(row["Date d'acceptation"]) else date.today()),
                                           key=f"mod_dt_acc_{idx}_{SID}") if st_accepte else None
            with s3:
                st_refuse = st.checkbox("Dossier refusé", value=bool(row["Dossier refusé"]), key=f"mod_ref_{idx}_{SID}")
                dt_refuse = st.date_input("Date de refus",
                                          value=(pd.to_datetime(row["Date de refus"]).date() if pd.notna(row["Date de refus"]) else date.today()),
                                          key=f"mod_dt_ref_{idx}_{SID}") if st_refuse else None
            with s4:
                st_annule = st.checkbox("Dossier annulé", value=bool(row["Dossier annulé"]), key=f"mod_ann_{idx}_{SID}")
                dt_annule = st.date_input("Date d’annulation",
                                          value=(pd.to_datetime(row["Date d'annulation"]).date() if pd.notna(row["Date d'annulation"]) else date.today()),
                                          key=f"mod_dt_ann_{idx}_{SID}") if st_annule else None
            with s5:
                st_rfe = st.checkbox("RFE", value=bool(row["RFE"]), key=f"mod_rfe_{idx}_{SID}")

            # Paiements
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
                # Règle RFE
                if st_rfe and not (st_envoye or st_accepte or st_refuse or st_annule):
                    st.error("RFE ne peut être cochée que si au moins un autre statut est coché."); st.stop()

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

                base.loc[idx,"Dossier envoyé"] = 1 if st_envoye else 0
                base.loc[idx,"Date d'envoi"] = dt_envoye if st_envoye else None
                base.loc[idx,"Dossier accepté"] = 1 if st_accepte else 0
                base.loc[idx,"Date d'acceptation"] = dt_accepte if st_accepte else None
                base.loc[idx,"Dossier refusé"] = 1 if st_refuse else 0
                base.loc[idx,"Date de refus"] = dt_refuse if st_refuse else None
                base.loc[idx,"Dossier annulé"] = 1 if st_annule else 0
                base.loc[idx,"Date d'annulation"] = dt_annule if st_annule else None
                base.loc[idx,"RFE"] = 1 if st_rfe else 0

                _write_clients(_normalize_clients(base), clients_path)
                st.success("✅ Modifications enregistrées."); st.rerun()

    # ---------- SUPPRIMER ----------
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

# =========================
# APP VISA MANAGER — PARTIE 3/3
# (Analyses, Escrow, Visa aperçu, export)
# =========================

# =========================
# 📈 ANALYSES
# =========================
with tabs[1]:
    st.subheader("📈 Analyses des clients")

    if df_all.empty:
        st.info("Aucune donnée à analyser.")
    else:
        # Graphique 1 — répartition par catégorie
        dfc = df_all.copy()
        dfc_grp = dfc.groupby("Categorie", as_index=False).agg({
            "Dossier N": "count",
            "Montant honoraires (US $)": "sum",
            "Payé": "sum",
            "Reste": "sum"
        }).rename(columns={"Dossier N": "Nombre"})

        st.markdown("#### Répartition des dossiers par catégorie")
        chart1 = (
            alt.Chart(dfc_grp)
            .mark_bar()
            .encode(
                x=alt.X("Categorie:N", sort="-y", title="Catégorie"),
                y=alt.Y("Nombre:Q", title="Nombre de dossiers"),
                tooltip=["Categorie", "Nombre"]
            )
        )
        st.altair_chart(chart1, use_container_width=True)

        # Graphique 2 — Honoraires par mois
        dfm = dfc.copy()
        dfm["Mois"] = dfm["Mois"].astype(str)
        dfm_grp = dfm.groupby("Mois", as_index=False)["Montant honoraires (US $)"].sum()
        st.markdown("#### Montants d’honoraires par mois")
        chart2 = (
            alt.Chart(dfm_grp)
            .mark_line(point=True)
            .encode(
                x=alt.X("Mois:N", sort=list(map(lambda m: f"{m:02d}", range(1,13)))),
                y=alt.Y("Montant honoraires (US $):Q", title="Honoraires (US $)"),
                tooltip=["Mois", "Montant honoraires (US $)"]
            )
        )
        st.altair_chart(chart2, use_container_width=True)

        # Graphique 3 — Payé vs Reste
        dfp = pd.DataFrame({
            "Type": ["Payé", "Reste"],
            "Montant": [
                _safe_num_series(dfc,"Payé").sum(),
                _safe_num_series(dfc,"Reste").sum()
            ]
        })
        st.markdown("#### Répartition Payé / Reste")
        chart3 = (
            alt.Chart(dfp)
            .mark_arc(innerRadius=60)
            .encode(
                theta=alt.Theta(field="Montant", type="quantitative"),
                color=alt.Color(field="Type", type="nominal"),
                tooltip=["Type","Montant"]
            )
        )
        st.altair_chart(chart3, use_container_width=True)

# =========================
# 🏦 ESCROW
# =========================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse financière")

    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfp = df_all.copy()
        dfp["Payé"] = _safe_num_series(dfp,"Payé")
        dfp["Reste"] = _safe_num_series(dfp,"Reste")
        dfp["Total (US $)"] = _safe_num_series(dfp,"Total (US $)")

        agg = dfp.groupby("Categorie", as_index=False).agg({
            "Total (US $)": "sum",
            "Payé": "sum",
            "Reste": "sum"
        })
        agg["% Payé"] = (agg["Payé"] / agg["Total (US $)"] * 100).round(1)
        agg["% Reste"] = (agg["Reste"] / agg["Total (US $)"] * 100).round(1)

        st.dataframe(agg, use_container_width=True)

        st.markdown("#### Montants cumulés (tous dossiers)")
        t1,t2,t3 = st.columns(3)
        t1.metric("Total honoraires", _fmt_money(dfp["Total (US $)"].sum()))
        t2.metric("Total payé", _fmt_money(dfp["Payé"].sum()))
        t3.metric("Total restant", _fmt_money(dfp["Reste"].sum()))

# =========================
# 📄 VISA (APERÇU)
# =========================
with tabs[3]:
    st.subheader("📄 Aperçu du fichier Visa")

    try:
        visa_file = pd.read_excel(visa_path, sheet_name=SHEET_VISA)
        visa_file = _uniquify_columns(visa_file)
        st.dataframe(visa_file, use_container_width=True)
    except Exception as e:
        st.error(f"Erreur de lecture du fichier Visa : {e}")

# =========================
# 💾 EXPORT GLOBAL
# =========================
st.markdown("---")
st.markdown("### 💾 Export complet du projet")

if st.button("📦 Télécharger Clients + Visa (ZIP)", type="primary"):
    try:
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(clients_path, "Clients.xlsx")
            zf.write(visa_path, "Visa.xlsx")
        st.download_button(
            label="⬇️ Télécharger le ZIP",
            data=buffer.getvalue(),
            file_name="Visa_Manager_Export.zip",
            mime="application/zip"
        )
    except Exception as e:
        st.error(f"Erreur export : {e}")