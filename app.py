# ==============================================
# 🛂 VISA MANAGER — Application Streamlit (Partie 1/3)
# ==============================================

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

# -------- Page & titre
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

# Espace de noms unique pour éviter les collisions de clés Streamlit
SID = st.session_state.setdefault("WIDGET_NS", str(uuid4()))

# ==============================================
# Constantes & colonnes attendues
# ==============================================
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

# ==============================================
# Utilitaires généraux
# ==============================================
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

# Crée des fichiers vides si absents
ensure_file(CLIENTS_FILE_DEFAULT, SHEET_CLIENTS, CLIENTS_COLS)
ensure_file(VISA_FILE_DEFAULT, SHEET_VISA, ["Categorie","Sous-categorie 1","COS","EOS"])

def _norm(s: str) -> str:
    s2 = unicodedata.normalize("NFKD", s)
    s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
    s2 = s2.strip().lower().replace("\u00a0", " ")
    s2 = s2.replace("-", " ").replace("_", " ")
    return " ".join(s2.split())

# ==============================================
# Parsing de la feuille Visa → visa_map {cat:{sub:[options]}}
# ==============================================
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    """
    Construit un mapping: {Categorie: {Sous-categorie: [options...]}}
    - Chaque option correspond à une colonne cochée (=1, x, oui, true...) sur la ligne de la sous-catégorie.
    - Si aucune option cochée, on conserve la sous-catégorie seule comme option.
    - Injection auto '2-Etudiants' => F-1/F-2 COS/EOS si absente.
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
                        # libellé d'option = "Sous-categorie + nom de colonne"
                        lab = (f"{sub} {cc}".strip() if sub else str(cc).strip())
                        opts.append(lab)
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

    # Si aucune catégorie étudiants trouvée, on ajoute '2-Etudiants'
    if not found_students:
        out.setdefault("2-Etudiants", {})
        out["2-Etudiants"].setdefault("F-1", ["F-1 COS", "F-1 EOS"])
        out["2-Etudiants"].setdefault("F-2", ["F-2 COS", "F-2 EOS"])

    # Nettoyage doublons
    for cat, subs in out.items():
        for sub, arr in subs.items():
            subs[sub] = sorted(set(arr))
    return out

# ==============================================
# I/O & normalisation des Clients
# ==============================================
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

    # Paiements (JSON list)
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
    stack = st.session_state.get("undo_stack", [])
    if not stack:
        st.warning("Aucune opération à annuler.")
        return
    prev_df = stack.pop()
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        _uniquify_columns(prev_df).to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
    st.success("Dernière action annulée.")

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if "Dossier N" in df.columns:
        s = pd.to_numeric(df["Dossier N"], errors="coerce")
        if s.notna().any():
            return int(s.max()) + 1
    return int(start)

# --- Identifiant client robuste (prend str/date/datetime/None) ---
def _make_client_id(nom: str, d) -> str:
    base = _safe_str(nom).strip()
    base = unicodedata.normalize("NFKD", base)
    base = "".join(ch for ch in base if not unicodedata.combining(ch))
    base = base.replace(" ", "_").replace("/", "-").replace("\\", "-")
    base = base.strip("_-") or "Client"
    try:
        if not isinstance(d, (date, datetime)):
            d = pd.to_datetime(d, errors="coerce")
        if pd.isna(d):
            d = datetime.today()
        if isinstance(d, date) and not isinstance(d, datetime):
            d = datetime(d.year, d.month, d.day)
        stamp = d.strftime("%Y%m%d")
    except Exception:
        stamp = datetime.today().strftime("%Y%m%d")
    return f"{base}-{stamp}"

# ==============================================
# Helper UI — sélecteur d'options VISA (radio + cases)
# ==============================================
def build_visa_option_selector(visa_map: dict, cat: str, sub: str, keyprefix: str, preselected: dict | None = None):
    """
    UI des options Visa pour (cat, sub) :
      - Radio 'exclusive' si options de la forme 'sub XXX' (ex: 'F-1 COS', 'F-1 EOS')
      - Cases à cocher pour les autres options
    Retourne (visa_final, opts_dict, info_msg).
    """
    arr = visa_map.get(cat, {}).get(sub, [])
    prefix = f"{sub} "
    suffixes = sorted({o[len(prefix):] for o in arr if o.startswith(prefix) and len(o) > len(prefix)})
    others = sorted([o for o in arr if not (o.startswith(prefix) and len(o) > len(prefix))])

    preselected = preselected or {}
    pre_excl = preselected.get("exclusive")
    pre_opts = preselected.get("options", []) if isinstance(preselected.get("options", []), list) else []

    chosen_excl = None
    if suffixes:
        radio_opts = [""] + suffixes
        default_idx = radio_opts.index(pre_excl) if pre_excl in radio_opts else 0
        chosen_excl = st.radio(
            f"Option exclusive — {sub}",
            options=radio_opts,
            index=default_idx,
            key=f"{keyprefix}_excl"
        )

    chosen_multi = []
    for i, lab in enumerate(others):
        default = lab in pre_opts
        if st.checkbox(lab, value=default, key=f"{keyprefix}_chk_{i}"):
            chosen_multi.append(lab)

    visa_final = sub
    if chosen_excl:
        visa_final = f"{sub} {chosen_excl}".strip()

    info_msg = "" if arr else "Aucune option cochée pour cette sous-catégorie dans la feuille Visa."
    return visa_final, {"exclusive": (chosen_excl or None), "options": chosen_multi}, info_msg

# ==============================================
# Barre latérale (fichiers, actions, UNDO)
# ==============================================
with st.sidebar:
    st.header("🧭 Navigation")
    clients_path = st.text_input("Fichier Clients", value=CLIENTS_FILE_DEFAULT, key=f"sb_clients_path_{SID}")
    visa_path    = st.text_input("Fichier Visa",    value=VISA_FILE_DEFAULT,    key=f"sb_visa_path_{SID}")
    st.markdown("---")
    st.subheader("🧾 Gestion")
    action = st.radio("Action", options=["Ajouter","Modifier","Supprimer"], key=f"sb_action_{SID}")
    st.markdown("---")
    if st.button("↩️ Annuler dernière action (UNDO)", key=f"undo_{SID}"):
        undo_last_write(clients_path)
        st.cache_data.clear()
        st.rerun()

# ==============================================
# Chargement des données
# ==============================================
visa_map = parse_visa_sheet(visa_path)
df_all   = _read_clients(clients_path)

# ==============================================
# Création des onglets
# ==============================================
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "📄 Visa (aperçu)", "👤 Clients"])

# ==============================================
# 📊 DASHBOARD (début)
# ==============================================
with tabs[0]:
    st.subheader("📊 Dashboard — tous les clients")

    # Filtres dashboard (barre latérale)
    with st.sidebar:
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


# ==============================================
# 📊 DASHBOARD (suite) — tableau détaillé
# ==============================================
with tabs[0]:
    st.markdown("### 📋 Détails (après filtres)")
    view = df.copy()

    # Mise en forme lisible
    for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Total (US $)", "Payé", "Reste"]:
        if c in view.columns:
            view[c] = _safe_num_series(view, c).apply(_fmt_money)
    if "Date" in view.columns:
        try:
            view["Date"] = pd.to_datetime(view["Date"], errors="coerce").dt.date.astype(str)
        except Exception:
            view["Date"] = view["Date"].astype(str)

    show_cols = [c for c in [
        "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
        "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
        "Payé","Reste",
        "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
    ] if c in view.columns]

    sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in view.columns]
    view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

    st.dataframe(
        view_sorted[show_cols].reset_index(drop=True),
        use_container_width=True,
        key=f"df_dash_{SID}",
    )


# ==============================================
# 👤 ONGLET : Clients — vue Compte Client (consultation + ajout paiement)
# ==============================================
def _render_status_dates_block(row: pd.Series):
    """Bloc statuts + dates — pas de f-strings pour éviter les quotes dans les labels."""
    st.markdown("### 📌 Statuts & dates")
    s1, s2, s3, s4, s5 = st.columns(5)

    sent_ok = bool(row.get("Dossier envoyé"))
    s1.write("**Envoyé** : " + ("✅" if sent_ok else "❌"))
    s1.write("• Date : " + _safe_str(row.get("Date d'envoi", "")))

    acc_ok = bool(row.get("Dossier accepté"))
    s2.write("**Accepté** : " + ("✅" if acc_ok else "❌"))
    s2.write("• Date : " + _safe_str(row.get("Date d'acceptation", "")))

    ref_ok = bool(row.get("Dossier refusé"))
    s3.write("**Refusé** : " + ("✅" if ref_ok else "❌"))
    s3.write("• Date : " + _safe_str(row.get("Date de refus", "")))

    ann_ok = bool(row.get("Dossier annulé"))
    s4.write("**Annulé** : " + ("✅" if ann_ok else "❌"))
    s4.write("• Date : " + _safe_str(row.get("Date d'annulation", "")))

    rfe_ok = bool(row.get("RFE"))
    s5.write("**RFE** : " + ("✅" if rfe_ok else "❌"))

with tabs[4]:
    st.subheader("👤 Compte client — suivi du dossier & paiements")

    if df_all.empty:
        st.info("Aucun client disponible.")
    else:
        sel_col1, sel_col2 = st.columns([2,2])
        names = sorted(df_all["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_all.columns else []
        ids   = sorted(df_all["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_all.columns else []

        chosen_name = sel_col1.selectbox("Nom", [""] + names, index=0, key=f"acc_nom_{SID}")
        chosen_id   = sel_col2.selectbox("ID_Client", [""] + ids, index=0, key=f"acc_id_{SID}")

        if chosen_id:
            row = df_all[df_all["ID_Client"].astype(str)==chosen_id].iloc[0]
        elif chosen_name:
            row = df_all[df_all["Nom"].astype(str)==chosen_name].iloc[0]
        else:
            row = None

        if row is None:
            st.stop()

        v_total = float(_safe_num_series(pd.DataFrame([row]), "Total (US $)").iloc[0])
        v_paye  = float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0])
        v_reste = float(_safe_num_series(pd.DataFrame([row]), "Reste").iloc[0])

        k1, k2, k3 = st.columns(3)
        k1.metric("Total (US $)", _fmt_money(v_total))
        k2.metric("Payé", _fmt_money(v_paye))
        k3.metric("Reste", _fmt_money(v_reste))

        st.markdown("### 🧾 Informations principales")
        i1, i2, i3 = st.columns(3)
        i1.write({"Dossier N": row.get("Dossier N",""), "ID_Client": row.get("ID_Client","")})
        i2.write({"Nom": row.get("Nom",""), "Categorie": row.get("Categorie","")})
        i3.write({"Sous-categorie": row.get("Sous-categorie",""), "Visa": row.get("Visa","")})

        _render_status_dates_block(row)

        # Paiements — historique + ajout
        st.markdown("### 💵 Paiements")

        hist = row.get("Paiements", [])
        if not isinstance(hist, list):
            try:
                hist = json.loads(_safe_str(hist) or "[]")
                if not isinstance(hist, list):
                    hist = []
            except Exception:
                hist = []
        if hist:
            df_hist = pd.DataFrame(hist)
            if "amount" in df_hist.columns:
                df_hist["amount"] = pd.to_numeric(df_hist["amount"], errors="coerce").fillna(0.0)
            if "date" in df_hist.columns:
                df_hist["date"] = pd.to_datetime(df_hist["date"], errors="coerce").dt.date.astype(str)
            st.dataframe(df_hist, use_container_width=True, height=200, key=f"hist_{SID}_{row.get('ID_Client','')}")
        else:
            st.caption("Aucun règlement enregistré.")

        st.markdown("#### ➕ Ajouter un paiement")
        ap1, ap2, ap3, ap4 = st.columns([1.2,1,1,1.2])
        pay_date = ap1.date_input("Date", value=date.today(), key=f"paydate_{SID}_{row.get('ID_Client','')}")
        pay_amt  = ap2.number_input("Montant (US $)", min_value=0.0, value=0.0, step=10.0, format="%.2f",
                                    key=f"payamt_{SID}_{row.get('ID_Client','')}")
        pay_meth = ap3.selectbox("Méthode", ["Cash","Chèque","CB","Virement","Venmo"],
                                 key=f"paymeth_{SID}_{row.get('ID_Client','')}")
        go_add   = ap4.button("Enregistrer le paiement", key=f"btn_add_pay_{SID}_{row.get('ID_Client','')}")

        if go_add:
            if pay_amt <= 0:
                st.warning("Le montant doit être supérieur à 0.")
            else:
                df_live = _read_clients(clients_path)
                mask = False
                if "ID_Client" in df_live.columns and row.get("ID_Client"):
                    mask = df_live["ID_Client"].astype(str) == str(row.get("ID_Client"))
                if (not mask.any()) and row.get("Nom"):
                    mask = (df_live["Nom"].astype(str) == str(row.get("Nom")))
                if not mask.any():
                    st.error("Ligne introuvable pour mise à jour.")
                else:
                    idx = df_live[mask].index[0]
                    current = df_live.at[idx, "Paiements"]
                    try:
                        cur_list = json.loads(_safe_str(current) or "[]") if not isinstance(current, list) else current
                        if not isinstance(cur_list, list):
                            cur_list = []
                    except Exception:
                        cur_list = []
                    cur_list.append({
                        "date": _safe_str(pay_date),
                        "amount": float(pay_amt),
                        "method": _safe_str(pay_meth),
                    })
                    df_live.at[idx, "Paiements"] = cur_list

                    def _sum_json(lst):
                        try:
                            return float(sum(float(it.get("amount",0.0) or 0.0) for it in (lst or [])))
                        except Exception:
                            return 0.0
                    sum_json = _sum_json(cur_list)
                    prev_paid = float(_safe_num_series(df_live.loc[[idx]], "Payé").iloc[0])
                    df_live.at[idx, "Payé"] = max(prev_paid, sum_json)

                    total_val = float(_safe_num_series(df_live.loc[[idx]], "Total (US $)").iloc[0])
                    df_live.at[idx, "Reste"] = max(0.0, total_val - float(df_live.at[idx, "Payé"]))

                    _write_clients(df_live, clients_path)
                    st.success("Paiement ajouté et soldes recalculés.")
                    st.cache_data.clear()
                    st.rerun()


# ==============================================
# 🧾 ONGLET : Visa (aperçu des structures pour contrôle)
# ==============================================
with tabs[3]:
    st.subheader("📄 Visa (aperçu)")
    st.caption("Contrôle rapide de la structure Visa lue depuis le fichier.")
    with st.expander("Voir le mapping brut (cat → sous-cat → options)"):
        st.json(visa_map)

    if visa_map:
        selc1, selc2 = st.columns(2)
        cats = sorted(list(visa_map.keys()))
        test_cat = selc1.selectbox("Categorie", [""] + cats, index=0, key=f"vz_cat_{SID}")
        if test_cat:
            subs = sorted(list(visa_map.get(test_cat, {}).keys()))
            test_sub = selc2.selectbox("Sous-categorie", [""] + subs, index=0, key=f"vz_sub_{SID}")
            if test_sub:
                st.markdown("**Options disponibles**")
                opt_list = visa_map.get(test_cat, {}).get(test_sub, [])
                if opt_list:
                    st.write(opt_list)
                else:
                    st.info("Aucune option cochée dans la feuille Visa pour cette sous-catégorie.")


# ==============================================
# 📈 ONGLET : Analyses (séries & détails)
# ==============================================
with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Aucune donnée client.")
    else:
        yearsA  = sorted([int(y) for y in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3, a4, a5 = st.columns(5)
        fy = a1.multiselect("Année", yearsA, default=[], key=f"a_years_{SID}")
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=f"a_months_{SID}")
        fc = a3.multiselect("Catégorie", catsA, default=[], key=f"a_cats_{SID}")
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=f"a_subs_{SID}")
        fv = a5.multiselect("Visa", visasA, default=[], key=f"a_visas_{SID}")

        dfA = df_all.copy()
        if fy: dfA = dfA[dfA["_Année_"].isin(fy)]
        if fm: dfA = dfA[dfA["Mois"].astype(str).isin(fm)]
        if fc: dfA = dfA[dfA["Categorie"].astype(str).isin(fc)]
        if fs: dfA = dfA[dfA["Sous-categorie"].astype(str).isin(fs)]
        if fv: dfA = dfA[dfA["Visa"].astype(str).isin(fv)]

        # KPI
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Dossiers", f"{len(dfA)}")
        k2.metric("Honoraires", _fmt_money(_safe_num_series(dfA, "Montant honoraires (US $)").sum()))
        k3.metric("Payé", _fmt_money(_safe_num_series(dfA, "Payé").sum()))
        k4.metric("Reste", _fmt_money(_safe_num_series(dfA, "Reste").sum()))

        # Graphique : dossiers par catégorie
        if not dfA.empty and "Categorie" in dfA.columns:
            st.markdown("### 📊 Dossiers par catégorie")
            vc = dfA["Categorie"].value_counts().reset_index()
            vc.columns = ["Categorie", "Nombre"]
            st.bar_chart(vc.set_index("Categorie"))

        # Graphique : honoraires par mois
        if not dfA.empty and "Mois" in dfA.columns:
            st.markdown("### 📈 Honoraires par mois")
            tmp = dfA.copy()
            tmp["Mois"] = tmp["Mois"].astype(str)
            gm = tmp.groupby("Mois", as_index=False)["Montant honoraires (US $)"].sum().sort_values("Mois")
            st.line_chart(gm.set_index("Mois"))

        # Tableau détaillé
        st.markdown("### 🧾 Détails des dossiers filtrés")
        det = dfA.copy()
        for c in ["Montant honoraires (US $)", "Autres frais (US $)", "Total (US $)", "Payé", "Reste"]:
            if c in det.columns:
                det[c] = _safe_num_series(det, c).apply(_fmt_money)
        if "Date" in det.columns:
            try:
                det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
            except Exception:
                det["Date"] = det["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste",
            "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
        ] if c in det.columns]

        sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
        det_sorted = det.sort_values(by=sort_keys) if sort_keys else det

        st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=f"a_detail_{SID}")


# ==============================================
# 🏦 ONGLET : Escrow — synthèse
# ==============================================
with tabs[2]:
    st.subheader("🏦 Escrow — synthèse")
    if df_all.empty:
        st.info("Aucun client.")
    else:
        dfE = df_all.copy()
        dfE["Payé"]  = _safe_num_series(dfE, "Payé")
        dfE["Reste"] = _safe_num_series(dfE, "Reste")
        dfE["Total (US $)"] = _safe_num_series(dfE, "Total (US $)")

        agg = dfE.groupby("Categorie", as_index=False)[["Total (US $)", "Payé", "Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg["Total (US $)"]).replace([pd.NA, pd.NaT], 0).fillna(0.0) * 100
        st.dataframe(agg, use_container_width=True, key=f"esc_agg_{SID}")

        t1, t2, t3 = st.columns(3)
        t1.metric("Total (US $)", _fmt_money(float(dfE["Total (US $)"].sum())))
        t2.metric("Payé", _fmt_money(float(dfE["Payé"].sum())))
        t3.metric("Reste", _fmt_money(float(dfE["Reste"].sum())))

        st.caption("NB : pour un escrow « strict », on peut isoler les honoraires perçus avant envoi, puis signaler les transferts à effectuer une fois « Dossier envoyé » coché.")


# ==============================================
# 🧾 GESTION (CRUD) — dans l’onglet Clients
# ==============================================
with tabs[4]:
    st.markdown("---")
    st.subheader("🧾 Gestion des clients (Ajouter / Modifier / Supprimer)")

    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=f"crud_op_{SID}")

    # --- Charger le dernier état pour édition
    df_live = _read_clients(clients_path)

    # ---------- AJOUT ----------
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")

        c1, c2, c3 = st.columns(3)
        nom = c1.text_input("Nom", "", key=f"add_nom_{SID}")
        dt  = c2.date_input("Date de création", value=date.today(), key=f"add_date_{SID}")
        mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                            index=int(date.today().month)-1, key=f"add_mois_{SID}")

        # Cascade Visa
        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        sel_cat = st.selectbox("Catégorie", [""] + cats, index=0, key=f"add_cat_{SID}")
        sel_sub = ""
        visa_final = ""
        opts_dict = {"exclusive": None, "options": []}
        info_msg = ""
        if sel_cat:
            subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
            sel_sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=f"add_sub_{SID}")
            if sel_sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=f"add_opts_{SID}", preselected={}
                )
        if info_msg:
            st.info(info_msg)

        f1, f2 = st.columns(2)
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f",
                                key=f"add_h_{SID}")
        other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f",
                                key=f"add_o_{SID}")

        st.markdown("#### 📌 Statuts initiaux")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent    = s1.checkbox("Dossier envoyé", key=f"add_sent_{SID}")
        sent_d  = s1.date_input("Date d'envoi", value=None, key=f"add_sentd_{SID}")
        acc     = s2.checkbox("Dossier accepté", key=f"add_acc_{SID}")
        acc_d   = s2.date_input("Date d'acceptation", value=None, key=f"add_accd_{SID}")
        ref     = s3.checkbox("Dossier refusé", key=f"add_ref_{SID}")
        ref_d   = s3.date_input("Date de refus", value=None, key=f"add_refd_{SID}")
        ann     = s4.checkbox("Dossier annulé", key=f"add_ann_{SID}")
        ann_d   = s4.date_input("Date d'annulation", value=None, key=f"add_annd_{SID}")
        rfe     = s5.checkbox("RFE", key=f"add_rfe_{SID}")

        if rfe and not any([sent, acc, ref, ann]):
            st.warning("⚠️ La case RFE ne peut être cochée qu’avec un autre statut (envoyé, accepté, refusé ou annulé).")

        save_add = st.button("💾 Enregistrer le client", key=f"btn_add_{SID}")
        if save_add:
            if not nom:
                st.warning("Veuillez saisir le nom.")
                st.stop()
            if not sel_cat or not sel_sub:
                st.warning("Veuillez choisir la catégorie et la sous-catégorie.")
            else:
                total = float(honor) + float(other)
                paye  = 0.0
                reste = max(0.0, total - paye)
                did = _make_client_id(nom, dt)
                dossier_n = _next_dossier(df_live, start=13057)

                new_row = {
                    "Dossier N": dossier_n,
                    "ID_Client": did,
                    "Nom": nom,
                    "Date": dt,
                    "Mois": f"{int(mois):02d}" if isinstance(mois, (int,str)) else _safe_str(mois),
                    "Categorie": sel_cat,
                    "Sous-categorie": sel_sub,
                    "Visa": visa_final if visa_final else sel_sub,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Total (US $)": total,
                    "Payé": paye,
                    "Reste": reste,
                    "Paiements": [],
                    "Options": opts_dict,
                    "Dossier envoyé": 1 if sent else 0,
                    "Date d'envoi": dt if sent and not sent_d else sent_d,
                    "Dossier accepté": 1 if acc else 0,
                    "Date d'acceptation": acc_d,
                    "Dossier refusé": 1 if ref else 0,
                    "Date de refus": ref_d,
                    "Dossier annulé": 1 if ann else 0,
                    "Date d'annulation": ann_d,
                    "RFE": 1 if rfe else 0,
                }
                df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                _write_clients(df_new, clients_path)
                st.success("Client ajouté.")
                st.cache_data.clear()
                st.rerun()

    # ---------- MODIFICATION ----------
    elif op == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if df_live.empty:
            st.info("Aucun client à modifier.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            sel1, sel2 = st.columns(2)
            target_name = sel1.selectbox("Nom", [""]+names, index=0, key=f"mod_nom_{SID}")
            target_id   = sel2.selectbox("ID_Client", [""]+ids, index=0, key=f"mod_id_{SID}")

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is None or not mask.any():
                st.stop()

            idx = df_live[mask].index[0]
            row = df_live.loc[idx].copy()

            c1, c2, c3 = st.columns(3)
            nom = c1.text_input("Nom", _safe_str(row.get("Nom","")), key=f"mod_nomv_{SID}")
            # valeur par défaut sûre pour la date
            dval = row.get("Date")
            if not isinstance(dval, (date, datetime)):
                try:
                    dval = pd.to_datetime(dval, errors="coerce")
                    dval = dval.date() if pd.notna(dval) else date.today()
                except Exception:
                    dval = date.today()
            dt  = c2.date_input("Date de création", value=dval, key=f"mod_date_{SID}")
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=(int(_safe_str(row.get("Mois","01"))) - 1), key=f"mod_mois_{SID}")

            # options déjà enregistrées
            preset_opts = row.get("Options", {})
            if not isinstance(preset_opts, dict):
                try:
                    preset_opts = json.loads(_safe_str(preset_opts) or "{}")
                    if not isinstance(preset_opts, dict):
                        preset_opts = {}
                except Exception:
                    preset_opts = {}

            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            preset_cat = _safe_str(row.get("Categorie",""))
            sel_cat = st.selectbox("Catégorie", [""] + cats,
                                   index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                                   key=f"mod_cat_{SID}")

            sel_sub = _safe_str(row.get("Sous-categorie",""))
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""] + subs,
                                       index=(subs.index(sel_sub)+1 if sel_sub in subs else 0),
                                       key=f"mod_sub_{SID}")

            visa_final, opts_dict, info_msg = "", {"exclusive": None, "options": []}, ""
            if sel_cat and sel_sub:
                visa_final, opts_dict, info_msg = build_visa_option_selector(
                    visa_map, sel_cat, sel_sub, keyprefix=f"mod_opts_{SID}", preselected=preset_opts
                )
            if info_msg:
                st.info(info_msg)

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), "Montant honoraires (US $)").iloc[0]),
                                    step=50.0, format="%.2f", key=f"mod_h_{SID}")
            other = f2.number_input("Autres frais (US $)", min_value=0.0,
                                    value=float(_safe_num_series(pd.DataFrame([row]), "Autres frais (US $)").iloc[0]),
                                    step=20.0, format="%.2f", key=f"mod_o_{SID}")

            st.markdown("#### 📌 Statuts")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent   = s1.checkbox("Dossier envoyé", value=bool(row.get("Dossier envoyé")), key=f"mod_sent_{SID}")
            # dates sûres
            def _date_or_none(val):
                if isinstance(val, (date, datetime)):
                    return val
                try:
                    d2 = pd.to_datetime(val, errors="coerce")
                    return d2.date() if pd.notna(d2) else None
                except Exception:
                    return None
            sent_d = s1.date_input("Date d'envoi", value=_date_or_none(row.get("Date d'envoi")), key=f"mod_sentd_{SID}")
            acc    = s2.checkbox("Dossier accepté", value=bool(row.get("Dossier accepté")), key=f"mod_acc_{SID}")
            acc_d  = s2.date_input("Date d'acceptation", value=_date_or_none(row.get("Date d'acceptation")), key=f"mod_accd_{SID}")
            ref    = s3.checkbox("Dossier refusé", value=bool(row.get("Dossier refusé")), key=f"mod_ref_{SID}")
            ref_d  = s3.date_input("Date de refus", value=_date_or_none(row.get("Date de refus")), key=f"mod_refd_{SID}")
            ann    = s4.checkbox("Dossier annulé", value=bool(row.get("Dossier annulé")), key=f"mod_ann_{SID}")
            ann_d  = s4.date_input("Date d'annulation", value=_date_or_none(row.get("Date d'annulation")), key=f"mod_annd_{SID}")
            rfe    = s5.checkbox("RFE", value=bool(row.get("RFE")), key=f"mod_rfe_{SID}")

            save_mod = st.button("💾 Enregistrer les modifications", key=f"btn_mod_{SID}")
            if save_mod:
                if not nom:
                    st.warning("Le nom est requis.")
                    st.stop()
                if not sel_cat or not sel_sub:
                    st.warning("Choisissez Catégorie et Sous-catégorie.")
                    st.stop()

                total = float(honor) + float(other)
                paye  = float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0])
                reste = max(0.0, total - paye)

                df_live.at[idx, "Nom"] = nom
                df_live.at[idx, "Date"] = dt
                df_live.at[idx, "Mois"] = f"{int(mois):02d}" if isinstance(mois,(int,str)) else _safe_str(mois)
                df_live.at[idx, "Categorie"] = sel_cat
                df_live.at[idx, "Sous-categorie"] = sel_sub
                df_live.at[idx, "Visa"] = (visa_final if visa_final else sel_sub)
                df_live.at[idx, "Montant honoraires (US $)"] = float(honor)
                df_live.at[idx, "Autres frais (US $)"] = float(other)
                df_live.at[idx, "Total (US $)"] = total
                df_live.at[idx, "Reste"] = reste
                df_live.at[idx, "Options"] = opts_dict
                df_live.at[idx, "Dossier envoyé"] = 1 if sent else 0
                df_live.at[idx, "Date d'envoi"] = dt if sent and not sent_d else sent_d
                df_live.at[idx, "Dossier accepté"] = 1 if acc else 0
                df_live.at[idx, "Date d'acceptation"] = acc_d
                df_live.at[idx, "Dossier refusé"] = 1 if ref else 0
                df_live.at[idx, "Date de refus"] = ref_d
                df_live.at[idx, "Dossier annulé"] = 1 if ann else 0
                df_live.at[idx, "Date d'annulation"] = ann_d
                df_live.at[idx, "RFE"] = 1 if rfe else 0

                _write_clients(df_live, clients_path)
                st.success("Modifications enregistrées.")
                st.cache_data.clear()
                st.rerun()

    # ---------- SUPPRESSION ----------
    elif op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client à supprimer.")
        else:
            names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
            ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
            sel1, sel2 = st.columns(2)
            target_name = sel1.selectbox("Nom", [""]+names, index=0, key=f"del_nom_{SID}")
            target_id   = sel2.selectbox("ID_Client", [""]+ids, index=0, key=f"del_id_{SID}")

            mask = None
            if target_id:
                mask = (df_live["ID_Client"].astype(str) == target_id)
            elif target_name:
                mask = (df_live["Nom"].astype(str) == target_name)

            if mask is not None and mask.any():
                row = df_live[mask].iloc[0]
                st.write({"Dossier N": row.get("Dossier N",""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                if st.button("❗ Confirmer la suppression", key=f"btn_del_{SID}"):
                    df_new = df_live[~mask].copy()
                    _write_clients(df_new, clients_path)
                    st.success("Client supprimé.")
                    st.cache_data.clear()
                    st.rerun()


# ==============================================
# 💾 Export global (Clients + Visa)
# ==============================================
st.markdown("---")
st.markdown("### 💾 Export global (Clients + Visa)")
colz1, colz2 = st.columns([1,3])
with colz1:
    if st.button("Préparer l’archive ZIP", key=f"zip_btn_{SID}"):
        try:
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients propre
                df_export = _read_clients(clients_path)
                with BytesIO() as xbuf:
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as wr:
                        df_export.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
                    zf.writestr("Clients.xlsx", xbuf.getvalue())
                # Visa (fichier tel quel si possible)
                try:
                    zf.write(visa_path, "Visa.xlsx")
                except Exception:
                    try:
                        dfv0 = pd.read_excel(visa_path, sheet_name=SHEET_VISA)
                        with BytesIO() as vb:
                            with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                                dfv0.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                            zf.writestr("Visa.xlsx", vb.getvalue())
                    except Exception:
                        pass
            st.session_state[f"zip_export_{SID}"] = buf.getvalue()
            st.success("Archive prête.")
        except Exception as e:
            st.error("Erreur de préparation : " + _safe_str(e))

with colz2:
    if st.session_state.get(f"zip_export_{SID}"):
        st.download_button(
            label="⬇️ Télécharger l’export (ZIP)",
            data=st.session_state[f"zip_export_{SID}"],
            file_name="Export_Visa_Manager.zip",
            mime="application/zip",
            key=f"zip_dl_{SID}",
        )