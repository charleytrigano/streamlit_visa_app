# =========================
# VISA MANAGER — APP COMPLETE
# =========================

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
import altair as alt

# ---- Page
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

# ---- Espace de noms unique pour éviter collisions de widgets
SID = st.session_state.setdefault("WIDGET_NS", str(uuid4()))

# =========================
# Constantes
# =========================
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

# =========================
# Utilitaires
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

# Crée les fichiers vides au besoin
ensure_file(CLIENTS_FILE_DEFAULT, SHEET_CLIENTS, CLIENTS_COLS)
ensure_file(VISA_FILE_DEFAULT, SHEET_VISA, ["Categorie","Sous-categorie 1","COS","EOS"])

def _norm(s: str) -> str:
    s2 = unicodedata.normalize("NFKD", s)
    s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
    s2 = s2.strip().lower().replace("\u00a0", " ")
    s2 = s2.replace("-", " ").replace("_", " ")
    return " ".join(s2.split())

# =========================
# Parsing de la feuille Visa
# =========================
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    """
    Retourne: {Categorie: {Sous-categorie: [options...]}}
    - Les options proviennent des colonnes dont la cellule = 1 (ou 'x' / 'oui'...)
    - Injection automatique "2-Etudiants" -> F-1/F-2 COS/EOS si aucune catégorie 'etudiant' détectée
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
                        opts.append(f"{sub} {cc}".strip())
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

    # Si aucune catégorie étudiants trouvée, on ajoute "2-Etudiants"
    if not found_students:
        out.setdefault("2-Etudiants", {})
        out["2-Etudiants"].setdefault("F-1", ["F-1 COS", "F-1 EOS"])
        out["2-Etudiants"].setdefault("F-2", ["F-2 COS", "F-2 EOS"])

    # Nettoyage
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
    # --- sauvegarde pour UNDO ---
    st.session_state.setdefault("undo_stack", [])
    try:
        prev = pd.read_excel(path, sheet_name=SHEET_CLIENTS)
    except Exception:
        prev = pd.DataFrame(columns=CLIENTS_COLS)
    st.session_state["undo_stack"].append(prev.copy())

    # --- écriture ---
    df2 = df.copy()
    df2["Options"] = df2["Options"].apply(lambda d: json.dumps(_normalize_options_json(d), ensure_ascii=False))
    df2["Paiements"] = df2["Paiements"].apply(lambda l: json.dumps(l, ensure_ascii=False))
    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        df2[c] = df2[c].apply(lambda b: 1 if bool(b) else 0)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as wr:
        _uniquify_columns(df2).to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)

def undo_last_write(path: str):
    """Annule la dernière écriture du fichier Clients, si possible."""
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

def _make_client_id(nom: str, d: date) -> str:
    base = _safe_str(nom).strip().replace(" ","_")
    return f"{base}-{d:%Y%m%d}"

# =========================
# UI — barre latérale
# =========================
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

# =========================
# Données
# =========================
visa_map = parse_visa_sheet(visa_path)
df_all   = _read_clients(clients_path)

# =========================
# Onglets
# =========================
tabs = st.tabs(["📊 Dashboard", "📈 Analyses", "🏦 Escrow", "📄 Visa (aperçu)", "👤 Clients"])

# =========================
# 📊 DASHBOARD
# =========================
with tabs[0]:
    st.subheader("📊 Dashboard — tous les clients")

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
            # Visa dynamique
            if cat and sub:
                visa_final, info_msg, opts = (lambda c=cat, s=sub: (
                    # rendu options: exclusif si "sub XXX", cases à cocher sinon
                    (lambda vmap=parse_visa_sheet(visa_path), kp=f"add_steps_{SID}":
                        (lambda arr=vmap.get(c, {}).get(s, []):
                            (lambda prefix=f"{s} ":
                                (lambda suffixes=[o[len(prefix):] for o in arr if o.startswith(prefix) and len(o)>len(prefix)],
                                         others=[o for o in arr if not (o.startswith(prefix) and len(o)>len(prefix))]):
                                    (lambda chosen_excl=st.radio(
                                        f"Option exclusive — {s}",
                                        options=[""]+sorted(set(suffixes)) if suffixes else [""],
                                        index=0, key=f"{kp}_excl_{SID}"
                                    ) if suffixes else None,
                                    chosen_multi=[lab for i,lab in enumerate(sorted(set(others)))
                                                  if st.checkbox(lab, value=False, key=f"{kp}_chk_{i}_{SID}")]):
                                        (f"{s} {chosen_excl}".strip() if chosen_excl else s,
                                         "" if arr else "Aucune option cochée pour cette sous-catégorie dans la feuille Visa.",
                                         {"exclusive": (chosen_excl or None),
                                          "options": chosen_multi})
                            )()
                        )()
                    )()
                ))()
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
    if action == "Modifier":
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
                # options pré-sélectionnées
                cur_opts = row.get("Options", {})
                visa_final, info_msg, opts = (lambda c=cat, s=sub, pre=cur_opts: (
                    (lambda vmap=parse_visa_sheet(visa_path), kp=f"mod_steps_{idx}_{SID}":
                        (lambda arr=vmap.get(c, {}).get(s, []):
                            (lambda prefix=f"{s} ":
                                (lambda suffixes=[o[len(prefix):] for o in arr if o.startswith(prefix) and len(o)>len(prefix)],
                                         others=[o for o in arr if not (o.startswith(prefix) and len(o)>len(prefix))]):
                                    (lambda sup_opts=[""]+sorted(set(suffixes)) if suffixes else [""],
                                            chosen_excl=st.radio(
                                                f"Option exclusive — {s}",
                                                options=sup_opts,
                                                index=(sup_opts.index(pre.get("exclusive")) if isinstance(pre,dict) and pre.get("exclusive") in sup_opts else 0),
                                                key=f"{kp}_excl_{SID}"
                                            ) if suffixes else None,
                                            chosen_multi=[lab for i,lab in enumerate(sorted(set(others)))
                                                          if st.checkbox(lab, value=(lab in (pre.get("options",[]) if isinstance(pre,dict) else [])),
                                                                         key=f"{kp}_chk_{i}_{SID}")]):
                                        (f"{s} {chosen_excl}".strip() if chosen_excl else s,
                                         "" if arr else "Aucune option cochée pour cette sous-catégorie dans la feuille Visa.",
                                         {"exclusive": (chosen_excl or None),
                                          "options": chosen_multi})
                            )()
                        )()
                    )()
                ))()
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

            # Paiements — ajout rapide
            st.markdown("##### 💳 Paiements — ajout rapide")
            p1,p2,p3,p4 = st.columns([1,1,1,2])
            with p1: pdt = st.date_input("Date paiement", value=date.today(), key=f"mod_paydt_{idx}_{SID}")
            with p2: pmd = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"mod_mode_{idx}_{SID}")
            with p3: pmt = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"mod_amt_{idx}_{SID}")
            with p4:
                note = st.text_input("Note (facultatif)", key=f"mod_note_{idx}_{SID}")
                if st.button("➕ Ajouter paiement", key=f"mod_addpay_{idx}_{SID}"):
                    base = _read_clients(clients_path)
                    plist = base.loc[idx,"Paiements"]
                    if not isinstance(plist, list):
                        try:
                            plist = json.loads(_safe_str(plist) or "[]")
                        except Exception:
                            plist = []
                    if float(pmt) > 0:
                        plist.append({"date": str(pdt), "mode": pmd, "amount": float(pmt), "note": _safe_str(note)})
                        base.loc[idx,"Paiements"] = plist
                        base = _normalize_clients(base)
                        _write_clients(base, clients_path)
                        st.success("Paiement ajouté."); st.rerun()
                    else:
                        st.warning("Montant > 0 requis.")

            # Historique détaillé (édition/suppression)
            st.markdown("##### 📜 Historique des paiements (édition & suppression)")
            base2 = _read_clients(clients_path)
            plist2 = base2.loc[idx, "Paiements"]
            if not isinstance(plist2, list):
                try:
                    plist2 = json.loads(_safe_str(plist2) or "[]")
                except Exception:
                    plist2 = []

            if plist2:
                for j, it in enumerate(plist2):
                    with st.expander(f"Paiement #{j+1} — {it.get('date','?')} • {it.get('mode','?')} • ${float(it.get('amount',0.0) or 0.0):,.2f}", expanded=False):
                        c1, c2, c3, c4, c5 = st.columns([1,1,1,2,1])
                        cur_dt = pd.to_datetime(it.get("date",""), errors="coerce")
                        cur_dt = (cur_dt.date() if pd.notna(cur_dt) else date.today())
                        cur_md = _safe_str(it.get("mode","CB")) or "CB"
                        try:
                            cur_amt = float(it.get("amount", 0.0) or 0.0)
                        except Exception:
                            cur_amt = 0.0
                        cur_note = _safe_str(it.get("note",""))

                        with c1:
                            edt_dt = st.date_input("Date", value=cur_dt, key=f"pay_dt_{idx}_{j}_{SID}")
                        with c2:
                            edt_md = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"],
                                                  index=["CB","Chèque","Cash","Virement","Venmo"].index(cur_md) if cur_md in ["CB","Chèque","Cash","Virement","Venmo"] else 0,
                                                  key=f"pay_md_{idx}_{j}_{SID}")
                        with c3:
                            edt_amt = st.number_input("Montant (US $)", min_value=0.0, value=cur_amt, step=10.0, format="%.2f",
                                                      key=f"pay_amt_{idx}_{j}_{SID}")
                        with c4:
                            edt_note = st.text_input("Note", value=cur_note, key=f"pay_note_{idx}_{j}_{SID}")
                        with c5:
                            upd = st.button("Mettre à jour", key=f"pay_upd_{idx}_{j}_{SID}")
                            rem = st.button("Supprimer",     key=f"pay_del_{idx}_{j}_{SID}")

                        if upd:
                            if float(edt_amt) <= 0:
                                st.warning("Le montant doit être > 0.")
                            else:
                                plist2[j] = {
                                    "date": str(edt_dt),
                                    "mode": edt_md,
                                    "amount": float(edt_amt),
                                    "note": edt_note,
                                }
                                base2.loc[idx, "Paiements"] = plist2
                                base2 = _normalize_clients(base2)
                                _write_clients(base2, clients_path)
                                st.success("✅ Paiement mis à jour.")
                                st.rerun()

                        if rem:
                            try:
                                del plist2[j]
                                base2.loc[idx, "Paiements"] = plist2
                                base2 = _normalize_clients(base2)
                                _write_clients(base2, clients_path)
                                st.success("✅ Paiement supprimé.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Suppression impossible : {e}")
            else:
                st.caption("Aucun paiement.")

            if st.button("💾 Sauvegarder", key=f"mod_save_{idx}_{SID}"):
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
    if action == "Supprimer":
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
# 📈 ANALYSES
# =========================
with tabs[1]:
    st.subheader("📈 Analyses des clients")
    if df_all.empty:
        st.info("Aucune donnée à analyser.")
    else:
        dfc = df_all.copy()
        dfc_grp = dfc.groupby("Categorie", as_index=False).agg({
            "Dossier N": "count",
            "Montant honoraires (US $)": "sum",
            "Payé": "sum",
            "Reste": "sum"
        }).rename(columns={"Dossier N": "Nombre"})

        st.markdown("#### Répartition des dossiers par catégorie")
        chart1 = (alt.Chart(dfc_grp).mark_bar()
                  .encode(x=alt.X("Categorie:N", sort="-y"),
                          y=alt.Y("Nombre:Q"),
                          tooltip=["Categorie","Nombre"]))
        st.altair_chart(chart1, use_container_width=True)

        dfm = dfc.copy()
        dfm["Mois"] = dfm["Mois"].astype(str)
        dfm_grp = dfm.groupby("Mois", as_index=False)["Montant honoraires (US $)"].sum()
        st.markdown("#### Montants d’honoraires par mois")
        chart2 = (alt.Chart(dfm_grp).mark_line(point=True)
                  .encode(x=alt.X("Mois:N", sort=[f"{m:02d}" for m in range(1,13)]),
                          y=alt.Y("Montant honoraires (US $):Q"),
                          tooltip=["Mois","Montant honoraires (US $)"]))
        st.altair_chart(chart2, use_container_width=True)

        dfp = pd.DataFrame({
            "Type": ["Payé", "Reste"],
            "Montant": [
                _safe_num_series(dfc,"Payé").sum(),
                _safe_num_series(dfc,"Reste").sum()
            ]
        })
        st.markdown("#### Répartition Payé / Reste")
        chart3 = (alt.Chart(dfp).mark_arc(innerRadius=60)
                  .encode(theta=alt.Theta(field="Montant", type="quantitative"),
                          color=alt.Color(field="Type", type="nominal"),
                          tooltip=["Type","Montant"]))
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

# =========================
# 👤 CLIENTS — Compte client (consultation & paiements)
# =========================
with tabs[4]:
    st.subheader("👤 Compte client — suivi du dossier & règlements")

    df_all = _read_clients(clients_path)
    if df_all.empty:
        st.info("Aucun client pour le moment.")
        st.stop()

    def _label(i: int) -> str:
        try:
            r = df_all.loc[i]
            return f"{r.get('Nom','?')} — {r.get('ID_Client','?')} (#{r.get('Dossier N','?')})"
        except Exception:
            return str(i)

    idx = st.selectbox("Choisis un client", options=list(df_all.index), format_func=_label, key=f"cli_sel_{SID}")
    if idx is None:
        st.stop()

    row = df_all.loc[idx]

    st.markdown("### 🗂️ Dossier")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Dossier N", _safe_str(row.get("Dossier N","")))
    c2.metric("ID Client", _safe_str(row.get("ID_Client","")))
    c3.metric("Nom", _safe_str(row.get("Nom","")))
    c4.metric("Date création", _safe_str(row.get("Date","")))

    cA, cB, cC = st.columns(3)
    cA.metric("Catégorie", _safe_str(row.get("Categorie","")))
    cB.metric("Sous-catégorie", _safe_str(row.get("Sous-categorie","")))
    cC.metric("Visa", _safe_str(row.get("Visa","")))

    st.markdown("### 📌 Statuts & dates")
    s1, s2, s3, s4, s5 = st.columns(5)
    s1.write(f"**Envoyé** : {'✅' if bool(row.get('Dossier envoyé')) else '❌'}")
    s1.write(f"• Date : {_safe_str(row.get(\"Date d'envoi\",\"\"))}")
    s2.write(f"**Accepté** : {'✅' if bool(row.get('Dossier accepté')) else '❌'}")
    s2.write(f"• Date : {_safe_str(row.get(\"Date d'acceptation\",\"\"))}")
    s3.write(f"**Refusé** : {'✅' if bool(row.get('Dossier refusé')) else '❌'}")
    s3.write(f"• Date : {_safe_str(row.get('Date de refus',''))}")
    s4.write(f"**Annulé** : {'✅' if bool(row.get('Dossier annulé')) else '❌'}")
    s4.write(f"• Date : {_safe_str(row.get(\"Date d'annulation\",\"\"))}")
    s5.write(f"**RFE** : {'✅' if bool(row.get('RFE')) else '❌'}")

    st.markdown("### 💵 Synthèse financière")
    hon = float(_safe_num_series(df_all.loc[[idx]], "Montant honoraires (US $)").iloc[0])
    aut = float(_safe_num_series(df_all.loc[[idx]], "Autres frais (US $)").iloc[0])
    tot = float(_safe_num_series(df_all.loc[[idx]], "Total (US $)").iloc[0])
    pay = float(_safe_num_series(df_all.loc[[idx]], "Payé").iloc[0])
    res = float(_safe_num_series(df_all.loc[[idx]], "Reste").iloc[0])

    f1, f2, f3, f4 = st.columns(4)
    f1.metric("Honoraires", _fmt_money(hon))
    f2.metric("Autres frais", _fmt_money(aut))
    f3.metric("Total", _fmt_money(tot))
    f4.metric("Reste à encaisser", _fmt_money(res))

    # Paiements — bloc édition/suppression + ajout + export
    st.markdown("### 📜 Historique des paiements (édition & suppression)")
    base = _read_clients(clients_path)
    plist = base.loc[idx, "Paiements"]
    if not isinstance(plist, list):
        try:
            plist = json.loads(_safe_str(plist) or "[]")
        except Exception:
            plist = []

    if plist:
        for j, it in enumerate(plist):
            with st.expander(f"Paiement #{j+1} — {it.get('date','?')} • {it.get('mode','?')} • ${float(it.get('amount',0.0) or 0.0):,.2f}", expanded=False):
                c1, c2, c3, c4, c5 = st.columns([1,1,1,2,1])
                cur_dt = pd.to_datetime(it.get("date",""), errors="coerce")
                cur_dt = (cur_dt.date() if pd.notna(cur_dt) else date.today())
                cur_md = _safe_str(it.get("mode","CB")) or "CB"
                try:
                    cur_amt = float(it.get("amount", 0.0) or 0.0)
                except Exception:
                    cur_amt = 0.0
                cur_note = _safe_str(it.get("note",""))

                with c1:
                    edt_dt = st.date_input("Date", value=cur_dt, key=f"cli_pay_dt_{idx}_{j}_{SID}")
                with c2:
                    edt_md = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"],
                                          index=["CB","Chèque","Cash","Virement","Venmo"].index(cur_md) if cur_md in ["CB","Chèque","Cash","Virement","Venmo"] else 0,
                                          key=f"cli_pay_md_{idx}_{j}_{SID}")
                with c3:
                    edt_amt = st.number_input("Montant (US $)", min_value=0.0, value=cur_amt, step=10.0, format="%.2f",
                                              key=f"cli_pay_amt_{idx}_{j}_{SID}")
                with c4:
                    edt_note = st.text_input("Note", value=cur_note, key=f"cli_pay_note_{idx}_{j}_{SID}")
                with c5:
                    upd = st.button("Mettre à jour", key=f"cli_pay_upd_{idx}_{j}_{SID}")
                    rem = st.button("Supprimer",     key=f"cli_pay_del_{idx}_{j}_{SID}")

                if upd:
                    if float(edt_amt) <= 0:
                        st.warning("Le montant doit être > 0.")
                    else:
                        plist[j] = {
                            "date": str(edt_dt),
                            "mode": edt_md,
                            "amount": float(edt_amt),
                            "note": edt_note,
                        }
                        base.loc[idx, "Paiements"] = plist
                        base = _normalize_clients(base)
                        _write_clients(base, clients_path)
                        st.success("✅ Paiement mis à jour.")
                        st.rerun()

                if rem:
                    try:
                        del plist[j]
                        base.loc[idx, "Paiements"] = plist
                        base = _normalize_clients(base)
                        _write_clients(base, clients_path)
                        st.success("✅ Paiement supprimé.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Suppression impossible : {e}")
    else:
        st.caption("Aucun paiement enregistré.")

    st.markdown("#### ➕ Ajouter un règlement")
    p1, p2, p3, p4 = st.columns([1,1,1,2])
    with p1:
        pay_dt = st.date_input("Date", value=date.today(), key=f"cli_add_dt_{idx}_{SID}")
    with p2:
        pay_md = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"cli_add_md_{idx}_{SID}")
    with p3:
        pay_amt = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"cli_add_amt_{idx}_{SID}")
    with p4:
        note = st.text_input("Note (facultatif)", key=f"cli_add_note_{idx}_{SID}")

    if st.button("💾 Enregistrer le paiement", key=f"cli_add_save_{idx}_{SID}"):
        if float(pay_amt) <= 0:
            st.warning("Le montant doit être > 0.")
        else:
            plist.append({"date": str(pay_dt), "mode": pay_md, "amount": float(pay_amt), "note": _safe_str(note)})
            base.loc[idx, "Paiements"] = plist
            base = _normalize_clients(base)
            _write_clients(base, clients_path)
            st.success("✅ Paiement enregistré.")
            st.rerun()

    st.markdown("#### ⬇️ Export du compte")
    base2 = _read_clients(clients_path)
    row2 = base2.loc[idx]
    plist2 = row2.get("Paiements", [])
    if not isinstance(plist2, list):
        try:
            plist2 = json.loads(_safe_str(plist2) or "[]")
        except Exception:
            plist2 = []

    hon2 = float(_safe_num_series(base2.loc[[idx]], "Montant honoraires (US $)").iloc[0])
    aut2 = float(_safe_num_series(base2.loc[[idx]], "Autres frais (US $)").iloc[0])

    export_rows = [{"Type": "Honoraires", "Date": row2.get("Date",""), "Montant": hon2}]
    if aut2:
        export_rows.append({"Type": "Autres frais", "Date": row2.get("Date",""), "Montant": aut2})
    for it in (plist2 or []):
        export_rows.append({"Type": f"Paiement {it.get('mode','')}",
                            "Date": it.get("date",""),
                            "Montant": -float(it.get("amount",0.0) or 0.0)})
    ledger = pd.DataFrame(export_rows)
    ledger_csv = ledger.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Télécharger le relevé (CSV)", data=ledger_csv,
                       file_name=f"Compte_{_safe_str(row2.get('ID_Client','client'))}.csv",
                       mime="text/csv")