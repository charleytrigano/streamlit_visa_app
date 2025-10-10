# ============================================
# VISA APP — PARTIE 1/5
# Imports, constantes, helpers, lecture fichiers, onglets
# ============================================

from __future__ import annotations

import json
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Tuple, List, Dict, Any

import numpy as np
import pandas as pd
import streamlit as st

# ---------- Noms de fichiers par défaut ----------
DEFAULT_CLIENTS_XLSX = "donnees_visa_clients1_adapte.xlsx"
DEFAULT_VISA_XLSX    = "donnees_visa_clients0.xlsx"

# ---------- Constantes colonnes ----------
HONO   = "Montant honoraires (US $)"
AUTRE  = "Autres frais (US $)"
TOTAL  = "Total (US $)"
PAY_JSON = "Paiements"
DOSSIER_COL = "Dossier N"

REF_LEVELS = ["Catégorie"] + [f"Sous-categories {i}" for i in range(1,9)]
DOSSIER_START = 13057

# ---------- Utils génériques ----------
def _fmt_money_us(x: float) -> str:
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"

def _safe_str(x) -> str:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return ""
        return str(x).strip()
    except Exception:
        return ""

def _norm_txt(x: str) -> str:
    s = _safe_str(x)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s*[/\-]\s*", " ", s)
    s = re.sub(r"[^a-zA-Z0-9\s]+", " ", s)
    return " ".join(s.lower().split())

def _to_num_series(s_like) -> pd.Series:
    """Convertit une colonne (Series/DataFrame/list) en Series float robuste."""
    if isinstance(s_like, pd.DataFrame):
        if s_like.shape[1] == 0:
            return pd.Series([], dtype=float)
        s = s_like.iloc[:, 0]
    else:
        s = pd.Series(s_like)
    s = s.astype(str).str.replace(r"[^\d,.\-]", "", regex=True)
    def _one(x: str) -> float:
        if x == "" or x == "-":
            return 0.0
        if x.count(",") == 1 and x.count(".") == 0:
            x = x.replace(",", ".")
        if x.count(".") == 1 and x.count(",") >= 1:
            x = x.replace(",", "")
        try:
            return float(x)
        except Exception:
            return 0.0
    return s.map(_one)

def _collapse_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Fusionne les colonnes dupliquées (somme si numérique, sinon première non vide)."""
    cols = df.columns.astype(str)
    if not cols.duplicated().any():
        return df
    out = pd.DataFrame(index=df.index)
    for col in pd.unique(cols):
        same = df.loc[:, cols == col]
        if same.shape[1] == 1:
            out[col] = same.iloc[:, 0]
            continue
        # Essai somme numérique
        try:
            same_num = same.apply(pd.to_numeric, errors="coerce")
            if same_num.notna().any().any():
                out[col] = same_num.sum(axis=1, skipna=True)
                continue
        except Exception:
            pass
        # Sinon, première non vide
        def _first_non_empty(row):
            for v in row:
                if pd.notna(v) and str(v).strip() != "":
                    return v
            return ""
        out[col] = same.apply(_first_non_empty, axis=1)
    return out

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    """Retourne une Series numérique pour 'col', même si dupliquée ou absente."""
    if col not in df.columns:
        return pd.Series([], dtype=float)
    s = df[col]
    if isinstance(s, pd.DataFrame):  # colonnes dupliquées
        s = s.iloc[:, 0]
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def _visa_code_only(v: str) -> str:
    s = _safe_str(v)
    if not s:
        return ""
    parts = s.split()
    if len(parts) >= 2 and parts[-1].upper() in {"COS", "EOS"}:
        return " ".join(parts[:-1]).strip()
    return s.strip()

def next_dossier_number(df: pd.DataFrame) -> int:
    if df is None or df.empty or DOSSIER_COL not in df.columns:
        return DOSSIER_START
    try:
        nums = pd.to_numeric(df[DOSSIER_COL], errors="coerce")
        m = int(nums.max()) if nums.notna().any() else DOSSIER_START - 1
    except Exception:
        m = DOSSIER_START - 1
    return max(m, DOSSIER_START - 1) + 1

def _make_client_id_from_row(row: dict) -> str:
    nom = _safe_str(row.get("Nom"))
    d = row.get("Date")
    try:
        d = pd.to_datetime(d).date() if pd.notna(d) else date.today()
    except Exception:
        d = date.today()
    base = f"{nom}-{d.strftime('%Y%m%d')}"
    base = re.sub(r"[^A-Za-z0-9\-]+", "", base.replace(" ", "-")).lower()
    return base

# ---------- Normalisation CLIENTS ----------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = _collapse_duplicate_columns(df.copy())

    # mapping souple
    ren = {}
    for c in df.columns:
        lc = _norm_txt(c)
        if "montant honoraires" in lc or lc == "honoraires":
            ren[c] = HONO
        elif "autres frais" in lc or lc == "autres":
            ren[c] = AUTRE
        elif lc.startswith("total"):
            ren[c] = TOTAL
        elif lc in {"reste", "solde"}:
            ren[c] = "Reste"
        elif "paye" in lc or "payé" in lc:
            ren[c] = "Payé"
        elif "categorie" in lc:
            ren[c] = "Catégorie"
        elif lc == "visa":
            ren[c] = "Visa"
        elif lc in {"dossier n", "dossier"}:
            ren[c] = DOSSIER_COL
        elif lc == "id_client":
            ren[c] = "ID_Client"
        elif lc == "nom":
            ren[c] = "Nom"
        elif lc == "date":
            ren[c] = "Date"
        elif lc == "mois":
            ren[c] = "Mois"
        elif lc == "paiements":
            ren[c] = PAY_JSON
        elif "escrow" in lc and "transf" in lc:
            ren[c] = "ESCROW transféré (US $)"
        elif "journal" in lc and "escrow" in lc:
            ren[c] = "Journal ESCROW"
        elif lc == "dossier envoye" or "envoy" in lc:
            ren[c] = "Dossier envoyé"
        elif "approuve" in lc:
            ren[c] = "Dossier approuvé"
        elif lc == "rfe":
            ren[c] = "RFE"
        elif "refuse" in lc:
            ren[c] = "Dossier refusé"
        elif "annule" in lc:
            ren[c] = "Dossier annulé"
        elif "date env" in lc:
            ren[c] = "Date envoyé"
        elif "date appr" in lc:
            ren[c] = "Date approuvé"
        elif "date rfe" in lc:
            ren[c] = "Date RFE"
        elif "date refus" in lc:
            ren[c] = "Date refusé"
        elif "date annul" in lc:
            ren[c] = "Date annulé"

    if ren:
        df = df.rename(columns=ren)

    # colonnes requises
    required = [
        DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois", "Catégorie", "Visa",
        HONO, AUTRE, TOTAL, "Payé", "Reste", PAY_JSON,
        "ESCROW transféré (US $)", "Journal ESCROW",
        "Dossier envoyé", "Date envoyé", "Dossier approuvé", "Date approuvé",
        "RFE", "Date RFE", "Dossier refusé", "Date refusé", "Dossier annulé", "Date annulé",
    ]
    for c in required:
        if c not in df.columns:
            if c in [HONO, AUTRE, TOTAL, "Payé", "Reste", "ESCROW transféré (US $)"]:
                df[c] = 0.0
            elif c in [PAY_JSON, "Journal ESCROW", "ID_Client", "Nom", "Catégorie", "Visa", "Mois", "Date"]:
                df[c] = ""
            elif c in ["Dossier envoyé", "Dossier approuvé", "RFE", "Dossier refusé", "Dossier annulé"]:
                df[c] = False
            else:
                df[c] = ""

    # Nettoyage Visa/Catégorie
    df["Visa"] = df["Visa"].astype(str).map(_visa_code_only)
    df["Catégorie"] = df["Catégorie"].replace("", pd.NA).fillna(df["Visa"]).astype(str)

    # Numériques
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste", "ESCROW transféré (US $)"]:
        df[c] = _to_num_series(df[c])

    # Dates + Mois
    def _to_date(x):
        try:
            if x == "" or pd.isna(x):
                return pd.NaT
            return pd.to_datetime(x).date()
        except Exception:
            return pd.NaT
    df["Date"] = df["Date"].map(_to_date)
    df["Mois"] = df["Date"].apply(lambda d: f"{d.month:02d}" if pd.notna(d) else pd.NA)

    # Totaux
    df[TOTAL] = df[HONO] + df[AUTRE]
    df["Payé"] = _to_num_series(df["Payé"])
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # N° de dossier
    if DOSSIER_COL in df.columns:
        nums = pd.to_numeric(df[DOSSIER_COL], errors="coerce")
        maxn = int(nums.max()) if nums.notna().any() else DOSSIER_START - 1
        for i in range(len(df)):
            if pd.isna(nums.iat[i]) or (isinstance(nums.iat[i], (int, float)) and int(nums.iat[i]) <= 0):
                maxn += 1
                df.at[i, DOSSIER_COL] = maxn
        try:
            df[DOSSIER_COL] = df[DOSSIER_COL].astype(int)
        except Exception:
            pass

    # ID client si manquant (Nom + date yyyymmdd + suffixe en cas de doublon)
    for i, r in df.iterrows():
        if not _safe_str(r.get("ID_Client", "")):
            base = _make_client_id_from_row(r.to_dict())
            cand = base
            j = 0
            while (df["ID_Client"].astype(str) == cand).any():
                j += 1
                cand = f"{base}-{j}"
            df.at[i, "ID_Client"] = cand

    # Champs dérivés (année/mois num)
    df["_Année_"] = df["Date"].apply(lambda x: x.year if pd.notna(x) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda x: int(x.month) if pd.notna(x) else pd.NA)

    # Tri
    try:
        df = df.sort_values(["Date", "Nom"], na_position="last").reset_index(drop=True)
    except Exception:
        pass
    return df

# ---------- Normalisation VISA (référentiel) ----------
def _ensure_visa_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoie et structure le fichier Visa (Catégorie + Sous-categories 1..8 + Actif)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=REF_LEVELS + ["Actif"])
    out = df.copy()

    # normalisation des noms
    norm = {re.sub(r"[^a-z0-9]+", "", str(c).lower()): str(c) for c in out.columns}
    def find_col(*cands):
        for cand in cands:
            key = re.sub(r"[^a-z0-9]+", "", cand.lower())
            if key in norm:
                return norm[key]
        for cand in cands:
            key = re.sub(r"[^a-z0-9]+", "", cand.lower())
            for k, orig in norm.items():
                if key in k:
                    return orig
        return None

    cat = find_col("Catégorie", "Categorie", "Category")
    out = out.rename(columns={cat: "Catégorie"}) if cat else out.assign(**{"Catégorie": ""})
    for i in range(1, 9):
        sc = find_col(f"Sous-categories {i}", f"Sous categorie {i}", f"SC{i}")
        if sc:
            out = out.rename(columns={sc: f"Sous-categories {i}"})
        else:
            out[f"Sous-categories {i}"] = ""
    act = find_col("Actif", "Active", "Inclure", "Afficher")
    out = out.rename(columns={act: "Actif"}) if act else out.assign(**{"Actif": 1})

    out = out.reindex(columns=REF_LEVELS + ["Actif"])
    for c in REF_LEVELS + ["Actif"]:
        out[c] = out[c].fillna("").astype(str).str.strip()
    out["Catégorie"] = out["Catégorie"].replace("", pd.NA).ffill().fillna("")
    try:
        out["Actif_num"] = pd.to_numeric(out["Actif"], errors="coerce").fillna(0).astype(int)
        out = out[out["Actif_num"] == 1].drop(columns=["Actif_num"])
    except Exception:
        pass
    mask = out[REF_LEVELS].apply(lambda r: "".join(r.values), axis=1) != ""
    return out[mask].reset_index(drop=True)

def _slug(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(s).lower()).strip("_")

def _multi_bool_inputs(options: List[str], label: str, keyprefix: str, as_toggle: bool=False) -> List[str]:
    """Affiche une sélection multi-case (checkbox ou toggle)"""
    options = [o for o in options if str(o).strip() != ""]
    if not options:
        st.caption(f"Aucune option pour **{label}**.")
        return []
    with st.expander(label, expanded=False):
        c1, c2 = st.columns(2)
        all_on  = c1.toggle("Tout sélectionner", value=False, key=f"{keyprefix}_all")
        none_on = c2.toggle("Tout désélectionner", value=False, key=f"{keyprefix}_none")
        selected = []
        cols = st.columns(3 if len(options) > 6 else 2)
        for i, opt in enumerate(sorted(options)):
            k = f"{keyprefix}_{i}"
            if all_on:
                st.session_state[k] = True
            if none_on:
                st.session_state[k] = False
            with cols[i % len(cols)]:
                val = st.toggle(opt, value=st.session_state.get(k, False), key=k) if as_toggle \
                      else st.checkbox(opt, value=st.session_state.get(k, False), key=k)
                if val:
                    selected.append(opt)
    return selected

def build_checkbox_filters_grouped(df_ref_in: pd.DataFrame, keyprefix: str, as_toggle: bool=False) -> dict:
    """Construit l’arborescence dynamique de filtres (Catégorie → SC1..SC8), renvoie une whitelist par Catégorie."""
    df_ref = _ensure_visa_columns(df_ref_in)
    res = {"Catégorie": [], "SC_map": {}, "__whitelist_visa__": []}
    if df_ref.empty:
        st.info("Référentiel Visa vide ou invalide.")
        return res

    cats = sorted([v for v in df_ref["Catégorie"].unique() if str(v).strip() != ""])
    sel_cats = _multi_bool_inputs(cats, "Catégories", f"{keyprefix}_cat", as_toggle=as_toggle)
    res["Catégorie"] = sel_cats

    whitelist = set()
    for cat in sel_cats:
        sub = df_ref[df_ref["Catégorie"] == cat].copy()
        res["SC_map"][cat] = {}
        st.markdown(f"#### 🧭 {cat}")
        for i in range(1, 9):
            col = f"Sous-categories {i}"
            options = sorted([v for v in sub[col].unique() if str(v).strip() != ""])
            picked = _multi_bool_inputs(options, f"{cat} — {col}", f"{keyprefix}_{_slug(cat)}_sc{i}", as_toggle=as_toggle)
            res["SC_map"][cat][col] = picked
            if picked:
                sub = sub[sub[col].isin(picked)]
        # Dans cette app, on filtre côté Clients par la Catégorie (code de base)
        whitelist.add(cat)

    res["__whitelist_visa__"] = sorted(whitelist)
    return res

def filter_clients_by_ref(df_clients: pd.DataFrame, sel: dict) -> pd.DataFrame:
    """Applique le filtre de sélection (whitelist Catégorie) au tableau des clients."""
    if df_clients is None or df_clients.empty:
        return df_clients
    f = df_clients.copy()
    wl = set(sel.get("__whitelist_visa__", []))
    if wl and "Catégorie" in f.columns:
        f = f[f["Catégorie"].astype(str).isin(wl)]
    return f

def _uniquify_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Renomme les colonnes dupliquées en ajoutant des suffixes (_2, _3, ...)."""
    cols = list(map(str, df.columns))
    seen = {}
    new_cols = []
    for c in cols:
        if c not in seen:
            seen[c] = 1
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
    df = df.copy()
    df.columns = new_cols
    return # --- 5) Tableau (montants formatés) ---
view = ff.copy()
for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
    if c in view.columns:
        view[c] = _safe_num_series(view, c).map(_fmt_money_us)
if "Date" in view.columns:
    view["Date"] = view["Date"].astype(str)

show_cols = [c for c in [
    DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", "Date", "Mois",
    HONO, AUTRE, TOTAL, "Payé", "Reste"
] if c in view.columns]

sort_keys = [c for c in ["_Année_", "_MoisNum_", "Catégorie", "Nom"] if c in view.columns]
view_sorted = view.sort_values(by=sort_keys) if sort_keys else view

# ✅ Sélection puis dédoublonnage (renommage des doublons)
df_disp = view_sorted[show_cols].copy()
df_disp = _uniquify_columns(df_disp)

st.dataframe(df_disp.reset_index(drop=True), use_container_width=True)


# ============================================
# VISA APP — PARTIE 3/5
# Clients : créer / modifier / supprimer / paiements multiples
# ============================================

with tab_clients:
    st.subheader("👥 Clients — créer / modifier / supprimer / paiements")

    live = df_clients.copy()

    # ---------- Sélection d’un client existant ----------
    cL, cR = st.columns([1, 1])

    with cL:
        st.markdown("### 🔎 Sélection")
        if live.empty:
            sel_idx = None
            sel_row = None
            st.caption("Aucun client.")
        else:
            labels = (live.get("Nom", "").astype(str) + " — " + live.get("ID_Client", "").astype(str)).fillna("")
            sel_idx = st.selectbox(
                "Client",
                options=list(live.index),
                format_func=lambda i: labels.iloc[i],
                key=f"cli_sel_{sheet_choice}",
            )
            sel_row = live.loc[sel_idx] if sel_idx is not None else None

    # ---------- Création d’un nouveau client ----------
    with cR:
        st.markdown("### ➕ Nouveau client")
        new_name = st.text_input("Nom", key=f"new_nom_{sheet_choice}")
        new_date = st.date_input("Date de création", value=date.today(), key=f"new_date_{sheet_choice}")

        # Choix du visa (code) — basé sur le référentiel si présent (Catégorie = code racine)
        if isinstance(df_visa, pd.DataFrame) and not df_visa.empty:
            codes = sorted(df_visa["Catégorie"].dropna().astype(str).unique().tolist())
        else:
            # fallback : utiliser les valeurs existantes
            codes = sorted(live.get("Catégorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
        new_code = st.selectbox("Catégorie / Code visa", options=[""] + codes, index=0, key=f"new_code_{sheet_choice}")

        new_hono = st.number_input(HONO, min_value=0.0, step=10.0, format="%.2f", key=f"new_hono_{sheet_choice}")
        new_autr = st.number_input(AUTRE, min_value=0.0, step=10.0, format="%.2f", key=f"new_autr_{sheet_choice}")

        if st.button("💾 Créer", key=f"btn_new_{sheet_choice}"):
            if not new_name:
                st.warning("Renseigne le **Nom**.")
            elif not new_code:
                st.warning("Choisis une **Catégorie / Code visa**.")
            else:
                base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)  # lecture brute
                base_norm = normalize_clients(base_raw.copy())

                dossier = next_dossier_number(base_norm)
                client_id = _make_client_id_from_row({"Nom": new_name, "Date": new_date})
                # éviter collision d’ID
                origin = client_id
                i = 0
                if "ID_Client" in base_norm.columns:
                    while (base_norm["ID_Client"].astype(str) == client_id).any():
                        i += 1
                        client_id = f"{origin}-{i}"

                total = float(new_hono) + float(new_autr)
                new_row = {
                    DOSSIER_COL: dossier,
                    "ID_Client": client_id,
                    "Nom": new_name,
                    "Date": pd.to_datetime(new_date).date(),
                    "Mois": f"{new_date.month:02d}",
                    "Catégorie": new_code,                 # Catégorie = code racine
                    "Visa": _visa_code_only(new_code),     # code de base (sans COS/EOS)
                    HONO: float(new_hono),
                    AUTRE: float(new_autr),
                    TOTAL: total,
                    "Payé": 0.0,
                    "Reste": total,
                    PAY_JSON: "[]",
                    "ESCROW transféré (US $)": 0.0,
                    "Journal ESCROW": "[]",
                    "Dossier envoyé": False,
                    "Date envoyé": "",
                    "Dossier approuvé": False,
                    "Date approuvé": "",
                    "RFE": False,
                    "Date RFE": "",
                    "Dossier refusé": False,
                    "Date refusé": "",
                    "Dossier annulé": False,
                    "Date annulé": "",
                }

                base_raw = pd.concat([base_raw, pd.DataFrame([new_row])], ignore_index=True)
                base_norm = normalize_clients(base_raw)
                write_sheet_inplace(clients_path, sheet_choice, base_norm)
                st.success("✅ Client créé.")
                st.rerun()

    st.markdown("---")

    # Si aucun client sélectionné on s’arrête ici
    if sel_row is None:
        st.stop()

    # ---------- Formulaire d’édition ----------
    idx = sel_idx
    ed = sel_row.to_dict()

    e1, e2, e3 = st.columns(3)
    with e1:
        ed_nom = st.text_input("Nom", value=_safe_str(ed.get("Nom", "")), key=f"ed_nom_{idx}_{sheet_choice}")
        ed_date = st.date_input(
            "Date de création",
            value=(pd.to_datetime(ed.get("Date")).date() if pd.notna(ed.get("Date")) else date.today()),
            key=f"ed_date_{idx}_{sheet_choice}",
        )
    with e2:
        # Catégorie / Code
        codes_all = sorted(df_visa["Catégorie"].dropna().astype(str).unique().tolist()) if isinstance(df_visa, pd.DataFrame) and not df_visa.empty \
                    else sorted(live.get("Catégorie", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
        current_code = _visa_code_only(ed.get("Catégorie", ed.get("Visa", "")))
        ed_code = st.selectbox(
            "Catégorie / Code visa",
            options=[""] + codes_all,
            index=(codes_all.index(current_code) + 1 if current_code in codes_all else 0),
            key=f"ed_code_{idx}_{sheet_choice}",
        )
    with e3:
        ed_hono = st.number_input(HONO, min_value=0.0, value=float(ed.get(HONO, 0.0)), step=10.0, format="%.2f", key=f"ed_hono_{idx}_{sheet_choice}")
        ed_autr = st.number_input(AUTRE, min_value=0.0, value=float(ed.get(AUTRE, 0.0)), step=10.0, format="%.2f", key=f"ed_autr_{idx}_{sheet_choice}")

    st.markdown("#### 🧾 Statuts du dossier")
    s1, s2, s3 = st.columns(3)
    with s1:
        ed_env = st.checkbox("Dossier envoyé", value=bool(ed.get("Dossier envoyé", False)), key=f"ed_env_{idx}_{sheet_choice}")
        ed_app = st.checkbox("Dossier approuvé", value=bool(ed.get("Dossier approuvé", False)), key=f"ed_app_{idx}_{sheet_choice}")
    with s2:
        ed_rfe = st.checkbox("RFE", value=bool(ed.get("RFE", False)), key=f"ed_rfe_{idx}_{sheet_choice}")
        ed_ref = st.checkbox("Dossier refusé", value=bool(ed.get("Dossier refusé", False)), key=f"ed_ref_{idx}_{sheet_choice}")
    with s3:
        ed_ann = st.checkbox("Dossier annulé", value=bool(ed.get("Dossier annulé", False)), key=f"ed_ann_{idx}_{sheet_choice}")

    st.caption("💡 Rappel : RFE ne peut être activé que si l’un des statuts **Envoyé / Approuvé / Refusé / Annulé** est vrai.")

    # ---------- Paiements (acomptes multiples) ----------
    st.markdown("### 💳 Paiements (acomptes)")

    p1, p2, p3, p4 = st.columns([1, 1, 1, 2])
    with p1:
        p_date = st.date_input("Date paiement", value=date.today(), key=f"p_date_{idx}_{sheet_choice}")
    with p2:
        p_mode = st.selectbox("Mode", ["CB", "Chèque", "Cash", "Virement", "Venmo"], key=f"p_mode_{idx}_{sheet_choice}")
    with p3:
        p_amt = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"p_amt_{idx}_{sheet_choice}")
    with p4:
        if st.button("➕ Ajouter paiement", key=f"btn_addpay_{idx}_{sheet_choice}"):
            base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
            # identifier la ligne via ID_Client si possible
            idc = _safe_str(ed.get("ID_Client", ""))
            if idc and "ID_Client" in base_raw.columns:
                idxs = base_raw.index[base_raw["ID_Client"].astype(str) == idc].tolist()
                real_idx = idxs[0] if idxs else idx
            else:
                real_idx = idx

            if float(p_amt) <= 0:
                st.warning("Le montant doit être > 0.")
            else:
                row = base_raw.loc[real_idx].to_dict()
                try:
                    plist = json.loads(_safe_str(row.get(PAY_JSON, "[]")) or "[]")
                    if not isinstance(plist, list):
                        plist = []
                except Exception:
                    plist = []
                plist.append({"date": str(p_date), "mode": p_mode, "amount": float(p_amt)})
                row[PAY_JSON] = json.dumps(plist, ensure_ascii=False)

                base_raw.loc[real_idx] = row
                base_norm = normalize_clients(base_raw.copy())
                write_sheet_inplace(clients_path, sheet_choice, base_norm)
                st.success("Paiement ajouté.")
                st.rerun()

    # Historique paiements
    try:
        hist = json.loads(_safe_str(sel_row.get(PAY_JSON, "[]")) or "[]")
        if not isinstance(hist, list):
            hist = []
    except Exception:
        hist = []
    st.write("**Historique des paiements**")
    if hist:
        h = pd.DataFrame(hist)
        if "amount" in h.columns:
            h["amount"] = h["amount"].astype(float).map(_fmt_money_us)
        st.dataframe(h, use_container_width=True)
    else:
        st.caption("Aucun paiement saisi.")

    st.markdown("---")

    # ---------- Boutons : Sauvegarder / Supprimer ----------
    a1, a2 = st.columns([1, 1])

    if a1.button("💾 Sauvegarder les modifications", key=f"btn_save_{idx}_{sheet_choice}"):
        base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
        # retrouver la ligne réelle via ID si possible
        idc = _safe_str(ed.get("ID_Client", ""))
        if idc and "ID_Client" in base_raw.columns:
            idx_real_list = base_raw.index[base_raw["ID_Client"].astype(str) == idc].tolist()
            real_idx = idx_real_list[0] if idx_real_list else None
        else:
            real_idx = idx

        if real_idx is None or not (0 <= real_idx < len(base_raw)):
            st.error("Ligne introuvable.")
        else:
            row = base_raw.loc[real_idx].to_dict()

            # maj champs
            row["Nom"] = ed_nom
            row["Date"] = pd.to_datetime(ed_date).date()
            row["Mois"] = f"{ed_date.month:02d}"
            if ed_code:
                row["Catégorie"] = ed_code
                row["Visa"] = _visa_code_only(ed_code)
            row[HONO] = float(ed_hono)
            row[AUTRE] = float(ed_autr)
            row[TOTAL] = float(ed_hono) + float(ed_autr)

            # statuts
            row["Dossier envoyé"] = bool(ed_env)
            row["Dossier approuvé"] = bool(ed_app)
            row["RFE"] = bool(ed_rfe)
            row["Dossier refusé"] = bool(ed_ref)
            row["Dossier annulé"] = bool(ed_ann)

            base_raw.loc[real_idx] = row
            base_norm = normalize_clients(base_raw.copy())
            write_sheet_inplace(clients_path, sheet_choice, base_norm)
            st.success("✅ Modifications sauvegardées.")
            st.rerun()

    if a2.button("🗑️ Supprimer ce client", key=f"btn_del_{idx}_{sheet_choice}"):
        base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
        # supprimer via ID si possible
        idc = _safe_str(ed.get("ID_Client", ""))
        if idc and "ID_Client" in base_raw.columns:
            mask = base_raw["ID_Client"].astype(str) == idc
            base_raw = base_raw.loc[~mask].reset_index(drop=True)
        else:
            if 0 <= idx < len(base_raw):
                base_raw = base_raw.drop(index=idx).reset_index(drop=True)
            else:
                st.error("Ligne introuvable."); st.stop()

        base_norm = normalize_clients(base_raw.copy())
        write_sheet_inplace(clients_path, sheet_choice, base_norm)
        st.success("🗑️ Client supprimé.")
        st.rerun()


# ============================================
# VISA APP — PARTIE 4/5
# Analyses : filtres + KPI + comparaisons + détails
# ============================================

with tab_analyses:
    st.subheader("📈 Analyses — Volumes & Financier")

    # --- 1) Filtres VISA hiérarchiques (référentiel) ---
    df_visa_safe = _ensure_visa_columns(df_visa)
    if df_visa_safe.empty:
        st.warning("⚠️ Référentiel Visa vide ou mal formé. Les filtres de catégories sont désactivés.")
        sel = {"__whitelist_visa__": [], "Catégorie": []}
        base = df_clients.copy()
    else:
        sel = build_checkbox_filters_grouped(
            df_visa_safe,
            keyprefix=f"flt_ana_{sheet_choice}",
            as_toggle=False,   # passe à True pour des toggles
        )
        base = filter_clients_by_ref(df_clients, sel)

    # --- 2) Filtres additionnels (Année / Mois / Solde / Recherche) ---
    base = base.copy()

    # sécurise les colonnes dérivées si besoin
    if "_Année_" not in base.columns:
        base["_Année_"] = base["Date"].apply(lambda x: x.year if pd.notna(x) else pd.NA)
    if "_MoisNum_" not in base.columns:
        base["_MoisNum_"] = base["Date"].apply(lambda x: int(x.month) if pd.notna(x) else pd.NA)
    base["_Mois_"] = base["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if pd.notna(m) else pd.NA)

    yearsA  = sorted([int(y) for y in base["_Année_"].dropna().unique()]) if not base.empty else []
    monthsA = [f"{m:02d}" for m in sorted([int(m) for m in base["_MoisNum_"].dropna().unique()])] if not base.empty else []

    cR1, cR2, cR3, cR4 = st.columns([1, 1, 1, 2])
    with cR1:
        sel_years  = st.multiselect("Année", yearsA, default=[], key=f"ana_year_{sheet_choice}")
    with cR2:
        sel_months = st.multiselect("Mois (MM)", monthsA, default=[], key=f"ana_month_{sheet_choice}")
    with cR3:
        solde_mode = st.selectbox(
            "Solde",
            ["Tous", "Soldé (Reste = 0)", "Non soldé (Reste > 0)"],
            index=0,
            key=f"ana_solde_{sheet_choice}"
        )
    with cR4:
        q = st.text_input("Recherche (nom, ID, visa…)", "", key=f"ana_q_{sheet_choice}")

    ff = base.copy()
    if sel_years:
        ff = ff[ff["_Année_"].isin(sel_years)]
    if sel_months:
        ff = ff[ff["_Mois_"].astype(str).isin(sel_months)]
    if solde_mode == "Soldé (Reste = 0)":
        ff = ff[_safe_num_series(ff, "Reste") <= 1e-9]
    elif solde_mode == "Non soldé (Reste > 0)":
        ff = ff[_safe_num_series(ff, "Reste") > 1e-9]
    if q:
        qn = q.lower().strip()
        def _row_match(r):
            hay = " ".join([
                _safe_str(r.get("Nom","")),
                _safe_str(r.get("ID_Client","")),
                _safe_str(r.get("Catégorie","")),
                _safe_str(r.get("Visa","")),
                str(r.get(DOSSIER_COL,"")),
            ]).lower()
            return qn in hay
        ff = ff[ff.apply(_row_match, axis=1)]

    # --- 3) KPI globaux ---
    st.markdown("""
    <style>
      .small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}
      .small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(ff)}")
    k2.metric("Honoraires", _fmt_money_us(_safe_num_series(ff, HONO).sum()))
    k3.metric("Encaissements (Payé)", _fmt_money_us(_safe_num_series(ff, "Payé").sum()))
    k4.metric("Solde à encaisser", _fmt_money_us(_safe_num_series(ff, "Reste").sum()))
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # --- 4) Comparaison Année → Année ---
    st.markdown("### 📆 Comparaison Année → Année")
    if not ff.empty and ff["_Année_"].notna().any():
        # agrégations robustes
        def _sum_col(df_loc, col):
            return _safe_num_series(df_loc, col).sum()

        grpY = ff.groupby("_Année_", dropna=True).apply(
            lambda g: pd.Series({
                "Dossiers": int(g.shape[0]),
                "Honoraires": _sum_col(g, HONO),
                "Paye": _sum_col(g, "Payé"),
                "Reste": _sum_col(g, "Reste"),
            })
        ).reset_index().rename(columns={"_Année_":"Année"}).sort_values("Année")

        st.dataframe(grpY, use_container_width=True)

        # graphiques (optionnels)
        try:
            import altair as alt
            ch1 = alt.Chart(grpY).mark_bar().encode(
                x=alt.X("Année:O", sort=None),
                y=alt.Y("Dossiers:Q")
            ).properties(height=220)
            st.altair_chart(ch1, use_container_width=True)
        except Exception:
            pass

        try:
            import altair as alt
            g_long = grpY.melt(id_vars=["Année"], value_vars=["Honoraires","Paye","Reste"],
                               var_name="Type", value_name="Montant")
            ch2 = alt.Chart(g_long).mark_line(point=True).encode(
                x=alt.X("Année:O", sort=None),
                y=alt.Y("Montant:Q"),
                color="Type:N"
            ).properties(height=240)
            st.altair_chart(ch2, use_container_width=True)
        except Exception:
            pass
    else:
        st.info("Aucune date exploitable pour la comparaison annuelle.")

    st.markdown("---")

    # --- 5) Par Mois (toutes années confondues) ---
    st.markdown("### 🗓️ Par mois (toutes années)")
    if not ff.empty and ff["_Mois_"].notna().any():
        def _sum_col(df_loc, col):
            return _safe_num_series(df_loc, col).sum()

        grpM = ff.groupby("_Mois_", dropna=True).apply(
            lambda g: pd.Series({
                "Dossiers": int(g.shape[0]),
                "Honoraires": _sum_col(g, HONO),
                "Paye": _sum_col(g, "Payé"),
                "Reste": _sum_col(g, "Reste"),
            })
        ).reset_index().rename(columns={"_Mois_":"Mois"}).sort_values("Mois")

        st.dataframe(grpM, use_container_width=True)

        try:
            import altair as alt
            ch3 = alt.Chart(grpM).mark_bar().encode(
                x=alt.X("Mois:O", sort=None),
                y=alt.Y("Dossiers:Q")
            ).properties(height=200)
            st.altair_chart(ch3, use_container_width=True)
        except Exception:
            pass
    else:
        st.info("Aucun mois exploitable.")

    st.markdown("---")

    (liste clients) ---
st.markdown("### 📋 Détails des dossiers filtrés")

detail = ff.copy()
for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
    if c in detail.columns:
        detail[c] = _safe_num_series(detail, c).map(_fmt_money_us)
if "Date" in detail.columns:
    detail["Date"] = detail["Date"].astype(str)

show_cols = [c for c in [
    DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", "Date", "Mois",
    HONO, AUTRE, TOTAL, "Payé", "Reste",
    "Dossier envoyé", "Dossier approuvé", "RFE", "Dossier refusé", "Dossier annulé"
] if c in detail.columns]

# Tri avant sélection
sort_keys = [c for c in ["_Année_", "_MoisNum_", "Catégorie", "Nom"] if c in detail.columns]
detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail

# ✅ Sélection + dédoublonnage des colonnes AVANT affichage
df_disp = detail_sorted[show_cols].copy()
df_disp = df_disp.loc[:, ~pd.Index(df_disp.columns).duplicated(keep="first")]

st.dataframe(df_disp.reset_index(drop=True), use_container_width=True)

# Récap filtres actifs
st.caption(
    "🧾 Filtres actifs — "
    f"Catégories={sel.get('Catégorie', [])} | "
    f"Années={sel_years} | Mois={sel_months} | "
    f"Solde={solde_mode} | Recherche='{q}'"
)

# ============================================
# VISA APP — PARTIE 5/5
# ESCROW : calculs, transferts, journal & alertes
# ============================================

with tab_escrow:
    st.subheader("🏦 ESCROW — dépôts, transferts & alertes")

    # ------- helpers locaux -------
    def _sum_payments_from_json(js: str) -> float:
        try:
            lst = json.loads(_safe_str(js) or "[]")
            if not isinstance(lst, list):
                return 0.0
            s = 0.0
            for it in lst:
                try:
                    s += float(it.get("amount", 0.0))
                except Exception:
                    pass
            return max(0.0, s)
        except Exception:
            return 0.0

    def _escrow_row_metrics(r: pd.Series) -> dict:
        """Calcule les éléments financiers ESCROW pour une ligne client."""
        hono = float(r.get(HONO, 0.0) or 0.0)
        pay_decl = float(r.get("Payé", 0.0) or 0.0)
        pay_js = _sum_payments_from_json(r.get(PAY_JSON, "[]"))
        pay = max(pay_decl, pay_js)  # tolérant: prend le plus grand des 2
        transf = float(r.get("ESCROW transféré (US $)", 0.0) or 0.0)
        dispo = max(min(pay, hono) - transf, 0.0)
        return {
            "honoraires": hono,
            "paye": pay,
            "transfere": transf,
            "dispo": dispo,
        }

    # Vue normalisée
    live = df_clients.copy()

    # ------- filtres rapides -------
    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    with c1:
        only_with_dispo = st.toggle("Uniquement ESCROW disponible", value=True, key="esc_onlydispo")
    with c2:
        only_sent = st.toggle("Uniquement dossiers envoyés", value=False, key="esc_onlysent")
    with c3:
        order_by_dispo = st.toggle("Trier par ESCROW disponible", value=True, key="esc_sortdispo")
    with c4:
        q = st.text_input("Recherche (Nom / ID / Dossier N / Visa)", "", key="esc_q")

    # ------- calculs par ligne -------
    rows = []
    for i, r in live.iterrows():
        m = _escrow_row_metrics(r)
        rows.append({
            "idx": i,
            DOSSIER_COL: r.get(DOSSIER_COL, ""),
            "ID_Client": r.get("ID_Client", ""),
            "Nom": r.get("Nom", ""),
            "Catégorie": r.get("Catégorie", ""),
            "Visa": r.get("Visa", ""),
            "Dossier envoyé": bool(r.get("Dossier envoyé", False)),
            HONO: m["honoraires"],
            "Payé_calc": m["paye"],
            "ESCROW transféré (US $)": m["transfere"],
            "ESCROW dispo": m["dispo"],
            "Journal ESCROW": _safe_str(r.get("Journal ESCROW", "[]")),
        })

    jdf = pd.DataFrame(rows)
    if only_with_dispo:
        jdf = jdf[jdf["ESCROW dispo"] > 0.0]
    if only_sent and "Dossier envoyé" in jdf.columns:
        jdf = jdf[jdf["Dossier envoyé"] == True]
    if q:
        qn = q.lower().strip()
        def _m(row):
            hay = " ".join([
                _safe_str(row.get("Nom","")),
                _safe_str(row.get("ID_Client","")),
                str(row.get(DOSSIER_COL,"")),
                _safe_str(row.get("Visa","")),
                _safe_str(row.get("Catégorie","")),
            ]).lower()
            return qn in hay
        jdf = jdf[jdf.apply(_m, axis=1)]

    if order_by_dispo and not jdf.empty:
        jdf = jdf.sort_values("ESCROW dispo", ascending=False)

    # ------- KPI globaux -------
    st.markdown("""
    <style>
      .small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}
      .small-kpi [data-testid="stMetricLabel"]{font-size:.85rem;opacity:.8}
    </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers listés", f"{len(jdf)}")
    k2.metric("Honoraires (périmètre)", _fmt_money_us(float(jdf[HONO].sum()) if not jdf.empty else 0.0))
    k3.metric("ESCROW transféré (périmètre)", _fmt_money_us(float(jdf["ESCROW transféré (US $)"].sum()) if not jdf.empty else 0.0))
    k4.metric("ESCROW dispo (périmètre)", _fmt_money_us(float(jdf["ESCROW dispo"].sum()) if not jdf.empty else 0.0))
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # ------- tableau & actions -------
    if jdf.empty:
        st.info("Aucun dossier à afficher avec les filtres actuels.")
        st.stop()

    # affichage tableau lisible
    show = jdf.copy()
    show[HONO] = show[HONO].map(_fmt_money_us)
    show["Payé_calc"] = show["Payé_calc"].map(_fmt_money_us)
    show["ESCROW transféré (US $)"] = show["ESCROW transféré (US $)"].map(_fmt_money_us)
    show["ESCROW dispo"] = show["ESCROW dispo"].map(_fmt_money_us)
    st.dataframe(
        show[[DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", HONO, "Payé_calc", "ESCROW transféré (US $)", "ESCROW dispo", "Dossier envoyé"]]
        .reset_index(drop=True),
        use_container_width=True
    )

    st.markdown("### ↗️ Enregistrer un transfert ESCROW")
    st.caption("Rappel : l’ESCROW disponible = min(Payé, Honoraires) − déjà transféré. Seule la partie **honoraires** est transférée.")

    for _, r in jdf.iterrows():
        # Un sous-formulaire par dossier
        st.markdown(f"**{r['Nom']} — {r['ID_Client']} — Dossier {r[DOSSIER_COL]}**")
        cA, cB, cC, cD = st.columns([1, 1, 1, 2])
        dispo = float(r["ESCROW dispo"])
        with cA:
            st.write("ESCROW disponible")
            st.write(_fmt_money_us(dispo))
        with cB:
            t_date = st.date_input("Date transfert", value=date.today(), key=f"esc_dt_{r['ID_Client']}")
        with cC:
            amt = st.number_input(
                "Montant à transférer (US $)",
                min_value=0.0, value=float(dispo),
                step=10.0, format="%.2f",
                key=f"esc_amt_{r['ID_Client']}"
            )
        with cD:
            note = st.text_input("Note (optionnel)", value="", key=f"esc_note_{r['ID_Client']}")

        can_transfer = dispo > 0 and amt > 0 and amt <= dispo + 1e-9
        if st.button("💸 Enregistrer transfert", key=f"esc_btn_{r['ID_Client']}"):
            if not can_transfer:
                st.warning("Montant invalide (doit être > 0 et ≤ ESCROW disponible).")
                st.stop()

            # écriture dans le classeur
            base_raw = pd.read_excel(clients_path, sheet_name=sheet_choice)
            # retrouver la ligne par ID_Client si possible
            if "ID_Client" in base_raw.columns:
                idxs = base_raw.index[base_raw["ID_Client"].astype(str) == str(r["ID_Client"])].tolist()
                real_idx = idxs[0] if idxs else None
            else:
                real_idx = int(r["idx"])

            if real_idx is None or not (0 <= real_idx < len(base_raw)):
                st.error("Ligne introuvable pour ce client.")
                st.stop()

            row = base_raw.loc[real_idx].to_dict()
            # maj transféré
            curr_tr = float(row.get("ESCROW transféré (US $)", 0.0) or 0.0)
            row["ESCROW transféré (US $)"] = curr_tr + float(amt)

            # append journal
            try:
                jlist = json.loads(_safe_str(row.get("Journal ESCROW", "[]")) or "[]")
                if not isinstance(jlist, list):
                    jlist = []
            except Exception:
                jlist = []
            jlist.append({
                "ts": datetime.now().isoformat(timespec="seconds"),
                "date": str(t_date),
                "amount": float(amt),
                "note": note,
            })
            row["Journal ESCROW"] = json.dumps(jlist, ensure_ascii=False)

            base_raw.loc[real_idx] = row
            base_norm = normalize_clients(base_raw.copy())
            write_sheet_inplace(clients_path, sheet_name=sheet_choice, df=base_norm)
            st.success("✅ Transfert enregistré.")
            st.rerun()

    st.markdown("---")

    # ------- alertes : dossiers envoyés à réclamer -------
    st.markdown("### 🚨 Alertes — dossiers envoyés à réclamer")
    alert_df = jdf[(jdf["Dossier envoyé"] == True) & (jdf["ESCROW dispo"] > 0.0)] if "Dossier envoyé" in jdf.columns else pd.DataFrame()
    if alert_df.empty:
        st.success("Aucune alerte : tous les dossiers envoyés ont leur ESCROW transféré ✅.")
    else:
        alert_view = alert_df.copy()
        alert_view["ESCROW dispo"] = alert_view["ESCROW dispo"].map(_fmt_money_us)
        st.dataframe(
            alert_view[[DOSSIER_COL, "ID_Client", "Nom", "Catégorie", "Visa", "ESCROW dispo"]],
            use_container_width=True
        )

    st.markdown("---")

    # ------- Journal ESCROW global -------
    st.markdown("### 📚 Journal ESCROW — global")
    logs = []
    for _, r in jdf.iterrows():
        try:
            jlist = json.loads(_safe_str(r.get("Journal ESCROW", "[]")) or "[]")
            if not isinstance(jlist, list):
                continue
            for it in jlist:
                logs.append({
                    "Dossier N": r.get(DOSSIER_COL, ""),
                    "ID_Client": r.get("ID_Client", ""),
                    "Nom": r.get("Nom", ""),
                    "Date": it.get("date",""),
                    "Horodatage": it.get("ts",""),
                    "Montant": float(it.get("amount", 0.0) or 0.0),
                    "Note": it.get("note",""),
                })
        except Exception:
            continue

    if logs:
        j = pd.DataFrame(logs).sort_values(["Horodatage","Date"], na_position="last").reset_index(drop=True)
        j["Montant"] = j["Montant"].map(_fmt_money_us)
        st.dataframe(j, use_container_width=True)
    else:
        st.caption("Aucun mouvement dans le journal ESCROW.")