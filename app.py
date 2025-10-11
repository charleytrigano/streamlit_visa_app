from __future__ import annotations

import streamlit as st
import pandas as pd
import json
from datetime import date, datetime
from pathlib import Path
from io import BytesIO
import openpyxl
import zipfile
import unicodedata

# ---------- CONFIG ----------
st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

# ---------- CONSTANTES ----------
CLIENTS_FILE = "donnees_visa_clients1_adapte.xlsx"
VISA_FILE    = "donnees_visa_clients1.xlsx"
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

# Colonnes attendues côté Clients (ajout de Options)
CLIENTS_COLS = [
    "Dossier N","ID_Client","Nom","Date","Mois",
    "Categorie","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
    "Payé","Reste","Paiements","Options"
]

# Widget toggles préférés (affichés en touche basculante)
TOGGLE_COLUMNS = {
    "AOS","CP","USCIS","I-130","I-140","I-140 & AOS","I-829","I-407",
    "Work Permit","Re-entry Permit","Consultation","Analysis","Referral",
    "Derivatives","Travel Permit","USC","LPR","Perm"
}

# ---------- HELPERS ----------
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
            seen[c] = 1
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}_{seen[c]}")
    out = df.copy()
    out.columns = new_cols
    return out

def ensure_file(path: str, sheet_name: str, cols: list[str]) -> None:
    p = Path(path)
    if not p.exists():
        df = pd.DataFrame(columns=cols)
        with pd.ExcelWriter(p, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=sheet_name, index=False)

# Créer fichiers si absents
ensure_file(CLIENTS_FILE, SHEET_CLIENTS, CLIENTS_COLS)
ensure_file(VISA_FILE, SHEET_VISA, ["Categorie","Sous-categorie 1"])

# ---------- LECTURE VISA (cellule=1 => option active ; libellé = entête colonne) ----------
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    """
    Retourne un mapping:
    {
      "Categorie": {
          "Sous-categorie": ["Sous-categorie COS", "Sous-categorie EOS", ...]
      }
    }
    """
    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)): 
            return False
        if isinstance(v, (int, float)): 
            return float(v) == 1.0
        s = str(v).strip().lower()
        return s in {"1", "true", "vrai", "oui", "yes", "x"}

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

        def _norm_header(s: str) -> str:
            s2 = unicodedata.normalize("NFKD", s)
            s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
            s2 = s2.strip().lower().replace("\u00a0", " ")
            s2 = s2.replace("-", " ").replace("_", " ")
            s2 = " ".join(s2.split())
            return s2

        colmap = { _norm_header(c): c for c in dfv.columns }
        cat_col = next((colmap[k] for k in colmap if "categorie" in k), None)
        sub_col = next((colmap[k] for k in colmap if "sous" in k), None)
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

            options = []
            for cc in check_cols:
                if _is_checked(row.get(cc, None)):
                    options.append(f"{sub} {cc}".strip())

            if not options and sub:
                options = [sub]

            if options:
                out.setdefault(cat, {})
                out[cat].setdefault(sub, [])
                out[cat][sub].extend(options)

        if out:
            for cat, subs in out.items():
                for sub, opts in subs.items():
                    subs[sub] = sorted(set(opts))
            return out
    return {}

# ---------- RAW VISA DF + utilitaires pour widgets dynamiques ----------
@st.cache_data(show_spinner=False)
def load_raw_visa_df(xlsx_path: str | Path, sheet_name: str = SHEET_VISA) -> pd.DataFrame:
    try:
        df = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()
    df = _uniquify_columns(df)
    df.columns = df.columns.map(str).str.strip()
    return df

def _find_cat_sub_columns(df: pd.DataFrame) -> tuple[str | None, str | None]:
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
            sub_col = cmap[k]
            break
    return cat_col, sub_col

def _normalize_options_json(x) -> dict:
    try:
        d = json.loads(_safe_str(x) or "{}")
        if not isinstance(d, dict):
            return {}
        # sanitize
        excl = d.get("exclusive", None)
        opts = d.get("options", [])
        if not isinstance(opts, list):
            opts = []
        return {"exclusive": excl, "options": [str(o) for o in opts]}
    except Exception:
        return {"exclusive": None, "options": []}

def render_dynamic_steps(cat: str, sub: str, keyprefix: str, preselected: dict | None = None) -> tuple[str, str, dict]:
    """
    Affiche dynamiquement:
      - radio pour duo exclusif (COS/EOS ou USCIS/CP)
      - toggles pour TOGGLE_COLUMNS
      - checkboxes pour le reste
    preselected = {"exclusive": "...", "options": ["...", ...]}
    Retourne (visa_final, message_info, selected_dict)
    """
    if not (cat and sub):
        return "", "Choisir d'abord Catégorie et Sous-catégorie.", {"exclusive": None, "options": []}

    vdf = load_raw_visa_df(VISA_FILE, SHEET_VISA)
    if vdf.empty:
        return "", "Feuille Visa introuvable ou vide.", {"exclusive": None, "options": []}

    cat_col, sub_col = _find_cat_sub_columns(vdf)
    if not cat_col:
        return "", "Colonne 'Catégorie' introuvable dans la feuille Visa.", {"exclusive": None, "options": []}
    if not sub_col:
        row = vdf[vdf[cat_col].astype(str).str.strip() == cat]
    else:
        row = vdf[
            (vdf[cat_col].astype(str).str.strip() == cat) &
            (vdf[sub_col].astype(str).str.strip() == sub)
        ]
    if row.empty:
        return "", "Combinaison Catégorie/Sous-catégorie non trouvée.", {"exclusive": None, "options": []}

    row = row.iloc[0]
    option_cols = [c for c in vdf.columns if c not in {cat_col, sub_col}]

    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)): 
            return False
        if isinstance(v, (int, float)): 
            return float(v) == 1.0
        s = str(v).strip().lower()
        return s in {"1","x","true","vrai","oui","yes"}

    possibles = [c for c in option_cols if _is_checked(row.get(c))]
    exclusive = None
    if {"COS","EOS"}.issubset(set(possibles)):
        exclusive = ("COS","EOS")
    elif {"USCIS","CP"}.issubset(set(possibles)):
        exclusive = ("USCIS","CP")

    pre = _normalize_options_json(preselected or {})
    visa_final = ""
    info_msg = ""
    selected_opts: list[str] = []
    selected_excl: str | None = None

    # Exclusif
    if exclusive:
        st.caption("Choix exclusif")
        # valeur par défaut
        default_index = 0
        if pre["exclusive"] in exclusive:
            default_index = list(exclusive).index(pre["exclusive"])
        choice = st.radio(
            "Sélectionner une option",
            options=list(exclusive),
            index=default_index,
            horizontal=True,
            key=f"{keyprefix}_exclusive",
        )
        selected_excl = choice
        visa_final = f"{sub} {choice}".strip()

    # Autres
    others = [c for c in possibles if not (exclusive and c in exclusive)]
    if others:
        st.caption("Options complémentaires")
    for i, col in enumerate(others):
        label = col
        default_val = label in pre["options"]
        if col in TOGGLE_COLUMNS:
            val = st.toggle(label, value=default_val, key=f"{keyprefix}_tog_{i}")
            if val:
                selected_opts.append(label)
        else:
            val = st.checkbox(label, value=default_val, key=f"{keyprefix}_chk_{i}")
            if val:
                selected_opts.append(label)

    if not visa_final:
        # pas de duo exclusif -> imposer un seul choix dans others
        if len(selected_opts) == 0:
            info_msg = "Coche une option (une seule)."
        elif len(selected_opts) > 1:
            info_msg = "Une seule option possible."
        else:
            visa_final = f"{sub} {selected_opts[0]}".strip()

    if not visa_final and not possibles:
        visa_final = sub

    return visa_final, info_msg, {"exclusive": selected_excl, "options": selected_opts}


# ---------- DONNÉES CLIENTS ----------
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

    # Options JSON
    df["Options"] = df["Options"].apply(_normalize_options_json)

    df["_Année_"]   = df["Date"].apply(lambda d: d.year if pd.notna(d) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda d: d.month if pd.notna(d) else pd.NA)
    return _uniquify_columns(df)

def _read_clients() -> pd.DataFrame:
    df = pd.read_excel(CLIENTS_FILE, sheet_name=SHEET_CLIENTS)
    return _normalize_clients(df)

def _write_clients(df: pd.DataFrame) -> None:
    df = df.copy()
    # re-sérialiser Options + Paiements
    df["Options"] = df["Options"].apply(lambda d: json.dumps(_normalize_options_json(d), ensure_ascii=False))
    df["Paiements"] = df["Paiements"].apply(lambda l: json.dumps(l, ensure_ascii=False))
    with pd.ExcelWriter(CLIENTS_FILE, engine="openpyxl", mode="w") as wr:
        _uniquify_columns(df).to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if "Dossier N" in df.columns:
        s = pd.to_numeric(df["Dossier N"], errors="coerce")
        if s.notna().any():
            return int(s.max()) + 1
    return int(start)

def _make_client_id(nom: str, d: date) -> str:
    base = f"{_safe_str(nom).strip().replace(' ', '_')}-{d:%Y%m%d}"
    return base

visa_map = parse_visa_sheet(VISA_FILE)

tab_clients, tab_visa = st.tabs(["👥 Clients", "📄 Visa"])

# ---------- ONGLET CLIENTS ----------
with tab_clients:
    st.subheader("Créer, modifier, supprimer un client — et gérer ses paiements + options")
    df_clients = _read_clients()

    left, right = st.columns([1,1], gap="large")

    # Liste / sélection
    with left:
        st.markdown("### 🔎 Sélection d’un client existant")
        if df_clients.empty:
            st.info("Aucun client pour l’instant.")
            sel_idx, sel_row = None, None
        else:
            labels = (df_clients.get("Nom","").astype(str) + " — " + df_clients.get("ID_Client","").astype(str)).fillna("")
            indices = list(df_clients.index)
            sel_idx = st.selectbox(
                "Client",
                options=indices,
                format_func=lambda i: labels.iloc[i],
                key="cli_sel_idx"
            )
            sel_row = df_clients.loc[sel_idx] if sel_idx is not None else None

        st.markdown("### 📋 Tous les clients")
        view = df_clients.copy()
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in view.columns:
                view[c] = _safe_num_series(view, c).map(_fmt_money)
        if "Date" in view.columns:
            view["Date"] = view["Date"].astype(str)
        # Afficher aussi un condensé d'options
        view["Options (résumé)"] = view["Options"].apply(
            lambda d: f"[{(d or {}).get('exclusive')}] + {', '.join((d or {}).get('options', []))}" if isinstance(d, dict) else ""
        )

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Date","Mois","Categorie","Sous-categorie","Visa",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste","Options (résumé)"
        ] if c in view.columns]
        sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
        view = view.sort_values(by=sort_cols) if sort_cols else view
        st.dataframe(_uniquify_columns(view[show_cols].reset_index(drop=True)), use_container_width=True)

    # Création
    with right:
        st.markdown("### ➕ Nouveau client")
        new_nom  = st.text_input("Nom", key="new_nom")
        new_date = st.date_input("Date de création", value=date.today(), key="new_date")

        cats = sorted(list(visa_map.keys()))
        new_cat = st.selectbox("Catégorie", options=[""]+cats, index=0, key="new_cat")

        subs = sorted(list(visa_map.get(new_cat, {}).keys())) if new_cat else []
        new_sub = st.selectbox("Sous-catégorie", options=[""]+subs, index=0, key="new_sub")

        st.caption("Étapes disponibles pour cette sous-catégorie")
        new_visa, err_new, new_opts = render_dynamic_steps(new_cat, new_sub, keyprefix="new_form", preselected=None)
        if err_new:
            st.info(err_new)

        new_hono = st.number_input("Montant honoraires (US $)", min_value=0.0, step=10.0, format="%.2f", key="new_hono")
        new_autr = st.number_input("Autres frais (US $)",     min_value=0.0, step=10.0, format="%.2f", key="new_autre")

        if st.button("💾 Créer", key="btn_create_new"):
            if not new_nom:
                st.warning("Nom obligatoire."); st.stop()
            if not new_cat:
                st.warning("Catégorie obligatoire."); st.stop()
            if not new_sub:
                st.warning("Sous-catégorie obligatoire."); st.stop()
            if new_visa == "":
                new_visa = new_sub or ""

            base = _read_clients()
            dossier = _next_dossier(base)
            cid_base = _make_client_id(new_nom, new_date)
            cid = cid_base; i = 0
            while (base["ID_Client"].astype(str) == cid).any():
                i += 1; cid = f"{cid_base}-{i}"

            total = float(new_hono) + float(new_autr)
            row = {
                "Dossier N": dossier,
                "ID_Client": cid,
                "Nom": new_nom,
                "Date": pd.to_datetime(new_date).date(),
                "Mois": f"{new_date.month:02d}",
                "Categorie": new_cat,
                "Sous-categorie": new_sub,
                "Visa": new_visa,
                "Montant honoraires (US $)": float(new_hono),
                "Autres frais (US $)": float(new_autr),
                "Total (US $)": total,
                "Payé": 0.0,
                "Reste": total,
                "Paiements": json.dumps([], ensure_ascii=False),
                "Options": json.dumps(_normalize_options_json(new_opts), ensure_ascii=False),
            }
            base = pd.concat([base, pd.DataFrame([row])], ignore_index=True)
            base = _normalize_clients(base)
            _write_clients(base)
            st.success("✅ Client créé.")
            st.rerun()

    st.markdown("---")

    # Modification + paiements
    if sel_row is not None:
        idx = sel_idx
        ed = sel_row.to_dict()
        cur_options = _normalize_options_json(ed.get("Options", {}))

        c1, c2, c3 = st.columns(3)

        with c1:
            ed_nom  = st.text_input("Nom", value=_safe_str(ed.get("Nom","")), key=f"ed_nom_{idx}")
            ed_date = st.date_input(
                "Date de création",
                value=(pd.to_datetime(ed.get("Date")).date() if pd.notna(ed.get("Date")) else date.today()),
                key=f"ed_date_{idx}"
            )

        with c2:
            cats = sorted(list(visa_map.keys()))
            curr_cat = _safe_str(ed.get("Categorie",""))
            ed_cat = st.selectbox(
                "Catégorie",
                options=[""]+cats,
                index=(cats.index(curr_cat)+1 if curr_cat in cats else 0),
                key=f"ed_cat_{idx}"
            )

            subs = sorted(list(visa_map.get(ed_cat, {}).keys())) if ed_cat else []
            curr_sub = _safe_str(ed.get("Sous-categorie",""))
            ed_sub = st.selectbox(
                "Sous-catégorie",
                options=[""]+subs,
                index=(subs.index(curr_sub)+1 if curr_sub in subs else 0),
                key=f"ed_sub_{idx}"
            )

        with c3:
            st.caption("Étapes disponibles pour cette sous-catégorie")
            ed_visa_final, err_ed, ed_opts = render_dynamic_steps(
                ed_cat, ed_sub, keyprefix=f"ed_{idx}", preselected=cur_options
            )
            if err_ed:
                st.info(err_ed)

            ed_hono = st.number_input(
                "Montant honoraires (US $)", min_value=0.0,
                value=float(ed.get("Montant honoraires (US $)",0.0)), step=10.0, format="%.2f",
                key=f"ed_hono_{idx}"
            )
            ed_autr = st.number_input(
                "Autres frais (US $)", min_value=0.0,
                value=float(ed.get("Autres frais (US $)",0.0)), step=10.0, format="%.2f",
                key=f"ed_autre_{idx}"
            )

        st.markdown("### 💳 Paiements (acomptes)")
        p1, p2, p3, p4 = st.columns([1,1,1,2])
        with p1:
            pay_date = st.date_input("Date paiement", value=date.today(), key=f"pay_dt_{idx}")
        with p2:
            pay_mode = st.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=f"pay_mode_{idx}")
        with p3:
            pay_amt = st.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=f"pay_amt_{idx}")
        with p4:
            if st.button("➕ Ajouter paiement", key=f"btn_addpay_{idx}"):
                if float(pay_amt) <= 0:
                    st.warning("Montant > 0 requis.")
                else:
                    base = _read_clients()
                    idc = _safe_str(ed.get("ID_Client",""))
                    if idc and "ID_Client" in base.columns:
                        hit = base.index[base["ID_Client"].astype(str) == idc]
                        if len(hit) == 0:
                            st.error("Ligne introuvable."); st.stop()
                        ridx = int(hit[0])
                    else:
                        ridx = int(idx)

                    row = base.loc[ridx].to_dict()
                    try:
                        plist = json.loads(_safe_str(row.get("Paiements","[]")) or "[]")
                        if not isinstance(plist, list):
                            plist = []
                    except Exception:
                        plist = []
                    plist.append({"date": str(pay_date), "mode": pay_mode, "amount": float(pay_amt)})
                    row["Paiements"] = json.dumps(plist, ensure_ascii=False)
                    base.loc[ridx] = row
                    base = _normalize_clients(base)
                    _write_clients(base)
                    st.success("Paiement ajouté.")
                    st.rerun()

        # Historique paiements
        st.markdown("#### Historique des paiements")
        try:
            hist = json.loads(_safe_str(sel_row.get("Paiements","[]")) or "[]")
            if not isinstance(hist, list):
                hist = []
        except Exception:
            hist = []
        if hist:
            h = pd.DataFrame(hist)
            if "amount" in h.columns:
                h["amount"] = h["amount"].astype(float).map(_fmt_money)
            st.dataframe(h, use_container_width=True)
        else:
            st.caption("Aucun paiement saisi.")

        st.markdown("---")
        b1, b2 = st.columns(2)
        if b1.button("💾 Sauvegarder les modifications", key=f"btn_save_{idx}"):
            base = _read_clients()
            idc = _safe_str(ed.get("ID_Client",""))
            if idc and "ID_Client" in base.columns:
                hit = base.index[base["ID_Client"].astype(str) == idc]
                if len(hit) == 0:
                    st.error("Ligne introuvable."); st.stop()
                ridx = int(hit[0])
            else:
                ridx = int(idx)

            if ridx < 0 or ridx >= len(base):
                st.error("Ligne introuvable."); st.stop()

            row = base.loc[ridx].to_dict()
            row["Nom"] = ed_nom
            row["Date"] = pd.to_datetime(ed_date).date()
            row["Mois"] = f"{ed_date.month:02d}"
            row["Categorie"] = ed_cat
            row["Sous-categorie"] = ed_sub
            row["Visa"] = ed_visa_final if ed_visa_final else (ed_sub or "")
            row["Montant honoraires (US $)"] = float(ed_hono)
            row["Autres frais (US $)"] = float(ed_autr)
            row["Total (US $)"] = float(ed_hono) + float(ed_autr)
            row["Options"] = json.dumps(_normalize_options_json(ed_opts), ensure_ascii=False)

            base.loc[ridx] = row
            base = _normalize_clients(base)
            _write_clients(base)
            st.success("✅ Modifications enregistrées.")
            st.rerun()

        if b2.button("🗑️ Supprimer ce client", key=f"btn_del_{idx}"):
            base = _read_clients()
            idc = _safe_str(ed.get("ID_Client",""))
            if idc and "ID_Client" in base.columns:
                base = base.loc[base["ID_Client"].astype(str) != idc].reset_index(drop=True)
            else:
                base = base.drop(index=ridx).reset_index(drop=True)
            _write_clients(_normalize_clients(base))
            st.success("🗑️ Client supprimé.")
            st.rerun()

# ---------- ONGLET VISA (aperçu) ----------
with tab_visa:
    st.subheader("Référentiel Visa (cellules = 1 → options)")
    visa_map = parse_visa_sheet(VISA_FILE)
    if not visa_map:
        st.warning("Aucune donnée Visa trouvée. Vérifie le fichier et l’onglet.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            cat_pick = st.selectbox("Catégorie", [""] + sorted(list(visa_map.keys())), index=0, key="vz_cat")
        with c2:
            subs = sorted(list(visa_map.get(cat_pick, {}).keys())) if cat_pick else []
            sub_pick = st.selectbox("Sous-catégorie", [""] + subs, index=0, key="vz_sub")

        if cat_pick and sub_pick:
            opts = sorted(list(visa_map.get(cat_pick, {}).get(sub_pick, [])))
            st.write("**Options détectées (depuis l’Excel, cellules = 1)** :")
            if opts:
                st.write(", ".join(opts))
                st.caption("En formulaire, les exclusifs COS/EOS ou USCIS/CP sont en **radio** ; le reste en **toggle/checkbox**.")
            else:
                st.caption("Aucune option : le Visa final = Sous-catégorie seule.")

        with st.expander("Aperçu complet du mapping"):
            for cat, submap in visa_map.items():
                st.write(f"**{cat}**")
                for sous, arr in submap.items():
                    st.caption(f"- {sous} → {', '.join(arr)}")


# ---------- EXPORTS & ANALYSES ----------
st.markdown("---")
st.subheader("📤 Exports & 📈 Analyses rapides")

def _excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        _uniquify_columns(df).to_excel(wr, sheet_name=sheet_name, index=False)
    bio.seek(0)
    return bio.getvalue()

def _zip_bytes(clients_df: pd.DataFrame, visa_file: str) -> bytes:
    bio = BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Clients.xlsx", _excel_bytes(clients_df.copy(), SHEET_CLIENTS))
        try:
            zf.write(visa_file, arcname="Visa.xlsx")
        except Exception:
            try:
                visa_df = pd.read_excel(visa_file, sheet_name=SHEET_VISA)
                zf.writestr("Visa.xlsx", _excel_bytes(visa_df, SHEET_VISA))
            except Exception:
                pass
    bio.seek(0)
    return bio.getvalue()

cE1, cE2, cE3 = st.columns(3)
with cE1:
    try:
        df_to_dl = _read_clients()
        st.download_button(
            "⬇️ Télécharger Clients.xlsx",
            data=_excel_bytes(df_to_dl, SHEET_CLIENTS),
            file_name="Clients.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_clients_xlsx",
        )
    except Exception as e:
        st.caption(f"Impossible d'exporter Clients.xlsx : {e}")

with cE2:
    try:
        st.download_button(
            "⬇️ Télécharger Visa.xlsx",
            data=Path(VISA_FILE).read_bytes(),
            file_name="Visa.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_visa_xlsx",
        )
    except Exception as e:
        st.caption(f"Impossible d'exporter Visa.xlsx : {e}")

with cE3:
    try:
        st.download_button(
            "📦 ZIP : Clients + Visa",
            data=_zip_bytes(_read_clients(), VISA_FILE),
            file_name="Visa_Manager_Export.zip",
            mime="application/zip",
            key="dl_zip_all",
        )
    except Exception as e:
        st.caption(f"Impossible de créer le ZIP : {e}")

st.markdown("---")
st.markdown("### 📊 Analyses rapides")

base = _read_clients()
if base.empty:
    st.info("Pas de données clients pour le moment.")
else:
    a1, a2, a3, a4 = st.columns(4)
    years  = sorted([int(y) for y in pd.to_numeric(base["_Année_"], errors="coerce").dropna().unique().tolist()])
    months = [f"{m:02d}" for m in range(1, 13)]
    cats   = sorted(base["Categorie"].dropna().astype(str).unique().tolist())
    visas  = sorted(base["Visa"].dropna().astype(str).unique().tolist())

    sel_years  = a1.multiselect("Année", years, default=[], key="an_years")
    sel_months = a2.multiselect("Mois (MM)", months, default=[], key="an_months")
    sel_cats   = a3.multiselect("Catégories", cats, default=[], key="an_cats")
    sel_visas  = a4.multiselect("Visa", visas, default=[], key="an_visas")

    f = base.copy()
    if sel_years:  f = f[f["_Année_"].isin(sel_years)]
    if sel_months: f = f[f["Mois"].isin(sel_months)]
    if sel_cats:   f = f[f["Categorie"].astype(str).isin(sel_cats)]
    if sel_visas:  f = f[f["Visa"].astype(str).isin(sel_visas)]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Honoraires", _fmt_money(_safe_num_series(f, "Montant honoraires (US $)").sum()))
    k3.metric("Payé", _fmt_money(_safe_num_series(f, "Payé").sum()))
    k4.metric("Solde", _fmt_money(_safe_num_series(f, "Reste").sum()))

    st.markdown("#### Volumes par catégorie")
    vol_cat = f.groupby(["Categorie"], dropna=True).size().reset_index(name="Dossiers")
    st.dataframe(vol_cat.sort_values("Dossiers", ascending=False).reset_index(drop=True), use_container_width=True)

    st.markdown("#### Volumes par sous-catégorie")
    vol_sub = f.groupby(["Sous-categorie"], dropna=True).size().reset_index(name="Dossiers")
    st.dataframe(vol_sub.sort_values("Dossiers", ascending=False).reset_index(drop=True), use_container_width=True)

    st.markdown("#### Volumes par Visa")
    vol_visa = f.groupby(["Visa"], dropna=True).size().reset_index(name="Dossiers")
    st.dataframe(vol_visa.sort_values("Dossiers", ascending=False).reset_index(drop=True), use_container_width=True)

    st.markdown("#### Détails (clients filtrés)")
    detail = f.copy()
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        if c in detail.columns:
            detail[c] = _safe_num_series(detail, c).map(_fmt_money)
    if "Date" in detail.columns:
        detail["Date"] = detail["Date"].astype(str)
    detail["Options"] = detail["Options"].apply(
        lambda d: json.dumps(_normalize_options_json(d), ensure_ascii=False) if not isinstance(d, str) else d
    )

    show_cols = [c for c in [
        "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
        "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste","Options"
    ] if c in detail.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in detail.columns]
    detail = detail.sort_values(by=sort_cols) if sort_cols else detail
    st.dataframe(_uniquify_columns(detail[show_cols].reset_index(drop=True)), use_container_width=True)