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

st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager (lecture des cases via 1 dans Visa.xlsx)")

# ==============================
# CONSTANTES
# ==============================
CLIENTS_FILE = "donnees_visa_clients1_adapte.xlsx"
VISA_FILE = "donnees_visa_clients1.xlsx"
SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"

# ==============================
# HELPERS GÉNÉRAUX
# ==============================
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

# ==============================
# PARSING VISA (lecture des "1")
# ==============================
@st.cache_data(show_spinner=False)
def parse_visa_sheet(xlsx_path: str | Path, sheet_name: str | None = None) -> dict[str, dict[str, list[str]]]:
    """
    Retourne:
    {
      "Categorie": {
          "Sous-categorie": ["Sous-categorie COS", "Sous-categorie EOS", ...]
      }
    }
    Une cellule vaut 1 (ou "1", True, "oui") => case cochée.
    """
    def _is_checked(v) -> bool:
        if v is None or (isinstance(v, float) and pd.isna(v)): 
            return False
        if isinstance(v, (int, float)): return float(v) == 1.0
        s = str(v).strip().lower()
        return s in {"1", "true", "vrai", "oui", "yes", "x"}

    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception:
        return {}

    sheets_to_try = [sheet_name] if sheet_name else xls.sheet_names
    for sn in sheets_to_try:
        if sn is None: continue
        try:
            dfv = pd.read_excel(xlsx_path, sheet_name=sn)
        except Exception:
            continue
        if dfv.empty:
            continue

        dfv = _uniquify_columns(dfv)
        dfv.columns = dfv.columns.map(str).str.strip()

        # Localiser colonnes principales
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
        if not cat_col: continue
        if not sub_col:
            dfv["_Sous_"] = ""
            sub_col = "_Sous_"

        check_cols = [c for c in dfv.columns if c not in {cat_col, sub_col}]
        out = {}
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

# ==============================
# RENDU DES CASES DYNAMIQUES
# ==============================
def render_dynamic_checkboxes(cat: str, sub: str, visa_map: dict, keyprefix: str) -> tuple[str, str]:
    if not (cat and sub):
        return "", "Choisir Catégorie et Sous-catégorie."
    raw_opts = sorted(list(visa_map.get(cat, {}).get(sub, [])))
    if raw_opts:
        sufs = []
        for lab in raw_opts:
            lab = _safe_str(lab)
            suf = lab[len(sub):].strip() if lab.startswith(sub) else lab
            if suf:
                sufs.append(suf)
        sufs = sorted(set(sufs))
        checked = []
        for i, s in enumerate(sufs):
            if st.checkbox(s, key=f"{keyprefix}_chk_{i}"):
                checked.append(s)
        if len(checked) == 0:
            return "", "Veuillez cocher une case."
        if len(checked) > 1:
            return "", "Une seule case possible."
        return f"{sub} {checked[0]}".strip(), ""
    else:
        st.caption("Aucune option : Visa = sous-catégorie seule.")
        return sub, ""

# ==============================
# LECTURE DES FICHIERS
# ==============================
def ensure_file(path: str, sheet_name: str, cols: list[str]) -> None:
    p = Path(path)
    if not p.exists():
        df = pd.DataFrame(columns=cols)
        with pd.ExcelWriter(p, engine="openpyxl") as wr:
            df.to_excel(wr, sheet_name=sheet_name, index=False)

CLIENTS_COLS = [
    "Dossier N","ID_Client","Nom","Date","Mois",
    "Categorie","Sous-categorie","Visa",
    "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
    "Payé","Reste","Paiements"
]
ensure_file(CLIENTS_FILE, SHEET_CLIENTS, CLIENTS_COLS)
ensure_file(VISA_FILE, SHEET_VISA, ["Categorie","Sous-categorie 1"])

visa_map = parse_visa_sheet(VISA_FILE)

# Aperçu visa dans la sidebar
with st.sidebar.expander("🔍 Aperçu Visa détecté"):
    if visa_map:
        for cat, submap in visa_map.items():
            st.write(f"**{cat}**")
            for sous, opts in submap.items():
                st.caption(f"{sous} → {', '.join(opts)}")
    else:
        st.caption("Aucune donnée Visa détectée.")


# ==============================
# ONGLETS
# ==============================
tab_clients, tab_visa = st.tabs(["👥 Clients", "📄 Visa"])

# ==============================
# FONCTIONS CLIENTS
# ==============================
def _normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Colonnes minimales
    base_cols = [
        "Dossier N","ID_Client","Nom","Date","Mois",
        "Categorie","Sous-categorie","Visa",
        "Montant honoraires (US $)","Autres frais (US $)","Total (US $)",
        "Payé","Reste","Paiements"
    ]
    for c in base_cols:
        if c not in df.columns:
            df[c] = None

    # Date / Mois
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df["Mois"] = df.apply(
        lambda r: f"{pd.to_datetime(r['Date']).month:02d}" if pd.notna(r["Date"]) else (_safe_str(r.get("Mois",""))[:2] or None),
        axis=1
    )

    # Numériques
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        df[c] = _safe_num_series(df, c)
    df["Total (US $)"] = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]

    # Paiements JSON
    def _parse_p(x):
        try:
            j = json.loads(_safe_str(x) or "[]")
            return j if isinstance(j, list) else []
        except Exception:
            return []
    df["Paiements"] = df["Paiements"].apply(_parse_p)

    # Payé = max(colonne Payé, somme JSON)
    def _sum_json(lst):
        try:
            return float(sum(float(it.get("amount",0.0) or 0.0) for it in (lst or [])))
        except Exception:
            return 0.0
    paid_json = df["Paiements"].apply(_sum_json)
    df["Payé"] = pd.concat([df["Payé"].fillna(0.0).astype(float), paid_json], axis=1).max(axis=1)

    df["Reste"] = (df["Total (US $)"] - df["Payé"]).clip(lower=0.0)

    # Clés de tri
    df["_Année_"]   = df["Date"].apply(lambda d: d.year if pd.notna(d) else pd.NA)
    df["_MoisNum_"] = df["Date"].apply(lambda d: d.month if pd.notna(d) else pd.NA)

    return _uniquify_columns(df)

def _read_clients() -> pd.DataFrame:
    df = pd.read_excel(CLIENTS_FILE, sheet_name=SHEET_CLIENTS)
    return _normalize_clients(df)

def _write_clients(df: pd.DataFrame) -> None:
    df = df.copy()
    with pd.ExcelWriter(CLIENTS_FILE, engine="openpyxl", mode="w") as wr:
        df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if "Dossier N" in df.columns:
        s = pd.to_numeric(df["Dossier N"], errors="coerce")
        if s.notna().any():
            return int(s.max()) + 1
    return int(start)

def _make_client_id(nom: str, d: date) -> str:
    # id = NOM-YYYYMMDD ; si déjà existant, suffixe -1, -2, ...
    base = f"{_safe_str(nom).strip().replace(' ', '_')}-{d:%Y%m%d}"
    return base

# ==============================
# DONNÉES CHARGÉES
# ==============================
df_clients = _read_clients()

# ==============================
# ONGLET CLIENTS
# ==============================
with tab_clients:
    st.subheader("Créer, modifier, supprimer un client — et gérer ses paiements")

    left, right = st.columns([1,1], gap="large")

    # ---------- LISTE / SÉLECTION ----------
    with left:
        st.markdown("### 🔎 Sélection d’un client existant")
        if df_clients.empty:
            st.info("Aucun client pour l’instant.")
            sel_idx = None
            sel_row = None
        else:
            labels = (df_clients.get("Nom","").astype(str) + " — " + df_clients.get("ID_Client","").astype(str)).fillna("")
            sel_idx = st.selectbox(
                "Client",
                options=list(df_clients.index),
                format_func=lambda i: labels.iloc[i],
                key="sel_cli_idx"
            )
            sel_row = df_clients.loc[sel_idx] if sel_idx is not None else None

        st.markdown("### 📋 Tous les clients")
        view = df_clients.copy()
        # format $ pour lecture
        for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
            if c in view.columns:
                view[c] = _safe_num_series(view, c).map(_fmt_money)
        if "Date" in view.columns:
            view["Date"] = view["Date"].astype(str)

        show_cols = [c for c in [
            "Dossier N","ID_Client","Nom","Date","Mois","Categorie","Sous-categorie","Visa",
            "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"
        ] if c in view.columns]
        sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in view.columns]
        view = view.sort_values(by=sort_cols) if sort_cols else view
        st.dataframe(_uniquify_columns(view[show_cols].reset_index(drop=True)), use_container_width=True)

    # ---------- CRÉATION ----------
    with right:
        st.markdown("### ➕ Nouveau client")
        new_nom  = st.text_input("Nom", key="new_nom")
        new_date = st.date_input("Date de création", value=date.today(), key="new_date")

        # Catégorie
        cats = sorted(list(visa_map.keys()))
        new_cat = st.selectbox("Catégorie", options=[""]+cats, index=0, key="new_cat")

        # Sous-catégorie dépendante
        sub_opts = sorted(list(visa_map.get(new_cat, {}).keys())) if new_cat else []
        new_sub  = st.selectbox("Sous-catégorie", options=[""]+sub_opts, index=0, key="new_sub")

        # Cases dynamiques lues depuis la feuille Visa (cellules = 1)
        st.caption("Options disponibles (cocher **une** seule si présentes)")
        new_visa, err_new = render_dynamic_checkboxes(new_cat, new_sub, visa_map, keyprefix="new")
        if err_new:
            st.info(err_new)

        new_hono = st.number_input("Montant honoraires (US $)", min_value=0.0, step=10.0, format="%.2f", key="new_hono")
        new_autr = st.number_input("Autres frais (US $)",     min_value=0.0, step=10.0, format="%.2f", key="new_autre")

        if st.button("💾 Créer", key="btn_create_client"):
            if not new_nom:
                st.warning("Nom obligatoire."); st.stop()
            if not new_cat:
                st.warning("Catégorie obligatoire."); st.stop()
            if not new_sub:
                st.warning("Sous-catégorie obligatoire."); st.stop()

            # Visa final : si pas d’options, la fonction renvoie la sous-cat seule.
            if new_visa == "":
                st.warning("Sélectionnez une case (si disponible) — sinon la sous-catégorie seule sera enregistrée.")
                # on accepte tout de même : remplace par sous-cat seule
                new_visa = new_sub or ""

            base = _read_clients()
            dossier = _next_dossier(base)
            cid_base = _make_client_id(new_nom, new_date)
            cid = cid_base
            i = 0
            while (base["ID_Client"].astype(str) == cid).any():
                i += 1
                cid = f"{cid_base}-{i}"

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
            }
            base = pd.concat([base, pd.DataFrame([row])], ignore_index=True)
            base = _normalize_clients(base)
            _write_clients(base)
            st.success("✅ Client créé.")
            st.rerun()

    st.markdown("---")

    # ---------- MODIFICATION & PAIEMENTS ----------
    if sel_row is not None:
        idx = sel_idx
        ed = sel_row.to_dict()

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
            st.caption("Options disponibles (cocher **une** seule si présentes)")
            # Pré-cocher selon visa actuel (une seule fois)
            if not st.session_state.get(f"prefill_{ed.get('ID_Client')}", False):
                curr_v = _safe_str(ed.get("Visa",""))
                raw_opts = sorted(list(visa_map.get(ed_cat, {}).get(curr_sub, [])))
                # suffixes
                sufs = []
                for lab in raw_opts:
                    lab = _safe_str(lab)
                    suf = lab[len(curr_sub):].strip() if (curr_sub and lab.startswith(curr_sub)) else lab
                    if suf:
                        sufs.append(suf)
                sufs = sorted(set(sufs))
                # suffixe actuel
                curr_suffix = curr_v[len(curr_sub):].strip() if (curr_sub and curr_v.startswith(curr_sub)) else ""
                for i, s in enumerate(sufs):
                    st.session_state[f"ed_{idx}_chk_{i}"] = (s == curr_suffix)
                st.session_state[f"prefill_{ed.get('ID_Client')}"] = True

            ed_visa_final, err_ed = render_dynamic_checkboxes(ed_cat, ed_sub, visa_map, keyprefix=f"ed_{idx}")
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

# ==============================
# ONGLET VISA (aperçu & test)
# ==============================
with tab_visa:
    st.subheader("Référentiel Visa (lecture des cases = 1)")

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
            st.write("**Options disponibles (depuis l’Excel, cellules = 1)** :")
            if opts:
                st.write(", ".join(opts))
                st.caption("En formulaire, ces options apparaissent sous forme de **cases**. Un seul choix autorisé.")
            else:
                st.caption("Aucune option : le Visa final = Sous-catégorie seule.")

        with st.expander("Aperçu complet du mapping"):
            for cat, submap in visa_map.items():
                st.write(f"**{cat}**")
                for sous, arr in submap.items():
                    st.caption(f"- {sous} → {', '.join(arr)}")


# ==============================
# EXPORTS & ANALYSES RAPIDES
# ==============================

st.markdown("---")
st.subheader("📤 Exports & 📈 Analyses rapides")

# --- Utilitaires d'export ---
def _excel_bytes(df: pd.DataFrame, sheet_name: str) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        _uniquify_columns(df).to_excel(wr, sheet_name=sheet_name, index=False)
    bio.seek(0)
    return bio.getvalue()

def _zip_bytes(clients_df: pd.DataFrame, visa_file: str) -> bytes:
    bio = BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        # Clients
        zf.writestr("Clients.xlsx", _excel_bytes(clients_df.copy(), SHEET_CLIENTS))
        # Visa (on prend le fichier tel quel pour préserver l'onglet, les formules, etc.)
        try:
            zf.write(visa_file, arcname="Visa.xlsx")
        except Exception:
            # sec fallback: on relit l'onglet SHEET_VISA si possible
            try:
                visa_df = pd.read_excel(visa_file, sheet_name=SHEET_VISA)
                zf.writestr("Visa.xlsx", _excel_bytes(visa_df, SHEET_VISA))
            except Exception:
                pass
    bio.seek(0)
    return bio.getvalue()

# --- Barre d'exports ---
cE1, cE2, cE3 = st.columns(3)
with cE1:
    # Export Clients (version normalisée actuelle en mémoire)
    try:
        df_to_dl = _read_clients()  # relit pour être sûr d'avoir la dernière écriture
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
    # Export Visa: on renvoie le fichier brut (préserve les entêtes exacts et les '1')
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
    # Export ZIP des deux
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

# ==============================
# ANALYSES (tableaux simples)
# ==============================
st.markdown("### 📊 Analyses rapides")

base = _read_clients()
if base.empty:
    st.info("Pas de données clients pour le moment.")
else:
    # Filtres
    a1, a2, a3, a4 = st.columns(4)
    years  = sorted([int(y) for y in pd.to_numeric(base["_Année_"], errors="coerce").dropna().unique().tolist()])
    months = [f"{m:02d}" for m in range(1, 13)]
    cats   = sorted(base["Categorie"].dropna().astype(str).unique().tolist())
    visas  = sorted(base["Visa"].dropna().astype(str).unique().tolist())

    sel_years  = a1.multiselect("Année", years, default=[], key="an3_years")
    sel_months = a2.multiselect("Mois (MM)", months, default=[], key="an3_months")
    sel_cats   = a3.multiselect("Catégories", cats, default=[], key="an3_cats")
    sel_visas  = a4.multiselect("Visa", visas, default=[], key="an3_visas")

    f = base.copy()
    if sel_years:  f = f[f["_Année_"].isin(sel_years)]
    if sel_months: f = f[f["Mois"].isin(sel_months)]
    if sel_cats:   f = f[f["Categorie"].astype(str).isin(sel_cats)]
    if sel_visas:  f = f[f["Visa"].astype(str).isin(sel_visas)]

    # KPIs
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Honoraires", _fmt_money(_safe_num_series(f, "Montant honoraires (US $)").sum()))
    k3.metric("Payé", _fmt_money(_safe_num_series(f, "Payé").sum()))
    k4.metric("Solde", _fmt_money(_safe_num_series(f, "Reste").sum()))

    st.markdown("#### Volumes par catégorie")
    vol_cat = f.groupby(["Categorie"], dropna=True).size().reset_index(name="Dossiers")
    if not vol_cat.empty:
        st.dataframe(vol_cat.sort_values("Dossiers", ascending=False).reset_index(drop=True), use_container_width=True)
    else:
        st.caption("Aucune donnée après filtres.")

    st.markdown("#### Volumes par sous-catégorie")
    vol_sub = f.groupby(["Sous-categorie"], dropna=True).size().reset_index(name="Dossiers")
    if not vol_sub.empty:
        st.dataframe(vol_sub.sort_values("Dossiers", ascending=False).reset_index(drop=True), use_container_width=True)
    else:
        st.caption("Aucune donnée après filtres.")

    st.markdown("#### Volumes par Visa")
    vol_visa = f.groupby(["Visa"], dropna=True).size().reset_index(name="Dossiers")
    if not vol_visa.empty:
        st.dataframe(vol_visa.sort_values("Dossiers", ascending=False).reset_index(drop=True), use_container_width=True)
    else:
        st.caption("Aucune donnée après filtres.")

    st.markdown("#### Détails (clients filtrés)")
    detail = f.copy()
    for c in ["Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"]:
        if c in detail.columns:
            detail[c] = _safe_num_series(detail, c).map(_fmt_money)
    if "Date" in detail.columns:
        detail["Date"] = detail["Date"].astype(str)

    show_cols = [c for c in [
        "Dossier N","ID_Client","Nom","Categorie","Sous-categorie","Visa","Date","Mois",
        "Montant honoraires (US $)","Autres frais (US $)","Total (US $)","Payé","Reste"
    ] if c in detail.columns]
    sort_cols = [c for c in ["_Année_","_MoisNum_","Categorie","Nom"] if c in detail.columns]
    detail = detail.sort_values(by=sort_cols) if sort_cols else detail
    st.dataframe(_uniquify_columns(detail[show_cols].reset_index(drop=True)), use_container_width=True)