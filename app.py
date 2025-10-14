# ================================
# 🛂 Visa Manager — PARTIE 1/4
# Imports, constantes, helpers, I/O, normalisation
# ================================
from __future__ import annotations

import json, os, re, zipfile
from io import BytesIO
from datetime import date, datetime
from typing import Any, Dict, List, Tuple

import pandas as pd
import streamlit as st

# ---------- Constantes colonnes ----------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

DOSSIER_COL = "Dossier N"
HONO = "Montant honoraires (US $)"
AUTRE = "Autres frais (US $)"
TOTAL = "Total (US $)"

# ---------- Dossier mémoire locale ----------
APP_STORE = ".visa_manager_store"
os.makedirs(APP_STORE, exist_ok=True)
LAST_STATE_JSON = os.path.join(APP_STORE, "last_paths.json")

# ---------- Petite identité de session pour les clés Streamlit ----------
def _sid() -> str:
    # 6 chars à partir de la session pour ne pas dupliquer les keys
    return st.session_state.get("_sid_", None) or st.session_state.setdefault("_sid_", hex(abs(hash(id(st.session_state))))[2:8])

def skey(*parts: str) -> str:
    return "k_" + _sid() + "_" + "_".join([re.sub(r"[^a-zA-Z0-9_]+", "", str(p)) for p in parts])

# ---------- Helpers sûrs ----------
def _safe_str(v: Any) -> str:
    try:
        if v is None: return ""
        return str(v)
    except Exception:
        return ""

def _fmt_money(v: float) -> str:
    try:
        return f"${v:,.2f}"
    except Exception:
        return "$0.00"

def _fmt_money_us(v: float) -> str:
    return _fmt_money(float(v or 0.0))

def _to_num_series(s: pd.Series) -> pd.Series:
    if s is None: 
        return pd.Series([], dtype=float)
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0)
    s = s.astype(str)
    s = s.str.replace(r"[^\d,.\-]+", "", regex=True)
    # si virgule utilisée comme séparateur décimal
    def to_float(x: str) -> float:
        if x.count(",") == 1 and x.count(".") == 0:
            x = x.replace(",", ".")
        elif x.count(",") > 1 and x.count(".") == 0:
            x = x.replace(".", "").replace(",", ".")
        else:
            x = x.replace(",", "")
        try:
            return float(x)
        except Exception:
            return 0.0
    return s.map(to_float).astype(float)

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0] * len(df), dtype=float)
    return _to_num_series(df[col])

def _date_for_widget(v: Any) -> date | None:
    # Retourne un objet date (ou None) accepté par st.date_input
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    if isinstance(v, datetime):
        return v.date()
    try:
        d = pd.to_datetime(v, errors="coerce")
        if pd.isna(d): 
            return None
        return d.date()
    except Exception:
        return None

def _make_client_id(nom: str, d: Any) -> str:
    base = re.sub(r"[^A-Za-z0-9\-]+", "-", _safe_str(nom)).strip("-").lower() or "client"
    dd = _date_for_widget(d) or date.today()
    # suffix court YYYYMMDD + 3 chars hasard
    suf = f"{dd:%Y%m%d}"
    # si le nom existe déjà, un nombre -0,-1… sera géré par l’appelant si besoin
    return f"{base}-{suf}"

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    if df is None or df.empty or DOSSIER_COL not in df.columns:
        return start
    try:
        cur = pd.to_numeric(df[DOSSIER_COL], errors="coerce")
        mx = int(pd.Series(cur).dropna().max()) if pd.Series(cur).notna().any() else start-1
        return max(start, mx + 1)
    except Exception:
        return start

# ---------- Mémoire chemins (dernier fichier utilisé) ----------
def _save_last_paths(clients_path: str | None, visa_path: str | None, mode: str) -> None:
    try:
        payload = {
            "mode": mode,
            "clients_path": clients_path or "",
            "visa_path": visa_path or "",
            "ts": datetime.now().isoformat(timespec="seconds"),
        }
        with open(LAST_STATE_JSON, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def _load_last_paths() -> Tuple[str | None, str | None, str]:
    try:
        if os.path.exists(LAST_STATE_JSON):
            with open(LAST_STATE_JSON, "r", encoding="utf-8") as f:
                obj = json.load(f)
                return (
                    obj.get("clients_path") or None,
                    obj.get("visa_path") or None,
                    obj.get("mode") or "two",
                )
    except Exception:
        pass
    return (None, None, "two")

# ---------- Lecture/écriture Excel en mémoire ----------
@st.cache_data(show_spinner=False)
def read_excel_sheet(path: str, sheet: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet)

def write_df_to_bytes(df: pd.DataFrame, sheet_name: str = SHEET_CLIENTS) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        df.to_excel(wr, index=False, sheet_name=sheet_name)
    return bio.getvalue()

def write_two_sheets_to_bytes(df_clients: pd.DataFrame, df_visa: pd.DataFrame) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as wr:
        df_clients.to_excel(wr, index=False, sheet_name=SHEET_CLIENTS)
        df_visa.to_excel(wr, index=False, sheet_name=SHEET_VISA)
    return bio.getvalue()

# ---------- Normalisation Clients ----------
def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
            "Categorie", "Sous-categorie", "Visa",
            HONO, AUTRE, TOTAL, "Payé", "Reste",
            "Paiements", "Options", "Commentaire",
            "Dossier envoyé","Date d'envoi",
            "Dossier accepté","Date d'acceptation",
            "Dossier refusé","Date de refus",
            "Dossier annulé","Date d'annulation",
            "RFE"
        ])
    df = df.copy()

    # Renommer colonnes tolérant variantes
    ren = {}
    for c in df.columns:
        cn = _safe_str(c).strip()
        if cn.lower() in ["categorie", "catégorie"]: ren[c] = "Categorie"
        elif cn.lower().startswith("sous-categorie"): ren[c] = "Sous-categorie"
        elif cn.lower().startswith("montant honoraires"): ren[c] = HONO
        elif cn.lower().startswith("autres frais"): ren[c] = AUTRE
        elif cn.lower().startswith("total"): ren[c] = TOTAL
        elif cn.lower() in ["paye","payé","payes","payés"]: ren[c] = "Payé"
        elif cn.lower() == "reste": ren[c] = "Reste"
        elif cn.lower() in ["paiement","paiements"]: ren[c] = "Paiements"
        elif cn.lower() in ["option","options"]: ren[c] = "Options"
        elif cn.lower() in ["commentaire","commentaires","notes"]: ren[c] = "Commentaire"
        elif cn.lower() in ["dossier n","dossier no","dossier"]: ren[c] = DOSSIER_COL
    if ren:
        df.rename(columns=ren, inplace=True)

    # Colonnes minimales
    for c in [DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
              "Categorie", "Sous-categorie", "Visa",
              HONO, AUTRE, TOTAL, "Payé", "Reste",
              "Paiements", "Options", "Commentaire",
              "Dossier envoyé","Date d'envoi",
              "Dossier accepté","Date d'acceptation",
              "Dossier refusé","Date de refus",
              "Dossier annulé","Date d'annulation",
              "RFE"]:
        if c not in df.columns:
            df[c] = "" if c in ["ID_Client","Nom","Categorie","Sous-categorie","Visa","Commentaire"] else 0

    # Numériques
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = _to_num_series(df[c])

    # Date / Mois
    try:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    except Exception:
        pass
    df["Mois"] = df["Mois"].astype(str).str.replace(r"[^\d]", "", regex=True).str.zfill(2)

    # Statuts → int(0/1)
    for c in ["Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    # Options / Paiements → objet structuré
    def _json_or_list(v):
        if isinstance(v, (list, dict)): return v
        s = _safe_str(v).strip()
        if not s: return [] if "Paiements" else {}
        try:
            obj = json.loads(s)
            return obj
        except Exception:
            return [] if "Paiements" else {}
    # Paiements
    def _parse_pays(v):
        if isinstance(v, list): return v
        s = _safe_str(v).strip()
        if not s: return []
        try:
            obj = json.loads(s)
            if isinstance(obj, list): return obj
        except Exception:
            pass
        return []

    df["Paiements"] = df["Paiements"].apply(_parse_pays)
    def _parse_opts(v):
        if isinstance(v, dict): return v
        s = _safe_str(v).strip()
        if not s: return {"exclusive": None, "options": []}
        try:
            obj = json.loads(s)
            if isinstance(obj, dict): return obj
        except Exception:
            pass
        return {"exclusive": None, "options": []}
    df["Options"] = df["Options"].apply(_parse_opts)

    # Recalcul Total si non cohérent
    fix_total = (df[TOTAL] == 0) | df[TOTAL].isna()
    df.loc[fix_total, TOTAL] = _to_num_series(df[HONO]) + _to_num_series(df[AUTRE])

    # Recalcul Reste = Total - Payé
    df["Payé"] = _to_num_series(df["Payé"])
    df["Reste"] = (df[TOTAL] - df["Payé"]).clip(lower=0.0)

    # Colonnes techniques pour tris
    df["_Année_"]   = pd.to_datetime(df["Date"], errors="coerce").dt.year
    df["_MoisNum_"] = pd.to_numeric(df["Mois"], errors="coerce")

    return df

# ---------- Lecture Clients (depuis chemin mémorisé) ----------
def _read_clients(path: str | None = None) -> pd.DataFrame:
    # priorité au path fourni, sinon dernier chemin mémorisé
    c_last, _, _mode = _load_last_paths()
    xls_path = path or c_last
    if not xls_path or not os.path.exists(xls_path):
        return pd.DataFrame()
    try:
        # accepter soit un classeur 2 onglets (Clients), soit un fichier mono-feuille
        try:
            df = pd.read_excel(xls_path, sheet_name=SHEET_CLIENTS)
        except Exception:
            df = pd.read_excel(xls_path)
        return normalize_clients(df)
    except Exception:
        return pd.DataFrame()

def _write_clients(df: pd.DataFrame, path: str | None = None) -> None:
    df = normalize_clients(df)
    c_last, _, _ = _load_last_paths()
    xls_path = path or c_last
    # si pas de chemin, stocker temporairement en mémoire (session) + proposer un download ailleurs
    if not xls_path:
        st.session_state["__clients_df_cache__"] = df.copy()
        return
    # si le chemin est un classeur 2 onglets, on écrase uniquement l’onglet Clients
    try:
        # lecture du classeur existant pour préserver l’onglet Visa si présent
        with pd.ExcelWriter(xls_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
            df.to_excel(wr, index=False, sheet_name=SHEET_CLIENTS)
    except Exception:
        # sinon, on écrit un fichier mono-onglet
        bytes_out = write_df_to_bytes(df, sheet_name=SHEET_CLIENTS)
        with open(xls_path, "wb") as f:
            f.write(bytes_out)

# ---------- Lecture Visa (structure) ----------
@st.cache_data(show_spinner=False)
def read_visa_raw(path: str | None) -> pd.DataFrame:
    if not path or not os.path.exists(path):
        return pd.DataFrame()
    try:
        try:
            df = pd.read_excel(path, sheet_name=SHEET_VISA)
        except Exception:
            df = pd.read_excel(path)
        return df
    except Exception:
        return pd.DataFrame()

# ---------- Construction du dictionnaire de visas ----------
def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    Fichier Visa attendu :
    - Colonnes obligatoires : 'Categorie', 'Sous-categorie'
    - Colonnes optionnelles : cases exclusives (ex: COS, EOS) → on les marque en 'exclusive'
                              autres cases → 'options'
    Convention : mettre 1 dans la cellule si l’option s’applique à la sous-catégorie.
    """
    if df_visa is None or df_visa.empty:
        return {}

    df = df_visa.copy()
    # normaliser noms
    rename = {}
    for c in df.columns:
        cl = _safe_str(c).strip()
        if cl.lower() in ["categorie","catégorie"]: rename[c] = "Categorie"
        elif cl.lower().startswith("sous-categorie"): rename[c] = "Sous-categorie"
    if rename:
        df.rename(columns=rename, inplace=True)

    for col in ["Categorie","Sous-categorie"]:
        if col not in df.columns:
            df[col] = ""

    # détecter colonnes d’options (toutes sauf Catégorie / Sous-catégorie)
    opt_cols = [c for c in df.columns if c not in ["Categorie","Sous-categorie"]]

    vm: Dict[str, Dict[str, Dict[str, Any]]] = {}
    for _, r in df.iterrows():
        cat = _safe_str(r.get("Categorie","")).strip()
        sub = _safe_str(r.get("Sous-categorie","")).strip()
        if not cat or not sub:
            continue

        # options présentes (valeur == 1)
        active = []
        for oc in opt_cols:
            val = r.get(oc, 0)
            try:
                v = float(val)
            except Exception:
                v = 0.0
            if v == 1.0:
                active.append(_safe_str(oc).strip())

        # appliquer heuristique : si l’ensemble {"COS","EOS"} ⊆ active → exclusif
        exclusive = []
        others = []
        if "COS" in active or "EOS" in active:
            # exclusif seulement si au moins un des deux cochés dans la ligne
            for x in ["COS","EOS"]:
                if x in active:
                    exclusive.append(x)
            # retirer COS/EOS des "others"
            others = [x for x in active if x not in ("COS","EOS")]
        else:
            exclusive = []
            others = active

        vm.setdefault(cat, {}).setdefault(sub, {"exclusive": None, "options": []})
        vm[cat][sub]["exclusive"] = exclusive if exclusive else None
        vm[cat][sub]["options"] = others

    return vm



# ================================
# 🧭 PARTIE 2/4 — Configuration Streamlit + chargement Excel + setup onglets
# ================================

st.set_page_config(
    page_title="Visa Manager",
    page_icon="🛂",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🛂 Visa Manager — Gestion Clients et Visas")

# ================================
# 📁 Chargement des fichiers
# ================================
st.sidebar.header("📂 Fichiers à charger")

mode = st.sidebar.radio(
    "Mode d'import",
    ["Fichier unique (Clients+Visa)", "Fichiers séparés"],
    horizontal=False,
    key=skey("mode","import")
)
mode_tag = "one" if mode.startswith("Fichier unique") else "two"

# chemins déjà mémorisés
c_last, v_last, _mode = _load_last_paths()

# --- Cas 1 : Fichier unique (2 onglets)
if mode_tag == "one":
    file_one = st.sidebar.file_uploader(
        "Classeur Excel unique (.xlsx, onglets Clients & Visa)",
        type=["xlsx"], key=skey("up","one")
    )
    if file_one:
        clients_path = os.path.join(APP_STORE, "import_clients_visa.xlsx")
        with open(clients_path, "wb") as f:
            f.write(file_one.getbuffer())
        visa_path = clients_path
        _save_last_paths(clients_path, visa_path, "one")
    else:
        clients_path = c_last if (_mode == "one" and c_last and os.path.exists(c_last)) else None
        visa_path = v_last if (_mode == "one" and v_last and os.path.exists(v_last)) else None

# --- Cas 2 : Fichiers séparés
else:
    file_clients = st.sidebar.file_uploader(
        "Fichier Clients (.xlsx)",
        type=["xlsx"], key=skey("up","clients")
    )
    file_visa = st.sidebar.file_uploader(
        "Fichier Visa (.xlsx)",
        type=["xlsx"], key=skey("up","visa")
    )

    clients_path, visa_path = None, None
    if file_clients:
        clients_path = os.path.join(APP_STORE, "import_clients.xlsx")
        with open(clients_path, "wb") as f:
            f.write(file_clients.getbuffer())
    elif c_last and os.path.exists(c_last):
        clients_path = c_last

    if file_visa:
        visa_path = os.path.join(APP_STORE, "import_visa.xlsx")
        with open(visa_path, "wb") as f:
            f.write(file_visa.getbuffer())
    elif v_last and os.path.exists(v_last):
        visa_path = v_last

    _save_last_paths(clients_path, visa_path, "two")

# ================================
# 🧾 Lecture des données
# ================================
df_clients_raw = _read_clients(clients_path)
df_visa_raw = read_visa_raw(visa_path)
visa_map = build_visa_map(df_visa_raw)

# ================================
# 🪟 Tabs principaux
# ================================
tab_labels = [
    "🏠 Accueil",
    "🏦 Escrow",
    "👤 Compte client",
    "🧾 Clients (CRUD)",
    "📄 Visa (aperçu)",
    "💾 Export"
]
tabs = st.tabs(tab_labels)

# ================================
# 🏠 ONGLET : Accueil
# ================================
with tabs[0]:
    st.subheader("🏠 Accueil — synthèse générale")

    if df_clients_raw.empty:
        st.info("Aucun client chargé. Veuillez importer un fichier dans la barre latérale.")
    else:
        df_all = normalize_clients(df_clients_raw)

        k1, k2, k3, k4 = st.columns(4)
        total_usd = float(df_all[TOTAL].sum())
        paye_usd = float(df_all["Payé"].sum())
        reste_usd = float(df_all["Reste"].sum())
        nb_clients = len(df_all)

        k1.metric("Total (US $)", _fmt_money(total_usd))
        k2.metric("Payé", _fmt_money(paye_usd))
        k3.metric("Reste", _fmt_money(reste_usd))
        k4.metric("Clients", nb_clients)

        st.markdown("### 📋 Aperçu du tableau Clients")
        st.dataframe(
            df_all[[DOSSIER_COL, "Nom", "Categorie", "Sous-categorie", "Visa", HONO, AUTRE, TOTAL, "Payé", "Reste"]],
            use_container_width=True,
            hide_index=True,
            key=skey("home","table")
        )

        # Histogramme simple par catégorie
        try:
            import plotly.express as px
            agg = df_all.groupby("Categorie", as_index=False)[TOTAL].sum().sort_values(TOTAL, ascending=False)
            fig = px.bar(agg, x="Categorie", y=TOTAL, text_auto=True)
            st.plotly_chart(fig, use_container_width=True, key=skey("home","chart"))
        except Exception:
            st.caption("Graphique non disponible (Plotly manquant ou données insuffisantes).")

    st.caption("ℹ️ Utilisez les autres onglets pour modifier les clients, gérer les paiements ou visualiser les visas.")



# ================================
# 🏦 ONGLET : Escrow — Synthèse financière
# ================================
with tabs[1]:
    st.subheader("🏦 Escrow — synthèse financière")

    if df_clients_raw.empty:
        st.info("Aucun client chargé.")
    else:
        df_all = normalize_clients(df_clients_raw)

        st.markdown("### 💰 Répartition par catégorie")
        agg = df_all.groupby("Categorie", as_index=False)[[TOTAL, "Payé", "Reste"]].sum()
        agg["% Payé"] = (agg["Payé"] / agg[TOTAL] * 100).round(1).fillna(0)
        st.dataframe(agg, use_container_width=True, hide_index=True, key=skey("escrow","table"))

        k1, k2, k3 = st.columns(3)
        k1.metric("Total (US $)", _fmt_money(agg[TOTAL].sum()))
        k2.metric("Payé", _fmt_money(agg["Payé"].sum()))
        k3.metric("Reste", _fmt_money(agg["Reste"].sum()))

        try:
            import plotly.express as px
            fig = px.pie(agg, values=TOTAL, names="Categorie", title="Répartition par catégorie")
            st.plotly_chart(fig, use_container_width=True, key=skey("escrow","pie"))
        except Exception:
            st.caption("Graphique non disponible (Plotly manquant ou données insuffisantes).")


# ================================
# 👤 ONGLET : Compte client — suivi & paiements
# ================================
with tabs[2]:
    st.subheader("👤 Compte client — suivi individuel")

    df_all = normalize_clients(df_clients_raw)
    if df_all.empty:
        st.info("Aucun client à afficher.")
    else:
        noms = sorted(df_all["Nom"].dropna().astype(str).unique().tolist())
        sel_nom = st.selectbox("Sélectionnez un client", [""] + noms, key=skey("cc","nom"))
        if sel_nom:
            row = df_all[df_all["Nom"] == sel_nom].iloc[0].to_dict()

            c1, c2, c3 = st.columns(3)
            c1.metric("Total (US $)", _fmt_money(row.get(TOTAL, 0)))
            c2.metric("Payé", _fmt_money(row.get("Payé", 0)))
            c3.metric("Reste", _fmt_money(row.get("Reste", 0)))

            st.markdown("### 📋 Détails du dossier")
            st.write({
                "Nom": row.get("Nom",""),
                "Catégorie": row.get("Categorie",""),
                "Sous-catégorie": row.get("Sous-categorie",""),
                "Visa": row.get("Visa",""),
                "Date": row.get("Date",""),
                "Mois": row.get("Mois",""),
            })

            st.markdown("### 💵 Paiements")
            paiements = row.get("Paiements", [])
            if isinstance(paiements, str):
                try:
                    paiements = json.loads(paiements)
                except Exception:
                    paiements = []

            dfp = pd.DataFrame(paiements)
            if not dfp.empty:
                dfp["Montant"] = dfp["Montant"].apply(_fmt_money)
                st.dataframe(dfp, use_container_width=True, hide_index=True, key=skey("cc","pmt"))
            else:
                st.info("Aucun paiement enregistré.")

            st.markdown("### ➕ Ajouter un paiement")
            pay_col1, pay_col2 = st.columns(2)
            p_montant = pay_col1.number_input("Montant (US $)", min_value=0.0, step=10.0, key=skey("cc","mont"))
            p_date = pay_col2.date_input("Date du paiement", value=date.today(), key=skey("cc","pdate"))
            p_btn = st.button("💾 Ajouter le paiement", key=skey("cc","addpay"))

            if p_btn:
                if p_montant <= 0:
                    st.warning("Le montant doit être supérieur à 0.")
                else:
                    new_pmt = {"Montant": float(p_montant),
                               "Date": p_date.strftime("%Y-%m-%d")}
                    if not isinstance(paiements, list):
                        paiements = []
                    paiements.append(new_pmt)
                    new_pay = row.get("Payé", 0) + p_montant
                    reste = max(0.0, row.get(TOTAL, 0) - new_pay)
                    df_all.loc[df_all["Nom"] == sel_nom, "Paiements"] = [paiements]
                    df_all.loc[df_all["Nom"] == sel_nom, "Payé"] = new_pay
                    df_all.loc[df_all["Nom"] == sel_nom, "Reste"] = reste
                    _write_clients(df_all, clients_path)
                    st.success("Paiement ajouté.")
                    st.cache_data.clear()
                    st.rerun()


# ================================
# 🧾 ONGLET : Gestion (CRUD)
# ================================
with tabs[3]:
    st.subheader("🧾 Gestion des clients (Ajouter / Modifier / Supprimer)")
    if df_clients_raw.empty:
        st.info("Aucun fichier Clients chargé.")
    else:
        op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"],
                      horizontal=True, key=skey("crud","op"))

        df_live = _read_clients(clients_path)

        # --- AJOUT CLIENT ---
        if op == "Ajouter":
            st.markdown("### ➕ Ajouter un client")
            c1, c2, c3 = st.columns(3)
            nom = c1.text_input("Nom", "", key=skey("add","nom"))
            dt = c2.date_input("Date de création", value=date.today(), key=skey("add","date"))
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)],
                                index=int(date.today().month)-1, key=skey("add","mois"))

            st.markdown("#### 🎯 Choix Visa")
            cats = sorted(list(visa_map.keys()))
            sel_cat = st.selectbox("Catégorie", [""] + cats, key=skey("add","cat"))
            sel_sub = ""
            visa_final = ""
            opts_dict = {"exclusive": None, "options": []}
            if sel_cat:
                subs = sorted(list(visa_map.get(sel_cat, {}).keys()))
                sel_sub = st.selectbox("Sous-catégorie", [""] + subs, key=skey("add","sub"))
                if sel_sub:
                    visa_final, opts_dict, _ = build_visa_option_selector(visa_map, sel_cat, sel_sub, keyprefix=skey("add","opts"))

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", 0.0, step=50.0, format="%.2f", key=skey("add","hon"))
            other = f2.number_input("Autres frais (US $)", 0.0, step=20.0, format="%.2f", key=skey("add","autres"))
            comment = st.text_area("Commentaires / détails (autres frais)", "", key=skey("add","comm"))

            st.markdown("#### 📌 Statuts initiaux")
            s1, s2, s3, s4, s5 = st.columns(5)
            sent = s1.checkbox("Dossier envoyé", key=skey("add","sent"))
            acc = s2.checkbox("Dossier accepté", key=skey("add","acc"))
            ref = s3.checkbox("Dossier refusé", key=skey("add","ref"))
            ann = s4.checkbox("Dossier annulé", key=skey("add","ann"))
            rfe = s5.checkbox("RFE", key=skey("add","rfe"))

            save_add = st.button("💾 Enregistrer le client", key=skey("add","btn"))
            if save_add:
                if not nom:
                    st.warning("Veuillez saisir un nom.")
                elif not sel_cat or not sel_sub:
                    st.warning("Choisissez la catégorie et la sous-catégorie.")
                else:
                    total = honor + other
                    reste = total
                    dossier_n = _next_dossier(df_live)
                    new_row = {
                        DOSSIER_COL: dossier_n,
                        "Nom": nom,
                        "Date": dt,
                        "Mois": mois,
                        "Categorie": sel_cat,
                        "Sous-categorie": sel_sub,
                        "Visa": visa_final if visa_final else sel_sub,
                        HONO: honor,
                        AUTRE: other,
                        TOTAL: total,
                        "Payé": 0.0,
                        "Reste": reste,
                        "Commentaires": comment,
                        "Dossier envoyé": int(sent),
                        "Dossier accepté": int(acc),
                        "Dossier refusé": int(ref),
                        "Dossier annulé": int(ann),
                        "RFE": int(rfe),
                        "Paiements": []
                    }
                    df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                    _write_clients(df_new, clients_path)
                    st.success("Client ajouté avec succès.")
                    st.cache_data.clear()
                    st.rerun()



# --- MODIFIER CLIENT ---
        if op == "Modifier":
            st.markdown("### ✏️ Modifier un client")
            if df_live.empty:
                st.info("Aucun client à modifier.")
            else:
                # Choix du client par Nom ou ID
                ncol1, ncol2 = st.columns(2)
                names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
                ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
                target_name = ncol1.selectbox("Nom", [""] + names, key=skey("mod","sel_nom"))
                target_id   = ncol2.selectbox("ID_Client", [""] + ids, key=skey("mod","sel_id"))

                mask = None
                if target_id:
                    mask = (df_live["ID_Client"].astype(str) == target_id)
                elif target_name:
                    mask = (df_live["Nom"].astype(str) == target_name)

                if mask is None or not mask.any():
                    st.stop()

                idx = df_live[mask].index[0]
                row = df_live.loc[idx].copy()

                # Infos principales
                d1, d2, d3, d4 = st.columns([1.2,1,1,1])
                d1.text_input("Nom", value=_safe_str(row.get("Nom","")), key=skey("mod","nom"))
                # ID_Client (affiché en lecture seule si manquant on en (re)générera à l'enregistrement)
                curr_id = _safe_str(row.get("ID_Client",""))
                d2.text_input("ID_Client", value=curr_id, key=skey("mod","id"), disabled=True)
                # date & mois
                dval = _date_for_widget(row.get("Date"))
                if dval is None:
                    dval = date.today()
                dt   = d2.date_input("Date de création", value=dval, key=skey("mod","date"))
                mois_str = _safe_str(row.get("Mois","")).zfill(2)
                try:
                    mois_idx = max(0, int(mois_str) - 1)
                except Exception:
                    mois_idx = date.today().month - 1
                d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=mois_idx, key=skey("mod","mois"))

                # Montants + commentaire
                f1, f2 = st.columns(2)
                honor_v = float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0])
                other_v = float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0])
                f1.number_input(HONO, min_value=0.0, value=honor_v, step=50.0, format="%.2f", key=skey("mod","hon"))
                f2.number_input(AUTRE, min_value=0.0, value=other_v, step=20.0, format="%.2f", key=skey("mod","autres"))
                st.text_area("Commentaire (autres frais / notes)", _safe_str(row.get("Commentaire", row.get("Commentaires",""))), key=skey("mod","comm"))

                # Catégorie / Sous-catégorie / Options
                st.markdown("#### 🎯 Visa — Catégorie, Sous-catégorie et options")
                mc1, mc2 = st.columns(2)
                cats = sorted(list(visa_map.keys()))
                preset_cat = _safe_str(row.get("Categorie",""))
                mc1.selectbox("Catégorie", [""] + cats,
                              index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                              key=skey("mod","cat"))
                sel_cat = st.session_state[skey("mod","cat")]
                subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
                preset_sub = _safe_str(row.get("Sous-categorie",""))
                mc2.selectbox("Sous-catégorie", [""] + subs,
                              index=(subs.index(preset_sub)+1 if preset_sub in subs else 0),
                              key=skey("mod","sub"))
                sel_sub = st.session_state[skey("mod","sub")]

                # Options pré-enregistrées
                preset_opts = row.get("Options", {})
                if not isinstance(preset_opts, dict):
                    try:
                        preset_opts = json.loads(_safe_str(preset_opts) or "{}")
                        if not isinstance(preset_opts, dict):
                            preset_opts = {}
                    except Exception:
                        preset_opts = {}

                visa_final = _safe_str(row.get("Visa",""))
                opts_dict  = {"exclusive": None, "options": []}
                info_msg   = ""
                if sel_cat and sel_sub:
                    visa_final, opts_dict, info_msg = build_visa_option_selector(
                        visa_map, sel_cat, sel_sub, keyprefix=skey("mod","opts"), preselected=preset_opts
                    )
                    if info_msg:
                        st.info(info_msg)
                else:
                    st.warning("Choisissez d’abord la catégorie puis la sous-catégorie.")

                # Statuts + dates
                st.markdown("#### 📌 Statuts & dates")
                s1, s2, s3, s4, s5 = st.columns(5)

                def _int01(x): 
                    try: return int(x) if pd.notna(x) else 0
                    except Exception: return 0

                envoye = _int01(row.get("Dossier envoyé",0)) == 1
                accepte = _int01(row.get("Dossier accepté",0)) == 1
                refuse  = _int01(row.get("Dossier refusé",0)) == 1
                annule  = _int01(row.get("Dossier annulé",0)) == 1
                rfe     = _int01(row.get("RFE",0)) == 1

                s1.checkbox("Dossier envoyé", value=envoye, key=skey("mod","sent"))
                s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=skey("mod","sentd"))
                s2.checkbox("Dossier accepté", value=accepte, key=skey("mod","acc"))
                s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=skey("mod","accd"))
                s3.checkbox("Dossier refusé", value=refuse, key=skey("mod","ref"))
                s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=skey("mod","refd"))
                s4.checkbox("Dossier annulé", value=annule, key=skey("mod","ann"))
                s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=skey("mod","annd"))
                s5.checkbox("RFE", value=rfe, key=skey("mod","rfe"))

                # Enregistrer
                if st.button("💾 Enregistrer les modifications", key=skey("mod","save")):
                    # Récup des champs
                    new_nom   = st.session_state[skey("mod","nom")].strip()
                    new_id    = curr_id or _make_client_id(new_nom, dt)
                    new_mois  = st.session_state[skey("mod","mois")]
                    new_hon   = float(st.session_state[skey("mod","hon")])
                    new_autre = float(st.session_state[skey("mod","autres")])
                    new_comm  = st.session_state[skey("mod","comm")]
                    new_cat   = st.session_state[skey("mod","cat")]
                    new_sub   = st.session_state[skey("mod","sub")]
                    if not new_nom:
                        st.warning("Le nom est requis.")
                        st.stop()
                    if not new_cat or not new_sub:
                        st.warning("Catégorie et sous-catégorie sont requises.")
                        st.stop()

                    total = float(new_hon + new_autre)
                    paye_old = float(_safe_num_series(pd.DataFrame([row]), "Payé").iloc[0])
                    reste = max(0.0, total - paye_old)

                    # Mise à jour dataframe
                    df_live.at[idx, "Nom"] = new_nom
                    df_live.at[idx, "ID_Client"] = new_id
                    df_live.at[idx, "Date"] = dt
                    df_live.at[idx, "Mois"] = new_mois
                    df_live.at[idx, "Categorie"] = new_cat
                    df_live.at[idx, "Sous-categorie"] = new_sub
                    df_live.at[idx, "Visa"] = (visa_final if visa_final else new_sub)
                    df_live.at[idx, HONO] = new_hon
                    df_live.at[idx, AUTRE] = new_autre
                    df_live.at[idx, TOTAL] = total
                    df_live.at[idx, "Reste"] = reste
                    df_live.at[idx, "Commentaire"] = new_comm
                    df_live.at[idx, "Options"] = opts_dict

                    df_live.at[idx, "Dossier envoyé"] = 1 if st.session_state[skey("mod","sent")] else 0
                    df_live.at[idx, "Date d'envoi"] = st.session_state[skey("mod","sentd")]
                    df_live.at[idx, "Dossier accepté"] = 1 if st.session_state[skey("mod","acc")] else 0
                    df_live.at[idx, "Date d'acceptation"] = st.session_state[skey("mod","accd")]
                    df_live.at[idx, "Dossier refusé"] = 1 if st.session_state[skey("mod","ref")] else 0
                    df_live.at[idx, "Date de refus"] = st.session_state[skey("mod","refd")]
                    df_live.at[idx, "Dossier annulé"] = 1 if st.session_state[skey("mod","ann")] else 0
                    df_live.at[idx, "Date d'annulation"] = st.session_state[skey("mod","annd")]
                    df_live.at[idx, "RFE"] = 1 if st.session_state[skey("mod","rfe")] else 0

                    _write_clients(df_live, clients_path)
                    st.success("Modifications enregistrées.")
                    st.cache_data.clear()
                    st.rerun()

        # --- SUPPRIMER CLIENT ---
        if op == "Supprimer":
            st.markdown("### 🗑️ Supprimer un client")
            if df_live.empty:
                st.info("Aucun client à supprimer.")
            else:
                sc1, sc2 = st.columns(2)
                names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
                ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
                target_name = sc1.selectbox("Nom", [""] + names, key=skey("del","nom"))
                target_id   = sc2.selectbox("ID_Client", [""] + ids, key=skey("del","id"))

                mask = None
                if target_id:
                    mask = (df_live["ID_Client"].astype(str) == target_id)
                elif target_name:
                    mask = (df_live["Nom"].astype(str) == target_name)

                if mask is not None and mask.any():
                    row = df_live[mask].iloc[0]
                    st.write({"Dossier N": row.get(DOSSIER_COL,""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
                    if st.button("❗ Confirmer la suppression", key=skey("del","go")):
                        df_new = df_live[~mask].copy()
                        _write_clients(df_new, clients_path)
                        st.success("Client supprimé.")
                        st.cache_data.clear()
                        st.rerun()


# ================================
# 📄 ONGLET : Visa (aperçu & test)
# ================================
with tabs[4]:
    st.subheader("📄 Visa — aperçu du référentiel")

    if df_visa_raw.empty:
        st.info("Aucun fichier Visa chargé.")
    else:
        st.markdown("#### Fichier Visa brut")
        st.dataframe(df_visa_raw, use_container_width=True, hide_index=True, key=skey("visa","raw"))

        st.markdown("#### Sélecteurs de test")
        cats = sorted(list(visa_map.keys()))
        t1, t2 = st.columns(2)
        t1.selectbox("Catégorie", [""] + cats, key=skey("vt","cat"))
        sel_cat = st.session_state[skey("vt","cat")]
        subs = sorted(list(visa_map.get(sel_cat, {}).keys())) if sel_cat else []
        t2.selectbox("Sous-catégorie", [""] + subs, key=skey("vt","sub"))
        sel_sub = st.session_state[skey("vt","sub")]

        if sel_cat and sel_sub:
            visa_final, opts_dict, info_msg = build_visa_option_selector(
                visa_map, sel_cat, sel_sub, keyprefix=skey("vt","opts"), preselected={}
            )
            st.success(f"Visa résultant : **{visa_final or sel_sub}**")
            st.json(opts_dict, expanded=False)


# ================================
# 💾 ONGLET : Export
# ================================
with tabs[5]:
    st.subheader("💾 Export")
    df_all = normalize_clients(df_clients_raw)

    cexp1, cexp2 = st.columns([1.2, 2])
    # Export "Clients.xlsx" (feuille unique)
    clients_bytes = write_df_to_bytes(df_all, sheet_name=SHEET_CLIENTS)
    cexp1.download_button(
        "⬇️ Télécharger Clients.xlsx",
        data=clients_bytes,
        file_name="Clients.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=skey("exp","clients")
    )

    # Export ZIP (Clients + Visa si dispo)
    if not df_visa_raw.empty:
        try:
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                # Clients
                zf.writestr("Clients.xlsx", clients_bytes)
                # Visa
                with BytesIO() as vb:
                    with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                        df_visa_raw.to_excel(wr, index=False, sheet_name=SHEET_VISA)
                    zf.writestr("Visa.xlsx", vb.getvalue())
            cexp2.download_button(
                "⬇️ Télécharger Export (ZIP)",
                data=buf.getvalue(),
                file_name="Export_Visa_Manager.zip",
                mime="application/zip",
                key=skey("exp","zip")
            )
        except Exception as e:
            st.error("Erreur export ZIP : " + _safe_str(e))
    else:
        st.caption("Chargez aussi le fichier Visa pour activer l’export ZIP (Clients + Visa).")


# ================================
# 🔧 Sélecteur d’options Visa (radio + cases) — utilitaire UI
# ================================
def build_visa_option_selector(visa_map: Dict[str, Dict[str, Dict[str, Any]]],
                               cat: str, sub: str,
                               keyprefix: str,
                               preselected: Dict[str, Any] | None = None
                               ) -> Tuple[str, Dict[str, Any], str]:
    """
    Rend dynamiquement :
    - un choix exclusif (radio) si la ligne comporte COS/EOS cochés dans le fichier Visa
    - des cases à cocher pour les autres colonnes (valeur = 1)
    Retourne (visa_final, opts_dict, info_message)
    """
    info = ""
    base = sub
    preset_excl = None
    preset_others: List[str] = []

    if isinstance(preselected, dict):
        preset_excl = preselected.get("exclusive")
        # conserver str simple si dict a mis une liste
        if isinstance(preset_excl, list) and preset_excl:
            preset_excl = preset_excl[0]
        preset_others = preselected.get("options", []) or []
        if not isinstance(preset_others, list):
            preset_others = []

    block = visa_map.get(cat, {}).get(sub, {"exclusive": None, "options": []})
    excl_candidates = block.get("exclusive") or None
    other_candidates: List[str] = block.get("options") or []

    chosen_excl = None
    if excl_candidates:
        # radio exclusif
        radio_opts = [""] + excl_candidates
        default_idx = 0
        if preset_excl and preset_excl in excl_candidates:
            default_idx = excl_candidates.index(preset_excl) + 1
        chosen_excl = st.radio(
            "Option exclusive",
            radio_opts,
            index=default_idx,
            horizontal=True,
            key=keyprefix + "_excl",
        ) or None

    chosen_others: List[str] = []
    if other_candidates:
        st.markdown("Options complémentaires")
        oc1, oc2, oc3 = st.columns(3)
        cols = [oc1, oc2, oc3]
        for i, opt in enumerate(other_candidates):
            preset = opt in preset_others
            if cols[i % 3].checkbox(opt, value=preset, key=f"{keyprefix}_opt_{i}"):
                chosen_others.append(opt)

    visa_final = base
    if chosen_excl:
        visa_final = f"{base} {chosen_excl}"

    opts_dict = {
        "exclusive": chosen_excl,
        "options": chosen_others
    }

    if excl_candidates and "COS" in excl_candidates and "EOS" in excl_candidates:
        info = "Cette sous-catégorie supporte un choix exclusif **COS/EOS**."

    return visa_final, opts_dict, info