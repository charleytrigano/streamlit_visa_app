# ================================
# 🛂 Visa Manager — PARTIE 1/4
# Imports, constantes, utilitaires, persistance fichiers
# ================================

from __future__ import annotations

import json
import re
from io import BytesIO
from pathlib import Path
from datetime import date, datetime
from typing import Any, Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st

# ----------------
# Constantes
# ----------------
SHEET_CLIENTS = "Clients"
SHEET_VISA    = "Visa"

DOSSIER_COL   = "Dossier N"
HONO          = "Montant honoraires (US $)"
AUTRE         = "Autres frais (US $)"
TOTAL         = "Total (US $)"

# répertoire de persistance des derniers fichiers chargés
LAST_DIR = Path("./last_files")
LAST_DIR.mkdir(parents=True, exist_ok=True)

# identifiant unique pour les clés de widgets
SID = st.session_state.get("_sid_", None)
if SID is None:
    SID = datetime.now().strftime("%Y%m%d%H%M%S%f")[-8:]
    st.session_state["_sid_"] = SID

def skey(*parts: str) -> str:
    """Génère une clé unique et stable pour Streamlit widgets."""
    return "key_" + SID + "_" + "_".join(str(p) for p in parts)

# ----------------
# Helpers sûrs
# ----------------
def _safe_str(x: Any) -> str:
    if x is None:
        return ""
    try:
        return str(x)
    except Exception:
        return ""

def _to_iso_date(d: Any) -> str:
    if isinstance(d, date) and not isinstance(d, datetime):
        return d.strftime("%Y-%m-%d")
    try:
        d2 = pd.to_datetime(d, errors="coerce")
        if pd.isna(d2):
            return ""
        return d2.date().strftime("%Y-%m-%d")
    except Exception:
        return ""

def _date_for_widget(v: Any) -> Optional[date]:
    """Renvoie un `date` (sans heure) ou None, sans planter les widgets."""
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    try:
        d2 = pd.to_datetime(v, errors="coerce")
        if pd.isna(d2):
            return None
        return d2.date()
    except Exception:
        return None

def _safe_num_series(df: pd.DataFrame, col: str) -> pd.Series:
    if col not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index, dtype=float)
    s = df[col]
    if pd.api.types.is_numeric_dtype(s):
        return s.fillna(0.0).astype(float)
    s = s.astype(str)
    s = s.str.replace(r"[^\d,.\-]", "", regex=True)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0).astype(float)

def _fmt_money(x: float) -> str:
    try:
        return f"${x:,.2f}"
    except Exception:
        return f"${x}"

# ----------------
# Persistance : derniers fichiers
# ----------------
def _save_last(kind: str, content_bytes: bytes) -> None:
    """kind ∈ {'clients','visa','both'}"""
    try:
        p = LAST_DIR / f"last_{kind}.bin"
        with open(p, "wb") as f:
            f.write(content_bytes or b"")
    except Exception:
        pass

def _load_last(kind: str) -> Optional[bytes]:
    try:
        p = LAST_DIR / f"last_{kind}.bin"
        if p.exists() and p.is_file():
            return p.read_bytes()
    except Exception:
        pass
    return None

# ----------------
# Lecture / Écriture Excel
# ----------------
@st.cache_data(show_spinner=False)
def read_sheet(xlsx_path_or_bytes: Any, sheet_name: str) -> pd.DataFrame:
    try:
        if isinstance(xlsx_path_or_bytes, (str, Path)):
            return pd.read_excel(xlsx_path_or_bytes, sheet_name=sheet_name)
        else:
            return pd.read_excel(BytesIO(xlsx_path_or_bytes), sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()

def normalize_clients(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    # colonnes minimales
    for c in [
        DOSSIER_COL, "ID_Client", "Nom", "Date", "Mois",
        "Categorie", "Sous-categorie", "Visa",
        HONO, AUTRE, TOTAL, "Payé", "Reste",
        "Paiements", "Options", "Commentaire",
        "Dossier envoyé", "Date d'envoi",
        "Dossier accepté", "Date d'acceptation",
        "Dossier refusé", "Date de refus",
        "Dossier annulé", "Date d'annulation",
        "RFE",
    ]:
        if c not in df.columns:
            df[c] = None

    # numéros
    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
        df[c] = _safe_num_series(df, c)

    # dates dérivées
    # _Année_, _MoisNum_ (pour tri, analyses)
    dt = pd.to_datetime(df["Date"], errors="coerce")
    df["_Année_"] = dt.dt.year
    df["_MoisNum_"] = dt.dt.month

    # Mois (MM — affichage)
    def _month_text(val, dflt):
        s = _safe_str(val).strip()
        if s.isdigit() and 1 <= int(s) <= 12:
            return f"{int(s):02d}"
        if isinstance(dflt, (date, datetime)):
            return f"{int(dflt.month):02d}"
        return "01"

    df["Mois"] = [
        _month_text(df.at[i, "Mois"], dt.iloc[i] if i < len(dt) else None)
        for i in range(len(df))
    ]

    # champs booléens / int
    for c in ["Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)

    return df

def write_clients_to_bytes(df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as wr:
        df.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
    return out.getvalue()

def write_two_sheets_to_bytes(df_clients: pd.DataFrame, df_visa: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as wr:
        df_clients.to_excel(wr, sheet_name=SHEET_CLIENTS, index=False)
        df_visa.to_excel(wr, sheet_name=SHEET_VISA, index=False)
    return out.getvalue()

# ----------------
# Visa map (Catégorie → Sous-catégorie → options)
#   On détecte toutes les colonnes != {Categorie, Sous-categorie};
#   si la cellule == 1, on ajoute le nom de colonne comme option dispo.
#   Si on trouve des colonnes "COS"/"EOS", on les marque comme "exclusives".
# ----------------
def build_visa_map(df_visa: pd.DataFrame) -> Dict[str, Dict[str, Dict[str, Any]]]:
    vm: Dict[str, Dict[str, Dict[str, Any]]] = {}
    if df_visa.empty:
        return vm

    cols = [c for c in df_visa.columns if c not in ("Categorie", "Sous-categorie")]
    for _, row in df_visa.iterrows():
        cat = _safe_str(row.get("Categorie", "")).strip()
        sub = _safe_str(row.get("Sous-categorie", "")).strip()
        if not cat or not sub:
            continue
        opt_cols = []
        for c in cols:
            v = row.get(c, 0)
            try:
                ok = float(v) == 1.0
            except Exception:
                ok = _safe_str(v).strip() == "1"
            if ok:
                opt_cols.append(c)

        exclusive_set = None
        if "COS" in opt_cols or "EOS" in opt_cols:
            exclusive_set = ["COS", "EOS"]

        vm.setdefault(cat, {})
        vm[cat].setdefault(sub, {
            "exclusive": exclusive_set,   # liste de labels exclusifs (radio), sinon None
            "options": [o for o in opt_cols if o not in (exclusive_set or [])]
        })
    return vm

# ----------------
# ID client & Dossier N
# ----------------
def _make_client_id(base_name: str, d: Any) -> str:
    base = re.sub(r"[^A-Za-z0-9\-]+", "-", _safe_str(base_name)).strip("-") or "Client"
    d2 = _date_for_widget(d) or date.today()
    return f"{base}-{d2.strftime('%Y%m%d')}"

def _next_dossier(df: pd.DataFrame, start: int = 13057) -> int:
    try:
        nums = pd.to_numeric(df.get(DOSSIER_COL, pd.Series(dtype=int)), errors="coerce")
        mx = int(nums.dropna().max()) if not nums.dropna().empty else (start - 1)
        return mx + 1
    except Exception:
        return start



# ================================
# 🛂 Visa Manager — PARTIE 2/4
# Chargement fichiers, lecture DF, visa_map, onglets
# ================================

st.set_page_config(page_title="Visa Manager", layout="wide")
st.title("🛂 Visa Manager")

# -------------------------------------------------
# 1) Raccourcis lecture/écriture dans la session
# -------------------------------------------------
def _read_clients() -> pd.DataFrame:
    """Lit Clients depuis la session (bytes) si présents, sinon DataFrame vide."""
    b = st.session_state.get("clients_bytes")
    if not b:
        return pd.DataFrame()
    try:
        df = read_sheet(b, SHEET_CLIENTS)
        return normalize_clients(df.copy())
    except Exception:
        return pd.DataFrame()

def _read_visa_raw() -> pd.DataFrame:
    """Lit Visa depuis la session (bytes) si présents, sinon DataFrame vide."""
    b = st.session_state.get("visa_bytes")
    if not b:
        return pd.DataFrame()
    try:
        return read_sheet(b, SHEET_VISA)
    except Exception:
        # si ce n'est pas un classeur 2 onglets mais un xlsx simple nommé Visa
        try:
            return pd.read_excel(BytesIO(b))
        except Exception:
            return pd.DataFrame()

def _write_clients(df_new: pd.DataFrame) -> None:
    """Écrit Clients dans la session + persistance disque."""
    bytes_out = write_clients_to_bytes(df_new)
    st.session_state["clients_bytes"] = bytes_out
    _save_last("clients", bytes_out)

# -------------------------------------------------
# 2) Auto-recharge des derniers fichiers persistés
# -------------------------------------------------
if not st.session_state.get("clients_bytes"):
    lastC = _load_last("clients")
    if lastC:
        st.session_state["clients_bytes"] = lastC

if not st.session_state.get("visa_bytes"):
    lastV = _load_last("visa")
    if lastV:
        st.session_state["visa_bytes"] = lastV

# -------------------------------------------------
# 3) Zone chargement — 2 modes : (A) deux fichiers, (B) un seul fichier à 2 onglets
# -------------------------------------------------
st.markdown("## 📂 Fichiers")
mode = st.radio("Mode de chargement", ["Deux fichiers (Clients & Visa)", "Un seul fichier (2 onglets)"], horizontal=True, key=skey("load","mode"))

c_up1, c_up2 = st.columns(2)

if mode.startswith("Deux fichiers"):
    with c_up1:
        upC = st.file_uploader("Clients (xlsx)", type=["xlsx"], key=skey("up","clients"))
        if upC is not None:
            b = upC.read()
            st.session_state["clients_bytes"] = b
            _save_last("clients", b)
            st.success("✅ Clients chargé.")
    with c_up2:
        upV = st.file_uploader("Visa (xlsx)", type=["xlsx"], key=skey("up","visa"))
        if upV is not None:
            b = upV.read()
            st.session_state["visa_bytes"] = b
            _save_last("visa", b)
            st.success("✅ Visa chargé.")
else:
    upBoth = st.file_uploader("Classeur unique (onglets 'Clients' et 'Visa')", type=["xlsx"], key=skey("up","both"))
    if upBoth is not None:
        b = upBoth.read()
        try:
            # on vérifie la présence des 2 onglets
            x = pd.ExcelFile(BytesIO(b))
            sheets = [s.lower() for s in x.sheet_names]
            if SHEET_CLIENTS.lower() in sheets and SHEET_VISA.lower() in sheets:
                # sépare en 2 buffers pour rester homogène avec le reste de l’app
                dfC = pd.read_excel(BytesIO(b), sheet_name=SHEET_CLIENTS)
                dfV = pd.read_excel(BytesIO(b), sheet_name=SHEET_VISA)

                cb = write_clients_to_bytes(dfC)
                st.session_state["clients_bytes"] = cb
                _save_last("clients", cb)

                vb = BytesIO()
                with pd.ExcelWriter(vb, engine="openpyxl") as wr:
                    dfV.to_excel(wr, sheet_name=SHEET_VISA, index=False)
                st.session_state["visa_bytes"] = vb.getvalue()
                _save_last("visa", vb.getvalue())

                st.success("✅ Classeur chargé (Clients & Visa).")
            else:
                st.error("Le classeur doit contenir 2 onglets : 'Clients' et 'Visa'.")
        except Exception as e:
            st.error(f"Lecture impossible : {e}")

# Raccourcis recharger derniers
col_lastC, col_lastV = st.columns(2)
with col_lastC:
    if st.button("↩️ Recharger dernier Clients", key=skey("last","clients")):
        last = _load_last("clients")
        if last:
            st.session_state["clients_bytes"] = last
            st.success("Dernier Clients rechargé.")
        else:
            st.info("Aucun Clients mémorisé.")

with col_lastV:
    if st.button("↩️ Recharger dernier Visa", key=skey("last","visa")):
        last = _load_last("visa")
        if last:
            st.session_state["visa_bytes"] = last
            st.success("Dernier Visa rechargé.")
        else:
            st.info("Aucun Visa mémorisé.")

st.markdown("---")

# -------------------------------------------------
# 4) Lecture DataFrames courant + normalisation
# -------------------------------------------------
df_clients_raw = _read_clients()
df_visa_raw    = _read_visa_raw()

if not df_clients_raw.empty:
    df_all = normalize_clients(df_clients_raw.copy())
else:
    df_all = pd.DataFrame()

# -------------------------------------------------
# 5) Construction de la carte des visas
#    Catégorie -> Sous-catégorie -> {"exclusive":[...], "options":[...]}
# -------------------------------------------------
visa_map: Dict[str, Dict[str, Dict[str, Any]]] = {}
try:
    if not df_visa_raw.empty:
        # Tenter de tolérer quelques variations de noms de colonnes
        dfv = df_visa_raw.copy()
        # Renommer si besoin (ex: "Catégorie" -> "Categorie")
        rename_map = {}
        for c in dfv.columns:
            if c.lower().replace("é","e") == "categorie":
                rename_map[c] = "Categorie"
            if c.lower().startswith("sous") and "categorie" in c.lower().replace("é","e"):
                rename_map[c] = "Sous-categorie"
        if rename_map:
            dfv = dfv.rename(columns=rename_map)
        visa_map = build_visa_map(dfv)
except Exception as e:
    st.warning(f"Impossible de construire la carte Visa : {e}")

# -------------------------------------------------
# 6) Création des onglets
# -------------------------------------------------
tabs = st.tabs([
    "📊 Dashboard",
    "📈 Analyses",
    "🏦 Escrow",
    "👤 Compte client",
    "🧾 Clients",
    "📄 Visa (aperçu)",
])



# ================================
# 🧭 PARTIE 3/4 — Dashboard & Analyses
# ================================

def _kpi_badge(label: str, value: str):
    st.markdown(
        f"""
        <div style="display:flex;flex-direction:column;gap:2px;padding:8px 10px;border:1px solid #e5e7eb;border-radius:10px;">
            <div style="font-size:11px;color:#6b7280;">{label}</div>
            <div style="font-size:16px;font-weight:700;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def _apply_filters(base: pd.DataFrame, fy, fm, fc, fs, fv) -> pd.DataFrame:
    df = base.copy()
    if fy:
        df = df[df["_Année_"].isin(fy)]
    if fm:
        df = df[df["Mois"].astype(str).isin([f"{int(m):02d}" for m in fm])]
    if fc:
        df = df[df["Categorie"].astype(str).isin(fc)]
    if fs:
        df = df[df["Sous-categorie"].astype(str).isin(fs)]
    if fv:
        df = df[df["Visa"].astype(str).isin(fv)]
    return df

with tabs[0]:
    st.subheader("📊 Dashboard")

    if df_all.empty:
        st.info("Charge un classeur Clients/Visa pour démarrer.")
    else:
        years = sorted([int(x) for x in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        cats  = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subs  = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visas = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        f1, f2, f3, f4, f5 = st.columns([1.2,1.2,1.6,1.6,1.6])
        fy = f1.multiselect("Année", years, default=[], key=skey("dash","years"))
        fm = f2.multiselect("Mois (MM)", monthsA, default=[], key=skey("dash","months"))
        fc = f3.multiselect("Catégorie", cats, default=[], key=skey("dash","cats"))
        fs = f4.multiselect("Sous-catégorie", subs, default=[], key=skey("dash","subs"))
        fv = f5.multiselect("Visa", visas, default=[], key=skey("dash","visas"))

        view = _apply_filters(df_all, fy, fm, fc, fs, fv)

        kA, kB, kC, kD, kE = st.columns([1,1,1,1,1])
        _kpi_badge("Dossiers", f"{len(view)}")
        _kpi_badge("Honoraires", _fmt_money(float(_safe_num_series(view, HONO).sum())))
        _kpi_badge("Autres frais", _fmt_money(float(_safe_num_series(view, AUTRE).sum())))
        _kpi_badge("Payé", _fmt_money(float(_safe_num_series(view, "Payé").sum())))
        _kpi_badge("Reste", _fmt_money(float(_safe_num_series(view, "Reste").sum())))

        st.divider()

        cL, cR = st.columns([1.2, 1.0])

        with cL:
            st.markdown("#### 📦 Dossiers par catégorie (% & volume)")
            if not view.empty and "Categorie" in view.columns:
                vc = view["Categorie"].value_counts().rename_axis("Categorie").reset_index(name="Nb")
                tot = max(int(vc["Nb"].sum()), 1)
                vc["%"] = (vc["Nb"] * 100.0 / tot).round(1)
                st.dataframe(vc, use_container_width=True, hide_index=True, key=skey("dash","tab_cat"))
                st.bar_chart(vc.set_index("Categorie")["Nb"])
            else:
                st.caption("Aucune donnée (catégorie).")

            st.markdown("#### 🧾 Dossiers par sous-catégorie (% & volume)")
            if not view.empty and "Sous-categorie" in view.columns:
                vs = view["Sous-categorie"].value_counts().rename_axis("Sous-categorie").reset_index(name="Nb")
                tot2 = max(int(vs["Nb"].sum()), 1)
                vs["%"] = (vs["Nb"] * 100.0 / tot2).round(1)
                st.dataframe(vs, use_container_width=True, hide_index=True, key=skey("dash","tab_sub"))
                st.bar_chart(vs.set_index("Sous-categorie")["Nb"])
            else:
                st.caption("Aucune donnée (sous-catégorie).")

        with cR:
            st.markdown("#### 💵 Honoraires par mois")
            if not view.empty:
                t = view.copy()
                t["_periode_"] = t["_Année_"].astype(str) + "-" + t["Mois"].astype(str).str.zfill(2)
                g = t.groupby("_periode_", as_index=False)[HONO].sum().sort_values("_periode_")
                if not g.empty:
                    st.line_chart(g.set_index("_periode_")[HONO])
                else:
                    st.caption("Aucune donnée.")
            else:
                st.caption("Aucune donnée.")

        st.markdown("#### 📋 Détails (tri par année/mois/catégorie)")
        if not view.empty:
            detail = view.copy()
            for cnum in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                if cnum in detail.columns:
                    detail[cnum] = _safe_num_series(detail, cnum).map(_fmt_money)

            if "Date" in detail.columns:
                try:
                    detail["Date"] = pd.to_datetime(detail["Date"], errors="coerce").dt.date.astype(str)
                except Exception:
                    detail["Date"] = detail["Date"].astype(str)

            cols_show = [c for c in [
                DOSSIER_COL, "ID_Client", "Nom", "Categorie", "Sous-categorie", "Visa",
                "Date", "Mois", HONO, AUTRE, TOTAL, "Payé", "Reste",
                "Dossier envoyé", "Dossier accepté", "Dossier refusé", "Dossier annulé", "RFE"
            ] if c in detail.columns]

            sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in detail.columns]
            detail_sorted = detail.sort_values(by=sort_keys) if sort_keys else detail
            st.dataframe(detail_sorted[cols_show].reset_index(drop=True), use_container_width=True, key=skey("dash","detail"))
        else:
            st.caption("Aucun résultat avec les filtres actuels.")


with tabs[1]:
    st.subheader("📈 Analyses")

    if df_all.empty:
        st.info("Charge des données pour activer les analyses.")
    else:
        yearsA  = sorted([int(x) for x in pd.to_numeric(df_all["_Année_"], errors="coerce").dropna().unique().tolist()])
        monthsA = [f"{m:02d}" for m in range(1, 13)]
        catsA   = sorted(df_all["Categorie"].dropna().astype(str).unique().tolist()) if "Categorie" in df_all.columns else []
        subsA   = sorted(df_all["Sous-categorie"].dropna().astype(str).unique().tolist()) if "Sous-categorie" in df_all.columns else []
        visasA  = sorted(df_all["Visa"].dropna().astype(str).unique().tolist()) if "Visa" in df_all.columns else []

        a1, a2, a3, a4, a5 = st.columns([1,1,1.3,1.3,1.3])
        fy = a1.multiselect("Année", yearsA, default=[], key=skey("ana","years"))
        fm = a2.multiselect("Mois (MM)", monthsA, default=[], key=skey("ana","months"))
        fc = a3.multiselect("Catégorie", catsA, default=[], key=skey("ana","cats"))
        fs = a4.multiselect("Sous-catégorie", subsA, default=[], key=skey("ana","subs"))
        fv = a5.multiselect("Visa", visasA, default=[], key=skey("ana","visas"))

        dfA = _apply_filters(df_all, fy, fm, fc, fs, fv)

        k1, k2, k3, k4, k5 = st.columns([1,1,1,1,1])
        _kpi_badge("Dossiers", f"{len(dfA)}")
        _kpi_badge("Honoraires", _fmt_money(float(_safe_num_series(dfA, HONO).sum())))
        _kpi_badge("Autres", _fmt_money(float(_safe_num_series(dfA, AUTRE).sum())))
        _kpi_badge("Payé", _fmt_money(float(_safe_num_series(dfA, "Payé").sum())))
        _kpi_badge("Reste", _fmt_money(float(_safe_num_series(dfA, "Reste").sum())))

        st.divider()

        cL, cR = st.columns([1.2, 1.0])
        with cL:
            st.markdown("#### 📦 Répartition catégories")
            if not dfA.empty and "Categorie" in dfA.columns:
                vc = dfA["Categorie"].value_counts().rename_axis("Categorie").reset_index(name="Nb")
                tot = max(int(vc["Nb"].sum()), 1)
                vc["%"] = (vc["Nb"] * 100.0 / tot).round(1)
                st.dataframe(vc, use_container_width=True, hide_index=True, key=skey("ana","tab_cat"))
                st.bar_chart(vc.set_index("Categorie")["Nb"])
            else:
                st.caption("Aucune donnée.")

            st.markdown("#### 🧾 Répartition sous-catégories")
            if not dfA.empty and "Sous-categorie" in dfA.columns:
                vs = dfA["Sous-categorie"].value_counts().rename_axis("Sous-categorie").reset_index(name="Nb")
                tot2 = max(int(vs["Nb"].sum()), 1)
                vs["%"] = (vs["Nb"] * 100.0 / tot2).round(1)
                st.dataframe(vs, use_container_width=True, hide_index=True, key=skey("ana","tab_sub"))
                st.bar_chart(vs.set_index("Sous-categorie")["Nb"])
            else:
                st.caption("Aucune donnée.")

        with cR:
            st.markdown("#### 💵 Évolution des montants (Honoraires + Autres)")
            if not dfA.empty:
                grp = (
                    dfA.groupby(["_Année_", "Mois"], dropna=True)[[HONO, AUTRE, TOTAL]]
                    .sum()
                    .reset_index()
                )
                if not grp.empty:
                    grp["Periode"] = grp["_Année_"].astype(str) + "-" + grp["Mois"].astype(str).str.zfill(2)
                    st.bar_chart(grp.set_index("Periode")[TOTAL], use_container_width=True)
                    st.line_chart(grp.set_index("Periode")[[HONO, AUTRE]], use_container_width=True)
                else:
                    st.caption("Aucune donnée.")
            else:
                st.caption("Aucune donnée.")

        st.markdown("#### 🧾 Détails des dossiers filtrés")
        det = dfA.copy()
        if not det.empty:
            for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                if c in det.columns:
                    det[c] = _safe_num_series(det, c).map(_fmt_money)
            if "Date" in det.columns:
                try:
                    det["Date"] = pd.to_datetime(det["Date"], errors="coerce").dt.date.astype(str)
                except Exception:
                    det["Date"] = det["Date"].astype(str)

            show_cols = [c for c in [
                DOSSIER_COL,"ID_Client","Nom","Categorie","Sous-categorie","Visa",
                "Date","Mois", HONO, AUTRE, TOTAL, "Payé","Reste",
                "Dossier envoyé","Dossier accepté","Dossier refusé","Dossier annulé","RFE"
            ] if c in det.columns]

            sort_keys = [c for c in ["_Année_", "_MoisNum_", "Categorie", "Nom"] if c in det.columns]
            det_sorted = det.sort_values(by=sort_keys) if sort_keys else det
            st.dataframe(det_sorted[show_cols].reset_index(drop=True), use_container_width=True, key=skey("ana","detail"))
        else:
            st.caption("Aucun résultat.")



# ================================
# ⚙️ PARTIE 4/4 — Escrow, Compte client (paiements),
#                 Clients (CRUD), Visa (aperçu)
# ================================

# --------- Outils "paiements" ---------
def _parse_payments(v: Any) -> List[Dict[str, Any]]:
    if isinstance(v, list):
        return v
    s = _safe_str(v).strip()
    if not s:
        return []
    # tolérer chaînes JSON
    try:
        data = json.loads(s)
        if isinstance(data, list):
            return data
    except Exception:
        pass
    return []

def _payments_total(lst: List[Dict[str, Any]]) -> float:
    tot = 0.0
    for r in lst:
        try:
            tot += float(r.get("amount", 0.0) or 0.0)
        except Exception:
            pass
    return float(tot)

# --------- Rendu des options Visa (cases/choix exclusifs) ---------
def build_visa_option_selector(
    vm: Dict[str, Dict[str, Dict[str, Any]]],
    cat: str, sub: str,
    keyprefix: str,
    preselected: Dict[str, Any] | None = None,
) -> Tuple[str, Dict[str, Any], str]:
    """
    Affiche dynamiquement les options de la sous-catégorie.
    - Si "exclusive" existe (ex: ["COS","EOS"]), on montre un radio exclusif.
    - Les autres options sont des cases à cocher.
    Retourne (visa_label, options_dict, info_msg)
    """
    preselected = preselected or {}
    exclusive_def = None
    options = []

    data = vm.get(cat, {}).get(sub, {})
    exclusive_def = data.get("exclusive") or None
    options = data.get("options", [])

    visa_label = sub
    info = ""
    if exclusive_def:
        st.markdown("##### ⚙️ Options exclusives")
        default_choice = _safe_str(preselected.get("exclusive", "")) if isinstance(preselected, dict) else ""
        if default_choice not in exclusive_def:
            default_choice = ""
        choice = st.radio(
            "Choix exclusif",
            [""] + exclusive_def,
            index=(exclusive_def.index(default_choice) + 1) if default_choice in exclusive_def else 0,
            horizontal=True,
            key=skey(keyprefix, "exo"),
        )
        if choice:
            visa_label = f"{sub} {choice}"
    else:
        choice = None

    sel_opts = []
    if options:
        st.markdown("##### 🔧 Options supplémentaires")
        # pré-sélection
        preset_list = preselected.get("options", []) if isinstance(preselected, dict) else []
        for c in options:
            checked = c in preset_list
            v = st.checkbox(c, value=checked, key=skey(keyprefix, "opt", c))
            if v:
                sel_opts.append(c)

    out = {"exclusive": choice, "options": sel_opts}
    return visa_label, out, info

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
        dfE[TOTAL]   = _safe_num_series(dfE, TOTAL)

        # KPI (compacts)
        k1, k2, k3 = st.columns([1,1,1])
        _kpi_badge("Total (US $)", _fmt_money(float(dfE[TOTAL].sum())))
        _kpi_badge("Payé", _fmt_money(float(dfE["Payé"].sum())))
        _kpi_badge("Reste", _fmt_money(float(dfE["Reste"].sum())))

        st.markdown("#### Par catégorie")
        agg = dfE.groupby("Categorie", as_index=False)[[TOTAL, "Payé", "Reste"]].sum()
        if not agg.empty:
            agg["% Payé"] = (agg["Payé"] / agg[TOTAL]).replace([pd.NA, pd.NaT], 0).fillna(0.0) * 100
            st.dataframe(agg, use_container_width=True, hide_index=True, key=skey("escrow","agg"))
        else:
            st.caption("Aucune donnée à agréger.")

        st.caption("NB : vous pouvez utiliser le suivi des statuts « Dossier envoyé/accepté/refusé/annulé + RFE » dans l’onglet Clients ; les transferts d'escrow se pilotent en pratique via la section paiements (onglet Compte client).")

# ==============================================
# 👤 ONGLET : Compte client — Détail + Paiements
# ==============================================
with tabs[3]:
    st.subheader("👤 Compte client — Détail & paiements")

    df_live = _read_clients()
    if df_live.empty:
        st.info("Aucun client.")
    else:
        # Sélecteurs
        names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
        ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
        s1, s2 = st.columns([1,1])
        sel_name = s1.selectbox("Nom", [""] + names, index=0, key=skey("acct","name"))
        sel_id   = s2.selectbox("ID_Client", [""] + ids, index=0, key=skey("acct","id"))

        mask = None
        if sel_id:
            mask = (df_live["ID_Client"].astype(str) == sel_id)
        elif sel_name:
            mask = (df_live["Nom"].astype(str) == sel_name)

        if mask is None or not mask.any():
            st.stop()

        idx = df_live[mask].index[0]
        row = df_live.loc[idx].copy()

        # Récap
        b1, b2, b3, b4, b5 = st.columns(5)
        b1.metric("Dossier", _safe_str(row.get(DOSSIER_COL,"")))
        b2.metric("Nom", _safe_str(row.get("Nom","")))
        b3.metric("Visa", _safe_str(row.get("Visa","")))
        b4.metric("Honoraires", _fmt_money(float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0])))
        b5.metric("Reste", _fmt_money(float(_safe_num_series(pd.DataFrame([row]), "Reste").iloc[0])))

        st.markdown("#### 🧾 Paiements")
        pay_list = _parse_payments(row.get("Paiements"))

        if pay_list:
            dfP = pd.DataFrame(pay_list)
            if "date" in dfP.columns:
                try:
                    dfP["date"] = pd.to_datetime(dfP["date"], errors="coerce").dt.date.astype(str)
                except Exception:
                    dfP["date"] = dfP["date"].astype(str)
            st.dataframe(dfP[["date","mode","amount","note"]], use_container_width=True, hide_index=True, key=skey("acct","paylist"))
        else:
            st.caption("Aucun paiement enregistré.")

        # Ajout paiement
        st.markdown("##### ➕ Ajouter un paiement")
        p1, p2, p3, p4 = st.columns([1,1,1,2])
        pdate = _date_for_widget(date.today())
        d_in  = p1.date_input("Date", value=pdate, key=skey("acct","pdate"))
        mode  = p2.selectbox("Mode", ["CB","Chèque","Cash","Virement","Venmo"], key=skey("acct","pmode"))
        amt   = p3.number_input("Montant (US $)", min_value=0.0, step=10.0, format="%.2f", key=skey("acct","pamt"))
        note  = p4.text_input("Note", "", key=skey("acct","pnote"))

        if st.button("💾 Enregistrer le paiement", key=skey("acct","savepay")):
            if float(amt) <= 0:
                st.warning("Montant invalide.")
                st.stop()

            pay_row = {
                "date": (d_in or date.today()).strftime("%Y-%m-%d"),
                "mode": mode,
                "amount": float(amt),
                "note": note,
            }
            pay_list.append(pay_row)

            # recalcul Payé / Reste
            honor = float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0])
            other = float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0])
            total = honor + other
            paye  = _payments_total(pay_list)
            reste = max(0.0, total - paye)

            df_live.at[idx, "Paiements"] = pay_list
            df_live.at[idx, "Payé"] = paye
            df_live.at[idx, "Reste"] = reste
            _write_clients(df_live)

            st.success("Paiement ajouté.")
            st.cache_data.clear()
            st.rerun()

# ==============================================
# 🧾 ONGLET : Clients — Ajouter / Modifier / Supprimer
# ==============================================
with tabs[4]:
    st.subheader("🧾 Gestion des clients")

    df_live = _read_clients()
    op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=skey("crud","op"))

    # ---------- AJOUT ----------
    if op == "Ajouter":
        st.markdown("### ➕ Ajouter un client")
        a1, a2, a3 = st.columns(3)
        nom  = a1.text_input("Nom", "", key=skey("add","nom"))
        dval = _date_for_widget(date.today())
        dt   = a2.date_input("Date de création", value=dval, key=skey("add","date"))
        mois = a3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=(dval.month-1 if dval else 0), key=skey("add","mois"))

        st.markdown("#### 🎯 Choix Visa")
        cats = sorted(list(visa_map.keys()))
        cat  = st.selectbox("Catégorie", [""] + cats, index=0, key=skey("add","cat"))
        sub  = ""
        visa_final = ""
        opts_dict = {"exclusive": None, "options": []}
        if cat:
            subs = sorted(list(visa_map.get(cat, {}).keys()))
            sub  = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=skey("add","sub"))
            if sub:
                visa_final, opts_dict, _ = build_visa_option_selector(visa_map, cat, sub, keyprefix="add_opts", preselected={})

        f1, f2 = st.columns(2)
        honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, step=50.0, format="%.2f", key=skey("add","honor"))
        other = f2.number_input("Autres frais (US $)", min_value=0.0, step=20.0, format="%.2f", key=skey("add","other"))
        comment = st.text_area("Commentaire (autres frais / remarques)", key=skey("add","comment"))

        st.markdown("#### 📌 Statuts initiaux")
        s1, s2, s3, s4, s5 = st.columns(5)
        sent = s1.checkbox("Dossier envoyé", key=skey("add","sent"))
        sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(None), key=skey("add","sentd"))
        acc  = s2.checkbox("Dossier accepté", key=skey("add","acc"))
        acc_d= s2.date_input("Date d'acceptation", value=_date_for_widget(None), key=skey("add","accd"))
        ref  = s3.checkbox("Dossier refusé", key=skey("add","ref"))
        ref_d= s3.date_input("Date de refus", value=_date_for_widget(None), key=skey("add","refd"))
        ann  = s4.checkbox("Dossier annulé", key=skey("add","ann"))
        ann_d= s4.date_input("Date d'annulation", value=_date_for_widget(None), key=skey("add","annd"))
        rfe  = s5.checkbox("RFE (avec un autre statut)", key=skey("add","rfe"))
        if rfe and not any([sent, acc, ref, ann]):
            st.warning("RFE doit être combiné à un autre statut.")

        if st.button("💾 Enregistrer le client", key=skey("add","save")):
            if not nom:
                st.warning("Le nom est requis.")
                st.stop()
            if not (cat and sub):
                st.warning("Choisissez Catégorie et Sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            dossier_n = _next_dossier(df_live, start=13057) if not df_live.empty else 13057
            did = _make_client_id(nom, dt)

            new_row = {
                DOSSIER_COL: dossier_n,
                "ID_Client": did,
                "Nom": nom,
                "Date": dt,
                "Mois": f"{int(mois):02d}" if isinstance(mois,(int,str)) else "01",
                "Categorie": cat,
                "Sous-categorie": sub,
                "Visa": (visa_final or sub),
                HONO: float(honor),
                AUTRE: float(other),
                TOTAL: total,
                "Payé": 0.0,
                "Reste": total,
                "Paiements": [],
                "Options": opts_dict,
                "Commentaire": comment,
                "Dossier envoyé": int(bool(sent)),
                "Date d'envoi": sent_d,
                "Dossier accepté": int(bool(acc)),
                "Date d'acceptation": acc_d,
                "Dossier refusé": int(bool(ref)),
                "Date de refus": ref_d,
                "Dossier annulé": int(bool(ann)),
                "Date d'annulation": ann_d,
                "RFE": int(bool(rfe)),
            }

            df_new = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True) if not df_live.empty else pd.DataFrame([new_row])
            _write_clients(normalize_clients(df_new))
            st.success("Client ajouté.")
            st.cache_data.clear()
            st.rerun()

    # ---------- MODIFICATION ----------
    elif op == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if df_live.empty:
            st.info("Aucun client.")
            st.stop()

        names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist()) if "Nom" in df_live.columns else []
        ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist()) if "ID_Client" in df_live.columns else []
        m1, m2 = st.columns(2)
        target_name = m1.selectbox("Nom", [""]+names, index=0, key=skey("mod","nm"))
        target_id   = m2.selectbox("ID_Client", [""]+ids, index=0, key=skey("mod","id"))

        mask = None
        if target_id:
            mask = (df_live["ID_Client"].astype(str) == target_id)
        elif target_name:
            mask = (df_live["Nom"].astype(str) == target_name)

        if mask is None or not mask.any():
            st.stop()

        idx = df_live[mask].index[0]
        row = df_live.loc[idx].copy()

        d1, d2, d3 = st.columns(3)
        nom  = d1.text_input("Nom", _safe_str(row.get("Nom","")), key=skey("mod","nomv"))
        # date sûre
        dval = _date_for_widget(row.get("Date")) or date.today()
        dt   = d2.date_input("Date de création", value=dval, key=skey("mod","date"))
        try:
            mois_default = int(_safe_str(row.get("Mois","01")))
            if not (1 <= mois_default <= 12): mois_default = dval.month
        except Exception:
            mois_default = dval.month
        mois = d3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=mois_default-1, key=skey("mod","mois"))

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
        cat = st.selectbox("Catégorie", [""] + cats,
                           index=(cats.index(preset_cat)+1 if preset_cat in cats else 0),
                           key=skey("mod","cat"))
        sub = _safe_str(row.get("Sous-categorie",""))
        if cat:
            subs = sorted(list(visa_map.get(cat, {}).keys()))
            sub = st.selectbox("Sous-catégorie", [""] + subs,
                               index=(subs.index(sub)+1 if sub in subs else 0),
                               key=skey("mod","sub"))

        visa_final, opts_dict, _ = "", {"exclusive": None, "options": []}, ""
        if cat and sub:
            visa_final, opts_dict, _ = build_visa_option_selector(visa_map, cat, sub, keyprefix="mod_opts", preselected=preset_opts)

        f1, f2 = st.columns(2)
        honor = f1.number_input(HONO, min_value=0.0,
                                value=float(_safe_num_series(pd.DataFrame([row]), HONO).iloc[0]),
                                step=50.0, format="%.2f", key=skey("mod","honor"))
        other = f2.number_input(AUTRE, min_value=0.0,
                                value=float(_safe_num_series(pd.DataFrame([row]), AUTRE).iloc[0]),
                                step=20.0, format="%.2f", key=skey("mod","other"))
        comment = st.text_area("Commentaire (autres frais / remarques)", value=_safe_str(row.get("Commentaire","")), key=skey("mod","comment"))

        st.markdown("#### 📌 Statuts")
        s1, s2, s3, s4, s5 = st.columns(5)
        envoye = s1.checkbox("Dossier envoyé", value=int(row.get("Dossier envoyé",0) or 0)==1, key=skey("mod","sent"))
        sent_d = s1.date_input("Date d'envoi", value=_date_for_widget(row.get("Date d'envoi")), key=skey("mod","sentd"))
        accepte = s2.checkbox("Dossier accepté", value=int(row.get("Dossier accepté",0) or 0)==1, key=skey("mod","acc"))
        acc_d  = s2.date_input("Date d'acceptation", value=_date_for_widget(row.get("Date d'acceptation")), key=skey("mod","accd"))
        refuse = s3.checkbox("Dossier refusé", value=int(row.get("Dossier refusé",0) or 0)==1, key=skey("mod","ref"))
        ref_d  = s3.date_input("Date de refus", value=_date_for_widget(row.get("Date de refus")), key=skey("mod","refd"))
        annule = s4.checkbox("Dossier annulé", value=int(row.get("Dossier annulé",0) or 0)==1, key=skey("mod","ann"))
        ann_d  = s4.date_input("Date d'annulation", value=_date_for_widget(row.get("Date d'annulation")), key=skey("mod","annd"))
        rfe    = s5.checkbox("RFE", value=int(row.get("RFE",0) or 0)==1, key=skey("mod","rfe"))
        if rfe and not any([envoye, accepte, refuse, annule]):
            st.warning("RFE doit être combiné à un autre statut.")

        if st.button("💾 Enregistrer les modifications", key=skey("mod","save")):
            if not nom:
                st.warning("Nom requis.")
                st.stop()
            if not (cat and sub):
                st.warning("Choisissez Catégorie et Sous-catégorie.")
                st.stop()

            total = float(honor) + float(other)
            pay_list = _parse_payments(row.get("Paiements"))
            paye = _payments_total(pay_list)
            reste = max(0.0, total - paye)

            df_live.at[idx, "Nom"] = nom
            df_live.at[idx, "Date"] = dt
            df_live.at[idx, "Mois"] = f"{int(mois):02d}"
            df_live.at[idx, "Categorie"] = cat
            df_live.at[idx, "Sous-categorie"] = sub
            df_live.at[idx, "Visa"] = (visa_final or sub)
            df_live.at[idx, HONO] = float(honor)
            df_live.at[idx, AUTRE] = float(other)
            df_live.at[idx, TOTAL] = total
            df_live.at[idx, "Payé"] = paye
            df_live.at[idx, "Reste"] = reste
            df_live.at[idx, "Options"] = opts_dict
            df_live.at[idx, "Commentaire"] = comment
            df_live.at[idx, "Dossier envoyé"] = int(bool(envoye))
            df_live.at[idx, "Date d'envoi"] = sent_d
            df_live.at[idx, "Dossier accepté"] = int(bool(accepte))
            df_live.at[idx, "Date d'acceptation"] = acc_d
            df_live.at[idx, "Dossier refusé"] = int(bool(refuse))
            df_live.at[idx, "Date de refus"] = ref_d
            df_live.at[idx, "Dossier annulé"] = int(bool(annule))
            df_live.at[idx, "Date d'annulation"] = ann_d
            df_live.at[idx, "RFE"] = int(bool(rfe))

            _write_clients(normalize_clients(df_live))
            st.success("Modifications enregistrées.")
            st.cache_data.clear()
            st.rerun()

    # ---------- SUPPRESSION ----------
    elif op == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client")
        if df_live.empty:
            st.info("Aucun client.")
            st.stop()

        names = sorted(df_live["Nom"].dropna().astype(str).unique().tolist())
        ids   = sorted(df_live["ID_Client"].dropna().astype(str).unique().tolist())
        s1, s2 = st.columns(2)
        target_name = s1.selectbox("Nom", [""]+names, index=0, key=skey("del","nm"))
        target_id   = s2.selectbox("ID_Client", [""]+ids, index=0, key=skey("del","id"))

        mask = None
        if target_id:
            mask = (df_live["ID_Client"].astype(str) == target_id)
        elif target_name:
            mask = (df_live["Nom"].astype(str) == target_name)

        if mask is not None and mask.any():
            row = df_live[mask].iloc[0]
            st.write({"Dossier N": row.get(DOSSIER_COL,""), "Nom": row.get("Nom",""), "Visa": row.get("Visa","")})
            if st.button("❗ Confirmer la suppression", key=skey("del","ok")):
                df_new = df_live[~mask].copy()
                _write_clients(normalize_clients(df_new))
                st.success("Client supprimé.")
                st.cache_data.clear()
                st.rerun()

# ==============================================
# 📄 ONGLET : Visa — aperçu (catégorie / sous-catégorie / options)
# ==============================================
with tabs[5]:
    st.subheader("📄 Visa — aperçu")

    if not visa_map:
        st.info("Aucune structure Visa chargée.")
    else:
        cats = sorted(list(visa_map.keys()))
        cat = st.selectbox("Catégorie", [""] + cats, index=0, key=skey("vprev","cat"))
        if cat:
            subs = sorted(list(visa_map.get(cat, {}).keys()))
            sub = st.selectbox("Sous-catégorie", [""] + subs, index=0, key=skey("vprev","sub"))
            if sub:
                data = visa_map.get(cat, {}).get(sub, {})
                exo = data.get("exclusive") or []
                opts = data.get("options") or []
                st.markdown("**Options exclusives possibles :** " + (", ".join(exo) if exo else "—"))
                st.markdown("**Autres options :** " + (", ".join(opts) if opts else "—"))

        st.markdown("#### Table Visa brute")
        df_show = df_visa_raw.copy()
        if not df_show.empty:
            st.dataframe(df_show, use_container_width=True, key=skey("vprev","table"))
        else:
            st.caption("Feuille Visa vide.")

# ==============================================
# 💾 Export — (optionnel) sauvegarder Clients et/ou classeur combiné
# ==============================================
st.divider()
st.markdown("### 💾 Sauvegarde / Export")

colE1, colE2 = st.columns(2)
with colE1:
    if not _read_clients().empty:
        data_bytes = write_clients_to_bytes(_read_clients())
        st.download_button(
            "⬇️ Télécharger Clients.xlsx",
            data=data_bytes,
            file_name="Clients.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=skey("dl","clients"),
        )
    else:
        st.caption("Aucun Clients à exporter.")

with colE2:
    if (not _read_clients().empty) and (not df_visa_raw.empty):
        both_bytes = write_two_sheets_to_bytes(_read_clients(), df_visa_raw)
        st.download_button(
            "⬇️ Télécharger classeur (Clients + Visa).xlsx",
            data=both_bytes,
            file_name="Visa_Manager.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=skey("dl","both"),
        )
    else:
        st.caption("Pour exporter les deux onglets, chargez Clients et Visa.")