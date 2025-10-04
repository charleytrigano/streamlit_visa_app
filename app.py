# app.py
import io
import json
import hashlib
from datetime import date
from pathlib import Path

import streamlit as st
import pandas as pd

st.set_page_config(page_title="📊 Visas — Edition directe", layout="wide")
st.title("📊 Visas — Edition DIRECTE du fichier")

# ---------- Choix robuste d'un dossier de travail réinscriptible ----------
def pick_workdir() -> Path | None:
    candidates = [Path("/mnt/data"), Path("/tmp/visa_workspace"), Path.cwd() / "visa_workspace"]
    for p in candidates:
        try:
            p.mkdir(parents=True, exist_ok=True)
            t = p / ".write_test"
            t.write_text("ok", encoding="utf-8")
            t.unlink(missing_ok=True)
            return p
        except Exception:
            continue
    return None

WORK_DIR = pick_workdir()
WS_FILE = (WORK_DIR / "_workspace.json") if WORK_DIR else None

# ---------- Helpers persistance ----------
def load_workspace_path() -> Path | None:
    if WS_FILE is None or not WS_FILE.exists():
        return None
    try:
        obj = json.loads(WS_FILE.read_text(encoding="utf-8"))
        p = Path(obj.get("last_path", ""))
        return p if p.exists() else None
    except Exception:
        return None

def save_workspace_path(p: Path):
    if WS_FILE is None:
        return
    try:
        WS_FILE.write_text(json.dumps({"last_path": str(p)}), encoding="utf-8")
    except Exception:
        pass

def copy_upload_to_workspace(upload) -> Path:
    """Copie l'upload dans le WORK_DIR (ou /tmp à défaut), sans écraser par défaut."""
    base_dir = WORK_DIR if WORK_DIR else Path("/tmp")
    base_dir.mkdir(parents=True, exist_ok=True)
    name = getattr(upload, "name", "donnees_visa_clients.xlsx")
    dest = base_dir / name
    if dest.exists():
        stem, suf = dest.stem, dest.suffix
        n = 1
        while True:
            cand = base_dir / f"{stem}_{n}{suf}"
            if not cand.exists():
                dest = cand
                break
            n += 1
    data = upload.read()
    dest.write_bytes(data)
    return dest

# ---------- Utils data ----------
def _safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def _to_num(s: pd.Series) -> pd.Series:
    cleaned = (
        s.astype(str)
         .str.replace("\u00a0", "", regex=False)
         .str.replace("\u202f", "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.replace("$", "", regex=False)
         .str.replace(",", "", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)

def _to_date(s: pd.Series) -> pd.Series:
    d = pd.to_datetime(s, errors="coerce")
    try: d = d.dt.tz_localize(None)
    except Exception: pass
    return d.dt.normalize().dt.date  # YYYY-MM-DD

def _fmt_money_us(v: float) -> str:
    try: return f"${float(v):,.2f}"
    except Exception: return "$0.00"

def _parse_paiements(x):
    if isinstance(x, list): return x
    if pd.isna(x) or str(x).strip()== "": return []
    try:
        v = json.loads(x)
        return v if isinstance(v, list) else []
    except Exception:
        return []

def _sum_payments(pay_list) -> float:
    tot = 0.0
    for p in (pay_list or []):
        try: amt = float(p.get("amount", 0) if isinstance(p, dict) else p)
        except Exception: amt = 0.0
        tot += amt
    return tot

def _make_client_id_from_row(row) -> str:
    base = "|".join([_safe_str(row.get("Nom")), _safe_str(row.get("Telephone")), _safe_str(row.get("Date"))])
    h = hashlib.sha1(base.encode("utf-8")).hexdigest()[:8].upper()
    return f"CL-{h}"

def looks_like_reference(df: pd.DataFrame) -> bool:
    cols = set(map(str.lower, df.columns.astype(str)))
    has_ref = {"categories", "visa"} <= cols
    no_money = not ({"montant", "honoraires", "acomptes", "payé", "reste", "solde"} & cols)
    return has_ref and no_money

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "Date" in df.columns: df["Date"] = _to_date(df["Date"])
    else: df["Date"] = pd.NaT

    df["Mois"] = df["Date"].apply(lambda x: f"{x.month:02d}" if pd.notna(x) else pd.NA)

    # Visa
    visa_col = None
    for c in ["Visa", "Categories", "Catégorie", "TypeVisa"]:
        if c in df.columns: visa_col = c; break
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"

    # Montant / Payé
    if "Montant" in df.columns: df["Montant"] = _to_num(df["Montant"])
    else:
        src = None
        for c in ["Honoraires","Total","Amount"]:
            if c in df.columns: src=c; break
        df["Montant"] = _to_num(df[src]) if src else 0.0

    if "Payé" in df.columns: df["Payé"] = _to_num(df["Payé"])
    else:
        if "Paiements" in df.columns:
            parsed = df["Paiements"].apply(_parse_paiements)
            df["Payé"] = parsed.apply(_sum_payments).astype(float)
        else:
            df["Payé"] = 0.0

    df["Reste"] = (df["Montant"] - df["Payé"]).fillna(0.0)

    if "ID_Client" not in df.columns: df["ID_Client"] = ""
    need_id = df["ID_Client"].astype(str).str.strip().eq("") | df["ID_Client"].isna()
    if need_id.any():
        df.loc[need_id, "ID_Client"] = df.loc[need_id].apply(_make_client_id_from_row, axis=1)

    return df

# ---------- IO Excel DIRECTEMENT SUR DISQUE ----------
def read_sheet(path: Path, sheet: str, normalize: bool) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    if sheet not in xls.sheet_names:
        base = pd.DataFrame(columns=[
            "ID_Client","Nom","Telephone","Email","Date","Visa","Montant","Payé","Reste",
            "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé","Paiements"
        ])
        return normalize_dataframe(base) if normalize else base
    df = pd.read_excel(xls, sheet_name=sheet)
    if normalize and not looks_like_reference(df):
        df = normalize_dataframe(df)
    return df

def write_sheet_inplace(path: Path, sheet_to_replace: str, new_df: pd.DataFrame):
    """Réécrit le fichier Excel en remplaçant une feuille, puis écrase le fichier SOURCE (même nom)."""
    xls = pd.ExcelFile(path)
    out = io.BytesIO()
    target_written = False
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name in xls.sheet_names:
            if name == sheet_to_replace:
                dfw = new_df.copy()
                for c in dfw.columns:
                    if dfw[c].dtype == "object":
                        dfw[c] = dfw[c].astype(str).fillna("")
                dfw.to_excel(writer, sheet_name=name, index=False)
                target_written = True
            else:
                pd.read_excel(xls, sheet_name=name).to_excel(writer, sheet_name=name, index=False)
        if not target_written:
            dfw = new_df.copy()
            for c in dfw.columns:
                if dfw[c].dtype == "object":
                    dfw[c] = dfw[c].astype(str).fillna("")
            dfw.to_excel(writer, sheet_name=sheet_to_replace, index=False)

    # ECRITURE DIRECTE sur le même fichier + snapshot pour téléchargement
    bytes_out = out.getvalue()
    path.write_bytes(bytes_out)
    try:
        st.session_state["download_bytes"] = bytes_out
        st.session_state["download_name"] = path.name
    except Exception:
        pass

def write_analyses_sheet(path: Path, blocks: list[tuple[str, pd.DataFrame]]):
    """
    Écrit/Remplace la feuille 'Analyses' en plaçant chaque DataFrame à la suite
    (avec un titre au-dessus). Ne modifie pas les autres feuilles.
    """
    xls = pd.ExcelFile(path)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        # Copier toutes les feuilles sauf 'Analyses'
        for name in xls.sheet_names:
            if name == "Analyses":
                continue
            pd.read_excel(xls, sheet_name=name).to_excel(writer, sheet_name=name, index=False)

        # Nouvelle feuille 'Analyses'
        startrow = 0
        for title, df in blocks:
            # Titre
            pd.DataFrame({title: []}).to_excel(writer, sheet_name="Analyses", index=False, startrow=startrow)
            startrow += 1
            # Table
            df_to_write = df.copy()
            if isinstance(df_to_write, pd.Series):
                df_to_write = df_to_write.to_frame()
            df_to_write.to_excel(writer, sheet_name="Analyses", index=True, startrow=startrow)
            startrow += (len(df_to_write) + 3)  # espace

    # ECRITURE + snapshot
    bytes_out = out.getvalue()
    path.write_bytes(bytes_out)
    try:
        st.session_state["download_bytes"] = bytes_out
        st.session_state["download_name"] = path.name
    except Exception:
        pass

# ---------- Source : dernier fichier auto + remplacement ----------
st.sidebar.header("Source")
current_path = load_workspace_path()

if current_path and current_path.exists():
    st.sidebar.success(f"Fichier courant : {current_path.name}")
else:
    DEFAULTS = [
        "/mnt/data/Visa_Clients_20251001-114844.xlsx",
        "/mnt/data/visa_analytics_datecol.xlsx",
    ]
    candidate = next((Path(p) for p in DEFAULTS if Path(p).exists()), None)
    if candidate:
        current_path = candidate
        save_workspace_path(current_path)
        st.sidebar.success(f"Fichier courant : {current_path.name}")
    else:
        if WORK_DIR is None:
            st.sidebar.warning("Aucun workspace persistant disponible. Importez un fichier pour travailler (non mémorisé au redémarrage).")
        else:
            st.sidebar.warning("Aucun fichier trouvé. Importez un Excel pour démarrer.")

# Uploader pour CHANGER de fichier
up = st.sidebar.file_uploader("Remplacer par un autre Excel (.xlsx, .xls)", type=["xlsx","xls"])
if up is not None:
    new_path = copy_upload_to_workspace(up)
    save_workspace_path(new_path)
    # snapshot pour téléchargement
    try:
        st.session_state["download_bytes"] = new_path.read_bytes()
        st.session_state["download_name"] = new_path.name
    except Exception:
        st.session_state["download_bytes"] = b""
        st.session_state["download_name"] = new_path.name
    st.sidebar.success(f"Nouveau fichier chargé : {new_path.name}")
    st.rerun()

# Si rien encore, on arrête proprement
if current_path is None or not current_path.exists():
    st.stop()

# snapshot au démarrage si absent
if "download_bytes" not in st.session_state or st.session_state.get("download_name") != current_path.name:
    try:
        st.session_state["download_bytes"] = current_path.read_bytes()
        st.session_state["download_name"] = current_path.name
    except Exception:
        st.session_state["download_bytes"] = b""
        st.session_state["download_name"] = current_path.name

# Liste des feuilles
try:
    sheet_names = pd.ExcelFile(current_path).sheet_names
except Exception as e:
    st.error(f"Impossible de lire l'Excel : {e}")
    st.stop()

# Choix des feuilles
preferred_order = ["Clients","Visa","Données normalisées"]
default_sheet = next((s for s in preferred_order if s in sheet_names), sheet_names[0])
sheet_choice = st.sidebar.selectbox("Feuille (Dashboard)", sheet_names, index=sheet_names.index(default_sheet))
client_sheet_default = "Clients" if "Clients" in sheet_names else sheet_choice
client_target_sheet = st.sidebar.selectbox("Feuille *Clients* (cible CRUD)", sheet_names,
                                           index=sheet_names.index(client_sheet_default))

ws_info = f"`{current_path}`" + ("" if WORK_DIR else "  \n(Mémorisation du dernier fichier indisponible : espace non réinscriptible)")
st.sidebar.caption(f"Édition **directe** dans : {ws_info}")

# bouton de téléchargement -> snapshot mémoire
st.sidebar.download_button(
    "⬇️ Télécharger une copie",
    data=st.session_state.get("download_bytes", b""),
    file_name=st.session_state.get("download_name", current_path.name),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="Télécharge le snapshot en mémoire (toujours à jour)."
)

# ---------- TABS ----------
tabs = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses"])

# ---------- DASHBOARD ----------
with tabs[0]:
    df = read_sheet(current_path, sheet_choice, normalize=True)

    # Filtres OFF par défaut
    with st.container():
        c1, c2, c3 = st.columns(3)
        years = sorted({d.year for d in df["Date"] if pd.notna(d)}) if "Date" in df.columns else []
        months = sorted(df["Mois"].dropna().unique()) if "Mois" in df.columns else []
        visas = sorted(df["Visa"].dropna().astype(str).unique()) if "Visa" in df.columns else []

        sel_years  = c1.multiselect("Année", years, default=[])
        sel_months = c2.multiselect("Mois (MM)", months, default=[])
        sel_visas  = c3.multiselect("Type de visa", visas, default=[])

        c4, c5, c6 = st.columns(3)
        include_na_dates = c6.checkbox("Inclure lignes sans date", value=True)

        def make_slider(_df: pd.DataFrame, col: str, lab: str, container):
            if col not in _df.columns or _df[col].dropna().empty:
                container.caption(f"{lab} : aucune donnée")
                return None
            vmin, vmax = float(_df[col].min()), float(_df[col].max())
            if not (vmin < vmax):
                container.caption(f"{lab} : valeur unique = {_fmt_money_us(vmin)}")
                return (vmin, vmax)
            step = 1.0 if (vmax - vmin) > 1000 else 0.1 if (vmax - vmin) > 10 else 0.01
            return container.slider(lab, min_value=vmin, max_value=vmax, value=(vmin, vmax), step=step)

        pay_range   = make_slider(df, "Payé", "Payé (min-max)", c4)
        reste_range = make_slider(df, "Reste", "Solde/Reste (min-max)", c5)

    f = df.copy()
    if "Date" in f.columns and sel_years:
        mask = f["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask |= f["Date"].isna()
        f = f[mask]
    if "Mois" in f.columns and sel_months:
        mask = f["Mois"].isin(sel_months)
        if include_na_dates: mask |= f["Mois"].isna()
        f = f[mask]
    if "Visa" in f.columns and sel_visas:
        f = f[f["Visa"].astype(str).isin(sel_visas)]
    if "Payé" in f.columns and pay_range is not None:
        f = f[(f["Payé"] >= pay_range[0]) & (f["Payé"] <= pay_range[1])]
    if "Reste" in f.columns and reste_range is not None:
        f = f[(f["Reste"] >= reste_range[0]) & (f["Reste"] <= reste_range[1])]

    hidden = len(df) - len(f)
    if hidden > 0:
        st.caption(f"🔎 {hidden} ligne(s) masquée(s) par les filtres.")

    # KPI compacts
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Montant total", _fmt_money_us(float(f.get("Montant", pd.Series(dtype=float)).sum())) )
    k3.metric("Payé", _fmt_money_us(float(f.get("Payé", pd.Series(dtype=float)).sum())) )
    k4.metric("Reste", _fmt_money_us(float(f.get("Reste", pd.Series(dtype=float)).sum())) )
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    # Paiements (ajouts successifs) — écrit DIRECTEMENT dans current_path
    st.subheader("➕ Ajouter un paiement (US $)")
    clients_norm = read_sheet(current_path, client_target_sheet, normalize=True)
    todo = clients_norm[clients_norm["Reste"] > 0.004].copy() if "Reste" in clients_norm.columns else pd.DataFrame()
    if todo.empty:
        st.success("Tous les dossiers sont soldés ✅")
    else:
        todo["_label"] = todo.apply(lambda r: f'{r.get("ID_Client","")} — {r.get("Nom","")} — Reste {_fmt_money_us(float(r.get("Reste",0)))}', axis=1)
        label_to_id = todo.set_index("_label")["ID_Client"].to_dict()

        csel, camt, cdate, cmode = st.columns([2,1,1,1])
        sel_label = csel.selectbox("Dossier à créditer", todo["_label"].tolist())
        amount = camt.number_input("Montant ($)", min_value=0.0, step=10.0, format="%.2f")
        pdate  = cdate.date_input("Date", value=date.today())
        mode   = cmode.selectbox("Mode", ["CB","Chèque","Espèces","Virement","Autre"])
        note   = st.text_input("Note (facultatif)", "")

        if st.button("💾 Ajouter le paiement (écrit dans le fichier)"):
            try:
                live = read_sheet(current_path, client_target_sheet, normalize=False)
                if "Paiements" not in live.columns: live["Paiements"] = ""
                # Trouver la ligne
                target_id = label_to_id.get(sel_label, "")
                idxs = live.index[live.get("ID_Client","").astype(str) == str(target_id)]
                if len(idxs)==0:
                    raise RuntimeError("Dossier introuvable.")
                idx = idxs[0]

                reste = float(todo.set_index("_label").loc[sel_label, "Reste"])
                add = float(amount or 0.0)
                if add <= 0:
                    st.warning("Le montant doit être > 0.")
                    st.stop()
                if add > reste + 1e-9:
                    st.info(f"Le paiement dépasse le reste. Plafonné à {_fmt_money_us(reste)}.")
                    add = reste

                pay_list = _parse_paiements(live.at[idx, "Paiements"])
                pay_list.append({"date": str(pdate), "amount": float(add), "mode": mode, "note": note})
                live.at[idx, "Paiements"] = json.dumps(pay_list, ensure_ascii=False)

                if "Payé" not in live.columns: live["Payé"] = 0.0
                total_paid = _sum_payments(pay_list)
                live.at[idx, "Payé"] = float(total_paid)

                if "Montant" not in live.columns: live["Montant"] = 0.0
                if "Reste" not in live.columns: live["Reste"] = 0.0
                try: m = float(live.at[idx, "Montant"])
                except Exception: m = 0.0
                live.at[idx, "Reste"] = max(m - float(total_paid), 0.0)

                write_sheet_inplace(current_path, client_target_sheet, live)
                st.success("Paiement enregistré **dans le fichier**. ✅")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur : {e}")

    # Tableau
    st.subheader("📋 Données")
    cols_show = [c for c in ["ID_Client","Nom","Telephone","Email","Date","Visa","Montant","Payé","Reste",
                             "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"] if c in f.columns]
    table = f.copy()
    for col in ["Montant","Payé","Reste"]:
        if col in table.columns: table[col] = table[col].map(_fmt_money_us)
    if "Date" in table.columns: table["Date"] = table["Date"].astype(str)
    st.dataframe(table[cols_show].sort_values(by=[c for c in ["Date","Visa"] if c in table.columns], na_position="last"),
                 use_container_width=True)

# ---------- CLIENTS (CRUD) ----------
with tabs[1]:
    st.subheader("👤 Clients — Créer / Modifier / Supprimer (écriture **directe**)")
    if st.button("🔄 Recharger le fichier"):
        st.rerun()

    live_raw = read_sheet(current_path, client_target_sheet, normalize=False).copy()
    live_raw["_RowID"] = range(len(live_raw))

    # Colonnes booléennes si présentes
    has_envoye  = "Dossier envoyé"  in live_raw.columns
    has_refuse  = "Dossier refusé"  in live_raw.columns
    has_annule  = "Dossier annulé"  in live_raw.columns
    has_appr    = "Dossier approuvé" in live_raw.columns
    has_rfe     = "RFE"             in live_raw.columns

    # Référentiel Visa (si dispo)
    try:
        visa_ref = read_sheet(current_path, "Visa", normalize=False)
        visa_options = sorted(visa_ref["Visa"].dropna().astype(str).unique()) if "Visa" in visa_ref.columns else []
    except Exception:
        visa_options = []

    action = st.radio("Action", ["Créer", "Modifier", "Supprimer"], horizontal=True)

    # ---- CREER ----
    if action == "Créer":
        st.markdown("### ➕ Nouveau client")
        # Cols minimales si besoin
        for must in ["ID_Client","Nom","Telephone","Email","Date","Visa","Montant","Payé","Reste","Paiements",
                     "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"]:
            if must not in live_raw.columns:
                if must in {"Montant","Payé","Reste"}: live_raw[must]=0.0
                elif must=="Paiements": live_raw[must]=""
                elif must in {"RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"}: live_raw[must]=False
                else: live_raw[must]=""

        with st.form("create_form", clear_on_submit=False):
            c1,c2 = st.columns(2)
            nom = c1.text_input("Nom")
            tel = c2.text_input("Telephone")
            c3,c4 = st.columns(2)
            email = c3.text_input("Email")
            d = c4.date_input("Date", value=date.today())
            visa = st.selectbox("Visa", visa_options, index=0) if visa_options else st.text_input("Visa")
            c5,c6 = st.columns(2)
            montant = c5.number_input("Montant (US $)", value=0.0, step=10.0, format="%.2f")
            paye    = c6.number_input("Payé (US $)", value=0.0, step=10.0, format="%.2f")

            if has_envoye or has_refuse or has_annule or has_appr or has_rfe:
                st.markdown("#### État du dossier")
            val_envoye  = st.checkbox("Dossier envoyé",  value=False) if has_envoye else False
            val_refuse  = st.checkbox("Dossier refusé",  value=False) if has_refuse else False
            val_annule  = st.checkbox("Dossier annulé",  value=False) if has_annule else False
            val_appr    = st.checkbox("Dossier approuvé",value=False) if has_appr   else False
            val_rfe     = st.checkbox("RFE",             value=False) if has_rfe    else False

            ok = st.form_submit_button("💾 Sauvegarder (dans le fichier)", type="primary")

        if ok:
            if val_rfe and not (val_envoye or val_refuse or val_annule):
                st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché.")
                st.stop()

            gen_id = _make_client_id_from_row({"Nom": nom, "Telephone": tel, "Date": d})
            existing = set(live_raw["ID_Client"].astype(str)) if "ID_Client" in live_raw.columns else set()
            new_id = gen_id; n=1
            while new_id in existing:
                n+=1; new_id=f"{gen_id}-{n:02d}"

            new_row = {
                "ID_Client": new_id, "Nom": _safe_str(nom), "Telephone": _safe_str(tel), "Email": _safe_str(email),
                "Date": str(d), "Visa": _safe_str(visa), "Montant": float(montant or 0), "Payé": float(paye or 0),
                "Reste": float(montant or 0) - float(paye or 0), "Paiements": ""
            }
            if has_envoye: new_row["Dossier envoyé"] = bool(val_envoye)
            if has_refuse: new_row["Dossier refusé"] = bool(val_refuse)
            if has_annule: new_row["Dossier annulé"] = bool(val_annule)
            if has_appr:   new_row["Dossier approuvé"]= bool(val_appr)
            if has_rfe:    new_row["RFE"]             = bool(val_rfe)

            live_after = pd.concat([live_raw.drop(columns=["_RowID"]), pd.DataFrame([new_row])], ignore_index=True)
            write_sheet_inplace(current_path, client_target_sheet, live_after)
            save_workspace_path(current_path)
            st.success("Client créé **dans le fichier**. ✅")
            st.rerun()

    # ---- MODIFIER ----
    if action == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            opts = [(int(r["_RowID"]), f"{_safe_str(r.get('ID_Client'))} — {_safe_str(r.get('Nom'))}") for _,r in live_raw.iterrows()]
            label = st.selectbox("Sélection", [lab for _,lab in opts])
            sel_rowid = [rid for rid,lab in opts if lab==label][0]
            idx = live_raw.index[live_raw["_RowID"]==sel_rowid][0]
            init = live_raw.loc[idx].to_dict()

            with st.form("edit_form", clear_on_submit=False):
                c1,c2 = st.columns(2)
                nom = c1.text_input("Nom", value=_safe_str(init.get("Nom")))
                tel = c2.text_input("Telephone", value=_safe_str(init.get("Telephone")))
                c3,c4 = st.columns(2)
                email = c3.text_input("Email", value=_safe_str(init.get("Email")))
                try: d_init = pd.to_datetime(init.get("Date")).date() if _safe_str(init.get("Date")) else date.today()
                except Exception: d_init = date.today()
                d = c4.date_input("Date", value=d_init)

                if visa_options:
                    try: idx_vis = visa_options.index(_safe_str(init.get("Visa")))
                    except Exception: idx_vis = 0
                    visa = st.selectbox("Visa", visa_options, index=idx_vis)
                else:
                    visa = st.text_input("Visa", value=_safe_str(init.get("Visa")))

                c5,c6 = st.columns(2)
                try: montant0=float(init.get("Montant",0))
                except Exception: montant0=0.0
                try: paye0=float(init.get("Payé",0))
                except Exception: paye0=0.0
                montant = c5.number_input("Montant (US $)", value=montant0, step=10.0, format="%.2f")
                paye    = c6.number_input("Payé (US $)", value=paye0,    step=10.0, format="%.2f")

                val_envoye  = st.checkbox("Dossier envoyé",  value=bool(init.get("Dossier envoyé")))   if "Dossier envoyé" in live_raw.columns else False
                val_refuse  = st.checkbox("Dossier refusé",  value=bool(init.get("Dossier refusé")))   if "Dossier refusé" in live_raw.columns else False
                val_annule  = st.checkbox("Dossier annulé",  value=bool(init.get("Dossier annulé")))   if "Dossier annulé" in live_raw.columns else False
                val_appr    = st.checkbox("Dossier approuvé",value=bool(init.get("Dossier approuvé"))) if "Dossier approuvé" in live_raw.columns else False
                val_rfe     = st.checkbox("RFE",             value=bool(init.get("RFE")))              if "RFE" in live_raw.columns else False

                ok = st.form_submit_button("💾 Enregistrer (dans le fichier)", type="primary")

            if ok:
                if val_rfe and not (val_envoye or val_refuse or val_annule):
                    st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché.")
                    st.stop()
                live = live_raw.drop(columns=["_RowID"]).copy()
                # Retrouver la ligne par ID si possible
                t_idx = None
                if "ID_Client" in live.columns and _safe_str(init.get("ID_Client")):
                    hits = live.index[live["ID_Client"].astype(str) == _safe_str(init.get("ID_Client"))]
                    if len(hits)>0: t_idx = hits[0]
                if t_idx is None:
                    msk = (live.get("Nom","").astype(str)==_safe_str(init.get("Nom"))) & \
                          (live.get("Telephone","").astype(str)==_safe_str(init.get("Telephone")))
                    hit2 = live.index[msk]
                    t_idx = hit2[0] if len(hit2)>0 else None
                if t_idx is None:
                    st.error("Ligne introuvable.")
                else:
                    live.at[t_idx,"Nom"]=_safe_str(nom)
                    live.at[t_idx,"Telephone"]=_safe_str(tel)
                    live.at[t_idx,"Email"]=_safe_str(email)
                    live.at[t_idx,"Date"]=str(d)
                    live.at[t_idx,"Visa"]=_safe_str(visa)
                    live.at[t_idx,"Montant"]=float(montant or 0)
                    live.at[t_idx,"Payé"]=float(paye or 0)
                    live.at[t_idx,"Reste"]=float(live.at[t_idx,"Montant"])-float(live.at[t_idx,"Payé"])
                    if "Dossier envoyé" in live.columns:   live.at[t_idx,"Dossier envoyé"]=bool(val_envoye)
                    if "Dossier refusé" in live.columns:   live.at[t_idx,"Dossier refusé"]=bool(val_refuse)
                    if "Dossier annulé" in live.columns:   live.at[t_idx,"Dossier annulé"]=bool(val_annule)
                    if "Dossier approuvé" in live.columns: live.at[t_idx,"Dossier approuvé"]=bool(val_appr)
                    if "RFE" in live.columns:              live.at[t_idx,"RFE"]=bool(val_rfe)

                    write_sheet_inplace(current_path, client_target_sheet, live)
                    save_workspace_path(current_path)
                    st.success("Modifications enregistrées **dans le fichier**. ✅")
                    st.rerun()

    # ---- SUPPRIMER ----
    if action == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client (écrit directement)")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            opts = [(int(r["_RowID"]), f"{_safe_str(r.get('ID_Client'))} — {_safe_str(r.get('Nom'))}") for _,r in live_raw.iterrows()]
            label = st.selectbox("Sélection", [lab for _,lab in opts])
            sel_rowid = [rid for rid,lab in opts if lab==label][0]
            idx = live_raw.index[live_raw["_RowID"]==sel_rowid][0]

            st.error("⚠️ Action irréversible.")
            if st.button("Supprimer (dans le fichier)"):
                live = live_raw.drop(columns=["_RowID"]).copy()
                key = _safe_str(live_raw.at[idx, "ID_Client"])
                if key and "ID_Client" in live.columns:
                    live = live[live["ID_Client"].astype(str)!=key].reset_index(drop=True)
                else:
                    nom = _safe_str(live_raw.at[idx,"Nom"]); tel=_safe_str(live_raw.at[idx,"Telephone"])
                    live = live[~((live.get("Nom","").astype(str)==nom)&(live.get("Telephone","").astype(str)==tel))].reset_index(drop=True)
                write_sheet_inplace(current_path, client_target_sheet, live)
                save_workspace_path(current_path)
                st.success("Client supprimé **dans le fichier**. ✅")
                st.rerun()

# ---------- ANALYSES ----------
with tabs[2]:
    st.subheader("📊 Analyses — Volumes & Financier")

    # On travaille sur la feuille Clients cible (normalisée)
    dfA = read_sheet(current_path, client_target_sheet, normalize=True).copy()

    if dfA.empty:
        st.info("Aucune donnée dans la feuille cible pour analyser.")
        st.stop()

    # Filtres (par défaut rien de sélectionné)
    with st.container():
        c1, c2, c3, c4 = st.columns([1,1,1,1])

        yearsA  = sorted({d.year for d in dfA["Date"] if pd.notna(d)}) if "Date" in dfA.columns else []
        monthsA = [f"{m:02d}" for m in range(1,13)]
        visasA  = sorted(dfA["Visa"].dropna().astype(str).unique()) if "Visa" in dfA.columns else []

        sel_years  = c1.multiselect("Année", yearsA, default=[])
        sel_months = c2.multiselect("Mois (MM)", monthsA, default=[])
        sel_visa   = c3.multiselect("Type de visa", visasA, default=[])
        include_na_dates = c4.checkbox("Inclure lignes sans date", value=True)

        c5, c6 = st.columns([1,1])
        agg_with_year = c5.toggle("Agrégation par Année-Mois (YYYY-MM)", value=False,
                                  help="Si OFF : agrégation par Mois (MM) toutes années confondues.")
        show_tables   = c6.toggle("Voir les tableaux en dessous des graphiques", value=False)

    # Application des filtres
    fA = dfA.copy()
    # Visa
    if sel_visa:
        fA = fA[fA["Visa"].astype(str).isin(sel_visa)]

    # Année
    if "Date" in fA.columns and sel_years:
        mask_year = fA["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates:
            mask_year = mask_year | fA["Date"].isna()
        fA = fA[mask_year]

    # Mois (MM)
    if "Mois" in fA.columns and sel_months:
        mask_month = fA["Mois"].isin(sel_months)
        if include_na_dates:
            mask_month = mask_month | fA["Mois"].isna()
        fA = fA[mask_month]

    # Période d'agrégation
    if agg_with_year:
        # YYYY-MM (si date manquante -> "NA")
        fA["Periode"] = fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")
        ordre_periodes = sorted([p for p in fA["Periode"].unique() if p != "NA"]) + (["NA"] if "NA" in fA["Periode"].values else [])
    else:
        # MM seulement (1..12), NA pour dates manquantes
        fA["Periode"] = fA["Mois"].fillna("NA")
        ordre_periodes = [f"{m:02d}" for m in range(1,13)]
        if "NA" in fA["Periode"].values:
            ordre_periodes = ordre_periodes + ["NA"]

    # Conversions numériques sûres
    for col in ["Montant","Payé","Reste"]:
        if col in fA.columns:
            try:
                fA[col] = pd.to_numeric(fA[col], errors="coerce").fillna(0.0)
            except Exception:
                fA[col] = 0.0

    # ===== KPI globaux (après filtres) =====
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers (filtrés)", f"{len(fA)}")
    k2.metric("Chiffre d’affaires", _fmt_money_us(float(fA.get("Montant", pd.Series(dtype=float)).sum())))
    k3.metric("Encaissements",       _fmt_money_us(float(fA.get("Payé",    pd.Series(dtype=float)).sum())))
    k4.metric("Solde à encaisser",   _fmt_money_us(float(fA.get("Reste",   pd.Series(dtype=float)).sum())))

    st.divider()

    # ===== Volumes par période =====
    st.markdown("### 📦 Volumes par période")

    def _safe_bool_sum(series):
        return series.fillna(False).astype(bool).sum()

    vol_all = fA.groupby("Periode").size().reindex(ordre_periodes).fillna(0).astype(int)

    vol_env = fA.groupby("Periode")["Dossier envoyé"].apply(_safe_bool_sum).reindex(ordre_periodes).fillna(0) if "Dossier envoyé" in fA.columns else pd.Series(dtype=int)
    vol_app = fA.groupby("Periode")["Dossier approuvé"].apply(_safe_bool_sum).reindex(ordre_periodes).fillna(0) if "Dossier approuvé" in fA.columns else pd.Series(dtype=int)
    vol_ref = fA.groupby("Periode")["Dossier refusé"].apply(_safe_bool_sum).reindex(ordre_periodes).fillna(0)   if "Dossier refusé" in fA.columns else pd.Series(dtype=int)
    vol_ann = fA.groupby("Periode")["Dossier annulé"].apply(_safe_bool_sum).reindex(ordre_periodes).fillna(0)   if "Dossier annulé" in fA.columns else pd.Series(dtype=int)

    cvol1, cvol2 = st.columns(2)
    cvol1.caption("Dossiers ouverts")
    st_data1 = pd.DataFrame({"Ouverts": vol_all})
    cvol1.bar_chart(st_data1)

    vols_dict = {}
    if not vol_env.empty: vols_dict["Envoyés"] = vol_env
    if not vol_app.empty: vols_dict["Approuvés"] = vol_app
    if not vol_ref.empty: vols_dict["Refusés"]  = vol_ref
    if not vol_ann.empty: vols_dict["Annulés"]  = vol_ann

    if vols_dict:
        st_data2 = pd.DataFrame(vols_dict)
        cvol2.caption("Statuts par période")
        cvol2.bar_chart(st_data2)
        if show_tables:
            with st.expander("Détail tableaux — Volumes"):
                st.write("Ouverts")
                st.dataframe(st_data1)
                st.write("Statuts")
                st.dataframe(st_data2)
    else:
        cvol2.info("Aucune colonne de statut trouvée (Envoyé/Approuvé/Refusé/Annulé).")

    st.divider()

    # ===== Financier par période =====
    st.markdown("### 💵 Financier par période")

    sums = fA.groupby("Periode")[["Montant","Payé","Reste"]].sum().reindex(ordre_periodes).fillna(0.0)
    cfin1, cfin2 = st.columns(2)
    cfin1.caption("Chiffre d'affaires (Montant)")
    cfin1.bar_chart(sums[["Montant"]])
    cfin2.caption("Encaissements (Payé) & Solde à encaisser (Reste)")
    cfin2.bar_chart(sums[["Payé","Reste"]])

    if show_tables:
        with st.expander("Détail tableaux — Financier"):
            st.dataframe(sums)

    st.divider()

    # ===== Top visas =====
    st.markdown("### 🏷️ Top visas")
    top_vol = fA.groupby("Visa").size().sort_values(ascending=False).head(15).rename("Dossiers")
    top_ca = fA.groupby("Visa")["Montant"].sum().sort_values(ascending=False).head(15)

    ctop1, ctop2 = st.columns(2)
    ctop1.caption("Top visas par nombre de dossiers")
    ctop1.bar_chart(pd.DataFrame(top_vol))
    ctop2.caption("Top visas par chiffre d'affaires")
    ctop2.bar_chart(pd.DataFrame({"CA": top_ca}))

    if show_tables:
        with st.expander("Détail tableaux — Top visas"):
            st.write("Par dossiers")
            st.dataframe(pd.DataFrame(top_vol))
            st.write("Par CA")
            st.dataframe(pd.DataFrame({"CA": top_ca}))

    st.divider()

    # ===== Export analyses -> Excel / feuille 'Analyses' =====
    st.markdown("### 📤 Export Excel")
    st.caption("Crée/Met à jour la feuille **'Analyses'** dans ton fichier courant avec l'ensemble des tableaux.")

    if st.button("Exporter vers l'Excel (feuille 'Analyses')"):
        try:
            blocks = [("KPI Globaux (après filtres)",
                       pd.DataFrame({
                           "KPI": ["Dossiers (filtrés)", "Chiffre d’affaires", "Encaissements", "Solde à encaisser"],
                           "Valeur": [
                               len(fA),
                               float(fA.get("Montant", pd.Series(dtype=float)).sum()),
                               float(fA.get("Payé", pd.Series(dtype=float)).sum()),
                               float(fA.get("Reste", pd.Series(dtype=float)).sum()),
                           ],
                       }))
            ]
            blocks.append(("Volumes — Dossiers ouverts par période", st_data1))
            if vols_dict:
                blocks.append(("Volumes — Statuts par période", st_data2))
            blocks.append(("Financier — Montant / Payé / Reste par période", sums))
            blocks.append(("Top visas — par nombre de dossiers", pd.DataFrame(top_vol)))
            blocks.append(("Top visas — par chiffre d'affaires", pd.DataFrame({"CA": top_ca})))

            write_analyses_sheet(current_path, blocks)
            save_workspace_path(current_path)
            st.success("Feuille **'Analyses'** exportée dans le fichier. ✅ (Téléchargement → sidebar)")
        except Exception as e:
            st.error(f"Échec export : {e}")
