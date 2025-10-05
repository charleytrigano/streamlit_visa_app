# app.py
import io
import json
import hashlib
from datetime import date, timedelta
from pathlib import Path
import zipfile

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
    return d.dt.normalize().dt.date

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
    base = "|".join([_safe_str(row.get("Nom")), _safe_str(row.get("Date"))])
    h = hashlib.sha1(base.encode("utf-8")).hexdigest()[:8].upper()
    return f"CL-{h}"

def looks_like_reference(df: pd.DataFrame) -> bool:
    cols = set(map(str.lower, df.columns.astype(str)))
    # Référentiel "Visa" = contient 'visa', pas de colonnes financières
    has_visa = "visa" in cols
    no_money = not ({"montant", "honoraires", "acomptes", "payé", "reste", "solde"} & cols)
    return has_visa and no_money

def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "Date" in df.columns: df["Date"] = _to_date(df["Date"])
    else: df["Date"] = pd.NaT
    df["Mois"] = df["Date"].apply(lambda x: f"{x.month:02d}" if pd.notna(x) else pd.NA)

    visa_col = None
    for c in ["Visa", "Categories", "Catégorie", "TypeVisa"]:
        if c in df.columns: visa_col = c; break
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"

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

    for b in ["RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"]:
        if b not in df.columns: df[b] = False

    # On s'assure d'un schéma cohérent (sans Téléphone/Email)
    for dropcol in ["Telephone","Email"]:
        if dropcol in df.columns: df = df.drop(columns=[dropcol])

    if "Nom" not in df.columns: df["Nom"] = ""
    if "ID_Client" not in df.columns: df["ID_Client"] = ""
    need_id = df["ID_Client"].astype(str).str.strip().eq("") | df["ID_Client"].isna()
    if need_id.any():
        df.loc[need_id, "ID_Client"] = df.loc[need_id].apply(_make_client_id_from_row, axis=1)
    if "Paiements" not in df.columns: df["Paiements"] = ""

    ordered = ["ID_Client","Nom","Date","Mois","Visa","Montant","Payé","Reste",
               "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé","Paiements"]
    cols = [c for c in ordered if c in df.columns] + [c for c in df.columns if c not in ordered]
    return df[cols]

def is_clients_like(df: pd.DataFrame) -> bool:
    cols = set(df.columns.astype(str))
    return {"Nom","Visa"}.issubset(cols)

# ---------- IO Excel ----------
def read_sheet(path: Path, sheet: str, normalize: bool) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    if sheet not in xls.sheet_names:
        base = pd.DataFrame(columns=[
            "ID_Client","Nom","Date","Mois","Visa","Montant","Payé","Reste",
            "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé","Paiements"
        ])
        return normalize_dataframe(base) if normalize else base
    df = pd.read_excel(xls, sheet_name=sheet)
    if normalize and not looks_like_reference(df):
        df = normalize_dataframe(df)
    return df

def write_sheet_inplace(path: Path, sheet_to_replace: str, new_df: pd.DataFrame):
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

    bytes_out = out.getvalue()
    path.write_bytes(bytes_out)
    try:
        st.session_state["download_bytes"] = bytes_out
        st.session_state["download_name"] = path.name
    except Exception:
        pass

def write_analyses_sheet(path: Path, blocks: list[tuple[str, pd.DataFrame]]):
    xls = pd.ExcelFile(path)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name in xls.sheet_names:
            if name == "Analyses":
                continue
            pd.read_excel(xls, sheet_name=name).to_excel(writer, sheet_name=name, index=False)

        startrow = 0
        for title, df in blocks:
            pd.DataFrame({title: []}).to_excel(writer, sheet_name="Analyses", index=False, startrow=startrow)
            startrow += 1
            df_to_write = df.copy()
            if isinstance(df_to_write, pd.Series):
                df_to_write = df_to_write.to_frame()
            df_to_write.to_excel(writer, sheet_name="Analyses", index=True, startrow=startrow)
            startrow += (len(df_to_write) + 3)

    bytes_out = out.getvalue()
    path.write_bytes(bytes_out)
    try:
        st.session_state["download_bytes"] = bytes_out
        st.session_state["download_name"] = path.name
    except Exception:
        pass

# ---------- Source ----------
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
            st.sidebar.warning("Aucun workspace persistant disponible. Importez un fichier.")
        else:
            st.sidebar.warning("Aucun fichier trouvé. Importez un Excel pour démarrer.")

up = st.sidebar.file_uploader("Remplacer par un Excel (.xlsx, .xls)", type=["xlsx","xls"])
if up is not None:
    new_path = copy_upload_to_workspace(up)
    save_workspace_path(new_path)
    try:
        st.session_state["download_bytes"] = new_path.read_bytes()
        st.session_state["download_name"] = new_path.name
    except Exception:
        st.session_state["download_bytes"] = b""
        st.session_state["download_name"] = new_path.name
    st.sidebar.success(f"Nouveau fichier chargé : {new_path.name}")
    st.rerun()

if current_path is None or not current_path.exists():
    st.stop()

if "download_bytes" not in st.session_state or st.session_state.get("download_name") != current_path.name:
    try:
        st.session_state["download_bytes"] = current_path.read_bytes()
        st.session_state["download_name"] = current_path.name
    except Exception:
        st.session_state["download_bytes"] = b""
        st.session_state["download_name"] = current_path.name

# ---------- Détection des feuilles + verrouillage CRUD ----------
try:
    sheet_names = pd.ExcelFile(current_path).sheet_names
except Exception as e:
    st.error(f"Impossible de lire l'Excel : {e}")
    st.stop()

# Feuille affichée dans le Dashboard (libre)
preferred_order = ["Clients","Visa","Données normalisées"]
default_sheet = next((s for s in preferred_order if s in sheet_names), sheet_names[0])
sheet_choice = st.sidebar.selectbox("Feuille (Dashboard)", sheet_names, index=sheet_names.index(default_sheet))

# Liste des feuilles éligibles CRUD = au moins Nom & Visa
valid_client_sheets = []
for s in sheet_names:
    try:
        df_tmp = read_sheet(current_path, s, normalize=False)
        if is_clients_like(df_tmp):
            valid_client_sheets.append(s)
    except Exception:
        pass

if not valid_client_sheets:
    st.sidebar.error("Aucune feuille 'clients' valide trouvée (il faut au minimum les colonnes Nom & Visa).")
    client_target_sheet = None
else:
    default_client_sheet = "Clients" if "Clients" in valid_client_sheets else valid_client_sheets[0]
    # Nettoyage session pour éviter clés invalides
    if "client_sheet_select" in st.session_state and st.session_state["client_sheet_select"] not in valid_client_sheets:
        del st.session_state["client_sheet_select"]
    client_target_sheet = st.sidebar.selectbox(
        "Feuille *Clients* (cible CRUD)",
        valid_client_sheets,
        index=valid_client_sheets.index(default_client_sheet),
        key="client_sheet_select"
    )
    if len(valid_client_sheets) < len(sheet_names):
        st.sidebar.caption("🔒 Seules les feuilles contenant **Nom** et **Visa** sont proposées ici.")

ws_info = f"`{current_path}`" + ("" if WORK_DIR else "  \n(Espace non réinscriptible : dernier fichier non mémorisé)")
st.sidebar.caption(f"Édition **directe** dans : {ws_info}")

st.sidebar.download_button(
    "⬇️ Télécharger une copie",
    data=st.session_state.get("download_bytes", b""),
    file_name=st.session_state.get("download_name", current_path.name),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ---------- TABS ----------
tabs = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses"])

# ---------- DASHBOARD ----------
with tabs[0]:
    # Si la feuille choisie est un référentiel Visa → UI dédiée
    df_raw = read_sheet(current_path, sheet_choice, normalize=False)
    if looks_like_reference(df_raw):
        st.subheader("📄 Référentiel — Types de Visa")
        if "Visa" not in df_raw.columns:
            df_raw = pd.DataFrame(columns=["Visa"])
        df_ref = pd.DataFrame({"Visa": df_raw["Visa"].astype(str).fillna("").str.strip()})
        st.dataframe(df_ref, use_container_width=True)

        st.markdown("### ✏️ Gérer les types")
        action = st.radio("Action", ["Ajouter", "Renommer", "Supprimer"], horizontal=True, key="visa_ref_action")
        options = sorted([v for v in df_ref["Visa"].unique() if v], key=str.lower)

        if action == "Ajouter":
            new_v = st.text_input("Nouveau type de visa").strip()
            if st.button("➕ Ajouter"):
                if not new_v:
                    st.warning("Saisis un libellé.")
                elif new_v in options:
                    st.info("Ce type existe déjà.")
                else:
                    out = pd.concat([df_ref, pd.DataFrame([{"Visa": new_v}])], ignore_index=True)
                    write_sheet_inplace(current_path, sheet_choice, out)
                    st.success("Type ajouté.")
                    st.rerun()

        elif action == "Renommer":
            if not options:
                st.info("Aucun type existant.")
            else:
                old = st.selectbox("Type à renommer", options)
                new = st.text_input("Nouveau libellé").strip()
                if st.button("📝 Renommer"):
                    if not new:
                        st.warning("Nouveau libellé requis.")
                    elif new == old:
                        st.info("Aucun changement.")
                    elif new in options:
                        st.info("Un type avec ce nom existe déjà.")
                    else:
                        out = df_ref.copy()
                        out.loc[out["Visa"] == old, "Visa"] = new
                        write_sheet_inplace(current_path, sheet_choice, out)
                        st.success("Type renommé.")
                        st.rerun()

        else:  # Supprimer
            if not options:
                st.info("Aucun type à supprimer.")
            else:
                rm = st.selectbox("Type à supprimer", options)
                st.error("⚠️ Cette action est irréversible.")
                if st.button("🗑️ Supprimer"):
                    out = df_ref[df_ref["Visa"] != rm].reset_index(drop=True)
                    write_sheet_inplace(current_path, sheet_choice, out)
                    st.success("Type supprimé.")
                    st.rerun()

        st.info("Astuce : sélectionne une autre feuille dans la sidebar pour revenir au Dashboard classique.")
        st.stop()

    # Dashboard classique (feuille transactionnelle)
    df = read_sheet(current_path, sheet_choice, normalize=True)

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
        reste_range = make_slider(df, "Reste", "Solde (min-max)", c5)

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

    # Paiements
    st.subheader("➕ Ajouter un paiement (US $)")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide dans la sidebar pour créditer des paiements.")
    else:
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
    cols_show = [c for c in ["ID_Client","Nom","Date","Visa","Montant","Payé","Reste",
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

    if client_target_sheet is None:
        st.warning("Aucune feuille *Clients* valide disponible. Ajoute une feuille avec au moins les colonnes **Nom** et **Visa**.")
        st.stop()

    if st.button("🔄 Recharger le fichier"):
        st.rerun()

    live_raw = read_sheet(current_path, client_target_sheet, normalize=False).copy()
    live_raw["_RowID"] = range(len(live_raw))

    has_envoye  = "Dossier envoyé"  in live_raw.columns
    has_appr    = "Dossier approuvé" in live_raw.columns
    has_rfe     = "RFE"             in live_raw.columns
    has_refuse  = "Dossier refusé"  in live_raw.columns
    has_annule  = "Dossier annulé"  in live_raw.columns

    try:
        visa_ref = read_sheet(current_path, "Visa", normalize=False)
        visa_options = sorted(visa_ref["Visa"].dropna().astype(str).unique()) if "Visa" in visa_ref.columns else []
    except Exception:
        visa_options = []

    action = st.radio("Action", ["Créer", "Modifier", "Supprimer"], horizontal=True)

    # ---- CREER ----
    if action == "Créer":
        st.markdown("### ➕ Nouveau client")
        for must in ["ID_Client","Nom","Date","Mois","Visa","Montant","Payé","Reste",
                     "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé","Paiements"]:
            if must not in live_raw.columns:
                if must in {"Montant","Payé","Reste"}: live_raw[must]=0.0
                elif must=="Paiements": live_raw[must]=""
                elif must in {"Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"}: live_raw[must]=False
                elif must=="Mois": live_raw[must]=""
                else: live_raw[must]=""

        with st.form("create_form", clear_on_submit=False):
            c1,c2 = st.columns(2)
            nom = c1.text_input("Nom")
            d = c2.date_input("Date", value=date.today())
            visa = st.selectbox("Visa", visa_options, index=0) if visa_options else st.text_input("Visa")
            c5,c6 = st.columns(2)
            montant = c5.number_input("Montant (US $)", value=0.0, step=10.0, format="%.2f")
            paye    = c6.number_input("Payé (US $)", value=0.0, step=10.0, format="%.2f")

            st.markdown("#### État du dossier")
            # Ordre: Envoyé > Approuvé > RFE > Refusé > Annulé
            val_envoye = st.checkbox("Dossier envoyé",  value=False) if has_envoye else False
            val_appr   = st.checkbox("Dossier approuvé",value=False) if has_appr   else False
            val_rfe    = st.checkbox("RFE",             value=False) if has_rfe    else False
            val_refuse = st.checkbox("Dossier refusé",  value=False) if has_refuse else False
            val_annule = st.checkbox("Dossier annulé",  value=False) if has_annule else False

            ok = st.form_submit_button("💾 Sauvegarder (dans le fichier)", type="primary")

        if ok:
            if val_rfe and not (val_envoye or val_refuse or val_annule):
                st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché.")
                st.stop()

            gen_id = _make_client_id_from_row({"Nom": nom, "Date": d})
            existing = set(live_raw["ID_Client"].astype(str)) if "ID_Client" in live_raw.columns else set()
            new_id = gen_id; n=1
            while new_id in existing:
                n+=1; new_id=f"{gen_id}-{n:02d}"

            new_row = {
                "ID_Client": new_id, "Nom": _safe_str(nom), "Date": str(d),
                "Mois": f"{d.month:02d}",
                "Visa": _safe_str(visa), "Montant": float(montant or 0), "Payé": float(paye or 0),
                "Reste": max(float(montant or 0) - float(paye or 0), 0.0), "Paiements": "",
                "Dossier envoyé": bool(val_envoye), "Dossier approuvé": bool(val_appr),
                "RFE": bool(val_rfe), "Dossier refusé": bool(val_refuse), "Dossier annulé": bool(val_annule),
            }

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
                try: d_init = pd.to_datetime(init.get("Date")).date() if _safe_str(init.get("Date")) else date.today()
                except Exception: d_init = date.today()
                d = c2.date_input("Date", value=d_init)

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

                st.markdown("#### État du dossier")
                val_envoye = st.checkbox("Dossier envoyé",  value=bool(init.get("Dossier envoyé")))   if has_envoye else False
                val_appr   = st.checkbox("Dossier approuvé",value=bool(init.get("Dossier approuvé"))) if has_appr   else False
                val_rfe    = st.checkbox("RFE",             value=bool(init.get("RFE")))              if has_rfe    else False
                val_refuse = st.checkbox("Dossier refusé",  value=bool(init.get("Dossier refusé")))   if has_refuse else False
                val_annule = st.checkbox("Dossier annulé",  value=bool(init.get("Dossier annulé")))   if has_annule else False

                ok = st.form_submit_button("💾 Enregistrer (dans le fichier)", type="primary")

            if ok:
                if val_rfe and not (val_envoye or val_refuse or val_annule):
                    st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché.")
                    st.stop()
                live = live_raw.drop(columns=["_RowID"]).copy()
                t_idx = None
                if "ID_Client" in live.columns and _safe_str(init.get("ID_Client")):
                    hits = live.index[live["ID_Client"].astype(str) == _safe_str(init.get("ID_Client"))]
                    if len(hits)>0: t_idx = hits[0]
                if t_idx is None:
                    msk = (live.get("Nom","").astype(str)==_safe_str(init.get("Nom"))) & \
                          (pd.to_datetime(live.get("Date",""), errors="coerce").dt.date == pd.to_datetime(_safe_str(init.get("Date")), errors="coerce").date())
                    hit2 = live.index[msk] if hasattr(msk, "__len__") else []
                    t_idx = hit2[0] if len(hit2)>0 else None
                if t_idx is None:
                    st.error("Ligne introuvable.")
                else:
                    live.at[t_idx,"Nom"]=_safe_str(nom)
                    live.at[t_idx,"Date"]=str(d)
                    live.at[t_idx,"Mois"]=f"{d.month:02d}"
                    live.at[t_idx,"Visa"]=_safe_str(visa)
                    live.at[t_idx,"Montant"]=float(montant or 0)
                    live.at[t_idx,"Payé"]=float(paye or 0)
                    live.at[t_idx,"Reste"]=max(float(live.at[t_idx,"Montant"])-float(live.at[t_idx,"Payé"]),0.0)
                    if has_envoye: live.at[t_idx,"Dossier envoyé"]=bool(val_envoye)
                    if has_appr:   live.at[t_idx,"Dossier approuvé"]=bool(val_appr)
                    if has_rfe:    live.at[t_idx,"RFE"]=bool(val_rfe)
                    if has_refuse: live.at[t_idx,"Dossier refusé"]=bool(val_refuse)
                    if has_annule: live.at[t_idx,"Dossier annulé"]=bool(val_annule)

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
                    nom = _safe_str(live_raw.at[idx,"Nom"])
                    dat = _safe_str(live_raw.at[idx,"Date"])
                    live = live[~((live.get("Nom","").astype(str)==nom)&(live.get("Date","").astype(str)==dat))].reset_index(drop=True)
                write_sheet_inplace(current_path, client_target_sheet, live)
                save_workspace_path(current_path)
                st.success("Client supprimé **dans le fichier**. ✅")
                st.rerun()

# ---------- ANALYSES ----------
with tabs[2]:
    st.subheader("📊 Analyses — Volumes & Financier")

    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa) pour lancer les analyses.")
        st.stop()

    # Normalisation FORCÉE ici
    dfA_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    dfA = normalize_dataframe(dfA_raw).copy()

    with st.container():
        c1, c2, c3, c4 = st.columns([1,1,1,1])
        yearsA  = sorted({d.year for d in dfA["Date"] if pd.notna(d)}) if "Date" in dfA.columns else []
        monthsA = [f"{m:02d}" for m in range(1,13)]
        visasA  = sorted(dfA["Visa"].dropna().astype(str).unique()) if "Visa" in dfA.columns else []
        sel_years  = c1.multiselect("Année", yearsA, default=[], key="anal_years")
        sel_months = c2.multiselect("Mois (MM)", monthsA, default=[], key="anal_months")
        sel_visa   = c3.multiselect("Type de visa", visasA, default=[], key="anal_visa")
        include_na_dates = c4.checkbox("Inclure lignes sans date", value=True, key="anal_na_dates")

    with st.container():
        d1, d2 = st.columns([1,1])
        today = date.today()
        if ("Date" in dfA.columns) and dfA["Date"].notna().any():
            dmin = min([d for d in dfA["Date"] if pd.notna(d)])
            dmax = max([d for d in dfA["Date"] if pd.notna(d)])
        else:
            dmin, dmax = today - timedelta(days=365), today
        date_from = d1.date_input("Du", value=dmin, key="anal_date_from")
        date_to   = d2.date_input("Au", value=dmax, key="anal_date_to")

        c3a, c3b = st.columns([1,1])
        agg_with_year = c3a.toggle("Agrégation par Année-Mois (YYYY-MM)", value=False, key="anal_agg_with_year")
        show_tables   = c3b.toggle("Voir les tableaux détaillés", value=False, key="anal_show_tables")

    fA = dfA.copy()
    if sel_visa:
        fA = fA[fA["Visa"].astype(str).isin(sel_visa)]
    if "Date" in fA.columns and sel_years:
        mask_year = fA["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask_year = mask_year | fA["Date"].isna()
        fA = fA[mask_year]
    if "Mois" in fA.columns and sel_months:
        mask_month = fA["Mois"].isin(sel_months)
        if include_na_dates: mask_month = mask_month | fA["Mois"].isna()
        fA = fA[mask_month]
    if "Date" in fA.columns and (date_from or date_to):
        mask_range = fA["Date"].apply(lambda x: pd.notna(x) and (x >= date_from) and (x <= date_to))
        if include_na_dates:
            mask_range = mask_range | fA["Date"].isna()
        fA = fA[mask_range]

    if agg_with_year:
        fA["Periode"] = fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")
        ordre_periodes = sorted([p for p in fA["Periode"].unique() if p != "NA"]) + (["NA"] if "NA" in fA["Periode"].values else [])
    else:
        fA["Periode"] = fA["Mois"].fillna("NA")
        ordre_periodes = [f"{m:02d}" for m in range(1,13)]
        if "NA" in fA["Periode"].values:
            ordre_periodes += ["NA"]

    for col in ["Montant","Payé","Reste"]:
        if col in fA.columns:
            fA[col] = pd.to_numeric(fA[col], errors="coerce").fillna(0.0)

    def derive_statut(row) -> str:
        if bool(row.get("Dossier approuvé", False)): return "Approuvé"
        if bool(row.get("Dossier refusé", False)):   return "Refusé"
        if bool(row.get("Dossier annulé", False)):   return "Annulé"
        return "En attente"

    details = fA.copy()
    details["Statut"] = details.apply(derive_statut, axis=1)
    details_display_cols = [c for c in ["Periode","ID_Client","Nom","Visa","Date","Montant","Payé","Reste","Statut"] if c in details.columns]

    # KPI
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers (filtrés)", f"{len(fA)}")
    k2.metric("Chiffre d’affaires", _fmt_money_us(float(fA.get("Montant", pd.Series(dtype=float)).sum())))
    k3.metric("Encaissements",       _fmt_money_us(float(fA.get("Payé",    pd.Series(dtype=float)).sum())))
    k4.metric("Solde à encaisser",   _fmt_money_us(float(fA.get("Reste",   pd.Series(dtype=float)).sum())))

    st.divider()

    # Volumes
    st.markdown("### 📦 Volumes par période")
    def _safe_bool_sum(series): return series.fillna(False).astype(bool).sum()
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
    else:
        cvol2.info("Aucune colonne de statut trouvée (Envoyé/Approuvé/Refusé/Annulé).")

    st.divider()

    # Financier
    st.markdown("### 💵 Financier par période")
    sums = fA.groupby("Periode")[["Montant","Payé","Reste"]].sum().reindex(ordre_periodes).fillna(0.0)
    cfin1, cfin2 = st.columns(2)
    cfin1.caption("Chiffre d'affaires (Montant)")
    cfin1.bar_chart(sums[["Montant"]])
    cfin2.caption("Encaissements (Payé) & Solde à encaisser (Reste)")
    cfin2.bar_chart(sums[["Payé","Reste"]])

    st.divider()

    # Détails clients + recherche
    st.markdown("### 🔎 Détails (clients)")
    d1, d2 = st.columns([1,1])
    statut_filter = d1.multiselect("Filtrer par statut", ["Approuvé","Refusé","Annulé","En attente"], key="anal_statut_filter")
    search = d2.text_input("Recherche (Nom / Visa / ID)", key="anal_search")

    details_to_show = details[details_display_cols].copy()
    if statut_filter:
        details_to_show = details_to_show[details_to_show["Statut"].isin(statut_filter)]
    if search:
        s = search.lower()
        def _match(row):
            return (s in str(row.get("Nom","")).lower()) or (s in str(row.get("Visa","")).lower()) or (s in str(row.get("ID_Client","")).lower())
        details_to_show = details_to_show[details_to_show.apply(_match, axis=1)]

    st.dataframe(details_to_show.sort_values(["Periode","Nom"]), use_container_width=True)

    with st.expander("Clients par période / par statut"):
        clients_par_periode = details.groupby("Periode")["Nom"].apply(lambda s: ", ".join(sorted(set(map(str, s.dropna())))))
        clients_par_statut  = details.groupby("Statut")["Nom"].apply(lambda s: ", ".join(sorted(set(map(str, s.dropna())))))
        st.write("**Clients par période**"); st.dataframe(clients_par_periode.to_frame("Clients"))
        st.write("**Clients par statut**");  st.dataframe(clients_par_statut.to_frame("Clients"))

    st.divider()

    # Fiche détaillée client (ajout de paiements aussi ici)
    st.markdown("### 🧾 Fiche détaillée — client")
    unique_clients = details_to_show.dropna(subset=["ID_Client"]).copy()
    unique_clients["_opt"] = unique_clients.apply(
        lambda r: f'{r.get("ID_Client","")} — {r.get("Nom","")} — {r.get("Visa","")} — {r.get("Date","")}', axis=1
    )
    opt = st.selectbox("Choisir un client", unique_clients["_opt"].tolist(), key="anal_client_select")
    if opt:
        sel_id = unique_clients.set_index("_opt").loc[opt, "ID_Client"]
        live_all_raw = read_sheet(current_path, client_target_sheet, normalize=False)
        live_all = normalize_dataframe(live_all_raw)
        idxs = live_all.index[live_all["ID_Client"].astype(str) == str(sel_id)]
        if len(idxs) == 0:
            st.info("Client introuvable dans la feuille cible.")
        else:
            idx = idxs[0]
            r0 = live_all.loc[idx].to_dict()

            cA, cB, cC = st.columns(3)
            try: m = float(r0.get("Montant", 0))
            except Exception: m = 0.0
            try: p = float(r0.get("Payé", 0))
            except Exception: p = 0.0
            rest = max(m - p, 0.0)
            cA.metric("Montant", _fmt_money_us(m)); cB.metric("Payé", _fmt_money_us(p)); cC.metric("Reste", _fmt_money_us(rest))

            def _b(v, lab):
                return f'<span style="padding:.2rem .5rem;border-radius:.6rem;border:1px solid #ddd;margin-right:.25rem">{lab}: {"✅" if v else "—"}</span>'
            st.markdown(
                _b(bool(r0.get("Dossier envoyé")),"Envoyé")+
                _b(bool(r0.get("Dossier approuvé")),"Approuvé")+
                _b(bool(r0.get("RFE")),"RFE")+
                _b(bool(r0.get("Dossier refusé")),"Refusé")+
                _b(bool(r0.get("Dossier annulé")),"Annulé"),
                unsafe_allow_html=True
            )

            st.markdown("#### 💳 Paiements")
            payments = _parse_paiements(r0.get("Paiements",""))
            if payments:
                pay_df = pd.DataFrame(payments)
                if "date" in pay_df.columns:
                    pay_df["date"] = pd.to_datetime(pay_df["date"], errors="coerce").dt.date
                    pay_df = pay_df.sort_values("date")
                st.dataframe(pay_df.rename(columns={"date":"Date","amount":"Montant ($)","mode":"Mode","note":"Note"}), use_container_width=True)
            else:
                st.info("Aucun paiement encore enregistré.")

            with st.form("add_payment_from_detail"):
                c1, c2, c3, c4 = st.columns([1,1,1,2])
                np_amount = c1.number_input("Montant ($)", min_value=0.0, step=10.0, format="%.2f", key="anal_cli_pay_amt")
                np_date   = c2.date_input("Date", value=date.today(), key="anal_cli_pay_date")
                np_mode   = c3.selectbox("Mode", ["CB","Chèque","Espèces","Virement","Autre"], key="anal_cli_pay_mode")
                np_note   = c4.text_input("Note", "", key="anal_cli_pay_note")
                ok_add = st.form_submit_button("💾 Ajouter le paiement (fiche)", type="primary")
            if ok_add:
                try:
                    live_write = read_sheet(current_path, client_target_sheet, normalize=False).copy()
                    for c in ["Paiements","Payé","Montant","Reste","ID_Client"]:
                        if c not in live_write.columns:
                            live_write[c] = "" if c=="Paiements" else 0.0
                    ids_series = live_write["ID_Client"].astype(str) if "ID_Client" in live_write.columns else pd.Series([""], index=live_write.index)
                    hit = live_write.index[ids_series == str(sel_id)]
                    if len(hit)==0:
                        hit = live_write.index[(live_write.get("Nom","").astype(str)==_safe_str(r0.get("Nom"))) &
                                               (live_write.get("Date","").astype(str)==_safe_str(r0.get("Date")))]
                    if len(hit)==0:
                        st.error("Impossible de localiser la ligne à mettre à jour.")
                    else:
                        i = hit[0]
                        pay_list = _parse_paiements(live_write.at[i, "Paiements"])
                        add = float(np_amount or 0.0)
                        if add <= 0:
                            st.warning("Le montant doit être > 0.")
                            st.stop()
                        try: m0 = float(live_write.at[i,"Montant"])
                        except Exception: m0 = 0.0
                        try: p0 = float(live_write.at[i,"Payé"])
                        except Exception: p0 = 0.0
                        reste0 = max(m0 - p0, 0.0)
                        if add > reste0 + 1e-9:
                            st.info(f"Le paiement dépasse le reste. Plafonné à {_fmt_money_us(reste0)}.")
                            add = reste0

                        pay_list.append({"date": str(np_date), "amount": float(add), "mode": np_mode, "note": np_note})
                        live_write.at[i, "Paiements"] = json.dumps(pay_list, ensure_ascii=False)
                        total_paid = _sum_payments(pay_list)
                        live_write.at[i, "Payé"]  = float(total_paid)
                        live_write.at[i, "Reste"] = max(m0 - float(total_paid), 0.0)

                        write_sheet_inplace(current_path, client_target_sheet, live_write)
                        st.success("Paiement ajouté **dans le fichier** depuis la fiche. ✅")
                        st.rerun()
                except Exception as e:
                    st.error(f"Erreur: {e}")

    st.divider()

    # Export Analyses
    st.markdown("### 📤 Export Excel")
    if st.button("Exporter vers l'Excel (feuille 'Analyses')", key="anal_export_btn"):
        try:
            blocks = [("KPI Globaux (après filtres)",
                       pd.DataFrame({
                           "KPI": ["Dossiers (filtrés)", "Chiffre d’affaires", "Encaissements", "Solde à encaisser"],
                           "Valeur": [
                               len(fA),
                               float(fA.get("Montant", pd.Series(dtype=float)).sum()),
                               float(fA.get("Payé",    pd.Series(dtype=float)).sum()),
                               float(fA.get("Reste",   pd.Series(dtype=float)).sum()),
                           ],
                       }))
            ]
            blocks.append(("Volumes — Dossiers ouverts par période", st_data1))
            if vols_dict:
                blocks.append(("Volumes — Statuts par période", st_data2))
            blocks.append(("Financier — Montant / Payé / Reste par période", sums))
            blocks.append(("Détails (clients)", details_to_show))
            blocks.append(("Clients par période (liste)", clients_par_periode.to_frame("Clients")))
            blocks.append(("Clients par statut (liste)", clients_par_statut.to_frame("Clients")))

            write_analyses_sheet(current_path, blocks)
            save_workspace_path(current_path)
            st.success("Feuille **'Analyses'** exportée dans le fichier. ✅ (Téléchargement → sidebar)")
        except Exception as e:
            st.error(f"Échec export : {e}")

    st.markdown("### 📥 Export CSV (ZIP)")
    if st.button("Préparer le ZIP des analyses (CSV)", key="anal_zip_btn"):
        try:
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                kpi_df = pd.DataFrame({
                    "KPI": ["Dossiers (filtrés)", "Chiffre d’affaires", "Encaissements", "Solde à encaisser"],
                    "Valeur": [
                        len(fA),
                        float(fA.get("Montant", pd.Series(dtype=float)).sum()),
                        float(fA.get("Payé",    pd.Series(dtype=float)).sum()),
                        float(fA.get("Reste",   pd.Series(dtype=float)).sum()),
                    ],
                })
                zf.writestr("kpi_globaux.csv", kpi_df.to_csv(index=False, encoding="utf-8-sig"))
                zf.writestr("volumes_ouverts_par_periode.csv", st_data1.to_csv(encoding="utf-8-sig"))
                if vols_dict:
                    zf.writestr("volumes_statuts_par_periode.csv", st_data2.to_csv(encoding="utf-8-sig"))
                zf.writestr("financier_par_periode.csv", sums.to_csv(encoding="utf-8-sig"))
                zf.writestr("details_clients.csv", details_to_show.to_csv(index=False, encoding="utf-8-sig"))
                zf.writestr("clients_par_periode.csv", clients_par_periode.to_frame("Clients").to_csv(encoding="utf-8-sig"))
                zf.writestr("clients_par_statut.csv", clients_par_statut.to_frame("Clients").to_csv(encoding="utf-8-sig"))
                zf.writestr("dataset_filtre_complet.csv", fA.to_csv(index=False, encoding="utf-8-sig"))
            mem.seek(0)
            st.session_state["anal_zip_bytes"] = mem.read()
            st.success("ZIP des analyses prêt. Utilise le bouton ci-dessous pour télécharger.")
        except Exception as e:
            st.error(f"Échec création ZIP : {e}")

    if "anal_zip_bytes" in st.session_state:
        st.download_button(
            "⬇️ Télécharger le ZIP des analyses (CSV)",
            data=st.session_state["anal_zip_bytes"],
            file_name="analyses_csv.zip",
            mime="application/zip",
            key="anal_zip_download"
        )
