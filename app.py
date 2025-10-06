# app.py
import io, json, hashlib
from datetime import date, datetime, timedelta
from pathlib import Path
import streamlit as st
import pandas as pd
import altair as alt

# ============ Config Altair / Streamlit ============
alt.data_transformers.disable_max_rows()
alt.renderers.set_embed_options(actions=False)
st.set_page_config(page_title="📊 Visas — Edition directe + ESCROW + Analyses", layout="wide")
st.title("📊 Visas — Edition DIRECTE du fichier (ESCROW + Analyses)")

# ============ Constantes colonnes ============
HONO   = "Honoraires (US $)"
AUTRE  = "Autres frais (US $)"
TOTAL  = "Total (US $)"
ESC_TR = "Escrow transféré (US $)"
ESC_JR = "Escrow journal"
DOSSIER_COL = "Dossier N"
DOSSIER_START = 13057

# ============ Workspace (mémoriser le dernier fichier) ============
def pick_workdir() -> Path | None:
    for p in [Path("/mnt/data"), Path("/tmp/visa_workspace"), Path.cwd() / "visa_workspace"]:
        try:
            p.mkdir(parents=True, exist_ok=True)
            t = p / ".write_test"; t.write_text("ok", encoding="utf-8"); t.unlink(missing_ok=True)
            return p
        except Exception:
            continue
    return None

WORK_DIR = pick_workdir()
WS_FILE  = (WORK_DIR / "_workspace.json") if WORK_DIR else None

def load_workspace_path() -> Path | None:
    if WS_FILE is None or not WS_FILE.exists():
        return None
    try:
        data = json.loads(WS_FILE.read_text(encoding="utf-8"))
        p = Path(data.get("last_path",""))
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
    dest.write_bytes(upload.read())
    return dest

# ============ Utils ============
def _safe_str(x): return "" if pd.isna(x) else str(x).strip()

def _to_num(s: pd.Series) -> pd.Series:
    cleaned = (s.astype(str)
                 .str.replace("\u00a0","",regex=False)
                 .str.replace("\u202f","",regex=False)
                 .str.replace(" ","",regex=False)
                 .str.replace("$","",regex=False)
                 .str.replace(",","",regex=False))
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)

def _to_int(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

def _to_date(s: pd.Series) -> pd.Series:
    d = pd.to_datetime(s, errors="coerce")
    try: d = d.dt.tz_localize(None)
    except Exception: pass
    return d.dt.normalize().dt.date

def _fmt_money_us(v: float) -> str:
    try: return f"${float(v):,.2f}"
    except Exception: return "$0.00"

def _parse_json_list(x):
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
    # Référentiel si Visa (et potentiellement Catégorie) mais pas de colonnes financières
    return ("visa" in cols) and not ({"montant","honoraires","payé","reste","solde"} & cols)

def is_clients_like(df: pd.DataFrame) -> bool:
    cols = set(df.columns.astype(str))
    return {"Nom","Visa"}.issubset(cols)

def _clean_for_chart(df: pd.DataFrame, str_cols, num_cols, drop_na_cols):
    df2 = df.copy()
    for c in str_cols:
        if c in df2.columns:
            df2[c] = df2[c].astype(str).fillna("")
    for c in num_cols:
        if c in df2.columns:
            df2[c] = pd.to_numeric(df2[c], errors="coerce").astype(float)
    keep = pd.Series(True, index=df2.index)
    for c in drop_na_cols:
        if c in df2.columns:
            keep &= df2[c].notna()
    return df2[keep]

# ============ ESCROW helpers ============
def escrow_available_from_row(row) -> float:
    try: hon = float(row.get(HONO, 0.0))
    except Exception: hon = 0.0
    try: paid = float(row.get("Payé", 0.0))
    except Exception: paid = 0.0
    try: moved = float(row.get(ESC_TR, 0.0))
    except Exception: moved = 0.0
    return max(min(paid, hon) - moved, 0.0)

def append_escrow_journal(row_raw: pd.Series, amount: float, note: str = "") -> str:
    journal = _parse_json_list(row_raw.get(ESC_JR, ""))
    journal.append({"ts": datetime.now().isoformat(timespec="seconds"),
                    "amount": float(amount), "note": note})
    return json.dumps(journal, ensure_ascii=False)

# ============ Numérotation Dossier N ============
def ensure_dossier_numbers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if DOSSIER_COL not in df.columns:
        start = DOSSIER_START
        df[DOSSIER_COL] = list(range(start, start + len(df)))
        return df
    s = _to_int(df[DOSSIER_COL]) if not df[DOSSIER_COL].empty else pd.Series([], dtype=int)
    df[DOSSIER_COL] = s
    existing = s[s > 0]
    next_num = int(existing.max() + 1) if not existing.empty else DOSSIER_START
    to_fill_idx = df.index[df[DOSSIER_COL] <= 0].tolist()
    for i in to_fill_idx:
        df.at[i, DOSSIER_COL] = next_num
        next_num += 1
    return df

def next_dossier_number(df_existing: pd.DataFrame) -> int:
    if DOSSIER_COL not in df_existing.columns or df_existing.empty:
        return DOSSIER_START
    s = _to_int(df_existing[DOSSIER_COL]); s = s[s > 0]
    return int(s.max() + 1) if not s.empty else DOSSIER_START

# ============ Référentiel (Visa ⇢ Catégorie) ============
def read_visa_reference(path: Path) -> pd.DataFrame:
    try:
        dfv = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=["Catégorie","Visa"])
    # normaliser
    if "Visa" not in dfv.columns:
        visa_col = next((c for c in dfv.columns if str(c).strip().lower()=="visa"), None)
        if visa_col: dfv = dfv.rename(columns={visa_col:"Visa"})
    cat_col = next((c for c in dfv.columns if str(c).strip().lower() in ("catégorie","categorie")), None)
    if not cat_col:
        dfv["Catégorie"] = ""
    elif cat_col != "Catégorie":
        dfv = dfv.rename(columns={cat_col:"Catégorie"})
    # typage
    for c in ["Catégorie","Visa"]:
        if c in dfv.columns: dfv[c] = dfv[c].astype(str).fillna("").str.strip()
    return dfv[["Catégorie","Visa"]]

def map_category_from_ref(visalib: str, ref_df: pd.DataFrame) -> str:
    if ref_df is None or ref_df.empty: return ""
    tmp = ref_df[ref_df["Visa"].astype(str).str.strip().str.lower() == _safe_str(visalib).lower()]
    if tmp.empty: return ""
    return _safe_str(tmp.iloc[0]["Catégorie"])

# ============ Normalisation ============
def normalize_dataframe(df: pd.DataFrame, visa_ref: pd.DataFrame | None = None) -> pd.DataFrame:
    df = df.copy()
    # Date / Mois
    df["Date"] = _to_date(df["Date"]) if "Date" in df.columns else pd.NaT
    df["Mois"] = df["Date"].apply(lambda x: f"{x.month:02d}" if pd.notna(x) else pd.NA)

    # Visa & Catégorie
    visa_col = next((c for c in ["Visa","Categories","Catégorie","Categorie","TypeVisa"] if c in df.columns), None)
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"
    if "Catégorie" not in df.columns: df["Catégorie"] = ""
    if visa_ref is not None and not visa_ref.empty:
        df["Catégorie"] = df.apply(lambda r: _safe_str(r.get("Catégorie")) or map_category_from_ref(_safe_str(r.get("Visa")), visa_ref), axis=1)

    # Honoraires alias
    hono_aliases = [HONO, "Honoraires", "Honoraires US $", "Montant honoraires us $", "Montant honoraires",
                    "Montant (US $)", "Montant"]
    hono_src = next((c for c in hono_aliases if c in df.columns), None)
    df[HONO] = _to_num(df[hono_src]) if hono_src else 0.0

    # Autres frais alias
    autre_aliases = [AUTRE, "Autres frais", "Frais", "Other fees", "Autres"]
    autre_src = next((c for c in autre_aliases if c in df.columns), None)
    df[AUTRE] = _to_num(df[autre_src]) if autre_src else 0.0

    # Paiements / Acomptes
    if "Paiements" in df.columns and df["Paiements"].astype(str).str.strip().ne("").any():
        parsed = df["Paiements"].apply(_parse_json_list)
        df["Payé"] = parsed.apply(_sum_payments).astype(float)
    else:
        acompte_cols = [c for c in df.columns if str(c).strip().lower().startswith("acompte")]
        if acompte_cols:
            df["Payé"] = _to_num(df[acompte_cols].fillna(0).sum(axis=1))
            def _build_paiements_json(row):
                entries = []
                for c in acompte_cols:
                    val = _to_num(pd.Series([row.get(c, 0)])).iloc[0]
                    if val > 0:
                        entries.append({"date": str(row.get("Date","")), "amount": float(val), "mode": "", "note": c})
                return json.dumps(entries, ensure_ascii=False) if entries else ""
            df["Paiements"] = df.apply(_build_paiements_json, axis=1)
        else:
            df["Payé"] = _to_num(df["Payé"]) if "Payé" in df.columns else 0.0

    # Totaux / Reste
    df[TOTAL] = (df[HONO] + df[AUTRE]).astype(float)
    df["Reste"] = (df[TOTAL] - df["Payé"]).fillna(0.0)

    # Statuts
    for b in ["RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"]:
        if b not in df.columns: df[b] = False

    # Identité
    if "Nom" not in df.columns: df["Nom"] = ""
    if "ID_Client" not in df.columns: df["ID_Client"] = ""
    need_id = df["ID_Client"].astype(str).str.strip().eq("") | df["ID_Client"].isna()
    if need_id.any(): df.loc[need_id,"ID_Client"] = df.loc[need_id].apply(_make_client_id_from_row, axis=1)

    if "Paiements" not in df.columns: df["Paiements"] = ""

    # ESCROW
    df[ESC_TR] = _to_num(df[ESC_TR]) if ESC_TR in df.columns else 0.0
    if ESC_JR not in df.columns: df[ESC_JR] = ""

    # Dossier N
    df = ensure_dossier_numbers(df)

    # Nettoyage (on retire Téléphone/Email si présents)
    for dropcol in ["Telephone","Email"]:
        if dropcol in df.columns: df = df.drop(columns=[dropcol])

    ordered = [DOSSIER_COL,"ID_Client","Nom","Date","Mois","Catégorie","Visa",
               HONO, AUTRE, TOTAL, "Payé","Reste",
               ESC_TR, ESC_JR,
               "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé","Paiements"]
    cols = [c for c in ordered if c in df.columns] + [c for c in df.columns if c not in ordered]
    return df[cols]

# ============ IO Excel ============
def read_sheet(path: Path, sheet: str, normalize: bool, visa_ref: pd.DataFrame | None = None) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    if sheet not in xls.sheet_names:
        base = pd.DataFrame(columns=[DOSSIER_COL,"ID_Client","Nom","Date","Mois","Catégorie","Visa",
                                     HONO, AUTRE, TOTAL, "Payé","Reste",
                                     ESC_TR, ESC_JR,
                                     "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé","Paiements"])
        return normalize_dataframe(base, visa_ref=visa_ref) if normalize else base
    df = pd.read_excel(xls, sheet_name=sheet)
    return normalize_dataframe(df, visa_ref=visa_ref) if (normalize and not looks_like_reference(df)) else df

def write_sheet_inplace(path: Path, sheet_to_replace: str, new_df: pd.DataFrame):
    xls = pd.ExcelFile(path)
    out = io.BytesIO()
    target_written = False
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name in xls.sheet_names:
            if name == sheet_to_replace:
                dfw = new_df.copy()
                for c in dfw.columns:
                    if dfw[c].dtype == "object": dfw[c] = dfw[c].astype(str).fillna("")
                dfw.to_excel(writer, sheet_name=name, index=False); target_written = True
            else:
                pd.read_excel(xls, sheet_name=name).to_excel(writer, sheet_name=name, index=False)
        if not target_written:
            dfw = new_df.copy()
            for c in dfw.columns:
                if dfw[c].dtype == "object": dfw[c] = dfw[c].astype(str).fillna("")
            dfw.to_excel(writer, sheet_name=sheet_to_replace, index=False)
    bytes_out = out.getvalue()
    path.write_bytes(bytes_out)
    try:
        st.session_state["download_bytes"] = bytes_out
        st.session_state["download_name"] = path.name
    except Exception:
        pass

# ============ Source (sidebar) ============
st.sidebar.header("Source")

def _find_latest_xlsx(paths: list[Path]) -> Path | None:
    cand = []
    for base in paths:
        try:
            if base and base.exists():
                cand.extend([p for p in base.glob("*.xlsx") if p.is_file()])
        except Exception:
            pass
    if not cand: return None
    return sorted(cand, key=lambda p: p.stat().st_mtime, reverse=True)[0]

current_path = load_workspace_path()
if (current_path is None) or (not current_path.exists()):
    search_dirs = [WORK_DIR] if WORK_DIR else []
    if Path("/mnt/data").exists(): search_dirs.append(Path("/mnt/data"))
    latest = _find_latest_xlsx(search_dirs)
    if latest:
        current_path = latest
        save_workspace_path(current_path)

if (current_path is None) or (not current_path.exists()):
    defaults = [
        Path("/mnt/data/donnees_visa_clients.xlsx"),
        Path("/mnt/data/modele_clients_visa.xlsx"),
        Path("/mnt/data/Visa_Clients_20251001-114844.xlsx"),
        Path("/mnt/data/visa_analytics_datecol.xlsx"),
    ]
    current_path = next((p for p in defaults if p.exists()), None)
    if current_path:
        save_workspace_path(current_path)

if current_path and current_path.exists():
    st.sidebar.success(f"Fichier courant : {current_path.name}")
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

# Sync bouton de téléchargement avec le fichier courant
if ("download_bytes" not in st.session_state) or (st.session_state.get("download_name") != current_path.name):
    try:
        st.session_state["download_bytes"] = current_path.read_bytes()
        st.session_state["download_name"] = current_path.name
    except Exception:
        st.session_state["download_bytes"] = b""
        st.session_state["download_name"] = current_path.name

try:
    sheet_names = pd.ExcelFile(current_path).sheet_names
except Exception as e:
    st.error(f"Impossible de lire l'Excel : {e}")
    st.stop()

preferred_order = ["Clients","Visa","Données normalisées"]
default_sheet = next((s for s in preferred_order if s in sheet_names), sheet_names[0])
sheet_choice = st.sidebar.selectbox("Feuille (Dashboard)", sheet_names, index=sheet_names.index(default_sheet))

# Détection de la feuille *Clients* (cible CRUD)
valid_client_sheets = []
for s in sheet_names:
    try:
        df_tmp = read_sheet(current_path, s, normalize=False)
        if is_clients_like(df_tmp):
            valid_client_sheets.append(s)
    except Exception:
        pass

if not valid_client_sheets:
    st.sidebar.error("Aucune feuille 'clients' valide (au minimum Nom & Visa).")
    client_target_sheet = None
else:
    default_client_sheet = "Clients" if "Clients" in valid_client_sheets else valid_client_sheets[0]
    if "client_sheet_select" in st.session_state and st.session_state["client_sheet_select"] not in valid_client_sheets:
        del st.session_state["client_sheet_select"]
    client_target_sheet = st.sidebar.selectbox(
        "Feuille *Clients* (cible CRUD)",
        valid_client_sheets,
        index=valid_client_sheets.index(default_client_sheet),
        key="client_sheet_select"
    )

st.sidebar.caption(f"Édition **directe** dans : `{current_path}`")
st.sidebar.download_button(
    "⬇️ Télécharger une copie",
    data=st.session_state.get("download_bytes", b""),
    file_name=st.session_state.get("download_name", current_path.name),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ============ Onglets ============
tabs = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"])

# ================= DASHBOARD =================
with tabs[0]:
    visa_ref = read_visa_reference(current_path)
    df_raw = read_sheet(current_path, sheet_choice, normalize=False)

    # Référentiel Visa (Catégorie + Visa)
    if looks_like_reference(df_raw) and sheet_choice == "Visa":
        st.subheader("📄 Référentiel — Catégories & Types de Visa")
        if "Visa" not in df_raw.columns:
            df_ref = pd.DataFrame(columns=["Catégorie","Visa"])
        else:
            df_ref = df_raw.copy()
            if "Catégorie" not in df_ref.columns:
                df_ref["Catégorie"] = ""
            df_ref = df_ref[["Catégorie","Visa"]].astype(str).fillna("").applymap(lambda s: s.strip())

        c1, c2 = st.columns(2)
        cats = sorted([c for c in df_ref["Catégorie"].astype(str).unique() if c!=""])
        visas = sorted([v for v in df_ref["Visa"].astype(str).unique() if v!=""])
        f_cat = c1.multiselect("Filtrer Catégorie", cats, default=[])
        f_vis = c2.multiselect("Filtrer Visa", visas, default=[])
        view = df_ref.copy()
        if f_cat: view = view[view["Catégorie"].isin(f_cat)]
        if f_vis: view = view[view["Visa"].isin(f_vis)]
        st.dataframe(view, use_container_width=True)

        st.markdown("### ✏️ Gérer le référentiel")
        mode = st.radio("Action", ["Ajouter", "Renommer", "Supprimer"], horizontal=True, key="visa_ref_action")
        options = df_ref.assign(_label=df_ref["Catégorie"].str.cat(df_ref["Visa"], sep=" — "))

        if mode == "Ajouter":
            cA, cB = st.columns(2)
            new_cat = cA.text_input("Catégorie (ex: Étudiant)").strip()
            new_vis = cB.text_input("Visa (ex: F1)").strip()
            if st.button("➕ Ajouter"):
                if not new_vis:
                    st.warning("Le champ Visa est requis.")
                else:
                    dup = ((df_ref["Catégorie"].str.lower()==new_cat.lower()) & (df_ref["Visa"].str.lower()==new_vis.lower()))
                    if dup.any():
                        st.info("Cette paire Catégorie/Visa existe déjà.")
                    else:
                        out = pd.concat([df_ref, pd.DataFrame([{"Catégorie": new_cat, "Visa": new_vis}])], ignore_index=True)
                        write_sheet_inplace(current_path, "Visa", out); st.success("Ajouté."); st.rerun()

        elif mode == "Renommer":
            if options.empty:
                st.info("Aucune entrée.")
            else:
                sel_lab = st.selectbox("Sélection à renommer", options["_label"].tolist())
                row = options.loc[options["_label"]==sel_lab].iloc[0]
                cA, cB = st.columns(2)
                new_cat = cA.text_input("Nouvelle catégorie", value=row["Catégorie"]).strip()
                new_vis = cB.text_input("Nouveau visa", value=row["Visa"]).strip()
                if st.button("📝 Renommer"):
                    if not new_vis:
                        st.warning("Le champ Visa est requis.")
                    else:
                        out = df_ref.copy()
                        mask = (out["Catégorie"]==row["Catégorie"]) & (out["Visa"]==row["Visa"])
                        out.loc[mask, ["Catégorie","Visa"]] = [new_cat, new_vis]
                        write_sheet_inplace(current_path, "Visa", out); st.success("Renommé."); st.rerun()

        else:  # Supprimer
            if options.empty:
                st.info("Aucune entrée à supprimer.")
            else:
                sel_lab = st.selectbox("Sélection à supprimer", options["_label"].tolist())
                st.error("⚠️ Action irréversible.")
                if st.button("🗑️ Supprimer"):
                    out = df_ref[options["_label"]!=sel_lab].reset_index(drop=True)
                    write_sheet_inplace(current_path, "Visa", out); st.success("Supprimé."); st.rerun()
        st.stop()

    # Sinon : données -> normaliser (avec référentiel)
    df = read_sheet(current_path, sheet_choice, normalize=True, visa_ref=read_visa_reference(current_path))

    # Filtres Catégorie & Visa & Date
    with st.container():
        c1, c2, c3 = st.columns(3)
        cats = sorted(df["Catégorie"].dropna().astype(str).unique()) if "Catégorie" in df.columns else []
        visas = sorted(df["Visa"].dropna().astype(str).unique()) if "Visa" in df.columns else []
        sel_cats  = c1.multiselect("Catégorie", cats, default=[])
        sel_visas = c2.multiselect("Type de visa", visas, default=[])
        years = sorted({d.year for d in df["Date"] if pd.notna(d)}) if "Date" in df.columns else []
        sel_years = c3.multiselect("Année", years, default=[])
        d1, d2, d3 = st.columns(3)
        months = sorted(df["Mois"].dropna().unique()) if "Mois" in df.columns else []
        sel_months = d1.multiselect("Mois (MM)", months, default=[])
        include_na_dates = d2.checkbox("Inclure lignes sans date", value=True)

        def make_slider(_df, col, lab, container):
            if col not in _df.columns or _df[col].dropna().empty:
                container.caption(f"{lab} : aucune donnée"); return None
            vmin, vmax = float(_df[col].min()), float(_df[col].max())
            if not (vmin < vmax):
                container.caption(f"{lab} : valeur unique = {_fmt_money_us(vmin)}"); return (vmin, vmax)
            step = 1.0 if (vmax - vmin) > 1000 else 0.1 if (vmax - vmin) > 10 else 0.01
            return container.slider(lab, min_value=vmin, max_value=vmax, value=(vmin, vmax), step=step)
        total_range = make_slider(df, TOTAL, "Total (US $) min-max", d3)

    f = df.copy()
    if "Catégorie" in f.columns and sel_cats:  f = f[f["Catégorie"].astype(str).isin(sel_cats)]
    if "Visa" in f.columns and sel_visas:      f = f[f["Visa"].astype(str).isin(sel_visas)]
    if "Date" in f.columns and sel_years:
        mask = f["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask |= f["Date"].isna()
        f = f[mask]
    if "Mois" in f.columns and sel_months:
        mask = f["Mois"].isin(sel_months)
        if include_na_dates: mask |= f["Mois"].isna()
        f = f[mask]
    if TOTAL in f.columns and total_range is not None:
        f = f[(f[TOTAL] >= total_range[0]) & (f[TOTAL] <= total_range[1])]

    hidden = len(df) - len(f)
    if hidden > 0: st.caption(f"🔎 {hidden} ligne(s) masquée(s) par les filtres.")

    # KPI
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(f)}")
    k2.metric("Total (US $)", _fmt_money_us(float(f.get(TOTAL, pd.Series(dtype=float)).sum())) )
    k3.metric("Payé (US $)", _fmt_money_us(float(f.get("Payé", pd.Series(dtype=float)).sum())) )
    k4.metric("Solde (US $)", _fmt_money_us(float(f.get("Reste", pd.Series(dtype=float)).sum())) )
    st.markdown('</div>', unsafe_allow_html=True)

    # Alerte ESCROW (dossiers envoyés avec solde ESCROW à réclamer)
    df_esc = f.copy()
    if ESC_TR not in df_esc.columns: df_esc[ESC_TR] = 0.0
    else: df_esc[ESC_TR] = pd.to_numeric(df_esc[ESC_TR], errors="coerce").fillna(0.0)
    df_esc["escrow_dispo"] = df_esc.apply(escrow_available_from_row, axis=1)
    alert = df_esc[(df_esc.get("Dossier envoyé", False) == True) & (df_esc["escrow_dispo"] > 0.004)]
    if not alert.empty:
        total_alert = float(alert["escrow_dispo"].sum())
        st.warning(f"💡 ESCROW à réclamer : {_fmt_money_us(total_alert)} sur {len(alert)} dossier(s) **envoyés**.")
        st.dataframe(alert[[DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Date",HONO,"Payé",ESC_TR,"escrow_dispo"]].assign(
            escrow_dispo=alert["escrow_dispo"].map(_fmt_money_us),
            **{HONO: alert[HONO].map(_fmt_money_us),
               "Payé": alert["Payé"].map(_fmt_money_us),
               ESC_TR: alert[ESC_TR].map(_fmt_money_us)}
        ), use_container_width=True)

    st.divider()
    st.subheader("📋 Données (aperçu)")
    cols_show = [c for c in [DOSSIER_COL,"ID_Client","Nom","Date","Catégorie","Visa", HONO, AUTRE, TOTAL, "Payé","Reste",
                             "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"] if c in f.columns]
    table = f.copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in table.columns: table[col] = table[col].map(_fmt_money_us)
    if "Date" in table.columns: table["Date"] = table["Date"].astype(str)
    st.dataframe(table[cols_show].sort_values(by=[c for c in ["Date","Visa"] if c in table.columns], na_position="last"),
                 use_container_width=True)

    st.divider()
    # Ajout rapide d'un paiement (dossiers non soldés)
    st.subheader("➕ Ajouter un paiement (US $)")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide dans la sidebar.")
    else:
        clients_norm = read_sheet(current_path, client_target_sheet, normalize=True, visa_ref=read_visa_reference(current_path))
        todo = clients_norm[clients_norm["Reste"] > 0.004].copy() if "Reste" in clients_norm.columns else pd.DataFrame()
        if todo.empty:
            st.success("Tous les dossiers sont soldés ✅")
        else:
            todo["_label"] = todo.apply(lambda r: f'{r.get("ID_Client","")} — {r.get("Nom","")} — Reste {_fmt_money_us(float(r.get("Reste",0)))}', axis=1)
            label_to_id = todo.set_index("_label")["ID_Client"].to_dict()
            csel, camt, cdate, cmode = st.columns([2,1,1,1])
            sel_label = csel.selectbox("Dossier à créditer", todo["_label"].tolist(), key="quick_pay_sel")
            amount = camt.number_input("Montant ($)", min_value=0.0, step=10.0, format="%.2f", key="quick_pay_amt")
            pdate  = cdate.date_input("Date", value=date.today(), key="quick_pay_date")
            mode   = cmode.selectbox("Mode", ["CB","Chèque","Espèces","Virement","Venmo","Autre"], key="quick_pay_mode")
            note   = st.text_input("Note (facultatif)", "", key="quick_pay_note")
            if st.button("💾 Ajouter le paiement (écrit dans le fichier)", key="quick_pay_btn"):
                try:
                    live = read_sheet(current_path, client_target_sheet, normalize=False)
                    if "Paiements" not in live.columns: live["Paiements"] = ""
                    target_id = label_to_id.get(sel_label, "")
                    idxs = live.index[live.get("ID_Client","").astype(str) == str(target_id)]
                    if len(idxs)==0: raise RuntimeError("Dossier introuvable.")
                    idx = idxs[0]
                    pay_list = _parse_json_list(live.at[idx, "Paiements"])
                    add = float(amount or 0.0)
                    if add <= 0: st.warning("Le montant doit être > 0."); st.stop()
                    live_norm = normalize_dataframe(live.copy(), visa_ref=read_visa_reference(current_path))
                    mask = live_norm["ID_Client"].astype(str) == str(live.at[idx, "ID_Client"])
                    reste_curr = float(live_norm.loc[mask, "Reste"].sum()) if mask.any() else 0.0
                    if add > reste_curr + 1e-9: add = reste_curr
                    pay_list.append({"date": str(pdate), "amount": float(add), "mode": mode, "note": note})
                    live.at[idx, "Paiements"] = json.dumps(pay_list, ensure_ascii=False)
                    for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                        if c not in live.columns: live[c] = 0.0
                    total_paid = _sum_payments(pay_list)
                    hono = _to_num(pd.Series([live.at[idx, HONO]])).iloc[0] if HONO in live.columns else 0.0
                    autr = _to_num(pd.Series([live.at[idx, AUTRE]])).iloc[0] if AUTRE in live.columns else 0.0
                    total = float(hono + autr)
                    live.at[idx, "Payé"]  = float(total_paid)
                    live.at[idx, "Reste"] = max(total - float(total_paid), 0.0)
                    live.at[idx, TOTAL]   = total
                    write_sheet_inplace(current_path, client_target_sheet, live)
                    st.success("Paiement enregistré **dans le fichier**. ✅"); st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")

# ================= CLIENTS (CRUD) =================
with tabs[1]:
    st.subheader("👤 Clients — Créer / Modifier / Supprimer (écriture **directe**)")
    if client_target_sheet is None:
        st.warning("Aucune feuille *Clients* valide disponible. Ajoute une feuille avec au moins les colonnes **Nom** et **Visa**."); st.stop()
    if st.button("🔄 Recharger le fichier", key="reload_btn"):
        st.rerun()

    visa_ref = read_visa_reference(current_path)
    live_raw = read_sheet(current_path, client_target_sheet, normalize=False).copy()
    live_raw = ensure_dossier_numbers(live_raw)
    live_raw["_RowID"] = range(len(live_raw))

    cats_ref  = sorted([c for c in visa_ref["Catégorie"].astype(str).unique() if c!=""]) if not visa_ref.empty else []
    visas_all = sorted(visa_ref["Visa"].astype(str).unique()) if not visa_ref.empty else []

    has_envoye  = "Dossier envoyé"  in live_raw.columns
    has_appr    = "Dossier approuvé" in live_raw.columns
    has_rfe     = "RFE"             in live_raw.columns
    has_refuse  = "Dossier refusé"  in live_raw.columns
    has_annule  = "Dossier annulé"  in live_raw.columns

    action = st.radio("Action", ["Créer", "Modifier", "Supprimer"], horizontal=True, key="crud_action")

    # --- CREER ---
    if action == "Créer":
        st.markdown("### ➕ Nouveau client")
        # s'assurer des colonnes nécessaires
        for must in [DOSSIER_COL,"ID_Client","Nom","Date","Mois","Catégorie","Visa",
                     HONO, AUTRE, TOTAL, "Payé","Reste", ESC_TR, ESC_JR,
                     "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé","Paiements"]:
            if must not in live_raw.columns:
                if must in {HONO, AUTRE, TOTAL, "Payé","Reste", ESC_TR}: live_raw[must]=0.0
                elif must in {"Paiements", ESC_JR}: live_raw[must]=""
                elif must in {"Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"}: live_raw[must]=False
                elif must=="Mois": live_raw[must]=""
                elif must in {"Catégorie"}: live_raw[must]=""
                elif must==DOSSIER_COL: live_raw[must]=0
                else: live_raw[must]=""

        next_num = next_dossier_number(live_raw)
        with st.form("create_form", clear_on_submit=False):
            c0, c1, c2 = st.columns([1,1,1])
            c0.metric("Prochain Dossier N", f"{next_num}")
            nom_in = c1.text_input("Nom")
            d = c2.date_input("Date", value=date.today())

            cC, cV = st.columns(2)
            sel_cat = cC.selectbox("Catégorie", [""] + cats_ref, index=0, key="create_cat")
            if sel_cat:
                visas_opt = sorted(visa_ref.loc[visa_ref["Catégorie"]==sel_cat, "Visa"].unique().tolist())
            else:
                visas_opt = visas_all
            if visas_opt:
                visa = cV.selectbox("Visa", visas_opt, key="create_visa")
            else:
                visa = cV.text_input("Visa", key="create_visa_txt")

            c5,c6 = st.columns(2)
            honoraires = c5.number_input("Montant honoraires (US $)", value=0.0, step=10.0, format="%.2f", key="create_hono")
            autres     = c6.number_input("Autres frais (US $)", value=0.0, step=10.0, format="%.2f", key="create_autre")
            c7,c8 = st.columns(2)
            total_preview = float(honoraires + autres); c7.metric("Total (US $)", _fmt_money_us(total_preview))
            paye_init = c8.number_input("Payé (US $)", value=0.0, step=10.0, format="%.2f", key="create_paye")

            st.markdown("#### État du dossier")
            val_envoye = st.checkbox("Dossier envoyé",  value=False, key="create_env") if has_envoye else False
            val_appr   = st.checkbox("Dossier approuvé",value=False, key="create_app") if has_appr   else False
            val_rfe    = st.checkbox("RFE",             value=False, key="create_rfe") if has_rfe    else False
            val_refuse = st.checkbox("Dossier refusé",  value=False, key="create_ref") if has_refuse else False
            val_annule = st.checkbox("Dossier annulé",  value=False, key="create_ann") if has_annule else False

            ok = st.form_submit_button("💾 Sauvegarder (dans le fichier)", type="primary", help="Ajoute la ligne directement dans l'Excel")
        if ok:
            if val_rfe and not (val_envoye or val_refuse or val_annule):
                st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché."); st.stop()

            # Nom dupliqué -> suffixes -0, -1, ...
            existing_names = set(live_raw["Nom"].dropna().astype(str))
            base_name = _safe_str(nom_in)
            use_name = base_name
            if base_name in existing_names:
                k = 0
                while f"{base_name}-{k}" in existing_names:
                    k += 1
                use_name = f"{base_name}-{k}"

            gen_id = _make_client_id_from_row({"Nom": use_name, "Date": d})
            existing_ids = set(live_raw["ID_Client"].astype(str)) if "ID_Client" in live_raw.columns else set()
            new_id = gen_id; n=1
            while new_id in existing_ids:
                n+=1; new_id=f"{gen_id}-{n:02d}"

            total = float((honoraires or 0.0)+(autres or 0.0))
            reste = max(total - float(paye_init or 0.0), 0.0)

            new_row = {
                DOSSIER_COL: int(next_num),
                "ID_Client": new_id,
                "Nom": use_name,
                "Date": str(d),
                "Mois": f"{d.month:02d}",
                "Catégorie": _safe_str(sel_cat),
                "Visa": _safe_str(visa if isinstance(visa, str) else ""),
                HONO: float(honoraires or 0.0),
                AUTRE: float(autres or 0.0),
                TOTAL: total, "Payé": float(paye_init or 0.0), "Reste": reste,
                ESC_TR: 0.0, ESC_JR: "", "Paiements": "",
                "Dossier envoyé": bool(val_envoye), "Dossier approuvé": bool(val_appr),
                "RFE": bool(val_rfe), "Dossier refusé": bool(val_refuse), "Dossier annulé": bool(val_annule)
            }
            live_after = pd.concat([live_raw.drop(columns=["_RowID"]), pd.DataFrame([new_row])], ignore_index=True)
            live_after = ensure_dossier_numbers(live_after)
            write_sheet_inplace(current_path, client_target_sheet, live_after); save_workspace_path(current_path)
            st.success(f"Client créé **dans le fichier** avec Dossier N {next_num}. ✅"); st.rerun()

    # --- MODIFIER ---
    if action == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            opts = [(int(r["_RowID"]), f'{int(r.get(DOSSIER_COL,0))} — {_safe_str(r.get("ID_Client"))} — {_safe_str(r.get("Nom"))}')
                    for _,r in live_raw.iterrows()]
            label = st.selectbox("Sélection", [lab for _,lab in opts], key="edit_select")
            sel_rowid = [rid for rid,lab in opts if lab==label][0]
            idx = live_raw.index[live_raw["_RowID"]==sel_rowid][0]
            init = live_raw.loc[idx].to_dict()

            with st.form("edit_form", clear_on_submit=False):
                c0, c1, c2 = st.columns([1,1,1])
                c0.metric("Dossier N", f'{int(init.get(DOSSIER_COL,0))}')
                nom = c1.text_input("Nom", value=_safe_str(init.get("Nom")), key="edit_nom")
                try: d_init = pd.to_datetime(init.get("Date")).date() if _safe_str(init.get("Date")) else date.today()
                except Exception: d_init = date.today()
                d = c2.date_input("Date", value=d_init, key="edit_date")

                cC, cV = st.columns(2)
                init_cat = _safe_str(init.get("Catégorie"))
                sel_cat = cC.selectbox("Catégorie", [""] + cats_ref, index=([""]+cats_ref).index(init_cat) if init_cat in ([""]+cats_ref) else 0, key="edit_cat")
                visas_opt = sorted(visa_ref.loc[visa_ref["Catégorie"]==sel_cat, "Visa"].unique().tolist()) if sel_cat else visas_all
                init_visa = _safe_str(init.get("Visa"))
                if visas_opt:
                    try: idxv = visas_opt.index(init_visa)
                    except Exception: idxv = 0
                    visa = cV.selectbox("Visa", visas_opt, index=idxv, key="edit_visa")
                else:
                    visa = cV.text_input("Visa", value=init_visa, key="edit_visa_txt")

                def _f(v, alt=0.0):
                    try: return float(v)
                    except Exception: return float(alt)
                hono0  = _f(init.get(HONO, init.get("Montant", 0.0)))
                autre0 = _f(init.get(AUTRE, 0.0))
                paye0  = _f(init.get("Payé", 0.0))
                moved0 = _f(init.get(ESC_TR, 0.0))
                c5,c6 = st.columns(2)
                honoraires = c5.number_input("Montant honoraires (US $)", value=hono0, step=10.0, format="%.2f", key="edit_hono")
                autres     = c6.number_input("Autres frais (US $)", value=autre0, step=10.0, format="%.2f", key="edit_autre")
                c7,c8 = st.columns(2)
                total_preview = float(honoraires + autres); c7.metric("Total (US $)", _fmt_money_us(total_preview))
                paye    = c8.number_input("Payé (US $)", value=paye0, step=10.0, format="%.2f", key="edit_paye")

                st.caption(f"ESCROW transféré (cumul) actuellement : {_fmt_money_us(moved0)} — (gérer les transferts dans l’onglet ESCROW)")
                st.markdown("#### État du dossier")
                val_envoye = st.checkbox("Dossier envoyé",  value=bool(init.get("Dossier envoyé")),  key="edit_env") if has_envoye else False
                val_appr   = st.checkbox("Dossier approuvé",value=bool(init.get("Dossier approuvé")), key="edit_app") if has_appr   else False
                val_rfe    = st.checkbox("RFE",             value=bool(init.get("RFE")),              key="edit_rfe") if has_rfe    else False
                val_refuse = st.checkbox("Dossier refusé",  value=bool(init.get("Dossier refusé")),   key="edit_ref") if has_refuse else False
                val_annule = st.checkbox("Dossier annulé",  value=bool(init.get("Dossier annulé")),   key="edit_ann") if has_annule else False

                ok = st.form_submit_button("💾 Enregistrer (dans le fichier)", type="primary")
            if ok:
                if val_rfe and not (val_envoye or val_refuse or val_annule):
                    st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché."); st.stop()
                live = live_raw.drop(columns=["_RowID"]).copy()
                # localiser la ligne
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
                    total = float((honoraires or 0.0)+(autres or 0.0))
                    live.at[t_idx,"Nom"]=_safe_str(nom)
                    live.at[t_idx,"Date"]=str(d); live.at[t_idx,"Mois"]=f"{d.month:02d}"
                    live.at[t_idx,"Catégorie"]=_safe_str(sel_cat)
                    live.at[t_idx,"Visa"]=_safe_str(visa if isinstance(visa,str) else "")
                    live.at[t_idx, HONO]=float(honoraires or 0.0)
                    live.at[t_idx, AUTRE]=float(autres or 0.0)
                    live[TOTAL] = live.get(TOTAL, 0.0); live.at[t_idx, TOTAL]=total
                    live.at[t_idx,"Payé"]=float(paye or 0.0)
                    live.at[t_idx,"Reste"]=max(total - float(paye or 0.0), 0.0)
                    for c in [ESC_TR, ESC_JR]:
                        if c not in live.columns: live[c] = 0.0 if c==ESC_TR else ""
                    if has_envoye: live.at[t_idx,"Dossier envoyé"]=bool(val_envoye)
                    if has_appr:   live.at[t_idx,"Dossier approuvé"]=bool(val_appr)
                    if has_rfe:    live.at[t_idx,"RFE"]=bool(val_rfe)
                    if has_refuse: live.at[t_idx,"Dossier refusé"]=bool(val_refuse)
                    if has_annule: live.at[t_idx,"Dossier annulé"]=bool(val_annule)
                    live = ensure_dossier_numbers(live)
                    write_sheet_inplace(current_path, client_target_sheet, live); save_workspace_path(current_path)
                    st.success("Modifications enregistrées **dans le fichier**. ✅"); st.rerun()

    # --- SUPPRIMER ---
    if action == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client (écrit directement)")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            opts = [(int(r["_RowID"]), f'{int(r.get(DOSSIER_COL,0))} — {_safe_str(r.get("ID_Client"))} — {_safe_str(r.get("Nom"))}')
                    for _,r in live_raw.iterrows()]
            label = st.selectbox("Sélection", [lab for _,lab in opts], key="del_select")
            sel_rowid = [rid for rid,lab in opts if lab==label][0]
            idx = live_raw.index[live_raw["_RowID"]==sel_rowid][0]
            st.error("⚠️ Action irréversible.")
            if st.button("Supprimer (dans le fichier)", key="del_btn"):
                live = live_raw.drop(columns=["_RowID"]).copy()
                key = _safe_str(live_raw.at[idx, "ID_Client"])
                if key and "ID_Client" in live.columns:
                    live = live[live["ID_Client"].astype(str)!=key].reset_index(drop=True)
                else:
                    nom = _safe_str(live_raw.at[idx,"Nom"]); dat = _safe_str(live_raw.at[idx,"Date"])
                    live = live[~((live.get("Nom","").astype(str)==nom)&(live.get("Date","").astype(str)==dat))].reset_index(drop=True)
                live = ensure_dossier_numbers(live)
                write_sheet_inplace(current_path, client_target_sheet, live); save_workspace_path(current_path)
                st.success("Client supprimé **dans le fichier**. ✅"); st.rerun()

# ================= ANALYSES =================
with tabs[2]:
    st.subheader("📊 Analyses — Volumes & Financier")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa)."); st.stop()
    visa_ref = read_visa_reference(current_path)
    dfA_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    dfA = normalize_dataframe(dfA_raw, visa_ref=visa_ref).copy()
    if dfA.empty:
        st.info("Aucune donnée pour analyser."); st.stop()

    with st.container():
        c1, c2, c3, c4, c5 = st.columns(5)
        catsA  = sorted(dfA["Catégorie"].dropna().astype(str).unique()) if "Catégorie" in dfA.columns else []
        visasA = sorted(dfA["Visa"].dropna().astype(str).unique()) if "Visa" in dfA.columns else []
        sel_cats  = c1.multiselect("Catégorie", catsA, default=[], key="anal_cats")
        sel_visas = c2.multiselect("Type de visa", visasA, default=[], key="anal_visa")
        yearsA  = sorted({d.year for d in dfA["Date"] if pd.notna(d)}) if "Date" in dfA.columns else []
        sel_years  = c3.multiselect("Année", yearsA, default=[], key="anal_years")
        monthsA = [f"{m:02d}" for m in range(1,13)]
        sel_months = c4.multiselect("Mois (MM)", monthsA, default=[], key="anal_months")
        include_na_dates = c5.checkbox("Inclure lignes sans date", value=True, key="anal_na_dates")

    with st.container():
        d1, d2 = st.columns(2)
        today = date.today()
        if ("Date" in dfA.columns) and dfA["Date"].notna().any():
            dmin = min([d for d in dfA["Date"] if pd.notna(d)]); dmax = max([d for d in dfA["Date"] if pd.notna(d)])
        else:
            dmin, dmax = today - timedelta(days=365), today
        date_from = d1.date_input("Du", value=dmin, key="anal_date_from")
        date_to   = d2.date_input("Au", value=dmax, key="anal_date_to")
        c3a, c3b = st.columns(2)
        agg_with_year = c3a.toggle("Agrégation par Année-Mois (YYYY-MM)", value=False, key="anal_agg_with_year")
        show_tables   = c3b.toggle("Voir les tableaux détaillés", value=True, key="anal_show_tables")

    fA = dfA.copy()
    if sel_cats:  fA = fA[fA["Catégorie"].astype(str).isin(sel_cats)]
    if sel_visas: fA = fA[fA["Visa"].astype(str).isin(sel_visas)]
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
        if include_na_dates: mask_range = mask_range | fA["Date"].isna()
        fA = fA[mask_range]

    # Période (YYYY-MM si demandé, sinon MM)
    if agg_with_year:
        fA["Periode"] = fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")
    else:
        fA["Periode"] = fA["Mois"].fillna("NA")

    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in fA.columns: fA[col] = pd.to_numeric(fA[col], errors="coerce").fillna(0.0)

    def derive_statut(row) -> str:
        if bool(row.get("Dossier approuvé", False)): return "Approuvé"
        if bool(row.get("Dossier refusé", False)):   return "Refusé"
        if bool(row.get("Dossier annulé", False)):   return "Annulé"
        return "En attente"
    fA["Statut"] = fA.apply(derive_statut, axis=1)

    # ---- Graphiques Volumes
    st.markdown("### 📈 Volumes")
    vol_crees = fA.groupby("Periode").size().reset_index(name="Créés")
    df_vol = vol_crees.rename(columns={"Créés":"Volume"}).assign(Indic="Créés")
    df_vol = _clean_for_chart(df_vol, ["Periode","Indic"], ["Volume"], ["Periode","Indic","Volume"])
    if not df_vol.empty:
        try:
            chart_vol = alt.Chart(df_vol).mark_line(point=True).encode(
                x=alt.X("Periode:N", sort=None, title="Période"),
                y=alt.Y("Volume:Q"),
                color=alt.Color("Indic:N", legend=alt.Legend(title="Statut")),
                tooltip=["Periode","Indic","Volume"]
            ).properties(height=280, use_container_width=True)
            st.altair_chart(chart_vol, use_container_width=True)
        except Exception:
            st.dataframe(df_vol, use_container_width=True)

    # ---- Graphiques Financier
    st.markdown("### 💵 Financier")
    fin = fA.groupby("Periode", dropna=False)[[HONO, AUTRE, TOTAL, "Payé","Reste"]].sum().reset_index()
    if not fin.empty:
        ca = fin.melt(id_vars="Periode", value_vars=[HONO, AUTRE], var_name="Type", value_name="Montant")
        ca = _clean_for_chart(ca, ["Periode","Type"], ["Montant"], ["Periode","Type","Montant"])
        try:
            chart_ca = alt.Chart(ca).mark_bar().encode(
                x=alt.X("Periode:N", title="Période"), y=alt.Y("Montant:Q"),
                color=alt.Color("Type:N", legend=alt.Legend(title="Composant")),
                tooltip=["Periode","Type", alt.Tooltip("Montant:Q", format="$.2f")]
            ).properties(title="Chiffre d'affaires (Honoraires + Autres)", height=280)
            st.altair_chart(chart_ca, use_container_width=True)
        except Exception:
            st.dataframe(ca, use_container_width=True)

        enc = fin.melt(id_vars="Periode", value_vars=["Payé","Reste"], var_name="Indicateur", value_name="Montant")
        enc = _clean_for_chart(enc, ["Periode","Indicateur"], ["Montant"], ["Periode","Indicateur","Montant"])
        try:
            chart_enc = alt.Chart(enc).mark_line(point=True).encode(
                x=alt.X("Periode:N", title="Période"),
                y=alt.Y("Montant:Q"),
                color=alt.Color("Indicateur:N", legend=alt.Legend(title="Indicateur")),
                tooltip=["Periode","Indicateur", alt.Tooltip("Montant:Q", format="$.2f")]
            ).properties(title="Encaissements vs Solde restant", height=280)
            st.altair_chart(chart_enc, use_container_width=True)
        except Exception:
            st.dataframe(enc, use_container_width=True)

    # ---- Répartition par Catégorie / Visa
    st.markdown("### 🧭 Répartition par Catégorie / Visa")
    rep = fA.groupby(["Catégorie","Visa"]).agg(
        Dossiers=("Visa","count"),
        Total_USD=(TOTAL,"sum"),
        Paye_USD=("Payé","sum"),
        Reste_USD=("Reste","sum")
    ).reset_index().sort_values(["Catégorie","Dossiers"], ascending=[True,False])
    if not rep.empty:
        repc = _clean_for_chart(rep, ["Catégorie","Visa"], ["Dossiers","Total_USD","Paye_USD","Reste_USD"], ["Visa","Dossiers"])
        try:
            chart_rep = alt.Chart(repc).mark_bar().encode(
                y=alt.Y("Visa:N", sort="-x", title="Visa"),
                x=alt.X("Dossiers:Q", title="Nb dossiers"),
                color=alt.Color("Catégorie:N"),
                tooltip=["Catégorie","Visa","Dossiers",
                         alt.Tooltip("Total_USD:Q", format="$.2f", title="Total"),
                         alt.Tooltip("Paye_USD:Q", format="$.2f", title="Payé"),
                         alt.Tooltip("Reste_USD:Q", format="$.2f", title="Reste")]
            ).properties(height=360)
            st.altair_chart(chart_rep, use_container_width=True)
        except Exception:
            st.dataframe(repc, use_container_width=True)

    st.divider()
    st.markdown("### 🔎 Détails (clients)")
    details_cols = [c for c in ["Periode",DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Date", HONO, AUTRE, TOTAL, "Payé","Reste","Statut"] if c in fA.columns]
    details = fA[details_cols].copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in details.columns:
            details[col] = details[col].apply(lambda x: _fmt_money_us(x) if pd.notna(x) else "")
    d1, d2 = st.columns(2)
    statut_filter = d1.multiselect("Filtrer par statut", ["Approuvé","Refusé","Annulé","En attente"], key="anal_statut_filter")
    search = d2.text_input("Recherche (Nom / Catégorie / Visa / ID / Dossier N)", key="anal_search")
    if statut_filter:
        mask_st = fA["Statut"].isin(statut_filter); details = details[mask_st.values]
    if search:
        s = search.lower()
        mask_s = fA.apply(lambda r: (s in str(r.get("Nom","")).lower()) or (s in str(r.get("Catégorie","")).lower()) or
                                   (s in str(r.get("Visa","")).lower()) or (s in str(r.get("ID_Client","")).lower()) or
                                   (s in str(r.get(DOSSIER_COL,"")).lower()), axis=1)
        details = details[mask_s.values]
    st.dataframe(details.sort_values(["Periode","Catégorie","Nom"]), use_container_width=True)

    # ---- Fiche & règlements du client (avec ajout de paiement) ----
    st.markdown("#### 🧾 Fiche & règlements du client")
    base_live = read_sheet(current_path, client_target_sheet, normalize=False).copy()
    base_norm = normalize_dataframe(base_live.copy(), visa_ref=visa_ref)
    if base_norm.empty:
        st.info("Aucune donnée client.")
    else:
        base_norm["_label"] = base_norm.apply(lambda r: f'{r.get(DOSSIER_COL,"")} — {r.get("ID_Client","")} — {r.get("Nom","")} — {r.get("Catégorie","")}/{r.get("Visa","")}', axis=1)
        labels = base_norm["_label"].tolist()
        sel_lab = st.selectbox("Sélectionne un client :", labels, index=0, key="detail_sel")
        sel_id = base_norm.loc[base_norm["_label"]==sel_lab, "ID_Client"].iloc[0]

        rowN = base_norm.loc[base_norm["ID_Client"]==sel_id].iloc[0]
        k1,k2,k3,k4,k5 = st.columns(5)
        k1.metric("Honoraires", _fmt_money_us(float(rowN.get(HONO,0.0))))
        k2.metric("Autres frais", _fmt_money_us(float(rowN.get(AUTRE,0.0))))
        k3.metric("Total", _fmt_money_us(float(rowN.get(TOTAL,0.0))))
        k4.metric("Payé", _fmt_money_us(float(rowN.get("Payé",0.0))))
        k5.metric("Reste", _fmt_money_us(float(rowN.get("Reste",0.0))))

        rlive = base_live.loc[base_live.get("ID_Client","").astype(str)==str(sel_id)]
        plist = _parse_json_list(rlive.iloc[0].get("Paiements","")) if not rlive.empty else []
        st.markdown("**Historique des règlements**")
        if plist:
            dfp = pd.DataFrame(plist)
            if "date" in dfp.columns: dfp["date"] = pd.to_datetime(dfp["date"], errors="coerce").dt.date.astype(str)
            if "amount" in dfp.columns: dfp["Montant (US $)"] = dfp["amount"].apply(lambda x: _fmt_money_us(float(x) if pd.notna(x) else 0.0))
            for col in ["mode","note"]:
                if col not in dfp.columns: dfp[col] = ""
            show_cols = [c for c in ["date","mode","Montant (US $)","note"] if c in dfp.columns]
            st.table(dfp[show_cols].rename(columns={"date":"Date","mode":"Mode","note":"Note"}))
        else:
            st.caption("Aucun paiement enregistré pour ce client.")

        st.markdown("**Ajouter un règlement**")
        cA, cB, cC, cD = st.columns([1,1,1,2])
        pay_date = cA.date_input("Date", value=date.today(), key=f"pay_date_{sel_id}")
        pay_mode = cB.selectbox("Mode", ["CB","Chèque","Espèces","Virement","Venmo","Autre"], key=f"pay_mode_{sel_id}")
        pay_amt  = cC.number_input("Montant ($)", min_value=0.0, step=10.0, format="%.2f", key=f"pay_amt_{sel_id}")
        pay_note = cD.text_input("Note", "", key=f"pay_note_{sel_id}")
        if st.button("💾 Enregistrer ce règlement (dans le fichier)", key=f"pay_add_btn_{sel_id}"):
            try:
                live = read_sheet(current_path, client_target_sheet, normalize=False)
                if "Paiements" not in live.columns: live["Paiements"] = ""
                idxs = live.index[live.get("ID_Client","").astype(str)==str(sel_id)]
                if len(idxs)==0:
                    st.error("Dossier introuvable."); st.stop()
                i = idxs[0]
                pay_list = _parse_json_list(live.at[i, "Paiements"])
                add = float(pay_amt or 0.0)
                if add <= 0:
                    st.warning("Le montant doit être > 0.")
                    st.stop()
                norm = normalize_dataframe(live.copy(), visa_ref=visa_ref)
                mask_id = norm["ID_Client"].astype(str) == str(sel_id)
                reste_curr = float(norm.loc[mask_id, "Reste"].sum()) if mask_id.any() else 0.0
                if add > reste_curr + 1e-9:
                    add = reste_curr
                pay_list.append({"date": str(pay_date), "amount": float(add), "mode": pay_mode, "note": pay_note})
                live.at[i, "Paiements"] = json.dumps(pay_list, ensure_ascii=False)
                for c in [HONO, AUTRE, TOTAL, "Payé", "Reste"]:
                    if c not in live.columns: live[c] = 0.0
                total_paid = _sum_payments(pay_list)
                hono = _to_num(pd.Series([live.at[i, HONO]])).iloc[0] if HONO in live.columns else 0.0
                autr = _to_num(pd.Series([live.at[i, AUTRE]])).iloc[0] if AUTRE in live.columns else 0.0
                total = float(hono + autr)
                live.at[i, "Payé"]  = float(total_paid)
                live.at[i, "Reste"] = max(total - float(total_paid), 0.0)
                live.at[i, TOTAL]   = total
                write_sheet_inplace(current_path, client_target_sheet, live)
                st.success("Règlement ajouté **dans le fichier**. ✅")
                st.rerun()
            except Exception as e:
                st.error(f"Erreur : {e}")

# ================= ESCROW =================
with tabs[3]:
    st.subheader("🏦 ESCROW — dépôts sur honoraires & transferts")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa)."); st.stop()
    visa_ref = read_visa_reference(current_path)
    live_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    live = normalize_dataframe(live_raw, visa_ref=visa_ref).copy()
    if ESC_TR not in live.columns: live[ESC_TR] = 0.0
    else: live[ESC_TR] = pd.to_numeric(live[ESC_TR], errors="coerce").fillna(0.0)
    live["ESCROW dispo"] = live.apply(escrow_available_from_row, axis=1)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Dossiers", f"{len(live)}")
    c2.metric("ESCROW total dispo", _fmt_money_us(float(live["ESCROW dispo"].sum())))
    envoyes = live[(live["Dossier envoyé"]==True)]
    a_transferer = envoyes[envoyes["ESCROW dispo"]>0.004]
    c3.metric("Dossiers envoyés (à réclamer)", f"{len(a_transferer)}")
    c4.metric("Montant à réclamer", _fmt_money_us(float(a_transferer["ESCROW dispo"].sum())))

    st.divider()
    st.markdown("### 📌 À transférer (dossiers **envoyés**)")
    if a_transferer.empty:
        st.success("Aucun transfert en attente pour des dossiers envoyés.")
    else:
        for _, r in a_transferer.sort_values("Date").iterrows():
            with st.expander(f'🔔 {r[DOSSIER_COL]} — {r["ID_Client"]} — {r["Nom"]} — {r.get("Catégorie","")} / {r["Visa"]} — ESCROW dispo: {_fmt_money_us(r["ESCROW dispo"])}'):
                cA, cB, cC = st.columns(3)
                cA.metric("Honoraires", _fmt_money_us(float(r.get(HONO,0.0))))
                cB.metric("Déjà transféré", _fmt_money_us(float(r.get(ESC_TR,0.0))))
                cC.metric("Payé", _fmt_money_us(float(r.get("Payé",0.0))))
                amt = st.number_input("Montant à marquer comme transféré (US $)",
                                      min_value=0.0, value=float(r["ESCROW dispo"]),
                                      step=10.0, format="%.2f", key=f"esc_amt_{r['ID_Client']}")
                note = st.text_input("Note (facultatif)", "", key=f"esc_note_{r['ID_Client']}")
                if st.button("✅ Marquer transféré (écrit dans le fichier)", key=f"esc_btn_{r['ID_Client']}"):
                    try:
                        live_w = read_sheet(current_path, client_target_sheet, normalize=False).copy()
                        for c in [ESC_TR, ESC_JR]:
                            if c not in live_w.columns: live_w[c] = 0.0 if c==ESC_TR else ""
                        idxs = live_w.index[live_w.get("ID_Client","").astype(str)==str(r["ID_Client"])]
                        if len(idxs)==0: st.error("Ligne introuvable."); st.stop()
                        i = idxs[0]
                        tmp = normalize_dataframe(live_w.copy(), visa_ref=visa_ref)
                        disp = float(tmp.loc[tmp["ID_Client"].astype(str)==str(r["ID_Client"]), :].apply(escrow_available_from_row, axis=1).iloc[0])
                        add = float(min(max(amt,0.0), disp))
                        live_w.at[i, ESC_TR] = float(pd.to_numeric(pd.Series([live_w.at[i, ESC_TR]]), errors="coerce").fillna(0.0).iloc[0] + add)
                        live_w.at[i, ESC_JR] = append_escrow_journal(live_w.loc[i], add, note)
                        live_w = ensure_dossier_numbers(live_w)
                        write_sheet_inplace(current_path, client_target_sheet, live_w)
                        st.success("Transfert ESCROW enregistré **dans le fichier**. ✅"); st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")

    st.divider()
    st.markdown("### 📥 En cours d’alimentation (dossiers **non envoyés**)")
    non_env = live[(live["Dossier envoyé"]!=True) & (live["ESCROW dispo"]>0.004)].copy()
    if non_env.empty:
        st.info("Rien en attente côté dossiers non envoyés.")
    else:
        show = non_env[[DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Date",HONO,"Payé",ESC_TR,"ESCROW dispo"]].copy()
        for col in [HONO,"Payé",ESC_TR,"ESCROW dispo"]:
            show[col] = show[col].map(_fmt_money_us)
        st.dataframe(show, use_container_width=True)

    st.divider()
    st.markdown("### 🧾 Historique des transferts (journal)")
    has_journal = live[live[ESC_JR].astype(str).str.len()>0]
    if has_journal.empty:
        st.caption("Aucun journal de transfert pour le moment.")
    else:
        rows = []
        for _, r in has_journal.iterrows():
            entries = _parse_json_list(r[ESC_JR])
            for e in entries:
                rows.append({
                    DOSSIER_COL: r.get(DOSSIER_COL, ""),
                    "ID_Client": r["ID_Client"], "Nom": r["Nom"], "Visa": r["Visa"],
                    "Date": r.get("Date"), "Horodatage": e.get("ts"),
                    "Montant (US $)": float(e.get("amount",0.0)), "Note": e.get("note","")
                })
        jdf = pd.DataFrame(rows).sort_values("Horodatage")
        jdf["Montant (US $)"] = jdf["Montant (US $)"].map(_fmt_money_us)
        st.dataframe(jdf, use_container_width=True)
