# app.py — COMPLET (Partie 1/2)
import io, json, hashlib
from datetime import date, datetime, timedelta
from pathlib import Path
import streamlit as st
import pandas as pd
import altair as alt

# -------- Config --------
alt.data_transformers.disable_max_rows()
alt.renderers.set_embed_options(actions=False)
st.set_page_config(page_title="📊 Visas — Edition directe + ESCROW + Analyses", layout="wide")
st.title("📊 Visas — Edition DIRECTE du fichier (ESCROW + Analyses + Cascade)")

# -------- Constantes colonnes --------
HONO   = "Honoraires (US $)"
AUTRE  = "Autres frais (US $)"
TOTAL  = "Total (US $)"
ESC_TR = "Escrow transféré (US $)"
ESC_JR = "Escrow journal"
DOSSIER_COL   = "Dossier N"
DOSSIER_START = 13057

# Statuts + dates
S_ENVOYE   = "Dossier envoyé";   D_ENVOYE   = "Date envoyé"
S_APPROUVE = "Dossier approuvé"; D_APPROUVE = "Date approuvé"
S_RFE      = "RFE";              D_RFE      = "Date RFE"
S_REFUSE   = "Dossier refusé";   D_REFUSE   = "Date refusé"
S_ANNULE   = "Dossier annulé";   D_ANNULE   = "Date annulé"
STATUS_COLS  = [S_ENVOYE, S_APPROUVE, S_RFE, S_REFUSE, S_ANNULE]
STATUS_DATES = [D_ENVOYE, D_APPROUVE, D_RFE, D_REFUSE, D_ANNULE]

# -------- Workspace (dernier fichier) --------
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
    if WS_FILE is None or not WS_FILE.exists(): return None
    try:
        data = json.loads(WS_FILE.read_text(encoding="utf-8"))
        p = Path(data.get("last_path","")); return p if p.exists() else None
    except Exception:
        return None

def save_workspace_path(p: Path):
    if WS_FILE is None: return
    try: WS_FILE.write_text(json.dumps({"last_path": str(p)}), encoding="utf-8")
    except Exception: pass

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
                dest = cand; break
            n += 1
    dest.write_bytes(upload.read())
    return dest

# -------- Utils --------
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
        v = json.loads(x); return v if isinstance(v, list) else []
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
    return ("visa" in cols) and not ({"montant","honoraires","payé","reste","solde"} & cols)

def is_clients_like(df: pd.DataFrame) -> bool:
    cols = set(df.columns.astype(str))
    return {"Nom","Visa"}.issubset(cols)

# ESCROW helpers
def escrow_available_from_row(row) -> float:
    hon  = float(row.get(HONO, 0.0)  or 0.0)
    paid = float(row.get("Payé", 0.0) or 0.0)
    moved= float(row.get(ESC_TR, 0.0) or 0.0)
    return max(min(paid, hon) - moved, 0.0)

def append_escrow_journal(row_raw: pd.Series, amount: float, note: str = "") -> str:
    journal = _parse_json_list(row_raw.get(ESC_JR, ""))
    journal.append({"ts": datetime.now().isoformat(timespec="seconds"),
                    "amount": float(amount), "note": note})
    return json.dumps(journal, ensure_ascii=False)

# Dossier N
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
        df.at[i, DOSSIER_COL] = next_num; next_num += 1
    return df

def next_dossier_number(df_existing: pd.DataFrame) -> int:
    if DOSSIER_COL not in df_existing.columns or df_existing.empty: return DOSSIER_START
    s = _to_int(df_existing[DOSSIER_COL]); s = s[s > 0]
    return int(s.max() + 1) if not s.empty else DOSSIER_START

# ---------- Référentiel Visa : plat & hiérarchique ----------
def read_visa_reference(path: Path) -> pd.DataFrame:
    try: dfv = pd.read_excel(path, sheet_name="Visa")
    except Exception: return pd.DataFrame(columns=["Catégorie","Visa"])
    if "Visa" not in dfv.columns:
        visa_col = next((c for c in dfv.columns if str(c).strip().lower()=="visa"), None)
        if visa_col: dfv = dfv.rename(columns={visa_col:"Visa"})
    cat_col = next((c for c in dfv.columns if str(c).strip().lower() in ("catégorie","categorie")), None)
    if not cat_col: dfv["Catégorie"] = ""
    elif cat_col != "Catégorie": dfv = dfv.rename(columns={cat_col:"Catégorie"})
    for c in ["Catégorie","Visa"]:
        if c in dfv.columns: dfv[c] = dfv[c].astype(str).fillna("").str.strip()
    return dfv[["Catégorie","Visa"]]

# --- Hiérarchie complète : Catégorie, Sous-catégorie 1..n, Visa (facultatif) ---
def read_visa_reference_hier(path: Path) -> pd.DataFrame:
    try:
        dfv = pd.read_excel(path, sheet_name="Visa")
    except Exception:
        return pd.DataFrame(columns=["Catégorie","Visa"])
    def _norm(c):
        c = str(c).strip()
        c_low = (c.lower()
                   .replace("é","e").replace("è","e").replace("ê","e")
                   .replace("à","a").replace("ô","o").replace("ï","i")
                   .replace("ç","c"))
        return c, c_low
    rename = {}
    for c in dfv.columns:
        raw, low = _norm(c)
        if low in ("categorie","catégorie"): rename[c] = "Catégorie"
        elif low.startswith("sous-categorie") or low.startswith("sous-catégorie") or low.startswith("sous categorie"):
            rename[c] = " ".join([w.capitalize() for w in raw.split()])
        elif low == "visa": rename[c] = "Visa"
    if rename: dfv = dfv.rename(columns=rename)
    # trim
    for c in dfv.columns:
        if dfv[c].dtype == "object":
            dfv[c] = dfv[c].fillna("").astype(str).str.strip()
    return dfv

def get_hierarchy_columns(df_ref: pd.DataFrame) -> list[str]:
    cols = [c for c in df_ref.columns if isinstance(c, str)]
    out = []
    if "Catégorie" in cols: out.append("Catégorie")
    subs = []
    for c in cols:
        low = c.lower().replace("é","e").replace("è","e")
        if low.startswith("sous-categorie"): subs.append(c)
    def _num_key(c):
        import re
        m = re.search(r"(\d+)", c)
        return int(m.group(1)) if m else 999
    subs = sorted(subs, key=_num_key)
    out.extend(subs)
    if "Visa" in cols: out.append("Visa")
    return out

def cascading_visa_picker(df_ref: pd.DataFrame, key_prefix: str, init: dict | None = None) -> dict:
    """Sélecteurs en cascade. Retourne {col: valeur} (Visa peut être '')."""
    if df_ref is None or df_ref.empty:
        st.info("Référentiel Visa vide."); return {"Catégorie":"", "Visa":""}
    cols = get_hierarchy_columns(df_ref)
    sel = {}
    df_work = df_ref.copy()
    for col in cols:
        # filtre par choix précédents
        for prev_col, prev_val in sel.items():
            if prev_val != "":
                df_work = df_work[df_work[prev_col].astype(str) == prev_val]
        # options du niveau
        options = sorted([v for v in df_work[col].astype(str).unique() if v != ""])
        if not options:
            if col == "Visa": sel[col] = ""
            break
        # défaut
        default_idx = 0
        if init and init.get(col, "") in options:
            default_idx = options.index(init[col])
        sel[col] = st.selectbox(col, [""] + options, index=default_idx+1 if options else 0, key=f"{key_prefix}_{col}")
        if sel[col] == "":
            for c2 in cols[cols.index(col)+1:]:
                sel[c2] = ""
            break
    for c in cols: sel.setdefault(c, "")
    sel.setdefault("Catégorie",""); sel.setdefault("Visa","")
    return sel

# -------- Normalisation des feuilles clients --------
def map_category_from_ref(visalib: str, ref_df: pd.DataFrame) -> str:
    if ref_df is None or ref_df.empty: return ""
    tmp = ref_df[ref_df["Visa"].astype(str).str.strip().str.lower() == _safe_str(visalib).lower()]
    if tmp.empty: return ""
    return _safe_str(tmp.iloc[0]["Catégorie"])

def normalize_dataframe(df: pd.DataFrame, visa_ref: pd.DataFrame | None = None) -> pd.DataFrame:
    df = df.copy()
    df["Date"] = _to_date(df["Date"]) if "Date" in df.columns else pd.NaT
    df["Mois"] = df["Date"].apply(lambda x: f"{x.month:02d}" if pd.notna(x) else pd.NA)
    visa_col = next((c for c in ["Visa","Categories","Catégorie","Categorie","TypeVisa"] if c in df.columns), None)
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"
    if "Catégorie" not in df.columns: df["Catégorie"] = ""
    if visa_ref is not None and not visa_ref.empty:
        df["Catégorie"] = df.apply(lambda r: _safe_str(r.get("Catégorie")) or map_category_from_ref(_safe_str(r.get("Visa")), visa_ref), axis=1)
    # Montants
    hono_alias = [HONO,"Honoraires","Honoraires US $","Montant honoraires us $","Montant honoraires","Montant (US $)","Montant"]
    autre_alias= [AUTRE,"Autres frais","Frais","Other fees","Autres"]
    hsrc = next((c for c in hono_alias if c in df.columns), None)
    asrc = next((c for c in autre_alias if c in df.columns), None)
    df[HONO] = _to_num(df[hsrc]) if hsrc else 0.0
    df[AUTRE]= _to_num(df[asrc]) if asrc else 0.0
    # Paiements
    if "Paiements" in df.columns and df["Paiements"].astype(str).str.strip().ne("").any():
        parsed = df["Paiements"].apply(_parse_json_list)
        df["Payé"] = parsed.apply(_sum_payments).astype(float)
    else:
        acompte_cols = [c for c in df.columns if str(c).strip().lower().startswith("acompte")]
        if acompte_cols:
            df["Payé"] = _to_num(df[acompte_cols].fillna(0).sum(axis=1))
            def _build_json(row):
                entries=[]
                for c in acompte_cols:
                    val = _to_num(pd.Series([row.get(c,0)])).iloc[0]
                    if val>0: entries.append({"date": str(row.get("Date","")), "amount": float(val), "mode":"", "note": c})
                return json.dumps(entries, ensure_ascii=False) if entries else ""
            df["Paiements"] = df.apply(_build_json, axis=1)
        else:
            df["Payé"] = _to_num(df["Payé"]) if "Payé" in df.columns else 0.0
    df[TOTAL] = (df[HONO] + df[AUTRE]).astype(float)
    df["Reste"] = (df[TOTAL] - df["Payé"]).fillna(0.0)
    # Statuts + dates
    for b in STATUS_COLS:
        if b not in df.columns: df[b]=False
    for dc in STATUS_DATES:
        if dc not in df.columns: df[dc] = ""
        else:
            tmp = _to_date(df[dc])
            df[dc] = tmp.apply(lambda x: str(x) if pd.notna(x) else "")
    # Identité
    if "Nom" not in df.columns: df["Nom"]=""
    if "ID_Client" not in df.columns: df["ID_Client"]=""
    need_id = df["ID_Client"].astype(str).str.strip().eq("") | df["ID_Client"].isna()
    if need_id.any(): df.loc[need_id,"ID_Client"] = df.loc[need_id].apply(_make_client_id_from_row, axis=1)
    if "Paiements" not in df.columns: df["Paiements"]=""
    # ESCROW
    df[ESC_TR] = _to_num(df[ESC_TR]) if ESC_TR in df.columns else 0.0
    if ESC_JR not in df.columns: df[ESC_JR]=""
    # Dossier N
    df = ensure_dossier_numbers(df)
    # Nettoyage (on retire Téléphone/Email si présents)
    for dropcol in ["Telephone","Email"]:
        if dropcol in df.columns: df = df.drop(columns=[dropcol])
    ordered = [DOSSIER_COL,"ID_Client","Nom","Date","Mois","Catégorie","Visa",
               HONO,AUTRE,TOTAL,"Payé","Reste",ESC_TR,ESC_JR] + \
              STATUS_COLS + STATUS_DATES + ["Paiements"]
    cols = [c for c in ordered if c in df.columns] + [c for c in df.columns if c not in ordered]
    return df[cols]

# -------- IO Excel --------
def read_sheet(path: Path, sheet: str, normalize: bool, visa_ref: pd.DataFrame | None = None) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    if sheet not in xls.sheet_names:
        base = pd.DataFrame(columns=[DOSSIER_COL,"ID_Client","Nom","Date","Mois","Catégorie","Visa",
                                     HONO,AUTRE,TOTAL,"Payé","Reste",ESC_TR,ESC_JR] + STATUS_COLS + STATUS_DATES + ["Paiements"])
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

# -------- Sidebar / Source --------
st.sidebar.header("Source")
def _find_latest_xlsx(paths: list[Path]) -> Path | None:
    cand = []
    for base in paths:
        try:
            if base and base.exists(): cand.extend([p for p in base.glob("*.xlsx") if p.is_file()])
        except Exception: pass
    if not cand: return None
    return sorted(cand, key=lambda p: p.stat().st_mtime, reverse=True)[0]

current_path = load_workspace_path()
if (current_path is None) or (not current_path.exists()):
    search_dirs = [WORK_DIR] if WORK_DIR else []
    if Path("/mnt/data").exists(): search_dirs.append(Path("/mnt/data"))
    latest = _find_latest_xlsx(search_dirs)
    if latest: current_path = latest; save_workspace_path(current_path)
if (current_path is None) or (not current_path.exists()):
    defaults = [Path("/mnt/data/donnees_visa_clients.xlsx"),
                Path("/mnt/data/modele_clients_visa.xlsx"),
                Path("/mnt/data/Visa_Clients_20251001-114844.xlsx"),
                Path("/mnt/data/visa_analytics_datecol.xlsx"),
                Path("/mnt/data/donnees_visa_clients1.xlsx")]
    current_path = next((p for p in defaults if p.exists()), None)
    if current_path: save_workspace_path(current_path)

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

if current_path is None or not current_path.exists(): st.stop()

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
    st.error(f"Impossible de lire l'Excel : {e}"); st.stop()

preferred_order = ["Clients","Visa","Données normalisées"]
default_sheet = next((s for s in preferred_order if s in sheet_names), sheet_names[0])
sheet_choice = st.sidebar.selectbox("Feuille (Dashboard)", sheet_names, index=sheet_names.index(default_sheet))

# Feuille cible Clients (CRUD)
valid_client_sheets = []
for s in sheet_names:
    try:
        df_tmp = read_sheet(current_path, s, normalize=False)
        if is_clients_like(df_tmp): valid_client_sheets.append(s)
    except Exception: pass
if not valid_client_sheets:
    st.sidebar.error("Aucune feuille 'clients' valide (au moins Nom & Visa).")
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

# -------- Onglets --------
tabs = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"])


# app.py — COMPLET (2/2) — Partie 2
# (Nécessite en Partie 1 : read_visa_reference_hier, get_hierarchy_columns, cascading_visa_picker)

# ---------- utilitaire filtre à partir de la cascade ----------
def visas_autorises_depuis_cascade(df_ref_full: pd.DataFrame, sel_path: dict) -> list[str]:
    """À partir de la sélection hiérarchique, calcule la liste des visas autorisés.
       Si la branche n'a pas de 'Visa', on renvoie [] (aucun filtre sur Visa).
    """
    if df_ref_full is None or df_ref_full.empty:
        return []
    cols = get_hierarchy_columns(df_ref_full)
    dfw = df_ref_full.copy()
    for c in cols:
        val = sel_path.get(c, "")
        if val:
            dfw = dfw[dfw[c].astype(str) == val]
    if "Visa" not in dfw.columns:
        return []
    visas = sorted([v for v in dfw["Visa"].astype(str).unique() if v != ""])
    return visas

# ================= DASHBOARD =================
with tabs[0]:
    visa_ref_full = read_visa_reference_hier(current_path)
    visa_ref_simple = read_visa_reference(current_path)
    df_raw = read_sheet(current_path, sheet_choice, normalize=False)

    # Référentiel Visa (CRUD complet si on est sur la feuille Visa)
    if looks_like_reference(df_raw) and sheet_choice == "Visa":
        st.subheader("📄 Référentiel — Catégories / Sous-catégories / Visa")
        df_ref = visa_ref_full.copy()
        if df_ref.empty:
            st.info("Feuille Visa vide.")
        else:
            st.dataframe(df_ref, use_container_width=True)

        st.markdown("### ✏️ Gestion simple (Catégorie / Visa)")
        # On propose la gestion de base Catégorie/Visa (les sous-catégories restent telles quelles)
        base = df_ref.copy()
        # Assure au moins ces deux colonnes pour l'édition minimale :
        if "Catégorie" not in base.columns: base["Catégorie"] = ""
        if "Visa" not in base.columns: base["Visa"] = ""
        base_min = base[["Catégorie","Visa"]].copy()

        mode = st.radio("Action", ["Ajouter", "Renommer", "Supprimer"], horizontal=True, key="visa_ref_action")
        options = base_min.assign(_label=base_min["Catégorie"].str.cat(base_min["Visa"], sep=" — "))

        if mode == "Ajouter":
            cA, cB = st.columns(2)
            new_cat = cA.text_input("Catégorie").strip()
            new_vis = cB.text_input("Visa (facultatif)").strip()
            if st.button("➕ Ajouter"):
                out = pd.concat([df_ref, pd.DataFrame([{"Catégorie": new_cat, "Visa": new_vis}])], ignore_index=True)
                write_sheet_inplace(current_path, "Visa", out); st.success("Ajouté."); st.rerun()

        elif mode == "Renommer":
            if options.empty: st.info("Aucune entrée.")
            else:
                sel_lab = st.selectbox("Sélection (Catégorie — Visa)", options["_label"].tolist())
                row = options.loc[options["_label"]==sel_lab].iloc[0]
                cA, cB = st.columns(2)
                new_cat = cA.text_input("Nouvelle catégorie", value=row["Catégorie"]).strip()
                new_vis = cB.text_input("Nouveau visa", value=row["Visa"]).strip()
                if st.button("📝 Renommer"):
                    out = df_ref.copy()
                    mask = (out["Catégorie"]==row["Catégorie"]) & (out["Visa"]==row["Visa"])
                    out.loc[mask, ["Catégorie","Visa"]] = [new_cat, new_vis]
                    write_sheet_inplace(current_path, "Visa", out); st.success("Renommé."); st.rerun()

        else:  # Supprimer
            if options.empty: st.info("Aucune entrée.")
            else:
                sel_lab = st.selectbox("Sélection (Catégorie — Visa)", options["_label"].tolist())
                st.error("⚠️ Action irréversible (ligne correspondante).")
                if st.button("🗑️ Supprimer"):
                    cat0, vis0 = sel_lab.split(" — ", 1)
                    out = df_ref[~((df_ref["Catégorie"]==cat0) & (df_ref["Visa"]==vis0))].reset_index(drop=True)
                    write_sheet_inplace(current_path, "Visa", out); st.success("Supprimé."); st.rerun()
        st.stop()

    # Données normalisées pour affichage Dashboard
    df = read_sheet(current_path, sheet_choice, normalize=True, visa_ref=visa_ref_simple)

    # Filtres (avec cascade)
    st.markdown("### 🔎 Filtres")
    with st.container():
        cL, cR = st.columns([1,2])
        cL.caption("Sélection hiérarchique (réduit la liste des visas)")
        with cL:
            sel_path_dash = cascading_visa_picker(visa_ref_full, key_prefix="dash_cascade")
        visas_aut = visas_autorises_depuis_cascade(visa_ref_full, sel_path_dash)

        cR1, cR2, cR3 = cR.columns(3)
        years = sorted({d.year for d in df["Date"] if pd.notna(d)}) if "Date" in df.columns else []
        sel_years = cR1.multiselect("Année", years, default=[])
        months = sorted(df["Mois"].dropna().unique()) if "Mois" in df.columns else []
        sel_months = cR2.multiselect("Mois (MM)", months, default=[])
        include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True)

    # Application des filtres
    f = df.copy()
    # Catégorie (si choisie dans la cascade)
    if sel_path_dash.get("Catégorie", ""):
        f = f[f["Catégorie"].astype(str) == sel_path_dash["Catégorie"]]
    # Visa (liste autorisée selon la branche)
    if visas_aut:
        f = f[f["Visa"].astype(str).isin(visas_aut)]
    # Année/Mois
    if "Date" in f.columns and sel_years:
        mask = f["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask |= f["Date"].isna()
        f = f[mask]
    if "Mois" in f.columns and sel_months:
        mask = f["Mois"].isin(sel_months)
        if include_na_dates: mask |= f["Mois"].isna()
        f = f[mask]

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

    # Aperçu table
    st.divider()
    st.subheader("📋 Données (aperçu)")
    cols_show = [c for c in [
        DOSSIER_COL,"ID_Client","Nom","Date","Catégorie","Visa",
        HONO, AUTRE, TOTAL, "Payé","Reste",
        S_ENVOYE, D_ENVOYE, S_APPROUVE, D_APPROUVE, S_RFE, D_RFE, S_REFUSE, D_REFUSE, S_ANNULE, D_ANNULE
    ] if c in f.columns]
    view = f.copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in view.columns: view[col] = view[col].map(_fmt_money_us)
    if "Date" in view.columns: view["Date"] = view["Date"].astype(str)
    st.dataframe(view[cols_show], use_container_width=True)

# ================= CLIENTS (CRUD) =================
with tabs[1]:
    st.subheader("👤 Clients — Créer / Modifier / Supprimer (écriture **directe**)")
    if client_target_sheet is None:
        st.warning("Aucune feuille *Clients* valide (Nom & Visa)."); st.stop()
    if st.button("🔄 Recharger le fichier", key="reload_btn"):
        st.rerun()

    visa_ref_full = read_visa_reference_hier(current_path)
    visa_ref_simple = read_visa_reference(current_path)

    live_raw = read_sheet(current_path, client_target_sheet, normalize=False).copy()
    live_raw = ensure_dossier_numbers(live_raw)
    live_raw["_RowID"] = range(len(live_raw))

    action = st.radio("Action", ["Créer", "Modifier", "Supprimer"], horizontal=True, key="crud_action")

    # --- CREER ---
    if action == "Créer":
        st.markdown("### ➕ Nouveau client")
        # garantir colonnes
        for must in [DOSSIER_COL,"ID_Client","Nom","Date","Mois","Catégorie","Visa",
                     HONO, AUTRE, TOTAL, "Payé","Reste", ESC_TR, ESC_JR] + STATUS_COLS + STATUS_DATES + ["Paiements"]:
            if must not in live_raw.columns:
                if must in {HONO, AUTRE, TOTAL, "Payé","Reste", ESC_TR}: live_raw[must]=0.0
                elif must in {"Paiements", ESC_JR, "Nom","Date","Mois","Catégorie","Visa"}: live_raw[must]=""
                elif must in STATUS_DATES: live_raw[must]=""
                elif must in STATUS_COLS: live_raw[must]=False
                elif must==DOSSIER_COL: live_raw[must]=0
                else: live_raw[must]=""

        next_num = next_dossier_number(live_raw)
        with st.form("create_form", clear_on_submit=False):
            c0, c1, c2 = st.columns([1,1,1])
            c0.metric("Prochain Dossier N", f"{next_num}")
            nom_in = c1.text_input("Nom")
            d = c2.date_input("Date", value=date.today())

            # --- Sélection en cascade ---
            st.caption("Sélection hiérarchique du visa")
            sel_path = cascading_visa_picker(visa_ref_full, key_prefix="create_cascade")
            sel_cat = sel_path.get("Catégorie","")
            visa    = sel_path.get("Visa","")  # peut être ""

            c5,c6 = st.columns(2)
            honoraires = c5.number_input("Montant honoraires (US $)", value=0.0, step=10.0, format="%.2f")
            autres     = c6.number_input("Autres frais (US $)", value=0.0, step=10.0, format="%.2f")

            st.markdown("#### État du dossier")
            r1c1, r1c2 = st.columns(2)
            v_env   = r1c1.checkbox(S_ENVOYE,  value=False)
            dt_env  = r1c2.date_input(D_ENVOYE, value=date.today(), disabled=not v_env, key="dt_env_cre")
            r2c1, r2c2 = st.columns(2)
            v_app   = r2c1.checkbox(S_APPROUVE, value=False)
            dt_app  = r2c2.date_input(D_APPROUVE, value=date.today(), disabled=not v_app, key="dt_app_cre")
            r3c1, r3c2 = st.columns(2)
            v_rfe   = r3c1.checkbox(S_RFE,      value=False)
            dt_rfe  = r3c2.date_input(D_RFE,    value=date.today(), disabled=not v_rfe, key="dt_rfe_cre")
            r4c1, r4c2 = st.columns(2)
            v_ref   = r4c1.checkbox(S_REFUSE,   value=False)
            dt_ref  = r4c2.date_input(D_REFUSE, value=date.today(), disabled=not v_ref, key="dt_ref_cre")
            r5c1, r5c2 = st.columns(2)
            v_ann   = r5c1.checkbox(S_ANNULE,   value=False)
            dt_ann  = r5c2.date_input(D_ANNULE, value=date.today(), disabled=not v_ann, key="dt_ann_cre")

            ok = st.form_submit_button("💾 Sauvegarder (dans le fichier)", type="primary")

        if ok:
            if v_rfe and not (v_env or v_ref or v_ann):
                st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché."); st.stop()
            existing_names = set(live_raw["Nom"].dropna().astype(str))
            base_name = _safe_str(nom_in); use_name = base_name
            if base_name in existing_names:
                k = 0
                while f"{base_name}-{k}" in existing_names: k += 1
                use_name = f"{base_name}-{k}"
            gen_id = _make_client_id_from_row({"Nom": use_name, "Date": d})
            existing_ids = set(live_raw["ID_Client"].astype(str)) if "ID_Client" in live_raw.columns else set()
            new_id = gen_id; n=1
            while new_id in existing_ids: n+=1; new_id=f"{gen_id}-{n:02d}"
            total = float((honoraires or 0.0)+(autres or 0.0))
            paye_init = 0.0
            reste = max(total - paye_init, 0.0)
            new_row = {
                DOSSIER_COL: int(next_num), "ID_Client": new_id, "Nom": use_name,
                "Date": str(d), "Mois": f"{d.month:02d}", "Catégorie": _safe_str(sel_cat), "Visa": _safe_str(visa),
                HONO: float(honoraires or 0.0), AUTRE: float(autres or 0.0),
                TOTAL: total, "Payé": paye_init, "Reste": reste,
                ESC_TR: 0.0, ESC_JR: "", "Paiements": "",
                S_ENVOYE: bool(v_env),   D_ENVOYE:   (str(dt_env) if v_env else ""),
                S_APPROUVE: bool(v_app), D_APPROUVE: (str(dt_app) if v_app else ""),
                S_RFE: bool(v_rfe),      D_RFE:      (str(dt_rfe) if v_rfe else ""),
                S_REFUSE: bool(v_ref),   D_REFUSE:   (str(dt_ref) if v_ref else ""),
                S_ANNULE: bool(v_ann),   D_ANNULE:   (str(dt_ann) if v_ann else "")
            }
            live_after = pd.concat([live_raw.drop(columns=["_RowID"]), pd.DataFrame([new_row])], ignore_index=True)
            live_after = ensure_dossier_numbers(live_after)
            write_sheet_inplace(current_path, client_target_sheet, live_after); save_workspace_path(current_path)
            st.success(f"Client créé **dans le fichier** (Dossier N {next_num}). ✅"); st.rerun()

    # --- MODIFIER ---
    if action == "Modifier":
        st.markdown("### ✏️ Modifier un client (fiche + paiements + dates)")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            options = [(int(r["_RowID"]),
                        int(r.name),
                        f'{int(r.get(DOSSIER_COL,0))} — { _safe_str(r.get("ID_Client")) } — { _safe_str(r.get("Nom")) }')
                       for _,r in live_raw.iterrows()]
            labels  = [lab for _,__,lab in options]
            sel_lab = st.selectbox("Sélection", labels, key="edit_sel_label")
            sel_rowid, orig_pos, _ = [t for t in options if t[2]==sel_lab][0]
            idx = live_raw.index[live_raw["_RowID"]==sel_rowid][0]
            init = live_raw.loc[idx].to_dict()

            with st.form(f"edit_form_{sel_rowid}", clear_on_submit=False):
                c0, c1, c2 = st.columns([1,1,1])
                c0.metric("Dossier N", f'{int(init.get(DOSSIER_COL,0))}')
                nom = c1.text_input("Nom", value=_safe_str(init.get("Nom")), key=f"edit_nom_{sel_rowid}")
                try:
                    d_init = pd.to_datetime(init.get("Date")).date() if _safe_str(init.get("Date")) else date.today()
                except Exception:
                    d_init = date.today()
                d = c2.date_input("Date", value=d_init, key=f"edit_date_{sel_rowid}")

                # --- Sélection en cascade (préremplie) ---
                st.caption("Sélection hiérarchique du visa")
                init_path = {"Catégorie": _safe_str(init.get("Catégorie")), "Visa": _safe_str(init.get("Visa"))}
                sel_path = cascading_visa_picker(visa_ref_full, key_prefix=f"edit_cascade_{sel_rowid}", init=init_path)
                sel_cat = sel_path.get("Catégorie","")
                visa    = sel_path.get("Visa","")

                def _f(v, alt=0.0):
                    try: return float(v)
                    except Exception: return float(alt)
                hono0  = _f(init.get(HONO, init.get("Montant", 0.0)))
                autre0 = _f(init.get(AUTRE, 0.0))
                paye0  = _f(init.get("Payé", 0.0))
                c5,c6 = st.columns(2)
                honoraires = c5.number_input("Montant honoraires (US $)", value=hono0, step=10.0, format="%.2f", key=f"edit_hono_{sel_rowid}")
                autres     = c6.number_input("Autres frais (US $)", value=autre0, step=10.0, format="%.2f", key=f"edit_autre_{sel_rowid}")
                c7,c8 = st.columns(2)
                total_preview = float(honoraires + autres); c7.metric("Total (US $)", _fmt_money_us(total_preview))
                st.caption(f"Payé actuel : {_fmt_money_us(paye0)} — Solde après sauvegarde : {_fmt_money_us(max(total_preview - paye0, 0.0))}")

                st.markdown("#### État du dossier (avec dates)")
                def _get_dt(key):
                    v = _safe_str(init.get(key))
                    try: return pd.to_datetime(v).date() if v else date.today()
                    except Exception: return date.today()

                r1c1, r1c2 = st.columns(2)
                v_env = r1c1.checkbox(S_ENVOYE, value=bool(init.get(S_ENVOYE)), key=f"edit_env_{sel_rowid}")
                dt_env = r1c2.date_input(D_ENVOYE, value=_get_dt(D_ENVOYE), disabled=not v_env, key=f"edit_dt_env_{sel_rowid}")

                r2c1, r2c2 = st.columns(2)
                v_app = r2c1.checkbox(S_APPROUVE, value=bool(init.get(S_APPROUVE)), key=f"edit_app_{sel_rowid}")
                dt_app = r2c2.date_input(D_APPROUVE, value=_get_dt(D_APPROUVE), disabled=not v_app, key=f"edit_dt_app_{sel_rowid}")

                r3c1, r3c2 = st.columns(2)
                v_rfe = r3c1.checkbox(S_RFE, value=bool(init.get(S_RFE)), key=f"edit_rfe_{sel_rowid}")
                dt_rfe = r3c2.date_input(D_RFE, value=_get_dt(D_RFE), disabled=not v_rfe, key=f"edit_dt_rfe_{sel_rowid}")

                r4c1, r4c2 = st.columns(2)
                v_ref = r4c1.checkbox(S_REFUSE, value=bool(init.get(S_REFUSE)), key=f"edit_ref_{sel_rowid}")
                dt_ref = r4c2.date_input(D_REFUSE, value=_get_dt(D_REFUSE), disabled=not v_ref, key=f"edit_dt_ref_{sel_rowid}")

                r5c1, r5c2 = st.columns(2)
                v_ann = r5c1.checkbox(S_ANNULE, value=bool(init.get(S_ANNULE)), key=f"edit_ann_{sel_rowid}")
                dt_ann = r5c2.date_input(D_ANNULE, value=_get_dt(D_ANNULE), disabled=not v_ann, key=f"edit_dt_ann_{sel_rowid}")

                ok_fiche = st.form_submit_button("💾 Enregistrer la fiche (dans le fichier)", type="primary")

            if ok_fiche:
                if v_rfe and not (v_env or v_ref or v_ann):
                    st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché."); st.stop()

                live = read_sheet(current_path, client_target_sheet, normalize=False).copy()

                # Relocalisation : ID -> Dossier N -> position
                t_idx = None
                key_id = _safe_str(init.get("ID_Client"))
                if key_id and "ID_Client" in live.columns:
                    hits = live.index[live["ID_Client"].astype(str) == key_id]
                    if len(hits)>0: t_idx = hits[0]
                if t_idx is None and (DOSSIER_COL in live.columns) and (init.get(DOSSIER_COL) not in [None, ""]):
                    try:
                        num = int(_to_int(pd.Series([init.get(DOSSIER_COL)])).iloc[0])
                        hits = live.index[_to_int(live[DOSSIER_COL]) == num]
                        if len(hits)>0: t_idx = hits[0]
                    except Exception:
                        pass
                if t_idx is None and (orig_pos is not None) and 0 <= int(orig_pos) < len(live):
                    t_idx = int(orig_pos)
                if t_idx is None:
                    st.error("Ligne introuvable."); st.stop()

                total = float((honoraires or 0.0)+(autres or 0.0))
                for c in [HONO, AUTRE, TOTAL, "Payé","Reste","Paiements", ESC_TR, ESC_JR,
                          "Nom","Date","Mois","Catégorie","Visa"] + STATUS_COLS + STATUS_DATES + [DOSSIER_COL]:
                    if c not in live.columns:
                        live[c] = 0.0 if c in [HONO,AUTRE,TOTAL,"Payé","Reste",ESC_TR] else ""
                for b in STATUS_COLS:
                    if b not in live.columns: live[b] = False

                live.at[t_idx,"Nom"]=_safe_str(nom)
                live.at[t_idx,"Date"]=str(d)
                live.at[t_idx,"Mois"]=f"{d.month:02d}"
                live.at[t_idx,"Catégorie"]=_safe_str(sel_cat)
                live.at[t_idx,"Visa"]=_safe_str(visa if isinstance(visa,str) else "")
                live.at[t_idx, HONO]=float(honoraires or 0.0)
                live.at[t_idx, AUTRE]=float(autres or 0.0)

                # Statuts + dates
                live.at[t_idx, S_ENVOYE]   = bool(v_env);  live.at[t_idx, D_ENVOYE]   = (str(dt_env) if v_env else "")
                live.at[t_idx, S_APPROUVE] = bool(v_app);  live.at[t_idx, D_APPROUVE] = (str(dt_app) if v_app else "")
                live.at[t_idx, S_RFE]      = bool(v_rfe);  live.at[t_idx, D_RFE]      = (str(dt_rfe) if v_rfe else "")
                live.at[t_idx, S_REFUSE]   = bool(v_ref);  live.at[t_idx, D_REFUSE]   = (str(dt_ref) if v_ref else "")
                live.at[t_idx, S_ANNULE]   = bool(v_ann);  live.at[t_idx, D_ANNULE]   = (str(dt_ann) if v_ann else "")

                pay_json = live.at[t_idx,"Paiements"]
                paid = _sum_payments(_parse_json_list(pay_json))
                live.at[t_idx, "Payé"]  = float(paid)
                live.at[t_idx, TOTAL]   = total
                live.at[t_idx, "Reste"] = max(total - float(paid), 0.0)

                live = ensure_dossier_numbers(live)
                write_sheet_inplace(current_path, client_target_sheet, live); save_workspace_path(current_path)
                st.success("Fiche enregistrée **dans le fichier**. ✅"); st.rerun()

            # Historique & gestion des paiements
            live_now = read_sheet(current_path, client_target_sheet, normalize=False)
            ixs = live_now.index[live_now.get("ID_Client","").astype(str)==_safe_str(init.get("ID_Client"))]
            st.markdown("#### 💳 Historique & gestion des règlements")
            if len(ixs)==0:
                st.info("Ligne introuvable pour les paiements.")
            else:
                i = ixs[0]
                if "Paiements" not in live_now.columns: live_now["Paiements"] = ""
                plist = _parse_json_list(live_now.at[i,"Paiements"])
                if plist:
                    dfp = pd.DataFrame(plist)
                    if "date" in dfp.columns: dfp["date"] = pd.to_datetime(dfp["date"], errors="coerce").dt.date.astype(str)
                    if "amount" in dfp.columns: dfp["Montant ($)"] = dfp["amount"].apply(lambda x: _fmt_money_us(float(x) if pd.notna(x) else 0.0))
                    for col in ["mode","note"]:
                        if col not in dfp.columns: dfp[col] = ""
                    show = dfp[["date","mode","Montant ($)","note"]] if set(["date","mode","note"]).issubset(dfp.columns) else dfp
                    with st.expander("Historique des règlements (cliquer pour ouvrir)", expanded=True):
                        st.table(show.rename(columns={"date":"Date","mode":"Mode","note":"Note"}))
                        if len(plist)>0:
                            del_idx = st.number_input("Supprimer la ligne n° (1..n)", min_value=1, max_value=len(plist), value=1, step=1, key=f"del_pay_idx_{i}")
                            if st.button("🗑️ Supprimer cette ligne", key=f"del_pay_btn_{i}"):
                                try:
                                    del plist[int(del_idx)-1]
                                    live_now.at[i,"Paiements"] = json.dumps(plist, ensure_ascii=False)
                                    total_paid = _sum_payments(plist)
                                    hono = _to_num(pd.Series([live_now.at[i, HONO] if HONO in live_now.columns else 0.0])).iloc[0]
                                    autr = _to_num(pd.Series([live_now.at[i, AUTRE] if AUTRE in live_now.columns else 0.0])).iloc[0]
                                    total = float(hono + autr)
                                    live_now.at[i,"Payé"]  = float(total_paid)
                                    live_now.at[i,"Reste"] = max(total - float(total_paid), 0.0)
                                    live_now.at[i,TOTAL]   = total
                                    write_sheet_inplace(current_path, client_target_sheet, live_now)
                                    st.success("Ligne supprimée et soldes recalculés. ✅"); st.rerun()
                                except Exception as e:
                                    st.error(f"Erreur suppression : {e}")
                else:
                    st.caption("Aucun paiement enregistré pour ce client.")
                cA, cB, cC, cD = st.columns([1,1,1,2])
                pay_date = cA.date_input("Date", value=date.today(), key=f"pay_date_{i}")
                pay_mode = cB.selectbox("Mode", ["CB","Chèque","Espèces","Virement","Venmo","Autre"], key=f"pay_mode_{i}")
                pay_amt  = cC.number_input("Montant ($)", min_value=0.0, step=10.0, format="%.2f", key=f"pay_amt_{i}")
                pay_note = cD.text_input("Note", "", key=f"pay_note_{i}")
                if st.button("💾 Enregistrer ce règlement (dans le fichier)", key=f"pay_add_btn_{i}"):
                    try:
                        add = float(pay_amt or 0.0)
                        if add <= 0: st.warning("Le montant doit être > 0."); st.stop()
                        norm = normalize_dataframe(live_now.copy(), visa_ref=visa_ref_simple)
                        mask_id = norm["ID_Client"].astype(str) == _safe_str(init.get("ID_Client"))
                        reste_curr = float(norm.loc[mask_id, "Reste"].sum()) if mask_id.any() else 0.0
                        if add > reste_curr + 1e-9: add = reste_curr
                        plist.append({"date": str(pay_date), "amount": float(add), "mode": pay_mode, "note": pay_note})
                        live_now.at[i,"Paiements"] = json.dumps(plist, ensure_ascii=False)
                        total_paid = _sum_payments(plist)
                        hono = _to_num(pd.Series([live_now.at[i, HONO] if HONO in live_now.columns else 0.0])).iloc[0]
                        autr = _to_num(pd.Series([live_now.at[i, AUTRE] if AUTRE in live_now.columns else 0.0])).iloc[0]
                        total = float(hono + autr)
                        live_now.at[i,"Payé"]  = float(total_paid)
                        live_now.at[i,"Reste"] = max(total - float(total_paid), 0.0)
                        live_now.at[i,TOTAL]   = total
                        write_sheet_inplace(current_path, client_target_sheet, live_now)
                        st.success("Règlement ajouté. ✅"); st.rerun()
                    except Exception as e:
                        st.error(f"Erreur : {e}")

    # --- SUPPRIMER ---
    if action == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client (écrit directement)")
        if live_raw.drop(columns=["_RowID"]).empty:
            st.info("Aucun client.")
        else:
            options = [(int(r["_RowID"]),
                        f'{int(r.get(DOSSIER_COL,0))} — { _safe_str(r.get("ID_Client")) } — { _safe_str(r.get("Nom")) }')
                       for _,r in live_raw.iterrows()]
            labels  = [lab for _,lab in options]
            sel_lab = st.selectbox("Sélection", labels, key="del_select")
            sel_rowid = [rid for rid,lab in options if lab==sel_lab][0]
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

# ================= ANALYSES (avec comparaisons + filtres cascade) =================
with tabs[2]:
    st.subheader("📊 Analyses — Volumes, Financier & Comparaisons")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa)."); st.stop()

    visa_ref_full = read_visa_reference_hier(current_path)
    visa_ref_simple = read_visa_reference(current_path)
    dfA_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    dfA = normalize_dataframe(dfA_raw, visa_ref=visa_ref_simple).copy()
    if dfA.empty: st.info("Aucune donnée pour analyser."); st.stop()

    # Filtres (avec cascade)
    with st.container():
        cL, cR = st.columns([1,2])
        cL.caption("Sélection hiérarchique (réduit la liste des visas)")
        with cL:
            sel_path_anal = cascading_visa_picker(visa_ref_full, key_prefix="anal_cascade")
        visas_aut = visas_autorises_depuis_cascade(visa_ref_full, sel_path_anal)

        cR1, cR2, cR3 = cR.columns(3)
        yearsA  = sorted({d.year for d in dfA["Date"] if pd.notna(d)}) if "Date" in dfA.columns else []
        sel_years  = cR1.multiselect("Année (filtre)", yearsA, default=[])
        monthsA = [f"{m:02d}" for m in range(1,13)]
        sel_months = cR2.multiselect("Mois (MM)", monthsA, default=[])
        include_na_dates = cR3.checkbox("Inclure lignes sans date", value=True)

    fA = dfA.copy()
    if sel_path_anal.get("Catégorie", ""):
        fA = fA[fA["Catégorie"].astype(str) == sel_path_anal["Catégorie"]]
    if visas_aut:
        fA = fA[fA["Visa"].astype(str).isin(visas_aut)]
    if "Date" in fA.columns and sel_years:
        mask_year = fA["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask_year = mask_year | fA["Date"].isna()
        fA = fA[mask_year]
    if "Mois" in fA.columns and sel_months:
        mask_month = fA["Mois"].isin(sel_months)
        if include_na_dates: mask_month = mask_month | fA["Mois"].isna()
        fA = fA[mask_month]

    # Colonnes temps & numériques
    fA["Année"] = fA["Date"].apply(lambda x: x.year if pd.notna(x) else pd.NA)
    fA["MoisNum"] = fA["Date"].apply(lambda x: int(x.month) if pd.notna(x) else pd.NA)
    fA["Periode"] = fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")

    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in fA.columns: fA[col] = pd.to_numeric(fA[col], errors="coerce").fillna(0.0)

    def derive_statut(row) -> str:
        if bool(row.get(S_APPROUVE, False)): return "Approuvé"
        if bool(row.get(S_REFUSE, False)):   return "Refusé"
        if bool(row.get(S_ANNULE, False)):   return "Annulé"
        return "En attente"
    fA["Statut"] = fA.apply(derive_statut, axis=1)

    # KPI
    st.markdown("""
    <style>.small-kpi [data-testid="stMetricValue"]{font-size:1.15rem}.small-kpi [data-testid="stMetricLabel"]{font-size:.8rem;opacity:.8}</style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="small-kpi">', unsafe_allow_html=True)
    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Dossiers", f"{len(fA)}")
    k2.metric("Total (US $)", _fmt_money_us(float(fA.get(TOTAL, pd.Series(dtype=float)).sum())) )
    k3.metric("Payé (US $)", _fmt_money_us(float(fA.get("Payé", pd.Series(dtype=float)).sum())) )
    k4.metric("Solde (US $)", _fmt_money_us(float(fA.get("Reste", pd.Series(dtype=float)).sum())) )
    st.markdown('</div>', unsafe_allow_html=True)

    # Volumes simples
    st.markdown("### 📈 Volumes (créations)")
    vol_crees = fA.groupby("Periode").size().reset_index(name="Créés")
    df_vol = vol_crees.rename(columns={"Créés":"Volume"}).assign(Indic="Créés")
    if not df_vol.empty:
        try:
            st.altair_chart(
                alt.Chart(df_vol).mark_line(point=True).encode(
                    x=alt.X("Periode:N", sort=None, title="Période"),
                    y=alt.Y("Volume:Q"),
                    color=alt.Color("Indic:N", legend=alt.Legend(title="Statut")),
                    tooltip=["Periode","Indic","Volume"]
                ).properties(height=260), use_container_width=True
            )
        except Exception:
            st.dataframe(df_vol, use_container_width=True)

    st.divider()

    # ==================== COMPARAISONS ====================
    st.markdown("## 🔁 Comparaisons (YoY & MoM)")

    # 1) YOY Global
    st.markdown("### 🗓️ Année sur année — Global")
    by_year = fA.dropna(subset=["Année"]).groupby("Année").agg(
        Dossiers=("Nom","count"),
        Honoraires=(HONO,"sum"),
        Autres=(AUTRE,"sum"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index().sort_values("Année")
    c1, c2 = st.columns(2)
    if not by_year.empty:
        try:
            c1.altair_chart(
                alt.Chart(by_year.melt("Année", ["Dossiers"])).mark_bar().encode(
                    x=alt.X("Année:N"), y=alt.Y("value:Q", title="Volume"),
                    color=alt.Color("variable:N", legend=None),
                    tooltip=["Année","value"]
                ).properties(title="Nombre de dossiers", height=260), use_container_width=True
            )
        except Exception:
            c1.dataframe(by_year[["Année","Dossiers"]], use_container_width=True)

        try:
            metric_vars = ["Honoraires","Autres","Total","Payé","Reste"]
            yo = by_year.melt("Année", metric_vars, var_name="Indicateur", value_name="Montant")
            c2.altair_chart(
                alt.Chart(yo).mark_bar().encode(
                    x=alt.X("Année:N"),
                    y=alt.Y("Montant:Q"),
                    color=alt.Color("Indicateur:N"),
                    tooltip=["Année","Indicateur", alt.Tooltip("Montant:Q", format="$.2f")]
                ).properties(title="Montants par année", height=260), use_container_width=True
            )
        except Exception:
            c2.dataframe(by_year.drop(columns=["Dossiers"]), use_container_width=True)
    else:
        st.info("Pas de dates exploitables pour la comparaison annuelle.")

    # 2) YOY par mois
    st.markdown("### 📅 Mois (1..12) — Année sur année")
    by_year_month = fA.dropna(subset=["Année","MoisNum"]).groupby(["Année","MoisNum"]).agg(
        Dossiers=("Nom","count"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index()

    c3, c4 = st.columns(2)
    if not by_year_month.empty:
        try:
            c3.altair_chart(
                alt.Chart(by_year_month).mark_line(point=True).encode(
                    x=alt.X("MoisNum:O", title="Mois"),
                    y=alt.Y("Dossiers:Q"),
                    color=alt.Color("Année:N"),
                    tooltip=["Année","MoisNum","Dossiers"]
                ).properties(title="Dossiers par mois (YoY)", height=260), use_container_width=True
            )
        except Exception:
            c3.dataframe(by_year_month.pivot(index="MoisNum", columns="Année", values="Dossiers"), use_container_width=True)

        try:
            c4.altair_chart(
                alt.Chart(by_year_month.melt(["Année","MoisNum"], ["Total","Payé","Reste"],
                                             var_name="Indicateur", value_name="Montant")
                ).mark_line(point=True).encode(
                    x=alt.X("MoisNum:O", title="Mois"),
                    y=alt.Y("Montant:Q"),
                    color=alt.Color("Année:N"),
                    tooltip=["Année","MoisNum","Indicateur", alt.Tooltip("Montant:Q", format="$.2f")]
                ).properties(title="Montants par mois (YoY)", height=260),
                use_container_width=True
            )
        except Exception:
            c4.dataframe(by_year_month.pivot_table(index="MoisNum", columns="Année", values="Total"), use_container_width=True)
    else:
        st.info("Pas de séries mois/année disponibles.")

    # 3) Par type de visa — YOY
    st.markdown("### 🛂 Par type de visa — Année sur année")
    topN = st.slider("Top N visas (par Total)", 3, 20, 10, 1, key="cmp_topn")
    metric_cmp = st.selectbox("Indicateur", ["Dossiers","Total","Payé","Reste","Honoraires","Autres"], index=1, key="cmp_metric")

    by_year_visa = fA.dropna(subset=["Année"]).groupby(["Année","Visa"]).agg(
        Dossiers=("Nom","count"),
        Honoraires=(HONO,"sum"),
        Autres=(AUTRE,"sum"),
        Total=(TOTAL,"sum"),
        Payé=("Payé","sum"),
        Reste=("Reste","sum"),
    ).reset_index()

    top_visas = (by_year_visa.groupby("Visa")["Total"].sum()
                 .sort_values(ascending=False).head(topN).index.tolist())
    by_year_visa_top = by_year_visa[by_year_visa["Visa"].isin(top_visas)].copy()

    if not by_year_visa_top.empty:
        try:
            st.altair_chart(
                alt.Chart(by_year_visa_top).mark_bar().encode(
                    x=alt.X("Visa:N", sort=top_visas),
                    y=alt.Y(f"{metric_cmp}:Q"),
                    color=alt.Color("Année:N"),
                    tooltip=["Visa","Année", alt.Tooltip(f"{metric_cmp}:Q", format="$.2f" if metric_cmp!="Dossiers" else "")],
                ).properties(height=300), use_container_width=True
            )
        except Exception:
            st.dataframe(by_year_visa_top.pivot_table(index="Visa", columns="Année", values=metric_cmp, aggfunc="sum"),
                         use_container_width=True)
    else:
        st.info("Pas de données par type de visa après filtres.")

    # 4) Comparer 2 années — barres par mois
    st.markdown("### ⚖️ Comparer 2 années — barres par mois")
    years_all = sorted([y for y in fA["Année"].dropna().unique()])
    colY1, colY2, colMet = st.columns([1,1,2])
    y1 = colY1.selectbox("Année A", years_all, index=0 if len(years_all)==0 else 0, key="cmp_y1")
    y2 = colY2.selectbox("Année B", years_all, index=1 if len(years_all)>1 else 0, key="cmp_y2")
    metric2 = colMet.selectbox("Indicateur", ["Dossiers","Total","Payé","Reste","Honoraires","Autres"], index=1, key="cmp_metric2")

    if y1 == y2:
        st.caption("Choisis deux années différentes pour cette vue.")
    else:
        b = by_year_month[by_year_month["Année"].isin([y1, y2])].copy()
        if metric2 == "Dossiers":
            plot_df = b[["Année","MoisNum","Dossiers"]].rename(columns={"Dossiers":"Valeur"})
        else:
            plot_df = b[["Année","MoisNum", metric2]].rename(columns={metric2:"Valeur"})
        if not plot_df.empty:
            try:
                st.altair_chart(
                    alt.Chart(plot_df).mark_bar().encode(
                        x=alt.X("MoisNum:O", title="Mois"),
                        y=alt.Y("Valeur:Q"),
                        column=alt.Column("Année:N", title=""),
                        tooltip=["Année","MoisNum", alt.Tooltip("Valeur:Q", format="$.2f" if metric2!="Dossiers" else "")]
                    ).properties(height=280), use_container_width=True
                )
            except Exception:
                st.dataframe(plot_df.pivot(index="MoisNum", columns="Année", values="Valeur"), use_container_width=True)

    # Détails
    st.divider()
    st.markdown("### 🔎 Détails (clients)")
    details_cols = [c for c in ["Periode",DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Date", HONO, AUTRE, TOTAL, "Payé","Reste","Statut","Année","MoisNum"] if c in fA.columns]
    details = fA[details_cols].copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in details.columns: details[col] = details[col].apply(lambda x: _fmt_money_us(x) if pd.notna(x) else "")
    st.dataframe(details.sort_values(["Année","MoisNum","Catégorie","Nom"]), use_container_width=True)

# ================= ESCROW =================
with tabs[3]:
    st.subheader("🏦 ESCROW — dépôts sur honoraires & transferts")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa)."); st.stop()
    visa_ref_simple = read_visa_reference(current_path)
    live_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    live = normalize_dataframe(live_raw, visa_ref=visa_ref_simple).copy()
    if ESC_TR not in live.columns: live[ESC_TR] = 0.0
    else: live[ESC_TR] = pd.to_numeric(live[ESC_TR], errors="coerce").fillna(0.0)
    live["ESCROW dispo"] = live.apply(escrow_available_from_row, axis=1)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Dossiers", f"{len(live)}")
    c2.metric("ESCROW total dispo", _fmt_money_us(float(live["ESCROW dispo"].sum())))
    envoyes = live[(live[S_ENVOYE]==True)]
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
                        tmp = normalize_dataframe(live_w.copy(), visa_ref=visa_ref_simple)
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
    non_env = live[(live[S_ENVOYE]!=True) & (live["ESCROW dispo"]>0.004)].copy()
    if non_env.empty:
        st.info("Rien en attente côté dossiers non envoyés.")
    else:
        show = non_env[[DOSSIER_COL,"ID_Client","Nom","Catégorie","Visa","Date",HONO,"Payé",ESC_TR,"ESCROW dispo"]].copy()
        for col in [HONO,"Payé",ESC_TR,"ESCROW dispo"]:
            show[col] = show[col].map(_fmt_money_us)
        st.dataframe(show, use_container_width=True)

    st.divider()
    st.markdown("### 🧾 Historique des transferts (journal)")
    has_journal = live[live[ESC_JR].astype(str).str.len() > 0]

    if has_journal.empty:
        st.caption("Aucun journal de transfert pour le moment.")
    else:
        rows = []
        for _, r in has_journal.iterrows():
            entries = _parse_json_list(r.get(ESC_JR, ""))
            for e in entries:
                rows.append({
                    DOSSIER_COL: r.get(DOSSIER_COL, ""),
                    "ID_Client": r.get("ID_Client", ""),
                    "Nom": r.get("Nom", ""),
                    "Visa": r.get("Visa", ""),
                    "Date": r.get("Date", ""),
                    "Horodatage": e.get("ts", ""),
                    "Montant (US $)": float(e.get("amount", 0.0)),
                    "Note": e.get("note", "")
                })
        jdf = pd.DataFrame(rows)
        if jdf.empty:
            st.caption("Aucun mouvement enregistré dans le journal.")
        else:
            if "Horodatage" in jdf.columns:
                jdf["Horodatage_dt"] = pd.to_datetime(jdf["Horodatage"], errors="coerce")
                jdf = jdf.sort_values(
                    by=["Horodatage_dt", DOSSIER_COL, "ID_Client"],
                    na_position="last"
                ).drop(columns=["Horodatage_dt"])
            jdf["Montant (US $)"] = jdf["Montant (US $)"].apply(lambda x: _fmt_money_us(float(x) if pd.notna(x) else 0.0))
            st.dataframe(jdf, use_container_width=True)
