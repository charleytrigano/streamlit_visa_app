# app.py
import io
import json
import hashlib
from datetime import date, datetime, timedelta
from pathlib import Path
import zipfile

import streamlit as st
import pandas as pd

st.set_page_config(page_title="📊 Visas — Edition directe + ESCROW", layout="wide")
st.title("📊 Visas — Edition DIRECTE du fichier (avec ESCROW)")

# ========================= Constantes colonnes =========================
HONO   = "Honoraires (US $)"
AUTRE  = "Autres frais (US $)"
TOTAL  = "Total (US $)"
ESC_TR = "Escrow transféré (US $)"
ESC_JR = "Escrow journal"

# ========================= Espace de travail =========================
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
        obj = json.loads(WS_FILE.read_text(encoding="utf-8"))
        p = Path(obj.get("last_path", ""))
        return p if p.exists() else None
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
            if not cand.exists(): dest = cand; break
            n += 1
    data = upload.read()
    dest.write_bytes(data)
    return dest

# ========================= Utils =========================
def _safe_str(x): return "" if pd.isna(x) else str(x).strip()

def _to_num(s: pd.Series) -> pd.Series:
    cleaned = (s.astype(str)
                 .str.replace("\u00a0", "", regex=False)
                 .str.replace("\u202f", "", regex=False)
                 .str.replace(" ", "", regex=False)
                 .str.replace("$", "", regex=False)
                 .str.replace(",", "", regex=False))
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)

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
    has_visa = "visa" in cols
    no_money = not ({"montant", "honoraires", "acomptes", "payé", "reste", "solde"} & cols)
    return has_visa and no_money

def is_clients_like(df: pd.DataFrame) -> bool:
    cols = set(df.columns.astype(str))
    return {"Nom","Visa"}.issubset(cols)

# ---------- ESCROW helpers ----------
def escrow_available_from_row(row) -> float:
    """ESCROW dispo = min(Payé, Honoraires) - Escrow transféré (>=0)"""
    try: hon = float(row.get(HONO, 0.0))
    except Exception: hon = 0.0
    try: paid = float(row.get("Payé", 0.0))
    except Exception: paid = 0.0
    try: moved = float(row.get(ESC_TR, 0.0))
    except Exception: moved = 0.0
    return max(min(paid, hon) - moved, 0.0)

def append_escrow_journal(row_raw: pd.Series, amount: float, note: str = "") -> str:
    journal = _parse_json_list(row_raw.get(ESC_JR, ""))
    journal.append({
        "ts": datetime.now().isoformat(timespec="seconds"),
        "amount": float(amount),
        "note": note
    })
    return json.dumps(journal, ensure_ascii=False)

# ========================= Normalisation =========================
def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Date / Mois
    if "Date" in df.columns: df["Date"] = _to_date(df["Date"])
    else: df["Date"] = pd.NaT
    df["Mois"] = df["Date"].apply(lambda x: f"{x.month:02d}" if pd.notna(x) else pd.NA)

    # Visa
    visa_col = None
    for c in ["Visa", "Categories", "Catégorie", "TypeVisa"]:
        if c in df.columns: visa_col = c; break
    df["Visa"] = df[visa_col].astype(str) if visa_col else "Inconnu"

    # Migration "Montant" -> HONO si besoin
    if "Montant" in df.columns and HONO not in df.columns:
        df[HONO] = _to_num(df["Montant"])
    else:
        if HONO in df.columns: df[HONO] = _to_num(df[HONO])
        else: df[HONO] = 0.0

    # Autres frais
    if AUTRE in df.columns: df[AUTRE] = _to_num(df[AUTRE])
    else: df[AUTRE] = 0.0

    # Total
    df[TOTAL] = (df[HONO] + df[AUTRE]).astype(float)

    # Payé
    if "Payé" in df.columns:
        df["Payé"] = _to_num(df["Payé"])
    else:
        if "Paiements" in df.columns:
            parsed = df["Paiements"].apply(_parse_json_list)
            df["Payé"] = parsed.apply(_sum_payments).astype(float)
        else:
            df["Payé"] = 0.0

    # Reste
    df["Reste"] = (df[TOTAL] - df["Payé"]).fillna(0.0)

    # Statuts
    for b in ["RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"]:
        if b not in df.columns: df[b] = False

    # Identité / Paiements
    if "Nom" not in df.columns: df["Nom"] = ""
    if "ID_Client" not in df.columns: df["ID_Client"] = ""
    need_id = df["ID_Client"].astype(str).str.strip().eq("") | df["ID_Client"].isna()
    if need_id.any():
        df.loc[need_id, "ID_Client"] = df.loc[need_id].apply(_make_client_id_from_row, axis=1)
    if "Paiements" not in df.columns: df["Paiements"] = ""

    # ESCROW
    if ESC_TR in df.columns: df[ESC_TR] = _to_num(df[ESC_TR])
    else: df[ESC_TR] = 0.0
    if ESC_JR not in df.columns: df[ESC_JR] = ""

    # Nettoyage
    for dropcol in ["Telephone","Email"]:
        if dropcol in df.columns: df = df.drop(columns=[dropcol])

    ordered = ["ID_Client","Nom","Date","Mois","Visa",
               HONO, AUTRE, TOTAL, "Payé","Reste",
               ESC_TR, ESC_JR,
               "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé","Paiements"]
    cols = [c for c in ordered if c in df.columns] + [c for c in df.columns if c not in ordered]
    return df[cols]

# ========================= IO Excel =========================
def read_sheet(path: Path, sheet: str, normalize: bool) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    if sheet not in xls.sheet_names:
        base = pd.DataFrame(columns=[
            "ID_Client","Nom","Date","Mois","Visa",
            HONO, AUTRE, TOTAL, "Payé","Reste",
            ESC_TR, ESC_JR,
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
            df2 = df.copy()
            if isinstance(df2, pd.Series): df2 = df2.to_frame()
            df2.to_excel(writer, sheet_name="Analyses", index=True, startrow=startrow)
            startrow += (len(df2) + 3)
    bytes_out = out.getvalue()
    path.write_bytes(bytes_out)
    try:
        st.session_state["download_bytes"] = bytes_out
        st.session_state["download_name"] = path.name
    except Exception:
        pass

# ========================= Source (sidebar) =========================
st.sidebar.header("Source")
current_path = load_workspace_path()
if current_path and current_path.exists():
    st.sidebar.success(f"Fichier courant : {current_path.name}")
else:
    defaults = ["/mnt/data/Visa_Clients_20251001-114844.xlsx", "/mnt/data/visa_analytics_datecol.xlsx"]
    cand = next((Path(p) for p in defaults if Path(p).exists()), None)
    if cand:
        current_path = cand; save_workspace_path(current_path)
        st.sidebar.success(f"Fichier courant : {current_path.name}")
    else:
        st.sidebar.warning("Aucun fichier trouvé. Importez un Excel pour démarrer.")

up = st.sidebar.file_uploader("Remplacer par un Excel (.xlsx, .xls)", type=["xlsx","xls"])
if up is not None:
    new_path = copy_upload_to_workspace(up); save_workspace_path(new_path)
    try:
        st.session_state["download_bytes"] = new_path.read_bytes()
        st.session_state["download_name"] = new_path.name
    except Exception:
        st.session_state["download_bytes"] = b""; st.session_state["download_name"] = new_path.name
    st.sidebar.success(f"Nouveau fichier chargé : {new_path.name}")
    st.rerun()

if current_path is None or not current_path.exists(): st.stop()

if "download_bytes" not in st.session_state or st.session_state.get("download_name") != current_path.name:
    try:
        st.session_state["download_bytes"] = current_path.read_bytes()
        st.session_state["download_name"] = current_path.name
    except Exception:
        st.session_state["download_bytes"] = b""; st.session_state["download_name"] = current_path.name

try:
    sheet_names = pd.ExcelFile(current_path).sheet_names
except Exception as e:
    st.error(f"Impossible de lire l'Excel : {e}"); st.stop()

# Feuille Dashboard libre
preferred_order = ["Clients","Visa","Données normalisées"]
default_sheet = next((s for s in preferred_order if s in sheet_names), sheet_names[0])
sheet_choice = st.sidebar.selectbox("Feuille (Dashboard)", sheet_names, index=sheet_names.index(default_sheet))

# Feuilles valides pour CRUD
valid_client_sheets = []
for s in sheet_names:
    try:
        df_tmp = read_sheet(current_path, s, normalize=False)
        if is_clients_like(df_tmp): valid_client_sheets.append(s)
    except Exception: pass

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
    if len(valid_client_sheets) < len(sheet_names):
        st.sidebar.caption("🔒 Seules les feuilles contenant **Nom** et **Visa** sont proposées ici.")

st.sidebar.caption(f"Édition **directe** dans : `{current_path}`")
st.sidebar.download_button(
    "⬇️ Télécharger une copie",
    data=st.session_state.get("download_bytes", b""),
    file_name=st.session_state.get("download_name", current_path.name),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

# ========================= Onglets =========================
tabs = st.tabs(["Dashboard", "Clients (CRUD)", "Analyses", "ESCROW"])

# ========================= DASHBOARD =========================
with tabs[0]:
    df_raw = read_sheet(current_path, sheet_choice, normalize=False)
    if looks_like_reference(df_raw):
        st.subheader("📄 Référentiel — Types de Visa")
        if "Visa" not in df_raw.columns: df_raw = pd.DataFrame(columns=["Visa"])
        df_ref = pd.DataFrame({"Visa": df_raw["Visa"].astype(str).fillna("").str.strip()})
        st.dataframe(df_ref, use_container_width=True)

        st.markdown("### ✏️ Gérer les types")
        action = st.radio("Action", ["Ajouter", "Renommer", "Supprimer"], horizontal=True, key="visa_ref_action")
        options = sorted([v for v in df_ref["Visa"].unique() if v], key=str.lower)
        if action == "Ajouter":
            new_v = st.text_input("Nouveau type de visa").strip()
            if st.button("➕ Ajouter"):
                if not new_v: st.warning("Saisis un libellé.")
                elif new_v in options: st.info("Ce type existe déjà.")
                else:
                    out = pd.concat([df_ref, pd.DataFrame([{"Visa": new_v}])], ignore_index=True)
                    write_sheet_inplace(current_path, sheet_choice, out); st.success("Type ajouté."); st.rerun()
        elif action == "Renommer":
            if not options: st.info("Aucun type existant.")
            else:
                old = st.selectbox("Type à renommer", options)
                new = st.text_input("Nouveau libellé").strip()
                if st.button("📝 Renommer"):
                    if not new: st.warning("Nouveau libellé requis.")
                    elif new == old: st.info("Aucun changement.")
                    elif new in options: st.info("Un type avec ce nom existe déjà.")
                    else:
                        out = df_ref.copy(); out.loc[out["Visa"] == old, "Visa"] = new
                        write_sheet_inplace(current_path, sheet_choice, out); st.success("Type renommé."); st.rerun()
        else:
            if not options: st.info("Aucun type à supprimer.")
            else:
                rm = st.selectbox("Type à supprimer", options)
                st.error("⚠️ Action irréversible.")
                if st.button("🗑️ Supprimer"):
                    out = df_ref[df_ref["Visa"] != rm].reset_index(drop=True)
                    write_sheet_inplace(current_path, sheet_choice, out); st.success("Type supprimé."); st.rerun()
        st.stop()

    # Dashboard classique
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
                container.caption(f"{lab} : aucune donnée"); return None
            vmin, vmax = float(_df[col].min()), float(_df[col].max())
            if not (vmin < vmax):
                container.caption(f"{lab} : valeur unique = {_fmt_money_us(vmin)}"); return (vmin, vmax)
            step = 1.0 if (vmax - vmin) > 1000 else 0.1 if (vmax - vmin) > 10 else 0.01
            return container.slider(lab, min_value=vmin, max_value=vmax, value=(vmin, vmax), step=step)

        total_range = make_slider(df, TOTAL, "Total (US $) min-max", c4)
        pay_range   = make_slider(df, "Payé", "Payé (US $) min-max", c5)
        reste_range = make_slider(df, "Reste", "Solde (US $) min-max", c6)

    f = df.copy()
    if "Date" in f.columns and sel_years:
        mask = f["Date"].apply(lambda x: (pd.notna(x) and x.year in sel_years))
        if include_na_dates: mask |= f["Date"].isna()
        f = f[mask]
    if "Mois" in f.columns and sel_months:
        mask = f["Mois"].isin(sel_months)
        if include_na_dates: mask |= f["Mois"].isna()
        f = f[mask]
    if "Visa" in f.columns and sel_visas: f = f[f["Visa"].astype(str).isin(sel_visas)]
    if TOTAL in f.columns and total_range is not None:
        f = f[(f[TOTAL] >= total_range[0]) & (f[TOTAL] <= total_range[1])]
    if "Payé" in f.columns and pay_range is not None:
        f = f[(f["Payé"] >= pay_range[0]) & (f["Payé"] <= pay_range[1])]
    if "Reste" in f.columns and reste_range is not None:
        f = f[(f["Reste"] >= reste_range[0]) & (f["Reste"] <= reste_range[1])]

    hidden = len(df) - len(f)
    if hidden > 0: st.caption(f"🔎 {hidden} ligne(s) masquée(s) par les filtres.")

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

    # ⚠️ Alerte ESCROW (dossiers envoyés à réclamer)
    df_esc = f.copy()
    if ESC_TR not in df_esc.columns: df_esc[ESC_TR] = 0.0
    else: df_esc[ESC_TR] = pd.to_numeric(df_esc[ESC_TR], errors="coerce").fillna(0.0)
    df_esc["escrow_dispo"] = df_esc.apply(escrow_available_from_row, axis=1)
    alert = df_esc[(df_esc.get("Dossier envoyé", False) == True) & (df_esc["escrow_dispo"] > 0.004)]
    if not alert.empty:
        total_alert = float(alert["escrow_dispo"].sum())
        st.warning(f"💡 ESCROW à réclamer : {_fmt_money_us(total_alert)} sur {len(alert)} dossier(s) **envoyés**.")
        st.dataframe(alert[["ID_Client","Nom","Visa","Date",HONO,"Payé",ESC_TR,"escrow_dispo"]].assign(
            **{"escrow_dispo": alert["escrow_dispo"].map(_fmt_money_us),
               HONO: alert[HONO].map(_fmt_money_us),
               "Payé": alert["Payé"].map(_fmt_money_us),
               ESC_TR: alert[ESC_TR].map(_fmt_money_us)}
        ), use_container_width=True)

    st.divider()

    # =============== ➕ Ajouter un paiement (US $) — dossiers non soldés ===============
    st.subheader("➕ Ajouter un paiement (US $)")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide dans la sidebar.")
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
                    # Localise la ligne
                    target_id = label_to_id.get(sel_label, "")
                    idxs = live.index[live.get("ID_Client","").astype(str) == str(target_id)]
                    if len(idxs)==0:
                        raise RuntimeError("Dossier introuvable.")
                    idx = idxs[0]

                    # Liste paiements existante
                    pay_list = _parse_json_list(live.at[idx, "Paiements"])
                    add = float(amount or 0.0)
                    if add <= 0:
                        st.warning("Le montant doit être > 0."); st.stop()

                    # Sécurité : ne pas dépasser le reste
                    live_norm = normalize_dataframe(live.copy())
                    idc = str(live.at[idx, "ID_Client"]) if "ID_Client" in live.columns else ""
                    reste_curr = float(live_norm.loc[live_norm["ID_Client"].astype(str)==idc, "Reste"].iloc[0]) if idc else 0.0
                    if add > reste_curr + 1e-9:
                        add = reste_curr

                    # Ajoute le paiement + recalcul Payé/Reste/Total
                    pay_list.append({"date": str(pdate), "amount": float(add), "mode": mode, "note": note})
                    live.at[idx, "Paiements"] = json.dumps(pay_list, ensure_ascii=False)

                    # Garantir colonnes
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

    # ==================== Tableau ====================
    st.subheader("📋 Données (aperçu)")
    cols_show = [c for c in ["ID_Client","Nom","Date","Visa", HONO, AUTRE, TOTAL, "Payé","Reste",
                             "RFE","Dossier envoyé","Dossier approuvé","Dossier refusé","Dossier annulé"] if c in f.columns]
    table = f.copy()
    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in table.columns: table[col] = table[col].map(_fmt_money_us)
    if "Date" in table.columns: table["Date"] = table["Date"].astype(str)
    st.dataframe(table[cols_show].sort_values(by=[c for c in ["Date","Visa"] if c in table.columns], na_position="last"),
                 use_container_width=True)

# ========================= CLIENTS (CRUD) =========================
with tabs[1]:
    st.subheader("👤 Clients — Créer / Modifier / Supprimer (écriture **directe**)")
    if client_target_sheet is None:
        st.warning("Aucune feuille *Clients* valide disponible. Ajoute une feuille avec au moins les colonnes **Nom** et **Visa**.")
        st.stop()

    if st.button("🔄 Recharger le fichier"): st.rerun()
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
        for must in ["ID_Client","Nom","Date","Mois","Visa",
                     HONO, AUTRE, TOTAL, "Payé","Reste", ESC_TR, ESC_JR,
                     "Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé","Paiements"]:
            if must not in live_raw.columns:
                if must in {HONO, AUTRE, TOTAL, "Payé","Reste", ESC_TR}: live_raw[must]=0.0
                elif must in {"Paiements", ESC_JR}: live_raw[must]=""
                elif must in {"Dossier envoyé","Dossier approuvé","RFE","Dossier refusé","Dossier annulé"}: live_raw[must]=False
                elif must=="Mois": live_raw[must]=""
                else: live_raw[must]=""

        with st.form("create_form", clear_on_submit=False):
            c1,c2 = st.columns(2)
            nom = c1.text_input("Nom")
            d = c2.date_input("Date", value=date.today())
            visa = st.selectbox("Visa", visa_options, index=0) if visa_options else st.text_input("Visa")
            c5,c6 = st.columns(2)
            honoraires = c5.number_input("Montant honoraires (US $)", value=0.0, step=10.0, format="%.2f")
            autres     = c6.number_input("Autres frais (US $)", value=0.0, step=10.0, format="%.2f")
            c7,c8 = st.columns(2)
            total_preview = float(honoraires + autres)
            c7.metric("Total (US $)", _fmt_money_us(total_preview))
            paye_init = c8.number_input("Payé (US $)", value=0.0, step=10.0, format="%.2f")

            st.markdown("#### État du dossier")
            val_envoye = st.checkbox("Dossier envoyé",  value=False) if has_envoye else False
            val_appr   = st.checkbox("Dossier approuvé",value=False) if has_appr   else False
            val_rfe    = st.checkbox("RFE",             value=False) if has_rfe    else False
            val_refuse = st.checkbox("Dossier refusé",  value=False) if has_refuse else False
            val_annule = st.checkbox("Dossier annulé",  value=False) if has_annule else False

            ok = st.form_submit_button("💾 Sauvegarder (dans le fichier)", type="primary")

        if ok:
            if val_rfe and not (val_envoye or val_refuse or val_annule):
                st.error("RFE ⇢ seulement si Envoyé/Refusé/Annulé est coché."); st.stop()

            gen_id = _make_client_id_from_row({"Nom": nom, "Date": d})
            existing = set(live_raw["ID_Client"].astype(str)) if "ID_Client" in live_raw.columns else set()
            new_id = gen_id; n=1
            while new_id in existing: n+=1; new_id=f"{gen_id}-{n:02d}"

            total = float((honoraires or 0.0)+(autres or 0.0))
            reste = max(total - float(paye_init or 0.0), 0.0)

            new_row = {
                "ID_Client": new_id, "Nom": _safe_str(nom), "Date": str(d), "Mois": f"{d.month:02d}",
                "Visa": _safe_str(visa),
                HONO: float(honoraires or 0.0), AUTRE: float(autres or 0.0), TOTAL: total,
                "Payé": float(paye_init or 0.0), "Reste": reste,
                ESC_TR: 0.0, ESC_JR: "",
                "Paiements": "",
                "Dossier envoyé": bool(val_envoye), "Dossier approuvé": bool(val_appr),
                "RFE": bool(val_rfe), "Dossier refusé": bool(val_refuse), "Dossier annulé": bool(val_annule),
            }
            live_after = pd.concat([live_raw.drop(columns=["_RowID"]), pd.DataFrame([new_row])], ignore_index=True)
            write_sheet_inplace(current_path, client_target_sheet, live_after); save_workspace_path(current_path)
            st.success("Client créé **dans le fichier**. ✅"); st.rerun()

    # ---- MODIFIER ----
    if action == "Modifier":
        st.markdown("### ✏️ Modifier un client")
        if live_raw.drop(columns=["_RowID"]).empty: st.info("Aucun client.")
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

                def _f(v, alt=0.0):
                    try: return float(v)
                    except Exception: return float(alt)

                hono0   = _f(init.get(HONO, init.get("Montant", 0.0)))
                autre0  = _f(init.get(AUTRE, 0.0))
                paye0   = _f(init.get("Payé", 0.0))
                moved0  = _f(init.get(ESC_TR, 0.0))

                c5,c6 = st.columns(2)
                honoraires = c5.number_input("Montant honoraires (US $)", value=hono0, step=10.0, format="%.2f")
                autres     = c6.number_input("Autres frais (US $)", value=autre0, step=10.0, format="%.2f")

                c7,c8 = st.columns(2)
                total_preview = float(honoraires + autres)
                c7.metric("Total (US $)", _fmt_money_us(total_preview))
                paye    = c8.number_input("Payé (US $)", value=paye0, step=10.0, format="%.2f")

                st.caption(f"ESCROW transféré (cumul) actuellement : {_fmt_money_us(moved0)} — (gérer les transferts dans l’onglet ESCROW)")

                st.markdown("#### État du dossier")
                val_envoye = st.checkbox("Dossier envoyé",  value=bool(init.get("Dossier envoyé")))   if has_envoye else False
                val_appr   = st.checkbox("Dossier approuvé",value=bool(init.get("Dossier approuvé"))) if has_appr   else False
                val_rfe    = st.checkbox("RFE",             value=bool(init.get("RFE")))              if has_rfe    else False
                val_refuse = st.checkbox("Dossier refusé",  value=bool(init.get("Dossier refusé")))   if has_refuse else False
                val_annule = st.checkbox("Dossier annulé",  value=bool(init.get("Dossier annulé")))   if has_annule else False

                # --- Alerte & transfert immédiat si 'Envoyé' ---
                escrow_dispo_after = max(min(float(paye or 0.0), float(honoraires or 0.0)) - float(moved0), 0.0)
                do_transfer_now = False
                transfer_amount = 0.0
                transfer_note = ""
                if val_envoye and escrow_dispo_after > 0.004:
                    st.warning(f"Ce dossier est *envoyé* : ESCROW à transférer possible = {_fmt_money_us(escrow_dispo_after)}")
                    do_transfer_now = st.checkbox("Transférer maintenant vers compte ordinaire ?", value=True)
                    transfer_amount = st.number_input("Montant à transférer (US $)", min_value=0.0,
                                                      max_value=float(escrow_dispo_after), value=float(escrow_dispo_after),
                                                      step=10.0, format="%.2f")
                    transfer_note = st.text_input("Note de transfert (facultatif)", "")

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
                    live.at[t_idx,"Visa"]=_safe_str(visa)
                    live.at[t_idx, HONO]=float(honoraires or 0.0)
                    live.at[t_idx, AUTRE]=float(autres or 0.0)
                    if TOTAL in live.columns: live.at[t_idx, TOTAL]=total
                    else:
                        live[TOTAL] = live.get(TOTAL, 0.0); live.at[t_idx, TOTAL]=total
                    live.at[t_idx,"Payé"]=float(paye or 0.0)
                    live.at[t_idx,"Reste"]=max(total - float(paye or 0.0), 0.0)
                    if ESC_TR not in live.columns: live[ESC_TR] = 0.0
                    if ESC_JR not in live.columns: live[ESC_JR] = ""
                    if has_envoye: live.at[t_idx,"Dossier envoyé"]=bool(val_envoye)
                    if has_appr:   live.at[t_idx,"Dossier approuvé"]=bool(val_appr)
                    if has_rfe:    live.at[t_idx,"RFE"]=bool(val_rfe)
                    if has_refuse: live.at[t_idx,"Dossier refusé"]=bool(val_refuse)
                    if has_annule: live.at[t_idx,"Dossier annulé"]=bool(val_annule)

                    write_sheet_inplace(current_path, client_target_sheet, live)
                    save_workspace_path(current_path)

                    # --- Transfert immédiat si demandé ---
                    if val_envoye and 'do_transfer_now' in locals() and do_transfer_now and (transfer_amount or 0.0) > 0:
                        try:
                            live_w = read_sheet(current_path, client_target_sheet, normalize=False).copy()
                            for c in [ESC_TR, ESC_JR]:
                                if c not in live_w.columns: live_w[c] = 0.0 if c==ESC_TR else ""
                            key = _safe_str(init.get("ID_Client"))
                            idxs = live_w.index[live_w.get("ID_Client","").astype(str)==key] if key else []
                            if len(idxs)==0:
                                msk = (live_w.get("Nom","").astype(str)==_safe_str(nom)) & \
                                      (live_w.get("Date","").astype(str)==str(d))
                                idxs = live_w.index[msk]
                            if len(idxs)==0:
                                st.info("Transfert demandé mais ligne introuvable après sauvegarde.")
                            else:
                                i = idxs[0]
                                tmp_norm = normalize_dataframe(live_w.copy())
                                disp = float(tmp_norm.loc[tmp_norm["ID_Client"].astype(str)==str(live_w.at[i,"ID_Client"])].apply(escrow_available_from_row, axis=1).iloc[0])
                                add = float(min(max(transfer_amount,0.0), disp))
                                live_w.at[i, ESC_TR] = float(pd.to_numeric(pd.Series([live_w.at[i, ESC_TR]]), errors="coerce").fillna(0.0).iloc[0] + add)
                                live_w.at[i, ESC_JR] = append_escrow_journal(live_w.loc[i], add, transfer_note)
                                write_sheet_inplace(current_path, client_target_sheet, live_w)
                                st.success(f"Transfert ESCROW immédiat effectué : {_fmt_money_us(add)} ✅")
                        except Exception as e:
                            st.error(f"Transfert immédiat — erreur : {e}")

                    st.success("Modifications enregistrées **dans le fichier**. ✅")
                    st.rerun()

    # ---- SUPPRIMER ----
    if action == "Supprimer":
        st.markdown("### 🗑️ Supprimer un client (écrit directement)")
        if live_raw.drop(columns=["_RowID"]).empty: st.info("Aucun client.")
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
                    nom = _safe_str(live_raw.at[idx,"Nom"]); dat = _safe_str(live_raw.at[idx,"Date"])
                    live = live[~((live.get("Nom","").astype(str)==nom)&(live.get("Date","").astype(str)==dat))].reset_index(drop=True)
                write_sheet_inplace(current_path, client_target_sheet, live); save_workspace_path(current_path)
                st.success("Client supprimé **dans le fichier**. ✅"); st.rerun()

# ========================= ANALYSES =========================
with tabs[2]:
    st.subheader("📊 Analyses — Volumes & Financier")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa) pour lancer les analyses."); st.stop()
    dfA_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    dfA = normalize_dataframe(dfA_raw).copy()
    if dfA.empty: st.info("Aucune donnée pour analyser."); st.stop()

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
    if sel_visa: fA = fA[fA["Visa"].astype(str).isin(sel_visa)]
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

    if agg_with_year:
        fA["Periode"] = fA["Date"].apply(lambda x: f"{x.year}-{x.month:02d}" if pd.notna(x) else "NA")
        ordre_periodes = sorted([p for p in fA["Periode"].unique() if p != "NA"]) + (["NA"] if "NA" in fA["Periode"].values else [])
    else:
        fA["Periode"] = fA["Mois"].fillna("NA")
        ordre_periodes = [f"{m:02d}" for m in range(1,13)]
        if "NA" in fA["Periode"].values: ordre_periodes += ["NA"]

    for col in [HONO, AUTRE, TOTAL, "Payé","Reste"]:
        if col in fA.columns: fA[col] = pd.to_numeric(fA[col], errors="coerce").fillna(0.0)

    def derive_statut(row) -> str:
        if bool(row.get("Dossier approuvé", False)): return "Approuvé"
        if bool(row.get("Dossier refusé", False)):   return "Refusé"
        if bool(row.get("Dossier annulé", False)):   return "Annulé"
        return "En attente"

    details = fA.copy()
    details["Statut"] = details.apply(derive_statut, axis=1)
    details_display_cols = [c for c in ["Periode","ID_Client","Nom","Visa","Date", HONO, AUTRE, TOTAL, "Payé","Reste","Statut"] if c in details.columns]

    # KPI
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Dossiers (filtrés)", f"{len(fA)}")
    k2.metric("Total (US $)", _fmt_money_us(float(fA.get(TOTAL, pd.Series(dtype=float)).sum())))
    k3.metric("Encaissements (US $)", _fmt_money_us(float(fA.get("Payé", pd.Series(dtype=float)).sum())))
    k4.metric("Solde à encaisser (US $)", _fmt_money_us(float(fA.get("Reste", pd.Series(dtype=float)).sum())))

    st.divider()

    st.markdown("### 📦 Volumes / périodes")
    def _safe_bool_sum(series): return series.fillna(False).astype(bool).sum()
    vol_all = fA.groupby("Periode").size().reindex(ordre_periodes).fillna(0).astype(int)
    vol_env = fA.groupby("Periode")["Dossier envoyé"].apply(_safe_bool_sum).reindex(ordre_periodes).fillna(0) if "Dossier envoyé" in fA.columns else pd.Series(dtype=int)
    vol_app = fA.groupby("Periode")["Dossier approuvé"].apply(_safe_bool_sum).reindex(ordre_periodes).fillna(0) if "Dossier approuvé" in fA.columns else pd.Series(dtype=int)
    vol_ref = fA.groupby("Periode")["Dossier refusé"].apply(_safe_bool_sum).reindex(ordre_periodes).fillna(0) if "Dossier refusé" in fA.columns else pd.Series(dtype=int)
    vol_ann = fA.groupby("Periode")["Dossier annulé"].apply(_safe_bool_sum).reindex(ordre_periodes).fillna(0) if "Dossier annulé" in fA.columns else pd.Series(dtype=int)

    cvol1, cvol2 = st.columns(2)
    cvol1.caption("Dossiers ouverts"); cvol1.bar_chart(pd.DataFrame({"Ouverts": vol_all}))
    vols_dict = {}
    if not vol_env.empty: vols_dict["Envoyés"] = vol_env
    if not vol_app.empty: vols_dict["Approuvés"] = vol_app
    if not vol_ref.empty: vols_dict["Refusés"]  = vol_ref
    if not vol_ann.empty: vols_dict["Annulés"]  = vol_ann
    if vols_dict: cvol2.caption("Statuts par période"); cvol2.bar_chart(pd.DataFrame(vols_dict))
    else: cvol2.info("Aucun statut disponible.")

    st.divider()

    st.markdown("### 💵 Financier par période")
    sums = fA.groupby("Periode")[[HONO, AUTRE, TOTAL, "Payé","Reste"]].sum().reindex(ordre_periodes).fillna(0.0)
    cfin1, cfin2 = st.columns(2)
    cfin1.caption("Honoraires & Autres frais (US $)"); cfin1.bar_chart(sums[[HONO, AUTRE]])
    cfin2.caption("Total vs Payé vs Reste (US $)");     cfin2.bar_chart(sums[[TOTAL, "Payé","Reste"]])

    st.divider()

    st.markdown("### 🔎 Détails (clients)")
    d1, d2 = st.columns([1,1])
    statut_filter = d1.multiselect("Filtrer par statut", ["Approuvé","Refusé","Annulé","En attente"], key="anal_statut_filter")
    search = d2.text_input("Recherche (Nom / Visa / ID)", key="anal_search")
    details_to_show = details[details_display_cols].copy()
    if statut_filter: details_to_show = details_to_show[details_to_show["Statut"].isin(statut_filter)]
    if search:
        s = search.lower()
        details_to_show = details_to_show[details_to_show.apply(lambda r: (s in str(r.get("Nom","")).lower()) or (s in str(r.get("Visa","")).lower()) or (s in str(r.get("ID_Client","")).lower()), axis=1)]
    st.dataframe(details_to_show.sort_values(["Periode","Nom"]), use_container_width=True)

# ========================= ESCROW =========================
with tabs[3]:
    st.subheader("🏦 ESCROW — dépôts sur honoraires & transferts")
    if client_target_sheet is None:
        st.info("Choisis d’abord une **feuille clients** valide (Nom & Visa)."); st.stop()

    live_raw = read_sheet(current_path, client_target_sheet, normalize=False)
    live = normalize_dataframe(live_raw).copy()

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
            with st.expander(f'🔔 {r["ID_Client"]} — {r["Nom"]} — {r["Visa"]} — ESCROW dispo: {_fmt_money_us(r["ESCROW dispo"])}'):
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
                        tmp = normalize_dataframe(live_w.copy())
                        disp = float(tmp.loc[tmp["ID_Client"].astype(str)==str(r["ID_Client"]), :].apply(escrow_available_from_row, axis=1).iloc[0])
                        add = float(min(max(amt,0.0), disp))
                        live_w.at[i, ESC_TR] = float(pd.to_numeric(pd.Series([live_w.at[i, ESC_TR]]), errors="coerce").fillna(0.0).iloc[0] + add)
                        live_w.at[i, ESC_JR] = append_escrow_journal(live_w.loc[i], add, note)
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
        show = non_env[["ID_Client","Nom","Visa","Date",HONO,"Payé",ESC_TR,"ESCROW dispo"]].copy()
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
                    "ID_Client": r["ID_Client"], "Nom": r["Nom"], "Visa": r["Visa"],
                    "Date": r.get("Date"), "Horodatage": e.get("ts"),
                    "Montant (US $)": float(e.get("amount",0.0)), "Note": e.get("note","")
                })
        jdf = pd.DataFrame(rows).sort_values("Horodatage")
        jdf["Montant (US $)"] = jdf["Montant (US $)"].map(_fmt_money_us)
        st.dataframe(jdf, use_container_width=True)
