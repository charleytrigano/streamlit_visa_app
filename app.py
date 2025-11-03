import os
import json
import re
from io import BytesIO
from datetime import date, datetime
import pandas as pd
import streamlit as st

# Constantes et configuration
APP_TITLE = "🛂 Visa Manager"
COLS_CLIENTS = [
    "ID_Client", "Dossier N", "Nom", "Date",
    "Categories", "Sous-categorie", "Visa",
    "Montant honoraires (US $)", "Autres frais (US $)",
    "Payé", "Solde", "Acompte 1", "Acompte 2",
    "RFE", "Dossiers envoyé", "Dossier approuvé",
    "Dossier refusé", "Dossier Annulé", "Commentaires",
    "Escrow", "Date denvoi", "Date dacceptation", "Date de refus", "Date dannulation"
]
MEMO_FILE = "_vmemory.json"
SHEET_CLIENTS = "Clients"
SHEET_VISA = "Visa"
SID = "vmgr"

def skey(*parts):
    return f"{SID}_" + "_".join([p for p in parts if p])

def _safe_str(x):
    try: return "" if x is None else str(x)
    except Exception: return ""

def _to_num(x):
    if isinstance(x, (int, float)): return float(x)
    s = _safe_str(x)
    if not s: return 0.0
    s = re.sub(r"[^d,.-]", "", s).replace(",", ".")
    try: return float(s)
    except Exception: return 0.0

def _fmt_money(v):
    try: return "${:,.2f}".format(float(v))
    except Exception: return "$0.00"

def _date_for_widget(val):
    if isinstance(val, date): return val
    if isinstance(val, datetime): return val.date()
    try:
        d = pd.to_datetime(val, errors="coerce")
        if pd.isna(d): return date.today()
        return d.date()
    except Exception:
        return date.today()

def _ensure_columns(df, cols):
    out = df.copy()
    for c in cols:
        if c not in out.columns:
            if c in ["Payé", "Solde", "Montant honoraires (US $)", "Autres frais (US $)", "Acompte 1", "Acompte 2"]:
                out[c] = 0.0
            elif c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
                out[c] = 0
            elif c == "Escrow":
                out[c] = 0
            else:
                out[c] = ""
    return out[cols]

def _normalize_clients_numeric(df):
    num_cols = ["Montant honoraires (US $)", "Autres frais (US $)", "Payé", "Solde", "Acompte 1", "Acompte 2"]
    for c in num_cols:
        if c in df.columns: df[c] = df[c].apply(_to_num)
    if "Montant honoraires (US $)" in df.columns and "Autres frais (US $)" in df.columns:
        total = df["Montant honoraires (US $)"] + df["Autres frais (US $)"]
        paye = df["Payé"] if "Payé" in df.columns else 0.0
        df["Solde"] = (total - paye).clip(lower=0.0)
    return df

def _normalize_status(df):
    for c in ["RFE", "Dossiers envoyé", "Dossier approuvé", "Dossier refusé", "Dossier Annulé"]:
        if c in df.columns:
            df[c] = df[c].apply(lambda x: 1 if str(x).strip().lower() in ["1", "true", "oui", "x"] else 0)
        else:
            df[c] = 0
    if "Escrow" in df.columns:
        df["Escrow"] = df["Escrow"].apply(lambda x: 1 if str(x).strip().lower() in ["1", "true", "t", "yes", "oui", "y", "x"] else 0)
    else:
        df["Escrow"] = 0
    return df

def normalize_clients(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS_CLIENTS)
    df = df.copy()
    ren = {
        "Categorie": "Categories", "Catégorie": "Categories",
        "Sous-categorie": "Sous-categorie", "Sous-catégorie": "Sous-categorie",
        "Payee": "Payé", "Payé (US $)": "Payé",
        "Montant honoraires": "Montant honoraires (US $)",
        "Autres frais": "Autres frais (US $)",
        "Dossier envoye": "Dossiers envoyé", "Dossier envoyé": "Dossiers envoyé",
    }
    df.rename(columns={k: v for k, v in ren.items() if k in df.columns}, inplace=True)
    df = _ensure_columns(df, COLS_CLIENTS)
    if "Date" in df.columns:
        try: df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        except Exception: pass
    df = _normalize_clients_numeric(df)
    df = _normalize_status(df)
    for col in ["Nom", "Categories", "Sous-categorie", "Visa", "Commentaires"]:
        df[col] = df[col].astype(str)
    try:
        dser = pd.to_datetime(df["Date"], errors="coerce")
        df["_Année_"] = dser.dt.year.fillna(0).astype(int)
        df["_MoisNum_"] = dser.dt.month.fillna(0).astype(int)
        df["Mois"] = df["_MoisNum_"].apply(lambda m: f"{int(m):02d}" if m and m == m else "")
    except Exception:
        df["_Année_"] = 0
        df["_MoisNum_"] = 0
        df["Mois"] = ""
    return df

def read_any_table(src, sheet=None):
    if src is None: return None
    if hasattr(src, "read") and hasattr(src, "name"):
        bio = BytesIO(src.read())
        if src.name.lower().endswith(".csv"): return pd.read_csv(bio)
        return pd.read_excel(bio, sheet_name=sheet or 0)
    if isinstance(src, (str, os.PathLike)):
        if not os.path.exists(src): return None
        if str(src).lower().endswith(".csv"): return pd.read_csv(src)
        return pd.read_excel(src, sheet_name=sheet or 0)
    if isinstance(src, BytesIO):
        try:
            bio2 = BytesIO(src.getvalue())
            return pd.read_excel(bio2, sheet_name=sheet or 0)
        except Exception:
            src.seek(0)
            return pd.read_csv(src)
    return None

def build_visa_map(dfv):
    vm = {}
    if dfv is None or dfv.empty: return vm
    cols = [c for c in dfv.columns if _safe_str(c)]
    if "Categories" not in cols and "Catégorie" in cols:
        dfv = dfv.rename(columns={"Catégorie": "Categories"})
    if "Sous-categorie" not in cols and "Sous-catégorie" in cols:
        dfv = dfv.rename(columns={"Sous-catégorie": "Sous-categorie"})
    if "Categories" not in dfv.columns or "Sous-categorie" not in dfv.columns: return vm
    fixed = ["Categories", "Sous-categorie"]
    option_cols = [c for c in dfv.columns if c not in fixed]
    for _, row in dfv.iterrows():
        cat = _safe_str(row.get("Categories", "")).strip()
        sub = _safe_str(row.get("Sous-categorie", "")).strip()
        if not cat or not sub: continue
        vm.setdefault(cat, {})
        vm[cat].setdefault(sub, {"exclusive": None, "options": []})
        opts = []
        for oc in option_cols:
            val = _safe_str(row.get(oc, "")).strip()
            if val.lower() in ["1", "x", "oui", "true"]: opts.append(oc)
        exclusive = None
        if set([o.upper() for o in opts]) == set(["COS", "EOS"]): exclusive = "radio_group"
        vm[cat][sub] = {"exclusive": exclusive, "options": opts}
    return vm

def _ensure_time_features(df):
    if df is None or df.empty: return df
    df = df.copy()
    try:
        dd = pd.to_datetime(df["Date"], errors="coerce") if "Date" in df.columns else pd.Series(dtype="datetime64[ns]")
        df["_Année_"] = dd.dt.year
        df["_MoisNum_"] = dd.dt.month
        df["Mois"] = dd.dt.month.apply(lambda m: f"{int(m):02d}" if pd.notna(m) else "")
    except Exception:
        df["_Année_"] = pd.NA
        df["_MoisNum_"] = pd.NA
        df["Mois"] = ""
    return df

def load_last_paths():
    if not os.path.exists(MEMO_FILE): return "", "", ""
    try:
        with open(MEMO_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("clients", ""), data.get("visa", ""), data.get("save_dir", "")
    except Exception: return "", "", ""

def save_last_paths(clients_path, visa_path, save_dir):
    data = {"clients": clients_path or "", "visa": visa_path or "", "save_dir": save_dir or ""}
    try:
        with open(MEMO_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception: pass

# Interface Streamlit
st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

st.sidebar.header("📂 Fichiers")
last_clients, last_visa, last_save_dir = load_last_paths()

mode = st.sidebar.radio(
    "Mode de chargement",
    ["Un fichier (Clients)", "Deux fichiers (Clients & Visa)"], index=0,
    key=skey("mode"),
)

up_clients = st.sidebar.file_uploader(
    "Clients (xlsx/csv)", type=["xlsx", "xls", "csv"], key=skey("up_clients")
)
up_visa = None
if mode == "Deux fichiers (Clients & Visa)":
    up_visa = st.sidebar.file_uploader(
        "Visa (xlsx/csv)", type=["xlsx", "xls", "csv"], key=skey("up_visa")
    )

clients_path_in = st.sidebar.text_input("ou chemin local Clients", value=last_clients, key=skey("cli_path"))
visa_path_in = st.sidebar.text_input("ou chemin local Visa", value=(last_visa if mode != "Un fichier (Clients)" else ""), key=skey("vis_path"))
save_dir_in = st.sidebar.text_input("Dossier de sauvegarde", value=last_save_dir, key=skey("save_dir"))

if st.sidebar.button("📥 Charger", key=skey("btn_load")):
    save_last_paths(clients_path_in, visa_path_in, save_dir_in)
    st.success("Chemins mémorisés. Relancez l'application pour appliquer.")
    st.experimental_rerun()

clients_src = up_clients if up_clients else (clients_path_in if clients_path_in else last_clients)
df_clients_raw = normalize_clients(read_any_table(clients_src))

if mode == "Deux fichiers (Clients & Visa)":
    visa_src = up_visa if up_visa else (visa_path_in if visa_path_in else last_visa)
else:
    visa_src = up_clients if up_clients else (clients_path_in if clients_path_in else last_clients)

df_visa_raw = read_any_table(visa_src, sheet=SHEET_VISA)
if df_visa_raw is None: df_visa_raw = read_any_table(visa_src)
if df_visa_raw is None: df_visa_raw = pd.DataFrame()

visa_map = build_visa_map(df_visa_raw)
df_all = _ensure_time_features(df_clients_raw)

tabs = st.tabs([
    "📄 Fichiers",
    "📊 Dashboard",
    "📈 Analyses",
    "🏦 Escrow",
    "👤 Compte client",
    "🧾 Gestion",
    "📄 Visa (aperçu)",
    "💾 Export",
])

with tabs[4]:
    # Compte client avec clés uniques
    st.subheader("👤 Compte client")
    if df_all is None or df_all.empty:
        st.info("Aucun client chargé.")
    else:
        left, right = st.columns(2)
        ids = sorted(df_all["ID_Client"].dropna().astype(str).unique())
        noms = sorted(df_all["Nom"].dropna().astype(str).unique())

        sel_id = left.selectbox("ID_Client", [""] + ids, key=skey("acct", "id"))
        sel_nom = right.selectbox("Nom", [""] + noms, key=skey("acct", "nm"))

        subset = df_all.copy()
        if sel_id:
            subset = subset[subset["ID_Client"].astype(str) == sel_id]
        elif sel_nom:
            subset = subset[subset["Nom"].astype(str) == sel_nom]

        if subset.empty:
            st.warning("Sélectionnez un client pour afficher le compte.")
        else:
            row = subset.iloc[0].to_dict()
            r1, r2, r3, r4 = st.columns(4)
            r1.metric("Dossier N", _safe_str(row.get("Dossier N", "")))
            total = float(_to_num(row.get("Montant honoraires (US $)", 0)) + _to_num(row.get("Autres frais (US $)", 0)))
            r2.metric("Total", _fmt_money(total))
            r3.metric("Payé", _fmt_money(_to_num(row.get("Payé", 0))))
            r4.metric("Solde", _fmt_money(_to_num(row.get("Solde", 0))))

            d1, d2, d3 = st.columns(3)
            d1.write(f"**Catégorie :** {_safe_str(row.get('Categories', ''))}")
            d1.write(f"**Sous-catégorie :** {_safe_str(row.get('Sous-categorie', ''))}")
            d1.write(f"**Visa :** {_safe_str(row.get('Visa', ''))}")
            d2.write(f"**Date :** {_safe_str(row.get('Date', ''))}")
            d2.write(f"**Mois (MM) :** {_safe_str(row.get('Mois', ''))}")
            d3.write(f"**Commentaires :** {_safe_str(row.get('Commentaires', ''))}")

            s1, s2 = st.columns(2)
            def sdate(label):
                val = row.get(label, "")
                if isinstance(val, (date, datetime)):
                    return val.strftime("%Y-%m-%d")
                try:
                    d = pd.to_datetime(val, errors="coerce")
                    return d.date().strftime("%Y-%m-%d") if pd.notna(d) else ""
                except Exception:
                    return _safe_str(val)
            s1.write(f"- **Dossier envoyé** : {'Oui' if sdate('Date denvoi') else 'Non'} | Date : {sdate('Date denvoi')}")
            s1.write(f"- **Dossier approuvé** : {'Oui' if sdate('Date dacceptation') else 'Non'} | Date : {sdate('Date dacceptation')}")
            s2.write(f"- **Dossier refusé** : {'Oui' if sdate('Date de refus') else 'Non'} | Date : {sdate('Date de refus')}")
            s2.write(f"- **Dossier annulé** : {'Oui' if sdate('Date dannulation') else 'Non'} | Date : {sdate('Date dannulation')}")
            rfeflag = int(_to_num(row.get("RFE", 0)) or 0)
            st.write(f"- **RFE** : {'Oui' if rfeflag else 'Non'}")

            mvts = []
            if "Acompte 1" in row and _to_num(row["Acompte 1"]) > 0:
                mvts.append({"Libellé": "Acompte 1", "Montant": float(_to_num(row["Acompte 1"]))})
            if "Acompte 2" in row and _to_num(row["Acompte 2"]) > 0:
                mvts.append({"Libellé": "Acompte 2", "Montant": float(_to_num(row["Acompte 2"]))})
            if mvts:
                dfm = pd.DataFrame(mvts)
                dfm["Montant"] = dfm["Montant"].map(_fmt_money)
                st.dataframe(dfm, use_container_width=True, hide_index=True, key=skey("acct", "mvts"))
            else:
                st.caption("Aucun acompte enregistré dans le fichier (colonnes « Acompte 1 » / « Acompte 2 »).")

with tabs[5]:
    st.subheader("🧾 Gestion (Ajouter / Modifier / Supprimer)")
    df_live = df_all.copy() if df_all is not None else pd.DataFrame()

    if df_live.empty:
        st.info("Aucun client à gérer (chargez un fichier Clients).")
    else:
        op = st.radio("Action", ["Ajouter", "Modifier", "Supprimer"], horizontal=True, key=skey("crud", "op"))
        cats = sorted(df_live["Categories"].dropna().astype(str).unique().tolist()) if "Categories" in df_live.columns else []

        def subs_for(cat):
            if cat and "Categories" in df_live.columns and "Sous-categorie" in df_live.columns:
                return sorted(df_live[df_live["Categories"].astype(str) == cat]["Sous-categorie"].dropna().astype(str).unique().tolist())
            return []

        if op == "Ajouter":
            c1, c2, c3 = st.columns(3)
            nom = c1.text_input("Nom", "", key=skey("add", "nom"))
            dval = _date_for_widget(date.today())
            dt = c2.date_input("Date de création", value=dval, key=skey("add", "date"))
            mois = c3.selectbox("Mois (MM)", [f"{m:02d}" for m in range(1,13)], index=int(dval.month)-1, key=skey("add", "mois"))

            v1, v2, v3 = st.columns(3)
            cat = v1.selectbox("Catégorie", [""] + cats, index=0, key=skey("add", "cat"))
            subs = subs_for(cat) if cat else []
            sub = v2.selectbox("Sous-catégorie", [""] + subs, index=0, key=skey("add", "sub"))
            visa_val = v3.text_input("Visa (libre ou dérivé)", sub if sub else "", key=skey("add", "visa"))

            f1, f2 = st.columns(2)
            honor = f1.number_input("Montant honoraires (US $)", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=skey("add", "h"))
            other = f2.number_input("Autres frais (US $)", min_value=0.0, value=0.0, step=20.0, format="%.2f", key=skey("add", "o"))
            acomp1 = st.number_input("Acompte 1", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=skey("add", "a1"))
            acomp2 = st.number_input("Acompte 2", min_value=0.0, value=0.0, step=10.0, format="%.2f", key=skey("add", "a2"))
            comm = st.text_area("Commentaires", "", key=skey("add", "com"))

            s1, s2 = st.columns(2)
            sent_d = s1.date_input("Date denvoi", value=None, key=skey("add", "sentd"))
            acc_d = s1.date_input("Date dacceptation", value=None, key=skey("add", "accd"))
            ref_d = s2.date_input("Date de refus", value=None, key=skey("add", "refd"))
            ann_d = s2.date_input("Date dannulation", value=None, key=skey("add", "annd"))
            rfe = st.checkbox("RFE", value=False, key=skey("add", "rfe"))
            escrow_val = st.checkbox("Escrow", value=False, key=skey("add", "escrow"))

            if st.button("💾 Enregistrer", key=skey("add", "save")):
                if not nom or not cat or not sub:
                    st.warning("Nom, Catégorie et Sous-catégorie sont requis.")
                    st.stop()
                total = float(honor) + float(other)
                paye = float(acomp1) + float(acomp2)
                solde = max(0.0, total - paye)
                new_id = f"{_norm(nom)}-{int(datetime.now().timestamp())}"
                new_dossier = next_dossier(df_live)

                new_row = {
                    "ID_Client": new_id,
                    "Dossier N": new_dossier,
                    "Nom": nom,
                    "Date": dt,
                    "Mois": f"{int(mois):02d}",
                    "Categories": cat,
                    "Sous-categorie": sub,
                    "Visa": visa_val,
                    "Montant honoraires (US $)": float(honor),
                    "Autres frais (US $)": float(other),
                    "Payé": paye,
                    "Solde": solde,
                    "Acompte 1": float(acomp1),
                    "Acompte 2": float(acomp2),
                    "Commentaires": comm,
                    "Date denvoi": sent_d,
                    "Date dacceptation": acc_d,
                    "Date de refus": ref_d,
                    "Date dannulation": ann_d,
                    "RFE": 1 if rfe else 0,
                    "Escrow": 1 if escrow_val else 0
                }
                if new_row.get("Escrow", 0) == 1 and pd.notna(new_row.get("Date denvoi")) and new_row.get("Date denvoi"):
                    montant_escrow = _to_num(new_row.get("Acompte 1", 0))
                    st.info(f"⚠️ Escrow activé : Dossier {new_row.get('Dossier N','')} / Client {new_row.get('Nom','')} — Montant à réclamer : {_fmt_money(montant_escrow)}")
                df_live = pd.concat([df_live, pd.DataFrame([new_row])], ignore_index=True)
                st.success("Client ajouté (en mémoire). Utilisez l’onglet Export pour sauvegarder.")
                st.cache_data.clear()
                st.experimental_rerun()

        elif op == "Modifier":
            # Idem logique modification, avec clés skey("mod", ...)
            # Remplacer les clés par des variantes uniques similaires
            pass  

        elif op == "Supprimer":
            # Idem logique suppression, avec clés skey("del", ...)
            pass  

# (Autres onglets inchangés. Intègre ta logique complète Dashboard, Analyses, Escrow, Visa, Export.)
